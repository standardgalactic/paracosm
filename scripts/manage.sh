#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION_FILE="$ROOT_DIR/VERSION"
README_FILE="$ROOT_DIR/README.md"

GENERATED_ARTIFACTS=(
  "$ROOT_DIR/protocols.txt"
  "$ROOT_DIR/protocols.html"
  "$ROOT_DIR/incomplete-protocols.txt"
  "$ROOT_DIR/incomplete-protocols.html"
  "$ROOT_DIR/protocols-draft.txt"
  "$ROOT_DIR/protocols-draft.html"
  "$ROOT_DIR/source-control.txt"
  "$ROOT_DIR/source-control.html"
  "$ROOT_DIR/workspace-projects.txt"
  "$ROOT_DIR/workspace-projects.html"
)

REQUIRED_SCRIPTS=(
  "$ROOT_DIR/processing/get-context.sh"
  "$ROOT_DIR/background/get-context.sh"
  "$ROOT_DIR/workspace/get-context.sh"
)

usage() {
  cat <<'EOF'
Repository manager

Usage:
  scripts/manage.sh clean
  scripts/manage.sh build
  scripts/manage.sh run <processing|background|workspace|all>
  scripts/manage.sh release
  scripts/manage.sh version show
  scripts/manage.sh version bump
  scripts/manage.sh version minor
  scripts/manage.sh version major
EOF
}

err() {
  echo "[error] $*" >&2
}

warn() {
  echo "[warn] $*" >&2
}

info() {
  echo "[info] $*"
}

# Portable relative-path helper. GNU realpath supports --relative-to, but
# BSD/macOS realpath does not, so fall back to python3, then perl, then
# just print the absolute path if neither is available.
relpath() {
  local target="$1"
  if realpath --relative-to="$ROOT_DIR" "$target" >/dev/null 2>&1; then
    realpath --relative-to="$ROOT_DIR" "$target"
  elif command -v python3 >/dev/null 2>&1; then
    python3 -c "import os,sys; print(os.path.relpath(sys.argv[1], sys.argv[2]))" "$target" "$ROOT_DIR"
  elif command -v perl >/dev/null 2>&1; then
    perl -MFile::Spec -e 'print File::Spec->abs2rel($ARGV[0], $ARGV[1]), "\n"' "$target" "$ROOT_DIR"
  else
    echo "$target"
  fi
}

require_version_file() {
  if [[ ! -f "$VERSION_FILE" ]]; then
    err "VERSION file not found at $VERSION_FILE"
    exit 1
  fi
}

current_version() {
  require_version_file
  tr -d '[:space:]' < "$VERSION_FILE"
}

is_semver() {
  local v="$1"
  [[ "$v" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]
}

set_version() {
  local v="$1"
  if ! is_semver "$v"; then
    err "Invalid semantic version: $v"
    exit 1
  fi

  printf '%s\n' "$v" > "$VERSION_FILE"

  if [[ -f "$README_FILE" ]]; then
    if grep -q '^Version: ' "$README_FILE"; then
      sed -i.bak -E "s/^Version: .*/Version: ${v}/" "$README_FILE"
      rm -f "$README_FILE.bak"
    else
      tmp_file="$(mktemp)"
      {
        awk 'NR==1 {print; print ""; print "Version: '"$v"'"; next} {print}' "$README_FILE"
      } > "$tmp_file"
      mv "$tmp_file" "$README_FILE"
    fi
  fi

  info "Version set to $v"
}

bump_version() {
  local mode="$1"
  local v
  v="$(current_version)"
  local major minor patch
  IFS='.' read -r major minor patch <<< "$v"

  case "$mode" in
    patch)
      patch=$((patch + 1))
      ;;
    minor)
      minor=$((minor + 1))
      patch=0
      ;;
    major)
      major=$((major + 1))
      minor=0
      patch=0
      ;;
    *)
      err "Unknown bump mode: $mode"
      exit 1
      ;;
  esac

  set_version "${major}.${minor}.${patch}"
}

run_mode() {
  local mode="$1"
  case "$mode" in
    processing|background|workspace)
      local script="$ROOT_DIR/$mode/get-context.sh"
      if [[ ! -f "$script" ]]; then
        err "Missing script: $script"
        return 1
      fi
      info "Running $script"
      (cd "$ROOT_DIR/$mode" && bash ./get-context.sh)
      ;;
    all)
      run_mode processing
      run_mode background
      run_mode workspace
      ;;
    *)
      err "Unknown run mode: $mode (expected one of: processing, background, workspace, all)"
      usage
      return 1
      ;;
  esac
}

clean() {
  local removed=0
  for artifact in "${GENERATED_ARTIFACTS[@]}"; do
    if [[ -f "$artifact" ]]; then
      rm -f "$artifact"
      info "Removed $(relpath "$artifact")"
      removed=$((removed + 1))
    fi
  done
  info "Clean complete ($removed files removed)"
}

verify_scripts() {
  local missing=0
  for script in "${REQUIRED_SCRIPTS[@]}"; do
    if [[ ! -f "$script" ]]; then
      err "Missing required script: $(relpath "$script")"
      missing=1
    fi
  done
  return $missing
}

verify_artifacts() {
  local missing=0
  for artifact in "${GENERATED_ARTIFACTS[@]}"; do
    if [[ ! -f "$artifact" ]]; then
      warn "Missing generated artifact: $(relpath "$artifact")"
      missing=1
    fi
  done
  return $missing
}

verify_tex_inputs() {
  local status=0
  local tex_main="$ROOT_DIR/main.tex"

  if [[ ! -f "$tex_main" ]]; then
    info "main.tex not present; skipping TeX input verification"
    return 0
  fi

  # Recursively follows \input{} references starting from main.tex, so that
  # a file missing several levels deep (e.g. a per-book file referenced by a
  # cycle index file that main.tex itself only reaches indirectly) is caught,
  # not just files main.tex references directly.
  local visited=$'main.tex\n'
  local queue=("main.tex")

  while [[ ${#queue[@]} -gt 0 ]]; do
    local current="${queue[0]}"
    queue=("${queue[@]:1}")
    local current_path="$ROOT_DIR/$current"
    [[ -f "$current_path" ]] || continue

    while IFS= read -r ref; do
      [[ -z "$ref" ]] && continue
      local target="${ref%.tex}.tex"
      case $'\n'"$visited" in
        *$'\n'"$target"$'\n'*)
          continue
          ;;
      esac
      visited+="$target"$'\n'
      if [[ ! -f "$ROOT_DIR/$target" ]]; then
        err "Unresolved TeX input: ${target} (referenced from ${current})"
        status=1
      else
        queue+=("$target")
      fi
    done < <(grep -oE '\\input\{[^}]+' "$current_path" | sed -E 's/^\\input\{//')
  done

  return $status
}

build() {
  local failed=0

  info "Verifying scripts..."
  if ! verify_scripts; then
    failed=1
  fi

  info "Verifying generated artifacts..."
  if ! verify_artifacts; then
    info "Some generated artifacts are missing (non-fatal). Run: scripts/manage.sh run <mode>"
  fi

  info "Verifying TeX input graph..."
  if ! verify_tex_inputs; then
    failed=1
  fi

  if [[ $failed -ne 0 ]]; then
    err "Build verification failed"
    return 1
  fi

  info "Build verification passed"
}

release() {
  build
  local v
  v="$(current_version)"
  info "Release checks passed for version $v"

  if command -v git >/dev/null 2>&1 && git -C "$ROOT_DIR" rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    local branch
    branch="$(git -C "$ROOT_DIR" rev-parse --abbrev-ref HEAD 2>/dev/null || echo "unknown")"
    info "Current branch: $branch"
    if [[ -n "$(git -C "$ROOT_DIR" status --porcelain 2>/dev/null)" ]]; then
      info "Working tree has uncommitted changes — commit before tagging"
    else
      info "Working tree is clean"
    fi
  fi

  info "Next steps:"
  info "  1) Commit pending changes"
  info "  2) Tag release: git tag v${v}"
  info "  3) Push tag: git push origin v${v}"
}

main() {
  local cmd="${1:-help}"

  case "$cmd" in
    help|-h|--help)
      usage
      ;;
    clean)
      clean
      ;;
    build)
      build
      ;;
    run)
      local mode="${2:-}"
      if [[ -z "$mode" ]]; then
        err "Missing run mode"
        usage
        exit 1
      fi
      run_mode "$mode"
      ;;
    release)
      release
      ;;
    version)
      local sub="${2:-show}"
      case "$sub" in
        show)
          current_version
          ;;
        bump)
          bump_version patch
          ;;
        minor)
          bump_version minor
          ;;
        major)
          bump_version major
          ;;
        *)
          err "Unknown version subcommand: $sub"
          usage
          exit 1
          ;;
      esac
      ;;
    *)
      err "Unknown command: $cmd"
      usage
      exit 1
      ;;
  esac
}

main "$@"

#!/usr/bin/env bash
#
# apply-paracosm-history.sh (v2)
#
# Reconstruct the REAL repository changes represented by a unified patch,
# committing them incrementally in semantic LaTeX sections.
#
# v2 adds resumability: Git itself is the recovery database. Every generated
# commit records base-commit, patch-sha256, classifier, filename, and
# state-sha256 in its message. On restart, if HEAD is a reconstruction
# commit from this same run (same base commit, same patch hash, same
# classifier), the script regenerates the plan from that immutable base,
# verifies the full chain of already-made commits against it, auto-recovers
# a partially-written file from a mid-stage crash, and resumes from the
# next stage. If HEAD is not a matching reconstruction commit, it starts a
# fresh run exactly like v1.
#
# Nothing is pushed.
#
# Example:
#
#   git switch -c paracosm-reconstructed-history <BASE_COMMIT>
#   ./apply-paracosm-history.sh --source paracosm-fixes.patch --dry-run
#   ./apply-paracosm-history.sh --source paracosm-fixes.patch
#
#   # if interrupted, just rerun the same command -- it resumes safely
#   ./apply-paracosm-history.sh --source paracosm-fixes.patch
#

set -Eeuo pipefail

###############################################################################
# Defaults
###############################################################################

SOURCE="paracosm-fixes.patch"
PREFIX="PARA"
DELAY="${DELAY:-0}"
DRY_RUN=0

CLASSIFIER_VERSION="paracosm-real-tree-v2"

###############################################################################
# Usage
###############################################################################

usage() {
    cat <<'EOF'
Usage:

  ./apply-paracosm-history.sh [options]

Options:

  --source FILE       Unified patch to reconstruct. Default: paracosm-fixes.patch
  --prefix NAME        Commit ID prefix. Default: PARA
  --delay SECONDS      Pause between commits. Default: 0
  --dry-run            Parse and classify without modifying files or Git.
  -h, --help            Show this help.

Recommended workflow:

  git switch -c paracosm-reconstructed-history <PRE_PATCH_COMMIT>
  ./apply-paracosm-history.sh --source /path/to/paracosm-fixes.patch

This version is resumable: if interrupted, rerun the identical command.
The script inspects HEAD; if it is a reconstruction commit from the same
base commit, same patch (by SHA-256), and same classifier version, it
verifies the completed chain and resumes from the next stage. Otherwise
it starts a fresh run (which requires a clean tracked working tree, same
as before).
EOF
}

###############################################################################
# Arguments
###############################################################################

while [[ $# -gt 0 ]]; do
    case "$1" in
        --source) SOURCE="${2:?Missing value for --source}"; shift 2 ;;
        --prefix) PREFIX="${2:?Missing value for --prefix}"; shift 2 ;;
        --delay) DELAY="${2:?Missing value for --delay}"; shift 2 ;;
        --dry-run) DRY_RUN=1; shift ;;
        -h|--help) usage; exit 0 ;;
        *)
            echo "Unknown argument: $1" >&2
            usage >&2
            exit 2
            ;;
    esac
done

###############################################################################
# Requirements
###############################################################################

for cmd in git python3 sha256sum; do
    command -v "$cmd" >/dev/null 2>&1 || {
        echo "error: required command not found: $cmd" >&2
        exit 1
    }
done

[[ -f "$SOURCE" ]] || {
    echo "error: source patch not found: $SOURCE" >&2
    exit 1
}

if ! [[ "$DELAY" =~ ^[0-9]+([.][0-9]+)?$ ]]; then
    echo "error: --delay must be a non-negative number" >&2
    exit 1
fi

###############################################################################
# Repository
###############################################################################

REPO_ROOT="$(git rev-parse --show-toplevel 2>/dev/null)" || {
    echo "error: not inside a Git repository" >&2
    exit 1
}

SOURCE="$(python3 - "$SOURCE" <<'PY'
import os, sys
print(os.path.abspath(sys.argv[1]))
PY
)"

cd "$REPO_ROOT"

SOURCE_SHA="$(sha256sum "$SOURCE" | awk '{print $1}')"
CURRENT_HEAD="$(git rev-parse HEAD 2>/dev/null || echo "")"

###############################################################################
# Temporary workspace
###############################################################################

TMPDIR="$(mktemp -d "${TMPDIR:-/tmp}/paracosm-apply.XXXXXXXX")"
cleanup() { rm -rf "$TMPDIR"; }
trap cleanup EXIT

PLAN="$TMPDIR/plan.tsv"
STATES="$TMPDIR/states"
mkdir -p "$STATES"

###############################################################################
# Detect resume state by inspecting HEAD's commit message
###############################################################################

RESUME=0
BASE_COMMIT="$CURRENT_HEAD"
COMPLETED=0
HEAD_MSG=""

if [[ -n "$CURRENT_HEAD" ]]; then
    HEAD_MSG="$(git log -1 --format=%B "$CURRENT_HEAD" 2>/dev/null || echo "")"
fi

# HEAD_MSG is passed to Python via a temp file (not stdin) because stdin is
# already used to supply the parser script itself via the heredoc below.
HEAD_MSG_FILE="$TMPDIR/head-msg.txt"
printf '%s' "$HEAD_MSG" > "$HEAD_MSG_FILE"

# Parse HEAD's message for reconstruction metadata. Prints five TSV fields:
# stage_id, filename, classifier, patch_sha, base_commit -- or a single
# empty line if HEAD does not look like a reconstruction commit.
HEAD_META="$(
    python3 - "$PREFIX" "$HEAD_MSG_FILE" <<'PY'
import re
import sys
from pathlib import Path

prefix = sys.argv[1]
msg = Path(sys.argv[2]).read_text(encoding="utf-8")

subject_re = re.compile(rf'^{re.escape(prefix)}-(\d+)\b')
lines = msg.splitlines()

if not lines or not subject_re.match(lines[0]):
    print("")
    sys.exit(0)

# Strip leading zeros so this can never be misread as octal by bash
# arithmetic later (e.g. "0200" as octal is 128 in decimal).
stage_id = str(int(subject_re.match(lines[0]).group(1)))

def field(name):
    m = re.search(rf'^{name}:\s*(.*)$', msg, re.MULTILINE)
    return m.group(1).strip() if m else ""

filename = field("file")
classifier = field("classifier")
patch_sha = field("patch-sha256")
base_commit = field("base-commit")

if not (filename and classifier and patch_sha and base_commit):
    print("")
    sys.exit(0)

print("\t".join([stage_id, filename, classifier, patch_sha, base_commit]))
PY
)"

if [[ -n "$HEAD_META" ]]; then
    IFS=$'\t' read -r HEAD_STAGE HEAD_FILE HEAD_CLASSIFIER HEAD_PATCH_SHA HEAD_BASE_COMMIT <<< "$HEAD_META"

    if [[ "$HEAD_CLASSIFIER" != "$CLASSIFIER_VERSION" ]]; then
        echo "error: HEAD is a reconstruction commit from classifier '$HEAD_CLASSIFIER'," >&2
        echo "       but this script is classifier '$CLASSIFIER_VERSION'." >&2
        echo "       Refusing to resume with a mismatched classifier." >&2
        exit 1
    fi

    if [[ "$HEAD_PATCH_SHA" != "$SOURCE_SHA" ]]; then
        echo "error: HEAD's reconstruction commit was built from a patch with" >&2
        echo "       SHA-256 $HEAD_PATCH_SHA," >&2
        echo "       but the supplied --source has SHA-256 $SOURCE_SHA." >&2
        echo "       Refusing to resume against a different patch." >&2
        exit 1
    fi

    RESUME=1
    BASE_COMMIT="$HEAD_BASE_COMMIT"
    COMPLETED="$HEAD_STAGE"

    echo "[info] Resuming reconstruction: HEAD is stage $PREFIX-$(printf '%04d' "$COMPLETED"), base commit $BASE_COMMIT"
fi

###############################################################################
# Safety (fresh runs only need full cleanliness; resumed runs are checked
# again below once we know which single file the next stage is allowed to
# have touched).
###############################################################################

if [[ "$RESUME" == "0" ]]; then
    if [[ -n "$(git status --porcelain --untracked-files=no)" ]]; then
        echo "error: tracked working tree is not clean." >&2
        echo >&2
        echo "Commit, stash, or restore existing tracked changes first." >&2
        exit 1
    fi
fi

###############################################################################
# Build final reference tree at BASE_COMMIT (immutable regardless of resume)
###############################################################################

REFERENCE_WORKTREE="$TMPDIR/reference"
git worktree add --detach --quiet "$REFERENCE_WORKTREE" "$BASE_COMMIT"
REFERENCE_ADDED=1
cleanup_worktree() {
    if [[ "${REFERENCE_ADDED:-0}" == "1" ]]; then
        git worktree remove --force "$REFERENCE_WORKTREE" >/dev/null 2>&1 || true
    fi
    rm -rf "$TMPDIR"
}
trap cleanup_worktree EXIT

( cd "$REFERENCE_WORKTREE" && git apply --whitespace=nowarn "$SOURCE" )

# File list at BASE_COMMIT, used for existed_before checks -- this must be
# evaluated against the immutable base commit, never against the live
# working tree, or plan regeneration on resume would misclassify files
# this same reconstruction already created.
BASE_TREE_FILES="$TMPDIR/base-tree-files.txt"
git ls-tree -r --name-only "$BASE_COMMIT" > "$BASE_TREE_FILES"

###############################################################################
# Parse patch and construct semantic plan
###############################################################################

python3 - \
    "$SOURCE" \
    "$BASE_TREE_FILES" \
    "$REFERENCE_WORKTREE" \
    "$STATES" \
    "$PLAN" \
    "$PREFIX" <<'PY'
from __future__ import annotations

import hashlib
import re
import sys
from dataclasses import dataclass
from pathlib import Path


patch_path = Path(sys.argv[1])
base_tree_files_path = Path(sys.argv[2])
reference_root = Path(sys.argv[3])
states_root = Path(sys.argv[4])
plan_path = Path(sys.argv[5])
prefix_name = sys.argv[6]

base_tree_files = set(
    line for line in base_tree_files_path.read_text(encoding="utf-8").splitlines() if line
)

raw_patch = patch_path.read_text(encoding="utf-8", errors="surrogateescape")

file_re = re.compile(r"^diff --git a/(.*?) b/(.*?)$", re.MULTILINE)
patch_files = []
for match in file_re.finditer(raw_patch):
    new_name = match.group(2)
    if new_name not in patch_files:
        patch_files.append(new_name)

HEADER_RE = re.compile(
    r"""
    ^
    \\(?P<kind>part|chapter|section|subsection|subsubsection)
    \*?
    \{
    (?P<title>.*)
    \}
    \s*$
    """,
    re.VERBOSE,
)


def clean_title(value: str) -> str:
    value = value.strip()
    for macro in (r"\textit{", r"\textbf{", r"\emph{", r"\mathbf{", r"\mathrm{"):
        value = value.replace(macro, "")
    value = value.replace("{", "").replace("}", "")
    value = value.replace("---", "—").replace("--", "–")
    value = re.sub(r"\s+", " ", value).strip()
    if len(value) > 86:
        value = value[:83].rstrip() + "..."
    return value or "untitled"


PATTERNS = {
    "OPENING": [r"\bpreface\b", r"\bprologue\b", r"\bintroduction\b", r"\breading note\b"],
    "ORIENTATION": [r"\boverview\b", r"\bconceptual map\b", r"\broadmap\b", r"\borientation\b", r"\bdiagrammatic\b"],
    "PROVENANCE": [r"\bprovenance\b", r"\bsource draft", r"\beditorial note\b", r"\btextual history\b"],
    "PROOF-OBJECT": [r"\bproof\b", r"\btheorem\b", r"\blemma\b", r"\bcorollary\b", r"\bproposition\b"],
    "DEFINITION": [r"\bdefinition\b", r"\baxiom\b", r"\bobjects?\b", r"\bmorphisms?\b", r"\bconstruction\b", r"\bformalism\b"],
    "DYNAMICS": [r"\bcurvature\b", r"\bflow\b", r"\bdynamics?\b", r"\bgeodesic", r"\bentropy\b", r"\bdiffusion\b", r"\bevolution\b", r"\btime\b", r"\bgradient\b", r"\bthermodynamic"],
    "BRIDGE": [r"\btoward\b", r"\btransition\b", r"\bcomparison\b", r"\bcorrespondence\b", r"\btranslation\b", r"\bconnection\b", r"\bfrom .* to\b"],
    "SYNTHESIS": [r"\bsynthesis\b", r"\bintegration\b", r"\bclosure\b", r"\bcomposition\b", r"\bcoherence\b", r"\bcolimit", r"\blimit", r"\bgluing\b", r"\buniversal"],
    "REFLECTION": [r"\bepilogue\b", r"\bethical\b", r"\bethics\b", r"\bcare\b", r"\bcompassion\b", r"\bmoral\b", r"\bmeaning\b"],
    "APPARATUS": [r"\bbibliograph", r"\bnotation\b", r"\bcomputational\b", r"\bimplementation\b", r"\breference\b", r"\btemplate"],
}

PRIORITY = ["PROVENANCE", "PROOF-OBJECT", "OPENING", "ORIENTATION", "APPARATUS", "REFLECTION", "DEFINITION", "DYNAMICS", "BRIDGE", "SYNTHESIS"]


def classify(title: str, body: str) -> str:
    title_l, body_l = title.lower(), body.lower()
    scores = {c: 0 for c in PATTERNS}
    for category, patterns in PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, title_l):
                scores[category] += 8
            scores[category] += min(len(re.findall(pattern, body_l)), 5)
    best = max(scores.values(), default=0)
    if best == 0:
        return "ARGUMENT"
    winners = {c for c, s in scores.items() if s == best}
    for category in PRIORITY:
        if category in winners:
            return category
    return "ARGUMENT"


@dataclass
class Stage:
    filename: str
    category: str
    unit: str
    title: str
    start_line: int
    end_line: int
    content: bytes
    final: bool = False


stages: list[Stage] = []

for filename in patch_files:
    final_path = reference_root / filename

    if not final_path.exists():
        stages.append(Stage(filename, "DELETE", "FILE", f"Remove {filename}", 0, 0, b"", True))
        continue

    final_bytes = final_path.read_bytes()

    try:
        final_text = final_bytes.decode("utf-8")
    except UnicodeDecodeError:
        stages.append(Stage(filename, "BINARY", "FILE", f"Materialize {filename}", 1, 1, final_bytes, True))
        continue

    # Existed-before check against the immutable BASE_COMMIT tree, not the
    # live working directory, so plan regeneration is identical whether this
    # is a fresh run or a resumed one.
    existed_before = filename in base_tree_files

    suffix = Path(filename).suffix.lower()

    if existed_before:
        category = "REVISION"
        if suffix in {".md", ".markdown"}:
            category = "EDITORIAL"
        elif Path(filename).name.lower() == "version":
            category = "RELEASE"
        stages.append(Stage(filename, category, "FILE", f"Revise {filename}", 1, len(final_text.splitlines()), final_bytes, True))
        continue

    if suffix != ".tex":
        category = "FILE-ADD"
        if suffix in {".md", ".markdown"}:
            category = "EDITORIAL"
        elif Path(filename).name.lower() == "version":
            category = "RELEASE"
        stages.append(Stage(filename, category, "FILE", f"Add {filename}", 1, len(final_text.splitlines()), final_bytes, True))
        continue

    lines = final_bytes.splitlines(keepends=True)
    boundaries = []
    for lineno, raw in enumerate(lines, start=1):
        line = raw.decode("utf-8", errors="replace").rstrip("\r\n")
        m = HEADER_RE.match(line)
        if m:
            boundaries.append((lineno, m.group("kind").upper(), clean_title(m.group("title"))))

    if not boundaries:
        stages.append(Stage(filename, "FILE-ADD", "FILE", f"Add {filename}", 1, len(lines), final_bytes, True))
        continue

    first_boundary = boundaries[0][0]
    accumulated = bytearray()

    if first_boundary > 1:
        preamble = b"".join(lines[:first_boundary - 1])
        accumulated.extend(preamble)
        stages.append(Stage(filename, "OPENING", "PREAMBLE", f"{Path(filename).name} preamble", 1, first_boundary - 1, preamble))

    for i, (start, unit, title) in enumerate(boundaries):
        end = boundaries[i + 1][0] - 1 if i + 1 < len(boundaries) else len(lines)
        section_bytes = b"".join(lines[start - 1:end])
        accumulated.extend(section_bytes)
        body = section_bytes.decode("utf-8", errors="replace")
        category = classify(title, body)
        stages.append(Stage(filename, category, unit, title, start, end, bytes(accumulated), i == len(boundaries) - 1))

rows = []
for index, stage in enumerate(stages, start=1):
    state_name = f"{index:05d}.state"
    (states_root / state_name).write_bytes(stage.content)
    digest = hashlib.sha256(stage.content).hexdigest()
    safe_title = stage.title.replace("\t", " ").replace("\n", " ")
    safe_filename = stage.filename.replace("\t", " ").replace("\n", " ")
    rows.append("\t".join([
        str(index), stage.category, stage.unit, safe_title, safe_filename,
        str(stage.start_line), str(stage.end_line), str(len(stage.content)),
        digest, state_name, "1" if stage.final else "0",
    ]))

plan_path.write_text("\n".join(rows) + "\n", encoding="utf-8")
print(f"files\t{len(patch_files)}")
print(f"stages\t{len(stages)}")
PY

TOTAL_STAGES="$(wc -l < "$PLAN" | tr -d ' ')"

echo
echo "Source patch: $SOURCE"
echo "Base commit:  $BASE_COMMIT"
echo "Stages:       $TOTAL_STAGES"
echo "Patch SHA:    $SOURCE_SHA"
echo

###############################################################################
# If resuming, verify the FULL completed chain against the regenerated plan,
# not just HEAD -- catches a tampered or unexpectedly divergent history.
###############################################################################

if [[ "$RESUME" == "1" && "$COMPLETED" -gt 0 ]]; then
    if [[ "$COMPLETED" -gt "$TOTAL_STAGES" ]]; then
        echo "error: HEAD claims stage $COMPLETED but the regenerated plan only has $TOTAL_STAGES stages." >&2
        echo "       This should not happen when classifier and patch hashes matched; refusing to proceed." >&2
        exit 1
    fi

    CHAIN_LOG="$TMPDIR/chain.log"
    git log --format='%H' -n "$COMPLETED" --reverse "$CURRENT_HEAD" > "$CHAIN_LOG"

    CHAIN_MSGS="$TMPDIR/chain-msgs.txt"
    : > "$CHAIN_MSGS"
    while IFS= read -r commit_hash; do
        {
            echo "===COMMIT-START==="
            git log -1 --format=%B "$commit_hash"
            echo "===COMMIT-END==="
        } >> "$CHAIN_MSGS"
    done < "$CHAIN_LOG"

    python3 - "$PREFIX" "$PLAN" "$CHAIN_MSGS" "$COMPLETED" <<'PY'
import re
import sys

prefix, plan_path, chain_path, completed = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4])

plan = {}
with open(plan_path, encoding="utf-8") as f:
    for line in f:
        parts = line.rstrip("\n").split("\t")
        if len(parts) < 11:
            continue
        idx = int(parts[0])
        plan[idx] = {"filename": parts[4], "state_sha256": parts[8]}

raw = open(chain_path, encoding="utf-8").read()
blocks = re.findall(r"===COMMIT-START===\n(.*?)\n===COMMIT-END===", raw, re.DOTALL)

if len(blocks) != completed:
    print(f"error: expected {completed} commits in chain, found {len(blocks)}", file=sys.stderr)
    sys.exit(1)

subject_re = re.compile(rf'^{re.escape(prefix)}-(\d+)\b')

for expected_idx, block in enumerate(blocks, start=1):
    lines = block.splitlines()
    if not lines or not subject_re.match(lines[0]):
        print(f"error: commit at chain position {expected_idx} is not a {prefix}-NNNN reconstruction commit", file=sys.stderr)
        sys.exit(1)

    actual_idx = int(subject_re.match(lines[0]).group(1))

    def field(name):
        m = re.search(rf'^{name}:\s*(.*)$', block, re.MULTILINE)
        return m.group(1).strip() if m else ""

    actual_file = field("file")
    actual_sha = field("state-sha256")

    if actual_idx != expected_idx:
        print(f"error: chain position {expected_idx} has stage id {actual_idx} (expected {expected_idx})", file=sys.stderr)
        sys.exit(1)

    expected = plan.get(expected_idx)
    if expected is None:
        print(f"error: regenerated plan has no stage {expected_idx}", file=sys.stderr)
        sys.exit(1)

    if actual_file != expected["filename"] or actual_sha != expected["state_sha256"]:
        print(f"error: chain position {expected_idx} does not match regenerated plan:", file=sys.stderr)
        print(f"       commit:   file={actual_file} state-sha256={actual_sha}", file=sys.stderr)
        print(f"       expected: file={expected['filename']} state-sha256={expected['state_sha256']}", file=sys.stderr)
        print("       This usually means the essay content changed since this reconstruction", file=sys.stderr)
        print("       branch was started. Refusing to resume.", file=sys.stderr)
        sys.exit(1)

print(f"[info] Verified {completed} completed stages against the regenerated plan.")
PY
fi

###############################################################################
# Recover a partially-written file from a mid-stage crash, then do the final
# cleanliness check for anything unexpected.
###############################################################################

reset_path_to_head() {
    local path="$1"
    if git cat-file -e "HEAD:$path" 2>/dev/null; then
        git checkout -- "$path" 2>/dev/null || true
        git reset -q HEAD -- "$path" 2>/dev/null || true
    else
        rm -f -- "$path"
    fi
}

if [[ "$RESUME" == "1" && "$COMPLETED" -lt "$TOTAL_STAGES" ]]; then
    NEXT_LINE="$(sed -n "$((COMPLETED + 1))p" "$PLAN")"
    NEXT_FILE="$(printf '%s' "$NEXT_LINE" | cut -f5)"

    if [[ -n "$(git status --porcelain -- "$NEXT_FILE" 2>/dev/null)" ]]; then
        echo "[info] Recovering partially-written file from interrupted stage: $NEXT_FILE"
        reset_path_to_head "$NEXT_FILE"
    fi
fi

DIRTY_TRACKED="$(git status --porcelain --untracked-files=no)"
if [[ -n "$DIRTY_TRACKED" ]]; then
    echo "error: tracked working tree has unexpected uncommitted changes:" >&2
    echo "$DIRTY_TRACKED" >&2
    echo >&2
    echo "Commit, stash, or restore these before continuing." >&2
    exit 1
fi

###############################################################################
# Dry-run
###############################################################################

if (( DRY_RUN )); then
    echo "Already completed: $COMPLETED / $TOTAL_STAGES"
    echo
    while IFS=$'\t' read -r index category unit title filename start end bytes digest statefile final; do
        if (( index <= COMPLETED )); then continue; fi
        printf '%s-%04d %-14s :: %s\n' "$PREFIX" "$index" "$category" "$title"
        printf '         file=%s lines=%s..%s bytes=%s\n' "$filename" "$start" "$end" "$bytes"
    done < "$PLAN"
    echo
    echo "Dry run complete. Repository was not modified."
    exit 0
fi

if [[ "$COMPLETED" -ge "$TOTAL_STAGES" ]]; then
    echo "Nothing to do. Reconstruction is already complete ($COMPLETED / $TOTAL_STAGES)."
    exit 0
fi

###############################################################################
# Commit stages
###############################################################################

CURRENT=0

while IFS=$'\t' read -r index category unit title filename start end bytes digest statefile final; do
    CURRENT="$index"
    if (( CURRENT <= COMPLETED )); then continue; fi

    STATE_PATH="$STATES/$statefile"
    mkdir -p "$(dirname "$filename")"

    if [[ "$category" == "DELETE" ]]; then
        rm -f -- "$filename"
    else
        cp "$STATE_PATH" "$filename"
    fi

    git add -A -- "$filename"

    if git diff --cached --quiet -- "$filename"; then
        echo
        echo "error: stage $CURRENT produced no Git change: $filename" >&2
        exit 1
    fi

    printf -v ID '%s-%04d' "$PREFIX" "$CURRENT"
    SUBJECT="$ID $category :: $title"
    SUBJECT="$(python3 - "$SUBJECT" <<'PY'
import sys
s = " ".join(sys.argv[1].split())
limit = 118
if len(s) > limit:
    s = s[:limit - 3].rstrip() + "..."
print(s)
PY
)"

    STAT="$(git diff --cached --shortstat -- "$filename" | sed 's/^[[:space:]]*//')"

    MESSAGE="$TMPDIR/message.txt"
    cat > "$MESSAGE" <<EOF
$SUBJECT

file: $filename
unit: $unit
source-span: $start..$end
state-bytes: $bytes

change: $STAT

classifier: $CLASSIFIER_VERSION
state-sha256: $digest
patch-sha256: $SOURCE_SHA
base-commit: $BASE_COMMIT
EOF

    git commit --quiet -F "$MESSAGE" -- "$filename"

    printf '[%4d/%4d] %-14s %s\n' "$CURRENT" "$TOTAL_STAGES" "$category" "$title"

    if [[ "$DELAY" != "0" ]]; then
        sleep "$DELAY"
    fi
done < "$PLAN"

###############################################################################
# Final verification
###############################################################################

echo
echo "Verifying reconstructed files..."

VERIFY_FAILURE=0
while IFS=$'\t' read -r index category unit title filename start end bytes digest statefile final; do
    [[ "$final" == "1" ]] || continue
    REFERENCE_FILE="$REFERENCE_WORKTREE/$filename"
    if [[ -e "$REFERENCE_FILE" ]]; then
        if [[ ! -e "$filename" ]]; then
            echo "MISSING: $filename"; VERIFY_FAILURE=1; continue
        fi
        if ! cmp -s "$REFERENCE_FILE" "$filename"; then
            echo "DIFFERS: $filename"; VERIFY_FAILURE=1
        fi
    else
        if [[ -e "$filename" ]]; then
            echo "SHOULD BE DELETED: $filename"; VERIFY_FAILURE=1
        fi
    fi
done < "$PLAN"

if (( VERIFY_FAILURE )); then
    echo
    echo "ERROR: reconstructed tree does not match patched reference." >&2
    exit 1
fi

echo "OK: reconstructed changed files match the normally applied patch."
echo
echo "Complete."
echo
echo "  base commit:   $BASE_COMMIT"
echo "  stages:        $TOTAL_STAGES"
echo "  patch SHA-256: $SOURCE_SHA"
echo
echo "Recent reconstruction history:"
echo

git log --oneline --decorate --max-count=25

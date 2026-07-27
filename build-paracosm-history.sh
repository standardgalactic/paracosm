#!/usr/bin/env bash
#
# build-paracosm-history.sh
#
# Reconstruct a unified Git patch in semantically classified chunks,
# committing each chunk separately.
#
# Primary boundaries:
#   1. diff --git                         -> new file context
#   2. @@ ... @@                          -> fallback patch-hunk boundary
#   3. +\part{...}
#      +\chapter{...}
#      +\section{...}
#      +\subsection{...}
#      +\subsubsection{...}               -> semantic LaTeX boundary
#
# The generated patch is always an exact prefix of the source patch.
# Git itself acts as the resume cursor: the version of TARGET stored in
# HEAD determines which chunks have already been committed.
#
# Nothing is pushed. All commits are local.
#
# Examples:
#
#   ./build-paracosm-history.sh
#
#   ./build-paracosm-history.sh \
#       --source ../paracosm-fixes.patch \
#       --target paracosm-history.patch
#
#   DELAY=1 ./build-paracosm-history.sh
#
#   ./build-paracosm-history.sh --dry-run
#
#   ./build-paracosm-history.sh -f
#
# -f / --force:
#   Pass -f to git add, useful if TARGET is ignored by .gitignore.
#

set -Eeuo pipefail

###############################################################################
# Defaults
###############################################################################

SOURCE="../paracosm-fixes.patch"
TARGET="paracosm-history.patch"
PREFIX="PARA"
DELAY="${DELAY:-0}"
FORCE_ADD=0
DRY_RUN=0

CLASSIFIER_VERSION="paracosm-section-hunk-v2"

###############################################################################
# Usage
###############################################################################

usage() {
    cat <<'EOF'
Usage:
  ./build-paracosm-history.sh [options]

Options:
  --source FILE       Source unified patch.
                      Default: paracosm-fixes.patch

  --target FILE       Patch reconstructed and committed chunk by chunk.
                      Default: workspace/paracosm-history.patch

  --prefix NAME       Commit ID prefix.
                      Default: PARA

  --delay SECONDS     Pause between commits.
                      Default: $DELAY or 0

  --dry-run           Classify and display all chunks without modifying Git.

  -f, --force         Use "git add -f" for the generated target.

  -h, --help          Show this help.

Examples:

  ./build-paracosm-history.sh

  ./build-paracosm-history.sh \
      --source ../paracosm-fixes.patch \
      --target paracosm-history.patch

  DELAY=2 ./build-paracosm-history.sh

  ./build-paracosm-history.sh --dry-run

Important:
  SOURCE is never modified.
  TARGET is considered a generated artifact.
  The script never runs git push.
EOF
}

###############################################################################
# Arguments
###############################################################################

while [[ $# -gt 0 ]]; do
    case "$1" in
        --source)
            SOURCE="${2:?Missing value for --source}"
            shift 2
            ;;
        --target)
            TARGET="${2:?Missing value for --target}"
            shift 2
            ;;
        --prefix)
            PREFIX="${2:?Missing value for --prefix}"
            shift 2
            ;;
        --delay)
            DELAY="${2:?Missing value for --delay}"
            shift 2
            ;;
        --dry-run)
            DRY_RUN=1
            shift
            ;;
        -f|--force)
            FORCE_ADD=1
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        *)
            echo "Unknown argument: $1" >&2
            echo >&2
            usage >&2
            exit 2
            ;;
    esac
done

###############################################################################
# Basic validation
###############################################################################

command -v git >/dev/null 2>&1 || {
    echo "error: git is required" >&2
    exit 1
}

command -v python3 >/dev/null 2>&1 || {
    echo "error: python3 is required" >&2
    exit 1
}

[[ -f "$SOURCE" ]] || {
    echo "error: source patch not found: $SOURCE" >&2
    exit 1
}

if [[ "$SOURCE" == "$TARGET" ]]; then
    echo "error: SOURCE and TARGET must be different files" >&2
    exit 1
fi

if ! [[ "$DELAY" =~ ^[0-9]+([.][0-9]+)?$ ]]; then
    echo "error: --delay must be a non-negative number" >&2
    exit 1
fi

###############################################################################
# Locate repository
###############################################################################

REPO_ROOT="$(git rev-parse --show-toplevel 2>/dev/null)" || {
    echo "error: not inside a Git repository" >&2
    exit 1
}

cd "$REPO_ROOT"

# Turn source into an absolute path because we have changed directory.
SOURCE="$(python3 - "$SOURCE" <<'PY'
import os
import sys
print(os.path.abspath(sys.argv[1]))
PY
)"

# Git pathspecs should use repository-relative paths.
if [[ "$TARGET" = /* ]]; then
    TARGET="$(python3 - "$REPO_ROOT" "$TARGET" <<'PY'
import os
import sys

root = os.path.abspath(sys.argv[1])
target = os.path.abspath(sys.argv[2])

try:
    rel = os.path.relpath(target, root)
except ValueError:
    raise SystemExit("target is not inside the repository")

if rel == ".." or rel.startswith("../"):
    raise SystemExit("target must be inside the Git repository")

print(rel)
PY
)"
fi

mkdir -p "$(dirname "$TARGET")"

###############################################################################
# Temporary workspace
###############################################################################

TMPDIR="$(mktemp -d "${TMPDIR:-/tmp}/paracosm-history.XXXXXXXX")"

cleanup() {
    rm -rf "$TMPDIR"
}

trap cleanup EXIT

CHUNK_DIR="$TMPDIR/chunks"
MANIFEST="$TMPDIR/manifest.tsv"

mkdir -p "$CHUNK_DIR"

###############################################################################
# Split and classify patch
###############################################################################

python3 - "$SOURCE" "$CHUNK_DIR" "$MANIFEST" "$PREFIX" <<'PY'
from __future__ import annotations

import hashlib
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path


source_path = Path(sys.argv[1])
chunk_dir = Path(sys.argv[2])
manifest_path = Path(sys.argv[3])
prefix = sys.argv[4]


###############################################################################
# Helpers
###############################################################################

def text(line: bytes) -> str:
    return line.decode("utf-8", errors="replace").rstrip("\r\n")


def clean_title(value: str) -> str:
    value = value.strip()

    # A few cheap LaTeX-to-log normalizations.
    value = value.replace(r"\textit{", "")
    value = value.replace(r"\textbf{", "")
    value = value.replace(r"\emph{", "")
    value = value.replace(r"\mathbf{", "")
    value = value.replace(r"\mathrm{", "")

    value = value.replace("---", "—")
    value = value.replace("--", "–")

    # Remove obvious remaining braces without attempting a LaTeX parser.
    value = value.replace("{", "")
    value = value.replace("}", "")

    value = re.sub(r"\s+", " ", value).strip()

    if len(value) > 86:
        value = value[:83].rstrip() + "..."

    return value or "untitled"


def patch_filename(line: str) -> str:
    """
    diff --git a/foo.tex b/foo.tex
                       ^^^^^^^^^
    """
    m = re.match(r"^diff --git a/(.*?) b/(.*)$", line)

    if not m:
        return "unknown"

    return m.group(2)


SECTION_RE = re.compile(
    r"""
    ^\+
    \\(?P<kind>
        part
        |chapter
        |section
        |subsection
        |subsubsection
    )
    \*?
    \{
    (?P<title>.*)
    \}
    \s*$
    """,
    re.VERBOSE,
)


def latex_header(line: str):
    if line.startswith("+++"):
        return None

    m = SECTION_RE.match(line)

    if not m:
        return None

    return m.group("kind"), clean_title(m.group("title"))


###############################################################################
# Semantic classifier
###############################################################################

# These are intentionally somewhat interpretive.
#
# Title matches are weighted much more strongly than body matches so the
# classification remains explainable rather than becoming a bag-of-words
# accident.

CATEGORY_PATTERNS = {
    "OPENING": [
        r"\bpreface\b",
        r"\bprologue\b",
        r"\bintroduction\b",
        r"\breading note\b",
        r"\bopening\b",
    ],

    "ORIENTATION": [
        r"\boverview\b",
        r"\bconceptual map\b",
        r"\bdiagrammatic overview\b",
        r"\broadmap\b",
        r"\bguide\b",
        r"\borientation\b",
    ],

    "PROVENANCE": [
        r"\bprovenance\b",
        r"\bsource drafts?\b",
        r"\beditorial note\b",
        r"\bsource note\b",
        r"\btextual history\b",
    ],

    "PROOF-OBJECT": [
        r"\bproof\b",
        r"\btheorem\b",
        r"\blemma\b",
        r"\bcorollary\b",
        r"\bproposition\b",
        r"\bformal verification\b",
        r"\btriangle identities\b",
    ],

    "DEFINITION": [
        r"\bdefinition\b",
        r"\bdefining\b",
        r"\baxiom\b",
        r"\bobjects?\b",
        r"\bmorphisms?\b",
        r"\bconstruction\b",
        r"\bformalism\b",
    ],

    "DYNAMICS": [
        r"\bcurvature\b",
        r"\bflow\b",
        r"\bdynamics?\b",
        r"\bgeodesic",
        r"\bentropy\b",
        r"\bdiffusion\b",
        r"\bevolution\b",
        r"\btime\b",
        r"\boscillat",
        r"\bthermodynamic",
        r"\bgradient\b",
    ],

    "BRIDGE": [
        r"\btoward\b",
        r"\btransition\b",
        r"\bcomparison\b",
        r"\bcomparative\b",
        r"\bfrom .* to\b",
        r"\bconnection\b",
        r"\bcorrespondence\b",
        r"\btranslation\b",
    ],

    "SYNTHESIS": [
        r"\bsynthesis\b",
        r"\bintegration\b",
        r"\bclosure\b",
        r"\bcomposition\b",
        r"\bcoherence\b",
        r"\bcolimit",
        r"\blimit",
        r"\bgluing\b",
        r"\buniversal",
        r"\bconsensus\b",
    ],

    "REFLECTION": [
        r"\breflective\b",
        r"\bepilogue\b",
        r"\bethical\b",
        r"\bethics\b",
        r"\bcare\b",
        r"\bcompassion\b",
        r"\bmoral\b",
        r"\bmeaning\b",
    ],

    "APPARATUS": [
        r"\bbibliograph",
        r"\bnotation\b",
        r"\bdiagram library\b",
        r"\btemplates?\b",
        r"\bcomputational\b",
        r"\bimplementation\b",
        r"\breference\b",
    ],
}


CATEGORY_PRIORITY = [
    "PROVENANCE",
    "PROOF-OBJECT",
    "OPENING",
    "ORIENTATION",
    "APPARATUS",
    "REFLECTION",
    "DEFINITION",
    "DYNAMICS",
    "BRIDGE",
    "SYNTHESIS",
    "ARGUMENT",
]


def semantic_category(title: str, body: str) -> str:
    title_low = title.lower()
    body_low = body.lower()

    scores = {category: 0 for category in CATEGORY_PATTERNS}

    for category, patterns in CATEGORY_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, title_low, flags=re.IGNORECASE):
                scores[category] += 8

            # Body evidence matters, but much less.
            hits = len(
                re.findall(
                    pattern,
                    body_low,
                    flags=re.IGNORECASE,
                )
            )

            scores[category] += min(hits, 5)

    best_score = max(scores.values(), default=0)

    if best_score == 0:
        return "ARGUMENT"

    tied = {
        category
        for category, score in scores.items()
        if score == best_score
    }

    for category in CATEGORY_PRIORITY:
        if category in tied:
            return category

    return "ARGUMENT"


###############################################################################
# Chunk object
###############################################################################

@dataclass
class Chunk:
    lines: list[bytes]
    start_line: int
    end_line: int
    filename: str
    unit_type: str
    title: str
    category: str = ""


chunks: list[Chunk] = []

current_lines: list[bytes] = []
current_start = 1
current_file = "patch"
current_unit = "PATCH"
current_title = "patch preamble"


def flush(end_line: int):
    global current_lines
    global current_start
    global current_unit
    global current_title

    if not current_lines:
        return

    chunks.append(
        Chunk(
            lines=current_lines,
            start_line=current_start,
            end_line=end_line,
            filename=current_file,
            unit_type=current_unit,
            title=current_title,
        )
    )

    current_lines = []


###############################################################################
# Boundary scan
###############################################################################

raw_lines = source_path.read_bytes().splitlines(keepends=True)

for lineno, raw in enumerate(raw_lines, start=1):
    s = text(raw)

    # File boundary.
    if s.startswith("diff --git "):
        flush(lineno - 1)

        current_file = patch_filename(s)
        current_start = lineno
        current_unit = "FILE"
        current_title = current_file
        current_lines = [raw]
        continue

    # Hunk boundary.
    if s.startswith("@@ "):
        flush(lineno - 1)

        current_start = lineno
        current_unit = "HUNK"

        hunk_label = s
        if len(hunk_label) > 80:
            hunk_label = hunk_label[:77] + "..."

        current_title = hunk_label
        current_lines = [raw]
        continue

    # Semantic LaTeX boundary.
    header = latex_header(s)

    if header is not None:
        kind, title_value = header

        flush(lineno - 1)

        current_start = lineno
        current_unit = kind.upper()
        current_title = title_value
        current_lines = [raw]
        continue

    # Ordinary line.
    if not current_lines:
        current_start = lineno
        current_unit = "PATCH"
        current_title = current_file

    current_lines.append(raw)


flush(len(raw_lines))


###############################################################################
# Categorize chunks
###############################################################################

for chunk in chunks:
    decoded = "".join(text(line) + "\n" for line in chunk.lines)

    if chunk.unit_type == "FILE":
        chunk.category = "FILE-OPEN"

    elif chunk.unit_type == "HUNK":
        ext = Path(chunk.filename).suffix.lower()
        base = Path(chunk.filename).name.lower()

        if base == "version":
            chunk.category = "RELEASE"

        elif ext in {".md", ".markdown", ".rst"}:
            chunk.category = "EDITORIAL"

        elif ext == ".tex":
            # A hunk with no LaTeX section boundary is treated as lower-level
            # patchwork rather than pretending it forms an essay section.
            chunk.category = "PATCHWORK"

        else:
            chunk.category = "PATCH-HUNK"

    elif chunk.unit_type in {
        "PART",
        "CHAPTER",
        "SECTION",
        "SUBSECTION",
        "SUBSUBSECTION",
    }:
        chunk.category = semantic_category(chunk.title, decoded)

    else:
        chunk.category = "PATCH"


###############################################################################
# Metrics
###############################################################################

def count_metrics(chunk: Chunk):
    additions = 0
    deletions = 0
    context = 0
    blank_additions = 0
    citations = 0
    labels = 0
    refs = 0
    equations = 0
    math_signals = 0

    math_tokens = [
        "$",
        r"\[",
        r"\]",
        r"\begin{equation",
        r"\begin{align",
        r"\mathrm",
        r"\mathbf",
        r"\mathcal",
        r"\Phi",
        r"\SEntropy",
        r"\vField",
        r"\grad",
        r"\lim",
        r"\mathrm{Hom}",
    ]

    for raw in chunk.lines:
        line = text(raw)

        if line.startswith("+++") or line.startswith("---"):
            continue

        if line.startswith("+"):
            additions += 1

            payload = line[1:]

            if not payload.strip():
                blank_additions += 1

            citations += len(
                re.findall(
                    r"\\(?:cite|citep|citet|autocite|parencite)\b",
                    payload,
                )
            )

            labels += len(re.findall(r"\\label\{", payload))
            refs += len(
                re.findall(
                    r"\\(?:ref|eqref|autoref|cref|Cref)\{",
                    payload,
                )
            )

            if re.search(
                r"\\begin\{(?:equation|align|gather|multline)",
                payload,
            ):
                equations += 1

            for token in math_tokens:
                math_signals += payload.count(token)

        elif line.startswith("-"):
            deletions += 1

        elif not (
            line.startswith("diff --git ")
            or line.startswith("index ")
            or line.startswith("@@ ")
            or line.startswith("new file mode ")
            or line.startswith("deleted file mode ")
            or line.startswith("similarity index ")
            or line.startswith("rename from ")
            or line.startswith("rename to ")
        ):
            context += 1

    return {
        "additions": additions,
        "deletions": deletions,
        "context": context,
        "blank_additions": blank_additions,
        "citations": citations,
        "labels": labels,
        "refs": refs,
        "equations": equations,
        "math_signals": math_signals,
    }


###############################################################################
# Write chunks and manifest
###############################################################################

manifest_rows = []

for i, chunk in enumerate(chunks, start=1):
    chunk_name = f"{i:05d}.patchchunk"
    chunk_path = chunk_dir / chunk_name

    chunk_bytes = b"".join(chunk.lines)
    chunk_path.write_bytes(chunk_bytes)

    metrics = count_metrics(chunk)

    chunk_hash = hashlib.sha256(chunk_bytes).hexdigest()

    title = chunk.title.replace("\t", " ").replace("\n", " ")
    filename = chunk.filename.replace("\t", " ").replace("\n", " ")

    row = [
        str(i),
        chunk.category,
        chunk.unit_type,
        title,
        filename,
        str(chunk.start_line),
        str(chunk.end_line),
        str(len(chunk.lines)),
        str(metrics["additions"]),
        str(metrics["deletions"]),
        str(metrics["context"]),
        str(metrics["blank_additions"]),
        str(metrics["citations"]),
        str(metrics["labels"]),
        str(metrics["refs"]),
        str(metrics["equations"]),
        str(metrics["math_signals"]),
        chunk_hash,
        chunk_name,
    ]

    manifest_rows.append("\t".join(row))


manifest_path.write_text(
    "\n".join(manifest_rows) + "\n",
    encoding="utf-8",
)


###############################################################################
# Summary
###############################################################################

source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()

print(f"source-lines\t{len(raw_lines)}")
print(f"chunks\t{len(chunks)}")
print(f"source-sha256\t{source_hash}")
PY

###############################################################################
# Read source summary
###############################################################################

SOURCE_LINES="$(wc -l < "$SOURCE" | tr -d ' ')"
SOURCE_SHA="$(sha256sum "$SOURCE" | awk '{print $1}')"
TOTAL_CHUNKS="$(wc -l < "$MANIFEST" | tr -d ' ')"

echo
echo "Source:      $SOURCE"
echo "Target:      $TARGET"
echo "Patch lines: $SOURCE_LINES"
echo "Chunks:      $TOTAL_CHUNKS"
echo "SHA-256:     $SOURCE_SHA"
echo

###############################################################################
# Helper: print planned history
###############################################################################

print_plan() {
    while IFS=$'\t' read -r \
        index \
        category \
        unit \
        title \
        file \
        start \
        end \
        lines \
        additions \
        deletions \
        context \
        blanks \
        citations \
        labels \
        refs \
        equations \
        mathsignals \
        chunksha \
        chunkfile
    do
        printf '%s-%04d %-14s :: %s\n' \
            "$PREFIX" \
            "$index" \
            "$category" \
            "$title"

        printf '         file=%s  patch=%s..%s  lines=%s  +%s/-%s\n' \
            "$file" \
            "$start" \
            "$end" \
            "$lines" \
            "$additions" \
            "$deletions"

    done < "$MANIFEST"
}

if (( DRY_RUN )); then
    print_plan
    echo
    echo "Dry run complete. Git was not modified."
    exit 0
fi

###############################################################################
# Determine progress from Git itself
###############################################################################

HEAD_COPY="$TMPDIR/head-target.patch"

if git cat-file -e "HEAD:$TARGET" 2>/dev/null; then
    git show "HEAD:$TARGET" > "$HEAD_COPY"
else
    : > "$HEAD_COPY"
fi

COMPLETED="$(
    python3 - "$HEAD_COPY" "$CHUNK_DIR" "$TOTAL_CHUNKS" <<'PY'
from pathlib import Path
import sys

head_path = Path(sys.argv[1])
chunk_dir = Path(sys.argv[2])
total = int(sys.argv[3])

head = head_path.read_bytes()

if not head:
    print(0)
    raise SystemExit(0)

acc = bytearray()

for i in range(1, total + 1):
    chunk = chunk_dir / f"{i:05d}.patchchunk"
    acc.extend(chunk.read_bytes())

    if bytes(acc) == head:
        print(i)
        raise SystemExit(0)

    if len(acc) > len(head):
        break

print(-1)
PY
)"

if [[ "$COMPLETED" == "-1" ]]; then
    echo "error:" >&2
    echo "The version of $TARGET stored in HEAD is not an exact chunk" >&2
    echo "prefix of the current source patch." >&2
    echo >&2
    echo "This usually means either:" >&2
    echo "  * the source patch changed after this history was started, or" >&2
    echo "  * TARGET was edited manually." >&2
    echo >&2
    echo "No files were changed." >&2
    exit 1
fi

###############################################################################
# Crash recovery
#
# HEAD is authoritative.
#
# If the previous run died after appending a chunk but before committing it,
# throw away the generated working-tree residue and restore the exact version
# represented by HEAD.
###############################################################################

if git cat-file -e "HEAD:$TARGET" 2>/dev/null; then
    if ! cmp -s "$HEAD_COPY" "$TARGET" 2>/dev/null; then
        echo "Recovering generated target from HEAD..."
        cp "$HEAD_COPY" "$TARGET"

        # Remove a stale staged version of TARGET without disturbing other
        # staged paths.
        git reset -q HEAD -- "$TARGET" 2>/dev/null || true
    fi
else
    # No committed target yet.
    if [[ -e "$TARGET" ]]; then
        if [[ -s "$TARGET" ]]; then
            echo "Removing uncommitted generated target from interrupted run..."
        fi
        rm -f "$TARGET"
    fi
fi

mkdir -p "$(dirname "$TARGET")"
touch "$TARGET"

echo "Already committed: $COMPLETED / $TOTAL_CHUNKS"
echo

if (( COMPLETED >= TOTAL_CHUNKS )); then
    echo "Nothing to do. The patch history is already complete."
    exit 0
fi

###############################################################################
# Commit chunks
###############################################################################

CURRENT=0

while IFS=$'\t' read -r \
    index \
    category \
    unit \
    title \
    file \
    start \
    end \
    lines \
    additions \
    deletions \
    context \
    blanks \
    citations \
    labels \
    refs \
    equations \
    mathsignals \
    chunksha \
    chunkfile
do
    CURRENT="$index"

    if (( CURRENT <= COMPLETED )); then
        continue
    fi

    CHUNK_PATH="$CHUNK_DIR/$chunkfile"

    ###########################################################################
    # Append exact source bytes
    ###########################################################################

    cat "$CHUNK_PATH" >> "$TARGET"

    ###########################################################################
    # Commit subject
    ###########################################################################

    printf -v ID '%s-%04d' "$PREFIX" "$CURRENT"

    SUBJECT="$ID $category :: $title"

    # Keep subjects readable in git log --oneline.
    SUBJECT="$(
        python3 - "$SUBJECT" <<'PY'
import sys

s = " ".join(sys.argv[1].split())

limit = 118

if len(s) > limit:
    s = s[:limit - 3].rstrip() + "..."

print(s)
PY
    )"

    ###########################################################################
    # Commit body
    ###########################################################################

    MESSAGE_FILE="$TMPDIR/commit-message.txt"

    cat > "$MESSAGE_FILE" <<EOF
$SUBJECT

file: $file
unit: $unit
patch-span: $start..$end
physical-lines: $lines

additions: $additions
deletions: $deletions
context-lines: $context
blank-additions: $blanks

citations: $citations
labels: $labels
references: $refs
equation-environments: $equations
math-signals: $mathsignals

classifier: $CLASSIFIER_VERSION
chunk-sha256: $chunksha
source-sha256: $SOURCE_SHA
EOF

    ###########################################################################
    # Stage generated patch
    ###########################################################################

    if (( FORCE_ADD )); then
        git add -f -- "$TARGET"
    else
        if ! git add -- "$TARGET"; then
            echo >&2
            echo "git add failed." >&2
            echo >&2
            echo "If TARGET is intentionally ignored, rerun with:" >&2
            echo >&2
            echo "    ./build-paracosm-history.sh -f" >&2
            echo >&2
            exit 1
        fi
    fi

    ###########################################################################
    # Commit only this generated path
    ###########################################################################

    git commit \
        --quiet \
        -F "$MESSAGE_FILE" \
        -- "$TARGET"

    printf '[%4d/%4d] %-14s %s\n' \
        "$CURRENT" \
        "$TOTAL_CHUNKS" \
        "$category" \
        "$title"

    ###########################################################################
    # Optional pacing
    ###########################################################################

    if [[ "$DELAY" != "0" ]]; then
        sleep "$DELAY"
    fi

done < "$MANIFEST"

###############################################################################
# Verification
###############################################################################

echo
echo "Verifying reconstructed patch..."

if cmp -s "$SOURCE" "$TARGET"; then
    echo "OK: generated patch is byte-for-byte identical to source."
else
    echo "ERROR: generated patch differs from source." >&2
    exit 1
fi

FINAL_SHA="$(sha256sum "$TARGET" | awk '{print $1}')"

echo
echo "Complete."
echo
echo "  source:       $SOURCE"
echo "  target:       $TARGET"
echo "  chunks:       $TOTAL_CHUNKS"
echo "  commits made: $((TOTAL_CHUNKS - COMPLETED))"
echo "  SHA-256:      $FINAL_SHA"
echo
echo "Recent history:"
echo

git log \
    --oneline \
    --decorate \
    --max-count=20 \
    -- "$TARGET"

#!/usr/bin/env bash
set -euo pipefail

# Symlink each skill directory containing SKILL.md under skills/ into the target
# local skills directory without overwriting existing unrelated entries.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
SOURCE_ROOT="$REPO_ROOT/skills"
TARGET_ROOT="${1:-${SKILLS_TARGET_DIR:-$HOME/.gemini/antigravity-cli/skills}}"

mkdir -p "$TARGET_ROOT"

found=0
for skill_dir in "$SOURCE_ROOT"/*/; do
  [ -d "$skill_dir" ] || continue
  skill_dir="${skill_dir%/}"

  if [ ! -f "$skill_dir/SKILL.md" ]; then
    continue
  fi

  found=1
  skill_name="$(basename "$skill_dir")"
  target_path="$TARGET_ROOT/$skill_name"

  if [ -L "$target_path" ]; then
    current_target="$(readlink "$target_path")"
    if [ "$current_target" = "$skill_dir" ]; then
      echo "ok: $skill_name (already linked)"
      continue
    fi
    echo "skip: $skill_name (symlink points elsewhere: $current_target)" >&2
    continue
  fi

  if [ -e "$target_path" ]; then
    echo "skip: $skill_name (existing non-symlink at $target_path)" >&2
    continue
  fi

  ln -s "$skill_dir" "$target_path"
  echo "linked: $target_path -> $skill_dir"
done

if [ "$found" -eq 0 ]; then
  echo "No skills found under $SOURCE_ROOT"
fi

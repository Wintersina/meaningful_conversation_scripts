#!/usr/bin/env bash
# Stop hook: keep this project's repo, origin, and live Apps Script in sync.
#
# Behavior: "push committed work only".
#   - No-op if the working tree is dirty  (never auto-commits WIP).
#   - If the tree is clean AND local is ahead of origin: `git push`, then
#     `clasp push -f` when the unpushed commits actually touched scripts/.
#   - Always exits 0 so a sync hiccup never blocks the turn.
#
# NOTE: this only fires during Claude Code sessions — it cannot catch edits
# made directly in the browser Apps Script editor.
set -uo pipefail

# Consume the hook's stdin JSON (unused).
cat >/dev/null 2>&1 || true

emit() { jq -cn --arg m "$1" '{systemMessage:$m, suppressOutput:true}'; }

repo="$(git rev-parse --show-toplevel 2>/dev/null)" || exit 0
cd "$repo" || exit 0

# 1) Never auto-commit — bail if there are uncommitted changes.
[ -z "$(git status --porcelain)" ] || exit 0

# 2) Require an upstream and only act when ahead of it.
upstream="$(git rev-parse --abbrev-ref --symbolic-full-name '@{u}' 2>/dev/null)" || exit 0
ahead="$(git rev-list --count "${upstream}..HEAD" 2>/dev/null || echo 0)"
[ "${ahead:-0}" -gt 0 ] || exit 0

# Capture which scripts/ files changed in the unpushed commits (before ref moves).
old_up="$(git rev-parse "$upstream" 2>/dev/null)"
scripts_changed="$(git diff --name-only "${old_up}..HEAD" -- scripts/ 2>/dev/null)"

# 3) Push to origin.
if ! git push --quiet 2>/tmp/mc_sync.err; then
  emit "Auto-sync: git push failed — run it manually. See /tmp/mc_sync.err"
  exit 0
fi
summary="Auto-sync: pushed ${ahead} commit(s) to ${upstream}"

# 4) clasp push only when scripts/ changed and clasp is set up.
if [ -n "$scripts_changed" ]; then
  if command -v clasp >/dev/null 2>&1 && [ -f .clasp.json ]; then
    if clasp push -f >/tmp/mc_clasp.out 2>&1; then
      summary="${summary}; clasp push → live OK"
    else
      summary="${summary}; clasp push FAILED (run 'clasp push -f'). See /tmp/mc_clasp.out"
    fi
  else
    summary="${summary}; clasp unavailable — skipped live push"
  fi
fi

emit "$summary"
exit 0

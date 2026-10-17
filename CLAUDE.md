# Excel MCP Server

Instructions for Claude Code working in this repository.

## Target platform — Windows only

This server runs on Windows and only on Windows: `windows_com.py` drives Excel
through COM, which does not exist off Windows. Linux and macOS are not
supported, not tested and not targeted.

- **Do not spend analysis, implementation, review or verification effort on
  Linux/macOS-only behavior** — POSIX paths, case-sensitive filesystems,
  permission bits and `chmod`, symlinks, signals, `sh`/`bash` quoting, POSIX
  locale and encoding defaults, LF-vs-CRLF differences that only appear off
  Windows.
- **A defect that can only occur off Windows is out of scope.** If you notice
  one, record it in a single line and move on: never widen scope to fix it,
  never let it block a review or a merge. Portability is not a requirement, so
  it never produces a review finding on its own.
- **Windows behavior itself is fully in scope.** COM object lifetime, a
  lingering `EXCEL.EXE` after a failed call, modal dialogs blocking automation,
  console encoding (`cp1252` vs UTF-8), file locking on an open workbook,
  drive-letter paths and path-length limits are Windows concerns, not
  portability concerns — they are where this server actually breaks.
- **Existing cross-platform code stays as it is.** This rule stops new effort;
  it is not a licence to strip working non-Windows branches out of code that
  already has them. Removing them is a behavior change that buys nothing here.

A requirement that explicitly asks for non-Windows support is a change to this
constraint: treat it as a requirement change and confirm with the user before
doing the work.

## Scope of this file

This file records the platform constraint, the repository layout and the *Pull request
lifecycle* below, nothing else. It deliberately does not define a routing policy or agent
roster — this repository has not adopted one, and inventing rules here that nobody agreed to
would be worse than the silence it replaces. The owner asked on 2026-10-12 that every
repository follow the same pull request loop, so that one section is the review process here,
and the only one. Follow the conventions already in the code and in `README.md`.

## Repository layout

The root holds only what a user needs to run the server: `server.py`, README,
`requirements.txt` and this file. The Excel backends live in `backends/`, the
demo in `demo/`, developer tooling in `scripts/`. When a change adds a Python
file, test or module, use `.claude/skills/repo-layout/SKILL.md`.

`python scripts/layout_check.py` is the verdict: a test file in the root, a
root `.py` with neither an `if __name__ == "__main__":` guard nor the
`# layout: entry-point` marker in its first lines, or more than three root
entry points fails it. Run it rather than judging the tree by eye.

## Pull request lifecycle

Every pull request you create or drive for a requirement goes through the loop below, whatever the size of the change. It is a standing rule, not a judgement call. It starts when a PR exists; it does not authorize opening one, so a PR is still opened only when the user asks for it or the task's own flow calls for it.

1. **Open, then Codex review — immediately.** The first action after the PR exists is the Codex review: a PR comment whose whole body is `@codex review`, and `@codex security review` as well when the change touches authentication, authorization, secrets, payments or personal data. Codex does not review a PR on its own unless the repository turns that on, so ask for it. A PR with no Codex review is not done and is never merged.
2. **Validate every finding before acting.** Check it against the code. A confirmed finding is fixed; a wrong one is answered on its thread with the evidence — never silently skipped, never "fixed" speculatively.
3. **Fix, then have the new head reviewed.** Fix each confirmed finding — with a regression test where the defect has testable behavior; a prose-only fix keeps the repository's normal checks instead — run the repository's own checks, push, and request the next review on the new head: each review command once per head commit, one request at a time, never before the previous response was validated.
4. **Stop when the review is clean.** Clean means every review command the change requires (the standard review, and the security review where it applies) has a review of the *current* head with no finding (other than a P2/P3 whose recorded rejection stands, see the first bullet below), every earlier finding was fixed or rejected with evidence, CI and the required checks are green, and there is no merge conflict. A fix pushed after the last review is unreviewed, so it is not clean — review it with every command the change requires.
5. **Merge it yourself — where the repository allows it.** A clean PR is merged without asking for approval, scoped to the validated head commit, provided the base branch requires branches to be up to date and the merging identity cannot bypass that rule (read the branch protection to confirm both). Where either is missing, merging is a genuine blocker: leave the PR open, say it is clean and why it was not merged, and do not merge. Record the merge commit and a one-line summary on the PR.
6. **Then continue.** Do not stop after the merge to ask what is next: start the next outstanding task from the latest default branch (a merged branch is finished; never stack new commits on it). An outstanding task is one the user explicitly asked for that is not done yet; never invent work or pick up an item the user did not ask for. A task that needs the code of a PR that is not merged waits: skip it rather than stack on that branch. When none remains, report what merged and stop.

What bounds the loop:

- **No review cap.** The loop runs until every review command the change requires is clean on the current head. A review is still bounded by step 3's rules. Two exits remain: a finding you reject with evidence — a P2/P3 stands at once and blocks nothing, even if Codex raises it again; a P0/P1 is a disagreement unless a review drops it: when other fixes make a new head, push it and let the next review decide (raised again, unchanged or reworded or moved when it is substantially the same issue, means escalate; not raised again means closed), and when no code change is pending escalate at once — report it to the repository owner and do not merge until they decide — and a genuine blocker.
- **Codex unreachable.** If Codex gives no reaction or review after a reasonable wait, or answers with an error or a usage-limit notice, the PR is BLOCKED: report it once, leave the PR open and continue with the next task that does not depend on it. Retry at most once on an unchanged head, and never call Codex another way (API, Action, CLI, SDK, browser automation).
- **The trigger string.** No PR description, comment or commit message carries the literal trigger string except the intentional review requests of steps 1 and 3: the standard review, the security review and the single retry.
- **Who may stop the loop.** A human is needed only for a genuine blocker — the task cannot proceed without a decision only a human can make — or a change to the original requirement.

<!-- team-onboarding:start -->
## Team agent process

`vohoailinh90/Excel_MCP_Server` is registered with the free-model team: `team` (Claude leads) and `team-codex` (Codex leads), with free models writing the tests, the code and a pre-review. The details are in `docs/TEAM.md`; the settings that differ per repository live in `repos.conf` on the VPS, not here.

- Base branch `main`; sensitivity `public`; gates setup `python3 -m venv --system-site-packages .venv && (grep -viE '^[[:space:]]*(pywin32|pywinauto|comtypes|wmi|pythonnet|pyinstaller[-_.]+hooks[-_.]+contrib[-_.]+win|winshell)([^A-Za-z0-9._-]|$)' requirements.txt > .team-linux-requirements.txt || true) && .venv/bin/pip install -q -r .team-linux-requirements.txt`; suggested test commands `.venv/bin/python scripts/layout_check.py`.
- If you are a team worker (tester, coder or reviewer), follow the handoff files you are given and touch only the files the task names. Never edit CI (`.github/`), agent instructions (`CLAUDE.md`, `AGENTS.md`, `.claude/`, `.codex/`), lockfiles or `repos.conf` unless the task says so.

If you are Claude Code or Codex working on this repository directly (a spec, an audit or an ordinary change), this is the pull-request process:

- When your work here is ready to ship, open the pull request yourself; do not wait to be asked, unless the user asked for a local change only or this repository's own rules say a pull request is opened only on request.
- As soon as it is open, ask Codex for a review with one comment on the PR: `@codex review` (add `@codex security review` where this repository's rules, or the change itself, call for one: authentication, authorization, secrets, payments, personal data). Codex never reviews on its own here. Ask once per head commit for each review command, never while a request is unanswered, and never mention Codex in the PR description or a commit message. Where the repository has `scripts/codex_round_count.py`, run it before each request.
- Check each finding against the code. Confirmed: fix it, with a regression test where the defect has testable behavior (a prose-only fix keeps the repository's normal checks instead), push, and ask again on the new head. Wrong: answer on its thread with the evidence. There is no cap on rounds; a rejected P2/P3 stands and blocks nothing; a P0/P1 you cannot settle goes to the repository owner and the PR is not merged.
- Merge it yourself, without waiting for approval, once all of these hold on the current head: every review this change requires (the standard review, plus the security review where it touches authentication, authorization, secrets, payments or personal data) has been requested of that exact commit and has no finding left (a bare thumbs-up is not a review, unless this repository's own rules accept one for that exact request), CI is green, there is no merge conflict, and the base branch's protection (or a merge queue) requires branches to be up to date and applies to admins, and the identity that merges has no bypass of it (ruleset bypass list, custom role or app privilege). Immediately before merging, re-read the PR head and the base head: if the PR head moved, review it again; if the base moved, merge it in and re-validate. Merge scoped to the validated head SHA. If that protection is missing, do not merge: leave the PR open, say it is clean and why it was not merged, and tell the owner.
- If Codex gives no sign of life within 10 minutes, retry once per head; an explicit error or usage-limit reply needs no wait before that one retry. If it acknowledges but never delivers a review, or its task fails, do not retry: the PR is BLOCKED. The PR is also BLOCKED when the retry gets no answer. A BLOCKED PR stays open and is reported once; never merge it unreviewed, and never use another route (API, CLI, Action) to get a review.

Exceptions in this repository, which replace the matching rule above:

- demo/demo_test.py is the README's smoke test, but it fails on main today (ModuleNotFoundError: backends.mock_openpyxl), so it is not registered as a gate; repairing or dropping it is a change of its own.

Where this repository documents its own rule for one of these steps elsewhere (who merges, what gates a merge, how the review is requested, a cap on rounds), that rule governs that step; this process covers whatever the repository leaves unsaid.
<!-- team-onboarding:end -->

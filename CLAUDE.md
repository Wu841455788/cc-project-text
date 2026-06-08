# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

A standalone, zero-dependency digital clock web page ([digital-clock.html](digital-clock.html)). Open the file directly in a browser — no build step, no server required.

## How to run

```bash
open digital-clock.html
```

## Architecture

Everything lives in a single HTML file:

- **CSS (lines 7–141)** — Dark theme by default. A `body.light` class toggles light mode. Responsive font sizing via `clamp()`. Monospace typeface.
- **HTML (lines 143–151)** — Clock container with time, date, weekday, a 12/24H format badge, and a theme hint.
- **JavaScript (lines 153–222)** — `setInterval` drives a per-second `update()` that reads `new Date()` and writes to the DOM. Two global booleans — `is24Hour` and `isDark` — track toggle state.

## Interactions

| Action | Trigger |
|---|---|
| Toggle 12H / 24H | Click the format badge, or press `F` |
| Toggle dark / light theme | Press `T` |

## Code conventions

- Chinese locale for date/weekday labels (`zh-CN`).
- `pad(n)` helper zero-pads with `String.padStart(2, '0')`.
- Seconds are rendered as a `<span class="seconds">` inside the time element (smaller, dimmer), not as plain text.

# CLAUDE.md

Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

---

**These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.

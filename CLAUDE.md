# Pramong Taidin — CLAUDE.md

## Project Overview
Kahoot-style real-time quiz system for fishery/aquaculture students.
Single-page app (`index.html`) + Google Apps Script backend (`Code.gs`).

## Architecture

### Frontend — `index.html`
- Self-contained: HTML + CSS + JS in one file, no build step
- Three roles rendered in a single page via screen toggling:
  - **Landing** → Teacher Login → Teacher Mode (game control + question editor)
  - **Display** (projector screen) — polls API every 4 seconds, shows questions/leaderboard
  - **Student** — joins via PIN/QR, answers questions, sees live score
- Screen switching: `showScreen(id)` toggles `.active` on `#screen-{id}` elements
- Panel switching within screens: `showStuPanel(id)`, `showDispPanel(id)` toggle `.hidden`
- **All `onclick` in HTML must call `App.*` methods** — standalone functions are also exposed as `window.*` globals for safety

### Backend — `Code.gs`
- Google Apps Script deployed as Web App (GET + POST)
- Data stored in Google Sheets (4 sheets: Players, Questions, Answers, GameSessions)
- Entry points: `doGet(e)` / `doPost(e)` → `handleGet()` / `handlePost()`

### API Communication
- GET: `apiGet({ action, ...params })` → builds URL query string, fetches
- POST: `apiPost({ action, ...data })` → JSON body
- `APPS_SCRIPT_URL` in `index.html` must be set to the deployed script URL

## Key Constants (both files must match)
| Constant | `index.html` | `Code.gs` |
|---|---|---|
| Teacher password | `TEACHER_PW_LOCAL = '1234'` | `TEACHER_PASSWORD = '1234'` |
| Spreadsheet ID | n/a | `SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'` |
| API URL | `APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE'` | (the deployment URL) |
| Poll interval | `REFRESH_INTERVAL = 4000` (ms) | n/a |

## Design System
- Style: Glassmorphism / iPhone Liquid Glass
- Background: deep-ocean gradient (no images, pure CSS)
- Glass cards: `backdrop-filter: blur(18px)`, `rgba(255,255,255,.13)` background, `border-radius: 28px`
- CSS custom properties in `:root` — always use variables (`--glass-blur`, `--glass-border`, etc.)
- Utility classes: `.px-box` (glass card), `.px-btn` + color modifiers (`-cyan`, `-yellow`, `-green`, `-red`, `-purple`, `-gray`)
- Answer buttons: `.ans-btn.A/B/C/D` with colored glass gradients
- **No emojis anywhere** — in HTML text, JS strings, or injected innerHTML
- Animations: `@keyframes` only — `screenFadeIn`, `tabFadeIn`, `modalSlide`, `bubbleUp`, `scorePop`, `bobMascot`
- Fonts: Sarabun (Thai UI) + Press Start 2P (pixel headings/numbers)

## Question Types
| Value | Label |
|---|---|
| `multiple_choice` | Multiple Choice (A/B/C/D) |
| `true_false` | True / False (A=ถูก, B=ผิด) |
| `poll` | Poll (A/B/C/D, no correct answer) |
| `brainstorm` | Brainstorm (free text) |
| `word_cloud` | Word Cloud (free text) |

## State Object (`S`)
```js
S.teacher   // gameId, pin, password, currentQ, totalQ, status
S.display   // gameId, lastStatus, lastQNum, refreshInterval
S.student   // gameId, playerId, nickname, pin, hasAnswered, answerStartTime
S.confirmCallback  // set before opening modal-confirm, called on confirmOk()
S.questionRows     // array of question objects being edited in modal
S.editingSetId     // null = new set, string = editing existing set
```

## Game Status Flow
`waiting` → `playing` → `closed` / `showing_answer` → `leaderboard` → (next question) → `ended`

## Rules
- Do not add emojis to any file
- Do not add inline `<style>` tags inside `<body>` — all CSS goes in the single `<head>` `<style>` block
- When adding `onclick` handlers in HTML, always use `App.*` form
- When adding new App methods that need to be called from HTML, also expose as `window.*` at the bottom of `<script>`
- Do not split into multiple files — keep everything in `index.html`
- `Code.gs` runs in Google Apps Script V8 runtime — use `var` / ES5 patterns, avoid `import`/`export`

# 🌶️ Chili Cook-Off Judge App

Judges scan a QR code, check in with a name + 4-digit PIN, and score each chili 1–10 on Taste, Texture, Appearance and Aroma (out of 40). The head judge sees a live, auto-tabulated scoreboard.

**Format (default):** 20 chilis (A–T) split into 4 groups of 5. Judges are auto-balanced across the groups as they check in. The top chili in each group (by average score) goes to the Final 4, every judge scores all four, and the highest average wins it all. Ties are broken by Taste, then Aroma, and you can override a finalist pick if needed.

- **Judges:** `https://samahlersmeyer-prog.github.io/chili-judge/`
- **Head judge:** same link + `#admin` (master PIN, default `2026`)

## Setup (about 5 minutes)

1. **Backend:** create a Google Sheet → Extensions → Apps Script → paste `Code.gs` → Deploy → New deployment → Web app (Execute as **Me**, access **Anyone**) → copy the `/exec` URL.
2. **Connect:** paste that URL into `config.js`.
3. **Host:** push this folder to a GitHub repo → Settings → Pages → Deploy from branch `main` / root.
4. Open `#admin`, enter the PIN, set the event name and (optionally) cook names, then print the QR code tab.

With `config.js` left empty, the app runs in **demo mode** (saves only on that device, plus a "Simulate 10 judges" button).

## Event-day flow
Warm-up (judges check in) → **Start Round 1** → **Lock Round 1 & pick finalists** → **Lock the Final & crown a winner** → **Start the reveal** (tap through 4th → champion with confetti).

Every scorecard also lands in the Google Sheet (`Scores` tab), and the Scores tab in the app has a CSV download.

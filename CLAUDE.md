# CLAUDE.md

Guidance for AI assistants working in this repository.

## Overview

This is the marketing website for **優學補習班 (Youxue Cram School)**, a
Taiwanese tutoring / cram school. It is a **static, multi-page website** —
plain HTML, CSS, and vanilla JavaScript with no build step, no framework,
no package manager, and no backend. All content is in **Traditional Chinese
(`lang="zh-TW"`)**.

The centerpiece feature is a client-side **course quote generator** on
`quote.html`, which lets staff build an itemized price quote and print it
(or save as PDF) directly from the browser.

## Tech Stack

- **HTML5** — one file per page, hand-authored.
- **Bootstrap 5.3.3** — loaded from jsDelivr CDN (CSS + `bootstrap.bundle.min.js`).
- **Bootstrap Icons 1.11.3** — CDN, used via `<i class="bi bi-...">`.
- **Noto Sans TC** — Google Fonts, the site's primary typeface.
- **Vanilla JavaScript** — no libraries beyond Bootstrap's bundle.
- **No build tooling** — no `package.json`, npm, bundler, or transpiler.

Everything runs directly in the browser. Because dependencies are all CDN
links, an internet connection is needed for correct styling/fonts.

## Project Structure

```
learning-class/
├── index.html        # Home / landing page (hero, features, courses, testimonials, CTA)
├── about.html        # 關於我們 — about the school
├── pricing.html      # 課程費用 — course pricing tables
├── promotions.html   # 文宣資料 — promotions / announcements / downloadable materials
├── quote.html        # 開立報價單 — interactive quote generator (uses js/quote.js)
├── css/
│   └── style.css     # Single shared stylesheet for all pages
├── js/
│   ├── main.js       # Global script: navbar shadow-on-scroll (loaded on every page)
│   └── quote.js      # Quote generator logic (loaded only on quote.html)
└── README.md         # Minimal — repo name only
```

Every page loads `css/style.css` and `js/main.js`. Only `quote.html`
additionally loads `js/quote.js`.

## How to Run

There is no build or server requirement. Options:

- Open any `.html` file directly in a browser, **or**
- Serve the directory for correct relative paths, e.g.
  `python3 -m http.server 8000` then visit `http://localhost:8000/`.

There are **no tests, linters, or CI** configured in this repo.

## The Quote Generator (`js/quote.js` + `quote.html`)

This is the only non-trivial logic in the project. Key facts:

- **Price data lives in `js/quote.js`** in three constants at the top:
  - `PRICE_TABLE` — NT$/hr by grade level (`國小`/`國中`/`高中`) and teaching
    format (`個人教學`/`小班教學`/`大班教學`).
  - `SUBJECTS` — list of subjects available per grade level, used to populate
    the subject dropdown when a level is chosen.
  - `DISCOUNT_MAP` — month-count → multiplier (`3 months → 0.95`,
    `6 months → 0.90`). The discount is based on the **largest** month value
    across all course rows (`getMaxMonths()`).
- **Monthly fee formula:** `unitPrice × hoursPerSession × sessionsPerWeek × 4 × months`
  (4 = weeks per month). See `calcRow()`.
- **Flow:** `addCourseRow()` builds a table row with cascading `<select>`s.
  Changing the level fires `updateSubjects()` → `updatePrice()` → `calcRow()`;
  `updateTotal()` recomputes the summary and applies the discount.
- **Preview & print:** `generateQuote()` validates input, `buildPreview()`
  renders the printable quote HTML into `#quote-preview` inside a Bootstrap
  modal, and `printQuote()` calls `window.print()`. Print styling is handled
  by the `@media print` block in `style.css` and the `no-print` class on
  chrome elements.
- Event handlers are wired via **inline `onclick`/`onchange` attributes** in
  the dynamically generated row HTML — keep new handlers consistent with this
  pattern (functions are global in `quote.js`).

> **Important:** There are **two independent sets of prices**. `quote.html`'s
> "費率速查" reference table and `PRICE_TABLE` in `quote.js` must stay in sync
> with each other (600/350/250, 700/400/300, 800/480/350). The homepage
> "授課方式" section on `index.html` shows separate marketing "起" (starting-from)
> rates ($550/$300/$220) that are intentionally different — don't try to
> reconcile those with the quote engine.

## Conventions

- **Language:** All user-facing text is Traditional Chinese. Match the
  existing tone and terminology (e.g. 報價單 = quote, 授課形式 = teaching format).
- **Styling:** Prefer Bootstrap utility classes. The brand color is indigo
  `#6366f1` (`--bs-primary`), overridden in `style.css`. Reusable custom
  classes live in `style.css` (`.hero-bg`, `.blob`, `.gradient-text`,
  `.glass-card`, `.feature-card`, `.course-card`, `.testimonial-card`,
  `.step-badge`, `.wave-bottom`, etc.). Inline `style="..."` is used liberally
  throughout for page-specific tweaks — this is the established pattern.
- **Shared layout:** The navbar and footer are **duplicated in each HTML
  file** (there is no templating/includes). If you change navigation links,
  branding, contact info, or footer content, update **every page**
  consistently.
- **Contact/brand constants** repeated across pages: 台北市中山區學習路100號,
  (02) 2345-6789, info@youxue.edu.tw, brand mark "優".
- **Naming:** JS uses camelCase functions and `SCREAMING_SNAKE_CASE` for the
  data-table constants. DOM element IDs in the quote table follow the
  `<field>-<rowId>` pattern (e.g. `level-3`, `subtotal-3`).

## Making Changes

- **Adding/adjusting prices or subjects:** edit the constants at the top of
  `js/quote.js`, and update the reference table in `quote.html` to match.
- **Adding a new page:** copy an existing page's `<head>`, navbar, and footer
  so the CDN links, fonts, and navigation stay consistent; add the new link to
  the navbar/footer on all pages.
- **Global behavior:** put site-wide scripts in `js/main.js` (loaded
  everywhere); keep page-specific logic in its own file loaded only where
  needed (as `quote.js` is).
- Verify changes by opening the affected page(s) in a browser; for the quote
  generator, walk through building a row and previewing/printing.

## Git Workflow

- Development branch for this work: `claude/claude-md-docs-efj87o`.
- Default branch: `main`.
- Commit with clear, descriptive messages; push to the designated feature
  branch. Do not open a pull request unless explicitly asked.

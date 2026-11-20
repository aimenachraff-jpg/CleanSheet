# Landing-page specification — CleanSheet

**Status:** specification only. No implementation of a paid service, telemetry, cloud
upload, or account system is implied or planned.

## Purpose

A single, trustworthy landing page that tells a nontechnical business user (shop
manager, accountant, small-business owner) what CleanSheet does, why it is safe to use
with business data, and how to install it in under two minutes.

## Audience

Primary: nontechnical business users in Algeria, Morocco, Tunisia, Egypt and France who
receive messy supplier/customer Excel/CSV files and need them cleaned before invoicing,
accounting, or reporting.

Secondary: technical evaluators who want to confirm the tool is offline, rule-based, and
free of hidden costs.

## Core messages (in order)

1. **Headline.** Clean your Excel and CSV files offline — on your computer, by you.
2. **The promise.** Local processing. Your data never leaves your machine. No account,
   no cloud, no telemetry, no hidden cost.
3. **What it fixes.** Extra spaces and inconsistent casing, mixed price and date
   formats, Arabic-Indic digits, country phone formats (+213 / +212 / +216 / +20 / +33),
   Arabic/French/English headers, exact duplicates, two-level headers, "Total" rows.
4. **How it works (3 steps).** Pick a file → confirm the guessed column types → Clean.
   You get the cleaned file, a change log, and a Review sheet of anything that needs
   your decision.
5. **Trust.** Rule-based (no AI). Ambiguous values are flagged, never guessed. Original
   file never modified. Outputs: cleaned workbook, change log, review sheet, markdown
   report.
6. **Install.** Windows: download `CleanSheet.exe` from Releases (unsigned → SmartScreen
   "More info, then Run anyway"). Any OS with Python: `pip install cleansheet`.
7. **Languages.** English / Arabic / French.

## Page sections

1. Navbar: CleanSheet logo · How it works · Locales · Install · GitHub
2. Hero: headline, one-line value prop, primary "Download for Windows" button,
   secondary "pip install cleansheet" code block, OS badges.
3. "What it cleans" — before/after style mini-table (messy → clean) with Arabic,
   French and English examples.
4. "How it works" — 3 illustrated steps (pick, confirm, clean) + screenshot/GIF.
5. "Your data stays yours" — privacy bullets (offline, no telemetry, no cloud, original
   untouched, flag-not-guess).
6. "Supported formats" — locales table (dz/ma/tn/eg/fr/generic) and file types.
7. Install section — Windows exe + pip command + SmartScreen note.
8. Footer: MIT license, GitHub link, privacy/security/roadmap links, "no AI / no
   telemetry / no accounts" badge.

## Design direction

- Clean, professional, business-appropriate (not playful). Dark-blue/white accent
  matching the workbook header color (#1F3A5F).
- RTL-safe where Arabic is shown; language switcher (EN / AR / FR).
- One screenshot or the demo GIF; no stock photos of people.
- Static HTML/CSS only — no build step, no tracking scripts, no cookies.

## What is explicitly NOT on the page

- No pricing, plans, or "upgrade" CTAs (the tool is free and MIT-licensed).
- No "create an account", "sign in", or "free trial".
- No telemetry/analytics scripts.
- No cloud-upload or "secure online processing" claims — processing is local.
- No AI/ML claims.

## Success measure

A business user can read the page, download and run CleanSheet, clean one file, and
understand the Review sheet within 10 minutes — without contacting support.

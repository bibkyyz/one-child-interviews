# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single static page (no build step, no dependencies, no package.json) that lets a candidate pick an available interview date/time and submit their contact info. There is no local dev server or test suite — open `index.html` directly in a browser to preview changes.

## Architecture

Three parts, two of which run in completely different environments:

1. **Frontend (`index.html`, `script.js`, `style.css`)** — static, deployed as-is (e.g. GitHub Pages or similar static host). All scheduling logic lives in `script.js`:
   - `availability` (top of `script.js`) is a hardcoded array of `{ date, times }` — this is the source of truth for which dates/times are offerable. To change bookable slots, edit this array directly.
   - The calendar is built from scratch (no library): `renderCalendar()` draws the month grid from `availability`, `selectDate()`/`renderTimeSlots()` handle picking a day then a time, `selectTime()` fills the hidden `timeSlot` input that actually gets submitted.
   - On load, `loadBookedSlots()` GETs the Apps Script endpoint to fetch already-booked slots (as `"<formatted date> at <time>"` strings) and disables/strikes them through in the UI. If the fetch fails, it fails open (shows everything as available) rather than blocking the form.
   - Form submit POSTs `{ name, email, phone, timeSlot }` as `text/plain` JSON to the same endpoint (this content-type is intentional — it avoids a CORS preflight against Apps Script).

2. **Backend (`apps-script.gs`)** — Google Apps Script, **not deployed from this repo**. It is pasted manually into the Apps Script editor attached to a Google Sheet ("One Child Interviews" sheet), then deployed as a web app from that editor. There is no CLI/CI path that pushes this file anywhere — after editing `apps-script.gs`, a human must copy/paste it into the Apps Script editor and redeploy for changes to take effect.
   - `doGet` returns already-booked time slots by reading column D (Time Slot) of the sheet.
   - `doPost` re-checks the slot isn't already booked (double-booking guard, serialized with `LockService`), then appends a row `[name, email, phone, timeSlot, timestamp]`.
   - The sheet's column order is load-bearing: `getBookedSlots()` hardcodes column D as the time slot column.

3. **The glue**: `SCRIPT_URL` at the top of `script.js` is the deployed Apps Script web app URL. If the Apps Script is redeployed (new deployment, not just "save"), this URL changes and must be updated in `script.js`.

## Key coupling to watch when editing

- The `timeSlot` string format (`"<full date label> at <time>"`, built in `renderTimeSlots()`) is the join key used to detect booked slots on both ends — it must stay identical between what's stored in the sheet and what the frontend re-generates on each load. Changing `formatDateLabel()` or the time strings in `availability` will silently break booked-slot detection for existing rows.
- `availability` dates are hardcoded strings; nothing expires or regenerates them automatically. Keeping the schedule current is a manual edit to `script.js`.

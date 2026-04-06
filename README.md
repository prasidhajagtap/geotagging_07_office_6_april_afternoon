# Seamex Geo Attendance

**Module:** Geo Attendance
**Author:** Prasidha Jagtap — Assistant Manager, IT · Aditya Birla Group (Seamex)
**Version:** v07 — Definitive Golden Build
**Stack:** Vanilla JS · Pico CSS v2 · Supabase JS v2 · Google Fonts (DM Sans + DM Mono)

---

## What This Module Does

Seamex Geo Attendance is a mobile-first web app that allows field employees to clock in and clock out with GPS-verified location tagging. Attendance records are submitted to a Supabase database. The app runs entirely in the browser with no native installation required.

**Core flow:**
1. Employee logs in with Full Name and Poornata ID
2. Enters clock-in location → GPS captured → shift timer starts
3. Enters clock-out location → GPS captured → shift duration calculated
4. Reviews the shift details before final submission
5. One-click submit stores the record in the `attendance` table

---

## File Structure

```
├── index.html    HTML structure + all element IDs documented inline
├── style.css     All CSS — design tokens, animations, dark/light theme
├── script.js     All JS — business logic, GPS, validation, Supabase calls
└── README.md     This file
```

The Seamex logo file must be placed in the same folder:
```
NEW Seamex logo_testing march 11.png
```

---

## First-Time Setup

### 1. Supabase Database

Create a table called `attendance` with the following columns:

| Column | Type |
|---|---|
| `id` | uuid (primary key, default gen_random_uuid()) |
| `user_name` | text |
| `employee_id` | text |
| `clock_in_time` | timestamptz |
| `clock_in_coords` | text |
| `clock_in_location_name` | text |
| `clock_out_time` | timestamptz |
| `clock_out_coords` | text |
| `clock_out_location_name` | text |
| `status` | text |
| `created_at` | timestamptz (default now()) |

### 2. Row Level Security (CRITICAL)

Enable RLS on the `attendance` table, then apply these policies:

```sql
-- Allow field employees to submit attendance
CREATE POLICY "Allow insert for anon"
ON attendance FOR INSERT TO anon
WITH CHECK (true);

-- Block all reads/updates/deletes from the client
CREATE POLICY "Deny select for anon"
ON attendance FOR SELECT TO anon
USING (false);
```

> **Important:** The Supabase anon key is visible in the browser. RLS is the only thing preventing unauthorized access. Never disable RLS on this table.

### 3. Deploy

The app is a static site. Deploy to:
- GitHub Pages (current)
- Any static host (Netlify, Vercel, Cloudflare Pages)
- SharePoint (see migration notes below)

---

## Features

| Feature | Details |
|---|---|
| GPS clock-in/out | `navigator.geolocation` with `maximumAge:0` (always fresh) |
| Location autofill | Last 6 used locations cached in `localStorage` |
| Quick-select chips | Last 3 locations shown as one-tap pill buttons |
| Live shift timer | `requestAnimationFrame` loop — no drift, no lag |
| Dark/light theme | Auto-detects time of day (6am–6pm = light), manual override |
| Ambient background | CRED-style CSS-only animations per time of day |
| Session persistence | State survives page refresh until midnight |
| Day reset | `localStorage` cleared at midnight for fresh daily attendance |
| Recovery button | Appears after 5s stuck — DOM-only restore, session preserved |
| Fix Clock-In | Modal: location name only. Time correction requires HR. |
| Redo Clock-Out | Modal disclaimer → full GPS redo. Timer resumes. |
| Responsive | Tested: 320px → foldables → tablets · iOS safe areas |

---

## Input Validation Rules

All validation is applied consistently on every input field, every page, and re-checked on every button click.

| Field | Rule |
|---|---|
| Full Name | `/^[a-zA-Z\s]{2,60}$/` — letters and spaces only, min 2 chars |
| Poornata ID | `/^[0-9]{3,12}$/` — digits only, 3–12 chars |
| Location Name | `/^[a-zA-Z0-9 \-]{2,60}$/` — blocks all SQL/XSS chars |

Sanitization via `sanitize()` is applied to all strings before any Supabase write as defence-in-depth, even though Supabase uses parameterized queries server-side.

---

## GPS Freeze Fix (v07)

The redo clock-out button was freezing on iOS Safari. Root cause: `getCurrentPosition` on a repeated call in the same session could silently fire neither the success nor the error callback.

**Fix applied in `script.js`:**

1. `getCoords()` now uses `Promise.race([geoPromise, safetyPromise])`. The safety promise resolves `null` at 9 seconds guaranteed.
2. The `btn-co` and `btn-ci` handlers both have a `finally` block that re-enables the button on every outcome: GPS success, GPS error, GPS timeout, or any thrown exception.

The button **cannot** stay frozen beyond 9 seconds under any circumstances.

---

## Migration Checklist

### GitHub Pages → SharePoint SPFx
- Bundle `supabase.min.js` locally in the SPFx webpart assets
- Remove the CDN `<script>` tag from `index.html`
- Add Supabase as a dependency in the SPFx `package.json`
- Change `submitted_via: 'web'` to `'sharepoint'` in the payload (script.js)

### Temporary Login → Azure AD SSO
- Replace the entire `#auth-sec` section in `index.html` with an MSAL.js redirect flow
- Map `Azure displayName` → `U.name` and `Azure employeeId` → `U.id` in script.js
- No changes needed in `#main-sec` after this migration

### Production Supabase
- Swap `SUPABASE_URL` and `SUPABASE_KEY` constants in `script.js`
- Keep the `attendance` table schema identical
- Enable RLS with the same policies on the new instance

### Multi-Office
- Uncomment `branch_code: 'HO_AIROLI'` in the payload object in `script.js`
- Parameterize the value per deployment

---

## Browser Support

| Browser | Status |
|---|---|
| Chrome Android | ✅ Fully supported |
| Safari iOS | ✅ Fully supported (GPS freeze fixed in v07) |
| Samsung Internet | ✅ Fully supported |
| Firefox Android | ✅ Supported |
| Desktop Chrome/Edge | ✅ Supported (for testing) |

---

## Changelog

| Version | Author | Notes |
|---|---|---|
| v04 | Prasidha Jagtap | Initial full-feature build |
| v05 | Prasidha Jagtap | CRED-style UI, Pico CSS, dark/light, ambient background |
| v06 | Prasidha Jagtap | Refresh-to-login fix, single rAF loop, redo timer fix |
| v07 | Prasidha Jagtap | GPS freeze fix (Promise.race + finally), recovery button, README |

---

*Maintained by Prasidha Jagtap · Aditya Birla Group (Seamex) · Reliable Tech Park, Airoli*

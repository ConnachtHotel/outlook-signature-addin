# Connacht Signatures Add-in — Maintainer's Guide

A practical guide to maintaining, extending and debugging the
`outlook-signature-addin` repository. Written for someone who has
never touched an Outlook add-in before. Read at least sections 1, 2,
3 and 8 before making changes.

---

## Table of contents

1. [What this thing is and how it works](#1-what-this-thing-is-and-how-it-works)
2. [Repository layout](#2-repository-layout)
3. [The signature flow, end to end](#3-the-signature-flow-end-to-end)
4. [Common maintenance tasks](#4-common-maintenance-tasks)
   - 4.1 [Adding a new hotel / email domain](#41-adding-a-new-hotel--email-domain)
   - 4.2 [Adding a new team override (e.g. Green Tourism)](#42-adding-a-new-team-override)
   - 4.3 [Changing colours / fonts for a hotel](#43-changing-colours--fonts-for-a-hotel)
   - 4.4 [Overriding details for a single person (`EMAIL_OVERRIDES`)](#44-overriding-details-for-a-single-person)
   - 4.5 [Adding or replacing a logo / banner](#45-adding-or-replacing-a-logo--banner)
   - 4.6 [Editing the signature layout (`buildSignatureHtml`)](#46-editing-the-signature-layout)
   - 4.7 [Editing the disclaimer or sign-off](#47-editing-the-disclaimer-or-sign-off)
   - 4.8 [Bumping the cache-bust `?v=` value](#48-bumping-the-cache-bust-v-value)
   - 4.9 [Editing the manifest](#49-editing-the-manifest)
5. [Deploying changes](#5-deploying-changes)
6. [Debugging](#6-debugging)
   - 6.1 [The status notification bar](#61-the-status-notification-bar)
   - 6.2 [Viewing the console logs](#62-viewing-the-console-logs)
   - 6.3 [What the log helpers do](#63-what-the-log-helpers-do)
   - 6.4 [Common symptoms and what they mean](#64-common-symptoms-and-what-they-mean)
   - 6.5 [Outlook Classic specifics](#65-outlook-classic-specifics)
   - 6.6 [The safety timeout](#66-the-safety-timeout)
   - 6.7 [The Azure Function — what to check](#67-the-azure-function--what-to-check)
   - 6.8 [Runtime logging on the desktop client](#68-runtime-logging-on-the-desktop-client)
7. [Reference: every function in `autorun.js`](#7-reference-every-function-in-autorunjs)
8. [Compatibility rules — DO NOT BREAK THESE](#8-compatibility-rules--do-not-break-these)
9. [Glossary](#9-glossary)
10. [Quick checklist for common changes](#10-quick-checklist-for-common-changes)

---

## 1. What this thing is and how it works

This repository is an **Outlook add-in** that automatically inserts a
branded email signature whenever any Connacht Hospitality Group
employee starts composing a new message. The user does nothing
manually — Outlook fires an event when a compose window opens, our
JavaScript runs in a tiny embedded webview, fetches the employee's
details, builds the signature HTML, and pushes it into the message
body.

### High-level architecture

```
  +-----------------------+
  | User clicks "New      |
  | mail" in Outlook      |
  +-----------+-----------+
              |
              v
  +-----------------------+
  | Outlook fires the     |
  | OnNewMessageCompose   |
  | LaunchEvent           |
  +-----------+-----------+
              |
              v
  +-----------------------------+        +----------------------+
  | autorun.js is loaded inside |  --->  | manifest.xml tells   |
  | an embedded webview         |        | Outlook where the    |
  |  - WebView2 (modern)        |        | JS and HTML live     |
  |  - Trident/IE11 (legacy)    |        +----------------------+
  +-----------+-----------------+
              |
              v
  +-----------------------------+
  | onNewMessageCompose() runs  |
  |  - reads the user's email   |
  |  - asks the Azure Function  |
  |    for their employee data  |
  |  - picks the hotel config   |
  |  - applies overrides        |
  |  - builds the HTML          |
  |  - calls setSignatureAsync  |
  +-----------+-----------------+
              |
              v
  +-----------------------------+        +----------------------+
  | Azure Function              |  --->  | Microsoft Graph      |
  | (NOT in this repo)          |  <---  | (Entra/Azure AD)     |
  | reads the user from Entra,  |        +----------------------+
  | returns name / title /      |
  | phone / banner URL / etc.   |
  +-----------------------------+
```

Two things to keep in mind:

- The **code in this repo is just the client side** of the add-in.
  The Azure Function — the thing that actually looks up an
  employee in Entra ID via Microsoft Graph — lives in a separate
  Azure project. Its URL is hard-coded in `autorun.js` as
  `API_URL`.
- The **manifest is registered with Microsoft 365 once** in the
  admin centre. After that, every change you push to GitHub Pages
  is picked up automatically the next time Outlook fetches the JS
  (controlled by the cache-bust value — see §4.8 and §5).

### Where the files are hosted at runtime

GitHub Pages serves the contents of the `main` branch at
`https://ConnachtHotel.github.io/outlook-signature-addin/`, so:

- `manifest.xml` → uploaded to the admin centre (and referenced
  from `https://ConnachtHotel.github.io/outlook-signature-addin/manifest.xml`)
- `src/autorun.html` → `https://…/src/autorun.html?v=1.0.4`
- `src/autorun.js` → `https://…/src/autorun.js?v=1.0.4`
- `assets/*.gif` and `assets/*.jpg` → referenced by full URL from
  the JS (banners) and from the Azure Function

---

## 2. Repository layout

```
outlook-signature-addin/
├── manifest.xml              ← The Office add-in manifest. Tells Outlook what the
│                               add-in does, where its JS lives, what permissions
│                               it needs, what events it listens for.
│
├── src/
│   ├── autorun.html          ← Tiny HTML shell. Loaded by older Outlook clients
│   │                           that use the Trident/IE11 webview. Its only job is
│   │                           to load office.js and then autorun.js.
│   │
│   └── autorun.js            ← The actual logic. All the hotel configs, the
│                               signature template, the event handler, everything.
│                               This is the file you'll edit 95 % of the time.
│
├── assets/
│   ├── newLogo.gif           ← Banner used by the Connacht hotels for the Green
│   │                           Tourism / Green Meetings team.
│   ├── logo.gif              ← Old generic Connacht banner (legacy / fallback).
│   ├── hawthornLogo.jpg      ← Banner for the Hawthorn hotel.
│   ├── hydeHotelLogo.jpg     ← Banner for the Hyde hotel.
│   ├── activeFitnessLogo.jpg ← Banner for Active Fitness.
│   ├── mfitzLogo.jpg         ← Banner for M. Fitzgerald's Bar.
│   ├── icon-16.png …         ← Outlook add-in icons (different sizes).
│   └── icon.png
│
└── data/
    └── signatures.json       ← LEGACY. Was used before the Azure Function existed,
                                when the add-in read employee data from this JSON.
                                Not read at runtime any more. Safe to leave; safe to
                                delete. Kept for reference.
```

There is no `package.json`, no build step, no test runner. This is
deliberately a zero-dependency static site — the JS runs as-is in
the browser/webview that Outlook gives it.

---

## 3. The signature flow, end to end

Walking through what happens when a user named, say, *Nathan
Duggan* clicks "New mail" in Outlook:

1. Outlook fires the **`OnNewMessageCompose`** launch event
   (declared in `manifest.xml`).
2. Outlook starts a small webview process and loads
   `autorun.js` (modern hosts) or `autorun.html` which then loads
   `autorun.js` (legacy hosts).
3. The file parses. Critically, `Office.actions.associate(
   "onNewMessageCompose", onNewMessageCompose)` runs **at the
   bottom of the file at top level** (see §8) and tells the
   runtime "if you fire that event, call this function."
4. The runtime invokes `onNewMessageCompose(event)`.
5. The function:
   1. Adds the "Signature add-in running…" status notification
      (so the user can see something happened).
   2. Starts a 55-second safety timer (mobile kills add-ins after
      60 seconds — this exits cleanly before that).
   3. Calls `getEmployeeData()`, which does an XHR GET to
      the Azure Function at `API_URL + "?email=" + userEmail`.
   4. The Azure Function looks Nathan up in Entra and returns
      JSON like:
      ```json
      {
        "name": "Nathan Duggan",
        "title": "IT Intern",
        "email": "NDuggan@chgl.ie",
        "phone": "+353 91 …",
        "Mphone": "+353 87 …",
        "banner": "https://ConnachtHotel.github.io/.../newLogo.gif",
        "teamCode": "GREEN"
      }
      ```
   6. `getConfigForEmail(employee.email)` finds the matching
      `HOTEL_CONFIG` entry by suffix (`@chgl.ie` → Connacht
      config). If nothing matches, `HOTEL_CONFIG["default"]`
      is used.
   7. **Single-user overrides** are applied from
      `EMAIL_OVERRIDES` (rarely used; useful for forcing a
      specific person's title or name).
   8. **Team overrides** are applied: if the employee has a
      `teamCode` (e.g. `"GREEN"`) and the matched hotel has a
      `teamOverrides[teamCode]` entry, the banner / team name
      gets replaced.
   9. Fallbacks fill in missing values — the hotel's default
      website and address are used if the employee record
      didn't supply them.
   10. `buildSignatureHtml(employee, config)` produces the
       final HTML string.
   11. The HTML is pushed into the message body via
       `body.setSignatureAsync()` on modern hosts, or
       `body.prependAsync()` on older desktop builds.
6. `event.completed()` is called so Outlook knows the add-in is
   done.

---

## 4. Common maintenance tasks

All of these are edits to `src/autorun.js` unless stated
otherwise.

### 4.1 Adding a new hotel / email domain

Say a new property opens at `@somenewhotel.ie`.

1. Open `src/autorun.js` and find the `HOTEL_CONFIG` object.
2. Copy one of the existing entries — `@hydehotel.ie` is a
   good template because it has no team overrides.
3. Paste it as a new key and edit:

   ```js
   "@somenewhotel.ie": {
       hotelName: "Some New Hotel",                 // shown under name + title
       website: "www.somenewhotel.ie",              // shown after "W:"
       address: "Street, Town, County, Eircode",    // shown after "A:"
       websiteUrl: "https://www.somenewhotel.ie/",  // where the banner image links to
       teamOverrides: {},                           // leave empty unless you need 4.2
       style: {
           nameColor: "#000000",
           nameSize: "14px",
           titleColor: "#666666",
           textColor: "#333333",
           dividerColor: "#cccccc",
           linkColor: "#333333",
           disclaimerColor: "#999999",
           fontFamily: "Arial,Helvetica,sans-serif"
       }
   },
   ```

4. The matching is done by **suffix** in `getConfigForEmail`:
   any user whose email ends in `@somenewhotel.ie` will get
   this config.
5. The banner image itself is **selected by the Azure
   Function** based on the email suffix, not by this file.
   If the new hotel needs a new banner, you also need to:
   - Drop the image into `assets/` (see §4.5).
   - Update the Azure Function so it returns the right
     banner URL for the new domain.
6. Bump the cache-bust (§4.8) and commit/push.

`hotelName` is **optional**. Some entries (`@chgl.ie`,
`@galwayhooker.ie`, `@connachthospitalitygroup.ie`) omit it on
purpose so no hotel line shows under the user's title — useful
for group-level addresses.

### 4.2 Adding a new team override

Team overrides let users in the *same email domain* get a
different banner or extra label — for example, the Green
Tourism team at Connacht. The team membership is read from the
**`faxNumber` field on the user's Entra account**, which the
Azure Function returns as `teamCode`.

Example: add a new "Family Friendly" team for the Connacht.

1. Decide the team code. Whatever string you put in the
   user's `faxNumber` field in Entra is what `teamCode` will
   be. Convention is upper-case. Say `"FAMILY"`.
2. Set that string on the relevant users in Entra / admin
   centre → user properties → fax number.
3. In `autorun.js`, find the hotel's `teamOverrides` object
   and add an entry:

   ```js
   teamOverrides: {
       "GREEN": {
           banner: "https://ConnachtHotel.github.io/outlook-signature-addin/assets/newLogo.gif",
           teamName: "Member of Green Tourism & Green Meetings"
       },
       "FAMILY": {
           banner: "https://ConnachtHotel.github.io/outlook-signature-addin/assets/familyFriendlyBanner.gif",
           teamName: "Family Friendly Member"
       }
   }
   ```

4. Both keys are optional:
   - `banner` — override the default banner for this team.
   - `teamName` — show an extra line in the signature.
5. The lookup is case-insensitive (the code does
   `teamCode.trim().toUpperCase()`), but it's still good
   hygiene to write the key in the same case as the Entra
   value.
6. Bump the cache-bust, commit, push.

### 4.3 Changing colours / fonts for a hotel

Each hotel has a `style` block. Every key maps to a specific
visible element:

| Key                 | Affects                                                |
|---------------------|--------------------------------------------------------|
| `nameColor`         | Colour of the bold employee name                       |
| `nameSize`          | Font size of the bold employee name (e.g. `"14px"`)    |
| `titleColor`        | Colour of the job title and the hotelName line         |
| `textColor`         | Default body text colour (the contact details, etc.)   |
| `dividerColor`      | The vertical line between the name and contact columns |
| `linkColor`         | Email, phone, and website links                        |
| `disclaimerColor`   | The grey disclaimer text at the bottom                 |
| `fontFamily`        | Font for the entire signature                          |

To change a colour, edit the hex value and push. Bump the
cache-bust if you want existing clients to refresh.

### 4.4 Overriding details for a single person

The `EMAIL_OVERRIDES` object (near the top of `autorun.js`) is
empty by default. It exists for cases where you want to force
specific values for one user without changing their Entra
record. Example, commented out:

```js
var EMAIL_OVERRIDES = {
    "conferencing@chgl.ie": {
        name: "Robyn O'Neill",
        title: "Meetings & Events Co-ordinator"
    }
};
```

When the add-in runs and the user's email matches a key
(case-insensitive), every property in the override object is
copied onto the employee object, replacing whatever came from
Entra. Use sparingly — the long-term fix is almost always to
fix the user's record in Entra.

### 4.5 Adding or replacing a logo / banner

1. Drop the file into `assets/` (keep it under ~200 KB —
   email clients dislike huge inline images).
2. Use a `.gif`, `.jpg`, or `.png`.
3. The full URL will be
   `https://ConnachtHotel.github.io/outlook-signature-addin/assets/<your-file>`.
4. Where to reference it depends on what you're changing:
   - **Default hotel banner** → update the Azure Function so
     it returns this URL for the matching email suffix. Not
     done in this repo.
   - **Team-override banner** → set the `banner` field in
     the relevant `teamOverrides` entry (see §4.2).
5. Commit + push. Banners are loaded fresh by Outlook each
   time the signature is built; you usually don't need a
   cache-bust just for a banner change, but bumping anyway is
   safer.

### 4.6 Editing the signature layout

The visible HTML lives in `buildSignatureHtml(emp, config)`.
The structure is three stacked `<table>`s (tables are used
deliberately — every email client renders tables consistently;
flexbox/grid does not):

1. The "Kind regards," row (sign-off — see §4.7).
2. A two-column row: name+title on the left, contact details
   on the right, separated by `dividerColor`.
3. The banner image, wrapped in an `<a>` linking to
   `config.websiteUrl`.
4. The disclaimer.

Rules of thumb when editing:

- Stick to **inline `style=""`** — many clients strip
  `<style>` blocks and class attributes.
- Use `<table>` for layout, not `<div>` + CSS.
- Wrap optional fields with `(emp.something ? '…' : '')` so
  the row collapses when the data is missing — this is the
  pattern used everywhere already.
- Don't string-template with backticks. Use `'string ' +
  variable + ' more string'` style concatenation — IE11 has
  no template literals and the whole file must parse on IE11
  (see §8).

### 4.7 Editing the disclaimer or sign-off

Both live inside `buildSignatureHtml`:

- **Sign-off** ("Kind regards,") is the first `<table>`. Edit
  the `<td>` text.
- **Disclaimer** is the last `<table>`. Edit between the
  `<td>` tags.

You can include `<br/>` for line breaks. Avoid links inside
the disclaimer unless you accept that some clients will
flag/rewrite them.

### 4.8 Bumping the cache-bust `?v=` value

The cache-bust query string is how you force Outlook to
re-download `autorun.js` and `autorun.html`. **You almost
always want to bump it when you change JS or HTML.**

Three places to keep in sync. The value must be identical
across all three:

| File              | Line                                                                                  |
|-------------------|---------------------------------------------------------------------------------------|
| `manifest.xml`    | `<bt:Url id="WebViewRuntime.Url" DefaultValue="…/autorun.html?v=1.0.4"/>`              |
| `manifest.xml`    | `<bt:Url id="JSRuntime.Url" DefaultValue="…/autorun.js?v=1.0.4"/>`                     |
| `src/autorun.html`| `<script type="text/javascript" src="autorun.js?v=1.0.4"></script>`                    |

To bump from `1.0.4` to `1.0.5`:

1. Edit all three occurrences (find-and-replace works fine).
2. Commit + push.
3. **Re-upload `manifest.xml`** to the admin centre. The
   manifest's URLs are what's actually registered with
   Microsoft 365; just pushing to GitHub Pages is **not
   enough** because users' Outlook clients keep using the
   `?v=` they saw when the app was first installed.

The cache-bust number is **only** a cache identifier. It has
no relation to the manifest's `<Version>` element. The comment
at the bottom of `manifest.xml` notes that `<Version>
1.0.3.0` and above failed validation — keep `<Version>` at
`1.0.1.0` unless you have a clear reason to change it.

### 4.9 Editing the manifest

Most of the time you do **not** need to touch `manifest.xml`.
Reasons you might:

- Adding a new permission (rare — `ReadWriteItem` is what we
  use and that covers everything).
- Adding a new launch event (e.g. on reply, on send).
- Changing the minimum mailbox version.
- Changing icons.

Things to know:

- `<Version>` is the manifest version visible in the admin
  centre — bump this only when the manifest itself changes
  meaningfully. Keep it ≤ `1.0.2.x` until the validation
  problem noted in the trailing comment is investigated.
- After any manifest change, you **must re-upload it to the
  admin centre**. GitHub Pages serving a new manifest is not
  enough — Microsoft 365 caches the manifest server-side.
- Always re-validate the manifest after editing. Microsoft's
  Office Add-in Validator (npm install -g
  office-addin-manifest then `office-addin-manifest validate
  manifest.xml`) is the canonical check.

---

## 5. Deploying changes

The flow for a typical JS-only change:

1. Edit `src/autorun.js`.
2. Bump the `?v=` cache-bust in all three places (§4.8).
3. Commit + push to `main` (or merge a PR into `main`).
4. GitHub Pages publishes the new file within ~1–2 minutes.
5. If you bumped the cache-bust in `manifest.xml`, re-upload
   `manifest.xml` to the Microsoft 365 admin centre:
   - Admin Centre → Settings → Integrated apps.
   - Find the Connacht Signatures app.
   - Edit / Update → upload the new manifest.
6. Wait 15–60 minutes for the admin centre to push the new
   manifest to all users (this is Microsoft-side, can't be
   sped up).
7. On a test machine, **fully restart Outlook**. Use Task
   Manager to kill `OUTLOOK.EXE` if necessary — minimising to
   tray is not enough.
8. Open a new mail compose window and verify.

If you're testing in a **sandbox app** (a second integrated
app you set up just for testing, recommended), all of the
above applies but only the test users who are assigned to the
sandbox app pick it up. This is the safe way to roll out
changes — never test in the production deployment.

For asset-only changes (e.g. swap a logo image), the cache-bust
bump is optional. Outlook re-fetches images on every send.

For manifest-only changes (e.g. a new icon URL), you don't
need to push to GitHub Pages at all, just re-upload the
manifest.

---

## 6. Debugging

### 6.1 The status notification bar

The first thing the add-in does is show a yellow info bar at
the top of the compose window: **"Signature add-in running…"**.

What it tells you:

- **You see the bar** → the event fired and the handler is
  executing. The problem is somewhere later (Azure Function,
  parsing, signature insertion).
- **You don't see the bar** → the event didn't fire, or the
  handler isn't being called. This is almost always a
  manifest, cache, or registration problem. See §6.4 case A.

After the bar, one of these follow-ups appears via
`notifyUser`:

| Notification text                                                  | What it means                                                |
|--------------------------------------------------------------------|--------------------------------------------------------------|
| `Your Outlook version doesn't support automatic signatures.`       | Neither `setSignatureAsync` nor `prependAsync` is available. |
| `No signature found for your account. Contact IT to get set up.`   | Azure Function returned 404 — user not in Entra or no match. |
| `Could not set signature: <message>`                               | The API call to insert the signature failed.                 |
| `Could not load signature. Check your connection.`                 | The XHR to the Azure Function failed (network or 5xx).       |

### 6.2 Viewing the console logs

Every step of the flow calls `logInfo` / `logWarn` /
`logError`. To see them you need to attach a debugger to the
add-in's webview process.

**On Outlook on the Web / New Outlook on the Web:**
1. Open Outlook in the browser.
2. Press **F12** → Console tab.
3. Start a new email — logs appear there, prefixed with
   `[ConnachtSig]`.

**On New Outlook for Windows (WebView2):**
1. Set the environment variable
   `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`
   for your user account (System → Environment variables).
2. Restart Outlook fully.
3. Open Edge → navigate to `edge://inspect`.
4. Click "inspect" on the Outlook target.
5. Start a new email — DevTools is now attached.

**On Outlook Classic for Windows (WebView2 mode):**
Same as above. If the user has Edge WebView2 installed,
Classic uses it.

**On Outlook Classic for Windows (Trident / IE11 fallback):**
There is no F12 in Trident from within Outlook. The only way
to see logs in this mode is the runtime log file — see §6.8.

**On Outlook for Mac:**
1. Outlook → Preferences → General → tick "Allow Developer
   Tools".
2. Right-click in the add-in pane → Inspect.

### 6.3 What the log helpers do

All three are thin wrappers around `console.*`. Every line is
prefixed with `[ConnachtSig]` so it's easy to filter:

```js
function logInfo(msg)  { console.log(  "[ConnachtSig]", "INFO:",  msg); }
function logWarn(msg)  { console.warn( "[ConnachtSig]", "WARN:",  msg); }
function logError(msg) { console.error("[ConnachtSig]", "ERROR:", msg); }
```

A successful run logs, in order:

```
[ConnachtSig] JS Loaded 2026-06-02T…
[ConnachtSig] INFO: OnNewMessageCompose triggered
[ConnachtSig] INFO: Current user email: NDuggan@chgl.ie
[ConnachtSig] INFO: Employee found: Nathan Duggan
[ConnachtSig] INFO: Looking up config for email: NDuggan@chgl.ie
[ConnachtSig] INFO: Team override applied — code: GREEN
[ConnachtSig] INFO: Matched config website: www.chgl.ie
[ConnachtSig] INFO: Config applied — website: www.chgl.ie
[ConnachtSig] INFO: Signature set successfully
```

If you don't see **"JS Loaded"**, the script isn't even
running — manifest/cache problem. If you see "JS Loaded" but
not "OnNewMessageCompose triggered", the event isn't reaching
the handler — registration problem (§8 rule about
`Office.actions.associate`).

### 6.4 Common symptoms and what they mean

**A. "Nothing happens when I click new email. No notification.
No signature."**

The handler isn't being invoked at all. Check in this order:

1. Is the add-in actually installed for this user?
   Admin centre → Integrated apps → Users tab.
2. Has Outlook been **fully restarted** since the install?
   The launch event registration happens at startup.
3. Did you bump the `?v=` cache-bust? Outlook may be running
   an old cached JS that has a bug.
4. Is `Office.actions.associate(…)` at top level of
   `autorun.js`? If someone re-wrapped the file in
   `Office.onReady`, the event handler registers too late on
   Classic — see §8.
5. Are you on Outlook Classic without WebView2? The JS may
   be hitting an IE11 syntax error and failing to parse.
   See §6.5 and §8.

**B. "I see the running… notification but no signature appears."**

The handler ran but either the Azure Function failed or the
HTML insertion failed.

1. Check which follow-up notification appears (§6.1).
2. If "No signature found" — the user isn't in Entra, or the
   Azure Function couldn't match them. Check §6.7.
3. If "Could not load signature" — network problem reaching
   the Azure Function. Try opening
   `https://connachtsignatures-bsbfakbbcjf6fnbb.westeurope-01.azurewebsites.net/api/signature?email=test@chgl.ie`
   in a browser. You should get JSON or 404, not a timeout.
4. If "Could not set signature: …" — the HTML was generated
   but Outlook refused to insert it. Look at the error
   message; common cause is an Outlook version that supports
   neither API.

**C. "Wrong hotel config is picked up."**

`getConfigForEmail` matches by **suffix**. Things to verify:

1. The user's "from" address really ends with the suffix
   you expect — shared mailboxes can sneak in a different
   domain.
2. The `HOTEL_CONFIG` key starts with `@` and matches the
   domain exactly. Typos here silently fall through to
   `default`.
3. Look at `INFO: Matched config website: …` in the log —
   that's what the code is using.

**D. "Banner not showing."**

1. Check `INFO: Employee found: …` then look for what
   `emp.banner` is. If it's empty, the Azure Function isn't
   returning a banner URL for this user.
2. Visit the banner URL in a browser. If it's 404, the
   asset isn't where the Azure Function thinks it is.
3. Some email clients block remote images by default —
   that's a *recipient* problem, not an add-in problem.

**E. "Wrong team override applied / no team override applied."**

1. The `teamCode` on the employee comes from the user's
   `faxNumber` field in Entra. Confirm that's set.
2. The lookup is case-insensitive and trimmed — but the
   key in `teamOverrides` should still be the same case as
   you intend.
3. Watch for `INFO: Team override applied — code: …` in the
   logs. If you don't see it, the code didn't match.

**F. "Old version is being served."**

Cache-bust value didn't get refreshed. Walk through §4.8 +
§5. Common gotchas:
- You bumped `?v=` in JS/HTML but didn't re-upload the
  manifest.
- You re-uploaded the manifest but Outlook wasn't restarted.
- GitHub Pages hasn't finished publishing yet — wait 2
  minutes and try again.
- The user is on a different "channel" of Microsoft 365 and
  the manifest hasn't propagated yet (this can take up to 24
  hours but usually under an hour).

### 6.5 Outlook Classic specifics

Outlook Classic on Windows picks one of these runtimes
depending on what's installed on the machine:

1. **Edge WebView2 runtime installed** → modern Chromium. ES6+
   works. This is the common case on Win10/11 with M365
   updated.
2. **Microsoft Edge (Chromium 79+) installed, no WebView2** →
   modern Chromium via Edge. Same situation.
3. **Neither of the above** → falls back to **Trident**, the
   IE11 engine. This is the breaking case.

Trident:
- Has no `async`/`await`. They are **parse-time syntax
  errors** — the whole file fails to load.
- Has no `fetch`.
- Has no `Object.assign`.
- Has no `String.prototype.endsWith` / `startsWith` /
  `includes`.
- Has no arrow functions, no template literals, no
  destructuring, no spread/rest, no `let`/`const` in the
  modern sense.
- `Promise` is polyfilled by `office.js`, so it's safe to
  use.
- `XMLHttpRequest` works.

Anyone touching `autorun.js` **must** keep it Trident-safe.
See §8 for the full list of "do not use".

To verify which engine a machine is using, look at the
`hostName` field in the `INFO: JS Loaded` area —
`Office.context.mailbox.diagnostics.hostName` will be
`Outlook` for Classic, but doesn't distinguish WebView2 vs
Trident directly. The easiest tell is whether modern syntax
works in the F12 console attached to the add-in process — if
F12 isn't available at all, you're almost certainly on
Trident.

### 6.6 The safety timeout

```js
var safetyTimeout = setTimeout(function () {
    logWarn("Safety timeout reached — completing event early");
    event.completed();
}, 55000);
```

Mobile Outlook kills any add-in that runs longer than 60
seconds. This 55-second guard ensures `event.completed()` is
called cleanly before the runtime force-kills the add-in.

If you see `WARN: Safety timeout reached`, something is
hanging — most likely:
- The Azure Function is unreachable or extremely slow.
- A `*Async` callback never fired (very unusual).

Investigate the Azure Function logs first.

### 6.7 The Azure Function — what to check

The Azure Function (separate Azure project) does:

1. Receives `GET …/api/signature?email=<address>`.
2. Looks up the user in Entra ID / Azure AD via Microsoft
   Graph.
3. Maps the user's email suffix to a banner URL.
4. Returns the employee JSON.

If the add-in shows "No signature found for your account" or
"Could not load signature", the issue is almost always
server-side. Quick checks:

1. **Direct GET in a browser**: open
   `https://connachtsignatures-bsbfakbbcjf6fnbb.westeurope-01.azurewebsites.net/api/signature?email=NDuggan@chgl.ie`
   (substitute the email of the user that's failing).
   - 200 + JSON → function works, problem is client-side.
   - 404 → user isn't in Entra (or has no usable record).
   - 5xx → function crashed. Check Azure Portal → the
     function app → Monitor / Log stream.
   - Timeout / connection refused → function is down.
2. **Azure Portal**: the function app is
   `connachtsignatures-bsbfakbbcjf6fnbb`. Look at its
   **Log stream** or **Application Insights** for the request
   that failed.
3. **Graph permissions**: if everything started failing at
   once, check that the service principal / app
   registration the function uses still has
   `User.Read.All` (or whichever scope it needs) granted
   tenant-wide.
4. **CORS**: the function must allow requests from
   `https://ConnachtHotel.github.io` and from any Outlook
   webview origin. If CORS is misconfigured, the XHR fails
   silently in the browser. Check Azure Portal → Function
   App → CORS.

Bear in mind the JSON contract the add-in expects:

```ts
{
  name:     string,           // "Nathan Duggan"
  title?:   string,           // "IT Intern"
  email:    string,           // "NDuggan@chgl.ie"   (REQUIRED — used for config lookup)
  phone?:   string,           // "+353 91 …"
  Mphone?:  string,           // "+353 87 …" (mobile)
  website?: string,           // "www.chgl.ie" — falls back to config.website
  address?: string,           // falls back to config.address
  banner?:  string,           // full URL to the banner image
  teamName?: string,
  teamCode?: string           // matched against HOTEL_CONFIG[*].teamOverrides
}
```

If the function changes the shape of this response, the
client will silently stop showing whichever field was renamed.

### 6.8 Runtime logging on the desktop client

For Outlook Classic on Windows you can enable a Microsoft
runtime log that captures launch-event activity (this is
Microsoft's own log, not our `logInfo` lines). Useful when
you can't attach DevTools.

1. Close Outlook.
2. Create the registry key:
   `HKCU\Software\Microsoft\Office\16.0\Outlook\Options\RuntimeLogging`
3. Add string value `Enabled` = `1`.
4. Add string value `LogFolderPath` =
   `C:\AddinLogs` (or any folder you have write access to).
5. Start Outlook. Open a new mail. Close it.
6. Look in that folder for log files — they'll show whether
   the launch event fired and whether the handler responded.

Remove these keys when finished. They generate a lot of
output.

---

## 7. Reference: every function in `autorun.js`

In order of appearance:

### `logInfo(msg)` / `logWarn(msg)` / `logError(msg)`
Thin wrappers around `console.log` / `warn` / `error` with
the `[ConnachtSig]` prefix. Use these instead of `console.*`
directly so logs are filterable.

### `notifyUser(type, message)`
Replaces the `connachtSigStatus` notification in the compose
window. `type` is `"error"` for the red error icon, anything
else for the blue info icon. Limited to ~150 characters in
practice.

### `getConfigForEmail(email)`
Walks the `HOTEL_CONFIG` object looking for a key the email
ends with. Returns the matched config object, or
`HOTEL_CONFIG["default"]` if nothing matches. Skips the
`"default"` key itself during iteration.

Implementation note: uses a manual `lastIndexOf` + length
check instead of `String.prototype.endsWith` because IE11
doesn't have `endsWith`. See §8.

### `getEmployeeData()`
Returns a `Promise` that resolves to the employee object (or
`null` if the Azure Function returned 404). Implemented with
`XMLHttpRequest`, not `fetch`, for IE11 compatibility.

Rejects on:
- HTTP status outside 200–299 (other than 404).
- Network errors (`xhr.onerror`).
- JSON parse errors.

### `buildSignatureHtml(emp, config)`
Pure function that returns the final HTML string. Takes the
merged employee object and the matched hotel config. No
side-effects. This is the one to edit when you want the
signature to look different.

### `onNewMessageCompose(event)`
The launch-event handler. Outlook calls this with an `event`
object that must eventually have `event.completed()` called on
it, otherwise the runtime considers the add-in hung and may
kill it.

Branches:
1. Neither `setSignatureAsync` nor `prependAsync` available →
   give up with a notification.
2. `setSignatureAsync` available (modern hosts) → use it.
3. `prependAsync` available but host isn't iOS → use it
   (older desktop fallback).
4. iOS without `setSignatureAsync` → exit cleanly without
   inserting anything (compose body insertion is restricted
   on iOS).

### `Office.actions.associate("onNewMessageCompose", onNewMessageCompose)`
The bottom line of the file. Registers our handler so the
runtime knows what to call when the event fires. **Must stay
at the top level of the file.** See §8.

---

## 8. Compatibility rules — DO NOT BREAK THESE

These are the rules that, if violated, will silently break
Outlook Classic for some or all users. They are not optional.

### Rule 1: Do not wrap the file in `Office.onReady`  --> this is false as of rolling back to a previous version on 02/06/26

`Office.actions.associate(…)` must run at script parse time.
If you wrap the file in `Office.onReady(function () { … })`,
the runtime fires the launch event before the `onReady`
callback runs, the handler isn't registered yet, the event is
dropped, and the add-in silently does nothing on Outlook
Classic. --> basically change it and it wont fire, but having it this way prevents outlook classic from picking it up. Bit annoying

This is Microsoft's documented requirement for event-based
add-ins.

### Rule 2: Do not use `async` or `await`

They are **parse-time syntax errors** on the IE11/Trident
webview. One stray `async function` and the entire file fails
to load. The add-in becomes a no-op.

Use `Promise` and `.then()`/`.catch()` (well, `["catch"]`)
chains instead.

### Rule 3: Do not use `fetch`

`fetch` doesn't exist on IE11. Use `XMLHttpRequest`.

### Rule 4: Do not use `Object.assign`

Not on IE11. Copy properties manually:

```js
for (var key in source) {
    if (Object.prototype.hasOwnProperty.call(source, key)) {
        target[key] = source[key];
    }
}
```

### Rule 5: Do not use ES6+ string methods

No `endsWith`, no `startsWith`, no `includes`. Use
`indexOf` / `lastIndexOf` / length checks.

### Rule 6: Do not use arrow functions, template literals, destructuring, spread/rest, classes, `let`, or `const`

All of these either parse-fail or behave unexpectedly on IE11.
Stick to ES5: `var`, `function`, plain object/array literals,
`+` for string concat.

(`const` is *partially* tolerated in IE11 but with quirks;
keep using `var` for consistency.)

### Rule 7: `Promise` is fine

`office.js` polyfills `Promise` globally on hosts that lack
it. Use it freely.

### Rule 8: `XMLHttpRequest` is fine

Available everywhere we care about.

### Rule 9: Keep the cache-bust in three places in sync

`manifest.xml` (two URLs) **and** `src/autorun.html` (script
tag). See §4.8.

### Rule 10: Don't bump `<Version>` past `1.0.2.x`

There's a comment at the bottom of `manifest.xml`:

> ManifestVersion 1.0.1.0 is the latest working version as of
> 10/03/2026, all version past 1.0.3.0 threw errors in
> validation.

Until that's investigated and resolved, leave `<Version>` at
`1.0.1.0`. The cache-bust (`?v=…`) is independent and can
move freely.

---

## 9. Glossary

| Term                          | Meaning                                                                                              |
|-------------------------------|------------------------------------------------------------------------------------------------------|
| **Outlook Classic**           | The legacy Win32 desktop Outlook. The one most users still have. Renders add-ins in a webview.       |
| **New Outlook**               | The modern WebView2-based desktop Outlook that Microsoft is gradually rolling out.                   |
| **OWA / Outlook on the Web**  | Outlook running in a browser at outlook.office.com.                                                  |
| **Edge WebView2**             | The Chromium-based webview Outlook uses on modern hosts. Supports ES6+ and modern web APIs.          |
| **Trident / IE11**            | The legacy webview Outlook Classic falls back to when WebView2 isn't installed. ES5 only.            |
| **office.js**                 | Microsoft's add-in library, loaded from a CDN. Provides `Office`, `Office.context`, `Office.actions`, the `Office.MailboxEnums`, and a `Promise` polyfill. |
| **Mailbox 1.10**              | The Office.js API requirement set we declare in the manifest. `setSignatureAsync` is in 1.10.        |
| **LaunchEvent**               | A manifest extension point that lets the add-in run automatically on certain Outlook events.         |
| **OnNewMessageCompose**       | The specific launch event we listen for: fired when a new mail compose window is opened.             |
| **Office.context.mailbox**    | The runtime object that gives access to the current user, the current item being composed, etc.     |
| **Office.context.mailbox.item.body** | The body of the message being composed. `.setSignatureAsync` / `.prependAsync` are methods on it. |
| **Office.actions.associate**  | Registers a JS function as the handler for a launch event by name (string used in the manifest).     |
| **faxNumber (Entra field)**   | Repurposed by this add-in to store team membership codes (e.g. `"GREEN"`). Returned as `teamCode`.   |
| **Cache-bust / ?v=**          | Query string suffix on the JS/HTML URLs that forces Outlook to fetch a new copy when changed.        |
| **Azure Function**            | Separate serverless backend that looks up users in Entra and returns their signature data.           |

---

## 10. Quick checklist for common changes

**I want to add a new hotel/domain:**
1. New entry in `HOTEL_CONFIG` in `autorun.js` (§4.1).
2. (If new banner) drop image in `assets/`, update Azure
   Function so it returns that URL for the new suffix.
3. Bump `?v=` in three places (§4.8).
4. Commit, push, wait for GH Pages, re-upload manifest if
   `?v=` bumped, restart Outlook.

**I want to change a colour for one hotel:**
1. Edit the relevant `style` block in `HOTEL_CONFIG` (§4.3).
2. Bump `?v=` (§4.8).
3. Push.

**I want to override one user's name/title without touching
Entra:**
1. Add an entry to `EMAIL_OVERRIDES` (§4.4).
2. Bump `?v=`, push.

**I want to give one team a different banner:**
1. Set their `faxNumber` in Entra to a team code.
2. Add a `teamOverrides[<CODE>]` entry under the right hotel
   in `HOTEL_CONFIG` (§4.2).
3. Bump `?v=`, push.

**I want to change the signature layout:**
1. Edit `buildSignatureHtml` in `autorun.js` (§4.6).
2. Stay table-based and inline-styled.
3. Bump `?v=`, push.

**A user reports the signature isn't appearing:**
1. Check what notification they see — §6.1.
2. Walk §6.4 starting at case A or B depending on
   notification.

**A new deploy isn't picking up:**
1. Did you bump `?v=` in all three places? §4.8.  -> cache buster if editing and new changes are not coming across on github pages
2. Did you re-upload the manifest to admin centre? §5.
3. Did the user fully restart Outlook? §5.
4. Did GH Pages finish publishing? §5.

**Outlook Classic specifically isn't working:**
1. Read §6.5.
2. Re-read §8 and audit `autorun.js` for any modern syntax.
3. Check that the file isn't wrapped in `Office.onReady`. -> stops working when it isnt wrapped...
4. Enable runtime logging (§6.8) on a failing machine.

---

*End of guide. If you change anything substantial in the
codebase, please update the relevant section here too.*

# Kalexius – Wrike ID Validator for Outlook
**Version 2.0.0**

An Outlook Web Add-in that validates a Wrike Task ID is present in the email subject before sending.

---

## How it works

| Scenario | Behaviour |
|---|---|
| Subject contains `[12345]` | ✅ Send allowed immediately |
| Subject has no Task ID — **1st** Send click | ⚠️ Blocked with warning banner |
| Subject has no Task ID — **2nd** Send click | ✅ Allowed (user confirmed intent) |

Users always need **at most 2 clicks** to send.

---

## File structure

```
/
├── manifest.xml       ← Upload to M365 Admin / Intune
├── commands.html      ← Loaded silently by Outlook (function file)
├── commands.js        ← Core validation logic
├── taskpane.html      ← Side panel UI
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    ├── icon-80.png
    └── icon-128.png
```

---

## Deployment

### Step 1 — Host the files on GitHub Pages

1. Create a new GitHub repository (e.g. `kal-wrike-validator`)
2. Push all files from this folder into the repo root
3. Go to **Settings → Pages → Source → main branch / root**
4. Note your Pages URL: `https://YOUR_USERNAME.github.io/kal-wrike-validator`

### Step 2 — Update the manifest

Open `manifest.xml` and replace **every** occurrence of:
```
YOUR_GITHUB_USERNAME  →  your actual GitHub username
YOUR_REPO_NAME        →  your repo name (e.g. kal-wrike-validator)
```

> Example: `https://jsmith.github.io/kal-wrike-validator/commands.html`

### Step 3 — Add icons

Place PNG icon files in the `assets/` folder:
- `icon-16.png` (16×16)
- `icon-32.png` (32×32)
- `icon-64.png` (64×64)
- `icon-80.png` (80×80)
- `icon-128.png` (128×128)

You can use any square PNG — the Kalexius logo or a simple checkmark works well.

### Step 4 — Deploy to Microsoft 365

#### Option A: Microsoft 365 Admin Centre (recommended for all users)

1. Go to [admin.microsoft.com](https://admin.microsoft.com)
2. Navigate to **Settings → Integrated apps → Upload custom apps**
3. Select **Office Add-in** and upload `manifest.xml`
4. Assign to users/groups as needed
5. Users see the add-in within ~24 hours (usually much faster)

#### Option B: Intune (for managed devices)

Use an Outlook policy or deploy the manifest via Exchange admin centre:
1. **Exchange Admin Centre → Organization → Add-ins → Add from file**
2. Upload `manifest.xml`
3. Set availability: Everyone / Specific users / Optional

#### Option C: Sideload for testing (individual)

1. Open Outlook on the web (OWA)
2. Click **Settings (gear) → View all Outlook settings → Mail → Customize actions**
3. Or go to: New email → three-dot menu → **Get Add-ins → My add-ins → Add a custom add-in → Add from file**
4. Upload `manifest.xml`

---

## Customisation

### Change the Task ID pattern

Edit `commands.js`, line with `var hasTaskId`:
```js
// Default: matches [123], [98765], etc.
var hasTaskId = /\[\d+\]/.test(subject);

// If you want alphanumeric IDs like [PROJ-123]:
var hasTaskId = /\[[A-Z0-9\-]+\]/i.test(subject);
```

### Change the warning message

Edit `commands.js`, the `errorMessage` string in the block section.

### Always block (never allow on 2nd click)

Remove the `else if (_userHasBeenWarned)` block and its `event.completed({ allowEvent: true })`.

---

## Troubleshooting

| Issue | Fix |
|---|---|
| Add-in not appearing | Wait up to 24h after M365 admin deployment. Try clearing Outlook cache. |
| Warning shown more than once | Ensure `commands.js` is the latest v2.0.0 version (check `_userHasBeenWarned` logic) |
| Icons not loading | Confirm PNG files exist in `assets/` and URLs in manifest are correct |
| Add-in not firing on Send | Confirm `Mailbox` requirement set `MinVersion="1.10"` and user's Outlook supports it |
| GitHub Pages 404 | Confirm Pages is enabled and files are committed to the correct branch/path |

---

## Requirements

- Microsoft 365 (Exchange Online)
- Outlook for Windows / Outlook on the Web (OWA)
- Mailbox requirement set 1.10+
- GitHub account (free) for hosting

> **Note:** The add-in does **not** work in Outlook for Mac with the legacy add-in engine. It works in new Outlook for Mac.

---

*Kalexius IT · 2025*

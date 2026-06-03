# LA COSTA HOTEL TIMELINE — Google Sheets Two-Way Sync
## Setup Guide

---

### What this gives you

- **One shared URL** — host the timeline HTML anywhere; every team member opens the same link
- **⬇ Load** — pulls the latest task data from the Google Sheet into the timeline
- **⬆ Save** — pushes all edits (dates, status, notes, new tasks) from the timeline back to the sheet
- **Direct sheet editing** — the team can also edit the Google Sheet directly (dates, notes, colors) and Load to see it in the timeline

---

## Step 1 — Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet
2. Name it something like **La Costa Hotel – Timeline Data**
3. Share it with your team members (Viewer or Editor access)

---

## Step 2 — Set up the Apps Script

1. In your Google Sheet, click **Extensions → Apps Script**
2. A new tab opens showing a code editor
3. **Delete** all the placeholder code (`function myFunction() {}`)
4. Open the file `Google_Apps_Script_Code.js` (in this same folder) and **copy all of its contents**
5. **Paste** it into the Apps Script editor
6. Click **Save** (the floppy disk icon or Ctrl+S)
7. Optional: run the `setupSheets` function once to pre-create the Tasks and Meta tabs with correct formatting
   - Click the function dropdown (shows `doGet`) → select `setupSheets`
   - Click **Run** → authorize when prompted

---

## Step 3 — Deploy as a Web App

1. Click **Deploy → New deployment**
2. Click the gear icon next to "Type" and select **Web app**
3. Fill in:
   - **Description:** `La Costa Timeline Sync v1`
   - **Execute as:** `Me`
   - **Who has access:** `Anyone` *(this allows the HTML timeline to connect without login)*
4. Click **Deploy**
5. Copy the **Web App URL** — it looks like:
   `https://script.google.com/macros/s/AKfycb.../exec`

> **Important:** Each time you change the Apps Script code, you must click **Deploy → Manage deployments → Edit → New version → Deploy** to update the live URL.

---

## Step 4 — Connect the Timeline

1. Open `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE.html` in a browser
2. In the top toolbar, find the **Google Sheets Sync** section (far right)
3. Click **⚙ Setup**
4. Paste your Web App URL into the field
5. Click **Test Connection** — you should see "Connection successful!"
6. Click **Save & Connect**

The status indicator will show a green dot and "Connected."

---

## Daily Workflow

### As editor / project manager:
1. Open the timeline HTML
2. Click **⬇ Load** to pull the latest data from the sheet
3. Drag bars to reschedule, add notes, update task status
4. Click **⬆ Save** to push changes back to the sheet
5. Team members can reload to see your updates

### As team member (view + light edit):
1. Open the timeline HTML (same file or hosted URL)
2. Click **⬇ Load** — see the current schedule
3. Make edits if needed → **⬆ Save**

### Direct sheet editing:
- Open the Google Sheet and edit the **Tasks** tab directly
- Change dates, notes, colors, or add rows
- Next time anyone clicks **⬇ Load** in the timeline, they see the updates

---

## Hosting the Timeline (optional)

For a permanent shareable link, host the HTML file on:

| Option | Cost | How |
|--------|------|-----|
| **Netlify Drop** | Free | Drag the HTML file to [app.netlify.com/drop](https://app.netlify.com/drop) — get a URL in 30 seconds |
| **GitHub Pages** | Free | Push the file to a repo, enable Pages in Settings |
| **Google Sites** | Free | Embed the HTML in a Google Site using "Embed" block |
| **Any web server** | Varies | Upload to your hosting, link directly |

Once hosted, share the URL with your team. Everyone loads and saves through the same Google Sheet.

---

## Keyboard Shortcut

**Ctrl+Shift+S** (or Cmd+Shift+S on Mac) = Save to Google Sheets

---

## Troubleshooting

**"Connection failed" when testing:**
- Make sure you selected "Anyone" (not "Anyone with Google account") for access
- Make sure you clicked Deploy → and got a URL ending in `/exec`
- Try opening the URL directly in a browser — you should see JSON data

**Save seems to work but sheet isn't updating:**
- The first save uses a fallback mode (no-CORS) which doesn't return a confirmation
- Check the Google Sheet directly — the data should be there within a few seconds
- If not, try re-deploying the Apps Script with a new version

**"No tasks found" after Loading:**
- The sheet's Tasks tab may be empty — use Save first to push the current timeline data into it

---

*Files in this folder:*
- `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE.html` — the timeline (now with sync built in)
- `Google_Apps_Script_Code.js` — paste this into Apps Script
- `SHEETS_SYNC_SETUP.md` — this guide

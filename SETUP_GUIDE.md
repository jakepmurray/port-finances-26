# 🥏 Ultimate Frisbee Finance Portal — Setup Guide

## What you'll have when done
- A **private Excel file** you manage on your computer (the source of truth)
- A **public website** your team can visit to see their finances
- Players log in with personal codes → see only their own balance
- Captains see full roster + attendance
- Coaches/Treasurers see everything

---

## Step 1 — Create a GitHub account
If you don't have one: https://github.com/signup (free)

---

## Step 2 — Create a new GitHub repository
1. Go to https://github.com/new
2. Name it something like `team-finances` (lowercase, no spaces)
3. Set it to **Public** ← required for GitHub Pages to work for free
4. Click **Create repository**

---

## Step 3 — Upload your files
In your new repo, click **Add file → Upload files** and upload:
- `index.html`
- `finance_data.xlsx` (rename your Excel file to exactly this)

Click **Commit changes**.

---

## Step 4 — Enable GitHub Pages
1. In your repo, click **Settings** (top menu)
2. In the left sidebar, click **Pages**
3. Under "Source", select **Deploy from a branch**
4. Branch: **main**, Folder: **/ (root)**
5. Click **Save**

After ~2 minutes, your site will be live at:
`https://YOUR_USERNAME.github.io/team-finances/`

---

## Step 5 — Test it
1. Visit your site URL
2. Enter one of the sample access codes from the Excel file:
   - `AJ2024` → Player view (Alex Johnson)
   - `JL2024` → Captain view (Jordan Lee)
   - `CK2024` → Coach/full view (Coach Kim)

---

## Step 6 — Set up your real roster

Open `finance_data.xlsx` and go to the **Roster** sheet:

1. **Delete the sample players** (rows 7–11)
2. **Add your real players** — one per row:
   - Full Name
   - Access Code (make something up — e.g. first initials + jersey # + year: `AJ7-2025`)
   - Jersey # (optional)
   - Dues Amount
   - Role: Player / Captain / Coach / Treasurer
   - Status: Active

3. **Share each player's code with them individually** (not the whole list!)
   - Text or email them their personal code
   - Never share the Excel file publicly

---

## Day-to-day workflow

### After every event:
1. Open `finance_data.xlsx`
2. **Events sheet** → add the event (name, date, category, cost)
3. **Attendance sheet** → mark Y/N/E for each player in the event's column
4. **Transactions sheet** → log any money paid or received
5. **Player Balances sheet** → update dues paid / upfront payments as needed
6. Save the file
7. Go to GitHub → **Add file → Upload files** → re-upload `finance_data.xlsx`
8. Website updates automatically within seconds!

### Or use GitHub Desktop (easier):
Download https://desktop.github.com — lets you drag-and-drop files and push with one click.

---

## Access Levels

| Role in Excel | What they see |
|---|---|
| Player | Their own balance, dues progress, event history |
| Captain | Full roster balances, events, attendance grid |
| Coach / Treasurer | Everything — dashboard, transactions, all balances |
| Manager | Same as Coach |

---


## Hiding the Access Codes sheet
The **Access Codes** tab is color-coded red so you don't forget it's there.

**Before sharing the file with players:**
1. Right-click the **Access Codes** tab at the bottom
2. Click **Hide**
3. The sheet is now invisible — your codes are safe

**To get it back:**
1. Right-click any sheet tab
2. Click **Unhide**
3. Select **Access Codes**

## Attendance Codes
- **1** = Attended
- **0** = Absent  
- **X** = Exempt (excused — cost can be waived)

---

## Tips
- The **Dashboard** sheet in Excel auto-calculates from all other sheets — don't edit it directly
- **Blue text** in Excel = your inputs. **Black text** = auto-calculated formulas (don't edit)
- If the website shows old data, try hard-refreshing: `Ctrl+Shift+R` (Windows) or `Cmd+Shift+R` (Mac)
- To change someone's access code, update it in the Roster sheet and re-upload the Excel file

---

## Troubleshooting

**"Could not load finance_data.xlsx"**
→ Make sure the file is named exactly `finance_data.xlsx` (case-sensitive) and is in the root of your repo (same folder as index.html)

**Code not found at login**
→ Codes are case-insensitive but must match exactly what's in column B of the Roster sheet. Check for extra spaces.

**Site not showing up**
→ GitHub Pages can take 2–5 minutes after first setup. Check Settings → Pages to confirm it's enabled.

**Data looks outdated**
→ Re-upload the Excel file to GitHub. Each upload triggers an update.

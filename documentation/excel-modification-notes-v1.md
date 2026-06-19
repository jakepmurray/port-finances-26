this directory contains an app that's intended to be used as a team frisbee finance website. it has the following goals:

1. be intuitive and user-friendly for everyone on the team
2. have different access levels for players and the treasurer (treasurer should get a helpful dashboard summary view along with their own personal page, everyone else should only see their own finances)
3. be easy to edit
4. be easy to pass of in the future

The excel file will live on google sheets. I'll download it into the project folder and upload it to github, where the html dashboard will be hosted on github sites. 

There are some changes I want to make off the bat. 

1. I want the roster page of the spreasheet to not have jersey # or date joined. Access codes should not exist on any sheet except the access codes sheet. Also remove the summary stats at the top of the page. 
2. Events sheet shouldn't have # expected on it and shouldn't have summary stats at top. it should have start and end date instead of just date
3. Attendance sheet should read off of roster and events to fill in rows and columns. total attendees and total cost is fine but the rest of the stuff at the top should go
4. transactions "linked events" should be a dropdown that directly links to the "events" page if possible
5. player balance tab shouldnt distinguish between events and dues. they should all be the same. dues owed shouldnt have to be manually filled. it should be rules-based with the attendance and events tab. 

I don't know the best way to address this yet, but i want an easy way for players to add items they paid for without editing the master sheet. Maybe a "pending approval" tab that is easy to migrate into the master transactions sheet? 

---

## Changes Implemented

### Session 1 — Structural Overhaul (`scripts/modify_xl.py`)

**Roster**
- Removed Jersey #, Date Joined, and Access Codes columns (codes kept only on Access Codes sheet)
- Removed summary stats block from the top of the sheet

**Events**
- Removed `# Expected` column and summary stats block
- Split `Date` into `Start Date` and `End Date` columns

**Attendance**
- Player rows auto-populate from Roster via INDEX/MATCH formulas
- Event columns auto-populate from Events via INDEX/MATCH formulas
- Removed header rows for Date and Category; kept Total Attendees and Total Cost rows

**Transactions**
- `Linked Event` column now uses a data validation dropdown sourced from Events!A5:A505

**Player Balances**
- Dues and event costs unified into a single `Total Charged` column (no separate columns)
- Dues formula-driven from Roster + Attendance (not manually entered)
- `Paid Upfront` pulled from Transactions via SUMIFS
- Net Balance and Status auto-calculated

**Pending Approval** (new tab)
- New sheet for player-submitted expenses; treasurer reviews and migrates entries to Transactions

---

### Session 2 — Formatting Fix (`scripts/fix_format.py`)

Applied the existing style system consistently across all modified areas:
- White-on-navy column headers
- Blue text for user input cells
- Black text for in-sheet formulas
- Green text for cross-sheet formula references
- Light blue fill for auto-calculated summary rows (Total Attendees, Total Cost, Cost per Person)
- Amber/light-yellow fill for input rows

---

### Session 3 — Rules Wiring (`scripts/wire_rules.py`)

**Attendance row 4 (Cost per Person)**
- Formula-driven from Events split method:
  - `Equal Split` → Actual Cost / active roster size (COUNTA of Roster)
  - `Attendance` → Actual Cost / number of attendees (COUNTIF of "1" in that event column)

**Transactions col J (Cost per Attendee)**
- New auto-calculated column using LET() to look up event cost and split method and return the per-person charge for each transaction row

---

### Session 4 — Bug Fixes (`scripts/fix_bugs.py`)

**Bug 1 — Total Charged not updating when attendance changes**
- Root cause: SUMPRODUCT in Player Balances compared attendance marks with `="1"` (strict string match), which fails when users type numeric `1`
- Fix: changed comparison to `=1` (numeric) across all 200 Player Balances rows

**Bug 2 — Events Actual Cost should auto-sum from Transactions**
- Root cause: Events col G (Actual Cost) was a hardcoded manual value, so it never updated when transactions were added
- Fix: replaced all 500 rows with `=IFERROR(SUMIFS(Transactions!$F..., type="Expense"), 0)`
- This completes the live chain: **Transactions → Events Actual Cost → Attendance Cost per Person → Player Balances Total Charged**
- Note: any event that previously had a hardcoded Actual Cost but no corresponding Expense transaction in Transactions will now show $0 until that transaction is logged


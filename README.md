# UKU SpiritHub

An Apps Script for Google Sheets designed to aid the spirit committee in understanding spirit data across a ultimate season.

It collects data from individual tournament results, maps teams to clubs, tracks issues and calculates statistics for spirit awards.

One hub covers one category (University, Club, ...) for one season.

## Folder layout

The hub reads results from Google Drive, laid out as:

```
<season folder>/                     the hub spreadsheet lives here
  <category folder>/                 e.g. University
    YYYYMMDD Tournament name/        e.g. 20251101 ELUXIR
      [regional folder/]
        <Any name> Spirit Results and Breakdown
```

- Each results file must sit in a folder named `YYYYMMDD Tournament name`, with a real date. The folder name gives the tournament's date and name.
- Each results file's name must contain `Spirit Results and Breakdown` (any capitalisation).
- Every Google Sheet in the category folder is listed on the Tournaments tab. Badly named files or folders are marked **ERROR** with the reason, and are not imported until renamed.
- The breakdown tab is found by its header row, not its name.
- International results are collected in the same layout as UK tournaments. Comments are ignored and monitoring for issue tracking is limited.

## Setting up a hub

### If using the online editor

1. Create a Google Sheet inside the season folder (e.g. 2025-26)
2. Open the Apps Script Editor from the top menu **Extensions > Apps Script**
3. Select **Settings** in the left hand menu and tick _"Show 'appsscript.json' manifest file in editor"_
4. Return to the **Editor** page from the left hand menu
5. Copy the contents of `appsscript.json` and `SpiritHub-<version>.js` into the editor, replacing what is there
6. Close the editor tab and refresh the Hub sheet
7. Run setup with the top menu **SpiritHub > SpiritHub Settings**. The first time, Google asks you to approve the hub's permissions:
   1. **Authorisation required** appears: click **Review permissions** (sometimes **OK**)
   2. Choose your Google account
   3. **Google hasn't verified this app** appears. This is expected, the hub is a private script, not a published app. Click **Advanced** (bottom left), then **Go to SpiritHub (unsafe)**
   4. The permission list appears. Tick **Select all** and click **Continue**. The hub asks for:
      - view and edit this spreadsheet only
      - see your Google Sheets, to read the results files (never to change them)
      - see information about your Google Drive files (names, folders and dates, not their contents), to find the results files
      - connect to an external service, to read all the results files at once from Google's own Drive and Sheets APIs
   5. The menu item does not carry on after approving: click **SpiritHub > SpiritHub Settings** again

   You only do this once per person. It happens again only when a new version of the hub asks for different permissions.
8. Wait until it stops saying Running Script
9. Select a Category of tournaments for this hub from the dropdown in cell **B6** (e.g. University, Club)
10. Check the other settings to see if everything looks about right
11. Gather tournaments, from the top menu click **SpiritHub > Refresh Tournaments**
12. Gather results, from the top menu click **SpiritHub > Refresh Results**
13. Look in the _Teams_ and _Clubs_ tabs, if you want to adjust a team's assigned club then type it in the _Club Override_ column
14. Automatically raise issues from results, from the top menu click **SpiritHub > Refresh Issues**
15. Generate statistics for determining Spirit Award Winner, from the top menu click **SpiritHub > Refresh Club Statistics**
16. If you want to adjust any of the settings for how the hub functions, from the top menu click **SpiritHub > SpiritHub Settings**, this will hide all tabs except settings, select any other menu items to hide settings and unhide all content tabs

### If using the clasp CLI

1. Create a Google Sheet inside the season folder (e.g. 2025-26)
2. Open the Apps Script Editor from the top menu **Extensions > Apps Script**
3. Select **Settings** in the left hand menu and copy the Script ID
4. Copy the script ID into `.clasp.json`
5. Run `just push` which will format, build and push the current code to the sheet
6. Continue from step 6 in the above instruction

## Using the hub

Everything runs from the **SpiritHub** menu.

| Menu item               | What it does                                                                                                                                    |
| ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| SpiritHub Settings      | Creates any missing tabs, updates the Category list, then shows only the Settings tab. Safe to rerun                                            |
| Refresh Tournaments     | Scans Drive and updates the Tournaments tab: new files are NEW, edited files are EDITED, removed files are MISSING, badly named files are ERROR |
| Refresh Results         | Imports every tournament marked NEW or REFRESH on the Tournaments tab, then updates Teams and Clubs. Does not add issues                        |
| Refresh Issues          | Adds any new issues from what is already in the spreadsheet, no data import                                                                     |
| Refresh Club Statistics | Recalculates the Club Statistics tab, no data import                                                                                            |

A normal update after uploading tournament results:

1. **Refresh Tournaments.**
2. On the Tournaments tab, tick **International** for international events, untick **Include** for anything to leave out, and change any EDITED or ERROR rows you want re-imported to **REFRESH**. Rename any files or folders marked ERROR, then run Refresh Tournaments again.
3. **Refresh Results.**
4. Check the Clubs tab for any duplicate looking clubs and set **Club Override** on the Teams tab to correct them. Clubs updates on edit.
5. **Refresh Issues.**
6. **Refresh Club Statistics** to calculate spirit award winners.

Each tab has an info row at the top saying what it is, which columns you can edit and how it refreshes.

Every refresh logs how long each step took to the Apps Script **Executions** log, to find slow steps.

## Tabs

| Tab             | What it's for                                                                                                                            |
| --------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| Issues          | A working tracker of automatically flagged spirit issues. Each row is a single issue, see [issues](#issues) for more detail              |
| Tournaments     | The tournaments found by the tool, their metadata and tournament level controls. Duplicate or blank tournament names are highlighted red |
| Results         | All spirit results imported into a single table, this is automatically generated and used as the data source                             |
| Teams           | Every team seen in results and its club. Based on Name Rules a club is suggested but can be overridden here if suggestion is wrong       |
| Clubs           | All the clubs present in the Teams table, used to check for duplicates                                                                   |
| Club Statistics | Statistics of club spirit across a season, used to determine spirit award winner                                                         |
| Club Report     | Pick a club to see its summary and charts of the scores it received                                                                      |
| Settings        | Hidden while refreshing. Sections for the Hub (Category), Spirit Award (minimum tournaments), Name Rules and Issue Rules                 |

## Issues

Each refresh runs every enabled check and adds issues that are not already on the tab. Existing rows are never changed, so sort, filter and edit freely.

| Check               | Policy                                           | Issue is for             |
| ------------------- | ------------------------------------------------ | ------------------------ |
| TOTAL-NO-COMMENT    | Total above 14 or below 6 needs a comment        | Scoring team             |
| CATEGORY-NO-COMMENT | A category scored 0 or 4 needs a comment         | Scoring team             |
| DANGEROUS-PLAY      | Comment mentions dangerous play                  | Receiving team           |
| CHEATING            | Comment mentions cheating                        | Receiving team           |
| HARASSMENT          | Comment mentions harassment or abuse             | Receiving team           |
| NOT-SUBMITTED       | Fewer scores given to an opponent than received  | Team that did not submit |
| LOW-SCORES          | 2 or more scores of 6 or below at a tournament   | Receiving team           |
| LOW-AVERAGE         | Average below 8 at a tournament                  | Receiving team           |
| MONITORING          | Club's teams average below 9 twice in the season | Club                     |
| MONITORING-BREACH   | Another average below 9 while on monitoring      | Club                     |
| MIN-CATEGORY        | Extra, not policy: received 0 in a category      | Receiving team           |
| SINGLE-LOW-SCORE    | Extra, not policy: received a total below 6      | Receiving team           |

Thresholds, keywords and on/off switches are in the Issue Rules section of the Settings tab. Rules added in later versions appear there on their own.

Differences from the published policy:

- Policy says "Consecutive" events are triggering, here we just look at multiple events in a season.
- MIN-CATEGORY and SINGLE-LOW-SCORE are not policy, but included for interest
- Comment checks and NON-SUBMITTED skip international tournaments. Comments not available and scoring teams are not necessarily UK teams.

## Club Statistics

Counts scores from included tournaments between different clubs (inter-club scores are excluded).

A club qualifies for the award once it has entered the minimum number of tournaments in the Spirit Award section of the Settings tab (default 3), and ranks are among qualifying clubs.

- **Mean, SD, median, min, max, % 6 or below, 95% CI** of the scores each club received. The CI uses the t distribution, as was used in R script `group.CI`.
- **Club Model / Team Model**: the R script mixed model, `lmer(total ~ club + (1 | scorer/tournament))`, fitted by REML in `src/Model.ts`. It adjusts each club's mean for scorers who give higher or lower scores than others. It is fitted twice, with scorers grouped by club (as in R) and by team. Checked against statsmodels' REML fit: means and scorer effects agree to 4d.p.
- **Average Given** and **Scorer Effect**: how a club scores others.

## Development

Needs Node.js and [just](https://github.com/casey/just).

```
just install     # install TypeScript, Apps Script types, dprint, clasp and ts-blank-space
just login       # log in to Google for clasp, once
just check       # format and type check
just push        # check, build and push to Apps Script
```

- `src/` contains all the source code in seperate files, written in TypeScript.
- `src/appsscript.json` is the Apps Script manifest, pushed as it is. It turns on the Drive and Sheets advanced services and lists the permissions (`oauthScopes`), so change them there rather than in the editor, where the next push would undo them. Because the permissions are listed, Apps Script no longer works them out from the code: new code that needs another permission fails with "You do not have permission to call …" until the scope is added here
- `just build` strips the types from all `.ts` files in `src/` and joins them into a single JavaScript file `dist/SpiritHub-<version>.js`
  - Types are removed rather than compiled (using `ts-blank-space`), so the code comes out as written, then dprint tidies the spacing
  - Comments are left out of the build, they live in `src/`. Only the version header and a `// ---- File.ts ----` marker per file are added
  - Only TypeScript that can be removed this way is allowed (`erasableSyntaxOnly` in `tsconfig.json`): no `enum`, `namespace` or constructor parameter properties
  - Files are plain scripts sharing one global scope, as in Apps Script: never add `import` or `export`
  - This is the code pushed by `just push`
  - The idea here is that someone can easily copy this script into a sheet without having to set up all the development environment. They can copy straight from a GitHub release version.
- The version is controlled by `HUB_VERSION` in `src/Consts.ts`. Each change is its own commit with its own version: patch for changes you would not notice in the sheet, minor for new features.

Conventions:

- Private functions are named `_name_` so they do not appear in the Apps Script run menu. `onOpen` and `onEdit` are reserved names
- Checks and calculations are pure functions with no Google calls, so they can be tested outside Apps Script
- Formula columns get a formula on every row (not an ARRAYFORMULA), so tabs can be sorted. `_fillFormulaColumns_` writes each formula column in one call, Sheets copies the first row's formula down
- Every tab has an info row, a header row, then data from row 3 (`INFO_ROW`, `HEADER_ROW`, `DATA_ROW`)

### Keeping it fast

Apps Script time goes almost entirely on calls to Google, not on the code itself, so:

- Drive is scanned one level of folders at a time, up to 40 folders per query (`_scanFolder_`)
- Results files are fetched all at once with `UrlFetchApp.fetchAll`: every modified time, then every file's contents (`_readResultsFiles_`). Requests Google asks to slow down are retried once
- Tables are written as one block of values, then one call per formula column
- Clubs are rebuilt from data already in memory, never by waiting for formulas and reading them back
- The Settings tab is read in one call and only repaired when a section or rule is missing
- Issues look scores up in an index built once per refresh, and the club model packs its data into arrays once per fit

Typical University season (30 tournaments, about 3,000 responses): Refresh Tournaments about 5s, Refresh Results about 10s.

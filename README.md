# UKU SpiritHub

An Apps Script for Google Sheets designed to aid the spirit committee in understanding spirit data across a ultimate season.

It collects data from individual tournament results, maps teams to clubs, tracks issues and calculates statistics for spirit awards.

One hub covers one season, or whatever dates are set on the Settings tab.

> **In progress:** international tournaments will be added by hand, not through the Schedule Manager, and publishing Results and Breakdown from the hub is still to come.

## Schedule Manager

Tournaments are read from the **Schedule List** tab of the UKU Schedule Manager, set on the Settings tab. The hub reads these columns from row 1, by name:

| Column       | Used for                                                  |
| ------------ | --------------------------------------------------------- |
| Event Name   | the tournament's name                                     |
| Start date   | the tournament's date, and whether it is within the dates |
| End date     | shown on the Tournaments tab                              |
| Schedule URL | the tournament's schedule                                 |
| SOTG Sheet   | the tournament's Form Responses, which identifies it      |

Rows without a Schedule URL or SOTG Sheet, or starting outside the Start Date and End Date on the Settings tab, are left out.

On the Tournaments tab grey columns are filled on refresh and white columns are yours to edit:

| Column              | What it's for                                                                                                                                                                                      |
| ------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Event Name Override | A name to use instead of the Schedule Manager's. Event Name is the name used on every other tab                                                                                                    |
| Schedule Status     | Whether the schedule needs importing, set REFRESH to import it again                                                                                                                               |
| Spirit Form Status  | Whether the Form Responses need importing, set REFRESH to import them again                                                                                                                        |
| Include Results     | Import this tournament's results. Unticked for new tournaments                                                                                                                                     |
| International       | Only the receiving teams are UK teams. The scoring teams are left off the Teams and Clubs tabs, and NO-SCORES and UNDER-SUBMITTED are not checked, as only one or two scores per game are expected |

Each file's status is **NEW** (not imported yet), **EDITED** (changed since it was imported), **REFRESH**, **ERROR** (cannot be read, see Message), **MISSING** (no longer on the Schedule List) or **OK**. Event Status is **NOT INCLUDED** when Include Results is unticked, otherwise the first of MISSING, ERROR, REFRESH, EDITED, NEW that either file has, otherwise OK.

A tournament that leaves the Schedule List is removed, unless something was imported from it, then it stays as MISSING.

Refresh Results reads the **Gamelist** tab of each included tournament's schedule onto the Games tab. Columns are read by position, as both team columns are headed "Team" and the score columns have no header: A Game, C Day, D Time, E Pitch, F and J the teams, G and I the scores. Games without both teams are left out, and 1-0 results are kept with Played unticked.

Refresh Results also reads each included tournament's **Form Responses**, finding two tabs by their header row rather than their name:

- the responses tab (today `Form responses 2`): every tab with the form's headers. There must be exactly one. Its rows go onto the Results tab. Rows with a missing team or a score that is not 0–4 are left out and listed in Message, placeholder rows with "Your Team Name" as the team are skipped
- the Deleted forms tab: the one with the form's headers and "Reason for deletion". Optional. Its rows go onto the Deleted Forms tab as they are, as a record of what the chair moved out. Some are scores taken out, others the original of a response that was edited and is still on Results, so they are never used in statistics

Then every response is **matched to its game**, for tournaments whose schedule has been imported and is not ERROR or MISSING. For each scoring team → receiving team at a tournament:

- as many forms count as the two teams played games: forms with a comment first, then the latest. Counted forms go to the games in timestamp order, the earliest to the earliest game
- forms beyond that are **Duplicate**, a team scoring itself is **Self score**, and a form with no game between the two teams is **No such game**. None of these count
- without a usable schedule every form is **No schedule** and counts as it is

Results shows the Game, a Counted tick and the Match. Only counted forms go into Counts for Award, Club Statistics and the Club Report. On Games, Team 1 Submitted and Team 2 Submitted are ticked when that team has a counted form for the game, and a 1-0 game is Played if anyone submitted for it

International tournaments will be added by hand, not through the Schedule Manager. Not built yet.

## Setting up a hub

### If using the online editor

1. Create a Google Sheet
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
      - see information about your Google Drive files (names and dates, not their contents), to spot files that changed since they were imported
      - connect to an external service, to read all the results files at once from Google's own Drive and Sheets APIs
   5. The menu item does not carry on after approving: click **SpiritHub > SpiritHub Settings** again

   You only do this once per person. It happens again only when a new version of the hub asks for different permissions.
8. Wait until it stops saying Running Script
9. In the Hub section, paste the Schedule Manager link and set the Start Date and End Date of the tournaments this hub covers
10. Check the other settings to see if everything looks about right
11. Gather tournaments, from the top menu click **SpiritHub > Refresh Tournaments**
12. Gather results, from the top menu click **SpiritHub > Refresh Results**
13. Look in the _Teams_ and _Clubs_ tabs, if you want to adjust a team's assigned club then type it in the _Club Override_ column
14. Automatically raise issues from results, from the top menu click **SpiritHub > Refresh Issues**
15. Generate statistics for determining Spirit Award Winner, from the top menu click **SpiritHub > Refresh Club Statistics**
16. If you want to adjust any of the settings for how the hub functions, from the top menu click **SpiritHub > SpiritHub Settings**, this will hide all tabs except settings, select any other menu items to hide settings and unhide all content tabs

### If using the clasp CLI

1. Create a Google Sheet
2. Open the Apps Script Editor from the top menu **Extensions > Apps Script**
3. Select **Settings** in the left hand menu and copy the Script ID
4. Copy the script ID into `.clasp.json`
5. Run `just push` which will format, build and push the current code to the sheet
6. Continue from step 6 in the above instruction

## Using the hub

Everything runs from the **SpiritHub** menu.

| Menu item               | What it does                                                                                                                                                                                    |
| ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| SpiritHub Settings      | Creates any missing tabs, then shows only the Settings tab. Safe to rerun                                                                                                                       |
| Refresh Tournaments     | Reads the Schedule List of the Schedule Manager and updates the Tournaments tab, see [Schedule Manager](#schedule-manager)                                                                      |
| Refresh Results         | Imports every included tournament's files marked NEW, EDITED or REFRESH: schedules onto Games, Form Responses onto Results and Deleted Forms. Then updates Teams and Clubs. Does not add issues |
| Refresh Issues          | Adds any new issues from what is already in the spreadsheet, no data import                                                                                                                     |
| Refresh Club Statistics | Recalculates the Club Statistics tab, no data import                                                                                                                                            |

A normal update after uploading tournament results:

1. **Refresh Tournaments.**
2. On the Tournaments tab, tick **Include Results** for this hub's tournaments and **International** for international events, and set any file you want re-imported to **REFRESH**.
3. **Refresh Results.**
4. Check the Clubs tab for any duplicate looking clubs and set **Club Override** on the Teams tab to correct them. Clubs updates on edit.
5. **Refresh Issues.**
6. **Refresh Club Statistics** to calculate spirit award winners.

Each tab has an info row at the top saying what it is, which columns you can edit and how it refreshes.

Every refresh logs how long each step took to the Apps Script **Executions** log, to find slow steps.

## Tabs

| Tab             | What it's for                                                                                                                                                                                                  |
| --------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Issues          | A working tracker of automatically flagged spirit issues. Each row is one team in one game, or one team at one tournament, see [issues](#issues) for more detail                                               |
| Tournaments     | The tournaments found by the tool, their metadata and tournament level controls. Duplicate or blank tournament names are highlighted red                                                                       |
| Games           | Every game on the Gamelist of each included schedule, one import per schedule even when divisions share it. Played is unticked for 1-0 results (forfeits and cancelled games)                                  |
| Results         | All spirit results imported into a single table, this is automatically generated and used as the data source                                                                                                   |
| Deleted Forms   | Every row of each included tournament's Deleted forms tab, with the chair's reason. A record only, never used in statistics                                                                                    |
| Teams           | Every team seen in results and its club. Based on Name Rules a club is suggested but can be overridden here if suggestion is wrong                                                                             |
| Clubs           | All the clubs present in the Teams table, used to check for duplicates                                                                                                                                         |
| Club Statistics | Statistics of club spirit across a season, used to determine spirit award winner                                                                                                                               |
| Club Report     | Pick a club, then one of its teams, to see the totals and category scores each received, totals split by tournament. Teams ignore trailing brackets, so Birmingham 1 (WW4) and Birmingham 1 (WW2) are one team |
| Club Contacts   | Email address for each club. New clubs are added at the bottom, the script never changes existing rows                                                                                                         |
| Email Templates | Hidden while refreshing. One subject and body per issue check, plus MULTIPLE for issues with several checks, with placeholders filled from the issue                                                           |
| Settings        | Hidden while refreshing. Sections for the Hub (Schedule Manager and dates), Spirit Award (minimum tournaments, international), Name Rules and Issue Rules                                                      |

## Issues

Issues are about one team, at one of three levels:

| Level      | One issue per        | From                                      |
| ---------- | -------------------- | ----------------------------------------- |
| Play       | team in a game       | the scores the team received in that game |
| Scoring    | team in a game       | the form the team gave for that game      |
| Tournament | team at a tournament | the whole tournament                      |

Everything found for the same issue goes into it: the Checks column lists every check, and Details has a section for each, then the comments. When a form has no game (a duplicate, a deleted form, or a tournament without a schedule) the issue is for the opponent instead of the game.

| Check               | Level      | Raised when                                                                                                                           | Source |
| ------------------- | ---------- | ------------------------------------------------------------------------------------------------------------------------------------- | ------ |
| LOW-SCORE           | Play       | received a total of 6 or below. Details list the team's other low scores at the tournament and say when policy's 2 or more is reached | policy |
| DANGEROUS-PLAY      | Play       | comment received mentions dangerous play                                                                                              | policy |
| CHEATING            | Play       | comment received mentions cheating                                                                                                    | extra  |
| HARASSMENT          | Play       | comment received mentions harassment or abuse                                                                                         | extra  |
| MIN-CATEGORY        | Play       | received 0 in a category                                                                                                              | extra  |
| TOTAL-NO-COMMENT    | Scoring    | gave a total above 14 or below 6 without a comment                                                                                    | policy |
| CATEGORY-NO-COMMENT | Scoring    | gave a 0 or 4 in a category without a comment                                                                                         | policy |
| DELETED-ROW         | Scoring    | a form is on the Deleted Forms tab, the details are the chair's reason                                                                | extra  |
| SELF-SCORE          | Scoring    | a team scored itself                                                                                                                  | extra  |
| NO-SUCH-GAME        | Scoring    | a form for two teams with no game on the schedule                                                                                     | extra  |
| DUPLICATE           | Scoring    | more forms to an opponent than games, every form is listed                                                                            | extra  |
| NO-SCORES           | Tournament | a team submitted no scores                                                                                                            | policy |
| UNDER-SUBMITTED     | Tournament | a team submitted for fewer than 50% of the games it played                                                                            | policy |
| IMMEDIATE-AVERAGE   | Tournament | a team averaged below 8. Details list the club's earlier averages below 9 this season                                                 | policy |
| MONITORING-AVERAGE  | Tournament | a team averaged below 9, but not below 8. Same details                                                                                | policy |

Thresholds, keywords and on/off switches are in the Issue Rules section of the Settings tab. Rules added in later versions appear there on their own, rules from earlier versions can be deleted.

- Only tournaments ticked **Include Results** are checked
- Play checks and averages only use counted forms, so a duplicate is never held against the opponent
- Submission checks need the tournament's schedule. A team's games are the played games on the schedule, so a 1-0 is not expected to have scores unless someone submitted for it
- Comment checks and submission checks skip international tournaments, comments are not available and scoring teams are not necessarily UK teams
- "Earlier" averages are from tournaments with an earlier start date, among included tournaments

Refresh Issues puts the tab back in order every time: club, then tournament (date, then name), team, level, checks and opponent, so a club's issues sit together. Filters are kept through every refresh, on this tab and every other.

Each refresh adds issues that are not already on the tab. When it finds a new check for an existing issue, the check is added to Checks and its section to Details, and a CLOSED issue goes back to NEW. Status, Committee Member and Notes are otherwise never changed, so filter and edit freely.

## Emails

The Issues tab drafts an email for each issue, ready to copy into whatever mail you use. SpiritHub never sends anything.

| Column        | What it holds                                                                                                                                                                                                                          |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Email Address | the club's address from Club Contacts, a formula so it is always current. "No email for club" when there is none, "Could not be determined" when DELETED-ROW is the only check, as a form can be deleted for something either team did |
| Email Subject | filled from the issue's template while the issue is NEW                                                                                                                                                                                |
| Email Body    | filled from the issue's template while the issue is NEW                                                                                                                                                                                |

1. Fill in each club's address on the **Club Contacts** tab (several addresses separated by commas)
2. Check the wording on the **Email Templates** tab (open it with **SpiritHub > SpiritHub Settings**). An issue with one check uses that check's template, one with several uses **MULTIPLE**, one with none (e.g. added by hand) uses **MANUAL**. Only NO-SCORES and UNDER-SUBMITTED are enabled to begin with, and checks that are not policy have no template
3. **Refresh Issues** fills Email Subject and Email Body for NEW issues whose template is enabled. Editing the Email Templates tab fills them again straight away
4. Copy the address, subject and body into your mail, check it and send it
5. Set the issue's Status to IN PROGRESS or CLOSED. Subject and body are then kept as they are, a record of what was sent

An issue added by hand gets an Email Address by copying the formula down from the row above.

Templates use placeholders such as `{club_name}`, `{team_name}`, `{tournament_name}`, `{game}`, `{opponent}`, `{checks}`, `{missing_count}` and `{missing_opponents}`; the full list is in the tab's info row. An unknown placeholder is reported by Refresh Issues, or when the template is edited, and nothing is filled until it is fixed.

## Club Statistics

Counts scores from tournaments ticked **Include Results**, between different clubs (inter-club scores are excluded). International tournaments count unless **Include International** is unticked in the Spirit Award section of the Settings tab.

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

- The Schedule List is read in one call, and every file's modified time is fetched at once with `UrlFetchApp.fetchAll` (`_readFileVersions_`)
- Results files are fetched all at once with `UrlFetchApp.fetchAll`: every modified time, then every file's contents (`_readResultsFiles_`). Requests Google asks to slow down are retried once
- Tables are written as one block of values, then one call per formula column
- Clubs are rebuilt from data already in memory, never by waiting for formulas and reading them back
- The Settings tab is read in one call and only repaired when a section or rule is missing
- Issues look scores up in an index built once per refresh, and the club model packs its data into arrays once per fit

Typical University season (30 tournaments, about 3,000 responses) with the old Drive scan: Refresh Tournaments about 5s, Refresh Results about 10s.

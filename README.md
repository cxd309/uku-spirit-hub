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
- The breakdown tab is found by its header row, not its name.
- International results are collected in the same layout as UK tournaments. Comments are ignored and monitoring for issue tracking is limited.

## Using the hub

Everything runs from the **SpiritHub** menu.

| Menu item               | What it does                                                                                                        |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------- |
| Setup SpiritHub         | Sets up the hub, creates all tables, safe to rerun                                                                  |
| Refresh Tournaments     | Scans Drive and updates the Tournaments tab: new files are NEW, edited files are EDITED, removed files are MISSING. |
| Refresh Results         | Imports every tournament marked NEW or REFRESH in Tournaments tab, updates Teams and Clubs and adds new issues.     |
| Refresh Issues          | Refreshes issues, no data import.                                                                                   |
| Refresh Club Statistics | Recalculates the Club Statistics tab, no data import.                                                               |

A normal update after uploading tournament results:

1. **Refresh Tournaments.**
2. On the Tournaments tab, tick **International** for international events, untick **Include** for anything to leave out, and change any EDITED or ERROR rows you want re-imported to **REFRESH**.
3. **Refresh Results.**
4. Check clubs tab for any duplicate looking clubs, set **Club Override** in teams tab to correct, Clubs table updates on edit
5. **Refresh Issues** if you manually changed anything.
6. **Refresh Club Statistics** to calculate spirit award winners.

Each tab has an info row at the top saying what it is, which columns you can edit and how it refreshes.

## Tabs

## Tabs

| Tab             | What it's for                                                                                                                                       |
| --------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| Issues          | A working tracker of automatically flagged spirit issues. Each row is a single issue, see [issues](#issues) for more detail                         |
| Tournaments     | The tournaments found by the tool, their metadata and tournament level controls                                                                     |
| Results         | All spirit results imported into a single table, this is automatically generated and used as the data source                                        |
| Teams           | Every team seen in results and its club. Based on Name Rules a club is suggested but can be overridden here if suggestion is wrong                  |
| Clubs           | All the clubs present in the Teams table, used to check for duplicates                                                                              |
| Club Statistics | Statistics of club spirit across a season, used to determine spirit award winner                                                                    |
| Name Rules      | The rules used by the hub to automatically suggest club names for a team. e.g. remove "2" or "(W7)" from end of name. Can be enabled for the season |
| Issue Rules     | The rules used to highlight spirit issues and if the are enabled                                                                                    |
| Config          | The hub's settings: which category folder it reads and how many tournaments a club needs to qualify for the spirit award                            |

## Issues

Each refresh runs every enabled check and adds issues that are not already on the tab. Existing rows are never changed, so sort, filter and edit freely.

| Check               | Policy                                           | Issue is for             |
| ------------------- | ------------------------------------------------ | ------------------------ |
| TOTAL-NO-COMMENT    | Total above 14 or below 6 needs a comment        | Scoring team             |
| CATEGORY-NO-COMMENT | A category scored 0 or 4 needs a comment         | Scoring team             |
| DANGEROUS-PLAY      | Comment mentions dangerous play                  | Receiving team           |
| NOT-SUBMITTED       | Fewer scores given to an opponent than received  | Team that did not submit |
| LOW-SCORES          | 2 or more scores of 6 or below at a tournament   | Receiving team           |
| LOW-AVERAGE         | Average below 8 at a tournament                  | Receiving team           |
| MONITORING          | Club's teams average below 9 twice in the season | Club                     |
| MONITORING-BREACH   | Another average below 9 while on monitoring      | Club                     |
| MIN-CATEGORY        | Extra, not policy: received 0 in a category      | Receiving team           |
| SINGLE-LOW-SCORE    | Extra, not policy: received a total below 6      | Receiving team           |

Thresholds and on/off switches are on the Issue Rules tab.

Differences from the published policy:

- Policy says "Consecutive" events are triggering, here we just look at multiple events in a season.
- MIN-CATEGORY and SINGLE-LOW-SCORE are not policy, but included for interest
- Comment checks and NON-SUBMITTED skip international tournaments. Comments not available and scoring teams are not necessarily UK teams.

## Club Statistics

Counts scores from included tournaments between different clubs (inter-club scores are excluded).

A club qualifies for the award once it has entered the Config tab's minimum number of tournaments (default 3), and ranks are among qualifying clubs.

- **Mean, SD, median, min, max, % 6 or below, 95% CI** of the scores each club received. The CI uses the t distribution, as was used in R script `group.CI`.
- **Club Model / Team Model**: the R script mixed model, `lmer(total ~ club + (1 | scorer/tournament))`, fitted by REML in `src/Model.js`. It adjusts each club's mean for scorers who give higher or lower scores than others. It is fitted twice, with scorers grouped by club (as in R) and by team. Checked against statsmodels' REML fit: means and scorer effects agree to 4d.p.
- **Average Given** and **Scorer Effect**: how a club scores others.

## Development

Needs Node.js and [just](https://github.com/casey/just).

```
just install     # install TypeScript, Apps Script types, dprint and clasp
just login       # log in to Google for clasp, once
just check       # format and type check
just push        # check, build and push to Apps Script
```

- `src/` contains all the source code in seperate files. Everything uses JSDoc for typehinting.
- `just build` joins all `.js` files in `src/` into a single file `dist/SpiritHub-<version>.js`
  - This is the code pushed by `just push`
  - The idea here is that someone can easily copy this script into a sheet without having to set up all the development environment. They can copy straight from a GitHub release version.
- The version is controlled by `HUB_VERSION` in `src/Consts.js`.

Conventions:

- Private functions are named `_name_` so they do not appear in the Apps Script run menu. `onOpen` and `onEdit` are reserved names
- Checks and calculations are pure functions with no Google calls, so they can be tested outside Apps Script
- Formula columns get a formula on every row (not an ARRAYFORMULA), so tabs can be sorted
- Every tab has an info row, a header row, then data from row 3 (`INFO_ROW`, `HEADER_ROW`, `DATA_ROW`)

## Legacy

`legacy/` keeps the old merge script (`code.gs`), the R award script (`SOTG.award.code.R`) and the published policy for reference. Data files are git-ignored.

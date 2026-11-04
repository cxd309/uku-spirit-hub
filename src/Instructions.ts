/**
 * name of the Instructions tab
 */
const INSTRUCTIONS_SHEET = "Instructions";

/**
 * one line of the Instructions tab: a heading or a line of text
 */
type InstructionLine = { heading: string } | { text: string };

/**
 * what the Instructions tab says, in order
 * a function so every tab name and setting constant has loaded before it is read
 *
 * @returns the lines
 */
function _instructionLines_(): InstructionLine[] {
  const steps = [
    "Make a copy of each tournament's Form Responses and add (BEFORE EDITS) to its name, for safety",
    "In the Schedule Manager, on each tournament's row, run SOTG After > Format Scores and Delete After Copying. "
    + "It creates the Deleted forms tab and moves forms that need a comment but have none",
    "In each Form Responses, go through the comments and unusually high scores. "
    + "To redact a comment, copy its row to Deleted forms first, then edit the comment in Form responses 2. "
    + "To take a form out, move its row to Deleted forms with a reason",
    `SpiritHub > Refresh Tournaments, then on the ${EVENTS_SHEET} tab tick Include Results for the new tournaments, `
    + "and International for international ones",
    "SpiritHub > Refresh Results and Issues",
    `Work through the ${ISSUES_SHEET} tab: copy Email Address, Email Subject and Email Body into an email and send it, `
    + "then set Status to IN PROGRESS or CLOSED",
    "Create the tournament's SOTG folder in Drive and put its link in the Schedule Manager's SOTG Folder column",
    "SpiritHub > Refresh Tournaments, so the hub sees the folder: Publish Status becomes NOT PUBLISHED",
    "SpiritHub > Publish Results",
    `Open the Results and Breakdown URL from the ${EVENTS_SHEET} tab, then File > Share > Publish to web`,
    "Copy the Results and Breakdown URL into the Schedule Manager's SOTG Results and Breakdown column",
  ];
  return [
    { heading: `SpiritHub v${HUB_VERSION}` },
    { text: "What to do after a tournament weekend, and what the hub does with what you do" },
    { text: "" },
    { heading: "After a tournament weekend" },
    ...steps.map((step, i) => ({ text: `${i + 1}. ${step}` })),
    { text: "" },
    { heading: "Reasons in Deleted forms" },
    { text: "Each row in Deleted forms raises an issue, depending on its reason:" },
    { text: "- a reason mentioning \"duplicate\": DUPLICATE, with the form that counts for that game" },
    {
      text: "- \"Total>14 & Comment required\", \"Total<6 & Comment required\" or \"Comment required\", "
        + "as step 2 writes them: TOTAL-NO-COMMENT or CATEGORY-NO-COMMENT",
    },
    { text: "- a copy of a row whose comment was then edited in Form responses 2: REDACTED-COMMENT" },
    { text: "- anything else: DELETED-ROW, with your reason" },
    { text: "" },
    { heading: "Good to know" },
    {
      text: "- SpiritHub only reads the Form Responses, the schedules and the Schedule Manager. "
        + "The only files it writes are the Results and Breakdown files it publishes",
    },
    {
      text: `- If a row on the ${EVENTS_SHEET} tab says ERROR, its Message says why. `
        + "Set Schedule Status or Spirit Form Status to REFRESH to import that file again",
    },
    {
      text: "- Each tab's first row says what it is, which columns you can edit and how it refreshes. "
        + "Grey columns are filled by the hub, white ones are yours",
    },
    {
      text: `- Settings and Email Templates are opened with SpiritHub > SpiritHub Settings, and hidden by any refresh`,
    },
  ];
}

/**
 * write the Instructions tab, creating it first if needed, and put it first
 * rewritten every time, so it always matches this version of the hub
 *
 * @returns the Instructions tab
 */
function _writeInstructions_(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INSTRUCTIONS_SHEET) ?? ss.insertSheet(INSTRUCTIONS_SHEET, 0);
  const lines = _instructionLines_();
  sheet.clear();
  _fitSheet_(sheet, lines.length, 1);
  sheet.setColumnWidth(1, 900);
  sheet
    .getRange(1, 1, lines.length, 1)
    .setValues(lines.map((line) => ["heading" in line ? line.heading : line.text]))
    .setWrap(true)
    .setVerticalAlignment("top");
  lines.forEach((line, i) => {
    if ("heading" in line) sheet.getRange(i + 1, 1).setFontWeight("bold").setFontSize(i === 0 ? 14 : 12);
  });
  sheet.showSheet();
  if (sheet.getIndex() !== 1) {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }
  return sheet;
}

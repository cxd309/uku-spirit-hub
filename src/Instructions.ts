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
 * one string per line, headings are their own lines
 *
 * @returns the lines
 */
function _instructionLines_(): InstructionLine[] {
  const text = (line: string): InstructionLine => ({ text: line });
  const heading = (line: string): InstructionLine => ({ heading: line });
  const blank = text("");
  return [
    heading(`SpiritHub v${HUB_VERSION}`),
    text("What to do after a tournament weekend"),
    blank,
    heading("General Guidance"),
    text("White cells are for you to edit, grey cells are computer generated and should not be changed by a human."),
    text("Any questions reach out to Chris Davis"),
    blank,
    heading("After a tournament weekend"),
    text("1. Make a copy of each tournament's Form Responses and add (BEFORE EDITS) to its name, for safety"),
    text(
      "2. In the Schedule Manager, on each tournament's row, run SOTG After > Format Scores and Delete After Copying. "
        + "It creates the Deleted forms tab and moves responses with very high or low scores that required a comment "
        + "but have no comment",
    ),
    text(
      "3. In each Form Responses, go through the comments and unusually high scores. To redact a comment, copy the row "
        + "to Deleted forms first and add a rationale, then edit the comment in Form responses 2. To remove a response "
        + "entirely, move its row to Deleted forms with a reason",
    ),
    text(
      "4. Create a SOTG folder that published results will go into. "
        + "Add a link to the folder in Schedule Manager \"SOTG Folder\" column",
    ),
    text(
      "5. Return to SpiritHub. Run SpiritHub > Refresh Tournaments. "
        + "This will update the tournaments tab based on information in the schedule manager.",
    ),
    text("5a. On the tournaments tab tick \"Include Results\" for new tournaments that now have spirit data."),
    text(
      "6. SpiritHub > Refresh Results and Issues. This will gather information from schedules and spirit form "
        + "responses and try to automatically raise issues based on the Settings tab. It will also raise all of the "
        + "manual issues from Deleted forms tab",
    ),
    text(
      "6a. If you determine an issue does not require further action, set status to CLOSED, add your name and a note "
        + "with your rationale for closing the issue",
    ),
    text(
      "6b. If you determine an issue requires further action, decide a lead committee member for the issue and put in "
        + "their name. An email template and contact is provided as a starting point and once sent a note should be "
        + "added with a timestamp and action taken. Further updates to the issue can be included here so that progress "
        + "can be tracked.",
    ),
    text(
      "7. SpiritHub > Publish Results. For all events that are ready to publish this will create the "
        + "\"Results and Breakdown\" sheet in the SOTG folder ready to be published to the web.",
    ),
    text("8. Open the Results and Breakdown URL from the Tournaments tab, then File > Share > Publish to web"),
    text("9. Copy the Results and Breakdown URL into the Schedule Manager's SOTG Results and Breakdown column"),
    blank,
    heading("Reasons in Deleted forms tab"),
    text("Each row in Deleted forms raises an issue, depending on its reason:"),
    text("- If the reason mentions \"duplicate\" it is raised as a DUPLICATE score in the issues for that game."),
    text(
      "- If the comment is exactly: \"Total>14 & Comment required\", \"Total<6 & Comment required\" or "
        + "\"Comment required\" this is raised as a no-comment issue, these match the text added by Schedule Manager "
        + "SOTG After > Format Scores and Delete After Copying",
    ),
    text(
      "- If a \"Deleted forms\" row matches one in \"Form responses 2\" but with a different comment then it is raised "
        + "as a redacted comment issue",
    ),
    text("- anything else gets a generic DELETED-ROW issue with the rationale provided"),
    blank,
    heading("Good to know"),
    text(
      "- SpiritHub only reads: Form Responses, Schedules and the Schedule Manager. "
        + "It only writes the Results and Breakdown files to publishes",
    ),
    text("- If a row on the Tournaments tab says ERROR, it will have a Message explaining why"),
    text("- A tournament can be set to REFRESH which will refetch the data from the schedule or form responses"),
    text(
      "- Each tab's first row says what it is, which columns you can edit and how it refreshes. "
        + "Grey columns are filled by the hub, white ones are yours",
    ),
    text(
      "- Settings and Email Templates are opened with SpiritHub > SpiritHub Settings, and are hidden by any refresh",
    ),
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

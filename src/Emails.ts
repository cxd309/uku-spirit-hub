/**
 * name of the Club Contacts tab
 */
const CLUB_CONTACTS_SHEET = "Club Contacts";

/**
 * text for the info row above the Club Contacts table
 */
const CLUB_CONTACTS_INFO = "Email address for each club, looked up by the Email Address column of the Issues tab\n\n"
  + "User editable columns:\n"
  + "- Email: one address, or several separated by commas.\n"
  + "- Notes\n\n"
  + "New clubs are added at the bottom whenever Clubs is rebuilt, rows are never changed or removed by the script. "
  + "If a club is renamed by a Club Override, copy its email to the new name";

/**
 * club contacts columns, in order
 */
const CLUB_CONTACT_HEADERS = Object.freeze(
  {
    club: "Club",
    email: "Email",
    notes: "Notes",
  } as const,
);

/**
 * keys of the Club Contacts tab in column order
 */
const CLUB_CONTACT_KEYS = Object.keys(CLUB_CONTACT_HEADERS) as (keyof typeof CLUB_CONTACT_HEADERS)[];

/**
 * 0-based Club Contacts columns a person edits, left white, the Club column is shaded
 */
const CLUB_CONTACT_EDITABLE = Object.freeze(
  (["email", "notes"] as const).map((key) => CLUB_CONTACT_KEYS.indexOf(key)),
);

/**
 * name of the Email Templates tab
 */
const EMAIL_TEMPLATES_SHEET = "Email Templates";

/**
 * every placeholder a template can use, with what it is replaced by
 */
const EMAIL_PLACEHOLDERS = Object.freeze(
  {
    club_name: "Club column of the issue",
    team_name: "Team column of the issue",
    tournament_name: "Tournament column of the issue",
    tournament_date: "Tournament Date, e.g. 1 November 2025",
    game: "Game column of the issue, blank for tournament issues and forms with no game",
    opponent: "Opponent column of the issue",
    checks: "Checks column of the issue",
    details: "Details column of the issue",
    committee_member: "Committee Member column of the issue",
    missing_count: "NO-SCORES and UNDER-SUBMITTED only: how many scores were not submitted",
    missing_opponents: "NO-SCORES and UNDER-SUBMITTED only: who they were for, e.g. Durham 1 (2 games) and York 2",
  } as const,
);

/**
 * a placeholder name, e.g. "club_name"
 */
type EmailPlaceholder = keyof typeof EMAIL_PLACEHOLDERS;

/**
 * text for the info row above the Email Templates table
 */
const EMAIL_TEMPLATES_INFO =
  "One email template per issue check, filling Email Subject and Email Body on the Issues tab\n"
  + "An issue with one check uses that check's template, one with several uses MULTIPLE, one with none uses MANUAL\n"
  + "Only NEW issues are filled, by Refresh Issues and whenever this tab is edited\n\n"
  + "User editable columns: Enabled, Subject and Body. Do not change the Issue Code column. "
  + "Only issues whose template is ticked Enabled get a draft. Missing templates are added back with starter wording\n\n"
  + "Placeholders, replaced from the issue row:\n"
  + Object.entries(EMAIL_PLACEHOLDERS).map(([name, what]) => `- {${name}}: ${what}`).join("\n")
  + "\n\nAny refresh hides this tab, open it again with SpiritHub > SpiritHub Settings";

/**
 * email templates columns, in order
 */
const EMAIL_TEMPLATE_HEADERS = Object.freeze(
  {
    code: "Issue Code",
    enabled: "Enabled",
    subject: "Subject",
    body: "Body",
  } as const,
);

/**
 * keys of the Email Templates tab in column order
 */
const EMAIL_TEMPLATE_KEYS = Object.keys(EMAIL_TEMPLATE_HEADERS) as (keyof typeof EMAIL_TEMPLATE_HEADERS)[];

/**
 * 0-based Email Templates columns a person edits, left white, the Issue Code column is shaded
 */
const EMAIL_TEMPLATE_EDITABLE = Object.freeze(
  (["enabled", "subject", "body"] as const).map((key) => EMAIL_TEMPLATE_KEYS.indexOf(key)),
);

/**
 * one row of the Email Templates tab
 */
interface EmailTemplate {
  /** issue check code, e.g. NO-SCORES, or MULTIPLE_EMAIL_CODE */
  code: string;
  /** ticked to draft emails for this check */
  enabled: boolean;
  /** subject, may contain placeholders */
  subject: string;
  /** body, may contain placeholders */
  body: string;
}

/**
 * the sign-off shared by every starter template
 */
const EMAIL_SIGN_OFF = "\n\nThanks,\nUKU Spirit Committee";

/**
 * code of the template used by an issue with more than one check
 */
const MULTIPLE_EMAIL_CODE = "MULTIPLE";

/**
 * starter wording written for each check when its template is missing
 * only NO-SCORES and UNDER-SUBMITTED are enabled to begin with
 * checks left out here never get a template or an email (e.g. CHEATING and HARASSMENT, which are not policy)
 */
const DEFAULT_EMAIL_TEMPLATES: Readonly<
  Partial<Record<IssueCategoryKey, { enabled: boolean; subject: string; body: string }>>
> = Object.freeze({
  lowScore: {
    enabled: false,
    subject: "Low spirit score - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received a low spirit score from {opponent} at {tournament_name}.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  dangerousPlay: {
    enabled: false,
    subject: "Spirit comment received - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "A spirit score comment {team_name} received from {opponent} at {tournament_name} mentions dangerous play."
      + "\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  categoryMinimum: {
    enabled: false,
    subject: "Spirit score received - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received a 0 in a spirit category from {opponent} at {tournament_name}.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  totalWithoutComment: {
    enabled: false,
    subject: "Spirit score without a comment - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "At {tournament_name}, {team_name} gave {opponent} a spirit score total that the UKU SOTG Policy says "
      + "needs a comment, but no comment was given.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  categoryWithoutComment: {
    enabled: false,
    subject: "Spirit score without a comment - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "At {tournament_name}, {team_name} gave {opponent} a category score that the UKU SOTG Policy says "
      + "needs a comment, but no comment was given.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  noScores: {
    enabled: true,
    subject: "No spirit scores submitted - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} did not submit any spirit scores at {tournament_name}, "
      + "for games against {missing_opponents}.\n\n"
      + "Please do not submit them now, but make sure your captains submit scores after every game at future events."
      + EMAIL_SIGN_OFF,
  },
  underSubmitted: {
    enabled: true,
    subject: "Missing spirit scores - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} did not submit {missing_count} spirit score(s) at {tournament_name}, "
      + "for games against {missing_opponents}.\n\n"
      + "Please do not submit them now, but make sure your captains submit scores after every game at future events."
      + EMAIL_SIGN_OFF,
  },
  immediateAverage: {
    enabled: false,
    subject: "Low spirit average - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received a low average spirit score at {tournament_name}.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  monitoringAverage: {
    enabled: false,
    subject: "Spirit average - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received an average spirit score at {tournament_name} that counts towards monitoring."
      + "\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
});

/**
 * starter wording for an issue with more than one check, to be filled in by hand
 */
const DEFAULT_MULTIPLE_EMAIL_TEMPLATE = Object.freeze({
  enabled: false,
  subject: "Spirit - {tournament_name}",
  body: "Dear {club_name},\n\n"
    + "We are getting in touch about {team_name} at {tournament_name}, in the game against {opponent}.\n\n"
    + "{checks}\n\n{details}"
    + EMAIL_SIGN_OFF,
});

/**
 * code of the template used by an issue with no checks, e.g. one added by hand
 */
const MANUAL_EMAIL_CODE = "MANUAL";

/**
 * starter wording for an issue with no checks, to be filled in by hand
 */
const DEFAULT_MANUAL_EMAIL_TEMPLATE = Object.freeze({
  enabled: false,
  subject: "Spirit - {tournament_name}",
  body: "Dear {club_name},\n\n"
    + "We are getting in touch about {team_name} at {tournament_name}.\n\n{details}"
    + EMAIL_SIGN_OFF,
});

/**
 * text of the Email Address cell when the club has no address on the Club Contacts tab
 */
const NO_CLUB_EMAIL = "No email for club";

/**
 * text of the Email Address cell when DELETED-ROW is the issue's only check
 * a form can be deleted for something either team did, so neither club's address is given
 */
const UNKNOWN_RECIPIENT = "Could not be determined";

/**
 * one drafted email, for one Issues row
 */
interface EmailDraft {
  /** index of the issue in the rows given to _draftEmails_ */
  index: number;
  /** filled subject */
  subject: string;
  /** filled body */
  body: string;
}

/**
 * formula for the Email Address cell of one Issues row: the club's email from the Club Contacts tab
 * a formula, so a changed address shows straight away and a row added by hand can copy it down
 * an issue whose only check is DELETED-ROW gives no address, see UNKNOWN_RECIPIENT
 *
 * @param row  1-based sheet row the formula is for
 * @returns the formula
 */
function _emailAddressFormula_(row: number): string {
  const cell = (key: keyof typeof ISSUE_HEADERS) => `$${_columnLetter_(ISSUE_KEYS.indexOf(key) + 1)}${row}`;
  const contacts = (key: keyof typeof CLUB_CONTACT_HEADERS) =>
    _columnBelowHeader_(CLUB_CONTACTS_SHEET, CLUB_CONTACT_KEYS.indexOf(key) + 1);
  return `=IF(TRIM(${cell("checks")})="${ISSUE_CATEGORIES.deletedRow.code}", "${UNKNOWN_RECIPIENT}", `
    + `LET(address, IFERROR(XLOOKUP(${cell("club")}, ${contacts("club")}, ${contacts("email")}), ""), `
    + `IF(TRIM(address)="", "${NO_CLUB_EMAIL}", address)))`;
}

/**
 * @returns the Club Contacts tab, created on first use
 */
function _getClubContactsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(CLUB_CONTACT_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), CLUB_CONTACTS_SHEET, headers, CLUB_CONTACTS_INFO);
}

/**
 * every row of the Club Contacts tab
 *
 * @param sheet  the Club Contacts tab
 * @returns one array per row, as wide as the table
 */
function _readClubContactRows_(sheet: GoogleAppsScript.Spreadsheet.Sheet): unknown[][] {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return [];
  return sheet.getRange(DATA_ROW, 1, lastRow - HEADER_ROW, CLUB_CONTACT_KEYS.length).getValues();
}

/**
 * clubs not yet on the Club Contacts tab
 * pure, never modifies its arguments
 * matching ignores case
 *
 * @param clubs     every club, in display order
 * @param existing  club names already on the tab
 * @returns the clubs to add, in the order given
 */
function _newContactClubs_(clubs: string[], existing: unknown[]): string[] {
  const present = new Set(existing.map((club) => String(club).trim().toLowerCase()));
  return clubs.filter((club) => !present.has(club.toLowerCase()));
}

/**
 * add any club missing from the Club Contacts tab at the bottom, with a blank email
 * existing rows are never changed, so people can sort and edit freely
 * does not take the lock, callers are responsible
 *
 * @param clubs  every club, in display order
 */
function _addClubContacts_(clubs: string[]) {
  const sheet = _getClubContactsSheet_();
  const added = _newContactClubs_(clubs, _readClubContactRows_(sheet).map((row) => row[0]));
  if (added.length === 0) return;
  const firstNew = Math.max(sheet.getLastRow(), HEADER_ROW) + 1;
  _fitSheet_(sheet, firstNew + added.length - 1, CLUB_CONTACT_KEYS.length);
  sheet.getRange(firstNew, 1, added.length, CLUB_CONTACT_KEYS.length).setValues(added.map((club) => [club, "", ""]));
  _applyFilter_(sheet, firstNew + added.length - 1 - HEADER_ROW, CLUB_CONTACT_KEYS.length);
  _styleRows_(sheet, HEADER_ROW, 1, CLUB_CONTACT_KEYS.length, CLUB_CONTACT_EDITABLE);
  _styleRows_(sheet, firstNew, added.length, CLUB_CONTACT_KEYS.length, CLUB_CONTACT_EDITABLE);
}

/**
 * default rows for the Email Templates tab
 *
 * @returns one row per issue check that has a template, in the order of ISSUE_CATEGORIES, then MULTIPLE and MANUAL
 */
function _defaultEmailTemplateRows_(): unknown[][] {
  const multiple = DEFAULT_MULTIPLE_EMAIL_TEMPLATE;
  const manual = DEFAULT_MANUAL_EMAIL_TEMPLATE;
  return [
    ..._issueCategoryKeys_().flatMap((key) => {
      const template = DEFAULT_EMAIL_TEMPLATES[key];
      return template ? [[ISSUE_CATEGORIES[key].code, template.enabled, template.subject, template.body]] : [];
    }),
    [MULTIPLE_EMAIL_CODE, multiple.enabled, multiple.subject, multiple.body],
    [MANUAL_EMAIL_CODE, manual.enabled, manual.subject, manual.body],
  ];
}

/**
 * get the Email Templates tab, creating it and adding back any missing template
 * missing templates go at the bottom, existing rows are never changed
 *
 * @returns the Email Templates tab
 */
function _getEmailTemplatesSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(EMAIL_TEMPLATE_HEADERS);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isNew = ss.getSheetByName(EMAIL_TEMPLATES_SHEET) === null;
  const sheet = _getOrCreateSheet_(ss, EMAIL_TEMPLATES_SHEET, headers, EMAIL_TEMPLATES_INFO);
  if (isNew) {
    sheet.setColumnWidth(EMAIL_TEMPLATE_KEYS.indexOf("subject") + 1, 300);
    sheet.setColumnWidth(EMAIL_TEMPLATE_KEYS.indexOf("body") + 1, 600);
  }

  const lastRow = sheet.getLastRow();
  const present = new Set(
    lastRow < DATA_ROW
      ? []
      : sheet.getRange(DATA_ROW, 1, lastRow - HEADER_ROW, 1).getValues().map((row) => String(row[0]).trim()),
  );
  const missing = _defaultEmailTemplateRows_().filter((row) => !present.has(String(row[0])));
  _styleRows_(sheet, HEADER_ROW, 1, EMAIL_TEMPLATE_KEYS.length, EMAIL_TEMPLATE_EDITABLE);
  if (missing.length === 0) return sheet;

  const firstNew = Math.max(lastRow, HEADER_ROW) + 1;
  _fitSheet_(sheet, firstNew + missing.length - 1, EMAIL_TEMPLATE_KEYS.length);
  sheet.getRange(firstNew, 1, missing.length, EMAIL_TEMPLATE_KEYS.length).setValues(missing).setVerticalAlignment(
    "top",
  );
  sheet
    .getRange(firstNew, EMAIL_TEMPLATE_KEYS.indexOf("enabled") + 1, missing.length, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sheet.getRange(firstNew, EMAIL_TEMPLATE_KEYS.indexOf("subject") + 1, missing.length, 2).setWrap(true);
  _styleRows_(sheet, firstNew, missing.length, EMAIL_TEMPLATE_KEYS.length, EMAIL_TEMPLATE_EDITABLE);
  return sheet;
}

/**
 * every template on the Email Templates tab
 *
 * @returns issue code → template
 */
function _readEmailTemplates_(): Map<string, EmailTemplate> {
  const sheet = _getEmailTemplatesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return new Map();
  const rows = sheet.getRange(DATA_ROW, 1, lastRow - HEADER_ROW, EMAIL_TEMPLATE_KEYS.length).getValues();
  const templates: Map<string, EmailTemplate> = new Map();
  for (const row of rows) {
    const code = String(row[0]).trim();
    if (code === "" || templates.has(code)) continue;
    templates.set(code, { code, enabled: row[1] === true, subject: String(row[2]), body: String(row[3]) });
  }
  return templates;
}

/**
 * fill a template's placeholders
 * pure, never modifies its arguments
 *
 * @param text    subject or body
 * @param values  value for every placeholder
 * @returns the filled text, and any {placeholder} that is not known, left in place
 */
function _fillTemplate_(text: string, values: Record<EmailPlaceholder, string>): { text: string; unknown: string[] } {
  const unknown: string[] = [];
  const filled = text.replace(/\{([^{}\s]+)\}/g, (match, name: string) => {
    if (Object.prototype.hasOwnProperty.call(values, name)) return values[name as EmailPlaceholder];
    unknown.push(match);
    return match;
  });
  return { text: filled, unknown };
}

/**
 * join names as readable text, e.g. "A, B and C"
 * pure
 *
 * @param names  the names
 * @returns the text, "" for none
 */
function _readableList_(names: string[]): string {
  if (names.length <= 1) return names.join("");
  return `${names.slice(0, -1).join(", ")} and ${names[names.length - 1]}`;
}

/**
 * how many scores were not submitted and for whom, from a NO-SCORES or UNDER-SUBMITTED issue's Details
 * pure
 * Details lists one opponent per missing score below "Missing for:", up to the next blank line
 *
 * @param details  the Details cell
 * @returns the count, and the opponents with a game count when more than one
 */
function _missingScores_(details: string): { count: number; opponents: string } {
  const [, after = ""] = details.split(/Missing for:\s*\n/);
  const names = after.split(/\n\s*\n/)[0].split("\n").map((line) => line.trim()).filter((line) => line !== "");
  const counts: Map<string, number> = new Map();
  for (const name of names) counts.set(name, (counts.get(name) ?? 0) + 1);
  const opponents = [...counts].map(([name, n]) => (n > 1 ? `${name} (${n} games)` : name));
  return { count: names.length, opponents: _readableList_(opponents) };
}

/**
 * one Issues row, as read back for drafting
 */
type IssueRow = Record<keyof typeof ISSUE_HEADERS, unknown>;

/**
 * every row of the Issues tab
 *
 * @returns one record per row, in sheet order
 */
function _readIssueRows_(): IssueRow[] {
  const sheet = _getIssuesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return [];
  return sheet
    .getRange(DATA_ROW, 1, lastRow - HEADER_ROW, ISSUE_KEYS.length)
    .getValues()
    .map((row) => Object.fromEntries(ISSUE_KEYS.map((key, i) => [key, row[i]])) as IssueRow);
}

/**
 * draft one email per NEW issue whose template is enabled
 * pure, never modifies its arguments
 * an issue with one check uses that check's template, one with several MULTIPLE, one with none MANUAL
 *
 * @param issues      every Issues row
 * @param templates   issue code → template
 * @param formatDate  turns a tournament date into text
 * @returns the emails, and every unknown placeholder as "CODE: {name}"
 */
function _draftEmails_(
  issues: IssueRow[],
  templates: Map<string, EmailTemplate>,
  formatDate: (date: Date) => string,
): { drafts: EmailDraft[]; unknown: string[] } {
  const drafts: EmailDraft[] = [];
  const unknown: Set<string> = new Set();
  const text = (value: unknown) => String(value ?? "").trim();

  issues.forEach((issue, index) => {
    if (text(issue.status).toUpperCase() !== ISSUE_STATUSES[0]) return;
    const codes = _checkCodes_(issue.checks);
    const code = codes.length === 0 ? MANUAL_EMAIL_CODE : codes.length === 1 ? codes[0] : MULTIPLE_EMAIL_CODE;
    const template = templates.get(code);
    if (!template || !template.enabled) return;

    const details = text(issue.details);
    const missingCodes: string[] = [ISSUE_CATEGORIES.noScores.code, ISSUE_CATEGORIES.underSubmitted.code];
    const missing = missingCodes.includes(code) ? _missingScores_(details) : { count: NaN, opponents: "" };
    const date = issue.tournamentDate;
    const values: Record<EmailPlaceholder, string> = {
      club_name: text(issue.club),
      team_name: text(issue.team),
      tournament_name: text(issue.tournament),
      tournament_date: date instanceof Date ? formatDate(date) : text(date),
      game: text(issue.game),
      opponent: text(issue.opponent),
      checks: text(issue.checks),
      details,
      committee_member: text(issue.committeeMember),
      missing_count: Number.isNaN(missing.count) ? "" : String(missing.count),
      missing_opponents: missing.opponents,
    };

    const subject = _fillTemplate_(template.subject, values);
    const body = _fillTemplate_(template.body, values);
    for (const name of [...subject.unknown, ...body.unknown]) unknown.add(`${code}: ${name}`);
    drafts.push({ index, subject: subject.text, body: body.text });
  });
  return { drafts, unknown: [...unknown] };
}

/**
 * the Email Subject and Email Body columns after drafting
 * pure, never modifies its arguments
 * NEW issues get their draft, or blank when their template is missing or disabled
 * every other issue keeps what it had, so the wording of an email already sent stays
 *
 * @param issues  every Issues row
 * @param drafts  from _draftEmails_
 * @returns one subject and one body per issue
 */
function _emailColumns_(issues: IssueRow[], drafts: EmailDraft[]): { subject: string; body: string }[] {
  const cells = issues.map((issue) =>
    String(issue.status ?? "").trim().toUpperCase() === ISSUE_STATUSES[0]
      ? { subject: "", body: "" }
      : { subject: String(issue.emailSubject ?? ""), body: String(issue.emailBody ?? "") }
  );
  for (const draft of drafts) cells[draft.index] = { subject: draft.subject, body: draft.body };
  return cells;
}

/**
 * fill Email Subject and Email Body for every NEW issue whose template is enabled
 * run by Refresh Issues, and whenever the Email Templates tab is edited
 * reads the Issues and Email Templates tabs, writes only those two columns, nothing is sent
 *
 * @returns a one-line summary
 * @throws {Error} if a template uses a placeholder that does not exist
 */
function _updateEmailDrafts_(): string {
  const sheet = _getIssuesSheet_();
  const issues = _readIssueRows_();
  if (issues.length === 0) return "no issues to email";

  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const { drafts, unknown } = _draftEmails_(
    issues,
    _readEmailTemplates_(),
    (date) => Utilities.formatDate(date, timeZone, "d MMMM yyyy"),
  );
  if (unknown.length > 0) {
    throw new Error(
      `Email Templates: unknown placeholder(s) ${unknown.join(", ")}. `
        + `Use one of ${Object.keys(EMAIL_PLACEHOLDERS).map((name) => `{${name}}`).join(", ")}`,
    );
  }

  const columns = _emailColumns_(issues, drafts);
  const write = (key: "emailSubject" | "emailBody", field: "subject" | "body") =>
    sheet
      .getRange(DATA_ROW, ISSUE_KEYS.indexOf(key) + 1, issues.length, 1)
      .setValues(columns.map((cell) => [cell[field]]));
  write("emailSubject", "subject");
  write("emailBody", "body");
  return `${drafts.length} email draft(s)`;
}

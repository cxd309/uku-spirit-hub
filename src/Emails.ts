/**
 * name of the Club Contacts tab
 */
const CLUB_CONTACTS_SHEET = "Club Contacts";

/**
 * text for the info row above the Club Contacts table
 */
const CLUB_CONTACTS_INFO = "Email address for each club, used by Draft Emails\n\n"
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
    tournament_name: "Tournament(s) column of the issue",
    tournament_date: "Tournament Date, e.g. 1 November 2025",
    issue_category: "Issue Category column of the issue",
    details: "Details column of the issue",
    committee_member: "Committee Member column of the issue",
    missing_count: "NOT-SUBMITTED only: how many scores were not submitted",
    missing_opponents: "NOT-SUBMITTED only: who they were for, e.g. Durham 1 (2 games) and York 2",
  } as const,
);

/**
 * a placeholder name, e.g. "club_name"
 */
type EmailPlaceholder = keyof typeof EMAIL_PLACEHOLDERS;

/**
 * text for the info row above the Email Templates table
 */
const EMAIL_TEMPLATES_INFO = "One email template per policy issue check, used by Draft Emails\n\n"
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
 * one row of the Email Templates tab
 */
interface EmailTemplate {
  /** issue check code, e.g. NOT-SUBMITTED */
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
 * starter wording written for each check when its template is missing
 * only NOT-SUBMITTED is enabled to begin with
 * checks left out here never get a template or an email (e.g. CHEATING and HARASSMENT, which are not policy)
 */
const DEFAULT_EMAIL_TEMPLATES: Readonly<
  Partial<Record<IssueCategoryKey, { enabled: boolean; subject: string; body: string }>>
> = Object.freeze({
  totalWithoutComment: {
    enabled: false,
    subject: "Spirit score without a comment - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "At {tournament_name}, {team_name} gave a spirit score total that the UKU SOTG Policy says needs a comment, "
      + "but no comment was given.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  categoryWithoutComment: {
    enabled: false,
    subject: "Spirit score without a comment - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "At {tournament_name}, {team_name} gave a category score that the UKU SOTG Policy says needs a comment, "
      + "but no comment was given.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  dangerousPlay: {
    enabled: false,
    subject: "Spirit comment received - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "A spirit score comment received by {team_name} at {tournament_name} mentions dangerous play.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  notSubmitted: {
    enabled: true,
    subject: "Missing spirit scores - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} did not submit {missing_count} spirit score(s) at {tournament_name}, "
      + "for games against {missing_opponents}.\n\n"
      + "Please do not submit them now, but make sure your captains submit scores after every game at future events."
      + EMAIL_SIGN_OFF,
  },
  twoLowScores: {
    enabled: false,
    subject: "Low spirit scores - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received several low spirit scores at {tournament_name}.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  lowAverage: {
    enabled: false,
    subject: "Low spirit average - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received a low average spirit score at {tournament_name}.\n\n{details}"
      + EMAIL_SIGN_OFF,
  },
  categoryMinimum: {
    enabled: false,
    subject: "Spirit score received - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received a 0 in a spirit category at {tournament_name}.\n\n{details}" + EMAIL_SIGN_OFF,
  },
  singleLowScore: {
    enabled: false,
    subject: "Low spirit score - {tournament_name}",
    body: "Dear {club_name},\n\n"
      + "{team_name} received a low spirit score at {tournament_name}.\n\n{details}" + EMAIL_SIGN_OFF,
  },
});

/**
 * text of the Email cell when the club has no address on the Club Contacts tab
 */
const NO_CLUB_EMAIL = "No email for club";

/**
 * one drafted email, for one Issues row
 */
interface EmailDraft {
  /** index of the issue in the rows given to _draftEmails_ */
  index: number;
  /** club's email, "" if it has none */
  to: string;
  /** filled subject */
  subject: string;
  /** filled body */
  body: string;
}

/**
 * the Email cell for a draft: an Open in mail link, or a note that the club has no email
 * pure
 * subject and body are percent-encoded and quotes are dropped from addresses, so the formula never contains a quote
 *
 * @param draft  the drafted email
 * @returns a formula, or the no email text
 */
function _emailCell_(draft: EmailDraft): string {
  if (draft.to === "") return NO_CLUB_EMAIL;
  const to = draft.to.split(",").map((address) => address.trim().replace(/[\s"?&]/g, "")).join(",");
  const url = `mailto:${to}?subject=${encodeURIComponent(draft.subject)}&body=${encodeURIComponent(draft.body)}`;
  return `=HYPERLINK("${url}","Open in mail")`;
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
}

/**
 * each club's email from the Club Contacts tab
 *
 * @returns lower case club name → email, blank emails left out
 */
function _readClubEmails_(): Map<string, string> {
  const emails: Map<string, string> = new Map();
  for (const row of _readClubContactRows_(_getClubContactsSheet_())) {
    const club = String(row[0]).trim().toLowerCase();
    const email = String(row[1]).trim();
    if (club !== "" && email !== "" && !emails.has(club)) emails.set(club, email);
  }
  return emails;
}

/**
 * default rows for the Email Templates tab
 *
 * @returns one row per issue check that has a template, in the order of ISSUE_CATEGORIES
 */
function _defaultEmailTemplateRows_(): unknown[][] {
  return _issueCategoryKeys_().flatMap((key) => {
    const template = DEFAULT_EMAIL_TEMPLATES[key];
    return template ? [[ISSUE_CATEGORIES[key].code, template.enabled, template.subject, template.body]] : [];
  });
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
 * how many scores were not submitted and for whom, from a NOT-SUBMITTED issue's Details
 * pure
 * Details lists one opponent per missing score below "Missing for:"
 *
 * @param details  the Details cell
 * @returns the count, and the opponents with a game count when more than one
 */
function _missingScores_(details: string): { count: number; opponents: string } {
  const [, after = ""] = details.split(/Missing for:\s*\n/);
  const names = after.split("\n").map((line) => line.trim()).filter((line) => line !== "");
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
 * the check is the code at the start of the Issue ID
 *
 * @param issues      every Issues row
 * @param templates   issue code → template
 * @param emails      lower case club → email
 * @param formatDate  turns a tournament date into text
 * @returns the emails, and every unknown placeholder as "CODE: {name}"
 */
function _draftEmails_(
  issues: IssueRow[],
  templates: Map<string, EmailTemplate>,
  emails: Map<string, string>,
  formatDate: (date: Date) => string,
): { drafts: EmailDraft[]; unknown: string[] } {
  const drafts: EmailDraft[] = [];
  const unknown: Set<string> = new Set();
  const text = (value: unknown) => String(value ?? "").trim();

  issues.forEach((issue, index) => {
    if (text(issue.status).toUpperCase() !== ISSUE_STATUSES[0]) return;
    const code = text(issue.issueId).split(" | ")[0];
    const template = templates.get(code);
    if (!template || !template.enabled) return;

    const details = text(issue.details);
    const missing = code === ISSUE_CATEGORIES.notSubmitted.code
      ? _missingScores_(details)
      : { count: NaN, opponents: "" };
    const date = issue.tournamentDate;
    const values: Record<EmailPlaceholder, string> = {
      club_name: text(issue.club),
      team_name: text(issue.team),
      tournament_name: text(issue.tournaments),
      tournament_date: date instanceof Date ? formatDate(date) : text(date),
      issue_category: text(issue.category),
      details,
      committee_member: text(issue.committeeMember),
      missing_count: Number.isNaN(missing.count) ? "" : String(missing.count),
      missing_opponents: missing.opponents,
    };

    const subject = _fillTemplate_(template.subject, values);
    const body = _fillTemplate_(template.body, values);
    for (const name of [...subject.unknown, ...body.unknown]) unknown.add(`${code}: ${name}`);

    drafts.push({
      index,
      to: emails.get(values.club_name.toLowerCase()) ?? "",
      subject: subject.text,
      body: body.text,
    });
  });
  return { drafts, unknown: [...unknown] };
}

/**
 * the whole Email column after drafting
 * pure, never modifies its arguments
 * NEW issues get their draft's link, or blank when their template is missing or disabled
 * every other issue keeps what it had, so the link to an email already sent stays
 *
 * @param issues   every Issues row
 * @param current  the Email column as it is, formulas where there are any
 * @param drafts   from _draftEmails_
 * @returns one cell per issue, formulas start with "="
 */
function _emailColumn_(issues: IssueRow[], current: string[], drafts: EmailDraft[]): string[] {
  const cells = issues.map((issue, i) =>
    String(issue.status ?? "").trim().toUpperCase() === ISSUE_STATUSES[0] ? "" : current[i] ?? ""
  );
  for (const draft of drafts) cells[draft.index] = _emailCell_(draft);
  return cells;
}

/**
 * fill the Email column of the Issues tab with an Open in mail link for every NEW issue whose template is enabled
 * reads the Issues, Email Templates and Club Contacts tabs, nothing is sent
 * only the Email column is written
 *
 * @returns a one-line summary
 * @throws {Error} if a template uses a placeholder that does not exist
 */
function _updateEmailLinks_(): string {
  const sheet = _getIssuesSheet_();
  const issues = _readIssueRows_();
  if (issues.length === 0) return "no issues to email";

  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const { drafts, unknown } = _draftEmails_(
    issues,
    _readEmailTemplates_(),
    _readClubEmails_(),
    (date) => Utilities.formatDate(date, timeZone, "d MMMM yyyy"),
  );
  if (unknown.length > 0) {
    throw new Error(
      `Email Templates: unknown placeholder(s) ${unknown.join(", ")}. `
        + `Use one of ${Object.keys(EMAIL_PLACEHOLDERS).map((name) => `{${name}}`).join(", ")}`,
    );
  }

  const range = sheet.getRange(DATA_ROW, ISSUE_KEYS.indexOf("email") + 1, issues.length, 1);
  const formulas = range.getFormulas();
  const values = range.getValues();
  const current = formulas.map((row, i) => (row[0] !== "" ? row[0] : String(values[i][0])));
  range.setValues(_emailColumn_(issues, current, drafts).map((cell) => [cell]));

  const noEmail = drafts.filter((d) => d.to === "").length;
  const summary = `${drafts.length} email link(s)`;
  return noEmail === 0 ? summary : `${summary}, ${noEmail} with no club email (see Club Contacts)`;
}

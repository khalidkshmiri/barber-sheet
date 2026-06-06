/***************
 * FORMATTING
 * Visual theme for all sheets. Safe to re-run at any time.
 ***************/

/**
 * Applies the visual theme to all sheets.
 * Run from the Apps Script editor dropdown. Safe to re-run at any time.
 */
function formatSpreadsheet() {
  const COLORS = {
    bg:        '#FAF7F2',
    surface:   '#F2EDE5',
    headerBg:  '#E8DDD0',
    text:      '#2C2017',
    textMuted: '#8C7B6B',
    accent: {
      blue:   '#4A7CA7',
      green:  '#5A8A5A',
      red:    '#B85050',
      yellow: '#B8920A',
      purple: '#8A6A9A',
      orange: '#C47040',
    },
    tint: {
      green:  '#E8F2E8',
      red:    '#F5E4E4',
      yellow: '#F5EDD0',
      purple: '#EDE8F2',
      blue:   '#E4EEF5',
      orange: '#F5EAE0',
    },
  };

  try {
    Logger.log('formatSpreadsheet started');
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    formatAppointments_(ss.getSheetByName('Appointments'), COLORS);
    formatClients_(ss.getSheetByName('Clients'), COLORS);
    formatServices_(ss.getSheetByName('Services'), COLORS);
    formatSubscriptions_(ss.getSheetByName('Subscriptions'), COLORS);
    formatDashboard_(ss.getSheetByName('Dashboard'), COLORS);

    Logger.log('Flushing...');
    SpreadsheetApp.flush();
    Logger.log('✅ formatSpreadsheet complete');
  } catch (e) {
    Logger.log('formatSpreadsheet error: ' + e.message + '\n' + e.stack);
    sendTelegramError_('Format spreadsheet failed: ' + e.message);
    throw e;
  }
}

/**
 * Sets the base theme on a sheet. Tables start at B2: row 1 and col A are spacers.
 * Returns the alternating-row rule to be appended last (lowest priority) by the caller.
 */
function applyBaseTheme_(sheet, COLORS) {
  if (!sheet) return [];

  Logger.log('Formatting: ' + sheet.getName());

  const maxRows = sheet.getMaxRows();
  const maxCols = Math.min(Math.max(sheet.getLastColumn(), 1), sheet.getMaxColumns());
  const styledRows = Math.min(Math.max(sheet.getLastRow(), 1) + 10, maxRows);

  sheet.clearConditionalFormatRules();
  sheet.showColumns(1, maxCols);

  // Base background for all cells (including spacer row/col)
  sheet.getRange(1, 1, styledRows, maxCols)
    .setBackground(COLORS.bg)
    .setFontColor(COLORS.text)
    .setFontWeight('normal');

  // Header row (row 2): styled tan background, bold text
  sheet.getRange(HEADER_ROW, 1, 1, maxCols)
    .setBackground(COLORS.headerBg)
    .setFontColor(COLORS.text)
    .setFontWeight('bold');

  // Freeze both the spacer row and the header row so they stay visible while scrolling
  sheet.setFrozenRows(HEADER_ROW);

  sheet.setRowHeight(1, 20);
  sheet.setRowHeight(HEADER_ROW, 150);
  if (maxRows > HEADER_ROW) {
    sheet.setRowHeightsForced(DATA_ROW, maxRows - DATA_ROW + 1, 260);
  }

  // Alternating row rule — applied only to data rows, lowest priority so status colors win
  const altRule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=MOD(ROW(),2)=0'])
    .setBackground(COLORS.surface)
    .setRanges([sheet.getRange(DATA_ROW, 1, Math.max(maxRows - DATA_ROW + 1, 1), maxCols)])
    .build();

  return [altRule];
}

/**
 * Formats the Appointments sheet.
 * Cols: A=spacer B=Date C=Time D=Name E=Price F=Payment G=Status H=Tips I=Late J=Notes K=Service
 *       L=CachedName(hidden) M=ClientID(hidden) N=EventID(hidden) O=spacer
 */
function formatAppointments_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules = applyBaseTheme_(sheet, COLORS);

  // A=spacer(20) B=Date(125) C=Time(125) D=Name(280) E=Price(125) F=Payment(225) G=Status(225)
  // H=Tips(125) I=Late(125) J=Notes(280) K=Service(125) L=CachedName(300) M=ClientID(125) N=EventID(300) O=spacer(20)
  const widths = [20, 125, 125, 280, 125, 225, 225, 125, 125, 180, 125, 300, 125, 300, 20];
  widths.forEach((w, i) => { if (w > 0) sheet.setColumnWidth(i + 1, w); });

  // Hide K=Service(11) + internal columns L=CachedName(12) M=ClientID(13) N=EventID(14)
  sheet.hideColumns(11, 4);

  const maxRows = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  // Data range: B3 (col 2, row DATA_ROW), 13 cols wide (B-N)
  const dataRange = sheet.getRange(DATA_ROW, 2, dataRows, 13);

  dataRange
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheet.getRange(HEADER_ROW, 2, 1, 11).setFontSize(18); // header labels B-L

  // Hero columns: Time (C=3) and Name (D=4)
  sheet.getRange(DATA_ROW, 3, dataRows, 1).setFontSize(25).setFontWeight('bold');
  sheet.getRange(DATA_ROW, 4, dataRows, 1).setFontSize(30).setFontWeight('bold');

  // Date (B=2): slightly smaller
  sheet.getRange(DATA_ROW, 2, dataRows, 1).setFontSize(18);

  // Price (E=5) and Payment (F=6)
  sheet.getRange(DATA_ROW, 5, dataRows, 1).setFontSize(25);
  sheet.getRange(DATA_ROW, 6, dataRows, 1).setFontSize(23);

  // Secondary info: muted colour
  sheet.getRange(DATA_ROW, 10, dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted); // Notes J=10
  sheet.getRange(DATA_ROW, 11, dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted); // Service K=11

  // Horizontal alignment
  const centred = [2, 3, 5, 6, 7, 8, 9]; // Date, Time, Price, Payment, Status, Tips, Late
  centred.forEach(col => sheet.getRange(DATA_ROW, col, dataRows, 1).setHorizontalAlignment('center'));

  const ruleDefs = [
    { formula: `=$I${DATA_ROW}=TRUE`,           bg: COLORS.tint.orange, fg: COLORS.accent.orange }, // Late
    { formula: `=$G${DATA_ROW}="No Show"`,      bg: COLORS.tint.red,    fg: COLORS.accent.red    },
    { formula: `=$G${DATA_ROW}="Cancelled"`,    bg: COLORS.bg,          fg: COLORS.textMuted     },
    { formula: `=$F${DATA_ROW}="Subscription"`, bg: COLORS.tint.purple, fg: COLORS.accent.purple },
    { formula: `=$G${DATA_ROW}="Paid"`,         bg: COLORS.tint.green,  fg: COLORS.accent.green  },
    { formula: `=$G${DATA_ROW}="Upcoming"`,     bg: COLORS.tint.blue,   fg: COLORS.accent.blue   },
    { formula: `=$G${DATA_ROW}="Not Paid"`,     bg: COLORS.tint.yellow, fg: COLORS.accent.yellow },
  ];

  const rules = ruleDefs.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, [r.formula])
      .setBackground(r.bg)
      .setFontColor(r.fg)
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules([...rules, ...baseRules]);
  sheet.setTabColor(COLORS.accent.green);
}

/**
 * Formats the Clients sheet.
 * Cols: A=spacer B=Name C=FavService D=LastVisit E=SocialMedia F=Notes G=NoShow H=Late
 *       I=Referral(hidden) J=TotalVisits K=TotalSpent L=TotalTips M=FirstVisit
 *       N=ConsecutivePaid O=VIP P=DoNotCut Q=ClientID(hidden) R=spacer
 */
function formatClients_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules = applyBaseTheme_(sheet, COLORS);

  // A=spacer(20) B=Name(280) C=FavService(300) D=LastVisit(125) E=SocialMedia(300) F=Notes(280)
  // G=NoShow(125) H=Late(125) I=Referral(125) J=TotalVisits(125) K=TotalSpent(125)
  // L=TotalTips(125) M=FirstVisit(125) N=ConsecutivePaid(125) O=VIP(125) P=DoNotCut(125)
  // Q=ClientID(125) R=spacer(20)
  const widths = [20, 280, 300, 125, 300, 280, 125, 125, 125, 125, 125, 125, 125, 125, 125, 125, 125, 20];
  widths.forEach((w, i) => { if (w > 0) sheet.setColumnWidth(i + 1, w); });

  // Hide: I=Referral(9), Q=ClientID(17)
  sheet.hideColumns(9, 1);   // I Referral
  sheet.hideColumns(17, 1);  // Q ClientID

  const maxRows = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  // Data range: B3 (col 2, row DATA_ROW), 16 cols (B-Q)
  const dataRange = sheet.getRange(DATA_ROW, 2, dataRows, 16);

  dataRange
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(HEADER_ROW, 2, 1, 16).setFontSize(18);

  // Name (B=2) is the hero column
  sheet.getRange(DATA_ROW, 2, dataRows, 1).setFontSize(30).setFontWeight('bold');

  // Centre-align: LastVisit(D=4), NoShow(G=7), Late(H=8), ConsecutivePaid(N=14), VIP(O=15), DoNotCut(P=16)
  [4, 7, 8, 14, 15, 16].forEach(col =>
    sheet.getRange(DATA_ROW, col, dataRows, 1).setHorizontalAlignment('center')
  );

  sheet.getRange(DATA_ROW, 3, dataRows, 1).setFontSize(22); // FavService C=3
  sheet.getRange(DATA_ROW, 6, dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted); // Notes F=6

  const ruleDefs = [
    { formula: `=$P${DATA_ROW}=TRUE`,
      bg: COLORS.tint.red,    fg: COLORS.accent.red    }, // DoNotCut P=16 — highest priority
    { formula: `=$G${DATA_ROW}+$H${DATA_ROW}>=3`,
      bg: COLORS.tint.orange, fg: COLORS.accent.orange }, // Unreliable: 3+ incidents
    { formula: `=AND($G${DATA_ROW}+$H${DATA_ROW}>=1,$G${DATA_ROW}+$H${DATA_ROW}<=2)`,
      bg: COLORS.tint.blue,   fg: COLORS.accent.blue   }, // Watch: 1-2 incidents
    { formula: `=$O${DATA_ROW}=TRUE`,
      bg: COLORS.tint.yellow, fg: COLORS.accent.yellow }, // VIP O=15
    { formula: `=$N${DATA_ROW}>=5`,
      bg: COLORS.tint.green,  fg: COLORS.accent.green  }, // Free-eligible: ConsecutivePaid ≥ 5
  ];

  const rules = ruleDefs.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, [r.formula])
      .setBackground(r.bg)
      .setFontColor(r.fg)
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules([...rules, ...baseRules]);
  sheet.setTabColor(COLORS.accent.purple);
}

/**
 * Formats the Services sheet.
 * Cols: A=spacer B=Service C=Price D=spacer
 */
function formatServices_(sheet, COLORS) {
  if (!sheet) return;
  const baseRules = applyBaseTheme_(sheet, COLORS);

  // A=spacer(20) B=Service(300) C=Price(300) D=spacer(20)
  [20, 300, 300, 20].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const maxRows = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  sheet.getRange(DATA_ROW, 2, dataRows, 2)
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(HEADER_ROW, 2, 1, 2).setFontSize(18);

  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.textMuted);
}

/**
 * Formats the Subscriptions sheet.
 * Cols: A=spacer B=Name C=Price D=Type E=Expiry F=Credits G=Status H=Notes I=StartDate J=ClientID K=spacer
 */
function formatSubscriptions_(sheet, COLORS) {
  if (!sheet) return;
  const baseRules = applyBaseTheme_(sheet, COLORS);

  // A=spacer(20) B=Name(280) C=Price(125) D=Type(160) E=Expiry(160) F=Credits(125)
  // G=Status(230) H=Notes(280) I=StartDate(125) J=ClientID(125) K=spacer(20)
  const widths = [20, 280, 125, 160, 160, 125, 230, 280, 125, 125, 20];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const maxRows = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  const dataRange = sheet.getRange(DATA_ROW, 2, dataRows, 9);

  dataRange
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(HEADER_ROW, 2, 1, 9).setFontSize(18);

  // Per-column font sizes
  sheet.getRange(DATA_ROW, 2, dataRows, 1).setFontSize(30).setFontWeight('bold'); // Name B=2
  sheet.getRange(DATA_ROW, 3, dataRows, 1).setFontSize(25);                       // Price C=3
  sheet.getRange(DATA_ROW, 4, dataRows, 1).setFontSize(25);                       // Type D=4
  sheet.getRange(DATA_ROW, 5, dataRows, 1).setFontSize(20);                       // Expiry E=5
  sheet.getRange(DATA_ROW, 7, dataRows, 1).setFontSize(25);                       // Status G=7
  sheet.getRange(DATA_ROW, 9, dataRows, 1).setFontSize(20);                       // StartDate I=9

  // Conditional formatting by Status (G=7)
  const ruleDefs = [
    { formula: `=$G${DATA_ROW}="Active"`,    bg: COLORS.tint.green,  fg: COLORS.accent.green  },
    { formula: `=$G${DATA_ROW}="Expired"`,   bg: COLORS.tint.red,    fg: COLORS.accent.red    },
    { formula: `=$G${DATA_ROW}="Completed"`, bg: COLORS.tint.blue,   fg: COLORS.accent.blue   },
  ];

  const rules = ruleDefs.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, [r.formula])
      .setBackground(r.bg)
      .setFontColor(r.fg)
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules([...rules, ...baseRules]);
  sheet.setTabColor(COLORS.accent.orange);
}

/**
 * Formats the Dashboard sheet.
 * Sync checkbox is at C3 — clearFormats is avoided to preserve data validation.
 */
function formatDashboard_(sheet, COLORS) {
  if (!sheet) return;
  // Dashboard is manually formatted — leave all colours, borders, and fonts untouched.
  // Only ensure rows are not frozen (user preference).
  sheet.setFrozenRows(0);
}

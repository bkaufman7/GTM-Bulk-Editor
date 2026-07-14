/**
 * GTM Bulk Trigger Assignment Editor
 * Standalone Google Sheets + Apps Script tool.
 *
 * MVP scope:
 * - Source of truth is pasted GTM export JSON in RAW_JSON
 * - Only mutates tag.firingTriggerId and tag.blockingTriggerId
 * - Does not modify trigger definitions, variables, folders, templates, consent, etc.
 */

var APP = {
  MENU: 'GTM Bulk Editor',
  SHEETS: {
    READ_ME: 'Read Me',
    RAW_JSON: 'RAW_JSON',
    CONTAINER_INFO: 'Container Info',
    TRIGGER_DIR: 'Trigger Directory',
    TAG_DIR: 'Tag Directory',
    ASSIGN_CURRENT: 'Assignments – Current',
    ASSIGN_ADD: 'Assignments – Add',
    BULK_RULES: 'Bulk Rules',
    EDIT_PREVIEW: 'Edit Preview',
    EXPORT_JSON: 'EXPORT_JSON'
  },
  TAB_COLORS: {
    GRAY: '#9e9e9e',
    BLUE: '#42a5f5',
    YELLOW: '#fbc02d',
    PURPLE: '#ab47bc',
    GREEN: '#43a047'
  },
  CHUNK_SIZE: 45000,
  HEADER_BG: '#f3f6fa',
  HEADER_FONT: '#1f2937'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP.MENU)
    .addItem('Open JSON Loader', 'openJsonLoader')
    .addItem('Build Editor Tabs', 'buildEditorTabs')
    .addItem('Build Preview', 'buildPreview')
    .addItem('Apply Edits & Create Export JSON', 'applyEditsAndCreateExportJson')
    .addSeparator()
    .addItem('Reset Editor Workspace', 'resetEditorWorkspace')
    .addItem('Rebuild Read Me', 'rebuildReadMe')
    .addToUi();
}

function openJsonLoader() {
  ensureCoreSheets_();
  var html = HtmlService.createHtmlOutput(getLoaderSidebarHtml_())
    .setTitle('GTM JSON Loader')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveRawJsonFromSidebar(rawJsonText) {
  if (!rawJsonText || !String(rawJsonText).trim()) {
    throw new Error('No JSON text provided. Paste the full GTM export JSON and try again.');
  }

  var parsed;
  try {
    parsed = JSON.parse(rawJsonText);
  } catch (err) {
    throw new Error('Invalid JSON. Please paste a valid GTM container export JSON.\n\n' + err.message);
  }

  var cv = getContainerVersion_(parsed);
  if (!cv) {
    throw new Error('JSON parsed but no containerVersion object was found at root.containerVersion or root.');
  }

  writeRawJsonChunks_(rawJsonText);

  return {
    success: true,
    message: 'JSON saved to RAW_JSON successfully. Use "Build Editor Tabs" next.',
    tagCount: asArray_(cv.tag).length,
    triggerCount: asArray_(cv.trigger).length
  };
}

function buildEditorTabs() {
  ensureCoreSheets_();

  var parsed = getParsedRootFromRawSheet_();
  var root = parsed.root;
  var cv = parsed.containerVersion;

  validateContainerCore_(cv);

  buildContainerInfoSheet_(root, cv);
  buildTriggerDirectorySheet_(cv);
  buildTagDirectorySheet_(cv);
  buildAssignmentsCurrentSheet_(cv);
  buildAssignmentsAddSheet_(cv);
  buildBulkRulesSheet_();
  resetPreviewSheet_();
  resetExportSheet_();

  SpreadsheetApp.getUi().alert(
    'Editor tabs built successfully.\n\nNext step: Run "Build Preview" after selecting manual rows and/or enabling rules.'
  );
}

function buildPreview() {
  ensureCoreSheets_();

  var previewRows;
  try {
    var parsed = getParsedRootFromRawSheet_();
    var cv = parsed.containerVersion;
    validateContainerCore_(cv);

    var ctx = buildContext_(cv);
    previewRows = buildPreviewRows_(ctx);
  } catch (err) {
    previewRows = [
      previewRowFromOperation_({
        severity: 'ERROR',
        source: 'System',
        action: '',
        assignmentType: '',
        tagId: '',
        triggerId: '',
        beforeIds: [],
        afterIds: [],
        message: err.message || String(err),
        apply: false,
        operationKey: 'ERROR|SYSTEM|' + Utilities.getUuid()
      }, null, null)
    ];
  }

  writePreviewSheet_(previewRows);

  var counts = summarizePreviewSeverities_(previewRows);
  SpreadsheetApp.getUi().alert(
    'Preview built.\n\n' +
      'INFO: ' + counts.INFO + '\n' +
      'WARNING: ' + counts.WARNING + '\n' +
      'ERROR: ' + counts.ERROR + '\n\n' +
      'Review Edit Preview before applying edits.'
  );
}

function applyEditsAndCreateExportJson() {
  ensureCoreSheets_();

  var parsed = getParsedRootFromRawSheet_();
  var root = parsed.root;
  var cv = parsed.containerVersion;
  validateContainerCore_(cv);

  var previewSheet = getOrCreateSheet_(APP.SHEETS.EDIT_PREVIEW);
  var data = getSheetData_(previewSheet);
  if (!data.rows.length) {
    throw new Error('Edit Preview is empty. Run "Build Preview" first.');
  }

  var headers = data.headers;
  var h = indexHeaders_(headers);

  var applicableOps = [];
  var errorsSkipped = 0;
  var warningsApplied = 0;

  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    var severity = String(row[h['Severity']] || '').toUpperCase();
    var apply = toBoolean_(row[h['Apply']]);

    if (severity === 'ERROR') {
      errorsSkipped++;
      continue;
    }
    if (!apply) {
      continue;
    }

    var op = {
      action: String(row[h['Action']] || '').toUpperCase(),
      assignmentType: String(row[h['Assignment Type']] || '').toUpperCase(),
      tagId: toId_(row[h['Tag ID']]),
      triggerId: toId_(row[h['Trigger ID']]),
      operationKey: String(row[h['Operation Key']] || '')
    };

    if (!op.action || !op.assignmentType || !op.tagId || !op.triggerId) {
      continue;
    }

    if (severity === 'WARNING') warningsApplied++;
    applicableOps.push(op);
  }

  if (!applicableOps.length) {
    throw new Error('No applicable preview operations selected. Nothing to apply.');
  }

  var applyResult = applyOperationsToContainer_(cv, applicableOps);

  var jsonOut = JSON.stringify(root, null, 2);
  var driveFile = createDriveJsonFile_(jsonOut);
  writeExportSheet_(jsonOut, {
    originalTagCount: applyResult.originalTagCount,
    modifiedTagCount: applyResult.modifiedTagCount,
    operationsApplied: applyResult.operationsApplied,
    warningsApplied: warningsApplied,
    errorsSkipped: errorsSkipped,
    driveFileUrl: driveFile.getUrl()
  });

  SpreadsheetApp.getUi().alert(
    'Export JSON created successfully.\n\n' +
      'Operations Applied: ' + applyResult.operationsApplied + '\n' +
      'Warnings Applied: ' + warningsApplied + '\n' +
      'Errors Skipped: ' + errorsSkipped + '\n' +
      'Drive File: ' + driveFile.getUrl() + '\n\n' +
      'GTM import reminder:\n' +
      'Import into a NEW workspace, choose Merge, choose overwrite conflicting tags/triggers/variables, then review GTM detailed changes before confirming.'
  );
}

function resetEditorWorkspace() {
  ensureCoreSheets_();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    'Reset Editor Workspace',
    'Choose YES to clear generated sheets and RAW_JSON.\nChoose NO to clear generated sheets but keep RAW_JSON.\nChoose CANCEL to abort.',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.CANCEL) return;

  var keepRaw = response === ui.Button.NO;
  var namesToDelete = [
    APP.SHEETS.CONTAINER_INFO,
    APP.SHEETS.TRIGGER_DIR,
    APP.SHEETS.TAG_DIR,
    APP.SHEETS.ASSIGN_CURRENT,
    APP.SHEETS.ASSIGN_ADD,
    APP.SHEETS.BULK_RULES,
    APP.SHEETS.EDIT_PREVIEW,
    APP.SHEETS.EXPORT_JSON
  ];

  for (var i = 0; i < namesToDelete.length; i++) {
    safeDeleteSheet_(namesToDelete[i]);
  }

  rebuildReadMe();

  if (!keepRaw) {
    var rawSheet = getOrCreateSheet_(APP.SHEETS.RAW_JSON, APP.TAB_COLORS.GRAY);
    rawSheet.clear();
    setupRawJsonSheet_(rawSheet);
  }

  ui.alert('Workspace reset complete.');
}

function rebuildReadMe() {
  var sheet = getOrCreateSheet_(APP.SHEETS.READ_ME, APP.TAB_COLORS.GRAY);
  sheet.clear();

  var lines = [
    ['GTM Bulk Trigger Assignment Editor - Read Me'],
    [''],
    ['Purpose'],
    ['This tool safely bulk-edits GTM tag trigger assignments from a pasted container export JSON.'],
    ['It only edits tag.firingTriggerId and tag.blockingTriggerId.'],
    [''],
    ['Workflow'],
    ['1) Export full container JSON from GTM.'],
    ['2) Use "GTM Bulk Editor -> Open JSON Loader" and paste JSON.'],
    ['3) Run "Build Editor Tabs."'],
    ['4) Use row-based edits and/or rule-based edits.'],
    ['5) Run "Build Preview."'],
    ['6) Review warnings and errors in Edit Preview.'],
    ['7) Run "Apply Edits & Create Export JSON."'],
    ['8) Import modified JSON into a NEW GTM workspace.'],
    ['9) Choose Merge and overwrite conflicting tags/triggers/variables.'],
    ['10) Review GTM detailed changes before confirming.'],
    [''],
    ['Important Notes'],
    ['- Original pasted JSON remains source of truth.'],
    ['- This tool does NOT delete triggers.'],
    ['- This tool does NOT publish changes.'],
    ['- Import into a NEW workspace for safe review.'],
    [''],
    ['MVP Limitations'],
    ['- No live GTM API calls.'],
    ['- No publishing workflow.'],
    ['- No trigger deletion module.'],
    ['- No edits to trigger definitions, tag parameters, variables, folders, templates, consent settings.'],
    ['- Server-side GTM differences are only supported if export JSON is structurally compatible.'],
    ['- Ambiguous trigger names are not auto-resolved. Use Trigger ID.']
  ];

  sheet.getRange(1, 1, lines.length, 1).setValues(lines);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sheet.setColumnWidth(1, 980);
  sheet.getRange(1, 1, lines.length, 1).setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setTabColor(APP.TAB_COLORS.GRAY);
}

// ---------------------------
// Build Tabs
// ---------------------------

function buildContainerInfoSheet_(root, cv) {
  var sheet = getOrCreateSheet_(APP.SHEETS.CONTAINER_INFO, APP.TAB_COLORS.BLUE);
  sheet.clear();

  var info = [
    ['Field', 'Value'],
    ['Account ID', asValue_(cv.accountId)],
    ['Container ID', asValue_(cv.containerId)],
    ['Container Name', asValue_(cv.container && cv.container.name)],
    ['Public ID', asValue_(cv.container && cv.container.publicId)],
    ['Container Version ID', asValue_(cv.containerVersionId)],
    ['Container Version Name', asValue_(cv.name)],
    ['Export Time', asValue_(root.exportTime || root.exportTimestamp || cv.exportTime || '')],
    ['Number of Tags', asArray_(cv.tag).length],
    ['Number of Triggers', asArray_(cv.trigger).length],
    ['Number of Variables', asArray_(cv.variable).length],
    ['Number of Folders', asArray_(cv.folder).length],
    ['Number of Templates', asArray_(cv.template).length],
    ['Number of Built-In Variables', asArray_(cv.builtInVariable).length]
  ];

  sheet.getRange(1, 1, info.length, info[0].length).setValues(info);
  formatHeaderRow_(sheet, 1, 2);
  sheet.autoResizeColumns(1, 2);
}

function buildTriggerDirectorySheet_(cv) {
  var sheet = getOrCreateSheet_(APP.SHEETS.TRIGGER_DIR, APP.TAB_COLORS.BLUE);
  sheet.clear();

  var headers = [
    'Trigger ID',
    'Trigger Name',
    'Trigger Type',
    'Folder ID',
    'Folder Name',
    'Used As Firing Count',
    'Used As Blocking Count',
    'Total Assignment Count',
    'Notes'
  ];

  var tags = asArray_(cv.tag);
  var triggers = asArray_(cv.trigger);
  var folderMap = buildFolderMap_(cv);

  var firingCounts = {};
  var blockingCounts = {};
  for (var i = 0; i < tags.length; i++) {
    var t = tags[i];
    asArray_(t.firingTriggerId).forEach(function(id) {
      id = toId_(id);
      if (!id) return;
      firingCounts[id] = (firingCounts[id] || 0) + 1;
    });
    asArray_(t.blockingTriggerId).forEach(function(id) {
      id = toId_(id);
      if (!id) return;
      blockingCounts[id] = (blockingCounts[id] || 0) + 1;
    });
  }

  var nameFreq = {};
  for (var j = 0; j < triggers.length; j++) {
    var name = String(triggers[j].name || '');
    if (!name) continue;
    nameFreq[name] = (nameFreq[name] || 0) + 1;
  }

  var rows = [];
  for (var k = 0; k < triggers.length; k++) {
    var tr = triggers[k];
    var trigId = toId_(tr.triggerId || tr.tagManagerUrl || tr.name);
    var trigName = asValue_(tr.name);
    var folderId = toId_(tr.parentFolderId);
    var fCount = firingCounts[trigId] || 0;
    var bCount = blockingCounts[trigId] || 0;
    var total = fCount + bCount;

    var notes = [];
    if (nameFreq[trigName] > 1) notes.push('Duplicate trigger name detected');
    if (total === 0) notes.push('Unused trigger');

    rows.push([
      trigId,
      trigName,
      asValue_(tr.type),
      folderId,
      folderMap[folderId] || '',
      fCount,
      bCount,
      total,
      notes.join('; ')
    ]);
  }

  writeTableSheet_(sheet, headers, rows, APP.TAB_COLORS.BLUE);
}

function buildTagDirectorySheet_(cv) {
  var sheet = getOrCreateSheet_(APP.SHEETS.TAG_DIR, APP.TAB_COLORS.BLUE);
  sheet.clear();

  var headers = [
    'Tag ID',
    'Tag Name',
    'Tag Type',
    'Folder ID',
    'Folder Name',
    'Paused',
    'Firing Trigger IDs',
    'Firing Trigger Names',
    'Blocking Trigger IDs',
    'Blocking Trigger Names',
    'Firing Count',
    'Blocking Count',
    'Notes'
  ];

  var tags = asArray_(cv.tag);
  var triggerMap = buildTriggerMap_(cv);
  var folderMap = buildFolderMap_(cv);

  var rows = [];
  for (var i = 0; i < tags.length; i++) {
    var tag = tags[i];
    var tagId = toId_(tag.tagId);
    var firingIds = normalizeIdArray_(tag.firingTriggerId);
    var blockingIds = normalizeIdArray_(tag.blockingTriggerId);

    var firingNames = firingIds.map(function(id) {
      return triggerMap[id] ? triggerMap[id].name : '(missing trigger ' + id + ')';
    });

    var blockingNames = blockingIds.map(function(id) {
      return triggerMap[id] ? triggerMap[id].name : '(missing trigger ' + id + ')';
    });

    var paused = isTagPaused_(tag);
    var notes = [];
    if (!paused && firingIds.length === 0) {
      notes.push('Non-paused tag has zero firing triggers');
    }

    var folderId = toId_(tag.parentFolderId);

    rows.push([
      tagId,
      asValue_(tag.name),
      asValue_(tag.type),
      folderId,
      folderMap[folderId] || '',
      paused,
      firingIds.join(', '),
      firingNames.join(', '),
      blockingIds.join(', '),
      blockingNames.join(', '),
      firingIds.length,
      blockingIds.length,
      notes.join('; ')
    ]);
  }

  writeTableSheet_(sheet, headers, rows, APP.TAB_COLORS.BLUE);
}

function buildAssignmentsCurrentSheet_(cv) {
  var sheet = getOrCreateSheet_(APP.SHEETS.ASSIGN_CURRENT, APP.TAB_COLORS.YELLOW);
  sheet.clear();

  var headers = [
    'Select',
    'Action',
    'Assignment Type',
    'Tag ID',
    'Tag Name',
    'Tag Type',
    'Tag Paused',
    'Folder ID',
    'Folder Name',
    'Trigger ID',
    'Trigger Name',
    'Trigger Type',
    'Result',
    'Notes'
  ];

  var tags = asArray_(cv.tag);
  var triggerMap = buildTriggerMap_(cv);
  var folderMap = buildFolderMap_(cv);
  var rows = [];

  for (var i = 0; i < tags.length; i++) {
    var tag = tags[i];
    var tagId = toId_(tag.tagId);
    var paused = isTagPaused_(tag);
    var folderId = toId_(tag.parentFolderId);

    normalizeIdArray_(tag.firingTriggerId).forEach(function(triggerId) {
      var tr = triggerMap[triggerId] || {};
      rows.push([
        false,
        'Remove',
        'Firing',
        tagId,
        asValue_(tag.name),
        asValue_(tag.type),
        paused,
        folderId,
        folderMap[folderId] || '',
        triggerId,
        asValue_(tr.name),
        asValue_(tr.type),
        '',
        ''
      ]);
    });

    normalizeIdArray_(tag.blockingTriggerId).forEach(function(triggerId) {
      var tr = triggerMap[triggerId] || {};
      rows.push([
        false,
        'Remove',
        'Blocking',
        tagId,
        asValue_(tag.name),
        asValue_(tag.type),
        paused,
        folderId,
        folderMap[folderId] || '',
        triggerId,
        asValue_(tr.name),
        asValue_(tr.type),
        '',
        ''
      ]);
    });
  }

  writeTableSheet_(sheet, headers, rows, APP.TAB_COLORS.YELLOW);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 1).insertCheckboxes();
  }

  var actionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Remove'], true)
    .setAllowInvalid(false)
    .build();
  var assignTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Firing', 'Blocking'], true)
    .setAllowInvalid(false)
    .build();

  if (rows.length) {
    sheet.getRange(2, 2, rows.length, 1).setDataValidation(actionRule);
    sheet.getRange(2, 3, rows.length, 1).setDataValidation(assignTypeRule);
  }
}

function buildAssignmentsAddSheet_(cv) {
  var sheet = getOrCreateSheet_(APP.SHEETS.ASSIGN_ADD, APP.TAB_COLORS.YELLOW);
  sheet.clear();

  var headers = [
    'Select',
    'Assignment Type',
    'Tag ID',
    'Tag Name',
    'Tag Type',
    'Tag Paused',
    'Folder ID',
    'Folder Name',
    'Trigger ID to Add',
    'Trigger Name to Add',
    'Result',
    'Notes'
  ];

  var tags = asArray_(cv.tag);
  var folderMap = buildFolderMap_(cv);

  var rows = [];
  for (var i = 0; i < tags.length; i++) {
    var tag = tags[i];
    var folderId = toId_(tag.parentFolderId);

    rows.push([
      false,
      'Firing',
      toId_(tag.tagId),
      asValue_(tag.name),
      asValue_(tag.type),
      isTagPaused_(tag),
      folderId,
      folderMap[folderId] || '',
      '',
      '',
      '',
      ''
    ]);
  }

  writeTableSheet_(sheet, headers, rows, APP.TAB_COLORS.YELLOW);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 1).insertCheckboxes();
    var assignTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Firing', 'Blocking'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 2, rows.length, 1).setDataValidation(assignTypeRule);

    // Auto-fill Trigger Name from Trigger ID where possible.
    // Uses row-specific formulas to keep references aligned.
    var formulas = [];
    for (var r = 0; r < rows.length; r++) {
      var rowNum = r + 2;
      formulas.push([
        '=IF($I' + rowNum + '="","",IFERROR(VLOOKUP($I' + rowNum + ',\'' + APP.SHEETS.TRIGGER_DIR + '\'!A:B,2,FALSE),"(unknown trigger id)"))'
      ]);
    }
    sheet.getRange(2, 10, rows.length, 1).setFormulas(formulas);
  }
}

function buildBulkRulesSheet_() {
  var sheet = getOrCreateSheet_(APP.SHEETS.BULK_RULES, APP.TAB_COLORS.YELLOW);
  sheet.clear();

  var headers = [
    'Enabled',
    'Action',
    'Assignment Type',
    'Match Field',
    'Match Operator',
    'Match Value',
    'Target Trigger ID',
    'Target Trigger Name',
    'Notes',
    'Result'
  ];

  writeTableSheet_(sheet, headers, [], APP.TAB_COLORS.YELLOW);

  // Pre-fill starter rows for convenience.
  var starterRows = 30;
  if (starterRows > 0) {
    var values = [];
    for (var i = 0; i < starterRows; i++) {
      values.push([false, 'Remove', 'Firing', 'Tag Name', 'Contains', '', '', '', '', '']);
    }
    sheet.getRange(2, 1, starterRows, headers.length).setValues(values);
    sheet.getRange(2, 1, starterRows, 1).insertCheckboxes();

    var actionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Remove', 'Add'], true)
      .setAllowInvalid(false)
      .build();

    var assignmentTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Firing', 'Blocking'], true)
      .setAllowInvalid(false)
      .build();

    var matchFieldRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Trigger ID',
        'Trigger Name',
        'Tag ID',
        'Tag Name',
        'Tag Type',
        'Folder Name',
        'Paused'
      ], true)
      .setAllowInvalid(false)
      .build();

    var matchOpRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Equals', 'Contains', 'Starts With', 'Ends With', 'Regex'], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, 2, starterRows, 1).setDataValidation(actionRule);
    sheet.getRange(2, 3, starterRows, 1).setDataValidation(assignmentTypeRule);
    sheet.getRange(2, 4, starterRows, 1).setDataValidation(matchFieldRule);
    sheet.getRange(2, 5, starterRows, 1).setDataValidation(matchOpRule);

    var formulas = [];
    for (var r = 0; r < starterRows; r++) {
      var rowNum = r + 2;
      formulas.push([
        '=IF($G' + rowNum + '="","",IFERROR(VLOOKUP($G' + rowNum + ',\'' + APP.SHEETS.TRIGGER_DIR + '\'!A:B,2,FALSE),"(unknown trigger id)"))'
      ]);
    }
    sheet.getRange(2, 8, starterRows, 1).setFormulas(formulas);
  }
}

function resetPreviewSheet_() {
  writePreviewSheet_([]);
}

function resetExportSheet_() {
  var sheet = getOrCreateSheet_(APP.SHEETS.EXPORT_JSON, APP.TAB_COLORS.GREEN);
  sheet.clear();

  var headers = [
    'Export Status',
    'Export Created At',
    'Original Tag Count',
    'Modified Tag Count',
    'Operations Applied',
    'Warnings Applied',
    'Errors Skipped',
    'Drive File URL',
    'JSON Chunks'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sheet, 1, headers.length);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setTabColor(APP.TAB_COLORS.GREEN);
}

// ---------------------------
// Preview Engine
// ---------------------------

function buildContext_(cv) {
  var tags = asArray_(cv.tag);
  var triggers = asArray_(cv.trigger);
  var folders = asArray_(cv.folder);

  var tagMap = {};
  tags.forEach(function(tag) {
    tagMap[toId_(tag.tagId)] = tag;
  });

  var triggerMap = {};
  var triggerNameMap = {};
  var triggerNameDupes = {};

  triggers.forEach(function(tr) {
    var id = toId_(tr.triggerId);
    var name = String(tr.name || '');

    triggerMap[id] = tr;

    if (!triggerNameMap[name]) triggerNameMap[name] = [];
    triggerNameMap[name].push(id);
    if (triggerNameMap[name].length > 1) triggerNameDupes[name] = true;
  });

  var folderMap = {};
  folders.forEach(function(f) {
    folderMap[toId_(f.folderId)] = asValue_(f.name);
  });

  return {
    cv: cv,
    tags: tags,
    triggers: triggers,
    tagMap: tagMap,
    triggerMap: triggerMap,
    triggerNameMap: triggerNameMap,
    triggerNameDupes: triggerNameDupes,
    folderMap: folderMap
  };
}

function buildPreviewRows_(ctx) {
  var rawOps = [];

  rawOps = rawOps.concat(readManualRemoveOps_(ctx));
  rawOps = rawOps.concat(readManualAddOps_(ctx));
  rawOps = rawOps.concat(readBulkRuleOps_(ctx));

  // De-dupe operations by key while preserving source trail.
  var byKey = {};
  for (var i = 0; i < rawOps.length; i++) {
    var op = rawOps[i];
    var key = op.operationKey || buildOperationKey_(op.action, op.assignmentType, op.tagId, op.triggerId);
    op.operationKey = key;

    if (!byKey[key]) {
      byKey[key] = op;
      byKey[key].sources = {};
      byKey[key].sources[op.source] = true;
    } else {
      byKey[key].sources[op.source] = true;
      if (!byKey[key].message && op.message) byKey[key].message = op.message;
      if (op.severityPriority > byKey[key].severityPriority) {
        byKey[key].severityPriority = op.severityPriority;
      }
    }
  }

  var deduped = Object.keys(byKey).map(function(k) {
    var item = byKey[k];
    item.source = Object.keys(item.sources).join(' + ');
    return item;
  });

  // Conflict detection: same tag + type + trigger has both ADD and REMOVE.
  var conflictMap = {};
  deduped.forEach(function(op) {
    var key = op.assignmentType + '|' + op.tagId + '|' + op.triggerId;
    if (!conflictMap[key]) conflictMap[key] = {};
    conflictMap[key][op.action] = true;
  });

  deduped.forEach(function(op) {
    var key = op.assignmentType + '|' + op.tagId + '|' + op.triggerId;
    if (conflictMap[key].ADD && conflictMap[key].REMOVE) {
      op.forceError = 'Conflicting operations: same trigger is both added and removed for this tag/assignment type.';
    }
  });

  // Build display rows with before/after arrays and validation severity.
  var rows = deduped.map(function(op) {
    return previewRowFromOperation_(op, ctx.tagMap[op.tagId], ctx.triggerMap[op.triggerId]);
  });

  // Sort by severity then source then tag.
  rows.sort(function(a, b) {
    var order = { 'ERROR': 0, 'WARNING': 1, 'INFO': 2 };
    var sa = order[a[0]]; // Severity
    var sb = order[b[0]];
    if (sa !== sb) return sa - sb;
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]));
    return String(a[4]).localeCompare(String(b[4]));
  });

  return rows;
}

function readManualRemoveOps_(ctx) {
  var sheet = getOrCreateSheet_(APP.SHEETS.ASSIGN_CURRENT, APP.TAB_COLORS.YELLOW);
  var data = getSheetData_(sheet);
  var h = indexHeaders_(data.headers);
  var ops = [];

  data.rows.forEach(function(row) {
    var selected = toBoolean_(row[h['Select']]);
    if (!selected) return;

    var action = String(row[h['Action']] || 'Remove').toUpperCase();
    var assignmentType = String(row[h['Assignment Type']] || '').toUpperCase();
    var tagId = toId_(row[h['Tag ID']]);
    var triggerId = toId_(row[h['Trigger ID']]);

    ops.push({
      source: APP.SHEETS.ASSIGN_CURRENT,
      action: action,
      assignmentType: assignmentType,
      tagId: tagId,
      triggerId: triggerId,
      severityPriority: 0,
      message: ''
    });
  });

  return ops;
}

function readManualAddOps_(ctx) {
  var sheet = getOrCreateSheet_(APP.SHEETS.ASSIGN_ADD, APP.TAB_COLORS.YELLOW);
  var data = getSheetData_(sheet);
  var h = indexHeaders_(data.headers);
  var ops = [];

  data.rows.forEach(function(row) {
    var selected = toBoolean_(row[h['Select']]);
    if (!selected) return;

    var assignmentType = String(row[h['Assignment Type']] || '').toUpperCase();
    var tagId = toId_(row[h['Tag ID']]);
    var triggerId = toId_(row[h['Trigger ID to Add']]);

    ops.push({
      source: APP.SHEETS.ASSIGN_ADD,
      action: 'ADD',
      assignmentType: assignmentType,
      tagId: tagId,
      triggerId: triggerId,
      severityPriority: 0,
      message: ''
    });
  });

  return ops;
}

function readBulkRuleOps_(ctx) {
  var sheet = getOrCreateSheet_(APP.SHEETS.BULK_RULES, APP.TAB_COLORS.YELLOW);
  var data = getSheetData_(sheet);
  var h = indexHeaders_(data.headers);
  var ops = [];

  data.rows.forEach(function(row, idx) {
    var enabled = toBoolean_(row[h['Enabled']]);
    if (!enabled) return;

    var action = String(row[h['Action']] || '').trim();
    var assignmentType = String(row[h['Assignment Type']] || '').trim();
    var matchField = String(row[h['Match Field']] || '').trim();
    var operator = String(row[h['Match Operator']] || '').trim();
    var matchValue = String(row[h['Match Value']] || '');
    var targetTriggerId = toId_(row[h['Target Trigger ID']]);
    var targetTriggerName = String(row[h['Target Trigger Name']] || '').trim();

    var ruleSource = APP.SHEETS.BULK_RULES + ' (row ' + (idx + 2) + ')';

    if (!action || !assignmentType || !matchField || !operator) {
      ops.push(ruleErrorOp_(ruleSource, 'Rule missing Action, Assignment Type, Match Field, or Match Operator.'));
      return;
    }

    action = action.toUpperCase();
    assignmentType = assignmentType.toUpperCase();

    // Validate name ambiguity if using trigger name references.
    if (matchField === 'Trigger Name' && ctx.triggerNameDupes[matchValue]) {
      ops.push(ruleErrorOp_(ruleSource, 'Rule references duplicate Trigger Name "' + matchValue + '". Use Trigger ID instead.'));
      return;
    }

    var resolvedTargetTriggerId = targetTriggerId;
    if (action === 'ADD') {
      if (!resolvedTargetTriggerId && targetTriggerName) {
        var ids = ctx.triggerNameMap[targetTriggerName] || [];
        if (ids.length > 1) {
          ops.push(ruleErrorOp_(ruleSource, 'Target Trigger Name is duplicated in container. Use Target Trigger ID.'));
          return;
        }
        resolvedTargetTriggerId = ids.length === 1 ? ids[0] : '';
      }

      if (!resolvedTargetTriggerId) {
        ops.push(ruleErrorOp_(ruleSource, 'Add rule requires Target Trigger ID (or unique Target Trigger Name).'));
        return;
      }
    }

    if (action === 'REMOVE') {
      buildOpsFromRemoveRule_(ctx, {
        source: ruleSource,
        assignmentType: assignmentType,
        matchField: matchField,
        operator: operator,
        matchValue: matchValue,
        targetTriggerId: resolvedTargetTriggerId
      }).forEach(function(op) {
        ops.push(op);
      });
      return;
    }

    if (action === 'ADD') {
      buildOpsFromAddRule_(ctx, {
        source: ruleSource,
        assignmentType: assignmentType,
        matchField: matchField,
        operator: operator,
        matchValue: matchValue,
        targetTriggerId: resolvedTargetTriggerId
      }).forEach(function(op) {
        ops.push(op);
      });
      return;
    }

    ops.push(ruleErrorOp_(ruleSource, 'Unsupported rule Action: ' + action));
  });

  return ops;
}

function buildOpsFromRemoveRule_(ctx, rule) {
  var ops = [];

  ctx.tags.forEach(function(tag) {
    var tagId = toId_(tag.tagId);
    var assignmentIds = rule.assignmentType === 'FIRING'
      ? normalizeIdArray_(tag.firingTriggerId)
      : normalizeIdArray_(tag.blockingTriggerId);

    assignmentIds.forEach(function(triggerId) {
      if (rule.targetTriggerId && rule.targetTriggerId !== triggerId) return;

      var fieldValue = getMatchFieldValue_(ctx, tag, triggerId, rule.matchField, rule.assignmentType);
      if (!matchByOperator_(fieldValue, rule.operator, rule.matchValue)) return;

      ops.push({
        source: rule.source,
        action: 'REMOVE',
        assignmentType: rule.assignmentType,
        tagId: tagId,
        triggerId: triggerId,
        severityPriority: 0,
        message: ''
      });
    });
  });

  if (!ops.length) {
    ops.push({
      source: rule.source,
      action: 'REMOVE',
      assignmentType: rule.assignmentType,
      tagId: '',
      triggerId: '',
      severityPriority: 3,
      message: 'Rule matched no current assignments.',
      forceWarningOnly: true,
      operationKey: 'RULE_NO_MATCH|' + Utilities.getUuid()
    });
  }

  return ops;
}

function buildOpsFromAddRule_(ctx, rule) {
  var ops = [];

  ctx.tags.forEach(function(tag) {
    var fieldValue = getMatchFieldValue_(ctx, tag, '', rule.matchField, rule.assignmentType);
    if (!matchByOperator_(fieldValue, rule.operator, rule.matchValue)) return;

    ops.push({
      source: rule.source,
      action: 'ADD',
      assignmentType: rule.assignmentType,
      tagId: toId_(tag.tagId),
      triggerId: toId_(rule.targetTriggerId),
      severityPriority: 0,
      message: ''
    });
  });

  if (!ops.length) {
    ops.push({
      source: rule.source,
      action: 'ADD',
      assignmentType: rule.assignmentType,
      tagId: '',
      triggerId: toId_(rule.targetTriggerId),
      severityPriority: 3,
      message: 'Rule matched no tags.',
      forceWarningOnly: true,
      operationKey: 'RULE_NO_MATCH|' + Utilities.getUuid()
    });
  }

  return ops;
}

function ruleErrorOp_(source, msg) {
  return {
    source: source,
    action: 'ERROR',
    assignmentType: '',
    tagId: '',
    triggerId: '',
    severityPriority: 5,
    message: msg,
    forceError: msg,
    operationKey: 'RULE_ERROR|' + Utilities.getUuid()
  };
}

function previewRowFromOperation_(op, tag, trigger) {
  var action = String(op.action || '').toUpperCase();
  var assignmentType = String(op.assignmentType || '').toUpperCase();

  var beforeIds = [];
  var afterIds = null;
  var severity = 'INFO';
  var message = op.message || 'Valid operation.';
  var apply = true;

  if (!tag && action !== 'ERROR') {
    severity = 'ERROR';
    message = 'Tag ID not found: ' + op.tagId;
    apply = false;
  }

  if (op.forceError) {
    severity = 'ERROR';
    message = op.forceError;
    apply = false;
  } else if (op.forceWarningOnly) {
    severity = 'WARNING';
    message = op.message || 'Rule produced no matching rows.';
    apply = true;
  }

  if (tag && !op.forceError && !op.forceWarningOnly) {
    var current = assignmentType === 'BLOCKING'
      ? normalizeIdArray_(tag.blockingTriggerId)
      : normalizeIdArray_(tag.firingTriggerId);

    beforeIds = current.slice();

    if (action === 'REMOVE') {
      if (!op.triggerId) {
        severity = 'ERROR';
        message = 'Missing Trigger ID for remove operation.';
        apply = false;
      } else if (current.indexOf(op.triggerId) === -1) {
        severity = 'WARNING';
        message = 'Trigger is not currently assigned to this tag.';
        afterIds = current.slice();
      } else {
        afterIds = removeId_(current, op.triggerId);
        if (assignmentType === 'FIRING' && !isTagPaused_(tag) && afterIds.length === 0) {
          severity = 'WARNING';
          message = 'Removing last firing trigger from a non-paused tag.';
        } else if (isTagPaused_(tag)) {
          severity = 'INFO';
          message = 'Editing paused tag.';
        } else {
          severity = 'INFO';
          message = 'Remove operation is valid.';
        }
      }
    } else if (action === 'ADD') {
      if (!op.triggerId) {
        severity = 'ERROR';
        message = 'Missing Trigger ID to add.';
        apply = false;
      } else if (!trigger) {
        severity = 'ERROR';
        message = 'Trigger ID to add does not exist: ' + op.triggerId;
        apply = false;
      } else if (current.indexOf(op.triggerId) !== -1) {
        severity = 'WARNING';
        message = 'Trigger is already assigned to this tag.';
        afterIds = current.slice();
      } else {
        afterIds = addIdNoDup_(current, op.triggerId);
        if (isTagPaused_(tag)) {
          severity = 'INFO';
          message = 'Editing paused tag.';
        } else {
          severity = 'INFO';
          message = 'Add operation is valid.';
        }
      }
    } else if (action === 'ERROR') {
      severity = 'ERROR';
      message = op.message || 'Rule validation error.';
      apply = false;
    } else {
      severity = 'ERROR';
      message = 'Unsupported action in operation: ' + action;
      apply = false;
    }

    if (afterIds === null && (action === 'ADD' || action === 'REMOVE')) {
      afterIds = current.slice();
    }
  }

  if (afterIds === null) afterIds = [];

  if (severity === 'ERROR') apply = false;

  return [
    severity,
    op.source || '',
    action,
    titleCase_(assignmentType),
    op.tagId || '',
    tag ? asValue_(tag.name) : '',
    tag ? asValue_(tag.type) : '',
    tag ? isTagPaused_(tag) : '',
    op.triggerId || '',
    trigger ? asValue_(trigger.name) : '',
    beforeIds.join(', '),
    afterIds.join(', '),
    message,
    apply,
    op.operationKey || buildOperationKey_(action, assignmentType, op.tagId, op.triggerId)
  ];
}

function writePreviewSheet_(rows) {
  var sheet = getOrCreateSheet_(APP.SHEETS.EDIT_PREVIEW, APP.TAB_COLORS.PURPLE);
  sheet.clear();

  var headers = [
    'Severity',
    'Source',
    'Action',
    'Assignment Type',
    'Tag ID',
    'Tag Name',
    'Tag Type',
    'Tag Paused',
    'Trigger ID',
    'Trigger Name',
    'Before Trigger IDs',
    'After Trigger IDs',
    'Message',
    'Apply',
    'Operation Key'
  ];

  writeTableSheet_(sheet, headers, rows, APP.TAB_COLORS.PURPLE);

  if (rows.length) {
    // Severity validation (display only)
    var sevRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['INFO', 'WARNING', 'ERROR'], true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 1, rows.length, 1).setDataValidation(sevRule);

    // Apply checkbox
    sheet.getRange(2, 14, rows.length, 1).insertCheckboxes();

    // Default Apply: true for INFO/WARNING, false for ERROR
    var applyValues = rows.map(function(r) {
      return [String(r[0] || '').toUpperCase() !== 'ERROR' && toBoolean_(r[13])];
    });
    sheet.getRange(2, 14, rows.length, 1).setValues(applyValues);
  }
}

function summarizePreviewSeverities_(rows) {
  var out = { INFO: 0, WARNING: 0, ERROR: 0 };
  rows.forEach(function(r) {
    var sev = String(r[0] || '').toUpperCase();
    if (out.hasOwnProperty(sev)) out[sev]++;
  });
  return out;
}

// ---------------------------
// Apply + Export
// ---------------------------

function applyOperationsToContainer_(cv, operations) {
  var tagMap = buildTagMap_(cv);

  var applied = 0;
  operations.forEach(function(op) {
    var tag = tagMap[toId_(op.tagId)];
    if (!tag) return;

    var prop = String(op.assignmentType || '').toUpperCase() === 'BLOCKING'
      ? 'blockingTriggerId'
      : 'firingTriggerId';

    var arr = normalizeIdArray_(tag[prop]);

    if (String(op.action || '').toUpperCase() === 'ADD') {
      arr = addIdNoDup_(arr, toId_(op.triggerId));
      tag[prop] = arr;
      applied++;
      return;
    }

    if (String(op.action || '').toUpperCase() === 'REMOVE') {
      arr = removeId_(arr, toId_(op.triggerId));
      tag[prop] = arr;
      applied++;
    }
  });

  return {
    originalTagCount: asArray_(cv.tag).length,
    modifiedTagCount: asArray_(cv.tag).length,
    operationsApplied: applied
  };
}

function writeExportSheet_(jsonOut, summary) {
  var sheet = getOrCreateSheet_(APP.SHEETS.EXPORT_JSON, APP.TAB_COLORS.GREEN);
  sheet.clear();

  var headers = [
    'Export Status',
    'Export Created At',
    'Original Tag Count',
    'Modified Tag Count',
    'Operations Applied',
    'Warnings Applied',
    'Errors Skipped',
    'Drive File URL',
    'JSON Chunks'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sheet, 1, headers.length);

  var statusRow = [
    'SUCCESS',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    summary.originalTagCount,
    summary.modifiedTagCount,
    summary.operationsApplied,
    summary.warningsApplied,
    summary.errorsSkipped,
    summary.driveFileUrl,
    ''
  ];
  sheet.getRange(2, 1, 1, statusRow.length).setValues([statusRow]);

  var chunks = chunkText_(jsonOut, APP.CHUNK_SIZE);
  var chunkRows = chunks.map(function(c) {
    return [protectChunkForSheet_(c)];
  });

  if (chunkRows.length) {
    sheet.getRange(4, 9, chunkRows.length, 1).setValues(chunkRows);
    sheet.getRange(4, 9, chunkRows.length, 1).setNumberFormat('@STRING@');
  }

  sheet.setFrozenRows(1);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(2, 3 + chunkRows.length), headers.length).createFilter();
  sheet.autoResizeColumns(1, headers.length);
  sheet.setTabColor(APP.TAB_COLORS.GREEN);
}

function createDriveJsonFile_(jsonOut) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  var name = 'gtm-bulk-edited-container-' + stamp + '.json';
  return DriveApp.createFile(name, jsonOut, MimeType.PLAIN_TEXT);
}

// ---------------------------
// Matching and Rule Helpers
// ---------------------------

function getMatchFieldValue_(ctx, tag, triggerId, matchField, assignmentType) {
  var field = String(matchField || '').trim();
  var folderId = toId_(tag.parentFolderId);
  var folderName = ctx.folderMap[folderId] || '';

  if (field === 'Tag ID') return toId_(tag.tagId);
  if (field === 'Tag Name') return asValue_(tag.name);
  if (field === 'Tag Type') return asValue_(tag.type);
  if (field === 'Folder Name') return folderName;
  if (field === 'Paused') return String(isTagPaused_(tag));

  if (field === 'Trigger ID') {
    if (triggerId) return toId_(triggerId);
    var arr = assignmentType === 'BLOCKING'
      ? normalizeIdArray_(tag.blockingTriggerId)
      : normalizeIdArray_(tag.firingTriggerId);
    return arr.join(',');
  }

  if (field === 'Trigger Name') {
    if (triggerId) {
      var t = ctx.triggerMap[toId_(triggerId)];
      return t ? asValue_(t.name) : '';
    }
    var ids = assignmentType === 'BLOCKING'
      ? normalizeIdArray_(tag.blockingTriggerId)
      : normalizeIdArray_(tag.firingTriggerId);
    return ids.map(function(id) {
      return ctx.triggerMap[id] ? asValue_(ctx.triggerMap[id].name) : '';
    }).join(',');
  }

  return '';
}

function matchByOperator_(actualValue, operator, expectedValue) {
  var actual = String(actualValue == null ? '' : actualValue);
  var expected = String(expectedValue == null ? '' : expectedValue);
  var op = String(operator || '').trim();

  if (op === 'Equals') return actual.toLowerCase() === expected.toLowerCase();
  if (op === 'Contains') return actual.toLowerCase().indexOf(expected.toLowerCase()) !== -1;
  if (op === 'Starts With') return actual.toLowerCase().indexOf(expected.toLowerCase()) === 0;
  if (op === 'Ends With') {
    var a = actual.toLowerCase();
    var e = expected.toLowerCase();
    return a.slice(Math.max(0, a.length - e.length)) === e;
  }
  if (op === 'Regex') {
    try {
      var re = new RegExp(expected, 'i');
      return re.test(actual);
    } catch (err) {
      return false;
    }
  }

  return false;
}

// ---------------------------
// JSON + Sheet Helpers
// ---------------------------

function ensureCoreSheets_() {
  rebuildReadMeIfMissing_();
  var raw = getOrCreateSheet_(APP.SHEETS.RAW_JSON, APP.TAB_COLORS.GRAY);
  if (raw.getLastRow() === 0 || !raw.getRange(1, 1).getValue()) {
    setupRawJsonSheet_(raw);
  }
}

function rebuildReadMeIfMissing_() {
  var ss = SpreadsheetApp.getActive();
  var readMe = ss.getSheetByName(APP.SHEETS.READ_ME);
  if (!readMe) rebuildReadMe();
}

function setupRawJsonSheet_(sheet) {
  sheet.clear();
  var headers = [
    'JSON Chunks',
    'Chunk Index',
    'Created At',
    'Notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sheet, 1, headers.length);
  sheet.getRange('A:A').setNumberFormat('@STRING@');
  sheet.setFrozenRows(1);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  sheet.autoResizeColumns(1, headers.length);
  sheet.setTabColor(APP.TAB_COLORS.GRAY);
}

function writeRawJsonChunks_(rawJsonText) {
  var sheet = getOrCreateSheet_(APP.SHEETS.RAW_JSON, APP.TAB_COLORS.GRAY);
  sheet.clear();
  setupRawJsonSheet_(sheet);

  var chunks = chunkText_(String(rawJsonText), APP.CHUNK_SIZE);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  var rows = chunks.map(function(chunk, i) {
    return [protectChunkForSheet_(chunk), i + 1, now, 'Pasted via JSON Loader'];
  });

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@STRING@');
  }

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(2, rows.length + 1), 4).createFilter();
  sheet.autoResizeColumns(1, 4);
}

function getRawJsonString_() {
  var sheet = getOrCreateSheet_(APP.SHEETS.RAW_JSON, APP.TAB_COLORS.GRAY);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('RAW_JSON is empty. Use "Open JSON Loader" to paste a GTM export JSON first.');
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  var chunks = values
    .map(function(r) { return String(r[0] || ''); })
    .filter(function(v) { return v !== ''; })
    .map(unprotectChunkFromSheet_);

  if (!chunks.length) {
    throw new Error('RAW_JSON has no chunk values in column A.');
  }

  return chunks.join('');
}

function getParsedRootFromRawSheet_() {
  var raw = getRawJsonString_();

  var root;
  try {
    root = JSON.parse(raw);
  } catch (err) {
    throw new Error('Failed to parse RAW_JSON chunks into valid JSON: ' + err.message);
  }

  var cv = getContainerVersion_(root);
  if (!cv) {
    throw new Error('No containerVersion found at root.containerVersion or root.');
  }

  return {
    root: root,
    containerVersion: cv
  };
}

function getContainerVersion_(root) {
  if (!root || typeof root !== 'object') return null;

  if (root.containerVersion && typeof root.containerVersion === 'object') {
    return root.containerVersion;
  }

  // Tolerate direct containerVersion shape as root.
  if (root.tag || root.trigger || root.variable || root.folder || root.containerId) {
    return root;
  }

  return null;
}

function validateContainerCore_(cv) {
  if (!cv) throw new Error('Container JSON is missing containerVersion.');
  if (!cv.tag) throw new Error('containerVersion.tag is missing.');
  if (!cv.trigger) throw new Error('containerVersion.trigger is missing.');
}

function buildTagMap_(cv) {
  var map = {};
  asArray_(cv.tag).forEach(function(tag) {
    map[toId_(tag.tagId)] = tag;
  });
  return map;
}

function buildTriggerMap_(cv) {
  var map = {};
  asArray_(cv.trigger).forEach(function(trigger) {
    map[toId_(trigger.triggerId)] = {
      id: toId_(trigger.triggerId),
      name: asValue_(trigger.name),
      type: asValue_(trigger.type),
      raw: trigger
    };
  });
  return map;
}

function buildFolderMap_(cv) {
  var map = {};
  asArray_(cv.folder).forEach(function(folder) {
    map[toId_(folder.folderId)] = asValue_(folder.name);
  });
  return map;
}

function getOrCreateSheet_(name, tabColor) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (tabColor) sheet.setTabColor(tabColor);
  return sheet;
}

function safeDeleteSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) return;
  ss.deleteSheet(sh);
}

function writeTableSheet_(sheet, headers, rows, tabColor) {
  var totalRows = Math.max(1, rows.length + 1);
  var totalCols = headers.length;

  sheet.getRange(1, 1, 1, totalCols).setValues([headers]);
  formatHeaderRow_(sheet, 1, totalCols);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, totalCols).setValues(rows);
  }

  sheet.setFrozenRows(1);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, totalRows, totalCols).createFilter();
  sheet.autoResizeColumns(1, totalCols);
  if (tabColor) sheet.setTabColor(tabColor);
}

function formatHeaderRow_(sheet, row, cols) {
  var r = sheet.getRange(row, 1, 1, cols);
  r.setFontWeight('bold')
    .setBackground(APP.HEADER_BG)
    .setFontColor(APP.HEADER_FONT)
    .setWrap(true);
}

function getSheetData_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    return { headers: [], rows: [] };
  }

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function(v) { return String(v || '').trim(); });
  var rows = values.slice(1);

  return { headers: headers, rows: rows };
}

function indexHeaders_(headers) {
  var idx = {};
  for (var i = 0; i < headers.length; i++) {
    idx[headers[i]] = i;
  }
  return idx;
}

function chunkText_(text, size) {
  var out = [];
  var str = String(text || '');
  if (!str.length) return [''];

  for (var i = 0; i < str.length; i += size) {
    out.push(str.slice(i, i + size));
  }
  return out;
}

function protectChunkForSheet_(chunk) {
  var c = String(chunk || '');
  // Guard against formula interpretation in Sheets.
  if (/^[=+\-@]/.test(c)) {
    return '\t' + c;
  }
  return c;
}

function unprotectChunkFromSheet_(chunk) {
  var c = String(chunk || '');
  if (/^\t[=+\-@]/.test(c)) {
    return c.slice(1);
  }
  return c;
}

function asArray_(value) {
  if (Array.isArray(value)) return value;
  if (value === null || value === undefined || value === '') return [];
  return [value];
}

function normalizeIdArray_(value) {
  return asArray_(value)
    .map(toId_)
    .filter(function(v) { return !!v; });
}

function addIdNoDup_(arr, id) {
  var out = normalizeIdArray_(arr);
  var s = toId_(id);
  if (!s) return out;
  if (out.indexOf(s) === -1) out.push(s);
  return out;
}

function removeId_(arr, id) {
  var s = toId_(id);
  return normalizeIdArray_(arr).filter(function(v) {
    return v !== s;
  });
}

function toId_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function asValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

function toBoolean_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}

function isTagPaused_(tag) {
  var v = tag && tag.paused;
  return v === true || String(v).toLowerCase() === 'true';
}

function titleCase_(value) {
  var s = String(value || '').toLowerCase();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function buildOperationKey_(action, assignmentType, tagId, triggerId) {
  return [
    String(action || '').toUpperCase(),
    String(assignmentType || '').toUpperCase(),
    toId_(tagId),
    toId_(triggerId)
  ].join('|');
}

function getLoaderSidebarHtml_() {
  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '  <meta charset="utf-8" />',
    '  <style>',
    '    body { font-family: Arial, sans-serif; margin: 12px; color: #1f2937; }',
    '    h2 { font-size: 15px; margin: 0 0 10px; }',
    '    p { font-size: 12px; margin: 0 0 10px; color: #4b5563; }',
    '    textarea { width: 100%; min-height: 340px; border: 1px solid #d1d5db; border-radius: 6px; padding: 8px; font-family: Consolas, monospace; font-size: 11px; }',
    '    .row { margin-top: 10px; display: flex; gap: 8px; }',
    '    button { border: 0; background: #2563eb; color: #fff; padding: 8px 12px; border-radius: 6px; cursor: pointer; font-size: 12px; }',
    '    button.secondary { background: #6b7280; }',
    '    .msg { margin-top: 10px; font-size: 12px; white-space: pre-wrap; }',
    '  </style>',
    '</head>',
    '<body>',
    '  <h2>GTM JSON Loader</h2>',
    '  <p>Paste the full GTM container export JSON, then click Save to RAW_JSON.</p>',
    '  <textarea id="jsonInput" placeholder="Paste GTM export JSON here..."></textarea>',
    '  <div class="row">',
    '    <button onclick="saveJson()">Save to RAW_JSON</button>',
    '    <button class="secondary" onclick="clearText()">Clear</button>',
    '  </div>',
    '  <div id="msg" class="msg"></div>',
    '  <script>',
    '    function setMsg(text, isError) {',
    '      var el = document.getElementById("msg");',
    '      el.style.color = isError ? "#b91c1c" : "#065f46";',
    '      el.textContent = text || "";',
    '    }',
    '    function clearText() {',
    '      document.getElementById("jsonInput").value = "";',
    '      setMsg("", false);',
    '    }',
    '    function saveJson() {',
    '      var text = document.getElementById("jsonInput").value || "";',
    '      setMsg("Saving...", false);',
    '      google.script.run',
    '        .withSuccessHandler(function(res) {',
    '          var msg = (res && res.message) ? res.message : "Saved.";',
    '          if (res) {',
    '            msg += "\\nTags: " + res.tagCount + " | Triggers: " + res.triggerCount;',
    '          }',
    '          setMsg(msg, false);',
    '        })',
    '        .withFailureHandler(function(err) {',
    '          setMsg((err && err.message) ? err.message : String(err), true);',
    '        })',
    '        .saveRawJsonFromSidebar(text);',
    '    }',
    '  </script>',
    '</body>',
    '</html>'
  ].join('\n');
}

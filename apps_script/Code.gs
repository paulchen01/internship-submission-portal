const CONFIG = {
  ROOT_FOLDER_ID: 'ROOT_FOLDER_ID',
  SPREADSHEET_ID: 'SPREADSHEET_ID',
  TEACHER_EMAILS: 'TEACHER_EMAILS',
  COURSE_TITLE: 'COURSE_TITLE'
};

const PRECONFIGURED = {
  ROOT_FOLDER_ID: '',
  SPREADSHEET_ID: '',
  TEACHER_EMAILS: ''
};

const TEST_STUDENT_EMAIL = '';
const SHARED_SOURCE_FILE_IDS = [];

const SHEETS = {
  submissions: 'Submissions',
  announcements: 'Announcements',
  students: 'Students',
  visits: 'Visits'
};

const DEFAULT_COURSE_TITLE = '實習作業上傳與評分系統';
const WEB_APP_URL = '';
const STATUS = ['已繳交', '退件', '已完成'];
const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;

const SUBMISSION_HEADERS = [
  'submissionId',
  'timestamp',
  'studentEmail',
  'studentUserKey',
  'studentName',
  'studentId',
  'className',
  'agency',
  'assignmentType',
  'week',
  'attempt',
  'status',
  'score',
  'teacherComment',
  'fileName',
  'fileUrl',
  'fileId',
  'mimeType',
  'reviewedFileName',
  'reviewedFileUrl',
  'reviewedFileId',
  'reviewedAt',
  'lastReviewedAt',
  'reviewerEmail',
  'agencySupervisor'
];

const ANNOUNCEMENT_HEADERS = [
  'announcementId',
  'updatedAt',
  'title',
  'body',
  'pinned',
  'visible',
  'sortOrder'
];

const STUDENT_HEADERS = [
  'studentKey',
  'studentId',
  'studentName',
  'degree',
  'className',
  'gender',
  'agency',
  'agencyAddress',
  'field',
  'supervisor',
  'studentEmail',
  'studentUserKey',
  'bindStatus',
  'boundAt',
  'notes'
];

const VISIT_HEADERS = [
  'visitId',
  'timestamp',
  'studentEmail',
  'studentUserKey',
  'studentName',
  'studentId',
  'className',
  'agency',
  'agencyAddress',
  'agencySupervisor',
  'visitDate',
  'startTime',
  'endTime',
  'contactName',
  'contactPhone',
  'studentNote',
  'status',
  'teacherNote',
  'lastReviewedAt',
  'reviewerEmail'
];

const DEFAULT_ROSTER_B64 = '';

const ASSIGNMENTS = [
  {
    type: '實習計畫書',
    label: '實習計畫書',
    requiresWeek: false,
    description: '請上傳完整實習計畫書。'
  },
  {
    type: '週誌',
    label: '8週週誌',
    requiresWeek: true,
    weeks: [1, 2, 3, 4, 5, 6, 7, 8],
    description: '請依週次上傳週誌。退件後請重新上傳修正版。'
  },
  {
    type: '讀書心得報告',
    label: '讀書心得報告',
    requiresWeek: false,
    description: '請上傳指定讀物或自選讀物的心得報告。'
  },
  {
    type: '個案分析報告',
    label: '個案分析報告',
    requiresWeek: false,
    description: '請上傳個案分析報告。'
  },
  {
    type: '實習總報告',
    label: '實習總報告 PDF',
    requiresWeek: false,
    description: '完成實習後，請上傳完整 PDF 實習總報告。'
  },
  {
    type: '實習照片1',
    label: '實習照片1',
    requiresWeek: false,
    description: '請上傳第1張學生實習照片。'
  },
  {
    type: '實習照片2',
    label: '實習照片2',
    requiresWeek: false,
    description: '請上傳第2張學生實習照片。'
  }
];

function doGet() {
  initializeConfigIfNeeded_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(DEFAULT_COURSE_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function runInitialSetup(teacherEmailsCsv) {
  const activeEmail = getCurrentUserEmail_();
  const teacherEmails = teacherEmailsCsv || activeEmail || PRECONFIGURED.TEACHER_EMAILS;

  if (!teacherEmails) {
    throw new Error('無法取得目前登入者 Email。請手動傳入老師 Email，例如 runInitialSetup("teacher@example.com")。');
  }

  const properties = PropertiesService.getScriptProperties();
  let rootFolderId = properties.getProperty(CONFIG.ROOT_FOLDER_ID);
  let spreadsheetId = properties.getProperty(CONFIG.SPREADSHEET_ID);

  if (!rootFolderId) {
    rootFolderId = PRECONFIGURED.ROOT_FOLDER_ID;
    if (!rootFolderId) {
      const root = DriveApp.createFolder('實習作業上傳');
      rootFolderId = root.getId();
    }
    properties.setProperty(CONFIG.ROOT_FOLDER_ID, rootFolderId);
  }

  if (!spreadsheetId) {
    spreadsheetId = PRECONFIGURED.SPREADSHEET_ID;
    if (!spreadsheetId) {
      const spreadsheet = SpreadsheetApp.create('實習作業上傳紀錄');
      spreadsheetId = spreadsheet.getId();
    }
    properties.setProperty(CONFIG.SPREADSHEET_ID, spreadsheetId);
  }

  properties.setProperty(CONFIG.TEACHER_EMAILS, teacherEmails);
  properties.setProperty(CONFIG.COURSE_TITLE, DEFAULT_COURSE_TITLE);

  const spreadsheet = getSpreadsheet_();
  ensureSheet_(spreadsheet, SHEETS.submissions, SUBMISSION_HEADERS);
  ensureSheet_(spreadsheet, SHEETS.announcements, ANNOUNCEMENT_HEADERS);
  ensureSheet_(spreadsheet, SHEETS.students, STUDENT_HEADERS);
  ensureSheet_(spreadsheet, SHEETS.visits, VISIT_HEADERS);
  seedAnnouncements_();
  seedDefaultStudents_();
  revokeDirectTestStudentShares_();

  return {
    rootFolderUrl: DriveApp.getFolderById(rootFolderId).getUrl(),
    spreadsheetUrl: spreadsheet.getUrl(),
    teacherEmails
  };
}

function getInitialData() {
  assertConfigured_();
  const userEmail = getCurrentUserEmail_();
  const userKey = getCurrentUserKey_();
  const teacher = isTeacher_(userEmail);

  return serializeForClient_({
    courseTitle: getCourseTitle_(),
    userEmail,
    userKey,
    isTeacher: teacher,
    assignments: ASSIGNMENTS,
    announcements: listAnnouncements_(teacher),
    myStudent: getStudentForCurrentUser_(),
    mySubmissions: [],
    myVisits: [],
    teacherData: null
  });
}

function uploadSubmission(payload) {
  assertConfigured_();
  validateUploadPayload_(payload);

  const loginEmail = getCurrentUserEmail_();
  const userKey = getCurrentUserKey_();
  const studentEmail = normalizeEmail_(payload.studentEmail || loginEmail);
  if (!studentEmail) {
    throw new Error('請填寫可收件的 Gmail，老師退件與回傳批改檔會使用這個信箱。');
  }
  if (!userKey) {
    throw new Error('無法取得登入身分。請確認 Web App 設定為需要 Google 帳號登入。');
  }

  const rosterStudent = assertAndBindRosterStudent_(payload, studentEmail, userKey);
  const storageStudentEmail = normalizeEmail_(rosterStudent.studentEmail || studentEmail);

  const assignment = ASSIGNMENTS.find(item => item.type === payload.assignmentType);
  if (!assignment) {
    throw new Error('作業類型不正確。');
  }
  validateAssignmentFileType_(assignment.type, payload.fileName, payload.mimeType);

  const week = assignment.requiresWeek ? Number(payload.week) : '';
  if (assignment.requiresWeek && (!week || week < 1 || week > 8)) {
    throw new Error('週誌必須選擇第1週到第8週。');
  }

  const decoded = Utilities.base64Decode(payload.fileData);
  if (decoded.length > MAX_UPLOAD_BYTES) {
    throw new Error('檔案超過25MB。請壓縮檔案或改用較小的 PDF/DOCX。');
  }

  const studentName = rosterStudent.studentName;
  const studentId = rosterStudent.studentId;
  const className = rosterStudent.className || String(payload.className || '').trim();
  const agency = rosterStudent.agency || String(payload.agency || '').trim();
  const agencySupervisor = String(payload.agencySupervisor || '').trim();
  const originalName = sanitizeFileName_(payload.fileName || 'submission');
  const mimeType = payload.mimeType || MimeType.PDF;
  const rootFolder = getRootFolder_();
  const studentFolder = getOrCreateFolder_(rootFolder, `${studentId}_${studentName}`);
  const assignmentFolder = getOrCreateFolder_(studentFolder, assignment.requiresWeek ? '週誌' : assignment.type);
  const targetFolder = assignment.requiresWeek ? getOrCreateFolder_(assignmentFolder, `第${week}週`) : assignmentFolder;
  const attempt = getNextAttempt_(storageStudentEmail, userKey, payload.assignmentType, week);
  const timestamp = new Date();
  const stamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const weekPart = assignment.requiresWeek ? `_第${week}週` : '';
  const fileName = `${studentId}_${studentName}_${payload.assignmentType}${weekPart}_第${attempt}次_${stamp}_${originalName}`;
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const file = targetFolder.createFile(blob);

  try {
    if (storageStudentEmail) file.addViewer(storageStudentEmail);
  } catch (error) {
    // 部分學校網域或個人帳戶權限可能不允許自動分享；不影響老師收件。
  }

  const row = [
    makeId_('SUB'),
    timestamp,
    storageStudentEmail,
    userKey,
    studentName,
    studentId,
    className,
    agency,
    payload.assignmentType,
    week,
    attempt,
    '已繳交',
    '',
    '',
    fileName,
    file.getUrl(),
    file.getId(),
    mimeType,
    '',
    '',
    '',
    '',
    '',
    '',
    agencySupervisor
  ];

  const sheet = getSheet_(SHEETS.submissions);
  sheet.appendRow(row);

  return serializeForClient_({
    ok: true,
    message: '上傳完成。',
    submission: rowToObject_(SUBMISSION_HEADERS, row)
  });
}

function getMySubmissions() {
  assertConfigured_();
  return listSubmissionsForCurrentUser_();
}

function getMyVisits() {
  assertConfigured_();
  return listVisitsForCurrentUser_();
}

function recoverMyIdentity(payload) {
  assertConfigured_();
  const userKey = getCurrentUserKey_();
  if (!userKey) {
    throw new Error('系統無法取得目前 Google 登入識別碼，請重新登入 Google 帳號後再試。');
  }

  const student = findStudentForRecovery_(payload || {});
  bindStudentIdentity_(student.studentKey, normalizeEmail_(student.studentEmail), userKey);
  student.studentUserKey = userKey;
  student.bindStatus = 'bound';
  student.boundAt = new Date();

  const submissions = getAllSubmissionObjects_()
    .filter(row => rowBelongsToCurrentStudent_(row, student, normalizeEmail_(student.studentEmail), userKey))
    .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
  const visits = getAllVisitObjects_()
    .filter(row => rowBelongsToCurrentStudent_(row, student, normalizeEmail_(student.studentEmail), userKey))
    .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  return serializeForClient_({
    ok: true,
    message: `已重新辨識為 ${student.studentName}，並讀取到 ${submissions.length} 筆上傳紀錄。`,
    myStudent: student,
    mySubmissions: submissions,
    myVisits: visits
  });
}

function getTeacherData(filters) {
  return getTeacherSubmissionData(filters);
}

function getTeacherSubmissionData(filters) {
  assertTeacher_();
  const rows = getAllSubmissionObjects_();
  const students = listStudents_();
  const filtered = applyTeacherFilters_(rows, filters || {});
  return serializeForClient_({
    submissions: filtered,
    summary: buildSummary_(rows),
    progress: buildProgressList_(rows, students),
    students
  });
}

function getVisitAdminData(filters) {
  assertTeacher_();
  const visits = getAllVisitObjects_();
  return serializeForClient_({
    visits: applyVisitFilters_(visits, filters || {}),
    visitSummary: buildVisitSummary_(visits),
    students: listStudents_()
  });
}

function getRosterData() {
  assertTeacher_();
  return serializeForClient_({
    students: listStudents_()
  });
}

function submitVisitPlan(payload) {
  assertConfigured_();
  validateVisitPayload_(payload);

  const loginEmail = getCurrentUserEmail_();
  const userKey = getCurrentUserKey_();
  const studentEmail = normalizeEmail_(payload.studentEmail || loginEmail);
  if (!studentEmail) {
    throw new Error('請填寫 Gmail，老師才能辨識訪視回報來源。');
  }
  if (!userKey) {
    throw new Error('系統無法取得登入識別，請確認 Web App 是以 Google 帳號登入後使用。');
  }

  const rosterStudent = assertAndBindRosterStudent_(payload, studentEmail, userKey);
  const row = [
    makeId_('VISIT'),
    new Date(),
    normalizeEmail_(rosterStudent.studentEmail || studentEmail),
    userKey,
    rosterStudent.studentName,
    rosterStudent.studentId,
    rosterStudent.className || String(payload.className || '').trim(),
    rosterStudent.agency || String(payload.agency || '').trim(),
    rosterStudent.agencyAddress || String(payload.agencyAddress || '').trim(),
    String(payload.agencySupervisor || rosterStudent.supervisor || '').trim(),
    String(payload.visitDate || '').trim(),
    String(payload.startTime || '').trim(),
    String(payload.endTime || '').trim(),
    String(payload.contactName || '').trim(),
    String(payload.contactPhone || '').trim(),
    String(payload.studentNote || '').trim(),
    '待老師確認',
    '',
    '',
    ''
  ];

  const sheet = getSheet_(SHEETS.visits);
  sheet.appendRow(row);

  return serializeForClient_({
    ok: true,
    message: '訪視時間已送出，請等待老師確認。',
    visit: rowToObject_(VISIT_HEADERS, row),
    myVisits: listVisitsForCurrentUser_()
  });
}

function updateVisitStatus(payload) {
  assertTeacher_();
  if (!payload || !payload.visitId) {
    throw new Error('缺少 visitId。');
  }
  const allowed = ['待老師確認', '已確認', '請學生修改'];
  if (allowed.indexOf(payload.status) === -1) {
    throw new Error('訪視狀態不正確。');
  }

  const sheet = getSheet_(SHEETS.visits);
  const values = sheet.getDataRange().getValues();
  const rowIndex = findRowIndexById_(values, payload.visitId, VISIT_HEADERS);
  if (rowIndex === -1) {
    throw new Error('找不到這筆訪視安排。');
  }

  const headerMap = headerIndexMap_(VISIT_HEADERS);
  sheet.getRange(rowIndex + 1, headerMap.status + 1).setValue(payload.status);
  sheet.getRange(rowIndex + 1, headerMap.teacherNote + 1).setValue(String(payload.teacherNote || '').trim());
  sheet.getRange(rowIndex + 1, headerMap.lastReviewedAt + 1).setValue(new Date());
  sheet.getRange(rowIndex + 1, headerMap.reviewerEmail + 1).setValue(getCurrentUserEmail_());

  return serializeForClient_({
    ok: true,
    message: payload.status === '已確認' ? '訪視安排已確認。' : '訪視安排狀態已更新。',
    visitAdminData: getVisitAdminData({})
  });
}

function updateReview(payload) {
  assertTeacher_();
  if (!payload || !payload.submissionId) {
    throw new Error('缺少 submissionId。');
  }
  if (STATUS.indexOf(payload.status) === -1) {
    throw new Error('狀態必須是：已繳交、退件、已完成。');
  }

  const score = payload.score === '' || payload.score === null || payload.score === undefined
    ? ''
    : Number(payload.score);
  if (score !== '' && (Number.isNaN(score) || score < 0 || score > 100)) {
    throw new Error('分數需為0到100。若不評分可留空。');
  }
  const teacherComment = String(payload.teacherComment || '').trim();
  if (payload.status === '退件' && !teacherComment) {
    throw new Error('退件時請填寫退件原因或評語，學生才知道要修改哪裡。');
  }

  const sheet = getSheet_(SHEETS.submissions);
  const values = sheet.getDataRange().getValues();
  const rowIndex = findRowIndexById_(values, payload.submissionId, SUBMISSION_HEADERS);
  if (rowIndex === -1) {
    throw new Error('找不到這筆上傳紀錄。');
  }

  const headerMap = headerIndexMap_(SUBMISSION_HEADERS);
  const rowObject = rowToObject_(SUBMISSION_HEADERS, values[rowIndex]);
  sheet.getRange(rowIndex + 1, headerMap.status + 1).setValue(payload.status);
  sheet.getRange(rowIndex + 1, headerMap.score + 1).setValue(score);
  sheet.getRange(rowIndex + 1, headerMap.teacherComment + 1).setValue(teacherComment);
  sheet.getRange(rowIndex + 1, headerMap.lastReviewedAt + 1).setValue(new Date());
  sheet.getRange(rowIndex + 1, headerMap.reviewerEmail + 1).setValue(getCurrentUserEmail_());

  let message = '評分狀態已更新。';
  if (payload.status === '退件') {
    try {
      sendReturnNotification_(rowObject, teacherComment);
      message = '退件完成，已寄出 Email 通知學生。';
    } catch (error) {
      message = `退件完成，但 Apps Script 自動寄信失敗：${error.message || error}。請在老師後台該筆退件旁點「用 Gmail 通知學生」手動寄出。`;
    }
  } else if (payload.status === '已完成') {
    message = '確認完成。';
  } else if (payload.status === '已繳交') {
    message = '評語已儲存。';
  }

  return serializeForClient_({
    ok: true,
    message,
    teacherData: getTeacherSubmissionData({})
  });
}

function sendReturnNotification_(submission, teacherComment) {
  const studentEmail = normalizeEmail_(submission.studentEmail);
  if (!studentEmail) {
    throw new Error('這筆作業沒有學生 Email。');
  }

  const weekText = submission.week ? `第${submission.week}週` : '';
  const assignmentText = [submission.assignmentType, weekText].filter(Boolean).join(' ');
  const subject = `[實習作業退件通知] ${assignmentText}`;
  const body = [
    `${submission.studentName || '同學'}您好：`,
    '',
    `您上傳的「${assignmentText}」已由老師退件。`,
    '',
    '退件原因／老師評語：',
    teacherComment,
    '',
    '請依照評語修正後，登入系統重新上傳修正版。',
    WEB_APP_URL,
    '',
    '此信件由實習作業上傳與評分系統自動寄出，請勿直接回覆本信。'
  ].join('\n');

  MailApp.sendEmail({
    to: studentEmail,
    subject,
    body,
    name: getCourseTitle_()
  });
}

function authorizeEmailSending() {
  return {
    remainingDailyQuota: MailApp.getRemainingDailyQuota(),
    message: 'Email 寄送權限已授權。'
  };
}

function testEmailNotification() {
  assertTeacher_();
  const recipient = getCurrentUserEmail_() || PRECONFIGURED.TEACHER_EMAILS.split(',')[0];
  if (!recipient) {
    throw new Error('無法取得測試收件者 Email，請確認老師帳號已登入。');
  }
  const quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) {
    throw new Error('今天的 Apps Script 寄信配額已用完，請明天再試。');
  }

  MailApp.sendEmail({
    to: recipient,
    subject: '[實習系統測試信] Email 通知功能檢查',
    body: [
      '這是一封實習作業上傳與評分系統的測試信。',
      '',
      '如果您收到這封信，代表 Apps Script 的 Email 寄送授權與配額目前可用。',
      '',
      `系統網址：${WEB_APP_URL}`,
      `剩餘寄信配額：${quota - 1}`
    ].join('\n'),
    name: getCourseTitle_()
  });

  return {
    ok: true,
    message: `測試信已寄出到 ${recipient}。若收件匣沒有看到，請檢查垃圾郵件或促銷分類。`,
    remainingDailyQuota: quota - 1
  };
}

function uploadReviewedFile(payload) {
  assertTeacher_();
  if (!payload || !payload.submissionId || !payload.fileData || !payload.fileName) {
    throw new Error('缺少回傳檔案資料。');
  }

  const decoded = Utilities.base64Decode(payload.fileData);
  if (decoded.length > MAX_UPLOAD_BYTES) {
    throw new Error('檔案超過25MB。請壓縮檔案或改用較小的 PDF/DOCX。');
  }

  const sheet = getSheet_(SHEETS.submissions);
  const values = sheet.getDataRange().getValues();
  const rowIndex = findRowIndexById_(values, payload.submissionId, SUBMISSION_HEADERS);
  if (rowIndex === -1) {
    throw new Error('找不到這筆上傳紀錄。');
  }

  const headerMap = headerIndexMap_(SUBMISSION_HEADERS);
  const rowObject = rowToObject_(SUBMISSION_HEADERS, values[rowIndex]);
  const studentEmail = normalizeEmail_(rowObject.studentEmail);
  const rootFolder = getRootFolder_();
  const studentFolder = getOrCreateFolder_(rootFolder, `${rowObject.studentId}_${rowObject.studentName}`);
  const reviewFolder = getOrCreateFolder_(studentFolder, '老師回傳檔案');
  const assignmentFolder = getOrCreateFolder_(reviewFolder, rowObject.assignmentType);
  const targetFolder = rowObject.week ? getOrCreateFolder_(assignmentFolder, `第${rowObject.week}週`) : assignmentFolder;
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const originalName = sanitizeFileName_(payload.fileName);
  const weekPart = rowObject.week ? `_第${rowObject.week}週` : '';
  const fileName = `${rowObject.studentId}_${rowObject.studentName}_${rowObject.assignmentType}${weekPart}_老師回傳_${stamp}_${originalName}`;
  const blob = Utilities.newBlob(decoded, payload.mimeType || 'application/octet-stream', fileName);
  const file = targetFolder.createFile(blob);

  if (studentEmail) {
    try {
      file.addViewer(studentEmail);
    } catch (error) {
      // 若網域政策不允許自動分享，老師仍可從後台取得檔案連結。
    }
  }

  sheet.getRange(rowIndex + 1, headerMap.reviewedFileName + 1).setValue(fileName);
  sheet.getRange(rowIndex + 1, headerMap.reviewedFileUrl + 1).setValue(file.getUrl());
  sheet.getRange(rowIndex + 1, headerMap.reviewedFileId + 1).setValue(file.getId());
  sheet.getRange(rowIndex + 1, headerMap.reviewedAt + 1).setValue(new Date());
  sheet.getRange(rowIndex + 1, headerMap.lastReviewedAt + 1).setValue(new Date());
  sheet.getRange(rowIndex + 1, headerMap.reviewerEmail + 1).setValue(getCurrentUserEmail_());

  return serializeForClient_({
    ok: true,
    message: '老師回傳檔案已上傳並分享給學生。',
    teacherData: getTeacherSubmissionData({})
  });
}

function importRosterText(pastedText) {
  assertTeacher_();
  if (!String(pastedText || '').trim()) {
    throw new Error('請貼上名單資料。');
  }

  const rows = parsePastedTable_(pastedText);
  if (rows.length < 2) {
    throw new Error('名單至少需要標題列與一筆學生資料。');
  }

  const headers = rows[0].map(cell => normalizeHeader_(cell));
  const indexes = {
    studentId: findHeaderIndex_(headers, ['學號', 'studentid', 'student_id']),
    studentName: findHeaderIndex_(headers, ['姓名', 'name', 'studentname']),
    degree: findHeaderIndex_(headers, ['學制', 'degree']),
    className: findHeaderIndex_(headers, ['年班', '班級', 'class']),
    gender: findHeaderIndex_(headers, ['性別', 'gender']),
    agency: findHeaderIndex_(headers, ['實習機構', '機構', 'agency']),
    agencyAddress: findHeaderIndex_(headers, ['機構地址', '地址', 'agencyaddress']),
    field: findHeaderIndex_(headers, ['領域', 'field']),
    supervisor: findHeaderIndex_(headers, ['學校督導', '督導', 'supervisor']),
    studentEmail: findHeaderIndex_(headers, ['email', 'googleemail', 'google帳號', '學生email'])
  };

  if (indexes.studentId === -1 || indexes.studentName === -1) {
    throw new Error('名單需要包含「學號」與「姓名」欄位。');
  }

  const sheet = getSheet_(SHEETS.students);
  const existing = getStudentIndexMap_();
  let imported = 0;
  let updated = 0;

  rows.slice(1).forEach(row => {
    const studentId = getCell_(row, indexes.studentId);
    const studentName = cleanStudentName_(getCell_(row, indexes.studentName));
    if (!studentId || !studentName) return;

    const record = [
      makeStudentKey_(studentId, studentName),
      studentId,
      studentName,
      getCell_(row, indexes.degree),
      getCell_(row, indexes.className),
      getCell_(row, indexes.gender),
      getCell_(row, indexes.agency),
      getCell_(row, indexes.agencyAddress),
      getCell_(row, indexes.field),
      getCell_(row, indexes.supervisor),
      normalizeEmail_(getCell_(row, indexes.studentEmail)),
      '',
      normalizeEmail_(getCell_(row, indexes.studentEmail)) ? 'bound' : 'unbound',
      '',
      ''
    ];

    const rowNumber = existing[record[0]];
    if (rowNumber) {
      const old = sheet.getRange(rowNumber, 1, 1, STUDENT_HEADERS.length).getValues()[0];
      record[10] = record[10] || old[10];
      record[11] = old[11];
      record[12] = record[10] || record[11] ? 'bound' : old[12] || 'unbound';
      record[13] = old[13];
      record[14] = old[14];
      sheet.getRange(rowNumber, 1, 1, STUDENT_HEADERS.length).setValues([record]);
      updated += 1;
    } else {
      sheet.appendRow(record);
      imported += 1;
    }
  });

  return serializeForClient_({
    ok: true,
    message: `名單匯入完成：新增${imported}筆，更新${updated}筆。`,
    rosterData: getRosterData()
  });
}

function updateStudentRecord(payload) {
  assertTeacher_();
  if (!payload) throw new Error('缺少學生資料。');

  const originalStudentKey = String(payload.studentKey || '').trim();
  const studentId = String(payload.studentId || '').trim();
  const studentName = cleanStudentName_(payload.studentName);
  if (!originalStudentKey) throw new Error('缺少要更新的學生代碼。');
  if (!studentId || !studentName) throw new Error('學號與姓名不可空白。');

  const sheet = getSheet_(SHEETS.students);
  const values = sheet.getDataRange().getValues();
  const rowIndex = findRowIndexById_(values, originalStudentKey, STUDENT_HEADERS);
  if (rowIndex === -1) throw new Error('找不到這位學生，請重新整理後再試。');

  const old = rowToObject_(STUDENT_HEADERS, values[rowIndex]);
  const getText = key => Object.prototype.hasOwnProperty.call(payload, key)
    ? String(payload[key] || '').trim()
    : String(old[key] || '').trim();
  const studentEmail = normalizeEmail_(payload.studentEmail || '');
  const userKey = studentEmail ? old.studentUserKey || '' : '';
  const record = [
    makeStudentKey_(studentId, studentName),
    studentId,
    studentName,
    getText('degree'),
    getText('className'),
    getText('gender'),
    getText('agency'),
    getText('agencyAddress'),
    getText('field'),
    getText('supervisor'),
    studentEmail,
    userKey,
    studentEmail || userKey ? 'bound' : 'unbound',
    studentEmail || userKey ? old.boundAt || new Date() : '',
    getText('notes')
  ];

  sheet.getRange(rowIndex + 1, 1, 1, STUDENT_HEADERS.length).setValues([record]);
  return serializeForClient_({
    ok: true,
    message: '學生資料已更新。',
    rosterData: getRosterData()
  });
}

function createStudentRecord(payload) {
  assertTeacher_();
  if (!payload) throw new Error('缺少學生資料。');

  const studentId = String(payload.studentId || '').trim();
  const studentName = cleanStudentName_(payload.studentName);
  if (!studentId || !studentName) throw new Error('新增學生時，學號與姓名不可空白。');

  const studentKey = makeStudentKey_(studentId, studentName);
  const existing = getStudentIndexMap_();
  if (existing[studentKey]) {
    throw new Error('這位學生已存在，請直接在名單中修改後按儲存。');
  }

  const studentEmail = normalizeEmail_(payload.studentEmail || '');
  const record = [
    studentKey,
    studentId,
    studentName,
    String(payload.degree || '').trim(),
    String(payload.className || '').trim(),
    String(payload.gender || '').trim(),
    String(payload.agency || '').trim(),
    String(payload.agencyAddress || '').trim(),
    String(payload.field || '').trim(),
    String(payload.supervisor || '').trim(),
    studentEmail,
    '',
    studentEmail ? 'bound' : 'unbound',
    studentEmail ? new Date() : '',
    String(payload.notes || '').trim()
  ];

  const sheet = getSheet_(SHEETS.students);
  sheet.appendRow(record);
  return serializeForClient_({
    ok: true,
    message: '學生已新增。',
    rosterData: getRosterData()
  });
}

function deleteStudentRecord(studentKey) {
  assertTeacher_();
  const key = String(studentKey || '').trim();
  if (!key) throw new Error('缺少要刪除的學生代碼。');

  const sheet = getSheet_(SHEETS.students);
  const values = sheet.getDataRange().getValues();
  const rowIndex = findRowIndexById_(values, key, STUDENT_HEADERS);
  if (rowIndex === -1) throw new Error('找不到這位學生，請重新整理後再試。');

  sheet.deleteRow(rowIndex + 1);
  return serializeForClient_({
    ok: true,
    message: '學生已從名單刪除。原有上傳紀錄與雲端檔案不會被刪除。',
    rosterData: getRosterData()
  });
}

function saveAnnouncement(payload) {
  assertTeacher_();
  if (!payload || !String(payload.title || '').trim()) {
    throw new Error('公告標題不可空白。');
  }

  const sheet = getSheet_(SHEETS.announcements);
  const values = sheet.getDataRange().getValues();
  const id = payload.announcementId || makeId_('ANN');
  const row = [
    id,
    new Date(),
    String(payload.title || '').trim(),
    String(payload.body || '').trim(),
    Boolean(payload.pinned),
    payload.visible !== false,
    Number(payload.sortOrder || 0)
  ];

  const rowIndex = findRowIndexById_(values, id, ANNOUNCEMENT_HEADERS);
  if (rowIndex === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
  }

  return serializeForClient_({
    ok: true,
    announcements: listAnnouncements_(true)
  });
}

function deleteAnnouncement(announcementId) {
  assertTeacher_();
  const sheet = getSheet_(SHEETS.announcements);
  const values = sheet.getDataRange().getValues();
  const rowIndex = findRowIndexById_(values, announcementId, ANNOUNCEMENT_HEADERS);
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 1);
  }
  return serializeForClient_({
    ok: true,
    announcements: listAnnouncements_(true)
  });
}

function listAnnouncements_(includeHidden) {
  const sheet = getSheet_(SHEETS.announcements);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).map(row => rowToObject_(ANNOUNCEMENT_HEADERS, row));
  return rows
    .filter(row => includeHidden || row.visible === true || row.visible === 'TRUE')
    .sort((a, b) => {
      const pinnedDiff = Number(Boolean(b.pinned)) - Number(Boolean(a.pinned));
      if (pinnedDiff !== 0) return pinnedDiff;
      const sortDiff = Number(b.sortOrder || 0) - Number(a.sortOrder || 0);
      if (sortDiff !== 0) return sortDiff;
      return new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime();
    });
}

function listSubmissionsForCurrentUser_() {
  const normalized = normalizeEmail_(getCurrentUserEmail_());
  const userKey = getCurrentUserKey_();
  const student = getStudentForCurrentUser_();
  return getAllSubmissionObjects_()
    .filter(row => rowBelongsToCurrentStudent_(row, student, normalized, userKey))
    .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
}

function getAllSubmissionObjects_() {
  const sheet = getSheet_(SHEETS.submissions);
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(row => row[0])
    .map(row => rowToObject_(SUBMISSION_HEADERS, row));
}

function listVisitsForCurrentUser_() {
  const normalized = normalizeEmail_(getCurrentUserEmail_());
  const userKey = getCurrentUserKey_();
  const student = getStudentForCurrentUser_();
  return getAllVisitObjects_()
    .filter(row => rowBelongsToCurrentStudent_(row, student, normalized, userKey))
    .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
}

function getAllVisitObjects_() {
  const sheet = getSheet_(SHEETS.visits);
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(row => row[0])
    .map(row => normalizeVisitObject_(rowToObject_(VISIT_HEADERS, row)));
}

function applyVisitFilters_(rows, filters) {
  const query = String(filters.query || '').trim().toLowerCase();
  const studentKey = String(filters.studentKey || '').trim();

  return rows
    .filter(row => {
      if (studentKey && makeStudentKey_(row.studentId, row.studentName) !== studentKey) return false;
      if (!query) return true;
      const haystack = [
        row.studentEmail,
        row.studentName,
        row.studentId,
        row.className,
        row.agency,
        row.agencyAddress,
        row.agencySupervisor,
        row.contactName,
        row.contactPhone,
        row.studentNote,
        row.status
      ].join(' ').toLowerCase();
      return haystack.indexOf(query) !== -1;
    })
    .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
}

function buildVisitSummary_(rows) {
  return {
    total: rows.length,
    pending: rows.filter(row => row.status === '待老師確認').length,
    confirmed: rows.filter(row => row.status === '已確認').length,
    revise: rows.filter(row => row.status === '請學生修改').length
  };
}

function applyTeacherFilters_(rows, filters) {
  const query = String(filters.query || '').trim().toLowerCase();
  const assignmentType = String(filters.assignmentType || '').trim();
  const status = String(filters.status || '').trim();
  const studentKey = String(filters.studentKey || '').trim();

  return rows
    .filter(row => {
      if (studentKey && makeStudentKey_(row.studentId, row.studentName) !== studentKey) return false;
      if (assignmentType && row.assignmentType !== assignmentType) return false;
      if (status && row.status !== status) return false;
      if (!query) return true;
      const haystack = [
        row.studentEmail,
        row.studentName,
        row.studentId,
        row.className,
        row.agency,
        row.assignmentType,
        row.week
      ].join(' ').toLowerCase();
      return haystack.indexOf(query) !== -1;
    })
    .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
}

function buildSummary_(rows) {
  const summary = {
    total: rows.length,
    submitted: 0,
    returned: 0,
    completed: 0,
    students: {}
  };

  rows.forEach(row => {
    if (row.status === '已完成') summary.completed += 1;
    if (row.status === '退件') summary.returned += 1;
    if (row.status === '已繳交') summary.submitted += 1;
    const key = row.studentEmail || `${row.studentId}_${row.studentName}`;
    if (!summary.students[key]) {
      summary.students[key] = {
        studentEmail: row.studentEmail,
        studentName: row.studentName,
        studentId: row.studentId,
        count: 0,
        returned: 0,
        completed: 0
      };
    }
    summary.students[key].count += 1;
    if (row.status === '退件') summary.students[key].returned += 1;
    if (row.status === '已完成') summary.students[key].completed += 1;
  });

  summary.studentCount = Object.keys(summary.students).length;
  summary.students = Object.values(summary.students)
    .sort((a, b) => String(a.studentId).localeCompare(String(b.studentId)));
  return summary;
}

function buildMissingList_(rows, studentRows) {
  const students = {};
  (studentRows || listStudents_()).forEach(student => {
    students[student.studentKey] = {
      studentEmail: student.studentEmail,
      studentName: student.studentName,
      studentId: student.studentId,
      submittedKeys: {}
    };
  });

  rows.forEach(row => {
    const key = makeStudentKey_(row.studentId, row.studentName);
    if (!students[key]) {
      students[key] = {
        studentEmail: row.studentEmail,
        studentName: row.studentName,
        studentId: row.studentId,
        submittedKeys: {}
      };
    }
    if (row.status === '已完成' || row.status === '已繳交') {
      students[key].submittedKeys[assignmentKey_(row.assignmentType, row.week)] = true;
    }
  });

  const requiredKeys = requiredAssignmentItems_().map(item => item.key);

  return Object.values(students).map(student => {
    const missing = requiredKeys.filter(key => !student.submittedKeys[key]).map(formatAssignmentKey_);
    return {
      studentEmail: student.studentEmail,
      studentName: student.studentName,
      studentId: student.studentId,
      missing
    };
  }).filter(item => item.missing.length > 0);
}

function buildProgressList_(rows, studentRows) {
  const requiredItems = requiredAssignmentItems_();
  const students = {};
  (studentRows || listStudents_()).forEach(student => {
    students[student.studentKey] = {
      studentEmail: student.studentEmail,
      studentName: student.studentName,
      studentId: student.studentId,
      className: student.className,
      itemsByKey: {}
    };
  });

  rows.forEach(row => {
    const studentKey = makeStudentKey_(row.studentId, row.studentName);
    if (!students[studentKey]) {
      students[studentKey] = {
        studentEmail: row.studentEmail,
        studentName: row.studentName,
        studentId: row.studentId,
        className: row.className,
        itemsByKey: {}
      };
    }

    const itemKey = assignmentKey_(row.assignmentType, row.week);
    const current = students[studentKey].itemsByKey[itemKey];
    const rowTime = new Date(row.timestamp).getTime() || 0;
    const currentTime = current ? new Date(current.timestamp).getTime() || 0 : -1;
    const rowAttempt = Number(row.attempt || 0);
    const currentAttempt = current ? Number(current.attempt || 0) : -1;
    if (!current || rowAttempt > currentAttempt || rowTime >= currentTime) {
      students[studentKey].itemsByKey[itemKey] = {
        label: formatAssignmentKey_(itemKey),
        status: row.status || '已繳交',
        attempt: row.attempt || '',
        timestamp: row.timestamp || ''
      };
    }
  });

  return Object.values(students)
    .sort((a, b) => String(a.studentId).localeCompare(String(b.studentId)))
    .map(student => {
      const items = requiredItems.map(required => {
        const item = student.itemsByKey[required.key];
        return item || {
          label: required.label,
          status: '未繳交',
          attempt: '',
          timestamp: ''
        };
      });
      return {
        studentEmail: student.studentEmail,
        studentName: student.studentName,
        studentId: student.studentId,
        className: student.className,
        missingCount: items.filter(item => item.status === '未繳交').length,
        returnedCount: items.filter(item => item.status === '退件').length,
        needsResubmissionCount: items.filter(item => item.status === '未繳交' || item.status === '退件').length,
        submittedCount: items.filter(item => item.status === '已繳交').length,
        completedCount: items.filter(item => item.status === '已完成').length,
        items
      };
    });
}

function getNextAttempt_(studentEmail, userKey, assignmentType, week) {
  const normalized = normalizeEmail_(studentEmail);
  const attempts = getAllSubmissionObjects_()
    .filter(row => {
      if (normalized) return normalizeEmail_(row.studentEmail) === normalized;
      return userKey && String(row.studentUserKey || '') === String(userKey);
    })
    .filter(row => row.assignmentType === assignmentType)
    .filter(row => String(row.week || '') === String(week || ''))
    .map(row => Number(row.attempt || 0));
  return attempts.length ? Math.max.apply(null, attempts) + 1 : 1;
}

function validateUploadPayload_(payload) {
  if (!payload) throw new Error('缺少上傳資料。');
  ['studentName', 'studentId', 'studentEmail', 'agencySupervisor', 'assignmentType', 'fileName', 'fileData'].forEach(key => {
    if (!String(payload[key] || '').trim()) {
      throw new Error(`缺少必要欄位：${key}`);
    }
  });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(payload.studentEmail || '').trim())) {
    throw new Error('Gmail 格式不正確，請填寫可收件的 Email。');
  }
}

function validateVisitPayload_(payload) {
  if (!payload) throw new Error('缺少訪視安排資料。');
  ['studentName', 'studentId', 'studentEmail', 'agency', 'agencyAddress', 'agencySupervisor', 'visitDate', 'startTime', 'endTime'].forEach(key => {
    if (!String(payload[key] || '').trim()) {
      throw new Error(`缺少必要欄位：${key}`);
    }
  });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(payload.studentEmail || '').trim())) {
    throw new Error('Gmail 格式不正確，請填寫可收件的 Email。');
  }
  const start = String(payload.startTime || '');
  const end = String(payload.endTime || '');
  if (start && end && start >= end) {
    throw new Error('結束時間必須晚於開始時間。');
  }
}

function validateAssignmentFileType_(assignmentType, fileName, mimeType) {
  const lowerName = String(fileName || '').toLowerCase();
  const type = String(mimeType || '').toLowerCase();
  if (assignmentType === '實習總報告' && !(type === 'application/pdf' || lowerName.endsWith('.pdf'))) {
    throw new Error('實習總報告請上傳 PDF 檔。');
  }
  if ((assignmentType === '實習照片1' || assignmentType === '實習照片2') &&
      !(type.indexOf('image/') === 0 || /\.(jpg|jpeg|png|heic|webp)$/i.test(lowerName))) {
    throw new Error('實習照片請上傳圖片檔，例如 JPG、PNG、HEIC 或 WEBP。');
  }
}

function assertAndBindRosterStudent_(payload, email, userKey) {
  const students = listStudents_();
  if (!students.length) {
    throw new Error('尚未匯入實習學生名單。請老師先到「名單管理」匯入名單。');
  }

  const normalizedEmail = normalizeEmail_(email);
  const normalizedUserKey = String(userKey || '').trim();
  const byEmail = normalizedEmail
    ? students.find(student => normalizeEmail_(student.studentEmail) === normalizedEmail)
    : null;
  if (byEmail) return byEmail;
  const byUserKey = students.find(student => normalizedUserKey && String(student.studentUserKey || '') === normalizedUserKey);
  if (byUserKey) return byUserKey;

  const studentId = String(payload.studentId || '').trim();
  const studentName = cleanStudentName_(payload.studentName);
  const matched = students.find(student =>
    String(student.studentId).trim() === studentId &&
    String(student.studentName).trim() === studentName
  );

  if (!matched) {
    throw new Error('這組學號與姓名不在實習名單中，請確認是否輸入正確。');
  }

  const sameBoundUser = normalizedUserKey && String(matched.studentUserKey || '') === normalizedUserKey;
  if (matched.studentEmail && normalizedEmail && normalizeEmail_(matched.studentEmail) !== normalizedEmail && !sameBoundUser) {
    throw new Error('這位學生已綁定其他 Google 帳號。若需更換帳號，請由老師在名單表中調整。');
  }

  if (matched.studentUserKey && normalizedUserKey && String(matched.studentUserKey) !== normalizedUserKey) {
    throw new Error('這位學生已綁定其他 Google 登入身分。若需更換帳號，請由老師在名單表中調整。');
  }

  bindStudentIdentity_(matched.studentKey, normalizedEmail, normalizedUserKey);
  matched.studentEmail = normalizedEmail;
  matched.studentUserKey = normalizedUserKey;
  matched.bindStatus = 'bound';
  matched.boundAt = new Date();
  return matched;
}

function findStudentForRecovery_(payload) {
  const studentId = String(payload.studentId || '').trim();
  const studentName = cleanStudentName_(payload.studentName);
  const studentEmail = normalizeEmail_(payload.studentEmail || '');
  if (!studentId || !studentName || !studentEmail) {
    throw new Error('請完整填寫姓名、學號與 Gmail，系統才能重新辨識你的身分。');
  }

  const matched = listStudents_().find(student =>
    String(student.studentId).trim() === studentId &&
    cleanStudentName_(student.studentName) === studentName
  );
  if (!matched) {
    throw new Error('名單中找不到這組姓名與學號，請確認是否輸入正確。');
  }

  const rosterEmail = normalizeEmail_(matched.studentEmail || '');
  if (rosterEmail && rosterEmail !== studentEmail) {
    throw new Error('Gmail 與老師名單中的帳號不一致，請用第一次登記或上傳時的 Gmail。');
  }
  if (!rosterEmail) {
    throw new Error('老師名單尚未登記這位學生的 Gmail，請先請老師在名單管理補上 Gmail。');
  }
  return matched;
}

function bindStudentIdentity_(studentKey, email, userKey) {
  const sheet = getSheet_(SHEETS.students);
  const values = sheet.getDataRange().getValues();
  const rowIndex = findRowIndexById_(values, studentKey, STUDENT_HEADERS);
  if (rowIndex === -1) return;
  const headerMap = headerIndexMap_(STUDENT_HEADERS);
  if (email) sheet.getRange(rowIndex + 1, headerMap.studentEmail + 1).setValue(email);
  if (userKey) sheet.getRange(rowIndex + 1, headerMap.studentUserKey + 1).setValue(userKey);
  sheet.getRange(rowIndex + 1, headerMap.bindStatus + 1).setValue('bound');
  sheet.getRange(rowIndex + 1, headerMap.boundAt + 1).setValue(new Date());
}

function listStudents_() {
  const sheet = getSheet_(SHEETS.students);
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(row => row[0])
    .map(row => rowToObject_(STUDENT_HEADERS, row));
}

function getStudentByEmail_(email) {
  const normalized = normalizeEmail_(email);
  if (!normalized) return null;
  return listStudents_().find(student => normalizeEmail_(student.studentEmail) === normalized) || null;
}

function getStudentForCurrentUser_() {
  const email = normalizeEmail_(getCurrentUserEmail_());
  const userKey = getCurrentUserKey_();
  return listStudents_().find(student => {
    if (email && normalizeEmail_(student.studentEmail) === email) return true;
    return userKey && String(student.studentUserKey || '') === String(userKey);
  }) || null;
}

function rowBelongsToCurrentStudent_(row, student, email, userKey) {
  const normalizedEmail = normalizeEmail_(email);
  const normalizedUserKey = String(userKey || '').trim();
  if (normalizedEmail && normalizeEmail_(row.studentEmail) === normalizedEmail) return true;
  if (normalizedUserKey && String(row.studentUserKey || '') === normalizedUserKey) return true;
  if (!student) return false;

  const studentEmail = normalizeEmail_(student.studentEmail);
  const studentUserKey = String(student.studentUserKey || '').trim();
  if (studentEmail && normalizeEmail_(row.studentEmail) === studentEmail) return true;
  if (studentUserKey && String(row.studentUserKey || '') === studentUserKey) return true;
  return makeStudentKey_(row.studentId, row.studentName) === student.studentKey;
}

function normalizeVisitObject_(visit) {
  visit.visitDate = normalizeDateField_(visit.visitDate);
  visit.startTime = normalizeTimeField_(visit.startTime);
  visit.endTime = normalizeTimeField_(visit.endTime);
  return visit;
}

function normalizeDateField_(value) {
  if (!value) return '';
  const timezone = Session.getScriptTimeZone();
  if (value instanceof Date) {
    return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  const isoMatch = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoMatch) return isoMatch[1];
  return text;
}

function normalizeTimeField_(value) {
  if (!value) return '';
  const timezone = Session.getScriptTimeZone();
  if (value instanceof Date) {
    return Utilities.formatDate(value, timezone, 'HH:mm');
  }
  const text = String(value).trim();
  const timeOnlyMatch = text.match(/^(\d{1,2}):(\d{2})/);
  if (timeOnlyMatch) {
    return `${timeOnlyMatch[1].padStart(2, '0')}:${timeOnlyMatch[2]}`;
  }
  const isoMatch = text.match(/T(\d{2}):(\d{2})/);
  if (isoMatch) return `${isoMatch[1]}:${isoMatch[2]}`;
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, timezone, 'HH:mm');
  }
  return text;
}

function getStudentIndexMap_() {
  const sheet = getSheet_(SHEETS.students);
  const values = sheet.getDataRange().getValues();
  const map = {};
  values.slice(1).forEach((row, index) => {
    if (row[0]) map[row[0]] = index + 2;
  });
  return map;
}

function assertConfigured_() {
  initializeConfigIfNeeded_();
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty(CONFIG.ROOT_FOLDER_ID) || !properties.getProperty(CONFIG.SPREADSHEET_ID)) {
    throw new Error('尚未完成初始化。請老師先在 Apps Script 編輯器執行 runInitialSetup("你的GoogleEmail")。');
  }
}

function initializeConfigIfNeeded_() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty(CONFIG.ROOT_FOLDER_ID) && PRECONFIGURED.ROOT_FOLDER_ID) {
    properties.setProperty(CONFIG.ROOT_FOLDER_ID, PRECONFIGURED.ROOT_FOLDER_ID);
  }
  if (!properties.getProperty(CONFIG.SPREADSHEET_ID) && PRECONFIGURED.SPREADSHEET_ID) {
    properties.setProperty(CONFIG.SPREADSHEET_ID, PRECONFIGURED.SPREADSHEET_ID);
  }
  if (!properties.getProperty(CONFIG.TEACHER_EMAILS) && PRECONFIGURED.TEACHER_EMAILS) {
    properties.setProperty(CONFIG.TEACHER_EMAILS, PRECONFIGURED.TEACHER_EMAILS);
  }
  if (!properties.getProperty(CONFIG.COURSE_TITLE)) {
    properties.setProperty(CONFIG.COURSE_TITLE, DEFAULT_COURSE_TITLE);
  }
}

function ensureProjectSheets_() {
  const spreadsheet = getSpreadsheet_();
  ensureSheet_(spreadsheet, SHEETS.submissions, SUBMISSION_HEADERS);
  ensureSheet_(spreadsheet, SHEETS.announcements, ANNOUNCEMENT_HEADERS);
  ensureSheet_(spreadsheet, SHEETS.students, STUDENT_HEADERS);
  seedAnnouncements_();
  seedDefaultStudents_();
  revokeDirectTestStudentShares_();
}

function assertTeacher_() {
  assertConfigured_();
  if (!isTeacher_(getCurrentUserEmail_())) {
    throw new Error('目前登入帳號沒有老師後台權限。');
  }
}

function isTeacher_(email) {
  const normalized = normalizeEmail_(email);
  if (!normalized) return false;
  const teachers = String(PropertiesService.getScriptProperties().getProperty(CONFIG.TEACHER_EMAILS) || '')
    .split(',')
    .map(normalizeEmail_)
    .filter(Boolean);
  return teachers.indexOf(normalized) !== -1;
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(CONFIG.SPREADSHEET_ID);
  return SpreadsheetApp.openById(spreadsheetId);
}

function getSheet_(name) {
  const spreadsheet = getSpreadsheet_();
  if (name === SHEETS.submissions) return ensureSheet_(spreadsheet, name, SUBMISSION_HEADERS);
  if (name === SHEETS.students) return ensureSheet_(spreadsheet, name, STUDENT_HEADERS);
  if (name === SHEETS.visits) return ensureSheet_(spreadsheet, name, VISIT_HEADERS);
  return ensureSheet_(spreadsheet, name, ANNOUNCEMENT_HEADERS);
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = firstRow.some(Boolean);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const currentLastColumn = sheet.getLastColumn();
    const existing = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0];
    const missing = headers.filter(header => existing.indexOf(header) === -1);
    if (missing.length) {
      sheet.getRange(1, currentLastColumn + 1, 1, missing.length).setValues([missing]);
    }
  }
  return sheet;
}

function seedAnnouncements_() {
  const sheet = getSheet_(SHEETS.announcements);
  if (sheet.getLastRow() > 1) return;
  sheet.appendRow([
    makeId_('ANN'),
    new Date(),
    '實習作業上傳公告',
    [
      '請依規定上傳實習計畫書、8週週誌、讀書心得報告與個案分析報告。',
      '週誌若被退件，請依老師評語修正後重新上傳。',
      '檔案建議使用 PDF 或 DOCX，單一檔案請勿超過25MB。'
    ].join('\n'),
    true,
    true,
    100
  ]);
}

function seedDefaultStudents_() {
  // Public template starts with an empty roster.
  // Use roster management or paste an Excel roster after setup.
}

function revokeDirectTestStudentShares_() {
  // Public template has no hard-coded test student or source files to revoke.
}

function getRootFolder_() {
  const rootFolderId = PropertiesService.getScriptProperties().getProperty(CONFIG.ROOT_FOLDER_ID);
  return DriveApp.getFolderById(rootFolderId);
}

function getOrCreateFolder_(parent, name) {
  const safeName = sanitizeFolderName_(name);
  const folders = parent.getFoldersByName(safeName);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(safeName);
}

function getCourseTitle_() {
  return PropertiesService.getScriptProperties().getProperty(CONFIG.COURSE_TITLE) || DEFAULT_COURSE_TITLE;
}

function getCurrentUserEmail_() {
  try {
    return normalizeEmail_(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function getCurrentUserKey_() {
  try {
    return Session.getTemporaryActiveUserKey();
  } catch (error) {
    return '';
  }
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function sanitizeFolderName_(name) {
  return String(name || '未命名')
    .replace(/[\\/:*?"<>|#%{}~&]/g, '_')
    .trim()
    .slice(0, 120) || '未命名';
}

function sanitizeFileName_(name) {
  return String(name || 'submission')
    .replace(/[\\/:*?"<>|#%{}~&]/g, '_')
    .trim()
    .slice(0, 160) || 'submission';
}

function cleanStudentName_(name) {
  return String(name || '')
    .replace(/\uE396/g, '')
    .trim();
}

function makeId_(prefix) {
  return `${prefix}-${Utilities.getUuid()}`;
}

function makeStudentKey_(studentId, studentName) {
  return `${String(studentId || '').trim()}::${cleanStudentName_(studentName)}`;
}

function rowToObject_(headers, row) {
  return headers.reduce((object, header, index) => {
    object[header] = row[index];
    return object;
  }, {});
}

function serializeForClient_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  if (Array.isArray(value)) {
    return value.map(serializeForClient_);
  }
  if (value && typeof value === 'object') {
    return Object.keys(value).reduce((object, key) => {
      object[key] = serializeForClient_(value[key]);
      return object;
    }, {});
  }
  return value;
}

function headerIndexMap_(headers) {
  return headers.reduce((map, header, index) => {
    map[header] = index;
    return map;
  }, {});
}

function findRowIndexById_(values, id, headers) {
  const idIndex = headerIndexMap_(headers)[headers[0]];
  for (let index = 1; index < values.length; index += 1) {
    if (String(values[index][idIndex]) === String(id)) return index;
  }
  return -1;
}

function parsePastedTable_(text) {
  return String(text || '')
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean)
    .map(line => line.indexOf('\t') !== -1 ? line.split('\t') : line.split(','))
    .map(row => row.map(cell => String(cell || '').trim()));
}

function normalizeHeader_(value) {
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

function findHeaderIndex_(headers, candidates) {
  return headers.findIndex(header => candidates.map(normalizeHeader_).indexOf(header) !== -1);
}

function getCell_(row, index) {
  if (index === -1) return '';
  return String(row[index] || '').trim();
}

function assignmentKey_(assignmentType, week) {
  return `${assignmentType}::${week || ''}`;
}

function requiredAssignmentItems_() {
  const items = [];
  ASSIGNMENTS.forEach(assignment => {
    if (assignment.requiresWeek) {
      (assignment.weeks || []).forEach(week => {
        items.push(assignmentKey_(assignment.type, week));
      });
    } else {
      items.push(assignmentKey_(assignment.type, ''));
    }
  });
  return items.map(key => ({
    key,
    label: formatAssignmentKey_(key)
  }));
}

function formatAssignmentKey_(key) {
  const parts = key.split('::');
  if (parts[0] === '週誌') return `第${parts[1]}週週誌`;
  return parts[0];
}

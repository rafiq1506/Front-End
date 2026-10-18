const TEACHER_EMAIL = 'mdrafiq@svc.ac.in';
const SENDER_DISPLAY_NAME = 'SEC: Front End'
const SPREADSHEET_ID = '1Y42vPwbyE7rK3WLVot2D-bsV2z7f33dczd5WptrHCts';
const SHEET_TAB_NAME = 'Submissions';

const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;
const MAX_EXTRACTED_BYTES = 25 * 1024 * 1024;
const MAX_PROJECT_FILES = 1000;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Practical Desk | Dept. of Mathematics')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function setup() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('ROOT_FOLDER_ID')) {
    const rootFolder = DriveApp.createFolder('Student Submissions');
    properties.setProperty('ROOT_FOLDER_ID', rootFolder.getId());
  }
}

function text_(value, name, max) {
  if (typeof value !== 'string' || !value.trim() || value.length > max) throw new Error('Invalid ' + name + '.');
  return value.trim();
}

function folderName_(value) {
  return value.replace(/[\\/\x00-\x1f]/g, '_');
}

function childFolder_(parent, name) {
  const matches = parent.getFoldersByName(name);
  return matches.hasNext() ? matches.next() : parent.createFolder(name);
}

// Writes directly to your configured Google Sheet
function appendSheetRow_(row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_TAB_NAME) || ss.getSheets()[0];
    
    // Read the headers in Row 1 to match keys dynamically
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      throw new Error('The submissions sheet has no header row.');
    }
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const rowData = headers.map(header => {
      const key = String(header).trim();
      return row[key] !== undefined ? row[key] : '';
    });
    
    sheet.appendRow(rowData);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function extractProject_(bytes, name) {
  let blobs;
  try { 
    blobs = Utilities.unzip(Utilities.newBlob(bytes, 'application/zip', name)); 
  } catch (error) { 
    throw new Error('Cannot extract this ZIP. Please submit a valid unencrypted ZIP archive.'); 
  }
  
  let total = 0;
  const entries = [];
  blobs.forEach(blob => {
    const path = blob.getName();
    if (typeof path !== 'string' || !path || /[\\\x00-\x1f]/.test(path) || path.charAt(0) === '/' || /^[A-Za-z]:/.test(path)) {
      throw new Error('The ZIP contains an invalid file path.');
    }
    const parts = path.replace(/\/$/, '').split('/');
    if (parts.some(part => !part || part === '.' || part === '..')) {
      throw new Error('The ZIP contains an unsafe file path.');
    }
    if (parts[0] === '__MACOSX' || parts[parts.length - 1] === '.DS_Store' || path.endsWith('/')) return;
    
    total += blob.getBytes().length;
    if (total > MAX_EXTRACTED_BYTES) throw new Error('Extracted project files exceed 25 MB.');
    entries.push({path: parts, blob: blob});
    if (entries.length > MAX_PROJECT_FILES) throw new Error('The ZIP contains more than 1,000 project files.');
  });
  
  if (!entries.length) throw new Error('The ZIP contains no project files.');
  const first = entries[0].path[0];
  if (entries.every(entry => entry.path.length > 1 && entry.path[0] === first)) {
    entries.forEach(entry => entry.path.shift());
  }
  
  const paths = new Set();
  entries.forEach(entry => {
    const path = entry.path.join('/');
    if (paths.has(path)) throw new Error('The ZIP contains duplicate file paths.');
    paths.add(path);
  });
  
  entries.forEach(entry => {
    for (let length = 1; length < entry.path.length; length++) {
      if (paths.has(entry.path.slice(0, length).join('/'))) throw new Error('The ZIP has a file and folder with the same path.');
    }
  });
  
  return {entries: entries, totalBytes: total};
}

function submitPractical(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Invalid submission.');
  const properties = PropertiesService.getScriptProperties();
  let rootId = properties.getProperty('ROOT_FOLDER_ID');
  
  // Auto-initialize folder if not yet run
  if (!rootId) {
    const rootFolder = DriveApp.createFolder('Student Submissions');
    rootId = rootFolder.getId();
    properties.setProperty('ROOT_FOLDER_ID', rootId);
  }

  const studentEmail = Session.getActiveUser().getEmail();
  if (!studentEmail) throw new Error('Could not identify authenticated student email. Sign in to your account.');

  const details = {};
  ['studentName', 'collegeRoll', 'examRoll', 'course', 'practicalName'].forEach(key => {
    details[key] = text_(payload[key], key, 100);
  });

  details.studentEmail = studentEmail;
  details.githubUrl = text_(payload.githubUrl, 'repository URL', 500);
  details.liveUrl = text_(payload.liveUrl, 'live URL', 500);
  if (!/^https:\/\/(www\.)?github\.com\/[A-Za-z0-9_-]+\/[A-Za-z0-9_.-]+\/?$/.test(details.githubUrl)) throw new Error('Invalid GitHub repository URL.');
  if (!/^https:\/\/[^\s/?#]+(?:[/?#][^\s]*)?$/i.test(details.liveUrl)) throw new Error('Invalid HTTPS live URL.');
  
  const name = text_(payload.archiveName, 'archive name', 200);
  if (!/\.zip$/i.test(name)) throw new Error('Only ZIP archives are accepted.');
  if (typeof payload.archiveBase64 !== 'string' || !payload.archiveBase64.length || payload.archiveBase64.length > Math.ceil(MAX_UPLOAD_BYTES / 3) * 4) {
    throw new Error('Upload exceeds 25 MB or is empty.');
  }
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(payload.archiveBase64) || payload.archiveBase64.length % 4) {
    throw new Error('Invalid file encoding.');
  }
  
  const bytes = Utilities.base64Decode(payload.archiveBase64);
  if (bytes.length < 4 || bytes.length > MAX_UPLOAD_BYTES) throw new Error('Invalid ZIP size.');
  const signature = bytes.slice(0, 4).map(value => value & 255).join(',');
  if (signature !== '80,75,3,4' && signature !== '80,75,5,6') throw new Error('The file is not a valid ZIP archive.');
  
  const project = extractProject_(bytes, name);

  let destination;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const root = DriveApp.getFolderById(rootId);
    const student = childFolder_(root, folderName_(details.studentName + '-' + details.collegeRoll));
    const practicalName = folderName_(details.practicalName);
    const existing = student.getFoldersByName(practicalName);
    if (existing.hasNext()) {
      destination = existing.next();
      // Optional: trash old files inside the folder so new ones replace them
      const oldFiles = destination.getFiles();
      while (oldFiles.hasNext()) {
        oldFiles.next().setTrashed(true);
      }
    } else {
      destination = student.createFolder(practicalName);
    }
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }

  const receipt = Utilities.getUuid();
  const submittedAt = new Date().toISOString();

  try {
    const folderCache = new Map([['', destination]]);
    project.entries.forEach(entry => {
      let parent = destination;
      for (let depth = 1; depth < entry.path.length; depth++) {
        const path = entry.path.slice(0, depth).join('/');
        if (!folderCache.has(path)) folderCache.set(path, parent.createFolder(entry.path[depth - 1]));
        parent = folderCache.get(path);
      }
      parent.createFile(entry.blob.setName(entry.path[entry.path.length - 1]));
    });
  } catch (error) {
    try { destination.setTrashed(true); } catch (cleanupError) {}
    throw new Error('Could not save all project files: ' + error.message);
  }

  const row = Object.assign(details, {
    receipt: receipt, 
    submittedAt: submittedAt, 
    folderUrl: destination.getUrl(), 
    fileCount: String(project.entries.length), 
    totalBytes: String(project.totalBytes)
  });

  try {
    appendSheetRow_(row);
    return { receipt: receipt, submittedAt: submittedAt, sheetSaved: true };
  } catch (error) {
    console.error('Sheet logging error for receipt ' + receipt + ': ' + error.message);
    return { receipt: receipt, submittedAt: submittedAt, sheetSaved: false };
  }
}

function sendBatchReceiptEmail(student, practicalList) {
  try {
    const remainingQuota = MailApp.getRemainingDailyQuota();
    if (remainingQuota < 1) {
      console.warn('Daily email quota reached. Skipping email delivery.');
      return { sent: false, error: 'Daily email quota exhausted' };
    }

    student = student || {};
    practicalList = Array.isArray(practicalList) ? practicalList : [];

    const recipient = Session.getActiveUser().getEmail() || student.studentEmail;
    if (!recipient) return { sent: false, error: 'No recipient email found' };

    const rowsHtml = practicalList.length > 0 
      ? practicalList.map((p, idx) => `
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 8px 12px; font-weight: bold; color: #0a4f85;">${idx + 1}</td>
            <td style="padding: 8px 12px;"><strong>${p.practicalName || 'Practical ' + (idx + 1)}</strong></td>
            <td style="padding: 8px 12px;">
              ${p.githubUrl ? `<a href="${p.githubUrl}" style="color: #0a7fb8; text-decoration: none;" target="_blank">Repository</a>` : '—'}
              &nbsp;|&nbsp;
              ${p.liveUrl ? `<a href="${p.liveUrl}" style="color: #0a7fb8; text-decoration: none;" target="_blank">Live Site</a>` : '—'}
            </td>
          </tr>
        `).join('')
      : `<tr><td colspan="3" style="padding: 12px; text-align: center; color: #888;">No individual practical items recorded.</td></tr>`;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 620px; margin: 0 auto; border: 1px solid #d4eaf4; border-radius: 12px; overflow: hidden; background: #ffffff;">
        <div style="background: linear-gradient(120deg, #0a6fb0, #17c3d6); color: #ffffff; padding: 24px; text-align: center;">
          <h2 style="margin: 0; font-size: 22px;">Practical Submission Receipt</h2>
          <p style="margin: 6px 0 0; font-size: 13.5px; opacity: 0.95;">${student.course || 'Front End Web Design and Development'}</p>
        </div>
        
        <div style="padding: 24px; color: #0b2a3f;">
          <p style="font-size: 15px; margin-top: 0;">Dear <strong>${student.studentName || 'Student'}</strong>,</p>
          <p style="font-size: 14px; color: #587a8f; margin-bottom: 20px;">Your practical submission batch has been successfully received and registered in the departmental submission records.</p>
          
          <table style="width: 100%; margin-bottom: 24px; font-size: 13px; background: #f6fbfe; border: 1px solid #c6e1ee; border-radius: 8px; padding: 12px;">
            <tr><td style="padding: 4px 6px; color: #587a8f; width: 40%;"><strong>College Roll Number:</strong></td><td style="padding: 4px 6px;"><strong>${student.collegeRoll || 'N/A'}</strong></td></tr>
            <tr><td style="padding: 4px 6px; color: #587a8f;"><strong>Exam Roll Number:</strong></td><td style="padding: 4px 6px;"><strong>${student.examRoll || 'N/A'}</strong></td></tr>
            <tr><td style="padding: 4px 6px; color: #587a8f;"><strong>Course:</strong></td><td style="padding: 4px 6px;"><strong>${student.course || 'N/A'}</strong></td></tr>
            <tr><td style="padding: 4px 6px; color: #587a8f;"><strong>Total Practicals Submitted:</strong></td><td style="padding: 4px 6px;"><strong>${practicalList.length}</strong></td></tr>
            <tr><td style="padding: 4px 6px; color: #587a8f;"><strong>Primary Receipt ID:</strong></td><td style="padding: 4px 6px;"><code>${practicalList[0]?.receipt || practicalList[0]?.result?.receipt || 'CONFIRMED'}</code></td></tr>
          </table>

          <h3 style="font-size: 14px; margin: 0 0 10px; color: #0a4f85;">Summary of Submitted Practicals</h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
            <thead>
              <tr style="background: #eaf7fc; text-align: left; color: #2f5f7c;">
                <th style="padding: 8px 12px;">#</th>
                <th style="padding: 8px 12px;">Practical Name</th>
                <th style="padding: 8px 12px;">Links</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>
          
          <div style="font-size: 12px; color: #64748b; margin-top: 26px; border-top: 1px solid #e2e8f0; padding-top: 14px; line-height: 1.6;">
            <strong>Course Instructor:</strong> Dr. Mohd Rafiq<br>
            <strong>Contact Email:</strong> <a href="mailto:${TEACHER_EMAIL}" style="color: #0a7fb8;">${TEACHER_EMAIL}</a><br>
            <em>Please preserve this email receipt for practical assessment verification.</em>
          </div>
        </div>
      </div>
    `;

    GmailApp.sendEmail(recipient, 'Practical Submission Confirmation — ' + (student.studentName || 'Student'), '', {
      htmlBody: htmlBody,
      replyTo: TEACHER_EMAIL,
      name: SENDER_DISPLAY_NAME
    });
    return { sent: true };
  } catch (err) {
    console.error('Email send failed: ' + err.message);
    return { sent: false, error: err.message };
  }
}

// Direct test function to verify writing into your specific sheet
function testSheetWrite() {
  const dummyRow = {
    studentName: 'Test Student',
    studentEmail: 'mdrafiq@svc.ac.in',
    collegeRoll: '24BCA001',
    examRoll: '23058000001',
    course: 'SEC: Front End Web Design',
    practicalName: 'PRACTICAL-TEST',
    githubUrl: 'https://github.com/example/test-repo',
    liveUrl: 'https://example.github.io/test-live/',
    folderUrl: 'https://drive.google.com',
    fileCount: '3',
    totalBytes: '40960',
    receipt: 'TEST-' + Utilities.getUuid().slice(0, 8),
    submittedAt: new Date().toISOString()
  };
  
  appendSheetRow_(dummyRow);
  Logger.log('Success! Check your Google Sheet to see the test row.');
}
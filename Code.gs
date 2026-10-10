/**
 * Deep Tubewell Survey – Google Apps Script backend
 * Data  -> Google Sheet "Deep Tubewell Survey Data" (auto-created in your Drive)
 * Photo -> Drive folder "Tubewell_Survey_Photos" (auto-created)
 */

const SHEET_NAME  = 'Tubewell_Data';
const FOLDER_NAME = 'Tubewell_Survey_Photos';
const HEADERS = [
  'Timestamp', 'Surveyor', 'Location (Address/Area)', 'Latitude', 'Longitude', 'Map Link',
  'Ward No.', 'Operator Name', 'Operator Phone', 'Collection Date',
  'Strainer Depth', 'Strainer Unit', 'Pumping Hour (hr/day)',
  'Pumping Capacity', 'Capacity Unit', 'Delivery', 'Delivery Unit',
  'Remarks', 'Photo'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Deep Tubewell Survey')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Receives submissions from the externally hosted form (GitHub Pages). */
function doPost(e) {
  let out;
  try {
    out = submitData(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to authorize & create the Sheet/Folder. */
function setup() {
  const sh = getSheet_();
  getFolder_();
  Logger.log('Sheet URL: ' + sh.getParent().getUrl());
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Deep Tubewell Survey Data');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(SHEET_NAME); }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#0b6e4f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function submitData(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const required = ['surveyor', 'ward', 'operatorName', 'operatorPhone', 'date',
                      'strainerDepth', 'pumpingHour', 'capacity', 'delivery'];
    required.forEach(k => {
      if (d[k] === undefined || String(d[k]).trim() === '') throw new Error('Missing field: ' + k);
    });
    if (!d.location && !(d.lat && d.lng)) throw new Error('Location or GPS is required');

    let photoUrl = '';
    if (d.photo && d.photo.data) {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(d.photo.data), d.photo.type || 'image/jpeg',
        'Ward' + d.ward + '_' + Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyyyMMdd_HHmmss') + '.jpg'
      );
      photoUrl = getFolder_().createFile(blob).getUrl();
    }

    const mapLink = (d.lat && d.lng) ? 'https://maps.google.com/?q=' + d.lat + ',' + d.lng : '';
    const sh = getSheet_();
    sh.appendRow([
      new Date(), d.surveyor || '', d.location || '', d.lat || '', d.lng || '', mapLink,
      d.ward, d.operatorName, "'" + d.operatorPhone, d.date,
      Number(d.strainerDepth), d.strainerUnit, Number(d.pumpingHour),
      Number(d.capacity), d.capacityUnit, d.delivery, d.deliveryUnit,
      d.remarks || '', photoUrl
    ]);
    return { ok: true, count: sh.getLastRow() - 1 };
  } finally {
    lock.releaseLock();
  }
}

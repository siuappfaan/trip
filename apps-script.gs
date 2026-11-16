// Paste into Extensions > Apps Script of a new Google Sheet, then Deploy > Web app (Execute as: Me, Access: Anyone).
const N=40000;
function sh(){const s=SpreadsheetApp.getActiveSpreadsheet();return s.getSheetByName('data')||s.insertSheet('data')}
function out(t){return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JSON)}
function doGet(){const s=sh();const v=s.getRange(1,1,Math.max(s.getLastRow(),1),1).getValues().map(r=>r[0]).join('');return out(v||'null')}
function doPost(e){const s=sh(),t=e.postData.contents;s.clear();for(let i=0,r=1;i<t.length;i+=N,r++)s.getRange(r,1).setValue(t.slice(i,i+N));return out(t)}

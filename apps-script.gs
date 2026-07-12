// エンタメ記録アプリ用の簡易バックエンド（Google Apps Script）
// 使い方は README_apps-script.md を参照。
// このスプレッドシートの拡張機能 > Apps Script にこの内容を貼り付けてデプロイする。

const SHEET_NAME = 'entries';
const HEADERS = ['id','genre','date','time','title','program','place','performers','rating','review','impression','updatedAt'];
const TOKEN = '350458bb8a044dcfacd9d0e3de131ed1'; // 簡易トークン。公開リポジトリのJSに埋め込むため、本当のセキュリティにはならない点に注意。

function getSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if(sheet.getLastRow() === 0){
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function rowsToObjects_(sheet){
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  return rows.filter(function(r){ return r[0]; }).map(function(r){
    const obj = {};
    headers.forEach(function(h,i){ obj[h] = r[i]; });
    obj.rating = Number(obj.rating) || 0;
    return obj;
  });
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  const sheet = getSheet_();
  return jsonOut_({ok:true, entries: rowsToObjects_(sheet)});
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.token !== TOKEN){
      return jsonOut_({ok:false, error:'unauthorized'});
    }
    const sheet = getSheet_();
    if(body.action === 'upsert'){
      upsertEntry_(sheet, body.entry);
    }else if(body.action === 'delete'){
      deleteEntry_(sheet, body.id);
    }else if(body.action === 'bulkUpsert'){
      (body.entries || []).forEach(function(entry){ upsertEntry_(sheet, entry); });
    }else{
      return jsonOut_({ok:false, error:'unknown action'});
    }
    return jsonOut_({ok:true, entries: rowsToObjects_(sheet)});
  }catch(err){
    return jsonOut_({ok:false, error:String(err)});
  }
}

function upsertEntry_(sheet, entry){
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  entry.updatedAt = new Date().toISOString();
  const rowValues = headers.map(function(h){ return entry[h] !== undefined ? entry[h] : ''; });
  for(let i=1;i<data.length;i++){
    if(data[i][idCol] === entry.id){
      sheet.getRange(i+1,1,1,headers.length).setValues([rowValues]);
      return;
    }
  }
  sheet.appendRow(rowValues);
}

function deleteEntry_(sheet, id){
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for(let i=1;i<data.length;i++){
    if(data[i][idCol] === id){
      sheet.deleteRow(i+1);
      return;
    }
  }
}

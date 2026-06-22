/**
 * 海のいきもの — 共有ランキング用 Google Apps Script（端末ID対応版）
 *
 * 特徴:
 *  - 1台のPC/スマホ＝1人として扱い（uid）、最初に登録した名前のまま、より良い記録で自動更新。
 *  - board="t"（タイピング, key=コース:秒数）と board="cw"（クロスワード進行度, key="progress"）に対応。
 *
 * 【更新手順】コードを貼り替えたら：デプロイ → デプロイを管理 → 鉛筆 → バージョン「新バージョン」→ デプロイ
 *  （URL は変わりません）
 */
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (id) ss = SpreadsheetApp.openById(id);
    else { ss = SpreadsheetApp.create('umi-leaderboard'); PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId()); }
  }
  var sh = ss.getSheetByName('scores');
  if (!sh) { sh = ss.insertSheet('scores'); sh.appendRow(['ts','board','key','uid','name','score','acc','kpm','label']); }
  return sh;
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e){
  var p = (e && e.parameter) || {};
  var board = String(p.board||'t'), key = String(p.key||'');
  var sh = sheet_(), rows = sh.getDataRange().getValues(), out=[];
  for (var i=1;i<rows.length;i++){
    var r=rows[i];
    if (String(r[1])===board && String(r[2])===key){
      out.push({ name:r[4], score:Number(r[5]), acc:Number(r[6]), kpm:Number(r[7]), label:r[8] });
    }
  }
  out.sort(function(a,b){ return b.score-a.score || b.acc-a.acc; });
  return json_(out.slice(0,20));
}

function doPost(e){
  var d; try{ d=JSON.parse(e.postData.contents); }catch(err){ return json_({ok:false,error:'bad json'}); }
  var board=String(d.board||'t'), key=String(d.key||''), uid=String(d.uid||'');
  var name=String(d.name||'ゲスト').slice(0,12);
  var score=Number(d.score)||0, acc=Number(d.acc)||0, kpm=Number(d.kpm)||0, label=String(d.label||'');
  var sh=sheet_(), rows=sh.getDataRange().getValues(), found=-1;
  for (var i=1;i<rows.length;i++){
    if (String(rows[i][1])===board && String(rows[i][2])===key && String(rows[i][3])===uid){ found=i; break; }
  }
  if (found<0){
    sh.appendRow([new Date(), board, key, uid, name, score, acc, kpm, label]);
    return json_({ok:true, action:'insert', name:name});
  }
  var oldScore=Number(rows[found][5])||0, oldAcc=Number(rows[found][6])||0;
  var better = (score>oldScore) || (score===oldScore && acc>oldAcc);
  if (better){
    var rownum=found+1;
    sh.getRange(rownum,6,1,4).setValues([[score,acc,kpm,label]]);
    sh.getRange(rownum,1).setValue(new Date());
  }
  return json_({ok:true, action: better?'update':'kept', name:rows[found][4]});
}

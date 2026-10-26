/* TSEKmate -> Google Sheets receiver. Paste into Extensions > Apps Script, then Deploy > Web app (Execute as: Me, Access: Anyone). */
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function col(n){var s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26)}return s}
function safe(v){v=String(v==null?'':v);return /^[=+\-@]/.test(v)?"'"+v:v}

function doGet(e){var c=e&&e.parameter&&e.parameter.c;
  if(!c)return out({ok:true,app:'TSEKmate',url:SpreadsheetApp.getActiveSpreadsheet().getUrl()});
  var v=pubSheet().getDataRange().getValues(),i;c=String(c).toUpperCase();
  for(i=1;i<v.length;i++)if(String(v[i][0])===c&&String(v[i][1])==='0'){var m=JSON.parse(v[i][4]);return out({ok:true,cls:v[i][2],roster:m.roster,at:m.at})}
  return out({ok:false,error:'Class not found. Ask your teacher to publish the results first.'})}
function pubSheet(){var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('_students');if(!sh){sh=ss.insertSheet('_students');sh.getRange(1,1,1,5).setValues([['Class code','Paper No','Name','PIN','Data']]);sh.hideSheet()}return sh}
function pub(d){var code=String(d.code||'').toUpperCase();if(!/^[A-Z0-9]{4,8}$/.test(code)||!d.students)throw new Error('Bad class data');
  var sh=pubSheet(),v=sh.getDataRange().getValues(),rows=[v[0]],i;
  for(i=1;i<v.length;i++)if(String(v[i][0])!==code)rows.push(v[i]);
  rows.push([code,0,String(d.cls||''),'',JSON.stringify({at:Date.now(),roster:d.students.map(function(s){return {n:s.n,name:s.name}})})]);
  d.students.forEach(function(s){rows.push([code,s.n,s.name,String(s.pin),JSON.stringify({res:s.res})])});
  sh.clearContents();var rg=sh.getRange(1,1,rows.length,5);rg.setNumberFormat('@');rg.setValues(rows);return {ok:true,count:d.students.length}}
function me(d){var v=pubSheet().getDataRange().getValues(),i,c=String(d.c||'').toUpperCase();
  for(i=1;i<v.length;i++)if(String(v[i][0])===c&&String(v[i][1])===String(d.n)){
    if(String(v[i][3])!==String(d.pin))return {ok:false,error:'That PIN is not right. Ask your teacher.'};
    return {ok:true,name:v[i][2],data:JSON.parse(v[i][4])}}
  return {ok:false,error:'Student not found.'}}

function doPost(e){
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    var d=JSON.parse(e.postData.contents),ss=SpreadsheetApp.getActiveSpreadsheet();
    if(d.type=='pub')return out(pub(d));if(d.type=='me')return out(me(d));
    if(!d.id||!d.key||!d.key.length)throw new Error('Incomplete data');
    var name=String(d.exam||'Exam').replace(/[\[\]*?:\/\\]/g,' ').substring(0,90);
    var sh=ss.getSheetByName(name)||setup(ss,name,d);
    var n=d.key.length,idc=6+n,last=col(5+n),rows=Math.max(sh.getLastRow()-3,1);
    var ids=sh.getRange(4,idc,rows,1).getValues(),r=0,i;
    for(i=0;i<ids.length;i++){if(ids[i][0]===d.id){r=4+i;break}}   /* same paper sent twice = update, never a duplicate */
    if(!r)r=Math.max(sh.getLastRow()+1,4);
    var ans=(d.answers||[]).slice(0,n);while(ans.length<n)ans.push('');
    sh.getRange(r,1,1,idc).setValues([[safe(d.student),'',d.total,'',new Date(d.at||Date.now())].concat(ans,[d.id])]);
    sh.getRange(r,2).setFormula('=SUMPRODUCT(--(F'+r+':'+last+r+'=F$2:'+last+'$2))');   /* edit an answer cell and the score updates */
    sh.getRange(r,4).setFormula('=IFERROR(B'+r+'/C'+r+',0)').setNumberFormat('0%');
    sh.getRange(r,5).setNumberFormat('dd mmm HH:mm');
    return out({ok:true,url:ss.getUrl()+'#gid='+sh.getSheetId(),row:r});
  }catch(err){return out({ok:false,error:String(err)})}
  finally{lock.releaseLock()}
}

function setup(ss,name,d){
  var n=d.key.length,sh=ss.insertSheet(name),last=col(5+n),hd=['Student','Score','Total','Percent','Scanned'],i;
  for(i=1;i<=n;i++)hd.push('Q'+i);hd.push('Paper ID');
  sh.getRange(1,1,1,hd.length).setValues([hd]).setFontWeight('bold').setBackground('#0B1530').setFontColor('#FFFFFF');
  sh.getRange(2,1,1,n+6).setValues([['ANSWER KEY','','','',''].concat(d.key,[''])]).setFontWeight('bold').setBackground('#FFF3BF');
  var f=['% CORRECT BY ITEM','','','',''];
  for(i=0;i<n;i++){var c=col(6+i);f.push('=IFERROR(SUMPRODUCT(--('+c+'$4:'+c+'$1000='+c+'$2))/COUNTA($A$4:$A$1000),"")')}
  f.push('');
  sh.getRange(3,1,1,f.length).setFormulas([f]).setNumberFormat('0%').setFontStyle('italic').setBackground('#E8EEFF');
  sh.setFrozenRows(3);sh.setFrozenColumns(1);sh.setColumnWidth(1,200);sh.hideColumns(6+n);
  var rg=sh.getRange('F4:'+last+'1000');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(F4<>"",F4=F$2)').setBackground('#C8F2DC').setRanges([rg]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(F4<>"",F4<>F$2)').setBackground('#FFD6D1').setRanges([rg]).build()]);
  return sh;
}

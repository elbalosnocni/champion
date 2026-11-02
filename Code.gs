/**
 * Champion – Google Apps Script backend
 * Cấu hình bắt buộc: Apps Script > Project Settings > Script Properties:
 * SHEET_ID = ID Google Spreadsheet
 * ADMIN_TOKEN = chuỗi bí mật ngẫu nhiên dành riêng cho BTC
 */
const CONFIG = {
  SHEET_ID: PropertiesService.getScriptProperties().getProperty('SHEET_ID') || '1j24IRF6hR5iD-aSM_5HS9Cv5oDK4TMayVzQIaB_5mfY',
  ADMIN_TOKEN: PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN') || ''
};
const TEAM_HEADERS = ['ID','TeamName','Captain','Players JSON','Status','Matches JSON'];
const MATCH_HEADERS = ['MatchID','HomeTeam','AwayTeam','HomeScore','AwayScore','Timeline JSON','UpdatedAt'];

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function spreadsheet_() {
  if (!CONFIG.SHEET_ID) throw new Error('Chưa cấu hình SHEET_ID.');
  return SpreadsheetApp.openById(CONFIG.SHEET_ID);
}
function teamsSheet_() {
  const ss=spreadsheet_();
  let sh=ss.getSheetByName('Teams');
  if (!sh) sh=ss.getActiveSheet();
  if (sh.getLastRow()===0) sh.appendRow(TEAM_HEADERS);
  return sh;
}
function matchesSheet_() {
  const ss=spreadsheet_();
  let sh=ss.getSheetByName('Matches');
  if (!sh) sh=ss.insertSheet('Matches');
  if (sh.getLastRow()===0) sh.appendRow(MATCH_HEADERS);
  return sh;
}
function parseJson_(value,fallback) {
  if (value===null || value==='' || typeof value==='undefined') return fallback;
  try { return JSON.parse(value); } catch (_) { return fallback; }
}
function readTeams_() {
  const values=teamsSheet_().getDataRange().getValues(), result=[];
  for (let i=1;i<values.length;i++) {
    const r=values[i];
    if (!r[1]) continue;
    result.push({id:String(r[0]),teamName:String(r[1]),captain:String(r[2]||''),
      players:parseJson_(r[3],[]),status:String(r[4]||'Chờ duyệt'),
      matches:parseJson_(r[5],[])});
  }
  return result;
}
function readMatches_() {
  const sh=matchesSheet_(), values=sh.getDataRange().getValues(), result=[];
  for(let i=1;i<values.length;i++) {
    const r=values[i]; if(!r[0]) continue;
    result.push({id:String(r[0]),homeTeam:String(r[1]),awayTeam:String(r[2]),
      homeScore:Number(r[3]),awayScore:Number(r[4]),timeline:parseJson_(r[5],[]),
      updatedAt:r[6] instanceof Date?r[6].toISOString():String(r[6]||'')});
  }
  return result;
}
function doGet() {
  try {
    const teams=readTeams_(), matches=readMatches_();
    // Hỗ trợ dữ liệu cũ: khôi phục trận từ JSON từng đội nếu sheet Matches chưa có.
    if (!matches.length) {
      const seen={};
      teams.forEach(t=>(t.matches||[]).forEach(m=>{
        const id=String(m.id); if(seen[id]) return;
        const home=m.role==='Home'?t.teamName:m.opponent;
        const away=m.role==='Away'?t.teamName:m.opponent;
        seen[id]={id:id,homeTeam:home,awayTeam:away,
          homeScore:Number(m.role==='Home'?m.goalsFor:m.goalsAgainst)||0,
          awayScore:Number(m.role==='Away'?m.goalsFor:m.goalsAgainst)||0,
          timeline:m.rawTimeline||[]};
      }));
      return json_({teams:teams.map(t=>Object.assign({},t,{matches:[]})),matches:Object.values(seen)});
    }
    return json_({teams:teams.map(t=>Object.assign({},t,{matches:[]})),matches:matches});
  } catch(err) { return json_({success:false,message:err.message}); }
}
function requireAdmin_(params) {
  if (!CONFIG.ADMIN_TOKEN) throw new Error('Máy chủ chưa cấu hình ADMIN_TOKEN trong Script Properties.');
  if (String(params.adminToken||'')!==CONFIG.ADMIN_TOKEN) throw new Error('Không có quyền thực hiện thao tác BTC.');
}
function cleanText_(value,label,max) {
  const s=String(value==null?'':value).trim();
  if(!s || s.length>max) throw new Error(label+' không hợp lệ (tối đa '+max+' ký tự).');
  return s;
}
function validatePlayers_(players) {
  if(!Array.isArray(players)||players.length<1||players.length>60) throw new Error('Danh sách cầu thủ phải có từ 1 đến 60 người.');
  return players.map(p=>({name:cleanText_(p&&p.name,'Tên cầu thủ',80),number:String(p&&p.number||'').trim().slice(0,8)}));
}
function saveMatch_(params) {
  requireAdmin_(params);
  const m=params.match||{};
  const id=cleanText_(m.id,'Mã trận',80);
  const home=cleanText_(m.homeTeam,'Đội nhà',80), away=cleanText_(m.awayTeam,'Đội khách',80);
  if(home===away) throw new Error('Hai đội phải khác nhau.');
  const hs=Number(m.homeScore), as=Number(m.awayScore);
  if(!Number.isInteger(hs)||!Number.isInteger(as)||hs<0||as<0||hs>99||as>99) throw new Error('Tỉ số phải là số nguyên từ 0 đến 99.');
  if(!Array.isArray(m.timeline)||m.timeline.length>300) throw new Error('Timeline không hợp lệ.');
  const timeline=m.timeline.map(e=>{
    const minute=Number(e.minute), team=String(e.team||''), type=String(e.type||'');
    if(!Number.isInteger(minute)||minute<1||minute>150) throw new Error('Phút sự kiện không hợp lệ.');
    if(team!==home&&team!==away) throw new Error('Sự kiện chứa đội không thuộc trận đấu.');
    if(['goal','yellow','red'].indexOf(type)<0) throw new Error('Loại sự kiện không hợp lệ.');
    return {minute:minute,team:team,type:type,player:cleanText_(e.player||'Chưa xác định','Tên cầu thủ',80)};
  });
  const homeGoals=timeline.filter(e=>e.type==='goal'&&e.team===home).length;
  const awayGoals=timeline.filter(e=>e.type==='goal'&&e.team===away).length;
  if(homeGoals!==hs||awayGoals!==as) throw new Error('Tỉ số không khớp số sự kiện bàn thắng trong timeline.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const teams=readTeams_();
    const ht=teams.find(t=>t.teamName===home), at=teams.find(t=>t.teamName===away);
    if(!ht||!at||ht.status!=='Đã duyệt'||at.status!=='Đã duyệt') throw new Error('Chỉ được lưu trận giữa hai đội đã được duyệt.');
    const sh=matchesSheet_(), last=sh.getLastRow();
    const row=last>=2?sh.getRange(2,1,last-1,MATCH_HEADERS.length).getValues():[];
    let target=-1;
    for(let i=0;i<row.length;i++) if(String(row[i][0])===id){target=i+2;break;}
    // ID hiện có chỉ được sửa nếu vẫn giữ nguyên cặp đội (không tráo vai).
    if(target>0) {
      const old=row[target-2];
      if(String(old[1])!==home||String(old[2])!==away) throw new Error('Không được đổi hai đội của trận đã lưu.');
    }
    const record=[id,home,away,hs,as,JSON.stringify(timeline),new Date()];
    if(target>0) sh.getRange(target,1,1,record.length).setValues([record]);
    else sh.appendRow(record);
    // Cập nhật bản sao tương thích ở cả hai dòng đội trong cùng khóa khóa.
    const teamSh=teamsSheet_(), values=teamSh.getDataRange().getValues();
    const homeEntry={id:id,role:'Home',opponent:away,goalsFor:hs,goalsAgainst:as,
      scorers:timeline.filter(e=>e.team===home&&e.type==='goal').map(e=>e.player),
      yellowCards:timeline.filter(e=>e.team===home&&e.type==='yellow').map(e=>e.player),
      redCards:timeline.filter(e=>e.team===home&&e.type==='red').map(e=>e.player),rawTimeline:timeline};
    const awayEntry={id:id,role:'Away',opponent:home,goalsFor:as,goalsAgainst:hs,
      scorers:timeline.filter(e=>e.team===away&&e.type==='goal').map(e=>e.player),
      yellowCards:timeline.filter(e=>e.team===away&&e.type==='yellow').map(e=>e.player),
      redCards:timeline.filter(e=>e.team===away&&e.type==='red').map(e=>e.player),rawTimeline:timeline};
    values.slice(1).forEach((r,i)=>{
      const entry=String(r[1])===home?homeEntry:String(r[1])===away?awayEntry:null;
      if(!entry)return;
      const oldMatches=parseJson_(r[5],[]).filter(x=>String(x.id)!==id);
      oldMatches.push(entry);teamSh.getRange(i+2,6).setValue(JSON.stringify(oldMatches));
    });
    return {success:true,message:'Đã lưu đồng bộ hai đội.',matchId:id};
  } finally { lock.releaseLock(); }
}
function doPost(e) {
  try {
    if(!e||!e.postData||!e.postData.contents) throw new Error('Thiếu dữ liệu gửi lên.');
    const params=JSON.parse(e.postData.contents);
    if(!params||typeof params!=='object') throw new Error('Payload không hợp lệ.');
    if(params.action==='register') {
      const name=cleanText_(params.teamName,'Tên đội',80), captain=cleanText_(params.captain,'Đội trưởng',80);
      const players=validatePlayers_(params.players);
      const lock=LockService.getScriptLock();lock.waitLock(30000);
      try {
        const sh=teamsSheet_(), rows=sh.getDataRange().getValues();
        if(rows.slice(1).some(r=>String(r[1]).toLocaleLowerCase()===name.toLocaleLowerCase())) throw new Error('Tên đội đã tồn tại.');
        sh.appendRow([cleanText_(params.id||Utilities.getUuid(),'ID',80),name,captain,JSON.stringify(players),'Chờ duyệt','[]']);
      } finally {lock.releaseLock();}
      return json_({success:true,message:'Đã nhận hồ sơ đăng ký.'});
    }
    if(params.action==='approve') {
      requireAdmin_(params);
      const lock=LockService.getScriptLock();lock.waitLock(30000);
      try {
        const sh=teamsSheet_(),rows=sh.getDataRange().getValues();let found=false;
        for(let i=1;i<rows.length;i++) if(String(rows[i][0])===String(params.id)){sh.getRange(i+1,5).setValue('Đã duyệt');found=true;break;}
        if(!found) throw new Error('Không tìm thấy đội cần duyệt.');
      } finally {lock.releaseLock();}
      return json_({success:true,message:'Đã duyệt đội.'});
    }
    if(params.action==='saveMatch') return json_(saveMatch_(params));
    throw new Error('Action không được hỗ trợ.');
  } catch(err) { return json_({success:false,message:err.message||String(err)}); }
}

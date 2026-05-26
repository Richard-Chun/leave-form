/* =========================================================
   柔美飯店 假單產生器  ·  核心邏輯
   - 解析「柔美班表」Google Sheet（姓名 + 月/日 橫向格式）
   - 也支援上傳同格式的 Excel
   - 依員工 + 日期自動帶入班別時間
   - 產生 請假單 / 加班單 / 未刷卡證明單（A4 列印清單）
   ========================================================= */

// 假單對象：用「姓名關鍵字」比對班表 C 欄，自動忽略姓名後的數字（如「小涵18」→「小涵」）
const TARGET_NAMES = ['緯宸','順正','小涵','淑雲','婉茹','玉美','玉樺','東志','敏智','莉莉'];
const DEPT = '房務';   // 職稱/單位預設

// 全域狀態
let SCHEDULE = {};   // { '小涵': { name:'小涵', raw:'小涵18', days:{ '5/1':'08-17', '5/2':'休', ... } } }
let state = {
  source:'gs', empName:null, dayKey:null, formType:'leave',
  rocYear:115, month:null
};

/* ---------- 工具：判斷與解析時間 ---------- */
// 是否為「非上班」（休假、代碼、空白）：只要格子裡沒有「時間樣式」就視為非上班
function isOff(code){
  if(code==null) return true;
  let c = String(code).trim();
  if(c==='') return true;
  // 去掉加班註記（+2H、+2.5H、+1H）與前綴註記（卡/國加/休加/加）後再看
  c = c.replace(/\s*\+\d+(\.\d+)?\s*H/gi,'').trim();
  c = c.replace(/^(卡|國加|休加|加)\s*/,'').trim();
  if(c==='') return true;
  // 含「數字-數字」或「數字/數字」或換行兩段數字 → 是上班時間
  if(/\d{1,2}\s*[-\/\r\n]\s*\d{1,2}/.test(c)) return false;
  return true;   // 其餘（休/例/國/年/套/三套/龍/寧夏/柔/xx/取消…）皆視為非上班
}

// 把班表格子解析成 {start:'08:00', end:'17:00'}；非上班回傳 null
function parseShift(code){
  if(code==null) return null;
  let s = String(code).trim();
  s = s.replace(/\s*\+\d+(\.\d+)?\s*H/gi,'').trim();   // 去加班註記 +2H/+2.5H
  s = s.replace(/^(卡|國加|休加|加)\s*/,'').trim();      // 去前綴註記 卡/國加/休加
  if(isOff(s)) return null;
  // 統一各種分隔符（含換行、空白、全形）為「-」
  s = s.replace(/[／]/g,'/').replace(/[–—~～]/g,'-')
       .replace(/[\r\n\t ]+/g,'-');
  let parts = s.split(/[\/\-]+/).filter(Boolean);
  if(parts.length < 2) return null;
  const a = toHM(parts[0]);
  const b = toHM(parts[1]);
  if(!a || !b) return null;
  return {start:a, end:b};
}

// 把 "08"→08:00, "1030"→10:30, "8"→08:00, "0930"→09:30, "17"→17:00
function toHM(raw){
  let t = String(raw).replace(/[^\d]/g,'');
  if(t==='') return null;
  let h, m;
  if(t.length<=2){ h=parseInt(t,10); m=0; }
  else if(t.length===3){ h=parseInt(t.slice(0,1),10); m=parseInt(t.slice(1),10); } // 930→9:30
  else { h=parseInt(t.slice(0,2),10); m=parseInt(t.slice(2,4),10); }
  if(isNaN(h)||isNaN(m)||h>24||m>59) return null;
  return String(h).padStart(2,'0')+':'+String(m).padStart(2,'0');
}

function hmParts(hm){ // '08:30' → {h:'08', m:'30'}
  if(!hm) return {h:'',m:''};
  const [h,m] = hm.split(':');
  return {h, m};
}

/* ---------- 載入來源切換 ---------- */
function switchSource(src){
  state.source = src;
  document.getElementById('tabFile').classList.toggle('on', src==='file');
  document.getElementById('tabGS').classList.toggle('on', src==='gs');
  document.getElementById('srcFile').style.display = src==='file'?'block':'none';
  document.getElementById('srcGS').style.display   = src==='gs'?'block':'none';
}

function setStatus(msg, kind){
  const el = document.getElementById('loadStatus');
  el.className = 'status show '+kind;
  el.innerHTML = (kind==='ok'?'✓ ':kind==='err'?'⚠ ':'⏳ ')+msg;
}

/* ---------- Excel 上傳 ---------- */
const fileInput = document.getElementById('fileInput');
const dropZone  = document.getElementById('dropZone');
fileInput.addEventListener('change', e=>{ if(e.target.files[0]) readExcel(e.target.files[0]); });
['dragover','dragleave','drop'].forEach(ev=>{
  dropZone.addEventListener(ev, e=>{
    e.preventDefault();
    if(ev==='dragover') dropZone.style.borderColor='var(--brand)';
    else dropZone.style.borderColor='var(--line)';
    if(ev==='drop' && e.dataTransfer.files[0]) readExcel(e.dataTransfer.files[0]);
  });
});

function readExcel(file){
  setStatus('讀取中…','load');
  const reader = new FileReader();
  reader.onload = e=>{
    try{
      const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array'});
      // 找含目標姓名最多的分頁；找不到就用第一個
      let best = wb.SheetNames[0], bestHit = -1;
      for(const sn of wb.SheetNames){
        const rws = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:null});
        let hit=0;
        rws.forEach(r=>(r||[]).forEach(c=>{ const v=String(c==null?'':c);
          if(TARGET_NAMES.some(t=>v.includes(t))) hit++; }));
        if(hit>bestHit){ bestHit=hit; best=sn; }
      }
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[best], {header:1, raw:false, defval:null});
      parseScheduleRows(rows);
    }catch(err){
      console.error(err); setStatus('讀取失敗：'+err.message,'err');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------- Google Sheet 公開讀取 ---------- */
let GS_ID = null;   // 記住試算表 ID，供切換分頁用

async function loadFromGoogleSheet(){
  const url = document.getElementById('gsUrl').value.trim();
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if(!m){ setStatus('連結格式不正確，請貼完整的 Google Sheet 網址','err'); return; }
  GS_ID = m[1];
  const gidMatch = url.match(/[#&?]gid=(\d+)/);
  setStatus('讀取 Google Sheet 中…','load');
  // 依序嘗試：連結指定的分頁(gid) → 預設第一個分頁
  const tries = [];
  if(gidMatch) tries.push(`https://docs.google.com/spreadsheets/d/${GS_ID}/export?format=csv&gid=${gidMatch[1]}`);
  tries.push(`https://docs.google.com/spreadsheets/d/${GS_ID}/export?format=csv`);
  for(const u of tries){
    try{
      const res = await fetch(u);
      if(!res.ok) continue;
      const text = await res.text();
      if(text.trim().startsWith('<')) continue; // 拿到HTML=權限不足
      const rows = csvToRows(text);
      parseScheduleRows(rows);
      return;
    }catch(e){ /* try next */ }
  }
  setStatus('無法讀取，請確認共用設定為「知道連結可檢視」','err');
}

function csvToRows(text){
  // 簡易 CSV 解析（支援引號內逗號與換行）
  const rows=[]; let row=[], cur='', q=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i];
    if(q){
      if(ch==='"'){ if(text[i+1]==='"'){cur+='"';i++;} else q=false; }
      else cur+=ch;
    }else{
      if(ch==='"') q=true;
      else if(ch===','){ row.push(cur); cur=''; }
      else if(ch==='\n'){ row.push(cur); rows.push(row); row=[]; cur=''; }
      else if(ch==='\r'){}
      else cur+=ch;
    }
  }
  if(cur!==''||row.length){ row.push(cur); rows.push(row); }
  return rows;
}

/* ---------- 解析班表（核心：姓名比對 + 月/日對應）---------- */
function parseScheduleRows(rows){
  SCHEDULE = {};
  // 1) 找日期列：某一列出現 ≥8 個「月/日」樣式（如 5/1、1/13）
  let dateRow = -1, dateCols = {};   // dateCols: { 欄索引: {mon, day, key:'5/1'} }
  for(let r=0;r<rows.length;r++){
    let cnt=0, cols={};
    (rows[r]||[]).forEach((c,ci)=>{
      const mm = String(c==null?'':c).trim().match(/^(\d{1,2})\/(\d{1,2})$/);
      if(mm){ cnt++; cols[ci]={mon:+mm[1], day:+mm[2], key:`${+mm[1]}/${+mm[2]}`}; }
    });
    if(cnt>=8){ dateRow=r; dateCols=cols; break; }
  }
  if(dateRow<0){
    setStatus('找不到日期列（需有「月/日」如 5/1 的橫向日期）','err'); return;
  }

  // 偵測表頭涵蓋的月份（取出現最多次的月份當預設）
  const monthCount={};
  Object.values(dateCols).forEach(d=>{ monthCount[d.mon]=(monthCount[d.mon]||0)+1; });
  const mainMonth = +Object.keys(monthCount).sort((a,b)=>monthCount[b]-monthCount[a])[0];

  // 2) 找姓名欄：掃前幾欄，哪一欄最常出現目標名字，就當姓名欄
  let nameCol = 0, bestHit = -1;
  for(let ci=0; ci<6; ci++){
    let hit=0;
    for(let r=0;r<rows.length;r++){
      const v=String((rows[r]||[])[ci]||'').trim();
      if(TARGET_NAMES.some(t=>v.includes(t))) hit++;
    }
    if(hit>bestHit){ bestHit=hit; nameCol=ci; }
  }

  // 3) 逐列抓目標員工（姓名比對，自動去掉姓名後的數字/空白）
  let found=0;
  for(let r=0;r<rows.length;r++){
    const raw = String((rows[r]||[])[nameCol]||'').trim();
    if(!raw) continue;
    const hit = TARGET_NAMES.find(t=>raw.includes(t));
    if(!hit) continue;
    if(SCHEDULE[hit]) continue;   // 同名只取第一次
    const days={};
    for(const ci in dateCols){
      const v = (rows[r]||[])[ci];
      days[dateCols[ci].key] = v==null?'':String(v).trim();
    }
    SCHEDULE[hit] = {name:hit, raw, days};
    found++;
  }

  if(found===0){ setStatus('班表中找不到指定的員工（請確認姓名與分頁）','err'); return; }
  state.month = mainMonth;
  document.getElementById('month').value = mainMonth;
  setStatus(`成功載入 ${found} 位員工（${mainMonth} 月班表）`,'ok');
  buildEmployeeSelect();
  const s2=document.getElementById('step2'); s2.style.opacity=1; s2.style.pointerEvents='auto';
}

function buildEmployeeSelect(){
  const sel = document.getElementById('empSelect');
  sel.innerHTML = '<option value="">— 請選擇員工 —</option>';
  // 依 TARGET_NAMES 的順序列出有抓到的員工
  TARGET_NAMES.forEach(nm=>{
    if(SCHEDULE[nm]) sel.innerHTML += `<option value="${nm}">${nm}</option>`;
  });
}

/* ---------- 選員工 / 選日期 ---------- */
function onEmpChange(){
  state.empName = document.getElementById('empSelect').value || null;
  state.dayKey = null;
  refreshDays();
}
function refreshDays(){
  const daySel = document.getElementById('daySelect');
  const preview = document.getElementById('schedPreview');
  if(!state.empName){ daySel.innerHTML='<option>—</option>'; preview.style.display='none'; return; }
  const emp = SCHEDULE[state.empName];
  daySel.innerHTML = '<option value="">— 請選擇日期 —</option>';
  // days 的 key 是「月/日」，依日期排序列出
  const keys = Object.keys(emp.days).sort((a,b)=>{
    const [ma,da]=a.split('/').map(Number), [mb,db]=b.split('/').map(Number);
    return ma-mb || da-db;
  });
  for(const key of keys){
    const code = emp.days[key];
    const sh = parseShift(code);
    const tag = sh ? `${sh.start}-${sh.end}` : (code||'—');
    const d = key.split('/')[1];
    daySel.innerHTML += `<option value="${key}">${d} 日 ｜ ${tag||'(空)'}</option>`;
  }
  renderSchedPreview(emp);
}
function renderSchedPreview(emp){
  const box = document.getElementById('schedDays');
  const preview = document.getElementById('schedPreview');
  box.innerHTML='';
  const keys = Object.keys(emp.days).sort((a,b)=>{
    const [ma,da]=a.split('/').map(Number), [mb,db]=b.split('/').map(Number);
    return ma-mb || da-db;
  });
  for(const key of keys){
    const code = emp.days[key];
    const off = isOff(code);
    const sh = parseShift(code);
    const txt = off? (code||'—') : (sh?`${sh.start.replace(':','')}` :code);
    const d = key.split('/')[1];
    const div = document.createElement('div');
    div.className = 'sd'+(off?' off':'')+(state.dayKey===key?' sel':'');
    div.innerHTML = `<div class="d">${d}</div><div class="t">${txt||''}</div>`;
    div.onclick = ()=>{ document.getElementById('daySelect').value=key; onDayChange(); };
    box.appendChild(div);
  }
  preview.style.display='block';
}
function onDayChange(){
  state.dayKey = document.getElementById('daySelect').value || null;
  if(state.dayKey){ state.month = +state.dayKey.split('/')[0]; }
  state.rocYear = parseInt(document.getElementById('rocYear').value,10)||115;
  if(state.empName) renderSchedPreview(SCHEDULE[state.empName]);
  if(state.dayKey){
    const s3=document.getElementById('step3'); s3.style.opacity=1; s3.style.pointerEvents='auto';
    renderDynFields();
  }
}

/* ---------- 假單類型切換 ---------- */
function setFormType(t){
  state.formType = t;
  ['leave','ot','miss'].forEach(k=>{
    document.getElementById('ft'+(k==='leave'?'Leave':k==='ot'?'OT':'Miss'))
      .classList.toggle('on', k===t);
  });
  renderDynFields();
}

// 取得目前選定日的資訊
function currentShift(){
  if(!state.empName||!state.dayKey) return null;
  return parseShift(SCHEDULE[state.empName].days[state.dayKey]);
}
// 選定日的時間：當天有班就用當天；當天是休假類(年/國/休…)就用該員工「平常標準工時」
function effectiveShift(){
  const s = currentShift();
  if(s) return s;
  return standardShift(state.empName);   // 回退到推算的標準工時
}
// 從該員工當月班表，推算最常出現的上下班時段當作「標準工時」
function standardShift(name){
  const emp = SCHEDULE[name];
  if(!emp) return null;
  const tally = {};   // 'start|end' -> 次數
  for(const key in emp.days){
    const sh = parseShift(emp.days[key]);
    if(!sh) continue;
    const k = sh.start+'|'+sh.end;
    tally[k] = (tally[k]||0)+1;
  }
  let best=null, max=0;
  for(const k in tally){ if(tally[k]>max){ max=tally[k]; best=k; } }
  if(!best) return null;
  const [start,end] = best.split('|');
  return {start, end};
}
function currentCode(){
  if(!state.empName||!state.dayKey) return '';
  return SCHEDULE[state.empName].days[state.dayKey]||'';
}
function currentDay(){   // 取「日」數字
  return state.dayKey ? +state.dayKey.split('/')[1] : '';
}

/* =========================================================
   動態輸入欄位（依假單類型，自動帶入班表時間，可手改）
   ========================================================= */
function renderDynFields(){
  const box = document.getElementById('dynFields');
  const todayShift = currentShift();      // 當天班別時間（休假類為 null）
  const sh = effectiveShift();            // 帶入用：休假日回退到標準工時
  const code = currentCode();
  // 提示：當天是休假類但有推算到標準工時 → 告知已帶平常時段；完全推不到 → 告知留白
  let offNote = '';
  if(!todayShift){
    if(sh){
      offNote = `<div class="mini-note">該日班表為「<b>${code||'—'}</b>」，已自動帶入 <b>${sh.start}–${sh.end}</b>（${state.empName} 平常上班時段），可自行修改。</div>`;
    }else{
      offNote = `<div class="mini-note">該日班表為「<b>${code||'—'}</b>」，且無法從班表推算平常時段，時間已留白供您手動填寫。</div>`;
    }
  }

  if(state.formType==='leave'){
    box.innerHTML = `
      ${offNote}
      <label class="fld">職稱</label>
      <input id="f_title" value="${DEPT}">
      <label class="fld">請假起 — 時 / 分</label>
      <div class="row2">
        <input id="f_sh" type="number" placeholder="時" value="${sh?+sh.start.split(':')[0]:''}">
        <input id="f_sm" type="number" placeholder="分" value="${sh?+sh.start.split(':')[1]:''}">
      </div>
      <label class="fld">請假迄 — 時 / 分</label>
      <div class="row2">
        <input id="f_eh" type="number" placeholder="時" value="${sh?+sh.end.split(':')[0]:''}">
        <input id="f_em" type="number" placeholder="分" value="${sh?+sh.end.split(':')[1]:''}">
      </div>
      <label class="fld">迄日（跨日請改，預設同起日）</label>
      <input id="f_endDay" type="number" value="${currentDay()||''}">
      <label class="fld">假別</label>
      <select id="f_reason">
        <option>事假</option><option>病假</option>
        <option selected>特休假</option><option>產假</option>
        <option>婚假</option><option>喪假</option><option>其他</option>
      </select>
      <label class="fld">共計</label>
      <div class="row3">
        <input id="f_days" type="number" placeholder="日" step="0.5">
        <input id="f_hrs" type="number" placeholder="時">
        <input id="f_mins" type="number" placeholder="分">
      </div>
      <label class="fld">職務代理人</label>
      <input id="f_agent" placeholder="（可留空，手寫）">
      <label class="fld">備註</label>
      <input id="f_note" placeholder="（可留空）">`;
  }
  else if(state.formType==='ot'){
    box.innerHTML = `
      <div class="mini-note">加班通常不在原班表內，時間請手動輸入。姓名、職稱已自動帶入。</div>
      <label class="fld">職稱</label>
      <input id="f_title" value="${DEPT}">
      <label class="fld">加班起 — 時 / 分</label>
      <div class="row2">
        <input id="f_sh" type="number" placeholder="時" value="${sh?+sh.end.split(':')[0]:''}">
        <input id="f_sm" type="number" placeholder="分" value="${sh?+sh.end.split(':')[1]:''}">
      </div>
      <label class="fld">加班迄 — 時 / 分</label>
      <div class="row2">
        <input id="f_eh" type="number" placeholder="時">
        <input id="f_em" type="number" placeholder="分">
      </div>
      <label class="fld">共計（時 / 分）</label>
      <div class="row2">
        <input id="f_oth" type="number" placeholder="時" step="0.5">
        <input id="f_otm" type="number" placeholder="分">
      </div>
      <label class="fld">加班事由</label>
      <input id="f_reason" value="人力需求">
      <label class="fld">補休 / 加班費</label>
      <select id="f_comp">
        <option value="補休">補休</option>
        <option value="加班費">加班費</option>
      </select>
      <label class="fld">備註</label>
      <input id="f_note" placeholder="（可留空）">`;
  }
  else { // miss 未刷卡
    box.innerHTML = `
      <label class="fld">職稱 / 單位</label>
      <input id="f_title" value="${DEPT}">
      <label class="fld">未刷卡別</label>
      <div class="row3">
        <label style="font-size:12px;display:flex;align-items:center;gap:4px"><input type="radio" name="misskind" value="上班" checked style="width:auto"> 上班</label>
        <label style="font-size:12px;display:flex;align-items:center;gap:4px"><input type="radio" name="misskind" value="下班" style="width:auto"> 下班</label>
        <label style="font-size:12px;display:flex;align-items:center;gap:4px"><input type="radio" name="misskind" value="休息空班" style="width:auto"> 空班</label>
      </div>
      <label class="fld">未刷卡時間（時 / 分）— 預設帶入上班時間</label>
      <div class="row2">
        <input id="f_mh" type="number" placeholder="時" value="${sh?+sh.start.split(':')[0]:''}">
        <input id="f_mm" type="number" placeholder="分" value="${sh?+sh.start.split(':')[1]:''}">
      </div>
      <label class="fld">原因</label>
      <select id="f_reason">
        <option selected>忘記打卡(漏打卡)</option>
        <option>打錯卡</option>
        <option>打卡時間重疊</option>
        <option>打卡鐘色帶不清</option>
        <option>打卡鐘故障</option>
        <option>其他</option>
      </select>`;
    // 監看上班/下班切換 → 自動換帶上班或下班時間
    setTimeout(()=>{
      document.querySelectorAll('input[name="misskind"]').forEach(r=>{
        r.addEventListener('change',()=>{
          const sh2=currentShift(); if(!sh2) return;
          const t = r.value==='下班'? sh2.end : sh2.start;
          if(document.querySelector('input[name="misskind"]:checked').value===r.value){
            document.getElementById('f_mh').value=+t.split(':')[0];
            document.getElementById('f_mm').value=+t.split(':')[1];
          }
        });
      });
    },0);
  }
}

/* =========================================================
   產生假單（寫入右側 A4）
   ========================================================= */
function val(id,d=''){ const e=document.getElementById(id); return e?e.value:d; }
function pad(n){ return String(n).padStart(2,'0'); }

function empName(){ return SCHEDULE[state.empName]?.name || ''; }
function rocY(){ return parseInt(val('rocYear'),10)||state.rocYear||115; }
function mon(){ return parseInt(val('month'),10)||state.month||''; }

/* =========================================================
   列印清單（收集籃）：每張假單獨立，列印時每 2 張排成一張 A4
   ========================================================= */
let PRINT_LIST = [];   // [{html, label}]

function generate(){
  if(!state.empName||!state.dayKey){ alert('請先選擇員工與日期'); return; }
  let unit='', kind='';
  if(state.formType==='leave'){ unit=buildLeave(); kind='請假單'; }
  else if(state.formType==='ot'){ unit=buildOT(); kind='加班單'; }
  else { unit=buildMiss(); kind='未刷卡證明單'; }

  const label = `${empName()}｜${rocY()}年${mon()}月${currentDay()}日｜${kind}`;
  PRINT_LIST.push({html:unit, label});
  renderStage();
}

// 渲染右側：上方清單管理 + 下方 A4 分頁預覽
function renderStage(){
  const head = document.getElementById('stageHead');
  const body = document.getElementById('stageBody');

  if(PRINT_LIST.length===0){
    head.style.display='none';
    body.innerHTML = `<div class="empty-stage"><div>
      <div class="big">🗂️</div><h3>列印清單是空的</h3>
      <p>依左側步驟填好一張假單後按「加入列印清單」。可重複加入不同員工、不同類別的假單，最後一起列印——系統會每 2 張排成一張 A4。</p>
    </div></div>`;
    return;
  }

  head.style.display='flex';
  document.getElementById('stageTitle').innerHTML =
    `列印清單 <span>共 ${PRINT_LIST.length} 張 · ${Math.ceil(PRINT_LIST.length/2)} 頁 A4</span>`;

  // 清單管理列（不會被列印）
  const listRows = PRINT_LIST.map((it,i)=>`
    <div class="qitem">
      <span class="qno">${i+1}</span>
      <span class="qlabel">${it.label}</span>
      <span class="qacts">
        <button onclick="moveItem(${i},-1)" ${i===0?'disabled':''} title="上移">↑</button>
        <button onclick="moveItem(${i},1)" ${i===PRINT_LIST.length-1?'disabled':''} title="下移">↓</button>
        <button onclick="removeItem(${i})" class="del" title="刪除">✕</button>
      </span>
    </div>`).join('');

  // A4 分頁：每 2 張一頁
  let pages='';
  for(let i=0;i<PRINT_LIST.length;i+=2){
    const a = PRINT_LIST[i].html;
    const b = PRINT_LIST[i+1] ? PRINT_LIST[i+1].html : '';
    pages += `<div class="a4">${a}${b}</div>`;
  }

  body.innerHTML = `
    <div class="queue-box no-print">
      <div class="queue-hd">列印清單（拖不會列印；✕ 可刪除、↑↓ 調順序）</div>
      ${listRows}
    </div>
    <div class="sheet-wrap">${pages}</div>`;
}

function removeItem(i){ PRINT_LIST.splice(i,1); renderStage(); }
function moveItem(i,dir){
  const j=i+dir; if(j<0||j>=PRINT_LIST.length) return;
  [PRINT_LIST[i],PRINT_LIST[j]]=[PRINT_LIST[j],PRINT_LIST[i]];
  renderStage();
}
function clearList(){
  if(PRINT_LIST.length && !confirm('確定清空列印清單？')) return;
  PRINT_LIST=[]; renderStage();
}

/* ---------- 請假單 ---------- */
function buildLeave(){
  const Y=rocY(), M=mon(), D=currentDay();
  const endDay = val('f_endDay')||D;
  const sh=val('f_sh'), sm=val('f_sm'), eh=val('f_eh'), em=val('f_em');
  const reason=val('f_reason'), agent=val('f_agent'), note=val('f_note');
  const days=val('f_days'), hrs=val('f_hrs'), mins=val('f_mins');
  const reasons=['事假','病假','特休假','產假','婚假','喪假','其他'];
  const ck = r => (r===reason?'■':'□')+' '+r;
  return `
  <div class="form-unit">
    <div class="fu-title">
      <div class="co">柔美飯店</div>
      <div class="nm">請假單</div>
    </div>
    <table class="ft">
      <tr><td colspan="4" style="border:none;padding:4px 8px">申請日期：　${Y}　年　${M}　月　${D}　日</td></tr>
    </table>
    <table class="ft">
      <tr>
        <td class="lb">姓　名</td><td class="editable" style="width:32%">${empName()}</td>
        <td class="lb">職　稱</td><td class="editable">${val('f_title')||DEPT}</td>
      </tr>
      <tr>
        <td class="lb">日　期</td>
        <td colspan="3" class="editable">自 ${Y} 年 ${M} 月 ${D} 日 ${sh||'　'} 時 ${sm||'　'} 分起，至 ${Y} 年 ${M} 月 ${endDay} 日 ${eh||'　'} 時 ${em||'　'} 分止</td>
      </tr>
      <tr>
        <td class="lb">共　計</td>
        <td colspan="3" class="editable">${days||'　'} 日 ／ ${hrs||'　'} 時 ${mins||'　'} 分</td>
      </tr>
      <tr>
        <td class="lb lb-sm">職務代理人</td>
        <td class="editable">${agent||''}</td>
        <td class="lb lb-sm">代理人簽名</td><td></td>
      </tr>
      <tr>
        <td class="lb">請假事由</td>
        <td colspan="3" class="ck editable" style="line-height:2">
          ${ck('事假')}　　${ck('病假')}　　${ck('特休假')}　　${ck('產假')}<br>
          ${ck('婚假')}　　${ck('喪假')}　　${ck('其他')} ${reason==='其他'?(note||''):''}
        </td>
      </tr>
      <tr>
        <td class="lb">備　註</td><td colspan="3" class="editable" style="height:34px">${reason!=='其他'?(note||''):''}</td>
      </tr>
      <tr class="sign-row">
        <td>副總經理：</td><td>權責主管：</td><td colspan="2">申請人：</td>
      </tr>
    </table>
  </div>`;
}

/* ---------- 加班單 ---------- */
function buildOT(){
  const Y=rocY(), M=mon(), D=currentDay();
  const sh=val('f_sh'), sm=val('f_sm'), eh=val('f_eh'), em=val('f_em');
  const oth=val('f_oth'), otm=val('f_otm');
  const reason=val('f_reason')||'人力需求';
  const comp=val('f_comp')||'補休';
  const note=val('f_note');
  return `
  <div class="form-unit">
    <div class="fu-title">
      <div class="co">柔美飯店</div>
      <div class="nm">加班單</div>
    </div>
    <table class="ft">
      <tr><td colspan="4" style="border:none;padding:4px 8px">申請日期：　${Y}　年　${M}　月　${D}　日</td></tr>
    </table>
    <table class="ft">
      <tr>
        <td class="lb">姓　名</td><td class="editable" style="width:32%">${empName()}</td>
        <td class="lb">職　稱</td><td class="editable">${val('f_title')||DEPT}</td>
      </tr>
      <tr>
        <td class="lb">加班時間</td>
        <td colspan="3" class="editable">自 ${Y} 年 ${M} 月 ${D} 日 ${sh||'　'} 時 ${sm||'　'} 分起，至 ${Y} 年 ${M} 月 ${D} 日 ${eh||'　'} 時 ${em||'　'} 分止</td>
      </tr>
      <tr>
        <td class="lb">共　計</td>
        <td colspan="3" class="editable">${oth||'　'} 時 ${otm||'　'} 分</td>
      </tr>
      <tr>
        <td class="lb">加班事由</td><td colspan="3" class="editable">${reason}</td>
      </tr>
      <tr>
        <td class="lb">權責主管<br>審　核</td>
        <td colspan="3" class="ck editable">■ 同意。　　□ 不同意。理由：</td>
      </tr>
      <tr>
        <td class="lb">備　註</td>
        <td colspan="3" class="ck editable">${comp==='補休'?'■':'□'} 補休。補休於　　　　　　　　　　　　　${comp==='加班費'?'■':'□'} 加班費。${note?'　'+note:''}</td>
      </tr>
      <tr class="sign-row">
        <td>副總經理：</td><td>權責主管：</td><td colspan="2">申請人：</td>
      </tr>
    </table>
  </div>`;
}

/* ---------- 未刷卡證明單 ---------- */
function buildMiss(){
  const Y=rocY(), M=mon(), D=currentDay();
  const kind = document.querySelector('input[name="misskind"]:checked')?.value || '上班';
  const mh=val('f_mh'), mm=val('f_mm');
  const reason=val('f_reason');
  const ck = k => (k===kind?'■':'□')+' '+k;
  const rk = r => (r===reason?'▓':'□');
  return `
  <div class="form-unit">
    <div class="fu-title">
      <div class="co">柔美飯店</div>
      <div class="nm" style="letter-spacing:4px">未刷卡證明單</div>
      <div class="dt">${Y} 年 ${M} 月 ${D} 日</div>
    </div>
    <table class="ft">
      <tr>
        <td class="lb">姓　名</td><td class="editable" style="width:28%">${empName()}</td>
        <td class="lb">單　位</td><td class="editable">${DEPT}</td>
      </tr>
      <tr>
        <td class="lb lb-sm">未刷卡時間</td>
        <td colspan="3" class="ck editable">
          ${ck('上班')}　${ck('下班')}　${ck('休息空班')}　　
          ${Y} 年 ${M} 月 ${D} 日 ${mh||'　'} 時 ${mm||'　'} 分
        </td>
      </tr>
      <tr>
        <td class="lb lb-sm">未刷卡原因<br>(請打 ˇ)</td>
        <td colspan="3" class="ck editable" style="line-height:1.9">
          ${rk('忘記打卡(漏打卡)')} 忘記打卡(漏打卡)。　${rk('打錯卡')} 打錯卡，打到　　　　　的出勤卡。<br>
          ${rk('打卡時間重疊')} 打卡時間重疊。　${rk('打卡鐘色帶不清')} 打卡鐘色帶不清。　${rk('打卡鐘故障')} 打卡鐘故障。<br>
          ${rk('其他')} 其他：
        </td>
      </tr>
      <tr>
        <td class="lb">說　明</td>
        <td colspan="3" class="note-cell">
          <ol>
            <li>漏打卡者如未填具「未刷卡證明單」以證明其出勤事實者，以曠職論。</li>
            <li>已填具者，每月漏打卡累計達三次（含）以上，無全勤獎金；達五次（含）以上，另記警告一次。</li>
            <li>主任級以上主管每月漏打卡超過三次(含)以上者，記警告一次。</li>
            <li>本單若無主管簽名，視為未填具，以曠職論。</li>
          </ol>
        </td>
      </tr>
      <tr class="sign-row">
        <td>申請人：</td><td>單位主管證明：</td><td colspan="2">館店最高主管審核：</td>
      </tr>
    </table>
  </div>`;
}

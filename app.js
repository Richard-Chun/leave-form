/* =========================================================
   柔美飯店 假單產生器  ·  核心邏輯
   - 解析 AUTO排班資料匯入.xlsx 的「房務」分頁班表貼上區
   - 支援 Excel 上傳 與 Google Sheet 公開連結
   - 依員工 + 日期自動帶入班別時間
   - 產生 請假單 / 加班單 / 未刷卡證明單（A4 可列印）
   ========================================================= */

// 只處理這 11 位房務人員（依需求）
const TARGET_EMPLOYEES = [
  {id:'R20', name:'小涵'}, {id:'R25', name:'淑雲'}, {id:'N7',  name:'婉茹'},
  {id:'F19', name:'玉美'}, {id:'R94', name:'玉樺'}, {id:'R27', name:'俊傑'},
  {id:'R59', name:'東志'}, {id:'F13', name:'敏智'}, {id:'R86', name:'莉莉'},
  {id:'R68', name:'蔡緯宸'}, {id:'R71', name:'蔡順正'}
];
const DEPT = '房務';   // 職稱/單位預設

// 非上班的代碼（休假類）— 這些日子沒有上下班時間
const OFF_CODES = ['休','例','國','休尚','例尚','補休','年','套','公假',
                   '休加','休加A','休加B','休加C','{res}','{sta}','{reg}'];

// 全域狀態
let SCHEDULE = {};        // { 'R20': { name:'小涵', days:{1:'08-17',2:'休',...} }, ... }
let state = {
  source:'file', empId:null, day:null, formType:'leave',
  rocYear:115, month:null
};

/* ---------- 工具：判斷與解析時間 ---------- */
function isOff(code){
  if(code==null) return true;
  const c = String(code).trim();
  if(c==='') return true;
  return OFF_CODES.includes(c);
}

// 把班表格子（如 "08-17"、"1030/1930"、"19-04"、"0930-1830"）解析成 {start:'08:00', end:'17:00'}
function parseShift(code){
  if(isOff(code)) return null;
  let s = String(code).trim();
  // 統一各種分隔符（含換行、空白、全形）為「-」：
  //   1030↵1930、1030/1930、08-17、0930-1830、19-04 都要能拆成兩段
  s = s.replace(/[／]/g,'/').replace(/[–—~～]/g,'-')
       .replace(/[\r\n\t ]+/g,'-');   // 換行/空白 → 分隔符
  // 形如 1030/1930 或 1030-1930 或 08-17 或 0930-1830 或 19-04
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
      const ws = wb.Sheets['房務'] || wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:null});
      parseScheduleRows(rows);
    }catch(err){
      console.error(err); setStatus('讀取失敗：'+err.message,'err');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------- Google Sheet 公開讀取 ---------- */
async function loadFromGoogleSheet(){
  const url = document.getElementById('gsUrl').value.trim();
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if(!m){ setStatus('連結格式不正確，請貼完整的 Google Sheet 網址','err'); return; }
  const id = m[1];
  const gidMatch = url.match(/[#&?]gid=(\d+)/);
  setStatus('讀取 Google Sheet 中…','load');
  // 優先嘗試「房務」分頁；若失敗則讀預設分頁
  const tries = [];
  tries.push(`https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent('房務')}`);
  if(gidMatch) tries.push(`https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gidMatch[1]}`);
  tries.push(`https://docs.google.com/spreadsheets/d/${id}/export?format=csv`);
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
  // 簡易 CSV 解析（支援引號內逗號）
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

/* ---------- 解析班表列（核心）---------- */
// 找到「班表貼上區」：表頭含 員工編號/員工姓名，後面 1..31 為日期欄
function parseScheduleRows(rows){
  SCHEDULE = {};
  // 找出表頭列（含「員工編號」或第一格是「班表貼上區」）
  let headerRow = -1, idCol = -1, nameCol = -1, dayCols = {};
  for(let r=0;r<rows.length;r++){
    const line = (rows[r]||[]).map(x=>x==null?'':String(x).trim());
    const joined = line.join('|');
    if(joined.includes('員工編號') && joined.includes('員工姓名')){
      headerRow = r;
      line.forEach((c,ci)=>{
        if(c==='員工編號') idCol=ci;
        else if(c==='員工姓名') nameCol=ci;
        else if(/^\d{1,2}$/.test(c)){ const d=parseInt(c,10); if(d>=1&&d<=31) dayCols[d]=ci; }
      });
      break; // 只取第一個（班表貼上區），複製區在下方不取
    }
  }
  if(headerRow<0 || idCol<0){
    setStatus('找不到班表貼上區（需有「員工編號/員工姓名」表頭）','err');
    return;
  }

  const targetIds = new Set(TARGET_EMPLOYEES.map(e=>e.id));
  let found=0;
  for(let r=headerRow+1;r<rows.length;r++){
    const line = rows[r]||[];
    const id = line[idCol]==null?'':String(line[idCol]).trim();
    // 遇到「複製區」區塊即停止（複製區的時間已轉成 {res} 等代碼，不適合假單）
    if(id==='複製區' || line.join('').includes('複製區')) break;
    if(id==='') continue;
    if(!targetIds.has(id)) continue;
    if(SCHEDULE[id]) continue;        // 已抓過 → 跳過重複貼上區的第二份
    const name = nameCol>=0 && line[nameCol]!=null ? String(line[nameCol]).trim() : id;
    const days={};
    for(const d in dayCols){
      const v = line[dayCols[d]];
      days[d] = v==null?'':String(v).trim();
    }
    SCHEDULE[id] = {name, days};
    found++;
    if(found>=TARGET_EMPLOYEES.length) break;   // 11 位抓滿即停
  }

  if(found===0){ setStatus('班表中找不到指定的 11 位房務人員','err'); return; }
  setStatus(`成功載入 ${found} 位員工班表`,'ok');
  buildEmployeeSelect();
  // 開放步驟2
  const s2=document.getElementById('step2'); s2.style.opacity=1; s2.style.pointerEvents='auto';
}

function buildEmployeeSelect(){
  const sel = document.getElementById('empSelect');
  sel.innerHTML = '<option value="">— 請選擇員工 —</option>';
  TARGET_EMPLOYEES.forEach(e=>{
    if(SCHEDULE[e.id]){
      const nm = SCHEDULE[e.id].name || e.name;
      sel.innerHTML += `<option value="${e.id}">${nm}（${e.id}）</option>`;
    }
  });
}

/* ---------- 選員工 / 選日期 ---------- */
function onEmpChange(){
  state.empId = document.getElementById('empSelect').value || null;
  refreshDays();
}
function refreshDays(){
  const daySel = document.getElementById('daySelect');
  const preview = document.getElementById('schedPreview');
  if(!state.empId){ daySel.innerHTML='<option>—</option>'; preview.style.display='none'; return; }
  const emp = SCHEDULE[state.empId];
  daySel.innerHTML = '<option value="">— 請選擇日期 —</option>';
  for(let d=1; d<=31; d++){
    const code = emp.days[d];
    if(code===undefined) continue;
    const sh = parseShift(code);
    const tag = isOff(code) ? (code||'—') : (sh?`${sh.start}-${sh.end}`:code);
    daySel.innerHTML += `<option value="${d}">${d} 日 ｜ ${tag||'(空)'}</option>`;
  }
  renderSchedPreview(emp);
}
function renderSchedPreview(emp){
  const box = document.getElementById('schedDays');
  const preview = document.getElementById('schedPreview');
  box.innerHTML='';
  for(let d=1; d<=31; d++){
    const code = emp.days[d];
    if(code===undefined) continue;
    const off = isOff(code);
    const sh = parseShift(code);
    const txt = off? (code||'—') : (sh?`${sh.start.replace(':','')}` :code);
    const div = document.createElement('div');
    div.className = 'sd'+(off?' off':'')+(state.day==d?' sel':'');
    div.innerHTML = `<div class="d">${d}</div><div class="t">${txt||''}</div>`;
    div.onclick = ()=>{ document.getElementById('daySelect').value=d; onDayChange(); };
    box.appendChild(div);
  }
  preview.style.display='block';
}
function onDayChange(){
  state.day = document.getElementById('daySelect').value || null;
  state.month = parseInt(document.getElementById('month').value,10)||state.month;
  state.rocYear = parseInt(document.getElementById('rocYear').value,10)||115;
  if(state.empId) renderSchedPreview(SCHEDULE[state.empId]);
  // 開放步驟3
  if(state.day){
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

// 取得目前選定日的班別時間
function currentShift(){
  if(!state.empId||!state.day) return null;
  const code = SCHEDULE[state.empId].days[state.day];
  return parseShift(code);
}

/* =========================================================
   動態輸入欄位（依假單類型，自動帶入班表時間，可手改）
   ========================================================= */
function renderDynFields(){
  const box = document.getElementById('dynFields');
  const sh = currentShift();           // {start,end} 或 null
  const code = state.empId&&state.day ? SCHEDULE[state.empId].days[state.day] : '';
  const offNote = (!sh) ? `<div class="mini-note">該日班表為「<b>${code||'—'}</b>」，非一般上下班時間，下方時間已留白供您手動填寫。</div>` : '';

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
      <input id="f_endDay" type="number" value="${state.day||''}">
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

function empName(){ return SCHEDULE[state.empId]?.name || ''; }
function rocY(){ return parseInt(val('rocYear'),10)||state.rocYear||115; }
function mon(){ return parseInt(val('month'),10)||state.month||''; }

/* =========================================================
   列印清單（收集籃）：每張假單獨立，列印時每 2 張排成一張 A4
   ========================================================= */
let PRINT_LIST = [];   // [{html, label}]

function generate(){
  if(!state.empId||!state.day){ alert('請先選擇員工與日期'); return; }
  let unit='', kind='';
  if(state.formType==='leave'){ unit=buildLeave(); kind='請假單'; }
  else if(state.formType==='ot'){ unit=buildOT(); kind='加班單'; }
  else { unit=buildMiss(); kind='未刷卡證明單'; }

  const label = `${empName()}｜${rocY()}年${mon()}月${state.day}日｜${kind}`;
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
  const Y=rocY(), M=mon(), D=state.day;
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
  const Y=rocY(), M=mon(), D=state.day;
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
  const Y=rocY(), M=mon(), D=state.day;
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

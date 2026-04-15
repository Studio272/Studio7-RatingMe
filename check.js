Chart.defaults.font.family = "'Sarabun', sans-serif";
Chart.defaults.color = '#9ca3af';

let state = { 
  sales: [], rating: [], fileNames: {sales:'',rating:''}, 
  selectedMonth: 'all', selectedYear: 'all', selectedBranch: 'all', selectedCat: 'all',
  availableMonths: new Set(), availableYears: new Set(), availableBranches: new Set(), availableCats: new Set()
};
let charts = {};

function switchTab(tabId, btn){
  document.querySelectorAll('.tab-pane').forEach(el=>el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el=>el.classList.remove('active'));
  document.getElementById('tab-'+tabId).classList.add('active');
  btn.classList.add('active');
  setTimeout(()=> { Object.values(charts).forEach(c=> { if(c) c.resize(); }); }, 50);
}

function classifyCategory(catName, subCat, mwatSinka) {
  let combined = (catName + " " + subCat + " " + mwatSinka).toLowerCase();
  let cCat = catName.toLowerCase();
  
  if (cCat.includes('mac') || combined.includes('mac')) return "Mac";
  if (cCat.includes('ipad') && !combined.includes('acc')) return "iPad";
  if (cCat.includes('iphone') && !combined.includes('acc')) return "iPhone";
  if (cCat.includes('watch')) return "Apple Watch";
  
  if (combined.includes('apple care') || combined.includes('applecare') || combined.includes('accessori') || combined.includes('apple acc')) return "BTB (Apple)";
  
  if (combined.includes('speaker') || combined.includes('gadget') || combined.includes('smart living') || combined.includes('health') || combined.includes('cases') || combined.includes('film') || combined.includes('smile') || combined.includes('non-apple')) return "BTB";
  
  return "Other";
}

const getVal = (v) => { if(typeof v === 'number') return v; if(!v) return 0; const p=parseFloat(String(v).replace(/,/g,'').replace(/[^\d.-]/g,'')); return isNaN(p)?0:p; };

function normalizeDate(val) {
  if(!val) return null;
  if(typeof val === 'number') return new Date(Math.round((val-25569)*86400*1000));
  let str = String(val).trim().replace(/^[\u0E00-\u0E7F]+\.\s*/,'');
  const m = str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if(m) { let d=m[1], mo=m[2], y=parseInt(m[3]); if(y>2500)y-=543; return new Date(`${y}-${mo.padStart(2,'0')}-${d.padStart(2,'0')}`); }
  const dObj = new Date(str.split(' ')[0]);
  if(!isNaN(dObj.getTime())) return dObj;
  return null;
}

const colorPalette = ['#1d4ed8','#7c3aed','#059669','#d97706','#e11d48','#0891b2','#6366f1','#84cc16','#f59e0b','#10b981'];

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
dropZone.addEventListener('dragover',(e)=>{e.preventDefault();dropZone.classList.add('dragover');});
dropZone.addEventListener('dragleave',()=>dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop',(e)=>{e.preventDefault();dropZone.classList.remove('dragover');handleFiles(e.dataTransfer.files);});
fileInput.addEventListener('change',(e)=>handleFiles(e.target.files));



const defaultUsers = { "admin": "1234", "studio7": "studio7" };
if(!localStorage.getItem('s7_users')) {
   localStorage.setItem('s7_users', JSON.stringify(defaultUsers));
}

function checkLogin() {
  if(sessionStorage.getItem('s7_logged_in') !== 'true') {
     document.getElementById('upload-view').style.display='none';
     document.getElementById('dashboard-view').style.display='none';
     if(document.getElementById('btn-admin')) document.getElementById('btn-admin').style.display='none';
  } else {
     document.getElementById('login-view').style.display='none';
     document.getElementById('upload-view').style.display='flex';
     if(sessionStorage.getItem('s7_role')==='admin') {
        document.getElementById('btn-admin').style.display = 'inline-block';
    }
    document.getElementById('btn-chgpass').style.display = 'inline-block';
  }
}

function toggleAuthView(viewName) {
    document.getElementById('box-login').style.display = 'none';
    document.getElementById('box-register').style.display = 'none';
    document.getElementById('box-forgot').style.display = 'none';
    
    // Reset inputs and messages
    ['regUser','regPass','forUser'].forEach(id => document.getElementById(id).value = '');
    ['reg-error','reg-success','for-error','for-success'].forEach(id => document.getElementById(id).style.display = 'none');
    document.getElementById('btnSubmitReg').style.display = 'block';
    document.getElementById('btnSubmitFor').style.display = 'block';
    
    document.getElementById('box-' + viewName).style.display = 'block';
}

function togglePass(inputId, btn) {
    const el = document.getElementById(inputId);
    if(el.type === 'password') {
        el.type = 'text';
        btn.textContent = '🙈';
    } else {
        el.type = 'password';
        btn.textContent = '👁️';
    }
}

function submitRegister() {
    const u = document.getElementById('regUser').value.trim().toLowerCase();
    const p = document.getElementById('regPass').value.trim();
    if(!u || !p) {
        document.getElementById('reg-error').textContent = '⚠️ กรุณากรอกขัอมูลให้ครบถ้วน';
        document.getElementById('reg-error').style.display = 'block';
        return;
    }
    
    // Check if user exists anywhere
    let users = {}; let pending = {};
    try { users = JSON.parse(localStorage.getItem('s7_users')) || {}; } catch(e){}
    try { pending = JSON.parse(localStorage.getItem('s7_pending_users')) || {}; } catch(e){}
    
    if(users[u] || pending[u] || u === 'admin' || u === 'studio7') {
        document.getElementById('reg-error').textContent = '❌ ชื่อผู้ใช้งานนี้ถูกใช้ไปแล้ว';
        document.getElementById('reg-error').style.display = 'block';
        return;
    }
    
    // Add to pending
    pending[u] = p;
    localStorage.setItem('s7_pending_users', JSON.stringify(pending));
    
    // Success State
    document.getElementById('reg-error').style.display = 'none';
    document.getElementById('reg-success').style.display = 'block';
    document.getElementById('btnSubmitReg').style.display = 'none';
}

function submitForgot() {
    const u = document.getElementById('forUser').value.trim().toLowerCase();
    if(!u) {
        document.getElementById('for-error').textContent = '⚠️ กรุณากรอกรหัสพนักงาน (Username)';
        document.getElementById('for-error').style.display = 'block';
        return;
    }
    
    let resets = [];
    try { resets = JSON.parse(localStorage.getItem('s7_reset_requests')) || []; } catch(e){}
    
    if(!resets.includes(u)) {
        resets.push(u);
        localStorage.setItem('s7_reset_requests', JSON.stringify(resets));
    }
    
    // Success State
    document.getElementById('for-error').style.display = 'none';
    document.getElementById('for-success').style.display = 'block';
    document.getElementById('btnSubmitFor').style.display = 'none';
}

function handleLogin() {
  const u = document.getElementById('loginUser').value.trim().toLowerCase();
  const p = document.getElementById('loginPass').value.trim();
  
  let valid = false;
  if (u === 'admin' && p === '1234') valid = true;
  if (u === 'studio7' && p === 'studio7') valid = true;
  
  if(!valid) {
      try {
          let users = JSON.parse(localStorage.getItem('s7_users'));
          if (users && users[u] && users[u] === p) valid = true;
      } catch(e) {}
  }
  
  if(valid) {
     sessionStorage.setItem('s7_logged_in', 'true');
     sessionStorage.setItem('s7_logged_user', u);
     sessionStorage.setItem('s7_session_start', Date.now()); // Telemetry Tracking Start
     if(u === 'admin') sessionStorage.setItem('s7_role', 'admin');
     else sessionStorage.setItem('s7_role', 'user');
     
     document.getElementById('login-error').style.display='none';
     document.getElementById('login-view').style.animation = 'fadeIn 0.3s reverse forwards';
     setTimeout(()=>{
         document.getElementById('login-view').style.display='none';
         document.getElementById('upload-view').style.display='flex';
         if(u === 'admin') {
             const btn = document.getElementById('btn-admin');
             if(btn) btn.style.display='inline-block';
         }
         const btnCp = document.getElementById('btn-chgpass');
         if(btnCp) btnCp.style.display='inline-block';
     }, 250);
  } else {
     document.getElementById('login-error').style.display='block';
     document.getElementById('login-view').animate([{transform:'translateX(-5px)'},{transform:'translateX(5px)'},{transform:'translateX(-5px)'},{transform:'translateX(5px)'},{transform:'translateX(0)'}], {duration:300});
  }
}

function openAdmin() { 
   renderAdminUsers(); 
   renderAdminPendings();
   renderAdminResets();
   renderAdminTelemetry();
   const apiEl = document.getElementById('gemini-api-key');
   if(apiEl) apiEl.value = localStorage.getItem('s7_gemini_api') || '';
   document.getElementById('adminModal').classList.add('active'); 
}
function saveApiKey() {
   let key = document.getElementById('gemini-api-key').value.trim();
   localStorage.setItem('s7_gemini_api', key);
   alert('บันทึก Gemini API Key เรียบร้อยแล้ว');
}
function closeAdmin() { document.getElementById('adminModal').classList.remove('active'); }

function renderAdminUsers() {
   let users = JSON.parse(localStorage.getItem('s7_users')) || {};
   let tbody = document.getElementById('admin-tbody');
   if(!tbody) return;
   tbody.innerHTML = '';
   for(let k in users) {
      tbody.innerHTML += `<tr>
         <td style="padding:12px; border-bottom:1px solid var(--border2);">${k}</td>
         <td style="padding:12px; border-bottom:1px solid var(--border2);">${k==='admin'?'****':users[k]}</td>
         <td style="padding:12px; text-align:right; border-bottom:1px solid var(--border2);">
            ${k==='admin' ? '<span style="color:var(--dim)">System Default</span>' : `<button onclick="deleteUser('${k}')" style="background:#fee2e2;color:#dc2626;border:none;border-radius:4px;padding:4px 8px;cursor:pointer;font-family:'Space Grotesk',sans-serif;font-weight:600;font-size:11px;">Remove</button>`}
         </td>
      </tr>`;
   }
}
function addNewUser() {
   let nu = document.getElementById('new-user').value.trim().toLowerCase();
   let np = document.getElementById('new-pass').value.trim();
   if(!nu || !np) return alert('โปรดกรอกข้อมูล Username และ Password ให้ครบถ้วน');
   let users = JSON.parse(localStorage.getItem('s7_users'));
   if(users[nu]) return alert('Username นี้มีอยู่ในระบบแล้ว');
   users[nu] = np;
   localStorage.setItem('s7_users', JSON.stringify(users));
   document.getElementById('new-user').value='';
   document.getElementById('new-pass').value='';
   renderAdminUsers();
}
function deleteUser(key) {
   if(key==='admin') return;
   if(confirm('ยืนยันการลบผู้ใช้: '+key+' ?')) {
       let users = JSON.parse(localStorage.getItem('s7_users'));
       delete users[key];
       localStorage.setItem('s7_users', JSON.stringify(users));
       renderAdminUsers();
   }
}

document.addEventListener('DOMContentLoaded', checkLogin);
document.getElementById('loginPass').addEventListener('keypress', function(e) { if(e.key === 'Enter') handleLogin(); });

async function handleFiles(files) {
  if(files.length===0) return;
  document.querySelector('.upload-btn').style.display='none';
  document.getElementById('loading-indicator').style.display='block';
  for(const file of files) { await processFile(file); }
  
  // Set default selection to latest year and month
  if (state.availableYears.size > 0) state.selectedYear = Math.max(...Array.from(state.availableYears)).toString();
  if (state.availableMonths.size > 0) state.selectedMonth = Math.max(...Array.from(state.availableMonths)).toString();
  
  populateFilters();
  document.getElementById('upload-view').style.display='none';
  document.getElementById('dashboard-view').style.display='block';
  if(typeof lucide!=='undefined') lucide.createIcons();
  renderDashboard();
}

function processFile(file) {
  return new Promise((resolve)=>{
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, {type:'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, {defval:''});
      parseData(json, file.name);
      resolve();
    };
    reader.readAsArrayBuffer(file);
  });
}

function parseData(data, filename) {
  if(!data.length) return;
  const headers = Object.keys(data[0]);
  const getCol = (kws) => { let c=headers.find(h=>kws.some(k=>h.toLowerCase()===k.toLowerCase())); if(c) return c; return headers.find(h=>kws.some(k=>h.toLowerCase().includes(k.toLowerCase()))); };
  
  const dateKey = getCol(['doc date','completed date (day)','date','วันที่']);
  const staffNameKey = getCol(['officer (name)','staff name','พนักงาน']);
  const staffIdKey = getCol(['officer (id)','staff id','รหัส']);
  const branchKey = getCol(['branch (name)','store name','สาขา']);
  const priceKey = getCol(['total price','ราคาขายตามบิล','ยอด']);
  const qtyKey = getCol(['number','จำนวน','qty']);
  const docNoKey = getCol(['doc no','เลขที่','bill']);
  
  const catKey = getCol(['category (name)', 'category']);
  // First column logic: assume catKey is the actual first, or we pick index 0 if undefined
  const catColHeader = catKey || headers[0];
  const subCatColHeader = getCol(['sub category', 'sub-category']) || headers[1] || '';
  const mwadColHeader = getCol(['หมวดสินค้า']) || headers[2] || '';
  
  const productKey = getCol(['product (name)','สินค้า']);
  const ratingKey = getCol(['กรุณาให้คะแนนความพึงพอใจ','คะแนน','rating']);
  const topicKey = getCol(['เรื่องใด','หัวข้อ','topic','นำเสนอข้อมูล']);
  const fb1Key = getCol(['ตอบกลับ discovery','discovery','ลูกค้า']);
  const fb2Key = getCol(['ตอบกลับ demo topic','demo topic','นำเสนอ']);

  const isSales = !!priceKey && !!qtyKey;
  const isRating = !!ratingKey;
  if(isSales) state.fileNames.sales = filename;
  if(isRating) state.fileNames.rating = filename;

  data.forEach(row=>{
    const dObj = normalizeDate(row[dateKey]);
    let dateStr = null;
    let year = null;
    let month = null;
    let day = null;
    
    if (dObj) {
        year = dObj.getFullYear();
        month = dObj.getMonth() + 1; // 1-12
        day = dObj.getDate();
        dateStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        state.availableYears.add(year);
        state.availableMonths.add(month);
    }
    
    let branch = String(row[branchKey]||'').trim();
    if(branch && branch!=='-'){ branch=branch.replace(/^ID\d+\s*:\s*/,''); state.availableBranches.add(branch); }
    else branch='ไม่ระบุสาขา';
    
    let cleanName = String(row[staffNameKey]||'ไม่ระบุ').trim();
    if(cleanName!=='ไม่ระบุ') cleanName=cleanName.split(/[0-9]/)[0].trim();
    
    let rawCat = String(row[catColHeader]||'').trim();
    let rawSubCat = String(row[subCatColHeader]||'').trim();
    let rawMwad = String(row[mwadColHeader]||'').trim();
    
    // IF Rating, map category using Topic
    if(isRating && !isSales) {
        let topicStr = String(row[topicKey]||'').trim();
        rawCat = topicStr; // Treat topic as main category string for classification
    }
    
    let groupCat = classifyCategory(rawCat, rawSubCat, rawMwad);
    
    state.availableCats.add(groupCat);

    const obj = {
      date: dateStr, year, month, day,
      staffName: cleanName, staffId: String(row[staffIdKey]||'-').trim(),
      branch,
      price: getVal(row[priceKey]), qty: getVal(row[qtyKey]), docNo: String(row[docNoKey]||'').trim(),
      category: groupCat, subCategory: rawCat || rawSubCat || 'Other', productName: String(row[productKey]||'').trim(),
      rating: getVal(row[ratingKey]), topic: String(row[topicKey]||'อื่นๆ').trim(),
      feedback1: String(row[fb1Key]||'').trim(), feedback2: String(row[fb2Key]||'').trim()
    };
    
    if(obj.staffName==='ไม่ระบุ' && obj.staffId!=='-') obj.staffName='รหัส '+obj.staffId;
    if(isSales && obj.qty===0 && obj.price>0) obj.qty=1;

    if(isRating && !isSales) { if(obj.staffName) state.rating.push(obj); }
    else if(isSales) { if(obj.staffName && (obj.price>0 || obj.qty>0)) state.sales.push(obj); }
    else { if(obj.rating>0) state.rating.push(obj); if(obj.price>0 || obj.qty>0) state.sales.push(obj); }
  });
}

function processCalcGrowth(currentVal, previousVal) {
    if (previousVal === 0 && currentVal > 0) return { pct: 100, class: 'g-up', arrow: '↑' };
    if (previousVal === 0 && currentVal === 0) return { pct: 0, class: 'g-neutral', arrow: '-' };
    const pct = ((currentVal - previousVal) / previousVal) * 100;
    if (pct > 0) return { pct: pct.toFixed(1), class: 'g-up', arrow: '↑' };
    if (pct < 0) return { pct: Math.abs(pct).toFixed(1), class: 'g-down', arrow: '↓' };
    return { pct: 0, class: 'g-neutral', arrow: '-' };
}

const monthNames = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];

function populateFilters() {
  const branchFilter=document.getElementById('branchFilter');
  const monthFilter=document.getElementById('monthFilter');
  const yearFilter=document.getElementById('yearFilter');
  const catFilter=document.getElementById('catFilter');

  branchFilter.innerHTML='<option value="all">รวมทุกสาขา (All Branches)</option>';
  Array.from(state.availableBranches).sort().forEach(b=>{const o=document.createElement('option');o.value=b;o.textContent=b;branchFilter.appendChild(o);});
  
  monthFilter.innerHTML='<option value="all">รวมทุกเดือน</option>';
  Array.from(state.availableMonths).sort((a,b)=>a-b).forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=monthNames[m];monthFilter.appendChild(o);});
  if (state.selectedMonth !== 'all') monthFilter.value = state.selectedMonth;

  yearFilter.innerHTML='<option value="all">รวมทุกปี</option>';
  Array.from(state.availableYears).sort((a,b)=>b-a).forEach(y=>{const o=document.createElement('option');o.value=y;o.textContent=y;yearFilter.appendChild(o);});
  if (state.selectedYear !== 'all') yearFilter.value = state.selectedYear;

  catFilter.innerHTML='<option value="all">รวมทุกหมวดหมู่ (All Categories)</option>';
  Array.from(state.availableCats).sort().forEach(c=>{const o=document.createElement('option');o.value=c;o.textContent=c;catFilter.appendChild(o);});
}

document.getElementById('branchFilter').addEventListener('change',(e)=>{state.selectedBranch=e.target.value;renderDashboard();});
document.getElementById('monthFilter').addEventListener('change',(e)=>{state.selectedMonth=e.target.value;renderDashboard();});
document.getElementById('yearFilter').addEventListener('change',(e)=>{state.selectedYear=e.target.value;renderDashboard();});
document.getElementById('catFilter').addEventListener('change',(e)=>{state.selectedCat=e.target.value;renderDashboard();});

function renderDashboard() {
  const tMonth = state.selectedMonth === 'all' ? null : parseInt(state.selectedMonth);
  const tYear = state.selectedYear === 'all' ? null : parseInt(state.selectedYear);
  const pMonth = tMonth ? (tMonth === 1 ? 12 : tMonth - 1) : null;
  const pYearForMoM = (tMonth === 1 && tYear) ? tYear - 1 : tYear;
  const pYearForYoY = tYear ? tYear - 1 : null;

  // Filter Functions
  const fFilter = (arr, ty, tm) => arr.filter(x => 
    (state.selectedBranch==='all' || x.branch===state.selectedBranch) &&
    (state.selectedCat==='all'|| x.category===state.selectedCat) &&
    (ty === null || x.year === ty) &&
    (tm === null || x.month === tm)
  );

  const curSales = fFilter(state.sales, tYear, tMonth);
  const momSales = pMonth!==null ? fFilter(state.sales, pYearForMoM, pMonth) : [];
  const yoySales = pYearForYoY!==null ? fFilter(state.sales, pYearForYoY, tMonth) : [];

  const curRating = fFilter(state.rating, tYear, tMonth);

  // Headers Update
  document.getElementById('brand-title').textContent=state.selectedBranch==='all'?'Studio 7 (รวมทุกสาขา)':state.selectedBranch;
  document.getElementById('sales-filename').textContent='ข้อมูลจากไฟล์: '+(state.fileNames.sales||'ไม่ได้อัปโหลด');
  document.getElementById('rating-filename').textContent='ข้อมูลจากไฟล์: '+(state.fileNames.rating||'ไม่ได้อัปโหลด');

  // Aggregation - Current
  let rRev=0, rQty=0, billsSet=new Set(), staffMap={}, prodMap={}, catAcMap={}, curDaily={};
  curSales.forEach(s=>{
    rRev+=s.price; rQty+=s.qty; if(s.docNo) billsSet.add(s.docNo);
    // Group By Category -> SubCat
    if(!catAcMap[s.category]) catAcMap[s.category] = { rev:0, subs:{} };
    catAcMap[s.category].rev += s.price;
    if(!catAcMap[s.category].subs[s.subCategory]) catAcMap[s.category].subs[s.subCategory] = 0;
    catAcMap[s.category].subs[s.subCategory] += s.price;

    // Staff
    const skey = s.staffName;
    if(!staffMap[skey]) staffMap[skey] = { rev:0, branch: s.branch };
    staffMap[skey].rev += s.price;
    
    // Prods
    if(s.productName) prodMap[s.productName] = (prodMap[s.productName]||0) + s.qty;

    // Daily
    if(s.day) curDaily[s.day] = (curDaily[s.day]||0) + s.price;
  });
  const rBills = billsSet.size || curSales.length;

  // Aggregation - MoM / YoY
  let momRev=0, momQty=0, momBillsSet=new Set(), momDaily={};
  momSales.forEach(s=> { momRev+=s.price; momQty+=s.qty; if(s.docNo) momBillsSet.add(s.docNo); if(s.day) momDaily[s.day] = (momDaily[s.day]||0) + s.price; });
  let momBills = momBillsSet.size || momSales.length;

  let yoyRev=0, yoyQty=0, yoyBillsSet=new Set(), yoyDaily={};
  yoySales.forEach(s=> { yoyRev+=s.price; yoyQty+=s.qty; if(s.docNo) yoyBillsSet.add(s.docNo); if(s.day) yoyDaily[s.day] = (yoyDaily[s.day]||0) + s.price; });
  let yoyBills = yoyBillsSet.size || yoySales.length;

  // Render Core KPIs
  document.getElementById('kpi-rev').textContent=rRev>=1000000?'฿'+(rRev/1000000).toFixed(1)+'M':'฿'+(rRev/1000).toFixed(0)+'K';
  document.getElementById('kpi-rev-sub').textContent=rRev.toLocaleString()+' บาท';
  
  document.getElementById('kpi-qty').textContent=rQty.toLocaleString();
  document.getElementById('kpi-bill').textContent=rBills.toLocaleString();

  // Growth Labels Generator
  const gYoyRev = processCalcGrowth(rRev, yoyRev);
  const gMomRev = processCalcGrowth(rRev, momRev);
  const elRyoy = document.getElementById('kpi-rev-yoy'); elRyoy.className = 'growth-label ' + gYoyRev.class; elRyoy.textContent = `YoY: ${gYoyRev.arrow}${gYoyRev.pct}%`;
  const elRmom = document.getElementById('kpi-rev-mom'); elRmom.className = 'growth-label ' + gMomRev.class; elRmom.textContent = `MoM: ${gMomRev.arrow}${gMomRev.pct}%`;

  const gYoyQty = processCalcGrowth(rQty, yoyQty);
  const gMomQty = processCalcGrowth(rQty, momQty);
  const elQyoy = document.getElementById('kpi-qty-yoy'); elQyoy.className = 'growth-label ' + gYoyQty.class; elQyoy.textContent = `YoY: ${gYoyQty.arrow}${gYoyQty.pct}%`;
  const elQmom = document.getElementById('kpi-qty-mom'); elQmom.className = 'growth-label ' + gMomQty.class; elQmom.textContent = `MoM: ${gMomQty.arrow}${gMomQty.pct}%`;

  const gYoyBill = processCalcGrowth(rBills, yoyBills);
  const gMomBill = processCalcGrowth(rBills, momBills);
  const elByoy = document.getElementById('kpi-bill-yoy'); elByoy.className = 'growth-label ' + gYoyBill.class; elByoy.textContent = `YoY: ${gYoyBill.arrow}${gYoyBill.pct}%`;
  const elBmom = document.getElementById('kpi-bill-mom'); elBmom.className = 'growth-label ' + gMomBill.class; elBmom.textContent = `MoM: ${gMomBill.arrow}${gMomBill.pct}%`;

  // Hidden if 'all' selected
  if (tYear==='all' || tMonth==='all') {
    document.querySelectorAll('.growth-label').forEach(e=>e.style.display='none');
  } else {
    document.querySelectorAll('.growth-label').forEach(e=>e.style.display='inline-block');
  }

  // Top Performance
  const staffArr = Object.entries(staffMap).sort((a,b)=>b[1].rev - a[1].rev);
  const topStaff = staffArr.length>0 ? staffArr[0][0] : '-';
  const topBranchList = {};
  curSales.forEach(s=> { topBranchList[s.branch] = (topBranchList[s.branch]||0) + s.price; });
  const bArr = Object.entries(topBranchList).sort((a,b)=>b[1]-a[1]);
  const topBranch = bArr.length>0 ? bArr[0][0] : '-';

  document.getElementById('kpi-topsale').textContent = topStaff.split(' ')[0];
  document.getElementById('kpi-topbranch').textContent = topBranch;
  
  const prodArr = Object.entries(prodMap).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const prodUl = document.getElementById('kpi-topprods');
  prodUl.innerHTML = '';
  if(prodArr.length===0) prodUl.innerHTML='<li>-ไม่มีสินค้า-</li>';
  else prodArr.forEach(p=> { prodUl.innerHTML+=`<li><strong>${p[1]}</strong>: ${p[0]}</li>`; });

  // Accordion (Drill-down)
  const catArr = Object.entries(catAcMap).sort((a,b)=>b[1].rev - a[1].rev);
  const catDiv = document.getElementById('catListAccordion');
  catDiv.innerHTML = '';
  catArr.forEach((cItem, i) => {
     const cName = cItem[0]; const cObj = cItem[1];
     window.toggleAccordion = function(el) {
         el.parentElement.classList.toggle('open');
     };
     
     let subHtml = '';
     const sortedSubs = Object.entries(cObj.subs).sort((a,b)=>b[1]-a[1]);
     sortedSubs.forEach(sItem => {
         subHtml += `<div class="subcat-row"><span class="subcat-name" title="${sItem[0]}">${sItem[0]}</span><span class="subcat-val">฿${sItem[1]>=1000000?(sItem[1]/1000000).toFixed(2)+'M':(sItem[1]/1000).toFixed(1)+'K'}</span></div>`;
     });

     const cColor = colorPalette[i % colorPalette.length];
     const rStr = cObj.rev >= 1000000 ? `฿${(cObj.rev/1000000).toFixed(2)}M` : `฿${(cObj.rev/1000).toFixed(1)}K`;
     
     catDiv.innerHTML += `
        <div class="cat-item">
          <div class="cat-header" onclick="window.toggleAccordion(this)">
             <div class="cat-header-left">
                <div class="cat-icon" style="background:${cColor}">${cName[0]}</div>
                <div class="cat-title">${cName}</div>
             </div>
             <div class="cat-rev">${rStr}</div>
          </div>
          <div class="cat-body">${subHtml}</div>
        </div>
     `;
  });

  // REST OF UNCHANGED VIEWS
  
  // Rating logic
  let totalRating=0, ratingCount=0, topicMap={}, rDaily={};
  let times = [];
  curRating.forEach(r=>{
      if(r.rating>0){totalRating+=r.rating; ratingCount++;}
      if(r.topic) topicMap[r.topic] = (topicMap[r.topic]||0)+1;
      if(r.day) rDaily[r.day] = (rDaily[r.day]||0)+1;
      if(r.date) times.push(new Date(r.date).getTime());
  });
  const tDemos = curRating.length;
  
  let wLabel = "ครั้ง";
  if(times.length > 0) {
      let minT = Math.min(...times);
      let maxT = Math.max(...times);
      let daysPassed = ((maxT - minT) / (1000 * 3600 * 24)) + 1; // inclusive days
      if(daysPassed < 1) daysPassed = 1;
      
      let dailyAvg = tDemos / daysPassed;
      let daysInPeriod = 30; // default for 'all'
      
      const st_month = state.selectedMonth;
      const st_year = state.selectedYear;
      if(st_month !== 'all' && st_year !== 'all') {
         daysInPeriod = new Date(parseInt(st_year), parseInt(st_month), 0).getDate();
      } else if(st_month !== 'all') {
         daysInPeriod = new Date(2024, parseInt(st_month), 0).getDate(); // leap year generic fallback
      }
      
      let forecastWeekly = Math.round(dailyAvg * 7);
      wLabel = `คาดการณ์ (รายสัปดาห์): ~${forecastWeekly.toLocaleString()} ครั้ง`;
  }
  
  document.getElementById('kpi-rating').textContent = ratingCount>0 ? (totalRating/ratingCount).toFixed(2) : '0.00';
  document.getElementById('kpi-rating-sub').textContent = `จาก ${ratingCount} / ${tDemos}`;
  document.getElementById('kpi-demo').textContent = tDemos.toLocaleString();
  let kpiDemoSub = document.getElementById('kpi-demo-sub');
  if(kpiDemoSub) kpiDemoSub.textContent = wLabel;
  
  const topicsSorted=Object.entries(topicMap).sort((a,b)=>b[1]-a[1]);
  if(topicsSorted.length>0){document.getElementById('kpi-topic').textContent=topicsSorted[0][0];}
  else {document.getElementById('kpi-topic').textContent='-';}

  // Chart Rendering
  if(charts.sales) charts.sales.destroy();
  if(charts.demo) charts.demo.destroy();

  const labels = Array.from(new Set([...Object.keys(curDaily), ...Object.keys(rDaily)])).sort((a,b)=>a-b);
  const salesData = labels.map(l => curDaily[l] || 0);
  const demoData = labels.map(l => rDaily[l] || 0);

  const ctxSales = document.getElementById('dailyChart');
  if(ctxSales) {
      charts.sales = new Chart(ctxSales, {
          type: 'line',
          data: {
              labels: labels.map(l=> l+''),
              datasets: [{
                  label: 'ยอดขาย (บาท)',
                  data: salesData,
                  borderColor: '#10b981',
                  backgroundColor: 'rgba(16, 185, 129, 0.1)',
                  fill: true, tension: 0.4
              }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
      });
  }

  const ctxDemo = document.getElementById('demoDailyChart');
  if(ctxDemo) {
      charts.demo = new Chart(ctxDemo, {
          type: 'bar',
          data: {
              labels: labels.map(l=> l+''),
              datasets: [{
                  label: 'จำนวน Demo (ครั้ง)',
                  data: demoData,
                  backgroundColor: '#f59e0b',
                  borderRadius: 4
              }]
          },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
      });
  }

  // Staff Table calculation
  const cMap={};
  curSales.forEach(s=>{
    const k=s.staffId!=='-'?s.staffId:s.staffName;
    if(!cMap[k]) cMap[k]={name:s.staffName,id:s.staffId,rev:0,items:0,bills:new Set(),demos:0,rSum:0,rCount:0};
    cMap[k].rev+=s.price; cMap[k].items+=s.qty; if(s.docNo) cMap[k].bills.add(s.docNo);
  });
  curRating.forEach(r=>{
    const k=r.staffId!=='-'?r.staffId:r.staffName;
    if(!cMap[k]) cMap[k]={name:r.staffName,id:r.staffId,rev:0,items:0,bills:new Set(),demos:0,rSum:0,rCount:0};
    cMap[k].demos+=1; if(r.rating>0){cMap[k].rSum+=r.rating;cMap[k].rCount++;}
  });

  const cSorted=Object.values(cMap).sort((a,b)=>b.rev-a.rev);
  const cb = document.getElementById('combinedBody'); cb.innerHTML='';
  const dStaff = [...cSorted].sort((a,b)=>b.demos-a.demos);
  if(dStaff.length>0 && dStaff[0].demos>0) { document.getElementById('kpi-topdemo').textContent=dStaff[0].name.split(' ')[0]; }
  else { document.getElementById('kpi-topdemo').textContent='-'; }
  document.getElementById('meta-staff-count').textContent=cSorted.length+' คน';

  const dg = document.getElementById('demoGrid');
  if(dg) {
      dg.innerHTML = '';
      if(dStaff.length === 0 || dStaff[0].demos === 0) {
          dg.innerHTML = '<div style="color:var(--dim);font-size:12px;padding:20px;grid-column:1/-1;text-align:center;">ไม่มีข้อมูล Demo</div>';
      } else {
          dStaff.filter(s=>s.demos>0).slice(0, 10).forEach(s => {
              dg.innerHTML += `
                <div class="demo-card" style="cursor:pointer" onclick="window.openModal('${s.id}','${s.name}')">
                  <div class="demo-avatar">${s.name.charAt(0)}</div>
                  <div style="min-width:0;">
                     <div class="demo-count">${s.demos} <span class="demo-label">ครั้ง</span></div>
                     <div class="demo-name" title="${s.name}">${s.name}</div>
                  </div>
                </div>
              `;
          });
      }
  }

  cSorted.forEach((s,i)=>{
    let badge='<span style="background:var(--bg);color:var(--dim);font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px">ทั่วไป</span>';
    if(s.rev>4000000&&s.demos>=12)badge='<span style="background:#fef3c7;color:#92400e;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px">⭐ TOP STAR</span>';
    
    let closeRate = '-';
    if(s.demos > 0) {
        closeRate = ((s.items / s.demos) * 100).toFixed(1) + '%';
    } else if (s.items > 0) {
        closeRate = '<span style="color:#10b981">No Demo</span>';
    }
    cb.innerHTML+=`<tr class="row-clickable" onclick="window.openModal('${s.id}','${s.name}')"><td>${i+1}</td><td><span class="staff-name">${s.name}</span><span class="staff-id">${s.id}</span></td><td class="text-r">฿${s.rev.toLocaleString()}</td><td class="text-r">${s.items.toLocaleString()}</td><td class="text-r">${s.bills.size||Math.ceil(s.items/2)}</td><td class="text-r">${s.demos}</td><td class="text-r">${closeRate}</td><td class="text-r">${s.rCount>0?(s.rSum/s.rCount).toFixed(2):'-'}</td><td class="text-r">${badge}</td></tr>`;

  });

  // Modal Function logic inside window.openModal so it's global
  window.openModal = function(sid, sname) {
    const mm = document.getElementById('staffModal');
    const key = sid!=='-'?sid:sname;
    document.getElementById('modal-name').textContent=sname;
    document.getElementById('modal-id').textContent='ID: '+sid;
    const sd = curSales.filter(s=>(s.staffId!=='-'?s.staffId:s.staffName)===key);
    const rd = curRating.filter(r=>(r.staffId!=='-'?r.staffId:r.staffName)===key);
    let mr=0, mi=0, ms={}, mrSum=0, mrCt=0, mf=[];
    sd.forEach(s=>{mr+=s.price; mi+=s.qty; if(s.productName) ms[s.productName]=(ms[s.productName]||0)+s.qty;});
    rd.forEach(r=>{if(r.rating>0){mrSum+=r.rating;mrCt++;} if(r.feedback1||r.feedback2)mf.push(r);});
    document.getElementById('mkpi-rev').textContent='฿'+mr.toLocaleString();
    document.getElementById('mkpi-items').textContent=mi.toLocaleString();
    document.getElementById('mkpi-demo').textContent=rd.length;
    document.getElementById('mkpi-rating').textContent=mrCt>0?(mrSum/mrCt).toFixed(2):'-';
    
    const lp = document.getElementById('modal-products'); lp.innerHTML='';
    Object.entries(ms).sort((a,b)=>b[1]-a[1]).slice(0,20).forEach(p=> lp.innerHTML+=`<li>${p[0]} - <strong>${p[1]} ชิ้น</strong></li>`);
    const lf = document.getElementById('modal-feedbacks'); lf.innerHTML='';
    mf.forEach(f=> { lf.innerHTML+=`<li><strong>หัวข้อ: ${f.topic} (★${f.rating})</strong><div class="feedback-text">${f.feedback1} <br>${f.feedback2}</div></li>`; });
    let transcripts = []; let objections = [];
    rd.forEach(r => {
        let text1 = String(r.feedback1||'').trim();
        let text2 = String(r.feedback2||'').trim();
        let combined = text1 + (text2 ? " | " + text2 : "");
        if (combined) transcripts.push(combined);
        if (combined.includes('แพง') || combined.includes('ไม่มีของ') || combined.includes('ลดได้ไหม') || combined.includes('ยังไม่แน่ใจ')) objections.push(combined);
    });
    
    document.getElementById('ai-json-container').style.display = 'none';
    window.currentStaffForAI = {
        name: sname, employee_id: sid,
        branch_id: ((sd[0]&&sd[0].branch) || (rd[0]&&rd[0].branch) || 'Unknown'),
        total_demos: rd.length, items_sold: mi, avg_rating: (mrCt>0?mrSum/mrCt:0),
        transcripts: transcripts.slice(0, 5), objections: objections.slice(0, 3) 
    };

    mm.classList.add('active');
  };
  window.closeModal = function() { document.getElementById('staffModal').classList.remove('active'); };
}

  window.generateStoreAIAnalysis = async function() {
      if(!window.state || !window.state.demo || window.state.demo.length === 0) { alert("ยังไม่มีข้อมูล Rating Demo ให้อ่านครับ"); return; }
      const apiKey = localStorage.getItem('s7_gemini_api');
      if(!apiKey) { alert("คุณยังไม่ได้ตั้งค่า API Key! กรุณาไปที่ปุ่ม 'จัดการ User' เพื่อตั้งค่า Google Gemini API Key ก่อนครับ"); return; }

      let filtered = window.state.demo;
      if (window.state.branch && window.state.branch !== 'ALL') {
          filtered = filtered.filter(d => d.branch_id === window.state.branch);
      }
      if(filtered.length === 0) { alert("ไม่มีข้อมูล Demo สำหรับสาขานี้"); return; }

      let combinedTranscripts = filtered.slice(0, 50).map(d => `EmpID: ${d.employee_id} (${d.branch_id})\nRating: ${d.rating}/5\nTranscript: ${d.demo_transcript || '-'}\nCustomer Objection: ${d.customer_objection || '-'}`).join('\n---\n');

      const systemPrompt = `You are a Senior Retail Area Coach for Studio 7 (Apple Premium Reseller).
Your task is to analyze the aggregated daily performance of a specific branch based on the combined Demo interactions of its staff.
Identify systematic issues, common strengths, and group the employees by what they need to develop.

[Strict Output Format]
You MUST output the analysis strictly in JSON format. Do NOT include Markdown codeblocks.
CRITICAL: All textual values MUST BE IN THAI LANGUAGE.
{
  "store_overview": { "overall_score": 85, "total_demos_analyzed": 15 },
  "store_strengths": ["...", "..."],
  "systematic_weaknesses": ["...", "..."],
  "development_groups": [
    { "focus_area": "...", "employee_ids": ["E101", "E103"] }
  ],
  "manager_action_plan": "..."
}`;
      const payloadContent = `Please evaluate these aggregated store demo interactions:\n\n${combinedTranscripts}`;

      document.getElementById('store-ai-json-output').innerHTML = '';
      document.getElementById('store-ai-header-title').textContent = "กำลังวิเคราะห์ภาพรวมสาขา... (Analyzing)";
      document.getElementById('storeAiModal').classList.add('active');

      try {
          const preferred = ["models/gemini-2.5-flash", "models/gemini-1.5-flash", "models/gemini-1.5-flash-8b", "models/gemini-1.5-pro", "models/gemini-pro"];
          let res = null; let targetModel = ""; let lastErr = "";
          for(let mName of preferred) {
              res = await fetch(`https://generativelanguage.googleapis.com/v1beta/${mName}:generateContent?key=${apiKey}`, {
                  method: "POST", headers: { "Content-Type": "application/json" },
                  body: JSON.stringify({ contents: [{ parts: [{ text: systemPrompt + "\n\n---\n" + payloadContent }] }] })
              });
              targetModel = mName; if(res.ok) break; lastErr = await res.text();
          }
          if (!res || !res.ok) throw new Error(`HTTP ${res ? res.status : 'Unknown'} on ${targetModel}. Last Error: ${lastErr}`);
          
          const ans = await res.json();
          let rawString = ans.candidates[0].content.parts[0].text.replace(/^```(json)?/, '').replace(/```$/, '').trim();
          
          let inTk = 0, outTk = 0;
          if(ans.usageMetadata) { inTk = ans.usageMetadata.promptTokenCount || 0; outTk = ans.usageMetadata.candidatesTokenCount || 0; }
          const lu = sessionStorage.getItem('s7_logged_user');
          if(lu) {
              let logs = JSON.parse(localStorage.getItem('s7_usage_logs')) || [];
              logs.push({ type: 'ai_call', user: lu, timestamp: Date.now(), inputTokens: inTk, outputTokens: outTk, costEst: ((inTk/1000000)*0.35 + (outTk/1000000)*1.05) });
              localStorage.setItem('s7_usage_logs', JSON.stringify(logs));
          }
          
          const jObj = JSON.parse(rawString);
          document.getElementById('store-ai-header-title').textContent = "✅ ผลการวิเคราะห์เสร็จสมบูรณ์";
          
          let groupsHtml = (jObj.development_groups || []).map(g => `<div style="background:#0f172a; padding:10px; border-radius:6px; margin-bottom:8px; border-left:3px solid #38bdf8;"><div style="color:#e2e8f0; font-weight:700;">${g.focus_area}</div><div style="color:#94a3b8; font-size:11px; margin-top:4px;">พนักงานเป้าหมาย: ${(g.employee_ids||[]).join(', ')}</div></div>`).join('');
          
          let html = `<div style="font-family:'Sarabun',sans-serif; color:#f8fafc;">
               <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; padding-bottom:12px; border-bottom:1px solid #334155;">
                  <div>
                    <div style="font-size:16px; font-weight:700; color:#38bdf8;">สรุปภาพรวมสาขา (Branch Overall Analysis)</div>
                    <div style="font-size:12px; color:#94a3b8;">สาขา: ${window.state.branch || 'ทุกสาขารวมกัน'} (${jObj.store_overview?.total_demos_analyzed || 0} Demos)</div>
                  </div>
                  <div style="text-align:right;">
                    <div style="font-size:24px; font-weight:700; color:${(jObj.store_overview?.overall_score || 0) >= 80 ? '#10b981' : '#f59e0b'}; line-height:1;">${jObj.store_overview?.overall_score || 0}</div>
                    <div style="font-size:10px; color:#cbd5e1;">คะแนนเฉลี่ยรวม</div>
                  </div>
               </div>
               <div style="margin-bottom:16px;">
                  <div style="font-size:13px; font-weight:700; color:#10b981; margin-bottom:6px;">🌟 จุดเด่นร่วมกัน (Store Strengths)</div>
                  <ul style="margin:0; padding-left:20px; font-size:12px; color:#cbd5e1; line-height:1.6;">
                     ${(jObj.store_strengths||[]).map(s => `<li>${s}</li>`).join('')}
                  </ul>
               </div>
               <div style="margin-bottom:16px;">
                  <div style="font-size:13px; font-weight:700; color:#f43f5e; margin-bottom:6px;">⚠️ ปัญหาเชิงระบบ (Systematic Weaknesses)</div>
                  <ul style="margin:0; padding-left:20px; font-size:12px; color:#cbd5e1; line-height:1.6;">
                     ${(jObj.systematic_weaknesses||[]).map(w => `<li>${w}</li>`).join('')}
                  </ul>
               </div>
               <div style="margin-bottom:16px;">
                  <div style="font-size:13px; font-weight:700; color:#e0e7ff; margin-bottom:6px;">👥 จัดกลุ่มพนักงานที่ต้องพัฒนา (Development Target)</div>
                  ${groupsHtml}
               </div>
               <div style="background:#1e293b; border-radius:8px; padding:12px; border:1px solid #334155;">
                  <div style="font-size:13px; font-weight:700; color:#fde047; margin-bottom:8px;">🎯 Manager Action Plan</div>
                  <div style="font-size:12px; color:#e2e8f0; line-height:1.5;">${jObj.manager_action_plan || ''}</div>
               </div>
               <div style="margin-top:16px; font-size:10px; color:#475569; text-align:right;">Generated by ${targetModel} (${inTk} In / ${outTk} Out)</div>
            </div>`;
          document.getElementById('store-ai-json-output').innerHTML = html;
      } catch(err) {
          document.getElementById('store-ai-header-title').textContent = "❌ AI Coach output (Error)";
          document.getElementById('store-ai-json-output').innerText = "เกิดข้อผิดพลาดในการเรียก API:\n" + err.message;
      }
  }  window.generateAIAnalysis = async function() {
      const data = window.currentStaffForAI;
      if(!data) return;
      const apiKey = localStorage.getItem('s7_gemini_api');
      if(!apiKey) { alert("คุณยังไม่ได้ตั้งค่า API Key! กรุณาไปที่ปุ่ม 'จัดการ User' เพื่อตั้งค่า Google Gemini API Key ก่อนครับ"); return; }
      
      document.getElementById('ai-json-output').textContent = "กำลังเชื่อมต่อไปยัง Google Gemini API... รอสักครู่ ⏳";
      document.getElementById('ai-json-container').style.display = 'block';
      
      let t_script = data.transcripts.length > 0 ? data.transcripts.join(' || ') : "No conversation available";
      let c_obj = data.objections.length > 0 ? data.objections.join(' || ') : "None recorded";
      let conv_status = `Sold ${data.items_sold} items from ${data.total_demos} demos`;
      
      const systemPrompt = `You are an advanced "AI Demo Performance Analyst". Your primary objective is to evaluate employee product demonstrations (Demos).
      
[Analysis Guidelines]
1. Individual Analysis: Evaluate opening connection, key features presentation, and closing. Evaluate Objection Handling.
2. Tone & Sentiment: Analyze tone from transcribed words.
3. Actionable Coaching: 1-2 immediate tips to apply.

[Strict Output Format]
You MUST output the analysis strictly in JSON format. Do NOT include any markdown codeblocks or 'json' prefixes outside the JSON block. 
CRITICAL: All textual values and descriptions inside the JSON MUST BE IN THAI LANGUAGE.
{
  "analysis_overview": { "employee_id": "...", "branch_id": "...", "overall_score": "0-100" },
  "performance_metrics": { "strengths": ["...", "..."], "weaknesses": ["...", "..."] },
  "demo_insights": { "objection_handling_quality": "Excellent/Good/Needs Improvement", "time_efficiency": "Unknown/Optimal" },
  "ai_coach_recommendation": { "immediate_action": "...", "training_focus": "..." }
}`;

      const payloadContent = `Please evaluate this employee data:
- employee_id: ${data.employee_id} (${data.name})
- branch_id: ${data.branch_id}
- demo_transcript: ${t_script}
- customer_objection: ${c_obj}
- conversion_status: ${conv_status}
- average_demo_rating: ${data.avg_rating}/5`;

      try {
          const apiKey = localStorage.getItem('s7_gemini_api');
          if(!apiKey) throw new Error("API Key not found");
          
          let targetModel = "models/gemini-2.5-flash";
          const preferred = [
              "models/gemini-2.5-flash",
              "models/gemini-1.5-flash",
              "models/gemini-1.5-flash-8b",
              "models/gemini-1.5-pro",
              "models/gemini-pro"
          ];

          let res = null;
          let lastErr = "";
          for(let mName of preferred) {
              res = await fetch(`https://generativelanguage.googleapis.com/v1beta/${mName}:generateContent?key=${apiKey}`, {
                  method: "POST",
                  headers: { "Content-Type": "application/json" },
                  body: JSON.stringify({
                      contents: [{ parts: [{ text: systemPrompt + "\n\n---\n" + payloadContent }] }]
                  })
              });
              targetModel = mName;
              if(res.ok) break; 
              lastErr = await res.text();
              console.warn(`Model ${mName} failed: HTTP ${res.status}. Retrying next available model...`);
          }
          
          if (!res || !res.ok) {
              throw new Error(`HTTP ${res ? res.status : 'Unknown'} on ${targetModel}. Last Error: ${lastErr}`);
          }
          
          const ans = await res.json();
          if(!ans.candidates || !ans.candidates[0].content) {
              throw new Error("Invalid API Response: " + JSON.stringify(ans));
          }
          
          let inTk = 0, outTk = 0;
          if(ans.usageMetadata) {
             inTk = ans.usageMetadata.promptTokenCount || 0;
             outTk = ans.usageMetadata.candidatesTokenCount || 0;
          }
          const lu = sessionStorage.getItem('s7_logged_user');
          if(lu) {
              let logs = JSON.parse(localStorage.getItem('s7_usage_logs')) || [];
              logs.push({
                  type: 'ai_call',
                  user: lu,
                  timestamp: Date.now(),
                  inputTokens: inTk,
                  outputTokens: outTk,
                  costEst: ((inTk/1000000)*0.35 + (outTk/1000000)*1.05)
              });
              localStorage.setItem('s7_usage_logs', JSON.stringify(logs));
          }

          let rawString = ans.candidates[0].content.parts[0].text;
          rawString = rawString.replace(/^```(json)?/, '').replace(/```$/, '').trim();
          
          try {
              const jObj = JSON.parse(rawString);
              document.getElementById('ai-header-title').textContent = "🤖 AI Coach Analysis Report";
              
              let html = `
                <div style="font-family:'Sarabun',sans-serif; color:#f8fafc;">
                   <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; padding-bottom:12px; border-bottom:1px solid #334155;">
                      <div>
                        <div style="font-size:16px; font-weight:700; color:#38bdf8;">ผลการวิเคราะห์ Demo (AI Analysis)</div>
                        <div style="font-size:12px; color:#94a3b8;">สาขา: ${jObj.analysis_overview?.branch_id || '-'}</div>
                      </div>
                      <div style="text-align:right;">
                        <div style="font-size:24px; font-weight:700; color:${(jObj.analysis_overview?.overall_score || 0) >= 80 ? '#10b981' : '#f59e0b'}; line-height:1;">${jObj.analysis_overview?.overall_score || 0}</div>
                        <div style="font-size:10px; color:#94a3b8;">คะแนนรวม (Overall Score)</div>
                      </div>
                   </div>
                   
                   <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:16px;">
                      <div style="background:#0f172a; padding:12px; border-radius:8px; border-left:3px solid #10b981;">
                         <div style="font-weight:600; font-size:13px; color:#10b981; margin-bottom:8px;">✅ จุดแข็ง (Strengths)</div>
                         <ul style="margin:0; padding-left:20px; font-size:12px; color:#cbd5e1; line-height:1.6;">
                            ${(jObj.performance_metrics?.strengths || []).map(s=>`<li style="margin-bottom:4px;">${s}</li>`).join('')}
                         </ul>
                      </div>
                      <div style="background:#0f172a; padding:12px; border-radius:8px; border-left:3px solid #f43f5e;">
                         <div style="font-weight:600; font-size:13px; color:#f43f5e; margin-bottom:8px;">💡 จุดที่ควรพัฒนา (Areas for Improvement)</div>
                         <ul style="margin:0; padding-left:20px; font-size:12px; color:#cbd5e1; line-height:1.6;">
                            ${(jObj.performance_metrics?.weaknesses || []).map(s=>`<li style="margin-bottom:4px;">${s}</li>`).join('')}
                         </ul>
                      </div>
                   </div>
                   
                   <div style="background:#1e1b4b; padding:16px; border-radius:8px; margin-bottom:16px; border:1px solid #4c1d95;">
                      <div style="font-weight:600; font-size:14px; color:#c084fc; margin-bottom:10px; display:flex; align-items:center; gap:6px;">🎯 คำแนะนำจาก AI (Actionable Coaching)</div>
                      <div style="font-size:13px; color:#e2e8f0; line-height:1.6; margin-bottom:12px;">
                         <strong style="color:#f8fafc">⚡ สิ่งที่ควรลงมือทำทันที:</strong> ${jObj.ai_coach_recommendation?.immediate_action || '-'}
                      </div>
                      <div style="font-size:13px; color:#e2e8f0; line-height:1.6;">
                         <strong style="color:#f8fafc">📚 หัวข้อที่ควรเน้นย้ำ/ฝึกอบรม:</strong> ${jObj.ai_coach_recommendation?.training_focus || '-'}
                      </div>
                   </div>
                   
                   <div style="display:flex; gap:16px; font-size:12px; margin-top:10px; color:#94a3b8; background:#0f172a; padding:10px 16px; border-radius:8px;">
                      <div><strong>การจัดการข้อโต้แย้ง (Objection Handling):</strong> <span style="color:#f8fafc;">${jObj.demo_insights?.objection_handling_quality || '-'}</span></div>
                      <div><strong>การบริหารเวลา (Time Efficiency):</strong> <span style="color:#f8fafc;">${jObj.demo_insights?.time_efficiency || '-'}</span></div>
                   </div>
                </div>
              `;
              document.getElementById('ai-json-output').innerHTML = html;
          } catch(e) {
              document.getElementById('ai-header-title').textContent = "AI Coach output (Raw Data)";
              document.getElementById('ai-json-output').innerHTML = `<pre style="margin:0; white-space:pre-wrap;">${rawString}</pre>`;
          }
      } catch(err) {
          console.error(err);
          document.getElementById('ai-header-title').textContent = "AI Coach output (Error)";
          document.getElementById('ai-json-output').innerHTML = "เกิดข้อผิดพลาดในการเรียก API: " + err.message;
      }
  };
  
  window.copyAIJson = function() {
      const container = document.getElementById('ai-json-output');
      const txt = container.innerText || container.textContent;
      navigator.clipboard.writeText(txt).then(()=>{ alert('Copied text to clipboard!'); });
  };

  function renderAdminPendings() {
       let pending = JSON.parse(localStorage.getItem('s7_pending_users')) || {};
       let tbody = document.getElementById('pending-tbody');
       if(!tbody) return;
       tbody.innerHTML = '';
       let count = 0;
       for(let k in pending) {
          count++;
          tbody.innerHTML += `<tr>
             <td style="padding:12px; border-bottom:1px solid var(--border2);">${k}</td>
             <td style="padding:12px; text-align:right; border-bottom:1px solid var(--border2);">
                <button onclick="approveUser('${k}')" style="background:var(--green);color:white;border:none;border-radius:4px;padding:4px 8px;cursor:pointer;font-weight:600;font-size:11px;">✅ Approve</button>
                <button onclick="rejectUser('${k}')" style="background:var(--rose);color:white;border:none;border-radius:4px;padding:4px 8px;cursor:pointer;font-weight:600;font-size:11px;margin-left:4px;">❌ Reject</button>
             </td>
          </tr>`;
       }
       if(count === 0) tbody.innerHTML = `<tr><td colspan="2" style="padding:12px; text-align:center; color:var(--dim); font-size:12px;">ไม่มีคำขอ (No pending requests)</td></tr>`;
  }

  function renderAdminResets() {
       let resets = JSON.parse(localStorage.getItem('s7_reset_requests')) || [];
       let tbody = document.getElementById('reset-tbody');
       if(!tbody) return;
       tbody.innerHTML = '';
       if(resets.length === 0) {
          tbody.innerHTML = `<tr><td colspan="2" style="padding:12px; text-align:center; color:var(--dim); font-size:12px;">ไม่มีคำขอ (No reset requests)</td></tr>`;
          return;
       }
       resets.forEach(k => {
          tbody.innerHTML += `<tr>
             <td style="padding:12px; border-bottom:1px solid var(--border2);">${k}</td>
             <td style="padding:12px; text-align:right; border-bottom:1px solid var(--border2); min-width:180px;">
                <div style="display:flex;gap:6px;">
                   <input type="password" id="reset-pass-${k}" placeholder="Set Password" style="flex:1; padding:4px 8px; border:1px solid var(--border); border-radius:4px; font-size:12px; font-family:var(--mono);">
                   <button onclick="fulfillReset('${k}')" style="background:var(--accent);color:white;border:none;border-radius:4px;padding:4px 8px;cursor:pointer;font-size:11px;">Update</button>
                </div>
             </td>
          </tr>`;
       });
  }

  function approveUser(u) {
       let users = JSON.parse(localStorage.getItem('s7_users')) || {};
       let pending = JSON.parse(localStorage.getItem('s7_pending_users')) || {};
       if(pending[u]) {
          users[u] = pending[u];
          delete pending[u];
          localStorage.setItem('s7_users', JSON.stringify(users));
          localStorage.setItem('s7_pending_users', JSON.stringify(pending));
          renderAdminUsers(); renderAdminPendings();
       }
  }

  function rejectUser(u) {
       let pending = JSON.parse(localStorage.getItem('s7_pending_users')) || {};
       if(pending[u]) {
          delete pending[u];
          localStorage.setItem('s7_pending_users', JSON.stringify(pending));
          renderAdminPendings();
       }
  }

  function fulfillReset(u) {
       const val = document.getElementById('reset-pass-'+u).value.trim();
       if(!val) return alert("กรุณากรอกรหัสผ่านใหม่");
       let users = JSON.parse(localStorage.getItem('s7_users')) || {};
       if(users[u]) {
           users[u] = val;
           localStorage.setItem('s7_users', JSON.stringify(users));
           
           let resets = JSON.parse(localStorage.getItem('s7_reset_requests')) || [];
           resets = resets.filter(x => x !== u);
           localStorage.setItem('s7_reset_requests', JSON.stringify(resets));
           
           alert("อัปเดตรหัสผ่านใหม่เรียบร้อยแล้ว");
           renderAdminUsers(); renderAdminResets();
       } else {
           alert("ไม่พบผู้ใช้งานในระบบ");
       }
  }

  function editUserPassword(u) {
       const val = document.getElementById('chg-pass-'+u).value.trim();
       if(!val) return alert("กรุณากรอกรหัสผ่านใหม่");
       let users = JSON.parse(localStorage.getItem('s7_users')) || {};
       if(users[u]) {
           users[u] = val;
           localStorage.setItem('s7_users', JSON.stringify(users));
           alert("อัปเดตรหัสผ่านเรียบร้อยแล้ว");
           document.getElementById('chg-pass-'+u).value = '';
       }
  }

function openChangePass() {
   const u = sessionStorage.getItem('s7_logged_user');
   if(u === 'admin' || u === 'studio7') {
      return alert("บัญชีพื้นฐานของระบบ (Admin/Studio7) ไม่สามารถเปลี่ยนรหัสผ่านจากเมนูนี้ได้ครับ");
   }
   document.getElementById('cp-old').value = '';
   document.getElementById('cp-new').value = '';
   document.getElementById('cp-new2').value = '';
   document.getElementById('cp-error').style.display = 'none';
   document.getElementById('changePassModal').classList.add('active');
}

function submitChangePassword() {
   const u = sessionStorage.getItem('s7_logged_user');
   const old = document.getElementById('cp-old').value.trim();
   const pwd1 = document.getElementById('cp-new').value.trim();
   const pwd2 = document.getElementById('cp-new2').value.trim();
   
   if(!old || !pwd1 || !pwd2) return showCpErr("กรุณากรอกข้อมูลให้ครบถ้วน");
   if(pwd1 !== pwd2) return showCpErr("รหัสผ่านใหม่ไม่ตรงกัน");
   
   let users = JSON.parse(localStorage.getItem('s7_users')) || {};
   if(!users[u]) return showCpErr("ข้อผิดพลาด: ไม่พบผู้ใช้นี้ในระบบ");
   
   if(users[u] !== old) {
       return showCpErr("รหัสผ่านเดิมไม่ถูกต้อง!");
   }
   
   users[u] = pwd1;
   localStorage.setItem('s7_users', JSON.stringify(users));
   
   alert("อัปเดตรหัสผ่านใหม่สำเร็จแล้ว!");
   document.getElementById('changePassModal').classList.remove('active');
}

function showCpErr(msg) {
   const el = document.getElementById('cp-error');
   el.textContent = '❌ ' + msg;
   el.style.display = 'block';
}

// Telemetry & Session Management
function calculateSessionDuration() {
    const start = sessionStorage.getItem('s7_session_start');
    const u = sessionStorage.getItem('s7_logged_user');
    if(start && u) {
        const end = Date.now();
        const durationMin = Math.round((end - parseInt(start)) / 60000); // Minutes
        
        let logs = JSON.parse(localStorage.getItem('s7_usage_logs')) || [];
        logs.push({
            type: 'session',
            user: u,
            loginTime: parseInt(start),
            logoutTime: end,
            durationMin: durationMin
        });
        localStorage.setItem('s7_usage_logs', JSON.stringify(logs));
        sessionStorage.removeItem('s7_session_start'); // Prevent double logging
    }
}

function handleLogout() {
    calculateSessionDuration();
    sessionStorage.clear();
    location.reload();
}

window.addEventListener('beforeunload', () => {
    calculateSessionDuration();
});

function renderAdminTelemetry() {
   let logs = JSON.parse(localStorage.getItem('s7_usage_logs')) || [];
   let tbody = document.getElementById('telemetry-tbody');
   if(!tbody) return;
   tbody.innerHTML = '';
   
   if(logs.length === 0) {
       tbody.innerHTML = `<tr><td colspan="4" style="padding:12px; text-align:center; color:var(--dim); font-size:12px;">ไม่มีข้อมูลการใช้งาน (No tracking data)</td></tr>`;
       return;
   }
   
   let agg = {};
   for(let log of logs) {
       if(!agg[log.user]) agg[log.user] = { min: 0, in: 0, out: 0, cost: 0 };
       if(log.type === 'session') agg[log.user].min += (log.durationMin || 0);
       if(log.type === 'ai_call') {
           agg[log.user].in += (log.inputTokens || 0);
           agg[log.user].out += (log.outputTokens || 0);
           agg[log.user].cost += parseFloat(log.costEst || 0);
       }
   }
   
   for(let u in agg) {
       let totalTokens = agg[u].in + agg[u].out;
       tbody.innerHTML += `<tr>
         <td style="padding:10px; border-bottom:1px solid var(--border2);">${u}</td>
         <td style="padding:10px; text-align:right; border-bottom:1px solid var(--border2);">${agg[u].min}m</td>
         <td style="padding:10px; text-align:right; border-bottom:1px solid var(--border2);">${totalTokens.toLocaleString()}</td>
         <td style="padding:10px; text-align:right; border-bottom:1px solid var(--border2);">$${agg[u].cost.toFixed(4)}</td>
       </tr>`;
   }
}

function clearTelemetry() {
    if(confirm("ลบข้อมูลประวัติการใช้งานและ Log ข้อมูลย้อนหลังทั้งหมด? (ข้อมูลนี้ไม่สามารถกู้คืนได้)")) {
        localStorage.removeItem('s7_usage_logs');
        renderAdminTelemetry();
    }
}

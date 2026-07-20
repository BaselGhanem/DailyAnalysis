const $=(s,root=document)=>root.querySelector(s), $$=(s,root=document)=>[...root.querySelectorAll(s)];let raw=[], filtered=[], meta={},tableExports={};
const denyNetwork=()=>{throw new Error(`تم تعطيل الاتصال الشبكي داخل أداة التحليل لحماية بيانات التقرير`)};
try{window.fetch=denyNetwork;window.XMLHttpRequest=class{constructor(){denyNetwork()}};window.WebSocket=class{constructor(){denyNetwork()}};if(navigator.sendBeacon)navigator.sendBeacon=()=>false}catch(_){ }const selectedFilters={team:[],mr:[],area:[],group:[],item:[]},availableFilters={team:[],mr:[],area:[],group:[],item:[]};
const fmt=n=>new Intl.NumberFormat(`en-US`,{maximumFractionDigits:2}).format(n||0), money=n=>`${fmt(n)} JOD`, norm=v=>String(v??``).trim(), num=v=>Number(v)||0;
const customerNameKey=value=>norm(value)
    .normalize(`NFKC`)
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g,``)
    .replace(/\s+/g,` `)
    .trim()
    .toLowerCase();
const countUniqueCustomerNames=data=>new Set(data.map(r=>customerNameKey(r.customer)).filter(Boolean)).size;
let toastTimer;
function showToast(message){const el=$(`#toast`);el.textContent=message;el.classList.add(`show`);clearTimeout(toastTimer);toastTimer=setTimeout(()=>el.classList.remove(`show`),2600)}
const pad2 = n => String(n).padStart(2, `0`);
const currentMonthRange=()=>{const now=new Date(),year=now.getFullYear(),month=now.getMonth()+1,lastDay=new Date(year,month,0).getDate();return{from:`${year}-${pad2(month)}-01`,to:`${year}-${pad2(month)}-${pad2(lastDay)}`}};
const setCurrentMonthRange=()=>{const range=currentMonthRange();$(`#from`).value=range.from;$(`#to`).value=range.to;return range};

const dateKey = value => {
    if (value === null || value === undefined || value === ``) return ``;

    // Excel serial date: prevents timezone shifting
    if (typeof value === `number` && Number.isFinite(value)) {
        const parsed = XLSX.SSF.parse_date_code(value);

        if (!parsed) return ``;

        return `${parsed.y}-${pad2(parsed.m)}-${pad2(parsed.d)}`;
    }

    // Date object: use UTC values to prevent one-day rollback
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return `${value.getUTCFullYear()}-${pad2(value.getUTCMonth() + 1)}-${pad2(value.getUTCDate())}`;
    }

    const text = norm(value).replace(/،/g, `,`).trim();

    const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);

    if (iso) {
        return `${iso[1]}-${pad2(iso[2])}-${pad2(iso[3])}`;
    }

    const dmy = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);

    if (dmy) {
        return `${dmy[3]}-${pad2(dmy[2])}-${pad2(dmy[1])}`;
    }

    return ``;
};
const val=(r,names)=>{for(const n of names){const k=Object.keys(r).find(x=>norm(x).toLowerCase()===n.toLowerCase());if(k!==undefined)return r[k]}return ``};
function parseRow(r){return{mr:norm(val(r,[`MR`,`مندوب`,`Representative`])),customerNo:norm(val(r,[`Cust No`,`Customer No`,`كود العميل`])),customer:norm(val(r,[`Cust Name`,`Customer Name`,`العميل`])),area:norm(val(r,[`Area Name`,`Area`,`المنطقة`])),date:dateKey(val(r,[`Transaction Date`,`Date`,`التاريخ`])),team:norm(val(r,[`Team`,`Supervisor`,`المشرف`])),item:norm(val(r,[`Item Name`,`Product`,`الصنف`])),group:norm(val(r,[`Product Group`,`Group`,`المجموعة`])),value:num(val(r,[`Value`,`Sales Value`,`القيمة`])),bonus:num(val(r,[`Bonus`,`Bonus Qty`,`البونص`])),units:num(val(r,[`Units`,`Qty`,`Quantity`,`الوحدات`]))}}
function unique(k,data=raw){return[...new Set(data.map(x=>x[k]).filter(Boolean))].sort((a,b)=>a.localeCompare(b,`ar`))}function options(el,vals,title){const old=el.value;el.innerHTML=`<option value="">${title}</option>`+vals.map(x=>`<option>${escapeHtml(x)}</option>`).join(``);if(vals.includes(old))el.value=old}function escapeHtml(s){return norm(s).replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]))}
function aggregate(data,key){const m=new Map;for(const r of data){const k=typeof key===`function`?key(r):r[key]||`غير محدد`;if(!m.has(k))m.set(k,{name:k,value:0,units:0,bonus:0,rows:0,customers:new Set,purchasingCustomers:new Set,dates:new Set,mrs:new Set,teams:new Set});const x=m.get(k),customerId=customerNameKey(r.customer);x.value+=r.value;x.units+=r.units;x.bonus+=r.bonus;x.rows++;if(customerId)x.customers.add(customerId);if(customerId&&r.value>0&&r.units>0)x.purchasingCustomers.add(customerId);if(r.date)x.dates.add(r.date);if(r.mr)x.mrs.add(r.mr);if(r.team)x.teams.add(r.team)}return[...m.values()].map(x=>({...x,customers:x.customers.size,purchasingCustomers:x.purchasingCustomers.size,dates:x.dates.size,mrs:x.mrs.size,teams:x.teams.size})).sort((a,b)=>b.value-a.value)}
function setProcessingStage(title,message){$(`#loadingTitle`).textContent=title;$(`#loadingMessage`).textContent=message}
function readFile(file){if(!file)return;$(`#loading`).classList.remove(`hidden`);setProcessingStage(`جاري قراءة التقرير...`,`الملف ${file.name} يبقى على جهازك، ولا يتم إرسال أي صف أو قيمة إلى الإنترنت.`);const reader=new FileReader;reader.onload=e=>{setProcessingStage(`جاري بناء التحليل...`,`يتم الآن تنظيف البيانات وحساب المؤشرات داخل المتصفح فقط.`);setTimeout(()=>{try{const wb=XLSX.read(e.target.result,{type:`array`,cellDates:false});const ws=wb.Sheets[wb.SheetNames[0]];if(!ws)throw Error(`لا توجد ورقة بيانات داخل الملف`);const rows=XLSX.utils.sheet_to_json(ws,{defval:``,raw:true});raw=rows.map(parseRow).filter(r=>r.mr||r.customer||r.item);if(!raw.length)throw Error(`لم يتم العثور على بيانات صالحة`);const dates=raw.map(r=>r.date).filter(Boolean).sort();if(!dates.length)throw Error(`لم يتم التعرف على عمود Transaction Date أو قيم التاريخ`);const latestDate=dates.at(-1),latestRows=raw.filter(r=>r.date===latestDate).length,invalidDates=raw.filter(r=>!r.date).length;meta={name:file.name,rows:raw.length,sheet:wb.SheetNames[0],firstDate:dates[0],latestDate,latestRows,invalidDates};init();setProcessingStage(`اكتمل التحليل`,`تمت معالجة ${fmt(raw.length)} حركة محليا. يتم الآن فتح لوحة القرار.`);setTimeout(()=>{$(`#loading`).classList.add(`hidden`);openDashboard()},520)}catch(err){$(`#loading`).classList.add(`hidden`);alert(`تعذر قراءة الملف: ${err.message}`);$(`#file`).value=``}},160)};reader.onerror=()=>{$(`#loading`).classList.add(`hidden`);alert(`تعذر فتح الملف المحدد`);$(`#file`).value=``};reader.readAsArrayBuffer(file)}
const multiMeta={team:{all:`جميع الفرق`,one:`Team`,many:`Teams`,select:`اختيار جميع الفرق`,clear:`إلغاء تحديد جميع الفرق`},mr:{all:`جميع المندوبين`,one:`مندوب`,many:`مندوبين`,select:`اختيار جميع المندوبين`,clear:`إلغاء تحديد جميع المندوبين`},area:{all:`جميع المناطق`,one:`منطقة`,many:`مناطق`,select:`اختيار جميع المناطق`,clear:`إلغاء تحديد جميع المناطق`},group:{all:`جميع المجموعات`,one:`مجموعة`,many:`مجموعات`,select:`اختيار جميع المجموعات`,clear:`إلغاء تحديد جميع المجموعات`},item:{all:`جميع الأصناف`,one:`صنف`,many:`أصناف`,select:`اختيار جميع الأصناف`,clear:`إلغاء تحديد جميع الأصناف`}};
function matchesFilter(r,key){return!selectedFilters[key].length||selectedFilters[key].includes(r[key])}function matchesExcept(r,except){return Object.keys(selectedFilters).every(k=>k===except||matchesFilter(r,k))}
function renderMulti(key,vals){availableFilters[key]=vals;selectedFilters[key]=selectedFilters[key].filter(x=>vals.includes(x));const selected=selectedFilters[key],allSelected=vals.length>0&&selected.length===vals.length,m=multiMeta[key],menu=$(`#${key}Menu`);menu.innerHTML=`<button class="multiAction" data-toggle-all="${key}" type="button">${allSelected?m.clear:m.select}</button>`+(vals.length?vals.map(x=>`<label class="teamOption"><input type="checkbox" value="${escapeHtml(x)}" ${selected.includes(x)?`checked`:``}><span>${escapeHtml(x)}</span></label>`).join(``):`<div class="noFilters">لا توجد خيارات متاحة</div>`);$(`#${key}Button`).textContent=!selected.length?m.all:selected.length===1?selected[0]:allSelected?`${m.all} (${vals.length})`:`${selected.length} ${m.many} محددة`;$$(`#${key}Menu input`).forEach(c=>c.onchange=()=>{selectedFilters[key]=$$(`#${key}Menu input:checked`).map(x=>x.value);apply()});const toggle=menu.querySelector(`[data-toggle-all]`);if(toggle)toggle.onclick=()=>{selectedFilters[key]=allSelected?[]:[...vals];renderMulti(key,vals);apply()}}
function init(){const dates=raw.map(x=>x.date).filter(Boolean).sort(),range=setCurrentMonthRange();Object.keys(selectedFilters).forEach(k=>selectedFilters[k]=[]);Object.keys(selectedFilters).forEach(k=>renderMulti(k,unique(k)));$(`#fileStatus`).textContent=`${meta.name} · ${fmt(meta.rows)} حركة · العرض ${range.from} ← ${range.to}`;$(`#sourceFileName`).textContent=`الملف: ${meta.name}`;$(`#sourceRows`).textContent=`${fmt(meta.rows)} حركة · Sheet: ${meta.sheet}`;$(`#sourceDateRange`).textContent=`نطاق الملف: ${meta.firstDate} ← ${meta.latestDate}`;$(`#sourceLatest`).textContent=`الشهر الحالي: ${range.from} ← ${range.to}`;$(`#exportBtn`).disabled=false}
function apply(){const f={from:$(`#from`).value,to:$(`#to`).value};filtered=raw.filter(r=>(!f.from||r.date>=f.from)&&(!f.to||r.date<=f.to)&&Object.keys(selectedFilters).every(k=>matchesFilter(r,k)));updateDependent(f);renderActiveFilters(f);render()}
function updateDependent(f){const base=raw.filter(r=>(!f.from||r.date>=f.from)&&(!f.to||r.date<=f.to));Object.keys(selectedFilters).forEach(key=>renderMulti(key,unique(key,base.filter(r=>matchesExcept(r,key)))))}
function renderActiveFilters(f){const chips=[];if(f.from||f.to)chips.push({key:`date`,type:`date`,label:`الفترة`,value:`${f.from||`البداية`} ← ${f.to||`النهاية`}`});const labels={team:`Teams`,mr:`المندوبون`,area:`المناطق`,group:`المجموعات`,item:`الأصناف`};Object.keys(selectedFilters).forEach(key=>{const a=selectedFilters[key];if(a.length)chips.push({key,type:key,label:labels[key],value:a.length<=2?a.join(`، `):`${a.slice(0,2).join(`، `)} +${a.length-2}`})});$(`#activeFilters`).innerHTML=`<span class="activeFiltersTitle">الفلاتر النشطة</span>`+(chips.length?chips.map(c=>`<span class="filterChip chip-${c.type}" title="${escapeHtml(c.value)}"><small>${c.label}</small><span>${escapeHtml(c.value)}</span><button type="button" data-remove-filter="${c.key}" aria-label="إزالة الفلتر">×</button></span>`).join(``):`<span class="noFilters">لا توجد فلاتر مخصصة</span>`);$$(`[data-remove-filter]`).forEach(b=>b.onclick=()=>{const k=b.dataset.removeFilter;if(k===`date`){$(`#from`).value=``;$(`#to`).value=``}else selectedFilters[k]=[];apply()})}

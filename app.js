/* Inventory + Batch + Committed Dashboard (v1.2.3)
   - Section banners + centered Stock Check (index/style handle visuals)
   - Stock Check cards with full-width headers (via CSS)
   - Committed (selected product) table with totals
   - Batches (selected product) table with sticky totals row
   - Category-wise Stock Value chart + Top products chart
   - Zero-stock (Medyra Onco) table
   - Committed items (by customer) table with rate/amount & totals
   - Details (first 2000 rows) defaults to Medyra Onco in dropdown
*/

/////////////////////// helpers ///////////////////////
const $ = id => document.getElementById(id);
const fmtIN = n => {
  if(n==null||isNaN(n)) return "–";
  const v=+n,a=Math.abs(v);
  if(a>=1e7) return (v/1e7).toFixed(2)+" Cr";
  if(a>=1e5) return (v/1e5).toFixed(2)+" L";
  if(a>=1e3) return (v/1e3).toFixed(2)+" K";
  return v.toLocaleString("en-IN");
};
const norm = s => String(s||"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();
const normalizeSKU = s => String(s||"").toUpperCase().replace(/[\s\-\._]/g,"").trim();

function num(v){
  if(v==null) return NaN;
  if(typeof v==="number") return Number.isFinite(v)?v:NaN;
  let s=String(v).trim(); if(!s) return NaN; let neg=false;
  if(/^\(.*\)$/.test(s)){neg=true; s=s.slice(1,-1);}
  if(/-$/.test(s)){neg=true; s=s.slice(0,-1);}
  s=s.replace(/(₹|rs\.?|inr)/gi,"");
  s=s.replace(/,/g,"").replace(/\s+/g,"");
  s=s.replace(/\.(?=.*\.)/g,"");
  s=s.replace(/[^0-9.\-]/g,"");
  if(s===""||s==="."||s==="-"||s==="-.") return NaN;
  const n=Number(s); return Number.isFinite(n)?(neg?-n:n):NaN;
}

function parseToISO(x){
  if(x==null) return null;
  const n=num(x);
  if(Number.isFinite(n)){
    if(n>=20000&&n<=80000){ // Excel serial
      const d=new Date(Date.UTC(1899,11,30)+n*86400000); return isNaN(d)?null:d.toISOString();
    }
    if(n>1e11&&n<2e12){ const d=new Date(n); return isNaN(d)?null:d.toISOString(); }
    if(n>1e9&&n<2e10){ const d=new Date(n*1000); return isNaN(d)?null:d.toISOString(); }
  }
  if(typeof x==="string"){
    const s=x.trim();
    let m=s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
    if(m){ let[,d,mo,y]=m; if(y.length===2) y=(+y<50?'20':'19')+y; const dt=new Date(Date.UTC(+y,+mo-1,+d)); return isNaN(dt)?null:dt.toISOString(); }
    m=s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
    if(m){ const[,y,mo,dd]=m; const dt=new Date(Date.UTC(+y,+mo-1,+dd)); return isNaN(dt)?null:dt.toISOString(); }
    const dt=new Date(s); if(!isNaN(dt)) return dt.toISOString();
  }
  return null;
}
function monYYYY(dlike){
  const m=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const iso=parseToISO(dlike);
  if(iso){ const d=new Date(iso); return `${m[d.getUTCMonth()]}-${d.getUTCFullYear()}`; }
  const d2=new Date(dlike); if(!isNaN(d2)) return `${m[d2.getMonth()]}-${d2.getFullYear()}`;
  return String(dlike||"");
}
function monthsLeft(expISO){
  const now=new Date(); const d=new Date(expISO);
  if(isNaN(d)) return NaN;
  return (d.getUTCFullYear()-now.getUTCFullYear())*12 + (d.getUTCMonth()-now.getUTCMonth());
}

/////////////////////// state ///////////////////////
const COUNT_MISSING_AS_MISMATCH = false; // only numerical mismatches
const QTY_TOLERANCE = 1;

let invRows=[], invCols=[];
let batchRows=[], batchCols=[];
let committedRows=[], committedCols=[];
let batchBySKU=new Map(), invBySKU=new Map(), mismatchRows=[];

let invMap={sku:"",product:"",salesDesc:"",soh:"",committed:"",availSale:"",cost:"",value:""};
let batchMap={sku:"",qty:"",batchNo:"",mfg:"",exp:""};
let commitMap={sku:"",product:"",customer:"",qty:"",unused:"",rate:"",soid:"",status:""};

/////////////////////// header detection ///////////////////////
const H_SKU=["SKU","Item Code","Code","Product Code","SKU Code"];
const H_PRODUCT=["Item Name","Product Name","Item","SKU Name","Material","Product"];
const H_SDESC=["Sales Description","Alias Name","Description"];
const H_SOH=["Stock On Hand","Closing Stock","On Hand","Stock","Qty","Quantity"];
const H_COMMIT=["Committed Stock","Committed","Reserved","Allocated"];
const H_AVAIL=["Available for Sale","Available For Sale","Available","For Sale"];
const H_COST=["Purchase Price","Unit Cost","Cost","Purchase Rate","Unit Price","Cost Price","Rate"];
const H_VALUE=["Stock Value","Value","Inventory Value","Total Value","Amount"];

const H_SKU_B=["SKU","Item Code","Code"];
const H_QTY_BATCH=["Quantity Available","Batch Qty","Qty","Quantity"];
const H_BATCHNO=["Batch Number","Batch No","Batch","Lot"];
const H_MFG=["Manufactured Date","Mfg Date","Manufacturing Date","MFD","Production Date","Mfg"];
const H_EXP=["Expiry Date","Exp Date","EXP","Best Before","Use By","Expiry"];

const H_SKU_C=["sku","SKU","Item Code","Code"];
const H_PROD_C=["Item Name","Product Name","Product","Item"];
const H_CUST_C=["Customer Name","Customer"];
const H_QTY_C=["Committed Stock","Committed Qty","Qty","Quantity"];
const H_UNUSED_C=["unused_credits_receivable_amount_bcy","Unused Credits","Unused Credit"];
const H_RATE_C=["Rate","Price","Unit Price","Sales Price","Selling Price"];
const H_SOID_C=["Sales Order ID","SO ID","Order ID"];
const H_STATUS_C=["status","Status"];

const bestCol=(cols,hints)=>{
  const nc=cols.map(c=>({raw:c,n:norm(c)}));
  for(const h of hints){
    const nh=norm(h);
    const exact=nc.find(c=>c.n===nh); if(exact) return exact.raw;
    const part=nc.find(c=>c.n.includes(nh)||nh.includes(c.n)); if(part) return part.raw;
  }
  return "";
};

/////////////////////// parse arrays -> rows ///////////////////////
function findHeaderRow(arr){ let best=0,bestCnt=0;
  for(let i=0;i<Math.min(25,arr.length);i++){
    const cnt=(arr[i]||[]).filter(x=>String(x||"").trim()!=="").length;
    if(cnt>bestCnt){bestCnt=cnt;best=i;}
  }
  return best;
}
function looksLikeLabel(text){
  const s=String(text||"").trim(); if(!s) return false;
  if(/\d/.test(s)) return false;
  if(s.length>2&&s===s.toUpperCase()) return true;
  return /^[A-Za-z][A-Za-z\s/&\-()]+$/.test(s);
}

function buildInvFromArrays(arr){
  if(!arr.length) return {cols:[],rows:[]};
  const clean=arr.map(r=>(r||[]).map(c=>typeof c==="string"?c.trim():c));
  const headerIndex=findHeaderRow(clean);
  const headers=(clean[headerIndex]||[]).map(h=>String(h||"").replace(/\s+/g," ").trim());
  const start=headerIndex+1;

  let prodIdx=headers.findIndex(h=>/item|product|name|description/i.test(h));
  if(prodIdx<0) prodIdx=0;
  const skuIdx=headers.findIndex(h=>/sku|code|item\s*code|product\s*code/i.test(h));

  const out=[]; let currentCategory="(Uncategorized)";
  for(let i=start;i<clean.length;i++){
    const r=clean[i]||[];
    const prodCell=r[prodIdx]; const skuCell=skuIdx>=0?r[skuIdx]:"";
    if((!skuCell||String(skuCell).trim()==="") && looksLikeLabel(prodCell)){
      currentCategory=String(prodCell||"").trim(); continue;
    }
    const o={Category:currentCategory};
    headers.forEach((h,idx)=>{ if(h) o[h]=r[idx]; });
    if(String(o[headers[prodIdx]]||"").trim()==="") continue;
    out.push(o);
  }
  const colSet=new Set(headers.filter(Boolean)); colSet.add("Category");
  return {cols:[...colSet],rows:out};
}
function buildBatchFromArrays(arr){
  if(!arr.length) return {cols:[],rows:[]};
  const clean=arr.map(r=>(r||[]).map(c=>typeof c==="string"?c.trim():c));
  const headerIndex=findHeaderRow(clean);
  const headers=(clean[headerIndex]||[]).map(h=>String(h||"").replace(/\s+/g," ").trim());
  const start=headerIndex+1; const out=[];
  for(let i=start;i<clean.length;i++){
    const r=clean[i]||[]; const o={}; headers.forEach((h,idx)=>{ if(h) o[h]=r[idx]; });
    const any=Object.values(o).some(v=>String(v||"").trim()!==""); if(any) out.push(o);
  }
  return {cols:headers,rows:out};
}
function buildCommittedFromArrays(arr){
  if(!arr.length) return {cols:[],rows:[]};
  const clean=arr.map(r=>(r||[]).map(c=>typeof c==="string"?c.trim():c));
  const headerIndex=findHeaderRow(clean);
  const headers=(clean[headerIndex]||[]).map(h=>String(h||"").replace(/\s+/g," ").trim());
  const start=headerIndex+1; const out=[];
  for(let i=start;i<clean.length;i++){
    const r=clean[i]||[]; const o={}; headers.forEach((h,idx)=>{ if(h) o[h]=r[idx]; });
    const any=Object.values(o).some(v=>String(v||"").trim()!==""); if(any) out.push(o);
  }
  return {cols:headers,rows:out};
}

/////////////////////// auto-detect maps ///////////////////////
function autoDetectInventory(){
  invMap.product  = bestCol(invCols,H_PRODUCT) || invCols[0] || "";
  invMap.salesDesc= bestCol(invCols,H_SDESC)   || "";
  invMap.sku      = bestCol(invCols,H_SKU);
  invMap.soh      = bestCol(invCols,H_SOH);
  invMap.committed= bestCol(invCols,H_COMMIT)  || "";
  invMap.availSale= bestCol(invCols,H_AVAIL)   || "";
  const pp = invCols.find(c=>norm(c)==="purchase price");
  invMap.cost     = pp || bestCol(invCols,H_COST) || "";
  invMap.value    = bestCol(invCols,H_VALUE)   || "";
}
function autoDetectBatch(){
  batchMap.sku    = bestCol(batchCols,H_SKU_B);
  batchMap.qty    = bestCol(batchCols,H_QTY_BATCH);
  batchMap.batchNo= bestCol(batchCols,H_BATCHNO);
  batchMap.mfg    = bestCol(batchCols,H_MFG);
  batchMap.exp    = bestCol(batchCols,H_EXP);
}
function autoDetectCommitted(){
  commitMap.sku     = bestCol(committedCols,H_SKU_C);
  commitMap.product = bestCol(committedCols,H_PROD_C);
  commitMap.customer= bestCol(committedCols,H_CUST_C);
  commitMap.qty     = bestCol(committedCols,H_QTY_C);
  commitMap.unused  = bestCol(committedCols,H_UNUSED_C);
  commitMap.rate    = bestCol(committedCols,H_RATE_C) || "";
  commitMap.soid    = bestCol(committedCols,H_SOID_C);
  commitMap.status  = bestCol(committedCols,H_STATUS_C);
}

/////////////////////// dropdowns ///////////////////////
function initDropdowns(){
  const cats=[...new Set(invRows.map(r=>String(r.Category||"(Uncategorized)")))].sort((a,b)=>a.localeCompare(b));
  const target="medyra onco";
  const matchCat = cats.find(c=>c.toLowerCase()===target) || cats.find(c=>c.toLowerCase().includes("medyra")&&c.toLowerCase().includes("onco")) || cats[0] || "";
  if($("scCategory")) $("scCategory").innerHTML=`${cats.map(c=>`<option ${c===matchCat?'selected':''}>${c}</option>`).join("")}`;
  rebuildSCProducts();

  // Details default = Medyra Onco
  if($("detailCat")){
    $("detailCat").innerHTML = `<option value="">All</option>` + cats.map(c=>`<option ${c===matchCat?'selected':''}>${c}</option>`).join("");
  }

  // Customer filter (for global committed list)
  const custCol = commitMap.customer || (committedCols.find(c=>/customer/i.test(c))||"");
  const customers = [...new Set(committedRows.map(r=>String(r[custCol]||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  if($("commitCustomerFilter")){
    $("commitCustomerFilter").innerHTML = `<option value="">All</option>` + customers.map(c=>`<option>${c}</option>`).join("");
  }
}
function rebuildSCProducts(){
  const scCategory=$("scCategory"), scProduct=$("scProduct"), scSub=$("scProductSub");
  if(!scCategory||!scProduct) return;
  const catSel=scCategory.value;
  const pcol=invMap.product||invMap.salesDesc, skuCol=invMap.sku;
  const items=invRows
    .filter(r=>String(r.Category||"(Uncategorized)")===(catSel||""))
    .map(r=>({name:String(r[pcol]||"").trim(),sku:String(r[skuCol]||"").trim(),sales:String(r[invMap.salesDesc]||"")}))
    .filter(o=>o.name)
    .sort((a,b)=>a.name.localeCompare(b.name));
  scProduct.innerHTML=items.map(x=>`<option value="${x.name}" data-sku="${x.sku}" data-sales="${x.sales.replace(/"/g,'&quot;')}">${x.name}</option>`).join("") || `<option>(No Options)</option>`;
  const sel=scProduct.options[0]; if(scSub) scSub.textContent = sel? (sel.getAttribute("data-sales")||"") : "";
}

/////////////////////// compute & merge ///////////////////////
function computeInvValue(r){
  const q=num(r[invMap.soh]), c=num(r[invMap.cost]);
  if(Number.isFinite(q) && Number.isFinite(c)) return q*c;
  const v=num(r[invMap.value]); if(Number.isFinite(v)) return v;
  return NaN;
}
function mergeAndCompute(){
  batchBySKU=new Map();
  const skuB=batchMap.sku, qB=batchMap.qty, bNo=batchMap.batchNo, mfg=batchMap.mfg, exp=batchMap.exp;
  for(const b of batchRows){
    const skuRaw=String(b[skuB]||"").trim(); if(!skuRaw) continue;
    const sku=normalizeSKU(skuRaw);
    const obj={
      _skuRaw:skuRaw,_sku:sku,_qty:num(b[qB]),
      _batch:String(b[bNo]||"").trim(),
      _mfgRaw:b[mfg],_mfgISO:parseToISO(b[mfg]),
      _expRaw:b[exp],_expISO:parseToISO(b[exp])
    };
    (batchBySKU.get(sku)||batchBySKU.set(sku,[]).get(sku)).push(obj);
  }

  invBySKU=new Map();
  for(const r of invRows){ const sku=normalizeSKU(r[invMap.sku]); if(!sku) continue; invBySKU.set(sku,r); }

  mismatchRows=[];
  const allSKUs=new Set([...invBySKU.keys(),...batchBySKU.keys()]);
  for(const sku of allSKUs){
    const inv=invBySKU.get(sku)||null;
    const batches=batchBySKU.get(sku)||[];
    const invQty=inv?num(inv[invMap.soh]):NaN;

    let batchQty=0;
    for(const b of batches){ const q=num(b._qty); if(Number.isFinite(q)) batchQty+=q; }

    const hasInv=inv!==null && Number.isFinite(invQty);
    const hasBatch=batches.length>0 && Number.isFinite(batchQty);

    if(hasInv && hasBatch){
      const diff=batchQty-invQty;
      if(Math.abs(diff)>QTY_TOLERANCE){
        mismatchRows.push({
          SKU: inv[invMap.sku]||sku,
          Product: inv[invMap.product]||inv[invMap.salesDesc]||"",
          Category: inv.Category||"(Uncategorized)",
          InvQty: invQty, BatchQty: batchQty, Diff: diff, Reason: "Quantity mismatch"
        });
      }
    } else if (COUNT_MISSING_AS_MISMATCH){
      let reason=""; if(hasInv && !hasBatch) reason="Missing in batch file"; if(!hasInv && hasBatch) reason="Missing in inventory";
      if(reason){
        mismatchRows.push({
          SKU:(inv&&(inv[invMap.sku]||sku))||(batches[0]?batches[0]._skuRaw:sku),
          Product:inv?(inv[invMap.product]||inv[invMap.salesDesc]||""):"(Unknown)",
          Category:inv?(inv.Category||"(Uncategorized)"):"(Unknown)",
          InvQty:hasInv?invQty:"", BatchQty:hasBatch?batchQty:"", Diff:"", Reason:reason
        });
      }
    }
  }
}

/////////////////////// KPIs & charts ///////////////////////
function kpis(){
  const pcol=invMap.product||invMap.salesDesc, soh=invMap.soh;
  const skuNames=new Set(); let totalSOH=0,totalValue=0;
  for(const r of invRows){
    const name=String(r[pcol]||""); if(name) skuNames.add(name);
    const q=num(r[soh]); if(Number.isFinite(q)) totalSOH+=q;
    const v=computeInvValue(r); if(Number.isFinite(v)) totalValue+=v;
  }
  const mismatchCount=Array.isArray(mismatchRows)?mismatchRows.length:0;

  if($("kpiSKUs"))$("kpiSKUs").textContent=skuNames.size.toLocaleString("en-IN");
  if($("kpiSOH"))$("kpiSOH").textContent=totalSOH.toLocaleString("en-IN");
  if($("kpiInvValue"))$("kpiInvValue").textContent=fmtIN(totalValue);
  if($("kpiMismatch"))$("kpiMismatch").textContent=mismatchCount.toLocaleString("en-IN");

  if($("status")) $("status").textContent =
    `Inventory: ${invRows.length.toLocaleString()} • SKUs: ${skuNames.size} • Qty: ${totalSOH.toLocaleString("en-IN")} • Value: ${fmtIN(totalValue)} • Batches: ${batchRows.length} • Committed: ${committedRows.length}`;
}
function topProductsChart(){
  const el="topProducts";
  const map=new Map(); const pc=invMap.product||invMap.salesDesc;
  for(const r of invRows){ const name=String(r[pc]||""); const v=computeInvValue(r); if(!name||!Number.isFinite(v)) continue; map.set(name,(map.get(name)||0)+v); }
  const arr=[...map.entries()].sort((a,b)=>b[1]-a[1]).slice(0,20);
  if(!arr.length){ $(el).innerHTML='<div class="muted">No value data.</div>'; return; }
  Plotly.newPlot(el,[{x:arr.map(a=>a[0]),y:arr.map(a=>a[1]),type:'bar',text:arr.map(a=>fmtIN(a[1])),hovertemplate:"%{x}<br>%{text}<extra></extra>"}],
    {margin:{t:10,b:100},xaxis:{tickangle:-30},yaxis:{title:"Value"}},{responsive:true});
}
function stockByCategoryChart(){
  const el="stockByCategory";
  const map=new Map();
  for(const r of invRows){
    const v=computeInvValue(r); if(!Number.isFinite(v)) continue;
    const cat=String(r.Category||"(Uncategorized)");
    map.set(cat,(map.get(cat)||0)+v);
  }
  const arr=[...map.entries()].sort((a,b)=>b[1]-a[1]);
  if(!arr.length){ $(el).innerHTML='<div class="muted">No category value data.</div>'; return; }
  Plotly.newPlot(el,[{x:arr.map(a=>a[0]),y:arr.map(a=>a[1]),type:'bar',text:arr.map(a=>fmtIN(a[1])),hovertemplate:"%{x}<br>%{text}<extra></extra>"}],
    {margin:{t:10,b:80},xaxis:{tickangle:-20},yaxis:{title:"Stock Value"}},{responsive:true});
}

/////////////////////// tables ///////////////////////
function mismatchTable(){
  const el=$("mismatchTable"); if(!el) return;
  if(!mismatchRows.length){ el.innerHTML='<div class="muted">No mismatches in current view.</div>'; return; }
  el.innerHTML=`<table class="table-mismatch"><thead><tr><th>SKU</th><th>Product</th><th>Category</th><th>Stock On Hand</th><th>Batch Available</th><th>Diff</th><th>Reason</th></tr></thead><tbody>
    ${mismatchRows.slice(0,2000).map(r=>`<tr>
      <td>${r.SKU}</td><td>${r.Product}</td><td>${r.Category}</td>
      <td>${r.InvQty!==""?Number(r.InvQty).toLocaleString("en-IN"):""}</td>
      <td>${r.BatchQty!==""?Number(r.BatchQty).toLocaleString("en-IN"):""}</td>
      <td>${r.Diff!==""?Number(r.Diff).toLocaleString("en-IN"):""}</td>
      <td>${r.Reason}</td>
    </tr>`).join("")}
  </tbody></table>`;
}

function expiryTable(){
  const el=$("expiryTable"); if(!el) return;
  if(!batchRows.length){ el.innerHTML='<div class="muted">Load batch file to see expiry list.</div>'; return; }
  const skuCol=batchMap.sku, bNo=batchMap.batchNo, mfg=batchMap.mfg, exp=batchMap.exp, qty=batchMap.qty;
  const invBy=new Map(invRows.map(r=>[normalizeSKU(r[invMap.sku]||""),r]));
  const rows=[];
  for(const b of batchRows){
    const sku=normalizeSKU(b[skuCol]); if(!sku) continue;
    const expISO=parseToISO(b[exp]); if(!expISO) continue;
    const mLeft=monthsLeft(expISO); if(!Number.isFinite(mLeft) || mLeft>=11) continue;
    const inv=invBy.get(sku);
    const prod=inv? (inv[invMap.product]||inv[invMap.salesDesc]||"") : "(Unknown)";
    rows.push({
      SKU:inv? (inv[invMap.sku]) : sku,
      Product:prod,
      Batch:String(b[bNo]||""),
      Mfg:monYYYY(b[mfg]),
      Exp:monYYYY(expISO),
      Avail:Number(num(b[qty])||0),
      Months:mLeft
    });
  }
  if(!rows.length){ el.innerHTML='<div class="muted">No products expiring in < 11 months.</div>'; return; }
  rows.sort((a,b)=>a.Months-b.Months);
  el.innerHTML=`<table class="table-expiry"><thead><tr><th></th><th>SKU</th><th>Product</th><th>Batch No</th><th>Mfg</th><th>Expiry</th><th>Available Qty</th><th>Months left</th></tr></thead>
    <tbody>${rows.map(r=>`<tr>
      <td>${r.Months<6?'<span class="dot dotPulse" title="Less than 6 months"></span>':''}</td>
      <td>${r.SKU}</td><td>${r.Product}</td><td>${r.Batch}</td><td>${r.Mfg}</td><td>${r.Exp}</td>
      <td>${Number(r.Avail).toLocaleString('en-IN')}</td><td>${r.Months}</td>
    </tr>`).join("")}</tbody></table>`;
}

function detailsTable(){
  const el=$("table"); if(!el || !invRows.length) return;
  const catSel= $("detailCat") ? $("detailCat").value : "";
  const rows=invRows.filter(r=>!catSel || String(r.Category||"(Uncategorized)")===(catSel));
  const c=Object.keys(rows[0]||{});
  el.innerHTML=`<table class="table-details"><thead><tr>${c.map(x=>`<th>${x}</th>`).join("")}</tr></thead>
    <tbody>${rows.slice(0,2000).map(r=>`<tr>${c.map(k=>`<td>${r[k]}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}

function zeroStockTable(){
  const el=$("zeroStockTable"); if(!el) return;
  if(!invRows.length){ el.innerHTML='<div class="muted">Load Inventory to see zero-stock list.</div>'; return; }
  const cat = "medyra onco";
  const pcol=invMap.product||invMap.salesDesc;
  const rows = invRows.filter(r=>{
    const catMatch = String(r.Category||"").trim().toLowerCase()===cat;
    const q = num(r[invMap.soh]);
    return catMatch && Number.isFinite(q) && q===0;
  });
  if(!rows.length){ el.innerHTML='<div class="muted">No zero-stock items found in Medyra Onco.</div>'; return; }
  el.innerHTML = `<table class="table-zerostock"><thead><tr><th>Product</th><th>SKU</th><th>Sales Description</th><th>Stock On Hand</th><th>Category</th></tr></thead>
    <tbody>${rows.map(r=>`<tr>
      <td>${r[pcol]||""}</td>
      <td>${r[invMap.sku]||""}</td>
      <td>${r[invMap.salesDesc]||""}</td>
      <td>${Number(num(r[invMap.soh])||0).toLocaleString('en-IN')}</td>
      <td>${r.Category||""}</td>
    </tr>`).join("")}</tbody></table>`;
}

// Committed Items (global, by customer)
function renderCommittedItemsByCustomer(){
  const el=$("commitItemsTable"); if(!el) return;
  if(!committedRows.length){ el.innerHTML='<div class="muted">Load Committed file to see items.</div>'; return; }

  const customerSel = $("commitCustomerFilter") ? $("commitCustomerFilter").value : "";
  const skuCol = commitMap.sku || (committedCols.find(c=>/sku/i.test(c))||"");
  const prodCol= commitMap.product || (committedCols.find(c=>/item name|product name|product/i.test(c))||"");
  const custCol= commitMap.customer || (committedCols.find(c=>/customer/i.test(c))||"");
  const qtyCol = commitMap.qty || (committedCols.find(c=>/committed/i.test(c))||"");
  const rateCol= commitMap.rate || (committedCols.find(c=>/(^| )rate$|price/i.test(c))||"");

  let rows = committedRows.map(r=>{
    const cust = String(r[custCol]||"").trim();
    const prod = String(r[prodCol]||"").trim();
    const qty  = num(r[qtyCol]);
    const rate = num(r[rateCol]);
    return { Customer:cust, Product:prod, Qty:Number.isFinite(qty)?qty:NaN, Rate:Number.isFinite(rate)?rate:NaN };
  });

  if(customerSel){ rows = rows.filter(r=>r.Customer===customerSel); }
  if(!rows.length){ el.innerHTML='<div class="muted">No committed rows for the selected customer.</div>'; return; }

  const withAmt = rows.map(r=>({ ...r, Amount: (Number.isFinite(r.Qty)&&Number.isFinite(r.Rate)) ? r.Qty*r.Rate : NaN }));
  const tQty = withAmt.reduce((a,b)=>a+(Number.isFinite(b.Qty)?b.Qty:0),0);
  const tAmt = withAmt.reduce((a,b)=>a+(Number.isFinite(b.Amount)?b.Amount:0),0);

  el.innerHTML = `<table class="table-committed">
    <thead><tr><th>Customer</th><th>Product</th><th>Qty</th><th>Rate</th><th>Amount</th></tr></thead>
    <tbody>
      ${withAmt.map(r=>`<tr>
        <td>${r.Customer||""}</td>
        <td>${r.Product||""}</td>
        <td>${Number.isFinite(r.Qty)?r.Qty.toLocaleString('en-IN'):"—"}</td>
        <td>${Number.isFinite(r.Rate)?fmtIN(r.Rate):"—"}</td>
        <td>${Number.isFinite(r.Amount)?fmtIN(r.Amount):"—"}</td>
      </tr>`).join("")}
    </tbody>
    <tfoot><tr><th colspan="2">Total</th><th>${tQty.toLocaleString('en-IN')}</th><th></th><th>${fmtIN(tAmt)}</th></tr></tfoot>
  </table>`;
}

// Stock Check → committed for selected product
function renderCommittedForSelected(invRow){
  const el=$("lookupCommitted"); if(!el) return;
  if(!committedRows.length){ el.innerHTML=""; return; }
  const skuInv = normalizeSKU(invRow[invMap.sku]||"");
  const prodInv = String(invRow[invMap.product]||invRow[invMap.salesDesc]||"").trim();

  const skuCol= commitMap.sku || (committedCols.find(c=>/sku/i.test(c))||"");
  const prodCol= commitMap.product || (committedCols.find(c=>/item name|product name|product/i.test(c))||"");
  const custCol= commitMap.customer || (committedCols.find(c=>/customer/i.test(c))||"");
  const qtyCol = commitMap.qty || (committedCols.find(c=>/committed/i.test(c))||"");
  const unusedCol = commitMap.unused || (committedCols.find(c=>/unused.*credit/i.test(c))||"");

  let rows = committedRows.filter(r=>{
    const csku = normalizeSKU(r[skuCol]||"");
    if(csku && skuInv) return csku===skuInv;
    const pname = String(r[prodCol]||"").trim();
    return pname && prodInv && pname===prodInv;
  }).map(r=>({
    Customer: String(r[custCol]||""),
    Qty: num(r[qtyCol]),
    Unused: num(r[unusedCol])
  }));

  if(!rows.length){ el.innerHTML='<div class="muted">No committed rows found for this product.</div>'; return; }

  const agg = new Map();
  for(const r of rows){
    const key=r.Customer||"(Unknown)";
    const cur=agg.get(key)||{Customer:key,Qty:0,Unused:0};
    cur.Qty += Number.isFinite(r.Qty)?r.Qty:0;
    cur.Unused += Number.isFinite(r.Unused)?r.Unused:0;
    agg.set(key,cur);
  }
  const out=[...agg.values()].sort((a,b)=>String(a.Customer).localeCompare(String(b.Customer)));
  const tQty = out.reduce((a,b)=>a+b.Qty,0);
  const tUnused = out.reduce((a,b)=>a+b.Unused,0);

  el.innerHTML = `
    <div class="muted" style="margin:6px 0">Committed Stock Details (for selected product)</div>
    <div class="box" style="max-height:320px">
      <table class="table-committed">
        <thead><tr><th>Customer Name</th><th>Committed Qty</th><th>Unused Credits</th></tr></thead>
        <tbody>
          ${out.map(r=>`<tr><td>${r.Customer}</td><td>${Number.isFinite(r.Qty)?r.Qty.toLocaleString('en-IN'):""}</td><td>${Number.isFinite(r.Unused)?fmtIN(r.Unused):""}</td></tr>`).join("")}
        </tbody>
        <tfoot><tr><th>Total</th><th>${tQty.toLocaleString('en-IN')}</th><th>${fmtIN(tUnused)}</th></tr></tfoot>
      </table>
    </div>`;
}

// Stock Check → batches for selected product (with sticky totals)
function renderBatchesForSelected(invRow){
  const el = $("lookupBatches"); if(!el) return;
  if(!batchRows.length){ el.innerHTML=""; return; }

  const sku = normalizeSKU(invRow[invMap.sku]); 
  if(!sku){ el.innerHTML=""; return; }

  const skuCol = batchMap.sku, qCol = batchMap.qty, bNo = batchMap.batchNo,
        mfg = batchMap.mfg, exp = batchMap.exp;

  const rows = batchRows
    .filter(r => normalizeSKU(r[skuCol]) === sku)
    .map(r => {
      const q = num(r[qCol]);
      return {
        Batch : String(r[bNo] || ""),
        Mfg   : monYYYY(r[mfg]),
        Exp   : monYYYY(parseToISO(r[exp]) || r[exp]),
        Qty   : Number.isFinite(q) ? q : 0
      };
    });

  if(!rows.length){
    el.innerHTML = '<div class="muted">No batch details found for this SKU.</div>';
    return;
  }

  const total = rows.reduce((a,b)=> a + (b.Qty || 0), 0);

  el.innerHTML = `
    <div class="muted" style="margin:6px 0">
      Batches for <b>${invRow[invMap.sku] || sku}</b> (${rows.length}) • 
      Total available qty: <b>${total.toLocaleString('en-IN')}</b>
    </div>
    <div class="box" style="max-height:320px">
      <table class="table-batches">
        <thead>
          <tr>
            <th>#</th><th>Batch No</th><th>Mfg</th><th>Expiry</th>
            <th>Available Qty</th><th>% of SKU</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((r,i)=>`
            <tr>
              <td>${i+1}</td>
              <td>${r.Batch}</td>
              <td>${r.Mfg}</td>
              <td>${r.Exp}</td>
              <td>${r.Qty.toLocaleString('en-IN')}</td>
              <td>${total ? ((r.Qty/total)*100).toFixed(1) : "0.0"}%</td>
            </tr>`).join("")}
        </tbody>
        <tfoot>
          <tr>
            <th colspan="4" class="t-right">Total Available</th>
            <th>${total.toLocaleString('en-IN')}</th>
            <th></th>
          </tr>
        </tfoot>
      </table>
    </div>`;
}

/////////////////////// grouped renderers ///////////////////////
function expiryRender(){ expiryTable(); }
function detailsRender(){ detailsTable(); }
function zeroStockRender(){ zeroStockTable(); }
function committedItemsRender(){ renderCommittedItemsByCustomer(); }

function refresh(){
  mergeAndCompute();
  kpis(); stockByCategoryChart(); topProductsChart(); mismatchTable(); expiryRender(); detailsRender();
  zeroStockRender(); committedItemsRender();
  initDropdowns();
  if($("scProduct") && $("scProduct").value) updateStockCheck();
}

/////////////////////// file IO ///////////////////////
async function readFileToArrays(file){
  const ext=(file.name.split(".").pop()||"").toLowerCase();
  const asArrayBuffer = f => new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(r.result); r.onerror=rej; r.readAsArrayBuffer(f); });
  if(ext==="csv"){
    const ab=await asArrayBuffer(file);
    const wb=XLSX.read(ab); const sheet=wb.SheetNames[0];
    return XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,defval:""});
  }
  const ab=await asArrayBuffer(file);
  const wb=XLSX.read(ab);
  let sheet=wb.SheetNames[0];
  for(const s of wb.SheetNames){
    const a=XLSX.utils.sheet_to_json(wb.Sheets[s],{header:1,raw:true,defval:""});
    if(a && a.length && a.some(r=>r.some(c=>String(c||"").trim()!==""))){ sheet=s; break; }
  }
  return XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,defval:""});
}

function autoDetectAll(){ autoDetectInventory(); autoDetectBatch(); autoDetectCommitted(); initDropdowns(); refresh(); }

/////////////////////// events ///////////////////////
const fileInv = $("fileInv");
const fileBatch = $("fileBatch");
const fileCommitted = $("fileCommitted");

if(fileInv) fileInv.addEventListener("change", async e=>{
  const f=e.target.files[0]; if(!f) return;
  try{ const arr=await readFileToArrays(f); const built=buildInvFromArrays(arr||[]); invCols=built.cols; invRows=built.rows; autoDetectAll(); }
  catch(err){ console.error(err); alert("Failed to read Inventory file."); }
});
if(fileBatch) fileBatch.addEventListener("change", async e=>{
  const f=e.target.files[0]; if(!f) return;
  try{ const arr=await readFileToArrays(f); const built=buildBatchFromArrays(arr||[]); batchCols=built.cols; batchRows=built.rows; autoDetectAll(); }
  catch(err){ console.error(err); alert("Failed to read Batch file."); }
});
if(fileCommitted) fileCommitted.addEventListener("change", async e=>{
  const f=e.target.files[0]; if(!f) return;
  try{ const arr=await readFileToArrays(f); const built=buildCommittedFromArrays(arr||[]); committedCols=built.cols; committedRows=built.rows; autoDetectAll(); }
  catch(err){ console.error(err); alert("Failed to read Committed file."); }
});

if($("scCategory")) $("scCategory").addEventListener("change",()=>{ rebuildSCProducts(); updateStockCheck(); });
if($("scProduct")) $("scProduct").addEventListener("change",updateStockCheck);
if($("detailCat")) $("detailCat").addEventListener("change",detailsRender);
if($("commitCustomerFilter")) $("commitCustomerFilter").addEventListener("change",committedItemsRender);

function updateStockCheck(){
  const scProduct=$("scProduct"); 
  if(!scProduct || !scProduct.value){
    const cards=$("scCards"); if(cards) cards.style.display="none";
    $("lookupCommitted").innerHTML=""; $("lookupBatches").innerHTML="";
    return;
  }
  const prod=scProduct.value;
  const inv=invRows.find(r=>String(r[invMap.product]||r[invMap.salesDesc]||"").trim()===prod);
  if(!inv){
    const cards=$("scCards"); if(cards) cards.style.display="none";
    $("lookupCommitted").innerHTML=""; $("lookupBatches").innerHTML="";
    return;
  }

  const soh = invMap.soh ? num(inv[invMap.soh]) : NaN;
  const com = invMap.committed ? num(inv[invMap.committed]) : NaN;
  const aval= invMap.availSale ? num(inv[invMap.availSale]) : NaN;
  const price = invMap.cost ? num(inv[invMap.cost]) : NaN;
  const value = (Number.isFinite(soh)&&Number.isFinite(price)) ? soh*price : NaN;

  const cards=$("scCards"); if(cards) cards.style.display="flex";
  $("lkProduct").textContent = prod;
  $("lkSOH").textContent = Number.isFinite(soh)?soh.toLocaleString("en-IN"):"—";
  $("lkCommitted").textContent = Number.isFinite(com)?com.toLocaleString("en-IN"):"—";
  $("lkAvail").textContent = Number.isFinite(aval)?aval.toLocaleString("en-IN"):"—";
  $("lkValue").textContent = Number.isFinite(value)?fmtIN(value):"—";

  renderCommittedForSelected(inv);
  renderBatchesForSelected(inv);
}
# Ultimate Sales & Profit Dashboard generator
# VS Code setup:
#   py -m pip install pandas openpyxl
#   py dashboard_sales.py
# The HTML output is saved beside the Excel file.

from pathlib import Path
import json
import os
import pandas as pd

# Prioritas lokasi: variabel lingkungan SALES_EXCEL_PATH, lalu folder script.
# Jika file Python dan Excel berada di folder AI II yang sama, tidak perlu diubah.
SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_PATH = Path(os.environ.get(
    "SALES_EXCEL_PATH",
    str(SCRIPT_DIR / "Ultimate_Sales_Profit_Dataset.xlsx")
)).expanduser()
OUTPUT_HTML = EXCEL_PATH.parent / "Ultimate_Sales_Profit_Dashboard.html"

def main():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(
            f"File Excel tidak ditemukan:\n{EXCEL_PATH}\n"
            "Pastikan Ultimate_Sales_Profit_Dataset.xlsx berada satu folder "
            "dengan dashboard_sales.py, atau atur SALES_EXCEL_PATH."
        )

    df = pd.read_excel(EXCEL_PATH, sheet_name="Data")
    required = ["Order ID","Order Date","Year","Month-Year","Product Name","Category",
                "Brand","Region","Sales Person","Discount %","Sales","Profit",
                "Profit Margin %","Order Status","Payment Mode","Return Status"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom tidak ditemukan: {missing}")

    df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
    df = df.dropna(subset=["Order Date","Sales","Profit"]).copy()
    df["Order Date"] = df["Order Date"].dt.strftime("%Y-%m-%d")
    # Bersihkan NaN/NaT menjadi null agar data tertanam sebagai JSON yang valid.
    clean_df = df.astype(object).where(pd.notna(df), None)
    data = json.dumps(clean_df.to_dict("records"), ensure_ascii=False, default=str, allow_nan=False).replace("</", "<\\/")

    html = r'''<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Executive Sales & Profit Dashboard</title><script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
:root{--navy:#142b4a;--blue:#3979b9;--teal:#0f9f95;--ink:#1e293b;--muted:#64748b;--line:#e2e8f0;--bg:#f4f7fb}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Inter,"Segoe UI",Arial,sans-serif}
header{background:linear-gradient(115deg,#142b4a,#1d4770);color:white;padding:25px 3%}
.head,main{max-width:1550px;margin:auto}.head{display:flex;justify-content:space-between;align-items:center;gap:15px}
h1{font-size:30px;margin:0 0 6px}.sub{font-size:13px;color:#d4e2f1}.tag{font-size:12px;border:1px solid #7893b1;border-radius:8px;padding:10px}
main{padding:20px 3% 40px}.filters,.panel,.kpi{background:white;border:1px solid var(--line);border-radius:11px;box-shadow:0 3px 12px #172b4d08}
.filters{position:sticky;top:0;z-index:5;padding:13px;margin-bottom:18px;background:#ffffffed;backdrop-filter:blur(7px)}
.fgrid{display:grid;grid-template-columns:repeat(8,minmax(95px,1fr));gap:9px;align-items:end}
label{display:block;font-size:10px;font-weight:700;color:var(--muted);margin-bottom:5px}
select,input{width:100%;height:35px;border:1px solid #d5deea;border-radius:6px;padding:5px;font-size:11px;background:white;color:var(--ink)}
select[multiple]{height:67px}button{height:35px;background:var(--navy);color:white;border:0;border-radius:6px;padding:0 13px;cursor:pointer}
.note{font-size:10px;color:#8793a4;margin-top:7px}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:11px;margin:17px 0}
.kpi{padding:14px;min-height:108px}.kl{font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase}
.kv{font-size:clamp(20px,2vw,27px);font-weight:750;color:var(--navy);margin:10px 0 5px;white-space:nowrap}.kf{font-size:10px;color:var(--muted)}
.up{color:#0c927e}.down{color:#d14d4d}.title{font-size:17px;font-weight:750;color:var(--navy);margin:24px 0 11px}
.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:13px}.panel{padding:16px;min-width:0}
.panel h3{font-size:14px;margin:0 0 4px;color:#263c58}.panel p{font-size:11px;color:var(--muted);margin:0 0 10px}
.wide{grid-column:1/-1}.chart{height:275px;position:relative}
.insights{display:grid;grid-template-columns:repeat(3,1fr);gap:11px}.insight{border-left:4px solid var(--teal)}
.insight h4{font-size:13px;margin:0 0 7px;color:var(--navy)}.insight p{font-size:12px;line-height:1.5;color:#536276;margin:0}
.tablebox{overflow:auto;max-height:390px}table{border-collapse:collapse;width:100%;font-size:11px;white-space:nowrap}
td,th{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left}th{position:sticky;top:0;background:#f1f5f9}
.foot{text-align:right;color:#8390a2;font-size:10px;margin-top:18px}
@media(max-width:1050px){.fgrid{grid-template-columns:repeat(4,1fr)}.kpis{grid-template-columns:repeat(3,1fr)}.insights{grid-template-columns:repeat(2,1fr)}}
@media(max-width:650px){.head{display:block}h1{font-size:23px}.tag{display:inline-block;margin-top:12px}.fgrid{grid-template-columns:repeat(2,1fr)}.kpis{grid-template-columns:repeat(2,1fr)}.grid,.insights{grid-template-columns:1fr}.wide{grid-column:auto}.chart{height:245px}}
</style></head><body>
<header><div class="head"><div><h1>Executive Sales &amp; Profit Dashboard</h1><div class="sub">Executive business performance overview · Transaction-level analysis</div></div>
<div class="tag" id="daterange">Dataset period: loading…<br>Data through Apr 2025</div></div></header>
<main>
<section class="filters"><div class="fgrid">
<div><label>FROM DATE</label><input id="from" type="date"></div><div><label>TO DATE</label><input id="to" type="date"></div>
<div><label>REGION (MULTI)</label><select id="region" multiple></select></div><div><label>CATEGORY (MULTI)</label><select id="category" multiple></select></div>
<div><label>BRAND (MULTI)</label><select id="brand" multiple></select></div><div><label>SALES PERSON</label><select id="rep"><option value="">All</option></select></div>
<div><label>ORDER STATUS</label><select id="status"><option value="">All</option></select></div><div><label>RETURN STATUS</label><select id="returned"><option value="">All</option><option>Yes</option><option>No</option></select></div>
<div><button id="reset">Reset Filters</button></div></div><div class="note">Untuk memilih beberapa nilai, tahan Ctrl (Windows) atau Cmd (Mac). Semua KPI dan grafik mengikuti filter.</div></section>
<section class="kpis" id="kpis"></section>
<div class="title">Sales &amp; Profit Performance</div>
<section class="grid">
<div class="panel wide"><h3>Sales &amp; Profit Trend Over Time</h3><p>Monthly sales and profit. Year 2025 is partial (January–April only).</p><div class="chart"><canvas id="trend"></canvas></div></div>
<div class="panel"><h3>Sales &amp; Margin by Category</h3><p>Revenue and profit margin by category.</p><div class="chart"><canvas id="categoryChart"></canvas></div></div>
<div class="panel"><h3>Regional Performance</h3><p>Sales and profit by region.</p><div class="chart"><canvas id="regionChart"></canvas></div></div>
<div class="panel"><h3>Brand Comparison</h3><p>Sales and profit margin by brand.</p><div class="chart"><canvas id="brandChart"></canvas></div></div>
<div class="panel"><h3>Top 10 Products by Sales</h3><p>Sales and profit ranking.</p><div class="chart"><canvas id="products"></canvas></div></div>
<div class="panel"><h3>Discount vs. Profit Margin</h3><p>Each point represents one order. Correlation is descriptive, not proof of causation.</p><div class="chart"><canvas id="scatter"></canvas></div></div>
<div class="panel"><h3>Order Status</h3><p>Operational status distribution.</p><div class="chart"><canvas id="statusChart"></canvas></div></div>
<div class="panel"><h3>Payment Mode</h3><p>Payment distribution.</p><div class="chart"><canvas id="paymentChart"></canvas></div></div>
</section>
<div class="title">Key Insights &amp; Management Questions</div><section class="insights" id="insights"></section>
<div class="title">Filtered Transaction Details</div><section class="panel"><div class="tablebox"><table><thead><tr>
<th>Order ID</th><th>Date</th><th>Product</th><th>Category</th><th>Region</th><th>Sales Person</th><th>Sales</th><th>Profit</th><th>Discount</th><th>Return</th>
</tr></thead><tbody id="details"></tbody></table></div></section>
<div class="foot">Source: Ultimate_Sales_Profit_Dataset.xlsx · Sheet: Data · Profit is counted once; Gross Profit and Net Profit are not duplicated.</div>
</main>
<script>
const DATA=__DATA__;
const $=id=>document.getElementById(id), money=n=>'$'+Math.round(n||0).toLocaleString('en-US'), pct=n=>(n||0).toFixed(1)+'%';
const sum=(a,k)=>a.reduce((s,r)=>s+(Number(r[k])||0),0), unique=(a,k)=>[...new Set(a.map(r=>r[k]).filter(v=>v!==null&&v!==''))].sort();
const charts={}, colors=['#3979b9','#0f9f95','#f59e0b','#8b5cf6','#e05b65','#64748b'];
function fill(id,vals){vals.forEach(v=>{let o=document.createElement('option');o.value=v;o.textContent=v;$(id).appendChild(o)})}
function sel(id){return [...$(id).selectedOptions].map(o=>o.value)}
function getRows(){let f=$('from').value,t=$('to').value;return DATA.filter(r=>
(!f||r['Order Date']>=f)&&(!t||r['Order Date']<=t)&&(!sel('region').length||sel('region').includes(r.Region))&&
(!sel('category').length||sel('category').includes(r.Category))&&(!sel('brand').length||sel('brand').includes(r.Brand))&&
(!$('rep').value||r['Sales Person']===$('rep').value)&&(!$('status').value||r['Order Status']===$('status').value)&&
(!$('returned').value||r['Return Status']===$('returned').value))}
function yoy(rows,metric){let ys=unique(rows,'Year').map(Number).sort((a,b)=>a-b);if(ys.length<2)return 'YoY: perlu ≥2 tahun';
let a=rows.filter(r=>+r.Year===ys.at(-1)),b=rows.filter(r=>+r.Year===ys.at(-2)),v,w;
if(metric==='orders'){v=a.length;w=b.length}else if(metric==='aov'){v=sum(a,'Sales')/(a.length||1);w=sum(b,'Sales')/(b.length||1)}
else if(metric==='margin'){v=sum(a,'Profit')/(sum(a,'Sales')||1);w=sum(b,'Profit')/(sum(b,'Sales')||1)}
else if(metric==='returns'){v=a.filter(r=>r['Return Status']==='Yes').length/(a.length||1);w=b.filter(r=>r['Return Status']==='Yes').length/(b.length||1)}
else{v=sum(a,metric);w=sum(b,metric)}let d=w?(v-w)/Math.abs(w)*100:0;return `<span class="${d>=0?'up':'down'}">${d>=0?'▲':'▼'} ${Math.abs(d).toFixed(1)}%</span> vs ${ys.at(-2)} (latest ${ys.at(-1)})`}
function kpis(r){let s=sum(r,'Sales'),p=sum(r,'Profit'),n=r.length,ret=r.filter(x=>x['Return Status']==='Yes').length;
let c=[['Total Sales',money(s),yoy(r,'Sales')],['Total Profit',money(p),yoy(r,'Profit')],['Overall Profit Margin',pct(s?p/s*100:0),yoy(r,'margin')],['Total Orders',n.toLocaleString(),yoy(r,'orders')],['Average Order Value',money(n?s/n:0),yoy(r,'aov')],['Return Rate',pct(n?ret/n*100:0),yoy(r,'returns')]];
$('kpis').innerHTML=c.map(x=>`<div class="kpi"><div class="kl">${x[0]}</div><div class="kv">${x[1]}</div><div class="kf">${x[2]}</div></div>`).join('')}
function group(r,k){let m={};r.forEach(x=>(m[x[k]??'Unknown']??=[]).push(x));return m}
function chart(id,type,data,options={}){if(charts[id])charts[id].destroy();charts[id]=new Chart($(id),{type,data,options:{responsive:true,maintainAspectRatio:false,animation:{duration:200},plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:10}}}},...options}})}
function draw(r){
let months={};r.forEach(x=>{let d=new Date(x['Order Date']+'T00:00:00'),k=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');(months[k]??=[]).push(x)});
let keys=Object.keys(months).sort(),labels=keys.map(k=>{let [y,m]=k.split('-');return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+m-1]+' '+y.slice(2)});
chart('trend','bar',{labels,datasets:[{type:'bar',label:'Sales ($)',data:keys.map(k=>sum(months[k],'Sales')),backgroundColor:'#8bb6e8',yAxisID:'y'},
{type:'line',label:'Profit ($)',data:keys.map(k=>sum(months[k],'Profit')),borderColor:'#0f9f95',tension:.25,yAxisID:'y1',pointRadius:2}]},
{scales:{y:{beginAtZero:true,title:{display:true,text:'Sales ($)'},ticks:{callback:v=>'$'+Math.round(v/1000)+'k'}},y1:{position:'right',beginAtZero:true,grid:{drawOnChartArea:false},title:{display:true,text:'Profit ($)'},ticks:{callback:v=>'$'+Math.round(v/1000)+'k'}}}});
let c=group(r,'Category'),ck=Object.keys(c).sort();
chart('categoryChart','bar',{labels:ck,datasets:[{label:'Sales ($)',data:ck.map(k=>sum(c[k],'Sales')),backgroundColor:'#5b91ce',yAxisID:'y'},
{type:'line',label:'Margin (%)',data:ck.map(k=>sum(c[k],'Profit')/(sum(c[k],'Sales')||1)*100),borderColor:'#e09a35',yAxisID:'y1'}]},
{scales:{y:{beginAtZero:true,ticks:{callback:v=>'$'+Math.round(v/1000)+'k'}},y1:{position:'right',grid:{drawOnChartArea:false},ticks:{callback:v=>v+'%'}}}});
let g=group(r,'Region'),gk=Object.keys(g).sort();
chart('regionChart','bar',{labels:gk,datasets:[{label:'Sales ($)',data:gk.map(k=>sum(g[k],'Sales')),backgroundColor:'#5b91ce'},{label:'Profit ($)',data:gk.map(k=>sum(g[k],'Profit')),backgroundColor:'#31a99a'}]},
{scales:{y:{beginAtZero:true,ticks:{callback:v=>'$'+Math.round(v/1000)+'k'}}}});
let b=group(r,'Brand'),bk=Object.keys(b).sort();
chart('brandChart','bar',{labels:bk,datasets:[{label:'Sales ($)',data:bk.map(k=>sum(b[k],'Sales')),backgroundColor:'#5b91ce',yAxisID:'y'},
{type:'line',label:'Margin (%)',data:bk.map(k=>sum(b[k],'Profit')/(sum(b[k],'Sales')||1)*100),borderColor:'#e09a35',yAxisID:'y1'}]},
{scales:{y:{beginAtZero:true,ticks:{callback:v=>'$'+Math.round(v/1000)+'k'}},y1:{position:'right',grid:{drawOnChartArea:false},ticks:{callback:v=>v+'%'}}}});
let p=group(r,'Product Name'),pk=Object.keys(p).sort((a,b)=>sum(p[b],'Sales')-sum(p[a],'Sales')).slice(0,10);
chart('products','bar',{labels:pk,datasets:[{label:'Sales ($)',data:pk.map(k=>sum(p[k],'Sales')),backgroundColor:'#5b91ce'},{label:'Profit ($)',data:pk.map(k=>sum(p[k],'Profit')),backgroundColor:'#31a99a'}]},
{indexAxis:'y',scales:{x:{beginAtZero:true,ticks:{callback:v=>'$'+Math.round(v/1000)+'k'}}}});
let points=r.map(x=>({x:+x['Discount %'],y:+x['Profit Margin %']})).filter(x=>Number.isFinite(x.x)&&Number.isFinite(x.y));
chart('scatter','scatter',{datasets:[{label:'Orders',data:points,backgroundColor:'#3979b9aa',pointRadius:4}]},
{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`Discount ${c.raw.x}% · Margin ${c.raw.y}%`}}},scales:{x:{title:{display:true,text:'Discount (%)'}},y:{title:{display:true,text:'Profit Margin (%)'}}}});
let st=group(r,'Order Status'),sk=Object.keys(st);chart('statusChart','doughnut',{labels:sk,datasets:[{data:sk.map(k=>st[k].length),backgroundColor:colors}]});
let pm=group(r,'Payment Mode'),pmk=Object.keys(pm);chart('paymentChart','doughnut',{labels:pmk,datasets:[{data:pmk.map(k=>pm[k].length),backgroundColor:colors}]});
}
function insights(r){
let n=r.length,ret=n?r.filter(x=>x['Return Status']==='Yes').length/n*100:0;
let c=group(r,'Category'),cats=Object.keys(c).map(k=>({name:k,sales:sum(c[k],'Sales'),margin:sum(c[k],'Profit')/(sum(c[k],'Sales')||1)*100})).sort((a,b)=>b.sales-a.sales);
let marginLeader=[...cats].sort((a,b)=>b.margin-a.margin)[0];
let rg=group(r,'Region'),regions=Object.keys(rg).map(k=>({name:k,sales:sum(rg[k],'Sales'),margin:sum(rg[k],'Profit')/(sum(rg[k],'Sales')||1)*100})).sort((a,b)=>b.sales-a.sales);
let reps=group(r,'Sales Person'),rep=Object.keys(reps).map(k=>({name:k,sales:sum(reps[k],'Sales')})).sort((a,b)=>b.sales-a.sales)[0];
let mx=n?r.reduce((s,x)=>s+(+x['Discount %']||0),0)/n:0,my=n?r.reduce((s,x)=>s+(+x['Profit Margin %']||0),0)/n:0,num=0,dx=0,dy=0;
r.forEach(x=>{let a=(+x['Discount %']||0)-mx,b=(+x['Profit Margin %']||0)-my;num+=a*b;dx+=a*a;dy+=b*b});
let corr=dx&&dy?(num/Math.sqrt(dx*dy)).toFixed(2):'n/a',ys=unique(r,'Year').sort(),growth='';
if(ys.length>1){let a=r.filter(x=>String(x.Year)===ys.at(-1)),b=r.filter(x=>String(x.Year)===ys.at(-2));growth=`Latest selected year sales changed ${((sum(a,'Sales')/(sum(b,'Sales')||1)-1)*100).toFixed(1)}% vs ${ys.at(-2)}. 2025 is partial through April.`}
let cards=[
['Growth & revenue',`${cats[0]?cats[0].name+' is the leading sales category at '+money(cats[0].sales)+'. ':''}${growth||'Select at least two years to compare annual sales.'} Check seasonality before interpreting a partial-year change.`],
['Returns & operations',`${pct(ret)} of filtered orders are marked returned. Which products, regions, or delivery processes explain the returns? Investigate root causes before treating this rate as normal.`],
['Revenue is not margin',marginLeader?`${marginLeader.name} has the highest category margin (${pct(marginLeader.margin)}). Compare it with the sales leader (${cats[0]?.name||'n/a'}) when allocating promotions.`:'Insufficient category data to compare margin.'],
['Discount discipline',`Discount and profit-margin correlation is ${corr}. A negative relationship may indicate margin compression; review deep discounts and consider a discount floor. Correlation is not causation.`],
['Regional opportunity',regions.length?`${regions[0].name} leads sales (${money(regions[0].sales)}). ${[...regions].sort((a,b)=>b.margin-a.margin)[0].name} has the strongest regional margin. Could investment improve underperforming regions?`:'No regional data in this selection.'],
['Sales team',rep?`${rep.name} leads the selected sales team with ${money(rep.sales)} in sales. Consider sharing effective practices while accounting for territory and account differences.`:'No sales-person data in this selection.']
];
$('insights').innerHTML=cards.map(x=>`<div class="panel insight"><h4>${x[0]}</h4><p>${x[1]}</p></div>`).join('');
}
function update(){let r=getRows();kpis(r);draw(r);insights(r);let rows=[...r].sort((a,b)=>b['Order Date'].localeCompare(a['Order Date'])).slice(0,250);
$('details').innerHTML=rows.map(x=>`<tr><td>${x['Order ID']}</td><td>${x['Order Date']}</td><td>${x['Product Name']}</td><td>${x.Category}</td><td>${x.Region}</td><td>${x['Sales Person']}</td><td>${money(x.Sales)}</td><td>${money(x.Profit)}</td><td>${x['Discount %']}%</td><td>${x['Return Status']}</td></tr>`).join('');
$('daterange').innerHTML=`${r.length} of ${DATA.length} orders<br>Dataset through Apr 2025`}
function init(){let dates=DATA.map(x=>x['Order Date']).sort();$('from').value=dates[0];$('to').value=dates.at(-1);
fill('region',unique(DATA,'Region'));fill('category',unique(DATA,'Category'));fill('brand',unique(DATA,'Brand'));fill('rep',unique(DATA,'Sales Person'));fill('status',unique(DATA,'Order Status'));
['from','to','region','category','brand','rep','status','returned'].forEach(id=>$(id).addEventListener('change',update));
$('reset').onclick=()=>{['region','category','brand'].forEach(id=>[...$(id).options].forEach(o=>o.selected=false));['rep','status','returned'].forEach(id=>$(id).value='');$('from').value=dates[0];$('to').value=dates.at(-1);update()};update()}
init();
</script></body></html>'''
    html = html.replace("__DATA__", data)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print("Dashboard berhasil dibuat:", OUTPUT_HTML)
    print("Jumlah order:", len(df))
    print(f"Total Sales: ${df['Sales'].sum():,.2f}")
    print(f"Total Profit: ${df['Profit'].sum():,.2f}")
    print(f"Overall Margin: {df['Profit'].sum()/df['Sales'].sum()*100:.2f}%")
    print(f"Return Rate: {df['Return Status'].eq('Yes').mean()*100:.1f}%")

if __name__ == "__main__":
    main()

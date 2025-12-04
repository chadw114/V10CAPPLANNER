import sys
import os
import json
import webbrowser
import random
import csv
import io
from threading import Timer
from flask import Flask, request, jsonify, render_template_string, Response

app = Flask(__name__)

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HTML_PART_1 = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Capacity Planner v10</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.1/jspdf.plugin.autotable.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@2.2.1/dist/chartjs-plugin-annotation.min.js"></script>
<style>
:root{--primary:#2563eb;--bg:#f8fafc;--card-bg:#ffffff;--border:#e2e8f0;--text:#1e293b;--text-light:#64748b;--danger-bg:#fee2e2;--danger-text:#991b1b;--warning-bg:#fef3c7;--warning-text:#92400e;--success-bg:#dcfce7;--success-text:#166534}
*{box-sizing:border-box}body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);margin:0;padding:20px;overflow-y:scroll}
.container{max-width:1600px;margin:0 auto;padding-bottom:100px}
.header{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:10px}
h1{margin:0;font-size:26px;color:#0f172a;letter-spacing:-.5px}.subtitle{color:var(--text-light);font-size:14px;margin-top:5px}
.btn{padding:10px 20px;border-radius:8px;font-weight:600;border:none;cursor:pointer;transition:.2s;font-size:14px;display:inline-flex;align-items:center;gap:8px}.btn-primary{background:#2563eb;color:#fff;box-shadow:0 2px 4px rgba(37,99,235,.15)}.btn-sec{background:#fff;border:1px solid #cbd5e1;color:#475569}.btn-add{background:#dbeafe;color:#1e40af}.btn-random{background:#fef3c7;color:#92400e;border:1px solid #fde68a}.btn-pdf{background:#fff;border:1px solid #cbd5e1;color:#475569}
.custom-select{appearance:none;padding:10px 30px 10px 15px;border-radius:8px;border:1px solid #cbd5e1;background-color:#fff;font-family:'Inter',sans-serif;font-size:13px;font-weight:600;color:#475569;cursor:pointer;background-image:url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");background-repeat:no-repeat;background-position:right 10px center;background-size:14px}
.cell-select{width:100%;padding:12px 8px;border:none;outline:none;text-align:center;background:transparent;font-family:inherit;color:#334155;appearance:none;cursor:pointer;font-weight:500}
.cell-select:focus{background:white;color:var(--primary);font-weight:700;box-shadow:inset 0 0 0 2px var(--primary);appearance:auto}
#global-tooltip{position:fixed;background:#1e293b;color:#fff;padding:14px 18px;border-radius:8px;font-size:14px;line-height:1.5;pointer-events:none;z-index:2147483647;opacity:0;transition:opacity .1s ease-out;box-shadow:0 10px 30px rgba(0,0,0,0.5);min-width:220px;max-width:320px;white-space:pre-line;text-align:center}
.card{background:var(--card-bg);border-radius:12px;box-shadow:0 2px 5px rgba(0,0,0,0.05);border:1px solid #e2e8f0;margin-bottom:20px;overflow:hidden;transition:all .2s ease}
.active-card{border-color:var(--primary);box-shadow:0 0 0 1px var(--primary),0 4px 12px rgba(37,99,235,.1)}
.card-header{padding:16px 24px;background:#fff;border-bottom:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center;cursor:pointer;border-radius:12px 12px 0 0}.card-header:hover{background:#f8fafc}
.card-title{font-weight:700;font-size:15px;color:#334155;display:flex;align-items:center}.arrow{transition:transform .3s ease;color:var(--text-light)}.card-body{display:none;padding:0}.card-body.open{display:block}
.grid-wrapper{overflow-x:auto;max-height:550px;border-radius:0 0 12px 12px}table.data-grid{width:100%;border-collapse:collapse;min-width:1300px;font-size:13px}.data-grid th{background:#ffffff;color:#1e40af;font-weight:700;text-transform:uppercase;font-size:11px;padding:12px;border-bottom:2px solid #e2e8f0;border-right:1px solid #f1f5f9;position:sticky;top:0;z-index:20}.data-grid td{border-bottom:1px solid #f1f5f9;padding:0;transition:background .1s}.data-grid tr:hover td{background:#dbeafe!important}
.col-name{position:sticky;left:0;z-index:10;background:#f8fafc;border-right:2px solid #e2e8f0;width:180px}.col-line{position:sticky;left:180px;z-index:10;background:#f0f9ff;border-right:2px solid #e2e8f0;width:120px}.col-rate{position:sticky;left:300px;z-index:10;background:#eff6ff;border-right:2px solid #dbeafe;width:80px}input.cell-inp{width:100%;padding:12px 8px;border:none;outline:0;text-align:center;background:0 0;font-family:inherit;color:#334155}input.cell-inp:focus{background:#fff;color:var(--primary);font-weight:700;box-shadow:inset 0 0 0 2px var(--primary)}
.del-wrapper{padding:4px;display:flex;justify-content:center;background:#fff;border-right:1px solid #f1f5f9}.del-row{cursor:pointer;color:#94a3b8;font-size:16px;width:28px;height:28px;border-radius:6px;background:#f8fafc;display:flex;align-items:center;justify-content:center}.del-row:hover{color:#ef4444;background:#fee2e2}
.hm-cell{font-weight:700;color:#334155;transition:.2s;cursor:help}.hm-red{background:#fee2e2;color:#991b1b;border:1px solid #fecaca}.hm-orange{background:#fef3c7;color:#92400e;border:1px solid #fde68a}.hm-green{background:#dcfce7;color:#166534;border:1px solid #bbf7d0}

/* UPDATED: Responsive Charts & Scroll */
.charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;padding:24px}
.chart-box{
    height:250px; /* Reduced size */
    position:relative;
    width:100%;
    overflow-x: auto; /* Enable scroll */
    padding-bottom: 20px; /* Bottom padding */
}
@media (max-width: 900px) { 
    .charts-grid { grid-template-columns: 1fr !important; gap: 25px; padding: 15px; } 
    .chart-box { height: 250px !important; } 
}

.res-table{width:100%;border-collapse:collapse;font-size:13px}.res-table th{background:#f8fafc;color:#475569;padding:12px 16px;text-align:left;font-weight:600;border-bottom:2px solid #e2e8f0;position:sticky;top:0}.res-table td{padding:10px 16px;border-bottom:1px solid #f1f5f9;color:#334155}.res-table tr:hover{background:#f8fafc}
.filters{display:flex;gap:10px;align-items:center;flex-wrap:wrap}.filter-group{display:flex;align-items:center}select.filter-sel{padding:6px 12px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px;outline:0;color:#475569;background:#fff}
.settings-grid{padding:20px;display:flex;gap:20px;align-items:center;flex-wrap:wrap}.setting-item{display:flex;flex-direction:column;gap:5px}.setting-label{font-size:12px;font-weight:600;color:#64748b}.setting-inp{padding:8px 12px;border:1px solid #cbd5e1;border-radius:6px;width:100px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:20px;margin-bottom:25px;display:none}

/* REMOVED HOVER EFFECTS */
.kpi-card{padding:24px;border-radius:12px;background:#ffffff;border:1px solid #e2e8f0;color:var(--text);}
.kpi-label{font-size:13px;font-weight:600;text-transform:uppercase;margin-bottom:8px;z-index:2;color:var(--text-light)}.kpi-value{font-size:32px;font-weight:800;z-index:2;color:#0f172a}

.input-footer{padding:20px;border-top:1px solid #e2e8f0;background:#fff}.footer-actions{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:10px}.input-help{font-size:12px;color:#64748b}.badge-key{background:#f1f5f9;border:1px solid #cbd5e1;border-radius:4px;padding:2px 6px;font-size:10px;font-weight:700;color:#475569}
.scheduler-container{padding:20px;overflow-x:auto}.cal-controls{margin-bottom:15px;display:flex;gap:10px;align-items:center}.cal-table{width:100%;border-collapse:collapse;min-width:800px;font-size:12px}.cal-table th{text-align:center;padding:8px;border:1px solid #e2e8f0;background:#f8fafc;width:3.5%;color:#64748b}.cal-table td{border:1px solid #e2e8f0;height:36px;padding:0}.cal-line-header{width:120px;font-weight:700;padding:0 10px!important;background:#fff;position:sticky;left:0;z-index:5;border-right:2px solid #e2e8f0!important}.day-cell.filled{background:#3b82f6;border-right:1px solid #fff;opacity:.9}
.day-cell.filled-red{background:#ef4444;border-right:1px solid #fff;opacity:.9}
.status-indicator{font-weight:700;font-size:13px;padding:6px 12px;border-radius:6px;background:#e2e8f0;color:#64748b;transition:0.3s}.stat-ok{background:#dcfce7;color:#166534;border:1px solid #bbf7d0}.stat-warn{background:#fef3c7;color:#92400e;border:1px solid #fde68a}.stat-crit{background:#fee2e2;color:#991b1b;border:1px solid #fecaca}
.info-icon{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#e2e8f0;color:#64748b;font-size:11px;font-weight:bold;cursor:help;margin-right:8px}.info-icon:hover{background:var(--primary);color:white}
.kpi-safe { background-color: #dcfce7 !important; border-color: #bbf7d0 !important; color: #166534 !important; }
.kpi-warn { background-color: #fef3c7 !important; border-color: #fde68a !important; color: #92400e !important; }
.kpi-crit { background-color: #fee2e2 !important; border-color: #fecaca !important; color: #991b1b !important; }
.status-pill{padding:6px 12px;border-radius:20px;font-weight:700;font-size:12px;display:inline-block;cursor:default;text-align:center;min-width:80px}
::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-button{display:none}::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:10px}::-webkit-scrollbar-thumb:hover{background:#94a3b8}
.hidden{display:none !important;} .rotate{transform:rotate(180deg)}
</style>
</head>
"""
HTML_PART_2 = r"""
<body>
<div id="global-tooltip"></div>
<div class="container">
  <div class="header">
    <div><h1>Capacity Planner v10</h1><div class="subtitle">Production Scheduler & Demand Analysis</div></div>
    <div style="display:flex; gap:10px;">
        <button class="btn btn-sec" onclick="window.location.href='/template'">⬇ CSV Template</button>
        <button class="btn btn-pdf" onclick="downloadPDF()" id="pdfBtn" disabled>Download Report</button>
    </div>
  </div>
  
  <div class="card" onclick="activateCard(this,event)"><div class="card-header" onclick="toggle('settings-body',this)"><div class="card-title"><span class="arrow">▼</span> 0. System Settings</div></div><div id="settings-body" class="card-body"><div class="settings-grid">
    <div class="setting-item"><div class="setting-label">Operating Days / Month</div><input type="number" id="sys-days" class="setting-inp" value="22" min="1" max="31"></div>
    <div class="setting-item"><div class="setting-label">Update Rates (CSV)</div><input type="file" id="csvFile" accept=".csv" class="setting-inp" style="width:200px;padding:4px;" onchange="loadCSV()"></div>
  </div></div></div>

  <div class="card" onclick="activateCard(this,event)"><div class="card-header" onclick="toggle('input-body',this)"><div class="card-title"><span class="arrow rotate">▼</span> 1. Input Data</div></div><div id="input-body" class="card-body open"><div class="grid-wrapper"><table class="data-grid" id="mainGrid"><thead><tr><th style="width:50px;background:white;border:none;"></th><th class="col-name" style="z-index:30;">Product Name</th><th class="col-line" style="z-index:30;">Line</th><th class="col-rate" style="z-index:30;">Rate</th><th>Jan</th><th>Feb</th><th>Mar</th><th>Apr</th><th>May</th><th>Jun</th><th>Jul</th><th>Aug</th><th>Sep</th><th>Oct</th><th>Nov</th><th>Dec</th></tr></thead><tbody id="gridBody"></tbody></table></div><div class="input-footer"><div class="footer-actions"><div style="display:flex;gap:10px;"><button class="btn btn-add" onclick="addProduct()">+ Add Product</button><button class="btn btn-random" onclick="fillRandom()">Random Scenario</button><select id="toolSelect" class="custom-select" onchange="handleDataAction(this)"><option value="" disabled selected>Data Tools ▼</option><option value="clear_demand">Clear Monthly Demand</option><option value="clean">Remove Empty Rows</option><option value="reset">Reset to Defaults</option><option value="clear">Delete All Data</option></select></div><button class="btn btn-primary" onclick="runCalc()" id="runBtn">▶ Run Optimization</button></div><div class="input-help"><span><span class="badge-key">Enter</span> confirm, <span class="badge-key">Arrows</span> navigate.</span></div></div></div></div>
  
  <div id="results-area" class="hidden">
    <div class="kpi-row" style="display:grid;">
        <div class="kpi-card" id="card-util"><div class="kpi-label">System Utilization</div><div class="kpi-value" id="kpi-util">-</div></div>
        <div class="kpi-card"><div class="kpi-label">Total Demand (MT)</div><div class="kpi-value" id="kpi-mt">-</div></div>
        <div class="kpi-card"><div class="kpi-label">Total Days Required</div><div class="kpi-value" id="kpi-days">-</div></div>
        <div class="kpi-card"><div class="kpi-label">Capacity (Slots)</div><div class="kpi-value" id="kpi-active">-</div></div>
    </div>

    <div class="card" onclick="activateCard(this,event)"><div class="card-header" onclick="toggle('cal-body',this)"><div class="card-title"><span class="arrow rotate">▼</span> 2. Visual Schedule <span id="cal-title-month" style="margin-left:10px;font-weight:400;color:#64748b">- Jan</span></div></div><div id="cal-body" class="card-body open"><div class="scheduler-container"><div class="cal-controls"><label style="font-weight:600;font-size:14px;">Select Month:</label><select id="cal-month" class="btn-sec" onchange="updateDashboard(this.value)" style="padding:8px;margin-left:10px;"><option value="jan">January</option><option value="feb">February</option><option value="mar">March</option><option value="apr">April</option><option value="may">May</option><option value="jun">June</option><option value="jul">July</option><option value="aug">August</option><option value="sep">September</option><option value="oct">October</option><option value="nov">November</option><option value="dec">December</option></select></div><div id="calendar-wrapper"></div></div></div></div>
    
    <div class="card" onclick="activateCard(this,event)"><div class="card-header" onclick="toggle('charts-body',this)"><div class="card-title"><span class="arrow rotate">▼</span> 3. Capacity Analytics</div></div><div id="charts-body" class="card-body open"><div class="charts-grid"><div class="chart-box"><canvas id="chart-trend"></canvas></div><div class="chart-box" style="overflow:auto;"><h4 style="margin:0 0 10px 0;font-size:14px;color:#475569;">Line Utilization Heatmap (%)</h4><table class="res-table" id="heatmapTable"><thead><tr><th>Line</th><th>Jan</th><th>Feb</th><th>Mar</th><th>Apr</th><th>May</th><th>Jun</th></tr></thead><tbody></tbody></table></div><div class="chart-box"><canvas id="chart-top"></canvas></div><div class="chart-box"><canvas id="chart-share"></canvas></div></div></div></div>
    <div class="card" onclick="activateCard(this,event)"><div class="card-header" onclick="stopProp(event);toggle('sum-body',this)"><div class="card-title"><span class="arrow rotate">▼</span> 4. Capacity Summary</div><div onclick="stopProp(event)" class="filters"><span id="overall-status" class="status-indicator">System Status: --</span><div class="filter-group"><span class="info-icon has-tooltip" data-tooltip="Filter table to show only products running on a specific production line.">i</span><select id="sum-line" class="filter-sel" onchange="renderSummary()"><option value="all">All Lines</option></select></div><div class="filter-group"><span class="info-icon has-tooltip" data-tooltip="Filter to see products that are Over Capacity (>100%) or Safe.">i</span><select id="sum-stat" class="filter-sel" onchange="renderSummary()"><option value="all">All Status</option><option value="high">Over 50% Load</option><option value="safe">Safe (&lt;50%)</option></select></div></div></div><div id="sum-body" class="card-body open"><div style="overflow-x:auto;max-height:400px;"><table class="res-table" id="resTable"><thead><tr><th>Product</th><th>Line</th><th>Annual Demand (MT)</th><th>Daily Rate</th><th>Line-Days Req</th><th>System Util %</th></tr></thead><tbody></tbody></table></div></div></div></div>
  </div>
</div>
<datalist id="productList"></datalist>
"""
HTML_PART_3 = r"""
<script>
const tooltip=document.getElementById('global-tooltip');document.body.addEventListener('mouseover',(e)=>{const t=e.target.closest('[data-tooltip], .has-tooltip');if(t){tooltip.innerText=t.getAttribute('data-tooltip');tooltip.style.opacity='1';tooltip.style.visibility='visible'}});document.body.addEventListener('mousemove',(e)=>{if(tooltip.style.opacity==='1'){const x=e.clientX+15,y=e.clientY<window.innerHeight/2?e.clientY+20:e.clientY-40;tooltip.style.left=x+'px';tooltip.style.top=y+'px'}});document.body.addEventListener('mouseout',(e)=>{if(e.target.closest('[data-tooltip], .has-tooltip')){tooltip.style.opacity='0';tooltip.style.visibility='hidden'}});

let RATES_DB = {
    "OF 2729": 9.78, "TS 443": 10.20, "OF 1443": 11.64, "TS JJU 2176": 12.54,
    "TS 056": 15.12, "TS 2112": 17.94, "TS 305": 21.78, "TS JJU 2083": 22.86,
    "TS 005": 31.32, "TS 655": 32.64, "TS 705": 32.64, "F 10030 A": 45.72,
    "TS 120 LV": 48.96, "TS 3018": 48.96, "T 3050 A": 51.66, "TAS 10030": 52.02,
    "TS 1000": 54.18, "TS BIS 4065": 59.76, "TS 3000": 59.76
};
const DEFAULTS=[{name:"OF 2729",line:"System (All Lines)",rate:9.78,jan:200,feb:0,mar:0,apr:0,may:0,jun:0,jul:0,aug:0,sep:0,oct:0,nov:0,dec:0},{name:"TS 705",line:"DDA",rate:32.64,jan:21.8,feb:0,mar:0,apr:0,may:0,jun:0,jul:0,aug:0,sep:0,oct:0,nov:0,dec:0},{name:"TS 1000",line:"DDA",rate:54.18,jan:18.1,feb:0,mar:0,apr:0,may:0,jun:0,jul:0,aug:0,sep:0,oct:0,nov:0,dec:0}];
const MONTHS=["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
const LINES=["DDA","DDB","DDC","DDD","DDE","DDF","Gen","System (All Lines)"];
let globalData={}; let charts={};

function init(){ 
    updateDatalist();
    resetData(); 
}

function updateDatalist(){
    const list = document.getElementById('productList'); list.innerHTML = '';
    Object.keys(RATES_DB).forEach(p => { const opt = document.createElement('option'); opt.value = p; list.appendChild(opt); });
}

function onProductInput(inp) {
    const val = inp.value;
    if(RATES_DB[val]) { updateRateCell(inp.closest('tr'), RATES_DB[val]); }
}

function updateRateCell(row, systemRate) {
    const line = row.querySelector('.col-line select').value;
    const rate = (line === "System (All Lines)") ? systemRate : (systemRate / 6);
    row.querySelector('.col-rate input').value = rate.toFixed(2);
}

function onLineChange(select) {
    const row = select.closest('tr');
    const prod = row.querySelector('.name-inp').value;
    if(RATES_DB[prod]) { updateRateCell(row, RATES_DB[prod]); }
}

function loadCSV() {
    const file = document.getElementById('csvFile').files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        const text = e.target.result;
        const rows = text.split('\n');
        RATES_DB = {};
        rows.forEach((r, i) => {
            if(i < 2) return; 
            const c = r.split(',');
            if(c.length > 1) {
                const name = c[0].trim();
                let rate = 0;
                for(let j=1; j<=6; j++) {
                    const v = parseFloat(c[j]);
                    if(!isNaN(v)) rate += v;
                }
                if(name && rate > 0) RATES_DB[name] = rate.toFixed(2);
            }
        });
        updateDatalist();
        document.getElementById('gridBody').innerHTML = '';
        let limit = 0;
        Object.keys(RATES_DB).forEach(p => {
            if(limit < 20) { renderRow({name: p, line: "System (All Lines)", rate: RATES_DB[p]}); limit++; }
        });
        alert("Database Updated from CSV!");
    };
    reader.readAsText(file);
}

function handleDataAction(s){const v=s.value;if(v==='clear_demand')document.querySelectorAll('.m-inp').forEach(i=>i.value='');else if(v==='clean')cleanRows();else if(v==='reset')resetData();else if(v==='clear')clearGrid();s.value=""}
function fillRandom(){document.querySelectorAll('.m-inp').forEach(i=>{if(!i.value)i.value=(Math.random()*400).toFixed(1)})}
function resetData(){document.getElementById('gridBody').innerHTML='';DEFAULTS.forEach(d=>renderRow(d))}
function clearGrid(){document.getElementById('gridBody').innerHTML='';addProduct()}
function cleanRows(){const r=document.querySelectorAll('#gridBody tr');r.forEach(tr=>{let s=0;tr.querySelectorAll('.m-inp').forEach(i=>s+=parseFloat(i.value)||0);if(s===0)tr.remove()});if(document.querySelectorAll('#gridBody tr').length===0)addProduct()}
function addProduct(){renderRow({name:"",line:"DDA",rate:0},true);document.querySelector('.grid-wrapper').scrollTop=0}

function renderRow(d,top){
    const tr=document.createElement('tr');
    let opts=LINES.map(l=>`<option value="${l}" ${d.line===l?'selected':''}>${l}</option>`).join('');
    tr.innerHTML=`<td><div class="del-wrapper"><div class="del-row" onclick="this.closest('tr').remove()">×</div></div></td><td class="col-name"><input type="text" class="cell-inp name-inp nav-item" value="${d.name||''}" list="productList" onclick="this.select()" oninput="onProductInput(this)"></td><td class="col-line"><select class="cell-select nav-item" onchange="onLineChange(this)">${opts}</select></td><td class="col-rate"><input type="number" step="0.01" class="cell-inp nav-item" value="${d.rate||0}" onclick="this.select()"></td>${MONTHS.map(m=>`<td class="month-cell"><input type="number" class="cell-inp m-inp nav-item" data-m="${m}" value="${d[m]!==undefined?d[m]:''}" onclick="this.select()"></td>`).join('')}`;
    const b=document.getElementById('gridBody');top?b.prepend(tr):b.appendChild(tr);bindKeys();
}

function bindKeys(){const i=document.querySelectorAll('.nav-item');i.forEach((el,x)=>{el.onkeydown=e=>{const c=14,curr=Array.from(i).indexOf(e.target);if(e.key==='ArrowRight'){e.preventDefault();if(i[curr+1])i[curr+1].focus()}else if(e.key==='ArrowLeft'){e.preventDefault();if(i[curr-1])i[curr-1].focus()}else if(e.key==='ArrowDown'){e.preventDefault();if(i[curr+c])i[curr+c].focus()}else if(e.key==='ArrowUp'){e.preventDefault();if(i[curr-c])i[curr-c].focus()}}})}
document.addEventListener('click',e=>{if(!e.target.closest('.card'))document.querySelectorAll('.card').forEach(c=>c.classList.remove('active-card'))});
function activateCard(el,e){if(e)e.stopPropagation();document.querySelectorAll('.card').forEach(c=>c.classList.remove('active-card'));el.classList.add('active-card')}
function toggle(id,h){document.getElementById(id).classList.toggle('open');h.querySelector('.arrow').classList.toggle('rotate')}
function stopProp(e){e.stopPropagation()}
async function runCalc(){const btn=document.getElementById('runBtn');btn.innerText="Calculating...";btn.disabled=true;const op=document.getElementById('sys-days').value||22;const rows=[];document.querySelectorAll('#gridBody tr').forEach(tr=>{rows.push({product:tr.querySelector('.name-inp').value,line:tr.querySelector('.col-line select').value,rate:tr.querySelector('.col-rate input').value,jan:tr.querySelector('[data-m="jan"]').value,feb:tr.querySelector('[data-m="feb"]').value,mar:tr.querySelector('[data-m="mar"]').value,apr:tr.querySelector('[data-m="apr"]').value,may:tr.querySelector('[data-m="may"]').value,jun:tr.querySelector('[data-m="jun"]').value,jul:tr.querySelector('[data-m="jul"]').value,aug:tr.querySelector('[data-m="aug"]').value,sep:tr.querySelector('[data-m="sep"]').value,oct:tr.querySelector('[data-m="oct"]').value,nov:tr.querySelector('[data-m="nov"]').value,dec:tr.querySelector('[data-m="dec"]').value})});try{const res=await fetch('/optimize',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows,op_days:op})});const data=await res.json();globalData=data;updateUI(data);document.getElementById('results-area').classList.remove('hidden');document.getElementById('results-area').scrollIntoView({behavior:'smooth'});document.getElementById('pdfBtn').disabled=false;updateFilters()}catch(e){alert(e)}btn.innerText="▶ Run Optimization";btn.disabled=false}

// NEW: Dynamic Dashboard Update based on Selected Month
function updateDashboard(month) {
    if (!globalData.monthly_data) return;
    const mData = globalData.monthly_data[month];
    
    // Update KPIs
    const util = document.getElementById('kpi-util');
    util.innerText = mData.util.toFixed(1) + "%";
    
    // COLOR LOGIC FOR KPI CARD
    const card = document.getElementById('card-util');
    if(mData.util === 0) card.className = 'kpi-card';
    else card.className = 'kpi-card ' + (mData.util > 100 ? 'kpi-crit' : (mData.util > 85 ? 'kpi-warn' : 'kpi-safe'));

    document.getElementById('kpi-mt').innerText = Math.round(mData.mt).toLocaleString() + " MT";
    document.getElementById('kpi-days').innerText = mData.days.toFixed(1);
    document.getElementById('kpi-active').innerText = Math.round(mData.capacity) + " Slots";
    
    renderCalendar(month);
    renderCharts(globalData); 
    renderHeatmap(globalData.line_heatmap);
    renderSummary();
}

function updateUI(d){updateDashboard('jan');} // Init with Jan

function renderCalendar(m){if(!m)m=document.getElementById('cal-month').value;const c=document.getElementById('calendar-wrapper'),d=globalData.calendar_data||{},l=Object.keys(d).sort(),lim=globalData.op_days||22;if(l.length===0){c.innerHTML='<div style="padding:20px;text-align:center;color:#64748b">No production.</div>';return}let h='<table class="cal-table"><thead><tr><th class="cal-line-header">Line</th>';for(let i=1;i<=lim;i++)h+=`<th>${i}</th>`;h+='</tr></thead><tbody>';l.forEach(ln=>{const p=d[ln][m]||[];let td=0;p.forEach(pr=>td+=pr.days);const isOver=td>lim;const bc=isOver?'day-cell filled-red has-tooltip':'day-cell filled has-tooltip';h+=`<tr><td class="cal-line-header">${ln}</td>`;let dc=0;p.forEach(pr=>{const dy=Math.ceil(pr.days);for(let i=0;i<dy;i++){if(dc<lim){h+=`<td class="${bc}" data-tooltip="Product: ${pr.product}\nDays: ${pr.days.toFixed(1)}"></td>`;dc++}}});for(let i=dc;i<lim;i++)h+=`<td class="day-cell"></td>`;h+='</tr>'});h+='</tbody></table>';c.innerHTML=h}
function renderHeatmap(d){const tb=document.querySelector('#heatmapTable tbody'),hd=document.querySelector('#heatmapTable thead tr');hd.innerHTML='<th>Line</th>'+MONTHS.slice(0,6).map(m=>`<th>${m.toUpperCase()}</th>`).join('');tb.innerHTML=d.map(r=>{let c=`<td style="font-weight:700">${r.line}</td>`;MONTHS.slice(0,6).forEach(m=>{const o=r[m]||{pct:0,mt:0,days:0};let cl='hm-cell has-tooltip';if(o.pct>100)cl+=' hm-red';else if(o.pct>85)cl+=' hm-orange';else if(o.pct>0)cl+=' hm-green';c+=`<td class="${cl}" data-tooltip="Month: ${m.toUpperCase()}\nLoad: ${o.pct.toFixed(0)}%\nDemand: ${o.mt.toFixed(0)} MT">${o.pct.toFixed(0)}%</td>`});return`<tr>${c}</tr>`}).join('')}
function updateFilters(){const l=[...new Set(globalData.capacity_view.map(i=>i.LINE))].sort(),s=document.getElementById('sum-line');s.innerHTML='<option value="all">All Lines</option>';l.forEach(n=>s.add(new Option(n,n)))}
function renderSummary(){const lv=document.getElementById('sum-line').value,sv=document.getElementById('sum-stat').value,f=globalData.capacity_view.filter(r=>r.Demand_MT>0);let final=f;if(lv!=='all')final=f.filter(r=>r.LINE===lv);if(sv==='high')final=f.filter(r=>r.System_Time_Utilization_Percent>=50);if(sv==='safe')final=f.filter(r=>r.System_Time_Utilization_Percent<50);const b=document.getElementById('overall-status'),cr=final.some(r=>r.System_Time_Utilization_Percent>100),wr=final.some(r=>r.System_Time_Utilization_Percent>85);if(cr){b.className='status-indicator stat-crit';b.innerText="System Status: CRITICAL"}else if(wr){b.className='status-indicator stat-warn';b.innerText="System Status: Warning"}else{b.className='status-indicator stat-ok';b.innerText="System Status: Safe"}const tb=document.querySelector('#resTable tbody');if(final.length===0){tb.innerHTML='<tr><td colspan="6" style="text-align:center;padding:20px;">No records.</td></tr>';return}tb.innerHTML=final.map(r=>{const u=r.System_Time_Utilization_Percent;let s='';if(u>100)s=`<span class="status-pill pill-red has-tooltip" data-tooltip="CRITICAL: ${u.toFixed(0)}%">${u.toFixed(0)}% ⚠️</span>`;else if(u>85)s=`<span class="status-pill pill-orange has-tooltip" data-tooltip="High Load">${u.toFixed(0)}%</span>`;else s=`<span class="status-pill pill-green">${u.toFixed(0)}%</span>`;return`<tr><td style="font-weight:600">${r.PRODUCT}</td><td>${r.LINE}</td><td>${r.Demand_MT.toFixed(1)}</td><td>${r.Daily_Total_Capacity_MT.toFixed(2)}</td><td>${r.Days_Required.toFixed(1)}</td><td>${s}</td></tr>`}).join('')}
function renderCharts(d){const ops={responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true}}};if(charts['trend'])charts['trend'].destroy();charts['trend']=new Chart(document.getElementById('chart-trend'),{type:'line',data:{labels:MONTHS.map(m=>m.toUpperCase()),datasets:[{data:d.monthly_util_trend,borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,0.1)',fill:true}]},options:ops});const s=[...d.capacity_view.filter(r=>r.Demand_MT>0)].sort((a,b)=>b.System_Time_Utilization_Percent-a.System_Time_Utilization_Percent).slice(0,6);if(charts['top'])charts['top'].destroy();charts['top']=new Chart(document.getElementById('chart-top'),{type:'bar',indexAxis:'y',data:{labels:s.map(r=>r.PRODUCT),datasets:[{data:s.map(r=>r.System_Time_Utilization_Percent),backgroundColor:'#3b82f6'}]},options:ops});if(charts['share'])charts['share'].destroy();charts['share']=new Chart(document.getElementById('chart-share'),{type:'doughnut',data:{labels:s.map(r=>r.PRODUCT),datasets:[{data:s.map(r=>r.Demand_MT),backgroundColor:['#3b82f6','#60a5fa','#93c5fd','#bfdbfe','#dbeafe','#eff6ff']}]},options:ops})}
function downloadPDF(){
    if (!globalData || !globalData.capacity_view) { alert("Run Optimization first!"); return; }
    const { jsPDF } = window.jspdf; const doc = new jsPDF();
    doc.setFontSize(22); doc.setTextColor(37, 99, 235); doc.text("Capacity Planning Report", 14, 20);
    doc.setFontSize(10); doc.setTextColor(100); doc.text(`Generated: ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`, 14, 28);
    doc.setFillColor(241, 245, 249); doc.rect(14, 35, 180, 30, 'F'); doc.setFontSize(12); doc.setTextColor(0);
    doc.text(`System Utilization: ${document.getElementById('kpi-util').innerText}`, 20, 48);
    doc.text(`Total Demand: ${document.getElementById('kpi-mt').innerText}`, 20, 58);
    doc.text(`Operating Days: ${globalData.op_days || 22}`, 110, 48);
    doc.text(`Active Slots: ${document.getElementById('kpi-active').innerText}`, 110, 58);
    const sumRows = globalData.capacity_view.filter(r=>r.Demand_MT>0).map(r => [
      r.PRODUCT, r.LINE, r.Demand_MT.toFixed(1), r.Daily_Total_Capacity_MT.toFixed(2), 
      r.Days_Required.toFixed(1), r.System_Time_Utilization_Percent.toFixed(1) + "%"
    ]);
    doc.autoTable({ startY: 85, head: [['Product', 'Line', 'MT', 'Rate', 'Days', 'Util %']], body: sumRows, theme: 'grid', headStyles: { fillColor: [37, 99, 235] } });
    let finalY = doc.lastAutoTable.finalY + 20;
    doc.text("2. Line Utilization (First 6 Months)", 14, finalY);
    const heatRows = globalData.line_heatmap.map(r => [
       r.line, r.jan.pct.toFixed(0)+"%", r.feb.pct.toFixed(0)+"%", r.mar.pct.toFixed(0)+"%", r.apr.pct.toFixed(0)+"%", 
       r.may.pct.toFixed(0)+"%", r.jun.pct.toFixed(0)+"%"
    ]);
    doc.autoTable({ startY: finalY + 5, head: [['Line', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']], body: heatRows, theme: 'grid', headStyles: { fillColor: [59, 130, 246] } });
    finalY = doc.lastAutoTable.finalY + 30; doc.setLineWidth(0.5); doc.setDrawColor(200); doc.line(14, finalY, 80, finalY); doc.text("Approved By", 14, finalY + 5); doc.line(120, finalY, 180, finalY); doc.text("Date", 120, finalY + 5);
    doc.save("Capacity_Report_v10.pdf");
}
init();
</script>
</body>
</html>
"""
# --- 4. PYTHON LOGIC ---
@app.route('/')
def index():
    return render_template_string(HTML_PART_1 + HTML_PART_2 + HTML_PART_3)

@app.route('/template')
def template():
    csv_content = "Product Name,Line DDA,Line DDB,Line DDC,Line DDD,Line DDE,Line DDF\nExample Product,1.5,1.5,1.5,1.5,1.5,1.5\n"
    return Response(csv_content, mimetype="text/csv", headers={"Content-disposition": "attachment; filename=Capacity_Template.csv"})

@app.route('/optimize', methods=['POST'])
def optimize():
    try:
        data = request.get_json() if request.is_json else {}
        rows = data.get('rows', [])
        try: 
            op_days = float(data.get('op_days', 22))
        except: 
            op_days = 22
            
        LINES_COUNT = 6
        months_list = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"]
        
        # PREPARE MONTHLY DATA STRUCTURE
        monthly_data = {m: {"mt": 0, "days": 0, "util": 0, "capacity": op_days * LINES_COUNT} for m in months_list}
        
        capacity_view = []; line_tracker = {}; calendar_data = {}; lines_set = set();
        monthly_demand_mt = {m: 0 for m in months_list} 

        for r in rows:
            l = r.get('line', 'General').strip()
            if l: lines_set.add(l)
        
        for l in ["DDA", "DDB", "DDC", "DDD", "DDE", "DDF"]:
            if l not in line_tracker: line_tracker[l] = {m: {"days":0, "mt":0} for m in months_list}

        for r in rows:
            prod = r.get('product', 'Unknown'); line = r.get('line', 'General').strip(); 
            try: daily_cap = float(r.get('rate', 0))
            except: daily_cap = 0
            
            if daily_cap > 0: single_line_rate = daily_cap / LINES_COUNT
            else: single_line_rate = 1 

            prod_demand = 0; prod_days = 0
            
            for m in months_list:
                val = float(r.get(m, 0) or 0)
                if val > 0:
                    prod_demand += val
                    duration = val / daily_cap if daily_cap > 0 else 0
                    slots_used = duration * (6 if line == "System (All Lines)" else 1)
                    
                    prod_days += duration # Keep track of raw duration
                    
                    # Update Monthly Aggregates
                    monthly_data[m]["mt"] += val
                    monthly_data[m]["days"] += slots_used
                    
                    if line == "System (All Lines)":
                         for l_key in ["DDA", "DDB", "DDC", "DDD", "DDE", "DDF"]:
                             line_tracker[l_key][m]["days"] += duration
                             line_tracker[l_key][m]["mt"] += val / 6
                    elif line in line_tracker:
                        line_tracker[line][m]["days"] += duration
                        line_tracker[line][m]["mt"] += val

                    if duration > 0:
                        if line == "System (All Lines)":
                            for l_key in ["DDA", "DDB", "DDC", "DDD", "DDE", "DDF"]:
                                 if l_key not in calendar_data: calendar_data[l_key] = {}
                                 if m not in calendar_data[l_key]: calendar_data[l_key][m] = []
                                 calendar_data[l_key][m].append({ "product": prod, "days": duration, "mt": val })
                        else:
                            if line not in calendar_data: calendar_data[line] = {}
                            if m not in calendar_data[line]: calendar_data[line][m] = []
                            calendar_data[line][m].append({ "product": prod, "days": duration, "mt": val })
            
            # Calc Annual Utilization for Table
            active_months_count = 0
            for m in months_list:
                if float(r.get(m, 0) or 0) > 0: active_months_count += 1
            
            if active_months_count > 0:
                 total_avail_slots = op_days * LINES_COUNT * active_months_count
                 # Calc Total Slots consumed by this product across all active months
                 total_slots_consumed = 0
                 for m in months_list:
                     if float(r.get(m,0)) > 0:
                         dur = float(r.get(m,0)) / daily_cap
                         total_slots_consumed += dur * (6 if line == "System (All Lines)" else 1)
                         
                 prod_util = (total_slots_consumed / total_avail_slots) * 100
            else:
                 prod_util = 0

            capacity_view.append({
                "PRODUCT": prod, "LINE": line, "Demand_MT": prod_demand, "Daily_Total_Capacity_MT": daily_cap,
                "Days_Required": prod_days, "System_Time_Utilization_Percent": prod_util
            })

        # Calculate Monthly Utilizations
        monthly_util_trend = []
        for m in months_list:
            capacity = op_days * LINES_COUNT
            util = (monthly_data[m]["days"] / capacity) * 100
            monthly_data[m]["util"] = util
            monthly_util_trend.append(util)

        line_heatmap = []
        for line, months_data in line_tracker.items():
            row_obj = {"line": line}
            for m, info in months_data.items(): 
                pct = (info["days"] / op_days) * 100
                row_obj[m] = { "pct": pct, "days": info["days"], "mt": info["mt"] }
            line_heatmap.append(row_obj)
            
        return jsonify({
            "monthly_data": monthly_data, # Send full breakdown for frontend filtering
            "capacity_view": capacity_view,
            "line_heatmap": line_heatmap,
            "calendar_data": calendar_data,
            "monthly_util_trend": monthly_util_trend,
            "op_days": op_days
        })

    except Exception as e: return jsonify({"error": str(e)}), 500

def open_browser(): webbrowser.open_new("http://127.0.0.1:5000")
if __name__ == '__main__': Timer(1, open_browser).start(); app.run(port=5000)
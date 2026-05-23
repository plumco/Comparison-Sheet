<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Huliot MLCP Price Advantage Tool</title>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700;900&family=DM+Mono:wght@300;400;500&family=Outfit:wght@300;400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0a0d14;
  --surface:#10141f;
  --surface2:#161c2d;
  --border:#1e2840;
  --huliot:#00c6a2;
  --huliot-dim:#00c6a230;
  --gold:#f0b429;
  --text:#e8eaf0;
  --muted:#7a82a0;
  --danger:#ff5a5a;
  --danger-dim:#ff5a5a22;
  --green:#00c6a2;
  --green-dim:#00c6a222;
  --tab-active:#1a2240;
}
body{background:var(--bg);color:var(--text);font-family:'Outfit',sans-serif;min-height:100vh}

/* HEADER */
.header{background:linear-gradient(135deg,#0d1b35 0%,#0a0d14 60%);border-bottom:1px solid var(--border);padding:28px 40px 20px;position:relative;overflow:hidden}
.header::before{content:'';position:absolute;top:-40px;right:-60px;width:300px;height:300px;background:radial-gradient(circle,#00c6a215 0%,transparent 70%);pointer-events:none}
.logo-row{display:flex;align-items:center;gap:16px;margin-bottom:12px}
.logo-badge{background:var(--huliot);color:#000;font-family:'Playfair Display',serif;font-weight:900;font-size:18px;padding:6px 16px;letter-spacing:1px}
.vs-badge{color:var(--muted);font-size:13px;letter-spacing:3px;text-transform:uppercase}
.competitors-badge{display:flex;gap:8px;flex-wrap:wrap}
.comp-tag{background:var(--surface2);border:1px solid var(--border);color:var(--muted);font-size:11px;padding:3px 10px;border-radius:2px;letter-spacing:1px;font-family:'DM Mono',monospace;text-transform:uppercase}
.header-title{font-family:'Playfair Display',serif;font-size:28px;font-weight:700;color:#fff;line-height:1.1;margin-bottom:6px}
.header-sub{color:var(--muted);font-size:13px;font-family:'DM Mono',monospace;letter-spacing:0.5px}

/* SUMMARY CARDS */
.summary-strip{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:1px;background:var(--border);margin:0;border-bottom:1px solid var(--border)}
.stat-card{background:var(--surface);padding:16px 20px;text-align:center}
.stat-num{font-family:'Playfair Display',serif;font-size:28px;font-weight:700;color:var(--huliot);line-height:1}
.stat-num.danger{color:var(--danger)}
.stat-label{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:1.5px;margin-top:4px;font-family:'DM Mono',monospace}

/* CONTROLS */
.controls{padding:20px 40px;display:flex;gap:12px;flex-wrap:wrap;align-items:center;background:var(--surface);border-bottom:1px solid var(--border)}
.search-wrap{position:relative;flex:1;min-width:220px;max-width:360px}
.search-wrap input{width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:9px 12px 9px 36px;font-family:'Outfit',sans-serif;font-size:13px;outline:none;transition:border-color .2s}
.search-wrap input:focus{border-color:var(--huliot)}
.search-wrap input::placeholder{color:var(--muted)}
.search-icon{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--muted);font-size:14px}
select{background:var(--bg);border:1px solid var(--border);color:var(--text);padding:9px 12px;font-family:'Outfit',sans-serif;font-size:13px;outline:none;cursor:pointer}
select:focus{border-color:var(--huliot)}
.filter-label{font-size:11px;color:var(--muted);letter-spacing:1px;text-transform:uppercase;font-family:'DM Mono',monospace;align-self:center}

/* TABS */
.tabs{display:flex;border-bottom:1px solid var(--border);background:var(--surface);padding:0 40px}
.tab{padding:14px 24px;cursor:pointer;font-size:13px;font-weight:500;color:var(--muted);border-bottom:2px solid transparent;transition:all .2s;letter-spacing:0.3px;white-space:nowrap}
.tab:hover{color:var(--text)}
.tab.active{color:var(--huliot);border-bottom-color:var(--huliot);background:var(--tab-active)}
.tab .tab-count{background:var(--border);color:var(--muted);font-size:10px;padding:1px 6px;border-radius:10px;margin-left:6px;font-family:'DM Mono',monospace}
.tab.active .tab-count{background:var(--huliot-dim);color:var(--huliot)}

/* TABLE */
.table-wrap{overflow-x:auto;padding:0 40px 40px}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:20px}
thead th{background:var(--surface2);color:var(--muted);font-family:'DM Mono',monospace;font-size:10px;letter-spacing:1.5px;text-transform:uppercase;padding:10px 12px;text-align:left;border-bottom:1px solid var(--border);cursor:pointer;user-select:none;white-space:nowrap}
thead th:hover{color:var(--text)}
thead th.sort-asc::after{content:' ↑';color:var(--huliot)}
thead th.sort-desc::after{content:' ↓';color:var(--huliot)}
tbody tr{border-bottom:1px solid #12172500;transition:background .15s}
tbody tr:hover{background:var(--surface2)}
tbody tr:nth-child(even){background:var(--surface)05}
td{padding:10px 12px;vertical-align:middle}
.td-name{font-weight:500;color:var(--text);max-width:320px;line-height:1.3;font-size:12.5px}
.td-name .cat-pill{display:inline-block;background:var(--surface2);color:var(--muted);font-size:9px;padding:1px 6px;border-radius:2px;margin-left:6px;font-family:'DM Mono',monospace;text-transform:uppercase;letter-spacing:0.5px;vertical-align:middle}
.td-pack{color:var(--muted);font-size:11px;font-family:'DM Mono',monospace;white-space:nowrap}
.td-price{font-family:'DM Mono',monospace;font-size:13px;white-space:nowrap}
.td-price.huliot{color:var(--huliot);font-weight:500}
.td-price.competitor{color:var(--text)}
.td-price.no-data{color:var(--muted);font-style:italic;font-size:11px}
.diff-cell{white-space:nowrap}
.diff-badge{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:2px;font-family:'DM Mono',monospace;font-size:11px;font-weight:500}
.diff-badge.cheaper{background:var(--green-dim);color:var(--green)}
.diff-badge.expensive{background:var(--danger-dim);color:var(--danger)}
.diff-bar-wrap{display:flex;align-items:center;gap:6px;margin-top:3px}
.diff-bar{height:3px;border-radius:2px;min-width:2px;max-width:100px;transition:width .3s}
.diff-bar.cheaper{background:var(--green)}
.diff-bar.expensive{background:var(--danger)}
.bar-pct{font-size:10px;font-family:'DM Mono',monospace}
.bar-pct.cheaper{color:var(--green)}
.bar-pct.expensive{color:var(--danger)}

/* NO DATA state */
.empty-state{text-align:center;padding:60px;color:var(--muted)}
.empty-state .big-icon{font-size:48px;margin-bottom:16px;opacity:0.3}

/* CHART SECTION */
.chart-section{padding:20px 40px;border-top:1px solid var(--border);background:var(--surface)}
.chart-title{font-family:'Playfair Display',serif;font-size:14px;color:var(--muted);margin-bottom:16px;letter-spacing:0.5px}
.chart-container{display:flex;align-items:flex-end;gap:8px;height:120px;overflow-x:auto;padding-bottom:8px}
.chart-bar-group{display:flex;flex-direction:column;align-items:center;gap:4px;min-width:60px}
.chart-bar-outer{display:flex;align-items:flex-end;height:90px;gap:4px}
.chart-bar{width:24px;border-radius:2px 2px 0 0;transition:height .5s ease;position:relative;cursor:pointer}
.chart-bar:hover::after{content:attr(data-val);position:absolute;top:-22px;left:50%;transform:translateX(-50%);background:var(--surface2);border:1px solid var(--border);color:var(--text);font-size:9px;padding:2px 5px;white-space:nowrap;font-family:'DM Mono',monospace}
.chart-bar-label{font-size:8px;color:var(--muted);text-align:center;font-family:'DM Mono',monospace;text-transform:uppercase;letter-spacing:0.5px}

/* FOOTER */
.footer{padding:16px 40px;border-top:1px solid var(--border);color:var(--muted);font-size:11px;font-family:'DM Mono',monospace;display:flex;justify-content:space-between;background:var(--surface)}

/* PRINT */
@media print{
  .controls,.tabs,.footer,.chart-section{display:none}
  .table-wrap{padding:0}
  body{background:#fff;color:#000}
  .header{background:#fff;border-bottom:2px solid #000}
  .header-title,.stat-num,.td-price.huliot{color:#000}
  tbody tr:hover{background:none}
}

/* Row highlight on click */
tr.selected-row{background:var(--tab-active) !important;outline:1px solid var(--huliot);outline-offset:-1px}
</style>
</head>
<body>

<div class="header">
  <div class="logo-row">
    <div class="logo-badge">HULIOT</div>
    <div class="vs-badge">vs</div>
    <div class="competitors-badge">
      <span class="comp-tag">Viega</span>
      <span class="comp-tag">Kitec</span>
      <span class="comp-tag">Geberit Mepla</span>
      <span class="comp-tag">Kantherm</span>
    </div>
  </div>
  <div class="header-title">MLCP Price Advantage Comparison</div>
  <div class="header-sub">Multilayer Composite Pipe System &mdash; April 2025 Pricing &mdash; INR</div>
</div>

<div class="summary-strip" id="summaryStrip"></div>

<div class="controls">
  <span class="filter-label">Filter</span>
  <div class="search-wrap">
    <span class="search-icon">&#9906;</span>
    <input type="text" id="searchInput" placeholder="Search by product name or size…">
  </div>
  <select id="categoryFilter">
    <option value="">All Categories</option>
    <option>Multilayer Pipe</option>
    <option>Adapters</option>
    <option>Elbows</option>
    <option>Tees</option>
    <option>Reducers</option>
    <option>End Caps</option>
    <option>Manifolds</option>
    <option>Valves</option>
    <option>Other Fittings</option>
  </select>
  <select id="savingsFilter">
    <option value="">All Items</option>
    <option value="cheaper">Huliot Cheaper</option>
    <option value="expensive">Huliot Higher</option>
    <option value="big">Savings &gt; 30%</option>
  </select>
  <span style="margin-left:auto;color:var(--muted);font-size:11px;font-family:'DM Mono',monospace" id="rowCount"></span>
  <button onclick="window.print()" style="background:var(--surface2);border:1px solid var(--border);color:var(--muted);padding:8px 14px;cursor:pointer;font-size:11px;font-family:'DM Mono',monospace;text-transform:uppercase;letter-spacing:1px">&#128438; Print</button>
  <button onclick="downloadCSV()" style="background:var(--huliot-dim);border:1px solid var(--huliot);color:var(--huliot);padding:8px 14px;cursor:pointer;font-size:11px;font-family:'DM Mono',monospace;text-transform:uppercase;letter-spacing:1px">&#8681; CSV</button>
  <button onclick="downloadExcel()" style="background:#1a2e1a;border:1px solid #2e7d32;color:#4caf50;padding:8px 14px;cursor:pointer;font-size:11px;font-family:'DM Mono',monospace;text-transform:uppercase;letter-spacing:1px">&#8681; Excel</button>
  <button onclick="downloadAllCSV()" style="background:var(--surface2);border:1px solid #4a8cff;color:#4a8cff;padding:8px 14px;cursor:pointer;font-size:11px;font-family:'DM Mono',monospace;text-transform:uppercase;letter-spacing:1px">&#8681; Full Sheet</button>
</div>

<div class="tabs" id="tabsBar"></div>

<div class="chart-section" id="chartSection" style="display:none">
  <div class="chart-title">Selected Product — Price Comparison Across Competitors</div>
  <div class="chart-container" id="chartContainer"></div>
</div>

<div class="table-wrap">
  <table id="mainTable">
    <thead id="tableHead"></thead>
    <tbody id="tableBody"></tbody>
  </table>
  <div id="emptyState" class="empty-state" style="display:none">
    <div class="big-icon">&#128269;</div>
    <div>No products match your filter</div>
  </div>
</div>

<div class="footer">
  <span>Huliot India &mdash; MLCP Price Advantage Tool &mdash; Confidential</span>
  <span id="footerDate"></span>
</div>

<script>
const PRODUCTS = [{"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 16x2.0", "category": "Multilayer Pipe", "pack": "100 Mtr Coil", "unit": "INR/MTR", "huliot": 170, "Viega": 271, "Kitec": 123, "Geberit Mepla": 540, "Kantherm": 225}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 20x2.0", "category": "Multilayer Pipe", "pack": "100 Mtr Coil", "unit": "INR/MTR", "huliot": 223, "Viega": 326, "Kitec": 155, "Geberit Mepla": 810, "Kantherm": 275}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA  25x2.5", "category": "Multilayer Pipe", "pack": "50 Mtr Coil", "unit": "INR/MTR", "huliot": 317, "Viega": 488, "Kitec": 203, "Geberit Mepla": 1230, "Kantherm": 412}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 32x3.0", "category": "Multilayer Pipe", "pack": "50 Mtr Coil", "unit": "INR/MTR", "huliot": 531, "Viega": 1070, "Kitec": 322, "Geberit Mepla": 2170, "Kantherm": 890}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 40x4.0", "category": "Multilayer Pipe", "pack": "5 Mtr Bar", "unit": "INR/MTR", "huliot": 820, "Viega": 3200, "Kitec": 426, "Geberit Mepla": 2650, "Kantherm": 2990}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 50x4.5", "category": "Multilayer Pipe", "pack": "5 Mtr Bar", "unit": "INR/MTR", "huliot": 1289, "Viega": 4350, "Kitec": 648, "Geberit Mepla": 3700, "Kantherm": 3930}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 63x6.0", "category": "Multilayer Pipe", "pack": "5 Mtr Bar", "unit": "INR/MTR", "huliot": 1905, "Viega": 6910, "Kitec": 882, "Geberit Mepla": 7140, "Kantherm": 5985}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 75x7.5", "category": "Multilayer Pipe", "pack": "5 Mtr Bar", "unit": "INR/MTR", "huliot": 6380, "Viega": null, "Kitec": 1221, "Geberit Mepla": null, "Kantherm": null}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 90x8.5", "category": "Multilayer Pipe", "pack": "5 Mtr Bar", "unit": "INR/MTR", "huliot": 18550, "Viega": null, "Kitec": 1644, "Geberit Mepla": null, "Kantherm": null}, {"name": "MULTILAYER PIPE PE-RT/AL/PERT HELIKLIMA 110x10.0", "category": "Multilayer Pipe", "pack": "5 Mtr Bar", "unit": "INR/MTR", "huliot": 21465, "Viega": null, "Kitec": 2196, "Geberit Mepla": null, "Kantherm": null}, {"name": "PERT-AL-PERT PIPE ISL 6 BL 16x2.0 - 50 MT (COIL)", "category": "Multilayer Pipe", "pack": "50 Mtr", "unit": "INR/MTR", "huliot": 233, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "PERT-AL-PERT PIPE ISL 6 BL 20x2.0 - 50 MT (COIL)", "category": "Multilayer Pipe", "pack": "50 Mtr", "unit": "INR/MTR", "huliot": 312, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "PERT-AL-PERT PIPE ISL 10 BL 25x2.5 - 25 MT (COIL)", "category": "Multilayer Pipe", "pack": "25 Mtr", "unit": "INR/MTR", "huliot": 461, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "PERT-AL-PERT PIPE ISL 10 BL 32x3.0 - 25 MT (COIL)", "category": "Multilayer Pipe", "pack": "25 Mtr", "unit": "INR/MTR", "huliot": 737, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "PERT EVOH PIPE 16x1.8 - 240 MT (COIL)", "category": "Multilayer Pipe", "pack": "240 Mtr", "unit": "INR/MTR", "huliot": 107, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "PERT EVOH PIPE 16x2.0 - 240 MT (COIL)", "category": "Multilayer Pipe", "pack": "240 Mtr", "unit": "INR/MTR", "huliot": 115, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "PERT EVOH PIPE 20x1.9 - 240 MT (COIL)", "category": "Multilayer Pipe", "pack": "240 Mtr", "unit": "INR/MTR", "huliot": 153, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 16x1/2 - KLIMAPRESS", "category": "Adapters", "pack": "40", "unit": "INR/UNIT", "huliot": 406, "Viega": 400, "Kitec": 137, "Geberit Mepla": 770, "Kantherm": 340}, {"name": "FEMALE THREADED ADAPTER 20x1/2 - KLIMAPRESS", "category": "Adapters", "pack": "35", "unit": "INR/UNIT", "huliot": 467, "Viega": 510, "Kitec": 155, "Geberit Mepla": 770, "Kantherm": 440}, {"name": "FEMALE THREADED ADAPTER 20x3/4 - KLIMAPRESS", "category": "Adapters", "pack": "30", "unit": "INR/UNIT", "huliot": 552, "Viega": 640, "Kitec": 284, "Geberit Mepla": null, "Kantherm": 580}, {"name": "FEMALE THREADED ADAPTER 25x1/2 - KLIMAPRESS", "category": "Adapters", "pack": "20", "unit": "INR/UNIT", "huliot": 641, "Viega": null, "Kitec": 364, "Geberit Mepla": 1000, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 25x3/4 - KLIMAPRESS", "category": "Adapters", "pack": "20", "unit": "INR/UNIT", "huliot": 713, "Viega": 830, "Kitec": 368, "Geberit Mepla": null, "Kantherm": 743}, {"name": "FEMALE THREADED ADAPTER 25x1 - KLIMAPRESS", "category": "Adapters", "pack": "18", "unit": "INR/UNIT", "huliot": 1078, "Viega": 800, "Kitec": 476, "Geberit Mepla": null, "Kantherm": 810}, {"name": "FEMALE THREADED ADAPTER 32x3/4 - KLIMAPRESS", "category": "Adapters", "pack": "18", "unit": "INR/UNIT", "huliot": 892, "Viega": null, "Kitec": 1485, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 32x1 - KLIMAPRESS", "category": "Adapters", "pack": "18", "unit": "INR/UNIT", "huliot": 1008, "Viega": 2810, "Kitec": 975, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 32x1'1/4 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 1393, "Viega": 2890, "Kitec": null, "Geberit Mepla": 1850, "Kantherm": 2734}, {"name": "FEMALE THREADED ADAPTER 40x1'1/4 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 2999, "Viega": 3470, "Kitec": 202, "Geberit Mepla": null, "Kantherm": 3090}, {"name": "FEMALE THREADED ADAPTER 40x1'1/2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 2832, "Viega": 3880, "Kitec": 194, "Geberit Mepla": null, "Kantherm": 3410}, {"name": "FEMALE THREADED ADAPTER 50x1'1/4 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 3669, "Viega": null, "Kitec": 333, "Geberit Mepla": 4270, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 50x1'1/2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 3984, "Viega": 6630, "Kitec": 468, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 50x2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 4335, "Viega": null, "Kitec": 632, "Geberit Mepla": 7030, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 63x2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 6988, "Viega": 8370, "Kitec": 731, "Geberit Mepla": 14510, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 75x2'1/2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 13283, "Viega": null, "Kitec": 1200, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 90x3 - KLIMAPRESS VX", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 36670, "Viega": null, "Kitec": 4397, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE THREADED ADAPTER 110x4 - KLIMAPRESS VX", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 48804, "Viega": null, "Kitec": 1487, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE THREADED ADAPTER  16x1/2 - KLIMAPRESS", "category": "Adapters", "pack": "40", "unit": "INR/UNIT", "huliot": 360, "Viega": 380, "Kitec": 75, "Geberit Mepla": 680, "Kantherm": 340}, {"name": "MALE THREADED ADAPTER  20x1/2 - KLIMAPRESS", "category": "Adapters", "pack": "35", "unit": "INR/UNIT", "huliot": 413, "Viega": 500, "Kitec": 92, "Geberit Mepla": 960, "Kantherm": 418}, {"name": "MALE THREADED ADAPTER  20x3/4 - KLIMAPRESS", "category": "Adapters", "pack": "30", "unit": "INR/UNIT", "huliot": 541, "Viega": 540, "Kitec": 95, "Geberit Mepla": null, "Kantherm": 495}, {"name": "MALE THREADED ADAPTER  25x1/2 - KLIMAPRESS", "category": "Adapters", "pack": "20", "unit": "INR/UNIT", "huliot": 621, "Viega": null, "Kitec": null, "Geberit Mepla": 960, "Kantherm": 726}, {"name": "MALE THREADED ADAPTER  25x3/4 - KLIMAPRESS", "category": "Adapters", "pack": "20", "unit": "INR/UNIT", "huliot": 701, "Viega": 830, "Kitec": 143, "Geberit Mepla": 960, "Kantherm": 737}, {"name": "MALE THREADED ADAPTER  25x1 - KLIMAPRESS", "category": "Adapters", "pack": "18", "unit": "INR/UNIT", "huliot": 929, "Viega": 800, "Kitec": 140, "Geberit Mepla": null, "Kantherm": 780}, {"name": "MALE THREADED ADAPTER  32x3/4 - KLIMAPRESS", "category": "Adapters", "pack": "18", "unit": "INR/UNIT", "huliot": 879, "Viega": null, "Kitec": 952, "Geberit Mepla": 1250, "Kantherm": null}, {"name": "MALE THREADED ADAPTER  32x1 - KLIMAPRESS", "category": "Adapters", "pack": "18", "unit": "INR/UNIT", "huliot": 879, "Viega": 2370, "Kitec": 128, "Geberit Mepla": 1880, "Kantherm": 2220}, {"name": "MALE THREADED ADAPTER  32x1'1/4 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 1208, "Viega": 2890, "Kitec": 138, "Geberit Mepla": null, "Kantherm": 2602}, {"name": "MALE THREADED ADAPTER  40x1'1/4 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 2989, "Viega": 4740, "Kitec": 192, "Geberit Mepla": null, "Kantherm": 2987}, {"name": "MALE THREADED ADAPTER  40x1'1/2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 2928, "Viega": null, "Kitec": 128, "Geberit Mepla": 4270, "Kantherm": 3410}, {"name": "MALE THREADED ADAPTER  50x1'1/2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 3890, "Viega": 6630, "Kitec": null, "Geberit Mepla": null, "Kantherm": 6537}, {"name": "MALE THREADED ADAPTER  50x2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 4336, "Viega": null, "Kitec": 429, "Geberit Mepla": 6590, "Kantherm": null}, {"name": "MALE THREADED ADAPTER  63x2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 7482, "Viega": 8370, "Kitec": 632, "Geberit Mepla": null, "Kantherm": 6910}, {"name": "MALE THREADED ADAPTER  75x2'1/2 - KLIMAPRESS", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 12523, "Viega": null, "Kitec": 1200, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE THREADED ADAPTER  90x3 - KLIMAPRESS VX", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 37754, "Viega": null, "Kitec": 1284, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE THREADED ADAPTER 110x4 - KLIMAPRESS VX", "category": "Adapters", "pack": "1", "unit": "INR/UNIT", "huliot": 48457, "Viega": null, "Kitec": 1487, "Geberit Mepla": null, "Kantherm": null}, {"name": "SOCKET  16x16 - KLIMAPRESS", "category": "Other Fittings", "pack": "45", "unit": "INR/UNIT", "huliot": 312, "Viega": 390, "Kitec": 142, "Geberit Mepla": 420, "Kantherm": 781}, {"name": "SOCKET  20x20 - KLIMAPRESS", "category": "Other Fittings", "pack": "36", "unit": "INR/UNIT", "huliot": 433, "Viega": 450, "Kitec": 173, "Geberit Mepla": 510, "Kantherm": 935}, {"name": "SOCKET  25x25 - KLIMAPRESS", "category": "Other Fittings", "pack": "20", "unit": "INR/UNIT", "huliot": 822, "Viega": 700, "Kitec": 264, "Geberit Mepla": 720, "Kantherm": 1309}, {"name": "SOCKET  32x32 - KLIMAPRESS", "category": "Other Fittings", "pack": "12", "unit": "INR/UNIT", "huliot": 1080, "Viega": 1900, "Kitec": 285, "Geberit Mepla": 1250, "Kantherm": 1640}, {"name": "SOCKET  40x40 - KLIMAPRESS", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 2893, "Viega": 3780, "Kitec": 402, "Geberit Mepla": 1570, "Kantherm": 3310}, {"name": "SOCKET  50x50 - KLIMAPRESS", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 3444, "Viega": 5060, "Kitec": 786, "Geberit Mepla": 3360, "Kantherm": 4486}, {"name": "SOCKET  63x63 - KLIMAPRESS", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 7801, "Viega": 7300, "Kitec": 962, "Geberit Mepla": 7690, "Kantherm": 6440}, {"name": "SOCKET  75x75 - KLIMAPRESS", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 11937, "Viega": null, "Kitec": 1240, "Geberit Mepla": null, "Kantherm": null}, {"name": "SOCKET  90x90 - KLIMAPRESS VX", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 60224, "Viega": null, "Kitec": 1880, "Geberit Mepla": null, "Kantherm": null}, {"name": "SOCKET  110x110 - KLIMAPRESS VX", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 89530, "Viega": null, "Kitec": 2356, "Geberit Mepla": null, "Kantherm": null}, {"name": "REDUCER 20x16 - KLIMAPRESS", "category": "Reducers", "pack": "37", "unit": "INR/UNIT", "huliot": 392, "Viega": 470, "Kitec": 159, "Geberit Mepla": null, "Kantherm": null}, {"name": "REDUCER 25x16 - KLIMAPRESS", "category": "Reducers", "pack": "20", "unit": "INR/UNIT", "huliot": 597, "Viega": 690, "Kitec": 207, "Geberit Mepla": 510, "Kantherm": null}, {"name": "REDUCER 25x20 - KLIMAPRESS", "category": "Reducers", "pack": "20", "unit": "INR/UNIT", "huliot": 637, "Viega": 610, "Kitec": 221, "Geberit Mepla": 720, "Kantherm": null}, {"name": "REDUCER 32x20 - KLIMAPRESS", "category": "Reducers", "pack": "15", "unit": "INR/UNIT", "huliot": 852, "Viega": 1350, "Kitec": 281, "Geberit Mepla": 720, "Kantherm": 1143}, {"name": "REDUCER 32x25 - KLIMAPRESS", "category": "Reducers", "pack": "12", "unit": "INR/UNIT", "huliot": 964, "Viega": 1410, "Kitec": 320, "Geberit Mepla": 1250, "Kantherm": 1145}, {"name": "REDUCER 40x20 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 1792, "Viega": 2821, "Kitec": 443, "Geberit Mepla": 1250, "Kantherm": 2234}, {"name": "REDUCER 40x25 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 2292, "Viega": 2940, "Kitec": 467, "Geberit Mepla": 1250, "Kantherm": 2343}, {"name": "REDUCER 40x32 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 2427, "Viega": 3660, "Kitec": 337, "Geberit Mepla": null, "Kantherm": 3201}, {"name": "REDUCER 50x25 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 3261, "Viega": 3949, "Kitec": 629, "Geberit Mepla": 1560, "Kantherm": null}, {"name": "REDUCER 50x32 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 3213, "Viega": 4710, "Kitec": 567, "Geberit Mepla": 1560, "Kantherm": 4142}, {"name": "REDUCER 50x40 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 3789, "Viega": 4750, "Kitec": 571, "Geberit Mepla": null, "Kantherm": 4280}, {"name": "REDUCER 63x40 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 8332, "Viega": null, "Kitec": 986, "Geberit Mepla": 3840, "Kantherm": 8696}, {"name": "REDUCER 63x50 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 8587, "Viega": null, "Kitec": 1082, "Geberit Mepla": 7690, "Kantherm": 6590}, {"name": "REDUCER 75x63 - KLIMAPRESS", "category": "Reducers", "pack": "1", "unit": "INR/UNIT", "huliot": 15916, "Viega": null, "Kitec": 1850, "Geberit Mepla": null, "Kantherm": null}, {"name": "CRAZY NUT SOCKET 16x1/2\" - KLIMAPRESS", "category": "Other Fittings", "pack": "40", "unit": "INR/UNIT", "huliot": 496, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "CRAZY NUT SOCKET 20x1/2\" - KLIMAPRESS", "category": "Other Fittings", "pack": "40", "unit": "INR/UNIT", "huliot": 548, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "CRAZY NUT SOCKET 20x3/4\" - KLIMAPRESS", "category": "Other Fittings", "pack": "30", "unit": "INR/UNIT", "huliot": 816, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "CRAZY NUT SOCKET 25x3/4\" - KLIMAPRESS", "category": "Other Fittings", "pack": "18", "unit": "INR/UNIT", "huliot": 876, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "CRAZY NUT SOCKET 32x1\" - KLIMAPRESS", "category": "Other Fittings", "pack": "14", "unit": "INR/UNIT", "huliot": 1324, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 16x16 - KLIMAPRESS", "category": "Elbows", "pack": "32", "unit": "INR/UNIT", "huliot": 379, "Viega": 430, "Kitec": 283, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 20x20 - KLIMAPRESS", "category": "Elbows", "pack": "22", "unit": "INR/UNIT", "huliot": 552, "Viega": 480, "Kitec": 176, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 25x25 - KLIMAPRESS", "category": "Elbows", "pack": "11", "unit": "INR/UNIT", "huliot": 1005, "Viega": 720, "Kitec": 269, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 32x32 - KLIMAPRESS", "category": "Elbows", "pack": "8", "unit": "INR/UNIT", "huliot": 1623, "Viega": 2970, "Kitec": 284, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 40x40 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 4265, "Viega": 5860, "Kitec": 362, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 50x50 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 5824, "Viega": 8080, "Kitec": 1240, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 63x63 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 10829, "Viega": 11910, "Kitec": 1677, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 75x75 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 16595, "Viega": null, "Kitec": 2402, "Geberit Mepla": null, "Kantherm": null}, {"name": "ELBOW 90\u00ba 90x90 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 101053, "Viega": null, "Kitec": 3064, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 16x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "30", "unit": "INR/UNIT", "huliot": 543, "Viega": 810, "Kitec": 141, "Geberit Mepla": null, "Kantherm": 732}, {"name": "FEMALE ELBOW 90\u00ba 16x3/4 - KLIMAPRESS", "category": "Elbows", "pack": "18", "unit": "INR/UNIT", "huliot": 618, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 20x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "25", "unit": "INR/UNIT", "huliot": 613, "Viega": 825, "Kitec": 427, "Geberit Mepla": null, "Kantherm": 1111}, {"name": "FEMALE ELBOW 90\u00ba 20x3/4 - KLIMAPRESS", "category": "Elbows", "pack": "20", "unit": "INR/UNIT", "huliot": 675, "Viega": 1220, "Kitec": 516, "Geberit Mepla": null, "Kantherm": 1755}, {"name": "FEMALE ELBOW 90\u00ba 25x3/4 - KLIMAPRESS", "category": "Elbows", "pack": "15", "unit": "INR/UNIT", "huliot": 936, "Viega": null, "Kitec": 203, "Geberit Mepla": null, "Kantherm": 2178}, {"name": "FEMALE ELBOW 90\u00ba 25x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "10", "unit": "INR/UNIT", "huliot": 756, "Viega": 911, "Kitec": 225, "Geberit Mepla": null, "Kantherm": 1991}, {"name": "FEMALE ELBOW 90\u00ba 25x1 - KLIMAPRESS", "category": "Elbows", "pack": "10", "unit": "INR/UNIT", "huliot": 1404, "Viega": null, "Kitec": 212, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 32x3/4 - KLIMAPRESS", "category": "Elbows", "pack": "12", "unit": "INR/UNIT", "huliot": 1312, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": 4565}, {"name": "FEMALE ELBOW 90\u00ba 32x1 - KLIMAPRESS", "category": "Elbows", "pack": "8", "unit": "INR/UNIT", "huliot": 1425, "Viega": null, "Kitec": 216, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 40x1'1/2 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 3620, "Viega": null, "Kitec": 241, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 40x1'1/4 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 3631, "Viega": null, "Kitec": 918, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 40x1 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 2809, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 50x1'1/2 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 5013, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 50x1 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 4368, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 63x2 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 10477, "Viega": null, "Kitec": 1131, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 75x2'1/2 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 19837, "Viega": null, "Kitec": 1642, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 90x3 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 53131, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE ELBOW 90\u00ba 110x4 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 69565, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE FEMALE CONNECTOR", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 694, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE ELBOW 90\u00ba 16x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "25", "unit": "INR/UNIT", "huliot": 526, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE ELBOW 90\u00ba 20x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "20", "unit": "INR/UNIT", "huliot": 616, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE ELBOW 90\u00ba 25x3/4 - KLIMAPRESS", "category": "Elbows", "pack": "12", "unit": "INR/UNIT", "huliot": 936, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE ELBOW 90\u00ba 25x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "12", "unit": "INR/UNIT", "huliot": 910, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE ELBOW 90\u00ba 32x1 - KLIMAPRESS", "category": "Elbows", "pack": "8", "unit": "INR/UNIT", "huliot": 1398, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE ELBOW 90\u00ba 75x21/2 - KLIMAPRESS", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 25428, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE ELBOW 90\u00ba 90x3 - KLIMAPRESS VX", "category": "Elbows", "pack": "1", "unit": "INR/UNIT", "huliot": 47644, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "WALL PLATE FEMALE ELBOW 16x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "14", "unit": "INR/UNIT", "huliot": 802, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "WALL PLATE FEMALE ELBOW 20x1/2 - KLIMAPRESS", "category": "Elbows", "pack": "12", "unit": "INR/UNIT", "huliot": 892, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "WALL PLATE FEMALE ELBOW 25x3/4 - KLIMAPRESS", "category": "Elbows", "pack": "10", "unit": "INR/UNIT", "huliot": 1782, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "TEE 16x16x16 - KLIMAPRESS", "category": "Tees", "pack": "20", "unit": "INR/UNIT", "huliot": 635, "Viega": 640, "Kitec": 188, "Geberit Mepla": 970, "Kantherm": 440}, {"name": "TEE 20x20x20 - KLIMAPRESS", "category": "Tees", "pack": "14", "unit": "INR/UNIT", "huliot": 777, "Viega": 670, "Kitec": 247, "Geberit Mepla": 1020, "Kantherm": 580}, {"name": "TEE 25x25x25 - KLIMAPRESS", "category": "Tees", "pack": "8", "unit": "INR/UNIT", "huliot": 1424, "Viega": 1010, "Kitec": 376, "Geberit Mepla": 1290, "Kantherm": 890}, {"name": "TEE 32x32x32 - KLIMAPRESS", "category": "Tees", "pack": "4", "unit": "INR/UNIT", "huliot": 2020, "Viega": 4090, "Kitec": 382, "Geberit Mepla": 1880, "Kantherm": 4135}, {"name": "TEE 40x40x40 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 5968, "Viega": 7910, "Kitec": 589, "Geberit Mepla": 2760, "Kantherm": 6530}, {"name": "TEE 50x50x50 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 8227, "Viega": 14210, "Kitec": 1672, "Geberit Mepla": 6590, "Kantherm": 9215}, {"name": "TEE 63x63x63 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 15255, "Viega": 20280, "Kitec": 2322, "Geberit Mepla": 14510, "Kantherm": 12310}, {"name": "TEE 75x75x75 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 22010, "Viega": null, "Kitec": 3290, "Geberit Mepla": null, "Kantherm": null}, {"name": "TEE 90x90x90 - KLIMAPRESS VX", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 93709, "Viega": null, "Kitec": 4282, "Geberit Mepla": null, "Kantherm": null}, {"name": "REDUCED TEE 16x20x16 - KLIMAPRESS", "category": "Tees", "pack": "18", "unit": "INR/UNIT", "huliot": 730, "Viega": null, "Kitec": 221, "Geberit Mepla": null, "Kantherm": 517}, {"name": "REDUCED TEE 16x25x16 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 1254, "Viega": null, "Kitec": 715, "Geberit Mepla": null, "Kantherm": null}, {"name": "REDUCED TEE 20x16x16 - KLIMAPRESS", "category": "Tees", "pack": "20", "unit": "INR/UNIT", "huliot": 630, "Viega": 710, "Kitec": 234, "Geberit Mepla": null, "Kantherm": 501}, {"name": "REDUCED TEE 20x16x20 - KLIMAPRESS", "category": "Tees", "pack": "18", "unit": "INR/UNIT", "huliot": 752, "Viega": 660, "Kitec": 277, "Geberit Mepla": null, "Kantherm": 539}, {"name": "REDUCED TEE 20x20x16 - KLIMAPRESS", "category": "Tees", "pack": "18", "unit": "INR/UNIT", "huliot": 789, "Viega": 750, "Kitec": null, "Geberit Mepla": null, "Kantherm": 539}, {"name": "REDUCED TEE 20x25x20 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 1147, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": 1298}, {"name": "REDUCED TEE 25x16x16 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 1497, "Viega": null, "Kitec": 462, "Geberit Mepla": null, "Kantherm": null}, {"name": "REDUCED TEE 25x16x20 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 1048, "Viega": 1000, "Kitec": 341, "Geberit Mepla": null, "Kantherm": 605}, {"name": "REDUCED TEE 25x16x25 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 1335, "Viega": 990, "Kitec": 471, "Geberit Mepla": null, "Kantherm": 754}, {"name": "REDUCED TEE 25x20x16 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 996, "Viega": null, "Kitec": 479, "Geberit Mepla": 1020, "Kantherm": 655}, {"name": "REDUCED TEE 25x20x20 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 968, "Viega": 830, "Kitec": 807, "Geberit Mepla": null, "Kantherm": 677}, {"name": "REDUCED TEE 25x20x25 - KLIMAPRESS", "category": "Tees", "pack": "8", "unit": "INR/UNIT", "huliot": 1206, "Viega": 920, "Kitec": 1258, "Geberit Mepla": 1020, "Kantherm": 710}, {"name": "REDUCED TEE 25x25x20 - KLIMAPRESS", "category": "Tees", "pack": "7", "unit": "INR/UNIT", "huliot": 1217, "Viega": null, "Kitec": 511, "Geberit Mepla": null, "Kantherm": 781}, {"name": "REDUCED TEE 25x32x25 - KLIMAPRESS", "category": "Tees", "pack": "7", "unit": "INR/UNIT", "huliot": 1876, "Viega": null, "Kitec": 1755, "Geberit Mepla": null, "Kantherm": 1854}, {"name": "REDUCED TEE 32x16x32 - KLIMAPRESS", "category": "Tees", "pack": "8", "unit": "INR/UNIT", "huliot": 1508, "Viega": null, "Kitec": 610, "Geberit Mepla": null, "Kantherm": 2360}, {"name": "REDUCED TEE 32x20x32 - KLIMAPRESS", "category": "Tees", "pack": "7", "unit": "INR/UNIT", "huliot": 1574, "Viega": 2376, "Kitec": 447, "Geberit Mepla": 1290, "Kantherm": 1903}, {"name": "REDUCED TEE 32x25x25 - KLIMAPRESS", "category": "Tees", "pack": "6", "unit": "INR/UNIT", "huliot": 2602, "Viega": null, "Kitec": null, "Geberit Mepla": 1290, "Kantherm": 2283}, {"name": "REDUCED TEE 32x25x32 - KLIMAPRESS", "category": "Tees", "pack": "5", "unit": "INR/UNIT", "huliot": 1783, "Viega": 3212, "Kitec": null, "Geberit Mepla": 1850, "Kantherm": 2410}, {"name": "REDUCED TEE 40X20X40 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 3607, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": 3586}, {"name": "REDUCED TEE 40x25x40 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 4783, "Viega": 5338, "Kitec": null, "Geberit Mepla": 1850, "Kantherm": null}, {"name": "REDUCED TEE 40x32x40 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 5409, "Viega": null, "Kitec": null, "Geberit Mepla": 2710, "Kantherm": null}, {"name": "REDUCED TEE 50x32x50 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 6106, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": 8561}, {"name": "REDUCED TEE 50x40x50 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 7814, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": 9312}, {"name": "REDUCED TEE 63x40x63 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 11655, "Viega": null, "Kitec": null, "Geberit Mepla": 6490, "Kantherm": 15290}, {"name": "REDUCED TEE 63x50x63 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 12414, "Viega": null, "Kitec": null, "Geberit Mepla": 14510, "Kantherm": 18079}, {"name": "REDUCED TEE 75x63x75 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 26272, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 16x1/2 - KLIMAPRESS", "category": "Tees", "pack": "16", "unit": "INR/UNIT", "huliot": 743, "Viega": 940, "Kitec": 21, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 20x1/2 - KLIMAPRESS", "category": "Tees", "pack": "15", "unit": "INR/UNIT", "huliot": 854, "Viega": 1350, "Kitec": 293, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 25x3/4 - KLIMAPRESS", "category": "Tees", "pack": "8", "unit": "INR/UNIT", "huliot": 1511, "Viega": 2370, "Kitec": 1101, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 25x1/2 - KLIMAPRESS", "category": "Tees", "pack": "10", "unit": "INR/UNIT", "huliot": 1103, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 32x1 - KLIMAPRESS", "category": "Tees", "pack": "4", "unit": "INR/UNIT", "huliot": 2959, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 40x1 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 4842, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 40x1'1/4 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 5786, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 50x1 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 6508, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 50x1'1/2 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 8115, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 63x1 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 11655, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 63x2 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 14220, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 63x2'1/2 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 15722, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 75x2'1/2 - KLIMAPRESS", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 28105, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 90x3 - KLIMAPRESS VX", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 82652, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "FEMALE TEE 110x4x110 - KLIMAPRESS VX", "category": "Tees", "pack": "1", "unit": "INR/UNIT", "huliot": 112418, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE TEE 16x1/2 - KLIMAPRESS", "category": "Tees", "pack": "16", "unit": "INR/UNIT", "huliot": 743, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE TEE 20x1/2 - KLIMAPRESS", "category": "Tees", "pack": "12", "unit": "INR/UNIT", "huliot": 858, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE TEE 25x1/2 - KLIMAPRESS", "category": "Tees", "pack": "5", "unit": "INR/UNIT", "huliot": 1495, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE TEE 25x3/4 - KLIMAPRESS", "category": "Tees", "pack": "5", "unit": "INR/UNIT", "huliot": 1511, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MALE TEE 32x1 - KLIMAPRESS", "category": "Tees", "pack": "4", "unit": "INR/UNIT", "huliot": 2959, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "Manifold with stop valve, 3/4\" main and 1/2\" branch", "category": "Manifolds", "pack": "1", "unit": "INR/UNIT", "huliot": 3553, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "Manifold with stop valve, 1\" main and 3/4\" branch", "category": "Manifolds", "pack": "1", "unit": "INR/UNIT", "huliot": 4463, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "Manifold with stop valve, 1\" main and 1/2\" branch", "category": "Manifolds", "pack": "1", "unit": "INR/UNIT", "huliot": 4266, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "Connector core for PEX. PE-RT/AL/PE-RT pipe", "category": "Multilayer Pipe", "pack": "1", "unit": "INR/UNIT", "huliot": 426, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "BALL VALVE 16 - KLIMAPRESS", "category": "Valves", "pack": "1", "unit": "INR/UNIT", "huliot": 1247, "Viega": null, "Kitec": null, "Geberit Mepla": 3150, "Kantherm": null}, {"name": "BALL VALVE 20 - KLIMAPRESS", "category": "Valves", "pack": "1", "unit": "INR/UNIT", "huliot": 1346, "Viega": null, "Kitec": null, "Geberit Mepla": 3500, "Kantherm": null}, {"name": "BALL VALVE 25 - KLIMAPRESS", "category": "Valves", "pack": "1", "unit": "INR/UNIT", "huliot": 2964, "Viega": null, "Kitec": null, "Geberit Mepla": 4270, "Kantherm": null}, {"name": "BALL VALVE 32 - KLIMAPRESS", "category": "Valves", "pack": "1", "unit": "INR/UNIT", "huliot": 4065, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "END CAP BRASS 16 MM", "category": "End Caps", "pack": "1", "unit": "INR/UNIT", "huliot": "Shortly Intr.", "Viega": 320, "Kitec": 76, "Geberit Mepla": null, "Kantherm": 215}, {"name": "END CAP BRASS 20 MM", "category": "End Caps", "pack": "1", "unit": "INR/UNIT", "huliot": 465, "Viega": 330, "Kitec": 106, "Geberit Mepla": 470, "Kantherm": 264}, {"name": "END CAP BRASS 25 MM", "category": "End Caps", "pack": "1", "unit": "INR/UNIT", "huliot": 565, "Viega": 600, "Kitec": 138, "Geberit Mepla": null, "Kantherm": 480}, {"name": "END CAP BRASS 32 MM", "category": "End Caps", "pack": "1", "unit": "INR/UNIT", "huliot": 665, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": 869}, {"name": "LOOSE NUT ADAPTER 25x1 - PPSU", "category": "Adapters", "pack": "10", "unit": "INR/UNIT", "huliot": 1210, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "LOOSE NUT ADAPTER 32x1\"1/4 - PPSU", "category": "Adapters", "pack": "5", "unit": "INR/UNIT", "huliot": 2946, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "CALIBRATOR - 16.20.25.32 MM", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 2000, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "CALIBRATOR - 40.50.63 MM", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 7500, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MANUAL PIPE DEBURRING TOOL 16 to 32 MM", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 4000, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MANUAL PRESS TOOL - 360\u00b0 ROTATING HEAD", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 7000, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "MANUAL PRESS TOOL (BLUE KIT)", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 7000, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "BATTERY POWERED PRESS TOOL WITH U JAWS", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 67500, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}, {"name": "BATTERY POWERED PRESS TOOL WITH V JAWS", "category": "Other Fittings", "pack": "1", "unit": "INR/UNIT", "huliot": 77600, "Viega": null, "Kitec": null, "Geberit Mepla": null, "Kantherm": null}];

const COMPETITORS = ['Viega','Kitec','Geberit Mepla','Kantherm'];
const COMP_COLORS = {
  'Viega':'#4a8cff',
  'Kitec':'#ff9c4a',
  'Geberit Mepla':'#c97eff',
  'Kantherm':'#ff6b9d',
};

let activeTab = 'Viega';
let sortCol = null;
let sortDir = 1;
let selectedRow = null;

function fmt(v){
  if(v==null) return null;
  return '₹' + Number(v).toLocaleString('en-IN');
}

function calcDiff(huliot, comp){
  if(!huliot||!comp) return null;
  return ((comp - huliot) / huliot * 100);
}

function buildSummary(){
  const strip = document.getElementById('summaryStrip');
  strip.innerHTML = '';
  COMPETITORS.forEach(comp=>{
    const items = PRODUCTS.filter(p=>p[comp]!=null && p.huliot!=null);
    const cheaper = items.filter(p=>p[comp]>p.huliot);
    const avgSav = cheaper.length ? cheaper.reduce((s,p)=>s+calcDiff(p.huliot,p[comp]),0)/cheaper.length : 0;
    strip.innerHTML += `<div class="stat-card">
      <div class="stat-num">${cheaper.length}</div>
      <div class="stat-label">Huliot Cheaper vs ${comp}</div>
    </div>
    <div class="stat-card">
      <div class="stat-num ${avgSav<0?'danger':''}">${avgSav>=0?'+':''}${avgSav.toFixed(0)}%</div>
      <div class="stat-label">Avg. Savings vs ${comp}</div>
    </div>`;
  });
}

function buildTabs(){
  const bar = document.getElementById('tabsBar');
  bar.innerHTML = '';
  COMPETITORS.forEach(comp=>{
    const items = PRODUCTS.filter(p=>p[comp]!=null);
    const t = document.createElement('div');
    t.className = 'tab' + (comp===activeTab?' active':'');
    t.innerHTML = `${comp} <span class="tab-count">${items.length}</span>`;
    t.onclick = ()=>{activeTab=comp; buildTabs(); renderTable(); buildChartForSelected();};
    bar.appendChild(t);
  });
}

function getFiltered(){
  const q = document.getElementById('searchInput').value.toLowerCase();
  const cat = document.getElementById('categoryFilter').value;
  const sav = document.getElementById('savingsFilter').value;
  let rows = PRODUCTS.filter(p=>p[activeTab]!=null);
  if(q) rows = rows.filter(p=>p.name.toLowerCase().includes(q));
  if(cat) rows = rows.filter(p=>p.category===cat);
  if(sav==='cheaper') rows = rows.filter(p=>p[activeTab]>p.huliot);
  if(sav==='expensive') rows = rows.filter(p=>p[activeTab]<p.huliot);
  if(sav==='big') rows = rows.filter(p=>p[activeTab] && calcDiff(p.huliot,p[activeTab])>30);
  if(sortCol!==null){
    rows.sort((a,b)=>{
      let va,vb;
      if(sortCol==='name') {va=a.name;vb=b.name;}
      else if(sortCol==='huliot') {va=a.huliot||0;vb=b.huliot||0;}
      else if(sortCol==='comp') {va=a[activeTab]||0;vb=b[activeTab]||0;}
      else if(sortCol==='diff') {va=calcDiff(a.huliot,a[activeTab])||0;vb=calcDiff(b.huliot,b[activeTab])||0;}
      return (va>vb?1:-1)*sortDir;
    });
  }
  return rows;
}

function setSort(col){
  if(sortCol===col) sortDir*=-1; else {sortCol=col;sortDir=1;}
  renderTable();
}

function renderTable(){
  const rows = getFiltered();
  const body = document.getElementById('tableBody');
  const head = document.getElementById('tableHead');
  const empty = document.getElementById('emptyState');
  const cnt = document.getElementById('rowCount');
  cnt.textContent = rows.length + ' items';
  
  const colClass = c=>sortCol===c?(sortDir===1?'sort-asc':'sort-desc'):'';
  head.innerHTML = `<tr>
    <th onclick="setSort('name')" class="${colClass('name')}">Product / SKU</th>
    <th>Pack</th>
    <th onclick="setSort('huliot')" class="${colClass('huliot')}">Huliot ₹</th>
    <th onclick="setSort('comp')" class="${colClass('comp')}">${activeTab} ₹</th>
    <th onclick="setSort('diff')" class="${colClass('diff')}">Savings</th>
  </tr>`;

  if(!rows.length){
    body.innerHTML='';
    empty.style.display='block';
    return;
  }
  empty.style.display='none';

  const maxAbsDiff = Math.max(...rows.map(r=>Math.abs(calcDiff(r.huliot,r[activeTab])||0)));

  body.innerHTML = rows.map((p,i)=>{
    const diff = calcDiff(p.huliot,p[activeTab]);
    const cheaper = diff>0;
    const barW = diff!=null ? Math.round(Math.abs(diff)/Math.max(maxAbsDiff,1)*90) : 0;
    const diffBadge = diff!=null ? `<div class="diff-badge ${cheaper?'cheaper':'expensive'}">${cheaper?'▼':'▲'}${Math.abs(diff).toFixed(1)}%</div>
      <div class="diff-bar-wrap"><div class="diff-bar ${cheaper?'cheaper':'expensive'}" style="width:${barW}px"></div><span class="bar-pct ${cheaper?'cheaper':'expensive'}">${cheaper?'Huliot saves '+fmt(p[activeTab]-p.huliot):'Higher by '+fmt(p.huliot-p[activeTab])}</span></div>` : '<span style="color:var(--muted);font-size:11px">N/A</span>';
    return `<tr onclick="selectRow(this,'${p.name.replace(/'/g,"\'")}')">
      <td class="td-name">${p.name} <span class="cat-pill">${p.category}</span></td>
      <td class="td-pack">${p.pack}</td>
      <td class="td-price huliot">${fmt(p.huliot)||'—'}<br><span style="font-size:9px;color:var(--muted)">${p.unit}</span></td>
      <td class="td-price competitor">${fmt(p[activeTab])||'—'}</td>
      <td class="diff-cell">${diffBadge}</td>
    </tr>`;
  }).join('');
}

function selectRow(tr, name){
  document.querySelectorAll('tr.selected-row').forEach(r=>r.classList.remove('selected-row'));
  tr.classList.add('selected-row');
  selectedRow = name;
  buildChartForSelected();
}

function buildChartForSelected(){
  const cs = document.getElementById('chartSection');
  const cc = document.getElementById('chartContainer');
  if(!selectedRow){cs.style.display='none';return;}
  const p = PRODUCTS.find(x=>x.name===selectedRow);
  if(!p){cs.style.display='none';return;}
  
  const entries = [['Huliot',p.huliot,'#00c6a2'],...COMPETITORS.map(c=>[c,p[c],COMP_COLORS[c]])].filter(e=>e[1]!=null);
  const maxVal = Math.max(...entries.map(e=>e[1]));
  
  cc.innerHTML = entries.map(([name,val,color])=>{
    const h = Math.round((val/maxVal)*80)+4;
    return `<div class="chart-bar-group">
      <div class="chart-bar-outer"><div class="chart-bar" style="height:${h}px;background:${color}" data-val="₹${Number(val).toLocaleString('en-IN')}"></div></div>
      <div class="chart-bar-label" style="color:${color}">${name.replace('Geberit Mepla','Geberit')}</div>
      <div class="chart-bar-label" style="color:var(--text);font-size:9px">₹${Number(val).toLocaleString('en-IN')}</div>
    </div>`;
  }).join('');
  cs.style.display='block';
}

// Init
buildSummary();
buildTabs();
renderTable();
document.getElementById('searchInput').addEventListener('input', ()=>renderTable());
document.getElementById('categoryFilter').addEventListener('change', ()=>renderTable());
document.getElementById('savingsFilter').addEventListener('change', ()=>renderTable());
document.getElementById('footerDate').textContent = 'Generated: ' + new Date().toLocaleDateString('en-IN',{year:'numeric',month:'long',day:'numeric'});

// ── DOWNLOAD HELPERS ──────────────────────────────────────────────

function escCSV(v){
  if(v==null) return '';
  const s=String(v);
  return s.includes(',')||s.includes('"')||s.includes('\n') ? '"'+s.replace(/"/g,'""')+'"' : s;
}

// Download currently visible/filtered tab as CSV
function downloadCSV(){
  const rows = getFiltered();
  const comp = activeTab;
  const headers = ['Product Name','Category','Pack','Unit','Huliot Price (INR)',comp+' Price (INR)','Diff Amount (INR)','Diff %','Huliot Advantage'];
  const lines = [headers.join(',')];
  rows.forEach(p=>{
    const diff = (p.huliot!=null && p[comp]!=null) ? (p[comp]-p.huliot) : null;
    const pct  = (p.huliot!=null && p[comp]!=null) ? ((p[comp]-p.huliot)/p.huliot*100) : null;
    const adv  = diff==null ? '' : diff>0 ? 'Huliot Cheaper by ₹'+Math.abs(diff).toFixed(0) : diff<0 ? comp+' Cheaper by ₹'+Math.abs(diff).toFixed(0) : 'Same Price';
    lines.push([
      escCSV(p.name), escCSV(p.category), escCSV(p.pack), escCSV(p.unit),
      p.huliot??'', p[comp]??'',
      diff!=null?diff.toFixed(0):'',
      pct!=null?pct.toFixed(1)+'%':'',
      escCSV(adv)
    ].join(','));
  });
  triggerDownload(lines.join('\n'), `Huliot_vs_${comp.replace(/ /g,'_')}_${yyyymmdd()}.csv`, 'text/csv');
}

// Download ALL competitors in one big CSV
function downloadAllCSV(){
  const headers = ['Product Name','Category','Pack','Unit','Huliot Price (INR)','Viega (INR)','Kitec (INR)','Geberit Mepla (INR)','Kantherm (INR)','Savings vs Viega %','Savings vs Kitec %','Savings vs Geberit %','Savings vs Kantherm %'];
  const lines = [headers.join(',')];
  PRODUCTS.forEach(p=>{
    const pct = c=>(p.huliot!=null&&p[c]!=null)?((p[c]-p.huliot)/p.huliot*100).toFixed(1)+'%':'';
    lines.push([
      escCSV(p.name),escCSV(p.category),escCSV(p.pack),escCSV(p.unit),
      p.huliot??'',p['Viega']??'',p['Kitec']??'',p['Geberit Mepla']??'',p['Kantherm']??'',
      pct('Viega'),pct('Kitec'),pct('Geberit Mepla'),pct('Kantherm')
    ].join(','));
  });
  triggerDownload(lines.join('\n'), `Huliot_Full_Comparison_${yyyymmdd()}.csv`, 'text/csv');
}

// ── TRUE XLSX WITH FORMULAS (SheetJS) ─────────────────────────────
function downloadExcel(){
  const rows = getFiltered();
  const comp = activeTab;
  if(!window.XLSX){alert('SheetJS not loaded. Check internet connection.');return;}

  const wb = XLSX.utils.book_new();

  // ── HELPER: cell style shorthands ───────────────────────────────
  // SheetJS Community doesn't support styles natively, but we can
  // set column widths and use proper formula cells.

  // ── SHEET 1: Comparison (filtered, with formulas) ───────────────
  const wsData = [];

  // Title rows
  wsData.push([`HULIOT vs ${comp.toUpperCase()} — MLCP Price Comparison`,'','','','','','','','']);
  wsData.push([`Generated: ${new Date().toLocaleDateString('en-IN',{year:'numeric',month:'long',day:'numeric'})} | Filtered: ${rows.length} products`,'','','','','','','','']);
  wsData.push([]); // blank row

  // Header row (row index 3 → Excel row 4)
  wsData.push([
    'Sr.No','Product Name','Category','Pack / Unit',
    `Huliot Price\n(INR)`,`${comp} Price\n(INR)`,
    'Diff Amount\n(INR)','Savings %','Huliot Advantage'
  ]);

  const dataStartRow = 5; // Excel row where data begins (1-indexed)
  rows.forEach((p, i)=>{
    const r = dataStartRow + i; // actual Excel row number
    const eCol = `E${r}`;      // Huliot price column
    const fCol = `F${r}`;      // Competitor price column
    wsData.push([
      i+1,
      p.name,
      p.category,
      p.pack ? `${p.pack} | ${p.unit}` : p.unit,
      p.huliot ?? '',
      p[comp] ?? '',
      // Diff Amount formula: =IF(AND(E5<>"",F5<>""),F5-E5,"N/A")
      { f: `IF(AND(${eCol}<>"",${fCol}<>""),${fCol}-${eCol},"N/A")` },
      // Savings % formula: =IF(AND(E5<>"",F5<>""),((F5-E5)/E5)*100,"N/A")
      { f: `IF(AND(${eCol}<>"",${fCol}<>""),((${fCol}-${eCol})/${eCol})*100,"N/A")` },
      // Advantage text formula
      { f: `IF(AND(${eCol}<>"",${fCol}<>""),IF(${fCol}>${eCol},"✔ Huliot Cheaper by ₹"&TEXT(${fCol}-${eCol},"#,##0"),IF(${fCol}<${eCol},"✘ ${comp} Cheaper by ₹"&TEXT(${eCol}-${fCol},"#,##0"),"= Same Price")),"—")` }
    ]);
  });

  // Summary rows at bottom
  const lastDataRow = dataStartRow + rows.length - 1;
  wsData.push([]); // blank
  wsData.push(['SUMMARY','','','','','','','','']);
  wsData.push([
    'Total Products Compared','','','',rows.length,'','','',''
  ]);
  wsData.push([
    `Huliot CHEAPER vs ${comp}`, '','','',
    { f: `COUNTIF(I${dataStartRow}:I${lastDataRow},"✔*")` },
    '','','',''
  ]);
  wsData.push([
    `${comp} Cheaper`, '','','',
    { f: `COUNTIF(I${dataStartRow}:I${lastDataRow},"✘*")` },
    '','','',''
  ]);
  wsData.push([
    'Avg Huliot Savings % (where cheaper)','','','',
    { f: `IFERROR(AVERAGEIF(H${dataStartRow}:H${lastDataRow},">"&0,H${dataStartRow}:H${lastDataRow}),"N/A")` },
    '','','',''
  ]);
  wsData.push([
    'Max Single-Item Savings %','','','',
    { f: `IFERROR(MAX(H${dataStartRow}:H${lastDataRow}),"N/A")` },
    '','','',''
  ]);

  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Column widths
  ws['!cols'] = [
    {wch:6},{wch:52},{wch:18},{wch:22},
    {wch:14},{wch:14},{wch:14},{wch:12},{wch:36}
  ];

  // Freeze header rows
  ws['!freeze'] = {xSplit:0, ySplit:4, topLeftCell:'A5', activePane:'bottomLeft'};

  XLSX.utils.book_append_sheet(wb, ws, `Huliot vs ${comp}`.substring(0,31));

  // ── SHEET 2: All-Competitor Summary (with formulas) ─────────────
  const ws2Data = [];
  ws2Data.push(['HULIOT — ALL COMPETITOR PRICE COMPARISON (MASTER SHEET)','','','','','','','','','','','','']);
  ws2Data.push([`Generated: ${new Date().toLocaleDateString('en-IN',{year:'numeric',month:'long',day:'numeric'})}`,'','','','','','','','','','','','']);
  ws2Data.push([]);
  ws2Data.push([
    'Sr.No','Product Name','Category','Pack','Unit',
    'Huliot (INR)','Viega (INR)','Kitec (INR)','Geberit Mepla (INR)','Kantherm (INR)',
    'Saving vs Viega %','Saving vs Kitec %','Saving vs Geberit %','Saving vs Kantherm %'
  ]);

  const allDataStart = 5;
  PRODUCTS.forEach((p,i)=>{
    const r = allDataStart + i;
    const f = col => `IF(AND(F${r}<>"",${col}${r}<>""),(${col}${r}-F${r})/F${r}*100,"N/A")`;
    ws2Data.push([
      i+1, p.name, p.category, p.pack, p.unit,
      p.huliot??'', p['Viega']??'', p['Kitec']??'', p['Geberit Mepla']??'', p['Kantherm']??'',
      {f:f('G')}, {f:f('H')}, {f:f('I')}, {f:f('J')}
    ]);
  });

  const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
  ws2['!cols'] = [
    {wch:6},{wch:50},{wch:18},{wch:20},{wch:10},
    {wch:13},{wch:13},{wch:13},{wch:18},{wch:14},
    {wch:16},{wch:16},{wch:18},{wch:18}
  ];
  ws2['!freeze'] = {xSplit:0, ySplit:4, topLeftCell:'A5', activePane:'bottomLeft'};
  XLSX.utils.book_append_sheet(wb, ws2, 'All Competitors');

  // Write & trigger download
  const wbOut = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const blob = new Blob([wbOut], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `Huliot_vs_${comp.replace(/ /g,'_')}_${yyyymmdd()}.xlsx`;
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),3000);
}

function yyyymmdd(){
  const d=new Date();
  return d.getFullYear()+('0'+(d.getMonth()+1)).slice(-2)+('0'+d.getDate()).slice(-2);
}

function triggerDownload(content, filename, mime){
  const blob = new Blob(['\uFEFF'+content],{type:mime+';charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),2000);
}
</script>
</body>
</html>
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>技術分析全攻略 · 個股評分系統（POC版）</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/plotly.js/2.26.0/plotly.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0;font-family:'Noto Sans TC',sans-serif}
body{background:#0a0e1a;color:#e0e0e0;min-height:100vh}
.app{display:flex;min-height:100vh}
.sidebar{width:270px;min-width:270px;background:#0f1525;border-right:1px solid rgba(255,255,255,.07);padding:24px 18px;position:sticky;top:0;height:100vh;overflow-y:auto}
.main{flex:1;padding:28px 32px;overflow-x:hidden}
.logo{font-size:20px;font-weight:700;color:#fff;margin-bottom:4px}
.logo-sub{font-size:11px;color:#666;margin-bottom:20px}
.sidebar hr{border:none;border-top:1px solid rgba(255,255,255,.08);margin:16px 0}
label{display:block;font-size:12px;color:#aaa;margin-bottom:6px}
input[type=text],input[type=password]{width:100%;padding:9px 12px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:8px;color:#fff;font-size:14px;outline:none}
input:focus{border-color:#4a9eff}
input::placeholder{color:#555}
.field{margin-bottom:14px}
.range-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.range-val{font-size:13px;font-weight:700;color:#4a9eff}
input[type=range]{width:100%;accent-color:#4a9eff;cursor:pointer}
.btn{width:100%;padding:12px;background:linear-gradient(135deg,#1a6fff,#0d4fd9);border:none;border-radius:10px;color:#fff;font-size:15px;font-weight:700;cursor:pointer;margin-top:10px;letter-spacing:1px;transition:opacity .2s}
.btn:hover{opacity:.88}
.btn:disabled{opacity:.5;cursor:not-allowed}
.legend-h{font-size:12px;font-weight:700;color:#ccc;margin:14px 0 6px}
.legend-i{font-size:12px;color:#999;margin:4px 0}
.page-title{font-size:22px;font-weight:700;color:#fff;margin-bottom:4px}
.page-sub{font-size:13px;color:#666;margin-bottom:24px}
.welcome{background:rgba(74,158,255,.08);border:1px solid rgba(74,158,255,.2);border-radius:12px;padding:40px;text-align:center;margin-top:40px}
.welcome-icon{font-size:48px;margin-bottom:12px}
.welcome-t{font-size:18px;font-weight:700;color:#4a9eff;margin-bottom:8px}
.welcome-d{font-size:14px;color:#888}
.loading{text-align:center;padding:60px 20px}
.spinner{width:40px;height:40px;border:3px solid rgba(74,158,255,.2);border-top-color:#4a9eff;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 16px}
@keyframes spin{to{transform:rotate(360deg)}}
.loading-t{color:#888;font-size:14px}
.err{background:rgba(255,60,60,.08);border:1px solid rgba(255,60,60,.25);border-radius:10px;padding:16px 20px;color:#ff8888;font-size:14px;line-height:1.8}
.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}
.mc{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:14px 18px}
.mc-l{font-size:11px;color:#888;letter-spacing:1px;margin-bottom:6px}
.mc-v{font-size:20px;font-weight:700;color:#fff}
.mc-d{font-size:12px;margin-top:3px}
.up{color:#ef5350}.dn{color:#26a69a}.flat{color:#888}
hr.div{border:none;border-top:1px solid rgba(255,255,255,.07);margin:24px 0}
.sec{font-size:16px;font-weight:700;color:#fff;margin-bottom:14px}
.score-wrap{display:grid;grid-template-columns:200px 1fr;gap:20px;margin-bottom:8px}
.score-card{background:linear-gradient(135deg,#1a1a2e,#16213e);border-radius:16px;padding:28px 16px;text-align:center;border:1px solid rgba(255,255,255,.1)}
.score-n{font-size:76px;font-weight:700;line-height:1}
.score-l{font-size:12px;color:#888;margin-top:6px;letter-spacing:2px}
.score-v{font-size:16px;font-weight:700;margin-top:12px;padding:7px 16px;border-radius:8px;display:inline-block}
.dim-card{background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:12px 16px;margin-bottom:10px}
.dim-row{display:flex;justify-content:space-between;align-items:center}
.dim-t{font-size:12px;color:#888;letter-spacing:1px;margin-bottom:3px}
.dim-d{font-size:11px;color:#666}
.dim-s{font-size:26px;font-weight:700}
.dim-s span{font-size:12px;color:#555}
.pbar{background:rgba(255,255,255,.05);border-radius:4px;height:5px;margin-top:8px}
.pfill{border-radius:4px;height:5px;transition:width .6s}
.sig-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
.sig-t{font-size:13px;font-weight:700;color:#ccc;margin-bottom:8px}
.tag{display:inline-block;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:600;margin:2px}
.bull{background:rgba(0,200,100,.12);color:#00c864;border:1px solid rgba(0,200,100,.3)}
.bear{background:rgba(255,60,60,.12);color:#ff5555;border:1px solid rgba(255,60,60,.3)}
.neu{background:rgba(180,180,180,.08);color:#aaa;border:1px solid rgba(180,180,180,.2)}
.adv{background:linear-gradient(135deg,#0d2137,#0a1a2e);border-left:4px solid #f0a500;border-radius:0 12px 12px 0;padding:20px 24px}
.adv-a{font-size:18px;font-weight:700;margin-bottom:10px}
.adv-t{font-size:14px;color:#ccc;line-height:1.8}
.adv-m{margin-top:12px;font-size:13px;color:#bbb;line-height:2}
.info-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:20px}
.ib{background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px 18px}
.ib-t{font-size:13px;font-weight:700;color:#ccc;margin-bottom:12px}
.rsi-v{font-size:34px;font-weight:700}
.rsi-n{font-size:12px;margin-top:5px;padding:5px 10px;border-radius:6px;display:inline-block}
.rsi-ob{background:rgba(255,60,60,.12);color:#ff5555}
.rsi-os{background:rgba(0,200,100,.12);color:#00c864}
.rsi-ok{background:rgba(240,165,0,.12);color:#f0a500}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;padding:7px 10px;color:#888;font-weight:600;font-size:11px;border-bottom:1px solid rgba(255,255,255,.07)}
td{padding:7px 10px;border-bottom:1px solid rgba(255,255,255,.04)}
.tup{color:#ef5350;font-weight:600}.tdn{color:#26a69a;font-weight:600}
.chart-box{background:rgba(15,20,35,1);border-radius:12px;overflow:hidden;margin-top:4px;width:100%}
.chart-box #chart,.chart-box .js-plotly-plot,.chart-box .plot-container{width:100% !important}
.raw-tog{cursor:pointer;font-size:13px;color:#888;padding:10px 0;border-top:1px solid rgba(255,255,255,.07);margin-top:24px;display:flex;align-items:center;gap:8px;user-select:none}
.raw-tog:hover{color:#bbb}
.raw-body{display:none;margin-top:10px;overflow-x:auto}
.raw-body.open{display:block}

/* ── Batch tabs ── */
.tab-bar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px}
.tab-btn{padding:7px 16px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);
  border-radius:8px;color:#aaa;font-size:13px;cursor:pointer;transition:all .2s;white-space:nowrap}
.tab-btn:hover{background:rgba(255,255,255,.1);color:#fff}
.tab-btn.active{background:rgba(74,158,255,.2);border-color:#4a9eff;color:#4a9eff;font-weight:700}
.tab-btn.pass{border-color:rgba(0,200,100,.5);color:#00c864}
.tab-btn.fail{border-color:rgba(255,60,60,.4);color:#ff5555}
.tab-btn.active.pass{background:rgba(0,200,100,.15)}
.tab-btn.active.fail{background:rgba(255,60,60,.1)}
.tab-panel{display:none}.tab-panel.active{display:block}

/* ── Progress ── */
.batch-progress{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);
  border-radius:10px;padding:16px 20px;margin-bottom:20px}
.progress-bar-wrap{background:rgba(255,255,255,.06);border-radius:4px;height:8px;margin-top:10px}
.progress-bar-fill{background:linear-gradient(90deg,#1a6fff,#00c864);border-radius:4px;
  height:8px;transition:width .4s}

/* ── Summary table ── */
.summary-wrap{width:100%;overflow:visible}
.summary-table{width:100%;border-collapse:collapse;font-size:11.5px;margin-bottom:4px;table-layout:auto}
.summary-table th{text-align:left;padding:6px 4px;color:#888;font-size:10.5px;font-weight:700;
  border-bottom:2px solid rgba(255,255,255,.1);letter-spacing:.3px;overflow-wrap:anywhere;vertical-align:bottom}
.summary-table td{padding:6px 4px;border-bottom:1px solid rgba(255,255,255,.05);cursor:pointer;overflow-wrap:anywhere;vertical-align:top;line-height:1.45}
.summary-table td span{overflow-wrap:anywhere}
.summary-table td.nw{white-space:nowrap;overflow-wrap:normal}
.summary-table .dmi-cell{position:relative}
.summary-table .dmi-pop{display:none;position:absolute;left:4px;top:calc(100% - 4px);z-index:60;background:#1b2238;
  border:1px solid rgba(124,58,237,.5);border-radius:6px;padding:6px 10px;white-space:nowrap;font-size:11.5px;
  font-weight:400;color:#ddd;box-shadow:0 6px 18px rgba(0,0,0,.5);line-height:1.6;pointer-events:none}
.summary-table .dmi-cell:hover .dmi-pop{display:block}
.summary-table tr:hover td{background:rgba(255,255,255,.03)}
.badge-pass{background:rgba(0,200,100,.15);color:#00c864;border-radius:4px;padding:2px 8px;font-size:11px;font-weight:700}
.badge-fail{background:rgba(255,60,60,.1);color:#ff5555;border-radius:4px;padding:2px 8px;font-size:11px;font-weight:700}
.badge-bull{background:rgba(0,200,100,.12);color:#00c864;border-radius:12px;padding:2px 8px;font-size:11px}
.badge-obs{background:rgba(240,165,0,.12);color:#f0a500;border-radius:12px;padding:2px 8px;font-size:11px}
.badge-bear{background:rgba(255,60,60,.1);color:#ff5555;border-radius:12px;padding:2px 8px;font-size:11px}
.badge-justbreak{background:rgba(255,152,0,.16);color:#ff9800;border-radius:4px;padding:2px 8px;font-size:11px;font-weight:700}

/* ── Filter bar ── */
.filter-bar{display:flex;flex-wrap:wrap;gap:8px;padding:12px 16px;border-bottom:1px solid rgba(255,255,255,.08);align-items:center}
.filter-btn{padding:5px 14px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:20px;color:#aaa;font-size:12px;cursor:pointer;transition:all .2s;white-space:nowrap}
.filter-btn:hover{background:rgba(255,255,255,.12);color:#fff}
.filter-btn.active{background:rgba(74,158,255,.2);border-color:#4a9eff;color:#4a9eff;font-weight:700}
.list-toggle-btn{padding:3px 10px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:14px;color:#aaa;font-size:11px;cursor:pointer;transition:all .2s;white-space:nowrap}
.list-toggle-btn:hover{background:rgba(255,255,255,.12);color:#fff}
.list-toggle-btn.active{background:rgba(74,158,255,.2);border-color:#4a9eff;color:#4a9eff;font-weight:700}
.lt-clear{margin-left:5px;opacity:.55;font-weight:700}
.lt-clear:hover{opacity:1;color:#ff5555}
.filter-input{padding:5px 12px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:20px;color:#fff;font-size:12px;outline:none;width:120px}
.filter-input::placeholder{color:#555}
.filter-sep{color:rgba(255,255,255,.15);font-size:18px}
.sort-icon{cursor:pointer;color:#666;margin-left:4px;font-size:11px}
.sort-icon.asc{color:#4a9eff}
.sort-icon.desc{color:#4a9eff}
tr.hidden-row{display:none}

@media(max-width:900px){.app{flex-direction:column}.sidebar{width:100%;height:auto;position:static}.metrics{grid-template-columns:repeat(2,1fr)}.score-wrap{grid-template-columns:1fr}.sig-grid{grid-template-columns:repeat(2,1fr)}.info-grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="app">
<aside class="sidebar">
  <div class="logo">📊 技術分析全攻略</div>
  <div class="logo-sub">朱家泓方法論 · 個股評分系統（POC分價量表版）</div>
  <hr>
  <div class="field"><label>FinMind API Token</label><input type="password" id="iToken" placeholder="輸入您的Token"></div>
  <div class="field">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px">
      <label style="margin:0">批次股票代號</label>
      <div style="display:flex;gap:6px;flex-wrap:wrap">
        <button type="button" class="list-toggle-btn active" id="btnListTWSE" data-list="twse">上市清單</button>
        <button type="button" class="list-toggle-btn" id="btnListTPEX" data-list="tpex">上櫃清單</button>
        <button type="button" class="list-toggle-btn" id="btnList0050" data-list="0050">0050成分股</button>
        <button type="button" class="list-toggle-btn" data-list="my"><span class="lt-label">我的清單</span><span class="lt-clear" data-clear="my" title="清除此清單">×</span></button>
        <button type="button" class="list-toggle-btn" data-list="my1"><span class="lt-label">進攻組合型態</span><span class="lt-clear" data-clear="my1" title="清除此清單">×</span></button>
        <button type="button" class="list-toggle-btn" data-list="my2"><span class="lt-label">穩健組合型態</span><span class="lt-clear" data-clear="my2" title="清除此清單">×</span></button>
        <button type="button" class="list-toggle-btn" data-list="my3"><span class="lt-label">彩券組合型態</span><span class="lt-clear" data-clear="my3" title="清除此清單">×</span></button>
        <button type="button" class="list-toggle-btn" data-list="my4"><span class="lt-label">我的清單4</span><span class="lt-clear" data-clear="my4" title="清除此清單">×</span></button>
      </div>
    </div>
    <textarea id="iStocks" rows="6" placeholder="每行一個，或逗號分隔&#10;例：&#10;2330&#10;2454&#10;2317&#10;0050" style="width:100%;margin-top:6px;padding:9px 12px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:8px;color:#fff;font-size:13px;outline:none;resize:vertical;line-height:1.6"></textarea>
    <div style="display:flex;gap:8px;margin-top:8px">
      <button type="button" class="list-toggle-btn" id="btnSaveMyList" style="flex:1;padding:7px 6px;font-size:12px;text-align:center">💾 更新我的清單</button>
    </div>
    <div style="display:flex;gap:8px;margin-top:8px">
      <button type="button" class="list-toggle-btn" id="btnSaveList1" style="flex:1;padding:7px 6px;font-size:12px;text-align:center">➕ 存為進攻</button>
      <button type="button" class="list-toggle-btn" id="btnSaveList2" style="flex:1;padding:7px 6px;font-size:12px;text-align:center">➕ 存為穩健</button>
      <button type="button" class="list-toggle-btn" id="btnSaveList3" style="flex:1;padding:7px 6px;font-size:12px;text-align:center">➕ 存為彩券</button>
      <button type="button" class="list-toggle-btn" id="btnSaveList4" style="flex:1;padding:7px 6px;font-size:12px;text-align:center">➕ 存為清單4</button>
    </div>
  </div>
  <div class="field">
    <div class="range-row"><label style="margin:0">分析天數</label><span class="range-val" id="sDays">180 天</span></div>
    <input type="range" id="iDays" min="90" max="365" value="180" step="30">
  </div>
  <div class="field" style="display:flex;align-items:center;gap:8px;flex-direction:row">
    <input type="checkbox" id="iRealtime" checked style="width:auto;margin:0">
    <label for="iRealtime" style="margin:0;font-size:12px;color:#aaa;cursor:pointer">🔴 加入盤中即時股價（需FinMind sponsor會員，非sponsor會自動略過）</label>
  </div>
  <div class="field" style="display:flex;flex-direction:column;gap:6px;margin-top:6px">
    <div style="font-size:11px;color:#666">以下每勾一項，批次分析都會為每檔股票多打1次API，股票數多時會明顯變慢：</div>
    <div style="display:flex;align-items:center;gap:8px">
      <input type="checkbox" id="iShowPE" style="width:auto;margin:0">
      <label for="iShowPE" style="margin:0;font-size:12px;color:#aaa;cursor:pointer">📐 近3年P/E區間</label>
    </div>
    <div style="display:flex;align-items:center;gap:8px">
      <input type="checkbox" id="iShowRevenue" style="width:auto;margin:0">
      <label for="iShowRevenue" style="margin:0;font-size:12px;color:#aaa;cursor:pointer">📈 近3月營收YoY/MoM</label>
    </div>
    <div style="display:flex;align-items:center;gap:8px">
      <input type="checkbox" id="iShowPriceYoY" style="width:auto;margin:0">
      <label for="iShowPriceYoY" style="margin:0;font-size:12px;color:#aaa;cursor:pointer">💹 近3月均價YoY（可獨立勾選；若同時勾營收，月份會對齊營收那組）</label>
    </div>
    <div style="display:flex;align-items:center;gap:8px">
      <input type="checkbox" id="iShowInst" style="width:auto;margin:0">
      <label for="iShowInst" style="margin:0;font-size:12px;color:#aaa;cursor:pointer">🏦 三大法人買賣超（近3月，逐月加總）</label>
    </div>
  </div>
  <button class="btn" id="btnRun">🔍 批次分析</button>
  <button class="btn" id="btnTop100Analysis" style="background:linear-gradient(135deg,#ff6b1a,#d9430d);margin-top:8px">🔥 漲幅前100分析</button>
  <button class="btn" id="btnVolume100Analysis" style="background:linear-gradient(135deg,#1abc9c,#0d8c6d);margin-top:8px">📊 成交量前100分析</button>
  <div class="field" style="display:flex;align-items:center;gap:8px;flex-direction:row;margin-top:8px">
    <input type="checkbox" id="iAutoTrack" checked style="width:auto;margin:0">
    <label for="iAutoTrack" style="margin:0;font-size:12px;color:#aaa;cursor:pointer">📌 批次分析後自動把命中組合股票加入追蹤</label>
  </div>
  <button class="btn" id="btnTrack" style="background:linear-gradient(135deg,#ff9800,#c76b00);margin-top:4px">📈 命中組合追蹤</button>
  <div id="batchStatus" style="margin-top:10px;font-size:12px;color:#888;text-align:center"></div>

  <div style="margin-top:14px;padding:10px;border:1px solid rgba(255,255,255,.1);border-radius:8px;background:rgba(255,255,255,.02)">
    <div style="font-size:13px;font-weight:700;color:#ddd;margin-bottom:6px">🔬 歷史回測分析</div>
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
      <label for="iBtMonthsBack" style="margin:0;font-size:11px;color:#aaa">回測天數（月）：</label>
      <input type="number" id="iBtMonthsBack" value="3" min="1" max="36" step="1" style="width:60px;padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px">
    </div>
    <div style="font-size:11px;color:#888;margin-bottom:8px">
      回溯過去N個月，用「上市清單」重新計算每個交易日『當時』的DMI評分（沒有偷看未來），對照5/10/20日後實際報酬。月數愈大，全市場規模跑的時間愈長，且不是單純等比例——評估天數增加，每檔股票要重算的次數也跟著變多（型態辨識、回後買上漲每個評估點都要重跑一次），拉到36個月時，光是評估紀錄就可能超過70-80萬筆，瀏覽器記憶體用量會明顯增加，實際跑完可能要1小時以上，建議先從6-12個月試跑，確認可行再拉長。先不用資料庫，結果只存在這次頁面的記憶體裡，重新整理頁面就會消失。
    </div>
    <div style="font-size:11px;color:#888;margin-bottom:4px">型態確認、回後買上漲免費算。以下兩項各自多抓一組資料，會拉長時間：</div>
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
      <input type="checkbox" id="iBtIncludeDiv" style="width:auto;margin:0">
      <label for="iBtIncludeDiv" style="margin:0;font-size:11px;color:#aaa;cursor:pointer">📈 近3月YoY乖離度（股價歷史要抓超過1年）</label>
    </div>
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
      <input type="checkbox" id="iBtIncludeInst" style="width:auto;margin:0">
      <label for="iBtIncludeInst" style="margin:0;font-size:11px;color:#aaa;cursor:pointer">🏦 三大法人買賣超（近3月）</label>
    </div>
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
      <label for="iBtMinLiquidity" style="margin:0;font-size:11px;color:#aaa">最低近20日均成交金額（萬元）：</label>
      <input type="number" id="iBtMinLiquidity" value="0" min="0" step="100" style="width:80px;padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px">
    </div>
    <div style="font-size:11px;color:#888;margin-bottom:8px">
      排除成交量太小的殭屍股：每個評估點各自檢查「當時」往前20天的平均成交金額（收盤價×成交量，不是純股數，避免低價股股數虛胖），低於門檻就跳過那個評估點——不是只看現在，避免用現在的流動性去篩過去的資料。0＝不篩選。
    </div>
    <button class="btn" id="btnRunBacktest" style="background:linear-gradient(135deg,#7c3aed,#5b21b6)">🔬 執行歷史回測</button>
  </div>
  <hr>
  <div class="field">
    <label>OpenAI API Key（選填）</label>
    <input type="password" id="iOpenAIKey" placeholder="sk-...">
    <div style="font-size:11px;color:#777;margin-top:5px;line-height:1.5">用於「🤖 AI 智能綜合分析」，金鑰只會從您的瀏覽器直接呼叫 OpenAI，不會被儲存或上傳到任何伺服器。</div>
  </div>
  <div class="field">
    <label>AI 模型</label>
    <select id="iOpenAIModel" style="width:100%;padding:9px 12px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:8px;color:#fff;font-size:13px;outline:none">
      <option value="gpt-4o-mini">gpt-4o-mini（快速／經濟）</option>
      <option value="gpt-4o">gpt-4o（進階／較貴）</option>
    </select>
  </div>
  <hr>
  <div class="legend-h">評分維度各25分</div>
  <div class="legend-i">📈 趨勢分析（轉折波）</div>
  <div class="legend-i">🕯️ K線型態分析</div>
  <div class="legend-i">📊 均線系統分析</div>
  <div class="legend-i">📦 成交量分析</div>
  <div class="legend-h">進場判斷</div>
  <div class="legend-i">🟢 80+ 積極做多</div>
  <div class="legend-i">🔵 65-79 可考慮進場</div>
  <div class="legend-i">🟡 50-64 觀望</div>
  <div class="legend-i">🔴 &lt;50 不適合進場</div>
</aside>
<main class="main">
  <div class="page-title">📊 技術分析全攻略 · 個股評分分析系統</div>
  <div class="page-sub">個股評分改為DMI趨向指標（+DI／-DI／ADX／ADXR）多方力道評分；回後買上漲、15種進場型態辨識仍沿用朱家泓《技術分析全攻略》方法論</div>
  <div id="output">
    <div class="welcome">
      <div class="welcome-icon">📈</div>
      <div class="welcome-t">請輸入分析條件</div>
      <div class="welcome-d">在左側輸入 FinMind API Token 與股票代號，點擊「開始分析」即可開始。</div>
    </div>
  </div>
  <div id="trackOutput" style="display:none;margin-top:24px"></div>
  <div id="backtestOutput" style="display:none;margin-top:24px"></div>
</main>
</div>
<script>
(function() {
  // ── DOM refs ──
  var iToken = document.getElementById('iToken');
  var batchStatus = document.getElementById('batchStatus');
  var btnTop100Analysis = document.getElementById('btnTop100Analysis');
  var btnVolume100Analysis = document.getElementById('btnVolume100Analysis');
  var iStock = document.getElementById('iStock');
  var iDays  = document.getElementById('iDays');
  var sDays  = document.getElementById('sDays');
  var btnRun = document.getElementById('btnRun');
  var output = document.getElementById('output');
  var iOpenAIKey   = document.getElementById('iOpenAIKey');
  var iOpenAIModel = document.getElementById('iOpenAIModel');
  var iStocks = document.getElementById('iStocks');
  var btnListTWSE = document.getElementById('btnListTWSE');
  var btnListTPEX = document.getElementById('btnListTPEX');
  function clearAllListToggleActive() {
    document.querySelectorAll('.list-toggle-btn[data-list]').forEach(function(el){ el.classList.remove('active'); });
  }

  // ── 預設股票清單（上市／上櫃）──
  var TWSE_LIST = ['1101','1101B','1102','1103','1104','1108','1109','1110','1201','1203','1210','1213','1215','1216','1217','1218','1219','1220','1225','1227','1229','1231','1232','1233','1234','1235','1236','1256','1301','1303','1304','1305','1307','1308','1309','1310','1312','1312A','1313','1314','1315','1316','1319','1321','1323','1324','1325','1326','1337','1338','1339','1340','1341','1342','1402','1409','1410','1413','1414','1416','1417','1418','1419','1423','1432','1434','1435','1436','1437','1438','1439','1440','1441','1442','1443','1444','1445','1446','1447','1449','1451','1452','1453','1454','1455','1456','1457','1459','1460','1463','1464','1465','1466','1467','1468','1470','1471','1472','1473','1474','1475','1476','1477','1503','1504','1506','1512','1513','1514','1515','1516','1517','1519','1521','1522','1522A','1524','1525','1526','1527','1528','1529','1530','1531','1532','1533','1535','1536','1537','1538','1539','1540','1541','1558','1560','1563','1568','1582','1583','1587','1590','1597','1598','1603','1604','1605','1608','1609','1611','1612','1614','1615','1616','1617','1618','1623','1626','1702','1707','1708','1709','1710','1711','1712','1713','1714','1717','1718','1720','1721','1722','1723','1725','1726','1727','1730','1731','1732','1733','1734','1735','1736','1737','1752','1760','1762','1773','1776','1783','1786','1789','1795','1802','1805','1806','1808','1809','1810','1817','1903','1904','1905','1906','1907','1909','2002','2002A','2006','2007','2008','2009','2010','2012','2013','2014','2015','2017','2020','2022','2023','2024','2025','2027','2028','2029','2030','2031','2032','2033','2034','2038','2049','2059','2062','2069','2072','2101','2102','2103','2104','2105','2106','2107','2108','2109','2114','2115','2201','2204','2206','2207','2208','2211','2227','2228','2231','2233','2236','2239','2241','2243','2247','2248','2250','2254','2258','2301','2302','2303','2305','2308','2312','2313','2314','2316','2317','2321','2323','2324','2327','2328','2329','2330','2331','2332','2337','2338','2340','2342','2344','2345','2347','2348','2348A','2349','2351','2352','2353','2354','2355','2356','2357','2359','2360','2362','2363','2364','2365','2367','2368','2369','2371','2373','2374','2375','2376','2377','2379','2382','2383','2385','2387','2388','2390','2392','2393','2395','2397','2399','2401','2402','2404','2405','2406','2408','2409','2412','2413','2414','2415','2417','2419','2420','2421','2423','2424','2425','2426','2427','2428','2429','2430','2431','2432','2433','2434','2436','2438','2439','2440','2441','2442','2444','2449','2450','2451','2453','2454','2455','2457','2458','2459','2460','2461','2462','2464','2465','2466','2467','2468','2471','2472','2474','2476','2477','2478','2480','2481','2482','2483','2484','2485','2486','2488','2489','2491','2492','2493','2495','2496','2497','2498','2501','2504','2505','2506','2509','2511','2514','2515','2516','2520','2524','2527','2528','2530','2534','2535','2536','2537','2538','2539','2540','2542','2543','2545','2546','2547','2548','2597','2601','2603','2605','2606','2607','2608','2609','2610','2611','2612','2613','2614','2615','2616','2617','2618','2630','2633','2634','2636','2637','2642','2645','2646','2701','2702','2704','2705','2706','2707','2712','2722','2723','2727','2731','2739','2748','2753','2762','2801','2812','2816','2820','2832','2834','2836','2836A','2838','2838A','2845','2849','2850','2851','2852','2855','2867','2880','2881','2881A','2881B','2881C','2882','2882A','2882B','2883','2883B','2884','2885','2886','2887','2887E','2887F','2887G','2887H','2887I','2887Z1','2889','2890','2891','2891B','2891C','2892','2897','2897B','2901','2903','2904','2905','2906','2908','2910','2911','2912','2913','2915','2923','2929','2939','2945','3002','3003','3004','3005','3006','3008','3010','3011','3013','3014','3015','3016','3017','3018','3019','3021','3022','3023','3024','3025','3026','3027','3028','3029','3030','3031','3032','3033','3034','3035','3036','3037','3038','3040','3041','3042','3043','3044','3045','3046','3047','3048','3049','3050','3051','3052','3054','3055','3056','3057','3058','3059','3060','3062','3090','3092','3094','3130','3135','3138','3149','3150','3164','3167','3168','3189','3209','3229','3231','3257','3266','3296','3305','3308','3311','3312','3321','3338','3346','3356','3376','3380','3406','3413','3416','3419','3432','3437','3443','3447','3450','3481','3494','3501','3504','3515','3518','3528','3530','3532','3533','3535','3543','3545','3550','3557','3563','3576','3583','3588','3591','3592','3593','3596','3605','3607','3617','3622','3645','3652','3653','3661','3665','3669','3673','3679','3686','3694','3701','3702','3703','3704','3705','3706','3708','3711','3712','3714','3715','3716','3717','4104','4106','4108','4119','4133','4137','4142','4148','4155','4164','4169','4178','4190','4195','4306','4414','4426','4438','4439','4440','4441','4526','4532','4536','4540','4545','4551','4552','4555','4557','4560','4562','4564','4566','4569','4571','4572','4576','4581','4582','4583','4585','4588','4590','4720','4722','4736','4737','4739','4746','4755','4763','4764','4766','4770','4771','4807','4904','4906','4912','4915','4916','4919','4927','4930','4934','4935','4938','4942','4943','4949','4952','4956','4958','4960','4961','4967','4968','4976','4977','4989','4994','4999','5007','5203','5215','5222','5225','5234','5243','5244','5258','5269','5283','5284','5285','5288','5292','5306','5388','5434','5469','5471','5484','5515','5519','5521','5522','5525','5531','5533','5534','5538','5546','5607','5608','5706','5871','5871A','5876','5880','5906','5907','6005','6024','6108','6112','6115','6116','6117','6120','6128','6133','6136','6139','6141','6142','6152','6153','6155','6164','6165','6166','6168','6176','6177','6183','6184','6189','6191','6192','6196','6197','6201','6202','6205','6206','6209','6213','6214','6215','6216','6224','6225','6226','6230','6235','6239','6243','6257','6269','6271','6272','6277','6278','6281','6282','6283','6285','6405','6409','6412','6414','6415','6416','6426','6431','6438','6442','6443','6446','6449','6451','6456','6464','6472','6477','6491','6504','6505','6515','6525','6526','6531','6533','6534','6541','6550','6552','6558','6573','6579','6581','6582','6585','6589','6591','6592','6592A','6592B','6598','6605','6606','6614','6625','6641','6645','6655','6657','6658','6666','6668','6669','6670','6671','6672','6674','6689','6691','6695','6698','6706','6715','6719','6722','6742','6743','6753','6754','6756','6757','6768','6770','6771','6776','6781','6782','6789','6790','6792','6794','6796','6799','6805','6807','6830','6831','6834','6835','6838','6854','6861','6862','6863','6869','6873','6885','6887','6890','6901','6902','6906','6908','6909','6914','6916','6918','6919','6921','6923','6924','6928','6931','6933','6934','6936','6937','6944','6949','6951','6952','6955','6957','6958','6958A','6962','6965','6969','6988','6994','7610','7631','7705','7711','7721','7722','7730','7732','7736','7740','7749','7750','7760','7765','7768','7769','7780','7786','7788','7791','7795','7799','7803','7818','7821','7822','7823','7827','8011','8016','8021','8028','8033','8039','8045','8046','8070','8072','8081','8101','8103','8104','8105','8110','8112','8112A','8114','8131','8150','8162','8163','8201','8210','8213','8215','8222','8249','8261','8271','8341','8367','8374','8404','8411','8422','8429','8438','8442','8443','8454','8462','8463','8464','8466','8467','8473','8476','8478','8481','8482','8487','8488','8499','8926','8940','8996','9103','910322','9105','910861','9110','911608','911622','911868','912000','9136','9802','9902','9904','9905','9906','9907','9908','9910','9911','9912','9914','9917','9918','9919','9921','9924','9925','9926','9927','9928','9929','9930','9931','9933','9934','9935','9937','9938','9939','9940','9941','9941A','9942','9943','9944','9945','9946','9955','9958'];
  var TPEX_LIST = ['1240','1259','1264','1268','1294','1295','1336','1565','1569','1570','1580','1584','1586','1591','1593','1595','1599','1742','1777','1780','1781','1784','1785','1788','1796','1799','1813','1815','2035','2061','2063','2064','2065','2066','2067','2070','2073','2221','2230','2235','2596','2640','2641','2643','2718','2719','2724','2726','2729','2732','2734','2736','2740','2743','2745','2751','2752','2754','2755','2756','2916','2924','2926','2937','2941','2947','2948','2949','3064','3066','3067','3071','3073','3078','3081','3083','3085','3086','3088','3093','3095','3105','3114','3115','3118','3122','3128','3131','3141','3147','3152','3158','3162','3163','3169','3171','3176','3178','3188','3191','3205','3206','3207','3211','3213','3217','3218','3219','3221','3224','3226','3227','3228','3230','3232','3234','3236','3252','3259','3260','3264','3265','3268','3272','3276','3284','3285','3287','3288','3289','3290','3293','3294','3297','3303','3306','3310','3313','3317','3322','3323','3324','3325','3332','3339','3349','3354','3357','3360','3362','3363','3372','3373','3374','3379','3388','3390','3402','3430','3434','3438','3441','3444','3455','3465','3466','3467','3479','3483','3484','3485','3489','3490','3491','3492','3498','3499','3508','3511','3512','3516','3520','3521','3522','3523','3526','3527','3529','3531','3537','3540','3541','3546','3548','3551','3552','3555','3556','3558','3564','3567','3570','3577','3580','3581','3587','3594','3597','3609','3611','3615','3623','3624','3625','3628','3629','3630','3631','3632','3646','3663','3664','3666','3672','3675','3680','3684','3685','3687','3689','3691','3693','3707','3709','3710','3713','4102','4105','4107','4109','4111','4113','4114','4116','4120','4121','4123','4126','4127','4128','4129','4131','4138','4139','4147','4153','4154','4157','4160','4161','4162','4163','4166','4167','4168','4171','4173','4174','4175','4183','4188','4192','4198','4205','4207','4303','4304','4305','4401','4402','4406','4413','4416','4417','4419','4420','4430','4432','4433','4442','4502','4503','4506','4510','4513','4523','4527','4528','4529','4530','4533','4534','4535','4538','4541','4542','4543','4549','4550','4554','4556','4558','4561','4563','4568','4577','4580','4584','4609','4702','4706','4707','4711','4714','4716','4721','4726','4728','4729','4735','4741','4743','4744','4745','4747','4749','4754','4760','4767','4768','4772','4806','4903','4905','4907','4908','4909','4911','4923','4924','4931','4933','4939','4946','4950','4951','4953','4966','4971','4972','4973','4974','4979','4991','4995','5009','5011','5013','5014','5015','5016','5201','5202','5205','5206','5209','5210','5211','5212','5213','5220','5223','5227','5228','5230','5245','5251','5263','5272','5274','5276','5278','5287','5289','5291','5299','5301','5302','5309','5310','5312','5314','5315','5321','5324','5328','5340','5344','5345','5347','5348','5351','5353','5355','5356','5364','5371','5381','5386','5392','5398','5403','5410','5425','5426','5432','5438','5439','5443','5450','5452','5455','5457','5460','5464','5465','5468','5474','5475','5478','5481','5483','5487','5488','5489','5490','5493','5498','5508','5511','5512','5514','5516','5520','5523','5529','5530','5536','5543','5547','5548','5601','5603','5604','5609','5701','5703','5704','5864','5878','5902','5903','5904','5905','6015','6016','6020','6021','6023','6026','6028','6101','6103','6104','6109','6111','6113','6114','6118','6121','6122','6123','6124','6125','6126','6127','6129','6130','6134','6138','6140','6143','6144','6146','6147','6148','6150','6151','6154','6156','6158','6160','6161','6163','6167','6169','6170','6171','6173','6174','6175','6179','6180','6182','6185','6186','6187','6188','6190','6194','6195','6198','6199','6203','6204','6207','6208','6210','6212','6217','6218','6219','6220','6221','6222','6223','6227','6228','6229','6231','6233','6234','6236','6237','6240','6241','6242','6244','6245','6246','6248','6259','6261','6263','6264','6265','6266','6270','6274','6275','6276','6279','6284','6290','6291','6292','6294','6411','6417','6418','6419','6423','6425','6432','6435','6441','6461','6462','6465','6469','6470','6474','6482','6485','6486','6488','6492','6494','6496','6498','6499','6506','6508','6509','6510','6512','6516','6517','6523','6527','6530','6532','6535','6538','6542','6546','6547','6548','6556','6560','6561','6568','6569','6570','6574','6576','6577','6578','6584','6588','6590','6593','6596','6597','6603','6609','6612','6613','6615','6616','6617','6620','6624','6629','6637','6640','6642','6643','6649','6651','6654','6661','6662','6664','6667','6679','6680','6683','6684','6690','6692','6693','6697','6703','6708','6712','6716','6720','6721','6725','6727','6728','6730','6732','6733','6735','6739','6741','6751','6752','6761','6762','6763','6767','6785','6788','6791','6803','6804','6811','6821','6823','6829','6840','6841','6843','6844','6846','6855','6856','6859','6865','6870','6872','6874','6875','6877','6881','6884','6894','6895','6899','6903','6904','6907','6910','6913','6922','6925','6929','6945','6953','6961','6967','6968','6971','6982','6983','6986','6996','6997','7402','7547','7556','7584','7642','7703','7704','7708','7709','7712','7713','7714','7715','7716','7717','7718','7723','7728','7734','7738','7743','7744','7747','7751','7753','7757','7767','7770','7772','7777','7782','7792','7794','7805','7810','7811','7814','7819','7820','7828','7839','7842','8024','8027','8032','8034','8038','8040','8042','8043','8044','8047','8048','8049','8050','8054','8059','8064','8066','8067','8068','8069','8071','8074','8076','8077','8080','8083','8084','8085','8086','8087','8088','8089','8091','8092','8093','8096','8097','8099','8102','8107','8109','8111','8121','8147','8155','8171','8176','8182','8183','8227','8234','8240','8255','8272','8277','8279','8284','8289','8291','8299','8342','8349','8349A','8354','8358','8383','8390','8401','8403','8409','8410','8415','8416','8421','8423','8424','8426','8431','8432','8433','8435','8436','8437','8440','8444','8446','8450','8455','8472','8477','8489','8905','8906','8908','8916','8917','8921','8923','8924','8927','8928','8929','8930','8931','8932','8933','8935','8936','8937','8938','8941','8942','9949','9950','9951','9960','9962'];
  var MY_LIST_DEFAULT = ['4979','6870','6451','3450','3163','4977','3363','6533','6757','2345','2376','2441','6285','2049','2303','3653','2467','2359','2308','3017','2368','4576','2327','6442','6531','6683','6257','6223','8150','3105','2634','8299','3532','6213','3033','4991','2449','6781','3189','8046','3037','3443','3491','2481','6770','2408','8271','3081','3665','6274','2455','2454','2360','3673','8021','4573','2351','1560','8028','6197','4958','8358','2330','2383','1303','6669','2137','3231','1815','5475','5340','3771','1519','3661','2059','6805','6510','3211','4931','8210','3013','6117','3693'];

  // 0050（元大台灣50）成分股：追蹤「臺灣50指數」，每季（3/6/9/12月）審核調整一次，
  // 是規則化的被動指數，成分股清單相對穩定、可公開查證，適合寫死維護。
  // 以下為 2026-09-04 公開持股快照（共51檔），下次調整約在2026年12月，屆時建議
  // 重新核對更新。注意：00981A／00991A是主動式ETF（經理人自由調整持股、無固定
  // 規則、且無完整成分股API），沒有一併做在這裡——硬編碼一份很快就會失準，
  // 如果你有這兩檔ETF官方月報揭露的完整持股清單，貼給我可以直接編進去。
  var TW0050_LIST = ['2330','2454','2308','2317','3711','2383','2303','2881','2891','3037','3017','1303','2882','2345','2887','2382','2327','2885','2357','2884','3008','2301','3231','2886','2883','2408','2890','2344','2412','2892','2449','2337','4938','2356','2368','5880','1301','1326','2002','1101','1216','3045','4904','2379','3034','1102','2610','2618','2603','2609','2615'];

  // ── 自訂清單（我的清單／清單1／清單2／清單3）：存在瀏覽器 localStorage，
  // 「我的清單」預設帶入原本內建的觀察名單（MY_LIST_DEFAULT），清單1/2/3預設空白。
  // 兩個動作：「更新我的清單」把目前批次股票代號的內容存回我的清單；
  // 「存為清單1／清單2／清單3」是三個各自獨立的按鈕，直接存進指定槽位（若該槽已有
  // 內容會先跳確認再覆蓋，不再是舊版那種自動找空槽的邏輯）。
  var CUSTOM_LIST_KEYS = { my:'customList_my', my1:'customList_1', my2:'customList_2', my3:'customList_3', my4:'customList_4' };
  // 清單1/2/3 改名為三種組合型態清單（批次分析摘要可一鍵加入），另加我的清單4
  var CUSTOM_LIST_NAMES = { my:'我的清單', my1:'進攻組合型態', my2:'穩健組合型態', my3:'彩券組合型態', my4:'我的清單4' };
  var STYLE_LIST_KIND = { '進攻':'my1', '穩健':'my2', '彩券':'my3' };
  function addToCustomList(kind, ids) {
    var cur = getCustomList(kind), seen = {}, added = 0;
    cur.forEach(function(x) { seen[x] = 1; });
    ids.forEach(function(x) { if (!seen[x]) { seen[x] = 1; cur.push(x); added++; } });
    setCustomList(kind, cur);
    return [added, cur.length];
  }

  function getCustomList(kind) {
    try {
      var raw = localStorage.getItem(CUSTOM_LIST_KEYS[kind]);
      if (raw === null) return kind === 'my' ? MY_LIST_DEFAULT.slice() : [];
      var arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch (e) { return kind === 'my' ? MY_LIST_DEFAULT.slice() : []; }
  }
  function setCustomList(kind, arr) {
    try { localStorage.setItem(CUSTOM_LIST_KEYS[kind], JSON.stringify(arr)); } catch (e) {}
  }
  function parseStocksFromTextarea() {
    return iStocks.value.split(/[\n,，、\s]+/).map(function(s){ return s.trim(); }).filter(Boolean);
  }
  function refreshCustomListLabels() {
    ['my','my1','my2','my3','my4'].forEach(function(kind) {
      var labelEl = document.querySelector('.list-toggle-btn[data-list="'+kind+'"] .lt-label');
      if (!labelEl) return;
      var n = getCustomList(kind).length;
      var base = CUSTOM_LIST_NAMES[kind];
      labelEl.textContent = base + (n ? '(' + n + ')' : '');
    });
  }

  function loadStockList(kind) {
    var list = kind === 'tpex' ? TPEX_LIST : (kind === 'twse' ? TWSE_LIST : (kind === '0050' ? TW0050_LIST : getCustomList(kind)));
    iStocks.value = list.join('\n');
    document.querySelectorAll('.list-toggle-btn[data-list]').forEach(function(el) {
      el.classList.toggle('active', el.dataset.list === kind);
    });
  }

  // Event delegation：整排清單按鈕（上市／上櫃／我的清單／清單1-3）
  document.querySelectorAll('.list-toggle-btn[data-list]').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      if (e.target.closest('.lt-clear')) return; // × 交給下面另一個監聽器處理
      loadStockList(btn.dataset.list);
    });
  });
  // × 清除單一自訂清單（不影響上市／上櫃這種固定清單，那兩個按鈕本來就沒有 .lt-clear）
  document.querySelectorAll('.lt-clear').forEach(function(x) {
    x.addEventListener('click', function(e) {
      e.stopPropagation();
      var kind = x.dataset.clear;
      if (!confirm('確定要清除「' + CUSTOM_LIST_NAMES[kind] + '」嗎？')) return;
      setCustomList(kind, []);
      refreshCustomListLabels();
      var activeBtn = document.querySelector('.list-toggle-btn[data-list="'+kind+'"]');
      if (activeBtn && activeBtn.classList.contains('active')) iStocks.value = '';
    });
  });

  var btnSaveMyList = document.getElementById('btnSaveMyList');
  btnSaveMyList.addEventListener('click', function() {
    var stocks = parseStocksFromTextarea();
    if (!stocks.length) { alert('批次股票代號目前是空的，沒有可儲存的內容'); return; }
    setCustomList('my', stocks);
    refreshCustomListLabels();
    var oldTxt = btnSaveMyList.textContent;
    btnSaveMyList.textContent = '✅ 已更新（' + stocks.length + '檔）';
    setTimeout(function(){ btnSaveMyList.textContent = oldTxt; }, 1500);
  });
  // 存為清單1／清單2／清單3：各自獨立按鈕，直接存進指定槽位（不再自動找空槽）。
  // 若該槽已有內容會先跳確認，避免手滑蓋掉舊清單。
  [1, 2, 3, 4].forEach(function(n) {
    var btn = document.getElementById('btnSaveList' + n);
    var key = 'my' + n;
    btn.addEventListener('click', function() {
      var stocks = parseStocksFromTextarea();
      if (!stocks.length) { alert('批次股票代號目前是空的，沒有可儲存的內容'); return; }
      var existing = getCustomList(key);
      if (existing.length && !confirm('「' + CUSTOM_LIST_NAMES['my' + n] + '」目前已有 ' + existing.length + ' 檔，確定要覆蓋嗎？')) return;
      setCustomList(key, stocks);
      refreshCustomListLabels();
      var oldTxt = btn.textContent;
      btn.textContent = '✅ 已存入' + CUSTOM_LIST_NAMES['my' + n];
      setTimeout(function(){ btn.textContent = oldTxt; }, 1500);
    });
  });
  refreshCustomListLabels();
  loadStockList('twse');

  // ── 漲幅前100分析：直接抓今日全市場漲幅排行前100名個股，套入清單後自動接著跑批次分析。
  // 改用 FinMind 自己的 API（跟本工具其他功能同一個資料來源），不帶 data_id 直接查全市場當天資料，
  // 此為 Backer/Sponsor 方案才能使用的用法（免費方案只能查單一股票）。
  // 流程：①先測試 API Token 是否有效 → ②再查全市場當天資料 → ③依需求（漲幅／成交量）排序取前100名，
  // 每一步都直接把訊息寫進主畫面（跟「批次分析」按鈕同一個做法），確保不會因為看漏小字說明而以為「沒反應」。
  // 「漲幅前100」「成交量前100」共用同一份全市場快照，只是排序依據不同，所以抽成共用函式避免重複打兩次API。
  async function fetchTaiwanMarketSnapshot() {
    var token = iToken.value.trim();
    if (!token) {
      showError('請先在上方輸入 FinMind API Token，才能抓取今日全市場資料（需 Backer/Sponsor 方案）');
      return null;
    }

    // ① 測試 API Token 是否有效（跟「批次分析」按鈕同一套測試方式）
    output.innerHTML = '<div class="batch-progress"><div style="font-size:14px;color:#ccc">🔌 測試 FinMind API Token 中...</div></div>';
    var testJson;
    try {
      testJson = await apiFetch('https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo&token=' + token);
      if (testJson.status !== 200) {
        showError('Token 無效或 API 錯誤：' + (testJson.msg || 'status=' + testJson.status) + '\n請確認 FinMind Token 是否正確。');
        return null;
      }
    } catch (connErr) {
      showError('❌ 無法連線至 FinMind API\n\n錯誤：' + connErr.message + '\n\n請確認：\n1. Token 是否正確\n2. 網路連線是否正常\n3. 是否從本機伺服器開啟（非 file:// 協定）');
      return null;
    }
    var nameJson = testJson;

    // ② 查全市場當天資料（往前找最近一個有資料的交易日，最多找7天，逐日獨立重試、單日失敗不影響其他天）
    var priceJson = null, usedDate = null, lastDayErr = '';
    for (var back = 0; back < 7 && !priceJson; back++) {
      var d = new Date();
      d.setDate(d.getDate() - back);
      var dateStr = d.toISOString().slice(0, 10);
      output.innerHTML = '<div class="batch-progress"><div style="font-size:14px;color:#ccc">📡 查詢 ' + dateStr + ' 全市場行情中...（此功能需 FinMind Backer/Sponsor 方案）</div></div>';
      try {
        var res = await apiFetch('https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&start_date=' + dateStr + '&token=' + token);
        if (res.status === 200 && res.data && res.data.length) { priceJson = res; usedDate = dateStr; }
        else { lastDayErr = res.msg || ('status=' + res.status); }
      } catch (dayErr) { lastDayErr = dayErr.message; }
    }
    if (!priceJson) {
      showError('❌ 近7天皆無法取得全市場資料\n\n錯誤：' + lastDayErr + '\n\n這個功能需要 FinMind「Backer」或「Sponsor」付費方案才能查詢全市場單日資料，免費方案只能查詢單一股票。\n請確認您的帳號方案，或改用「上市清單／上櫃清單」搭配「批次分析」。');
      return null;
    }

    // ③ 股票名稱對照（用①已經拿到的 TaiwanStockInfo 結果，不用再查一次）
    var nameMap = {};
    (nameJson.data || []).forEach(function(r) { nameMap[r.stock_id] = r.stock_name; });

    // 只保留出現在本工具「上市清單」「上櫃清單」裡的代碼——這兩份清單本來就是精選過的
    // 一般個股代碼，直接拿來當白名單，能同時排除ETF、權證、TDR、公司債等各種非個股商品，
    // 比單純判斷代碼開頭字元（例如「00」開頭）可靠很多。
    var validStockSet = {};
    TWSE_LIST.forEach(function(s) { validStockSet[s] = true; });
    TPEX_LIST.forEach(function(s) { validStockSet[s] = true; });

    var results = [];
    priceJson.data.forEach(function(r) {
      var sid = r.stock_id;
      if (!sid || !validStockSet[sid]) return;
      var close = parseFloat(r.close), spread = parseFloat(r.spread);
      if (!isFinite(close) || !isFinite(spread) || close <= 0) return;
      var prevClose = close - spread;
      if (prevClose <= 0) return;
      var volume = parseFloat(r.Trading_Volume);
      results.push({ id: sid, name: nameMap[sid] || sid, pct: (spread / prevClose) * 100, volume: isFinite(volume) ? volume : 0 });
    });
    if (!results.length) {
      showError('取得的全市場資料是空的，請稍後再試一次。');
      return null;
    }
    return { results: results, usedDate: usedDate };
  }

  async function loadTop100Gainers() {
    batchStatus.textContent = '';
    var snap = await fetchTaiwanMarketSnapshot();
    if (!snap) return false;
    var top100 = snap.results.slice().sort(function(a, b) { return b.pct - a.pct; }).slice(0, 100);
    iStocks.value = top100.map(function(r) { return r.id; }).join('\n');
    clearAllListToggleActive();
    batchStatus.textContent = '✅ 已套入 ' + snap.usedDate + ' 漲幅前' + top100.length + '（最高：' + top100[0].id + ' ' + top100[0].name + ' +' + top100[0].pct.toFixed(2) + '%）';
    return true;
  }

  async function loadTop100Volume() {
    batchStatus.textContent = '';
    var snap = await fetchTaiwanMarketSnapshot();
    if (!snap) return false;
    var top100 = snap.results.slice().sort(function(a, b) { return b.volume - a.volume; }).slice(0, 100);
    iStocks.value = top100.map(function(r) { return r.id; }).join('\n');
    clearAllListToggleActive();
    batchStatus.textContent = '✅ 已套入 ' + snap.usedDate + ' 成交量前' + top100.length + '（最高：' + top100[0].id + ' ' + top100[0].name + ' 量=' + Math.round(top100[0].volume).toLocaleString() + ' 股）';
    return true;
  }

  async function runListAnalysisButton(btn, loadFn) {
    btn.disabled = true;
    var origText = btn.textContent;
    btn.textContent = '⏳ 抓取中...';
    try {
      var ok = await loadFn();
      if (ok) {
        // 套入清單成功後，自動接著跑批次分析
        await batchAnalysis().catch(function(e){ showError(e.message); });
      }
    } finally {
      btn.disabled = false;
      btn.textContent = origText;
    }
  }
  btnTop100Analysis.addEventListener('click', function(){ runListAnalysisButton(btnTop100Analysis, loadTop100Gainers); });
  btnVolume100Analysis.addEventListener('click', function(){ runListAnalysisButton(btnVolume100Analysis, loadTop100Volume); });


  iDays.addEventListener('input', function() { sDays.textContent = this.value + ' 天'; });
  btnRun.addEventListener('click', function() { batchAnalysis().catch(function(e){ showError(e.message); }); });

  var btnRunBacktest = document.getElementById('btnRunBacktest');
  var backtestOutput = document.getElementById('backtestOutput');
  var gridSearchHorizon = 10;   // 參數網格搜尋目前選的天數，下拉選單切換用
  var comboSearchHorizon = 10;  // 多因子複選搜尋目前選的天數，下拉選單切換用
  var comboMinWinRate = 50;
  var comboMinT = 1;
  var comboBothHalves = false;   // 只列前後半段 t 值都 ≥2 的組合   // 多因子複選搜尋：排除 t值(同日調整) 低於此門檻的組合
  var comboSortBy = 't';   // 多因子複選搜尋排序依據：'t'＝t值(同日調整)、'win'＝勝率     // 多因子複選搜尋目前選的勝率門檻(%)——只列出達標的組合，不是固定前幾名
  var moonshotHorizon = 10;     // 飆股搜尋目前選的天數
  var moonshotThreshold = 30;   // 飆股搜尋目前選的漲幅門檻(%)
  var moonshotMinPct = 10;      // 飆股搜尋目前選的飆股比例門檻(%)——只列出達標的組合
  var skipRedundantCombos = true; // 排除冗餘組合（多加條件後樣本完全沒變）

  btnRunBacktest.addEventListener('click', function() {
    var token = iToken.value.trim();
    if (!token) { alert('請先輸入 FinMind API Token'); return; }
    var monthsBack = parseInt(document.getElementById('iBtMonthsBack').value, 10) || 3;
    monthsBack = Math.max(1, Math.min(36, monthsBack));
    var includeDiv = document.getElementById('iBtIncludeDiv').checked;
    var includeInst = document.getElementById('iBtIncludeInst').checked;
    var minLiquidityWan = parseFloat(document.getElementById('iBtMinLiquidity').value) || 0;
    var minLiquidity = minLiquidityWan * 10000; // 萬元轉成元，跟close*volume算出來的單位一致

    btnRunBacktest.disabled = true;
    var oldTxt = btnRunBacktest.textContent;
    backtestOutput.style.display = 'block';
    backtestOutput.innerHTML = '<div class="batch-progress">' +
      '<div style="font-size:14px;color:#ccc">🔬 歷史回測執行中（過去' + monthsBack + '個月）：<strong id="btProgText">0 / ' + TWSE_LIST.length + '</strong></div>' +
      '<div class="progress-bar-wrap"><div class="progress-bar-fill" id="btProgBar" style="width:0%"></div></div>' +
      '<div style="font-size:12px;color:#888;margin-top:6px" id="btProgDetail"></div>' +
      '</div>';
    backtestOutput.scrollIntoView({ behavior: 'smooth', block: 'start' });

    var btProgText = document.getElementById('btProgText');
    var btProgBar = document.getElementById('btProgBar');
    var btProgDetail = document.getElementById('btProgDetail');
    finmindQuotaNote = function(msg) { btProgDetail.textContent = msg; };

    runHistoricalBacktestHtml(token, monthsBack, includeDiv, includeInst, minLiquidity, function(i, total, saved, failed, sid) {
      btProgText.textContent = i + ' / ' + total;
      btProgBar.style.width = (i / total * 100) + '%';
      btProgDetail.textContent = sid ? ('目前：' + sid + '　已存 ' + saved.toLocaleString() + ' 筆　失敗 ' + failed + ' 檔') :
        ('完成　已存 ' + saved.toLocaleString() + ' 筆　失敗 ' + failed + ' 檔');
    }).then(function(result) {
      finmindQuotaNote = null;
      btnRunBacktest.disabled = false;
      btnRunBacktest.textContent = oldTxt;
      renderBacktestResults();
    }).catch(function(e) {
      finmindQuotaNote = null;
      btnRunBacktest.disabled = false;
      btnRunBacktest.textContent = oldTxt;
      backtestOutput.innerHTML = '<div class="err">❌ 歷史回測失敗：' + e.message + '</div>';
    });
  });

  function renderBacktestResults() {
    if (!backtestResults.length) {
      backtestOutput.innerHTML = '<div style="color:#888;font-size:13px">這次回測沒有產生任何評估紀錄。</div>';
      return;
    }
    var dates = backtestResults.map(function(r){ return r.evalDate; }).sort();
    var html = '<div class="sec" style="font-size:18px">🔬 歷史回測分析結果</div>' +
      '<div style="color:#888;font-size:12px;margin-bottom:16px">共 ' + backtestResults.length.toLocaleString() +
      ' 筆評估紀錄（' + dates[0] + ' ～ ' + dates[dates.length-1] + '）</div>';

    html += '<div class="sec">📏 全體基準（所有評估紀錄）</div>' +
      '<div style="color:#888;font-size:12px;margin-bottom:8px">以下各訊號都跟這一列比，高於基準才代表訊號有用。</div>' +
      objArrayToTable([buildStatRow('標記', '全體基準', backtestResults)]);
    html += '<div class="sec">📊 評分區間 vs 實際報酬</div>' + objArrayToTable(analyzeScoreBuckets(backtestResults));
    if (backtestResults.some(function(r){ return r.confluenceCount != null; })) {
      html += '<div class="sec">🧮 訊號堆疊數量 vs 實際報酬（Confluence Count）</div>' +
        '<div style="color:#888;font-size:12px;margin-bottom:8px">不是窮舉「哪幾個條件」的組合，而是算「同一個時間點總共同時觸發了幾個獨立訊號」（型態剛形成、KDJ交叉、MACD黃金交叉、布林通道極端位置、爆量，最多5個），測試「訊號數量越多，勝率是否線性遞增」。如果各區間勝率大致隨數量遞增，代表這套系統的優勢比較像是「訊號密度」帶來的，不是特定某個巧合組合。</div>' +
        objArrayToTable(analyzeConfluenceCount(backtestResults));
    }
    html += '<div class="sec">🚀 「強勢突破盤」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'isBreakout', '強勢突破盤'));
    html += '<div class="sec">🎯 「跌深反彈盤」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'isPullbackRebound', '跌深反彈盤'));
    html += '<div class="sec">⚡ 「MACD近3日內黃金交叉」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'goldenCrossRecent', 'MACD黃金交叉'));
    if (backtestResults.some(function(r){ return r.kdjGoldenCrossRecent != null; })) {
      html += '<div class="sec">🟢 「KDJ近3日內黃金交叉」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'kdjGoldenCrossRecent', 'KDJ黃金交叉'));
    }
    if (backtestResults.some(function(r){ return r.kdjDeathCrossRecent != null; })) {
      html += '<div class="sec">🔴 「KDJ近3日內死亡交叉」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'kdjDeathCrossRecent', 'KDJ死亡交叉'));
    }
    if (backtestResults.some(function(r){ return r.relStrengthPositive != null; })) {
      html += '<div class="sec">💪 「相對強弱(vs大盤0050，20日)」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'relStrengthPositive', '相對強弱為正'));
    }
    if (backtestResults.some(function(r){ return r.volSurge15 != null; })) {
      html += '<div class="sec">📊 「爆量(≥1.5倍均量)」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'volSurge15', '爆量1.5倍'));
    }
    if (backtestResults.some(function(r){ return r.volSurge2 != null; })) {
      html += '<div class="sec">📊 「爆量(≥2倍均量)」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'volSurge2', '爆量2倍'));
    }
    if (backtestResults.some(function(r){ return r.volLow05 != null; })) {
      html += '<div class="sec">📉 「地量(≤0.5倍均量)」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'volLow05', '地量'));
    }
    if (backtestResults.some(function(r){ return r.volRangeHigh90 != null; })) {
      html += '<div class="sec">📈 「量能區間高檔(≥90百分位)」標記 vs 實際報酬</div>' +
        '<div style="color:#888;font-size:12px;margin-bottom:8px">不是跟20日均量比，是跟這檔股票自己過去約1年的成交量分布比——今天的量在它自己的歷史裡排第幾百分位。抓的是「平常冷門股票突然爆量」這種相對異常，跟固定倍數的「爆量」概念不同。</div>' +
        objArrayToTable(analyzeTagHitrate(backtestResults, 'volRangeHigh90', '量能區間高檔'));
    }
    if (backtestResults.some(function(r){ return r.volRangeLow10 != null; })) {
      html += '<div class="sec">📉 「量能區間低檔(≤10百分位)」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'volRangeLow10', '量能區間低檔'));
    }
    if (backtestResults.some(function(r){ return r.vpBuySupport != null; })) {
      html += '<div class="sec">📊 個股分價量表（Volume Profile）訊號 vs 實際報酬</div>' +
        '<div style="color:#888;font-size:12px;margin-bottom:8px">用近期股價區間的成交量分布找出成交量最大的價位區（POC）。以POC區為界判斷：守穩POC買進（由上回測POC、長下影線、收盤守住POC價）、突破POC追價買進（帶量長紅突破POC區）、反彈POC遇壓賣出（由下反彈到POC、長上影線、收盤壓回POC價下）、破位停損賣出（帶量長黑跌破POC區）。參數在程式碼VP_PARAMS。</div>' +
        objArrayToTable(analyzeTagHitrate(backtestResults, 'vpBuySupport', '分價量表-守穩POC買進')) +
        objArrayToTable(analyzeTagHitrate(backtestResults, 'vpBuyBreakout', '分價量表-突破POC追價買進')) +
        objArrayToTable(analyzeTagHitrate(backtestResults, 'vpSellResistance', '分價量表-反彈POC遇壓賣出')) +
        objArrayToTable(analyzeTagHitrate(backtestResults, 'vpSellBreakdown', '分價量表-破位停損賣出'));
    }
    if (backtestResults.some(function(r){ return r.benchmarkAbove20 != null; })) {
      html += '<div class="sec">🌐 「大盤站上20日均線」標記 vs 實際報酬（市場狀態濾網）</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'benchmarkAbove20', '大盤站上20日均線'));
    }
    if (backtestResults.some(function(r){ return r.benchmarkAbove60 != null; })) {
      html += '<div class="sec">🌐 「大盤站上60日均線」標記 vs 實際報酬（市場狀態濾網）</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'benchmarkAbove60', '大盤站上60日均線'));
    }

    if (backtestResults.some(function(r){ return r.patternBreakout != null; })) {
      html += '<div class="sec">🔍 「型態突破確認」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'patternBreakout', '型態突破確認'));
    }
    if (backtestResults.some(function(r){ return r.patternJustBroke != null; })) {
      html += '<div class="sec">🔥 「型態剛形成(剛突破)」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'patternJustBroke', '型態剛形成'));
    }
    if (backtestResults.some(function(r){ return r.pbAllPass != null; })) {
      html += '<div class="sec">✅ 「回後買上漲全通過」標記 vs 實際報酬</div>' + objArrayToTable(analyzeTagHitrate(backtestResults, 'pbAllPass', '回後買上漲全通過'));
    }
    if (backtestResults.some(function(r){ return r.patternHits != null; })) {
      html += '<div class="sec">📐 15種型態各自「剛形成」vs 實際報酬</div>' +
        '<div style="color:#888;font-size:12px;margin-bottom:8px">依樣本數由多到少排序，樣本數&lt;10筆的型態不列出（資料太少沒有參考意義）。這是每種型態單獨、不跟其他條件混在一起的乾淨表現。</div>' +
        objArrayToTable(analyzePatternHits(backtestResults));
    if (backtestResults.some(function(r){ return r.maBull != null; })) {
      html += '<div class="sec">🆕 新增技術面條件 vs 實際報酬（52週高點／均線多頭排列／布林收窄／向上跳空）</div>' +
        '<div style="color:#888;font-size:12px;margin-bottom:8px">52週高點需要約一年歷史，回測會自動多抓資料；資料不足的評估點不列入「是／否」。</div>' +
        objArrayToTable(analyzeNewTech(backtestResults));
    }
    }

    html += '<div class="sec">🎛️ 參數網格搜尋（單一評分公式的權重調整）</div>' +
      '<div style="margin-bottom:8px"><label style="font-size:12px;color:#aaa;margin-right:6px">優化目標天數：</label>' +
      '<select id="selGridHorizon" style="padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px">' +
        BACKTEST_HORIZONS.map(function(h){ return '<option value="'+h+'"'+(h===gridSearchHorizon?' selected':'')+'>'+h+'日</option>'; }).join('') +
      '</select></div>' +
      '<div style="color:#f0a500;font-size:12px;margin-bottom:8px">⚠️ 這是在已收集的歷史資料上找「表現較好」的參數組合，樣本數有限時容易過度適配——建議當作方向參考，人工確認合理後再手動調整正式評分公式，不要照單全收直接套用。</div>' +
      objArrayToTable(gridSearchParamsHtml(backtestResults, gridSearchHorizon));

    var comboResult = comboSearchHtml(backtestResults, comboSearchHorizon, 3, 50, comboMinWinRate, skipRedundantCombos, comboSortBy);
    var comboNBeforeT = comboResult.results.length;
    comboResult.results = comboResult.results.filter(function(r) { var v = parseFloat(r['t值(同日調整)']); return !isNaN(v) && v >= comboMinT; });
    if (comboBothHalves) comboResult.results = comboResult.results.filter(function(r) { return r['前後一致'] === '✅'; });
    var redundantCb = function(id) {
      return '<span><label style="font-size:12px;color:#aaa;cursor:pointer"><input type="checkbox" id="' + id + '"' + (skipRedundantCombos ? ' checked' : '') +
        ' style="width:auto;margin:0 4px 0 0;vertical-align:middle">排除冗餘組合</label></span>';
    };
    var cb = comboResult.base;
    html += '<div class="sec">🧩 多因子複選搜尋（找出哪幾項欄位組合起來勝率最高）</div>' +
      '<div style="margin-bottom:8px;display:flex;align-items:center;flex-wrap:wrap;gap:12px">' +
        '<span><label style="font-size:12px;color:#aaa;margin-right:6px">優化目標天數：</label>' +
        '<select id="selComboHorizon" style="padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px">' +
          BACKTEST_HORIZONS.map(function(h){ return '<option value="'+h+'"'+(h===comboSearchHorizon?' selected':'')+'>'+h+'日</option>'; }).join('') +
        '</select></span>' +
        '<span><label style="font-size:12px;color:#aaa;margin-right:6px">勝率門檻(%)：</label>' +
        '<input type="number" id="iComboMinWinRate" value="'+comboMinWinRate+'" min="0" max="100" step="5" style="width:60px;padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px"></span>' +
        '<span><label style="font-size:12px;color:#aaa;cursor:pointer" title="回測期間依評估日切成前後兩半，各自算同日調整t值；兩段都≥2代表不是單一段行情造成的巧合"><input type="checkbox" id="cbBothHalves"' + (comboBothHalves ? ' checked' : '') + ' style="width:auto;margin:0 4px 0 0;vertical-align:middle">只列前後一致(✅)</label></span>' +
        '<span><label style="font-size:12px;color:#aaa;margin-right:6px">t值門檻：</label>' +
        '<input type="number" id="iComboMinT" value="'+comboMinT+'" min="-10" max="20" step="0.5" title="排除 t值(同日調整) 低於門檻的組合（t<1：扣掉同一天大盤後幾乎沒有超額報酬）" style="width:60px;padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px"></span>' +
        '<span><label style="font-size:12px;color:#aaa;margin-right:6px">排序：</label>' +
        '<select id="selComboSort" style="padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px">' +
          '<option value="t"' + (comboSortBy === 't' ? ' selected' : '') + '>t值(同日調整)</option>' +
          '<option value="win"' + (comboSortBy === 'win' ? ' selected' : '') + '>勝率</option></select></span>' +
        redundantCb('cbRedundantCombo') +
        '<button type="button" id="btnExportCombo" style="padding:5px 12px;font-size:12px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:6px;color:#ddd;cursor:pointer">📥 匯出 Excel</button>' +
      '</div>' +
      '<div style="color:#f0a500;font-size:12px;margin-bottom:8px">窮舉1~3個條件旗標的AND組合。⚠️ 測試的組合越多，純粹運氣好而表現突出的組合也會越多（多重比較問題），且條件數越多、樣本數通常越少——加了15種型態各自的「剛形成」條件後，旗標總數變多很多，這裡把組合上限改回3個條件（4個條件的話組合數會爆到5萬多組，瀏覽器可能會跑很久甚至卡住）——共測試了 ' +
      comboResult.combosTested.toLocaleString() + ' 種組合，其中勝率≥'+comboMinWinRate+'%的 ' + comboNBeforeT.toLocaleString() + ' 組，再排除 t值<' + comboMinT + ' 後剩 ' + comboResult.results.length.toLocaleString() + ' 組列在下面' + (comboResult.base && comboResult.base.mid ? '（前後半段以 ' + comboResult.base.mid + ' 為界：✅＝兩段t都≥2，⚠️＝只有一段）' : '') + '，組合數越多，排在前面的結果就越需要保留懷疑，不代表真的有效，「同日超額報酬」＝每筆報酬扣掉同一天全部紀錄的平均，t值也用它算，已排除大盤齊漲齊跌；t>2較可信。' +
      (cb ? '<br><span style="color:#4a9eff">全體基準（' + comboSearchHorizon + '日）：樣本 ' + cb.n.toLocaleString() + '　平均報酬 ' + cb.avg.toFixed(2) + '%　勝率 ' + cb.win.toFixed(1) + '%　中位數 ' + cb.median.toFixed(2) + '%</span>' : '') + '</div>' +
      objArrayToTable(comboResult.results, 'comboSearchTable');


    var moonshotResult = findMoonshotCombos(backtestResults, moonshotHorizon, moonshotThreshold, 3, 20, moonshotMinPct, skipRedundantCombos);
    var mb = moonshotResult.base;
    html += '<div class="sec">🚀 飆股搜尋（找出最容易出現大行情的組合）</div>' +
      '<div style="margin-bottom:8px;display:flex;align-items:center;flex-wrap:wrap;gap:12px">' +
        '<span><label style="font-size:12px;color:#aaa;margin-right:6px">天數：</label>' +
        '<select id="selMoonshotHorizon" style="padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px">' +
          BACKTEST_HORIZONS.map(function(h){ return '<option value="'+h+'"'+(h===moonshotHorizon?' selected':'')+'>'+h+'日</option>'; }).join('') +
        '</select></span>' +
        '<span><label style="font-size:12px;color:#aaa;margin-right:6px">漲幅門檻(%)：</label>' +
        '<input type="number" id="iMoonshotThreshold" value="'+moonshotThreshold+'" min="5" max="200" step="5" style="width:60px;padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px"></span>' +
        '<span><label style="font-size:12px;color:#aaa;margin-right:6px">飆股比例門檻(%)：</label>' +
        '<input type="number" id="iMoonshotMinPct" value="'+moonshotMinPct+'" min="0" max="100" step="5" style="width:60px;padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px"></span>' +
        redundantCb('cbRedundantMoon') +
        '<button type="button" id="btnExportMoonshot" style="padding:5px 12px;font-size:12px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:6px;color:#ddd;cursor:pointer">📥 匯出 Excel</button>' +
      '</div>' +
      '<div style="color:#888;font-size:12px;margin-bottom:8px">' +
        '這裡看的不是「平均勝率」，是「這個組合出現後，有多高比例會在'+moonshotHorizon+'天內飆漲超過'+moonshotThreshold+'%」——一個組合平均報酬普通，只要常常噴出大行情，一樣會排在前面。' +
        '⚠️ 飆股本來就是稀有事件，樣本數少（門檻設20筆）時「飆股比例」很容易被少數幾次極端行情撐出虛高的數字，務必搭配「飆股次數」一起看，次數只有個位數的不建議當真。' +
      '</div>' +
      (mb ? '<div style="color:#4a9eff;font-size:12px;margin-bottom:6px">全體基準（' + moonshotHorizon + '日漲幅>' + moonshotThreshold + '%）：' + mb.n.toLocaleString() + ' 筆中 ' + mb.moonN.toLocaleString() + ' 次飆股，基準飆股比例 ' + mb.pct.toFixed(2) + '%。「倍數」＝組合飆股比例÷基準。' + (mb.mid ? '前後半段以 ' + mb.mid + ' 為界，兩段飆股比例都高才可信。' : '') + '</div>' : '') +
      '共測試了 ' + moonshotResult.combosTested.toLocaleString() + ' 種組合，其中飆股比例≥'+moonshotMinPct+'%的共 ' + moonshotResult.results.length.toLocaleString() + ' 組列在下面：' +
      objArrayToTable(moonshotResult.results, 'moonshotSearchTable');


    backtestOutput.innerHTML = html;

    var selGrid = document.getElementById('selGridHorizon');
    if (selGrid) selGrid.addEventListener('change', function() {
      gridSearchHorizon = parseInt(this.value, 10);
      renderBacktestResults();
    });
    var selCombo = document.getElementById('selComboHorizon');
    if (selCombo) selCombo.addEventListener('change', function() {
      comboSearchHorizon = parseInt(this.value, 10);
      renderBacktestResults();
    });
    var cbBothHalves = document.getElementById('cbBothHalves');
    if (cbBothHalves) cbBothHalves.addEventListener('change', function() { comboBothHalves = this.checked; renderBacktestResults(); });
    var iComboMinT = document.getElementById('iComboMinT');
    if (iComboMinT) iComboMinT.addEventListener('change', function() {
      var v = parseFloat(this.value); comboMinT = isNaN(v) ? 1 : v;
      renderBacktestResults();
    });
    var iComboMinWr = document.getElementById('iComboMinWinRate');
    if (iComboMinWr) iComboMinWr.addEventListener('change', function() {
      comboMinWinRate = parseFloat(this.value);
      if (isNaN(comboMinWinRate)) comboMinWinRate = 60;
      renderBacktestResults();
    });
    var btnExportCombo = document.getElementById('btnExportCombo');
    if (btnExportCombo) btnExportCombo.addEventListener('click', function() {
      exportTableToExcel('#comboSearchTable', '多因子複選搜尋_勝率' + comboMinWinRate + '%以上_t' + comboMinT + '以上', '多因子複選搜尋');
    });
    var selMoonshot = document.getElementById('selMoonshotHorizon');
    if (selMoonshot) selMoonshot.addEventListener('change', function() {
      moonshotHorizon = parseInt(this.value, 10);
      renderBacktestResults();
    });
    var iMoonshotThr = document.getElementById('iMoonshotThreshold');
    if (iMoonshotThr) iMoonshotThr.addEventListener('change', function() {
      moonshotThreshold = parseFloat(this.value) || 30;
      renderBacktestResults();
    });
    var iMoonshotMinPct = document.getElementById('iMoonshotMinPct');
    if (iMoonshotMinPct) iMoonshotMinPct.addEventListener('change', function() {
      moonshotMinPct = parseFloat(this.value);
      if (isNaN(moonshotMinPct)) moonshotMinPct = 10;
      renderBacktestResults();
    });
    var selComboSort = document.getElementById('selComboSort');
    if (selComboSort) selComboSort.addEventListener('change', function() { comboSortBy = this.value; renderBacktestResults(); });
    ['cbRedundantCombo', 'cbRedundantMoon'].forEach(function(id) {
      var el = document.getElementById(id);
      if (el) el.addEventListener('change', function() { skipRedundantCombos = this.checked; renderBacktestResults(); });
    });
    var btnExportMoonshot = document.getElementById('btnExportMoonshot');
    if (btnExportMoonshot) btnExportMoonshot.addEventListener('click', function() {
      exportTableToExcel('#moonshotSearchTable', '飆股搜尋_比例' + moonshotMinPct + '%以上', '飆股搜尋');
    });
  }

  function showLoading(msg) {
    output.innerHTML = '<div class="loading"><div class="spinner"></div><div class="loading-t">' + msg + '</div></div>';
  }

  function showError(msg) {
    output.innerHTML = '<div class="err">❌ ' + msg + '</div>';
  }

  // ── API ──
  // ── FinMind API v4（與 BBW Scanner 相同方式）──
  function getDateRange(days) {
    var end = new Date();
    var start = new Date();
    start.setDate(start.getDate() - days);
    var fmt = function(d) { return d.toISOString().slice(0, 10); };
    return { start: fmt(start), end: fmt(end) };
  }

  async function apiFetch(url) {
    // 嘗試多個 endpoint
    var endpoints = [
      url,
      url.replace('api.finmindtrade.com/api/v4', 'api.finmind.tw/api/latest'),
    ];
    // 主站、備援站的錯誤都列出來（以前只回報最後一個，主站的真正原因會被蓋掉）。
    // FinMind 每小時額度用完（HTTP 402）時不當成失敗：每 30 秒重試，最多等 65 分鐘，等待狀態透過 finmindQuotaNote 顯示。
    var waited = 0;
    while (true) {
      var errs = [], quotaMsg = null;
      for (var i = 0; i < endpoints.length; i++) {
        var host = endpoints[i].split('/')[2];
        try {
          var res = await fetch(endpoints[i], { mode: 'cors' });
          var json = null;
          try { json = await res.json(); } catch (je) {}
          var msg = json && json.msg ? String(json.msg) : '';
          if (res.status === 402 || (json && json.status === 402) || /upper limit/i.test(msg)) { quotaMsg = msg || 'HTTP 402'; break; }
          if (!res.ok) { errs.push(host + ' HTTP ' + res.status + (msg ? '：' + msg : '')); continue; }
          if (json && json.status === 200) return json;
          errs.push(host + '：' + (msg || 'status=' + (json ? json.status : '?')));
        } catch(e) { errs.push(host + ' 連線失敗（' + e.message + '）'); }
      }
      if (!quotaMsg) throw new Error(errs.join('；') || 'fetch failed');
      if (waited >= 65 * 60) throw new Error('FinMind 請求額度已用完（' + quotaMsg + '），等了 65 分鐘仍未恢復');
      if (typeof finmindQuotaNote === 'function') finmindQuotaNote('⏸ FinMind 請求額度已用完（' + quotaMsg + '），暫停等待額度恢復（已等 ' + Math.floor(waited / 60) + ' 分鐘，最多 65 分鐘），恢復後會自動繼續');
      await new Promise(function(r) { setTimeout(r, 30000); });
      waited += 30;
    }
  }
  var finmindQuotaNote = null;   // 回測／批次分析執行時設定，用來顯示「額度用完、等待中」

  async function fetchPriceData(stockId, token, days) {
    var range = getDateRange(days);
    var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=' + stockId + '&start_date=' + range.start + '&end_date=' + range.end + '&token=' + token;
    var json = await apiFetch(url);
    if (!json.data || !json.data.length) throw new Error('無資料（可能代號錯誤）');
    // FinMind 在暫停交易日可能回傳收盤價 0，會讓報酬被算成 -100%，這種日子直接略過
    return json.data.filter(function(d) { return +d.close > 0 && +d.max > 0 && +d.min > 0; });
  }

  // ── 相對強弱（vs 大盤）：用0050（元大台灣50）當大盤代理，不用另外接指數API，
  // 直接當一般股票代號抓，同一套 fetchPriceData 就能用。只抓一次，所有股票共用
  // 同一份大盤資料，不是每檔股票都各抓一次。──
  var BENCHMARK_ID = '0050';
  async function fetchBenchmarkSeries(token, days) {
    var rows = await fetchPriceData(BENCHMARK_ID, token, days);
    var arr = rows.map(function(d) {
      return { date: d.date, close: +d.close };
    }).sort(function(a,b){ return a.date<b.date?-1:1; });
    // 順便把大盤自己的20日/60日均線也算好，市場狀態濾網（大盤站上/跌破均線）
    // 要用，這裡一次算完整條序列，比逐個評估點現算快很多。
    for (var i = 0; i < arr.length; i++) {
      if (i >= 19) {
        var sum20 = 0;
        for (var j = i - 19; j <= i; j++) sum20 += arr[j].close;
        arr[i].ma20 = sum20 / 20;
      }
      if (i >= 59) {
        var sum60 = 0;
        for (var k = i - 59; k <= i; k++) sum60 += arr[k].close;
        arr[i].ma60 = sum60 / 60;
      }
    }
    var dateIdx = {};
    arr.forEach(function(bar, i) { dateIdx[bar.date] = i; });
    return { arr: arr, dateIdx: dateIdx };
  }

  // 找「小於等於dateStr」的最新一筆大盤資料索引（股票的交易日曆理論上跟大盤一致，
  // 但用「往回找」保險起見處理極少數對不齊的情況，例如新股上市初期）
  function findBenchmarkIdx(benchmark, dateStr) {
    if (benchmark.dateIdx[dateStr] != null) return benchmark.dateIdx[dateStr];
    for (var i = benchmark.arr.length - 1; i >= 0; i--) {
      if (benchmark.arr[i].date <= dateStr) return i;
    }
    return -1;
  }

  // 市場狀態濾網：評估當下大盤自己站在20日/60日均線的上面還下面。跟相對強弱
  // 不同——相對強弱看的是「這檔股票 vs 大盤」，這個看的是「大盤自己的多空位置」，
  // 用同一份benchmark資料算，不用多打API。
  function computeBenchmarkRegime(benchmark, dateStr) {
    if (!benchmark) return { above20: null, above60: null };
    var idx = findBenchmarkIdx(benchmark, dateStr);
    if (idx < 0) return { above20: null, above60: null };
    var bar = benchmark.arr[idx];
    return {
      above20: bar.ma20 != null ? (bar.close > bar.ma20) : null,
      above60: bar.ma60 != null ? (bar.close > bar.ma60) : null,
    };
  }

  // 算「股票N日報酬 - 大盤N日報酬」的相對強弱，正值代表跑贏大盤。用20日窗口，
  // 跟其他「近20日」類指標（流動性門檻、vm20量比）的天數一致，方便橫向比較。
  function computeRelativeStrength(data, benchmark, lookback) {
    lookback = lookback || 20;
    if (!benchmark || !benchmark.arr.length || data.length <= lookback) return null;
    var last = data[data.length - 1];
    var prevBar = data[data.length - 1 - lookback];
    if (!prevBar || !prevBar.close) return null;
    var stockRet = (last.close - prevBar.close) / prevBar.close * 100;

    var benchIdx = findBenchmarkIdx(benchmark, last.date);
    if (benchIdx < lookback) return null;
    var benchNow = benchmark.arr[benchIdx];
    var benchPrev = benchmark.arr[benchIdx - lookback];
    if (!benchPrev || !benchPrev.close) return null;
    var benchRet = (benchNow.close - benchPrev.close) / benchPrev.close * 100;

    return stockRet - benchRet;
  }

  // 成交量確認：當天成交量 ÷ 近20日均量（vm20已經在enrich()裡算好），
  // ≥1.5倍算「溫和爆量」，≥2倍算「明顯爆量」。
  function computeVolRatio(data) {
    var last = data[data.length - 1];
    if (!last.vm20) return null;
    return last.volume / last.vm20;
  }

  // 量能區間位置（概念上跟「近3年P/E區間」一樣：不是跟固定倍數比，是跟這檔股票
  // 自己過去的歷史分布比）——今天的成交量，在這檔股票自己過去N天（最多約1年，
  // 用這次批次分析設定的「分析天數」抓多少算多少，天數設越長，這個百分位越可靠）
  // 的成交量分布裡排第幾百分位。跟volRatio（只跟20日均量比）不一樣的地方：一檔
  // 平常量能就偏大的股票，20日均量比可能怎麼算都不極端；但如果它是「平常冷門，
  // 這次突然爆量」，這裡的百分位會直接跳到接近100，能抓到真正相對於這檔股票
  // 自己常態的異常，不是絕對值的大小。
  function computeVolumeRangePosition(data) {
    if (data.length < 20) return null; // 資料太少，百分位沒有參考意義
    var window = data.slice(Math.max(0, data.length - 252)); // 最多抓約1年交易日
    var volumes = window.map(function(d){ return d.volume; }).filter(function(v){ return v != null; });
    if (volumes.length < 20) return null;
    var todayVol = data[data.length - 1].volume;
    var below = volumes.filter(function(v){ return v < todayVol; }).length;
    return below / volumes.length * 100; // 0~100，越高代表今天的量在這檔股票的歷史裡算越大
  }

  // 量能斜率：近5日均量 vs 近10日均量，判斷量能本身是在放大還是縮小的趨勢中
  // （不是單看某一天的量，是看「量的變化方向」）。正值代表近期量能在擴張。
  function computeVolumeTrend(data) {
    if (data.length < 10) return null;
    var last10 = data.slice(-10);
    var last5 = data.slice(-5);
    var avg10 = last10.reduce(function(s,d){ return s + d.volume; }, 0) / 10;
    var avg5 = last5.reduce(function(s,d){ return s + d.volume; }, 0) / 5;
    if (!avg10) return null;
    return (avg5 - avg10) / avg10 * 100;
  }

  // ── 近3年 P/E 區間（TaiwanStockPER，每股票每次批次分析都會多打一次API）──
  // ETF（如0050）本身沒有EPS，FinMind不會回傳PER資料，這種情況直接顯示「無資料」，
  // 不是bug。取得的PER陣列先過濾掉0、負值、非數字（虧損股常見PER<=0或缺值），
  // 再取這3年內的最大最小值作為區間，最新一筆視為目前P/E。
  function getDateRangeYearsAgo(years) {
    var end = new Date();
    var start = new Date();
    start.setFullYear(start.getFullYear() - years);
    var fmt = function(d) { return d.toISOString().slice(0, 10); };
    return { start: fmt(start), end: fmt(end) };
  }

  async function fetchPERange(stockId, token, years) {
    var range = getDateRangeYearsAgo(years || 3);
    var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPER&data_id=' + stockId + '&start_date=' + range.start + '&end_date=' + range.end + '&token=' + token;
    var json = await apiFetch(url);
    var rows = json.data || [];
    if (!rows.length) return null;
    rows.sort(function(a, b) { return a.date < b.date ? -1 : 1; });
    var values = rows.map(function(r) { return parseFloat(r.PER); }).filter(function(v) { return isFinite(v) && v > 0; });
    if (!values.length) return null;
    var min = Math.min.apply(null, values);
    var max = Math.max.apply(null, values);
    var current = parseFloat(rows[rows.length - 1].PER);
    if (!isFinite(current) || current <= 0) current = values[values.length - 1];
    return { min: min, max: max, current: current };
  }

  // ── 近3個月營收 YoY／MoM（TaiwanStockMonthRevenue，每股票每次批次分析都會多打一次API）──
  // YoY／MoM 是自己算的（FinMind只給原始revenue，不含增減率），需要抓夠長的區間
  // （往前16個月）才能算出最近3個月當中「最舊那個月」的YoY（要對比去年同月）。
  function fetchRevenueYoYMoM(stockId, token) {
    var end = new Date();
    var start = new Date();
    start.setMonth(start.getMonth() - 16);
    var fmt = function(d) { return d.toISOString().slice(0, 10); };
    var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockMonthRevenue&data_id=' + stockId + '&start_date=' + fmt(start) + '&end_date=' + fmt(end) + '&token=' + token;
    return apiFetch(url).then(function(json) {
      var rows = json.data || [];
      if (!rows.length) return null;
      var map = {};
      rows.forEach(function(r) {
        var key = r.revenue_year + '-' + (r.revenue_month < 10 ? '0' : '') + r.revenue_month;
        map[key] = r.revenue;
      });
      var months = Object.keys(map).sort();
      var last3 = months.slice(-3);
      if (!last3.length) return null;
      return last3.map(function(key) {
        var parts = key.split('-');
        var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10);
        var rev = map[key];
        var prevKey = (m === 1) ? (y - 1) + '-12' : y + '-' + ((m - 1) < 10 ? '0' : '') + (m - 1);
        var yoyKey = (y - 1) + '-' + parts[1];
        var mom = map[prevKey] ? (rev - map[prevKey]) / map[prevKey] * 100 : null;
        var yoy = map[yoyKey] ? (rev - map[yoyKey]) / map[yoyKey] * 100 : null;
        return { year: y, month: m, revenue: rev, mom: mom, yoy: yoy };
      });
    });
  }

  // ── 三大法人買賣超（近3個月，逐月加總）──
  // TaiwanStockInstitutionalInvestorsBuySell是「每日」資料，不是月資料，這裡自己依日期
  // 分組加總成月度買賣超。抓約100天回來，涵蓋最近3個完整月份綽綽有餘。
  // 買賣超 = buy - sell（單日、單一法人類別），三大法人合計 = 外資+投信+自營商
  // 三類加總。正值=買超（淨買進），負值=賣超（淨賣出）。單位是FinMind原始股數，
  // 沒有換算成「張」。
  function fetchInstitutionalMonthly(stockId, token) {
    var end = new Date();
    var start = new Date();
    start.setDate(start.getDate() - 100);
    var fmt = function(d) { return d.toISOString().slice(0, 10); };
    var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id=' + stockId + '&start_date=' + fmt(start) + '&end_date=' + fmt(end) + '&token=' + token;
    return apiFetch(url).then(function(json) {
      var rows = json.data || [];
      if (!rows.length) throw new Error('回傳空陣列（可能此帳號方案未開通此資料集，或該股票近100天無法人交易紀錄）');
      var monthMap = {}; // "YYYY-MM" -> 累加淨買賣超
      rows.forEach(function(r) {
        var d = String(r.date || '').slice(0, 10);
        if (d.length < 7) return;
        var key = d.slice(0, 7); // "YYYY-MM"
        var net = (parseFloat(r.buy) || 0) - (parseFloat(r.sell) || 0);
        monthMap[key] = (monthMap[key] || 0) + net;
      });
      var months = Object.keys(monthMap).sort();
      var last3 = months.slice(-3);
      if (!last3.length) {
        var sampleKeys = Object.keys(rows[0]).join(',');
        throw new Error('取得 ' + rows.length + ' 筆資料，但解析不出日期/買賣超欄位（欄位範例：' + sampleKeys + '）');
      }
      var outArr = last3.map(function(key) {
        var parts = key.split('-');
        return { year: parseInt(parts[0], 10), month: parseInt(parts[1], 10), net: monthMap[key] };
      });
      // 投信／外資近5日、投信連3日買超（2026-10-03 新增，掛在陣列屬性 flow 上，不影響原本的月份資料；跟回測同一套算法）
      var lDaily = {}, lTrust = {}, lForeign = {};
      rows.forEach(function(r) {
        var d = String(r.date || '').slice(0, 10); if (d.length < 10) return;
        var net = (parseFloat(r.buy) || 0) - (parseFloat(r.sell) || 0);
        lDaily[d] = (lDaily[d] || 0) + net;
        if (isTrustRow(r)) lTrust[d] = (lTrust[d] || 0) + net;
        else if (isForeignRow(r)) lForeign[d] = (lForeign[d] || 0) + net;
      });
      var lDates = Object.keys(lDaily).sort().map(function(d) { return { date: d }; });
      outArr.flow = lDates.length ? computeInstFlowAsof(lDaily, lTrust, lForeign, lDates, lDates.length - 1) : null;
      return outArr;
    });
  }
  function isTrustRow(r) { var nm = String(r.name || ''); return nm.indexOf('Investment_Trust') >= 0 || nm.indexOf('投信') >= 0; }
  // 外資＝外資及陸資（Foreign_Investor）＋外資自營商（Foreign_Dealer_Self）
  function isForeignRow(r) { var nm = String(r.name || ''); return nm.indexOf('Foreign') >= 0 || nm.indexOf('外資') >= 0; }
  // 截至 idx（含）：投信／外資近5個交易日買賣超合計、投信是否連續3日買超；這5天都沒有法人資料時全部回傳 null
  function computeInstFlowAsof(dailyNet, trustNet, foreignNet, data, idx) {
    var any = false, t5 = 0, f5 = 0;
    for (var k = Math.max(0, idx - 4); k <= idx; k++) {
      var d = data[k].date;
      if (dailyNet[d] != null) any = true;
      t5 += trustNet[d] || 0; f5 += foreignNet[d] || 0;
    }
    if (!any) return { trust5: null, foreign5: null, trustStreak3: null };
    var st = idx >= 2 ? 1 : 0;
    for (var q = Math.max(0, idx - 2); q <= idx; q++) { if (!((trustNet[data[q].date] || 0) > 0)) st = 0; }
    return { trust5: t5, foreign5: f5, trustStreak3: st };
  }

  // ── 回測專用：長時間歷史版本的月營收／法人買賣超抓取，以及「當時已公告」的
  // 判斷邏輯（避免回測偷看未來才公告的資料）。──
  function fetchRevenueHistoryForBacktest(stockId, token, daysBack) {
    var end = new Date();
    var start = new Date();
    start.setDate(start.getDate() - daysBack);
    var fmt = function(d) { return d.toISOString().slice(0, 10); };
    var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockMonthRevenue&data_id=' + stockId + '&start_date=' + fmt(start) + '&end_date=' + fmt(end) + '&token=' + token;
    return apiFetch(url).then(function(json) {
      var rows = json.data || [];
      var revMap = {}; // "year-month" -> revenue
      rows.forEach(function(r) {
        if (r.revenue_year == null || r.revenue_month == null || r.revenue == null) return;
        revMap[r.revenue_year + '-' + r.revenue_month] = r.revenue;
      });
      return revMap;
    });
  }

  function fetchInstitutionalHistoryForBacktest(stockId, token, daysBack) {
    var end = new Date();
    var start = new Date();
    start.setDate(start.getDate() - daysBack);
    var fmt = function(d) { return d.toISOString().slice(0, 10); };
    var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id=' + stockId + '&start_date=' + fmt(start) + '&end_date=' + fmt(end) + '&token=' + token;
    return apiFetch(url).then(function(json) {
      var rows = json.data || [];
      var dailyNet = {}, trustNet = {}, foreignNet = {}; // "YYYY-MM-DD" -> 當日淨買賣超（全部法人／投信／外資）
      rows.forEach(function(r) {
        var d = String(r.date || '').slice(0, 10);
        if (d.length < 10) return;
        var net = (parseFloat(r.buy) || 0) - (parseFloat(r.sell) || 0);
        dailyNet[d] = (dailyNet[d] || 0) + net;
        if (isTrustRow(r)) trustNet[d] = (trustNet[d] || 0) + net;
        else if (isForeignRow(r)) foreignNet[d] = (foreignNet[d] || 0) + net;
      });
      return { all: dailyNet, trust: trustNet, foreign: foreignNet };
    });
  }

  // 判斷 year年month月 的月營收，在 asOfDateStr 這天是不是已經依法公告
  // （台灣規定月營收要在次月10日前公告，這裡用次月10日當保守的「已知」門檻，
  // 避免回測用到「當下根本還沒公告」的營收資料造成偷看未來）。
  function revenueKnownBy(year, month, asOfDateStr) {
    var discloseYear = month < 12 ? year : year + 1;
    var discloseMonth = month < 12 ? month + 1 : 1;
    var discloseDate = new Date(discloseYear, discloseMonth - 1, 10);
    var asOf = new Date(asOfDateStr + 'T00:00:00');
    return asOf >= discloseDate;
  }

  // 回傳 asOfDateStr 當下「已公告」的最近n個月份 [{year,month}, ...]（舊到新）
  function lastNKnownMonths(asOfDateStr, n) {
    var asOf = new Date(asOfDateStr + 'T00:00:00');
    var y = asOf.getFullYear(), m = asOf.getMonth() + 1;
    var result = [];
    for (var i = 0; i < 18; i++) {  // 最多往前找18個月，避免極端情況下無限迴圈
      if (revenueKnownBy(y, m, asOfDateStr)) {
        result.push({ year: y, month: m });
        if (result.length >= n) break;
      }
      m -= 1;
      if (m < 1) { m = 12; y -= 1; }
    }
    result.reverse();
    return result;
  }

  // 月營收動能（只用 asOfDateStr 當下已公告的月份，跟 PY 版 rev_feats 同一套）：
  // revYoyTurn＝最新月YoY由負轉正；revYoy3Pos＝近3個月YoY皆為正；revHigh24＝最新月營收創24個月新高
  function computeRevFeats(revMap, asOfDateStr) {
    var out = { revYoyTurn: null, revYoy3Pos: null, revHigh24: null };
    var months = lastNKnownMonths(asOfDateStr, 4);
    if (!revMap || !Object.keys(revMap).length || !months.length) return out;
    var rv = function(y, m) { var v = revMap[y + '-' + m]; return v == null ? null : parseFloat(v); };
    var yoy = function(y, m) { var rt = rv(y, m), rl = rv(y - 1, m); return (rt != null && rl) ? (rt - rl) / rl * 100 : null; };
    var ys = months.map(function(x) { return yoy(x.year, x.month); });
    var n = ys.length;
    if (n >= 2 && ys[n - 1] != null && ys[n - 2] != null) out.revYoyTurn = (ys[n - 1] > 0 && ys[n - 2] <= 0) ? 1 : 0;
    if (n >= 3 && ys.slice(-3).every(function(v) { return v != null; })) out.revYoy3Pos = ys.slice(-3).every(function(v) { return v > 0; }) ? 1 : 0;
    var last = months[n - 1], cur = rv(last.year, last.month);
    if (cur != null) {
      var prev = [], y = last.year, m = last.month;
      for (var k = 1; k < 24; k++) { m -= 1; if (m < 1) { m = 12; y -= 1; } var p = rv(y, m); if (p != null) prev.push(p); }
      if (prev.length >= 18) out.revHigh24 = cur > Math.max.apply(null, prev) ? 1 : 0;
    }
    return out;
  }

  // 把已經抓好的完整日K資料，依年月分組算平均收盤價，回傳 {"year-month": 均價}。
  // 這樣算月均價YoY不用再多打一次API——直接沿用DMI計算已經抓到的長天期價格資料。
  function buildPriceMonthAvgMap(data) {
    var sums = {}, counts = {};
    data.forEach(function(d) {
      var y = parseInt(d.date.slice(0, 4), 10), m = parseInt(d.date.slice(5, 7), 10);
      var key = y + '-' + m;
      sums[key] = (sums[key] || 0) + d.close;
      counts[key] = (counts[key] || 0) + 1;
    });
    var out = {};
    Object.keys(sums).forEach(function(k) { out[k] = sums[k] / counts[k]; });
    return out;
  }

  // 算某個評估日「當下」能看到的近3月營收YoY vs 均價YoY 乖離度加總（近3月合計）。
  // 只用 as_of 當天已公告的月份，年增率的去年同月比較資料不存在就跳過那個月。
  // 回傳 {divTotal, priceYoyTotal}：divTotal是原本的「營收YoY-股價YoY」乖離度加總，
  // priceYoyTotal是單獨的股價YoY加總（不跟營收相減），讓兩者可以分開當條件旗標用
  // ——乖離度看的是「營收跟股價有沒有脫節」，股價YoY單獨看的是「股價本身年增率
  // 強不強」，是兩件不同的事，不該只看合併後的差值。
  function computeDivergenceAsof(revMap, priceMonthAvgMap, evalDateStr) {
    var months = lastNKnownMonths(evalDateStr, 3);
    if (!months.length) return { divTotal: null, priceYoyTotal: null };
    var divTotal = 0, priceYoyTotal = 0, anyValid = false, anyPriceValid = false;
    months.forEach(function(mo) {
      var revThis = revMap[mo.year + '-' + mo.month];
      var revLast = revMap[(mo.year - 1) + '-' + mo.month];
      var pxThis = priceMonthAvgMap[mo.year + '-' + mo.month];
      var pxLast = priceMonthAvgMap[(mo.year - 1) + '-' + mo.month];
      if (pxLast && pxThis != null) {
        priceYoyTotal += (pxThis - pxLast) / pxLast * 100;
        anyPriceValid = true;
      }
      if (!revLast || revThis == null || !pxLast || pxThis == null) return;
      var revYoy = (revThis - revLast) / revLast * 100;
      var pxYoy = (pxThis - pxLast) / pxLast * 100;
      divTotal += (revYoy - pxYoy);
      anyValid = true;
    });
    return { divTotal: anyValid ? divTotal : null, priceYoyTotal: anyPriceValid ? priceYoyTotal : null };
  }

  // 算eval_date_str往前3個月（約90天）的三大法人買賣超加總。法人日資料沒有
  // 公告延遲的疑慮，直接抓截至當天的歷史資料即可。
  function computeInst3mAsof(instDailyMap, evalDateStr) {
    var end = new Date(evalDateStr + 'T00:00:00');
    var start = new Date(end.getTime() - 90 * 86400000);
    var total = 0, anyValid = false;
    Object.keys(instDailyMap).forEach(function(dStr) {
      var d = new Date(dStr + 'T00:00:00');
      if (d >= start && d <= end) { total += instDailyMap[dStr]; anyValid = true; }
    });
    return anyValid ? total : null;
  }


  // 的去年同月，通常要抓超過一年），每股票每次批次分析又會多打一次API。
  function lastNCalendarMonths(n) {
    var out = [];
    var d = new Date();
    d.setDate(1); // 避免月底跨月造成 setMonth 邊界問題
    for (var i = 0; i < n; i++) {
      out.unshift({ year: d.getFullYear(), month: d.getMonth() + 1 });
      d.setMonth(d.getMonth() - 1);
    }
    return out;
  }

  function fetchMonthlyAvgPriceYoY(stockId, token, monthsList) {
    if (!monthsList || !monthsList.length) monthsList = lastNCalendarMonths(3);
    var oldest = monthsList[0];
    var start = new Date(oldest.year - 1, oldest.month - 1, 1);
    var end = new Date();
    var fmt = function(d) { return d.toISOString().slice(0, 10); };
    var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=' + stockId + '&start_date=' + fmt(start) + '&end_date=' + fmt(end) + '&token=' + token;
    return apiFetch(url).then(function(json) {
      var rows = json.data || [];
      if (!rows.length) return null;
      var sums = {}, counts = {};
      rows.forEach(function(r) {
        var key = String(r.date).slice(0, 7); // "YYYY-MM"
        var c = parseFloat(r.close);
        if (!isFinite(c) || c <= 0) return;
        sums[key] = (sums[key] || 0) + c;
        counts[key] = (counts[key] || 0) + 1;
      });
      var avgMap = {};
      Object.keys(sums).forEach(function(k) { avgMap[k] = sums[k] / counts[k]; });
      return monthsList.map(function(m) {
        var key = m.year + '-' + (m.month < 10 ? '0' : '') + m.month;
        var yoyKey = (m.year - 1) + '-' + (m.month < 10 ? '0' : '') + m.month;
        var thisAvg = avgMap[key], lastAvg = avgMap[yoyKey];
        var yoy = (thisAvg != null && lastAvg) ? (thisAvg - lastAvg) / lastAvg * 100 : null;
        return { year: m.year, month: m.month, avg: thisAvg, avgLastYear: lastAvg, yoy: yoy };
      });
    });
  }

  async function fetchName(token, stockId) {
    try {
      var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo&token=' + token;
      var json = await apiFetch(url);
      if (json.status === 200) {
        var row = (json.data || []).find(function(d) { return d.stock_id === stockId; });
        if (row) return row.stock_name || stockId;
      }
    } catch(e) {}
    return stockId;
  }

  // 一次抓全部股票代號→名稱的對照表，給全市場規模的迴圈用（例如回測1,100多檔），
  // 避免像 fetchName() 那樣每檔股票都重新下載一次完整清單。
  async function fetchStockNameMap(token) {
    var map = {};
    try {
      var url = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo&token=' + token;
      var json = await apiFetch(url);
      if (json.status === 200) {
        (json.data || []).forEach(function(d) { map[d.stock_id] = d.stock_name || d.stock_id; });
      }
    } catch (e) {}
    return map;
  }

  // ── 盤中即時股價快照（taiwan_stock_tick_snapshot，約10秒更新一次）──
  // 注意：這是 FinMind「只限sponsor會員使用」的端點，跟一般 dataset= 資料不同路徑。
  // 非sponsor會員呼叫會被拒絕（401/403），這裡設計成失敗就整批跳過、不中斷主流程，
  // 並在畫面上用一則訊息說明原因，而不是每檔股票都各自報一次錯誤洗版。
  var REALTIME_CHUNK_SIZE = 80;

  function chunkArray(arr, size) {
    var out = [];
    for (var i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }

  async function fetchRealtimeSnapshots(token, stockIds) {
    var chunks = chunkArray(stockIds, REALTIME_CHUNK_SIZE);
    var map = {};
    var lastErr = null;
    for (var i = 0; i < chunks.length; i++) {
      var params = chunks[i].map(function(id) { return 'data_id=' + encodeURIComponent(id); }).join('&');
      var url = 'https://api.finmindtrade.com/api/v4/taiwan_stock_tick_snapshot?' + params + '&token=' + encodeURIComponent(token);
      try {
        var resp = await fetch(url, { mode: 'cors' });
        if (!resp.ok) {
          var bodyTxt = '';
          try { bodyTxt = await resp.text(); } catch (e2) {}
          throw new Error('HTTP ' + resp.status + (bodyTxt ? '：' + bodyTxt.slice(0, 200) : ''));
        }
        var json = await resp.json();
        if (json.status !== 200) throw new Error(json.msg || ('status=' + json.status));
        (json.data || []).forEach(function(row) { if (row && row.stock_id) map[row.stock_id] = row; });
      } catch (e) {
        lastErr = e.message;
      }
    }
    return { map: map, error: Object.keys(map).length ? null : lastErr };
  }

  // 把即時快照併入歷史資料的最後一筆：只有在「今天還沒被包含在EOD資料裡」時才附加一筆
  // 合成的當日K棒（開高低收用快照的open/high/low/close，成交量用累積成交量total_volume），
  // 讓評分／型態辨識／圖表在盤中也能反映當下股價，而不是停在昨天收盤。
  // ── 盤中即時量換算成全天估計量（2026-10-19）：盤中的累積成交量只是「到目前為止」的量，直接拿來跟均量比，
  //    幾乎每檔都會被誤判成「地量」。依開盤後經過的時間推估全天量；成交量前段較多，累積比例 ≈ 時間比例^0.8；
  //    開盤15分鐘內以15分鐘計（避免放大太多），收盤後（或開盤前）不換算。
  function intradayVolumeFactor(tsMs, tz, openMin, closeMin) {
    try {
      var parts = new Intl.DateTimeFormat('en-US', { timeZone: tz, hour: '2-digit', minute: '2-digit', hour12: false }).formatToParts(new Date(tsMs));
      var h = 0, m = 0;
      parts.forEach(function(p) { if (p.type === 'hour') h = (+p.value) % 24; if (p.type === 'minute') m = +p.value; });
      var t = h * 60 + m;
      if (t >= closeMin || t <= openMin) return 1;
      var frac = Math.max((t - openMin) / (closeMin - openMin), 15 / (closeMin - openMin));
      return 1 / Math.pow(frac, 0.8);
    } catch (e) { return 1; }
  }
  function mergeRealtimeSnapshot(rawData, snapshot) {
    if (!snapshot || !snapshot.date) return { data: rawData, injected: false };
    var todayStr = String(snapshot.date).slice(0, 10);
    var lastDate = rawData.length ? rawData[rawData.length - 1].date : null;
    if (lastDate && todayStr <= lastDate) return { data: rawData, injected: false };
    var open = +snapshot.open, high = +snapshot.high, low = +snapshot.low, close = +snapshot.close;
    var vol = snapshot.total_volume != null ? +snapshot.total_volume : (snapshot.volume != null ? +snapshot.volume : null);
    if (!close || !isFinite(close)) return { data: rawData, injected: false };
    // 盤中累積量 → 全天估計量（台股 09:00～13:30）；快照有時間就用快照時間，否則用現在時間
    var snapTs = /\d{2}:\d{2}/.test(String(snapshot.date)) ? Date.parse(String(snapshot.date).replace(' ', 'T').slice(0, 19) + '+08:00') : Date.now();
    if (vol != null && isFinite(vol)) vol = Math.round(vol * intradayVolumeFactor(isFinite(snapTs) ? snapTs : Date.now(), 'Asia/Taipei', 540, 810));
    var merged = rawData.concat([{
      date: todayStr,
      open: open || close, high: high || close, low: low || close, close: close,
      volume: (vol != null && isFinite(vol)) ? vol : 0
    }]);
    return { data: merged, injected: true };
  }

  // ── 批次分析主函數 ──
  var batchResults = [];   // { stockId, name, data, tr, kl, ma, vl, pb, total }
  var activeTab = 0;
  var batchExtras = { pe: false, rev: false, pxyoy: false, inst: false }; // 本次批次分析實際使用的可選欄位（renderBatch會依此決定要不要畫出欄位）

  async function batchAnalysis() {
    var token = iToken.value.trim();
    var days  = parseInt(iDays.value, 10);
    var raw   = document.getElementById('iStocks').value;

    if (!token) { showError('請輸入 FinMind API Token'); return; }

    // 先測試 API 連線
    output.innerHTML = '<div class="batch-progress"><div style="font-size:14px;color:#ccc">🔌 測試 API 連線中...</div></div>';
    try {
      var testUrl = 'https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo&token=' + token;
      var testJson = await apiFetch(testUrl);
      if (testJson.status !== 200) {
        showError('Token 無效或 API 錯誤：' + (testJson.msg || 'status=' + testJson.status) + '\n請確認 FinMind Token 是否正確。');
        btnRun.disabled = false; btnRun.textContent = '🔍 批次分析'; return;
      }
    } catch(connErr) {
      showError('❌ 無法連線至 FinMind API\n\n錯誤：' + connErr.message + '\n\n請確認：\n1. Token 是否正確\n2. 網路連線是否正常\n3. 是否從本機伺服器開啟（非 file:// 協定）');
      btnRun.disabled = false; btnRun.textContent = '🔍 批次分析'; return;
    }

    // 解析股票清單（換行或逗號分隔）
    var stocks = raw.split(/[\n,，、\s]+/).map(function(s){ return s.trim(); }).filter(Boolean);
    if (!stocks.length) { showError('請輸入至少一個股票代號'); return; }

    // 相對強弱要用的大盤(0050)資料，整批只抓一次，不是每檔股票各抓一次
    currentBenchmark = null;
    try {
      currentBenchmark = await fetchBenchmarkSeries(token, days);
    } catch (e) {
      currentBenchmark = null; // 抓不到大盤資料，相對強弱那組欄位/旗標就不會出現，不影響其他分析
    }

    // 盤中即時股價快照：整批股票一次（或分批）呼叫，避免每檔各打一次API
    var realtimeMap = {};
    var realtimeErr = null;
    var useRealtime = document.getElementById('iRealtime').checked;
    batchExtras = {
      pe: document.getElementById('iShowPE').checked,
      rev: document.getElementById('iShowRevenue').checked,
      pxyoy: document.getElementById('iShowPriceYoY').checked,
      inst: document.getElementById('iShowInst').checked
    };
    if (useRealtime) {
      output.innerHTML = '<div class="batch-progress"><div style="font-size:14px;color:#ccc">🔴 抓取盤中即時股價快照中...</div></div>';
      var rtResult = await fetchRealtimeSnapshots(token, stocks);
      realtimeMap = rtResult.map;
      realtimeErr = rtResult.error;
    }

    btnRun.disabled = true;
    btnRun.textContent = '⏳ 分析中...';
    batchResults = [];
    document.getElementById('batchStatus').textContent = '';

    // 顯示進度
    output.innerHTML = '<div class="batch-progress">' +
      '<div style="font-size:14px;color:#ccc">批次分析中：<strong id="progText">0 / ' + stocks.length + '</strong></div>' +
      '<div class="progress-bar-wrap"><div class="progress-bar-fill" id="progBar" style="width:0%"></div></div>' +
      '</div>' +
      '<div id="batchLog" style="font-size:12px;color:#888;margin-top:8px"></div>';

    var progText = document.getElementById('progText');
    var progBar  = document.getElementById('progBar');
    var batchLog = document.getElementById('batchLog');
    finmindQuotaNote = function(msg) { batchLog.innerHTML += '<div style="color:#f0a500">' + msg + '</div>'; };

    if (useRealtime) {
      if (realtimeErr) {
        batchLog.innerHTML += '<div style="color:#f0a500">⚠️ 即時快照抓取失敗（已略過，僅用歷史資料分析）：' + realtimeErr + '　※此功能限FinMind sponsor會員使用</div>';
      } else {
        batchLog.innerHTML += '<div style="color:#4a9eff">🔴 已取得 ' + Object.keys(realtimeMap).length + ' 檔即時快照</div>';
      }
    }

    var realtimeInjectedCount = 0;
    for (var i = 0; i < stocks.length; i++) {
      var sid = stocks[i];
      batchLog.innerHTML += '<div>📡 分析 ' + sid + '...</div>';
      try {
        var rows = await fetchPriceData(sid, token, Math.max(days, LONG_HISTORY_DAYS));   // 多抓的部分只用來算52週高點等欄位
        if (!rows || !rows.length) throw new Error('無資料');
        var rawData = rows.map(function(d) {
          return { date:d.date, open:+d.open, high:+d.max, low:+d.min, close:+d.close, volume:+d.Trading_Volume };
        }).sort(function(a,b){ return a.date<b.date?-1:1; });
        var rtInjected = false;
        if (useRealtime && realtimeMap[sid]) {
          var mergeResult = mergeRealtimeSnapshot(rawData, realtimeMap[sid]);
          rawData = mergeResult.data;
          rtInjected = mergeResult.injected;
          if (rtInjected) realtimeInjectedCount++;
        }
        var name  = await fetchName(token, sid);
        var peRange = null;
        if (batchExtras.pe) {
          try { peRange = await fetchPERange(sid, token, 3); } catch (peErr) { /* PER抓不到就顯示無資料，不影響其他分析 */ }
        }
        var revRange = null, revFeats = null;
        if (batchExtras.rev) {
          try { revRange = await fetchRevenueYoYMoM(sid, token); } catch (revErr) { /* 營收抓不到就顯示無資料，不影響其他分析 */ }
          try { revFeats = computeRevFeats(await fetchRevenueHistoryForBacktest(sid, token, 800), new Date().toISOString().slice(0, 10)); } catch (rfErr) { /* 略過 */ }
        }
        var priceYoYRange = null;
        if (batchExtras.pxyoy) {
          try { priceYoYRange = await fetchMonthlyAvgPriceYoY(sid, token, revRange); } catch (pyErr) { /* 均價YoY抓不到就顯示無資料，不影響其他分析 */ }
        }
        var instRange = null, instRangeErr = null;
        if (batchExtras.inst) {
          try { instRange = await fetchInstitutionalMonthly(sid, token); } catch (instErr) { instRangeErr = instErr.message; }
        }
        var dataLong = enrich(rawData);
        var techX = computeTechExtras(dataLong, dataLong.length - 1);
        var cutDate = getDateRange(days).start;
        rawData = rawData.filter(function(d) { return d.date >= cutDate; });   // 其餘分析仍只用使用者設定的天數
        var data  = enrich(rawData);
        var dmi = calcDMI(data, 14);
        var dm = scoreDMI(dmi);
        var pb=checkPullbackBuy(data);
        var pt=detectPatterns(data, pb);
        var total=dm.score;
        batchResults.push({ tech:techX, stockId:sid, name:name, data:data, dm:dm, pb:pb, pt:pt, total:total, realtime:rtInjected, peRange:peRange, revRange:revRange, revFeats:revFeats, priceYoYRange:priceYoYRange, instRange:instRange, instRangeErr:instRangeErr });
        if (instRangeErr) batchLog.innerHTML += '<div style="color:#f0a500">  ⚠️ ' + sid + ' 三大法人買賣超讀取失敗：' + instRangeErr + '</div>';
        batchLog.innerHTML += '<div style="color:#00c864">  ✅ ' + sid + ' ' + name + '  得分:' + total + (rtInjected ? '　🔴即時' : '') + '</div>';
      } catch(ex) {
        batchLog.innerHTML += '<div style="color:#ff8888">  ❌ ' + sid + ' 失敗：' + ex.message + '</div>';
      }
      progText.textContent = (i+1) + ' / ' + stocks.length;
      progBar.style.width = ((i+1)/stocks.length*100) + '%';
      // 小延遲避免 API 速率限制
      if (i < stocks.length-1) await new Promise(function(r){ setTimeout(r, 600); });
    }

    btnRun.disabled = false;
    btnRun.textContent = '🔍 批次分析';

    if (!batchResults.length) { showError('所有股票均無法取得資料'); return; }

    // 依總分排序（高→低）
    batchResults.sort(function(a,b){ return b.total - a.total; });
    if (document.getElementById('iAutoTrack').checked) {
      var autoN = appendHitLog(buildHitRecords());
      document.getElementById('batchStatus').textContent = '📌 命中組合 ' + buildHitRecords().length + ' 檔，新增 ' + autoN + ' 筆追蹤紀錄';
    }
    renderBatch();
  }

  // ── 批次結果渲染 ──
  // ── 篩選與排序狀態 ──
  var filterState = { pb: 'all', score: 'all', pt: 'all', rg: 'all', kw: '', sortCol: 'total', sortDir: -1 };

  function applyFilter() {
    var rows = document.querySelectorAll('#summaryTbody tr');
    var visCount = 0;
    rows.forEach(function(row) {
      var idx  = parseInt(row.dataset.idx, 10);
      var r    = batchResults[idx];
      var last = r.data[r.data.length-1];
      var prev = r.data[r.data.length-2]||last;

      var okPb    = filterState.pb    === 'all' || (filterState.pb === 'pass' && r.pb.allPass) || (filterState.pb === 'fail' && !r.pb.allPass);
      var okScore = filterState.score === 'all' ||
        (filterState.score === '80' && r.total >= 80) ||
        (filterState.score === '65' && r.total >= 65 && r.total < 80) ||
        (filterState.score === '50' && r.total >= 50 && r.total < 65) ||
        (filterState.score === 'lt50' && r.total < 50);
      var okKw    = !filterState.kw || r.stockId.includes(filterState.kw) || r.name.includes(filterState.kw);
      var okPt    = filterState.pt === 'all' ||
        (filterState.pt === 'breakout'  && r.pt.anyBreakout) ||
        (filterState.pt === 'justbreak' && r.pt.anyJustBroke) ||
        (filterState.pt === 'forming'   && r.pt.anyFormed && !r.pt.anyBreakout) ||
        (filterState.pt === 'none'      && !r.pt.anyFormed);

      var okRg = true;
      if (filterState.rg && filterState.rg !== 'all') {
        var rf = regimeFit(r);
        okRg = filterState.rg === 'ok' ? rf.ok > 0 : rf.bad === 0;
      }
      var show = okPb && okScore && okKw && okPt && okRg;
      row.classList.toggle('hidden-row', !show);
      if (show) visCount++;
    });
    var ct = document.getElementById('filterCount');
    if (ct) ct.textContent = '顯示 ' + visCount + ' / ' + batchResults.length + ' 檔';
  }

  function sortTable(col) {
    if (filterState.sortCol === col) filterState.sortDir *= -1;
    else { filterState.sortCol = col; filterState.sortDir = -1; }
    // Update sort icons
    document.querySelectorAll('.sort-icon').forEach(function(el){ el.textContent='⇅'; el.classList.remove('asc','desc'); });
    var icon = document.getElementById('sort_'+col);
    if (icon) { icon.textContent = filterState.sortDir===-1?'↓':'↑'; icon.classList.add(filterState.sortDir===-1?'desc':'asc'); }

    var tbody = document.getElementById('summaryTbody');
    var rows  = Array.from(tbody.querySelectorAll('tr'));
    rows.sort(function(a, b) {
      var ia = parseInt(a.dataset.idx), ib = parseInt(b.dataset.idx);
      var ra = batchResults[ia], rb = batchResults[ib];
      var va, vb;
      if (col==='stock') { va = ra.stockId; vb = rb.stockId; }
      else if (col==='name') { va = ra.name || ''; vb = rb.name || ''; }
      else if (col==='price') { va = ra.data[ra.data.length-1].close || 0; vb = rb.data[rb.data.length-1].close || 0; }
      else if (col==='sphits') {
        var getSpHits = function(r) { try { return matchedStockPickCombos(r).length; } catch (e) { return 0; } };
        va = getSpHits(ra); vb = getSpHits(rb);
      }
      else if (col==='total')  { va=ra.total; vb=rb.total; }
      else if (col==='plusDI'){ va=ra.dm.plusDI||0; vb=rb.dm.plusDI||0; }
      else if (col==='minusDI'){ va=ra.dm.minusDI||0; vb=rb.dm.minusDI||0; }
      else if (col==='adx'){ va=ra.dm.adx||0; vb=rb.dm.adx||0; }
      else if (col==='adxr'){ va=ra.dm.adxr||0; vb=rb.dm.adxr||0; }
      else if (col==='bbpos'){
        var getBBPos = function(r) {
          var l = r.data[r.data.length-1];
          if (l.bbU == null || l.bbL == null) return 50;
          var w = l.bbU - l.bbL;
          return w ? (l.close - l.bbL) / w * 100 : 50;
        };
        va = getBBPos(ra); vb = getBBPos(rb);
      }
      else if (col==='volraw'){
        var lva=ra.data[ra.data.length-1], lvb=rb.data[rb.data.length-1];
        va = lva.volume||0; vb = lvb.volume||0;
      }
      else if (col==='chg'){
        var la=ra.data[ra.data.length-1], pa=ra.data[ra.data.length-2]||la;
        var lb=rb.data[rb.data.length-1], pb=rb.data[rb.data.length-2]||lb;
        va=pa.close?(la.close-pa.close)/pa.close:0;
        vb=pb.close?(lb.close-pb.close)/pb.close:0;
      }
      else if (col==='pe'){
        va = ra.peRange ? ra.peRange.current : 0;
        vb = rb.peRange ? rb.peRange.current : 0;
      }
      else if (col==='yoy'){
        va = (ra.revRange && ra.revRange.length) ? (ra.revRange[ra.revRange.length-1].yoy||0) : 0;
        vb = (rb.revRange && rb.revRange.length) ? (rb.revRange[rb.revRange.length-1].yoy||0) : 0;
      }
      else if (col==='pxyoy'){
        va = (ra.priceYoYRange && ra.priceYoYRange.length) ? (ra.priceYoYRange[ra.priceYoYRange.length-1].yoy||0) : 0;
        vb = (rb.priceYoYRange && rb.priceYoYRange.length) ? (rb.priceYoYRange[rb.priceYoYRange.length-1].yoy||0) : 0;
      }
      else if (col==='div'){
        var getDivTotal = function(r) {
          if (!(r.revRange && r.revRange.length && r.priceYoYRange && r.priceYoYRange.length)) return 0;
          var pxByKey = {};
          r.priceYoYRange.forEach(function(p) { pxByKey[p.year + '-' + p.month] = p; });
          var sum = 0, any = false;
          r.revRange.forEach(function(rv) {
            var px = pxByKey[rv.year + '-' + rv.month];
            if (rv.yoy != null && px && px.yoy != null) { sum += rv.yoy - px.yoy; any = true; }
          });
          return any ? sum : 0;
        };
        va = getDivTotal(ra); vb = getDivTotal(rb);
      }
      else if (col==='mom'){
        va = (ra.revRange && ra.revRange.length) ? (ra.revRange[ra.revRange.length-1].mom||0) : 0;
        vb = (rb.revRange && rb.revRange.length) ? (rb.revRange[rb.revRange.length-1].mom||0) : 0;
      }
      else if (col==='inst'){
        var getInstTotal = function(r) {
          if (!(r.instRange && r.instRange.length)) return 0;
          return r.instRange.reduce(function(sum, m) { return sum + m.net; }, 0);
        };
        va = getInstTotal(ra); vb = getInstTotal(rb);
      }
      else if (col==='cstyle'){
        var sa = comboStyleOf(ra), sb = comboStyleOf(rb);
        va = sa ? COMBO_STYLE_LOOK[sa][2] : 0; vb = sb ? COMBO_STYLE_LOOK[sb][2] : 0;
      }
      else if (col==='regfit'){
        var fa = regimeFit(ra), fb2 = regimeFit(rb);
        va = fa.ok - fa.bad; vb = fb2.ok - fb2.bad;
      }
      else if (col==='hitcats'){
        va = hitCategories(ra).length; vb = hitCategories(rb).length;
      }
      else if (col==='combohits'){
        var getComboHits = function(r) {
          var n = 0;
          try { n += matchedPinnedCombos(r).length; } catch (e) {}
          try { n += matchedMoonshotCombos(r).length; } catch (e) {}
          try { n += matchedBtWinCombos(r).length + matchedBtHotCombos(r).length; } catch (e) {}
          return n;
        };
        va = getComboHits(ra); vb = getComboHits(rb);
      }
      var cmp = (col === 'stock' || col === 'name') ? String(va).localeCompare(String(vb), undefined, { numeric: true, sensitivity: 'base' }) : (va - vb);
      return cmp * filterState.sortDir;
    });
    rows.forEach(function(r){ tbody.appendChild(r); });
  }

  function setFilter(type, val) {
    filterState[type] = val;
    // Update active button style
    document.querySelectorAll('.flt-'+type).forEach(function(el){ el.classList.toggle('active', el.dataset.val===val); });
    applyFilter();
  }

  // 摘要表欄寬：依欄位名稱給相對權重，換算成百分比，整張表剛好塞滿畫面寬度（不需要左右捲動）；
  // 內容太長就自動換行。勾選的附加欄位（P/E、營收、法人…）越多，每欄分到的寬度越窄。
  var SUMMARY_COL_WEIGHTS = {
    '股票': 52, '名稱': 84, '股價': 56, '漲跌幅': 56, '命中類別': 44, '組合型態': 58, '大盤適用': 64, '選股型命中': 70, '指定組合命中': 110, '型態確認': 100,
    '多方力道': 50, '布林通道位置': 64, 'MACD狀態': 92, 'KDJ狀態': 80, '相對強弱/量比': 72, '分價量表': 104,
    '成交量': 60, '回後買進場': 72, '近3年P/E區間': 76,
    '近3月營收YoY': 104, '近3月均價YoY': 104, '近3月YoY乖離度': 112, '近3月營收MoM': 104, '三大法人買賣超(近3月)': 112,
    '近3季營收YoY': 104, '近3季均價YoY': 104, '近3季YoY乖離度': 112, '近3季營收QoQ': 104
  };
  function fitSummaryColumns() {
    var ths = document.querySelectorAll('#summaryThead th');
    if (!ths.length) return;
    var labels = Array.prototype.map.call(ths, function(th) {
      var c = th.cloneNode(true);
      c.querySelectorAll('.sort-icon').forEach(function(s) { s.remove(); });
      return c.textContent.trim();
    });
    var ws = labels.map(function(l) { return SUMMARY_COL_WEIGHTS[l] || 80; });
    var tot = ws.reduce(function(a, b) { return a + b; }, 0);
    Array.prototype.forEach.call(ths, function(th, i) { th.style.width = (ws[i] / tot * 100).toFixed(2) + '%'; });
  }

  function renderBatch() {
    if (!batchResults.length) return;

    // 摘要表 rows（data-idx 保留原始排序索引）
    var summaryRows = batchResults.map(function(r, i) {
      var scoreColor = r.total>=80?'#00c864':r.total>=65?'#2196f3':r.total>=50?'#f0a500':'#ff5555';
      var scoreLbl   = r.total>=80?'積極做多':r.total>=65?'可考慮進場':r.total>=50?'觀望':'不建議進場';
      var scoreCls   = r.total>=80?'badge-bull':r.total>=65?'badge-bull':r.total>=50?'badge-obs':'badge-bear';
      var pbIcon     = r.pb.allPass ? '✅' : '❌';
      var pbCls      = r.pb.allPass ? 'badge-pass' : 'badge-fail';
      var pbTxt      = r.pb.allPass ? '符合進場' : r.pb.requiredPassed+'/'+r.pb.requiredTotal+' 通過';
      var ptIcon     = r.pt.anyJustBroke ? '🔥' : (r.pt.anyBreakout ? '✅' : (r.pt.anyFormed ? '🕒' : '－'));
      var ptCls      = r.pt.anyJustBroke ? 'badge-justbreak' : (r.pt.anyBreakout ? 'badge-pass' : (r.pt.anyFormed ? 'badge-obs' : 'badge-fail'));
      var ptNames    = r.pt.results.filter(function(x){ return r.pt.anyBreakout ? x.breakout : x.formed; }).map(function(x){ return (x.justBroke ? '🔥' : '') + x.name + (HIGH_CONSOLIDATION_IDS[x.id] ? '(高檔整理)' : ''); });
      var ptTxt      = ptNames.length ? ptNames.join('、') : '無';
      var last       = r.data[r.data.length-1];
      var prev       = r.data[r.data.length-2]||last;
      var chgp       = prev.close ? (last.close-prev.close)/prev.close*100 : 0;
      var chgColor   = chgp>=0?'#ef5350':'#26a69a';
      var peHtml     = '<span style="color:#666">無資料</span>';
      if (r.peRange) {
        var pr = r.peRange;
        var peSpan = pr.max - pr.min;
        var pePos = peSpan ? (pr.current - pr.min) / peSpan * 100 : 50;
        var peColor = pePos < 25 ? '#00c864' : pePos > 75 ? '#ff5555' : '#f0a500';
        peHtml = '<span style="color:'+peColor+';font-weight:700">' + pr.current.toFixed(1) + '</span>' +
          ' <span style="color:#888;font-size:11px">(' + pr.min.toFixed(1) + '~' + pr.max.toFixed(1) + ')</span>';
      }
      var yoyHtml = '<span style="color:#666">無資料</span>';
      var momHtml = '<span style="color:#666">無資料</span>';
      if (r.revRange && r.revRange.length) {
        var rv3 = r.revRange;
        var latest = rv3[rv3.length - 1];
        var priorMonths = rv3.slice(0, -1).reverse();
        var fmtPct = function(v) {
          if (v == null) return '<span style="color:#666">N/A</span>';
          var c = v >= 0 ? '#ef5350' : '#26a69a';
          return '<span style="color:'+c+'">' + (v>=0?'+':'') + v.toFixed(1) + '%</span>';
        };
        var fmtPctPlain = function(v) { return v!=null ? ((v>=0?'+':'')+v.toFixed(1)+'%') : 'N/A'; };

        yoyHtml = latest.month + '月 YoY ' + fmtPct(latest.yoy);
        var yoySub = priorMonths.map(function(m) { return m.month + '月 ' + fmtPctPlain(m.yoy); }).join('　');
        if (yoySub) yoyHtml += '<br><span style="color:#888;font-size:11px">' + yoySub + '</span>';

        momHtml = latest.month + '月 MoM ' + fmtPct(latest.mom);
        var momSub = priorMonths.map(function(m) { return m.month + '月 ' + fmtPctPlain(m.mom); }).join('　');
        if (momSub) momHtml += '<br><span style="color:#888;font-size:11px">' + momSub + '</span>';
      }
      // 三大法人買賣超（近3月，逐月加總）：主要顯示改成「近3月合計」，排序也用這個
      // 加總值，跟其他「近3月」欄位維持一致的呈現方式。下面小字保留3個月各自數字。
      // 顏色跟BBW掃描工具同一套慣例：正值（買超）＝綠，負值（賣超）＝紅。
      var instHtml = '<span style="color:#666">無資料</span>';
      if (r.instRange && r.instRange.length) {
        var instTotal = r.instRange.reduce(function(sum, m) { return sum + m.net; }, 0);
        var fmtInst = function(v) {
          var c = v >= 0 ? '#00c864' : '#ff5555';
          return '<span style="color:'+c+';font-weight:700">' + (v>=0?'+':'') + Math.round(v).toLocaleString() + '</span>';
        };
        var fmtInstPlain = function(v) { return (v>=0?'+':'') + Math.round(v).toLocaleString(); };
        instHtml = '近3月合計 ' + fmtInst(instTotal);
        var instMonthsDesc = r.instRange.slice().reverse();
        var instSub = instMonthsDesc.map(function(m) { return m.month + '月 ' + fmtInstPlain(m.net); }).join('　');
        if (instSub) instHtml += '<br><span style="color:#888;font-size:11px">' + instSub + '</span>';
      } else if (r.instRangeErr) {
        instHtml = '<span style="color:#f0a500;font-size:11px" title="' + r.instRangeErr.replace(/"/g, '&quot;') + '">⚠️ 讀取失敗</span>';
      }
      var pxYoyHtml = '<span style="color:#666">無資料</span>';
      if (r.priceYoYRange && r.priceYoYRange.length) {
        var py3 = r.priceYoYRange;
        var pyLatest = py3[py3.length - 1];
        var pyPrior = py3.slice(0, -1).reverse();
        var fmtPct2 = function(v) {
          if (v == null) return '<span style="color:#666">N/A</span>';
          var c = v >= 0 ? '#ef5350' : '#26a69a';
          return '<span style="color:'+c+'">' + (v>=0?'+':'') + v.toFixed(1) + '%</span>';
        };
        var fmtPct2Plain = function(v) { return v!=null ? ((v>=0?'+':'')+v.toFixed(1)+'%') : 'N/A'; };
        pxYoyHtml = pyLatest.month + '月 均價YoY ' + fmtPct2(pyLatest.yoy);
        var pxSub = pyPrior.map(function(m) { return m.month + '月 ' + fmtPct2Plain(m.yoy); }).join('　');
        if (pxSub) pxYoyHtml += '<br><span style="color:#888;font-size:11px">' + pxSub + '</span>';
      }
      // YoY乖離度＝營收YoY − 均價YoY，3個月都算（不是只算最新月）。
      // 用(year,month)配對，不用陣列位置對應，避免兩邊月份萬一沒對齊時算錯。
      // 主要顯示改成「近3月加總乖離度」（3個月各自的乖離度加總），排序也改用這個
      // 加總值——單一月份容易受單月雜訊干擾，加總後比較能看出持續性的乖離趨勢。
      // 3個月各自的乖離度數字還是保留顯示（在加總值下面）。
      // 正值大＝營收成長比股價快（可能還沒完全反映在股價上）；負值大＝股價漲幅超前營收成長（留意是否透支）。
      var divHtml = '<span style="color:#666">需同時勾營收與均價YoY</span>';
      if (r.revRange && r.revRange.length && r.priceYoYRange && r.priceYoYRange.length) {
        var pxByKey = {};
        r.priceYoYRange.forEach(function(p) { pxByKey[p.year + '-' + p.month] = p; });
        var divList = r.revRange.map(function(rv) {
          var px = pxByKey[rv.year + '-' + rv.month];
          var d = (rv.yoy != null && px && px.yoy != null) ? (rv.yoy - px.yoy) : null;
          return { year: rv.year, month: rv.month, div: d };
        });
        var validDivs = divList.filter(function(d) { return d.div != null; });
        var divTotal = validDivs.length ? validDivs.reduce(function(sum, d) { return sum + d.div; }, 0) : null;
        var fmtDiv = function(d) {
          if (d == null) return '<span style="color:#666">N/A</span>';
          var c = d > 45 ? '#00c864' : d < -45 ? '#ff5555' : '#f0a500';
          return '<span style="color:'+c+';font-weight:700">' + (d>=0?'+':'') + d.toFixed(1) + 'pp</span>';
        };
        var fmtDivPlain = function(d) { return d!=null ? ((d>=0?'+':'')+d.toFixed(1)+'pp') : 'N/A'; };
        if (divTotal != null) {
          var divLbl = divTotal > 45 ? '　💚 營收優於股價' : divTotal < -45 ? '　⚠️ 股價超前營收' : '　大致同步';
          divHtml = '近3月合計 ' + fmtDiv(divTotal) + divLbl;
        } else {
          divHtml = '<span style="color:#666">當月資料不足</span>';
        }
        var divMonthsDesc = divList.slice().reverse();
        var divSub = divMonthsDesc.map(function(d) { return d.month + '月 ' + fmtDivPlain(d.div); }).join('　');
        if (divSub) divHtml += '<br><span style="color:#888;font-size:11px">' + divSub + '</span>';
      }
      // 布林通道股價位置：跟個股詳細面板用同一套算法（不用多打API，last.bbU/bbL
      // 在enrich()時就已經算好了）。0%=貼著下軌，100%=貼著上軌，50%=通道中央。
      var bbPosHtml = '<span style="color:#666">無資料</span>';
      if (last.bbU != null && last.bbL != null) {
        var bbWidthRow = last.bbU - last.bbL;
        var bbPosRow = bbWidthRow ? (last.close - last.bbL) / bbWidthRow * 100 : 50;
        var bbPosColor = bbPosRow > 80 ? '#ff5555' : bbPosRow < 20 ? '#00c864' : '#f0a500';
        var bbWidthPctRow = last.close ? bbWidthRow / last.close * 100 : 0;
        bbPosHtml = '<span style="color:'+bbPosColor+';font-weight:700">' + bbPosRow.toFixed(0) + '%</span>' +
          ' <span style="color:#888;font-size:11px">(寬度' + bbWidthPctRow.toFixed(1) + '%)</span>';
      }

      // ── MACD狀態 ＋ 布林通道×MACD 情境判斷（強勢突破盤／跌深反彈盤）──
      // 這是課程常見的「布林通道找位置、MACD做確認」的組合判斷法：
      //   強勢突破盤＝通道開口放大＋股價貼近上軌（布林找爆發）
      //             ＋MACD在零軸上且紅柱持續增長（MACD確認不是假突破）
      //   跌深反彈盤＝股價貼近下軌（布林找相對低點）
      //             ＋MACD剛黃金交叉或綠柱縮短（MACD確認下跌動能衰竭）
      // 這裡的「動能衰竭」是用MACD柱狀圖轉折＋黃金交叉當代理指標，不是嚴格定義
      // 的價格/指標背離偵測（真正的背離要比對兩個價格低點，這裡沒有做那麼細）。
      var macdStateTxt = '<span style="color:#666">無資料</span>';
      var comboTagHtml = '';
      if (last.macd != null && last.macdSig != null && last.macdHist != null && prev.macdHist != null) {
        var aboveZero = last.macd > 0;
        var histGrowing = last.macdHist > prev.macdHist;
        var justGoldenCross = prev.macd != null && prev.macdSig != null && prev.macd <= prev.macdSig && last.macd > last.macdSig;
        // 近3天內的黃金交叉（不是只看今天這一根，容許背離成形後晚1-2天才交叉確認）
        var goldenCrossRecent = false;
        for (var gci = Math.max(1, r.data.length - 3); gci < r.data.length; gci++) {
          var gcCur = r.data[gci], gcPrev = r.data[gci-1];
          if (gcCur.macd != null && gcCur.macdSig != null && gcPrev.macd != null && gcPrev.macdSig != null &&
              gcPrev.macd <= gcPrev.macdSig && gcCur.macd > gcCur.macdSig) { goldenCrossRecent = true; break; }
        }

        if (aboveZero && last.macdHist > 0) {
          macdStateTxt = histGrowing
            ? '<span style="color:#ff5555">零軸上・紅柱增長</span>'
            : '<span style="color:#f0a500">零軸上・紅柱縮短</span>';
        } else if (!aboveZero && last.macdHist < 0) {
          macdStateTxt = histGrowing
            ? '<span style="color:#f0a500">零軸下・綠柱縮短</span>'
            : '<span style="color:#00c864">零軸下・綠柱增長</span>';
        } else {
          macdStateTxt = '<span style="color:#888">交叉轉換中</span>';
        }
        if (justGoldenCross) macdStateTxt += '<br><span style="color:#888;font-size:11px">⚡剛黃金交叉</span>';

        // 通道開口是否放大：跟5天前的通道寬度比較
        var widthNowPct = null, widthExpanding = false;
        if (last.bbU != null && last.bbL != null) {
          widthNowPct = last.close ? (last.bbU - last.bbL) / last.close * 100 : 0;
          var refIdx = r.data.length - 6;
          var refBar = refIdx >= 0 ? r.data[refIdx] : null;
          if (refBar && refBar.bbU != null && refBar.bbL != null && refBar.close) {
            var widthRefPct = (refBar.bbU - refBar.bbL) / refBar.close * 100;
            widthExpanding = widthNowPct > widthRefPct;
          }
        }
        var bbPosForCombo = (last.bbU != null && last.bbL != null && (last.bbU - last.bbL))
          ? (last.close - last.bbL) / (last.bbU - last.bbL) * 100 : null;

        var isBreakout = bbPosForCombo != null && bbPosForCombo >= 80 && widthExpanding && aboveZero && last.macdHist > 0 && histGrowing;

        // ── MACD低檔背離（真正的背離偵測，不是代理指標）──
        // 用跟型態辨識、圖表轉折波同一套 buildZigzag 找出最近兩個轉折低點：
        // 如果「股價創更低（後面那個低點價格更低）」但「MACD的DIF值卻是更高
        // （沒有跟著創新低）」，就是典型的「量價背離／指標背離」，代表下跌力道
        // 已經在衰竭，即使股價還在破底。只看最近60根K棒內的低點，太久以前的
        // 轉折點對現在已經沒有參考意義。
        var divergenceDetected = false;
        var zzForDiv = buildZigzag(r.data);
        var zzLows = zzForDiv.filter(function(p){ return p.type === 'L'; });
        if (zzLows.length >= 2) {
          var recentLow = zzLows[zzLows.length - 1];
          var priorLow = zzLows[zzLows.length - 2];
          var withinLookback = recentLow.idx >= r.data.length - 1 - 60;
          var macdAtRecent = r.data[recentLow.idx] ? r.data[recentLow.idx].macd : null;
          var macdAtPrior = r.data[priorLow.idx] ? r.data[priorLow.idx].macd : null;
          if (withinLookback && macdAtRecent != null && macdAtPrior != null) {
            var priceLowerLow = recentLow.price < priorLow.price;
            var macdHigherLow = macdAtRecent > macdAtPrior;
            divergenceDetected = priceLowerLow && macdHigherLow;
          }
        }

        // 跌深反彈盤＝股價貼近下軌＋MACD低檔背離＋（近期）黃金交叉，三個條件都要
        // 成立，對應圖裡「低檔背離+黃金交叉」的雙重確認，不是只看代理訊號。
        var isPullbackRebound = bbPosForCombo != null && bbPosForCombo <= 20 && divergenceDetected && goldenCrossRecent;

        if (isBreakout) {
          comboTagHtml = '<br><span style="display:inline-block;margin-top:3px;padding:1px 8px;border-radius:10px;font-size:11px;color:#ff5555;border:1px solid #ff555555">🚀 強勢突破盤</span>';
        } else if (isPullbackRebound) {
          var divTip = 'MACD低檔背離：前低點DIF=' + macdAtPrior.toFixed(3) + '(價' + priorLow.price.toFixed(1) + ') → 新低點DIF=' + macdAtRecent.toFixed(3) + '(價' + recentLow.price.toFixed(1) + ')，股價破底但DIF未破底';
          comboTagHtml = '<br><span style="display:inline-block;margin-top:3px;padding:1px 8px;border-radius:10px;font-size:11px;color:#00c864;border:1px solid #00c86455" title="'+divTip+'">🎯 跌深反彈盤</span>';
        }
      }

      // KDJ(6,3,3) 黃金交叉／死亡交叉狀態
      var kdjStateHtml = '<span style="color:#666">無資料</span>';
      if (last.kdjK != null && last.kdjD != null) {
        var kdjBullish = last.kdjK > last.kdjD;
        var kdjColor = kdjBullish ? '#00c864' : '#ff5555';
        kdjStateHtml = '<span style="color:'+kdjColor+'">' + (kdjBullish ? 'K>D（多方）' : 'K<D（空方）') + '</span>';
        var kdjGold = false, kdjDead = false;
        for (var kdi = Math.max(1, r.data.length - 3); kdi < r.data.length; kdi++) {
          var kdCur = r.data[kdi], kdPrev = r.data[kdi-1];
          if (kdCur.kdjK != null && kdCur.kdjD != null && kdPrev.kdjK != null && kdPrev.kdjD != null) {
            if (kdPrev.kdjK <= kdPrev.kdjD && kdCur.kdjK > kdCur.kdjD) kdjGold = true;
            if (kdPrev.kdjK >= kdPrev.kdjD && kdCur.kdjK < kdCur.kdjD) kdjDead = true;
          }
        }
        if (kdjGold) kdjStateHtml += '<br><span style="color:#888;font-size:11px">⚡近3日黃金交叉</span>';
        if (kdjDead) kdjStateHtml += '<br><span style="color:#888;font-size:11px">💀近3日死亡交叉</span>';
      }

      // 相對強弱(vs大盤0050,20日) + 成交量確認(量比)
      var rsVolHtml = '';
      var rs20 = currentBenchmark ? computeRelativeStrength(r.data, currentBenchmark, 20) : null;
      if (rs20 != null) {
        var rsColor = rs20 > 0 ? '#00c864' : '#ff5555';
        rsVolHtml += '<span style="color:'+rsColor+'">RS' + (rs20>=0?'+':'') + rs20.toFixed(1) + '%</span>';
      } else {
        rsVolHtml += '<span style="color:#666">RS無資料</span>';
      }
      var vRatio = computeVolRatio(r.data);
      if (vRatio != null) {
        var volColor = vRatio >= 2 ? '#ff5555' : (vRatio >= 1.5 ? '#f0a500' : '#888');
        rsVolHtml += '<br><span style="color:'+volColor+'">量比' + vRatio.toFixed(1) + 'x' + (vRatio>=1.5?' 🔥':'') + '</span>';
      }

      // 個股分價量表（Volume Profile）訊號
      var vpProfile = computeVolumeProfile(r.data, VP_PARAMS.lookback, VP_PARAMS.bins);
      var vpResult = detectVolumeProfileSignal(r.data, vpProfile);
      var vpHtml = '';
      if (vpProfile) {
        vpHtml += '<span style="color:#888;font-size:11px">POC ' + vpProfile.pocPrice.toFixed(1) + '（' + vpProfile.pocLow.toFixed(1) + '~' + vpProfile.pocHigh.toFixed(1) + '）</span>';
        if (vpResult.signal === 'buy') {
          vpHtml += '<br><span style="color:#00c864">🟢 ' + vpResult.detail.split('：')[0] + '</span>';
        } else if (vpResult.signal === 'sell') {
          vpHtml += '<br><span style="color:#ff5555">🔴 ' + vpResult.detail.split('：')[0] + '</span>';
        }
      } else {
        vpHtml = '<span style="color:#666">無資料</span>';
      }

      return '<tr data-idx="'+i+'" title="點擊查看詳細分析" style="cursor:pointer">' +
        '<td class="nw"><strong style="color:#fff">' + r.stockId + '</strong></td>' +
        '<td style="color:#ccc">' + r.name + (r.realtime ? ' <span style="color:#ff5555;font-size:10px" title="含盤中即時股價">🔴即時</span>' : '') + '</td>' +
        '<td class="nw" style="font-weight:700">' + last.close.toFixed(2) + '</td>' +
        '<td class="nw" style="color:'+chgColor+'">' + (chgp>=0?'+':'') + chgp.toFixed(2) + '%</td>' +
        renderHitCategories(r) +
        renderComboStyle(r) +
        renderRegimeFit(r) +
        '<td>' + renderMatchedComboBadges(r, 'S') + '</td>' +
        '<td>' + renderMatchedComboBadges(r, 'PM') + '</td>' +
        '<td><span class="'+ptCls+'" title="'+ptTxt+'">' + ptIcon + ' ' + (ptNames.length ? ptTxt : '無') + '</span></td>' +
        '<td class="dmi-cell nw" style="font-weight:700;color:'+scoreColor+'">' + r.total +
          '<div class="dmi-pop">+DI ' + (r.dm.plusDI!=null?r.dm.plusDI.toFixed(1):'--') + '　-DI ' + (r.dm.minusDI!=null?r.dm.minusDI.toFixed(1):'--') +
          '<br>ADX ' + (r.dm.adx!=null?r.dm.adx.toFixed(1):'--') + '　ADXR ' + (r.dm.adxr!=null?r.dm.adxr.toFixed(1):'--') + '</div></td>' +
        '<td>' + bbPosHtml + '</td>' +
        '<td>' + macdStateTxt + comboTagHtml + '</td>' +
        '<td>' + kdjStateHtml + '</td>' +
        '<td>' + rsVolHtml + '</td>' +
        '<td>' + vpHtml + '</td>' +
        '<td class="nw">' + fv(last.volume) + '</td>' +
        (batchExtras.pe ? '<td>' + peHtml + '</td>' : '') +
        (batchExtras.rev ? '<td>' + yoyHtml + '</td>' : '') +
        (batchExtras.pxyoy ? '<td>' + pxYoyHtml + '</td>' : '') +
        (batchExtras.rev && batchExtras.pxyoy ? '<td>' + divHtml + '</td>' : '') +
        (batchExtras.rev ? '<td>' + momHtml + '</td>' : '') +
        (batchExtras.inst ? '<td>' + instHtml + '</td>' : '') +
        '</tr>';
    }).join('');

    var filterBar =
      '<div class="filter-bar" id="filterBar">' +
        '<span style="font-size:12px;color:#888;margin-right:4px">進場條件：</span>' +
        '<button class="filter-btn flt-pb active" data-type="pb" data-val="all">全部</button>' +
        '<button class="filter-btn flt-pb" data-type="pb" data-val="pass">✅ 符合進場</button>' +
        '<button class="filter-btn flt-pb" data-type="pb" data-val="fail">❌ 不符合</button>' +
        '<span class="filter-sep">|</span>' +
        '<span style="font-size:12px;color:#888;margin-right:4px">評分：</span>' +
        '<button class="filter-btn flt-score active" data-type="score" data-val="all">全部</button>' +
        '<button class="filter-btn flt-score" data-type="score" data-val="80" style="color:#00c864">80+</button>' +
        '<button class="filter-btn flt-score" data-type="score" data-val="65" style="color:#2196f3">65-79</button>' +
        '<button class="filter-btn flt-score" data-type="score" data-val="50" style="color:#f0a500">50-64</button>' +
        '<button class="filter-btn flt-score" data-type="score" data-val="lt50" style="color:#ff5555">&lt;50</button>' +
        '<span class="filter-sep">|</span>' +
        '<span style="font-size:12px;color:#888;margin-right:4px">型態確認：</span>' +
        '<button class="filter-btn flt-pt active" data-type="pt" data-val="all">全部</button>' +
        '<button class="filter-btn flt-pt" data-type="pt" data-val="breakout" style="color:#00c864">✅ 已突破</button>' +
        '<button class="filter-btn flt-pt" data-type="pt" data-val="justbreak" style="color:#ff9800">🔥 剛突破</button>' +
        '<button class="filter-btn flt-pt" data-type="pt" data-val="forming" style="color:#f0a500">🕒 成形中</button>' +
        '<button class="filter-btn flt-pt" data-type="pt" data-val="none" style="color:#666">－ 無</button>' +
        '<span class="filter-sep">|</span>' +
        '<span style="font-size:12px;color:#888;margin-right:4px">大盤適用：</span>' +
        '<button class="filter-btn flt-rg' + ((filterState.rg || 'all') === 'all' ? ' active' : '') + '" data-type="rg" data-val="all">全部</button>' +
        '<button class="filter-btn flt-rg' + (filterState.rg === 'ok' ? ' active' : '') + '" data-type="rg" data-val="ok" style="color:#4caf50">有✅</button>' +
        '<button class="filter-btn flt-rg' + (filterState.rg === 'nobad' ? ' active' : '') + '" data-type="rg" data-val="nobad" style="color:#ff5555">不含❌</button>' +
        '<span class="filter-sep">|</span>' +
        '<input class="filter-input" id="kwInput" placeholder="搜尋代號/名稱">' +
        '<span id="filterCount" style="font-size:12px;color:#888;margin-left:8px"></span>' +
      '</div>';

    var summary =
      '<div style="background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:12px;overflow:hidden;margin-bottom:20px">' +
        '<div style="padding:12px 16px;font-size:14px;font-weight:700;color:#fff;border-bottom:1px solid rgba(255,255,255,.08);display:flex;justify-content:space-between;align-items:center">' +
          '<span>📋 批次分析摘要</span>' +
          '<span style="display:flex;gap:8px">' +
          '<button type="button" id="btnExportExcel" style="padding:6px 14px;font-size:12px;font-weight:400;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:6px;color:#ddd;cursor:pointer">📥 匯出 Excel</button>' +
          '<button type="button" id="btnExportHits" style="padding:6px 14px;font-size:12px;font-weight:400;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:6px;color:#ddd;cursor:pointer;color:#ff9800;border-color:#ff980066">🎯 匯出命中組合股票（' + buildHitRecords().length + '檔）</button>' +
          '<button type="button" id="btnAddTrack" style="padding:6px 14px;font-size:12px;font-weight:400;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:6px;color:#ddd;cursor:pointer">📌 加入追蹤</button>' +
          '</span>' +
        '</div>' +
        '<div style="padding:8px 16px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;border-bottom:1px solid rgba(255,255,255,.06);font-size:12px;color:#888">一鍵加入清單：' +
          ['進攻', '穩健', '彩券'].map(function(st) {
            var lk = COMBO_STYLE_LOOK[st], n = batchResults.filter(function(r) { return comboStyleOf(r) === st; }).length;
            return '<button type="button" class="btnAddStyleList" data-style="' + st + '"' + (n ? '' : ' disabled') + ' style="padding:5px 12px;font-size:12px;background:rgba(255,255,255,.04);border:1px solid ' + lk[1] + '88;border-radius:6px;color:' + lk[1] + ';cursor:' + (n ? 'pointer' : 'default') + ';opacity:' + (n ? 1 : .45) + '">' + lk[0] + ' → 「' + CUSTOM_LIST_NAMES[STYLE_LIST_KIND[st]] + '」（' + n + '檔）</button>';
          }).join('') +
          '<span style="font-size:11px;color:#666">（加入＝保留原有、重複略過；要重來請先按清單上的 × 清空）</span>' +
        '</div>' +
        (function() {
          var g = batchMarketRegime();
          if (!g) return '';
          var warn = g === 2;
          return '<div style="padding:10px 16px;font-size:13px;color:' + (warn ? '#ff9800' : '#8ec5ff') + ';background:' + (warn ? 'rgba(255,152,0,.08)' : 'rgba(74,158,255,.06)') +
            ';border-bottom:1px solid rgba(255,255,255,.06)">🧭 今日大盤狀態：<b>' + MARKET_REGIMES[g][0] + '</b>（' + MARKET_REGIMES[g][1] + '）　→　' + REGIME_ADVICE[g] +
            '。<span style="color:#888;font-size:11px">「大盤適用」欄：✅＝這種大盤下三年回測有效、❌＝回測為負、⚠️＝普通。</span></div>';
        })() +
        (function() {
          var mp;
          try { mp = marketPanel(); } catch (e) { return ''; }
          return '<div style="padding:8px 16px;font-size:12.5px;color:#cfd8e3;background:rgba(255,255,255,.02);border-bottom:1px solid rgba(255,255,255,.06);line-height:1.8">' +
            mp.lines.map(function(l) { return l[0] + ' ' + l[1] + '：<b style="color:#ffd54f">' + l[2] + '</b>　<span style="color:#9aa3b2">' + l[3] + '</span>'; }).join('<br>') +
            '<br><b style="color:' + mp.verdict[0] + '">' + mp.verdict[1] + '</b>' +
            '<details open style="margin-top:4px"><summary style="cursor:pointer;color:#8ec5ff;font-weight:700">📋 大盤指引（白話解讀與觀察重點）</summary><div style="color:#cfd8e3;padding-left:8px">' + mp.guide.join('<br>') + '</div></details>' +
            '<span style="color:#666;font-size:11px">　（K線位置的歷史統計用等權大盤近似；只是部位參考，不是精準買賣點）</span></div>';
        })() +
        '<details style="padding:8px 16px;font-size:11px;color:#888;border-bottom:1px solid rgba(255,255,255,.05)">' +
          '<summary style="cursor:pointer;color:#aaa">「指定組合命中」編號對照 <span style="color:#7c3aed">[高勝率]</span> 共'+PINNED_COMBOS.length+'組 ／ <span style="color:#ff9800">[高標股]</span> 共'+MOONSHOT_COMBOS.length+'組（點開查看／滑鼠移到徽章上也能看到完整條件）</summary>' +
          '<div style="margin-top:6px;line-height:1.8">' +
          '<div style="color:#00c864;font-weight:700;margin-bottom:2px">⬥ 選股型（S編號，2023-10～2026-09 三年回測分三段，每段扣掉同一天大盤後都仍有超額報酬，最值得看）</div>' +
          (STOCK_PICK_COMBOS.length ? STOCK_PICK_COMBOS.map(function(c, idx){
            return 'S'+(idx+1)+' '+c.join('＋')+' <span style="color:#00c864">('+formatStockPickStats(STOCK_PICK_STATS[c.join(' ＋ ')])+')</span>';
          }).join('<br>') : '<span style="color:#666">（尚未設定，請用這個市場自己的回測，挑 t值(同日調整)≥3 的組合填入 STOCK_PICK_COMBOS）</span>') +
          '<div style="color:#7c3aed;font-weight:700;margin-top:10px;margin-bottom:2px">⬥ 高勝率（#編號；標⏱＝擇時型，勝率主要來自大盤反彈，適合判斷大盤落底，不適合選股）</div>' +
          PINNED_COMBOS.map(function(c, idx){
            var key = c.join(' ＋ ');
            var wrText = formatWinrates(PINNED_COMBO_WINRATES[key]);
            return '#'+(idx+1)+(isTimingCombo(c)?'⏱':'')+' '+c.join('＋')+(wrText?' <span style="color:#7c3aed">(歷史勝率：'+wrText+')</span>':'');
          }).join('<br>') +
          '<div style="color:#00bcd4;font-weight:700;margin-top:10px;margin-bottom:2px">⬥ 回測高勝率（W編號，青色，' + BT_WIN_COMBOS.length + '組：三年回測任一天期勝率>55%，且三年各自依日 t≥2、六個半年至少五個為正；標⏱＝擇時型）</div>' +
          BT_WIN_COMBOS.map(function(c, idx){ return 'W'+(idx+1)+(isTimingCombo(c)?'⏱':'')+' '+c.join('＋')+' <span style="color:#00bcd4;font-size:10px">'+btWinText(BT_WIN_STATS[c.join(' ＋ ')])+'</span>'; }).join('<br>') +
          '<div style="color:#e91e63;font-weight:700;margin-top:10px;margin-bottom:2px">⬥ 回測飆股（F編號，粉紅色，' + BT_HOT_COMBOS.length + '組：10/20日飆股比例>10%、飆股次數≥10，且三段飆股比例都≥2倍基準；摘要表每類只顯示最強的' + BT_HIT_SHOW + '個，其餘以 +N 表示）</div>' +
          BT_HOT_COMBOS.map(function(c, idx){ return 'F'+(idx+1)+' '+c.join('＋')+' <span style="color:#e91e63;font-size:10px">'+formatMoonshotStats(BT_HOT_STATS[c.join(' ＋ ')])+'</span>'; }).join('<br>') +
          '<div style="color:#ddd;font-weight:700;margin-top:10px;margin-bottom:2px">🏷️ 組合型態（依命中類別 S/#/M/W/F 的組合分級；其他組合顯示「－」）</div>' +
          ['進攻', '穩健', '彩券'].map(function(k) {
            var keys = Object.keys(COMBO_STYLE_RULES).filter(function(c) { return COMBO_STYLE_RULES[c] === k; });
            return '<span style="color:' + COMBO_STYLE_LOOK[k][1] + ';font-weight:700">' + COMBO_STYLE_LOOK[k][0] + '</span>＝' + keys.join('、') + ' <span style="color:#888;font-size:10px">（' + COMBO_STYLE_NOTE[k] + '）</span>';
          }).join('<br>') +
          '<div style="color:#ff9800;font-weight:700;margin-top:10px;margin-bottom:2px">⬥ 高標股（M編號，容易命中大行情，不代表整體勝率高）</div>' +
          MOONSHOT_COMBOS.map(function(c, idx){
            var key = c.join(' ＋ ');
            var statsText = formatMoonshotStats(MOONSHOT_COMBO_STATS[key]);
            return 'M'+(idx+1)+' '+c.join('＋')+(statsText?'<br><span style="color:#ff9800;font-size:10px;padding-left:12px">'+statsText+'</span>':'');
          }).join('<br>') +
          '</div>' +
        '</details>' +
        filterBar +
        '<div class="summary-wrap"><table class="summary-table">' +
          '<thead id="summaryThead"><tr>' +
            '<th data-col="stock" style="cursor:pointer">股票 <span class="sort-icon" id="sort_stock">⇅</span></th>' +
            '<th data-col="name" style="cursor:pointer">名稱 <span class="sort-icon" id="sort_name">⇅</span></th>' +
            '<th data-col="price" style="cursor:pointer">股價 <span class="sort-icon" id="sort_price">⇅</span></th>' +
            '<th data-col="chg"   style="cursor:pointer">漲跌幅 <span class="sort-icon" id="sort_chg">⇅</span></th>' +
            '<th data-col="hitcats" style="cursor:pointer" title="S/#/M/W/F 五類中命中幾類">命中類別 <span class="sort-icon" id="sort_hitcats">⇅</span></th>' +
            '<th data-col="cstyle" style="cursor:pointer" title="依命中類別組合分級：進攻／穩健／彩券（滑鼠移到格子上看回測數字）">組合型態 <span class="sort-icon" id="sort_cstyle">⇅</span></th>' +
            '<th data-col="regfit" style="cursor:pointer" title="依今日大盤狀態（①多頭 ②多頭回檔 ③空頭反彈 ④空頭），看命中的組合在這種大盤下的三年回測：✅有效 ❌為負 ⚠️普通（滑鼠移到格子上看編號）">大盤適用 <span class="sort-icon" id="sort_regfit">⇅</span></th>' +
            '<th data-col="sphits" style="cursor:pointer">選股型命中 <span class="sort-icon" id="sort_sphits">⇅</span></th>' +
            '<th data-col="combohits" style="cursor:pointer">指定組合命中 <span class="sort-icon" id="sort_combohits">⇅</span></th>' +
            '<th>型態確認</th>' +
            '<th data-col="total" style="cursor:pointer"title="滑鼠移到分數上可看 +DI／-DI／ADX／ADXR">多方力道 <span class="sort-icon desc" id="sort_total">↓</span></th>' +
            '<th data-col="bbpos"   style="cursor:pointer">布林通道位置 <span class="sort-icon" id="sort_bbpos">⇅</span></th>' +
            '<th>MACD狀態</th>' +
            '<th>KDJ狀態</th>' +
            '<th>相對強弱/量比</th>' +
            '<th>分價量表</th>' +
            '<th data-col="volraw" style="cursor:pointer">成交量 <span class="sort-icon" id="sort_volraw">⇅</span></th>' +
            (batchExtras.pe ? '<th data-col="pe"    style="cursor:pointer">近3年P/E區間 <span class="sort-icon" id="sort_pe">⇅</span></th>' : '') +
            (batchExtras.rev ? '<th data-col="yoy"   style="cursor:pointer">近3月營收YoY <span class="sort-icon" id="sort_yoy">⇅</span></th>' : '') +
            (batchExtras.pxyoy ? '<th data-col="pxyoy" style="cursor:pointer">近3月均價YoY <span class="sort-icon" id="sort_pxyoy">⇅</span></th>' : '') +
            (batchExtras.rev && batchExtras.pxyoy ? '<th data-col="div"   style="cursor:pointer">近3月YoY乖離度 <span class="sort-icon" id="sort_div">⇅</span></th>' : '') +
            (batchExtras.rev ? '<th data-col="mom"   style="cursor:pointer">近3月營收MoM <span class="sort-icon" id="sort_mom">⇅</span></th>' : '') +
            (batchExtras.inst ? '<th data-col="inst"  style="cursor:pointer">三大法人買賣超(近3月) <span class="sort-icon" id="sort_inst">⇅</span></th>' : '') +
          '</tr></thead>' +
          '<tbody id="summaryTbody">' + summaryRows + '</tbody>' +
        '</table></div>' +
      '</div>';

    // Tab 標籤
    var tabs = batchResults.map(function(r, i) {
      var cls = r.pb.allPass ? 'pass' : (r.total<50?'fail':'');
      return '<button class="tab-btn ' + cls + (i===0?' active':'') + '" id="tab'+i+'" onclick="switchTab('+i+')">' +
        r.stockId + ' <span style="font-size:11px;opacity:.8">' + r.total + 'pt</span></button>';
    }).join('');

    // 各股詳細 panel
    var panels = batchResults.map(function(r, i) {
      return '<div class="tab-panel' + (i===0?' active':'') + '" id="panel'+i+'"></div>';
    }).join('');

    output.innerHTML = summary +
      '<div class="tab-bar">' + tabs + '</div>' +
      '<div id="tabPanels">' + panels + '</div>';

    fitSummaryColumns();
    applyFilter();

    // Event delegation for filter buttons
    var fb = document.getElementById('filterBar');
    if (fb) {
      fb.addEventListener('click', function(e) {
        var btn = e.target.closest('[data-type]');
        if (!btn) return;
        var type = btn.dataset.type, val = btn.dataset.val;
        filterState[type] = val;
        document.querySelectorAll('.flt-'+type).forEach(function(el){ el.classList.toggle('active', el.dataset.val===val); });
        applyFilter();
      });
      var kw = document.getElementById('kwInput');
      if (kw) kw.addEventListener('input', function(){ filterState.kw=this.value; applyFilter(); });
    }

    // Event delegation for sortable column headers
    var th = document.getElementById('summaryThead');
    if (th) {
      th.addEventListener('click', function(e) {
        var cell = e.target.closest('[data-col]');
        if (!cell) return;
        sortTable(cell.dataset.col);
      });
    }

    // Event delegation for summary table row click → switch tab
    var tbody = document.getElementById('summaryTbody');
    if (tbody) {
      tbody.addEventListener('click', function(e) {
        var row = e.target.closest('tr');
        if (!row || row.classList.contains('hidden-row')) return;
        var idx = parseInt(row.dataset.idx, 10);
        if (!isNaN(idx)) {
          switchTab(idx);
          // Scroll to tab section
          var tabBar = document.querySelector('.tab-bar');
          if (tabBar) tabBar.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      });
    }

    var exportBtn = document.getElementById('btnExportExcel');
    if (exportBtn) exportBtn.addEventListener('click', exportSummaryToExcel);
    var hitsBtn = document.getElementById('btnExportHits');
    if (hitsBtn) hitsBtn.addEventListener('click', function() {
      var recs = buildHitRecords();
      if (!recs.length) { alert('這次批次分析沒有股票命中指定組合'); return; }
      exportHitRecords(recs, '命中組合股票', '命中組合股票');
    });
    document.querySelectorAll('.btnAddStyleList').forEach(function(b) {
      b.addEventListener('click', function() {
        var st = b.dataset.style, kind = STYLE_LIST_KIND[st];
        var ids = batchResults.filter(function(r) { return comboStyleOf(r) === st; }).map(function(r) { return r.stockId; });
        if (!ids.length) return;
        var res = addToCustomList(kind, ids);
        refreshCustomListLabels();
        var old = b.innerHTML;
        b.textContent = '✅ 新增 ' + res[0] + ' 檔，共 ' + res[1] + ' 檔';
        setTimeout(function() { b.innerHTML = old; }, 1800);
      });
    });
    var addTrackBtn = document.getElementById('btnAddTrack');
    if (addTrackBtn) addTrackBtn.addEventListener('click', function() {
      var n = appendHitLog(buildHitRecords());
      addTrackBtn.textContent = '✅ 新增 ' + n + ' 筆';
      setTimeout(function(){ addTrackBtn.textContent = '📌 加入追蹤'; }, 1500);
    });

    switchTab(0);
  }

  // 匯出Excel：直接讀指定的表格（跟你看到的一致——目前的篩選、排序、哪些欄位有
  // 顯示都會反映在匯出結果裡），把每個儲存格的 <br> 換成換行，其餘HTML標籤去掉，
  // 只留文字內容，用SheetJS產生.xlsx檔案下載。selector可以是CSS選擇器字串，也可以
  // 直接傳table元素本身。
  // ── 帶格式的 Excel 匯出（ExcelJS）：每個儲存格保留畫面上的文字顏色、粗體、分行與深色底色，
  //    跟網頁顯示一致；ExcelJS 載入失敗時退回 SheetJS 純文字匯出 ──
  var XL_FONT = 'Microsoft JhengHei';
  function xlParseColor(c) {
    var m = /rgba?\(([^)]+)\)/.exec(c || '');
    if (!m) return null;
    var p = m[1].split(',').map(function(x) { return parseFloat(x); });
    return { r: p[0], g: p[1], b: p[2], a: p.length > 3 ? p[3] : 1 };
  }
  function xlHex(c) {
    return 'FF' + [c.r, c.g, c.b].map(function(v) { return ('0' + Math.round(v).toString(16)).slice(-2); }).join('').toUpperCase();
  }
  // 由外往內疊加各層半透明底色，算出儲存格實際看到的背景色
  function xlEffectiveBg(el) {
    var chain = [];
    for (var n = el; n && n.nodeType === 1; n = n.parentElement) chain.push(n);
    var base = { r: 10, g: 14, b: 26, a: 1 };
    chain.reverse().forEach(function(n) {
      var c = xlParseColor(getComputedStyle(n).backgroundColor);
      if (c && c.a > 0) base = { r: base.r * (1 - c.a) + c.r * c.a, g: base.g * (1 - c.a) + c.g * c.a, b: base.b * (1 - c.a) + c.b * c.a, a: 1 };
    });
    return xlHex(base);
  }
  function xlFontOf(node) {
    var cs = getComputedStyle(node), col = xlParseColor(cs.color);
    var px = parseFloat(cs.fontSize) || 12;
    return { name: XL_FONT, size: Math.max(8, Math.round(px * 0.8)), bold: parseInt(cs.fontWeight, 10) >= 600,
             color: { argb: col ? xlHex(col) : 'FFE0E0E0' } };
  }
  // 把一個 td/th 轉成 ExcelJS 的 richText（每段文字各自的顏色／粗體），<br> 與區塊元素轉成換行
  function xlRichText(el) {
    var runs = [];
    var nl = function() { if (runs.length && !/\n$/.test(runs[runs.length - 1].text)) runs.push({ text: '\n', font: runs[runs.length - 1].font }); };
    (function walk(node) {
      Array.prototype.forEach.call(node.childNodes, function(ch) {
        if (ch.nodeType === 3) {
          var t = ch.textContent.replace(/\s+/g, ' ');
          if (t.trim()) {
            if (runs.length && /\n$/.test(runs[runs.length - 1].text)) t = t.replace(/^ /, '');
            runs.push({ text: t, font: xlFontOf(ch.parentElement) });
          }
        } else if (ch.nodeType === 1) {
          if (ch.classList.contains('dmi-pop') || ch.classList.contains('sort-icon')) return;
          var cs = getComputedStyle(ch);
          if (cs.display === 'none') return;
          if (ch.tagName === 'BR') { nl(); return; }
          var block = cs.display === 'block' || cs.display === 'flex' || ch.tagName === 'DIV';
          if (block) nl();
          walk(ch);
          if (block) nl();
          else if (cs.display === 'inline-block' && runs.length && !/[\s]$/.test(runs[runs.length - 1].text)) runs.push({ text: ' ', font: runs[runs.length - 1].font });
        }
      });
    })(el);
    while (runs.length && !runs[runs.length - 1].text.trim()) runs.pop();
    runs.forEach(function(r) { r.text = r.text.replace(/ +\n/g, '\n'); });
    return runs;
  }
  function xlDownload(wb, filename) {
    return wb.xlsx.writeBuffer().then(function(buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click();
      setTimeout(function() { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    });
  }
  function xlStyleCell(cell, runs, bg, isHeader) {
    if (!runs.length) cell.value = '';
    else if (runs.length === 1) { cell.value = runs[0].text; cell.font = runs[0].font; }
    else { cell.value = { richText: runs }; cell.font = runs[0].font; }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    cell.alignment = { wrapText: true, vertical: isHeader ? 'middle' : 'top' };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF2A2E3A' } } };
  }
  function xlRowHeight(texts, widths) {
    var lines = 1;
    texts.forEach(function(t, i) {
      var w = widths[i] || 10, n = 0;
      String(t).split('\n').forEach(function(seg) {
        var len = 0; for (var k = 0; k < seg.length; k++) len += seg.charCodeAt(k) > 255 ? 2 : 1;
        n += Math.max(1, Math.ceil(len / Math.max(w - 1, 4)));
      });
      if (n > lines) lines = n;
    });
    return Math.min(15 * lines + 4, 400);
  }
  // ── 匯出附加內容（2026-10-06）：表頭放大盤摘要表、右側加上所有技術訊號條件欄、另附「t值與組合型態索引」工作表 ──
  function xlF(t, col, bold, size) { return { text: t, font: { name: XL_FONT, size: size || 10, bold: !!bold, color: { argb: col || 'FFE0E0E0' } } }; }
  function xlCats() {
    return [['S', '選股型', 'FF00C864', STOCK_PICK_COMBOS], ['#', '高勝率', 'FF7C3AED', PINNED_COMBOS], ['M', '高標股', 'FFFF9800', MOONSHOT_COMBOS],
            ['W', '回測高勝率', 'FF00BCD4', BT_WIN_COMBOS], ['F', '回測飆股', 'FFE91E63', BT_HOT_COMBOS]];
  }
  var XL_TAG_COLORS = { '✅': 'FF4CAF50', '❌': 'FFFF5555', '⚠️': 'FFF0A500' };
  // 每檔股票今日成立的技術訊號條件（跟「多因子複選搜尋」用的條件清單相同）
  function xlFlagMatrix() {
    var frs = batchResults.map(function(r) { try { return computeLiveFlagsRow(r); } catch (e) { return null; } });
    var valid = frs.filter(Boolean), defs = valid.length ? buildConditionFlags(valid) : {}, names = Object.keys(defs), byId = {};
    batchResults.forEach(function(r, i) {
      var hit = {};
      if (frs[i]) names.forEach(function(n) { try { if (defs[n](frs[i])) hit[n] = 1; } catch (e) {} });
      byId[r.stockId] = hit;
    });
    return { names: names, byId: byId };
  }
  function xlFillRow(ws, r, ncols, bg) {
    for (var c = 1; c <= ncols; c++) ws.getRow(r).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
  }
  // 表頭：標題、今日大盤狀態、四種大盤狀態的摘要表（標出今天）、各清單在今日大盤下的 ✅/⚠️/❌ 數、「大盤適用」欄說明
  function xlWriteRegimeHeader(ws, title, ncols) {
    var BG = 'FF0A0E1A', r = 1, g = batchMarketRegime(), span = Math.max(ncols, 10);
    var put = function(runs, h) {
      xlFillRow(ws, r, span, BG);
      ws.mergeCells(r, 1, r, span);
      var c = ws.getRow(r).getCell(1);
      c.value = runs.length === 1 ? runs[0].text : { richText: runs };
      c.font = runs[0].font;
      c.alignment = { wrapText: true, vertical: 'middle' };
      var tw = 0; for (var ci = 1; ci <= span; ci++) tw += (ws.getColumn(ci).width || 10);
      var cw = runs.reduce(function(a, x) { return a + Array.from(x.text).reduce(function(b, ch) { return b + (ch.charCodeAt(0) > 255 ? 2 : 1); }, 0) * ((x.font && x.font.size) || 10) / 10; }, 0);
      ws.getRow(r).height = Math.max(h || 20, Math.ceil(cw / (tw * 0.95)) * 17 + 4); r++;
    };
    put([xlF(title + '　' + new Date().toISOString().slice(0, 10), 'FFFFFFFF', true, 14)], 26);
    if (g) {
      put([xlF('🧭 今日大盤狀態：', 'FF8EC5FF', true, 11), xlF(MARKET_REGIMES[g][0], 'FFFFD54F', true, 12), xlF('（' + MARKET_REGIMES[g][1] + '）　→　' + REGIME_ADVICE[g], g === 2 ? 'FFFF9800' : 'FF8EC5FF', false, 11)], 34);
    } else {
      put([xlF('🧭 今日大盤狀態：無大盤資料（這次抓不到指數），「大盤適用」欄不標示', 'FFFF9800', true, 11)]);
    }
    try {
      var mp = marketPanel();
      mp.lines.forEach(function(l) { put([xlF(l[0] + ' ' + l[1] + '：', 'FF8EC5FF', true, 11), xlF(l[2], 'FFFFD54F', true, 11), xlF('　' + l[3], 'FFB0B8C8', false, 10)], 22); });
      put([xlF(mp.verdict[1], 'FF' + mp.verdict[0].replace('#', '').toUpperCase(), true, 11)], 22);
      put([xlF('📋 大盤指引', 'FF8EC5FF', true, 11)], 20);
      mp.guide.forEach(function(t, i) { var bold = t.indexOf('建議持股') >= 0; put([xlF(t.replace(/<[^>]+>/g, ''), bold ? 'FFFFD54F' : 'FFCFD8E3', bold, bold ? 11 : 10)], 20); });
    } catch (e) {}
    // 大盤摘要表（四種大盤狀態）
    var hdr = ws.getRow(r);
    xlFillRow(ws, r, span, 'FF1C2236');
    [['大盤狀態', 1, 2], ['判斷方式', 3, 5], ['這種大盤下的回測建議', 6, span]].forEach(function(h) {
      ws.mergeCells(r, h[1], r, h[2]);
      var c = hdr.getCell(h[1]); c.value = h[0]; c.font = xlF('', 'FF8A8F9C', true).font; c.alignment = { vertical: 'middle' };
    });
    hdr.height = 18; r++;
    [1, 2, 3, 4].forEach(function(k) {
      var cur = k === g, row = ws.getRow(r);
      xlFillRow(ws, r, span, cur ? 'FF2A2410' : 'FF12172A');
      [[MARKET_REGIMES[k][0] + (cur ? '　◀ 今日' : ''), 1, 2], [MARKET_REGIMES[k][1], 3, 5], [REGIME_ADVICE[k], 6, span]].forEach(function(v, j) {
        ws.mergeCells(r, v[1], r, v[2]);
        var c = row.getCell(v[1]);
        c.value = v[0];
        c.font = xlF('', cur ? 'FFFFD54F' : (j === 0 ? 'FFE0E0E0' : 'FFB0B8C8'), cur || j === 0, 10).font;
        c.alignment = { wrapText: true, vertical: 'middle' };
      });
      row.height = 30; r++;
    });
    if (g) {
      var runs = [xlF('各清單在「' + MARKET_REGIMES[g][0] + '」下的三年回測：', 'FFB0B8C8', true)];
      xlCats().forEach(function(ct) {
        var n = { '✅': 0, '⚠️': 0, '❌': 0 };
        ct[3].forEach(function(c) { var t = regimeTag(c, g); if (t) n[t]++; });
        runs.push(xlF('　' + ct[0] + ct[1] + ' ', ct[2], true));
        ['✅', '⚠️', '❌'].forEach(function(t) { runs.push(xlF(t + n[t] + ' ', XL_TAG_COLORS[t])); });
      });
      put(runs, 22);
    }
    put([xlF('「大盤適用」欄：', 'FFB0B8C8', true), xlF('✅', 'FF4CAF50', true), xlF('＝這種大盤下三年回測有效（依日t≥2且超額為正）　', 'FFB0B8C8'),
         xlF('❌', 'FFFF5555', true), xlF('＝回測超額為負　', 'FFB0B8C8'), xlF('⚠️', 'FFF0A500', true), xlF('＝普通。　主表右側「技術訊號條件」欄：', 'FFB0B8C8'),
         xlF('✓', 'FF00C864', true), xlF('＝今日條件成立。　組合編號的 t 值、組合型態請看「t值與組合型態索引」工作表。', 'FFB0B8C8')], 22);
    xlFillRow(ws, r, span, BG); r++;
    return r;
  }
  // 批次分析摘要（大盤狀態／溫度計／K線／大盤指引／四種大盤摘要表）放到獨立的「大盤摘要」工作表
  function xlAddSummarySheet(wb, title) {
    var ws = wb.addWorksheet('大盤摘要', { properties: { tabColor: { argb: 'FFFFD54F' } } });
    ws.columns = [12, 12, 14, 14, 14, 22, 22, 22, 22, 22].map(function(w) { return { width: w }; });
    var r = xlWriteRegimeHeader(ws, title, 10);
    ws.views = [{ showGridLines: false }];
    return ws;
  }
  // 技術訊號條件欄的表頭（橫書、自動換行，欄寬15）與內容
  function xlFlagHeader(cell, name) {
    cell.value = name;
    cell.font = { name: XL_FONT, size: 9, bold: true, color: { argb: 'FF8EC5FF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16203A' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF2A2E3A' } }, left: { style: 'hair', color: { argb: 'FF2A2E3A' } } };
  }
  function xlFlagCell(cell, on, bg) {
    cell.value = on ? '✓' : '';
    cell.font = { name: XL_FONT, size: 10, bold: true, color: { argb: 'FF00C864' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: on ? 'FF0F2A1C' : bg } };
    cell.alignment = { horizontal: 'center', vertical: 'top' };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF2A2E3A' } }, left: { style: 'hair', color: { argb: 'FF2A2E3A' } } };
  }
  function xlRegimeCellText(combo) {
    var out = [];
    [1, 2, 3, 4].forEach(function(k) {
      var row = REGIME_STATS[comboNormKey(combo)], st = row ? row[k - 1] : null, t = regimeTag(combo, k);
      out.push(st && st[0] >= REGIME_MIN_N ? { tag: t, text: (t || '') + ' ' + (st[2] >= 0 ? '+' : '') + st[2] + '% t' + (st[3] == null ? '-' : st[3]) } : { tag: '', text: '樣本不足' });
    });
    return out;
  }
  function xlFmtHz(o, pct) {
    var ks = Object.keys(o || {}).filter(function(h) { return o[h] != null; }).sort(function(a, b) { return a - b; });
    return ks.map(function(h) { return h + '日' + o[h] + (pct ? '%' : ''); }).join(' / ');
  }
  // 「t值與組合型態索引」工作表：上半＝組合型態分級；下半＝所有組合編號的 t 值、勝率、飆股比例與四種大盤下的表現
  function xlAddIndexSheet(wb) {
    var ws = wb.addWorksheet('t值與組合型態索引'), g = batchMarketRegime(), BG = 'FF12172A', HBG = 'FF1C2236';
    var widths = [8, 11, 7, 66, 9, 9, 9, 30, 20, 17, 17, 17, 17, 10];
    ws.columns = widths.map(function(w) { return { width: w }; });
    var NC = widths.length, r = 1;
    var title = function(t) {
      xlFillRow(ws, r, NC, 'FF0A0E1A'); ws.mergeCells(r, 1, r, NC);
      var c = ws.getRow(r).getCell(1); c.value = t; c.font = xlF('', 'FFFFFFFF', true, 13).font; c.alignment = { vertical: 'middle' };
      ws.getRow(r).height = 24; r++;
    };
    var header = function(cols) {
      var row = ws.getRow(r);
      cols.forEach(function(h, i) { xlStyleCell(row.getCell(i + 1), [xlF(h, 'FF8A8F9C', true)], HBG, true); });
      for (var c = cols.length + 1; c <= NC; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HBG } };
      row.height = 20; return r++;
    };
    // A. 組合型態
    title('🏷️ 組合型態（依命中類別 S/#/M/W/F 的組合分級）');
    header(['型態', '類別組合', '', '回測說明']);
    ['進攻', '穩健', '彩券'].forEach(function(k) {
      var row = ws.getRow(r), lk = COMBO_STYLE_LOOK[k], col = 'FF' + lk[1].replace('#', '').toUpperCase();
      var keys = Object.keys(COMBO_STYLE_RULES).filter(function(c) { return COMBO_STYLE_RULES[c] === k; });
      xlStyleCell(row.getCell(1), [xlF(lk[0], col, true, 11)], BG, false);
      ws.mergeCells(r, 2, r, 3); xlStyleCell(row.getCell(2), [xlF(keys.join('、'), col, true)], BG, false);
      ws.mergeCells(r, 4, r, NC); xlStyleCell(row.getCell(4), [xlF(COMBO_STYLE_NOTE[k], 'FFB0B8C8')], BG, false);
      row.height = 34; r++;
    });
    var lr = ws.getRow(r); ws.mergeCells(r, 1, r, NC);
    var lg = [xlF('類別代號：', 'FFB0B8C8', true)];
    xlCats().forEach(function(ct) { lg.push(xlF(ct[0] + '＝' + ct[1] + '（' + ct[3].length + '組）　', ct[2], true)); });
    xlStyleCell(lr.getCell(1), lg, BG, false); lr.height = 20; r++;
    xlFillRow(ws, r, NC, 'FF0A0E1A'); r++;
    // B. 組合 t 值索引
    title('📇 組合 t 值索引（t＝扣掉同一天大盤後的超額報酬 t 值，>2 較可信；四種大盤欄＝20日超額與依日 t）');
    var hr = header(['編號', '類別', '擇時', '條件組合', 't(5日)', 't(10日)', 't(20日)', '勝率', '飆股比例', '①多頭', '②多頭回檔', '③空頭反彈', '④空頭', '今日適用']);
    xlCats().forEach(function(ct) {
      ct[3].forEach(function(c, i) {
        var st = comboStatsFor(c, ct[0]), row = ws.getRow(r), rg = xlRegimeCellText(c), today = g ? regimeTag(c, g) : '';
        var timing = isTimingCombo(c);
        var cells = [
          [xlF(ct[0] + (i + 1), ct[2], true)], [xlF(ct[1], ct[2])], [xlF(timing ? '⏱' : '', 'FFF0A500')], [xlF(c.join('＋'), 'FFE0E0E0')],
          [xlF(st.t[5] != null ? String(st.t[5]) : '', st.t[5] >= 2 ? 'FFFFFFFF' : 'FF8A8F9C', st.t[5] >= 2)],
          [xlF(st.t[10] != null ? String(st.t[10]) : '', st.t[10] >= 2 ? 'FFFFFFFF' : 'FF8A8F9C', st.t[10] >= 2)],
          [xlF(st.t[20] != null ? String(st.t[20]) : '', st.t[20] >= 2 ? 'FFFFFFFF' : 'FF8A8F9C', st.t[20] >= 2)],
          [xlF(xlFmtHz(st.win, true), 'FFB0B8C8')], [xlF(xlFmtHz(st.moon, true), 'FFE91E63')]
        ];
        rg.forEach(function(x) { cells.push([xlF(x.text, x.tag ? XL_TAG_COLORS[x.tag] : 'FF555555', !!x.tag)]); });
        cells.push([xlF(today || '－', today ? XL_TAG_COLORS[today] : 'FF555555', true, 11)]);
        cells.forEach(function(runs, j) {
          runs = runs.filter(function(x) { return x.text !== ''; });
          xlStyleCell(row.getCell(j + 1), runs, BG, false);
        });
        row.height = 18; r++;
      });
    });
    ws.autoFilter = { from: { row: hr, column: 1 }, to: { row: r - 1, column: NC } };
    ws.views = [{ state: 'frozen', ySplit: hr }];
  }
  function xlRegimeFitText(r) {
    var f = regimeFit(r);
    if (!f.reg) return [];
    var runs = [];
    ['✅', '❌', '⚠️'].forEach(function(t) {
      if (!f.b[t].length) return;
      if (runs.length) runs.push(xlF('\n', 'FFE0E0E0'));
      runs.push(xlF(t + f.b[t].length + ' ', XL_TAG_COLORS[t], true));
      runs.push(xlF(f.b[t].join(' '), XL_TAG_COLORS[t]));
    });
    return runs;
  }

  function exportTableToExcelStyled(table, filenamePrefix, sheetName, opts) {
    opts = opts || {};
    var wb = new ExcelJS.Workbook(), ws = wb.addWorksheet(sheetName || '工作表1');
    var ths = Array.prototype.slice.call(table.querySelectorAll('thead th'));
    var widths = ths.map(function(th) { return Math.min(Math.max(Math.round(th.getBoundingClientRect().width / 6.5), 8), 45); });
    var FM = opts.summary ? xlFlagMatrix() : null, nBase = widths.length;
    if (FM) FM.names.forEach(function() { widths.push(15); });
    ws.columns = widths.map(function(w) { return { width: w }; });
    if (opts.summary) xlAddSummarySheet(wb, '批次分析摘要');
    var top = 1;
    var hRow = ws.getRow(top);
    var hTexts = ths.map(function(th, i) {
      var runs = xlRichText(th);
      runs.forEach(function(r) { r.text = r.text.replace(/\n/g, ''); });
      if (runs.length) runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      runs = runs.filter(function(r) { return r.text !== ''; });
      xlStyleCell(hRow.getCell(i + 1), runs, xlEffectiveBg(th), true);
      return runs.map(function(r) { return r.text; }).join('');
    });
    if (FM) FM.names.forEach(function(n, k) { xlFlagHeader(hRow.getCell(nBase + k + 1), n); });
    if (FM) for (var hc = 1; hc <= nBase; hc++) hRow.getCell(hc).alignment = { wrapText: true, vertical: 'middle' };
    hRow.height = FM ? 50 : xlRowHeight(hTexts, widths);
    var rIdx = top + 1;
    Array.prototype.forEach.call(table.querySelectorAll('tbody tr'), function(tr) {
      if (tr.classList.contains('hidden-row')) return;
      var row = ws.getRow(rIdx++), texts = [], bg = null;
      Array.prototype.forEach.call(tr.querySelectorAll('td'), function(td, i) {
        var runs = xlRichText(td);
        bg = xlEffectiveBg(td);
        xlStyleCell(row.getCell(i + 1), runs, bg, false);
        texts.push(runs.map(function(r) { return r.text; }).join(''));
      });
      if (FM) {
        var r0 = batchResults[parseInt(tr.dataset.idx, 10)], hit = r0 ? (FM.byId[r0.stockId] || {}) : {};
        FM.names.forEach(function(n, k) { xlFlagCell(row.getCell(nBase + k + 1), !!hit[n], bg || 'FF12172A'); });
      }
      row.height = xlRowHeight(texts, widths.slice(0, nBase));
    });
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: top }];
    if (rIdx > top + 1) ws.autoFilter = { from: { row: top, column: 1 }, to: { row: rIdx - 1, column: widths.length } };   // 每個欄位都可篩選
    if (opts.summary) xlAddIndexSheet(wb);
    return xlDownload(wb, (filenamePrefix || '匯出資料') + '_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  }

  function exportTableToExcel(selector, filenamePrefix, sheetName, opts) {
    var table = typeof selector === 'string' ? document.querySelector(selector) : selector;
    if (!table) { alert('目前沒有可匯出的分析結果'); return; }
    if (window.ExcelJS) {
      try { return exportTableToExcelStyled(table, filenamePrefix, sheetName, opts); }
      catch (e) { console.warn('ExcelJS 匯出失敗，改用純文字匯出', e); }
    }

    function cellToText(el) {
      var clone = el.cloneNode(true);
      clone.querySelectorAll('.dmi-pop').forEach(function(x) { x.remove(); });
      clone.querySelectorAll('br').forEach(function(br) { br.replaceWith('\n'); });
      clone.querySelectorAll('div').forEach(function(dv) { dv.prepend('\n'); });
      return clone.textContent.replace(/\s+\n/g, '\n').replace(/\n{2,}/g, '\n').trim();
    }

    var headerCells = table.querySelectorAll('thead th');
    var headers = Array.prototype.map.call(headerCells, function(th) {
      var clone = th.cloneNode(true);
      clone.querySelectorAll('.sort-icon').forEach(function(s) { s.remove(); });
      return clone.textContent.trim();
    });

    var bodyRows = table.querySelectorAll('tbody tr');
    var aoa = [headers];
    bodyRows.forEach(function(tr) {
      if (tr.classList.contains('hidden-row')) return;
      var cells = tr.querySelectorAll('td');
      aoa.push(Array.prototype.map.call(cells, cellToText));
    });

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var colWidths = headers.map(function(h, colIdx) {
      var maxLen = h.length;
      aoa.slice(1).forEach(function(row) {
        var cellLines = String(row[colIdx] || '').split('\n');
        cellLines.forEach(function(line) { if (line.length > maxLen) maxLen = line.length; });
      });
      return { wch: Math.min(Math.max(maxLen + 2, 8), 40) };
    });
    ws['!cols'] = colWidths;

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName || '工作表1');
    var todayStr = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, (filenamePrefix || '匯出資料') + '_' + todayStr + '.xlsx');
  }

  // 舊名稱保留相容，內部直接呼叫泛用版本
  function exportSummaryToExcel() {
    exportTableToExcel('.summary-table', '批次分析摘要', '批次分析摘要', { summary: true });
  }

  function switchTab(idx) {
    // 更新 tab 按鈕狀態
    batchResults.forEach(function(_, i) {
      var tb = document.getElementById('tab'+i);
      var pn = document.getElementById('panel'+i);
      if (tb) tb.classList.toggle('active', i===idx);
      if (pn) pn.classList.toggle('active', i===idx);
    });
    activeTab = idx;

    // Lazy render：首次切換才渲染
    var panel = document.getElementById('panel'+idx);
    if (panel && !panel.dataset.rendered) {
      panel.dataset.rendered = '1';
      var r = batchResults[idx];
      var res = renderSingle(r.data, r.stockId + ' ' + r.name, r.dm, r.pb, r.total, r.pt);
      panel.innerHTML = res.html;
      var aiBtn = panel.querySelector('.ai-analysis-btn');
      if (aiBtn) aiBtn.addEventListener('click', function() { runAIAnalysis(r, panel); });
      // 等 DOM 寫入完成再畫圖
      setTimeout(function() {
        drawChart(r.data, r.stockId + ' ' + r.name, res.chartId, r.pt);
      }, 50);
    }

    // 確保被選中的 tab button 滾入視野
    var activeBtn = document.getElementById('tab'+idx);
    if (activeBtn) activeBtn.scrollIntoView({ behavior:'smooth', block:'nearest', inline:'center' });
  }


  // ────────────────────────────────────────────────────────────────
  // 命中組合：匯出與每日追蹤
  // 追蹤紀錄存在這個瀏覽器的 localStorage（換電腦/清快取會不見，請定期用「匯出紀錄」備份，
  // 另一台電腦用「匯入紀錄」還原）。同一天同一檔只記一次。
  // ────────────────────────────────────────────────────────────────
  var HIT_LOG_KEY = 'comboHitsLog_v1';
  var HIT_COLS = ['日期', '股票代號', '股票名', '股價', '漲跌%', '命中組合', '勝率', '飆股比例', 't值', '命中編號', '盤中'];
  // 盤中（即時資料）加入追蹤的命中：2026-10-19 前盤中成交量沒換算成全天量，容易誤判「地量」；
  // 舊紀錄沒有「盤中」欄位，就把含「地量」條件的命中視為疑似盤中假訊號
  function isIntradayHit(r) {
    if (r['盤中'] === '是') return true;
    if (r['盤中'] == null) return /地量/.test(String(r['命中組合'] || ''));
    return false;
  }

  function buildHitRecords() {
    var out = [];
    batchResults.forEach(function(r) {
      var mp = [], mm = [];
      try { mp = matchedPinnedCombos(r); } catch (e) {}
      try { mm = matchedMoonshotCombos(r); } catch (e) {}
      var ms = [];
      try { ms = matchedStockPickCombos(r); } catch (e) {}
      var mwh = [], mfh = [];
      try { mwh = matchedBtWinCombos(r); } catch (e) {}
      try { mfh = matchedBtHotCombos(r); } catch (e) {}
      if (!mp.length && !mm.length && !ms.length && !mwh.length && !mfh.length) return;
      var last = r.data[r.data.length - 1], prev = r.data[r.data.length - 2] || last;
      if (!(last.close > 0)) return;   // 0價（無成交／暫停交易）不列入命中
      var lines = [], codes = [], wins = [], moons = [], ts = [];
      var add = function(code, tag, c, kind) {
        var st = comboStatsFor(c, kind), bt = bestT(st);
        lines.push(code + tag + ' ' + c.join('＋') + (bt != null ? '（t=' + bt + '）' : ''));
        codes.push(code);
        wins.push(fmtHz(st.win, true));
        moons.push(fmtHz(st.moon, true));
        ts.push(fmtHz(st.t, false));
      };
      ms.forEach(function(i) { add('S' + (i+1), '[選股型]', STOCK_PICK_COMBOS[i], 'S'); });
      mp.forEach(function(i) { var c = PINNED_COMBOS[i]; add('#' + (i+1), isTimingCombo(c) ? '[高勝率·擇時型]' : '[高勝率]', c, '#'); });
      mm.forEach(function(i) { add('M' + (i+1), '[高標股]', MOONSHOT_COMBOS[i], 'M'); });
      // 回測自動入選：每類只取最強的5組（避免一檔命中幾十組時紀錄過長）
      var mx = function(st) { var v = Object.keys(st.moon || {}).map(function(h){ return st.moon[h]; }).filter(function(x){ return x != null; }); return v.length ? Math.max.apply(null, v) : 0; };
      mwh.sort(function(a, b){ return (bestT(comboStatsFor(BT_WIN_COMBOS[b], 'W')) || 0) - (bestT(comboStatsFor(BT_WIN_COMBOS[a], 'W')) || 0); })
        .slice(0, 5).forEach(function(i) { var c = BT_WIN_COMBOS[i]; add('W' + (i+1), isTimingCombo(c) ? '[回測高勝率·擇時型]' : '[回測高勝率]', c, 'W'); });
      mfh.sort(function(a, b){ return mx(comboStatsFor(BT_HOT_COMBOS[b], 'F')) - mx(comboStatsFor(BT_HOT_COMBOS[a], 'F')); })
        .slice(0, 5).forEach(function(i) { add('F' + (i+1), '[回測飆股]', BT_HOT_COMBOS[i], 'F'); });
      out.push({ '日期': last.date, '股票代號': r.stockId, '股票名': r.name, '股價': +last.close.toFixed(2),
        '漲跌%': prev.close ? +((last.close - prev.close) / prev.close * 100).toFixed(2) : null,
        '命中組合': lines.join('\n'), '勝率': wins.join('\n'), '飆股比例': moons.join('\n'), 't值': ts.join('\n'),
        '命中編號': codes.join(' '), '盤中': r.realtime ? '是' : '' });
    });
    return out;
  }

  // 命中組合匯出的顏色跟網頁徽章一致：S 綠、# 紫、M 橘、W 青、F 粉紅；漲跌%：漲紅跌綠
  var HIT_CODE_COLORS = { 'S': 'FF00C864', '#': 'FF7C3AED', 'M': 'FFFF9800', 'W': 'FF00BCD4', 'F': 'FFE91E63' };
  function exportHitRecordsStyled(recs, filenamePrefix, sheetName) {
    var wb = new ExcelJS.Workbook(), ws = wb.addWorksheet(sheetName);
    var SUMCOLS = ['命中類別', '組合型態', '大盤適用', '選股型命中', '指定組合命中'];
    var COLS = HIT_COLS.slice(0, 5).concat(SUMCOLS).concat(HIT_COLS.slice(5));
    var widths = [12, 10, 16, 10, 9, 9, 10, 24, 16, 24, 80, 30, 26, 30, 22, 6], nBase = widths.length, FM = xlFlagMatrix();
    // 跟網頁摘要表同一套畫法：把「命中類別／組合型態／選股型命中／指定組合命中」畫在隱藏表格裡，再照顏色轉成 Excel
    var tmpTbl = document.createElement('table');
    tmpTbl.className = 'summary-table';
    tmpTbl.style.cssText = 'position:absolute;left:-99999px;top:0;width:1600px';
    document.body.appendChild(tmpTbl);
    var sumRuns = function(r) {
      if (!r) return {};
      var tr = document.createElement('tr');
      tr.innerHTML = renderHitCategories(r) + renderComboStyle(r) + '<td>' + renderMatchedComboBadges(r, 'S') + '</td><td>' + renderMatchedComboBadges(r, 'PM') + '</td>';
      tmpTbl.appendChild(tr);
      var tds = tr.querySelectorAll('td'), o = {};
      ['命中類別', '組合型態', '選股型命中', '指定組合命中'].forEach(function(k, i) { o[k] = tds[i] ? xlRichText(tds[i]) : []; });
      tr.remove();
      return o;
    };
    FM.names.forEach(function() { widths.push(15); });
    ws.columns = widths.map(function(w) { return { width: w }; });
    xlAddSummarySheet(wb, '批次分析摘要');
    var top = 1;
    var byId = {};
    batchResults.forEach(function(r) { byId[r.stockId] = r; });
    var BG = 'FF12172A', HBG = 'FF1C2236', base = { name: XL_FONT, size: 10, color: { argb: 'FFE0E0E0' } };
    var hRow = ws.getRow(top);
    COLS.forEach(function(h, i) {
      xlStyleCell(hRow.getCell(i + 1), [{ text: h, font: { name: XL_FONT, size: 10, bold: true, color: { argb: 'FF8A8F9C' } } }], HBG, true);
    });
    FM.names.forEach(function(n, k) { xlFlagHeader(hRow.getCell(nBase + k + 1), n); });
    for (var hc = 1; hc <= nBase; hc++) hRow.getCell(hc).alignment = { wrapText: true, vertical: 'middle' };
    hRow.height = 50;
    recs.forEach(function(rec, ri) {
      var row = ws.getRow(top + 1 + ri), codes = String(rec['命中編號'] || '').split(' ');
      var lineColor = function(line, k) {
        var c = (String(line).match(/^[S#MWF]/) || [codes[k] ? codes[k][0] : ''])[0];
        return HIT_CODE_COLORS[c] || 'FFE0E0E0';
      };
      var texts = [], SR = null, cellCenter = [];
      COLS.forEach(function(col, i) {
        var v = rec[col], runs;
        if (col === '大盤適用') {
          runs = byId[rec['股票代號']] ? xlRegimeFitText(byId[rec['股票代號']]) : [];
        } else if (SUMCOLS.indexOf(col) >= 0) {
          if (!SR) SR = sumRuns(byId[rec['股票代號']]);
          runs = (SR[col] || []).slice();
          if (col === '命中類別' || col === '組合型態') cellCenter.push(i + 1);
        } else if (col === '命中組合' || col === '勝率' || col === '飆股比例' || col === 't值') {
          var lines = String(v == null ? '' : v).split('\n');
          runs = [];
          lines.forEach(function(line, k) {
            if (k) runs.push({ text: '\n', font: base });
            runs.push({ text: line, font: { name: XL_FONT, size: 10, bold: col === '命中組合', color: { argb: lineColor(col === '命中組合' ? line : '', k) } } });
          });
        } else if (col === '命中編號') {
          runs = [];
          codes.forEach(function(cd, k) {
            if (k) runs.push({ text: ' ', font: base });
            runs.push({ text: cd, font: { name: XL_FONT, size: 10, bold: true, color: { argb: HIT_CODE_COLORS[cd[0]] || 'FFE0E0E0' } } });
          });
        } else if (col === '漲跌%') {
          var num = +v;
          runs = [{ text: v == null ? '' : ((num > 0 ? '+' : '') + num.toFixed(2) + '%'), font: { name: XL_FONT, size: 10, color: { argb: num >= 0 ? 'FFEF5350' : 'FF26A69A' } } }];
        } else if (col === '股票代號') {
          runs = [{ text: String(v == null ? '' : v), font: { name: XL_FONT, size: 10, bold: true, color: { argb: 'FFFFFFFF' } } }];
        } else if (col === '股價') {
          runs = [{ text: v == null ? '' : (+v).toFixed(2), font: { name: XL_FONT, size: 10, bold: true, color: { argb: 'FFFFFFFF' } } }];
        } else {
          runs = [{ text: String(v == null ? '' : v), font: base }];
        }
        runs = runs.filter(function(r) { return r.text !== ''; });
        xlStyleCell(row.getCell(i + 1), runs, BG, false);
        texts.push(runs.map(function(r) { return r.text; }).join(''));
      });
      cellCenter.forEach(function(ci) { row.getCell(ci).alignment = { wrapText: true, vertical: 'top', horizontal: 'center' }; });
      var hit = FM.byId[rec['股票代號']] || {};
      FM.names.forEach(function(n, k) { xlFlagCell(row.getCell(nBase + k + 1), !!hit[n], BG); });
      row.height = xlRowHeight(texts, widths.slice(0, nBase));
    });
    tmpTbl.remove();
    ws.views = [{ state: 'frozen', xSplit: 3, ySplit: top }];
    if (recs.length) ws.autoFilter = { from: { row: top, column: 1 }, to: { row: top + recs.length, column: widths.length } };   // 每個欄位都可篩選
    xlAddIndexSheet(wb);
    return xlDownload(wb, filenamePrefix + '_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  }

  function exportHitRecords(recs, filenamePrefix, sheetName) {
    if (window.ExcelJS) {
      try { return exportHitRecordsStyled(recs, filenamePrefix, sheetName); }
      catch (e) { console.warn('ExcelJS 匯出失敗，改用純文字匯出', e); }
    }
    var ws = XLSX.utils.json_to_sheet(recs, { header: HIT_COLS });
    ws['!cols'] = [{wch:12},{wch:10},{wch:12},{wch:10},{wch:9},{wch:80},{wch:26},{wch:26},{wch:26},{wch:22}];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, filenamePrefix + '_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  }

  function getHitLog() {
    // 股價≤0 的紀錄是資料源當天回傳0價（無成交／暫停交易）造成的假命中，直接排除
    try { var a = JSON.parse(localStorage.getItem(HIT_LOG_KEY) || '[]'); return Array.isArray(a) ? a.filter(function(r) { return !(+r['股價'] <= 0); }) : []; }
    catch (e) { return []; }
  }
  function setHitLog(arr) {
    try { localStorage.setItem(HIT_LOG_KEY, JSON.stringify(arr)); } catch (e) { alert('追蹤紀錄儲存失敗（瀏覽器儲存空間不足？）'); }
  }
  function appendHitLog(recs) {
    if (!recs || !recs.length) return 0;
    var log = getHitLog(), idx = {}, added = 0;
    log.forEach(function(r, i) { idx[r['日期'] + '|' + r['股票代號']] = i; });
    recs.forEach(function(r) {
      var k = r['日期'] + '|' + r['股票代號'];
      if (idx[k] != null) log[idx[k]] = r; else { idx[k] = log.length; log.push(r); added++; }
    });
    log.sort(function(a, b) { return (a['日期'] + a['股票代號']) < (b['日期'] + b['股票代號']) ? -1 : 1; });
    setHitLog(log);
    return added;
  }

  var trackPerf = null; // 最近一次「更新追蹤報酬」的結果
  var trackMeta = null;  // { at: 更新時間, fail: [抓不到股價的代號] }
  // 上次更新的報酬結果也存在瀏覽器裡，重新開啟頁面不用再按一次「更新追蹤報酬」
  var HIT_PERF_KEY = HIT_LOG_KEY + '_perf';
  function saveTrackPerf() {
    try {
      if (trackPerf) localStorage.setItem(HIT_PERF_KEY, JSON.stringify({ perf: trackPerf, meta: trackMeta }));
      else localStorage.removeItem(HIT_PERF_KEY);
    } catch (e) { /* 空間不足就只保留在記憶體 */ }
  }
  (function loadTrackPerf() {
    try {
      var o = JSON.parse(localStorage.getItem(HIT_PERF_KEY) || 'null');
      if (o && Array.isArray(o.perf)) { trackPerf = o.perf; trackMeta = o.meta || null; }
    } catch (e) {}
  })();

  async function computeTrackPerf(token, onProgress) {
    var log = getHitLog();
    if (!log.length) return [];
    var first = log.reduce(function(m, r){ return r['日期'] < m ? r['日期'] : m; }, log[0]['日期']);
    var days = Math.ceil((Date.now() - new Date(first + 'T00:00:00').getTime()) / 86400000) + 10;
    var sids = Object.keys(log.reduce(function(o, r){ o[r['股票代號']] = 1; return o; }, {})).sort();
    var px = {}, failSids = [];
    for (var i = 0; i < sids.length; i++) {
      try {
        var rows = await fetchPriceData(sids[i], token, days);
        px[sids[i]] = rows.map(function(d){ return { date: d.date, high: +d.max, low: +d.min, close: +d.close }; })
          .sort(function(a,b){ return a.date < b.date ? -1 : 1; });
      } catch (e) { px[sids[i]] = []; }
      if (!px[sids[i]].length) failSids.push(sids[i]);
      if (onProgress) onProgress(i + 1, sids.length);
      if (i < sids.length - 1) await new Promise(function(r){ setTimeout(r, 300); });
    }
    trackMeta = { at: new Date().toLocaleString(), fail: failSids };
    return log.map(function(r) {
      var rows = px[r['股票代號']] || [];
      var rec = { '日期': r['日期'], '股票代號': r['股票代號'], '股票名': r['股票名'], '股價': r['股價'], '命中編號': r['命中編號'], '命中組合': r['命中組合'] , '盤中': r['盤中'] == null ? null : r['盤中'] };
      var k = -1;
      for (var j = rows.length - 1; j >= 0; j--) { if (rows[j].date <= r['日期']) { k = j; break; } }
      if (k < 0) { rec['持有天數'] = null; rec['最新價'] = null; rec['目前報酬%'] = null; return rec; }
      var ent = rows[k].close, after = rows.slice(k + 1), lastRow = rows[rows.length - 1];
      var pct = function(v){ return +((v - ent) / ent * 100).toFixed(2); };
      rec['持有天數'] = after.length;
      rec['最新價'] = lastRow.close;
      rec['目前報酬%'] = pct(lastRow.close);
      [5, 10, 20].forEach(function(h){ rec[h + '日報酬%'] = after.length >= h ? pct(after[h-1].close) : null; });
      rec['最大漲幅%'] = after.length ? pct(Math.max.apply(null, after.map(function(x){ return x.high; }))) : null;
      rec['最大回檔%'] = after.length ? pct(Math.min.apply(null, after.map(function(x){ return x.low; }))) : null;
      return rec;
    });
  }

  function comboTrackSummary(perf) {
    var groups = {};
    // 從追蹤紀錄當時寫下的「命中組合」文字取出各編號的組合（清單改版後舊紀錄仍對得上）
    var loggedCombo = function(code, txt) {
      var lines = String(txt || '').split('\n');
      for (var i = 0; i < lines.length; i++) {
        var ln = lines[i].trim();
        if (ln.indexOf(code + '[') === 0) return ln.replace(/^[^\]]*\]\s*/, '').replace(/（t=.*）$/, '').trim();
      }
      return null;
    };
    perf.forEach(function(r) {
      String(r['命中編號'] || '').split(/\s+/).filter(Boolean).forEach(function(code) {
        var key = code + '|' + (loggedCombo(code, r['命中組合']) || '');
        (groups[key] = groups[key] || []).push(r);
      });
    });
    var avg = function(a){ return a.length ? +(a.reduce(function(s,v){return s+v;},0) / a.length).toFixed(2) : null; };
    var rows = Object.keys(groups).map(function(gk) {
      var code = gk.split('|')[0], logged = gk.slice(code.length + 1);
      var g = groups[gk], idx = parseInt(code.slice(1), 10) - 1, combo, rec;
      if (code[0] === 'S') { combo = STOCK_PICK_COMBOS[idx] || []; rec = formatStockPickStats(STOCK_PICK_STATS[combo.join(' ＋ ')]); }
      else if (code[0] === '#') { combo = PINNED_COMBOS[idx] || []; rec = (isTimingCombo(combo) ? '[擇時型] ' : '') + formatWinrates(PINNED_COMBO_WINRATES[combo.join(' ＋ ')]); }
      else if (code[0] === 'W') { combo = BT_WIN_COMBOS[idx] || []; rec = btWinText(BT_WIN_STATS[combo.join(' ＋ ')]); }
      else if (code[0] === 'F') { combo = BT_HOT_COMBOS[idx] || []; rec = formatMoonshotStats(BT_HOT_STATS[combo.join(' ＋ ')]); }
      else { combo = MOONSHOT_COMBOS[idx] || []; rec = formatMoonshotStats(MOONSHOT_COMBO_STATS[combo.join(' ＋ ')]); }
      var shownCode = code, comboTxt = combo.join('＋');
      if (logged && logged !== comboTxt) { shownCode = code + '(舊版)'; comboTxt = logged; rec = '（舊版清單組合）'; }
      var row = { '編號': shownCode, '條件組合': comboTxt, '命中次數': g.length,
        '目前平均報酬%': avg(g.map(function(r){ return r['目前報酬%']; }).filter(function(v){ return v != null; })) };
      [5, 10, 20].forEach(function(h) {
        var v = g.map(function(r){ return r[h + '日報酬%']; }).filter(function(x){ return x != null; });
        row[h + '日已滿筆數'] = v.length;
        row[h + '日平均%'] = avg(v);
        row[h + '日勝率%'] = v.length ? +(v.filter(function(x){ return x > 0; }).length / v.length * 100).toFixed(1) : null;
      });
      row['回測記錄'] = rec;
      return row;
    });
    rows.sort(function(a, b){ return b['命中次數'] - a['命中次數']; });
    return rows;
  }

  var trackOutput = document.getElementById('trackOutput');
  var trackDayFilter = '全部';
  var trackExclIntraday = true;   // 組合彙總排除盤中／疑似盤中假訊號命中

  function renderTracking() {
    var log = getHitLog();
    trackOutput.style.display = 'block';
    var smallBtn = 'padding:6px 12px;font-size:12px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:6px;color:#ddd;cursor:pointer';
    var html = '<div class="sec" style="font-size:18px">📈 命中組合每日追蹤</div>' +
      '<div style="color:#888;font-size:12px;margin-bottom:10px">每次批次分析會自動把命中組合的股票累加進來（側邊欄可關閉），同一天同一檔只記一次。進場價＝命中當天收盤價，' +
      '「5/10/20日報酬」要等滿天數才會出現。紀錄存在這個瀏覽器裡，請定期「匯出紀錄」備份。</div>';
    if (!log.length) {
      trackOutput.innerHTML = html + '<div style="color:#888">目前沒有追蹤紀錄。跑一次批次分析（建議用「漲幅前100／成交量前100」），命中組合的股票會自動加入。</div>' +
        '<div style="margin-top:10px"><label style="' + smallBtn + '">📂 匯入紀錄<input type="file" id="iImportTrack" accept=".json" style="display:none"></label></div>';
      bindTrackButtons();
      return;
    }
    var days = Object.keys(log.reduce(function(o, r){ o[r['日期']] = 1; return o; }, {})).sort().reverse();
    html += '<div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px">' +
      '<span style="color:#ccc;font-size:13px">共 ' + log.length + ' 筆，' + Object.keys(log.reduce(function(o, r){ o[r['股票代號']] = 1; return o; }, {})).length +
      ' 檔（' + days[days.length-1] + ' ～ ' + days[0] + '）</span>' +
      '<button type="button" id="btnUpdTrack" style="' + smallBtn + ';color:#ff9800;border-color:#ff980066">🔄 更新追蹤報酬</button>' +
      '<label style="font-size:12px;color:#aaa">檢視日期：<select id="selTrackDay" style="padding:3px 6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);border-radius:4px;color:#fff;font-size:12px">' +
        ['全部'].concat(days).map(function(d){ return '<option' + (d === trackDayFilter ? ' selected' : '') + '>' + d + '</option>'; }).join('') + '</select></label>' +
      '<button type="button" id="btnExpTrack" style="' + smallBtn + '">📥 匯出追蹤報表</button>' +
      '<button type="button" id="btnBackupTrack" style="' + smallBtn + '">💾 匯出紀錄(備份)</button>' +
      '<label style="' + smallBtn + '">📂 匯入紀錄<input type="file" id="iImportTrack" accept=".json" style="display:none"></label>' +
      '<button type="button" id="btnPruneTrack" style="' + smallBtn + '">🧹 刪除N天前紀錄</button>' +
      '<span id="trackStatus" style="font-size:12px;color:#888"></span></div>';
    var flt = function(arr){ return trackDayFilter === '全部' ? arr : arr.filter(function(r){ return r['日期'] === trackDayFilter; }); };
    if (!trackPerf) {
      html += '<div style="color:#ff9800;font-size:13px;margin-bottom:6px">⚠️ 尚未計算報酬，下表只是原始命中紀錄。請按上方「🔄 更新追蹤報酬」抓最新股價，才會算出 5/10/20 日報酬。</div>' +
        objArrayToTable(flt(log).map(function(r){ return { '日期': r['日期'], '股票代號': r['股票代號'], '股票名': r['股票名'], '股價': r['股價'], '漲跌%': r['漲跌%'], '命中編號': r['命中編號'] }; }));
    } else {
      var perfKeys = {}; trackPerf.forEach(function(r) { perfKeys[r['日期'] + '|' + r['股票代號']] = 1; });
      var newN = log.filter(function(r) { return !perfKeys[r['日期'] + '|' + r['股票代號']]; }).length;
      var failN = trackMeta && trackMeta.fail ? trackMeta.fail.length : 0;
      html += '<div style="font-size:12px;color:#888;margin-bottom:4px">報酬計算時間：' + (trackMeta && trackMeta.at ? trackMeta.at : '--') +
        (newN ? '　<span style="color:#ff9800">｜之後新增 ' + newN + ' 筆尚未計算，請再按一次「更新追蹤報酬」</span>' : '') + '</div>' +
        (failN ? '<div style="font-size:12px;color:#ff8888;margin-bottom:4px">⚠️ ' + failN + ' 檔抓不到股價（多半是 API 每小時額度用完），這些的報酬顯示「--」，請過一陣子再按「更新追蹤報酬」：' +
          trackMeta.fail.slice(0, 30).join('、') + (failN > 30 ? '…' : '') + '</div>' : '');
      html += '<div class="sec">🧾 追蹤明細</div>' + objArrayToTable(flt(trackPerf), 'trackPerfTable') +
        '<div class="sec">🧩 組合實盤表現（命中後實際報酬 vs 回測記錄）</div>' +
        '<div style="color:#888;font-size:12px;margin-bottom:6px">實盤勝率明顯低於回測記錄的組合，代表可能是過度適配，可考慮移除；樣本少於10筆前先別下結論。</div>' +
        '<label style="font-size:12px;color:#aaa;cursor:pointer;display:inline-block;margin-bottom:6px" title="2026-10-19 以前，盤中加入追蹤時成交量沒換算成全天量，很多股票被誤判成「地量」；排除後的組合表現較接近實際">' +
          '<input type="checkbox" id="cbTrackExIntra"' + (trackExclIntraday ? ' checked' : '') + ' style="width:auto;margin:0 4px 0 0;vertical-align:middle">' +
          '組合彙總排除盤中命中（舊紀錄：含「地量」的命中視為疑似盤中假訊號）' +
          (trackExclIntraday ? '　<span style="color:#888">已排除 ' + trackPerf.filter(isIntradayHit).length + ' 筆（追蹤明細仍全部顯示）</span>' : '') + '</label>' +
        objArrayToTable(comboTrackSummary(trackExclIntraday ? trackPerf.filter(function(r) { return !isIntradayHit(r); }) : trackPerf), 'trackSummaryTable');
    }
    trackOutput.innerHTML = html;
    bindTrackButtons();
  }

  // ── 命中追蹤匯出：命中紀錄右邊接上追蹤報酬（持有天數、最新價、目前／5／10／20日報酬、最大漲幅、最大回檔），深色彩色＋篩選 ──
  var TRACK_PERF_COLS = ['持有天數', '最新價', '目前報酬%', '5日報酬%', '10日報酬%', '20日報酬%', '最大漲幅%', '最大回檔%'];
  function trackLogWithPerf() {
    var log = getHitLog(), idx = {};
    (trackPerf || []).forEach(function(p) { idx[p['日期'] + '|' + p['股票代號']] = p; });
    return log.map(function(r) {
      var o = Object.assign({}, r), p = idx[r['日期'] + '|' + r['股票代號']] || {};
      TRACK_PERF_COLS.forEach(function(c) { o[c] = p[c] == null ? null : p[c]; });
      return o;
    });
  }
  function xlObjectSheet(wb, name, rows, cols, widths) {
    var ws = wb.addWorksheet(name), BG = 'FF12172A', HBG = 'FF1C2236';
    ws.columns = cols.map(function(c, i) { return { width: (widths && widths[i]) || Math.min(Math.max(c.length * 2 + 2, 10), 18) }; });
    var hRow = ws.getRow(1);
    cols.forEach(function(c, i) { xlStyleCell(hRow.getCell(i + 1), [xlF(c, 'FF8A8F9C', true)], HBG, true); });
    hRow.height = 30;
    rows.forEach(function(r, ri) {
      var row = ws.getRow(ri + 2), texts = [];
      cols.forEach(function(c, i) {
        var v = r[c], cell = row.getCell(i + 1), col = 'FFE0E0E0', bold = false;
        if (v == null || v === '') v = '';
        var isRet = /報酬|漲幅|回檔|漲跌/.test(c) && typeof v === 'number';
        if (isRet) col = v > 0 ? 'FFEF5350' : (v < 0 ? 'FF26A69A' : 'FFE0E0E0');
        if (c === '股票代號' || c === '最新價' || c === '股價') { col = 'FFFFFFFF'; bold = true; }
        if (c === '命中編號') {
          var runs = [];
          String(v).split(/\s+/).filter(Boolean).forEach(function(cd, k) {
            if (k) runs.push(xlF(' ', 'FFE0E0E0'));
            runs.push(xlF(cd, HIT_CODE_COLORS[cd[0]] || 'FFE0E0E0', true));
          });
          xlStyleCell(cell, runs, BG, false);
        } else if (c === '命中組合' || c === '勝率' || c === '飆股比例' || c === 't值') {
          var lines = String(v).split('\n'), codes = String(r['命中編號'] || '').split(/\s+/), rr = [];
          lines.forEach(function(ln, k) {
            if (k) rr.push(xlF('\n', 'FFE0E0E0'));
            var cc = (ln.match(/^[S#MWF]/) || [codes[k] ? codes[k][0] : ''])[0];
            rr.push(xlF(ln, HIT_CODE_COLORS[cc] || 'FFE0E0E0', c === '命中組合'));
          });
          xlStyleCell(cell, rr.filter(function(x) { return x.text !== ''; }), BG, false);
        } else {
          xlStyleCell(cell, [xlF(typeof v === 'number' ? String(v) : String(v), col, bold)], BG, false);
          if (typeof v === 'number') { cell.value = v; cell.alignment = { vertical: 'top', horizontal: 'right' }; if (isRet) cell.numFmt = '+0.00;-0.00;0.00'; }
        }
        texts.push(String(v));
      });
      row.height = xlRowHeight(texts, cols.map(function(c, i) { return (widths && widths[i]) || 12; }));
    });
    ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 1 }];
    if (rows.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: rows.length + 1, column: cols.length } };
    return ws;
  }
  function exportTrackingExcel() {
    var fname = '命中追蹤_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    var log = trackLogWithPerf(), cols = HIT_COLS.concat(TRACK_PERF_COLS);
    if (window.ExcelJS) {
      try {
        var wb = new ExcelJS.Workbook();
        xlObjectSheet(wb, '命中紀錄', log, cols, [12, 10, 16, 10, 9, 70, 28, 22, 28, 22, 9, 10, 11, 10, 10, 10, 11, 11]);
        if (trackPerf) {
          var pc = Object.keys(trackPerf.reduce(function(o, r) { Object.keys(r).forEach(function(k) { o[k] = 1; }); return o; }, {}));
          xlObjectSheet(wb, '追蹤明細', trackPerf, pc, pc.map(function(c) { return c === '命中組合' ? 70 : (c === '命中編號' ? 22 : (c === '股票名' ? 16 : 11)); }));
          var sm = comboTrackSummary(trackExclIntraday ? trackPerf.filter(function(r) { return !isIntradayHit(r); }) : trackPerf), scs = Object.keys(sm.reduce(function(o, r) { Object.keys(r).forEach(function(k) { o[k] = 1; }); return o; }, {}));
          xlObjectSheet(wb, '組合彙總', sm, scs, scs.map(function(c) { return /組合/.test(c) ? 60 : 12; }));
        }
        xlAddIndexSheet(wb);
        return xlDownload(wb, fname);
      } catch (e) { console.warn('ExcelJS 匯出失敗，改用純文字匯出', e); }
    }
    var wb2 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.json_to_sheet(log, { header: cols }), '命中紀錄');
    if (trackPerf) {
      XLSX.utils.book_append_sheet(wb2, XLSX.utils.json_to_sheet(trackPerf), '追蹤明細');
      XLSX.utils.book_append_sheet(wb2, XLSX.utils.json_to_sheet(comboTrackSummary(trackExclIntraday ? trackPerf.filter(function(r) { return !isIntradayHit(r); }) : trackPerf)), '組合彙總');
    }
    XLSX.writeFile(wb2, fname);
  }

  function bindTrackButtons() {
    var el;
    if ((el = document.getElementById('btnUpdTrack'))) el.addEventListener('click', async function() {
      var token = iToken.value.trim();
      if (!token) { alert('請先輸入 FinMind API Token'); return; }
      var st = document.getElementById('trackStatus');
      this.disabled = true;
      try {
        trackPerf = await computeTrackPerf(token, function(i, n){ st.textContent = '抓取最新股價：' + i + ' / ' + n; });
        saveTrackPerf();
      } catch (e) { alert('更新失敗：' + e.message); }
      renderTracking();
    });
    if ((el = document.getElementById('cbTrackExIntra'))) el.addEventListener('change', function() { trackExclIntraday = this.checked; renderTracking(); });
    if ((el = document.getElementById('selTrackDay'))) el.addEventListener('change', function() { trackDayFilter = this.value; renderTracking(); });
    if ((el = document.getElementById('btnExpTrack'))) el.addEventListener('click', function() { exportTrackingExcel(); });
    if ((el = document.getElementById('btnBackupTrack'))) el.addEventListener('click', function() {
      var blob = new Blob([JSON.stringify(getHitLog())], { type: 'application/json' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = '命中追蹤紀錄備份_' + new Date().toISOString().slice(0, 10) + '.json';
      a.click();
    });
    if ((el = document.getElementById('iImportTrack'))) el.addEventListener('change', function() {
      var f = this.files && this.files[0];
      if (!f) return;
      var rd = new FileReader();
      rd.onload = function() {
        try {
          var arr = JSON.parse(rd.result);
          var n = appendHitLog(Array.isArray(arr) ? arr : []);
          alert('已匯入，新增 ' + n + ' 筆（重複的日期＋代號會以匯入檔為準）');
          trackPerf = null; trackMeta = null; saveTrackPerf(); renderTracking();
        } catch (e) { alert('匯入失敗：檔案格式不正確'); }
      };
      rd.readAsText(f);
    });
    if ((el = document.getElementById('btnPruneTrack'))) el.addEventListener('click', function() {
      var n = parseInt(prompt('只保留最近幾天的紀錄？', '120'), 10);
      if (!n || n < 1) return;
      var cut = new Date(Date.now() - n * 86400000).toISOString().slice(0, 10);
      var log = getHitLog(), keep = log.filter(function(r){ return r['日期'] >= cut; });
      if (!confirm('將刪除 ' + (log.length - keep.length) + ' 筆 ' + cut + ' 之前的紀錄，確定嗎？')) return;
      setHitLog(keep); trackPerf = null; trackMeta = null; saveTrackPerf(); renderTracking();
    });
  }

  document.getElementById('btnTrack').addEventListener('click', function() {
    renderTracking();
    trackOutput.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  // ── Indicators ──
  function ma(data, p) {
    return data.map(function(_, i) {
      if (i < p - 1) return null;
      var s = 0;
      for (var j = i - p + 1; j <= i; j++) s += data[j].close;
      return s / p;
    });
  }

  function ema(data, p) {
    var k = 2 / (p + 1), e = data[0].close, res = [];
    for (var i = 0; i < data.length; i++) {
      e = i === 0 ? data[i].close : data[i].close * k + e * (1 - k);
      res.push(e);
    }
    return res;
  }

  function macd(data) {
    var e12 = ema(data, 12), e26 = ema(data, 26);
    var dif = e12.map(function(v, i) { return v - e26[i]; });
    var k = 2 / 10, sig = [], s = dif[0];
    for (var i = 0; i < dif.length; i++) {
      s = i === 0 ? dif[i] : dif[i] * k + s * (1 - k);
      sig.push(s);
    }
    return { dif: dif, sig: sig, hist: dif.map(function(v, i) { return v - sig[i]; }) };
  }

  function rsi(data, p) {
    p = p || 14;
    var closes = data.map(function(d) { return d.close; });
    var res = new Array(closes.length).fill(null);
    var ag = 0, al = 0;
    for (var i = 1; i <= p; i++) {
      var d = closes[i] - closes[i - 1];
      if (d > 0) ag += d; else al += Math.abs(d);
    }
    ag /= p; al /= p;
    res[p] = 100 - 100 / (1 + ag / (al || 0.001));
    for (var i = p + 1; i < closes.length; i++) {
      var d = closes[i] - closes[i - 1];
      ag = (ag * (p - 1) + (d > 0 ? d : 0)) / p;
      al = (al * (p - 1) + (d < 0 ? Math.abs(d) : 0)) / p;
      res[i] = 100 - 100 / (1 + ag / (al || 0.001));
    }
    return res;
  }

  function bb(data, p, std) {
    p = p || 20; std = std || 2;
    var mid = ma(data, p);
    return data.map(function(_, i) {
      if (i < p - 1) return { u: null, l: null };
      var s = 0, m = mid[i];
      for (var j = i - p + 1; j <= i; j++) s += (data[j].close - m) * (data[j].close - m);
      var sigma = Math.sqrt(s / p);
      return { u: m + std * sigma, l: m - std * sigma };
    });
  }

  function volma(data, p) {
    return data.map(function(_, i) {
      if (i < p - 1) return null;
      var s = 0;
      for (var j = i - p + 1; j <= i; j++) s += data[j].volume;
      return s / p;
    });
  }

  function pivots(data, w) {
    w = w || 5;
    var highs = [], lows = [];
    for (var i = w; i < data.length - w; i++) {
      var h = data[i].high, l = data[i].low, isH = true, isL = true;
      for (var j = i - w; j <= i + w; j++) {
        if (data[j].high > h) isH = false;
        if (data[j].low < l) isL = false;
      }
      if (isH) highs.push(i);
      if (isL) lows.push(i);
    }
    return { highs: highs, lows: lows };
  }

  // KD (Stochastic) 計算
  function calcKD(data, period) {
    period = period || 9;
    var kArr = [], dArr = [];
    var prevK = 50, prevD = 50;
    for (var i = 0; i < data.length; i++) {
      if (i < period - 1) { kArr.push(null); dArr.push(null); continue; }
      var slice = data.slice(i - period + 1, i + 1);
      var lowest  = slice.reduce(function(m,d){ return Math.min(m, d.low);  },  Infinity);
      var highest = slice.reduce(function(m,d){ return Math.max(m, d.high); }, -Infinity);
      var rsv = highest === lowest ? 50 : (data[i].close - lowest) / (highest - lowest) * 100;
      var k = prevK * 2/3 + rsv * 1/3;
      var d = prevD * 2/3 + k   * 1/3;
      kArr.push(k); dArr.push(d);
      prevK = k; prevD = d;
    }
    return { k: kArr, d: dArr };
  }

  // ── KDJ(6,3,3)：跟上面的calcKD（9,3,3，給K線圖用）是兩套獨立參數，不要共用──
  // RSV回看6天（不是9天），K/D平滑因子都是3。J = 3K - 2D，J線放大K/D的乖離，
  // 常用來看超買超賣，這裡主要拿J線的存在來配合K/D抓黃金交叉／死亡交叉。
  function calcKDJ(data, period, kN, dN) {
    period = period || 6; kN = kN || 3; dN = dN || 3;
    var kArr = [], dArr = [], jArr = [];
    var prevK = 50, prevD = 50;
    for (var i = 0; i < data.length; i++) {
      if (i < period - 1) { kArr.push(null); dArr.push(null); jArr.push(null); continue; }
      var slice = data.slice(i - period + 1, i + 1);
      var lowest  = slice.reduce(function(m,d){ return Math.min(m, d.low);  },  Infinity);
      var highest = slice.reduce(function(m,d){ return Math.max(m, d.high); }, -Infinity);
      var rsv = highest === lowest ? 50 : (data[i].close - lowest) / (highest - lowest) * 100;
      var k = prevK * (kN-1)/kN + rsv * 1/kN;
      var d = prevD * (dN-1)/dN + k   * 1/dN;
      var j = 3*k - 2*d;
      kArr.push(k); dArr.push(d); jArr.push(j);
      prevK = k; prevD = d;
    }
    return { k: kArr, d: dArr, j: jArr };
  }

  function enrich(data) {
    var m5 = ma(data,5), m10 = ma(data,10), m20 = ma(data,20), m60 = ma(data,60);
    var md = macd(data), rs = rsi(data), bbs = bb(data);
    var vm5 = volma(data,5), vm20 = volma(data,20);
    var kd  = calcKD(data, 9);
    var kdj = calcKDJ(data, 6, 3, 3);
    return data.map(function(d, i) {
      return { date:d.date, open:d.open, high:d.high, low:d.low, close:d.close, volume:d.volume,
        ma5:m5[i], ma10:m10[i], ma20:m20[i], ma60:m60[i],
        macd:md.dif[i], macdSig:md.sig[i], macdHist:md.hist[i],
        rsi:rs[i], bbU:bbs[i].u, bbL:bbs[i].l,
        vm5:vm5[i], vm20:vm20[i],
        kdK:kd.k[i], kdD:kd.d[i],
        kdjK:kdj.k[i], kdjD:kdj.d[i], kdjJ:kdj.j[i] };
    });
  }

  // ── Scoring ──
  function scoreTrend(data) {
    var score = 0, sigs = [], last = data[data.length-1], tdir = '盤整';
    var pv = pivots(data);
    if (pv.highs.length >= 2 && pv.lows.length >= 2) {
      var rh = [data[pv.highs[pv.highs.length-2]].high, data[pv.highs[pv.highs.length-1]].high];
      var rl = [data[pv.lows[pv.lows.length-2]].low, data[pv.lows[pv.lows.length-1]].low];
      if (rh[1]>rh[0] && rl[1]>rl[0]) { tdir='多頭'; score+=10; sigs.push(['多頭趨勢確立（高高低低）','bull']); }
      else if (rh[1]<rh[0] && rl[1]<rl[0]) { tdir='空頭'; sigs.push(['空頭趨勢確立（低高低低）','bear']); }
      else { score+=2; sigs.push(['盤整區間','neu']); }
    } else { score+=2; }
    if (last.ma20!=null) { if(last.close>last.ma20){score+=5;sigs.push(['收盤站上20MA','bull']);}else sigs.push(['收盤跌破20MA','bear']); }
    if (last.ma60!=null) { if(last.close>last.ma60){score+=5;sigs.push(['收盤站上60MA','bull']);}else sigs.push(['收盤跌破60MA','bear']); }
    var p10 = data[data.length-10];
    if (last.ma60!=null && p10 && p10.ma60!=null) {
      var sl = (last.ma60-p10.ma60)/p10.ma60*100;
      if(sl>0.5){score+=5;sigs.push(['季線向上 +'+sl.toFixed(1)+'%','bull']);}
      else if(sl<-0.5){sigs.push(['季線向下 '+sl.toFixed(1)+'%','bear']);}
      else{score+=2;sigs.push(['季線走平','neu']);}
    }
    return { score:Math.min(score,25), max:25, sigs:sigs, tdir:tdir };
  }

  function scoreKline(data) {
    var score=0, sigs=[], c=data[data.length-1], p1=data[data.length-2]||c, p2=data[data.length-3]||c;
    var body=Math.abs(c.close-c.open), total=(c.high-c.low)||0.01;
    var upSh=c.high-Math.max(c.close,c.open), dnSh=Math.min(c.close,c.open)-c.low;
    var isBull=c.close>c.open;
    if(isBull){score+=5; if(body/total>0.7){score+=3;sigs.push(['實體長紅棒','bull']);}else sigs.push(['紅K棒','bull']);}
    else{if(body/total>0.7)sigs.push(['實體長黑棒','bear']);else sigs.push(['黑K棒','bear']);}
    var slice20=data.slice(-20);
    var rLow=slice20.reduce(function(m,d){return Math.min(m,d.close)},Infinity);
    var rHigh=slice20.reduce(function(m,d){return Math.max(m,d.close)},-Infinity);
    var pos=(c.close-rLow)/((rHigh-rLow)||0.01);
    if(pos<0.3){ if(dnSh>body*1.5){score+=8;sigs.push(['低檔長下影線（變盤訊號）','bull']);} if(isBull&&c.close>p1.high){score+=5;sigs.push(['低檔紅K突破前高','bull']);} }
    else if(pos>0.7){ if(upSh>body*1.5)sigs.push(['高檔長上影線（變盤訊號）','bear']); }
    var threeBull=p2.close<p2.open&&p1.close<p1.open&&isBull&&c.close>p1.high;
    if(threeBull&&pos<0.4){score+=6;sigs.push(['三K底部反轉組合','bull']);}
    var threeBear=p2.close>p2.open&&p1.close>p1.open&&!isBull&&c.close<p1.low;
    if(threeBear&&pos>0.6)sigs.push(['三K頂部反轉組合','bear']);
    var half=(c.high+c.low)/2;
    if(isBull&&c.close>half){score+=3;sigs.push(['收盤超過1/2價位 '+half.toFixed(1),'bull']);}
    else if(!isBull&&c.close<half)sigs.push(['收盤低於1/2價位 '+half.toFixed(1),'bear']);
    return {score:Math.min(score,25),max:25,sigs:sigs};
  }

  function scoreMA(data) {
    var score=0, sigs=[], last=data[data.length-1], prev=data[data.length-2]||last;
    if(last.ma5!=null&&last.ma10!=null&&last.ma20!=null&&last.ma60!=null){
      if(last.ma5>last.ma10&&last.ma10>last.ma20&&last.ma20>last.ma60){score+=10;sigs.push(['均線多頭排列','bull']);}
      else if(last.ma5<last.ma10&&last.ma10<last.ma20&&last.ma20<last.ma60)sigs.push(['均線空頭排列','bear']);
      else{score+=2;sigs.push(['均線糾結','neu']);}
    }
    if(data.length>=3){
      var p1=data[data.length-2];
      if(p1.ma5!=null&&p1.ma5<p1.ma20&&last.ma5>last.ma20){score+=8;sigs.push(['MA5 黃金交叉 MA20','bull']);}
      else if(p1.ma5!=null&&p1.ma5>p1.ma20&&last.ma5<last.ma20)sigs.push(['MA5 死亡交叉 MA20','bear']);
    }
    if(last.ma20!=null){
      var sl20=data.slice(-10).filter(function(d){return d.ma20!=null;});
      var slope=sl20.length>=2?(sl20[sl20.length-1].ma20-sl20[0].ma20)/sl20[0].ma20*100:0;
      if(slope>0&&last.close>last.ma20&&prev.close<prev.ma20){score+=5;sigs.push(['葛蘭畢買點1（突破均線）','bull']);}
      else if(slope>0&&last.close>last.ma20){
        var diff=(last.close-last.ma20)/last.ma20*100;
        if(diff>0&&diff<3){score+=4;sigs.push(['葛蘭畢買點2（均線支撐）','bull']);}
        else if(diff>=3){score+=2;sigs.push(['均線上揚股價強勢','bull']);}
      }
    }
    if(last.ma60!=null&&last.close>last.ma60){score+=2;sigs.push(['股價位於季線上方','bull']);}
    return {score:Math.min(score,25),max:25,sigs:sigs};
  }

  function scoreVol(data) {
    var score=0, sigs=[], last=data[data.length-1], p1=data[data.length-2]||last;
    var vr=last.vm20?last.volume/last.vm20:1, v5r=(last.vm5&&last.vm20)?last.vm5/last.vm20:1;
    var isUp=last.close>p1.close;
    if(isUp){
      if(vr>=1.5){score+=10;sigs.push(['上漲爆量 '+vr.toFixed(1)+'倍（多頭確認）','bull']);}
      else if(vr>=1.0){score+=6;sigs.push(['上漲放量 '+vr.toFixed(1)+'倍','bull']);}
      else{score+=2;sigs.push(['上漲縮量（動能不足）','neu']);}
    } else {
      if(vr>=1.5)sigs.push(['下跌爆量 '+vr.toFixed(1)+'倍（賣壓沉重）','bear']);
      else if(vr>=1.0)sigs.push(['下跌放量','bear']);
      else{score+=5;sigs.push(['下跌縮量（賣壓減輕）','neu']);}
    }
    var low20=data.slice(-20).reduce(function(m,d){return Math.min(m,d.low)},Infinity);
    if(last.close<low20*1.15&&isUp&&vr>=1.3){score+=8;sigs.push(['底部放量起漲訊號','bull']);}
    if(v5r>1.2){score+=5;sigs.push(['近5日均量擴增 '+v5r.toFixed(1)+'x','bull']);}
    else if(v5r<0.8){score+=1;sigs.push(['近5日均量萎縮','neu']);}
    score+=2;
    return {score:Math.min(score,25),max:25,sigs:sigs};
  }

  // ── DMI（趨向指標，Wilder 原始方法）：+DI／-DI衡量上升與下降方向的動能強弱，
  // ADX衡量趨勢強度（不分方向），ADXR是ADX跟N天前ADX的平均，用來看趨勢是在
  // 增強還是減弱。用「Wilder's smoothing」對TR／+DM／-DM做平滑化（是對累計
  // 總和做平滑，不是簡單移動平均，這是DMI的標準算法，跟MACD用的EMA不同）。
  function calcDMI(data, period) {
    var n = data.length;
    if (n < period + 1) return null;
    var plusDM = [], minusDM = [], tr = [];
    for (var i = 1; i < n; i++) {
      var upMove = data[i].high - data[i-1].high;
      var downMove = data[i-1].low - data[i].low;
      plusDM.push(upMove > downMove && upMove > 0 ? upMove : 0);
      minusDM.push(downMove > upMove && downMove > 0 ? downMove : 0);
      tr.push(Math.max(
        data[i].high - data[i].low,
        Math.abs(data[i].high - data[i-1].close),
        Math.abs(data[i].low - data[i-1].close)
      ));
    }
    if (tr.length < period) return null;
    function wilderSum(arr, p) {
      var out = [], sum = 0;
      for (var j = 0; j < p; j++) sum += arr[j];
      out.push(sum);
      for (var j2 = p; j2 < arr.length; j2++) { sum = sum - sum/p + arr[j2]; out.push(sum); }
      return out;
    }
    var smTR = wilderSum(tr, period), smPlusDM = wilderSum(plusDM, period), smMinusDM = wilderSum(minusDM, period);
    var plusDI = smPlusDM.map(function(v,i){ return smTR[i] ? v/smTR[i]*100 : 0; });
    var minusDI = smMinusDM.map(function(v,i){ return smTR[i] ? v/smTR[i]*100 : 0; });
    var dx = plusDI.map(function(v,i){ var s=v+minusDI[i]; return s ? Math.abs(v-minusDI[i])/s*100 : 0; });
    var adx = null;
    if (dx.length >= period) {
      adx = [];
      var avg = dx.slice(0, period).reduce(function(a,b){return a+b;}, 0) / period;
      adx.push(avg);
      for (var k = period; k < dx.length; k++) { avg = (avg*(period-1) + dx[k]) / period; adx.push(avg); }
    }
    var adxr = null;
    if (adx && adx.length > period) {
      adxr = [];
      for (var m = period; m < adx.length; m++) adxr.push((adx[m] + adx[m-period]) / 2);
    }
    return {
      plusDI: plusDI[plusDI.length-1],
      minusDI: minusDI[minusDI.length-1],
      adx: adx ? adx[adx.length-1] : null,
      adxr: adxr && adxr.length ? adxr[adxr.length-1] : null
    };
  }

  // ── 多方力道評分（0-100，取代原本朱家泓四維度總分）──
  // 組成：①方向性（+DI相對-DI的優勢程度，最高50分）②趨勢強度（ADX，最高30分）
  // ③趨勢轉強加分（ADX>ADXR，20分）。DMI資料不足（新股/資料太短）時各項給0分，
  // 不是「無資料」而是保守給0分，因為總分要能排序／篩選，不能是非數值。
  // sigs／tdir 是為了相容原本detail面板的顯示結構（訊號標籤、多空判斷文字）而設計。
  function scoreDMI(dmi) {
    if (!dmi || dmi.plusDI == null || dmi.minusDI == null) {
      return { score: 0, max: 100, plusDI: null, minusDI: null, adx: null, adxr: null,
        diPts: 0, adxPts: 0, adxrPts: 0, tdir: '資料不足', sigs: [['DMI資料不足（可能資料天數太短）', 'neu']] };
    }
    var diSum = dmi.plusDI + dmi.minusDI;
    var diDominancePct = diSum ? (dmi.plusDI / diSum * 100) : 50; // 50%=中性，100%=完全多方主導
    var diPts = diDominancePct * 0.5; // 0-50分
    var adxPts = dmi.adx != null ? Math.min(dmi.adx, 40) / 40 * 30 : 0; // 0-30分，ADX≥40視為滿分
    var adxrPts = (dmi.adx != null && dmi.adxr != null && dmi.adx > dmi.adxr) ? 20 : 0; // 0或20分
    var score = Math.round(Math.max(0, Math.min(100, diPts + adxPts + adxrPts)));

    var bullish = dmi.plusDI > dmi.minusDI;
    var tdir = bullish ? '多頭' : (dmi.plusDI < dmi.minusDI ? '空頭' : '盤整');

    var sigs = [];
    sigs.push(['+DI ' + dmi.plusDI.toFixed(1) + (bullish?' > ':' < ') + '-DI ' + dmi.minusDI.toFixed(1) + '（多方力道' + (bullish?'較強':'較弱') + '）', bullish?'bull':'bear']);
    if (dmi.adx != null) {
      var adxLbl = dmi.adx >= 25 ? '趨勢明確' : dmi.adx >= 20 ? '趨勢成形中' : '盤整';
      sigs.push(['ADX ' + dmi.adx.toFixed(1) + '（' + adxLbl + '）', dmi.adx >= 25 ? 'bull' : 'neu']);
    }
    if (dmi.adx != null && dmi.adxr != null) {
      var strengthening = dmi.adx > dmi.adxr;
      sigs.push(['ADX' + (strengthening?'>':'<') + 'ADXR，趨勢' + (strengthening?'轉強':'轉弱'), strengthening?'bull':'bear']);
    }

    return { score: score, max: 100, plusDI: dmi.plusDI, minusDI: dmi.minusDI, adx: dmi.adx, adxr: dmi.adxr,
      diPts: diPts, adxPts: adxPts, adxrPts: adxrPts, tdir: tdir, sigs: sigs };
  }

  // ── 個股分價量表（Volume Profile）──把近期股價區間切成幾個價格區間（bin），
  // 累加每個區間的成交量，找出成交量最大的那個區間（POC，Point of Control）。
  // 再從POC往上下擴張，每次納入成交量較大的那一側，直到納入的成交量達到總量的
  // 70%（業界標準的「價值區」Value Area定義），上緣是VAH（Value Area High）、
  // 下緣是VAL（Value Area Low）。概念上這段區間是「市場公認合理」的價格帶，
  // VAL是價值區低點（買盤支撐），VAH是價值區高點（解套賣壓）。
  function computeVolumeProfile(data, lookback, numBins) {
    lookback = lookback || 120;
    numBins = numBins || 24;
    if (data.length < 20) return null;
    var window = data.slice(Math.max(0, data.length - lookback));
    var lo = Math.min.apply(null, window.map(function(d){ return d.low; }));
    var hi = Math.max.apply(null, window.map(function(d){ return d.high; }));
    if (!(hi > lo)) return null;
    var binSize = (hi - lo) / numBins;
    var bins = new Array(numBins).fill(0);
    window.forEach(function(d) {
      var mid = (d.high + d.low) / 2;
      var idx = Math.min(numBins - 1, Math.max(0, Math.floor((mid - lo) / binSize)));
      bins[idx] += d.volume;
    });
    var pocIdx = 0;
    for (var i = 1; i < numBins; i++) { if (bins[i] > bins[pocIdx]) pocIdx = i; }

    var totalVolume = bins.reduce(function(s,v){ return s + v; }, 0);
    var lowIdx = pocIdx, highIdx = pocIdx;
    var includedVolume = bins[pocIdx];
    while (includedVolume < totalVolume * 0.7 && (lowIdx > 0 || highIdx < numBins - 1)) {
      var nextLowVol = lowIdx > 0 ? bins[lowIdx - 1] : -1;
      var nextHighVol = highIdx < numBins - 1 ? bins[highIdx + 1] : -1;
      if (nextHighVol >= nextLowVol) {
        highIdx++;
        includedVolume += bins[highIdx];
      } else {
        lowIdx--;
        includedVolume += bins[lowIdx];
      }
    }

    return {
      pocLow: lo + pocIdx * binSize,
      pocHigh: lo + (pocIdx + 1) * binSize,
      pocPrice: lo + (pocIdx + 0.5) * binSize,
      valLow: lo + lowIdx * binSize,
      valHigh: lo + (highIdx + 1) * binSize,
      binSize: binSize, lo: lo, hi: hi,
    };
  }

  // ── 分價量表訊號（POC版）：以成交量最大的價位區（POC區）為界，不再用VAH/VAL ──
  // 1) 守穩POC買進：前一日收在POC區上方，今日回測POC區、長下影線、收盤守住POC價
  // 2) 突破POC追價買進：帶量長紅K，收盤由POC區下方/區內突破到POC區上方
  // 3) 反彈POC遇壓賣出：前一日收在POC區下方，今日反彈到POC區、長上影線、收盤壓回POC價之下
  // 4) 破位停損賣出：帶量長黑K收盤跌破POC區（2026-09-27 回測：單獨看5/10/20日報酬
  //    皆低於全體基準，確認為偏空訊號）
  // 參數集中在VP_PARAMS，要調整直接改這裡。
  var VP_PARAMS = { lookback: 120, bins: 24, shadowMult: 1.0, volMult: 1.2, bodyMult: 1.2, touchTol: 0.5 };
  function detectVolumeProfileSignal(data, profile) {
    var base = { signal: null, detail: null, sigType: null,
      pocPrice: profile ? profile.pocPrice : null, pocLow: profile ? profile.pocLow : null, pocHigh: profile ? profile.pocHigh : null,
      vah: profile ? profile.valHigh : null, val: profile ? profile.valLow : null };
    if (!profile || data.length < 22) return base;
    var P = VP_PARAMS;
    var last = data[data.length - 1];
    var prev = data[data.length - 2];
    var isRedK = last.close > last.open;
    var isBlackK = last.close < last.open;
    var bodySize = Math.abs(last.close - last.open);
    var upperShadow = last.high - Math.max(last.open, last.close);
    var lowerShadow = Math.min(last.open, last.close) - last.low;
    var recentBars = data.slice(-21, -1); // 不含今天，近20日
    var avgBody = recentBars.reduce(function(s,d){ return s + Math.abs(d.close - d.open); }, 0) / recentBars.length;
    var avgVol = recentBars.reduce(function(s,d){ return s + d.volume; }, 0) / recentBars.length;
    var isLongRed = isRedK && bodySize > avgBody * P.bodyMult;
    var isLongBlack = isBlackK && bodySize > avgBody * P.bodyMult;
    var hasVolume = avgVol > 0 && last.volume > avgVol * P.volMult;
    var pL = profile.pocLow, pH = profile.pocHigh, pP = profile.pocPrice;
    var tol = profile.binSize * P.touchTol;
    var longLower = lowerShadow > 0 && lowerShadow >= bodySize * P.shadowMult;
    var longUpper = upperShadow > 0 && upperShadow >= bodySize * P.shadowMult;

    var support   = prev.close > pH && last.low <= pH && last.low >= pL - tol && last.close >= pP && longLower;
    var breakout  = prev.close <= pH && last.close > pH && isLongRed && hasVolume;
    var resist    = prev.close < pL && last.high >= pL && last.high <= pH + tol && last.close <= pP && longUpper;
    var breakdown = prev.close >= pL && last.close < pL && isLongBlack && hasVolume;

    if (support) {
      base.signal = 'buy'; base.sigType = 'buy_support';
      base.detail = '守穩POC買進：股價由上方回測POC成交密集區，長下影線且收盤守住POC價（停損設POC區之下）';
    } else if (breakout) {
      base.signal = 'buy'; base.sigType = 'buy_breakout';
      base.detail = '突破POC追價買進：帶量長紅K由POC區突破到上方，買方消化最大套牢量';
    } else if (resist) {
      base.signal = 'sell'; base.sigType = 'sell_resistance';
      base.detail = '反彈POC遇壓賣出：股價由下方反彈到POC成交密集區，長上影線且收盤壓回POC價之下';
    } else if (breakdown) {
      base.signal = 'sell'; base.sigType = 'sell_breakdown';
      base.detail = '破位停損賣出：帶量長黑K跌破POC區，賣方潰散，進入成交量真空區（停損/減碼）';
    }
    return base;
  }

  // 算布林通道位置、MACD狀態、強勢突破盤／跌深反彈盤——跟表格顯示邏輯完全一致，
  // 抽成獨立函式給回測用（表格渲染那邊維持原樣不動，避免改動已經在跑的顯示邏輯，
  // 這裡是刻意重複而不是共用，降低風險）。
  function computeCoreSignals(data, dm) {
    var last = data[data.length - 1];
    var prev = data[data.length - 2] || last;

    var bbPos = null;
    if (last.bbU != null && last.bbL != null) {
      var bbW = last.bbU - last.bbL;
      bbPos = bbW ? (last.close - last.bbL) / bbW * 100 : 50;
    }

    var macdState = null, isBreakout = false, isPullbackRebound = false, goldenCrossRecent = false;
    if (last.macd != null && last.macdSig != null && last.macdHist != null && prev.macdHist != null) {
      var aboveZero = last.macd > 0;
      var histGrowing = last.macdHist > prev.macdHist;
      if (aboveZero && last.macdHist > 0) {
        macdState = histGrowing ? '零軸上・紅柱增長' : '零軸上・紅柱縮短';
      } else if (!aboveZero && last.macdHist < 0) {
        macdState = histGrowing ? '零軸下・綠柱縮短' : '零軸下・綠柱增長';
      } else {
        macdState = '交叉轉換中';
      }

      for (var gci = Math.max(1, data.length - 3); gci < data.length; gci++) {
        var gcCur = data[gci], gcPrev = data[gci-1];
        if (gcCur.macd != null && gcCur.macdSig != null && gcPrev.macd != null && gcPrev.macdSig != null &&
            gcPrev.macd <= gcPrev.macdSig && gcCur.macd > gcCur.macdSig) { goldenCrossRecent = true; break; }
      }

      var widthExpanding = false;
      if (last.bbU != null && last.bbL != null) {
        var widthNowPct = last.close ? (last.bbU - last.bbL) / last.close * 100 : 0;
        var refIdx = data.length - 6;
        var refBar = refIdx >= 0 ? data[refIdx] : null;
        if (refBar && refBar.bbU != null && refBar.bbL != null && refBar.close) {
          var widthRefPct = (refBar.bbU - refBar.bbL) / refBar.close * 100;
          widthExpanding = widthNowPct > widthRefPct;
        }
      }

      isBreakout = bbPos != null && bbPos >= 80 && widthExpanding && aboveZero && last.macdHist > 0 && histGrowing;

      var divergenceDetected = false;
      var zzLows = buildZigzag(data).filter(function(p){ return p.type === 'L'; });
      if (zzLows.length >= 2) {
        var recentLow = zzLows[zzLows.length - 1];
        var priorLow = zzLows[zzLows.length - 2];
        var withinLookback = recentLow.idx >= data.length - 1 - 60;
        var macdAtRecent = data[recentLow.idx] ? data[recentLow.idx].macd : null;
        var macdAtPrior = data[priorLow.idx] ? data[priorLow.idx].macd : null;
        if (withinLookback && macdAtRecent != null && macdAtPrior != null) {
          divergenceDetected = (recentLow.price < priorLow.price) && (macdAtRecent > macdAtPrior);
        }
      }

      isPullbackRebound = bbPos != null && bbPos <= 20 && divergenceDetected && goldenCrossRecent;
    }

    // ── KDJ(6,3,3) 黃金交叉／死亡交叉：K由下往上穿越D＝黃金交叉（多方），
    // K由上往下穿越D＝死亡交叉（空方）。跟MACD黃金交叉用同一套「近3天內」判斷法。
    var kdjState = null, kdjGoldenCrossRecent = false, kdjDeathCrossRecent = false;
    if (last.kdjK != null && last.kdjD != null) {
      kdjState = last.kdjK > last.kdjD ? 'K>D（多方）' : (last.kdjK < last.kdjD ? 'K<D（空方）' : 'K=D');
      for (var kci = Math.max(1, data.length - 3); kci < data.length; kci++) {
        var kcCur = data[kci], kcPrev = data[kci-1];
        if (kcCur.kdjK != null && kcCur.kdjD != null && kcPrev.kdjK != null && kcPrev.kdjD != null) {
          if (kcPrev.kdjK <= kcPrev.kdjD && kcCur.kdjK > kcCur.kdjD) kdjGoldenCrossRecent = true;
          if (kcPrev.kdjK >= kcPrev.kdjD && kcCur.kdjK < kcCur.kdjD) kdjDeathCrossRecent = true;
        }
      }
      if (kdjGoldenCrossRecent) kdjState += '　⚡近3日黃金交叉';
      if (kdjDeathCrossRecent) kdjState += '　💀近3日死亡交叉';
    }

    return { bbPos: bbPos, macdState: macdState, isBreakout: isBreakout, isPullbackRebound: isPullbackRebound,
      goldenCrossRecent: goldenCrossRecent,
      kdjState: kdjState, kdjGoldenCrossRecent: kdjGoldenCrossRecent, kdjDeathCrossRecent: kdjDeathCrossRecent };
  }

  // 把批次分析結果的一列（今天的即時資料）轉成跟回測資料同樣結構的「旗標列」，
  // 這樣就能直接重複用回測那邊已經寫好、測過的 buildConditionFlags／PINNED_COMBOS，
  // 兩邊對「型態突破確認＋近3月均價YoY為正」這類條件的定義才會完全一致，不會
  // 出現「回測說符合、即時掃描卻對不上」的落差。
  var currentBenchmark = null; // 這次批次分析/回測用的大盤(0050)資料，跑之前先抓一次，全部股票共用

  function computeLiveFlagsRow(r) {
    // 同一檔的條件旗標只算一次（命中比對、排序時會被呼叫很多次）
    if (r._flagRowCache && r._flagRowBm === currentBenchmark) return r._flagRowCache;
    var fr = computeLiveFlagsRowRaw(r);
    r._flagRowCache = fr; r._flagRowBm = currentBenchmark;
    return fr;
  }
  function computeLiveFlagsRowRaw(r) {
    var dm = r.dm;
    var sig = computeCoreSignals(r.data, dm);
    var relStrength20 = currentBenchmark ? computeRelativeStrength(r.data, currentBenchmark, 20) : null;
    var volRatio = computeVolRatio(r.data);
    var volRangePos = computeVolumeRangePosition(r.data);
    var volTrend = computeVolumeTrend(r.data);
    var volProfile = computeVolumeProfile(r.data, VP_PARAMS.lookback, VP_PARAMS.bins);
    var volProfileSig = detectVolumeProfileSignal(r.data, volProfile);
    var lastDate = r.data[r.data.length - 1].date;
    var regime = currentBenchmark ? computeBenchmarkRegime(currentBenchmark, lastDate) : { above20: null, above60: null };

    var divTotal = null;
    if (r.revRange && r.revRange.length && r.priceYoYRange && r.priceYoYRange.length) {
      var pxByKey = {};
      r.priceYoYRange.forEach(function(p) { pxByKey[p.year + '-' + p.month] = p; });
      var sum = 0, any = false;
      r.revRange.forEach(function(rv) {
        var px = pxByKey[rv.year + '-' + rv.month];
        if (rv.yoy != null && px && px.yoy != null) { sum += rv.yoy - px.yoy; any = true; }
      });
      divTotal = any ? sum : null;
    }

    var priceYoy3m = null;
    if (r.priceYoYRange && r.priceYoYRange.length) {
      var vals = r.priceYoYRange.map(function(p){ return p.yoy; }).filter(function(v){ return v != null; });
      priceYoy3m = vals.length ? vals.reduce(function(a,b){return a+b;}, 0) : null;
    }

    var inst3m = null;
    if (r.instRange && r.instRange.length) {
      inst3m = r.instRange.reduce(function(sum, m) { return sum + m.net; }, 0);
    }

    var patternHits = {};
    if (r.pt && r.pt.results) {
      r.pt.results.forEach(function(pr) { patternHits[pr.id] = pr.justBroke ? 1 : 0; });
    }

    var flagRow = {
      score: dm.score, plusDI: dm.plusDI, minusDI: dm.minusDI, adx: dm.adx, adxr: dm.adxr,
      trust5: (r.instRange && r.instRange.flow) ? r.instRange.flow.trust5 : null,
      foreign5: (r.instRange && r.instRange.flow) ? r.instRange.flow.foreign5 : null,
      trustStreak3: (r.instRange && r.instRange.flow) ? r.instRange.flow.trustStreak3 : null,
      isBreakout: sig.isBreakout ? 1 : 0,
      isPullbackRebound: sig.isPullbackRebound ? 1 : 0,
      bbPos: sig.bbPos,
      goldenCrossRecent: sig.goldenCrossRecent ? 1 : 0,
      kdjGoldenCrossRecent: sig.kdjGoldenCrossRecent ? 1 : 0,
      kdjDeathCrossRecent: sig.kdjDeathCrossRecent ? 1 : 0,
      relStrength20: relStrength20,
      volRatio: volRatio,
      volRangePos: volRangePos,
      volTrend: volTrend,
      vpBuySupport: volProfileSig.sigType === 'buy_support' ? 1 : 0,
      vpBuyBreakout: volProfileSig.sigType === 'buy_breakout' ? 1 : 0,
      vpSellResistance: volProfileSig.sigType === 'sell_resistance' ? 1 : 0,
      vpSellBreakdown: volProfileSig.sigType === 'sell_breakdown' ? 1 : 0,
      vpPocPrice: volProfileSig.pocPrice,
      benchmarkAbove20: regime.above20 == null ? null : (regime.above20 ? 1 : 0),
      benchmarkAbove60: regime.above60 == null ? null : (regime.above60 ? 1 : 0),
      patternFormed: r.pt.anyFormed ? 1 : 0,
      patternBreakout: r.pt.anyBreakout ? 1 : 0,
      patternJustBroke: r.pt.anyJustBroke ? 1 : 0,
      patternHits: patternHits,
      pbAllPass: r.pb.allPass ? 1 : 0,
      divTotal: divTotal, priceYoy3m: priceYoy3m, inst3m: inst3m,
      revYoyTurn: r.revFeats ? r.revFeats.revYoyTurn : null, revYoy3Pos: r.revFeats ? r.revFeats.revYoy3Pos : null, revHigh24: r.revFeats ? r.revFeats.revHigh24 : null,
    };
    if (r.tech) Object.keys(r.tech).forEach(function(k) { flagRow[k] = r.tech[k]; });
    flagRow.confluenceCount = computeConfluenceCount(flagRow);
    return flagRow;
  }

  // 這檔股票「今天」符合PINNED_COMBOS裡的哪幾組，回傳組合的index陣列
  function matchedPinnedCombos(r) {
    var fs = liveFlagSet(r);
    var matched = [];
    PINNED_COMBOS.forEach(function(combo, idx) {
      var ok = combo.every(function(name) { return fs[name] === 1; });
      if (ok) matched.push(idx);
    });
    return matched;
  }

  // 每檔股票的條件只算一次就快取（2026-10-19）：舊寫法每比對一類清單就重算全部條件，
  // 摘要表、大盤適用欄、溫度計、排序／篩選每次都重算，股票一多（1000檔）畫面會非常慢
  function liveFlagSet(r) {
    if (r.__fs) return r.__fs;
    var fs = {};
    try {
      var fr = computeLiveFlagsRow(r), defs = buildConditionFlags([fr]);
      Object.keys(defs).forEach(function(n) { try { if (defs[n](fr)) fs[n] = 1; } catch (e) {} });
    } catch (e) {}
    try { Object.defineProperty(r, '__fs', { value: fs, writable: true, configurable: true, enumerable: false }); } catch (e) { r.__fs = fs; }
    return fs;
  }

  function matchedMoonshotCombos(r) {
    var fs = liveFlagSet(r);
    var matched = [];
    MOONSHOT_COMBOS.forEach(function(combo, idx) {
      var ok = combo.every(function(name) { return fs[name] === 1; });
      if (ok) matched.push(idx);
    });
    return matched;
  }

  // 選股型指定組合（S編號）：2023-10～2026-09 台股三年回測（分三段：偏多／空頭／多頭），三年的依日 t 值都 ≥2 的組合（S4、S7、S8 為 2026-10-05 加入的營收創新高系列）
  // （2026-09-30 改版；舊版 S1～S7 多數只在多頭年有效，已移除或改列高標股 M）
  var STOCK_PICK_COMBOS = [
    ["距52週高點≤5%", "近3日向上跳空缺口", "近3月乖離度為正(營收優於股價)"],
    ["創52週新高", "近3日向上跳空缺口", "近3月乖離度為正(營收優於股價)"],
    ["相對強弱為正(強於大盤)", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["近3日漲停", "月營收創24個月新高", "外資近5日買超"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "距52週高點≤5%"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "創52週新高"],
    ["大盤站上20日均線", "漲時量≥跌時量1.5倍(近20日)", "月營收創24個月新高"],
    ["近20日強於大盤≥10%", "大盤站上20日均線", "月營收創24個月新高"],

  ];
  var STOCK_PICK_STATS = {
    "距52週高點≤5% ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 4029, "win": 54.6, "ret": 3.33, "ex": 2.73, "t": 12.61}, "20": {"n": 4029, "win": 54.7, "ret": 5.17, "ex": 3.76, "t": 12.27}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2838, "win": 53.6, "ret": 3.79, "ex": 3.17, "t": 11.06}, "20": {"n": 2838, "win": 52.9, "ret": 5.53, "ex": 4.1, "t": 10.09}},
    "相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"10": {"n": 644, "win": 57.8, "ret": 2.28, "ex": 2.15, "t": 6.11}, "20": {"n": 644, "win": 56.7, "ret": 3.48, "ex": 2.29, "t": 4.93}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ 外資近5日買超": {"10": {"n": 2970, "win": 55.8, "ret": 4.88, "ex": 3.42, "t": 11.37}, "20": {"n": 2970, "win": 57.0, "ret": 8.59, "ex": 5.81, "t": 13.21}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 距52週高點≤5%": {"10": {"n": 2004, "win": 53.4, "ret": 2.69, "ex": 2.25, "t": 7.85}, "20": {"n": 2004, "win": 55.7, "ret": 5.25, "ex": 4.14, "t": 9.5}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 創52週新高": {"10": {"n": 267, "win": 59.2, "ret": 6.59, "ex": 6.22, "t": 5.5}, "20": {"n": 267, "win": 64.4, "ret": 12.11, "ex": 10.46, "t": 6.38}},
    "大盤站上20日均線 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 月營收創24個月新高": {"10": {"n": 16520, "win": 50.5, "ret": 2.37, "ex": 1.72, "t": 18.15}, "20": {"n": 16520, "win": 51.3, "ret": 4.62, "ex": 3.23, "t": 22.56}},
    "近20日強於大盤≥10% ＋ 大盤站上20日均線 ＋ 月營收創24個月新高": {"10": {"n": 8373, "win": 51.8, "ret": 3.09, "ex": 2.42, "t": 15.2}, "20": {"n": 8373, "win": 52.2, "ret": 5.35, "ex": 3.83, "t": 16.44}},

  };

  // 擇時型因子：描述的是「大盤／整體位置」，含這些條件的高勝率組合，勝率多半來自「出現在大盤反彈的
  // 日子」（2026-09-27 台股回測：扣掉同日大盤後超額報酬≈0），適合判斷大盤落底，不適合當選股依據。
  var TIMING_FACTORS = { '布林通道低檔(≤20%)': true, '大盤跌破20日均線': true, '大盤跌破60日均線': true, '類股跌破20日均線': true, '類股跌破60日均線': true };
  function isTimingCombo(combo) { return combo.some(function(c){ return TIMING_FACTORS[c]; }); }
  function formatStockPickStats(st) {
    if (!st) return '';
    return Object.keys(st).sort(function(a,b){ return a-b; }).map(function(h) {
      var e = st[h];
      return h + '日：' + e.n + '筆 勝率' + e.win + '% 同日超額' + (e.ex >= 0 ? '+' : '') + e.ex.toFixed(2) + '% t=' + e.t;
    }).join('　｜　');
  }

  // ── 回測自動入選組合（2026-10-05 台股三年回測，含月營收／外資／漲停新條件）：W＝回測高勝率（任一天期勝率>55%、三年各自依日 t≥2、六個半年至少五個為正，依最低年度t排序）；
  //    F＝回測飆股（10/20日飆股比例>10%、飆股次數≥10，且三段飆股比例都≥2倍基準，依最高飆股比例排序）。已在 S／#／M 裡的不重複列入。
  var BT_WIN_COMBOS = [
    ["大盤跌破60日均線", "月營收連3月年增", "月營收創24個月新高"],
    ["近3日漲停", "月營收連3月年增", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "均線多頭排列(5>20>60且站上月線)"],
    ["近3日向上跳空缺口", "近3日漲停", "月營收創24個月新高"],
    ["近3日漲停", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "漲時量≥跌時量1.5倍(近20日)"],
    ["近3日漲停", "三大法人近3月買超", "月營收創24個月新高"],
    ["強勢突破盤", "近3日漲停", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "外資近5日買超"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "三大法人近3月買超"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "近3月均價YoY為正"],
    ["強勢突破盤", "創52週新高", "月營收創24個月新高"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)"],
    ["近3日向上跳空缺口", "月營收連3月年增", "月營收創24個月新高"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "近3月乖離度為負(股價超前營收)"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日漲停"],
    ["地量(≤0.5倍均量)", "創52週新高", "近3月均價YoY為負"],
    ["大盤跌破60日均線", "月營收創24個月新高", "外資近5日買超"],
    ["多方力道≥80", "量能區間低檔(≤10百分位)", "月營收連3月年增"],
    ["回後買上漲全通過", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["多方力道≥65", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["強勢突破盤", "近20日強於大盤≥10%", "地量(≤0.5倍均量)"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "外資近5日買超"],
    ["短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)", "N字底剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日向上跳空缺口"],

  ];
  var BT_WIN_STATS = {
    "大盤跌破60日均線 ＋ 月營收連3月年增 ＋ 月營收創24個月新高": {"5": {"n": 10070, "win": 60.3, "ex": 0.8, "t": 12.15}, "10": {"n": 10070, "win": 60.6, "ex": 1.29, "t": 13.26}, "20": {"n": 10070, "win": 60.5, "ex": 2.15, "t": 14.99}},
    "近3日漲停 ＋ 月營收連3月年增 ＋ 月營收創24個月新高": {"5": {"n": 3713, "win": 53.4, "ex": 1.85, "t": 9.9}, "10": {"n": 3713, "win": 55.4, "ex": 3.23, "t": 12.16}, "20": {"n": 3713, "win": 57.9, "ex": 5.44, "t": 14.52}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 均線多頭排列(5>20>60且站上月線)": {"5": {"n": 752, "win": 57.3, "ex": 1.29, "t": 5.92}, "10": {"n": 752, "win": 57.8, "ex": 2.06, "t": 6.49}, "20": {"n": 752, "win": 58.0, "ex": 2.47, "t": 5.68}},
    "近3日向上跳空缺口 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"5": {"n": 2745, "win": 55.0, "ex": 2.02, "t": 9.15}, "10": {"n": 2745, "win": 57.2, "ex": 3.32, "t": 11.01}, "20": {"n": 2745, "win": 58.3, "ex": 5.23, "t": 12.06}},
    "近3日漲停 ＋ 月營收創24個月新高": {"5": {"n": 4781, "win": 53.4, "ex": 1.81, "t": 10.83}, "10": {"n": 4781, "win": 54.8, "ex": 3.0, "t": 12.72}, "20": {"n": 4781, "win": 56.7, "ex": 5.06, "t": 14.81}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 漲時量≥跌時量1.5倍(近20日)": {"5": {"n": 506, "win": 55.3, "ex": 1.31, "t": 4.96}, "10": {"n": 506, "win": 54.7, "ex": 2.15, "t": 5.45}, "20": {"n": 506, "win": 53.0, "ex": 1.56, "t": 3.05}},
    "近3日漲停 ＋ 三大法人近3月買超 ＋ 月營收創24個月新高": {"5": {"n": 3267, "win": 54.0, "ex": 2.01, "t": 9.62}, "10": {"n": 3267, "win": 54.7, "ex": 3.33, "t": 11.23}, "20": {"n": 3267, "win": 56.4, "ex": 5.68, "t": 13.31}},
    "強勢突破盤 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"5": {"n": 1862, "win": 54.2, "ex": 2.8, "t": 9.76}, "10": {"n": 1862, "win": 54.4, "ex": 4.16, "t": 10.51}, "20": {"n": 1862, "win": 55.4, "ex": 6.62, "t": 11.17}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 外資近5日買超": {"5": {"n": 667, "win": 55.8, "ex": 0.95, "t": 4.43}, "10": {"n": 667, "win": 57.4, "ex": 1.83, "t": 5.98}, "20": {"n": 667, "win": 57.6, "ex": 2.0, "t": 4.84}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 三大法人近3月買超": {"5": {"n": 909, "win": 54.8, "ex": 0.89, "t": 5.03}, "10": {"n": 909, "win": 55.8, "ex": 1.52, "t": 5.93}, "20": {"n": 909, "win": 55.7, "ex": 1.86, "t": 5.16}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"5": {"n": 1475, "win": 53.9, "ex": 2.68, "t": 8.13}, "10": {"n": 1475, "win": 56.2, "ex": 4.51, "t": 9.85}, "20": {"n": 1475, "win": 56.0, "ex": 7.61, "t": 10.97}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 近3月均價YoY為正": {"5": {"n": 1041, "win": 54.7, "ex": 0.93, "t": 5.69}, "10": {"n": 1041, "win": 54.8, "ex": 1.59, "t": 6.57}, "20": {"n": 1041, "win": 56.2, "ex": 1.9, "t": 5.71}},
    "強勢突破盤 ＋ 創52週新高 ＋ 月營收創24個月新高": {"5": {"n": 2077, "win": 51.4, "ex": 1.69, "t": 6.92}, "10": {"n": 2077, "win": 53.7, "ex": 2.9, "t": 8.59}, "20": {"n": 2077, "win": 55.1, "ex": 5.7, "t": 11.12}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": {"5": {"n": 311, "win": 55.6, "ex": 3.67, "t": 4.78}, "10": {"n": 311, "win": 60.1, "ex": 6.44, "t": 5.7}, "20": {"n": 311, "win": 63.3, "ex": 11.14, "t": 6.81}},
    "近3日向上跳空缺口 ＋ 月營收連3月年增 ＋ 月營收創24個月新高": {"5": {"n": 5985, "win": 53.2, "ex": 1.0, "t": 8.79}, "10": {"n": 5985, "win": 55.4, "ex": 1.98, "t": 12.42}, "20": {"n": 5985, "win": 56.0, "ex": 3.34, "t": 14.76}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 近3月乖離度為負(股價超前營收)": {"5": {"n": 681, "win": 55.9, "ex": 1.13, "t": 5.26}, "10": {"n": 681, "win": 55.1, "ex": 1.71, "t": 5.59}, "20": {"n": 681, "win": 56.5, "ex": 2.26, "t": 5.03}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"5": {"n": 217, "win": 57.6, "ex": 4.68, "t": 4.79}, "10": {"n": 217, "win": 59.9, "ex": 6.92, "t": 4.72}, "20": {"n": 217, "win": 59.4, "ex": 10.72, "t": 5.05}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 近3月均價YoY為負": {"5": {"n": 129, "win": 60.5, "ex": 4.71, "t": 4.45}, "10": {"n": 129, "win": 62.0, "ex": 6.35, "t": 4.59}, "20": {"n": 129, "win": 61.2, "ex": 8.37, "t": 4.12}},
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 外資近5日買超": {"5": {"n": 6770, "win": 59.4, "ex": 0.65, "t": 8.39}, "10": {"n": 6770, "win": 59.7, "ex": 1.02, "t": 8.84}, "20": {"n": 6770, "win": 59.2, "ex": 1.86, "t": 10.65}},
    "多方力道≥80 ＋ 量能區間低檔(≤10百分位) ＋ 月營收連3月年增": {"5": {"n": 352, "win": 52.3, "ex": 1.34, "t": 3.34}, "10": {"n": 352, "win": 55.7, "ex": 3.28, "t": 5.2}, "20": {"n": 352, "win": 54.5, "ex": 3.5, "t": 3.92}},
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"5": {"n": 579, "win": 51.1, "ex": 2.35, "t": 4.38}, "10": {"n": 579, "win": 54.2, "ex": 3.03, "t": 4.21}, "20": {"n": 579, "win": 56.0, "ex": 6.48, "t": 5.77}},
    "多方力道≥65 ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"5": {"n": 463, "win": 55.9, "ex": 1.19, "t": 4.05}, "10": {"n": 463, "win": 57.0, "ex": 2.13, "t": 4.98}, "20": {"n": 463, "win": 58.5, "ex": 2.2, "t": 3.72}},
    "強勢突破盤 ＋ 近20日強於大盤≥10% ＋ 地量(≤0.5倍均量)": {"5": {"n": 402, "win": 54.5, "ex": 2.86, "t": 4.59}, "10": {"n": 402, "win": 57.2, "ex": 4.93, "t": 5.43}, "20": {"n": 402, "win": 57.7, "ex": 8.25, "t": 6.18}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"5": {"n": 193, "win": 61.7, "ex": 2.71, "t": 4.33}, "10": {"n": 193, "win": 61.7, "ex": 5.04, "t": 5.69}, "20": {"n": 193, "win": 65.3, "ex": 5.54, "t": 4.37}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"5": {"n": 368, "win": 57.1, "ex": 2.92, "t": 4.96}, "10": {"n": 368, "win": 54.3, "ex": 3.61, "t": 4.26}, "20": {"n": 368, "win": 56.8, "ex": 6.39, "t": 5.2}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 外資近5日買超": {"5": {"n": 1682, "win": 46.8, "ex": 0.31, "t": 2.4}, "10": {"n": 1682, "win": 50.3, "ex": 0.86, "t": 4.42}, "20": {"n": 1682, "win": 55.5, "ex": 1.45, "t": 4.76}},
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價) ＋ N字底剛形成": {"5": {"n": 138, "win": 55.1, "ex": 4.87, "t": 3.8}, "10": {"n": 138, "win": 51.4, "ex": 5.75, "t": 3.2}, "20": {"n": 138, "win": 52.2, "ex": 7.55, "t": 3.06}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"5": {"n": 291, "win": 55.7, "ex": 3.63, "t": 4.76}, "10": {"n": 291, "win": 58.4, "ex": 6.09, "t": 5.51}, "20": {"n": 291, "win": 59.8, "ex": 9.76, "t": 6.07}},

  };
  var BT_HOT_COMBOS = [
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "三大法人近3月賣超"],
    ["距52週高點≤5%", "近3日漲停", "月營收創24個月新高"],
    ["短線過熱(高於月線≥20%)", "月營收連3月年增", "N字底剛形成"],
    ["近20日強於大盤≥10%", "地量(≤0.5倍均量)", "近3日漲停"],
    ["近3日漲停", "突破ABC修正下降切線剛形成", "K線橫盤的突破剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日漲停"],
    ["地量(≤0.5倍均量)", "OBV能量潮創60日新高", "近3日漲停"],
    ["地量(≤0.5倍均量)", "創52週新高", "三大法人近3月賣超"],
    ["短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)", "N字底剛形成"],
    ["近3日漲停", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "CMF資金流買方佔優(近20日≥0.1)", "近3日漲停"],
    ["相對強弱為正(強於大盤)", "地量(≤0.5倍均量)", "近3日漲停"],
    ["多方力道≥65", "地量(≤0.5倍均量)", "近3日漲停"],
    ["多方力道≥80", "近3日漲停", "K線橫盤的突破剛形成"],
    ["OBV能量潮創60日新高", "近3日漲停", "N字底剛形成"],
    ["大盤站上20日均線", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["地量(≤0.5倍均量)", "創52週新高", "近3日漲停"],
    ["地量(≤0.5倍均量)", "創52週新高", "月營收創24個月新高"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["CMF資金流買方佔優(近20日≥0.1)", "近3日漲停", "N字底剛形成"],
    ["創52週新高", "近3日漲停", "K線橫盤的突破剛形成"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "K線橫盤的突破剛形成"],
    ["量能區間高檔(≥90百分位)", "近3日漲停", "N字底剛形成"],
    ["布林通道高檔(≥80%)", "地量(≤0.5倍均量)", "近3日漲停"],
    ["距52週高點≤5%", "近3日漲停", "K線橫盤的突破剛形成"],
    ["創52週新高", "近3日漲停", "N字底剛形成"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "N字底剛形成"],
    ["相對強弱為負(弱於大盤)", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["多方力道≥80", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "近3日漲停", "N字底剛形成"],
    ["大盤站上20日均線", "近3日漲停", "N字底剛形成"],
    ["漲時量≥跌時量1.5倍(近20日)", "近3日漲停", "N字底剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "近3日漲停", "N字底剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "近3日漲停", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "N字底剛形成"],
    ["近3日漲停", "月營收創24個月新高", "圓弧底剛形成"],
    ["地量(≤0.5倍均量)", "OBV能量潮創60日新高", "短線過熱(高於月線≥20%)"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "近3日向上跳空缺口"],
    ["多方力道≥80", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["多方力道≥80", "近3日漲停", "N字底剛形成"],
    ["多方力道≥65", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "N字底剛形成"],
    ["多方力道≥65", "近3日漲停", "N字底剛形成"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "近3日漲停"],
    ["近3日漲停", "外資近5日買超", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "近3日漲停", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "CMF資金流買方佔優(近20日≥0.1)", "晨星剛形成"],
    ["強勢突破盤", "近3日漲停", "N字底剛形成"],
    ["大盤跌破20日均線", "月營收創24個月新高", "夜星剛形成"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)"],
    ["地量(≤0.5倍均量)", "近3日向上跳空缺口", "短線過熱(高於月線≥20%)"],
    ["地量(≤0.5倍均量)", "均線多頭排列(5>20>60且站上月線)", "近3日漲停"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["布林通道高檔(≥80%)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["大盤跌破20日均線", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["近20日強於大盤≥10%", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["近20日強於大盤≥10%", "近3日漲停", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "近3日漲停", "三大法人近3月買超"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "突破ABC修正下降切線剛形成"],
    ["短線過熱(高於月線≥20%)", "三大法人近3月買超", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "漲時量≥跌時量1.5倍(近20日)", "近3日漲停"],
    ["近3日漲停", "月營收連3月年增", "突破ABC修正下降切線剛形成"],
    ["布林通道高檔(≥80%)", "近3日漲停", "N字底剛形成"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["強勢突破盤", "地量(≤0.5倍均量)", "月營收創24個月新高"],
    ["創52週新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "母子懷抱(低檔)剛形成"],
    ["距52週高點≤5%", "近3日漲停", "N字底剛形成"],
    ["量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["回後買上漲全通過", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["地量(≤0.5倍均量)", "距52週高點≤5%", "近3日漲停"],
    ["大盤站上60日均線", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["大盤跌破60日均線", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "近3日漲停", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "晨星剛形成"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "創52週新高"],
    ["量能區間低檔(≤10百分位)", "漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)"],
    ["距52週高點≤5%", "月營收創24個月新高", "晨星剛形成"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "近3日漲停", "N字底剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "月營收創24個月新高"],
    ["均線多頭排列(5>20>60且站上月線)", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["強勢突破盤", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["近3日漲停", "月營收連3月年增", "N字底剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "K線橫盤的突破剛形成"],
    ["爆量(≥2倍均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "距52週高點≤5%", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "CMF資金流買方佔優(近20日≥0.1)", "晨星剛形成"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["多方力道≥65", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "OBV能量潮創60日新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["KDJ近3日內死亡交叉", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["近3日向上跳空缺口", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "晨星剛形成"],
    ["創52週新高", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["MACD近3日內黃金交叉", "爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)"],
    ["地量(≤0.5倍均量)", "漲時量≥跌時量1.5倍(近20日)", "夜星剛形成"],
    ["量能區間低檔(≤10百分位)", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["多方力道≥65", "量能區間低檔(≤10百分位)", "創52週新高"],
    ["量能區間低檔(≤10百分位)", "大盤站上60日均線", "短線過熱(高於月線≥20%)"],
    ["量能區間低檔(≤10百分位)", "CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)"],
    ["量能區間低檔(≤10百分位)", "均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)", "近3月均價YoY為正"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)", "三大法人近3月買超"],
    ["地量(≤0.5倍均量)", "量能區間高檔(≥90百分位)", "近3日漲停"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "距52週高點≤5%"],
    ["近3月均價YoY為正", "月營收創24個月新高", "夜星剛形成"],
    ["OBV能量潮創60日新高", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)"],
    ["月營收連3月年增", "月營收創24個月新高", "夜星剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)", "月營收創24個月新高"],
    ["大盤跌破60日均線", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "晨星剛形成"],
    ["近20日強於大盤≥10%", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["創52週新高", "近3月均價YoY為負", "突破ABC修正下降切線剛形成"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "近3日漲停", "月營收連3月年增"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["多方力道≥65", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "大盤跌破60日均線", "短線過熱(高於月線≥20%)"],
    ["近3日漲停", "月營收創24個月新高", "N字底剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["多方力道≥80", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "漲時量≥跌時量1.5倍(近20日)", "晨星剛形成"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)"],
    ["地量(≤0.5倍均量)", "大盤站上20日均線", "近3日漲停"],
    ["布林通道高檔(≥80%)", "月營收創24個月新高", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["短線過熱(高於月線≥20%)", "外資近5日買超", "N字底剛形成"],
    ["MACD近3日內黃金交叉", "OBV能量潮創60日新高", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "爆量(≥2倍均量)", "短線過熱(高於月線≥20%)"],
    ["大盤站上20日均線", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月均價YoY為負", "N字底剛形成"],
    ["創52週新高", "近3日漲停", "月營收創24個月新高"],
    ["CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "月營收連3月年增"],
    ["近20日強於大盤≥10%", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["多方力道≥65", "MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["短線過熱(高於月線≥20%)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "突破ABC修正下降切線剛形成"],
    ["地量(≤0.5倍均量)", "量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)"],
    ["多方力道≥80", "短線過熱(高於月線≥20%)", "突破ABC修正下降切線剛形成"],
    ["月營收創24個月新高", "夜星剛形成"],
    ["短線過熱(高於月線≥20%)", "三大法人近3月買超", "月營收創24個月新高"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["近20日強於大盤≥10%", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["距52週高點≤5%", "近3日漲停", "近3月均價YoY為負"],
    ["近20日強於大盤≥10%", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)"],
    ["布林通道高檔(≥80%)", "量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["大盤跌破20日均線", "月營收連3月年增", "夜星剛形成"],
    ["近20日強於大盤≥10%", "地量(≤0.5倍均量)", "夜星剛形成"],
    ["多方力道≥65", "量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["近20日強於大盤≥10%", "量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "月營收創24個月新高", "晨星剛形成"],
    ["量能區間低檔(≤10百分位)", "短線過熱(高於月線≥20%)"],
    ["均線多頭排列(5>20>60且站上月線)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "均線多頭排列(5>20>60且站上月線)", "晨星剛形成"],
    ["相對強弱為正(強於大盤)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能區間高檔(≥90百分位)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["大盤站上60日均線", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "三大法人近3月買超", "母子懷抱(低檔)剛形成"],
    ["多方力道≥65", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["距52週高點≤5%", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["三大法人近3月買超", "月營收創24個月新高", "夜星剛形成"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["強勢突破盤", "月營收連3月年增", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "近3日漲停", "三大法人近3月賣超"],
    ["MACD近3日內黃金交叉", "KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["創52週新高", "近3日向上跳空缺口", "三大法人近3月賣超"],
    ["創52週新高", "近3日漲停", "近3月均價YoY為負"],
    ["漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日向上跳空缺口", "近3日漲停"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "均線多頭排列(5>20>60且站上月線)", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "三大法人近3月買超", "突破ABC修正下降切線剛形成"],
    ["相對強弱為負(弱於大盤)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "近3日漲停"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "突破飆股大量黑K最高點剛形成"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3日漲停"],
    ["大盤站上60日均線", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["相對強弱為正(強於大盤)", "月營收創24個月新高", "夜星剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "創52週新高"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "距52週高點≤5%", "近3日漲停"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "三大法人近3月買超"],
    ["布林通道高檔(≥80%)", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["MACD近3日內黃金交叉", "漲時量≥跌時量1.5倍(近20日)", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "近3日向上跳空缺口", "短線過熱(高於月線≥20%)"],
    ["回後買上漲全通過", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["爆量(≥2倍均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["近3日向上跳空缺口", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "外資近5日買超"],
    ["大盤跌破60日均線", "外資近5日買超", "夜星剛形成"],
    ["KDJ近3日內死亡交叉", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["OBV能量潮創60日新高", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近20日強於大盤≥10%", "創52週新高", "突破ABC修正下降切線剛形成"],
    ["爆量(≥2倍均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近3日漲停", "突破ABC修正下降切線剛形成", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "月營收創24個月新高", "外資近5日買超"],
    ["創52週新高", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["MACD近3日內黃金交叉", "大盤站上20日均線", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "相對強弱為正(強於大盤)", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "近20日強於大盤≥10%", "短線過熱(高於月線≥20%)"],
    ["創52週新高", "近3日漲停", "近3月乖離度為正(營收優於股價)"],
    ["KDJ近3日內死亡交叉", "爆量(≥2倍均量)", "短線過熱(高於月線≥20%)"],
    ["多方力道≥80", "創52週新高", "突破ABC修正下降切線剛形成"],
    ["近20日強於大盤≥10%", "近3日向上跳空缺口", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "三大法人近3月賣超"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "突破飆股大量黑K最高點剛形成", "晨星剛形成"],
    ["MACD近3日內黃金交叉", "量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["強勢突破盤", "月營收連3月年增", "晨星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "創52週新高", "月營收創24個月新高"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)", "月營收創24個月新高"],
    ["爆量(≥2倍均量)", "月營收連3月年增", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "三大法人近3月賣超"],
    ["CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近3日漲停", "N字底剛形成", "K線橫盤的突破剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停"],
    ["MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)", "三大法人近3月買超"],
    ["MACD近3日內黃金交叉", "CMF資金流買方佔優(近20日≥0.1)", "短線過熱(高於月線≥20%)"],
    ["創52週新高", "近3日向上跳空缺口", "近3月均價YoY為負"],
    ["量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["地量(≤0.5倍均量)", "創52週新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)"],
    ["地量(≤0.5倍均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為負(股價超前營收)"],
    ["地量(≤0.5倍均量)", "近3日向上跳空缺口", "近3日漲停"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "近3月均價YoY為正"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "大盤跌破20日均線", "近3日漲停"],
    ["創52週新高", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["強勢突破盤", "MACD近3日內黃金交叉", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["近3日漲停", "近3月乖離度為正(營收優於股價)", "K線橫盤的突破剛形成"],
    ["近20日強於大盤≥10%", "地量(≤0.5倍均量)", "晨星剛形成"],
    ["距52週高點≤5%", "K線橫盤的突破剛形成", "晨星剛形成"],
    ["近20日強於大盤≥10%", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月均價YoY為負", "月營收創24個月新高"],
    ["均線多頭排列(5>20>60且站上月線)", "K線橫盤的突破剛形成", "母子懷抱(低檔)剛形成"],
    ["相對強弱為正(強於大盤)", "創52週新高", "母子懷抱(低檔)剛形成"],
    ["創52週新高", "母子懷抱(低檔)剛形成"],
    ["地量(≤0.5倍均量)", "量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)"],
    ["布林通道高檔(≥80%)", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["距52週高點≤5%", "近3日漲停", "母子懷抱(低檔)剛形成"],
    ["KDJ近3日內死亡交叉", "月營收創24個月新高", "夜星剛形成"],
    ["量能斜率轉弱(近5日均量<近10日均量20%以上)", "近3日漲停", "近3月乖離度為負(股價超前營收)"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["爆量(≥2倍均量)", "創52週新高", "月營收創24個月新高"],
    ["近3日漲停", "三大法人近3月買超", "突破ABC修正下降切線剛形成"],
    ["創52週新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "突破ABC修正下降切線剛形成"],
    ["強勢突破盤", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["強勢突破盤", "月營收創24個月新高", "突破ABC修正下降切線剛形成"],
    ["大盤跌破60日均線", "近3日漲停", "月營收創24個月新高"],
    ["地量(≤0.5倍均量)", "近3日向上跳空缺口", "OBV能量潮創60日新高"],
    ["距52週高點≤5%", "短線過熱(高於月線≥20%)", "突破ABC修正下降切線剛形成"],
    ["近3日向上跳空缺口", "近3日漲停", "月營收創24個月新高"],
    ["近3日漲停", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "母子懷抱(低檔)剛形成"],
    ["爆量(≥2倍均量)", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["OBV能量潮創60日新高", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["連續放量(近3日均量≥1.5倍前20日均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["近20日強於大盤≥10%", "三大法人近3月買超", "夜星剛形成"],
    ["OBV能量潮創60日新高", "近3日漲停", "月營收創24個月新高"],
    ["量能斜率轉強(近5日均量>近10日均量20%以上)", "距52週高點≤5%", "晨星剛形成"],
    ["均線多頭排列(5>20>60且站上月線)", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["距52週高點≤5%", "近3日漲停", "近3月乖離度為正(營收優於股價)"],
    ["爆量(≥1.5倍均量)", "短線過熱(高於月線≥20%)", "近3月乖離度為正(營收優於股價)"],
    ["KDJ近3日內黃金交叉", "量能斜率轉強(近5日均量>近10日均量20%以上)", "短線過熱(高於月線≥20%)"],
    ["近3日漲停", "近3月乖離度為正(營收優於股價)", "月營收創24個月新高"],
    ["回後買上漲全通過", "創52週新高", "三大法人近3月賣超"],
    ["爆量(≥2倍均量)", "大盤跌破60日均線", "短線過熱(高於月線≥20%)"],
    ["KDJ近3日內黃金交叉", "短線過熱(高於月線≥20%)", "N字底剛形成"],
    ["多方力道≥80", "地量(≤0.5倍均量)", "OBV能量潮創60日新高"],
    ["多方力道≥80", "CMF資金流買方佔優(近20日≥0.1)", "晨星剛形成"],
    ["短線過熱(高於月線≥20%)", "近3月均價YoY為負", "外資近5日買超"],
    ["大盤站上20日均線", "距52週高點≤5%", "短線過熱(高於月線≥20%)"],
    ["近20日強於大盤≥10%", "短線過熱(高於月線≥20%)", "近3月均價YoY為負"],
    ["MACD近3日內黃金交叉", "量能區間高檔(≥90百分位)", "短線過熱(高於月線≥20%)"],
    ["強勢突破盤", "月營收創24個月新高", "突破飆股大量黑K最高點剛形成"],
    ["KDJ近3日內黃金交叉", "近3日漲停", "突破ABC修正下降切線剛形成"],
    ["月營收創24個月新高", "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)", "夜星剛形成"],
    ["地量(≤0.5倍均量)", "連續放量(近3日均量≥1.5倍前20日均量)", "OBV能量潮創60日新高"],

  ];
  var BT_HOT_STATS = {
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月賣超": {"10": {"n": 167, "moonshotN": 20, "pct": 12.0, "avg": 50.4}, "20": {"n": 167, "moonshotN": 21, "pct": 12.6, "avg": 65.7}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1680, "moonshotN": 170, "pct": 10.1, "avg": 46.2}, "20": {"n": 1680, "moonshotN": 294, "pct": 17.5, "avg": 56.1}},
    "短線過熱(高於月線≥20%) ＋ 月營收連3月年增 ＋ N字底剛形成": {"10": {"n": 157, "moonshotN": 22, "pct": 14.0, "avg": 43.4}, "20": {"n": 157, "moonshotN": 28, "pct": 17.8, "avg": 57.9}},
    "近20日強於大盤≥10% ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 796, "moonshotN": 95, "pct": 11.9, "avg": 46.9}, "20": {"n": 796, "moonshotN": 159, "pct": 20.0, "avg": 55.2}},
    "近3日漲停 ＋ 突破ABC修正下降切線剛形成 ＋ K線橫盤的突破剛形成": {"10": {"n": 220, "moonshotN": 24, "pct": 10.9, "avg": 47.3}, "20": {"n": 220, "moonshotN": 24, "pct": 10.9, "avg": 57.8}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 217, "moonshotN": 33, "pct": 15.2, "avg": 47.5}, "20": {"n": 217, "moonshotN": 52, "pct": 24.0, "avg": 59.2}},
    "地量(≤0.5倍均量) ＋ OBV能量潮創60日新高 ＋ 近3日漲停": {"10": {"n": 172, "moonshotN": 22, "pct": 12.8, "avg": 44.9}, "20": {"n": 172, "moonshotN": 42, "pct": 24.4, "avg": 50.7}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 三大法人近3月賣超": {"10": {"n": 61, "moonshotN": 9, "pct": 14.8, "avg": 49.5}, "20": {"n": 61, "moonshotN": 17, "pct": 27.9, "avg": 49.9}},
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價) ＋ N字底剛形成": {"10": {"n": 138, "moonshotN": 16, "pct": 11.6, "avg": 51.7}, "20": {"n": 138, "moonshotN": 27, "pct": 19.6, "avg": 56.5}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 251, "moonshotN": 31, "pct": 12.4, "avg": 48.5}, "20": {"n": 251, "moonshotN": 48, "pct": 19.1, "avg": 55.0}},
    "地量(≤0.5倍均量) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 近3日漲停": {"10": {"n": 603, "moonshotN": 74, "pct": 12.3, "avg": 47.3}, "20": {"n": 603, "moonshotN": 126, "pct": 20.9, "avg": 58.0}},
    "相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 973, "moonshotN": 108, "pct": 11.1, "avg": 46.6}, "20": {"n": 973, "moonshotN": 174, "pct": 17.9, "avg": 55.4}},
    "多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 685, "moonshotN": 80, "pct": 11.7, "avg": 47.0}, "20": {"n": 685, "moonshotN": 135, "pct": 19.7, "avg": 57.7}},
    "多方力道≥80 ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 431, "moonshotN": 49, "pct": 11.4, "avg": 46.4}, "20": {"n": 431, "moonshotN": 56, "pct": 13.0, "avg": 56.4}},
    "OBV能量潮創60日新高 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 297, "moonshotN": 36, "pct": 12.1, "avg": 47.0}, "20": {"n": 297, "moonshotN": 49, "pct": 16.5, "avg": 52.9}},
    "大盤站上20日均線 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 299, "moonshotN": 36, "pct": 12.0, "avg": 45.7}, "20": {"n": 299, "moonshotN": 53, "pct": 17.7, "avg": 53.5}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 581, "moonshotN": 70, "pct": 12.0, "avg": 47.5}, "20": {"n": 581, "moonshotN": 129, "pct": 22.2, "avg": 57.6}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 近3日漲停": {"10": {"n": 285, "moonshotN": 33, "pct": 11.6, "avg": 46.3}, "20": {"n": 285, "moonshotN": 57, "pct": 20.0, "avg": 56.1}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 月營收創24個月新高": {"10": {"n": 244, "moonshotN": 27, "pct": 11.1, "avg": 47.7}, "20": {"n": 244, "moonshotN": 54, "pct": 22.1, "avg": 56.1}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 249, "moonshotN": 30, "pct": 12.0, "avg": 44.0}, "20": {"n": 249, "moonshotN": 47, "pct": 18.9, "avg": 53.7}},
    "CMF資金流買方佔優(近20日≥0.1) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 341, "moonshotN": 35, "pct": 10.3, "avg": 43.3}, "20": {"n": 341, "moonshotN": 50, "pct": 14.7, "avg": 50.5}},
    "創52週新高 ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 328, "moonshotN": 33, "pct": 10.1, "avg": 48.3}, "20": {"n": 328, "moonshotN": 39, "pct": 11.9, "avg": 56.0}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成": {"10": {"n": 523, "moonshotN": 56, "pct": 10.7, "avg": 47.1}, "20": {"n": 523, "moonshotN": 65, "pct": 12.4, "avg": 55.3}},
    "量能區間高檔(≥90百分位) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 396, "moonshotN": 43, "pct": 10.9, "avg": 45.3}, "20": {"n": 396, "moonshotN": 61, "pct": 15.4, "avg": 51.4}},
    "布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 513, "moonshotN": 55, "pct": 10.7, "avg": 47.0}, "20": {"n": 513, "moonshotN": 100, "pct": 19.5, "avg": 56.3}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 515, "moonshotN": 52, "pct": 10.1, "avg": 46.5}, "20": {"n": 515, "moonshotN": 64, "pct": 12.4, "avg": 55.4}},
    "創52週新高 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 281, "moonshotN": 35, "pct": 12.5, "avg": 46.8}, "20": {"n": 281, "moonshotN": 54, "pct": 19.2, "avg": 53.5}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成": {"10": {"n": 358, "moonshotN": 42, "pct": 11.7, "avg": 45.3}, "20": {"n": 358, "moonshotN": 54, "pct": 15.1, "avg": 51.7}},
    "相對強弱為負(弱於大盤) ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 135, "moonshotN": 18, "pct": 13.3, "avg": 44.0}, "20": {"n": 135, "moonshotN": 16, "pct": 11.9, "avg": 60.2}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 490, "moonshotN": 58, "pct": 11.8, "avg": 48.2}, "20": {"n": 490, "moonshotN": 98, "pct": 20.0, "avg": 58.1}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 176, "moonshotN": 20, "pct": 11.4, "avg": 44.7}, "20": {"n": 176, "moonshotN": 28, "pct": 15.9, "avg": 57.8}},
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 289, "moonshotN": 30, "pct": 10.4, "avg": 45.1}, "20": {"n": 289, "moonshotN": 49, "pct": 17.0, "avg": 52.7}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 298, "moonshotN": 31, "pct": 10.4, "avg": 46.6}, "20": {"n": 298, "moonshotN": 39, "pct": 13.1, "avg": 53.7}},
    "大盤站上20日均線 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 463, "moonshotN": 51, "pct": 11.0, "avg": 45.9}, "20": {"n": 463, "moonshotN": 68, "pct": 14.7, "avg": 52.3}},
    "漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 421, "moonshotN": 43, "pct": 10.2, "avg": 45.6}, "20": {"n": 421, "moonshotN": 59, "pct": 14.0, "avg": 51.3}},
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 314, "moonshotN": 35, "pct": 11.1, "avg": 46.2}, "20": {"n": 314, "moonshotN": 55, "pct": 17.5, "avg": 53.4}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 209, "moonshotN": 29, "pct": 13.9, "avg": 46.5}, "20": {"n": 209, "moonshotN": 34, "pct": 16.3, "avg": 57.1}},
    "短線過熱(高於月線≥20%) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 237, "moonshotN": 29, "pct": 12.2, "avg": 47.2}, "20": {"n": 237, "moonshotN": 41, "pct": 17.3, "avg": 52.3}},
    "均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 408, "moonshotN": 47, "pct": 11.5, "avg": 47.2}, "20": {"n": 408, "moonshotN": 63, "pct": 15.4, "avg": 52.9}},
    "短線過熱(高於月線≥20%) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成": {"10": {"n": 304, "moonshotN": 32, "pct": 10.5, "avg": 45.8}, "20": {"n": 304, "moonshotN": 51, "pct": 16.8, "avg": 53.2}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ 圓弧底剛形成": {"10": {"n": 128, "moonshotN": 13, "pct": 10.2, "avg": 50.4}, "20": {"n": 128, "moonshotN": 17, "pct": 13.3, "avg": 49.5}},
    "地量(≤0.5倍均量) ＋ OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 335, "moonshotN": 34, "pct": 10.1, "avg": 44.2}, "20": {"n": 335, "moonshotN": 69, "pct": 20.6, "avg": 49.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 413, "moonshotN": 44, "pct": 10.7, "avg": 45.0}, "20": {"n": 413, "moonshotN": 72, "pct": 17.4, "avg": 59.1}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"10": {"n": 291, "moonshotN": 36, "pct": 12.4, "avg": 46.0}, "20": {"n": 291, "moonshotN": 58, "pct": 19.9, "avg": 57.2}},
    "多方力道≥80 ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 353, "moonshotN": 37, "pct": 10.5, "avg": 45.3}, "20": {"n": 353, "moonshotN": 55, "pct": 15.6, "avg": 58.2}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 272, "moonshotN": 33, "pct": 12.1, "avg": 45.3}, "20": {"n": 272, "moonshotN": 51, "pct": 18.8, "avg": 53.9}},
    "多方力道≥80 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 315, "moonshotN": 37, "pct": 11.7, "avg": 44.5}, "20": {"n": 315, "moonshotN": 51, "pct": 16.2, "avg": 50.8}},
    "多方力道≥65 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 321, "moonshotN": 34, "pct": 10.6, "avg": 46.0}, "20": {"n": 321, "moonshotN": 55, "pct": 17.1, "avg": 53.2}},
    "漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 306, "moonshotN": 31, "pct": 10.1, "avg": 44.7}, "20": {"n": 306, "moonshotN": 52, "pct": 17.0, "avg": 52.4}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ N字底剛形成": {"10": {"n": 78, "moonshotN": 9, "pct": 11.5, "avg": 39.4}, "20": {"n": 78, "moonshotN": 14, "pct": 17.9, "avg": 50.7}},
    "多方力道≥65 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 418, "moonshotN": 49, "pct": 11.7, "avg": 45.6}, "20": {"n": 418, "moonshotN": 62, "pct": 14.8, "avg": 52.1}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 192, "moonshotN": 22, "pct": 11.5, "avg": 48.0}, "20": {"n": 192, "moonshotN": 35, "pct": 18.2, "avg": 55.1}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 145, "moonshotN": 14, "pct": 9.7, "avg": 44.1}, "20": {"n": 145, "moonshotN": 27, "pct": 18.6, "avg": 51.1}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": {"10": {"n": 528, "moonshotN": 61, "pct": 11.6, "avg": 46.0}, "20": {"n": 528, "moonshotN": 116, "pct": 22.0, "avg": 56.4}},
    "近3日漲停 ＋ 外資近5日買超 ＋ N字底剛形成": {"10": {"n": 373, "moonshotN": 39, "pct": 10.5, "avg": 47.3}, "20": {"n": 373, "moonshotN": 55, "pct": 14.7, "avg": 53.9}},
    "短線過熱(高於月線≥20%) ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1575, "moonshotN": 167, "pct": 10.6, "avg": 47.8}, "20": {"n": 1575, "moonshotN": 291, "pct": 18.5, "avg": 57.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成": {"10": {"n": 116, "moonshotN": 10, "pct": 8.6, "avg": 45.0}, "20": {"n": 116, "moonshotN": 24, "pct": 20.7, "avg": 48.9}},
    "強勢突破盤 ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 302, "moonshotN": 38, "pct": 12.6, "avg": 46.8}, "20": {"n": 302, "moonshotN": 38, "pct": 12.6, "avg": 55.4}},
    "大盤跌破20日均線 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 64, "moonshotN": 4, "pct": 6.2, "avg": 44.1}, "20": {"n": 64, "moonshotN": 12, "pct": 18.8, "avg": 47.0}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 311, "moonshotN": 38, "pct": 12.2, "avg": 46.8}, "20": {"n": 311, "moonshotN": 73, "pct": 23.5, "avg": 55.3}},
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 553, "moonshotN": 64, "pct": 11.6, "avg": 45.0}, "20": {"n": 553, "moonshotN": 114, "pct": 20.6, "avg": 57.3}},
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停": {"10": {"n": 748, "moonshotN": 85, "pct": 11.4, "avg": 46.7}, "20": {"n": 748, "moonshotN": 145, "pct": 19.4, "avg": 56.2}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 143, "moonshotN": 16, "pct": 11.2, "avg": 47.0}, "20": {"n": 143, "moonshotN": 22, "pct": 15.4, "avg": 58.9}},
    "短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 342, "moonshotN": 38, "pct": 11.1, "avg": 45.5}, "20": {"n": 342, "moonshotN": 57, "pct": 16.7, "avg": 53.3}},
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 341, "moonshotN": 38, "pct": 11.1, "avg": 45.5}, "20": {"n": 341, "moonshotN": 57, "pct": 16.7, "avg": 53.3}},
    "大盤跌破20日均線 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 287, "moonshotN": 31, "pct": 10.8, "avg": 45.6}, "20": {"n": 287, "moonshotN": 39, "pct": 13.6, "avg": 61.0}},
    "近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 335, "moonshotN": 35, "pct": 10.4, "avg": 45.0}, "20": {"n": 335, "moonshotN": 55, "pct": 16.4, "avg": 52.5}},
    "近20日強於大盤≥10% ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 423, "moonshotN": 44, "pct": 10.4, "avg": 47.0}, "20": {"n": 423, "moonshotN": 59, "pct": 13.9, "avg": 52.4}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 三大法人近3月買超": {"10": {"n": 1083, "moonshotN": 112, "pct": 10.3, "avg": 46.6}, "20": {"n": 1083, "moonshotN": 189, "pct": 17.5, "avg": 54.7}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 425, "moonshotN": 43, "pct": 10.1, "avg": 45.6}, "20": {"n": 425, "moonshotN": 64, "pct": 15.1, "avg": 58.0}},
    "短線過熱(高於月線≥20%) ＋ 三大法人近3月買超 ＋ N字底剛形成": {"10": {"n": 270, "moonshotN": 31, "pct": 11.5, "avg": 45.4}, "20": {"n": 270, "moonshotN": 49, "pct": 18.1, "avg": 54.8}},
    "地量(≤0.5倍均量) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停": {"10": {"n": 690, "moonshotN": 78, "pct": 11.3, "avg": 45.6}, "20": {"n": 690, "moonshotN": 132, "pct": 19.1, "avg": 55.8}},
    "近3日漲停 ＋ 月營收連3月年增 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 503, "moonshotN": 55, "pct": 10.9, "avg": 44.7}, "20": {"n": 503, "moonshotN": 71, "pct": 14.1, "avg": 53.5}},
    "布林通道高檔(≥80%) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 521, "moonshotN": 54, "pct": 10.4, "avg": 46.3}, "20": {"n": 521, "moonshotN": 72, "pct": 13.8, "avg": 52.0}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1669, "moonshotN": 149, "pct": 8.9, "avg": 45.2}, "20": {"n": 1669, "moonshotN": 287, "pct": 17.2, "avg": 55.1}},
    "強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 月營收創24個月新高": {"10": {"n": 203, "moonshotN": 25, "pct": 12.3, "avg": 49.9}, "20": {"n": 203, "moonshotN": 49, "pct": 24.1, "avg": 59.3}},
    "創52週新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 65, "moonshotN": 6, "pct": 9.2, "avg": 49.6}, "20": {"n": 65, "moonshotN": 10, "pct": 15.4, "avg": 66.8}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 320, "moonshotN": 39, "pct": 12.2, "avg": 46.7}, "20": {"n": 320, "moonshotN": 55, "pct": 17.2, "avg": 54.0}},
    "量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 232, "moonshotN": 25, "pct": 10.8, "avg": 45.8}, "20": {"n": 232, "moonshotN": 41, "pct": 17.7, "avg": 52.7}},
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 262, "moonshotN": 28, "pct": 10.7, "avg": 44.3}, "20": {"n": 262, "moonshotN": 44, "pct": 16.8, "avg": 52.5}},
    "地量(≤0.5倍均量) ＋ 距52週高點≤5% ＋ 近3日漲停": {"10": {"n": 355, "moonshotN": 40, "pct": 11.3, "avg": 45.3}, "20": {"n": 355, "moonshotN": 73, "pct": 20.6, "avg": 54.9}},
    "大盤站上60日均線 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 304, "moonshotN": 36, "pct": 11.8, "avg": 45.2}, "20": {"n": 304, "moonshotN": 55, "pct": 18.1, "avg": 53.1}},
    "大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 224, "moonshotN": 15, "pct": 6.7, "avg": 37.3}, "20": {"n": 224, "moonshotN": 36, "pct": 16.1, "avg": 51.4}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 433, "moonshotN": 58, "pct": 13.4, "avg": 48.7}, "20": {"n": 433, "moonshotN": 100, "pct": 23.1, "avg": 56.6}},
    "地量(≤0.5倍均量) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成": {"10": {"n": 66, "moonshotN": 6, "pct": 9.1, "avg": 37.7}, "20": {"n": 66, "moonshotN": 13, "pct": 19.7, "avg": 60.1}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 創52週新高": {"10": {"n": 59, "moonshotN": 3, "pct": 5.1, "avg": 30.5}, "20": {"n": 59, "moonshotN": 11, "pct": 18.6, "avg": 44.4}},
    "量能區間低檔(≤10百分位) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 60, "moonshotN": 3, "pct": 5.0, "avg": 37.5}, "20": {"n": 60, "moonshotN": 11, "pct": 18.3, "avg": 42.3}},
    "距52週高點≤5% ＋ 月營收創24個月新高 ＋ 晨星剛形成": {"10": {"n": 101, "moonshotN": 8, "pct": 7.9, "avg": 39.4}, "20": {"n": 101, "moonshotN": 18, "pct": 17.8, "avg": 45.6}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 近3日漲停 ＋ N字底剛形成": {"10": {"n": 177, "moonshotN": 18, "pct": 10.2, "avg": 46.4}, "20": {"n": 177, "moonshotN": 22, "pct": 12.4, "avg": 46.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 425, "moonshotN": 44, "pct": 10.4, "avg": 48.1}, "20": {"n": 425, "moonshotN": 77, "pct": 18.1, "avg": 59.4}},
    "均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 135, "moonshotN": 13, "pct": 9.6, "avg": 45.9}, "20": {"n": 135, "moonshotN": 19, "pct": 14.1, "avg": 61.2}},
    "強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 363, "moonshotN": 37, "pct": 10.2, "avg": 46.0}, "20": {"n": 363, "moonshotN": 55, "pct": 15.2, "avg": 60.4}},
    "近3日漲停 ＋ 月營收連3月年增 ＋ N字底剛形成": {"10": {"n": 244, "moonshotN": 25, "pct": 10.2, "avg": 44.8}, "20": {"n": 244, "moonshotN": 34, "pct": 13.9, "avg": 51.2}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ K線橫盤的突破剛形成": {"10": {"n": 78, "moonshotN": 9, "pct": 11.5, "avg": 44.6}, "20": {"n": 78, "moonshotN": 13, "pct": 16.7, "avg": 56.1}},
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 876, "moonshotN": 83, "pct": 9.5, "avg": 46.9}, "20": {"n": 876, "moonshotN": 154, "pct": 17.6, "avg": 57.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 距52週高點≤5% ＋ 晨星剛形成": {"10": {"n": 55, "moonshotN": 8, "pct": 14.5, "avg": 44.7}, "20": {"n": 55, "moonshotN": 12, "pct": 21.8, "avg": 46.2}},
    "地量(≤0.5倍均量) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成": {"10": {"n": 68, "moonshotN": 7, "pct": 10.3, "avg": 38.3}, "20": {"n": 68, "moonshotN": 14, "pct": 20.6, "avg": 58.9}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1531, "moonshotN": 139, "pct": 9.1, "avg": 46.4}, "20": {"n": 1531, "moonshotN": 266, "pct": 17.4, "avg": 56.1}},
    "多方力道≥65 ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 111, "moonshotN": 10, "pct": 9.0, "avg": 48.2}, "20": {"n": 111, "moonshotN": 18, "pct": 16.2, "avg": 62.5}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ OBV能量潮創60日新高": {"10": {"n": 109, "moonshotN": 8, "pct": 7.3, "avg": 44.5}, "20": {"n": 109, "moonshotN": 15, "pct": 13.8, "avg": 55.5}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 116, "moonshotN": 12, "pct": 10.3, "avg": 43.5}, "20": {"n": 116, "moonshotN": 17, "pct": 14.7, "avg": 55.4}},
    "KDJ近3日內死亡交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 671, "moonshotN": 59, "pct": 8.8, "avg": 46.0}, "20": {"n": 671, "moonshotN": 104, "pct": 15.5, "avg": 53.9}},
    "近3日向上跳空缺口 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 274, "moonshotN": 25, "pct": 9.1, "avg": 43.4}, "20": {"n": 274, "moonshotN": 42, "pct": 15.3, "avg": 50.1}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成": {"10": {"n": 162, "moonshotN": 14, "pct": 8.6, "avg": 44.1}, "20": {"n": 162, "moonshotN": 29, "pct": 17.9, "avg": 49.8}},
    "創52週新高 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 276, "moonshotN": 27, "pct": 9.8, "avg": 45.9}, "20": {"n": 276, "moonshotN": 48, "pct": 17.4, "avg": 54.2}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 591, "moonshotN": 43, "pct": 7.3, "avg": 44.4}, "20": {"n": 591, "moonshotN": 82, "pct": 13.9, "avg": 52.9}},
    "地量(≤0.5倍均量) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 夜星剛形成": {"10": {"n": 56, "moonshotN": 11, "pct": 19.6, "avg": 39.8}, "20": {"n": 56, "moonshotN": 12, "pct": 21.4, "avg": 68.3}},
    "量能區間低檔(≤10百分位) ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 55, "moonshotN": 3, "pct": 5.5, "avg": 37.5}, "20": {"n": 55, "moonshotN": 11, "pct": 20.0, "avg": 41.2}},
    "多方力道≥65 ＋ 量能區間低檔(≤10百分位) ＋ 創52週新高": {"10": {"n": 60, "moonshotN": 2, "pct": 3.3, "avg": 30.8}, "20": {"n": 60, "moonshotN": 11, "pct": 18.3, "avg": 44.4}},
    "量能區間低檔(≤10百分位) ＋ 大盤站上60日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 67, "moonshotN": 3, "pct": 4.5, "avg": 37.5}, "20": {"n": 67, "moonshotN": 12, "pct": 17.9, "avg": 41.7}},
    "量能區間低檔(≤10百分位) ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 56, "moonshotN": 1, "pct": 1.8, "avg": 30.9}, "20": {"n": 56, "moonshotN": 10, "pct": 17.9, "avg": 42.9}},
    "量能區間低檔(≤10百分位) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 69, "moonshotN": 3, "pct": 4.3, "avg": 37.5}, "20": {"n": 69, "moonshotN": 12, "pct": 17.4, "avg": 41.7}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為正": {"10": {"n": 69, "moonshotN": 3, "pct": 4.3, "avg": 37.5}, "20": {"n": 69, "moonshotN": 12, "pct": 17.4, "avg": 41.7}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月買超": {"10": {"n": 67, "moonshotN": 3, "pct": 4.5, "avg": 37.5}, "20": {"n": 67, "moonshotN": 11, "pct": 16.4, "avg": 40.1}},
    "地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位) ＋ 近3日漲停": {"10": {"n": 134, "moonshotN": 13, "pct": 9.7, "avg": 45.6}, "20": {"n": 134, "moonshotN": 20, "pct": 14.9, "avg": 50.3}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 距52週高點≤5%": {"10": {"n": 193, "moonshotN": 5, "pct": 2.6, "avg": 34.4}, "20": {"n": 193, "moonshotN": 27, "pct": 14.0, "avg": 42.2}},
    "近3月均價YoY為正 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 178, "moonshotN": 14, "pct": 7.9, "avg": 45.0}, "20": {"n": 178, "moonshotN": 25, "pct": 14.0, "avg": 50.6}},
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1342, "moonshotN": 107, "pct": 8.0, "avg": 47.4}, "20": {"n": 1342, "moonshotN": 222, "pct": 16.5, "avg": 56.8}},
    "MACD近3日內黃金交叉 ＋ 連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 596, "moonshotN": 51, "pct": 8.6, "avg": 45.2}, "20": {"n": 596, "moonshotN": 93, "pct": 15.6, "avg": 54.2}},
    "月營收連3月年增 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 180, "moonshotN": 16, "pct": 8.9, "avg": 44.5}, "20": {"n": 180, "moonshotN": 26, "pct": 14.4, "avg": 52.2}},
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 月營收創24個月新高": {"10": {"n": 1247, "moonshotN": 136, "pct": 10.9, "avg": 48.4}, "20": {"n": 1247, "moonshotN": 252, "pct": 20.2, "avg": 54.2}},
    "大盤跌破60日均線 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成": {"10": {"n": 83, "moonshotN": 5, "pct": 6.0, "avg": 41.3}, "20": {"n": 83, "moonshotN": 11, "pct": 13.3, "avg": 44.2}},
    "近20日強於大盤≥10% ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"10": {"n": 186, "moonshotN": 16, "pct": 8.6, "avg": 43.3}, "20": {"n": 186, "moonshotN": 33, "pct": 17.7, "avg": 50.2}},
    "創52週新高 ＋ 近3月均價YoY為負 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 213, "moonshotN": 15, "pct": 7.0, "avg": 45.4}, "20": {"n": 213, "moonshotN": 31, "pct": 14.6, "avg": 51.8}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 654, "moonshotN": 50, "pct": 7.6, "avg": 41.3}, "20": {"n": 654, "moonshotN": 95, "pct": 14.5, "avg": 52.9}},
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 月營收連3月年增": {"10": {"n": 759, "moonshotN": 85, "pct": 11.2, "avg": 44.3}, "20": {"n": 759, "moonshotN": 135, "pct": 17.8, "avg": 51.1}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1175, "moonshotN": 107, "pct": 9.1, "avg": 46.9}, "20": {"n": 1175, "moonshotN": 198, "pct": 16.9, "avg": 56.7}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 239, "moonshotN": 21, "pct": 8.8, "avg": 44.3}, "20": {"n": 239, "moonshotN": 33, "pct": 13.8, "avg": 53.1}},
    "多方力道≥65 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2364, "moonshotN": 208, "pct": 8.8, "avg": 46.4}, "20": {"n": 2364, "moonshotN": 373, "pct": 15.8, "avg": 55.5}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 87, "moonshotN": 8, "pct": 9.2, "avg": 36.7}, "20": {"n": 87, "moonshotN": 11, "pct": 12.6, "avg": 44.9}},
    "近3日漲停 ＋ 月營收創24個月新高 ＋ N字底剛形成": {"10": {"n": 118, "moonshotN": 12, "pct": 10.2, "avg": 40.8}, "20": {"n": 118, "moonshotN": 18, "pct": 15.3, "avg": 46.6}},
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2351, "moonshotN": 213, "pct": 9.1, "avg": 46.6}, "20": {"n": 2351, "moonshotN": 388, "pct": 16.5, "avg": 55.7}},
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1965, "moonshotN": 161, "pct": 8.2, "avg": 46.8}, "20": {"n": 1965, "moonshotN": 313, "pct": 15.9, "avg": 55.4}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 晨星剛形成": {"10": {"n": 204, "moonshotN": 17, "pct": 8.3, "avg": 45.0}, "20": {"n": 204, "moonshotN": 32, "pct": 15.7, "avg": 53.5}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1223, "moonshotN": 114, "pct": 9.3, "avg": 46.4}, "20": {"n": 1223, "moonshotN": 213, "pct": 17.4, "avg": 55.2}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 656, "moonshotN": 69, "pct": 10.5, "avg": 44.8}, "20": {"n": 656, "moonshotN": 114, "pct": 17.4, "avg": 54.8}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 近3日漲停": {"10": {"n": 812, "moonshotN": 71, "pct": 8.7, "avg": 44.5}, "20": {"n": 812, "moonshotN": 139, "pct": 17.1, "avg": 53.9}},
    "布林通道高檔(≥80%) ＋ 月營收創24個月新高 ＋ 晨星剛形成": {"10": {"n": 173, "moonshotN": 11, "pct": 6.4, "avg": 38.2}, "20": {"n": 173, "moonshotN": 27, "pct": 15.6, "avg": 44.5}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 2065, "moonshotN": 172, "pct": 8.3, "avg": 46.5}, "20": {"n": 2065, "moonshotN": 337, "pct": 16.3, "avg": 54.7}},
    "短線過熱(高於月線≥20%) ＋ 外資近5日買超 ＋ N字底剛形成": {"10": {"n": 222, "moonshotN": 27, "pct": 12.2, "avg": 46.6}, "20": {"n": 222, "moonshotN": 41, "pct": 18.5, "avg": 56.4}},
    "MACD近3日內黃金交叉 ＋ OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 628, "moonshotN": 50, "pct": 8.0, "avg": 45.2}, "20": {"n": 628, "moonshotN": 89, "pct": 14.2, "avg": 54.1}},
    "MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 492, "moonshotN": 39, "pct": 7.9, "avg": 44.7}, "20": {"n": 492, "moonshotN": 67, "pct": 13.6, "avg": 54.6}},
    "大盤站上20日均線 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 88, "moonshotN": 6, "pct": 6.8, "avg": 41.6}, "20": {"n": 88, "moonshotN": 12, "pct": 13.6, "avg": 57.2}},
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ N字底剛形成": {"10": {"n": 112, "moonshotN": 10, "pct": 8.9, "avg": 51.4}, "20": {"n": 112, "moonshotN": 22, "pct": 19.6, "avg": 49.9}},
    "創52週新高 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1621, "moonshotN": 155, "pct": 9.6, "avg": 48.1}, "20": {"n": 1621, "moonshotN": 279, "pct": 17.2, "avg": 57.2}},
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1847, "moonshotN": 160, "pct": 8.7, "avg": 48.2}, "20": {"n": 1847, "moonshotN": 307, "pct": 16.6, "avg": 55.8}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 月營收連3月年增": {"10": {"n": 376, "moonshotN": 35, "pct": 9.3, "avg": 45.9}, "20": {"n": 376, "moonshotN": 55, "pct": 14.6, "avg": 52.4}},
    "近20日強於大盤≥10% ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 126, "moonshotN": 10, "pct": 7.9, "avg": 43.7}, "20": {"n": 126, "moonshotN": 17, "pct": 13.5, "avg": 45.3}},
    "多方力道≥65 ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 599, "moonshotN": 50, "pct": 8.3, "avg": 43.1}, "20": {"n": 599, "moonshotN": 80, "pct": 13.4, "avg": 54.6}},
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2465, "moonshotN": 207, "pct": 8.4, "avg": 46.9}, "20": {"n": 2465, "moonshotN": 399, "pct": 16.2, "avg": 54.8}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 1785, "moonshotN": 131, "pct": 7.3, "avg": 43.6}, "20": {"n": 1785, "moonshotN": 274, "pct": 15.4, "avg": 50.7}},
    "短線過熱(高於月線≥20%) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 306, "moonshotN": 23, "pct": 7.5, "avg": 46.8}, "20": {"n": 306, "moonshotN": 46, "pct": 15.0, "avg": 51.4}},
    "地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 184, "moonshotN": 12, "pct": 6.5, "avg": 40.0}, "20": {"n": 184, "moonshotN": 28, "pct": 15.2, "avg": 45.2}},
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 294, "moonshotN": 22, "pct": 7.5, "avg": 45.3}, "20": {"n": 294, "moonshotN": 41, "pct": 13.9, "avg": 52.2}},
    "月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 226, "moonshotN": 16, "pct": 7.1, "avg": 44.5}, "20": {"n": 226, "moonshotN": 28, "pct": 12.4, "avg": 51.0}},
    "短線過熱(高於月線≥20%) ＋ 三大法人近3月買超 ＋ 月營收創24個月新高": {"10": {"n": 2127, "moonshotN": 181, "pct": 8.5, "avg": 46.2}, "20": {"n": 2127, "moonshotN": 368, "pct": 17.3, "avg": 53.8}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1452, "moonshotN": 128, "pct": 8.8, "avg": 48.0}, "20": {"n": 1452, "moonshotN": 243, "pct": 16.7, "avg": 55.4}},
    "近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2499, "moonshotN": 213, "pct": 8.5, "avg": 46.2}, "20": {"n": 2499, "moonshotN": 401, "pct": 16.0, "avg": 54.8}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月均價YoY為負": {"10": {"n": 2077, "moonshotN": 173, "pct": 8.3, "avg": 45.7}, "20": {"n": 2077, "moonshotN": 292, "pct": 14.1, "avg": 53.6}},
    "近20日強於大盤≥10% ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 132, "moonshotN": 13, "pct": 9.8, "avg": 46.7}, "20": {"n": 132, "moonshotN": 16, "pct": 12.1, "avg": 60.9}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 58, "moonshotN": 3, "pct": 5.2, "avg": 37.5}, "20": {"n": 58, "moonshotN": 12, "pct": 20.7, "avg": 41.7}},
    "布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 66, "moonshotN": 3, "pct": 4.5, "avg": 37.5}, "20": {"n": 66, "moonshotN": 12, "pct": 18.2, "avg": 41.7}},
    "大盤跌破20日均線 ＋ 月營收連3月年增 ＋ 夜星剛形成": {"10": {"n": 158, "moonshotN": 13, "pct": 8.2, "avg": 40.3}, "20": {"n": 158, "moonshotN": 28, "pct": 17.7, "avg": 53.3}},
    "近20日強於大盤≥10% ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"10": {"n": 63, "moonshotN": 11, "pct": 17.5, "avg": 39.8}, "20": {"n": 63, "moonshotN": 11, "pct": 17.5, "avg": 68.2}},
    "多方力道≥65 ＋ 量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 70, "moonshotN": 2, "pct": 2.9, "avg": 39.6}, "20": {"n": 70, "moonshotN": 12, "pct": 17.1, "avg": 41.7}},
    "近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 71, "moonshotN": 2, "pct": 2.8, "avg": 39.6}, "20": {"n": 71, "moonshotN": 12, "pct": 16.9, "avg": 41.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 月營收創24個月新高 ＋ 晨星剛形成": {"10": {"n": 71, "moonshotN": 4, "pct": 5.6, "avg": 40.4}, "20": {"n": 71, "moonshotN": 12, "pct": 16.9, "avg": 49.7}},
    "量能區間低檔(≤10百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 72, "moonshotN": 3, "pct": 4.2, "avg": 37.5}, "20": {"n": 72, "moonshotN": 12, "pct": 16.7, "avg": 41.7}},
    "均線多頭排列(5>20>60且站上月線) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 152, "moonshotN": 14, "pct": 9.2, "avg": 43.2}, "20": {"n": 152, "moonshotN": 25, "pct": 16.4, "avg": 47.5}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2576, "moonshotN": 223, "pct": 8.7, "avg": 46.6}, "20": {"n": 2576, "moonshotN": 421, "pct": 16.3, "avg": 54.9}},
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": {"10": {"n": 86, "moonshotN": 8, "pct": 9.3, "avg": 37.2}, "20": {"n": 86, "moonshotN": 14, "pct": 16.3, "avg": 58.7}},
    "相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2567, "moonshotN": 221, "pct": 8.6, "avg": 46.2}, "20": {"n": 2567, "moonshotN": 417, "pct": 16.2, "avg": 54.6}},
    "量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 715, "moonshotN": 52, "pct": 7.3, "avg": 44.1}, "20": {"n": 715, "moonshotN": 113, "pct": 15.8, "avg": 52.1}},
    "大盤站上60日均線 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 96, "moonshotN": 8, "pct": 8.3, "avg": 47.2}, "20": {"n": 96, "moonshotN": 15, "pct": 15.6, "avg": 61.8}},
    "創52週新高 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 91, "moonshotN": 8, "pct": 8.8, "avg": 47.2}, "20": {"n": 91, "moonshotN": 14, "pct": 15.4, "avg": 63.8}},
    "多方力道≥65 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 80, "moonshotN": 8, "pct": 10.0, "avg": 47.2}, "20": {"n": 80, "moonshotN": 12, "pct": 15.0, "avg": 66.5}},
    "距52週高點≤5% ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 414, "moonshotN": 31, "pct": 7.5, "avg": 46.3}, "20": {"n": 414, "moonshotN": 61, "pct": 14.7, "avg": 52.4}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 1004, "moonshotN": 66, "pct": 6.6, "avg": 46.5}, "20": {"n": 1004, "moonshotN": 142, "pct": 14.1, "avg": 48.2}},
    "三大法人近3月買超 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 135, "moonshotN": 11, "pct": 8.1, "avg": 45.4}, "20": {"n": 135, "moonshotN": 19, "pct": 14.1, "avg": 54.4}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 325, "moonshotN": 25, "pct": 7.7, "avg": 40.9}, "20": {"n": 325, "moonshotN": 44, "pct": 13.5, "avg": 53.1}},
    "強勢突破盤 ＋ 月營收連3月年增 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 112, "moonshotN": 7, "pct": 6.2, "avg": 48.2}, "20": {"n": 112, "moonshotN": 15, "pct": 13.4, "avg": 47.6}},
    "創52週新高 ＋ 近3日漲停 ＋ 三大法人近3月賣超": {"10": {"n": 1259, "moonshotN": 115, "pct": 9.1, "avg": 44.8}, "20": {"n": 1259, "moonshotN": 161, "pct": 12.8, "avg": 54.3}},
    "MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 323, "moonshotN": 22, "pct": 6.8, "avg": 39.8}, "20": {"n": 323, "moonshotN": 40, "pct": 12.4, "avg": 52.5}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 三大法人近3月賣超": {"10": {"n": 1065, "moonshotN": 86, "pct": 8.1, "avg": 44.7}, "20": {"n": 1065, "moonshotN": 128, "pct": 12.0, "avg": 52.8}},
    "創52週新高 ＋ 近3日漲停 ＋ 近3月均價YoY為負": {"10": {"n": 2026, "moonshotN": 162, "pct": 8.0, "avg": 46.0}, "20": {"n": 2026, "moonshotN": 279, "pct": 13.8, "avg": 53.1}},
    "漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 2260, "moonshotN": 192, "pct": 8.5, "avg": 47.5}, "20": {"n": 2260, "moonshotN": 364, "pct": 16.1, "avg": 56.4}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日向上跳空缺口 ＋ 近3日漲停": {"10": {"n": 912, "moonshotN": 71, "pct": 7.8, "avg": 47.3}, "20": {"n": 912, "moonshotN": 144, "pct": 15.8, "avg": 56.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": {"10": {"n": 206, "moonshotN": 16, "pct": 7.8, "avg": 45.8}, "20": {"n": 206, "moonshotN": 32, "pct": 15.5, "avg": 50.1}},
    "短線過熱(高於月線≥20%) ＋ 三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 360, "moonshotN": 31, "pct": 8.6, "avg": 44.3}, "20": {"n": 360, "moonshotN": 52, "pct": 14.4, "avg": 49.2}},
    "相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 543, "moonshotN": 36, "pct": 6.6, "avg": 44.5}, "20": {"n": 543, "moonshotN": 74, "pct": 13.6, "avg": 51.7}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": {"10": {"n": 470, "moonshotN": 54, "pct": 11.5, "avg": 46.6}, "20": {"n": 470, "moonshotN": 98, "pct": 20.9, "avg": 55.8}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1475, "moonshotN": 139, "pct": 9.4, "avg": 47.3}, "20": {"n": 1475, "moonshotN": 257, "pct": 17.4, "avg": 57.7}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ 突破飆股大量黑K最高點剛形成": {"10": {"n": 184, "moonshotN": 16, "pct": 8.7, "avg": 48.6}, "20": {"n": 184, "moonshotN": 29, "pct": 15.8, "avg": 55.4}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": {"10": {"n": 795, "moonshotN": 67, "pct": 8.4, "avg": 45.3}, "20": {"n": 795, "moonshotN": 115, "pct": 14.5, "avg": 55.0}},
    "大盤站上60日均線 ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 201, "moonshotN": 16, "pct": 8.0, "avg": 47.6}, "20": {"n": 201, "moonshotN": 29, "pct": 14.4, "avg": 58.7}},
    "相對強弱為正(強於大盤) ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 154, "moonshotN": 14, "pct": 9.1, "avg": 44.8}, "20": {"n": 154, "moonshotN": 21, "pct": 13.6, "avg": 54.4}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 創52週新高": {"10": {"n": 501, "moonshotN": 42, "pct": 8.4, "avg": 43.1}, "20": {"n": 501, "moonshotN": 93, "pct": 18.6, "avg": 53.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 距52週高點≤5% ＋ 近3日漲停": {"10": {"n": 474, "moonshotN": 44, "pct": 9.3, "avg": 43.4}, "20": {"n": 474, "moonshotN": 82, "pct": 17.3, "avg": 54.0}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 三大法人近3月買超": {"10": {"n": 1317, "moonshotN": 114, "pct": 8.7, "avg": 46.7}, "20": {"n": 1317, "moonshotN": 216, "pct": 16.4, "avg": 57.2}},
    "布林通道高檔(≥80%) ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 102, "moonshotN": 8, "pct": 7.8, "avg": 47.2}, "20": {"n": 102, "moonshotN": 15, "pct": 14.7, "avg": 61.8}},
    "MACD近3日內黃金交叉 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 682, "moonshotN": 56, "pct": 8.2, "avg": 45.9}, "20": {"n": 682, "moonshotN": 97, "pct": 14.2, "avg": 56.0}},
    "MACD近3日內黃金交叉 ＋ 近3日向上跳空缺口 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 684, "moonshotN": 46, "pct": 6.7, "avg": 43.3}, "20": {"n": 684, "moonshotN": 92, "pct": 13.5, "avg": 51.9}},
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 937, "moonshotN": 69, "pct": 7.4, "avg": 44.6}, "20": {"n": 937, "moonshotN": 127, "pct": 13.6, "avg": 51.1}},
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 98, "moonshotN": 8, "pct": 8.2, "avg": 46.9}, "20": {"n": 98, "moonshotN": 13, "pct": 13.3, "avg": 51.5}},
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2340, "moonshotN": 168, "pct": 7.2, "avg": 44.3}, "20": {"n": 2340, "moonshotN": 284, "pct": 12.1, "avg": 53.3}},
    "近3日向上跳空缺口 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 140, "moonshotN": 20, "pct": 14.3, "avg": 43.6}, "20": {"n": 140, "moonshotN": 23, "pct": 16.4, "avg": 55.3}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 外資近5日買超": {"10": {"n": 1288, "moonshotN": 112, "pct": 8.7, "avg": 45.8}, "20": {"n": 1288, "moonshotN": 213, "pct": 16.5, "avg": 55.0}},
    "大盤跌破60日均線 ＋ 外資近5日買超 ＋ 夜星剛形成": {"10": {"n": 93, "moonshotN": 8, "pct": 8.6, "avg": 44.8}, "20": {"n": 93, "moonshotN": 15, "pct": 16.1, "avg": 57.1}},
    "KDJ近3日內死亡交叉 ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 944, "moonshotN": 73, "pct": 7.7, "avg": 44.1}, "20": {"n": 944, "moonshotN": 134, "pct": 14.2, "avg": 51.0}},
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2284, "moonshotN": 146, "pct": 6.4, "avg": 44.2}, "20": {"n": 2284, "moonshotN": 299, "pct": 13.1, "avg": 50.8}},
    "近20日強於大盤≥10% ＋ 創52週新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 685, "moonshotN": 49, "pct": 7.2, "avg": 45.7}, "20": {"n": 685, "moonshotN": 83, "pct": 12.1, "avg": 52.9}},
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2088, "moonshotN": 130, "pct": 6.2, "avg": 43.7}, "20": {"n": 2088, "moonshotN": 252, "pct": 12.1, "avg": 51.3}},
    "近3日漲停 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成": {"10": {"n": 71, "moonshotN": 9, "pct": 12.7, "avg": 38.6}, "20": {"n": 71, "moonshotN": 13, "pct": 18.3, "avg": 48.0}},
    "短線過熱(高於月線≥20%) ＋ 月營收創24個月新高 ＋ 外資近5日買超": {"10": {"n": 1774, "moonshotN": 154, "pct": 8.7, "avg": 46.4}, "20": {"n": 1774, "moonshotN": 299, "pct": 16.9, "avg": 55.5}},
    "創52週新高 ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 420, "moonshotN": 41, "pct": 9.8, "avg": 45.9}, "20": {"n": 420, "moonshotN": 62, "pct": 14.8, "avg": 55.5}},
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 697, "moonshotN": 58, "pct": 8.3, "avg": 45.2}, "20": {"n": 697, "moonshotN": 100, "pct": 14.3, "avg": 54.6}},
    "MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 823, "moonshotN": 67, "pct": 8.1, "avg": 43.9}, "20": {"n": 823, "moonshotN": 116, "pct": 14.1, "avg": 53.9}},
    "MACD近3日內黃金交叉 ＋ 近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 763, "moonshotN": 63, "pct": 8.3, "avg": 44.4}, "20": {"n": 763, "moonshotN": 107, "pct": 14.0, "avg": 54.1}},
    "創52週新高 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2850, "moonshotN": 241, "pct": 8.5, "avg": 45.5}, "20": {"n": 2850, "moonshotN": 371, "pct": 13.0, "avg": 53.8}},
    "KDJ近3日內死亡交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 577, "moonshotN": 39, "pct": 6.8, "avg": 47.9}, "20": {"n": 577, "moonshotN": 69, "pct": 12.0, "avg": 54.1}},
    "多方力道≥80 ＋ 創52週新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 489, "moonshotN": 37, "pct": 7.6, "avg": 44.1}, "20": {"n": 489, "moonshotN": 57, "pct": 11.7, "avg": 54.0}},
    "近20日強於大盤≥10% ＋ 近3日向上跳空缺口 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 100, "moonshotN": 11, "pct": 11.0, "avg": 41.7}, "20": {"n": 100, "moonshotN": 17, "pct": 17.0, "avg": 58.7}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月賣超": {"10": {"n": 200, "moonshotN": 18, "pct": 9.0, "avg": 42.4}, "20": {"n": 200, "moonshotN": 37, "pct": 18.5, "avg": 53.2}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成": {"10": {"n": 158, "moonshotN": 11, "pct": 7.0, "avg": 40.8}, "20": {"n": 158, "moonshotN": 25, "pct": 15.8, "avg": 54.2}},
    "MACD近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 593, "moonshotN": 48, "pct": 8.1, "avg": 46.4}, "20": {"n": 593, "moonshotN": 86, "pct": 14.5, "avg": 56.7}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 837, "moonshotN": 68, "pct": 8.1, "avg": 45.1}, "20": {"n": 837, "moonshotN": 118, "pct": 14.1, "avg": 54.8}},
    "布林通道高檔(≥80%) ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 836, "moonshotN": 68, "pct": 8.1, "avg": 45.1}, "20": {"n": 836, "moonshotN": 118, "pct": 14.1, "avg": 54.8}},
    "強勢突破盤 ＋ 月營收連3月年增 ＋ 晨星剛形成": {"10": {"n": 180, "moonshotN": 8, "pct": 4.4, "avg": 42.7}, "20": {"n": 180, "moonshotN": 23, "pct": 12.8, "avg": 45.9}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 創52週新高 ＋ 月營收創24個月新高": {"10": {"n": 348, "moonshotN": 31, "pct": 8.9, "avg": 43.7}, "20": {"n": 348, "moonshotN": 56, "pct": 16.1, "avg": 56.5}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 1958, "moonshotN": 149, "pct": 7.6, "avg": 43.5}, "20": {"n": 1958, "moonshotN": 305, "pct": 15.6, "avg": 51.0}},
    "量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%) ＋ 月營收創24個月新高": {"10": {"n": 1488, "moonshotN": 130, "pct": 8.7, "avg": 46.5}, "20": {"n": 1488, "moonshotN": 229, "pct": 15.4, "avg": 56.2}},
    "爆量(≥2倍均量) ＋ 月營收連3月年增 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 105, "moonshotN": 5, "pct": 4.8, "avg": 48.3}, "20": {"n": 105, "moonshotN": 15, "pct": 14.3, "avg": 49.0}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月賣超": {"10": {"n": 978, "moonshotN": 90, "pct": 9.2, "avg": 45.2}, "20": {"n": 978, "moonshotN": 137, "pct": 14.0, "avg": 54.9}},
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3154, "moonshotN": 201, "pct": 6.4, "avg": 43.6}, "20": {"n": 3154, "moonshotN": 436, "pct": 13.8, "avg": 50.4}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2415, "moonshotN": 156, "pct": 6.5, "avg": 44.6}, "20": {"n": 2415, "moonshotN": 301, "pct": 12.5, "avg": 52.8}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3050, "moonshotN": 179, "pct": 5.9, "avg": 44.3}, "20": {"n": 3050, "moonshotN": 378, "pct": 12.4, "avg": 51.4}},
    "近3日漲停 ＋ N字底剛形成 ＋ K線橫盤的突破剛形成": {"10": {"n": 138, "moonshotN": 13, "pct": 9.4, "avg": 49.4}, "20": {"n": 138, "moonshotN": 17, "pct": 12.3, "avg": 52.5}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停": {"10": {"n": 622, "moonshotN": 56, "pct": 9.0, "avg": 49.1}, "20": {"n": 622, "moonshotN": 114, "pct": 18.3, "avg": 55.6}},
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 三大法人近3月買超": {"10": {"n": 670, "moonshotN": 48, "pct": 7.2, "avg": 42.9}, "20": {"n": 670, "moonshotN": 97, "pct": 14.5, "avg": 52.5}},
    "MACD近3日內黃金交叉 ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 573, "moonshotN": 44, "pct": 7.7, "avg": 49.1}, "20": {"n": 573, "moonshotN": 79, "pct": 13.8, "avg": 56.6}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月均價YoY為負": {"10": {"n": 1787, "moonshotN": 129, "pct": 7.2, "avg": 45.3}, "20": {"n": 1787, "moonshotN": 233, "pct": 13.0, "avg": 53.0}},
    "量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3467, "moonshotN": 209, "pct": 6.0, "avg": 43.4}, "20": {"n": 3467, "moonshotN": 437, "pct": 12.6, "avg": 51.0}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)": {"10": {"n": 514, "moonshotN": 43, "pct": 8.4, "avg": 43.1}, "20": {"n": 514, "moonshotN": 95, "pct": 18.5, "avg": 53.1}},
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 748, "moonshotN": 80, "pct": 10.7, "avg": 46.9}, "20": {"n": 748, "moonshotN": 135, "pct": 18.0, "avg": 54.3}},
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ 近3日漲停": {"10": {"n": 879, "moonshotN": 87, "pct": 9.9, "avg": 45.6}, "20": {"n": 879, "moonshotN": 149, "pct": 17.0, "avg": 55.8}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 近3月均價YoY為正": {"10": {"n": 1373, "moonshotN": 108, "pct": 7.9, "avg": 46.7}, "20": {"n": 1373, "moonshotN": 214, "pct": 15.6, "avg": 56.0}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 近3日漲停": {"10": {"n": 737, "moonshotN": 66, "pct": 9.0, "avg": 47.5}, "20": {"n": 737, "moonshotN": 112, "pct": 15.2, "avg": 53.5}},
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2493, "moonshotN": 194, "pct": 7.8, "avg": 43.8}, "20": {"n": 2493, "moonshotN": 341, "pct": 13.7, "avg": 53.6}},
    "強勢突破盤 ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 433, "moonshotN": 35, "pct": 8.1, "avg": 41.9}, "20": {"n": 433, "moonshotN": 58, "pct": 13.4, "avg": 50.1}},
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3835, "moonshotN": 239, "pct": 6.2, "avg": 43.6}, "20": {"n": 3835, "moonshotN": 503, "pct": 13.1, "avg": 51.0}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2566, "moonshotN": 162, "pct": 6.3, "avg": 43.8}, "20": {"n": 2566, "moonshotN": 316, "pct": 12.3, "avg": 51.3}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 2638, "moonshotN": 157, "pct": 6.0, "avg": 42.7}, "20": {"n": 2638, "moonshotN": 317, "pct": 12.0, "avg": 51.0}},
    "近3日漲停 ＋ 近3月乖離度為正(營收優於股價) ＋ K線橫盤的突破剛形成": {"10": {"n": 596, "moonshotN": 53, "pct": 8.9, "avg": 46.9}, "20": {"n": 596, "moonshotN": 71, "pct": 11.9, "avg": 52.0}},
    "近20日強於大盤≥10% ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"10": {"n": 82, "moonshotN": 8, "pct": 9.8, "avg": 37.8}, "20": {"n": 82, "moonshotN": 16, "pct": 19.5, "avg": 56.3}},
    "距52週高點≤5% ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"10": {"n": 110, "moonshotN": 11, "pct": 10.0, "avg": 41.4}, "20": {"n": 110, "moonshotN": 19, "pct": 17.3, "avg": 45.8}},
    "近20日強於大盤≥10% ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 82, "moonshotN": 8, "pct": 9.8, "avg": 47.2}, "20": {"n": 82, "moonshotN": 14, "pct": 17.1, "avg": 63.8}},
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 月營收創24個月新高": {"10": {"n": 657, "moonshotN": 51, "pct": 7.8, "avg": 44.8}, "20": {"n": 657, "moonshotN": 105, "pct": 16.0, "avg": 53.0}},
    "均線多頭排列(5>20>60且站上月線) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 65, "moonshotN": 4, "pct": 6.2, "avg": 41.2}, "20": {"n": 65, "moonshotN": 10, "pct": 15.4, "avg": 45.0}},
    "相對強弱為正(強於大盤) ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 99, "moonshotN": 8, "pct": 8.1, "avg": 47.2}, "20": {"n": 99, "moonshotN": 15, "pct": 15.2, "avg": 61.8}},
    "創52週新高 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 103, "moonshotN": 8, "pct": 7.8, "avg": 47.2}, "20": {"n": 103, "moonshotN": 15, "pct": 14.6, "avg": 61.8}},
    "地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 208, "moonshotN": 16, "pct": 7.7, "avg": 41.1}, "20": {"n": 208, "moonshotN": 31, "pct": 14.9, "avg": 53.1}},
    "布林通道高檔(≥80%) ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 142, "moonshotN": 12, "pct": 8.5, "avg": 46.6}, "20": {"n": 142, "moonshotN": 20, "pct": 14.1, "avg": 60.3}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 81, "moonshotN": 7, "pct": 8.6, "avg": 44.6}, "20": {"n": 81, "moonshotN": 11, "pct": 13.6, "avg": 56.0}},
    "KDJ近3日內死亡交叉 ＋ 月營收創24個月新高 ＋ 夜星剛形成": {"10": {"n": 145, "moonshotN": 13, "pct": 9.0, "avg": 45.3}, "20": {"n": 145, "moonshotN": 16, "pct": 11.0, "avg": 51.7}},
    "量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3日漲停 ＋ 近3月乖離度為負(股價超前營收)": {"10": {"n": 1120, "moonshotN": 88, "pct": 7.9, "avg": 46.8}, "20": {"n": 1120, "moonshotN": 178, "pct": 15.9, "avg": 57.2}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 1209, "moonshotN": 100, "pct": 8.3, "avg": 48.0}, "20": {"n": 1209, "moonshotN": 169, "pct": 14.0, "avg": 53.7}},
    "爆量(≥2倍均量) ＋ 創52週新高 ＋ 月營收創24個月新高": {"10": {"n": 1356, "moonshotN": 106, "pct": 7.8, "avg": 46.8}, "20": {"n": 1356, "moonshotN": 189, "pct": 13.9, "avg": 57.3}},
    "近3日漲停 ＋ 三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 828, "moonshotN": 81, "pct": 9.8, "avg": 46.8}, "20": {"n": 828, "moonshotN": 113, "pct": 13.6, "avg": 54.9}},
    "創52週新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 559, "moonshotN": 42, "pct": 7.5, "avg": 46.3}, "20": {"n": 559, "moonshotN": 71, "pct": 12.7, "avg": 53.2}},
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2928, "moonshotN": 204, "pct": 7.0, "avg": 45.7}, "20": {"n": 2928, "moonshotN": 362, "pct": 12.4, "avg": 54.2}},
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 371, "moonshotN": 20, "pct": 5.4, "avg": 45.2}, "20": {"n": 371, "moonshotN": 43, "pct": 11.6, "avg": 52.6}},
    "大盤跌破60日均線 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 984, "moonshotN": 62, "pct": 6.3, "avg": 42.6}, "20": {"n": 984, "moonshotN": 108, "pct": 11.0, "avg": 47.8}},
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ OBV能量潮創60日新高": {"10": {"n": 265, "moonshotN": 24, "pct": 9.1, "avg": 43.0}, "20": {"n": 265, "moonshotN": 47, "pct": 17.7, "avg": 50.6}},
    "距52週高點≤5% ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 334, "moonshotN": 27, "pct": 8.1, "avg": 42.5}, "20": {"n": 334, "moonshotN": 48, "pct": 14.4, "avg": 50.8}},
    "近3日向上跳空缺口 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 2745, "moonshotN": 201, "pct": 7.3, "avg": 45.6}, "20": {"n": 2745, "moonshotN": 382, "pct": 13.9, "avg": 52.7}},
    "近3日漲停 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 母子懷抱(低檔)剛形成": {"10": {"n": 79, "moonshotN": 7, "pct": 8.9, "avg": 52.8}, "20": {"n": 79, "moonshotN": 11, "pct": 13.9, "avg": 70.7}},
    "爆量(≥2倍均量) ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 2694, "moonshotN": 203, "pct": 7.5, "avg": 45.5}, "20": {"n": 2694, "moonshotN": 349, "pct": 13.0, "avg": 55.0}},
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2755, "moonshotN": 188, "pct": 6.8, "avg": 44.3}, "20": {"n": 2755, "moonshotN": 341, "pct": 12.4, "avg": 52.1}},
    "連續放量(近3日均量≥1.5倍前20日均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 3514, "moonshotN": 242, "pct": 6.9, "avg": 44.9}, "20": {"n": 3514, "moonshotN": 431, "pct": 12.3, "avg": 53.0}},
    "近20日強於大盤≥10% ＋ 三大法人近3月買超 ＋ 夜星剛形成": {"10": {"n": 341, "moonshotN": 21, "pct": 6.2, "avg": 43.1}, "20": {"n": 341, "moonshotN": 40, "pct": 11.7, "avg": 59.0}},
    "OBV能量潮創60日新高 ＋ 近3日漲停 ＋ 月營收創24個月新高": {"10": {"n": 1489, "moonshotN": 130, "pct": 8.7, "avg": 48.2}, "20": {"n": 1489, "moonshotN": 240, "pct": 16.1, "avg": 58.3}},
    "量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 距52週高點≤5% ＋ 晨星剛形成": {"10": {"n": 180, "moonshotN": 13, "pct": 7.2, "avg": 44.8}, "20": {"n": 180, "moonshotN": 25, "pct": 13.9, "avg": 50.6}},
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3596, "moonshotN": 235, "pct": 6.5, "avg": 42.8}, "20": {"n": 3596, "moonshotN": 483, "pct": 13.4, "avg": 51.6}},
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2956, "moonshotN": 261, "pct": 8.8, "avg": 45.2}, "20": {"n": 2956, "moonshotN": 384, "pct": 13.0, "avg": 54.2}},
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": {"10": {"n": 2956, "moonshotN": 215, "pct": 7.3, "avg": 44.4}, "20": {"n": 2956, "moonshotN": 359, "pct": 12.1, "avg": 53.4}},
    "KDJ近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 1053, "moonshotN": 80, "pct": 7.6, "avg": 45.2}, "20": {"n": 1053, "moonshotN": 126, "pct": 12.0, "avg": 54.0}},
    "近3日漲停 ＋ 近3月乖離度為正(營收優於股價) ＋ 月營收創24個月新高": {"10": {"n": 2305, "moonshotN": 147, "pct": 6.4, "avg": 46.5}, "20": {"n": 2305, "moonshotN": 271, "pct": 11.8, "avg": 53.5}},
    "回後買上漲全通過 ＋ 創52週新高 ＋ 三大法人近3月賣超": {"10": {"n": 613, "moonshotN": 42, "pct": 6.9, "avg": 42.9}, "20": {"n": 613, "moonshotN": 66, "pct": 10.8, "avg": 46.8}},
    "爆量(≥2倍均量) ＋ 大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 774, "moonshotN": 39, "pct": 5.0, "avg": 40.4}, "20": {"n": 774, "moonshotN": 80, "pct": 10.3, "avg": 47.9}},
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ N字底剛形成": {"10": {"n": 283, "moonshotN": 33, "pct": 11.7, "avg": 45.0}, "20": {"n": 283, "moonshotN": 52, "pct": 18.4, "avg": 54.3}},
    "多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ OBV能量潮創60日新高": {"10": {"n": 480, "moonshotN": 41, "pct": 8.5, "avg": 41.0}, "20": {"n": 480, "moonshotN": 82, "pct": 17.1, "avg": 48.7}},
    "多方力道≥80 ＋ CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成": {"10": {"n": 251, "moonshotN": 18, "pct": 7.2, "avg": 42.3}, "20": {"n": 251, "moonshotN": 40, "pct": 15.9, "avg": 52.8}},
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 外資近5日買超": {"10": {"n": 2434, "moonshotN": 169, "pct": 6.9, "avg": 43.7}, "20": {"n": 2434, "moonshotN": 370, "pct": 15.2, "avg": 51.0}},
    "大盤站上20日均線 ＋ 距52週高點≤5% ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 5254, "moonshotN": 409, "pct": 7.8, "avg": 44.5}, "20": {"n": 5254, "moonshotN": 755, "pct": 14.4, "avg": 52.7}},
    "近20日強於大盤≥10% ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": {"10": {"n": 3927, "moonshotN": 254, "pct": 6.5, "avg": 43.5}, "20": {"n": 3927, "moonshotN": 524, "pct": 13.3, "avg": 51.1}},
    "MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 短線過熱(高於月線≥20%)": {"10": {"n": 640, "moonshotN": 47, "pct": 7.3, "avg": 44.8}, "20": {"n": 640, "moonshotN": 84, "pct": 13.1, "avg": 55.0}},
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破飆股大量黑K最高點剛形成": {"10": {"n": 405, "moonshotN": 23, "pct": 5.7, "avg": 50.2}, "20": {"n": 405, "moonshotN": 51, "pct": 12.6, "avg": 53.5}},
    "KDJ近3日內黃金交叉 ＋ 近3日漲停 ＋ 突破ABC修正下降切線剛形成": {"10": {"n": 575, "moonshotN": 49, "pct": 8.5, "avg": 44.2}, "20": {"n": 575, "moonshotN": 70, "pct": 12.2, "avg": 54.5}},
    "月營收創24個月新高 ＋ DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 夜星剛形成": {"10": {"n": 113, "moonshotN": 9, "pct": 8.0, "avg": 47.5}, "20": {"n": 113, "moonshotN": 13, "pct": 11.5, "avg": 57.5}},
    "地量(≤0.5倍均量) ＋ 連續放量(近3日均量≥1.5倍前20日均量) ＋ OBV能量潮創60日新高": {"10": {"n": 65, "moonshotN": 6, "pct": 9.2, "avg": 51.4}, "20": {"n": 65, "moonshotN": 12, "pct": 18.5, "avg": 60.9}},

  };
  var BT_HIT_SHOW = 3;   // 每類最多顯示幾個徽章，其餘以「+N」表示
  function matchedComboList(r, list) {
    var fs = liveFlagSet(r);
    var matched = [];
    list.forEach(function(combo, idx) {
      if (combo.every(function(name) { return fs[name] === 1; })) matched.push(idx);
    });
    return matched;
  }
  function matchedBtWinCombos(r) { return matchedComboList(r, BT_WIN_COMBOS); }
  function matchedBtHotCombos(r) { return matchedComboList(r, BT_HOT_COMBOS); }
  function btWinText(e) { return Object.keys(e || {}).sort(function(a,b){return a-b;}).map(function(h){ return h + '日 勝率' + e[h].win + '% t=' + e[h].t; }).join('　'); }

  function matchedStockPickCombos(r) {
    var fs = liveFlagSet(r);
    var matched = [];
    STOCK_PICK_COMBOS.forEach(function(combo, idx) {
      if (combo.every(function(name) { return fs[name] === 1; })) matched.push(idx);
    });
    return matched;
  }

  // 2026-09-27 台股回測（勝率榜5/10/20日＋飆股榜5/10/20日）中，指定組合的勝率／t值(同日調整)／飆股比例記錄值。
  // key＝條件依字元排序後以「 ＋ 」連接。
  var COMBO_REF_STATS = {
    // ── 2026-10-05 台股三年三段回測（含月營收／外資／漲停等新條件）：S4 換成「漲停＋營收創24個月新高＋外資買超」，新增 S7、S8 營收創新高系列 ──
    "距52週高點≤5% ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 52.7, "10": 54.6, "20": 54.7}, "t": {"5": 9.65, "10": 12.61, "20": 12.27}, "moon": {"10": 5.0, "20": 9.5}},
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 52.3, "10": 53.6, "20": 52.9}, "t": {"5": 7.57, "10": 11.06, "20": 10.09}, "moon": {"10": 6.6, "20": 11.7}},
    "相對強弱為正(強於大盤) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": {"win": {"5": 58.1, "10": 57.8, "20": 56.7}, "t": {"5": 5.44, "10": 6.11, "20": 4.93}, "moon": {"10": 1.2, "20": 4.7}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"win": {"5": 53.8, "10": 55.1, "20": 55.1}, "t": {"5": 5.96, "10": 7.31, "20": 6.41}, "moon": {"10": 3.8, "20": 7.4}},
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 多方力道≥80": {"win": {"5": 49.9, "10": 52.3, "20": 49.2}, "t": {"5": 4.7, "10": 5.41, "20": 4.74}, "moon": {"10": 5.7, "20": 9.9}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 強勢突破盤": {"win": {"20": 60.2}, "t": {"20": 6.52}, "moon": {"10": 10.7, "20": 21.0}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 11.0, "20": 22.0}},
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": {"win": {"5": 55.4, "10": 55.7, "20": 57.5}, "t": {"5": 4.63, "10": 5.41, "20": 5.27}, "moon": {"10": 11.7, "20": 19.3}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3日向上跳空缺口": {"win": {"5": 57.6, "10": 71.1, "20": 72.9}, "t": {"5": 0.2, "10": 4.87, "20": 7.35}, "moon": {"10": 0.7, "20": 4.3}},
    "KDJ近3日內死亡交叉 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成": {"moon": {"20": 10.5}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"moon": {"20": 15.3}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 11.5, "20": 23.1}},
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.0}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ N字底剛形成": {"win": {"20": 60.8}, "t": {"20": 0.66}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線": {"win": {"20": 61.4}, "t": {"20": 0.97}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"win": {"5": 63.6}, "t": {"5": 2.96}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.1}},
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 61.2, "10": 62.3, "20": 62.0}, "t": {"5": 1.87, "10": 1.07, "20": -0.25}},
    "KDJ近3日內死亡交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"win": {"10": 61.1, "20": 63.9}, "t": {"10": 2.25, "20": 2.19}},
    "KDJ近3日內死亡交叉 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 62.2}, "t": {"10": 0.86}},
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.5}},
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 60.2}, "t": {"10": 1.8}},
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 10.3}},
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.8}},
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"moon": {"10": 17.4, "20": 13.0}},
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 10.5}},
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 13.1}},
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.0}},
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 10.3}},
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 14.9, "20": 10.6}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 63.2, "10": 60.9, "20": 62.7}, "t": {"5": 4.59, "10": 3.52, "20": 1.07}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 62.9, "10": 66.3}, "t": {"5": 0.11, "10": -0.01}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 62.1}, "t": {"10": 0.73}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 61.8, "20": 61.7}, "t": {"5": 7.09, "20": 3.76}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 64.9, "10": 63.1, "20": 64.9}, "t": {"5": 0.32, "10": -1.13, "20": -0.69}},
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 60.5}, "t": {"5": 1.78}},
    "KDJ近3日內死亡交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 10.1}},
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.7}},
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.3}},
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.8}},
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 60.6}, "t": {"10": 0.19}},
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.0}},
    "KDJ近3日內死亡交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.8}},
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": {"moon": {"20": 11.7}},
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 10.8, "20": 16.2}},
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成": {"moon": {"20": 13.4}},
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"moon": {"20": 10.0}},
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 61.7}, "t": {"10": 2.28}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"win": {"20": 62.2}, "t": {"20": 1.59}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": {"moon": {"20": 10.5}},
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.0}},
    "KDJ近3日內黃金交叉 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 12.2}},
    "KDJ近3日內黃金交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 13.5}},
    "KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 10.8}},
    "KDJ近3日內黃金交叉 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 10.3}},
    "KDJ近3日內黃金交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"10": 10.3, "20": 14.4}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 13.2}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.8}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.4}},
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 13.3, "20": 13.3}},
    "KDJ近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.2}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"5": 65.4, "10": 65.6, "20": 63.3}, "t": {"5": 1.13, "10": 0.59, "20": 1.67}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"20": 60.1}, "t": {"20": 1.39}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": {"win": {"10": 61.3}, "t": {"10": -0.38}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"10": 60.4, "20": 60.2}, "t": {"10": 1.65, "20": -0.14}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 71.9, "10": 72.8, "20": 66.3}, "t": {"5": 1.5, "10": 0.0, "20": -0.02}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 65.9, "10": 69.9, "20": 67.0}, "t": {"5": 1.41, "10": 2.92, "20": 2.13}},
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 60.7}, "t": {"20": 0.11}},
    "KDJ近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 78.5, "10": 82.5, "20": 74.9}, "t": {"5": -0.27, "10": -0.52, "20": -0.19}},
    "KDJ近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.8}},
    "KDJ近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"win": {"20": 60.4}, "t": {"20": 3.03}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 64.1, "10": 64.3, "20": 61.6}, "t": {"5": 1.54, "10": 0.92, "20": 1.11}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.2}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.4}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 60.4}, "t": {"10": 3.29}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"win": {"10": 60.2}, "t": {"10": 3.48}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 60.1}, "t": {"5": 1.78}},
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.6}},
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.6}},
    "K線橫盤的突破剛形成 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": {"moon": {"10": 11.6, "20": 23.3}},
    "K線橫盤的突破剛形成 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": {"win": {"20": 60.0}, "t": {"20": 1.96}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 晨星剛形成": {"moon": {"20": 12.2}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.3}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"moon": {"20": 15.0}},
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.1}},
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 晨星剛形成": {"moon": {"20": 10.9}},
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.8}},
    "K線橫盤的突破剛形成 ＋ 圓弧底剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.3}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 12.1}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.5}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 11.9}},
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 12.9, "20": 12.9}},
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 14.6}},
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 12.6}},
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"moon": {"20": 17.2}},
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 15.0}},
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 13.3}},
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 13.9}},
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 15.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成": {"moon": {"20": 13.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 13.9}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 12.0}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.3}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.9}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 16.1}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 16.4}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 11.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 14.2}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 10.9, "20": 16.3}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.7}},
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 12.9}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.6}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 11.5}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 11.5}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.3}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 21.9}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 60.0}, "t": {"20": 1.83}, "moon": {"20": 14.0}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 12.8}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.0}},
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.9}},
    "K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.9}},
    "MACD近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": {"moon": {"20": 14.3}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.8}},
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": {"moon": {"20": 12.5}},
    "MACD近3日內黃金交叉 ＋ 三重底剛形成 ＋ 晨星剛形成": {"moon": {"20": 14.3}},
    "MACD近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": {"win": {"10": 61.9}, "t": {"10": -1.37}},
    "MACD近3日內黃金交叉 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.2}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 15.4}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 13.2}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 12.8}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.6}},
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 15.8}},
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.6}},
    "MACD近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"20": 60.7}, "t": {"20": -2.3}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 61.3, "10": 61.3, "20": 61.4}, "t": {"5": 1.09, "10": 0.05, "20": -0.71}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 60.2, "20": 61.9}, "t": {"10": -0.51, "20": 2.04}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能區間低檔(≤10百分位)": {"win": {"5": 62.9, "10": 62.9, "20": 61.4}, "t": {"5": 1.69, "10": 0.91, "20": -0.8}},
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 64.2, "10": 64.7, "20": 63.4}, "t": {"5": -0.17, "10": -1.18, "20": -1.75}},
    "MACD近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 量能區間低檔(≤10百分位)": {"win": {"5": 60.5}, "t": {"5": 1.22}},
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 12.0}},
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.0}},
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.3}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.6}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 12.1}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 11.4}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.9}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.2}},
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 11.3}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 11.2}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.8}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 61.1}, "t": {"10": 1.05}, "moon": {"20": 11.1}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.4}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.9}},
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 13.9}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 三大法人近3月買超": {"win": {"20": 60.6}, "t": {"20": 2.96}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 大盤站上20日均線": {"win": {"20": 60.6}, "t": {"20": 2.35}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為正(強於大盤)": {"win": {"20": 60.2}, "t": {"20": 2.29}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 61.1}, "t": {"10": 1.01}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 60.3, "20": 61.7}, "t": {"10": 2.08, "20": 2.01}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月均價YoY為正": {"win": {"10": 60.8, "20": 65.1}, "t": {"10": 1.82, "20": 2.43}},
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": {"win": {"10": 61.2, "20": 61.2}, "t": {"10": 2.53, "20": 1.6}},
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.2}},
    "N字底剛形成 ＋ 三大法人近3月賣超 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 12.5, "20": 15.0}},
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 12.7}},
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 多方力道≥80": {"moon": {"20": 11.6}},
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 10.3, "20": 10.3}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": {"moon": {"20": 11.1}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": {"moon": {"20": 10.0}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"10": 10.3, "20": 13.8}},
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 61.8}, "t": {"10": -0.01}},
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破ABC修正下降切線剛形成": {"moon": {"5": 11.8, "10": 17.6, "20": 14.7}},
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.4}},
    "N字底剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"10": 10.0}},
    "N字底剛形成 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.5}},
    "N字底剛形成 ＋ 多方力道≥80 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.4}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"10": 10.9}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 10.3, "20": 10.3}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 10.3}},
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.9}},
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"10": 11.1}},
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 10.8, "20": 12.2}},
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.8}},
    "N字底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"10": 12.9}},
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": {"moon": {"5": 10.3, "10": 13.8, "20": 10.3}},
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.5}},
    "N字底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.0}},
    "N字底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 10.5}},
    "N字底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 12.5, "20": 12.5}},
    "N字底剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.4}},
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 11.5}},
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.5}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 16.3}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.5}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": {"moon": {"10": 11.1}},
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"10": 11.6}},
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 13.4}},
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.2}},
    "一字底(均線糾結)剛形成 ＋ 三大法人近3月買超 ＋ 頭肩底剛形成": {"win": {"20": 60.0}, "t": {"20": 1.61}},
    "一字底(均線糾結)剛形成 ＋ 三重底剛形成 ＋ 頭肩底剛形成": {"win": {"5": 60.7, "10": 66.1, "20": 60.7}, "t": {"5": 1.83, "10": 1.73, "20": 0.79}},
    "一字底(均線糾結)剛形成 ＋ 回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 66.1}, "t": {"10": 2.93}},
    "一字底(均線糾結)剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"5": 60.0, "10": 69.1}, "t": {"5": 1.88, "10": 0.82}},
    "一字底(均線糾結)剛形成 ＋ 大盤站上20日均線 ＋ 頭肩底剛形成": {"win": {"10": 61.1, "20": 65.3}, "t": {"10": 2.07, "20": 1.91}},
    "一字底(均線糾結)剛形成 ＋ 大盤站上60日均線 ＋ 頭肩底剛形成": {"win": {"10": 63.1, "20": 63.1}, "t": {"10": 1.99, "20": 1.54}},
    "一字底(均線糾結)剛形成 ＋ 布林通道高檔(≥80%) ＋ 頭肩底剛形成": {"win": {"20": 62.3}, "t": {"20": 1.58}},
    "一字底(均線糾結)剛形成 ＋ 強勢突破盤 ＋ 頭肩底剛形成": {"win": {"10": 61.7, "20": 63.3}, "t": {"10": 2.11, "20": 1.22}},
    "一字底(均線糾結)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成": {"win": {"20": 61.2}, "t": {"20": 1.41}},
    "一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": {"win": {"20": 61.5}, "t": {"20": 1.53}},
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 12.6}},
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 多方力道≥80": {"moon": {"20": 10.8}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 17.5}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 16.2}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.5}},
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 頭肩底剛形成": {"moon": {"20": 13.9}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": {"win": {"10": 62.7}, "t": {"10": 2.44}, "moon": {"10": 16.9, "20": 20.3}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 12.8}},
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": {"moon": {"20": 13.0}},
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 12.9}},
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 夜星剛形成": {"moon": {"20": 10.9}},
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 15.8}},
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 13.9}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 14.2}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.6}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.0}},
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"10": 61.5}, "t": {"10": 2.02}, "moon": {"10": 13.8, "20": 16.9}},
    "三大法人近3月買超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "三大法人近3月買超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 11.1}},
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"5": 62.5, "10": 60.6}, "t": {"5": 1.67, "10": 0.53}},
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": {"win": {"5": 66.0}, "t": {"5": 2.88}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 60.5, "10": 60.5}, "t": {"5": 4.1, "10": 4.99}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 69.3, "10": 69.1, "20": 61.8}, "t": {"5": 2.35, "10": 0.9, "20": 0.51}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 60.5}, "t": {"5": 0.97}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 60.2, "20": 60.2}, "t": {"10": 7.91, "20": 7.06}},
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 72.9, "10": 67.8, "20": 64.4}, "t": {"5": 1.8, "10": -0.36, "20": -0.87}},
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 76.3, "10": 81.3, "20": 77.0}, "t": {"5": 0.05, "10": -0.23, "20": 0.14}},
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": {"win": {"20": 60.2}, "t": {"20": -0.94}},
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 10.8}},
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.3}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 62.1, "10": 62.1}, "t": {"5": 2.23, "10": 2.33}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 13.2}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.8}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.2}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.2}},
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 15.9}},
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.2}},
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.5}},
    "三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.6}},
    "三大法人近3月賣超 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.1}},
    "三大法人近3月賣超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 10.7}},
    "三大法人近3月賣超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.2}},
    "三大法人近3月賣超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"10": 60.9}, "t": {"10": -1.2}},
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 64.5, "10": 66.3, "20": 63.0}, "t": {"5": -0.65, "10": -1.36, "20": -0.64}},
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 61.5}, "t": {"10": 0.54}},
    "三大法人近3月賣超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 75.9, "10": 77.6, "20": 69.0}, "t": {"5": -0.27, "10": -1.4, "20": -0.82}},
    "三大法人近3月賣超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 13.0}},
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.1}},
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.1}},
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 18.9}},
    "三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 18.5}},
    "三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.4}},
    "三重底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 23.1}},
    "三重底剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.0}},
    "三重底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.9}},
    "三重底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.8}},
    "三重底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.3}},
    "三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 12.0}},
    "分價量表-突破POC追價買進 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"win": {"5": 66.1}, "t": {"5": 3.55}, "moon": {"20": 20.3}},
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.3}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量)": {"moon": {"20": 11.4}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥65": {"moon": {"20": 13.9}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": {"moon": {"10": 10.0, "20": 17.5}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上20日均線": {"moon": {"20": 12.1}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上60日均線": {"moon": {"20": 13.0}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": {"win": {"10": 62.5, "20": 64.3}, "t": {"10": -1.13, "20": 0.16}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": {"moon": {"20": 16.2}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"win": {"20": 61.4}, "t": {"20": 3.04}, "moon": {"10": 12.0, "20": 24.1}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 13.5}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"5": 10.0, "10": 16.7, "20": 16.7}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 14.3}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為正": {"moon": {"20": 12.0}},
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.5}},
    "回後買上漲全通過 ＋ 多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.2}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上20日均線": {"moon": {"20": 10.9}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上60日均線": {"moon": {"20": 10.3}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.4}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.4}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月均價YoY為負": {"moon": {"20": 10.8}},
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.7}},
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 10.3}},
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.2}},
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 60.3}, "t": {"10": 0.86}},
    "回後買上漲全通過 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.0}},
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.2}},
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.6}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"win": {"10": 60.7}, "t": {"10": 2.01}, "moon": {"10": 10.7, "20": 14.3}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.1}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.9}},
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.5, "20": 12.5}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 11.5}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.1}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.8}},
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.5}},
    "回後買上漲全通過 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.5}},
    "回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.5}},
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.3}},
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.3}},
    "圓弧底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 15.1}},
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 17.0}},
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.0}},
    "圓弧底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 13.6}},
    "圓弧底剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 13.6}},
    "圓弧底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 14.6}},
    "圓弧底剛形成 ＋ 晨星剛形成": {"moon": {"20": 12.7}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.0}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 10.4}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.8}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.3}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 12.8}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.9}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.0}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 14.6}},
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.2}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 16.0}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 12.5}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 15.4}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 14.3}},
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.7}},
    "圓弧底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.3}},
    "圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.4}},
    "圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.3}},
    "圓弧底剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 頭肩底剛形成": {"moon": {"20": 11.9}},
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 強勢突破盤": {"moon": {"20": 12.9}},
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 夜星剛形成": {"win": {"10": 60.3}, "t": {"10": 2.36}, "moon": {"10": 14.3, "20": 14.3}},
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 強勢突破盤": {"moon": {"20": 10.2}},
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 強勢突破盤": {"moon": {"20": 12.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 夜星剛形成": {"moon": {"10": 14.5, "20": 17.7}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 10.1}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 強勢突破盤": {"moon": {"20": 15.5}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 18.8}},
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 13.3}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 夜星剛形成": {"moon": {"10": 17.4, "20": 17.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 12.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 晨星剛形成": {"moon": {"10": 10.3, "20": 22.4}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.5}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 11.4}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成": {"moon": {"10": 13.1, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上20日均線": {"moon": {"10": 14.0, "20": 14.0}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上60日均線": {"moon": {"10": 12.5, "20": 17.2}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"10": 11.8, "20": 20.6}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"5": 10.0, "10": 15.0, "20": 15.0}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": {"moon": {"10": 16.1, "20": 19.4}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"10": 14.5, "20": 17.1}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"10": 13.9, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"10": 13.3, "20": 17.8}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"win": {"10": 65.5}, "t": {"10": 2.41}, "moon": {"10": 12.1, "20": 17.2}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為負": {"moon": {"10": 16.7, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 17.6, "20": 26.5}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 強勢突破盤": {"moon": {"20": 11.8}},
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 16.0}},
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 強勢突破盤": {"moon": {"20": 11.4}},
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 12.4}},
    "地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 10.4}},
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 60.9, "10": 64.1}, "t": {"5": 0.96, "10": 1.13}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 13.3}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 11.1, "20": 16.7}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(高檔)剛形成": {"moon": {"10": 10.3}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 10.0, "20": 13.3}},
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.2}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤": {"moon": {"20": 11.1}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 13.6}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 13.0, "20": 30.4}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.3}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為正": {"moon": {"20": 11.7}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為負": {"moon": {"20": 10.2}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.3}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.6}},
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.7}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"10": 11.4, "20": 11.4}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.0}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.4}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.1}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 10.2}},
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 16.8}},
    "地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.6}},
    "型態剛形成(剛突破) ＋ 多方力道≥80 ＋ 量能區間低檔(≤10百分位)": {"moon": {"20": 16.7}},
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 62.8, "10": 66.3, "20": 61.7}, "t": {"5": -1.49, "10": -1.71, "20": -0.75}},
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 70.9, "10": 71.8, "20": 71.8}, "t": {"5": 0.27, "10": -1.4, "20": -1.11}},
    "型態成形中 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"10": 60.3}, "t": {"10": 3.0}, "moon": {"10": 11.8, "20": 14.7}},
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 61.3, "10": 61.3, "20": 60.6}, "t": {"5": -1.06, "10": -2.74, "20": -1.5}},
    "型態突破確認 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 10.7}},
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 62.7, "10": 65.6, "20": 61.7}, "t": {"5": -1.07, "10": -1.72, "20": -0.4}},
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 68.1, "10": 69.9, "20": 69.9}, "t": {"5": 0.16, "10": -1.36, "20": -1.0}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 10.6}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 10.1}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.1}},
    "多方力道≥65 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 11.9, "20": 11.9}},
    "多方力道≥65 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 12.7}},
    "多方力道≥65 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 11.7}},
    "多方力道≥65 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 10.3}},
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 64.7, "20": 61.7}, "t": {"10": 0.6, "20": 0.61}},
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.6}},
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.0}},
    "多方力道≥65 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.5}},
    "多方力道≥65 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.3}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.9}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.1}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.9}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.4}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.1}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.1}},
    "多方力道≥65 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 15.9}},
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.5}},
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 12.9}},
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.4}},
    "多方力道≥65 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.7}},
    "多方力道≥80 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 16.3, "20": 16.3}},
    "多方力道≥80 ＋ 大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 14.6}},
    "多方力道≥80 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 13.8}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": {"moon": {"20": 11.1}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.3}},
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": {"moon": {"20": 10.3}},
    "多方力道≥80 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"moon": {"20": 15.3}},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 12.7}},
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 61.1}, "t": {"10": 2.48}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.8}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能區間低檔(≤10百分位)": {"moon": {"20": 12.1}},
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.3}},
    "多方力道≥80 ＋ 晨星剛形成": {"moon": {"20": 14.0}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 11.0}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.9}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 14.4}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"10": 13.2, "20": 15.8}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 10.7}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.9}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 15.5}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"10": 10.0, "20": 15.2}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 11.5}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.2}},
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.3}},
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.0}},
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.6}},
    "多方力道≥80 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "多方力道≥80 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"win": {"10": 63.4}, "t": {"10": 1.06}},
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.2}},
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.5}},
    "多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": {"moon": {"20": 10.5}},
    "夜星剛形成 ＋ 大盤站上20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 10.6, "20": 10.6}},
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 大盤跌破20日均線": {"moon": {"20": 12.1}},
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "夜星剛形成 ＋ 大盤跌破20日均線": {"moon": {"20": 12.2}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線": {"moon": {"20": 12.3}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 10.2}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.0}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.7}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.8}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.4}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月均價YoY為正": {"moon": {"20": 15.4}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 11.7}},
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.9}},
    "夜星剛形成 ＋ 大盤跌破60日均線": {"moon": {"20": 10.5}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 布林通道高檔(≥80%)": {"moon": {"20": 13.3}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.0}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.5}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正": {"win": {"5": 60.6}, "t": {"5": 2.21}, "moon": {"20": 12.4}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 11.0}},
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 15.0, "20": 15.0}},
    "夜星剛形成 ＋ 布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤)": {"moon": {"20": 16.0}},
    "夜星剛形成 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.5}},
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月均價YoY為正": {"moon": {"20": 11.7}},
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.2, "20": 13.4}},
    "夜星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 14.3}},
    "夜星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 10.9, "20": 10.9}},
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.6}},
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.3}},
    "夜星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.1, "20": 15.2}},
    "夜星剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 12.9}},
    "夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.8}},
    "大盤站上20日均線 ＋ 大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 65.3, "20": 61.8}, "t": {"10": 1.64, "20": 2.22}},
    "大盤站上20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.7}},
    "大盤站上20日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 12.1}},
    "大盤站上20日均線 ＋ 晨星剛形成": {"moon": {"20": 11.1}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.1}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 10.2}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.5}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 12.0}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 15.7}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 15.9}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.1}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.9}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.3}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 14.3}},
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.7}},
    "大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.5}},
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 12.1}},
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.5}},
    "大盤站上60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.8}},
    "大盤站上60日均線 ＋ 晨星剛形成": {"moon": {"20": 10.7}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.7}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 11.2}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 13.9}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 16.2}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.8}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.4}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.0}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 14.0}},
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 10.1}},
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 11.3}},
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 16.1}},
    "大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 頭肩底剛形成": {"moon": {"10": 10.0, "20": 10.0}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": {"win": {"5": 60.2}, "t": {"5": 4.18}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 70.8, "10": 70.8, "20": 65.2}, "t": {"5": 1.61, "10": -0.19, "20": 0.31}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 60.5, "10": 63.2, "20": 64.6}, "t": {"5": 8.94, "10": 9.01, "20": 10.12}},
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 66.9, "10": 65.4, "20": 61.7}, "t": {"5": 0.28, "10": -1.85, "20": -1.59}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 76.6, "10": 80.2, "20": 74.2}, "t": {"5": -0.22, "10": -0.92, "20": -0.29}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 62.8, "10": 62.4, "20": 64.9}, "t": {"5": 1.46, "10": 3.03, "20": 1.59}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": {"win": {"5": 66.0, "10": 63.7, "20": 68.0}, "t": {"5": 1.75, "10": 3.09, "20": 1.07}},
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": {"win": {"5": 64.3, "10": 62.8, "20": 64.9}, "t": {"5": 0.12, "10": 2.34, "20": -0.84}},
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 10.6}},
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 15.4}},
    "大盤跌破20日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 13.2}},
    "大盤跌破20日均線 ＋ 晨星剛形成": {"win": {"5": 60.5, "10": 60.7}, "t": {"5": 0.26, "10": -0.29}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 60.5, "20": 63.6}, "t": {"5": -0.68, "20": 0.59}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 60.2, "20": 64.5}, "t": {"5": 1.36, "20": 0.86}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"win": {"5": 60.6}, "t": {"5": 1.68}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 73.4, "10": 72.7, "20": 72.5}, "t": {"5": 2.12, "10": 1.5, "20": 1.91}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.2}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 64.6, "10": 64.2, "20": 60.7}, "t": {"5": 0.86, "10": -0.42, "20": -0.14}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 60.2}, "t": {"20": 1.68}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"win": {"5": 65.7, "10": 64.4, "20": 64.4}, "t": {"5": 1.66, "10": 0.53, "20": 0.48}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"20": 60.0}, "t": {"20": 2.08}},
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 66.1, "10": 65.6, "20": 62.6}, "t": {"5": 1.15, "10": -0.41, "20": -0.88}},
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"20": 60.7}, "t": {"20": 0.32}},
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.3}},
    "大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.1}},
    "大盤跌破20日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 61.8}, "t": {"20": 3.65}},
    "大盤跌破20日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 62.4}, "t": {"20": 2.27}},
    "大盤跌破20日均線 ＋ 複式頭肩底剛形成 ＋ 量能區間高檔(≥90百分位)": {"win": {"10": 61.2}, "t": {"10": 2.1}},
    "大盤跌破20日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": {"win": {"5": 62.5}, "t": {"5": 0.9}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 79.6, "10": 83.0, "20": 77.0}, "t": {"5": 0.4, "10": -0.33, "20": 0.21}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 母子懷抱(高檔)剛形成": {"win": {"20": 60.0}, "t": {"20": 1.15}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": {"win": {"5": 65.4, "10": 65.6, "20": 68.5}, "t": {"5": -0.3, "10": 0.1, "20": 0.51}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": {"win": {"5": 67.9, "10": 66.0, "20": 71.1}, "t": {"5": -0.6, "10": -0.0, "20": -0.3}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 62.1}, "t": {"20": 3.83}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 61.0}, "t": {"20": 3.43}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": {"win": {"5": 68.1, "10": 66.6, "20": 68.6}, "t": {"5": -2.56, "10": -1.88, "20": -2.94}},
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"20": 61.2}, "t": {"20": -2.76}},
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 11.4}},
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 12.1}},
    "大盤跌破60日均線 ＋ 晨星剛形成": {"win": {"5": 67.2, "10": 67.8, "20": 62.3}, "t": {"5": 1.45, "10": -0.11, "20": 0.03}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"win": {"5": 66.7, "10": 66.1, "20": 70.1}, "t": {"5": 0.3, "10": 0.78, "20": 0.8}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 64.0, "10": 62.3, "20": 68.0}, "t": {"5": 0.49, "10": 0.05, "20": 0.2}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": {"win": {"5": 60.2}, "t": {"5": -0.1}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 80.0, "10": 80.0, "20": 76.7}, "t": {"5": 3.17, "10": 2.2, "20": 1.47}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.8}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 65.7, "10": 65.4, "20": 62.3}, "t": {"5": 0.61, "10": -0.79, "20": 0.08}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 71.2, "10": 72.3, "20": 64.6}, "t": {"5": 1.9, "10": 0.95, "20": 0.42}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"win": {"5": 68.1, "10": 69.4, "20": 60.9}, "t": {"5": 1.25, "10": 0.41, "20": -0.15}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"win": {"5": 66.9, "10": 65.8, "20": 65.1}, "t": {"5": 0.9, "10": -0.72, "20": 0.19}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 64.8, "10": 61.5, "20": 62.6}, "t": {"5": 1.85, "10": 1.42, "20": 1.49}},
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 67.1, "10": 68.9, "20": 64.3}, "t": {"5": 0.49, "10": -0.46, "20": -0.62}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": {"win": {"10": 60.2}, "t": {"10": 1.07}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 66.4, "10": 62.2, "20": 68.1}, "t": {"5": 0.65, "10": -0.03, "20": 0.54}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 61.4, "10": 67.3, "20": 63.7}, "t": {"5": 1.57, "10": 2.74, "20": 3.07}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"10": 60.4}, "t": {"10": 1.24}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"10": 60.7, "20": 60.7}, "t": {"10": 0.73, "20": 1.08}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": {"win": {"10": 60.1}, "t": {"10": 0.96}},
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"10": 62.3}, "t": {"10": 0.62}},
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 64.7, "10": 62.4}, "t": {"5": 2.33, "10": 1.95}},
    "大盤跌破60日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 61.4, "20": 64.6}, "t": {"10": 1.03, "20": 2.62}},
    "大盤跌破60日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"10": 60.3, "20": 64.8}, "t": {"10": -0.51, "20": 0.98}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 跌深反彈盤": {"win": {"10": 62.4, "20": 64.1}, "t": {"10": -2.31, "20": -0.9}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": {"win": {"20": 61.2}, "t": {"20": 3.64}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正": {"win": {"10": 60.0, "20": 61.3}, "t": {"10": 10.93, "20": 11.91}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": {"win": {"20": 60.0}, "t": {"20": 1.85}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"20": 60.3}, "t": {"20": -0.12}},
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"20": 60.7}, "t": {"20": 1.53}},
    "大盤跌破60日均線 ＋ 跌深反彈盤": {"win": {"5": 61.8, "10": 61.8, "20": 60.5}, "t": {"5": -0.4, "10": -2.23, "20": -1.68}},
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 64.5, "10": 63.2, "20": 60.5}, "t": {"5": -0.29, "10": -1.88, "20": -2.79}},
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": {"win": {"5": 71.3, "10": 64.9, "20": 64.9}, "t": {"5": 0.43, "10": -2.07, "20": -0.99}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成": {"win": {"5": 76.1, "10": 79.6, "20": 73.3}, "t": {"5": -0.15, "10": -1.0, "20": -0.36}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"win": {"5": 84.8, "10": 87.9, "20": 87.9}, "t": {"5": 0.95, "10": 0.63, "20": 0.14}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 77.6, "10": 81.2, "20": 76.7}, "t": {"5": 0.06, "10": -0.41, "20": -0.55}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 81.9, "10": 84.0, "20": 81.3}, "t": {"5": 0.04, "10": -0.68, "20": -0.16}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 70.1, "10": 73.8, "20": 63.6}, "t": {"5": 0.05, "10": -0.59, "20": -0.12}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"win": {"5": 69.9, "10": 74.1, "20": 64.3}, "t": {"5": -0.8, "10": -1.38, "20": -1.03}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": {"win": {"5": 83.9, "10": 86.6, "20": 84.8}, "t": {"5": 0.73, "10": 0.21, "20": 0.76}},
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 75.2, "10": 79.5, "20": 74.4}, "t": {"5": -1.29, "10": -1.84, "20": -1.41}},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為正(強於大盤)": {"win": {"5": 66.6, "10": 63.0, "20": 64.0}, "t": {"5": -6.36, "10": -5.92, "20": -6.71}},
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 60.3, "20": 63.2}, "t": {"5": 0.86, "20": 1.75}},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 相對強弱為正(強於大盤)": {"win": {"5": 72.2, "10": 65.2, "20": 68.3}, "t": {"5": -5.33, "10": -3.87, "20": -6.06}},
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"win": {"5": 68.4}, "t": {"5": 0.15}},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能區間高檔(≥90百分位)": {"win": {"5": 69.8, "10": 66.6, "20": 67.9}, "t": {"5": -3.91, "10": -3.0, "20": -5.43}},
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 61.2}, "t": {"5": -5.3}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成": {"moon": {"20": 11.5}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.6}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": {"moon": {"20": 10.1}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 11.6}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"moon": {"20": 10.9}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.1}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.9}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 12.8}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 12.2}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.4}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 10.3, "20": 16.1}},
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 11.2}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 10.0}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.9}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.2}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 10.8}},
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.5}},
    "強勢突破盤 ＋ 晨星剛形成": {"moon": {"20": 10.9}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": {"moon": {"20": 10.4}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": {"moon": {"20": 11.6}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.3}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 11.4}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.3}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.6}},
    "強勢突破盤 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 17.6, "20": 20.6}},
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"10": 16.0, "20": 12.0}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 12.1}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 61.1, "10": 60.3, "20": 62.7}, "t": {"5": 1.34, "10": 2.12, "20": 2.79}, "moon": {"20": 10.3}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 18.6}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 11.1}},
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.0}},
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 12.1}},
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 15.7}},
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 14.0}},
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.5}},
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 10.4}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 13.9}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.3}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 10.6}},
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.6}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤)": {"win": {"5": 60.4}, "t": {"5": 1.58}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為正(營收優於股價)": {"win": {"5": 63.4, "10": 60.2, "20": 61.5}, "t": {"5": -0.01, "10": -0.12, "20": 0.33}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 60.1, "10": 60.7}, "t": {"5": 2.23, "10": 2.35}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為負": {"win": {"5": 62.9, "20": 61.8}, "t": {"5": -1.08, "20": -1.52}},
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"win": {"5": 64.8, "10": 62.3, "20": 63.6}, "t": {"5": 1.05, "10": -0.03, "20": 0.06}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 12.5}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 20.0}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.3}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 13.3}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 11.2}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 13.0}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.3}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 17.9}},
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 12.7}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 15.3}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 12.0}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"win": {"5": 60.3}, "t": {"5": 2.33}, "moon": {"10": 11.5, "20": 20.5}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 15.7}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 15.2}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 14.2}},
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 16.0}},
    "晨星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.0}},
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 近3月均價YoY為負": {"moon": {"20": 12.8}},
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 11.0}},
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 13.0}},
    "晨星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.8}},
    "晨星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 11.3}},
    "晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 12.6}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.9}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": {"moon": {"20": 10.2}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": {"win": {"20": 60.7}, "t": {"20": 1.72}, "moon": {"20": 13.1}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": {"moon": {"20": 12.1}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.7}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 10.2}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為正": {"moon": {"20": 10.2}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為負": {"moon": {"20": 10.4}},
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 11.2}},
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.3}},
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位)": {"win": {"10": 60.0}, "t": {"10": 1.72}},
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 13.3}},
    "母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 13.6}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": {"moon": {"20": 14.7}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": {"moon": {"20": 11.1}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": {"moon": {"20": 18.9}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": {"moon": {"20": 13.0}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": {"moon": {"20": 17.4}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": {"moon": {"20": 12.2}},
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": {"moon": {"20": 13.8}},
    "母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": {"moon": {"20": 10.1}},
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 距52週高點≤5%": {"win": {"5": 52.2, "10": 53.4, "20": 55.7}, "t": {"5": 6.49, "10": 7.85, "20": 9.5}, "moon": {"10": 4.1, "20": 10.3}},
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": {"win": {"5": 58.8, "10": 59.2, "20": 64.4}, "t": {"5": 5.08, "10": 5.5, "20": 6.38}, "moon": {"10": 11.6, "20": 21.7}},
    "外資近5日買超 ＋ 月營收創24個月新高 ＋ 近3日漲停": {"win": {"5": 54.7, "10": 55.8, "20": 57.0}, "t": {"5": 9.96, "10": 11.37, "20": 13.21}, "moon": {"10": 7.9, "20": 15.3}},
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 漲時量≥跌時量1.5倍(近20日)": {"win": {"5": 48.9, "10": 50.5, "20": 51.3}, "t": {"5": 12.03, "10": 18.15, "20": 22.56}, "moon": {"10": 3.9, "20": 8.2}},
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 近20日強於大盤≥10%": {"win": {"5": 49.9, "10": 51.8, "20": 52.2}, "t": {"5": 10.84, "10": 15.2, "20": 16.44}, "moon": {"10": 5.7, "20": 11.4}},

  };

  // ── 每個組合的勝率／飆股比例／t值(同日調整) ──
  // 優先用「這個頁面跑過的回測」即時計算（只算這次批次有命中的組合，算過就快取）；
  // 沒跑回測時退回 COMBO_REF_STATS 記錄值，再退回各組合原本的歷史記錄。
  function comboNormKey(combo) { return combo.slice().sort().join(' ＋ '); }
  var liveStatsCache = { sig: null, map: {}, ctx: null };
  function liveCtx() {
    var rows = backtestResults;
    if (!rows || !rows.length) return null;
    var sig = rows.length + '|' + rows[0].evalDate + '|' + rows[rows.length - 1].evalDate;
    if (liveStatsCache.sig !== sig) {
      var dayMean = {};
      [5, 10, 20].forEach(function(h) {
        var col = 'ret' + h + 'd', sum = {}, cnt = {};
        rows.forEach(function(r) { if (r[col] != null) { sum[r.evalDate] = (sum[r.evalDate] || 0) + r[col]; cnt[r.evalDate] = (cnt[r.evalDate] || 0) + 1; } });
        var m = {}; Object.keys(sum).forEach(function(d){ m[d] = sum[d] / cnt[d]; });
        dayMean[h] = m;
      });
      liveStatsCache = { sig: sig, map: {}, ctx: { rows: rows, flagDefs: buildConditionFlags(rows), dayMean: dayMean } };
    }
    return liveStatsCache.ctx;
  }
  function liveComboStats(combo) {
    var ctx = liveCtx();
    if (!ctx) return null;
    var k = comboNormKey(combo);
    if (liveStatsCache.map[k] !== undefined) return liveStatsCache.map[k];
    if (combo.some(function(c){ return !ctx.flagDefs[c]; })) { liveStatsCache.map[k] = null; return null; }
    var matched = ctx.rows.filter(function(r){ return combo.every(function(c){ return ctx.flagDefs[c](r); }); });
    var st = { win: {}, t: {}, moon: {}, src: '本次回測' };
    [5, 10, 20].forEach(function(h) {
      var col = 'ret' + h + 'd';
      var g = matched.filter(function(r){ return r[col] != null; });
      var n = g.length;
      if (n < 5) return;
      var ex = g.map(function(r){ return r[col] - ctx.dayMean[h][r.evalDate]; });
      var mean = ex.reduce(function(a,b){ return a+b; }, 0) / n;
      var sd = n > 1 ? Math.sqrt(ex.reduce(function(s,v){ return s + (v-mean)*(v-mean); }, 0) / (n-1)) : 0;
      st.win[h] = +(g.filter(function(r){ return r[col] > 0; }).length / n * 100).toFixed(1);
      st.t[h] = sd > 0 ? +(mean / (sd / Math.sqrt(n))).toFixed(2) : null;
      st.moon[h] = +(g.filter(function(r){ return r[col] > 30; }).length / n * 100).toFixed(1);
    });
    liveStatsCache.map[k] = st;
    return st;
  }
  function comboStatsFor(combo, kind) {
    var live = liveComboStats(combo);
    if (live) return live;
    var ref = COMBO_REF_STATS[comboNormKey(combo)] || {};
    var st = { win: Object.assign({}, ref.win || {}), t: Object.assign({}, ref.t || {}), moon: Object.assign({}, ref.moon || {}),
      src: Object.keys(ref).length ? '記錄值' : '' };
    var key = combo.join(' ＋ ');
    if (kind === 'S') {
      var sps = STOCK_PICK_STATS[key] || {};
      Object.keys(sps).forEach(function(h){ if (st.win[h] == null) st.win[h] = sps[h].win; if (st.t[h] == null) st.t[h] = sps[h].t; });
    } else if (kind === '#' && !Object.keys(st.win).length) {
      var wr = PINNED_COMBO_WINRATES[key] || {};
      Object.keys(wr).forEach(function(h){ st.win[h] = wr[h]; });
    } else if (kind === 'M' && !Object.keys(st.moon).length) {
      var ms = MOONSHOT_COMBO_STATS[key] || {};
      Object.keys(ms).forEach(function(h){ if (ms[h]) st.moon[h] = ms[h].pct; });
    } else if (kind === 'W') {
      var ws = BT_WIN_STATS[key] || {};
      Object.keys(ws).forEach(function(h){ if (st.win[h] == null) st.win[h] = ws[h].win; if (st.t[h] == null) st.t[h] = ws[h].t; });
      st.src = st.src || '三年回測';
    } else if (kind === 'F') {
      var fs = BT_HOT_STATS[key] || {};
      Object.keys(fs).forEach(function(h){ if (st.moon[h] == null) st.moon[h] = fs[h].pct; });
      st.src = st.src || '三年回測';
    }
    return st;
  }
  function fmtHz(d, pct) {
    var ks = Object.keys(d || {}).filter(function(h){ return d[h] != null; }).sort(function(a,b){ return a-b; });
    return ks.length ? ks.map(function(h){ return h + '日' + d[h] + (pct ? '%' : ''); }).join('/') : '—';
  }
  function bestT(st) {
    var v = Object.keys(st.t || {}).map(function(h){ return st.t[h]; }).filter(function(x){ return x != null; });
    return v.length ? Math.max.apply(null, v) : null;
  }

  // part：'S'＝只列選股型；'PM'＝只列高勝率(#)＋高標股(M)；省略＝全部
  // 命中類別數：S/#/M/W/F 五類中命中幾類（同類組合常共用同一訊號，類別數比總命中數更能反映訊號強弱）
  function hitCategories(r) {
    var cats = [], fns = [['S', matchedStockPickCombos], ['#', matchedPinnedCombos], ['M', matchedMoonshotCombos], ['W', matchedBtWinCombos], ['F', matchedBtHotCombos]];
    fns.forEach(function(p) { try { if (p[1](r).length) cats.push(p[0]); } catch (e) {} });
    return cats;
  }
  // 組合型態：依命中類別（S/#/M/W/F）的組合分級，依據三年回測（20日報酬，含 t(依日) 檢驗，2026-10-04 三段回測重新統計）
  var COMBO_STYLE_RULES = { 'S#MWF': '進攻', 'SMWF': '進攻', 'SWF': '進攻', 'SF': '進攻', 'SM': '進攻', 'WF': '進攻', 'S': '穩健', 'SW': '穩健', '#W': '穩健', 'S#WF': '穩健', '#WF': '穩健', '#F': '穩健', 'F': '彩券', 'MF': '彩券', 'M': '彩券', 'SMF': '彩券', '#MF': '彩券', '#M': '彩券' };
  var COMBO_STYLE_NOTE = { '進攻': '台股三年三段回測（2023-10～2026-09，2026-10-05 重算）：20日超額+3.3~7.8%、飆股率約2.4~5.2倍、t(依日)3.1~5.9；S#MWF最強（勝率60%、超額+7.8%），波動大（跌>10%約15~27%）', '穩健': '台股三年三段回測：20日勝率51~85%、跌>10%約5~14%（全體13.5%）；S單獨命中最穩（超額+2.7%、t(依日)8.3），#W勝率66%', '彩券': '台股三年三段回測：20日勝率42~51%、跌>10%約17~29%，飆股率約1.6~3.3倍但期望值小（t(依日)1.2~3.6）；小部位分散＋停損' };
  var COMBO_STYLE_LOOK = { '進攻': ['🚀進攻', '#ff7043', 3], '穩健': ['🛡️穩健', '#4caf50', 2], '彩券': ['🎲彩券', '#ba68c8', 1] };
  // ── 大盤狀態：先判斷大盤，再看命中的組合在「這種大盤」下回測是否有效（2026-10-05 台股三年回測，20日超額報酬）──
  //    REGIME_STATS：每個清單組合在 ①多頭 ②多頭回檔 ③空頭反彈 ④空頭 下的 [筆數, 勝率%, 同日超額%, t(依日)]；樣本<30筆為 null
  var REGIME_STATS = {
    "距52週高點≤5% ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": [[2988, 55.2, 4.14, 8.04], [472, 53.8, 3.07, 1.39], [364, 55.5, 2.97, 2.72], [205, 47.8, 1.21, -0.4]],
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月乖離度為正(營收優於股價)": [[2148, 52.9, 4.11, 6.42], [303, 56.1, 5.94, 2.31], [262, 50.8, 2.56, 1.97], [125, 48.8, 2.74, 1.25]],
    "相對強弱為正(強於大盤) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[389, 60.7, 2.44, 4.17], [93, 53.8, 0.91, 1.44], [70, 54.3, 2.08, 1.4], [92, 44.6, 3.21, 2.14]],
    "外資近5日買超 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[2041, 54.3, 6.5, 7.64], [346, 50.6, 2.55, 0.59], [163, 54.6, 3.63, 1.48], [420, 76.7, 5.95, 2.01]],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 距52週高點≤5%": [[1586, 56.1, 4.04, 6.24], [182, 52.7, 3.55, 1.28], [108, 63.0, 7.98, 3.31], [128, 49.2, 3.03, 3.02]],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[229, 65.1, 10.6, 4.88], null, null, null],
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 漲時量≥跌時量1.5倍(近20日)": [[14517, 51.0, 3.42, 13.7], null, [2003, 53.7, 1.9, 5.97], null],
    "大盤站上20日均線 ＋ 月營收創24個月新高 ＋ 近20日強於大盤≥10%": [[7289, 51.5, 3.85, 9.92], null, [1084, 56.5, 3.69, 5.25], null],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [null, null, null, [62, 91.9, 0.44, 0.11]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [null, null, null, [110, 85.5, 0.51, -0.91]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, null, null, [139, 83.5, 0.06, -1.62]],
    "KDJ近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [207, 77.8, 0.02, -0.06]],
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [128, 81.2, 0.55, -1.75]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [235, 77.0, 0.11, 0.81]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [null, null, null, [203, 81.3, 0.03, -0.18]],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [235, 77.0, 0.11, 0.81]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [null, null, null, [380, 78.7, 0.89, 0.59]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [235, 77.0, 0.11, 0.81]],
    "型態成形中 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [235, 77.0, 0.11, 0.81]],
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [235, 77.0, 0.11, 0.81]],
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [235, 77.0, 0.11, 0.81]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, null, null, [112, 76.8, -0.69, 0.63]],
    "三大法人近3月賣超 ＋ 布林通道低檔(≤20%) ＋ 晨星剛形成": [null, null, null, [107, 72.0, -0.42, 0.98]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [null, null, null, [94, 68.1, 0.27, 1.2]],
    "布林通道低檔(≤20%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [null, null, null, [125, 69.6, -0.25, 0.96]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 相對強弱為正(強於大盤)": [null, [52, 73.1, -0.64, -0.35], null, [845, 69.0, -1.64, -1.42]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [null, [97, 48.5, 1.57, 1.15], null, [380, 78.7, 0.89, 0.59]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [null, null, [39, 51.3, 1.13, -0.68], [305, 66.6, 0.44, 0.24]],
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": [null, null, null, [84, 66.7, -0.51, -1.0]],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [null, null, null, [50, 66.0, -1.02, -1.98]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [null, null, [352, 50.6, -3.55, -4.68], [2296, 73.8, 0.43, 0.56]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, null, [653, 65.2, 0.13, -0.94]],
    "布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能區間高檔(≥90百分位)": [[88, 50.0, -0.75, -0.29], [211, 61.1, 1.08, 0.76], null, [1143, 70.7, -1.93, -1.82]],
    "夜星剛形成 ＋ 布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤)": [null, null, null, null],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [66, 51.5, 1.8, -1.14], [358, 64.2, 0.3, -0.89]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [null, null, [54, 46.3, -2.02, -1.56], [412, 63.1, 0.3, -0.5]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [85, 68.2, 3.4, 1.74], [267, 67.4, 1.53, 0.62]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [null, null, [30, 50.0, 0.48, -0.34], [143, 74.8, 1.0, -0.07]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [null, null, [48, 52.1, 3.92, 0.74], [197, 72.6, -0.36, -0.86]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": [null, null, [366, 49.2, -2.52, -3.04], [2190, 71.6, -0.36, 0.27]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, null, [47, 46.8, 1.56, -0.07], [278, 68.0, -0.43, -0.08]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [null, null, [684, 49.9, -3.11, -6.19], [4666, 70.7, 0.5, 0.22]],
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [null, null, null, [36, 77.8, -0.35, -0.74]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, [74, 68.9, 3.14, 1.62]],
    "夜星剛形成 ＋ 布林通道低檔(≤20%) ＋ 近3月均價YoY為負": [null, null, null, null],
    "大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [104, 46.2, 0.31, -1.53], [653, 65.2, 0.13, -0.94]],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [null, [1010, 54.3, -0.53, 0.3], null, [2296, 73.8, 0.43, 0.56]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [null, null, null, [131, 62.6, -1.04, -2.0]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": [null, null, [204, 64.7, 2.28, 0.55], [292, 63.0, 1.33, 0.24]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [188, 50.0, 3.12, 1.5], null, [511, 68.1, -0.05, -2.09]],
    "一字底(均線糾結)剛形成 ＋ 三重底剛形成 ＋ 頭肩底剛形成": [[48, 64.6, 1.44, 1.75], null, null, null],
    "一字底(均線糾結)剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[45, 60.0, 0.14, -0.24], null, null, null],
    "地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, null, [60, 56.7, 2.04, 0.32]],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, null, null],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [null, null, [33, 54.5, 0.76, -1.0], [85, 74.1, 0.78, -0.01]],
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為正(強於大盤)": [[35, 48.6, -2.16, -1.36], [182, 69.2, -0.78, 0.14], null, [1710, 64.2, -1.33, -2.45]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [null, null, [45, 42.2, 2.49, 0.25], [232, 69.8, -0.07, 0.12]],
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [null, null, null, [36, 77.8, -0.35, -0.74]],
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 爆量(≥1.5倍均量)": [null, null, null, [67, 47.8, 1.78, 0.29]],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, null, null],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [null, [107, 49.5, 2.53, 0.27], null, [197, 72.6, -0.36, -0.86]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [null, [76, 47.4, 0.86, 0.17], null, [232, 69.8, -0.07, 0.12]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, [88, 45.5, -0.91, -0.87], null, [278, 68.0, -0.43, -0.08]],
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [38, 36.8, -2.26, -2.09], [295, 66.4, -0.07, 0.01]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [314, 62.1, 1.94, -0.23], null],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [43, 60.5, 4.98, 0.86], [193, 47.2, -1.2, -1.83]],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[243, 41.6, -0.15, 0.12], [75, 52.0, 3.19, 1.61], null, [345, 78.6, 0.7, -0.65]],
    "大盤跌破60日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": [null, null, [1289, 50.7, -1.47, -2.49], [1877, 73.8, 1.39, 0.33]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, null, [57, 40.4, -0.19, -0.74], [320, 66.6, 0.18, -0.44]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, [1169, 51.5, -0.57, -2.42], [2014, 70.2, -0.18, -1.45]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 布林通道低檔(≤20%)": [null, null, null, null],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, [107, 43.0, -0.94, 0.18], null, [320, 66.6, 0.18, -0.44]],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位)": [null, [1508, 55.1, 0.07, -0.13], null, [2190, 71.6, -0.36, 0.27]],
    "一字底(均線糾結)剛形成 ＋ 回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成": [[92, 62.0, 2.52, 1.19], null, null, null],
    "大盤跌破20日均線 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [null, [2483, 53.9, -0.49, -0.42], null, [4666, 70.7, 0.5, 0.22]],
    "大盤跌破60日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": [null, null, [2697, 51.2, -1.1, -2.32], [3853, 73.0, 1.43, -0.71]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, null, [30974, 64.6, 0.61, -2.87]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [null, null, null, [95, 66.3, -0.74, -1.43]],
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月乖離度為負(股價超前營收)": [null, null, null, [67, 62.7, -1.34, -3.0]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[121, 45.5, 1.34, 1.49], null, null, [171, 81.3, 0.13, 1.21]],
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [2736, 56.0, -0.04, -2.18], [1383, 72.7, 0.35, 0.52]],
    "夜星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[35, 42.9, 2.94, 0.95], null, null, null],
    "型態剛形成(剛突破) ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [[33, 36.4, -1.61, -0.63], null, null, [94, 80.9, -0.54, -0.66]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月均價YoY為正": [[158, 67.1, 1.89, 2.31], null, null, null],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [null, [85, 44.7, -0.12, 1.7], null, [143, 74.8, 1.0, -0.07]],
    "一字底(均線糾結)剛形成 ＋ 大盤站上20日均線 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], null, null, null],
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": [null, [45, 40.0, -1.1, -0.86], null, [50, 66.0, -1.02, -1.98]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": [[53, 45.3, 1.18, -0.04], null, null, null],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 量能區間低檔(≤10百分位)": [null, null, [699, 54.5, -0.56, -2.97], [1106, 65.8, 0.02, -0.14]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": [null, null, [41, 73.2, 1.85, 0.6], null],
    "布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[113, 47.8, 0.36, 0.73], [46, 65.2, -1.3, -0.92], null, [136, 75.7, 2.5, 1.61]],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": [[64, 56.2, 1.98, 0.64], null, null, null],
    "KDJ近3日內死亡交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": [[92, 65.2, 2.67, 1.94], null, null, null],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為正(營收優於股價)": [[142, 38.7, 0.07, 0.73], null, null, [189, 83.1, 0.77, 0.33]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, [47, 57.4, 2.66, 1.01], [567, 61.2, -0.24, 0.38]],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[56, 37.5, 3.06, 0.87], null, null, null],
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, [1232, 53.8, -2.12, -2.69], [10213, 63.3, 0.4, -1.36]],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 布林通道低檔(≤20%)": [null, null, null, null],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [114, 55.3, 3.77, 1.89], null, [88, 55.7, 4.0, 1.68]],
    "一字底(均線糾結)剛形成 ＋ 大盤站上60日均線 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], null, null, null],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, [66, 50.0, 2.89, 1.07], null, [74, 68.9, 3.14, 1.62]],
    "型態突破確認 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量)": [[39, 41.0, -1.03, -0.45], null, null, [97, 79.4, -0.6, -0.79]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[208, 40.9, 1.52, 0.84], [51, 49.0, 1.28, 0.58], null, [200, 79.5, 1.24, -1.54]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 跌深反彈盤": [null, null, null, [93, 66.7, -0.63, -0.79]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ N字底剛形成": [[44, 54.5, 0.32, 0.47], null, null, null],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為負": [[112, 35.7, -3.7, -1.9], null, null, [141, 84.4, 0.04, -0.15]],
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, [51, 58.8, 2.2, 0.95], [591, 61.1, -0.1, 0.17]],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[71, 49.3, 1.45, 0.95], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [163, 66.9, 1.68, -0.74], [128, 43.8, -1.64, -1.99]],
    "大盤跌破20日均線 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": [null, [1846, 50.9, -0.57, 0.15], null, [1877, 73.8, 1.39, 0.33]],
    "三大法人近3月買超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [181, 49.7, 1.04, 1.3], null, [358, 64.2, 0.3, -0.89]],
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間低檔(≤10百分位)": [[35, 28.6, -3.29, -1.47], null, [32, 75.0, 1.24, 0.72], null],
    "KDJ近3日內死亡交叉 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[56, 55.4, 0.01, -0.2], null, null, null],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": [[54, 63.0, 1.79, 1.9], null, null, null],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤)": [null, null, [3688, 51.1, -2.13, -8.31], [17443, 63.6, 0.77, -1.82]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[662, 46.8, -0.55, -2.26], [183, 46.4, -0.15, 0.27], [113, 56.6, 0.22, 0.75], [673, 82.0, 0.29, 0.81]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": [[38, 50.0, 5.41, 1.31], null, null, null],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [3742, 54.8, -0.19, -3.16], [2768, 71.4, 0.94, 0.61]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": [[235, 41.3, 0.82, 1.34], [63, 50.8, 1.77, 1.22], null, [181, 76.2, 1.28, 0.8]],
    "一字底(均線糾結)剛形成 ＋ 強勢突破盤 ＋ 頭肩底剛形成": [[53, 62.3, 0.84, 1.17], null, null, null],
    "大盤跌破60日均線 ＋ 跌深反彈盤": [null, null, null, [131, 62.6, -1.04, -2.0]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 跌深反彈盤": [null, null, null, [131, 62.6, -1.04, -2.0]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [null, null, null, [131, 62.6, -1.04, -2.0]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, [90, 53.3, -2.95, -2.23], [513, 62.4, 0.2, -1.04]],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[47, 42.6, 3.76, 1.44], null, null, null],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[45, 46.7, 6.04, 2.05], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[310, 47.7, 2.66, 2.77], [105, 54.3, 5.25, 1.83], null, [246, 67.5, -0.14, -1.71]],
    "大盤跌破20日均線 ＋ 爆量(≥1.5倍均量) ＋ 相對強弱為負(弱於大盤)": [null, [4154, 51.4, -0.47, -0.46], null, [3853, 73.0, 1.43, -0.71]],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤)": [[130, 46.9, 3.08, 2.43], null, null, [88, 89.8, 3.05, 0.82]],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[385, 47.5, 2.88, 2.7], [135, 50.4, 3.21, 1.31], null, [318, 65.7, 0.28, -2.48]],
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位)": [[161, 56.5, 2.29, 1.86], null, null, null],
    "MACD近3日內黃金交叉 ＋ N字底剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[53, 49.1, -0.45, -0.25], null, null, null],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[121, 63.6, 1.17, 1.81], null, null, null],
    "布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[97, 46.4, -1.62, -1.66], [307, 67.8, 1.37, 0.12], [31, 64.5, 0.39, 0.4], [2001, 57.9, -1.01, 0.6]],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [null, null, null, [96, 64.6, -1.15, -2.12]],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正": [null, null, [44, 45.5, 0.94, 0.07], [91, 50.5, 3.06, 0.85]],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": [[65, 47.7, 6.33, 1.92], null, null, null],
    "MACD近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線": [null, null, [1102, 53.0, -0.73, -4.09], [1410, 65.4, -0.21, -0.61]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[74, 62.2, 10.94, 2.99], null, null, null],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, null, [975, 51.3, -2.73, -5.29], [5759, 62.2, 0.03, 0.38]],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [7262, 54.0, -0.48, -2.77], [10568, 66.3, 0.91, -1.49]],
    "大盤跌破20日均線 ＋ 複式頭肩底剛形成 ＋ 量能區間高檔(≥90百分位)": [null, [34, 61.8, 2.35, 0.08], null, [33, 42.4, 3.72, 1.45]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正": [null, null, [10978, 53.2, 0.21, 1.46], [19634, 65.2, 1.11, -1.42]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[56, 60.7, 5.7, 2.27], null, null, null],
    "大盤跌破20日均線 ＋ 晨星剛形成": [null, [312, 47.4, 1.06, 1.47], null, [653, 65.2, 0.13, -0.94]],
    "型態成形中 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [312, 47.4, 1.06, 1.47], null, [653, 65.2, 0.13, -0.94]],
    "型態突破確認 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [312, 47.4, 1.06, 1.47], null, [653, 65.2, 0.13, -0.94]],
    "型態剛形成(剛突破) ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [312, 47.4, 1.06, 1.47], null, [653, 65.2, 0.13, -0.94]],
    "大盤跌破60日均線 ＋ 跌深反彈盤 ＋ 近3月乖離度為正(營收優於股價)": [null, null, null, [61, 60.7, -1.79, -1.94]],
    "一字底(均線糾結)剛形成 ＋ 布林通道高檔(≥80%) ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], null, null, null],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": [[140, 62.9, 1.39, 1.71], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[191, 52.4, 3.61, 2.2], [56, 46.4, 0.99, 1.52], null, [116, 81.9, 2.01, 0.33]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [null, null, null, [119, 63.0, -0.99, -1.77]],
    "夜星剛形成 ＋ 布林通道低檔(≤20%) ＋ 近3月乖離度為正(營收優於股價)": [null, null, null, null],
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, null, null, [37, 54.1, 0.45, 0.04]],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[58, 55.2, 8.8, 2.51], null, null, null],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, [80, 67.5, 4.48, 0.72], [128, 46.1, -1.41, -3.75]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 跌深反彈盤": [null, null, null, [36, 63.9, -1.28, -0.88]],
    "KDJ近3日內黃金交叉 ＋ 跌深反彈盤 ＋ 近3月乖離度為正(營收優於股價)": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線": [null, null, [655, 53.0, -0.63, -1.2], [1894, 64.1, 0.42, 0.22]],
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能區間低檔(≤10百分位)": [null, null, null, [47, 61.7, 1.26, 1.64]],
    "圓弧底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], null, null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, null, null],
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 量能區間低檔(≤10百分位)": [[41, 48.8, 0.95, 0.5], null, null, null],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[51, 56.9, 2.81, 1.05], null, null, null],
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [null, [65, 43.1, -0.2, -0.27], null, [85, 74.1, 0.78, -0.01]],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[39, 48.7, 5.95, 1.62], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[175, 43.4, 2.38, 1.7], [43, 46.5, 0.64, -0.44], null, [231, 70.1, -0.71, -2.42]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [null, [64, 45.3, 1.08, 0.31], null, [96, 64.6, -1.15, -2.12]],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 量能區間低檔(≤10百分位)": [[33, 48.5, -1.18, -0.79], null, [31, 67.7, 1.31, 0.73], null],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 強勢突破盤": [[273, 59.3, 9.97, 4.05], null, null, null],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[65, 58.5, 4.83, 1.91], null, null, null],
    "大盤跌破20日均線 ＋ 跌深反彈盤 ＋ 近3月均價YoY為正": [null, [52, 38.5, -2.05, -2.0], null, [84, 66.7, -0.51, -1.0]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正": [[294, 42.2, 1.95, 2.03], [82, 47.6, 1.25, 0.69], null, [236, 75.4, 1.38, 0.44]],
    "三大法人近3月賣超 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [131, 44.3, 1.09, 0.78], null, [295, 66.4, -0.07, 0.01]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [null, [187, 49.7, 2.12, 2.18], null, [305, 66.6, 0.44, 0.24]],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": [null, [282, 53.5, 0.6, 0.91], null, [267, 67.4, 1.53, 0.62]],
    "KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 晨星剛形成": [[340, 48.2, 2.08, 2.2], [103, 51.5, 1.88, 0.5], [30, 56.7, 2.08, 0.26], [268, 67.5, 0.42, -2.49]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, [1227, 45.5, -0.82, -2.76], null, [2014, 70.2, -0.18, -1.45]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, null, [4369, 52.7, -0.61, -3.21], [5774, 65.9, 0.75, -0.29]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收)": [null, null, [9328, 50.8, -0.87, -4.85], [16021, 66.3, 0.87, 0.52]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 三大法人近3月買超": [[153, 60.1, 2.74, 2.59], null, null, null],
    "MACD近3日內黃金交叉 ＋ 布林通道低檔(≤20%) ＋ 量能區間低檔(≤10百分位)": [[86, 37.2, -0.8, -0.24], null, null, [54, 74.1, 1.01, -0.05]],
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [160, 61.9, 0.69, -2.79], [268, 55.2, -0.2, -0.55]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], null, null, null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [36, 47.2, 1.78, 0.7], null, null],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [7238, 53.8, -0.79, -2.94], [9571, 65.6, 0.63, -1.2]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 相對強弱為負(弱於大盤)": [[55, 56.4, 4.93, 2.1], null, null, null],
    "K線橫盤的突破剛形成 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], null, null, null],
    "K線橫盤的突破剛形成 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], null, null, null],
    "K線橫盤的突破剛形成 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], null, null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價)": [null, null, [36, 50.0, 1.85, 0.75], [77, 41.6, 0.11, -0.27]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 近3月乖離度為負(股價超前營收)": [null, null, [1316, 45.5, -3.16, -8.11], [13437, 61.6, 0.62, 0.05]],
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[60, 60.0, 5.55, 2.2], null, null, null],
    "MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 量能區間低檔(≤10百分位)": [null, [784, 51.0, -0.81, 0.18], null, [1106, 65.8, 0.02, -0.14]],
    "晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], null, [380, 78.7, 0.89, 0.59]],
    "型態成形中 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], null, [380, 78.7, 0.89, 0.59]],
    "型態突破確認 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], null, [380, 78.7, 0.89, 0.59]],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], null, [380, 78.7, 0.89, 0.59]],
    "K線橫盤的突破剛形成 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成": [[113, 59.3, 2.13, 2.12], null, null, null],
    "一字底(均線糾結)剛形成 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], null, null, null],
    "一字底(均線糾結)剛形成 ＋ 型態成形中 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], null, null, null],
    "一字底(均線糾結)剛形成 ＋ 型態突破確認 ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], null, null, null],
    "一字底(均線糾結)剛形成 ＋ 型態剛形成(剛突破) ＋ 頭肩底剛形成": [[63, 63.5, 1.56, 1.99], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 相對強弱為負(弱於大盤)": [null, [42, 57.1, 2.34, 1.36], null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": [null, null, null, [51, 47.1, 2.16, -0.24]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 布林通道低檔(≤20%)": [null, null, null, null],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, null, [172, 55.2, 1.76, -0.52], [278, 53.2, 0.33, -0.15]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [null, null, [129, 69.0, 1.4, -0.17], [236, 56.8, 0.38, -0.08]],
    "三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[51, 54.9, 4.0, 1.48], null, null, null],
    "三大法人近3月賣超 ＋ 大盤跌破60日均線 ＋ 跌深反彈盤": [null, null, null, [81, 60.5, -1.06, -1.54]],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [154, 62.3, 3.24, 0.59], [274, 52.2, -0.09, -1.76]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為負(弱於大盤)": [[134, 53.7, -0.01, 0.3], null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": [[84, 44.0, 3.53, 1.23], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, [741, 52.1, -2.07, -5.42], [7558, 60.5, 0.18, -0.38]],
    "大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 相對強弱為負(弱於大盤)": [null, null, [99, 51.5, 1.5, 1.78], [169, 62.1, 0.76, 0.08]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": [null, null, [1223, 48.7, -0.67, -1.81], [1486, 68.3, 1.35, -0.51]],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, [6077, 52.0, -0.39, -1.5], [8584, 65.8, 0.21, -1.16]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 大盤站上20日均線": [[244, 58.6, 1.23, 1.82], null, null, null],
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[69, 58.0, 4.52, 2.01], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線": [null, null, null, [155, 48.4, 2.66, 0.74]],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內黃金交叉 ＋ 型態成形中 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內黃金交叉 ＋ 型態突破確認 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], [49, 51.0, 2.47, 0.26], [511, 68.1, -0.05, -2.09]],
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[68, 47.1, 3.59, 1.22], null, null, null],
    "N字底剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [268, 53.7, -0.82, -3.02], [31, 61.3, -1.57, -2.52]],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [null, [228, 47.8, 1.21, 2.2], null, [412, 63.1, 0.3, -0.5]],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[70, 45.7, 4.62, 1.52], null, null, null],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [10738, 52.8, 0.03, 0.39], [13073, 66.0, 0.9, -0.84]],
    "三大法人近3月買超 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, [1312, 54.3, -0.85, -3.66], [12102, 59.6, 0.35, 0.73]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [null, null, [147, 55.1, 0.94, -0.18], [209, 56.0, 1.03, 0.68]],
    "三大法人近3月買超 ＋ 布林通道低檔(≤20%) ＋ 爆量(≥2倍均量)": [[890, 42.7, -0.92, -2.09], [377, 58.4, -0.55, -1.29], [103, 51.5, -3.14, -2.2], [968, 76.5, 0.69, 1.4]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[48, 58.3, 5.11, 1.83], null, null, null],
    "一字底(均線糾結)剛形成 ＋ 三大法人近3月買超 ＋ 頭肩底剛形成": [[43, 62.8, 2.42, 2.21], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, [2025, 53.1, 0.23, -0.18], [5979, 61.8, -0.0, -0.48]],
    "大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 多方力道≥65": [[129, 58.9, 2.53, 1.66], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[220, 42.3, 1.15, 0.9], [75, 42.7, 0.15, 0.52], null, [251, 70.9, 0.47, -0.7]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "型態突破確認 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, [314, 62.1, 1.94, -0.23], [542, 53.7, -0.15, -1.69]],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 母子懷抱(高檔)剛形成": [null, null, null, [131, 57.3, 0.51, 1.07]],
    "一字底(均線糾結)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成": [[54, 64.8, 1.59, 1.7], null, null, null],
    "大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能區間低檔(≤10百分位)": [null, null, null, [47, 61.7, 1.26, 1.64]],
    "大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 布林通道低檔(≤20%)": [null, null, null, [27704, 59.0, 0.25, 0.8]],
    "型態成形中 ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [22364, 53.1, -0.58, -2.38], [26340, 65.0, 0.63, -2.64]],
    "型態剛形成(剛突破) ＋ 大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤)": [null, null, [1648, 53.5, -0.03, -0.15], [1014, 64.9, 0.29, -0.8]],
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成": [null, null, [198, 57.1, 1.62, 1.59], null],
    "三重底剛形成 ＋ 多方力道≥65 ＋ 頭肩底剛形成": [[80, 45.0, -0.79, -0.5], null, null, null],
    "圓弧底剛形成 ＋ 大盤跌破20日均線 ＋ 突破飆股大量黑K最高點剛形成": [null, [41, 51.2, 3.76, 0.78], null, [30, 43.3, 3.27, 0.85]],
    "多方力道≥80 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": [[71, 56.3, 2.79, 1.01], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[70, 48.6, 4.97, 1.52], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位)": [null, null, null, null],
    "三大法人近3月賣超 ＋ 爆量(≥1.5倍均量) ＋ 跌深反彈盤": [[47, 48.9, -1.05, -0.18], null, null, null],
    "大盤跌破60日均線 ＋ 布林通道低檔(≤20%) ＋ 近3月均價YoY為正": [null, null, [1407, 48.6, -1.43, -5.23], [16616, 60.0, 0.62, 1.23]],
    "N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 相對強弱為正(強於大盤)": [[110, 64.5, 2.74, 2.48], null, null, null],
    "大盤跌破60日均線 ＋ 相對強弱為負(弱於大盤) ＋ 近3日向上跳空缺口": [null, null, [1909, 58.7, 0.64, 1.56], [3426, 82.2, 1.6, 0.61]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[88, 51.1, 5.76, 1.85], null, null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": [[34, 47.1, 3.09, 1.07], null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 晨星剛形成": [[46, 56.5, 9.41, 1.27], null, null, null],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[53, 52.8, 4.54, 1.78], null, null, null],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [null, null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[39, 41.0, 1.08, 0.36], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[46, 52.2, 8.64, 2.45], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], null, null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [null, null, null, null],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[30, 30.0, -0.39, 0.11], null, null, null],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[33, 60.6, 11.08, 1.96], null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 晨星剛形成": [[51, 54.9, 7.42, 1.01], null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[34, 47.1, 1.71, 0.18], null, null, null],
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "三重底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [null, null, null, null],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [null, null, null, null],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[91, 49.5, 6.49, 2.49], null, null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 夜星剛形成": [[33, 48.5, 6.95, 1.49], null, null, null],
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [[38, 60.5, 7.29, 2.51], null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], null, null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[134, 53.7, 4.97, 1.74], null, null, null],
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[102, 48.0, 6.03, 2.41], null, null, null],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], null, null, null],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": [[33, 48.5, 5.59, 1.99], null, null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[37, 56.8, 6.04, 1.61], null, null, null],
    "夜星剛形成 ＋ 布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤)": [null, null, null, null],
    "三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], null, null, null],
    "型態成形中 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], null, null, null],
    "型態突破確認 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], null, null, null],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 46.6, 4.91, 2.03], null, null, null],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": [[54, 40.7, 4.02, 1.14], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[41, 46.3, 4.11, 1.57], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[93, 41.9, 4.79, 1.68], null, null, null],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[54, 57.4, 6.06, 2.19], null, null, null],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "圓弧底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [[33, 63.6, 7.24, 2.08], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[113, 46.0, 4.33, 1.78], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[32, 34.4, -1.51, -0.5], null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [null, null, null, null],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[82, 41.5, 2.5, 1.03], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[72, 50.0, 2.01, 0.87], null, null, null],
    "K線橫盤的突破剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[137, 48.9, 5.61, 2.72], [31, 38.7, -1.15, -0.87], null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[134, 50.7, 3.83, 1.83], null, null, null],
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": [[50, 40.0, 3.98, 1.22], null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[36, 58.3, 7.6, 2.04], null, null, null],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[50, 46.0, 3.18, 0.61], null, null, null],
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [null, [68, 58.8, 6.98, 2.21], null, [36, 44.4, -2.33, -0.93]],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[73, 46.6, 1.33, 0.56], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[71, 45.1, 1.48, 0.61], null, null, null],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[37, 54.1, 5.8, 1.71], null, null, null],
    "N字底剛形成 ＋ 三大法人近3月賣超 ＋ 突破ABC修正下降切線剛形成": [[30, 46.7, 4.26, 1.25], null, null, null],
    "圓弧底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [null, null, null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[34, 47.1, 2.54, 0.42], null, null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": [[163, 52.1, 3.74, 1.63], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[121, 50.4, 6.54, 2.93], [31, 38.7, -1.67, -1.12], null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], null, null, null],
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 晨星剛形成": [[251, 52.2, 5.0, 3.06], [38, 39.5, -0.78, -0.34], null, null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為負": [[33, 48.5, 5.47, 1.09], null, null, null],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": [[108, 44.4, 4.49, 1.56], null, null, null],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[60, 45.0, 3.7, 1.46], null, null, null],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[149, 50.3, 2.79, 2.27], [45, 57.8, 4.69, 1.28], null, [40, 70.0, 5.54, 1.95]],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], null, null, null],
    "型態成形中 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], null, null, null],
    "型態突破確認 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], null, null, null],
    "型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[52, 55.8, 5.54, 1.81], null, null, null],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 夜星剛形成": [null, null, null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[57, 50.9, 5.0, 1.69], null, null, null],
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[58, 41.4, 0.36, 0.4], null, null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[207, 48.3, 4.2, 2.12], [39, 33.3, -2.4, -0.38], null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[182, 47.8, 4.93, 2.36], null, null, null],
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[77, 51.9, 6.36, 2.28], null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 強勢突破盤": [[400, 57.5, 7.41, 4.34], [46, 54.3, 1.74, 0.71], [31, 58.1, 10.86, 3.01], null],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[50, 46.0, 5.05, 1.32], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": [[128, 49.2, 5.25, 2.17], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[71, 46.5, 2.77, 0.84], null, null, null],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": [[102, 47.1, 5.71, 2.05], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": [[40, 37.5, -2.88, -1.23], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[49, 51.0, 2.62, 0.72], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[86, 39.5, 3.54, 1.3], null, null, null],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [null, null, null, null],
    "多方力道≥80 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [34, 47.1, 2.53, 0.03], null],
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成": [[68, 44.1, 1.74, 0.61], null, null, null],
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": [[42, 50.0, 5.28, 1.79], null, null, null],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], null, null, null],
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": [[37, 54.1, 5.96, 1.65], null, null, null],
    "夜星剛形成 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "KDJ近3日內黃金交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [null, null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [null, null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[62, 48.4, 2.69, 0.69], null, null, null],
    "圓弧底剛形成 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[35, 57.1, 5.93, 1.6], null, null, null],
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], null, null, null],
    "多方力道≥80 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], null, [34, 47.1, 2.53, 0.03], null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 布林通道高檔(≥80%)": [null, null, null, null],
    "圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[53, 54.7, 4.11, 1.44], null, null, null],
    "多方力道≥80 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[33, 48.5, 6.8, 1.59], null, null, null],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[44, 45.5, 2.7, 0.65], null, null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[330, 47.9, 4.11, 2.36], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], null],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成": [[37, 45.9, 2.22, 1.02], null, null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收)": [[138, 47.8, 3.4, 1.48], null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, null, null, null],
    "大盤跌破20日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 母子懷抱(高檔)剛形成": [[108, 31.5, 0.14, -0.19], [37, 56.8, 10.33, 1.66], null, null],
    "K線橫盤的突破剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [[87, 47.1, 5.2, 2.06], null, null, null],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": [[93, 49.5, 5.92, 2.16], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[116, 51.7, 5.39, 2.38], null, null, null],
    "多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], null],
    "多方力道≥65 ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], null],
    "型態成形中 ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], null],
    "型態突破確認 ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], null],
    "型態剛形成(剛突破) ＋ 多方力道≥80 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], [34, 47.1, 2.53, 0.03], null],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[56, 57.1, 3.57, 1.24], null, null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], null, null, null],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[86, 51.2, 5.7, 2.46], null, null, null],
    "N字底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [null, null, null, null],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], null, null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], null, null],
    "型態剛形成(剛突破) ＋ 多方力道≥80 ＋ 量能區間低檔(≤10百分位)": [[36, 36.1, 0.76, 0.28], null, null, null],
    "多方力道≥80 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[343, 46.6, 3.44, 2.14], [55, 38.2, -1.38, -0.56], null, null],
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [null, null, null, null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[46, 41.3, 2.39, 0.8], null, null, null],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[167, 50.3, 3.23, 1.98], [46, 50.0, 2.31, 0.67], null, [44, 56.8, 3.09, 1.47]],
    "N字底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [null, null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": [[72, 54.2, 4.72, 1.7], null, null, null],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], null, null],
    "大盤跌破60日均線 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "圓弧底剛形成 ＋ 型態成形中 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], null, null, null],
    "K線橫盤的突破剛形成 ＋ 型態成形中 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], null, null],
    "圓弧底剛形成 ＋ 型態突破確認 ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], null, null, null],
    "K線橫盤的突破剛形成 ＋ 型態突破確認 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], null, null],
    "圓弧底剛形成 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[44, 59.1, 5.07, 1.57], null, null, null],
    "K線橫盤的突破剛形成 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [[31, 71.0, 7.56, 1.87], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[52, 36.5, 0.81, 0.17], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 晨星剛形成": [[135, 47.4, 2.08, 1.58], null, null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[75, 40.0, 4.76, 2.06], null, null, null],
    "K線橫盤的突破剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[156, 46.2, 4.27, 2.16], [34, 44.1, -0.53, -0.88], null, null],
    "大盤跌破20日均線 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[66, 50.0, 3.96, 0.62], null, null, [31, 58.1, 3.02, 0.71]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[59, 44.1, 1.42, 0.46], null, null, null],
    "多方力道≥80 ＋ 強勢突破盤 ＋ 晨星剛形成": [[149, 41.6, 2.7, 1.49], null, null, null],
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 強勢突破盤": [[104, 56.7, 4.43, 1.67], null, null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[171, 46.8, 3.73, 2.3], [31, 41.9, 0.67, 0.13], null, null],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[63, 52.4, 8.45, 1.97], null, null, null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 相對強弱為正(強於大盤)": [[430, 57.4, 6.69, 3.94], [68, 47.1, -0.09, 0.5], [43, 58.1, 7.25, 2.31], [47, 46.8, 5.2, 2.69]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 相對強弱為正(強於大盤)": [null, [172, 50.0, 2.27, 1.32], null, [143, 47.6, 3.19, 0.91]],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[37, 45.9, 2.68, 0.69], null, null, null],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為正": [[42, 45.2, 1.52, 0.2], null, null, null],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥65": [[164, 51.8, 3.58, 1.55], null, null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[296, 48.0, 4.15, 2.89], [62, 37.1, -0.36, 0.65], null, null],
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [[103, 46.6, 3.94, 1.67], null, null, null],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[53, 54.7, 4.6, 1.44], null, null, null],
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[39, 56.4, 5.03, 1.47], null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[53, 43.4, 2.84, 1.19], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[100, 49.0, 5.48, 2.46], null, null, null],
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[137, 47.4, 2.37, 2.26], [39, 48.7, 1.23, 0.28], null, [35, 71.4, 3.41, 1.3]],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能區間高檔(≥90百分位)": [[32, 43.8, 1.52, 0.31], null, null, null],
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 晨星剛形成": [[380, 48.4, 3.35, 2.37], [66, 39.4, 0.27, 0.72], [30, 63.3, 6.62, 0.48], [51, 54.9, 0.27, 0.21]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 相對強弱為正(強於大盤)": [[187, 49.7, 2.97, 1.42], null, null, null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月乖離度為負(股價超前營收)": [[343, 56.3, 6.29, 3.36], null, null, null],
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": [null, null, null, null],
    "多方力道≥65 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], null, [47, 48.9, 2.73, 0.06], null],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": [[66, 42.4, 5.48, 1.54], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": [[146, 44.5, 3.65, 1.8], [34, 47.1, 1.67, 0.11], null, null],
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 晨星剛形成": [[97, 46.4, 3.68, 1.26], null, null, null],
    "三大法人近3月賣超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[195, 46.2, 2.88, 1.73], [45, 51.1, 4.73, 0.86], null, null],
    "N字底剛形成 ＋ 相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成": [[46, 41.3, 2.3, 0.49], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[153, 44.4, 1.03, -0.02], [30, 36.7, 8.51, 0.94], null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上60日均線": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [null, null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], null, null, null],
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[259, 45.2, 1.7, 1.39], [35, 37.1, -0.74, -0.46], null, null],
    "三大法人近3月賣超 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [[63, 42.9, 1.54, 0.42], null, null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[89, 40.4, 1.44, 0.58], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[336, 44.9, 2.08, 1.99], [61, 50.8, 5.72, 1.83], null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月均價YoY為正": [null, [137, 55.5, 5.07, 2.59], null, [91, 50.5, 3.06, 0.85]],
    "型態成形中 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], null, null, null],
    "型態突破確認 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], null, null, null],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[123, 45.5, 3.92, 1.67], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上60日均線": [[48, 45.8, 3.52, 1.47], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能區間高檔(≥90百分位)": [null, [93, 46.2, 0.62, 0.26], null, [70, 42.9, 2.7, 1.33]],
    "三重底剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(高檔)剛形成": [[58, 32.8, -2.41, -1.31], null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [[77, 42.9, 3.25, 1.3], null, null, null],
    "大盤站上20日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], null, [37, 45.9, 1.33, -0.14], null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 布林通道低檔(≤20%)": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成": [[36, 41.7, 3.17, 0.96], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[80, 45.0, 4.77, 2.14], null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": [[36, 50.0, 4.94, 1.52], null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[352, 45.2, 2.45, 1.32], [72, 37.5, -1.06, 0.6], null, [36, 47.2, -3.32, -1.52]],
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 強勢突破盤": [[305, 56.4, 5.65, 3.24], [36, 61.1, 2.2, 0.75], [32, 56.2, 8.42, 2.12], null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 爆量(≥2倍均量)": [null, null, null, null],
    "晨星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[167, 44.9, 1.91, 1.39], [43, 48.8, 1.82, 0.67], null, [47, 70.2, 4.27, 1.24]],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[40, 37.5, 0.64, 0.48], null, null, null],
    "晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], null, [74, 68.9, 3.14, 1.62]],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 晨星剛形成": [[60, 36.7, -0.48, 0.1], null, null, null],
    "KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": [[65, 44.6, 3.62, 1.47], null, null, null],
    "晨星剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成": [[67, 43.3, 2.54, 1.01], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[126, 45.2, 3.38, 1.71], [34, 44.1, -0.53, -0.88], null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[35, 42.9, 2.17, 1.04], null, null, null],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": [[63, 47.6, 4.03, 0.72], null, null, [34, 58.8, 3.03, 1.22]],
    "型態成形中 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], null, [74, 68.9, 3.14, 1.62]],
    "型態突破確認 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], null, [74, 68.9, 3.14, 1.62]],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[241, 45.6, 2.0, 1.55], [66, 50.0, 2.89, 1.07], null, [74, 68.9, 3.14, 1.62]],
    "圓弧底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[37, 51.4, 1.51, 0.36], null, null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[87, 43.7, 3.36, 1.14], null, null, null],
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[115, 47.8, 1.97, 1.13], null, null, null],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成": [[39, 38.5, -1.39, -0.25], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[412, 44.7, 1.91, 0.9], [70, 50.0, 4.65, 1.89], null, null],
    "母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": [[66, 43.9, 1.71, 0.56], null, null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[505, 46.3, 2.86, 1.79], [104, 35.6, -1.25, 0.06], [46, 50.0, 3.01, 0.18], [61, 44.3, -1.25, -0.44]],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[428, 55.6, 5.8, 3.22], [48, 50.0, 0.08, 0.3], [32, 68.8, 13.04, 2.87], null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為正": [[185, 47.0, 2.01, 1.37], null, null, null],
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[185, 49.7, 2.26, 0.74], null, null, null],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 晨星剛形成": [[102, 49.0, 4.39, 2.26], null, null, null],
    "N字底剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[40, 37.5, 0.64, 0.48], null, null, null],
    "夜星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[44, 56.8, 4.42, 1.3], null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 布林通道高檔(≥80%)": [[1519, 49.8, 3.41, 3.73], [184, 43.5, 1.79, -0.58], [140, 62.1, 8.4, 2.76], [101, 42.6, 2.8, 1.91]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[396, 46.2, 2.48, 1.37], [47, 55.3, 7.23, 1.09], [31, 58.1, 2.15, 1.32], null],
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": [[63, 46.0, 4.31, 1.38], null, null, null],
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": [[35, 48.6, 1.11, 0.45], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [64, 51.6, 1.29, -0.19], null, [50, 48.0, 2.43, 0.04]],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月均價YoY為正": [[283, 42.0, 0.52, 0.48], [113, 55.8, 5.55, 2.41], null, [85, 51.8, 3.6, 1.14]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 大盤站上20日均線": [[240, 46.7, 1.84, 1.15], null, [41, 73.2, 1.85, 0.6], null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[654, 45.1, 2.19, 2.52], null, [39, 51.3, 1.13, -0.68], null],
    "大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[78, 46.2, 1.37, 0.61], null, null, null],
    "圓弧底剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], null, null, null],
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], null, null],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[103, 46.6, 1.56, 0.94], null, null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], null, [80, 46.2, 0.41, -1.46], null],
    "三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[101, 48.5, 3.52, 1.51], null, null, null],
    "大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, [64, 45.3, 3.23, 0.84], null, [89, 56.2, 2.78, -0.55]],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[111, 50.5, 3.56, 1.79], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], null, null, null],
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成": [null, [46, 47.8, 6.58, 1.42], null, null],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位)": [[105, 51.4, 3.91, 1.5], null, null, null],
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], null, null, null],
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], null, null, null],
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": [[48, 45.8, 3.52, 1.47], null, null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[799, 44.8, 2.19, 1.65], null, [54, 46.3, -2.02, -1.56], null],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[73, 39.7, 2.33, 1.3], null, null, null],
    "夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], null, null, null],
    "多方力道≥65 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], null, null],
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [55, 38.2, -1.38, -0.56], null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[288, 41.3, 2.51, 1.51], [58, 41.4, -0.58, -0.83], [36, 47.2, 1.59, -0.08], [30, 46.7, 0.6, 0.64]],
    "KDJ近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[289, 47.8, 3.18, 2.58], [59, 52.5, 5.06, 1.01], null, null],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[63, 49.2, 2.55, 0.8], null, null, null],
    "大盤站上20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], null, [62, 48.4, 1.08, -0.28], null],
    "KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[167, 47.9, 3.9, 2.1], null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [null, null, null, null],
    "型態成形中 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], null, null, null],
    "型態突破確認 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], null, null, null],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], null, null, null],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[81, 44.4, 1.96, 1.34], null, null, null],
    "三大法人近3月賣超 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": [[42, 54.8, 4.23, 1.48], null, null, null],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": [[44, 56.8, 4.48, 1.83], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成": [null, null, null, null],
    "夜星剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[418, 44.0, 1.42, 1.13], [67, 37.3, -0.34, 0.11], [36, 50.0, 3.02, 0.34], [33, 42.4, -0.94, 0.09]],
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能區間低檔(≤10百分位)": [null, null, null, null],
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[161, 50.9, 4.12, 2.41], null, null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[197, 42.1, 1.66, 1.23], [34, 44.1, 0.4, 0.11], null, null],
    "三重底剛形成 ＋ 強勢突破盤 ＋ 晨星剛形成": [[30, 50.0, 4.33, 0.77], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[517, 45.6, 2.07, 1.73], [100, 47.0, 3.17, 1.13], [59, 47.5, 1.11, -0.34], [55, 40.0, -2.42, -0.96]],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[654, 45.1, 2.19, 2.52], [187, 49.7, 2.12, 2.18], null, null],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 頭肩底剛形成": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 晨星剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], null, [54, 51.9, 5.91, 1.84], null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為正": [[437, 55.4, 5.29, 2.95], [34, 44.1, -1.97, 0.06], null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[102, 45.1, 0.89, 0.54], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[102, 35.3, -2.98, -0.77], [51, 47.1, 0.65, 0.17], null, [44, 47.7, 3.5, 0.42]],
    "KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[118, 46.6, 2.6, 2.12], null, null, null],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[94, 50.0, 3.71, 1.72], null, null, null],
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], null, null],
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能區間高檔(≥90百分位)": [[178, 41.0, 0.37, -0.07], [51, 49.0, 4.48, 1.2], null, [40, 45.0, 3.32, 0.6]],
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 晨星剛形成": [null, null, null, null],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[63, 46.0, 3.42, 1.21], null, null, null],
    "K線橫盤的突破剛形成 ＋ 回後買上漲全通過 ＋ 晨星剛形成": [[90, 42.2, 3.19, 1.31], null, null, null],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[289, 45.0, 1.08, 1.06], [78, 62.8, 7.44, 2.58], null, [57, 56.1, 4.41, 1.19]],
    "多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "夜星剛形成 ＋ 大盤跌破20日均線": [null, [214, 51.4, 2.28, 1.51], null, [155, 48.4, 2.66, 0.74]],
    "型態成形中 ＋ 多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "型態突破確認 ＋ 多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "型態剛形成(剛突破) ＋ 多方力道≥65 ＋ 晨星剛形成": [[566, 45.1, 2.18, 1.39], [110, 37.3, -0.74, 0.08], [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "MACD近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": [[76, 40.8, 3.99, 1.85], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[101, 49.5, 4.57, 1.92], null, null, null],
    "大盤跌破60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": [null, null, [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "型態成形中 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "型態突破確認 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "型態剛形成(剛突破) ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[629, 45.3, 1.86, 1.31], [103, 48.5, 3.46, 1.07], [62, 48.4, 1.08, -0.28], [57, 38.6, -2.47, -0.9]],
    "KDJ近3日內死亡交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[50, 54.0, 3.26, 1.42], null, null, null],
    "型態成形中 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [214, 51.4, 2.28, 1.51], null, [155, 48.4, 2.66, 0.74]],
    "型態突破確認 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [214, 51.4, 2.28, 1.51], null, [155, 48.4, 2.66, 0.74]],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [214, 51.4, 2.28, 1.51], null, [155, 48.4, 2.66, 0.74]],
    "N字底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [null, null, null, null],
    "圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], null, null, null],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [131, 48.9, 0.75, 0.89], null, [86, 37.2, 1.89, 0.72]],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 量能區間高檔(≥90百分位)": [[151, 38.4, -0.71, 0.34], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[305, 45.6, 2.92, 2.42], [59, 45.8, 2.05, 0.87], [34, 47.1, 1.16, -0.45], null],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成": [[75, 54.7, 3.01, 1.3], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[101, 40.6, -0.36, -0.22], [36, 58.3, 8.25, 1.85], null, [42, 71.4, 4.8, 1.09]],
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, null, null],
    "三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[96, 44.8, -0.05, -0.52], [37, 51.4, 5.09, 1.15], null, [55, 56.4, 4.96, 0.02]],
    "K線橫盤的突破剛形成 ＋ 圓弧底剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 大盤跌破20日均線": [null, [214, 51.4, 2.28, 1.51], null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[799, 44.8, 2.19, 1.65], [228, 47.8, 1.21, 2.2], null, null],
    "大盤跌破60日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, null, [61, 39.3, -0.29, 0.28], [44, 31.8, 1.33, 0.52]],
    "圓弧底剛形成 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], null, null, null],
    "圓弧底剛形成 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], null, null, null],
    "圓弧底剛形成 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": [[44, 47.7, 2.74, 0.44], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[61, 52.5, 4.3, 1.41], null, null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[148, 41.9, 3.72, 2.24], [35, 37.1, -1.36, -0.75], null, null],
    "多方力道≥65 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[94, 41.5, -0.56, -1.13], [32, 43.8, 6.44, 1.2], null, [34, 55.9, 6.13, 0.54]],
    "多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": [[273, 51.6, 2.6, 2.8], [50, 40.0, 5.22, 1.19], null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[133, 43.6, 2.15, 0.76], null, null, null],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[283, 49.1, 3.33, 2.05], [38, 71.1, 14.55, 3.28], null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[204, 50.0, 1.51, 1.25], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[69, 44.9, 3.12, 0.43], null, null, [41, 53.7, 1.21, 0.29]],
    "大盤站上20日均線 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], null, null, null],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[197, 49.2, 3.56, 1.94], [48, 60.4, 5.7, 2.85], null, [64, 79.7, 1.95, 0.96]],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[32, 53.1, 6.77, 1.76], null, null, null],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], null, null, null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], [41, 73.2, 1.85, 0.6], null],
    "多方力道≥65 ＋ 強勢突破盤 ＋ 晨星剛形成": [[240, 40.8, 1.46, 1.12], [45, 33.3, -3.36, -2.61], null, null],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, [43, 51.2, 2.79, 0.57], [86, 37.2, 1.89, 0.72]],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[187, 47.6, 1.76, 1.04], null, null, null],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 多方力道≥80": [[536, 47.6, 3.14, 2.71], [79, 39.2, -1.83, -0.45], [53, 58.5, 3.0, 1.37], [51, 43.1, 4.35, 1.72]],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[334, 49.1, 2.4, 1.51], [46, 58.7, 10.67, 2.69], null, null],
    "KDJ近3日內黃金交叉 ＋ 回後買上漲全通過 ＋ 地量(≤0.5倍均量)": [[133, 45.1, 2.11, 0.83], null, null, null],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[296, 40.9, 0.79, 1.21], [93, 52.7, 3.72, 1.7], null, [81, 55.6, 4.6, 1.88]],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 型態成形中": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], [41, 73.2, 1.85, 0.6], null],
    "回後買上漲全通過 ＋ 地量(≤0.5倍均量) ＋ 型態突破確認": [[240, 46.7, 1.84, 1.15], [38, 42.1, 5.43, 0.5], [41, 73.2, 1.85, 0.6], null],
    "晨星剛形成 ＋ 近3月均價YoY為負 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[71, 46.5, 1.7, 0.95], null, null, null],
    "回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[121, 52.1, 4.06, 2.07], null, null, null],
    "強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[354, 44.1, 1.69, 1.36], [37, 48.6, 7.67, 1.46], null, null],
    "布林通道高檔(≥80%) ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], [74, 50.0, 0.28, 0.69], null, null],
    "型態成形中 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "型態突破確認 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "型態剛形成(剛突破) ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], [37, 45.9, 1.33, -0.14], [30, 46.7, 0.6, 0.64]],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[62, 58.1, 1.7, 0.78], null, null, null],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[47, 51.1, 1.95, 0.81], null, null, null],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[89, 48.3, 6.45, 2.04], null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], null, null],
    "夜星剛形成 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[45, 42.2, 2.5, 1.18], null, null, null],
    "三大法人近3月買超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[725, 45.1, 1.96, 1.38], null, [66, 51.5, 1.8, -1.14], null],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], null, null, null],
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 近3月均價YoY為負": [[73, 46.6, 5.59, 1.94], null, null, null],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 母子懷抱(高檔)剛形成": [[91, 40.7, 0.8, 0.39], [35, 48.6, 4.19, 0.99], null, null],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 晨星剛形成": [[31, 32.3, -0.9, 0.34], null, null, null],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成": [[38, 39.5, -0.67, -0.5], null, null, null],
    "地量(≤0.5倍均量) ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], null, [60, 56.7, 2.04, 0.32]],
    "大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], null, [104, 46.2, 0.31, -1.53], null],
    "MACD近3日內黃金交叉 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[34, 41.2, -1.82, -0.09], null, null, null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能區間高檔(≥90百分位)": [[32, 28.1, -1.81, -0.26], null, null, null],
    "大盤站上60日均線 ＋ 強勢突破盤 ＋ 晨星剛形成": [[324, 41.4, 1.58, 0.94], [58, 41.4, -0.58, -0.83], null, null],
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[68, 52.9, 2.28, 0.73], null, null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 爆量(≥2倍均量)": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], null, [60, 56.7, 2.04, 0.32]],
    "地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], null, [60, 56.7, 2.04, 0.32]],
    "地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成": [[96, 45.8, 3.06, 0.95], [33, 36.4, -4.25, -0.78], null, [60, 56.7, 2.04, 0.32]],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近3月均價YoY為負": [null, null, null, null],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [null, null, null, null],
    "型態成形中 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], null, [104, 46.2, 0.31, -1.53], null],
    "型態突破確認 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], null, [104, 46.2, 0.31, -1.53], null],
    "型態剛形成(剛突破) ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], null, [104, 46.2, 0.31, -1.53], null],
    "三大法人近3月賣超 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[438, 43.2, 1.49, 0.91], null, [38, 36.8, -2.26, -2.09], null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)": [null, null, [36, 44.4, 1.85, -0.03], [59, 54.2, 4.86, 0.6]],
    "K線橫盤的突破剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[277, 41.5, 2.56, 1.78], [43, 58.1, 4.3, 0.9], [64, 43.8, 5.09, 1.43], [38, 44.7, 6.63, 1.94]],
    "夜星剛形成 ＋ 大盤跌破60日均線": [null, null, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[139, 41.7, 0.27, 0.05], null, null, null],
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 夜星剛形成": [[196, 42.3, 1.66, 1.67], [41, 51.2, 4.75, 1.52], null, null],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[116, 45.7, 2.79, 1.68], null, null, null],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%)": [[65, 49.2, -0.02, 0.1], null, null, null],
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[433, 44.8, 1.4, 0.73], [58, 46.6, 2.47, 1.39], [38, 55.3, 4.33, 0.36], [37, 40.5, -2.58, -0.89]],
    "KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成": [[103, 37.9, -1.7, -0.29], [30, 53.3, 1.34, -0.09], null, null],
    "地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[112, 38.4, 0.39, -0.04], [51, 52.9, 7.46, 1.54], null, [71, 53.5, 1.76, -0.54]],
    "大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [null, [35, 40.0, -0.43, 0.26], null, [33, 75.8, 2.91, 0.43]],
    "三大法人近3月買超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[725, 45.1, 1.96, 1.38], [181, 49.7, 1.04, 1.3], null, null],
    "型態成形中 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "型態突破確認 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, [80, 47.5, 1.07, -0.58], [155, 48.4, 2.66, 0.74]],
    "圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[119, 47.9, 1.18, 0.51], null, null, null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], [74, 50.0, 0.28, 0.69], [54, 51.9, 5.91, 1.84], [47, 46.8, 5.2, 2.69]],
    "多方力道≥80 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[770, 47.8, 3.37, 2.42], [108, 42.6, 0.26, -0.04], [73, 47.9, -0.14, 0.06], [50, 46.0, 2.36, 0.86]],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 強勢突破盤": [[585, 53.5, 5.02, 3.75], [74, 50.0, 0.28, 0.69], [54, 51.9, 5.91, 1.84], [47, 46.8, 5.2, 2.69]],
    "KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 強勢突破盤": [[194, 48.5, 3.96, 2.31], [32, 34.4, -1.67, 0.24], null, null],
    "大盤站上60日均線 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], null, null],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[74, 44.6, 1.75, 0.94], null, null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能區間高檔(≥90百分位)": [null, null, [38, 47.4, 0.39, 0.01], [70, 42.9, 2.7, 1.33]],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[377, 49.3, 2.5, 1.88], [85, 44.7, -0.12, 1.7], null, null],
    "三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[111, 48.6, 5.59, 2.57], null, null, null],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 晨星剛形成": [null, null, null, null],
    "多方力道≥65 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[271, 41.3, 0.44, 0.04], [84, 58.3, 4.47, 1.8], null, [48, 43.8, 3.1, 1.28]],
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 夜星剛形成": [[301, 40.5, 0.63, 0.52], [74, 52.7, 2.28, 1.39], null, [49, 44.9, 2.93, 1.12]],
    "N字底剛形成 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[83, 51.8, 3.62, 1.35], null, null, null],
    "多方力道≥80 ＋ 母子懷抱(高檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[49, 42.9, -1.19, -1.11], null, null, null],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[60, 41.7, -0.67, -0.06], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[112, 43.8, 0.89, 0.17], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 布林通道高檔(≥80%)": [null, null, null, null],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": [[51, 43.1, 0.92, 0.42], null, null, null],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[92, 39.1, -2.83, -0.93], null, null, null],
    "MACD近3日內黃金交叉 ＋ 三重底剛形成 ＋ 晨星剛形成": [null, null, null, null],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[41, 43.9, 0.66, 0.62], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價)": [null, [80, 53.8, 2.36, 1.06], null, [77, 41.6, 0.11, -0.27]],
    "圓弧底剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [null, null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], null, null],
    "KDJ近3日內黃金交叉 ＋ 多方力道≥65 ＋ 晨星剛形成": [[234, 44.4, 2.35, 1.23], [49, 36.7, 1.11, 0.4], null, [36, 47.2, -3.08, -2.01]],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成": [[104, 48.1, 4.7, 1.62], null, null, null],
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[106, 45.3, 1.37, 0.81], null, null, null],
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[321, 48.6, 2.99, 0.55], [47, 48.9, 4.4, 1.09], [32, 46.9, -0.99, -0.19], null],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成": [[102, 52.0, 5.83, 1.95], null, null, null],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成": [[66, 30.3, -2.84, -2.11], null, null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 相對強弱為正(強於大盤)": [null, null, [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[517, 46.8, 2.14, 1.72], [130, 50.0, 0.95, 0.98], [48, 52.1, 2.72, -0.67], [158, 44.9, -0.9, -0.37]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[133, 53.4, 5.52, 2.46], null, null, null],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], null, null],
    "晨星剛形成 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間高檔(≥90百分位)": [[249, 47.0, 3.28, 2.21], [57, 38.6, 1.15, 0.65], null, [39, 56.4, -3.85, -1.17]],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[411, 44.3, 2.04, 2.5], null, [47, 46.8, 1.56, -0.07], null],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[71, 53.5, 3.29, 1.4], null, null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[377, 49.3, 2.5, 1.88], null, [30, 50.0, 0.48, -0.34], null],
    "型態成形中 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], null, null],
    "型態突破確認 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], null, null],
    "型態剛形成(剛突破) ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[1165, 44.4, 1.78, 1.54], [312, 47.4, 1.06, 1.47], null, null],
    "大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [null, null, [188, 47.3, 2.8, 0.79], [138, 44.2, 3.51, 1.79]],
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成": [null, null, [90, 43.3, 3.53, 1.6], [86, 38.4, 4.45, 2.15]],
    "回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[139, 48.9, 2.47, 0.97], null, null, null],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為正": [[80, 58.8, 7.01, 2.22], null, null, null],
    "多方力道≥65 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[292, 46.9, 2.42, 2.13], [50, 52.0, 4.11, 1.44], [40, 37.5, 0.67, 0.66], null],
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[374, 46.5, 3.07, 2.62], [69, 44.9, 1.72, 0.74], [35, 42.9, 4.28, 1.0], [30, 30.0, 2.81, 0.74]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 複式頭肩底剛形成": [[32, 50.0, 1.59, 0.25], null, null, null],
    "多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[178, 41.6, 1.96, 1.1], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 強勢突破盤 ＋ 晨星剛形成": [[178, 41.0, 0.79, 0.74], [33, 42.4, -1.66, -0.87], null, null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 相對強弱為負(弱於大盤)": [[70, 31.4, 1.39, 0.63], null, null, null],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收)": [[419, 47.3, 2.95, 2.55], [124, 49.2, 2.3, 1.72], null, [124, 52.4, -0.8, 0.2]],
    "晨星剛形成 ＋ 爆量(≥2倍均量) ＋ 量能區間高檔(≥90百分位)": [[216, 45.8, 3.6, 2.36], [47, 40.4, 0.6, 0.28], null, [51, 52.9, -1.52, -1.84]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月乖離度為正(營收優於股價)": [[89, 49.4, 1.87, 0.92], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收)": [null, [118, 52.5, 2.18, 1.3], null, [59, 54.2, 4.86, 0.6]],
    "回後買上漲全通過 ＋ 母子懷抱(低檔)剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[45, 42.2, -0.01, -0.32], null, null, null],
    "夜星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 近3月均價YoY為正": [[93, 43.0, 0.26, -0.13], [30, 56.7, 9.43, 1.86], null, [31, 41.9, -1.54, -0.08]],
    "三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[157, 51.0, 2.56, 1.55], null, null, null],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 突破ABC修正下降切線剛形成": [[36, 33.3, -0.2, -0.18], null, null, null],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(高檔)剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[30, 33.3, 0.42, -0.2], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥65 ＋ 晨星剛形成": [[216, 41.7, 0.91, 1.55], [38, 34.2, -1.46, -0.3], null, [30, 46.7, -3.16, -1.82]],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成": [[267, 48.3, 2.91, 1.52], [46, 37.0, -1.1, -0.45], [34, 58.8, 5.21, 1.33], [34, 32.4, 5.24, 2.12]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[259, 52.9, 3.44, 1.94], [31, 45.2, 1.06, 0.81], null, null],
    "大盤跌破20日均線 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [null, [103, 48.5, 3.46, 1.07], null, [57, 38.6, -2.47, -0.9]],
    "N字底剛形成 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[30, 30.0, -0.19, 0.13], null, null, null],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 爆量(≥2倍均量)": [[56, 42.9, 0.43, 0.05], null, null, null],
    "回後買上漲全通過 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[189, 44.4, 2.08, 0.63], null, null, null],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成": [[90, 44.4, 5.51, 1.63], null, null, null],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [null, null, null, null],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], null, null, null],
    "三大法人近3月買超 ＋ 回後買上漲全通過 ＋ 多方力道≥80": [[2767, 47.9, 3.0, 4.52], [370, 43.0, -0.44, -0.42], [248, 57.3, 4.13, 3.5], [227, 40.1, 1.08, 0.85]],
    "回後買上漲全通過 ＋ 強勢突破盤 ＋ 晨星剛形成": [[145, 42.8, 2.46, 0.44], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成": [[295, 43.4, 0.9, 1.37], [49, 49.0, 2.35, 1.39], [34, 64.7, 6.45, 0.86], null],
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], [188, 50.0, 3.12, 1.5], null, null],
    "KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[158, 44.3, 1.88, 1.51], null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[274, 44.2, 2.56, 2.07], [64, 45.3, 1.08, 0.31], null, null],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 量能區間高檔(≥90百分位)": [[98, 44.9, 4.78, 1.58], null, null, null],
    "夜星剛形成 ＋ 大盤站上60日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], null, null, null],
    "N字底剛形成 ＋ 型態成形中 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], null, null, null],
    "N字底剛形成 ＋ 型態突破確認 ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], null, null, null],
    "N字底剛形成 ＋ 型態剛形成(剛突破) ＋ 突破ABC修正下降切線剛形成": [[70, 41.4, 1.1, 0.23], null, null, null],
    "K線橫盤的突破剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[95, 43.2, 1.25, 0.58], null, null, null],
    "晨星剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 突破ABC修正下降切線剛形成": [[47, 42.6, 1.56, 0.84], null, null, null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], null, null],
    "三大法人近3月買超 ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 多方力道≥80": [[623, 49.0, 3.33, 3.64], [76, 42.1, -0.09, 0.46], [61, 47.5, 0.69, 0.47], [53, 37.7, 3.36, 1.52]],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 量能區間高檔(≥90百分位)": [[601, 48.8, 3.18, 3.19], [92, 44.6, 0.01, 0.4], [57, 50.9, 1.92, 1.16], [65, 32.3, 1.76, 1.51]],
    "三大法人近3月買超 ＋ 強勢突破盤 ＋ 晨星剛形成": [[227, 39.2, 0.68, 0.57], [35, 37.1, -2.25, -0.82], null, null],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[410, 48.3, 2.32, 1.31], [65, 56.9, 6.6, 2.16], [39, 53.8, 0.77, -0.15], [36, 44.4, -2.33, -0.93]],
    "三大法人近3月買超 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[358, 47.2, 2.05, 1.49], [39, 61.5, 11.61, 2.12], null, null],
    "MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], null, null],
    "MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], null, null],
    "MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": [[107, 34.6, -0.74, -0.79], [36, 47.2, 1.78, 0.7], null, null],
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[116, 52.6, 3.92, 2.09], [30, 60.0, 3.88, 0.98], null, null],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成": [[48, 37.5, -0.86, -0.17], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 爆量(≥2倍均量)": [null, null, null, null],
    "夜星剛形成 ＋ 大盤跌破60日均線 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, null, [47, 48.9, -0.63, -1.22], [88, 48.9, -0.27, -0.18]],
    "晨星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[89, 40.4, 0.16, 0.01], null, null, [33, 66.7, 2.73, 0.27]],
    "三大法人近3月賣超 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[438, 43.2, 1.49, 0.91], [131, 44.3, 1.09, 0.78], null, null],
    "三大法人近3月買超 ＋ 型態成形中 ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "三大法人近3月買超 ＋ 型態突破確認 ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "三大法人近3月買超 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成": [[434, 41.7, 0.48, 0.83], [114, 55.3, 3.77, 1.89], [38, 55.3, 4.92, 0.7], [88, 55.7, 4.0, 1.68]],
    "K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[38, 47.4, 2.04, 0.98], null, null, null],
    "多方力道≥65 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": [null, [191, 47.6, 2.34, 0.71], null, [160, 43.1, 3.39, 1.2]],
    "圓弧底剛形成 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[32, 46.9, 1.68, 0.7], null, null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[233, 43.8, 1.9, 1.66], [42, 38.1, -1.97, -1.58], null, null],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上20日均線": [[3625, 46.7, 2.48, 4.03], null, [369, 56.4, 3.55, 3.74], null],
    "多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": [null, [90, 45.6, 2.71, 0.97], null, [67, 31.3, 2.26, 1.28]],
    "三大法人近3月買超 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [[254, 43.7, 1.5, 1.71], null, null, null],
    "強勢突破盤 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[71, 40.8, -3.26, -2.38], null, null, null],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[91, 46.2, 1.74, 1.08], null, null, null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[70, 35.7, 1.72, 0.39], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[80, 45.0, 2.04, 1.04], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 近3月均價YoY為正": [[110, 41.8, -0.62, 0.1], [43, 58.1, 4.55, 1.02], null, [30, 46.7, 0.89, -0.98]],
    "母子懷抱(高檔)剛形成 ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[97, 43.3, -0.02, -0.75], [46, 43.5, 4.63, 1.06], null, [60, 50.0, 2.91, -0.1]],
    "N字底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[119, 50.4, 4.47, 2.09], null, null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[274, 44.2, 2.56, 2.07], null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[411, 44.3, 2.04, 2.5], [88, 45.5, -0.91, -0.87], null, null],
    "夜星剛形成 ＋ 大盤站上20日均線 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[61, 45.9, 3.89, 1.86], null, null, null],
    "大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [null, null, [194, 48.5, 2.9, 0.4], [166, 36.1, 1.75, 1.14]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[407, 46.7, 2.98, 2.74], [100, 56.0, 2.85, 2.18], null, [165, 66.7, 0.41, -0.36]],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[62, 53.2, 3.66, 1.26], null, null, null],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "KDJ近3日內黃金交叉 ＋ 多方力道≥65 ＋ 夜星剛形成": [[109, 38.5, -1.66, -0.68], [40, 40.0, -1.25, -0.57], null, null],
    "三大法人近3月買超 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[486, 45.1, 2.12, 1.98], [56, 46.4, 6.1, 1.64], [45, 66.7, 6.54, 0.47], [50, 44.0, -1.83, -0.82]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月均價YoY為負": [[1247, 45.1, 2.69, 3.55], [171, 48.5, -0.82, -1.49], [249, 59.4, 4.88, 3.42], [139, 40.3, 3.28, 2.09]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[110, 50.0, 5.23, 2.74], null, null, null],
    "地量(≤0.5倍均量) ＋ 型態成形中 ＋ 強勢突破盤": [[515, 51.8, 4.29, 3.21], [70, 50.0, 0.55, 0.75], [51, 49.0, 5.84, 1.8], [44, 45.5, 5.06, 2.61]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[231, 45.5, 3.07, 2.12], [55, 43.6, 1.33, 0.21], null, [33, 30.3, -3.31, -1.27]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[412, 45.1, 2.27, 1.56], [68, 38.2, 0.65, 0.14], [42, 47.6, -1.81, -0.93], [42, 38.1, -1.87, -0.53]],
    "型態成形中 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "型態突破確認 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "型態剛形成(剛突破) ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[545, 48.4, 2.19, 1.48], [68, 58.8, 6.98, 2.21], [48, 50.0, -0.24, -0.48], [36, 44.4, -2.33, -0.93]],
    "N字底剛形成 ＋ 布林通道高檔(≥80%) ＋ 突破飆股大量黑K最高點剛形成": [[130, 50.0, 4.47, 1.89], [31, 41.9, 0.08, 0.14], null, null],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(高檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[38, 39.5, 1.91, 0.52], null, null, null],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[84, 50.0, 0.25, 0.23], null, null, null],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[109, 40.4, -1.0, -0.0], null, null, null],
    "KDJ近3日內死亡交叉 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成": [[131, 39.7, -0.76, 0.07], [52, 51.9, 0.5, -0.48], null, [42, 47.6, 1.83, -0.32]],
    "KDJ近3日內黃金交叉 ＋ 夜星剛形成 ＋ 量能區間高檔(≥90百分位)": [[71, 45.1, -0.13, 0.27], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[566, 45.8, 2.11, 2.18], null, [49, 51.0, 2.47, 0.26], null],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 大盤跌破20日均線": [null, [141, 47.5, 0.27, 0.97], null, [112, 48.2, 1.87, -0.19]],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月均價YoY為正": [[505, 46.3, 2.34, 1.2], [146, 47.9, 1.2, 2.23], [40, 45.0, -2.6, -1.49], [176, 46.6, -1.15, -0.67]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為負": [[82, 51.2, 3.32, 1.5], null, null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[556, 45.3, 1.82, 1.43], null, [54, 46.3, -2.06, -1.28], null],
    "大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成": [null, null, [278, 47.1, 1.78, 0.51], null],
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[486, 48.6, 4.11, 2.95], [95, 52.6, -0.89, -0.39], [87, 51.7, 4.25, 1.39], [60, 36.7, 1.1, 0.01]],
    "三大法人近3月買超 ＋ 晨星剛形成 ＋ 近3月均價YoY為正": [[532, 45.5, 2.45, 1.55], [132, 50.0, 0.83, 1.73], [32, 43.8, -4.35, -2.14], [226, 62.4, 0.46, -0.8]],
    "晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[361, 44.9, 2.54, 2.01], [50, 52.0, 7.37, 1.67], null, null],
    "多方力道≥80 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[124, 44.4, 2.24, 1.46], null, null, null],
    "三大法人近3月賣超 ＋ 多方力道≥80 ＋ 晨星剛形成": [[92, 31.5, -0.81, -0.2], null, null, null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[370, 43.5, 1.79, 1.59], [79, 46.8, 3.44, 0.81], [39, 51.3, 3.74, 0.52], [45, 37.8, -2.07, -0.86]],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成": [[41, 34.1, 2.86, 0.4], null, null, null],
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成": [null, null, null, null],
    "MACD近3日內黃金交叉 ＋ 夜星剛形成 ＋ 爆量(≥1.5倍均量)": [[40, 35.0, -2.63, -1.2], null, null, null],
    "MACD近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], null, null, null],
    "型態成形中 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "型態突破確認 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "型態剛形成(剛突破) ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[755, 46.6, 2.56, 1.9], [215, 47.0, 0.83, 0.95], [80, 46.2, 0.41, -1.46], [273, 46.5, -0.92, -0.8]],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收)": [[239, 38.1, -0.38, -0.04], [98, 51.0, 1.31, 0.92], null, [55, 54.5, 5.44, 0.77]],
    "圓弧底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[74, 32.4, 0.35, 0.43], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤站上20日均線": [[48, 45.8, 3.52, 1.47], null, null, null],
    "地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成": [null, [90, 51.1, 3.23, 0.82], null, [131, 57.3, 1.49, 0.05]],
    "夜星剛形成 ＋ 近3月乖離度為正(營收優於股價) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, null, null, null],
    "三重底剛形成 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[37, 62.2, 4.75, 1.52], null, null, null],
    "大盤跌破60日均線 ＋ 複式頭肩底剛形成 ＋ 頭肩底剛形成": [null, null, null, null],
    "K線橫盤的突破剛形成 ＋ 大盤跌破60日均線 ＋ 複式頭肩底剛形成": [null, null, null, null],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], null, null, null],
    "N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 布林通道高檔(≥80%)": [[2088, 49.3, 2.61, 4.48], [302, 47.0, 2.0, -0.25], [202, 53.5, 5.06, 2.03], [162, 40.1, 1.65, 1.58]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[96, 44.8, 2.37, 0.87], null, null, null],
    "多方力道≥65 ＋ 晨星剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[215, 42.8, 1.82, 1.73], [45, 37.8, -2.0, -0.39], null, null],
    "多方力道≥65 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [47, 48.9, 2.73, 0.06], [76, 51.3, -0.68, -0.02]],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為正": [[129, 42.6, 0.56, 0.68], null, null, null],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[126, 39.7, 1.16, -0.27], null, null, null],
    "MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], null, null, null],
    "MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], null, null, null],
    "MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成": [[133, 53.4, 1.72, 0.66], null, null, null],
    "KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 量能區間高檔(≥90百分位)": [[144, 45.8, 2.9, 1.42], null, null, null],
    "三重底剛形成 ＋ 晨星剛形成 ＋ 相對強弱為正(強於大盤)": [[35, 54.3, 5.05, 1.29], null, null, null],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 相對強弱為負(弱於大盤)": [[410, 40.2, 0.33, 0.59], [97, 48.5, 1.57, 1.15], null, null],
    "大盤站上20日均線 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[469, 42.4, 1.5, 1.45], null, [48, 52.1, 3.92, 0.74], null],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量) ＋ 近3月均價YoY為正": [[96, 50.0, 3.4, 1.25], null, null, null],
    "N字底剛形成 ＋ 爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成": [[89, 50.6, 4.07, 1.67], null, null, null],
    "晨星剛形成 ＋ 近3月均價YoY為正 ＋ 量能區間高檔(≥90百分位)": [[312, 46.2, 2.21, 1.41], [73, 39.7, 0.36, 0.81], null, [56, 57.1, -2.79, -1.8]],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, [95, 49.5, 0.77, 0.67], null, [88, 48.9, -0.27, -0.18]],
    "N字底剛形成 ＋ 型態成形中 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], null, null],
    "N字底剛形成 ＋ 型態突破確認 ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], null, null],
    "N字底剛形成 ＋ 型態剛形成(剛突破) ＋ 突破飆股大量黑K最高點剛形成": [[131, 49.6, 4.26, 1.78], [32, 43.8, 0.19, 0.17], null, null],
    "回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月均價YoY為負": [[408, 44.9, 2.26, 0.91], [58, 51.7, 0.17, 0.08], [83, 53.0, 5.26, 1.66], [56, 46.4, 2.82, 0.98]],
    "夜星剛形成 ＋ 母子懷抱(高檔)剛形成 ＋ 近3月均價YoY為正": [[114, 44.7, 1.14, 0.45], [32, 46.9, 5.71, 1.55], null, null],
    "夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], null, [33, 75.8, 2.91, 0.43]],
    "回後買上漲全通過 ＋ 多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[551, 45.9, 1.7, 0.92], [69, 50.7, 5.13, 1.03], [63, 63.5, 0.21, 0.21], [68, 52.9, 3.23, 1.82]],
    "多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[201, 40.8, 0.43, 0.85], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥80 ＋ 母子懷抱(低檔)剛形成": [[138, 35.5, -1.4, -0.94], null, null, null],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 相對強弱為正(強於大盤)": [[3439, 46.5, 2.56, 4.35], [492, 42.7, -0.87, -1.05], [366, 56.8, 3.66, 3.85], [310, 37.4, 0.86, 0.7]],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 爆量(≥1.5倍均量)": [[497, 46.3, 2.73, 2.57], [81, 35.8, -1.46, -0.35], [48, 52.1, 1.73, 0.82], [57, 38.6, 3.75, 1.9]],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 大盤站上60日均線": [[3625, 46.7, 2.48, 4.03], [492, 42.7, -0.87, -1.05], null, null],
    "回後買上漲全通過 ＋ 多方力道≥80 ＋ 近3月乖離度為正(營收優於股價)": [[1612, 47.8, 3.15, 4.07], [242, 48.8, -0.89, -0.64], [266, 58.6, 4.72, 2.67], [162, 35.8, 2.39, 1.69]],
    "多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[505, 47.1, 3.36, 2.69], [108, 46.3, -0.6, -0.86], [93, 48.4, 3.66, 0.63], [66, 30.3, 3.73, 1.79]],
    "強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[174, 42.5, 1.36, 0.7], null, null, null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3月均價YoY為負": [[148, 48.0, 4.21, 3.42], [34, 47.1, 2.97, 1.31], null, null],
    "MACD近3日內黃金交叉 ＋ 晨星剛形成 ＋ 爆量(≥2倍均量)": [[60, 50.0, 2.38, 0.96], null, null, null],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 夜星剛形成": [[41, 41.5, 2.63, 0.51], null, null, null],
    "型態成形中 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "型態突破確認 ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "型態剛形成(剛突破) ＋ 夜星剛形成 ＋ 相對強弱為正(強於大盤)": [[448, 40.6, 0.34, 0.59], [172, 50.0, 2.27, 1.32], [56, 48.2, 1.8, 0.29], [143, 47.6, 3.19, 0.91]],
    "母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量) ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], null, [33, 75.8, 2.91, 0.43]],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[337, 48.4, 2.27, 2.05], [65, 43.1, -0.2, -0.27], null, null],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 爆量(≥1.5倍均量)": [[177, 42.9, 2.13, 1.44], [54, 51.9, 3.58, 0.73], null, [85, 76.5, 0.28, 0.66]],
    "型態成形中 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], null, [33, 75.8, 2.91, 0.43]],
    "型態突破確認 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], null, [33, 75.8, 2.91, 0.43]],
    "型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[179, 50.3, 3.27, 1.97], [35, 40.0, -0.43, 0.26], null, [33, 75.8, 2.91, 0.43]],
    "晨星剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[279, 45.9, 2.56, 1.74], [53, 41.5, -1.09, -0.98], [30, 43.3, -2.7, -1.09], [54, 64.8, 0.73, 0.57]],
    "大盤站上60日均線 ＋ 晨星剛形成 ＋ 量能區間高檔(≥90百分位)": [[556, 45.3, 1.82, 1.43], [121, 40.5, -0.28, -0.34], null, null],
    "大盤跌破20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成": [null, [80, 57.5, 4.67, 2.07], null, [44, 31.8, 1.33, 0.52]],
    "回後買上漲全通過 ＋ 大盤跌破60日均線 ＋ 複式頭肩底剛形成": [null, null, [50, 52.0, 1.58, 1.0], null],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近3月乖離度為負(股價超前營收)": [[230, 43.5, -0.3, 0.58], [64, 59.4, 3.7, 1.81], null, [35, 62.9, 6.5, 0.97]],
    "三大法人近3月賣超 ＋ 晨星剛形成 ＋ 母子懷抱(低檔)剛形成": [[135, 51.1, 2.32, 1.25], [38, 47.4, 0.71, 1.79], null, [66, 78.8, -0.19, 0.03]],
    "N字底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [null, null, null, null],
    "三重底剛形成 ＋ 大盤跌破60日均線 ＋ 量能區間低檔(≤10百分位)": [null, null, null, null],
    "大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 複式頭肩底剛形成": [null, null, null, null],
    "大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 頭肩底剛形成": [null, null, null, null],
    "KDJ近3日內死亡交叉 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[272, 40.1, 0.29, 1.03], [36, 55.6, 9.39, 1.69], null, null],
    "MACD近3日內黃金交叉 ＋ 多方力道≥80 ＋ 量能區間低檔(≤10百分位)": [null, null, null, null],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 大盤站上20日均線": [[988, 48.8, 2.56, 3.72], null, [97, 49.5, 1.54, 1.49], null],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 多方力道≥80": [[765, 49.8, 3.04, 3.69], [117, 45.3, 1.45, 1.06], [61, 57.4, 4.52, 2.06], [79, 36.7, 2.18, 0.84]],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 近3月均價YoY為負": [[212, 45.8, 1.53, 0.33], [33, 45.5, 0.93, 0.46], [34, 47.1, 4.35, 0.55], [32, 43.8, 1.0, 0.74]],
    "三大法人近3月賣超 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[187, 50.8, 2.47, 1.36], null, null, null],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成": [[85, 44.7, 1.34, 0.59], null, null, null],
    "KDJ近3日內死亡交叉 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(高檔)剛形成": [[50, 38.0, 2.74, 1.36], null, null, null],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, null, [37, 54.1, 0.45, 0.04]],
    "MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 近3月均價YoY為負": [[48, 60.4, 4.63, 1.74], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[404, 41.8, 0.13, 1.01], null, [43, 60.5, 4.98, 0.86], null],
    "大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥1.5倍均量)": [[337, 48.4, 2.27, 2.05], null, [33, 54.5, 0.76, -1.0], null],
    "回後買上漲全通過 ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[92, 45.7, 1.64, 0.95], null, null, null],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, null, null],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成": [[137, 38.7, 0.41, 0.02], [55, 47.3, 3.27, 0.69], null, [67, 52.2, 2.15, -0.25]],
    "圓弧底剛形成 ＋ 量能區間高檔(≥90百分位) ＋ 頭肩底剛形成": [[33, 57.6, 4.37, 1.41], null, null, null],
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成": [[983, 46.5, 2.13, 2.2], null, [120, 50.8, 3.78, 1.72], null],
    "回後買上漲全通過 ＋ 大盤站上20日均線 ＋ 晨星剛形成": [[274, 44.5, 1.88, 0.56], null, [36, 44.4, -1.97, -1.38], null],
    "K線橫盤的突破剛形成 ＋ 大盤跌破20日均線 ＋ 晨星剛形成": [null, [34, 44.1, -0.53, -0.88], null, null],
    "回後買上漲全通過 ＋ 大盤站上60日均線 ＋ 晨星剛形成": [[274, 44.5, 1.88, 0.56], [46, 54.3, 2.98, 0.51], null, null],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 晨星剛形成": [[38, 36.8, 0.06, 0.52], null, null, null],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月乖離度為正(營收優於股價)": [[47, 44.7, 3.42, 0.97], null, null, null],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": [[329, 56.8, 6.99, 3.27], null, null, null],
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 月營收連3月年增": [null, null, [3867, 57.6, 2.18, 10.41], [6203, 62.4, 2.13, 6.18]],
    "月營收創24個月新高 ＋ 月營收連3月年增 ＋ 近3日漲停": [[2408, 55.0, 6.0, 8.29], [497, 51.1, 3.68, -0.34], [247, 58.3, 3.22, 1.55], [561, 76.3, 5.6, 2.3]],
    "均線多頭排列(5>20>60且站上月線) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[554, 60.5, 2.6, 4.93], [78, 55.1, 0.82, 0.98], [66, 53.0, 1.54, 1.23], [54, 42.6, 4.62, 2.02]],
    "月營收創24個月新高 ＋ 近3日向上跳空缺口 ＋ 近3日漲停": [[1828, 54.0, 5.57, 6.11], [327, 48.6, 2.08, 0.08], [160, 60.0, 6.07, 2.42], [430, 83.3, 5.9, 1.61]],
    "月營收創24個月新高 ＋ 近3日漲停": [[3172, 53.8, 5.74, 7.08], [625, 48.8, 2.35, -0.78], [283, 57.2, 4.74, 2.37], [701, 76.5, 4.58, 1.35]],
    "漲時量≥跌時量1.5倍(近20日) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[353, 53.8, 1.74, 2.96], [49, 59.2, 0.0, 0.44], [61, 54.1, 0.74, 0.42], [43, 37.2, 2.97, 1.45]],
    "三大法人近3月買超 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[2244, 55.1, 6.58, 5.95], [447, 46.8, 2.09, -0.01], [162, 54.9, 4.9, 2.39], [414, 74.2, 4.98, 2.21]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[1446, 57.6, 8.08, 7.78], [199, 39.7, -0.41, -2.48], [128, 57.8, 5.29, 2.26], [89, 50.6, 0.51, -0.48]],
    "外資近5日買超 ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[475, 60.4, 2.08, 4.64], [79, 55.7, 0.3, 0.65], [66, 50.0, 1.05, 0.39], [47, 42.6, 5.44, 3.24]],
    "三大法人近3月買超 ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[652, 56.9, 1.92, 4.09], [102, 56.9, 1.22, 1.24], [87, 52.9, 1.57, 1.85], [68, 45.6, 2.7, 1.74]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1235, 57.4, 8.26, 6.49], [117, 38.5, 0.82, -1.53], [84, 59.5, 6.79, 2.96], [39, 56.4, 9.29, 2.37]],
    "距52週高點≤5% ＋ 近3月均價YoY為正 ＋ 量能區間低檔(≤10百分位)": [[742, 57.8, 1.88, 4.61], [123, 54.5, 0.66, 0.8], [85, 54.1, 1.91, 2.08], [91, 47.3, 3.71, 2.47]],
    "創52週新高 ＋ 強勢突破盤 ＋ 月營收創24個月新高": [[1679, 56.7, 6.66, 6.44], [183, 44.8, -0.08, -0.48], [118, 55.1, 4.15, 2.55], [97, 46.4, 1.95, 1.17]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 短線過熱(高於月線≥20%)": [[263, 63.1, 11.25, 4.81], null, null, null],
    "月營收創24個月新高 ＋ 月營收連3月年增 ＋ 近3日向上跳空缺口": [[3735, 52.1, 3.41, 5.96], [805, 50.6, 1.73, 0.5], [572, 63.1, 3.68, 3.57], [873, 73.2, 4.34, 2.56]],
    "距52週高點≤5% ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": [[498, 58.8, 2.53, 4.36], [82, 59.8, 1.48, 1.15], [57, 47.4, 1.11, 0.95], [44, 36.4, 2.21, 0.99]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3日漲停": [[183, 59.0, 11.14, 3.73], null, null, null],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為負": [[90, 61.1, 6.92, 2.9], null, null, null],
    "外資近5日買超 ＋ 大盤跌破60日均線 ＋ 月營收創24個月新高": [null, null, [2838, 56.6, 1.75, 5.33], [3932, 61.1, 1.94, 5.78]],
    "多方力道≥80 ＋ 月營收連3月年增 ＋ 量能區間低檔(≤10百分位)": [[272, 56.2, 3.46, 3.62], null, null, null],
    "回後買上漲全通過 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[495, 56.4, 6.45, 4.52], [41, 41.5, 2.51, -0.88], null, null],
    "多方力道≥65 ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[354, 58.5, 2.25, 3.83], [38, 65.8, 4.22, 2.06], [41, 56.1, 0.72, 0.19], [30, 53.3, 1.13, -0.08]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近20日強於大盤≥10%": [[306, 59.5, 9.21, 4.25], [37, 48.6, -1.93, -0.33], null, [31, 45.2, 7.03, 2.54]],
    "距52週高點≤5% ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[129, 67.4, 5.36, 3.0], null, null, null],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[244, 54.9, 4.61, 2.94], [51, 62.7, 10.08, 1.62], [55, 65.5, 13.53, 3.21], null],
    "外資近5日買超 ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[308, 53.9, 2.04, 2.65], [123, 55.3, 6.01, 2.63], [149, 58.4, 1.0, 2.21], [1102, 55.5, 0.84, 0.64]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[105, 50.5, 7.35, 2.44], null, null, null],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 近3日向上跳空缺口": [[252, 59.5, 10.02, 3.73], null, null, null],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 短線過熱(高於月線≥20%)": [[115, 49.6, 6.08, 2.34], null, null, null],
    "月營收創24個月新高 ＋ 距52週高點≤5% ＋ 近3日漲停": [[1377, 56.4, 7.26, 6.47], [164, 39.6, -1.07, -1.18], [86, 54.7, 5.99, 1.83], [53, 58.5, 4.8, 1.41]],
    "N字底剛形成 ＋ 月營收連3月年增 ＋ 短線過熱(高於月線≥20%)": [[129, 51.2, 5.65, 1.7], null, null, null],
    "地量(≤0.5倍均量) ＋ 近20日強於大盤≥10% ＋ 近3日漲停": [[549, 52.8, 5.23, 3.07], [138, 59.4, 9.71, 1.69], [31, 61.3, 8.63, 2.2], [78, 57.7, 1.74, 0.3]],
    "K線橫盤的突破剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[152, 54.6, 5.97, 3.54], [30, 63.3, 7.42, 1.67], null, null],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[143, 60.8, 9.79, 2.93], null, null, null],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 地量(≤0.5倍均量)": [[49, 69.4, 13.89, 3.4], null, null, null],
    "月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[188, 59.6, 9.62, 4.22], null, null, null],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[439, 53.5, 6.01, 2.7], [92, 59.8, 12.21, 2.34], null, [51, 58.8, -2.89, -1.19]],
    "地量(≤0.5倍均量) ＋ 相對強弱為正(強於大盤) ＋ 近3日漲停": [[616, 52.3, 5.08, 2.98], [163, 59.5, 9.59, 2.4], [35, 57.1, 6.82, 1.82], [159, 66.0, 0.58, -0.0]],
    "地量(≤0.5倍均量) ＋ 多方力道≥65 ＋ 近3日漲停": [[518, 53.7, 6.22, 3.18], [94, 54.3, 6.27, 1.44], null, [45, 60.0, 2.7, 0.88]],
    "K線橫盤的突破剛形成 ＋ 多方力道≥80 ＋ 近3日漲停": [[305, 49.5, 5.68, 4.2], [53, 43.4, -0.17, 0.05], [42, 50.0, 6.46, 1.18], [31, 35.5, -0.28, -0.28]],
    "N字底剛形成 ＋ OBV能量潮創60日新高 ＋ 近3日漲停": [[215, 47.9, 7.02, 3.23], [30, 40.0, -0.54, 0.12], [36, 50.0, 2.36, -0.21], null],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], null, [31, 45.2, 1.21, 0.51], null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[442, 54.1, 7.27, 2.9], [78, 55.1, 8.25, 1.59], null, [36, 61.1, 5.0, 1.34]],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[247, 57.9, 7.86, 3.39], null, null, null],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 月營收創24個月新高": [[219, 65.8, 11.17, 4.88], null, null, null],
    "N字底剛形成 ＋ 創52週新高 ＋ 短線過熱(高於月線≥20%)": [[201, 49.3, 5.28, 2.06], null, null, null],
    "CMF資金流買方佔優(近20日≥0.1) ＋ N字底剛形成 ＋ 近3日漲停": [[255, 45.1, 4.58, 2.61], [34, 47.1, -0.5, -0.01], [34, 47.1, -1.83, -0.27], null],
    "K線橫盤的突破剛形成 ＋ 創52週新高 ＋ 近3日漲停": [[247, 51.4, 4.65, 3.47], [41, 53.7, 3.36, 0.75], null, null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成 ＋ 近3日漲停": [[364, 49.5, 4.93, 4.46], [67, 44.8, 0.43, -0.02], [53, 54.7, 6.78, 1.42], [39, 41.0, -0.22, -0.2]],
    "N字底剛形成 ＋ 近3日漲停 ＋ 量能區間高檔(≥90百分位)": [[287, 44.6, 4.84, 2.97], [44, 38.6, -0.74, 0.29], [41, 61.0, 5.99, 1.47], null],
    "地量(≤0.5倍均量) ＋ 布林通道高檔(≥80%) ＋ 近3日漲停": [[418, 53.1, 6.45, 3.25], [47, 48.9, 2.67, 0.6], null, null],
    "K線橫盤的突破剛形成 ＋ 距52週高點≤5% ＋ 近3日漲停": [[383, 50.9, 4.81, 4.19], [70, 51.4, 3.19, 1.12], [37, 51.4, 5.83, 0.75], null],
    "N字底剛形成 ＋ 創52週新高 ＋ 近3日漲停": [[219, 49.8, 7.06, 3.11], [36, 47.2, 0.43, 0.28], null, null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成 ＋ 近3日漲停": [[265, 45.3, 4.06, 2.29], [40, 40.0, -1.72, -0.11], [33, 57.6, 5.7, 1.34], null],
    "相對強弱為負(弱於大盤) ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[122, 51.6, 3.96, 1.45], null, null, null],
    "地量(≤0.5倍均量) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[429, 59.0, 8.77, 4.61], [35, 31.4, -0.55, -0.22], null, null],
    "N字底剛形成 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[129, 48.8, 6.92, 2.35], null, null, null],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 短線過熱(高於月線≥20%)": [[225, 46.2, 3.66, 1.72], [34, 44.1, 1.7, 0.35], null, null],
    "N字底剛形成 ＋ 近3日漲停 ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[204, 45.1, 4.51, 2.59], [33, 39.4, 0.61, 0.45], [37, 56.8, 3.77, 0.64], null],
    "N字底剛形成 ＋ 大盤站上20日均線 ＋ 近3日漲停": [[398, 45.5, 4.12, 3.18], null, [65, 55.4, 2.78, 1.01], null],
    "N字底剛形成 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停": [[306, 45.8, 4.27, 2.79], [41, 39.0, -1.69, -0.36], [50, 54.0, 3.95, 0.86], null],
    "N字底剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%)": [[250, 47.2, 4.53, 1.88], [33, 48.5, 2.61, 0.59], null, null],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 短線過熱(高於月線≥20%)": [[162, 53.1, 8.45, 3.08], null, null, null],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[185, 47.6, 5.37, 1.86], null, null, null],
    "N字底剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停": [[308, 46.1, 4.83, 2.69], [43, 44.2, 0.01, 0.34], [37, 48.6, 0.6, 0.07], null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ N字底剛形成 ＋ 短線過熱(高於月線≥20%)": [[238, 46.2, 3.87, 1.67], [35, 45.7, 1.7, 0.36], null, null],
    "圓弧底剛形成 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[93, 59.1, 7.66, 2.61], null, null, null],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[285, 56.5, 5.56, 2.7], null, null, null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[353, 56.1, 6.83, 3.49], [32, 25.0, 0.07, -0.13], null, null],
    "多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[265, 54.7, 7.53, 4.45], [44, 59.1, 6.58, 1.65], null, null],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[219, 48.9, 5.41, 1.92], [31, 51.6, 3.17, 0.48], null, null],
    "N字底剛形成 ＋ 多方力道≥80 ＋ 近3日漲停": [[240, 45.8, 4.7, 2.47], [36, 33.3, -3.87, -0.95], null, null],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 短線過熱(高於月線≥20%)": [[254, 46.9, 4.18, 1.78], [35, 45.7, 1.7, 0.36], null, null],
    "N字底剛形成 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": [[238, 47.5, 4.06, 1.96], [32, 43.8, 1.24, 0.33], [30, 43.3, 1.05, 0.48], null],
    "N字底剛形成 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[70, 52.9, 5.0, 1.13], null, null, null],
    "N字底剛形成 ＋ 多方力道≥65 ＋ 近3日漲停": [[313, 45.4, 4.54, 3.01], [42, 38.1, -2.02, -0.29], [36, 55.6, 5.24, 1.16], null],
    "MACD近3日內黃金交叉 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[167, 53.3, 5.43, 2.32], null, null, null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[118, 67.8, 8.62, 3.5], null, null, null],
    "地量(≤0.5倍均量) ＋ 多方力道≥80 ＋ 近3日漲停": [[410, 53.9, 7.05, 3.05], [67, 53.7, 5.55, 1.21], null, null],
    "N字底剛形成 ＋ 外資近5日買超 ＋ 近3日漲停": [[270, 45.6, 4.64, 2.86], [41, 43.9, -0.38, -0.03], [42, 61.9, 6.38, 1.47], null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[1325, 55.1, 7.3, 5.82], [142, 36.6, 0.33, -1.16], [77, 51.9, 4.31, 1.81], [31, 48.4, 1.91, 0.88]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[84, 58.3, 6.89, 2.47], null, null, null],
    "N字底剛形成 ＋ 強勢突破盤 ＋ 近3日漲停": [[223, 46.6, 5.47, 2.81], [37, 45.9, 0.51, 0.68], null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 月營收創24個月新高": [null, [40, 50.0, 4.49, 1.88], null, null],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3日向上跳空缺口": [[472, 55.3, 6.62, 3.49], [49, 42.9, 6.88, 0.76], null, null],
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 近3日漲停": [[564, 52.0, 5.12, 3.24], [110, 58.2, 9.47, 1.51], [31, 61.3, 8.63, 2.2], [43, 58.1, 1.43, 0.6]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[108, 44.4, 6.15, 1.97], null, null, null],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], [36, 47.2, 2.33, 0.59], [31, 45.2, 1.21, 0.51], null],
    "N字底剛形成 ＋ 布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], [35, 48.6, 2.26, 0.54], [31, 45.2, 1.21, 0.51], null],
    "大盤跌破20日均線 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [null, [211, 34.6, -1.12, -1.71], null, [76, 51.3, 5.35, 2.12]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10%": [[261, 47.5, 3.98, 1.62], [36, 47.2, 2.33, 0.59], [31, 45.2, 1.21, 0.51], null],
    "N字底剛形成 ＋ 近20日強於大盤≥10% ＋ 近3日漲停": [[307, 46.3, 4.29, 2.46], [45, 44.4, 0.05, -0.01], [43, 55.8, 3.7, 0.95], null],
    "三大法人近3月買超 ＋ 地量(≤0.5倍均量) ＋ 近3日漲停": [[577, 51.1, 4.37, 2.65], [192, 64.6, 11.6, 2.82], [36, 66.7, 10.26, 2.59], [278, 78.4, 3.63, 0.84]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[323, 53.6, 6.43, 4.47], [52, 61.5, 9.65, 2.46], [30, 36.7, 6.62, 1.31], null],
    "N字底剛形成 ＋ 三大法人近3月買超 ＋ 短線過熱(高於月線≥20%)": [[216, 49.1, 5.56, 2.13], null, null, null],
    "地量(≤0.5倍均量) ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 近3日漲停": [[487, 52.8, 6.14, 2.95], [98, 57.1, 6.6, 1.82], null, [77, 61.0, -1.97, -0.89]],
    "月營收連3月年增 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[370, 55.4, 6.15, 4.27], [65, 61.5, 6.92, 1.96], [43, 44.2, 2.0, -0.03], null],
    "N字底剛形成 ＋ 布林通道高檔(≥80%) ＋ 近3日漲停": [[384, 45.8, 4.35, 3.07], [52, 46.2, -0.14, 0.22], [58, 55.2, 2.92, 0.61], null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[1423, 56.0, 6.52, 5.97], [124, 38.7, 0.46, -0.96], [74, 54.1, 7.03, 2.79], [48, 56.2, 11.39, 2.85]],
    "地量(≤0.5倍均量) ＋ 強勢突破盤 ＋ 月營收創24個月新高": [[172, 66.3, 14.11, 4.52], null, null, null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": [[56, 46.4, 5.81, 1.93], null, null, null],
    "N字底剛形成 ＋ 距52週高點≤5% ＋ 近3日漲停": [[250, 50.4, 6.73, 3.37], [39, 46.2, 0.42, 0.05], null, null],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[177, 47.5, 5.84, 2.22], null, null, null],
    "N字底剛形成 ＋ 回後買上漲全通過 ＋ 短線過熱(高於月線≥20%)": [[207, 48.8, 5.64, 2.24], null, null, null],
    "地量(≤0.5倍均量) ＋ 距52週高點≤5% ＋ 近3日漲停": [[303, 55.8, 7.32, 3.9], null, null, null],
    "N字底剛形成 ＋ 大盤站上60日均線 ＋ 短線過熱(高於月線≥20%)": [[268, 48.1, 4.56, 1.91], [36, 47.2, 2.33, 0.59], null, null],
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [null, null, [148, 50.7, 3.8, 2.18], [76, 51.3, 5.35, 2.12]],
    "地量(≤0.5倍均量) ＋ 月營收創24個月新高 ＋ 近3日漲停": [[252, 61.1, 10.57, 4.28], [81, 70.4, 12.81, 1.22], null, [90, 87.8, 7.87, -0.37]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": [[49, 55.1, 8.23, 1.12], null, null, null],
    "創52週新高 ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[51, 72.5, 7.4, 2.83], null, null, null],
    "漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[53, 69.8, 6.31, 2.14], null, null, null],
    "晨星剛形成 ＋ 月營收創24個月新高 ＋ 距52週高點≤5%": [[89, 55.1, 4.86, 1.89], null, null, null],
    "N字底剛形成 ＋ 近3日漲停 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[118, 37.3, 2.63, 1.58], null, null, null],
    "月營收創24個月新高 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[261, 57.5, 7.94, 3.73], [75, 56.0, 8.09, 1.38], null, [77, 89.6, 10.7, 1.82]],
    "均線多頭排列(5>20>60且站上月線) ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[101, 41.6, 1.72, 0.54], null, null, null],
    "強勢突破盤 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[271, 54.6, 6.24, 2.78], [48, 52.1, 4.68, 1.31], null, null],
    "N字底剛形成 ＋ 月營收連3月年增 ＋ 近3日漲停": [[182, 48.9, 5.04, 2.3], null, null, null],
    "K線橫盤的突破剛形成 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[47, 68.1, 12.74, 3.15], null, null, null],
    "月營收創24個月新高 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [[690, 51.2, 7.01, 4.26], [80, 36.2, -1.77, -0.84], [61, 49.2, 5.97, 2.16], [45, 60.0, 12.5, 3.38]],
    "晨星剛形成 ＋ 距52週高點≤5% ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[43, 62.8, 4.91, 1.27], null, null, null],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成": [[45, 51.1, 7.92, 1.08], null, null, null],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1306, 56.0, 6.83, 6.11], [116, 38.8, 0.45, -1.06], [64, 54.7, 7.83, 3.07], [45, 55.6, 8.89, 2.5]],
    "多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[82, 45.1, 3.03, 0.92], null, null, null],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[74, 59.5, 9.0, 3.45], null, null, null],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[92, 44.6, 0.6, -0.24], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[533, 53.8, 5.49, 3.26], [77, 35.1, -0.54, -2.32], null, [32, 50.0, 2.05, 0.65]],
    "月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日向上跳空缺口": [[214, 54.2, 5.7, 3.32], null, null, null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[120, 51.7, 4.63, 1.88], null, null, null],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": [[217, 59.4, 8.95, 4.25], null, null, null],
    "MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[450, 43.6, 1.27, 1.98], [68, 45.6, 0.18, -0.6], [37, 51.4, 3.88, 1.24], [36, 47.2, 0.37, -0.12]],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 漲時量≥跌時量1.5倍(近20日)": [null, null, null, null],
    "短線過熱(高於月線≥20%) ＋ 距52週高點≤5% ＋ 量能區間低檔(≤10百分位)": [[49, 75.5, 7.95, 3.03], null, null, null],
    "創52週新高 ＋ 多方力道≥65 ＋ 量能區間低檔(≤10百分位)": [[50, 68.0, 7.0, 2.72], null, null, null],
    "大盤站上60日均線 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[59, 71.2, 5.89, 2.21], null, null, null],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[46, 73.9, 7.6, 2.94], null, null, null],
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[57, 73.7, 6.28, 2.26], null, null, null],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為正 ＋ 量能區間低檔(≤10百分位)": [[58, 72.4, 6.51, 2.51], null, null, null],
    "三大法人近3月買超 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[57, 70.2, 5.16, 1.91], null, null, null],
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 量能區間高檔(≥90百分位)": [[87, 37.9, -2.94, -0.63], null, null, null],
    "夜星剛形成 ＋ 月營收創24個月新高 ＋ 近3月均價YoY為正": [[110, 52.7, 4.85, 2.17], [34, 52.9, 5.82, 2.04], null, null],
    "OBV能量潮創60日新高 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1118, 55.0, 6.86, 5.93], [98, 34.7, -0.54, -0.73], [81, 48.1, 3.3, 1.72], [45, 55.6, 8.57, 2.33]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[455, 48.1, 4.7, 3.44], [65, 44.6, -0.87, -0.96], [39, 53.8, 3.1, 0.81], [37, 48.6, 0.7, -0.0]],
    "夜星剛形成 ＋ 月營收創24個月新高 ＋ 月營收連3月年增": [[112, 55.4, 5.65, 2.57], [33, 48.5, 3.76, 1.25], null, null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": [[1091, 60.3, 8.73, 5.24], [91, 34.1, 0.75, -1.78], [56, 53.6, 1.77, -0.98], null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 大盤跌破60日均線 ＋ 晨星剛形成": [null, null, [40, 52.5, 4.02, 0.6], [43, 41.9, -0.04, 0.62]],
    "晨星剛形成 ＋ 近20日強於大盤≥10% ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[115, 54.8, 5.63, 2.37], [31, 51.6, 4.12, 0.92], null, null],
    "創52週新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3月均價YoY為負": [[140, 55.0, 4.12, 1.86], [32, 43.8, -1.78, -0.94], null, null],
    "KDJ近3日內黃金交叉 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[544, 54.8, 3.7, 2.88], [50, 32.0, -2.0, -1.68], null, [32, 65.6, 13.56, 3.22]],
    "地量(≤0.5倍均量) ＋ 月營收連3月年增 ＋ 近3日漲停": [[377, 50.9, 3.02, 2.06], [143, 69.9, 13.71, 1.64], null, [218, 83.5, 4.97, 0.37]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[940, 52.8, 7.08, 5.06], [106, 38.7, -0.1, -1.15], [85, 52.9, 4.02, 2.3], [44, 56.8, 9.76, 2.66]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[168, 48.2, 3.18, 2.17], [32, 75.0, 13.66, 1.67], null, null],
    "多方力道≥65 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1965, 54.2, 5.89, 5.25], [205, 35.1, -1.21, -1.82], [122, 44.3, 2.46, 1.56], [72, 50.0, 5.0, 2.01]],
    "MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 短線過熱(高於月線≥20%)": [null, null, [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "N字底剛形成 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[96, 52.1, 5.66, 2.01], null, null, null],
    "均線多頭排列(5>20>60且站上月線) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1959, 53.6, 6.07, 5.37], [207, 34.8, -1.01, -1.54], [112, 43.8, 2.75, 1.91], [73, 50.7, 4.91, 1.9]],
    "多方力道≥80 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1626, 53.0, 5.78, 4.21], [183, 34.4, -1.67, -2.6], [89, 36.0, 1.08, 1.22], [67, 47.8, 2.89, 1.33]],
    "晨星剛形成 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[136, 52.9, 5.9, 2.46], [30, 50.0, 2.66, 1.02], null, null],
    "月營收創24個月新高 ＋ 爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[972, 52.3, 6.16, 4.72], [110, 39.1, 1.22, -0.33], [84, 52.4, 4.43, 1.99], [57, 59.6, 10.19, 3.25]],
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[536, 47.4, 2.98, 2.54], [78, 35.9, -0.5, 0.49], null, null],
    "地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 近3日漲停": [[766, 49.3, 3.78, 2.04], null, [46, 63.0, 7.5, 1.95], null],
    "布林通道高檔(≥80%) ＋ 晨星剛形成 ＋ 月營收創24個月新高": [[149, 52.3, 3.7, 2.11], null, null, null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1714, 54.1, 6.04, 4.64], [186, 35.5, -1.42, -2.55], [97, 39.2, 2.12, 1.6], [68, 48.5, 3.08, 1.42]],
    "N字底剛形成 ＋ 外資近5日買超 ＋ 短線過熱(高於月線≥20%)": [[172, 47.1, 5.3, 1.87], null, null, null],
    "MACD近3日內黃金交叉 ＋ OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%)": [[497, 46.9, 2.71, 2.87], [65, 49.2, 1.57, -1.13], [36, 50.0, 2.73, 0.97], [30, 50.0, 1.14, 0.16]],
    "MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [[366, 43.4, 1.78, 1.87], [59, 49.2, 1.65, -0.4], [32, 50.0, 2.88, 1.0], [35, 48.6, 0.72, -0.12]],
    "創52週新高 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], null, null, null],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[79, 49.4, 6.83, 2.44], null, null, null],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[1348, 55.8, 7.15, 5.51], [150, 38.7, -0.21, -0.46], [70, 55.7, 6.89, 1.81], [53, 58.5, 4.27, 1.43]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1537, 53.9, 6.5, 5.64], [155, 30.3, -0.92, -3.09], [94, 46.8, 4.62, 2.48], [61, 47.5, 2.57, 1.12]],
    "MACD近3日內黃金交叉 ＋ 月營收連3月年增 ＋ 短線過熱(高於月線≥20%)": [[314, 47.1, 2.09, 1.34], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 近20日強於大盤≥10%": [[84, 48.8, 4.05, 1.92], null, null, null],
    "MACD近3日內黃金交叉 ＋ 多方力道≥65 ＋ 短線過熱(高於月線≥20%)": [[469, 51.0, 2.93, 2.26], [62, 48.4, 0.49, -0.53], [36, 50.0, 2.74, 1.12], [32, 62.5, 7.78, 1.2]],
    "布林通道高檔(≥80%) ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[2061, 55.1, 6.16, 6.2], [191, 35.6, -0.25, -1.62], [141, 51.8, 3.9, 2.29], [72, 54.2, 6.87, 2.5]],
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1243, 48.3, 3.27, 3.66], [223, 57.4, 5.32, 0.8], [201, 56.2, 6.1, 2.92], [118, 50.0, 4.54, 3.4]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[223, 52.0, 5.12, 3.47], [47, 51.1, 2.82, 0.17], null, null],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[124, 38.7, -1.83, -0.38], [33, 48.5, 1.13, 0.37], null, null],
    "多方力道≥80 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[212, 52.4, 4.95, 2.92], [46, 50.0, 1.17, -0.49], null, null],
    "夜星剛形成 ＋ 月營收創24個月新高": [[144, 53.5, 4.44, 2.32], [40, 50.0, 4.49, 1.88], null, null],
    "三大法人近3月買超 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1769, 56.0, 6.73, 5.67], [188, 37.2, 0.67, -0.8], [101, 57.4, 8.24, 3.88], [69, 50.7, 5.95, 2.18]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[1148, 52.7, 6.56, 5.65], [138, 38.4, -0.37, -1.19], [102, 53.9, 4.45, 2.05], [64, 54.7, 6.84, 2.31]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10%": [[2070, 54.8, 6.08, 5.76], [210, 34.3, -1.18, -1.7], [143, 49.0, 3.49, 2.15], [76, 51.3, 5.35, 2.12]],
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月均價YoY為負": [[1488, 48.6, 3.81, 4.64], [253, 54.9, 4.33, 1.23], [204, 56.9, 6.65, 2.86], [132, 50.8, 2.38, 1.24]],
    "母子懷抱(低檔)剛形成 ＋ 近20日強於大盤≥10% ＋ 近3日漲停": [[85, 44.7, 1.57, 0.47], null, null, null],
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 量能區間低檔(≤10百分位)": [[49, 77.6, 8.61, 2.84], null, null, null],
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[55, 72.7, 6.8, 2.57], null, null, null],
    "夜星剛形成 ＋ 大盤跌破20日均線 ＋ 月營收連3月年增": [null, [93, 59.1, 6.65, 2.18], null, [65, 49.2, 2.47, 0.2]],
    "地量(≤0.5倍均量) ＋ 夜星剛形成 ＋ 近20日強於大盤≥10%": [[32, 53.1, 6.48, 1.55], null, null, null],
    "多方力道≥65 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[57, 70.2, 5.51, 2.09], null, null, null],
    "短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10% ＋ 量能區間低檔(≤10百分位)": [[58, 70.7, 5.67, 2.2], null, null, null],
    "晨星剛形成 ＋ 月營收創24個月新高 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[45, 48.9, 2.2, 0.37], null, null, null],
    "短線過熱(高於月線≥20%) ＋ 量能區間低檔(≤10百分位)": [[59, 71.2, 5.89, 2.21], null, null, null],
    "K線橫盤的突破剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": [[114, 47.4, 5.67, 2.38], null, null, null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[2141, 54.8, 6.26, 5.81], [211, 34.6, -1.12, -1.71], [148, 50.7, 3.8, 2.18], [76, 51.3, 5.35, 2.12]],
    "地量(≤0.5倍均量) ＋ 均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成": [[62, 51.6, 4.91, 0.52], null, null, null],
    "月營收創24個月新高 ＋ 相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%)": [[2132, 54.8, 6.16, 5.79], [211, 34.6, -1.12, -1.71], [148, 50.7, 3.8, 2.18], [76, 51.3, 5.35, 2.12]],
    "近3日漲停 ＋ 量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[466, 47.6, 4.47, 3.06], [134, 47.0, 2.18, 1.2], [41, 34.1, -2.34, -0.4], [74, 51.4, 0.87, 0.23]],
    "創52週新高 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], null, null, null],
    "三大法人近3月買超 ＋ 創52週新高 ＋ 母子懷抱(低檔)剛形成": [[76, 52.6, 6.12, 2.35], null, null, null],
    "創52週新高 ＋ 多方力道≥65 ＋ 母子懷抱(低檔)剛形成": [[67, 47.8, 4.64, 1.78], null, null, null],
    "月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成 ＋ 距52週高點≤5%": [[313, 54.6, 6.23, 4.1], [47, 48.9, 3.78, 0.73], [34, 50.0, 8.01, 1.67], null],
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[657, 48.9, 3.08, 3.52], [108, 49.1, 1.82, -0.72], [154, 47.4, -0.16, -0.5], [85, 51.8, 4.31, 1.93]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 月營收創24個月新高": [[85, 51.8, 5.28, 2.36], null, null, null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[252, 52.4, 3.45, 2.88], [37, 51.4, 1.4, 0.37], null, null],
    "強勢突破盤 ＋ 月營收連3月年增 ＋ 母子懷抱(低檔)剛形成": [[89, 46.1, 3.78, 1.84], null, null, null],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 近3日漲停": [[920, 45.3, 2.96, 2.96], [186, 49.5, 1.76, 1.29], [80, 47.5, 2.86, 1.18], [73, 49.3, -1.96, 0.23]],
    "KDJ近3日內黃金交叉 ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[245, 51.8, 1.62, 1.35], [43, 44.2, -1.91, -0.29], null, null],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 近3日向上跳空缺口": [[829, 46.4, 2.56, 3.25], [134, 50.7, 3.11, 1.54], [64, 43.8, 1.02, -0.48], [38, 50.0, 0.12, 0.82]],
    "創52週新高 ＋ 近3日漲停 ＋ 近3月均價YoY為負": [[1419, 46.2, 2.74, 3.67], [267, 55.1, 4.75, 1.37], [200, 54.0, 6.01, 1.95], [140, 48.6, 2.44, 2.22]],
    "月營收創24個月新高 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": [[1877, 54.4, 6.58, 5.55], [184, 33.7, -0.62, -1.6], [130, 50.8, 3.56, 2.09], [69, 47.8, 2.34, 1.11]],
    "近3日向上跳空缺口 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[534, 50.2, 4.38, 3.22], [122, 59.8, 10.5, 2.73], [35, 57.1, 3.45, 0.42], [221, 86.4, 6.22, 0.9]],
    "均線多頭排列(5>20>60且站上月線) ＋ 晨星剛形成 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[154, 49.4, 2.99, 1.57], null, null, null],
    "三大法人近3月買超 ＋ 短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成": [[274, 48.5, 3.46, 2.59], [49, 53.1, 1.97, 0.62], null, null],
    "相對強弱為負(弱於大盤) ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[228, 48.2, 1.22, 1.0], [56, 64.3, 11.76, 2.62], null, [238, 89.5, 6.11, 0.71]],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[383, 52.5, 6.45, 2.99], [43, 46.5, -0.41, 0.33], null, null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 突破飆股大量黑K最高點剛形成": [[148, 45.9, 3.74, 1.57], null, null, null],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3日漲停": [[624, 48.7, 3.19, 3.15], [88, 48.9, 1.52, -0.46], [40, 47.5, 3.11, 1.0], [43, 55.8, 3.16, 0.79]],
    "大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[149, 48.3, 3.2, 1.45], [52, 61.5, 6.82, 1.88], null, null],
    "夜星剛形成 ＋ 月營收創24個月新高 ＋ 相對強弱為正(強於大盤)": [[90, 50.0, 4.62, 2.35], [32, 50.0, 3.28, 1.25], null, null],
    "創52週新高 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[417, 57.1, 6.27, 3.6], [42, 50.0, 10.11, 1.19], null, null],
    "距52週高點≤5% ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[371, 54.2, 5.14, 3.54], [61, 59.0, 9.41, 2.16], null, null],
    "三大法人近3月買超 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[774, 48.8, 4.1, 3.49], [253, 57.3, 8.46, 2.54], [65, 43.1, -0.33, -0.02], [225, 73.8, 5.52, 1.77]],
    "創52週新高 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], null, null, null],
    "MACD近3日內黃金交叉 ＋ 漲時量≥跌時量1.5倍(近20日) ＋ 短線過熱(高於月線≥20%)": [[537, 49.5, 3.89, 3.22], [74, 47.3, -1.44, -1.45], [34, 47.1, 2.77, 0.8], [37, 51.4, 1.39, 0.14]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3日向上跳空缺口": [[540, 47.2, 2.18, 3.18], [76, 47.4, 1.2, -0.39], [31, 48.4, 3.68, 0.8], [37, 54.1, 3.54, 0.66]],
    "回後買上漲全通過 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[640, 46.7, 2.94, 2.45], [101, 52.5, 0.7, -0.04], [129, 62.8, 5.69, 2.95], [67, 50.7, 6.21, 1.93]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成": [[71, 50.7, 5.08, 1.78], null, null, null],
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[1571, 47.0, 3.03, 4.64], [283, 53.7, 3.11, 1.72], [300, 50.0, 3.3, 1.28], [186, 47.8, 5.02, 2.86]],
    "N字底剛形成 ＋ 短線過熱(高於月線≥20%) ＋ 近3日向上跳空缺口": [[111, 46.8, 4.05, 1.19], null, null, null],
    "外資近5日買超 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[717, 48.3, 4.23, 3.31], [225, 58.7, 8.24, 2.74], [62, 38.7, -1.02, 0.36], [284, 79.6, 5.87, 2.09]],
    "外資近5日買超 ＋ 夜星剛形成 ＋ 大盤跌破60日均線": [null, null, null, [68, 52.9, 5.03, 0.99]],
    "KDJ近3日內死亡交叉 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[715, 51.5, 3.9, 3.51], [103, 44.7, -0.63, -1.56], [61, 49.2, 2.43, 1.61], [65, 41.5, -2.02, -0.1]],
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1573, 45.3, 2.09, 2.29], [262, 50.8, 0.94, -0.87], [291, 49.8, 3.72, 2.06], [158, 50.6, 3.68, 1.6]],
    "創52週新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近20日強於大盤≥10%": [[461, 49.7, 3.12, 3.12], [115, 53.0, 4.62, 1.14], [61, 50.8, 8.27, 2.27], [48, 41.7, -1.05, -0.52]],
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1378, 45.3, 1.99, 2.66], [246, 52.0, 2.25, 0.61], [278, 50.7, 3.5, 1.51], [186, 47.3, 3.85, 1.89]],
    "晨星剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[50, 52.0, 8.04, 2.37], null, null, null],
    "外資近5日買超 ＋ 月營收創24個月新高 ＋ 短線過熱(高於月線≥20%)": [[1460, 54.9, 6.95, 5.83], [151, 36.4, 0.25, -1.71], [101, 52.5, 4.95, 2.71], [62, 46.8, 4.86, 1.79]],
    "創52週新高 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[323, 49.5, 4.0, 2.93], [59, 57.6, 7.47, 1.74], null, null],
    "MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 短線過熱(高於月線≥20%)": [[654, 48.3, 3.01, 3.12], null, [43, 51.2, 3.4, 1.22], null],
    "MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 短線過熱(高於月線≥20%)": [[641, 48.2, 2.83, 3.02], [95, 49.5, 1.03, -0.66], [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10%": [[583, 49.4, 2.95, 2.9], [94, 50.0, 1.3, -0.61], [42, 50.0, 2.78, 1.14], [44, 56.8, 4.71, 0.93]],
    "創52週新高 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[2037, 47.9, 3.56, 6.71], [377, 54.4, 4.23, 1.97], [262, 54.6, 4.93, 1.94], [174, 46.0, 1.2, 1.17]],
    "KDJ近3日內死亡交叉 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [[381, 47.5, 3.63, 2.82], [79, 43.0, -0.95, -0.91], [59, 55.9, 3.67, 1.34], [58, 43.1, 0.01, 0.27]],
    "創52週新高 ＋ 多方力道≥80 ＋ 突破ABC修正下降切線剛形成": [[372, 49.2, 3.73, 3.07], [57, 42.1, 1.62, -0.02], [37, 54.1, 10.7, 2.08], null],
    "母子懷抱(低檔)剛形成 ＋ 近20日強於大盤≥10% ＋ 近3日向上跳空缺口": [[60, 48.3, 3.69, -0.19], null, null, null],
    "三大法人近3月賣超 ＋ 地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%)": [[139, 49.6, 7.13, 2.25], null, [31, 45.2, -3.47, -2.09], null],
    "晨星剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[107, 48.6, 6.04, 2.4], null, null, null],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[452, 46.9, 3.77, 2.4], [73, 47.9, -0.65, -1.13], [35, 54.3, 3.35, 0.86], [33, 48.5, -0.2, 0.07]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[654, 48.3, 3.01, 3.12], [96, 49.0, 0.67, -0.67], [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "MACD近3日內黃金交叉 ＋ 布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%)": [[653, 48.2, 3.02, 3.14], [96, 49.0, 0.67, -0.67], [43, 51.2, 3.4, 1.22], [44, 56.8, 4.71, 0.93]],
    "強勢突破盤 ＋ 晨星剛形成 ＋ 月營收連3月年增": [[130, 46.2, 2.07, 1.48], null, null, null],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[292, 58.9, 6.76, 4.06], null, null, null],
    "短線過熱(高於月線≥20%) ＋ 距52週高點≤5% ＋ 近3月均價YoY為負": [[1381, 48.4, 3.78, 3.84], [236, 57.2, 4.86, 0.83], [220, 57.7, 7.98, 3.52], [121, 52.1, 5.05, 3.5]],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[1208, 50.4, 5.19, 4.53], [138, 37.0, -1.3, -1.44], [77, 50.6, 3.96, 2.21], [65, 49.2, 6.03, 1.93]],
    "月營收連3月年增 ＋ 母子懷抱(低檔)剛形成 ＋ 爆量(≥2倍均量)": [[69, 56.5, 7.37, 2.57], null, null, null],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 短線過熱(高於月線≥20%)": [[734, 45.9, 3.27, 2.23], [126, 49.2, 2.21, 1.26], [69, 47.8, 1.45, 0.61], [49, 44.9, -1.46, 0.1]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[2144, 47.0, 2.53, 4.15], [398, 49.7, 2.16, 0.62], [405, 54.6, 5.42, 2.37], [207, 44.9, 3.65, 2.33]],
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1651, 45.2, 2.32, 2.42], [281, 53.0, 2.22, 0.52], [301, 53.5, 4.66, 2.35], [182, 51.1, 4.14, 1.29]],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[2031, 45.3, 1.81, 2.0], [370, 50.5, 2.52, 1.06], [395, 52.2, 3.81, 1.61], [254, 44.9, 2.84, 1.66]],
    "K線橫盤的突破剛形成 ＋ N字底剛形成 ＋ 近3日漲停": [[87, 46.0, 4.06, 1.75], null, null, null],
    "地量(≤0.5倍均量) ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[352, 48.9, 2.6, 1.64], [114, 64.9, 13.63, 2.75], null, [140, 77.9, 4.38, 0.64]],
    "MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 短線過熱(高於月線≥20%)": [[539, 48.1, 2.36, 2.86], [67, 44.8, 0.61, -0.48], [33, 48.5, 2.3, 0.79], [31, 54.8, 4.82, 0.99]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%)": [[453, 47.9, 3.63, 3.09], [64, 48.4, 1.54, 0.22], null, [30, 50.0, -0.61, 0.16]],
    "創52週新高 ＋ 近3日向上跳空缺口 ＋ 近3月均價YoY為負": [[1307, 49.1, 3.34, 5.05], [213, 59.2, 7.27, 1.55], [179, 47.5, 2.63, 1.61], [88, 44.3, 1.91, 1.44]],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 量能區間高檔(≥90百分位)": [[2346, 44.9, 1.89, 2.85], [431, 50.3, 2.3, 1.17], [417, 49.9, 2.88, 1.54], [273, 45.1, 2.5, 1.85]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 創52週新高 ＋ 地量(≤0.5倍均量)": [[426, 57.3, 6.23, 3.5], [44, 50.0, 9.3, 1.11], null, null],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為負(股價超前營收)": [[639, 51.5, 4.67, 2.91], [74, 41.9, -0.8, 0.47], null, null],
    "地量(≤0.5倍均量) ＋ 近3日向上跳空缺口 ＋ 近3日漲停": [[484, 53.9, 6.52, 3.59], [106, 70.8, 12.94, 2.66], null, [261, 90.8, 4.38, 1.66]],
    "近3日漲停 ＋ 近3月均價YoY為正 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[819, 47.5, 2.82, 2.41], [222, 53.2, 8.5, 2.38], [66, 50.0, 1.67, 0.5], [266, 81.2, 6.26, 0.91]],
    "大盤跌破20日均線 ＋ 近3日漲停 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [null, [338, 56.2, 7.04, 2.74], null, [399, 76.7, 4.15, 1.03]],
    "創52週新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[1786, 46.9, 3.14, 5.71], [298, 55.0, 5.4, 2.73], [259, 52.5, 4.26, 2.61], [150, 44.7, 3.43, 2.81]],
    "MACD近3日內黃金交叉 ＋ 強勢突破盤 ＋ 短線過熱(高於月線≥20%)": [[326, 45.1, 2.43, 2.26], [55, 47.3, -2.17, -0.39], null, null],
    "布林通道高檔(≥80%) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[2605, 46.7, 2.45, 2.93], [458, 53.1, 3.36, 0.85], [489, 51.5, 3.72, 1.36], [283, 45.2, 2.43, 1.58]],
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1724, 46.3, 2.2, 2.46], [303, 50.5, 1.92, 0.42], [325, 49.8, 3.18, 1.54], [214, 46.3, 3.54, 2.26]],
    "短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[1770, 45.6, 1.78, 1.93], [319, 49.8, 2.23, 0.61], [336, 51.2, 3.73, 1.66], [213, 48.8, 4.72, 2.84]],
    "K線橫盤的突破剛形成 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[381, 48.0, 3.96, 3.53], [83, 51.8, 0.89, 0.06], [80, 52.5, 6.47, 1.62], [52, 44.2, -1.12, 0.23]],
    "地量(≤0.5倍均量) ＋ 晨星剛形成 ＋ 近20日強於大盤≥10%": [[52, 51.9, 5.83, 0.87], null, null, null],
    "K線橫盤的突破剛形成 ＋ 晨星剛形成 ＋ 距52週高點≤5%": [[81, 54.3, 7.38, 2.9], null, null, null],
    "創52週新高 ＋ 母子懷抱(低檔)剛形成 ＋ 近20日強於大盤≥10%": [[65, 49.2, 6.05, 2.08], null, null, null],
    "月營收創24個月新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[545, 47.9, 3.52, 1.9], [40, 22.5, -2.67, -1.69], [47, 48.9, 10.18, 1.49], null],
    "K線橫盤的突破剛形成 ＋ 均線多頭排列(5>20>60且站上月線) ＋ 母子懷抱(低檔)剛形成": [[54, 55.6, 5.58, 1.96], null, null, null],
    "創52週新高 ＋ 母子懷抱(低檔)剛形成 ＋ 相對強弱為正(強於大盤)": [[81, 51.9, 5.99, 2.34], null, null, null],
    "創52週新高 ＋ 母子懷抱(低檔)剛形成": [[85, 50.6, 5.08, 2.11], null, null, null],
    "地量(≤0.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[167, 50.3, 3.9, 2.17], null, null, null],
    "布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[108, 43.5, 2.32, 0.95], null, null, null],
    "母子懷抱(低檔)剛形成 ＋ 距52週高點≤5% ＋ 近3日漲停": [[65, 41.5, 1.61, 0.68], null, null, null],
    "KDJ近3日內死亡交叉 ＋ 夜星剛形成 ＋ 月營收創24個月新高": [[94, 56.4, 4.96, 1.74], null, null, null],
    "近3日漲停 ＋ 近3月乖離度為負(股價超前營收) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)": [[678, 49.1, 3.31, 3.24], [192, 54.7, 9.26, 2.48], [46, 56.5, 3.44, 1.14], [204, 80.9, 4.86, -0.03]],
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[828, 48.7, 3.14, 3.34], [139, 59.0, 7.9, 2.87], [146, 50.7, 0.97, -0.01], [96, 52.1, 5.91, 2.69]],
    "創52週新高 ＋ 月營收創24個月新高 ＋ 爆量(≥2倍均量)": [[1043, 51.0, 6.0, 4.53], [140, 40.0, -1.88, -0.57], [82, 50.0, 6.3, 2.76], [91, 54.9, 4.21, 1.32]],
    "三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[606, 50.5, 4.51, 4.11], [123, 63.4, 8.1, 2.27], [54, 46.3, 5.2, 1.32], [45, 40.0, -0.17, -0.16]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 創52週新高 ＋ 突破ABC修正下降切線剛形成": [[429, 51.3, 4.31, 3.98], [69, 46.4, 3.99, 1.15], [39, 51.3, 9.5, 1.94], null],
    "強勢突破盤 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[2059, 47.4, 3.49, 5.69], [333, 53.2, 2.42, 1.22], [345, 51.6, 3.54, 2.01], [191, 46.6, 4.16, 2.24]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破ABC修正下降切線剛形成": [[277, 55.2, 5.37, 3.56], [42, 57.1, 5.55, 1.31], [40, 50.0, 5.63, 1.83], null],
    "大盤跌破60日均線 ＋ 月營收創24個月新高 ＋ 近3日漲停": [null, null, [283, 57.2, 4.74, 2.37], [701, 76.5, 4.58, 1.35]],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 近3日向上跳空缺口": [[218, 59.2, 7.3, 3.27], null, null, null],
    "短線過熱(高於月線≥20%) ＋ 突破ABC修正下降切線剛形成 ＋ 距52週高點≤5%": [[262, 48.5, 3.13, 1.71], [44, 52.3, 3.33, 0.64], null, null],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 母子懷抱(低檔)剛形成 ＋ 近3日漲停": [[64, 46.9, 3.23, 0.98], null, null, null],
    "爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[1995, 44.6, 2.67, 2.96], [347, 46.1, 1.3, 0.46], [195, 56.9, 5.07, 2.27], [157, 49.0, 3.87, 2.37]],
    "OBV能量潮創60日新高 ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[1935, 45.3, 2.17, 3.94], [300, 48.3, 1.57, 0.94], [334, 47.0, 2.4, 0.66], [186, 47.3, 3.82, 1.99]],
    "短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[2365, 47.1, 2.93, 5.25], [437, 52.2, 3.36, 2.35], [442, 50.0, 3.13, 1.16], [270, 44.1, 2.44, 2.22]],
    "三大法人近3月買超 ＋ 夜星剛形成 ＋ 近20日強於大盤≥10%": [[198, 42.4, 0.77, 1.04], [61, 55.7, 5.38, 1.39], null, [63, 55.6, 4.88, 1.92]],
    "OBV能量潮創60日新高 ＋ 月營收創24個月新高 ＋ 近3日漲停": [[1161, 55.6, 7.73, 6.87], [146, 36.3, -1.98, -1.11], [107, 50.5, 3.56, 1.46], [75, 54.7, 3.03, 0.55]],
    "晨星剛形成 ＋ 距52週高點≤5% ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[134, 47.0, 3.95, 2.21], null, null, null],
    "均線多頭排列(5>20>60且站上月線) ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[2415, 45.7, 2.26, 2.65], [454, 53.5, 3.99, 1.37], [455, 51.4, 4.13, 2.56], [272, 44.5, 2.09, 2.0]],
    "距52週高點≤5% ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[2116, 49.4, 4.2, 6.77], [386, 53.9, 3.45, 1.34], [278, 55.0, 4.92, 2.4], [176, 46.6, 0.84, 1.14]],
    "爆量(≥1.5倍均量) ＋ 短線過熱(高於月線≥20%) ＋ 近3月乖離度為正(營收優於股價)": [[2013, 47.6, 3.05, 5.41], [355, 52.1, 2.9, 1.99], [364, 48.6, 2.52, 1.11], [224, 46.9, 3.99, 2.6]],
    "KDJ近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)": [[700, 46.7, 2.15, 2.6], [139, 36.0, -0.85, -1.09], [124, 51.6, 0.92, -0.25], [90, 48.9, 6.3, 2.82]],
    "月營收創24個月新高 ＋ 近3日漲停 ＋ 近3月乖離度為正(營收優於股價)": [[1409, 50.2, 4.25, 4.92], [287, 50.5, 1.47, -0.08], [188, 55.3, 5.61, 2.48], [421, 75.3, 4.23, 1.44]],
    "三大法人近3月賣超 ＋ 創52週新高 ＋ 回後買上漲全通過": [[480, 50.4, 1.57, 1.37], [78, 46.2, -0.52, -0.11], [36, 61.1, 7.51, 1.82], null],
    "大盤跌破60日均線 ＋ 爆量(≥2倍均量) ＋ 短線過熱(高於月線≥20%)": [null, null, [459, 52.3, 2.26, 1.66], [315, 45.1, 2.13, 1.63]],
    "KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 短線過熱(高於月線≥20%)": [[229, 48.9, 5.62, 2.27], null, null, null],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 多方力道≥80": [[381, 56.2, 6.03, 3.74], [30, 50.0, -0.88, -0.39], [38, 52.6, 9.99, 2.13], [31, 35.5, -0.08, -0.23]],
    "CMF資金流買方佔優(近20日≥0.1) ＋ 多方力道≥80 ＋ 晨星剛形成": [[199, 48.7, 3.94, 2.14], null, null, null],
    "外資近5日買超 ＋ 短線過熱(高於月線≥20%) ＋ 近3月均價YoY為負": [[1655, 48.6, 3.98, 4.7], [311, 58.8, 5.81, 1.63], [311, 55.6, 5.56, 1.19], [157, 47.1, 3.27, 2.28]],
    "大盤站上20日均線 ＋ 短線過熱(高於月線≥20%) ＋ 距52週高點≤5%": [[4832, 48.4, 3.38, 4.49], null, [422, 53.6, 3.81, 2.35], null],
    "短線過熱(高於月線≥20%) ＋ 近20日強於大盤≥10% ＋ 近3月均價YoY為負": [[2643, 46.4, 2.47, 2.97], [492, 53.0, 3.39, 1.13], [500, 51.6, 3.74, 1.49], [292, 44.5, 2.1, 1.6]],
    "MACD近3日內黃金交叉 ＋ 短線過熱(高於月線≥20%) ＋ 量能區間高檔(≥90百分位)": [[492, 45.5, 1.94, 2.06], [76, 44.7, -1.1, -1.31], [37, 51.4, 3.76, 1.38], [35, 51.4, 2.67, -0.13]],
    "強勢突破盤 ＋ 月營收創24個月新高 ＋ 突破飆股大量黑K最高點剛形成": [[304, 49.0, 5.0, 3.21], [49, 55.1, 2.71, -0.39], null, null],
    "KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 近3日漲停": [[402, 52.0, 4.46, 2.19], [90, 55.6, 4.35, 0.76], [44, 45.5, 3.15, 0.84], [39, 43.6, 1.16, 0.76]],
    "DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR) ＋ 夜星剛形成 ＋ 月營收創24個月新高": [[76, 47.4, 3.19, 1.55], null, null, null],
    "OBV能量潮創60日新高 ＋ 地量(≤0.5倍均量) ＋ 連續放量(近3日均量≥1.5倍前20日均量)": [[44, 59.1, 13.21, 3.12], null, null, null],
  };
  var MARKET_REGIMES = { 1: ['①多頭', '加權指數站上月線（MA20）與季線（MA60）'], 2: ['②多頭回檔', '加權指數跌破月線、仍守住季線'],
    3: ['③空頭反彈', '加權指數站上月線、但仍在季線下'], 4: ['④空頭', '加權指數跌破月線與季線'] };
  var REGIME_ADVICE = { 1: '各類組合都有效：S選股型、W高勝率、F飆股、近20日強於大盤、距52週高點、營收創新高',
    2: 'W高勝率與#清單失效；S選股型、月營收連3月年增、外資買超仍有效，宜挑營收題材',
    3: '#清單明顯為負（避開）；S選股型、F飆股、營收創新高、近20日強勢股有效',
    4: '追強勢（近20日強於大盤、距52週高點）失效、投信買超為負；S選股型、月營收連3月年增、外資買超仍有效' };
  var REGIME_MIN_N = 50;
  function marketRegimeOf(r) {
    var fr; try { fr = computeLiveFlagsRow(r); } catch (e) { return null; }
    if (!fr || fr.benchmarkAbove20 == null || fr.benchmarkAbove60 == null) return null;
    var a20 = fr.benchmarkAbove20 === 1, a60 = fr.benchmarkAbove60 === 1;
    return (a20 && a60) ? 1 : (a60 ? 2 : (a20 ? 3 : 4));
  }
  function regimeTag(combo, reg) {
    if (!reg) return '';
    var row = REGIME_STATS[comboNormKey(combo)], st = row ? row[reg - 1] : null;
    if (!st || st[0] < REGIME_MIN_N) return '';
    if (st[3] != null && st[3] >= 2 && st[2] > 0) return '✅';
    return st[2] < 0 ? '❌' : '⚠️';
  }
  function regimeFit(r) {
    if (r.__rf) return r.__rf;   // 排序時會被呼叫很多次，算一次就快取
    var res = regimeFitCalc(r);
    try { Object.defineProperty(r, '__rf', { value: res, writable: true, configurable: true, enumerable: false }); } catch (e) {}
    return res;
  }
  function regimeFitCalc(r) {
    var reg = marketRegimeOf(r), b = { '✅': [], '⚠️': [], '❌': [] };
    if (!reg) return { reg: null, ok: 0, bad: 0, b: b };
    var groups = [['S', STOCK_PICK_COMBOS, matchedStockPickCombos], ['#', PINNED_COMBOS, matchedPinnedCombos], ['M', MOONSHOT_COMBOS, matchedMoonshotCombos],
                  ['W', BT_WIN_COMBOS, matchedBtWinCombos], ['F', BT_HOT_COMBOS, matchedBtHotCombos]];
    groups.forEach(function(g) {
      var idxs; try { idxs = g[2](r); } catch (e) { idxs = []; }
      idxs.forEach(function(i) { var t = regimeTag(g[1][i], reg); if (t) b[t].push(g[0] + (i + 1)); });
    });
    return { reg: reg, ok: b['✅'].length, bad: b['❌'].length, b: b };
  }
  function renderRegimeFit(r) {
    var f = regimeFit(r);
    if (!f.reg || !(f.ok + f.bad + f.b['⚠️'].length)) return '<td class="nw" style="color:#555;text-align:center">－</td>';
    var cols = { '✅': '#4caf50', '❌': '#ff5555', '⚠️': '#f0a500' }, html = [], tip = [];
    ['✅', '❌', '⚠️'].forEach(function(t) {
      var v = f.b[t];
      if (!v.length) return;
      html.push('<span style="color:' + cols[t] + ';font-weight:700">' + t + v.length + '</span>');
      tip.push(t + ' ' + v.join(' '));
    });
    return '<td class="nw" title="' + MARKET_REGIMES[f.reg][0] + '下回測：' + tip.join('｜') + '">' + html.join(' ') +
      '<div style="font-size:9px;color:#888">' + (f.b['✅'].slice(0, 3).join(' ') || f.b['❌'].slice(0, 3).join(' ')) + '</div></td>';
  }
  function batchMarketRegime() {
    for (var i = 0; i < batchResults.length; i++) { var g = marketRegimeOf(batchResults[i]); if (g) return g; }
    return null;
  }
  // ════════════════════════════════════════════════════════════════════
  //  大盤溫度計（全市場廣度）＋大盤K線位置（2026-10-19；與 PY 版同一套計算）
  //  溫度計：這次批次分析所有股票的訊號加總成「市場廣度」，跟三年回測每天的廣度分布比較（百分位），
  //          合成 0（極冷）～100（過熱）；三年回測：過熱後未來20日大盤偏弱、極冷後偏強。需要 ≥150 檔。
  //  K線位置：大盤收盤價判斷：距60日高點回檔≥10%、RSI(14)<35 或均線空頭排列＝超賣；RSI>65＝過熱。
  // ════════════════════════════════════════════════════════════════════
  var MKT_CALIB = {"Q": {"平均多方力道": [41.81004, 47.37472, 49.17032, 50.17712, 50.86521, 51.33941, 51.85213, 52.2347, 52.66536, 53.05976, 53.55392, 54.1873, 55.17925, 56.00244, 57.06301, 57.76912, 58.58097, 59.50855, 60.47151, 62.44357, 69.05623], "力道≥65比例": [0.02019, 0.09584, 0.14063, 0.16721, 0.18327, 0.19887, 0.21171, 0.22293, 0.23414, 0.24429, 0.2552, 0.26974, 0.28749, 0.31306, 0.34173, 0.35797, 0.37755, 0.39913, 0.43094, 0.48761, 0.67325], "買盤量比例": [0.03962, 0.13312, 0.17955, 0.20354, 0.2159, 0.22857, 0.24039, 0.2511, 0.26586, 0.27847, 0.28989, 0.31479, 0.32804, 0.34432, 0.35704, 0.36903, 0.38611, 0.41019, 0.44006, 0.53452, 0.70706], "多頭排列比例": [0.00269, 0.08024, 0.13075, 0.15731, 0.17567, 0.19429, 0.20871, 0.22538, 0.23934, 0.25572, 0.2678, 0.28338, 0.3008, 0.31966, 0.34411, 0.36726, 0.38978, 0.40705, 0.43262, 0.46508, 0.62158], "近高點比例": [0.00137, 0.0279, 0.03784, 0.04431, 0.04823, 0.05092, 0.05523, 0.05888, 0.0634, 0.06834, 0.07244, 0.07739, 0.08253, 0.09061, 0.10669, 0.12097, 0.14024, 0.15381, 0.17145, 0.18867, 0.2375], "W命中比例": [0.0, 0.00491, 0.00756, 0.00924, 0.01107, 0.0132, 0.01471, 0.01666, 0.01905, 0.02185, 0.02482, 0.02864, 0.03406, 0.04364, 0.05228, 0.06688, 0.07833, 0.08588, 0.12418, 0.13807, 0.25968], "布林低檔比例": [0.00141, 0.02456, 0.04156, 0.05935, 0.07656, 0.08724, 0.09994, 0.11608, 0.1266, 0.13598, 0.15121, 0.17487, 0.19652, 0.21156, 0.23768, 0.26744, 0.30134, 0.35104, 0.42405, 0.52224, 0.98385], "F命中比例": [0.0, 0.01032, 0.01337, 0.01605, 0.01813, 0.02086, 0.02257, 0.02464, 0.02703, 0.02898, 0.03231, 0.03534, 0.03855, 0.04212, 0.04569, 0.04907, 0.05468, 0.0676, 0.07848, 0.10461, 0.43952]}, "HEAT": ["平均多方力道", "力道≥65比例", "買盤量比例", "多頭排列比例", "近高點比例"], "COLD": ["W命中比例", "布林低檔比例", "F命中比例"], "TEMP": {"偏冷": {"days": 237, "mkt20": 1.25, "up": 59.1, "allWin": 48.3, "sWin": 53.4, "sN": 7085}, "中性": {"days": 181, "mkt20": 1.51, "up": 65.2, "allWin": 46.7, "sWin": 54.1, "sN": 7791}, "偏熱": {"days": 191, "mkt20": 0.85, "up": 64.9, "allWin": 45.4, "sWin": 48.2, "sN": 8139}, "過熱": {"days": 50, "mkt20": 0.76, "up": 66.0, "allWin": 45.2, "sWin": 52.5, "sN": 1329}, "極冷": {"days": 38, "mkt20": 5.78, "up": 89.5, "allWin": 65.3, "sWin": 69.7, "sN": 653}}, "KLINE": {"正常": {"days": 392, "mkt20": 0.91, "up": 61.5, "allWin": 45.4, "sWin": 52.1, "sN": 13238}, "過熱": {"days": 131, "mkt20": 0.97, "up": 69.5, "allWin": 44.5, "sWin": 50.9, "sN": 7873}, "超賣": {"days": 115, "mkt20": 3.16, "up": 67.8, "allWin": 56.9, "sWin": 58.5, "sN": 2711}}};
  var MKT_FLAG_OF = { '力道≥65比例': '多方力道≥65', '買盤量比例': '漲時量≥跌時量1.5倍(近20日)', '多頭排列比例': '均線多頭排列(5>20>60且站上月線)',
    '近高點比例': '距52週高點≤5%', '布林低檔比例': '布林通道低檔(≤20%)' };
  var MKT_MIN_STOCKS = 150;
  var TEMP_ZONES = [[25, '極冷'], [45, '偏冷'], [55, '中性'], [75, '偏熱'], [101, '過熱']];
  function calibPct(q, x) {
    if (x == null) return null;
    var inc = true, i;
    for (i = 0; i < 20; i++) if (!(q[i] < q[i + 1])) { inc = false; break; }
    if (inc) {
      if (x <= q[0]) return 0;
      if (x >= q[20]) return 1;
      for (i = 0; i < 20; i++) if (q[i] <= x && x <= q[i + 1]) return (i + (x - q[i]) / (q[i + 1] - q[i])) / 20;
    }
    var lo = q.filter(function(v) { return v < x; }).length / 20, hi = q.filter(function(v) { return v <= x; }).length / 20;
    return Math.min(1, (lo + hi) / 2);
  }
  function marketTemperature() {
    var rs = batchResults.filter(function(r) { return r && r.data && r.data.length; });
    var rows = [];
    rs.forEach(function(r) { try { rows.push(computeLiveFlagsRow(r)); } catch (e) {} });
    var n = rows.length;
    if (n < MKT_MIN_STOCKS) return { n: n, temp: null, zone: null, comps: {} };
    var defs = buildConditionFlags(rows), vals = {};
    vals['平均多方力道'] = rs.reduce(function(s, r) { return s + (+r.total || 0); }, 0) / rs.length;
    Object.keys(MKT_FLAG_OF).forEach(function(nm) {
      var fn = defs[MKT_FLAG_OF[nm]];
      if (fn) vals[nm] = rows.filter(function(fr) { try { return !!fn(fr); } catch (e) { return false; } }).length / n;
    });
    var cnt = function(f) { return rs.filter(function(r) { try { return f(r).length > 0; } catch (e) { return false; } }).length / rs.length; };
    vals['W命中比例'] = cnt(matchedBtWinCombos);
    vals['F命中比例'] = cnt(matchedBtHotCombos);
    var comps = {}, hp = [], cp = [];
    [['HEAT', hp], ['COLD', cp]].forEach(function(sd) {
      MKT_CALIB[sd[0]].forEach(function(nm) {
        if (vals[nm] == null) return;
        var p = calibPct(MKT_CALIB.Q[nm], vals[nm]);
        comps[nm] = [vals[nm], p, sd[0]]; sd[1].push(p);
      });
    });
    if (!hp.length || !cp.length) return { n: n, temp: null, zone: null, comps: comps };
    var mean = function(a) { return a.reduce(function(s, x) { return s + x; }, 0) / a.length; };
    var temp = (mean(hp) - mean(cp) + 1) / 2 * 100, zone = null;
    for (var i = 0; i < TEMP_ZONES.length; i++) if (temp < TEMP_ZONES[i][0]) { zone = TEMP_ZONES[i][1]; break; }
    return { n: n, temp: Math.round(temp * 10) / 10, zone: zone, comps: comps };
  }
  function marketKlinePosition(closes) {
    var c = (closes || []).filter(function(x) { return x != null && isFinite(x); }).map(Number);
    if (c.length < 61) return null;
    var ma = function(n) { var s = 0; for (var i = c.length - n; i < c.length; i++) s += c[i]; return s / n; };
    var ma5 = ma(5), ma20 = ma(20), ma60 = ma(60), last = c[c.length - 1];
    var dd60 = last / Math.max.apply(null, c.slice(-60)) - 1, up = null, dn = null;
    for (var i = 1; i < c.length; i++) {
      var d = c[i] - c[i - 1], u = Math.max(d, 0), w = Math.max(-d, 0);
      up = up == null ? u : up + (u - up) / 14; dn = dn == null ? w : dn + (w - dn) / 14;
    }
    var rsi = !dn ? 100 : 100 - 100 / (1 + up / dn);
    var bear = ma5 < ma20 && ma20 < ma60 && last < ma20, bull = ma5 > ma20 && ma20 > ma60 && last > ma20;
    var zone = (dd60 <= -0.10 || rsi < 35 || bear) ? '超賣' : (rsi > 65 ? '過熱' : '正常');
    return { zone: zone, dd60: Math.round(dd60 * 10000) / 100, rsi: Math.round(rsi * 10) / 10, ma: bull ? '多頭排列' : (bear ? '空頭排列' : '糾結') };
  }
  function zoneHist(kind, zone) {
    var s = MKT_CALIB[kind][zone];
    if (!s) return '';
    return '歷史上未來20日等權大盤平均 ' + (s.mkt20 >= 0 ? '+' : '') + s.mkt20.toFixed(2) + '%、上漲機率 ' + Math.round(s.up) + '%' +
      (s.sWin != null ? '、S選股型20日勝率 ' + s.sWin.toFixed(1) + '%' : '') + '（' + s.days + '天）';
  }
  // 回傳 { lines: [[icon, 標題, 重點, 說明], ...], verdict: [顏色, 文字], mt, kp }
  function marketPanel() {
    var closes = currentBenchmark && currentBenchmark.arr ? currentBenchmark.arr.map(function(b) { return b.close; }) : null;
    var kp = marketKlinePosition(closes), mt = marketTemperature(), lines = [];
    if (mt.temp != null) lines.push(['🌡️', '大盤溫度計', Math.round(mt.temp) + '（' + mt.zone + '）', '由 ' + mt.n + ' 檔股票的訊號廣度算出 → ' + zoneHist('TEMP', mt.zone)]);
    else lines.push(['🌡️', '大盤溫度計', '樣本不足', '這次只分析 ' + mt.n + ' 檔，少於 ' + MKT_MIN_STOCKS + ' 檔無法代表全市場（請用全市場／大型指數清單）']);
    if (kp) lines.push(['📉', '大盤K線位置', kp.zone, '距60日高 ' + (kp.dd60 >= 0 ? '+' : '') + kp.dd60.toFixed(1) + '%、RSI ' + Math.round(kp.rsi) + '、均線' + kp.ma + ' → ' + zoneHist('KLINE', kp.zone)]);
    var hot = ((mt.zone === '偏熱' || mt.zone === '過熱') ? 1 : 0) + (kp && kp.zone === '過熱' ? 1 : 0);
    var cold = ((mt.zone === '極冷' || mt.zone === '偏冷') ? 1 : 0) + (kp && kp.zone === '超賣' ? 1 : 0);
    var verdict = cold === 2 ? ['#4caf50', '🟢 溫度計與K線都偏冷／超賣：歷史上之後反彈機率高，可提高部位，S選股型與W擇時型都適用'] :
      hot === 2 ? ['#ff5555', '🔴 溫度計與K線都偏熱／過熱：歷史上之後一個月偏弱，宜減碼、偏重穩健型組合'] :
      ['#aaaaaa', '⚪ 訊號中性或不一致：維持正常部位，依「大盤適用」欄挑組合'];
    return { lines: lines, verdict: verdict, mt: mt, kp: kp, guide: marketGuideLines(mt, kp) };
  }


  var MKT_POS_TEMP = { '極冷': 2, '偏冷': 1, '中性': 0, '偏熱': -1, '過熱': -2 }, MKT_POS_KLINE = { '超賣': 2, '正常': 0, '過熱': -2 };
  var MKT_POS_MAP = { '4': [90, '大幅加碼'], '3': [80, '加碼'], '2': [70, '加碼'], '1': [60, '小幅加碼'], '0': [50, '維持'],
    '-1': [40, '小幅減碼'], '-2': [30, '減碼'], '-3': [20, '減碼'], '-4': [10, '大幅減碼'] };
  // 溫度計＋K線位置 → [建議持股%, 加減碼, 分數, 說明]；中性＝50%
  function marketPosition(mt, kp) {
    var parts = [], sc = 0, v;
    if (mt && mt.zone) { v = MKT_POS_TEMP[mt.zone]; sc += v; parts.push('溫度計' + mt.zone + ' ' + (v >= 0 ? '+' : '') + v); }
    if (kp) { v = MKT_POS_KLINE[kp.zone]; sc += v; parts.push('K線' + kp.zone + ' ' + (v >= 0 ? '+' : '') + v); }
    if (!parts.length) return null;
    var m = MKT_POS_MAP[String(Math.max(-4, Math.min(4, sc)))];
    return [m[0], m[1], sc, parts.join('、')];
  }
  // 大盤指引：把溫度計＋K線位置翻成白話（判定理由、操作建議、接下來要觀察什麼）
  function marketGuideLines(mt, kp) {
    var out = [], f1 = function(x) { return (x >= 0 ? '+' : '') + x.toFixed(1); };
    var ps = marketPosition(mt, kp);
    if (ps) out.push('・<b style="color:#ffd54f">建議持股 ' + ps[0] + '%（' + ps[1] + '）</b>　分數 ' + (ps[2] >= 0 ? '+' : '') + ps[2] + '（' + ps[3] + '；中性＝50%，每 ±1 分調 10%）' + (mt && mt.zone ? '' : '　⚠️ 溫度計樣本不足，只依K線判斷'));
    if (kp) {
      var dd = kp.dd60, rsi = kp.rsi, ma = kp.ma;
      var why = ['距60日高 ' + f1(dd) + '%（' + (dd <= -10 ? '已達' : '未達') + ' -10% 超賣門檻）',
        'RSI ' + Math.round(rsi) + '（' + (rsi < 35 ? '<35 超賣' : rsi > 65 ? '>65 過熱' : '介於35～65') + '）',
        '均線' + ma + (ma === '空頭排列' ? '（空頭排列＝超賣條件）' : '')];
      var mean = { '正常': '不在超賣也不在過熱，沒有明顯方向優勢，歷史統計大致等於平常水準',
        '超賣': '已進入超賣區，歷史上之後20日反彈機率高，是較好的進場區',
        '過熱': 'RSI 過熱，歷史上之後一個月大盤平均偏弱，追高風險較大' }[kp.zone];
      out.push('・K線位置「' + kp.zone + '」：' + why.join('、') + ' → ' + mean);
    }
    var tz = mt && mt.zone, kz = kp ? kp.zone : null;
    if (tz) out.push('・溫度計「' + tz + '」：' + { '極冷': '多數個股弱勢、低檔訊號多，歷史上是之後報酬最好的區間',
      '偏冷': '整體偏弱、略偏低檔，歷史上之後報酬略優於平常', '中性': '多空力道均衡',
      '偏熱': '多數個股已走強，之後報酬略低於平常', '過熱': '廣度過熱、強勢股比例偏高，歷史上之後報酬最差' }[tz]);
    var coldT = tz === '極冷' || tz === '偏冷', hotT = tz === '偏熱' || tz === '過熱', act;
    if (kz === '超賣' && coldT) act = '兩者都在低檔：較好的進場區，可提高部位，S選股型與W擇時型都適用';
    else if (kz === '過熱' && hotT) act = '兩者都過熱：宜減碼，新部位偏重S選股型中「大盤適用✅」的組合';
    else if (kz === '超賣') act = 'K線超賣但廣度未冷：可能是少數權值股拖累，分批試單，等溫度計轉冷再加碼';
    else if (kz === '過熱') act = 'K線過熱但廣度未熱：指數靠少數權值股撐盤，個股不宜追高';
    else if (coldT) act = '整理中、略偏低檔：部位可略高於中性，照常使用S選股型組合';
    else if (hotT) act = '指數不過熱但個股廣度偏熱：照常操作，留意K線RSI是否轉過熱';
    else act = '中性：維持正常部位，依「大盤適用」欄挑組合';
    out.push('・操作建議：' + act);
    if (kp) {
      var w = [];
      if (kz !== '超賣') w.push('指數再跌約 ' + Math.abs((0.9 / (1 + kp.dd60 / 100) - 1) * 100).toFixed(1) + '%（距60日高達-10%）或 RSI 跌破35 → 轉「超賣」');
      if (kz !== '過熱') w.push('RSI 升破65 → 轉「過熱」');
      w.push(kp.ma === '糾結' ? '均線糾結待選邊：站穩20MA轉多頭排列偏多；跌破並形成空頭排列即轉超賣訊號' :
        kp.ma === '多頭排列' ? '均線多頭排列：跌破20MA才需轉保守' : '均線空頭排列：重新站回20MA才算止跌');
      if (kz !== '超賣' && !coldT) w.push('最佳時機＝K線「超賣」且溫度計同時「偏冷／極冷」');
      out.push('・接下來觀察：' + w.join('；'));
    }
    return out;
  }
  function comboStyleOf(r) {
    var c = hitCategories(r), key = ['S', '#', 'M', 'W', 'F'].filter(function(k) { return c.indexOf(k) >= 0; }).join('');
    return COMBO_STYLE_RULES[key] || '';
  }
  function renderComboStyle(r) {
    var st = comboStyleOf(r);
    if (!st) return '<td class="nw" style="color:#555;text-align:center">－</td>';
    var lk = COMBO_STYLE_LOOK[st];
    return '<td class="nw" style="font-weight:700;color:' + lk[1] + '" title="' + COMBO_STYLE_NOTE[st] + '">' + lk[0] + '</td>';
  }
  function renderHitCategories(r) {
    var c = hitCategories(r), n = c.length;
    var col = n >= 3 ? '#ffd700' : (n === 2 ? '#fff' : (n === 1 ? '#aaa' : '#555'));
    return '<td class="nw" style="font-weight:700;text-align:center;color:' + col + '" title="命中類別：' + (n ? c.join('、') : '無') +
      '（回測：同時命中3類以上，台股20日超額約+2~4%、美股+0.6~1.8%，但波動也較大）">' + n + (n ? '<div style="font-size:9px;font-weight:400;color:#888">' + c.join('') + '</div>' : '') + '</td>';
  }
  function renderMatchedComboBadges(r, part) {
    var matched, matchedMoonshot;
    try { matched = matchedPinnedCombos(r); } catch (e) { matched = []; }
    try { matchedMoonshot = matchedMoonshotCombos(r); } catch (e) { matchedMoonshot = []; }
    var matchedSP;
    try { matchedSP = matchedStockPickCombos(r); } catch (e) { matchedSP = []; }
    if (part === 'S') { matched = []; matchedMoonshot = []; }
    else if (part === 'PM') { matchedSP = []; }
    var spBadges = matchedSP.map(function(idx) {
      var combo = STOCK_PICK_COMBOS[idx];
      var key = combo.join(' ＋ ');
      var stS = comboStatsFor(combo, 'S'), btS = bestT(stS);
      var title = '[選股型] ' + key + '　勝率 ' + fmtHz(stS.win, true) + '　t ' + fmtHz(stS.t) + '　(' + stS.src + ')';
      return '<span style="display:inline-block;margin:1px 2px 1px 0;padding:1px 6px;border-radius:10px;font-size:10px;font-weight:700;color:#00c864;border:1px solid #00c86488" title="'+title+'">S' + (idx+1) + (btS != null ? ' t' + btS.toFixed(1) : '') + '</span>';
    });
    var badges = matched.map(function(idx) {
      var combo = PINNED_COMBOS[idx];
      var comboKey = combo.join(' ＋ ');
      var wrObj = PINNED_COMBO_WINRATES[comboKey];
      var wrText = formatWinrates(wrObj);
      var timing = isTimingCombo(combo);
      var stP = comboStatsFor(combo, '#'), btP = bestT(stP);
      var title = (timing ? '[高勝率·擇時型：勝率主要來自大盤反彈] ' : '[高勝率] ') + comboKey + (wrText ? '　(歷史勝率：' + wrText + ')' : '') +
        '　勝率 ' + fmtHz(stP.win, true) + '　t ' + fmtHz(stP.t) + (stP.src ? '　(' + stP.src + ')' : '');
      var badgeText = '#' + (idx+1) + (timing ? '⏱' : '') + (btP != null ? ' t' + btP.toFixed(1) : '');   // 勝率等完整數字放在滑鼠提示（title）
      return '<span style="display:inline-block;margin:1px 2px 1px 0;padding:1px 6px;border-radius:10px;font-size:10px;color:#7c3aed;border:1px solid #7c3aed55" title="'+title+'">'+badgeText+'</span>';
    });
    var moonshotBadges = matchedMoonshot.map(function(idx) {
      var combo = MOONSHOT_COMBOS[idx];
      var comboKey = combo.join(' ＋ ');
      var statsObj = MOONSHOT_COMBO_STATS[comboKey];
      var statsText = formatMoonshotStats(statsObj);
      var stM = comboStatsFor(combo, 'M'), btM = bestT(stM);
      var title = '[高標股] ' + comboKey + (statsText ? '　' + statsText : '') + '　飆股比例 ' + fmtHz(stM.moon, true) + '　t ' + fmtHz(stM.t) + (stM.src ? '　(' + stM.src + ')' : '');
      // 徽章上簡短顯示最長天期（優先20日，其次10日、5日）的飆股比例，完整資訊放title
      var shortHorizon = [20,10,5].filter(function(h){ return statsObj && statsObj[h]; })[0];
      var badgeText = 'M' + (idx+1) + (shortHorizon ? ' ' + statsObj[shortHorizon].pct + '%' : '') + (btM != null ? ' t' + btM.toFixed(1) : '');
      return '<span style="display:inline-block;margin:1px 2px 1px 0;padding:1px 6px;border-radius:10px;font-size:10px;color:#ff9800;border:1px solid #ff980055" title="'+title+'">'+badgeText+'</span>';
    });
    // 回測自動入選：W＝青色、F＝粉紅色，每類只顯示最強的 BT_HIT_SHOW 個
    var extra = '';
    if (part !== 'S') {
      var mw = [], mf = [];
      try { mw = matchedBtWinCombos(r); } catch (e) {}
      try { mf = matchedBtHotCombos(r); } catch (e) {}
      var maxMoon = function(st) { var v = Object.keys(st.moon || {}).map(function(h){ return st.moon[h]; }).filter(function(x){ return x != null; }); return v.length ? Math.max.apply(null, v) : 0; };
      var wl = mw.map(function(i){ return { i: i, st: comboStatsFor(BT_WIN_COMBOS[i], 'W') }; }).sort(function(a, b){ return (bestT(b.st) || 0) - (bestT(a.st) || 0); });
      var fl = mf.map(function(i){ return { i: i, st: comboStatsFor(BT_HOT_COMBOS[i], 'F') }; }).sort(function(a, b){ return maxMoon(b.st) - maxMoon(a.st); });
      wl.slice(0, BT_HIT_SHOW).forEach(function(o) {
        var c = BT_WIN_COMBOS[o.i], bt = bestT(o.st), timing = isTimingCombo(c);
        var title = (timing ? '[回測高勝率·擇時型] ' : '[回測高勝率] ') + c.join(' ＋ ') + '　' + btWinText(BT_WIN_STATS[c.join(' ＋ ')]);
        extra += '<span style="display:inline-block;margin:1px 2px 1px 0;padding:1px 6px;border-radius:10px;font-size:10px;color:#00bcd4;border:1px solid #00bcd477" title="' + title + '">W' + (o.i+1) + (timing ? '⏱' : '') + (bt != null ? ' t' + bt.toFixed(1) : '') + '</span>';
      });
      if (wl.length > BT_HIT_SHOW) extra += '<span style="font-size:10px;color:#00bcd4" title="其餘 ' + (wl.length - BT_HIT_SHOW) + ' 組回測高勝率組合">+' + (wl.length - BT_HIT_SHOW) + '</span> ';
      fl.slice(0, BT_HIT_SHOW).forEach(function(o) {
        var c = BT_HOT_COMBOS[o.i];
        var title = '[回測飆股] ' + c.join(' ＋ ') + '　' + formatMoonshotStats(BT_HOT_STATS[c.join(' ＋ ')]);
        extra += '<span style="display:inline-block;margin:1px 2px 1px 0;padding:1px 6px;border-radius:10px;font-size:10px;color:#e91e63;border:1px solid #e91e6377" title="' + title + '">F' + (o.i+1) + ' ' + maxMoon(o.st) + '%</span>';
      });
      if (fl.length > BT_HIT_SHOW) extra += '<span style="font-size:10px;color:#e91e63" title="其餘 ' + (fl.length - BT_HIT_SHOW) + ' 組回測飆股組合">+' + (fl.length - BT_HIT_SHOW) + '</span>';
    }
    var html = spBadges.join('') + badges.join('') + moonshotBadges.join('') + extra;
    return html || '<span style="color:#555;font-size:11px">--</span>';
  }

  // ── 回後買上漲 8條件判斷 ──
  function checkPullbackBuy(data) {
    var last  = data[data.length - 1];
    var prev  = data[data.length - 2] || last;
    var prev2 = data[data.length - 3] || prev;

    var results = [];
    var allPass = true;

    // 轉折波（跟型態辨識、圖表轉折波用同一套 buildZigzag，不再用另一套獨立的分形判斷法）
    var zz = buildZigzag(data);
    var zzHighs = zz.filter(function(p){ return p.type === 'H'; });
    var zzLows  = zz.filter(function(p){ return p.type === 'L'; });

    // ── 條件1：趨勢多頭（轉折波 高高低低）──
    var c1 = false;
    if (zzHighs.length >= 2 && zzLows.length >= 2) {
      var rh = [zzHighs[zzHighs.length-2].price, zzHighs[zzHighs.length-1].price];
      var rl = [zzLows[zzLows.length-2].price,  zzLows[zzLows.length-1].price];
      c1 = (rh[1] > rh[0] && rl[1] > rl[0]);
    }
    results.push({ label:'①趨勢多頭（高高低低）', pass:c1, required:true });
    if (!c1) allPass = false;

    // ── 條件2：位置回後上漲（前1~5日曾出現縮量回檔，今日反彈）──
    // 依課程講義：回檔幅度用費波那契回檔比例分級，回檔愈淺（愈接近0.382）代表股票愈強，
    // 愈要把握機會進場；回檔愈深（接近或超過0.618）力道愈弱。
    var pullbackDays = data.slice(-6, -1); // 前5日
    var hadPullback = pullbackDays.some(function(d, i) {
      return d.close < (i>0 ? pullbackDays[i-1].close : d.close) || d.close < d.open;
    });
    var c2 = hadPullback && last.close > prev.close;
    var pullbackDetail = '';
    if (zzHighs.length >= 1) {
      var peakPoint = zzHighs[zzHighs.length - 1];
      var priorLowCands = zzLows.filter(function(p){ return p.idx < peakPoint.idx; });
      if (priorLowCands.length) {
        var priorLow = priorLowCands[priorLowCands.length - 1].price;
        var peak = peakPoint.price;
        if (peak > priorLow) {
          var pullbackSegment = data.slice(peakPoint.idx);
          var pullbackLow = Math.min.apply(null, pullbackSegment.map(function(d){ return d.low; }));
          var retrace = (peak - pullbackLow) / (peak - priorLow);
          var grade = retrace <= 0.382 ? '最強' : retrace <= 0.5 ? '強' : retrace <= 0.618 ? '弱' : '回檔過深';
          pullbackDetail = '回檔幅度' + (retrace * 100).toFixed(1) + '%（' + grade + '，費波0.382/0.5/0.618分級）';
        }
      }
    }
    results.push({ label:'②位置回後上漲（近期有回檔，今轉上）', pass:c2, required:true, detail: pullbackDetail });
    if (!c2) allPass = false;

    // ── 條件3：站上5MA（平價不算，必須 > ）──
    var c3 = last.ma5 != null && last.close > last.ma5;
    results.push({ label:'③收盤站上5MA（平價不算）', pass:c3, required:true,
      detail: last.ma5 != null ? '5MA=' + last.ma5.toFixed(2) + '  收盤=' + last.close.toFixed(2) : '' });
    if (!c3) allPass = false;

    // ── 條件4：過前一日高點（含上影線，即今高 > 昨高）──
    var c4 = last.high > prev.high;
    results.push({ label:'④突破前一日高點（含上影線）', pass:c4, required:true,
      detail:'今高=' + last.high.toFixed(2) + '  昨高=' + prev.high.toFixed(2) });
    if (!c4) allPass = false;

    // ── 條件5：漲幅2%以上（收盤 vs 昨收）──
    var chgPct = prev.close > 0 ? (last.close - prev.close) / prev.close * 100 : 0;
    var c5 = chgPct >= 2.0;
    results.push({ label:'⑤漲幅2%以上', pass:c5, required:true,
      detail:'漲幅=' + chgPct.toFixed(2) + '%' });
    if (!c5) allPass = false;

    // ── 條件6：實體紅K，影線不大於實體──
    var body    = last.close - last.open;                          // 必須 > 0 才是紅K
    var upSh    = last.high  - last.close;                         // 上影線
    var dnSh    = last.open  - last.low;                           // 下影線
    var maxSh   = Math.max(upSh, dnSh);
    var c6      = body > 0 && maxSh <= body;                       // 紅K & 影線≤實體
    results.push({ label:'⑥實體紅K，影線不大於實體', pass:c6, required:true,
      detail:'實體=' + body.toFixed(2) + '  最大影線=' + maxSh.toFixed(2) });
    if (!c6) allPass = false;

    // ── 條件7：成交量增（加分選項）──
    var volRatio = last.vm20 ? last.volume / last.vm20 : 1;
    var c7 = volRatio >= 1.0;
    results.push({ label:'⑦成交量增（加分項）', pass:c7, required:false,
      detail:'量比MA20=' + volRatio.toFixed(2) + 'x' });
    // 不影響 allPass

    // ── 條件8：KD指標確認（K值向上，K > 昨K）──
    var prevKD = data[data.length - 2];
    var c8 = false;
    if (last.kdK != null && prevKD && prevKD.kdK != null) {
      c8 = last.kdK > prevKD.kdK;  // K值向上
    }
    results.push({ label:'⑧指標確認（K值向上）', pass:c8, required:true,
      detail: last.kdK != null ? 'K=' + last.kdK.toFixed(1) + '  昨K=' + (prevKD&&prevKD.kdK!=null?prevKD.kdK.toFixed(1):'--') : 'KD資料不足' });
    if (!c8) allPass = false;

    // 計算通過數（必要項）
    var requiredTotal   = results.filter(function(r){ return r.required; }).length;
    var requiredPassed  = results.filter(function(r){ return r.required && r.pass; }).length;
    var bonusPassed     = results.filter(function(r){ return !r.required && r.pass; }).length;

    return { results:results, allPass:allPass, requiredPassed:requiredPassed, requiredTotal:requiredTotal, bonusPassed:bonusPassed };
  }

  // ── 型態確認：朱家泓 進場型態（6種底部型態＋ABC切線＋上升軌道＋大量黑K＋回後買上漲）──
  // (1)頭肩底 (2)複式頭肩底 (3)N字底 (4)三重底 (5)圓弧底 (6)一字底(均線糾結)
  // 夜星、母子懷抱(高檔)：原本視為看空而排除在彙總旗標外；2026-09-27 台股＋美股回測都顯示強勢股出現
  // 這兩種型態多半是洗盤後續漲，改稱「高檔強勢整理」，照常計入彙總旗標（BEARISH_PATTERN_IDS 留空保留介面）。
  var BEARISH_PATTERN_IDS = {};
  var HIGH_CONSOLIDATION_IDS = { harami_bear: true, evening_star: true };
  function tolerant(a, b, pct) {
    var base = Math.max(Math.abs(a), Math.abs(b), 1e-6);
    return Math.abs(a - b) / base <= pct;
  }

  function detectPatterns(data, pb, skipJustBroke) {
    // 型態辨識固定看跟圖表一致的完整資料範圍（也就是「分析天數」實際抓到的全部K棒），
    // 並套用同一套壞資料過濾規則，確保型態辨識用的轉折波，跟圖表上實際畫出來的
    // 轉折波／輔助線是同一組資料算出來的結果——避免用不同範圍算出對不起來的轉折點。
    data = data.filter(function(d) {
      return d && d.open > 0 && d.high > 0 && d.low > 0 && d.close > 0 &&
        isFinite(d.open) && isFinite(d.high) && isFinite(d.low) && isFinite(d.close);
    });
    var last = data.length - 1;
    var lastClose = data[last].close;
    var lastVol   = data[last].volume;
    var vm20      = data[last].vm20;
    var volConfirm = vm20 ? (lastVol / vm20 >= 1.3) : false;

    // 型態辨識所用的高低點，改成直接沿用「轉折波」(buildZigzag) 算出來的同一組轉折點，
    // 不再用另一套獨立的分形視窗判斷法——這樣圖表上畫出來的轉折波，就是型態辨識實際依據的高低點，
    // 兩者完全一致，不會有「圖上看到的轉折」跟「型態判斷用的轉折」對不起來的狀況。
    var zz = buildZigzag(data);
    // 型態辨識用的轉折點範圍，直接沿用整個(已限制在120根K棒內的)資料範圍，跟圖表顯示範圍完全一致，
    // 不再另外疊加一層90天子視窗限制——避免「圖表上看得到的高低點」卻被排除在型態判斷之外，
    // 導致畫出來的頸線/壓力線跟圖上真正的高低點對不起來。
    var floor = 0;
    var lows  = zz.filter(function(p) { return p.type === 'L' && p.idx >= floor && p.idx < last; }).map(function(p){ return p.idx; });
    var highs = zz.filter(function(p) { return p.type === 'H' && p.idx >= floor && p.idx < last; }).map(function(p){ return p.idx; });

    function breakoutCheck(resistance) {
      if (resistance == null) return { confirmed: false, detail: '' };
      return {
        confirmed: lastClose > resistance,
        detail: '頸線/壓力＝' + resistance.toFixed(2) + '　現價＝' + lastClose.toFixed(2) + (volConfirm ? '　(帶量突破)' : '')
      };
    }

    // MA 是否維持上揚（近 n 根K棒比較）
    function maSlopeUp(key, n) {
      var p = data[Math.max(0, last - n)], c = data[last];
      if (c[key] == null || p[key] == null) return false;
      return c[key] > p[key];
    }

    var results = [];

    // ── (1) 頭肩底：左肩-頭-右肩，頭最低，左右肩相近，頸線＝兩肩間高點連線，右肩不破頭→頸線1/2 ──
    (function () {
      var id = 'hs', name = '頭肩底';
      if (lows.length >= 3) {
        var l3 = lows.slice(-3);
        var L1 = data[l3[0]].low, L2 = data[l3[1]].low, L3 = data[l3[2]].low;
        var shouldersSimilar = tolerant(L1, L3, 0.06);
        var headLower = L2 < L1 * 0.985 && L2 < L3 * 0.985;
        if (shouldersSimilar && headLower) {
          var hBetween = highs.filter(function(i) { return i > l3[0] && i < l3[2]; });
          var neck = hBetween.length ? Math.max.apply(null, hBetween.map(function(i){ return data[i].high; })) : null;
          // 依課程講義：右肩(L3)不能跌破 頭(L2)→頸線 漲幅的 1/2，才是有效的右肩（不是單純比頭高就好）
          var validRightShoulder = neck != null && L3 > (L2 + neck) / 2;
          if (validRightShoulder) {
            var bo = breakoutCheck(neck);
            results.push({ id:id, name:name, formed:true, breakout:bo.confirmed,
              detail: bo.detail || '型態成形，等待突破頸線', desc:'左右肩低點相近，頭部最低，右肩不破1/2，突破頸線為買點',
              line: { i1: l3[0], p1: neck, slope: 0 } });
            return;
          }
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構', desc:'左右肩低點相近，頭部最低，右肩不破1/2，突破頸線為買點' });
    })();

    // ── (2) 複式頭肩底：多個相近肩部低點夾一個最低頭部，最近肩部不破頭→頸線1/2 ──
    (function () {
      var id = 'chs', name = '複式頭肩底';
      if (lows.length >= 4) {
        var lastLows = lows.slice(-5);
        var lowVals = lastLows.map(function(i){ return data[i].low; });
        var minVal = Math.min.apply(null, lowVals);
        var headIdxPos = lowVals.indexOf(minVal);
        var hasLeft  = headIdxPos > 0;
        var hasRight = headIdxPos < lastLows.length - 1;
        var shoulderVals = lowVals.filter(function(_, idx) { return idx !== headIdxPos; });
        var shouldersOK = hasLeft && hasRight && shoulderVals.every(function(v) {
          return tolerant(v, shoulderVals[0], 0.08) && v > minVal * 1.02;
        });
        if (shouldersOK) {
          var headIdx = lastLows[headIdxPos];
          var hBetween = highs.filter(function(i) { return i > lastLows[0] && i < lastLows[lastLows.length-1]; });
          var neck = hBetween.length ? Math.max.apply(null, hBetween.map(function(i){ return data[i].high; })) : null;
          // 依課程講義：最近（最右）一個肩部不能跌破 頭→頸線 漲幅的 1/2
          var lastShoulderVal = lowVals[lowVals.length - 1];
          var validLastShoulder = neck != null && lastShoulderVal > (minVal + neck) / 2;
          if (validLastShoulder) {
            var bo = breakoutCheck(neck);
            results.push({ id:id, name:name, formed:true, breakout:bo.confirmed,
              detail: bo.detail || '型態成形，等待突破頸線', desc:'多重肩部低點環繞單一最低頭部，最近肩部不破1/2，突破頸線為買點',
              line: { i1: lastLows[0], p1: neck, slope: 0 } });
            return;
          }
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構', desc:'多重肩部低點環繞單一最低頭部，最近肩部不破1/2，突破頸線為買點' });
    })();

    // ── (3) N字底：低點A → 反彈高點B → 拉回不破1/2（依課程講義：拉回不深於A-B漲幅的一半）C → 突破B為買點 ──
    (function () {
      var id = 'nb', name = 'N字底';
      if (lows.length >= 2 && highs.length >= 1) {
        var A = lows[lows.length - 2], C = lows[lows.length - 1];
        var Bcands = highs.filter(function(i) { return i > A && i < C; });
        if (Bcands.length) {
          // 取A、C之間「最高」的確認高點，而不是「最近」的一個——
          // 若下跌過程中出現次要反彈小高點，會比真正的主高點更晚被確認，
          // 用「最近」選到的話會抓到錯誤（偏低）的壓力位置。
          var B = Bcands.reduce(function(best, i) { return data[i].high > data[best].high ? i : best; }, Bcands[0]);
          var lowA = data[A].low, lowC = data[C].low, highB = data[B].high;
          // 依課程講義：C（拉回低點）不能跌破 A→B 漲幅的 1/2，才是有效的淺拉回（不是單純比A高就好）
          var halfPoint = (lowA + highB) / 2;
          var shallowPullback = lowC > halfPoint;
          if (shallowPullback) {
            var bo = breakoutCheck(highB);
            results.push({ id:id, name:name, formed:true, breakout:bo.confirmed,
              detail: bo.detail || ('拉回未破1/2（' + halfPoint.toFixed(2) + '），等待突破反彈高點'), desc:'低點反彈後拉回不破1/2，再突破反彈高點為買點',
              line: { i1: B, p1: highB, slope: 0 } });
            return;
          }
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構', desc:'低點反彈後拉回不破1/2，再突破反彈高點為買點' });
    })();

    // ── (4) 三重底：三個相近低點，兩個中間高點也要相近（形成真正的水平頸線）──
    (function () {
      var id = 'tb', name = '三重底';
      if (lows.length >= 3) {
        var l3 = lows.slice(-3);
        var vals = l3.map(function(i){ return data[i].low; });
        var allSimilar = tolerant(vals[0], vals[1], 0.08) && tolerant(vals[1], vals[2], 0.08) && tolerant(vals[0], vals[2], 0.08);
        if (allSimilar) {
          var hBetween1 = highs.filter(function(i) { return i > l3[0] && i < l3[1]; });
          var hBetween2 = highs.filter(function(i) { return i > l3[1] && i < l3[2]; });
          var peak1 = hBetween1.length ? Math.max.apply(null, hBetween1.map(function(i){ return data[i].high; })) : null;
          var peak2 = hBetween2.length ? Math.max.apply(null, hBetween2.map(function(i){ return data[i].high; })) : null;
          // 依課程講義：兩個中間高點要相近，才是真正的水平頸線（不是隨便夾兩個高低不一的高點）
          var necklineOK = peak1 != null && peak2 != null && tolerant(peak1, peak2, 0.05);
          if (necklineOK) {
            var res = Math.max(peak1, peak2);
            var bo = breakoutCheck(res);
            results.push({ id:id, name:name, formed:true, breakout:bo.confirmed,
              detail: bo.detail || '型態成形，等待突破壓力', desc:'三個低點高度相近，中間兩高點形成水平頸線，突破頸線為買點',
              line: { i1: l3[0], p1: res, slope: 0 } });
            return;
          }
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構', desc:'三個低點高度相近，中間兩高點形成水平頸線，突破頸線為買點' });
    })();

    // ── (5) 圓弧底：價格緩跌後緩升，呈U型弧線，突破起跌點壓力為買點 ──
    (function () {
      var id = 'rb', name = '圓弧底';
      var win = data.slice(-40);
      if (win.length >= 30) {
        var seg = Math.floor(win.length / 3);
        var first = win.slice(0, seg), mid = win.slice(seg, win.length - seg), tail = win.slice(win.length - seg);
        function avg(arr, key) { return arr.reduce(function(s,d){ return s + d[key]; }, 0) / arr.length; }
        function slope(arr) {
          var n = arr.length, sx=0, sy=0, sxy=0, sxx=0;
          arr.forEach(function(d, i) { sx+=i; sy+=d.close; sxy+=i*d.close; sxx+=i*i; });
          var denom = (n*sxx - sx*sx) || 1;
          return (n*sxy - sx*sy) / denom;
        }
        var slopeFirst = slope(first), slopeTail = slope(tail);
        var midLow = Math.min.apply(null, mid.map(function(d){ return d.low; }));
        var isConvex = avg(first,'close') > midLow * 1.01 && avg(tail,'close') > midLow * 1.01;
        var shapeOK = slopeFirst < 0 && slopeTail > 0 && isConvex;
        // 低波動：日均振幅相對價格不宜過大
        var avgRange = win.reduce(function(s,d){ return s + (d.high-d.low)/d.close; }, 0) / win.length;
        var lowVol = avgRange < 0.05;
        if (shapeOK && lowVol) {
          var resistance = Math.max.apply(null, first.map(function(d){ return d.high; }));
          var bo = breakoutCheck(resistance);
          var winStartIdx = data.length - win.length;
          // 依課程講義：目標價 = 突破點 + 型態高度（起跌點高點 - 最低點），即測量移動法
          var target = resistance + (resistance - midLow);
          results.push({ id:id, name:name, formed:true, breakout:bo.confirmed,
            detail: (bo.detail || '弧形築底中，等待突破起跌壓力') + '　目標價≈' + target.toFixed(2),
            desc:'價格緩跌後緩升成U型，突破起跌點高點為買點',
            line: { i1: winStartIdx, p1: resistance, slope: 0 } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構', desc:'價格緩跌後緩升成U型，突破起跌點高點為買點' });
    })();

    // ── (6) 一字底（均線糾結）：MA5/10/20/60 長期糾結收斂，整理區間範圍約10%內，帶量突破整理區間 ──
    (function () {
      var id = 'fb', name = '一字底(均線糾結)';
      var N = 10;
      var win = data.slice(-N - 1, -1); // 不含今日，前N日整理區
      var ok = win.length === N && win.every(function(d) {
        return d.ma5 != null && d.ma10 != null && d.ma20 != null && d.ma60 != null;
      });
      if (ok) {
        var tangled = win.every(function(d) {
          var vals = [d.ma5, d.ma10, d.ma20, d.ma60];
          var mx = Math.max.apply(null, vals), mn = Math.min.apply(null, vals);
          return (mx - mn) / mn <= 0.05;
        });
        // 課程講義定義：整理區間範圍（整段最高與最低價的差距）約在 10% 以內，才算「一字底」
        var winHigh = Math.max.apply(null, win.map(function(d){ return d.high; }));
        var winLow  = Math.min.apply(null, win.map(function(d){ return d.low; }));
        var narrowRange = (winHigh - winLow) / winLow <= 0.10;
        if (tangled && narrowRange) {
          var resistance = winHigh;
          var last4 = [data[last].ma5, data[last].ma10, data[last].ma20, data[last].ma60];
          var aboveAllMA = last4.every(function(v){ return v != null && lastClose > v; });
          var breakout = lastClose > resistance && aboveAllMA && volConfirm;
          var winStartIdx6 = data.length - N - 1;
          // 依課程講義：目標價 = 突破點 + 型態高度（整理區間高點 - 整理區間低點）
          var target6 = resistance + (resistance - winLow);
          results.push({ id:id, name:name, formed:true, breakout:breakout,
            detail: '整理區間高點＝' + resistance.toFixed(2) + '　現價＝' + lastClose.toFixed(2) + (volConfirm ? '　(帶量突破)' : '　(尚未帶量)') + '　目標價≈' + target6.toFixed(2),
            desc:'均線糾結、價格窄幅整理（區間範圍約10%內），帶量突破整理區間為買點',
            line: { i1: winStartIdx6, p1: resistance, slope: 0 } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構', desc:'均線糾結、價格窄幅整理（區間範圍約10%內），帶量突破整理區間為買點' });
    })();

    // ── (7) 突破ABC修正下降切線：多頭回檔呈ABC三段式下跌(約20天內)，A、B兩高點畫下降切線，
    //      B之後要有C段拉回低點，MA20上揚，帶量紅K突破切線為買點 ──
    (function () {
      var id = 'abc', name = '突破ABC修正下降切線';
      // 只看「最近」的轉折高點，取最後兩個（A、B），避免抓到太久遠、已經沒有參考意義的舊高點
      var recentHighs = highs.filter(function(i) { return i >= last - 40; });
      if (recentHighs.length >= 2) {
        var h1 = recentHighs[recentHighs.length - 2]; // A：較早、較高的反彈高點
        var h2 = recentHighs[recentHighs.length - 1]; // B：較近、較低的反彈高點
        var y1 = data[h1].high, y2 = data[h2].high;
        // A、B之間要有拉回的低點（確認A→低點→B是有效的一次反彈，不是隨便兩個高點連線）
        var hasLowBetween = lows.some(function(li) { return li > h1 && li < h2; });
        // 依課程講義：B之後還要有一段「C」低點（真正的拉回），突破才是站在C低點之上完成的，
        // 不能B之後價格根本沒拉回就直接算突破——那樣就只是A、B兩個高點連線，不是完整的ABC三段修正。
        var cLow = null;
        for (var ci = h2 + 1; ci < last; ci++) {
          if (cLow === null || data[ci].low < cLow) cLow = data[ci].low;
        }
        // 至少要有2%以上的拉回幅度，才算真正的C段（避免單純的價格雜訊被誤判成有效拉回）
        var hasCLeg = cLow !== null && cLow < data[h2].close * 0.98;
        if (y2 < y1 && h2 > h1 && hasLowBetween && hasCLeg && (h2 - h1) <= 20 && (last - h2) <= 20) {
          var slope = (y2 - y1) / (h2 - h1);
          var lineAtLast = y1 + slope * (last - h1);
          var ma20up = maSlopeUp('ma20', 10);
          var isRed  = data[last].close > data[last].open;
          var breakout = lastClose > lineAtLast && ma20up && isRed;
          results.push({ id:id, name:name, formed:true, breakout:breakout,
            detail: '下降切線位置≈' + lineAtLast.toFixed(2) + '　C低點＝' + cLow.toFixed(2) + '　現價＝' + lastClose.toFixed(2) + (ma20up ? '　MA20上揚' : '　MA20未上揚'),
            desc:'多頭回檔呈ABC三段式下跌，A、B高點畫下降切線，C段拉回後帶量紅K突破切線為買點',
            line: { i1: h1, p1: y1, slope: slope } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'多頭回檔呈ABC三段式下跌，A、B高點畫下降切線，C段拉回後帶量紅K突破切線為買點' });
    })();

    // ── (8) 突破上升軌道線：股價沿上升通道緩步上漲，帶量長紅收盤突破軌道上緣 ──
    (function () {
      var id = 'channel', name = '突破上升軌道線';
      var floor2 = Math.max(0, data.length - 60);
      var lowsIn  = lows.filter(function(i) { return i >= floor2; });
      var highsIn = highs.filter(function(i) { return i >= floor2; });
      if (lowsIn.length >= 2) {
        var l1 = lowsIn[lowsIn.length - 2], l2 = lowsIn[lowsIn.length - 1];
        var ly1 = data[l1].low, ly2 = data[l2].low;
        if (ly2 > ly1 && l2 > l1) {
          var slope2 = (ly2 - ly1) / (l2 - l1);
          var offsetCands = highsIn.filter(function(i) { return i > l1; }).map(function(i) {
            return data[i].high - (ly1 + slope2 * (i - l1));
          });
          // 依課程講義：上升軌道線要有「至少2個高點」貼著同一條平行線，才是真正的軌道線
          // （不能只是單一個高點恰好離支撐線最遠，那只是巧合，不是真的通道）
          // 容忍度用股價的3%來算（而不是用offset自身的百分比），避免offset數值太小時誤判過嚴
          var offset = null;
          if (offsetCands.length >= 2) {
            var tol = lastClose * 0.03;
            var bestOffset = null, bestCount = 0;
            offsetCands.forEach(function(o) {
              var count = offsetCands.filter(function(o2) { return Math.abs(o - o2) <= tol; }).length;
              if (count > bestCount || (count === bestCount && (bestOffset === null || o > bestOffset))) {
                bestCount = count; bestOffset = o;
              }
            });
            if (bestCount >= 2) offset = bestOffset;
          }
          if (offset != null && offset > 0) {
            var upperAtLast = ly1 + slope2 * (last - l1) + offset;
            var ma20up2 = maSlopeUp('ma20', 10);
            var isRed2  = data[last].close > data[last].open;
            var breakout2 = lastClose > upperAtLast && ma20up2 && isRed2 && volConfirm;
            results.push({ id:id, name:name, formed:true, breakout:breakout2,
              detail: '軌道上緣≈' + upperAtLast.toFixed(2) + '　現價＝' + lastClose.toFixed(2) + (volConfirm ? '　帶量' : '　量未放大'),
              desc:'股價沿上升軌道緩步上漲，MA20上揚下帶量長紅收盤突破軌道上緣為買點',
              line: { i1: l1, p1: ly1 + offset, slope: slope2 },
              line2: { i1: l1, p1: ly1, slope: slope2 } });
            return;
          }
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'股價沿上升軌道緩步上漲，MA20上揚下帶量長紅收盤突破軌道上緣為買點' });
    })();

    // ── (9) 突破飆股大量黑K最高點：急漲後大量黑K回檔，3日內帶量長紅突破其最高點 ──
    (function () {
      var id = 'blackk', name = '突破飆股大量黑K最高點';
      var lookback = 10, confirmWindow = 3;
      var floor3 = Math.max(0, data.length - 1 - lookback);
      var candidates = [];
      for (var i = floor3; i < last; i++) {
        var d = data[i];
        if (d.close < d.open && d.vm20 && d.volume / d.vm20 >= 1.6) candidates.push(i);
      }
      if (candidates.length) {
        var bkIdx = candidates[candidates.length - 1];
        var bkHigh = data[bkIdx].high;
        var withinWindow = (last - bkIdx) >= 0 && (last - bkIdx) <= confirmWindow;
        if (withinWindow) {
          var isRed3 = data[last].close > data[last].open;
          var ma20up3 = maSlopeUp('ma20', 10);
          var breakout3 = lastClose > bkHigh && isRed3 && volConfirm && ma20up3;
          results.push({ id:id, name:name, formed:true, breakout:breakout3,
            detail: '大量黑K高點＝' + bkHigh.toFixed(2) + '　現價＝' + lastClose.toFixed(2) + (volConfirm ? '　帶量' : '　量未放大'),
            desc:'飆股急漲後出現大量黑K回檔，3日內帶量長紅突破其最高點為買點',
            line: { i1: bkIdx, p1: bkHigh, slope: 0 } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'飆股急漲後出現大量黑K回檔，3日內帶量長紅突破其最高點為買點' });
    })();

    // ── (10) K線橫盤的突破：三天以上(含首日)收盤未突破/跌破首日K線高低點，帶量突破首日高點為買點 ──
    (function () {
      var id = 'kbp', name = 'K線橫盤的突破';
      if (data.length >= 4) {
        var anchorIdx = last - 3; // 最小需求：3天(含首日)
        var minOK = true;
        for (var j = anchorIdx + 1; j < last; j++) {
          if (data[j].close > data[anchorIdx].high || data[j].close < data[anchorIdx].low) { minOK = false; break; }
        }
        if (minOK) {
          // 課本圖上的橫盤區間常常不只3天——只要收盤價持續守在「首日K線」的高低範圍內，
          // 就往前延伸找到最早、仍然成立的起點，涵蓋較長的橫盤整理段。
          var maxLookback = 20;
          var candidate = anchorIdx - 1;
          while (candidate >= 0 && (last - candidate) <= maxLookback) {
            var inRange = true;
            for (var k = candidate + 1; k < last; k++) {
              if (data[k].close > data[candidate].high || data[k].close < data[candidate].low) { inRange = false; break; }
            }
            if (!inRange) break;
            anchorIdx = candidate;
            candidate--;
          }
          var anchor = data[anchorIdx];
          var isRedK = data[last].close > data[last].open;
          var breakoutK = lastClose > anchor.high && isRedK && volConfirm;
          results.push({ id:id, name:name, formed:true, breakout:breakoutK,
            detail: '首日K線高點＝' + anchor.high.toFixed(2) + '　整理天數＝' + (last - anchorIdx) + '天　現價＝' + lastClose.toFixed(2) + (volConfirm ? '　(帶量)' : '　(量未放大)'),
            desc:'三天以上(含首日)收盤未突破首日K線高低點，帶量紅K突破首日高點為買點',
            line: { i1: anchorIdx, p1: anchor.high, slope: 0 } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'三天以上(含首日)收盤未突破首日K線高低點，帶量紅K突破首日高點為買點' });
    })();

    // ── (11) 高檔母子懷抱：中長紅K(母)＋隔日不過高不破低的黑K/變盤線(子)，次日確認轉折向下 ──
    // 依課程講義：屬於2根K線構成，上漲高檔出現中長紅，次日出現不過高也不破低的黑K線，
    // 中長紅K線稱為母線，次日K線稱為子線；代表多空力量突然拉鋸，多頭上漲力道減弱。
    (function () {
      var id = 'harami_bear', name = '母子懷抱(高檔)';
      if (data.length >= 3) {
        var mother = data[last - 2], child = data[last - 1], confirmDay = data[last];
        var motherBodyPct = Math.abs(mother.close - mother.open) / mother.close;
        var motherIsRed = mother.close > mother.open;
        var motherIsMedLong = motherBodyPct >= 0.035;
        var childContained = child.high <= mother.high && child.low >= mother.low;
        var childSmaller = Math.abs(child.close - child.open) < Math.abs(mother.close - mother.open) * 0.6;
        var atHigh = data[last - 3] ? mother.close >= data[last - 3].close : true;
        if (motherIsRed && motherIsMedLong && childContained && childSmaller && atHigh) {
          var breakoutH = confirmDay.close < child.close;
          results.push({ id:id, name:name, formed:true, breakout:breakoutH,
            detail: '母K(中長紅)高點＝' + mother.high.toFixed(2) + '　子K收於母K範圍內　' + (breakoutH ? '次日已確認轉折向下' : '等待次日確認轉折向下'),
            desc:'上漲高檔出現中長紅K，次日不過高不破低的黑K線為母子懷抱。回測顯示多為高檔強勢整理後續漲，不是反轉',
            marker: { idx: last - 1, price: mother.high, dir: 'up', label: '母子懷抱' } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'上漲高檔出現中長紅K，次日不過高不破低的黑K線為母子懷抱。回測顯示多為高檔強勢整理後續漲，不是反轉' });
    })();

    // ── (12) 低檔母子懷抱：中長黑K(母)＋隔日不過高不破低的紅K/變盤線(子)，次日確認轉折向上 ──
    (function () {
      var id = 'harami_bull', name = '母子懷抱(低檔)';
      if (data.length >= 3) {
        var mother = data[last - 2], child = data[last - 1], confirmDay = data[last];
        var motherBodyPct = Math.abs(mother.close - mother.open) / mother.close;
        var motherIsBlack = mother.close < mother.open;
        var motherIsMedLong = motherBodyPct >= 0.035;
        var childContained = child.high <= mother.high && child.low >= mother.low;
        var childSmaller = Math.abs(child.close - child.open) < Math.abs(mother.close - mother.open) * 0.6;
        var atLow = data[last - 3] ? mother.close <= data[last - 3].close : true;
        if (motherIsBlack && motherIsMedLong && childContained && childSmaller && atLow) {
          var breakoutH2 = confirmDay.close > child.close;
          results.push({ id:id, name:name, formed:true, breakout:breakoutH2,
            detail: '母K(中長黑)低點＝' + mother.low.toFixed(2) + '　子K收於母K範圍內　' + (breakoutH2 ? '次日已確認轉折向上' : '等待次日確認轉折向上'),
            desc:'下跌低檔出現中長黑K，次日不過高不破低的紅K線為母子懷抱，空頭下跌力道轉弱',
            marker: { idx: last - 1, price: mother.low, dir: 'down', label: '母子懷抱' } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'下跌低檔出現中長黑K，次日不過高不破低的紅K線為母子懷抱，空頭下跌力道轉弱' });
    })();

    // ── (13) 晨星：下跌中長黑K＋變盤線＋中長紅K，收盤站上首日黑K實體中點為低檔轉折向上 ──
    // 依課程講義：下跌低檔出現左邊中長黑K，右邊中長紅K，中間夾一根「變盤線」，是低檔轉折向上確認的K線組合。
    (function () {
      var id = 'morning_star', name = '晨星';
      if (data.length >= 3) {
        var s1 = data[last - 2], s2 = data[last - 1], s3 = data[last];
        var s1BodyPct = Math.abs(s1.close - s1.open) / s1.close;
        var s1IsBlack = s1.close < s1.open;
        var s1MedLong = s1BodyPct >= 0.035;
        var s2BodyPct = Math.abs(s2.close - s2.open) / s2.close;
        var isStar1 = s2BodyPct < 0.035;
        var s3BodyPct = Math.abs(s3.close - s3.open) / s3.close;
        var s3IsRed = s3.close > s3.open;
        var s3MedLong = s3BodyPct >= 0.035;
        var s1Mid = (s1.open + s1.close) / 2;
        var closesAboveMid = s3.close > s1Mid;
        if (s1IsBlack && s1MedLong && isStar1 && s3IsRed && s3MedLong && closesAboveMid) {
          results.push({ id:id, name:name, formed:true, breakout:true,
            detail: '首日黑K實體中點＝' + s1Mid.toFixed(2) + '　收盤＝' + s3.close.toFixed(2) + '（已站上中點，轉折確認）',
            desc:'下跌出現中長黑K+變盤線+中長紅K，收盤站上首日實體中點為低檔轉折向上訊號',
            marker: { idx: last - 1, price: s2.low, dir: 'down', label: '晨星' } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'下跌出現中長黑K+變盤線+中長紅K，收盤站上首日實體中點為低檔轉折向上訊號' });
    })();

    // ── (14) 夜星：上漲中長紅K＋變盤線＋中長黑K，收盤跌破首日紅K實體中點為高檔轉折向下 ──
    // 晨星的鏡像型態（課程講義未獨立列出，依同一套邏輯對稱推導）
    (function () {
      var id = 'evening_star', name = '夜星';
      if (data.length >= 3) {
        var e1 = data[last - 2], e2 = data[last - 1], e3 = data[last];
        var e1BodyPct = Math.abs(e1.close - e1.open) / e1.close;
        var e1IsRed = e1.close > e1.open;
        var e1MedLong = e1BodyPct >= 0.035;
        var e2BodyPct = Math.abs(e2.close - e2.open) / e2.close;
        var isStar2 = e2BodyPct < 0.035;
        var e3BodyPct = Math.abs(e3.close - e3.open) / e3.close;
        var e3IsBlack = e3.close < e3.open;
        var e3MedLong = e3BodyPct >= 0.035;
        var e1Mid = (e1.open + e1.close) / 2;
        var closesBelowMid = e3.close < e1Mid;
        // 修正：夜星必須出現在高檔——首日紅K收盤位於前20日(含首日)收盤區間的70%以上
        var esWin = data.slice(Math.max(0, last - 2 - 19), last - 1).map(function(d){ return d.close; });
        var esMax = Math.max.apply(null, esWin), esMin = Math.min.apply(null, esWin);
        var esPos = esMax > esMin ? (e1.close - esMin) / (esMax - esMin) : 0;
        var esAtHigh = esWin.length >= 10 && esPos >= 0.70;
        if (esAtHigh && e1IsRed && e1MedLong && isStar2 && e3IsBlack && e3MedLong && closesBelowMid) {
          results.push({ id:id, name:name, formed:true, breakout:true,
            detail: '首日紅K位於前20日區間' + (esPos * 100).toFixed(0) + '%　實體中點＝' + e1Mid.toFixed(2) + '　收盤＝' + e3.close.toFixed(2) + '（已跌破中點，轉折確認）',
            desc:'高檔(前20日區間70%以上)出現中長紅K+變盤線+中長黑K。回測顯示強勢股出現此型態多為洗盤後續漲（高檔強勢整理），不是反轉',
            marker: { idx: last - 1, price: e2.high, dir: 'up', label: '夜星' } });
          return;
        }
      }
      results.push({ id:id, name:name, formed:false, breakout:false, detail:'尚未偵測到符合結構',
        desc:'高檔(前20日區間70%以上)出現中長紅K+變盤線+中長黑K。回測顯示強勢股出現此型態多為洗盤後續漲（高檔強勢整理），不是反轉' });
    })();

    // ── (15) 回後買上漲：趨勢多頭、回檔量縮價穩，今日紅K放量突破前高 ──
    // 直接沿用 checkPullbackBuy() 的 8 條件判斷結果，統一併入型態確認清單
    if (pb) {
      var pbFormed = pb.allPass || pb.requiredPassed >= pb.requiredTotal - 1;
      results.push({
        id: 'pbup', name: '回後買上漲',
        formed: pbFormed, breakout: pb.allPass,
        detail: '必要條件 ' + pb.requiredPassed + '/' + pb.requiredTotal + ' 通過' + (pb.bonusPassed ? '　+成交量增加分' : ''),
        desc: '趨勢多頭，回檔量縮價穩後，今日紅K放量突破前高為買點'
      });
    }

    // BEARISH_PATTERN_IDS 目前為空：所有型態都計入「型態成形中／突破確認／剛形成」彙總旗標
    var isBullPattern = function(r){ return !BEARISH_PATTERN_IDS[r.id]; };
    var anyBreakout = results.some(function(r){ return isBullPattern(r) && r.breakout; });
    var anyFormed   = results.some(function(r){ return isBullPattern(r) && r.formed; });

    // ── 剛突破：與前一交易日比較，突破訊號是「今天才發生」（昨天用同一套資料尚未突破）──
    var anyJustBroke = false;
    if (!skipJustBroke && data.length > 1) {
      var prevData = data.slice(0, -1);
      var prevPb   = checkPullbackBuy(prevData);
      var prevPt   = detectPatterns(prevData, prevPb, true);
      results.forEach(function(r, i) {
        var pr = prevPt.results[i];
        r.justBroke = !!(r.breakout && (!pr || !pr.breakout));
      });
      anyJustBroke = results.some(function(r){ return isBullPattern(r) && r.justBroke; });
    } else {
      results.forEach(function(r) { r.justBroke = false; });
    }

    return { results:results, anyBreakout:anyBreakout, anyFormed:anyFormed, anyJustBroke:anyJustBroke };
  }

  // ────────────────────────────────────────────────────────────────
  // 歷史回測（回溯過去N個月＋事後N天報酬驗證＋參數/多因子搜尋）
  // 先不用資料庫，結果全部存在記憶體的陣列裡，跑完這次就看這次的結果。
  // ────────────────────────────────────────────────────────────────
  var BACKTEST_HORIZONS = [5, 10, 20];
  // detectPatterns() 裡14種型態的id/name對照（回後買上漲是獨立算的pbAllPass，
  // 不在這14種裡——外面常講的「15種進場型態」是這14種圖形型態+回後買上漲）。
  var PATTERN_DEFS = [
    { id:'hs', name:'頭肩底' }, { id:'chs', name:'複式頭肩底' }, { id:'nb', name:'N字底' },
    { id:'tb', name:'三重底' }, { id:'rb', name:'圓弧底' }, { id:'fb', name:'一字底(均線糾結)' },
    { id:'abc', name:'突破ABC修正下降切線' }, { id:'channel', name:'突破上升軌道線' },
    { id:'blackk', name:'突破飆股大量黑K最高點' }, { id:'kbp', name:'K線橫盤的突破' },
    { id:'harami_bear', name:'母子懷抱(高檔)' }, { id:'harami_bull', name:'母子懷抱(低檔)' },
    { id:'morning_star', name:'晨星' }, { id:'evening_star', name:'夜星' },
  ];
  var backtestResults = []; // 每筆：{evalDate, stockId, name, score, plusDI, minusDI, adx, adxr, bbPos,
                             //        isBreakout, isPullbackRebound, patternFormed, patternBreakout,
                             //        pbAllPass, divTotal, inst3m, entryClose, ret5d, ret10d, ret20d}

  async function runHistoricalBacktestHtml(token, monthsBack, includeDiv, includeInst, minLiquidity, onProgress) {
    backtestResults = [];
    var nameMap = await fetchStockNameMap(token);
    // 相對強弱要用的大盤(0050)資料，整個回測只抓一次，不是每檔股票各抓一次
    var lookbackBuffer = 60;
    var maxHorizon = Math.max.apply(null, BACKTEST_HORIZONS);
    var priceFetchDays = includeDiv
      ? (monthsBack * 30 + lookbackBuffer + maxHorizon + 10 + 400)
      : (monthsBack * 30 + lookbackBuffer + maxHorizon + 10 + 400);   // +400天：52週高點、布林收窄需要約一年歷史
    try {
      currentBenchmark = await fetchBenchmarkSeries(token, priceFetchDays);
    } catch (e) {
      currentBenchmark = null; // 抓不到大盤資料，相對強弱那組欄位/旗標就不會出現，不影響其他分析
    }
    var revenueFetchDays = monthsBack * 30 + 800;   // 營收創24個月新高需要多兩年的月營收
    var instFetchDays = monthsBack * 30 + 100;
    var evalWindowStartDate = new Date();
    evalWindowStartDate.setDate(evalWindowStartDate.getDate() - monthsBack * 30);
    var evalWindowStartStr = evalWindowStartDate.toISOString().slice(0, 10);

    var total = TWSE_LIST.length;
    var saved = 0, failed = 0;

    for (var i = 0; i < total; i++) {
      var sid = TWSE_LIST[i];
      if (onProgress) onProgress(i, total, saved, failed, sid);
      try {
        var rows = await fetchPriceData(sid, token, priceFetchDays);
        var rawData = rows.map(function(d) {
          return { date:d.date, open:+d.open, high:+d.max, low:+d.min, close:+d.close, volume:+d.Trading_Volume };
        }).sort(function(a,b){ return a.date<b.date?-1:1; });
        var data = enrich(rawData);
        var n = data.length;
        var evalStart = lookbackBuffer;
        var evalEnd = n - maxHorizon; // 不含，確保每個評估點後面都還有滿20天可以驗證
        if (evalEnd <= evalStart) { failed++; continue; }

        var revMap = {}, priceMonthAvgMap = {};
        if (includeDiv) {
          try {
            revMap = await fetchRevenueHistoryForBacktest(sid, token, revenueFetchDays);
            priceMonthAvgMap = buildPriceMonthAvgMap(data);
          } catch (e) { revMap = {}; }
        }
        var instDailyMap = {}, instTrustMap = {}, instForeignMap = {};
        if (includeInst) {
          try { var instRes = await fetchInstitutionalHistoryForBacktest(sid, token, instFetchDays); instDailyMap = instRes.all; instTrustMap = instRes.trust; instForeignMap = instRes.foreign; }
          catch (e) { instDailyMap = {}; instTrustMap = {}; instForeignMap = {}; }
        }

        var name = nameMap[sid] || sid;
        for (var idx = evalStart; idx < evalEnd; idx++) {
          var evalDate = data[idx].date;
          if (evalDate < evalWindowStartStr) continue;
          if (minLiquidity > 0) {
            var liqStart = Math.max(0, idx - 19);
            var liqBars = data.slice(liqStart, idx + 1);
            var avgLiquidity = liqBars.reduce(function(sum, b) { return sum + b.close * b.volume; }, 0) / liqBars.length;
            if (avgLiquidity < minLiquidity) continue; // 當時的流動性不夠，跳過這個評估點
          }
          var dataSlice = data.slice(0, idx + 1);
          var techX = computeTechExtras(data, idx);   // 只用到 idx（含）以前的資料 // 只給「當時」以前的資料，不含未來
          var dmi = calcDMI(dataSlice, 14);
          var dm = scoreDMI(dmi);
          var sig = computeCoreSignals(dataSlice, dm);
          var relStrength20 = currentBenchmark ? computeRelativeStrength(dataSlice, currentBenchmark, 20) : null;
          var volRatio = computeVolRatio(dataSlice);
          var volRangePos = computeVolumeRangePosition(dataSlice);
          var volTrend = computeVolumeTrend(dataSlice);
          var volProfile = computeVolumeProfile(dataSlice, VP_PARAMS.lookback, VP_PARAMS.bins);
          var volProfileSig = detectVolumeProfileSignal(dataSlice, volProfile);
          var regime = currentBenchmark ? computeBenchmarkRegime(currentBenchmark, evalDate) : { above20: null, above60: null };

          var pb = checkPullbackBuy(dataSlice);
          var pt = detectPatterns(dataSlice, pb);
          var patternFormed = pt.anyFormed ? 1 : 0;
          var patternBreakout = pt.anyBreakout ? 1 : 0;
          var patternJustBroke = pt.anyJustBroke ? 1 : 0;
          var pbAllPass = pb.allPass ? 1 : 0;
          // 15種型態各自的「剛形成」（justBroke）：跟aggregate的anyJustBroke同一套
          // 判斷邏輯（今天才突破、昨天用同一套資料還沒突破），只是拆成每個型態各自
          // 記錄，而不是只看「有沒有任一種」。用id當key，跟PATTERN_DEFS對應。
          var patternHits = {};
          pt.results.forEach(function(pr) { patternHits[pr.id] = pr.justBroke ? 1 : 0; });

          var divResult = includeDiv ? computeDivergenceAsof(revMap, priceMonthAvgMap, evalDate) : { divTotal: null, priceYoyTotal: null };
          var revF = includeDiv ? computeRevFeats(revMap, evalDate) : { revYoyTurn: null, revYoy3Pos: null, revHigh24: null };
          var inst3m = includeInst ? computeInst3mAsof(instDailyMap, evalDate) : null;
          var flow = includeInst ? computeInstFlowAsof(instDailyMap, instTrustMap, instForeignMap, data, idx) : { trust5: null, foreign5: null, trustStreak3: null };

          var entryClose = data[idx].close;
          var rets = {};
          BACKTEST_HORIZONS.forEach(function(h) {
            rets[h] = (idx + h < n && entryClose) ? (data[idx + h].close - entryClose) / entryClose * 100 : null;
          });

          backtestResults.push({
            evalDate: evalDate, stockId: sid, name: name, score: dm.score,
            plusDI: dm.plusDI, minusDI: dm.minusDI, adx: dm.adx, adxr: dm.adxr, bbPos: sig.bbPos,
            isBreakout: sig.isBreakout ? 1 : 0, isPullbackRebound: sig.isPullbackRebound ? 1 : 0,
            goldenCrossRecent: sig.goldenCrossRecent ? 1 : 0,
            kdjGoldenCrossRecent: sig.kdjGoldenCrossRecent ? 1 : 0, kdjDeathCrossRecent: sig.kdjDeathCrossRecent ? 1 : 0,
            relStrength20: relStrength20, volRatio: volRatio,
            volRangePos: volRangePos, volTrend: volTrend,
            vpBuySupport: volProfileSig.sigType === 'buy_support' ? 1 : 0,
            vpBuyBreakout: volProfileSig.sigType === 'buy_breakout' ? 1 : 0,
            vpSellResistance: volProfileSig.sigType === 'sell_resistance' ? 1 : 0,
            vpSellBreakdown: volProfileSig.sigType === 'sell_breakdown' ? 1 : 0,
            benchmarkAbove20: regime.above20 == null ? null : (regime.above20 ? 1 : 0),
            benchmarkAbove60: regime.above60 == null ? null : (regime.above60 ? 1 : 0),
            relStrengthPositive: relStrength20 != null ? (relStrength20 > 0 ? 1 : 0) : null,
            volSurge15: volRatio != null ? (volRatio >= 1.5 ? 1 : 0) : null,
            volSurge2: volRatio != null ? (volRatio >= 2 ? 1 : 0) : null,
            volLow05: volRatio != null ? (volRatio <= 0.5 ? 1 : 0) : null,
            volRangeHigh90: volRangePos != null ? (volRangePos >= 90 ? 1 : 0) : null,
            volRangeLow10: volRangePos != null ? (volRangePos <= 10 ? 1 : 0) : null,
            confluenceCount: computeConfluenceCount({
              patternJustBroke: patternJustBroke, kdjGoldenCrossRecent: sig.kdjGoldenCrossRecent ? 1 : 0,
              kdjDeathCrossRecent: sig.kdjDeathCrossRecent ? 1 : 0, goldenCrossRecent: sig.goldenCrossRecent ? 1 : 0,
              bbPos: sig.bbPos, volRatio: volRatio,
            }),
            patternFormed: patternFormed, patternBreakout: patternBreakout, patternJustBroke: patternJustBroke,
            patternHits: patternHits,
            pbAllPass: pbAllPass,
            high52Dist: techX.high52Dist, newHigh52: techX.newHigh52, maBull: techX.maBull, bbwRank: techX.bbwRank, gapUp3: techX.gapUp3,
            vol3Ratio: techX.vol3Ratio,   // crossMa60_3／crossMa100_3／gcBelow0_3 已確認無效，不再存進回測紀錄
            udVolRatio20: techX.udVolRatio20, cmf20: techX.cmf20, obvNewHigh60: techX.obvNewHigh60,
            bias20: techX.bias20, pullMa60: techX.pullMa60, limitUp3: techX.limitUp3, drop5: techX.drop5, nr7: techX.nr7,
            divTotal: divResult.divTotal, priceYoy3m: divResult.priceYoyTotal, revYoyTurn: revF.revYoyTurn, revYoy3Pos: revF.revYoy3Pos, revHigh24: revF.revHigh24, inst3m: inst3m, trust5: flow.trust5, foreign5: flow.foreign5, trustStreak3: flow.trustStreak3,
            entryClose: entryClose,
            ret5d: rets[5], ret10d: rets[10], ret20d: rets[20],
          });
          saved++;
        }
      } catch (ex) {
        failed++;
      }
      if (i < total - 1) await new Promise(function(r){ setTimeout(r, 300); });
    }
    if (onProgress) onProgress(total, total, saved, failed, null);
    return { saved: saved, failed: failed, total: total };
  }

  // ── 分析函式：評分區間／標記 vs 實際報酬 ──
  function bucketByScore(score) {
    if (score >= 80) return '80-100（積極做多）';
    if (score >= 65) return '65-79（可考慮進場）';
    if (score >= 50) return '50-64（觀望）';
    return '0-49（不建議）';
  }

  function analyzeScoreBuckets(rows) {
    var order = ['0-49（不建議）', '50-64（觀望）', '65-79（可考慮進場）', '80-100（積極做多）'];
    var groups = {};
    rows.forEach(function(r) {
      var b = bucketByScore(r.score);
      (groups[b] = groups[b] || []).push(r);
    });
    return order.filter(function(b){ return groups[b]; }).map(function(b) {
      return buildStatRow('評分區間', b, groups[b]);
    });
  }

  function analyzeTagHitrate(rows, field, label) {
    var groups = {};
    rows.forEach(function(r) {
      var v = r[field];
      if (v == null) return;
      var key = v === 1 ? (label + '＝是') : (label + '＝否');
      (groups[key] = groups[key] || []).push(r);
    });
    return Object.keys(groups).map(function(k) { return buildStatRow('標記', k, groups[k]); });
  }

  // 15種型態各自「剛形成」單獨列一列（只看「是」的情況，不像analyzeTagHitrate
  // 那樣拆是/否兩列——這裡是要橫向比較15種型態彼此的表現，不是看單一型態內部
  // 有沒有效，表格才不會一次擠出28列）。依樣本數由多到少排序，樣本太少
  // （<10筆）的型態直接跳過，不然單一兩筆資料的100%勝率沒有參考意義。
  function analyzePatternHits(rows) {
    var stats = [];
    PATTERN_DEFS.forEach(function(pdef) {
      var group = rows.filter(function(r) { return r.patternHits && r.patternHits[pdef.id] === 1; });
      if (group.length < 10) return;
      stats.push(buildStatRow('型態', pdef.name + '剛形成', group));
    });
    stats.sort(function(a,b){ return b['樣本數'] - a['樣本數']; });
    return stats;
  }

  function buildStatRow(labelKey, labelVal, group) {
    var row = {};
    row[labelKey] = labelVal;
    row['樣本數'] = group.length;
    BACKTEST_HORIZONS.forEach(function(h) {
      var col = 'ret' + h + 'd';
      var valid = group.map(function(r){ return r[col]; }).filter(function(v){ return v != null; });
      row[h + '日平均報酬%'] = valid.length ? (valid.reduce(function(a,b){return a+b;},0)/valid.length).toFixed(2) : null;
      row[h + '日勝率%'] = valid.length ? ((valid.filter(function(v){return v>0;}).length/valid.length*100).toFixed(1)) : null;
    });
    return row;
  }

  // ── 參數網格搜尋：在已收集的原始+DI/-DI/ADX/ADXR上，換不同權重試算 ──
  function gridSearchParamsHtml(rows, targetHorizon) {
    var retCol = 'ret' + targetHorizon + 'd';
    var valid = rows.filter(function(r) { return r[retCol] != null && r.plusDI != null && r.minusDI != null && r.adx != null; });
    if (!valid.length) return [];
    var results = [];
    [30,40,50].forEach(function(adxCap) {
      [15,20,25].forEach(function(adxrBonus) {
        [0.4,0.5,0.6].forEach(function(diWeight) {
          var adxMaxPts = Math.max(0, 100 - 100*diWeight - adxrBonus);
          var scored = valid.map(function(r) {
            var diSum = r.plusDI + r.minusDI;
            var diDom = diSum > 0 ? (r.plusDI / diSum * 100) : 50;
            var diPts = diDom * diWeight;
            var adxPts = Math.min(r.adx, adxCap) / adxCap * adxMaxPts;
            var hasAdxr = r.adxr != null;
            var adxrPts = (hasAdxr && r.adx > r.adxr) ? adxrBonus : 0;
            var newScore = Math.max(0, Math.min(100, diPts + adxPts + adxrPts));
            return { r: r, newScore: newScore };
          });
          [50,65,80].forEach(function(buyThr) {
            var matched = scored.filter(function(x){ return x.newScore >= buyThr; });
            if (matched.length < 20) return;
            var rets = matched.map(function(x){ return x.r[retCol]; });
            var avgRet = rets.reduce(function(a,b){return a+b;},0) / rets.length;
            var winRate = rets.filter(function(v){return v>0;}).length / rets.length * 100;
            results.push({
              '方向性權重': diWeight, 'ADX滿分上限': adxCap, 'ADXR加分': adxrBonus, '買進門檻': buyThr,
              '訊號數': matched.length,
            });
            results[results.length-1][targetHorizon + '日平均報酬%'] = avgRet.toFixed(2);
            results[results.length-1][targetHorizon + '日勝率%'] = winRate.toFixed(1);
          });
        });
      });
    });
    results.sort(function(a,b){ return parseFloat(b[targetHorizon+'日平均報酬%']) - parseFloat(a[targetHorizon+'日平均報酬%']); });
    return results.slice(0, 15);
  }

  // ── 訊號堆疊數量（Confluence Count）──不是窮舉「哪幾個條件」的組合，而是直接
  // 算「同一個時間點，總共同時觸發了幾個獨立訊號」，測試「訊號數量越多，勝率是否
  // 線性遞增」這個假設。跟多因子複選搜尋是互補的兩種角度：複選搜尋要人工圈定
  // 具體是哪幾個條件、容易有多重比較問題；這個角度不管是哪幾個訊號，只看「密度」，
  // 比較不受單一巧合組合影響。
  //
  // 特意挑5個彼此「概念上不重疊」的訊號類別（型態、KDJ交叉、MACD交叉、布林通道
  // 極端位置、爆量），不用「型態成形中/型態突破確認/型態剛形成」這種同一件事的
  // 不同變體去灌水，也不會被同一種訊號重複計分。
  function computeConfluenceCount(flagRow) {
    var count = 0;
    if (flagRow.patternJustBroke === 1) count++;
    if (flagRow.kdjGoldenCrossRecent === 1 || flagRow.kdjDeathCrossRecent === 1) count++;
    if (flagRow.goldenCrossRecent === 1) count++;
    if (flagRow.bbPos != null && (flagRow.bbPos <= 20 || flagRow.bbPos >= 80)) count++;
    if (flagRow.volRatio != null && flagRow.volRatio >= 1.5) count++;
    return count;
  }

  // 依訊號堆疊數量(0~5)分桶，看勝率是否隨數量遞增——這是驗證假設本身要看的表格，
  // 不是又一組「指定組合」。
  function analyzeConfluenceCount(rows) {
    var groups = {};
    rows.forEach(function(r) {
      if (r.confluenceCount == null) return;
      var key = r.confluenceCount + '個訊號同時觸發';
      (groups[key] = groups[key] || []).push(r);
    });
    var order = [0,1,2,3,4,5].map(function(n){ return n + '個訊號同時觸發'; });
    return order.filter(function(k){ return groups[k]; }).map(function(k) {
      return buildStatRow('訊號堆疊數量', k, groups[k]);
    });
  }

  // ── 多因子複選搜尋：窮舉1~3個條件旗標的AND組合 ──

  // ── 新增技術面欄位（2026-09-30）：52週高點、均線多頭排列、布林收窄、向上跳空 ──
  var LONG_HISTORY_DAYS = 400;   // 52週高點約需250根K棒；批次分析會多抓到400天只用來算這幾個欄位
  function computeTechExtras(data, idx) {
    var out = { high52Dist: null, newHigh52: null, maBull: null, bbwRank: null, gapUp3: null,
                crossMa60_3: null, crossMa100_3: null, vol3Ratio: null, gcBelow0_3: null };
    var bar = data[idx];
    if (!bar) return out;
    if (idx >= 249) {
      var mx = -Infinity, prev = -Infinity;
      for (var i = idx - 249; i <= idx; i++) {
        if (data[i].high > mx) mx = data[i].high;
        if (i < idx && data[i].high > prev) prev = data[i].high;
      }
      out.high52Dist = mx > 0 ? (bar.close / mx - 1) * 100 : null;
      out.newHigh52 = bar.high > prev ? 1 : 0;
    }
    if (bar.ma5 != null && bar.ma20 != null && bar.ma60 != null) {
      out.maBull = (bar.ma5 > bar.ma20 && bar.ma20 > bar.ma60 && bar.close > bar.ma20) ? 1 : 0;
    }
    var bw = function(b) { return (b.bbU != null && b.bbL != null && (b.bbU + b.bbL) > 0) ? (b.bbU - b.bbL) / ((b.bbU + b.bbL) / 2) : null; };
    var cur = bw(bar);
    if (cur != null && idx >= 119) {
      var cnt = 0, le = 0;
      for (var j = idx - 119; j <= idx; j++) { var w = bw(data[j]); if (w != null) { cnt++; if (w <= cur) le++; } }
      if (cnt >= 100) out.bbwRank = le / cnt * 100;
    }
    if (idx >= 3) {
      var g = 0;
      for (var k = 0; k < 3; k++) { if (data[idx - k].low > data[idx - k - 1].high) g = 1; }
      out.gapUp3 = g;
    }
    // 3日內收盤由下往上突破季線(MA60)／半年線(MA100)，且今天仍站在線上（2026-10-03 參考 BBW+MACD 掃描器新增）
    var smaAt = function(p, i) {
      if (i < p - 1) return null;
      var s = 0; for (var q = i - p + 1; q <= i; q++) s += data[q].close;
      return s / p;
    };
    [[60, 'crossMa60_3'], [100, 'crossMa100_3']].forEach(function(pp) {
      var p = pp[0], ma = {};
      for (var j = idx - 3; j <= idx; j++) ma[j] = j >= 0 ? smaAt(p, j) : null;
      if (idx >= 3 && ma[idx - 3] != null && ma[idx] != null) {
        var cr = 0;
        for (var j2 = idx - 2; j2 <= idx; j2++) { if (data[j2 - 1].close < ma[j2 - 1] && data[j2].close >= ma[j2]) cr = 1; }
        out[pp[1]] = (cr && bar.close >= ma[idx]) ? 1 : 0;
      }
    });
    // 連續放量：近3日均量 ÷ 再往前20日均量
    if (idx >= 22) {
      var base = 0; for (var v = idx - 22; v < idx - 2; v++) base += data[v].volume;
      base /= 20;
      out.vol3Ratio = base > 0 ? ((data[idx - 2].volume + data[idx - 1].volume + bar.volume) / 3) / base : null;
    }
    // MACD 零軸下（DIF<0）3日內黃金交叉＝低檔轉折
    if (idx >= 3) {
      var gb = 0;
      for (var m = idx - 2; m <= idx; m++) {
        var a0 = data[m - 1], a1 = data[m];
        if (a0.macd != null && a0.macdSig != null && a1.macd != null && a1.macdSig != null &&
            a0.macd <= a0.macdSig && a1.macd > a1.macdSig && a1.macd < 0) gb = 1;
      }
      out.gcBelow0_3 = gb;
    }
    // 2026-10-04 買賣力道近似（日線沒有內外盤，用收盤在當日區間的位置／漲跌日量推估）
    out.udVolRatio20 = null; out.cmf20 = null; out.obvNewHigh60 = null;
    if (idx >= 20) {
      var upV = 0, dnV = 0, mfv = 0, vsum = 0;
      for (var q2 = idx - 19; q2 <= idx; q2++) {
        var bq = data[q2], bp = data[q2 - 1], vq = bq.volume;
        if (bq.close > bp.close) upV += vq; else if (bq.close < bp.close) dnV += vq;
        var rg = bq.high - bq.low;
        if (rg > 0) mfv += ((bq.close - bq.low) - (bq.high - bq.close)) / rg * vq;
        vsum += vq;
      }
      out.udVolRatio20 = dnV > 0 ? Math.min(upV / dnV, 99) : (upV > 0 ? 99 : null);
      out.cmf20 = vsum > 0 ? mfv / vsum : null;
    }
    if (idx >= 61) {
      // OBV 從 idx-60 起算（起點＝0），今天的 OBV 高於前60天每一天＝OBV創60日新高
      var obvv = 0, obMax = 0;
      var stepObv = function(j) { return data[j].close > data[j - 1].close ? data[j].volume : (data[j].close < data[j - 1].close ? -data[j].volume : 0); };
      for (var q3 = idx - 59; q3 < idx; q3++) { obvv += stepObv(q3); obMax = Math.max(obMax, obvv); }
      obvv += stepObv(idx);
      out.obvNewHigh60 = obvv > obMax ? 1 : 0;
    }
    // 2026-10-05 新增：短線乖離、多頭回測季線、近3日漲停（台股）、近5日跌幅與NR7窄幅日（美股）
    out.bias20 = null; out.pullMa60 = null; out.limitUp3 = null; out.drop5 = null; out.nr7 = null;
    if (bar.ma20) out.bias20 = (bar.close / bar.ma20 - 1) * 100;
    if (idx >= 10 && bar.ma60 && data[idx - 10].ma60) out.pullMa60 = (bar.ma60 > data[idx - 10].ma60 && Math.abs(bar.close / bar.ma60 - 1) <= 0.03) ? 1 : 0;
    if (idx >= 3) {
      var lu = 0;
      for (var q4 = idx - 2; q4 <= idx; q4++) { if (data[q4 - 1].close > 0 && data[q4].close / data[q4 - 1].close - 1 >= 0.095) lu = 1; }
      out.limitUp3 = lu;
    }
    if (idx >= 5 && data[idx - 5].close) out.drop5 = (bar.close / data[idx - 5].close - 1) * 100;
    if (idx >= 6) {
      var rmin = Infinity;
      for (var q5 = idx - 6; q5 < idx; q5++) rmin = Math.min(rmin, data[q5].high - data[q5].low);
      out.nr7 = (bar.high - bar.low) < rmin ? 1 : 0;
    }
    return out;
  }
  // 新增的技術面條件旗標
  var NEW_TECH_FLAGS = [
    ['newHigh52', '創52週新高', function(v) { return v === 1; }],
    ['high52Dist', '距52週高點≤5%', function(v) { return v >= -5; }],
    ['maBull', '均線多頭排列(5>20>60且站上月線)', function(v) { return v === 1; }],
    // 布林通道收窄：2026-09-30 台股三年回測報酬低於基準，已移除（bbwRank 欄位仍保留）
    ['gapUp3', '近3日向上跳空缺口', function(v) { return v === 1; }],
    // 2026-10-03 參考 BBW+MACD 掃描器新增；2026-10-04 三年三段回測：突破季線／半年線、MACD零軸下金叉無效或反向，
    // 已移除旗標（crossMa60_3／crossMa100_3／gcBelow0_3 欄位仍保留在回測紀錄裡）
    ['vol3Ratio', '連續放量(近3日均量≥1.5倍前20日均量)', function(v) { return v >= 1.5; }],
    // 2026-10-04 買賣力道近似（待回測驗證）
    ['udVolRatio20', '漲時量≥跌時量1.5倍(近20日)', function(v) { return v >= 1.5; }],
    ['cmf20', 'CMF資金流買方佔優(近20日≥0.1)', function(v) { return v >= 0.1; }],
    ['obvNewHigh60', 'OBV能量潮創60日新高', function(v) { return v === 1; }],
    // 2026-10-05 新增（待回測驗證）
    ['pullMa60', '多頭回測季線(季線上揚、距季線±3%)', function(v) { return v === 1; }],
    ['bias20', '短線過熱(高於月線≥20%)', function(v) { return v >= 20; }],
    ['limitUp3', '近3日漲停', function(v) { return v === 1; }]
  ];
  // 不放進「多因子複選搜尋／飆股搜尋」的條件（回測貢獻極低或與個別型態重複；指定組合比對與統計仍可使用）
  var SEARCH_EXCLUDE_FLAGS = { '型態成形中': 1, '型態突破確認': 1, '型態剛形成(剛突破)': 1, '突破上升軌道線剛形成': 1,
    '分價量表-守穩POC買進': 1, '分價量表-突破POC追價買進': 1, '分價量表-反彈POC遇壓賣出': 1, '分價量表-破位停損賣出': 1 };
  // 2026-10-05 台股三年三段回測稽核：下列條件單獨無效（或只在單一行情有效），且幾乎不出現在三段都穩定的選股／飆股組合，
  // 不再放進搜尋以減少雜訊與運算量（指定組合比對、追蹤統計、單一條件表仍保留）
  ['跌深反彈盤', '頭肩底剛形成', '複式頭肩底剛形成', '一字底(均線糾結)剛形成', '三重底剛形成', '母子懷抱(高檔)剛形成', '布林通道低檔(≤20%)', '投信近5日買超', '投信連續買超≥3日', '土洋同買(外資、投信近5日皆買超)'].forEach(function(k) { SEARCH_EXCLUDE_FLAGS[k] = 1; });
  // 2026-10-05 台股三年回測：多頭回測季線、月營收YoY由負轉正 三年都明顯負報酬，不再放進搜尋（單一條件表仍顯示，可當警示）
  ['多頭回測季線(季線上揚、距季線±3%)', '月營收YoY由負轉正'].forEach(function(k) { SEARCH_EXCLUDE_FLAGS[k] = 1; });
  function searchFlagNames(flagDefs) { return Object.keys(flagDefs).filter(function(n) { return !SEARCH_EXCLUDE_FLAGS[n]; }); }
  // 前後半段驗證：依評估日切兩半（同一天不會被切開）
  function halfSplitDate(valid) {
    var ds = {}; valid.forEach(function(r) { ds[r.evalDate] = 1; });
    var u = Object.keys(ds).sort();
    return u.length >= 4 ? u[Math.floor(u.length / 2)] : null;
  }
  function tOfEx(exs) {
    var n = exs.length; if (n < 5) return null;
    var m = exs.reduce(function(a, b) { return a + b; }, 0) / n;
    var sd = Math.sqrt(exs.reduce(function(s, v) { return s + (v - m) * (v - m); }, 0) / (n - 1));
    return sd > 0 ? m / (sd / Math.sqrt(n)) : null;
  }
  function analyzeNewTech(rows) {
    var flagDefs = buildConditionFlags(rows), out = [];
    NEW_TECH_FLAGS.forEach(function(f) {
      var fld = f[0], lbl = f[1];
      if (!flagDefs[lbl]) return;
      var ok = rows.filter(function(r) { return r[fld] != null; });
      var yes = ok.filter(function(r) { return flagDefs[lbl](r); });
      var no = ok.filter(function(r) { return !flagDefs[lbl](r); });
      if (yes.length) out.push(buildStatRow('條件', lbl + '＝是', yes));
      if (no.length) out.push(buildStatRow('條件', lbl + '＝否', no));
    });
    return out;
  }

  function buildConditionFlags(rows) {
    var flagDefs = {
      '多方力道≥65': function(r){ return r.score >= 65; },
      '多方力道≥80': function(r){ return r.score >= 80; },
      '強勢突破盤': function(r){ return r.isBreakout === 1; },
      '跌深反彈盤': function(r){ return r.isPullbackRebound === 1; },
      '布林通道低檔(≤20%)': function(r){ return r.bbPos != null && r.bbPos <= 20; },
      '布林通道高檔(≥80%)': function(r){ return r.bbPos != null && r.bbPos >= 80; },
      'MACD近3日內黃金交叉': function(r){ return r.goldenCrossRecent === 1; },
    };
    if (rows.some(function(r){ return r.kdjGoldenCrossRecent != null; })) {
      flagDefs['KDJ近3日內黃金交叉'] = function(r){ return r.kdjGoldenCrossRecent === 1; };
    }
    if (rows.some(function(r){ return r.kdjDeathCrossRecent != null; })) {
      flagDefs['KDJ近3日內死亡交叉'] = function(r){ return r.kdjDeathCrossRecent === 1; };
    }
    if (rows.some(function(r){ return r.relStrength20 != null; })) {
      flagDefs['相對強弱為正(強於大盤)'] = function(r){ return r.relStrength20 != null && r.relStrength20 > 0; };
      flagDefs['相對強弱為負(弱於大盤)'] = function(r){ return r.relStrength20 != null && r.relStrength20 < 0; };
      flagDefs['近20日強於大盤≥10%'] = function(r){ return r.relStrength20 != null && r.relStrength20 >= 10; };
    }
    if (rows.some(function(r){ return r.volRatio != null; })) {
      flagDefs['爆量(≥1.5倍均量)'] = function(r){ return r.volRatio != null && r.volRatio >= 1.5; };
      flagDefs['爆量(≥2倍均量)'] = function(r){ return r.volRatio != null && r.volRatio >= 2; };
      flagDefs['地量(≤0.5倍均量)'] = function(r){ return r.volRatio != null && r.volRatio <= 0.5; };
    }
    if (rows.some(function(r){ return r.volRangePos != null; })) {
      flagDefs['量能區間高檔(≥90百分位)'] = function(r){ return r.volRangePos != null && r.volRangePos >= 90; };
      flagDefs['量能區間低檔(≤10百分位)'] = function(r){ return r.volRangePos != null && r.volRangePos <= 10; };
    }
    if (rows.some(function(r){ return r.volTrend != null; })) {
      flagDefs['量能斜率轉強(近5日均量>近10日均量20%以上)'] = function(r){ return r.volTrend != null && r.volTrend >= 20; };
      flagDefs['量能斜率轉弱(近5日均量<近10日均量20%以上)'] = function(r){ return r.volTrend != null && r.volTrend <= -20; };
    }
    if (rows.some(function(r){ return r.vpBuySupport != null; })) {
      flagDefs['分價量表-守穩POC買進'] = function(r){ return r.vpBuySupport === 1; };
      flagDefs['分價量表-突破POC追價買進'] = function(r){ return r.vpBuyBreakout === 1; };
      flagDefs['分價量表-反彈POC遇壓賣出'] = function(r){ return r.vpSellResistance === 1; };
      flagDefs['分價量表-破位停損賣出'] = function(r){ return r.vpSellBreakdown === 1; };
    }
    if (rows.some(function(r){ return r.benchmarkAbove20 != null; })) {
      flagDefs['大盤站上20日均線'] = function(r){ return r.benchmarkAbove20 === 1; };
      flagDefs['大盤跌破20日均線'] = function(r){ return r.benchmarkAbove20 === 0; };
    }
    if (rows.some(function(r){ return r.benchmarkAbove60 != null; })) {
      flagDefs['大盤站上60日均線'] = function(r){ return r.benchmarkAbove60 === 1; };
      flagDefs['大盤跌破60日均線'] = function(r){ return r.benchmarkAbove60 === 0; };
    }
    if (rows.some(function(r){ return r.patternFormed != null; })) {
      flagDefs['型態成形中'] = function(r){ return r.patternFormed === 1; };
    }
    if (rows.some(function(r){ return r.patternBreakout != null; })) {
      flagDefs['型態突破確認'] = function(r){ return r.patternBreakout === 1; };
    }
    if (rows.some(function(r){ return r.patternJustBroke != null; })) {
      flagDefs['型態剛形成(剛突破)'] = function(r){ return r.patternJustBroke === 1; };
    }
    if (rows.some(function(r){ return r.pbAllPass != null; })) {
      flagDefs['回後買上漲全通過'] = function(r){ return r.pbAllPass === 1; };
    }
    NEW_TECH_FLAGS.forEach(function(f) {
      var fld = f[0], fn = f[2];
      if (rows.some(function(r){ return r[fld] != null; })) {
        flagDefs[f[1]] = function(r) { return r[fld] != null && fn(r[fld]); };
      }
    });
    if (rows.some(function(r){ return r.divTotal != null; })) {
      flagDefs['近3月乖離度為正(營收優於股價)'] = function(r){ return r.divTotal != null && r.divTotal > 0; };
      flagDefs['近3月乖離度為負(股價超前營收)'] = function(r){ return r.divTotal != null && r.divTotal < 0; };
    }
    if (rows.some(function(r){ return r.priceYoy3m != null; })) {
      flagDefs['近3月均價YoY為正'] = function(r){ return r.priceYoy3m != null && r.priceYoy3m > 0; };
      flagDefs['近3月均價YoY為負'] = function(r){ return r.priceYoy3m != null && r.priceYoy3m < 0; };
    }
    if (rows.some(function(r){ return r.inst3m != null; })) {
      flagDefs['三大法人近3月買超'] = function(r){ return r.inst3m != null && r.inst3m > 0; };
      flagDefs['三大法人近3月賣超'] = function(r){ return r.inst3m != null && r.inst3m < 0; };
    }
    // 2026-10-05 新增：月營收動能（待回測驗證）
    if (rows.some(function(r){ return r.revYoyTurn != null; })) flagDefs['月營收YoY由負轉正'] = function(r){ return r.revYoyTurn === 1; };
    if (rows.some(function(r){ return r.revYoy3Pos != null; })) flagDefs['月營收連3月年增'] = function(r){ return r.revYoy3Pos === 1; };
    if (rows.some(function(r){ return r.revHigh24 != null; })) flagDefs['月營收創24個月新高'] = function(r){ return r.revHigh24 === 1; };
    if (rows.some(function(r){ return r.trust5 != null; })) {
      flagDefs['投信近5日買超'] = function(r){ return r.trust5 != null && r.trust5 > 0; };
    }
    if (rows.some(function(r){ return r.trustStreak3 != null; })) {
      flagDefs['投信連續買超≥3日'] = function(r){ return r.trustStreak3 === 1; };
    }
    if (rows.some(function(r){ return r.foreign5 != null; })) {
      flagDefs['外資近5日買超'] = function(r){ return r.foreign5 != null && r.foreign5 > 0; };
      if (rows.some(function(r){ return r.trust5 != null; })) {
        flagDefs['土洋同買(外資、投信近5日皆買超)'] = function(r){ return r.foreign5 != null && r.trust5 != null && r.foreign5 > 0 && r.trust5 > 0; };
      }
    }
    if (rows.some(function(r){ return r.adx != null && r.adxr != null && r.plusDI != null; })) {
      // 2026-10-03 台股三年回測：飆股率6.0%（基準3.6%）、t(依日)6.1，但勝率略低於基準
      flagDefs['DMI趨勢增強(+DI>-DI、ADX≥25且>ADXR)'] = function(r){
        return r.plusDI != null && r.minusDI != null && r.adx != null && r.adxr != null && r.plusDI > r.minusDI && r.adx >= 25 && r.adx > r.adxr;
      };
    }
    // 15種進場型態裡的14種圖形型態，各自的「剛形成」單獨當一個條件旗標
    // （回後買上漲已經是獨立的「回後買上漲全通過」旗標，這裡不重複）。
    if (rows.some(function(r){ return r.patternHits != null; })) {
      PATTERN_DEFS.forEach(function(pdef) {
        flagDefs[pdef.name + '剛形成'] = function(r) {
          return r.patternHits != null && r.patternHits[pdef.id] === 1;
        };
      });
    }
    return flagDefs;
  }

  function combinations(arr, size) {
    if (size === 0) return [[]];
    if (arr.length < size) return [];
    var [first, ...rest] = arr;
    var withFirst = combinations(rest, size - 1).map(function(c){ return [first].concat(c); });
    var withoutFirst = combinations(rest, size);
    return withFirst.concat(withoutFirst);
  }

  // 共用：漸進式窮舉1~3個條件的AND組合。先篩出符合條件i的列，再在裡面篩條件j、條件k，
  // 比每個組合都從全部資料重新篩一次快很多。skipRedundant=true 時，多加一個條件後樣本數
  // 完全沒變的組合（例如「母子懷抱剛形成」必然也是「型態剛形成」）不列出。
  function enumerateCombos(valid, flagDefs, flagNames, maxComboSize, minSamples, skipRedundant, visit) {
    var k = flagNames.length, tested = 0;
    var fn = flagNames.map(function(n){ return flagDefs[n]; });
    var single = [], rowsI = [];
    for (var i = 0; i < k; i++) {
      rowsI[i] = valid.filter(function(r){ return fn[i](r); });
      single[i] = rowsI[i].length;
    }
    var pairCnt = {};
    var bySize = { 1: [], 2: [], 3: [] };
    for (var a = 0; a < k; a++) {
      tested++;
      if (single[a] >= minSamples) { var v1 = visit([a], rowsI[a]); if (v1) bySize[1].push(v1); }
    }
    if (maxComboSize >= 2) {
      for (var i2 = 0; i2 < k; i2++) {
        for (var j2 = i2 + 1; j2 < k; j2++) {
          tested++;
          var rij = rowsI[i2].filter(function(r){ return fn[j2](r); });
          pairCnt[i2 + '_' + j2] = rij.length;
          if (rij.length < minSamples) { if (maxComboSize >= 3) tested += k - j2 - 1; continue; }
          if (!(skipRedundant && (rij.length === single[i2] || rij.length === single[j2]))) {
            var v2 = visit([i2, j2], rij); if (v2) bySize[2].push(v2);
          }
          if (maxComboSize >= 3) {
            for (var l = j2 + 1; l < k; l++) {
              tested++;
              var rijl = rij.filter(function(r){ return fn[l](r); });
              if (rijl.length < minSamples) continue;
              bySize[3].push({ i: i2, j: j2, l: l, rows: rijl });
            }
          }
        }
      }
      // 三因子的冗餘判斷要用到所有兩兩組合的樣本數，所以最後才處理
      var tri = bySize[3]; bySize[3] = [];
      tri.forEach(function(t) {
        var n = t.rows.length;
        if (skipRedundant) {
          var pil = pairCnt[t.i + '_' + t.l], pjl = pairCnt[t.j + '_' + t.l];
          if (pil == null) pil = rowsI[t.i].filter(function(r){ return fn[t.l](r); }).length;
          if (pjl == null) pjl = rowsI[t.j].filter(function(r){ return fn[t.l](r); }).length;
          if (n === pairCnt[t.i + '_' + t.j] || n === pil || n === pjl) return;
        }
        var v3 = visit([t.i, t.j, t.l], t.rows); if (v3) bySize[3].push(v3);
      });
    }
    return { results: bySize[1].concat(bySize[2], bySize[3]), combosTested: tested };
  }

  function comboSearchHtml(rows, targetHorizon, maxComboSize, minSamples, minWinRate, skipRedundant, sortBy) {
    var retCol = 'ret' + targetHorizon + 'd';
    var valid = rows.filter(function(r){ return r[retCol] != null; });
    var flagDefs = buildConditionFlags(rows);
    var flagNames = searchFlagNames(flagDefs);
    if (!valid.length || !flagNames.length) return { results: [], combosTested: 0, base: null };
    var midDate = halfSplitDate(valid);
    var allRets = valid.map(function(r){ return r[retCol]; });
    var baseAvg = allRets.reduce(function(a,b){return a+b;},0) / allRets.length;
    var base = { n: allRets.length, avg: baseAvg,
      win: allRets.filter(function(v){return v>0;}).length / allRets.length * 100, median: median(allRets) };
    // 同日超額報酬：每筆報酬扣掉同一評估日全部紀錄的平均報酬（去除大盤齊漲齊跌的共同成分），t值用它算
    var exKey = '_ex' + targetHorizon;
    var daySum = {}, dayCnt = {};
    valid.forEach(function(r) { daySum[r.evalDate] = (daySum[r.evalDate] || 0) + r[retCol]; dayCnt[r.evalDate] = (dayCnt[r.evalDate] || 0) + 1; });
    valid.forEach(function(r) { r[exKey] = r[retCol] - daySum[r.evalDate] / dayCnt[r.evalDate]; });
    var out = enumerateCombos(valid, flagDefs, flagNames, maxComboSize, minSamples, skipRedundant !== false, function(combo, matched) {
      var rets = matched.map(function(r){ return r[retCol]; });
      var n = rets.length;
      var avgRet = rets.reduce(function(a,b){return a+b;},0) / n;
      var winRate = rets.filter(function(v){return v>0;}).length / n * 100;
      if (winRate < minWinRate) return null; // 只保留勝率達門檻的組合
      var exs = matched.map(function(r){ return r[exKey]; });
      var exMean = exs.reduce(function(a,b){return a+b;},0) / n;
      var sd = n > 1 ? Math.sqrt(exs.reduce(function(s,v){ return s + (v-exMean)*(v-exMean); }, 0) / (n-1)) : 0;
      var t = sd > 0 ? exMean / (sd / Math.sqrt(n)) : null;
      var comboNames = combo.map(function(i){ return flagNames[i]; });
      var row = { '條件組合': comboNames.join(' ＋ '), '類型': isTimingCombo(comboNames) ? '擇時型' : '', '條件數': combo.length, '樣本數': n };
      row[targetHorizon+'日平均報酬%'] = avgRet.toFixed(2);
      row[targetHorizon+'日勝率%'] = winRate.toFixed(1);
      row['同日超額報酬%'] = exMean.toFixed(2);
      row['中位數報酬%'] = median(rets).toFixed(2);
      row['t值(同日調整)'] = t != null ? t.toFixed(2) : null;
      if (midDate) {   // 前後半段驗證：兩段 t 都 ≥2 才比較像真訊號
        var t1 = tOfEx(matched.filter(function(r){ return r.evalDate < midDate; }).map(function(r){ return r[exKey]; }));
        var t2 = tOfEx(matched.filter(function(r){ return r.evalDate >= midDate; }).map(function(r){ return r[exKey]; }));
        var ok1 = t1 != null && t1 >= 2, ok2 = t2 != null && t2 >= 2;
        row['t值(前半)'] = t1 != null ? t1.toFixed(2) : null;
        row['t值(後半)'] = t2 != null ? t2.toFixed(2) : null;
        row['前後一致'] = (ok1 && ok2) ? '✅' : ((ok1 || ok2) ? '⚠️' : '');
      }
      return row;
    });
    if (sortBy === 'win') {
      out.results.sort(function(a,b){ return parseFloat(b[targetHorizon+'日勝率%']) - parseFloat(a[targetHorizon+'日勝率%']); });
    } else { // 預設：依同日調整t值排序，優先列出「扣掉同一天大盤後仍然比較強」的選股型組合
      var tv = function(r){ var v = parseFloat(r['t值(同日調整)']); return isNaN(v) ? -1e9 : v; };
      out.results.sort(function(a,b){ return tv(b) - tv(a); });
    }
    base.mid = midDate;
    out.base = base;
    return out;
  }

  function median(arr) {
    if (!arr.length) return 0;
    var s = arr.slice().sort(function(a,b){ return a-b; });
    var m = Math.floor(s.length / 2);
    return s.length % 2 ? s[m] : (s[m-1] + s[m]) / 2;
  }

  // 飆股搜尋：只看「命中大行情的比例」，並附上全體基準與倍數
  function findMoonshotCombos(rows, targetHorizon, threshold, maxComboSize, minSamples, minMoonshotPct, skipRedundant) {
    var retCol = 'ret' + targetHorizon + 'd';
    var valid = rows.filter(function(r){ return r[retCol] != null; });
    var flagDefs = buildConditionFlags(rows);
    var flagNames = searchFlagNames(flagDefs);
    if (!valid.length || !flagNames.length) return { results: [], combosTested: 0, base: null };
    var midDate = halfSplitDate(valid);
    var allMoon = valid.filter(function(r){ return r[retCol] > threshold; }).length;
    var basePct = allMoon / valid.length * 100;
    var out = enumerateCombos(valid, flagDefs, flagNames, maxComboSize, minSamples, skipRedundant !== false, function(combo, matched) {
      var moonshots = matched.filter(function(r) { return r[retCol] > threshold; });
      if (!moonshots.length) return null;
      var rate = moonshots.length / matched.length * 100;
      if (rate < minMoonshotPct) return null;
      var avgMoon = moonshots.reduce(function(a,r){return a+r[retCol];},0) / moonshots.length;
      var res = {
        '條件組合': combo.map(function(i){ return flagNames[i]; }).join(' ＋ '), '條件數': combo.length, '樣本數': matched.length,
        '飆股次數': moonshots.length, '飆股比例%': rate.toFixed(1), '飆股平均漲幅%': avgMoon.toFixed(1),
        '倍數(vs基準)': basePct > 0 ? (rate / basePct).toFixed(2) : null,
      };
      if (midDate) {
        [['前半', function(r){ return r.evalDate < midDate; }], ['後半', function(r){ return r.evalDate >= midDate; }]].forEach(function(hh) {
          var g = matched.filter(hh[1]);
          res['飆股比例(' + hh[0] + ')%'] = g.length ? (g.filter(function(r){ return r[retCol] > threshold; }).length / g.length * 100).toFixed(1) : null;
        });
      }
      return res;
    });
    out.results.sort(function(a,b){ return parseFloat(b['飆股比例%']) - parseFloat(a['飆股比例%']); });
    out.base = { n: valid.length, moonN: allMoon, pct: basePct, mid: midDate };
    return out;
  }

  // 使用者指定要追蹤的特定條件組合——不管樣本數多少、排名有沒有進前幾名，
  // 都一定會顯示出來，方便針對特定假設做比較。
  // ✅ 2026-09-22 台股3年回測（含相對強弱vs大盤0050、成交量確認兩個新指標）的
  // 完整結果，24組去重後的高勝率組合。核心發現：「布林通道低檔+晨星剛形成」
  // 這個家族在新指標加入後表現進一步驗證，加「爆量≥1.5倍均量」或「近3月均價
  // YoY為負」後20日勝率衝到85-88%（樣本數約90-110筆）。相對強弱這個新指標
  // 也確實有進入排行榜（見#5、#9、#12、#13、#14），但不是最強因子，晨星+布林
  // 低檔仍是核心骨幹，相對強弱比較像加分項。
  // ✅ 2026-09-22 台股3年回測（含相對強弱vs大盤0050、成交量確認）結果，篩選條件：
  // 勝率>60%（5/10/20日任一天數達標即列入）。24組。
  // ✅ 2026-09-22 台股3年回測最新結果（勝率5日檔案跟10日檔案完全重複，判斷是
  // 檔名標錯，這批不含真正的5日勝率資料，只有10日/20日）。移除重複後32組。
  // ✅ 2026-09-22 台股3年回測最新結果，10日/20日資料 + 補上的5日勝率資料合併。
  // 42組（原32組裡有10組補上5日數字，另外10組是5日資料裡才出現的全新組合）。
  var PINNED_COMBOS = [
    ['布林通道低檔(≤20%)', '爆量(≥1.5倍均量)', '晨星剛形成'],  // 5日83.1%，10日86.4%，20日86.4%
    ['布林通道低檔(≤20%)', '近3月均價YoY為負', '晨星剛形成'],  // 5日82.8%，10日85.9%，20日85.9%
    ['布林通道低檔(≤20%)', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],  // 5日81.1%，10日83.5%，20日81.1%
    ['布林通道低檔(≤20%)', 'KDJ近3日內黃金交叉', '晨星剛形成'],  // 5日77.9%，10日82.4%，20日74.4%
    ['布林通道低檔(≤20%)', '三大法人近3月買超', '晨星剛形成'],  // 5日76.5%，10日82.4%，20日77.3%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '晨星剛形成'],  // 5日78.9%，10日82.3%，20日77.0%
    ['布林通道低檔(≤20%)', '相對強弱為負(弱於大盤)', '晨星剛形成'],  // 5日77.7%，10日81.7%，20日77.2%
    ['布林通道低檔(≤20%)', '大盤跌破20日均線', '晨星剛形成'],  // 5日76.3%，10日80.4%，20日74.4%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '晨星剛形成'],  // 5日79.7%，10日80.0%，20日76.4%
    ['布林通道低檔(≤20%)', '晨星剛形成'],  // 5日75.7%，10日79.6%，20日73.5%
    ['布林通道低檔(≤20%)', '型態成形中', '晨星剛形成'],  // 5日75.7%，10日79.6%，20日73.5%
    ['布林通道低檔(≤20%)', '型態突破確認', '晨星剛形成'],  // 5日75.7%，10日79.6%，20日73.5%
    ['布林通道低檔(≤20%)', '型態剛形成(剛突破)', '晨星剛形成'],  // 5日75.7%，10日79.6%，20日73.5%
    ['布林通道低檔(≤20%)', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],  // 5日74.8%，10日78.5%，20日73.8%
    ['布林通道低檔(≤20%)', '三大法人近3月賣超', '晨星剛形成'],  // 5日74.8%，10日76.6%，20日69.2%
    ['布林通道低檔(≤20%)', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],  // 5日69.8%，10日75.0%，20日63.5%
    ['布林通道低檔(≤20%)', '近3月均價YoY為正', '晨星剛形成'],  // 5日70.1%，10日74.8%，20日63.8%
    ['KDJ近3日內黃金交叉', '大盤跌破60日均線', '晨星剛形成'],  // 5日73.0%，10日73.8%，20日66.9%
    ['布林通道低檔(≤20%)', '相對強弱為正(強於大盤)', '爆量(≥2倍均量)'],  // 5日73.6%，10日65.6%，20日68.9%
    ['相對強弱為負(弱於大盤)', '大盤跌破20日均線', '晨星剛形成'],  // 5日73.4%，10日73.6%，20日73.2%
    ['大盤跌破60日均線', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],  // 5日72.3%，10日73.6%，20日65.1%
    ['跌深反彈盤', '大盤跌破60日均線', '近3月均價YoY為正'],  // 5日72.6%，10日63.1%，20日63.1%
    ['跌深反彈盤', '大盤跌破60日均線', '三大法人近3月買超'],  // 5日71.7%，10日66.0%，20日60.4%
    ['布林通道低檔(≤20%)', '爆量(≥2倍均量)', '大盤跌破60日均線'],  // 5日68.5%，10日66.3%，20日71.5%
    ['大盤跌破20日均線', '大盤跌破60日均線', '晨星剛形成'],  // 5日70.7%，10日71.0%，20日65.2%
    ['布林通道低檔(≤20%)', '爆量(≥2倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)'],  // 5日70.6%，10日60.8%
    ['布林通道低檔(≤20%)', '相對強弱為正(強於大盤)', '量能區間高檔(≥90百分位)'],  // 5日70.2%，10日66.9%，20日68.2%
    ['布林通道低檔(≤20%)', '相對強弱為正(強於大盤)', '夜星剛形成'],  // 5日70.0%，10日64.0%，20日60.0%
    ['大盤跌破60日均線', '三大法人近3月買超', '晨星剛形成'],  // 5日69.5%，10日70.0%，20日61.6%
    ['大盤跌破60日均線', '近3月均價YoY為正', '晨星剛形成'],  // 5日68.6%，10日69.8%，20日61.0%
    ['KDJ近3日內黃金交叉', '大盤跌破60日均線', '母子懷抱(低檔)剛形成'],  // 5日66.6%，10日69.5%，20日67.5%
    ['大盤跌破60日均線', '母子懷抱(低檔)剛形成', '晨星剛形成'],  // 5日64.9%，10日65.6%，20日69.5%
    ['爆量(≥1.5倍均量)', '大盤跌破60日均線', '晨星剛形成'],  // 5日64.9%，10日63.5%，20日69.2%
    ['布林通道低檔(≤20%)', '量能區間高檔(≥90百分位)', '大盤跌破60日均線'],  // 5日68.4%，10日66.7%，20日68.9%
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤跌破60日均線', '晨星剛形成'],  // 5日67.8%，10日68.9%，20日64.4%
    ['布林通道低檔(≤20%)', '爆量(≥1.5倍均量)', '大盤跌破60日均線'],  // 5日65.6%，10日65.7%，20日68.8%
    ['布林通道低檔(≤20%)', '爆量(≥2倍均量)', '型態剛形成(剛突破)'],  // 5日68.7%
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破60日均線', '晨星剛形成'],  // 5日68.4%，10日64.6%，20日64.6%
    ['布林通道低檔(≤20%)', '近3月均價YoY為負', '夜星剛形成'],  // 5日68.3%
    ['大盤跌破60日均線', '晨星剛形成'],  // 5日67.6%，10日68.3%，20日62.4%
    ['大盤跌破60日均線', '型態成形中', '晨星剛形成'],  // 5日67.6%，10日68.3%，20日62.4%
    ['大盤跌破60日均線', '型態突破確認', '晨星剛形成'],  // 5日67.6%，10日68.3%，20日62.4%
    ['大盤跌破60日均線', '型態剛形成(剛突破)', '晨星剛形成'],  // 5日67.6%，10日68.3%，20日62.4%
    ['布林通道低檔(≤20%)', '爆量(≥2倍均量)', '大盤跌破20日均線'],  // 5日66.6%，10日63.9%，20日68.3%
    ['跌深反彈盤', '大盤跌破20日均線', '大盤跌破60日均線'],  // 5日67.8%，10日65.3%，20日61.9%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '母子懷抱(低檔)剛形成'],  // 5日62.3%，10日67.5%，20日64.2%
    ['KDJ近3日內黃金交叉', '大盤跌破20日均線', '晨星剛形成'],  // 5日66.6%，10日67.3%，20日64.8%
    ['頭肩底剛形成', '三重底剛形成', '一字底(均線糾結)剛形成'],  // 5日61.5%，10日67.3%
    ['圓弧底剛形成', '一字底(均線糾結)剛形成', '突破飆股大量黑K最高點剛形成'],  // 10日67.3%
    ['地量(≤0.5倍均量)', '大盤跌破60日均線', '晨星剛形成'],  // 5日63.8%，10日67.2%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '夜星剛形成'],  // 5日67.0%，10日60.4%，20日63.2%
    ['爆量(≥1.5倍均量)', '大盤跌破60日均線', '母子懷抱(低檔)剛形成'],  // 5日66.0%，10日61.2%，20日67.0%
    ['布林通道低檔(≤20%)', '相對強弱為正(強於大盤)', '爆量(≥1.5倍均量)'],  // 5日66.8%，10日62.7%，20日64.3%
    ['大盤跌破60日均線', '近3月均價YoY為負', '晨星剛形成'],  // 5日66.8%，10日66.0%，20日64.8%
    ['布林通道低檔(≤20%)', '爆量(≥2倍均量)', '型態突破確認'],  // 5日66.7%
    ['爆量(≥1.5倍均量)', '大盤跌破60日均線', '母子懷抱(高檔)剛形成'],  // 5日66.7%，10日62.5%
    ['MACD近3日內黃金交叉', '大盤跌破60日均線', '夜星剛形成'],  // 5日61.1%，20日66.7%
    ['爆量(≥1.5倍均量)', '大盤跌破20日均線', '晨星剛形成'],  // 5日61.6%，10日60.5%，20日66.5%
    ['大盤跌破20日均線', '近3月均價YoY為負', '晨星剛形成'],  // 5日66.4%，10日65.7%，20日64.6%
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤跌破20日均線', '晨星剛形成'],  // 5日66.3%，10日65.7%，20日62.9%
    ['大盤跌破60日均線', '三大法人近3月賣超', '晨星剛形成'],  // 5日65.3%，10日66.2%，20日63.3%
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破20日均線', '夜星剛形成'],  // 5日66.0%
    ['大盤站上20日均線', '大盤跌破60日均線', '母子懷抱(低檔)剛形成'],  // 10日65.9%，20日63.0%
    ['KDJ近3日內死亡交叉', '大盤跌破60日均線', '晨星剛形成'],  // 5日63.2%，10日65.6%
    ['KDJ近3日內黃金交叉', '相對強弱為負(弱於大盤)', '晨星剛形成'],  // 5日64.5%，10日65.4%，20日61.3%
    ['相對強弱為負(弱於大盤)', '爆量(≥2倍均量)', '大盤跌破60日均線'],  // 10日60.7%，20日65.4%
    ['大盤跌破60日均線', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],  // 5日65.3%，10日65.3%，20日61.7%
    ['MACD近3日內黃金交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破60日均線'],  // 5日64.4%，10日65.3%，20日63.9%
    ['布林通道低檔(≤20%)', '大盤跌破20日均線', '夜星剛形成'],  // 5日65.2%
    ['大盤跌破20日均線', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],  // 5日64.9%，10日65.2%，20日60.9%
    ['布林通道低檔(≤20%)', '量能區間高檔(≥90百分位)', '大盤跌破20日均線'],  // 5日64.4%，10日62.6%，20日65.2%
    ['回後買上漲全通過', '一字底(均線糾結)剛形成', '突破飆股大量黑K最高點剛形成'],  // 10日65.2%
    ['布林通道低檔(≤20%)', '爆量(≥1.5倍均量)', '大盤跌破20日均線'],  // 5日62.9%，10日62.5%，20日65.1%
    ['相對強弱為負(弱於大盤)', '爆量(≥1.5倍均量)', '大盤跌破60日均線'],  // 10日61.7%，20日65.1%
    ['相對強弱為負(弱於大盤)', '大盤跌破20日均線', '大盤跌破60日均線'],  // 5日60.6%，10日63.4%，20日65.0%
    ['大盤跌破60日均線', '三大法人近3月買超', '夜星剛形成'],  // 5日64.8%，10日60.3%，20日60.3%
    ['跌深反彈盤', 'KDJ近3日內死亡交叉', '大盤跌破60日均線'],  // 5日64.6%，10日62.6%，20日63.6%
    ['跌深反彈盤', '大盤跌破60日均線', '近3月乖離度為負(股價超前營收)'],  // 5日64.6%，10日60.0%
    ['相對強弱為負(弱於大盤)', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],  // 5日64.6%，10日61.5%，20日61.9%
    ['多方力道≥65', '相對強弱為負(弱於大盤)', '大盤跌破60日均線'],  // 10日64.5%，20日60.8%
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月乖離度為負(股價超前營收)', '夜星剛形成'],  // 5日64.4%，10日61.5%
    ['布林通道低檔(≤20%)', '爆量(≥1.5倍均量)', '型態剛形成(剛突破)'],  // 5日64.3%，10日61.7%，20日63.6%
    ['近3月均價YoY為正', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 10日61.4%，20日64.2%
    ['大盤跌破20日均線', '母子懷抱(低檔)剛形成', '晨星剛形成'],  // 20日64.0%
    ['大盤站上20日均線', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 10日60.9%，20日63.8%
    ['跌深反彈盤', '大盤跌破20日均線', '三大法人近3月買超'],  // 5日63.7%
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', '晨星剛形成'],  // 5日63.7%
    ['MACD近3日內黃金交叉', '量能區間低檔(≤10百分位)', '大盤跌破60日均線'],  // 5日63.1%，10日63.6%，20日62.2%
    ['地量(≤0.5倍均量)', '大盤跌破60日均線', '回後買上漲全通過'],  // 10日61.5%，20日63.5%
    ['布林通道低檔(≤20%)', '爆量(≥1.5倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)'],  // 5日60.4%，20日63.4%
    ['MACD近3日內黃金交叉', 'KDJ近3日內黃金交叉', '母子懷抱(低檔)剛形成'],  // 10日63.4%，20日61.3%
    ['KDJ近3日內死亡交叉', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 10日63.3%，20日63.3%
    ['相對強弱為負(弱於大盤)', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],  // 5日63.2%，10日60.8%，20日61.7%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '型態剛形成(剛突破)'],  // 5日61.9%，10日63.2%，20日62.2%
    ['多方力道≥80', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],  // 5日63.1%
    ['大盤站上20日均線', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日63.1%
    ['布林通道低檔(≤20%)', 'KDJ近3日內死亡交叉', '大盤跌破60日均線'],  // 5日63.0%，10日61.0%，20日62.8%
    ['布林通道低檔(≤20%)', 'KDJ近3日內死亡交叉', '夜星剛形成'],  // 5日63.0%，20日61.6%
    ['大盤跌破20日均線', '三大法人近3月買超', '夜星剛形成'],  // 5日62.9%，20日60.2%
    ['大盤站上60日均線', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 10日62.9%，20日61.3%
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破20日均線', '晨星剛形成'],  // 20日62.9%
    ['布林通道低檔(≤20%)', '爆量(≥1.5倍均量)', '型態突破確認'],  // 5日62.8%，10日61.0%，20日62.8%
    ['相對強弱為負(弱於大盤)', '三大法人近3月買超', '晨星剛形成'],  // 5日61.7%，10日62.7%
    ['跌深反彈盤', '相對強弱為負(弱於大盤)', '大盤跌破60日均線'],  // 10日60.9%，20日62.7%
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', 'N字底剛形成'],  // 20日62.7%
    ['相對強弱為負(弱於大盤)', '近3月均價YoY為負', '晨星剛形成'],  // 5日62.6%，20日61.3%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '型態突破確認'],  // 5日61.7%，10日62.6%，20日62.0%
    ['MACD近3日內黃金交叉', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],  // 10日62.6%
    ['KDJ近3日內死亡交叉', '大盤跌破60日均線', '母子懷抱(低檔)剛形成'],  // 10日62.6%
    ['相對強弱為負(弱於大盤)', '爆量(≥2倍均量)', '大盤跌破20日均線'],  // 20日62.6%
    ['大盤跌破20日均線', '三大法人近3月買超', '晨星剛形成'],  // 5日62.5%，10日61.9%
    ['KDJ近3日內死亡交叉', '量能區間低檔(≤10百分位)', '母子懷抱(低檔)剛形成'],  // 10日62.5%
    ['KDJ近3日內死亡交叉', '一字底(均線糾結)剛形成', '突破飆股大量黑K最高點剛形成'],  // 10日62.5%
    ['頭肩底剛形成', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 10日62.5%，20日60.9%
    ['布林通道低檔(≤20%)', '相對強弱為負(弱於大盤)', '大盤跌破60日均線'],  // 20日62.4%
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)'],  // 5日60.9%，10日62.3%，20日62.0%
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', '母子懷抱(低檔)剛形成'],  // 20日62.3%
    ['MACD近3日內黃金交叉', '相對強弱為負(弱於大盤)', '大盤跌破60日均線'],  // 10日60.7%，20日62.2%
    ['相對強弱為負(弱於大盤)', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],  // 5日61.0%，10日62.1%
    ['強勢突破盤', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 10日60.3%，20日62.1%
    ['跌深反彈盤', '大盤跌破60日均線'],  // 5日62.0%，10日61.3%，20日60.6%
    ['跌深反彈盤', '布林通道低檔(≤20%)', '大盤跌破60日均線'],  // 5日62.0%，10日61.3%，20日60.6%
    ['跌深反彈盤', 'MACD近3日內黃金交叉', '大盤跌破60日均線'],  // 5日62.0%，10日61.3%，20日60.6%
    ['布林通道低檔(≤20%)', 'MACD近3日內黃金交叉', '大盤跌破60日均線'],  // 5日62.0%，10日61.9%，20日61.9%
    ['相對強弱為正(強於大盤)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],  // 5日62.0%
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '三大法人近3月買超', '夜星剛形成'],  // 5日62.0%，10日60.2%
    ['KDJ近3日內黃金交叉', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],  // 10日62.0%
    ['相對強弱為負(弱於大盤)', '爆量(≥1.5倍均量)', '大盤跌破20日均線'],  // 20日62.0%
    ['相對強弱為負(弱於大盤)', '母子懷抱(低檔)剛形成', '晨星剛形成'],  // 5日61.0%，10日60.5%，20日61.9%
    ['KDJ近3日內黃金交叉', '近3月均價YoY為正', '晨星剛形成'],  // 5日60.1%，10日61.9%
    ['多方力道≥80', '布林通道高檔(≥80%)', '量能區間低檔(≤10百分位)'],  // 10日61.9%
    ['MACD近3日內黃金交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)', 'N字底剛形成'],  // 10日61.9%
    ['近3月乖離度為負(股價超前營收)', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 10日61.9%，20日61.9%
    ['布林通道低檔(≤20%)', '相對強弱為正(強於大盤)', '量能斜率轉強(近5日均量>近10日均量20%以上)'],  // 5日61.8%
    ['爆量(≥2倍均量)', '大盤跌破60日均線', '晨星剛形成'],  // 5日61.8%，20日60.9%
    ['大盤跌破60日均線', '近3月均價YoY為正', '夜星剛形成'],  // 5日61.8%
    ['近3月均價YoY為正', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],  // 5日61.8%
    ['MACD近3日內黃金交叉', '地量(≤0.5倍均量)', '大盤跌破60日均線'],  // 10日61.8%，20日60.3%
    ['強勢突破盤', '地量(≤0.5倍均量)', '回後買上漲全通過'],  // 20日61.8%
    ['布林通道低檔(≤20%)', '量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤跌破60日均線'],  // 20日61.8%
    ['KDJ近3日內死亡交叉', '相對強弱為負(弱於大盤)', '大盤跌破60日均線'],  // 5日61.7%，20日61.7%
    ['量能區間高檔(≥90百分位)', '大盤跌破20日均線', '複式頭肩底剛形成'],  // 10日61.7%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '近3月均價YoY為正'],  // 10日60.3%，20日61.7%
    ['量能區間高檔(≥90百分位)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日61.7%
    ['大盤跌破20日均線', '晨星剛形成'],  // 5日60.9%，10日61.6%，20日60.2%
    ['大盤跌破20日均線', '型態成形中', '晨星剛形成'],  // 5日60.9%，10日61.6%，20日60.2%
    ['大盤跌破20日均線', '型態突破確認', '晨星剛形成'],  // 5日60.9%，10日61.6%，20日60.2%
    ['大盤跌破20日均線', '型態剛形成(剛突破)', '晨星剛形成'],  // 5日60.9%，10日61.6%，20日60.2%
    ['跌深反彈盤', '大盤跌破60日均線', '近3月乖離度為正(營收優於股價)'],  // 10日61.6%，20日60.3%
    ['布林通道高檔(≥80%)', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 10日60.3%，20日61.6%
    ['KDJ近3日內黃金交叉', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 20日61.6%
    ['KDJ近3日內黃金交叉', '母子懷抱(低檔)剛形成', '晨星剛形成'],  // 20日61.6%
    ['跌深反彈盤', '大盤跌破60日均線', '型態成形中'],  // 5日61.5%，10日60.8%，20日60.8%
    ['布林通道低檔(≤20%)', '近3月乖離度為正(營收優於股價)', '夜星剛形成'],  // 5日61.5%
    ['MACD近3日內黃金交叉', '大盤跌破20日均線', '晨星剛形成'],  // 5日61.5%
    ['近3月乖離度為負(股價超前營收)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],  // 5日61.5%
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破60日均線', '母子懷抱(低檔)剛形成'],  // 10日61.5%
    ['跌深反彈盤', 'KDJ近3日內黃金交叉', '大盤跌破20日均線'],  // 20日61.5%
    ['跌深反彈盤', 'KDJ近3日內黃金交叉', '近3月乖離度為正(營收優於股價)'],  // 20日61.5%
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', '大盤跌破60日均線'],  // 20日61.5%
    ['量能區間低檔(≤10百分位)', '大盤跌破60日均線', '母子懷抱(高檔)剛形成'],  // 20日61.5%
    ['大盤站上60日均線', '圓弧底剛形成', '晨星剛形成'],  // 20日61.5%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '夜星剛形成'],  // 5日61.4%，20日61.4%
    ['量能區間低檔(≤10百分位)', '大盤站上20日均線', '回後買上漲全通過'],  // 10日61.4%
    ['MACD近3日內黃金交叉', '相對強弱為正(強於大盤)', '母子懷抱(低檔)剛形成'],  // 20日61.4%
    ['爆量(≥1.5倍均量)', '大盤跌破20日均線', '母子懷抱(低檔)剛形成'],  // 20日61.4%
    ['多方力道≥65', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],  // 5日61.3%
    ['KDJ近3日內黃金交叉', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],  // 5日61.3%
    ['爆量(≥2倍均量)', '大盤跌破20日均線', '晨星剛形成'],  // 5日61.3%
    ['地量(≤0.5倍均量)', '量能區間低檔(≤10百分位)', 'N字底剛形成'],  // 10日61.3%
    ['多方力道≥80', '強勢突破盤', '地量(≤0.5倍均量)'],  // 10日60.3%，20日61.3%
    ['爆量(≥1.5倍均量)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日61.3%
    ['跌深反彈盤', '大盤跌破20日均線', '近3月均價YoY為正'],  // 5日61.2%
    ['相對強弱為負(弱於大盤)', '近3月均價YoY為正', '晨星剛形成'],  // 5日60.5%，10日61.2%
    ['大盤跌破20日均線', '三大法人近3月賣超', '晨星剛形成'],  // 10日61.2%，20日60.5%
    ['大盤跌破20日均線', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],  // 10日60.8%，20日61.2%
    ['KDJ近3日內黃金交叉', '大盤跌破20日均線', '母子懷抱(低檔)剛形成'],  // 20日61.2%
    ['KDJ近3日內黃金交叉', '三大法人近3月買超', '晨星剛形成'],  // 10日61.1%
    ['MACD近3日內黃金交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破20日均線'],  // 20日61.1%
    ['相對強弱為負(弱於大盤)', '量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤跌破60日均線'],  // 20日61.1%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '近3月乖離度為負(股價超前營收)'],  // 20日61.1%
    ['三大法人近3月買超', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 20日61.1%
    ['布林通道低檔(≤20%)', 'MACD近3日內黃金交叉', '量能區間低檔(≤10百分位)'],  // 5日61.0%
    ['大盤跌破60日均線', '三大法人近3月賣超', '母子懷抱(低檔)剛形成'],  // 10日61.0%
    ['K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日61.0%
    ['MACD近3日內黃金交叉', '大盤跌破20日均線', '夜星剛形成'],  // 20日61.0%
    ['KDJ近3日內黃金交叉', '相對強弱為負(弱於大盤)', '大盤跌破60日均線'],  // 20日61.0%
    ['相對強弱為負(弱於大盤)', '爆量(≥2倍均量)', '母子懷抱(低檔)剛形成'],  // 20日61.0%
    ['型態成形中', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日61.0%
    ['型態突破確認', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日61.0%
    ['型態剛形成(剛突破)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日61.0%
    ['大盤跌破60日均線', '近3月乖離度為正(營收優於股價)', '夜星剛形成'],  // 5日60.9%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '近3月乖離度為負(股價超前營收)'],  // 10日60.0%，20日60.9%
    ['布林通道高檔(≥80%)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日60.9%
    ['MACD近3日內黃金交叉', '量能區間低檔(≤10百分位)', '大盤跌破20日均線'],  // 20日60.9%
    ['相對強弱為負(弱於大盤)', '晨星剛形成'],  // 5日60.8%，10日60.4%
    ['相對強弱為負(弱於大盤)', '型態成形中', '晨星剛形成'],  // 5日60.8%，10日60.4%
    ['相對強弱為負(弱於大盤)', '型態突破確認', '晨星剛形成'],  // 5日60.8%，10日60.4%
    ['相對強弱為負(弱於大盤)', '型態剛形成(剛突破)', '晨星剛形成'],  // 5日60.8%，10日60.4%
    ['N字底剛形成', '一字底(均線糾結)剛形成', 'K線橫盤的突破剛形成'],  // 10日60.0%，20日60.8%
    ['頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.8%
    ['型態成形中', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.8%
    ['型態突破確認', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.8%
    ['型態剛形成(剛突破)', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.8%
    ['相對強弱為負(弱於大盤)', '大盤跌破20日均線', '夜星剛形成'],  // 5日60.7%
    ['大盤跌破60日均線', '母子懷抱(高檔)剛形成', '夜星剛形成'],  // 5日60.7%
    ['布林通道低檔(≤20%)', '三大法人近3月買超', '夜星剛形成'],  // 10日60.7%
    ['大盤跌破60日均線', '近3月乖離度為正(營收優於股價)', '母子懷抱(低檔)剛形成'],  // 10日60.7%
    ['大盤跌破60日均線', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],  // 10日60.7%，20日60.1%
    ['三重底剛形成', '圓弧底剛形成', '突破飆股大量黑K最高點剛形成'],  // 10日60.7%
    ['跌深反彈盤', '大盤跌破60日均線', '三大法人近3月賣超'],  // 20日60.7%
    ['大盤跌破60日均線', '三大法人近3月買超', '母子懷抱(低檔)剛形成'],  // 5日60.6%
    ['相對強弱為負(弱於大盤)', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 10日60.6%
    ['近3月均價YoY為正', '突破ABC修正下降切線剛形成', '晨星剛形成'],  // 10日60.6%
    ['布林通道低檔(≤20%)', 'KDJ近3日內黃金交叉', '大盤跌破60日均線'],  // 10日60.4%，20日60.6%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '母子懷抱(高檔)剛形成'],  // 10日60.2%，20日60.6%
    ['相對強弱為負(弱於大盤)', '量能區間高檔(≥90百分位)', '大盤跌破60日均線'],  // 20日60.6%
    ['相對強弱為負(弱於大盤)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破60日均線'],  // 20日60.6%
    ['大盤站上20日均線', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.6%
    ['大盤站上60日均線', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日60.6%
    ['大盤跌破20日均線', '大盤跌破60日均線', '夜星剛形成'],  // 5日60.5%
    ['KDJ近3日內黃金交叉', '晨星剛形成'],  // 10日60.5%
    ['KDJ近3日內黃金交叉', '型態成形中', '晨星剛形成'],  // 10日60.5%
    ['KDJ近3日內黃金交叉', '型態突破確認', '晨星剛形成'],  // 10日60.5%
    ['KDJ近3日內黃金交叉', '型態剛形成(剛突破)', '晨星剛形成'],  // 10日60.5%
    ['KDJ近3日內死亡交叉', '三重底剛形成', '突破飆股大量黑K最高點剛形成'],  // 10日60.5%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', 'N字底剛形成'],  // 10日60.5%
    ['大盤跌破20日均線', '近3月均價YoY為正', '晨星剛形成'],  // 10日60.5%
    ['近3月乖離度為負(股價超前營收)', '突破ABC修正下降切線剛形成', '晨星剛形成'],  // 10日60.5%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '三大法人近3月買超'],  // 10日60.2%，20日60.5%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '三大法人近3月買超'],  // 5日60.3%，10日60.4%
    ['大盤跌破60日均線', '近3月均價YoY為負', '母子懷抱(低檔)剛形成'],  // 10日60.4%
    ['布林通道低檔(≤20%)', '爆量(≥2倍均量)', '三大法人近3月買超'],  // 20日60.4%
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],  // 20日60.4%
    ['三大法人近3月買超', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.4%
    ['KDJ近3日內死亡交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破60日均線'],  // 5日60.3%
    ['大盤跌破60日均線', '母子懷抱(低檔)剛形成'],  // 10日60.3%
    ['多方力道≥65', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 10日60.3%
    ['KDJ近3日內黃金交叉', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],  // 10日60.3%
    ['大盤跌破60日均線', '型態成形中', '母子懷抱(低檔)剛形成'],  // 10日60.3%
    ['大盤跌破60日均線', '型態突破確認', '母子懷抱(低檔)剛形成'],  // 10日60.3%
    ['大盤跌破60日均線', '型態剛形成(剛突破)', '母子懷抱(低檔)剛形成'],  // 10日60.3%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '母子懷抱(高檔)剛形成'],  // 20日60.3%
    ['爆量(≥1.5倍均量)', '頭肩底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.3%
    ['量能區間低檔(≤10百分位)', '大盤跌破20日均線', '母子懷抱(高檔)剛形成'],  // 20日60.3%
    ['布林通道低檔(≤20%)', '大盤跌破20日均線', '大盤跌破60日均線'],  // 5日60.2%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '型態成形中'],  // 20日60.2%
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '型態剛形成(剛突破)'],  // 10日60.1%
    ['大盤站上20日均線', '大盤跌破60日均線', '母子懷抱(高檔)剛形成'],  // 10日60.1%，20日60.1%
    ['多方力道≥65', '頭肩底剛形成', '三重底剛形成'],  // 5日60.0%
    ['大盤跌破20日均線', '圓弧底剛形成', '突破飆股大量黑K最高點剛形成'],  // 5日60.0%
    ['多方力道≥80', '相對強弱為負(弱於大盤)', '突破飆股大量黑K最高點剛形成'],  // 10日60.0%
    ['KDJ近3日內黃金交叉', '突破ABC修正下降切線剛形成', '晨星剛形成'],  // 10日60.0%
    ['相對強弱為正(強於大盤)', '量能區間低檔(≤10百分位)', '母子懷抱(低檔)剛形成'],  // 10日60.0%
    ['跌深反彈盤', '爆量(≥1.5倍均量)', '三大法人近3月賣超'],  // 20日60.0%
    ['布林通道低檔(≤20%)', '大盤跌破60日均線', '近3月均價YoY為正'],  // 20日60.0%
    ['相對強弱為正(強於大盤)', 'N字底剛形成', '一字底(均線糾結)剛形成'],  // 20日60.0%
    // 2026-09-30 三年回測：大盤跌深時弱勢股開始跳空上漲（擇時型，三段 t 都 ≥2）
    ["相對強弱為負(弱於大盤)", "大盤跌破60日均線", "近3日向上跳空缺口"],
  ];

  var PINNED_COMBO_WINRATES = {
    '布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 晨星剛形成': { 5: 83.1, 10: 86.4, 20: 86.4 },
    '布林通道低檔(≤20%) ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 5: 82.8, 10: 85.9, 20: 85.9 },
    '布林通道低檔(≤20%) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 5: 81.1, 10: 83.5, 20: 81.1 },
    '布林通道低檔(≤20%) ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成': { 5: 77.9, 10: 82.4, 20: 74.4 },
    '布林通道低檔(≤20%) ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 5: 76.5, 10: 82.4, 20: 77.3 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 78.9, 10: 82.3, 20: 77.0 },
    '布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤) ＋ 晨星剛形成': { 5: 77.7, 10: 81.7, 20: 77.2 },
    '布林通道低檔(≤20%) ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 5: 76.3, 10: 80.4, 20: 74.4 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 79.7, 10: 80.0, 20: 76.4 },
    '布林通道低檔(≤20%) ＋ 晨星剛形成': { 5: 75.7, 10: 79.6, 20: 73.5 },
    '布林通道低檔(≤20%) ＋ 型態成形中 ＋ 晨星剛形成': { 5: 75.7, 10: 79.6, 20: 73.5 },
    '布林通道低檔(≤20%) ＋ 型態突破確認 ＋ 晨星剛形成': { 5: 75.7, 10: 79.6, 20: 73.5 },
    '布林通道低檔(≤20%) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 5: 75.7, 10: 79.6, 20: 73.5 },
    '布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 5: 74.8, 10: 78.5, 20: 73.8 },
    '布林通道低檔(≤20%) ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 5: 74.8, 10: 76.6, 20: 69.2 },
    '布林通道低檔(≤20%) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 5: 69.8, 10: 75.0, 20: 63.5 },
    '布林通道低檔(≤20%) ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 5: 70.1, 10: 74.8, 20: 63.8 },
    'KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 73.0, 10: 73.8, 20: 66.9 },
    '布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 爆量(≥2倍均量)': { 5: 73.6, 10: 65.6, 20: 68.9 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 5: 73.4, 10: 73.6, 20: 73.2 },
    '大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 5: 72.3, 10: 73.6, 20: 65.1 },
    '跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正': { 5: 72.6, 10: 63.1, 20: 63.1 },
    '跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 三大法人近3月買超': { 5: 71.7, 10: 66.0, 20: 60.4 },
    '布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 大盤跌破60日均線': { 5: 68.5, 10: 66.3, 20: 71.5 },
    '大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 70.7, 10: 71.0, 20: 65.2 },
    '布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)': { 5: 70.6, 10: 60.8 },
    '布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能區間高檔(≥90百分位)': { 5: 70.2, 10: 66.9, 20: 68.2 },
    '布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 夜星剛形成': { 5: 70.0, 10: 64.0, 20: 60.0 },
    '大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 5: 69.5, 10: 70.0, 20: 61.6 },
    '大盤跌破60日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 5: 68.6, 10: 69.8, 20: 61.0 },
    'KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 5: 66.6, 10: 69.5, 20: 67.5 },
    '大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 5: 64.9, 10: 65.6, 20: 69.5 },
    '爆量(≥1.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 64.9, 10: 63.5, 20: 69.2 },
    '布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位) ＋ 大盤跌破60日均線': { 5: 68.4, 10: 66.7, 20: 68.9 },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 67.8, 10: 68.9, 20: 64.4 },
    '布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破60日均線': { 5: 65.6, 10: 65.7, 20: 68.8 },
    '布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 型態剛形成(剛突破)': { 5: 68.7 },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 68.4, 10: 64.6, 20: 64.6 },
    '布林通道低檔(≤20%) ＋ 近3月均價YoY為負 ＋ 夜星剛形成': { 5: 68.3 },
    '大盤跌破60日均線 ＋ 晨星剛形成': { 5: 67.6, 10: 68.3, 20: 62.4 },
    '大盤跌破60日均線 ＋ 型態成形中 ＋ 晨星剛形成': { 5: 67.6, 10: 68.3, 20: 62.4 },
    '大盤跌破60日均線 ＋ 型態突破確認 ＋ 晨星剛形成': { 5: 67.6, 10: 68.3, 20: 62.4 },
    '大盤跌破60日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 5: 67.6, 10: 68.3, 20: 62.4 },
    '布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 大盤跌破20日均線': { 5: 66.6, 10: 63.9, 20: 68.3 },
    '跌深反彈盤 ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線': { 5: 67.8, 10: 65.3, 20: 61.9 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 5: 62.3, 10: 67.5, 20: 64.2 },
    'KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 5: 66.6, 10: 67.3, 20: 64.8 },
    '頭肩底剛形成 ＋ 三重底剛形成 ＋ 一字底(均線糾結)剛形成': { 5: 61.5, 10: 67.3 },
    '圓弧底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 10: 67.3 },
    '地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 63.8, 10: 67.2 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 5: 67.0, 10: 60.4, 20: 63.2 },
    '爆量(≥1.5倍均量) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 5: 66.0, 10: 61.2, 20: 67.0 },
    '布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 爆量(≥1.5倍均量)': { 5: 66.8, 10: 62.7, 20: 64.3 },
    '大盤跌破60日均線 ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 5: 66.8, 10: 66.0, 20: 64.8 },
    '布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 型態突破確認': { 5: 66.7 },
    '爆量(≥1.5倍均量) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成': { 5: 66.7, 10: 62.5 },
    'MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 5: 61.1, 20: 66.7 },
    '爆量(≥1.5倍均量) ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 5: 61.6, 10: 60.5, 20: 66.5 },
    '大盤跌破20日均線 ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 5: 66.4, 10: 65.7, 20: 64.6 },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 5: 66.3, 10: 65.7, 20: 62.9 },
    '大盤跌破60日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 5: 65.3, 10: 66.2, 20: 63.3 },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 5: 66.0 },
    '大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 10: 65.9, 20: 63.0 },
    'KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 63.2, 10: 65.6 },
    'KDJ近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 晨星剛形成': { 5: 64.5, 10: 65.4, 20: 61.3 },
    '相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 大盤跌破60日均線': { 10: 60.7, 20: 65.4 },
    '大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 5: 65.3, 10: 65.3, 20: 61.7 },
    'MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線': { 5: 64.4, 10: 65.3, 20: 63.9 },
    '布林通道低檔(≤20%) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 5: 65.2 },
    '大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 5: 64.9, 10: 65.2, 20: 60.9 },
    '布林通道低檔(≤20%) ＋ 量能區間高檔(≥90百分位) ＋ 大盤跌破20日均線': { 5: 64.4, 10: 62.6, 20: 65.2 },
    '回後買上漲全通過 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 10: 65.2 },
    '布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破20日均線': { 5: 62.9, 10: 62.5, 20: 65.1 },
    '相對強弱為負(弱於大盤) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破60日均線': { 10: 61.7, 20: 65.1 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線': { 5: 60.6, 10: 63.4, 20: 65.0 },
    '大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 5: 64.8, 10: 60.3, 20: 60.3 },
    '跌深反彈盤 ＋ KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線': { 5: 64.6, 10: 62.6, 20: 63.6 },
    '跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)': { 5: 64.6, 10: 60.0 },
    '相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 5: 64.6, 10: 61.5, 20: 61.9 },
    '多方力道≥65 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線': { 10: 64.5, 20: 60.8 },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成': { 5: 64.4, 10: 61.5 },
    '布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 型態剛形成(剛突破)': { 5: 64.3, 10: 61.7, 20: 63.6 },
    '近3月均價YoY為正 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 61.4, 20: 64.2 },
    '大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: 64.0 },
    '大盤站上20日均線 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 60.9, 20: 63.8 },
    '跌深反彈盤 ＋ 大盤跌破20日均線 ＋ 三大法人近3月買超': { 5: 63.7 },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成': { 5: 63.7 },
    'MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位) ＋ 大盤跌破60日均線': { 5: 63.1, 10: 63.6, 20: 62.2 },
    '地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 回後買上漲全通過': { 10: 61.5, 20: 63.5 },
    '布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)': { 5: 60.4, 20: 63.4 },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成': { 10: 63.4, 20: 61.3 },
    'KDJ近3日內死亡交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 63.3, 20: 63.3 },
    '相對強弱為負(弱於大盤) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 5: 63.2, 10: 60.8, 20: 61.7 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 型態剛形成(剛突破)': { 5: 61.9, 10: 63.2, 20: 62.2 },
    '多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 5: 63.1 },
    '大盤站上20日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 63.1 },
    '布林通道低檔(≤20%) ＋ KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線': { 5: 63.0, 10: 61.0, 20: 62.8 },
    '布林通道低檔(≤20%) ＋ KDJ近3日內死亡交叉 ＋ 夜星剛形成': { 5: 63.0, 20: 61.6 },
    '大盤跌破20日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 5: 62.9, 20: 60.2 },
    '大盤站上60日均線 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 62.9, 20: 61.3 },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 20: 62.9 },
    '布林通道低檔(≤20%) ＋ 爆量(≥1.5倍均量) ＋ 型態突破確認': { 5: 62.8, 10: 61.0, 20: 62.8 },
    '相對強弱為負(弱於大盤) ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 5: 61.7, 10: 62.7 },
    '跌深反彈盤 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線': { 10: 60.9, 20: 62.7 },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ N字底剛形成': { 20: 62.7 },
    '相對強弱為負(弱於大盤) ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 5: 62.6, 20: 61.3 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 型態突破確認': { 5: 61.7, 10: 62.6, 20: 62.0 },
    'MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 10: 62.6 },
    'KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 10: 62.6 },
    '相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 大盤跌破20日均線': { 20: 62.6 },
    '大盤跌破20日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 5: 62.5, 10: 61.9 },
    'KDJ近3日內死亡交叉 ＋ 量能區間低檔(≤10百分位) ＋ 母子懷抱(低檔)剛形成': { 10: 62.5 },
    'KDJ近3日內死亡交叉 ＋ 一字底(均線糾結)剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 10: 62.5 },
    '頭肩底剛形成 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 62.5, 20: 60.9 },
    '布林通道低檔(≤20%) ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線': { 20: 62.4 },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)': { 5: 60.9, 10: 62.3, 20: 62.0 },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成': { 20: 62.3 },
    'MACD近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線': { 10: 60.7, 20: 62.2 },
    '相對強弱為負(弱於大盤) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 5: 61.0, 10: 62.1 },
    '強勢突破盤 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 60.3, 20: 62.1 },
    '跌深反彈盤 ＋ 大盤跌破60日均線': { 5: 62.0, 10: 61.3, 20: 60.6 },
    '跌深反彈盤 ＋ 布林通道低檔(≤20%) ＋ 大盤跌破60日均線': { 5: 62.0, 10: 61.3, 20: 60.6 },
    '跌深反彈盤 ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線': { 5: 62.0, 10: 61.3, 20: 60.6 },
    '布林通道低檔(≤20%) ＋ MACD近3日內黃金交叉 ＋ 大盤跌破60日均線': { 5: 62.0, 10: 61.9, 20: 61.9 },
    '相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 5: 62.0 },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 5: 62.0, 10: 60.2 },
    'KDJ近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 10: 62.0 },
    '相對強弱為負(弱於大盤) ＋ 爆量(≥1.5倍均量) ＋ 大盤跌破20日均線': { 20: 62.0 },
    '相對強弱為負(弱於大盤) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 5: 61.0, 10: 60.5, 20: 61.9 },
    'KDJ近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 5: 60.1, 10: 61.9 },
    '多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 量能區間低檔(≤10百分位)': { 10: 61.9 },
    'MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ N字底剛形成': { 10: 61.9 },
    '近3月乖離度為負(股價超前營收) ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 61.9, 20: 61.9 },
    '布林通道低檔(≤20%) ＋ 相對強弱為正(強於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)': { 5: 61.8 },
    '爆量(≥2倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 5: 61.8, 20: 60.9 },
    '大盤跌破60日均線 ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 5: 61.8 },
    '近3月均價YoY為正 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 5: 61.8 },
    'MACD近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 大盤跌破60日均線': { 10: 61.8, 20: 60.3 },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 20: 61.8 },
    '布林通道低檔(≤20%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線': { 20: 61.8 },
    'KDJ近3日內死亡交叉 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線': { 5: 61.7, 20: 61.7 },
    '量能區間高檔(≥90百分位) ＋ 大盤跌破20日均線 ＋ 複式頭肩底剛形成': { 10: 61.7 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正': { 10: 60.3, 20: 61.7 },
    '量能區間高檔(≥90百分位) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 61.7 },
    '大盤跌破20日均線 ＋ 晨星剛形成': { 5: 60.9, 10: 61.6, 20: 60.2 },
    '大盤跌破20日均線 ＋ 型態成形中 ＋ 晨星剛形成': { 5: 60.9, 10: 61.6, 20: 60.2 },
    '大盤跌破20日均線 ＋ 型態突破確認 ＋ 晨星剛形成': { 5: 60.9, 10: 61.6, 20: 60.2 },
    '大盤跌破20日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 5: 60.9, 10: 61.6, 20: 60.2 },
    '跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價)': { 10: 61.6, 20: 60.3 },
    '布林通道高檔(≥80%) ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 60.3, 20: 61.6 },
    'KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 61.6 },
    'KDJ近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: 61.6 },
    '跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 型態成形中': { 5: 61.5, 10: 60.8, 20: 60.8 },
    '布林通道低檔(≤20%) ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成': { 5: 61.5 },
    'MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 5: 61.5 },
    '近3月乖離度為負(股價超前營收) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 5: 61.5 },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 10: 61.5 },
    '跌深反彈盤 ＋ KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線': { 20: 61.5 },
    '跌深反彈盤 ＋ KDJ近3日內黃金交叉 ＋ 近3月乖離度為正(營收優於股價)': { 20: 61.5 },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 大盤跌破60日均線': { 20: 61.5 },
    '量能區間低檔(≤10百分位) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成': { 20: 61.5 },
    '大盤站上60日均線 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: 61.5 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 5: 61.4, 20: 61.4 },
    '量能區間低檔(≤10百分位) ＋ 大盤站上20日均線 ＋ 回後買上漲全通過': { 10: 61.4 },
    'MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成': { 20: 61.4 },
    '爆量(≥1.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: 61.4 },
    '多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 5: 61.3 },
    'KDJ近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 5: 61.3 },
    '爆量(≥2倍均量) ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 5: 61.3 },
    '地量(≤0.5倍均量) ＋ 量能區間低檔(≤10百分位) ＋ N字底剛形成': { 10: 61.3 },
    '多方力道≥80 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)': { 10: 60.3, 20: 61.3 },
    '爆量(≥1.5倍均量) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 61.3 },
    '跌深反彈盤 ＋ 大盤跌破20日均線 ＋ 近3月均價YoY為正': { 5: 61.2 },
    '相對強弱為負(弱於大盤) ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 5: 60.5, 10: 61.2 },
    '大盤跌破20日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 10: 61.2, 20: 60.5 },
    '大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 10: 60.8, 20: 61.2 },
    'KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: 61.2 },
    'KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 10: 61.1 },
    'MACD近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線': { 20: 61.1 },
    '相對強弱為負(弱於大盤) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線': { 20: 61.1 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)': { 20: 61.1 },
    '三大法人近3月買超 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 61.1 },
    '布林通道低檔(≤20%) ＋ MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位)': { 5: 61.0 },
    '大盤跌破60日均線 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成': { 10: 61.0 },
    'K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 61.0 },
    'MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: 61.0 },
    'KDJ近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線': { 20: 61.0 },
    '相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: 61.0 },
    '型態成形中 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 61.0 },
    '型態突破確認 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 61.0 },
    '型態剛形成(剛突破) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 61.0 },
    '大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成': { 5: 60.9 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收)': { 10: 60.0, 20: 60.9 },
    '布林通道高檔(≥80%) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 60.9 },
    'MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位) ＋ 大盤跌破20日均線': { 20: 60.9 },
    '相對強弱為負(弱於大盤) ＋ 晨星剛形成': { 5: 60.8, 10: 60.4 },
    '相對強弱為負(弱於大盤) ＋ 型態成形中 ＋ 晨星剛形成': { 5: 60.8, 10: 60.4 },
    '相對強弱為負(弱於大盤) ＋ 型態突破確認 ＋ 晨星剛形成': { 5: 60.8, 10: 60.4 },
    '相對強弱為負(弱於大盤) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 5: 60.8, 10: 60.4 },
    'N字底剛形成 ＋ 一字底(均線糾結)剛形成 ＋ K線橫盤的突破剛形成': { 10: 60.0, 20: 60.8 },
    '頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.8 },
    '型態成形中 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.8 },
    '型態突破確認 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.8 },
    '型態剛形成(剛突破) ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.8 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 5: 60.7 },
    '大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成': { 5: 60.7 },
    '布林通道低檔(≤20%) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 10: 60.7 },
    '大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 母子懷抱(低檔)剛形成': { 10: 60.7 },
    '大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 10: 60.7, 20: 60.1 },
    '三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 10: 60.7 },
    '跌深反彈盤 ＋ 大盤跌破60日均線 ＋ 三大法人近3月賣超': { 20: 60.7 },
    '大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成': { 5: 60.6 },
    '相對強弱為負(弱於大盤) ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 60.6 },
    '近3月均價YoY為正 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 10: 60.6 },
    '布林通道低檔(≤20%) ＋ KDJ近3日內黃金交叉 ＋ 大盤跌破60日均線': { 10: 60.4, 20: 60.6 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成': { 10: 60.2, 20: 60.6 },
    '相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位) ＋ 大盤跌破60日均線': { 20: 60.6 },
    '相對強弱為負(弱於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線': { 20: 60.6 },
    '大盤站上20日均線 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.6 },
    '大盤站上60日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 60.6 },
    '大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 5: 60.5 },
    'KDJ近3日內黃金交叉 ＋ 晨星剛形成': { 10: 60.5 },
    'KDJ近3日內黃金交叉 ＋ 型態成形中 ＋ 晨星剛形成': { 10: 60.5 },
    'KDJ近3日內黃金交叉 ＋ 型態突破確認 ＋ 晨星剛形成': { 10: 60.5 },
    'KDJ近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 10: 60.5 },
    'KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 10: 60.5 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ N字底剛形成': { 10: 60.5 },
    '大盤跌破20日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 10: 60.5 },
    '近3月乖離度為負(股價超前營收) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 10: 60.5 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 三大法人近3月買超': { 10: 60.2, 20: 60.5 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 三大法人近3月買超': { 5: 60.3, 10: 60.4 },
    '大盤跌破60日均線 ＋ 近3月均價YoY為負 ＋ 母子懷抱(低檔)剛形成': { 10: 60.4 },
    '布林通道低檔(≤20%) ＋ 爆量(≥2倍均量) ＋ 三大法人近3月買超': { 20: 60.4 },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: 60.4 },
    '三大法人近3月買超 ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.4 },
    'KDJ近3日內死亡交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線': { 5: 60.3 },
    '大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 10: 60.3 },
    '多方力道≥65 ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 10: 60.3 },
    'KDJ近3日內黃金交叉 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 10: 60.3 },
    '大盤跌破60日均線 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成': { 10: 60.3 },
    '大盤跌破60日均線 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成': { 10: 60.3 },
    '大盤跌破60日均線 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成': { 10: 60.3 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成': { 20: 60.3 },
    '爆量(≥1.5倍均量) ＋ 頭肩底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.3 },
    '量能區間低檔(≤10百分位) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成': { 20: 60.3 },
    '布林通道低檔(≤20%) ＋ 大盤跌破20日均線 ＋ 大盤跌破60日均線': { 5: 60.2 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 型態成形中': { 20: 60.2 },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 型態剛形成(剛突破)': { 10: 60.1 },
    '大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成': { 10: 60.1, 20: 60.1 },
    '多方力道≥65 ＋ 頭肩底剛形成 ＋ 三重底剛形成': { 5: 60.0 },
    '大盤跌破20日均線 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 5: 60.0 },
    '多方力道≥80 ＋ 相對強弱為負(弱於大盤) ＋ 突破飆股大量黑K最高點剛形成': { 10: 60.0 },
    'KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 10: 60.0 },
    '相對強弱為正(強於大盤) ＋ 量能區間低檔(≤10百分位) ＋ 母子懷抱(低檔)剛形成': { 10: 60.0 },
    '跌深反彈盤 ＋ 爆量(≥1.5倍均量) ＋ 三大法人近3月賣超': { 20: 60.0 },
    '布林通道低檔(≤20%) ＋ 大盤跌破60日均線 ＋ 近3月均價YoY為正': { 20: 60.0 },
    '相對強弱為正(強於大盤) ＋ N字底剛形成 ＋ 一字底(均線糾結)剛形成': { 20: 60.0 },
    "相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 近3日向上跳空缺口": {"5": 57.6, "10": 71.1, "20": 72.9},
  };

  function formatWinrates(wrObj) {
    if (!wrObj) return '';
    return [5, 10, 20].filter(function(h){ return wrObj[h] != null; })
      .map(function(h){ return h + '日' + wrObj[h] + '%'; })
      .join(' / ');
  }

  var MOONSHOT_COMBOS = [
    ['突破飆股大量黑K最高點剛形成', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '突破ABC修正下降切線剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '回後買上漲全通過'],
    ['KDJ近3日內死亡交叉', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月乖離度為負(股價超前營收)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥80', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['強勢突破盤', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥80', '地量(≤0.5倍均量)', '晨星剛形成'],
    ['多方力道≥80', '強勢突破盤', '地量(≤0.5倍均量)'],
    ['三大法人近3月買超', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['三大法人近3月買超', '圓弧底剛形成', '晨星剛形成'],
    ['近3月乖離度為負(股價超前營收)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['突破飆股大量黑K最高點剛形成', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['大盤站上20日均線', '圓弧底剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '回後買上漲全通過', '突破ABC修正下降切線剛形成'],
    ['爆量(≥1.5倍均量)', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['布林通道高檔(≥80%)', '地量(≤0.5倍均量)', '母子懷抱(低檔)剛形成'],
    ['三大法人近3月賣超', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥65', '地量(≤0.5倍均量)', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['突破ABC修正下降切線剛形成', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['近3月乖離度為負(股價超前營收)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['近3月均價YoY為正', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '三重底剛形成', '晨星剛形成'],
    ['近3月均價YoY為負', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['大盤站上20日均線', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥65', '地量(≤0.5倍均量)', '夜星剛形成'],
    ['強勢突破盤', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['大盤站上60日均線', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥80', '地量(≤0.5倍均量)', '回後買上漲全通過'],
    ['爆量(≥1.5倍均量)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '三大法人近3月買超', '夜星剛形成'],
    ['大盤站上60日均線', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['地量(≤0.5倍均量)', '三大法人近3月買超', '夜星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破60日均線', '夜星剛形成'],
    ['近3月均價YoY為負', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '相對強弱為負(弱於大盤)', '夜星剛形成'],
    ['三大法人近3月賣超', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['型態成形中', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['型態突破確認', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['型態剛形成(剛突破)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['三大法人近3月賣超', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '地量(≤0.5倍均量)', '夜星剛形成'],
    ['近3月均價YoY為正', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '圓弧底剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月均價YoY為負', '圓弧底剛形成', '晨星剛形成'],
    ['三大法人近3月買超', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['布林通道高檔(≥80%)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['多方力道≥65', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '圓弧底剛形成', '晨星剛形成'],
    ['近3月均價YoY為負', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '大盤跌破20日均線', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['爆量(≥1.5倍均量)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月乖離度為負(股價超前營收)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['三大法人近3月賣超', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['強勢突破盤', '圓弧底剛形成', '晨星剛形成'],
    ['強勢突破盤', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '地量(≤0.5倍均量)', '回後買上漲全通過'],
    ['相對強弱為正(強於大盤)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '三大法人近3月買超', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '近3月均價YoY為負', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '三大法人近3月買超', '晨星剛形成'],
    ['大盤站上20日均線', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['型態成形中', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['型態突破確認', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['型態剛形成(剛突破)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['強勢突破盤', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['大盤跌破60日均線', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['多方力道≥80', '地量(≤0.5倍均量)', '夜星剛形成'],
    ['多方力道≥65', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '大盤跌破20日均線', '夜星剛形成'],
    ['三大法人近3月買超', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '近3月均價YoY為正', '晨星剛形成'],
    ['多方力道≥80', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['爆量(≥2倍均量)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥65', '強勢突破盤', '地量(≤0.5倍均量)'],
    ['KDJ近3日內黃金交叉', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥80', 'KDJ近3日內黃金交叉', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['量能區間高檔(≥90百分位)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['近3月均價YoY為正', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['地量(≤0.5倍均量)', '母子懷抱(高檔)剛形成', '夜星剛形成'],
    ['近3月乖離度為負(股價超前營收)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['強勢突破盤', 'N字底剛形成', '晨星剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', 'N字底剛形成'],
    ['多方力道≥80', '大盤跌破60日均線', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['大盤站上20日均線', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['大盤站上60日均線', '圓弧底剛形成', '晨星剛形成'],
    ['回後買上漲全通過', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(高檔)剛形成', '夜星剛形成'],
    ['KDJ近3日內黃金交叉', '圓弧底剛形成', '晨星剛形成'],
    ['爆量(≥2倍均量)', '圓弧底剛形成', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '圓弧底剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '大盤站上20日均線', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['多方力道≥80', '大盤站上20日均線', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '大盤跌破60日均線', '夜星剛形成'],
    ['量能區間高檔(≥90百分位)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['圓弧底剛形成', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['多方力道≥80', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['回後買上漲全通過', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['多方力道≥80', '相對強弱為正(強於大盤)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', 'KDJ近3日內黃金交叉', '夜星剛形成'],
    ['MACD近3日內黃金交叉', '大盤跌破60日均線', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '回後買上漲全通過', '近3月乖離度為負(股價超前營收)'],
    ['近3月乖離度為負(股價超前營收)', '圓弧底剛形成', '晨星剛形成'],
    ['近3月乖離度為負(股價超前營收)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['多方力道≥65', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '大盤跌破20日均線', '夜星剛形成'],
    ['大盤跌破20日均線', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['近3月乖離度為正(營收優於股價)', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '地量(≤0.5倍均量)', '母子懷抱(高檔)剛形成'],
    ['強勢突破盤', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['多方力道≥80', '晨星剛形成'],
    ['多方力道≥65', '多方力道≥80', '晨星剛形成'],
    ['多方力道≥80', '型態成形中', '晨星剛形成'],
    ['多方力道≥80', '型態突破確認', '晨星剛形成'],
    ['多方力道≥80', '型態剛形成(剛突破)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '量能區間高檔(≥90百分位)', '母子懷抱(低檔)剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤站上20日均線', '晨星剛形成'],
    ['大盤站上60日均線', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['三大法人近3月買超', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['近3月均價YoY為正', 'N字底剛形成', '母子懷抱(低檔)剛形成'],
    ['大盤站上60日均線', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['多方力道≥80', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['圓弧底剛形成', '晨星剛形成'],
    ['K線橫盤的突破剛形成', '晨星剛形成'],
    ['多方力道≥80', '量能區間低檔(≤10百分位)', '型態剛形成(剛突破)'],
    ['多方力道≥80', '大盤站上60日均線', '晨星剛形成'],
    ['強勢突破盤', '大盤跌破60日均線', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '地量(≤0.5倍均量)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '母子懷抱(高檔)剛形成', '夜星剛形成'],
    ['KDJ近3日內死亡交叉', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['爆量(≥2倍均量)', 'N字底剛形成', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤站上60日均線', '晨星剛形成'],
    ['大盤跌破60日均線', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['型態成形中', '圓弧底剛形成', '晨星剛形成'],
    ['型態成形中', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['型態突破確認', '圓弧底剛形成', '晨星剛形成'],
    ['型態突破確認', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['型態剛形成(剛突破)', '圓弧底剛形成', '晨星剛形成'],
    ['型態剛形成(剛突破)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['近3月乖離度為正(營收優於股價)', '圓弧底剛形成', '晨星剛形成'],
    ['近3月均價YoY為負', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月乖離度為正(營收優於股價)', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['多方力道≥80', 'KDJ近3日內死亡交叉', '晨星剛形成'],
    ['多方力道≥80', '爆量(≥2倍均量)', '晨星剛形成'],
    ['大盤站上60日均線', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['大盤跌破60日均線', '三大法人近3月買超', '夜星剛形成'],
    ['強勢突破盤', '大盤跌破20日均線', '母子懷抱(低檔)剛形成'],
    ['相對強弱為正(強於大盤)', '地量(≤0.5倍均量)', '晨星剛形成'],
    ['大盤跌破20日均線', '三大法人近3月買超', '夜星剛形成'],
    ['多方力道≥65', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '強勢突破盤', '晨星剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '型態剛形成(剛突破)'],
    ['強勢突破盤', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破20日均線', '夜星剛形成'],
    ['近3月乖離度為正(營收優於股價)', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['強勢突破盤', '相對強弱為正(強於大盤)', '地量(≤0.5倍均量)'],
    ['相對強弱為正(強於大盤)', '大盤跌破20日均線', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '回後買上漲全通過', 'N字底剛形成'],
    ['近3月均價YoY為正', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '回後買上漲全通過', '晨星剛形成'],
    ['多方力道≥65', '地量(≤0.5倍均量)', '回後買上漲全通過'],
    ['多方力道≥65', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['多方力道≥65', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '量能斜率轉強(近5日均量>近10日均量20%以上)', '母子懷抱(低檔)剛形成'],
    ['相對強弱為負(弱於大盤)', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['爆量(≥2倍均量)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['回後買上漲全通過', '突破ABC修正下降切線剛形成', '母子懷抱(低檔)剛形成'],
    ['量能區間高檔(≥90百分位)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['多方力道≥65', '三大法人近3月買超', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '地量(≤0.5倍均量)', '回後買上漲全通過'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月乖離度為負(股價超前營收)', '夜星剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '近3月乖離度為負(股價超前營收)'],
    ['多方力道≥80', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月均價YoY為負', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['多方力道≥65', '大盤站上20日均線', '晨星剛形成'],
    ['多方力道≥80', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['強勢突破盤', 'KDJ近3日內黃金交叉', '晨星剛形成'],
    ['強勢突破盤', '三大法人近3月賣超', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '三大法人近3月賣超', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['爆量(≥1.5倍均量)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '突破ABC修正下降切線剛形成', '母子懷抱(低檔)剛形成'],
    ['地量(≤0.5倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '回後買上漲全通過'],
    ['地量(≤0.5倍均量)', '大盤站上60日均線', '回後買上漲全通過'],
    ['地量(≤0.5倍均量)', '近3月乖離度為負(股價超前營收)', '夜星剛形成'],
    ['突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['多方力道≥80', '布林通道高檔(≥80%)', '晨星剛形成'],
    ['強勢突破盤', '三大法人近3月賣超', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['大盤跌破20日均線', '近3月均價YoY為正', '夜星剛形成'],
    ['型態成形中', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['型態突破確認', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['型態剛形成(剛突破)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '大盤站上60日均線', '夜星剛形成'],
    ['量能區間高檔(≥90百分位)', '大盤跌破20日均線', '夜星剛形成'],
    ['多方力道≥65', '三重底剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '地量(≤0.5倍均量)', '母子懷抱(高檔)剛形成'],
    ['多方力道≥80', '地量(≤0.5倍均量)', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '大盤站上20日均線', '晨星剛形成'],
    ['布林通道低檔(≤20%)', 'MACD近3日內黃金交叉', '夜星剛形成'],
    ['KDJ近3日內死亡交叉', '三重底剛形成', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內死亡交叉', '圓弧底剛形成', '突破ABC修正下降切線剛形成'],
    ['相對強弱為負(弱於大盤)', '爆量(≥2倍均量)', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['近3月均價YoY為負', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['圓弧底剛形成', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['多方力道≥65', '近3月均價YoY為正', '晨星剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '型態突破確認'],
    ['MACD近3日內黃金交叉', '相對強弱為正(強於大盤)', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '爆量(≥2倍均量)', '夜星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月均價YoY為正', '晨星剛形成'],
    ['大盤站上20日均線', 'N字底剛形成', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '三大法人近3月賣超', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['爆量(≥1.5倍均量)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['爆量(≥1.5倍均量)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '近3月均價YoY為正', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '三大法人近3月買超', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '型態成形中', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '型態突破確認', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '型態剛形成(剛突破)', '晨星剛形成'],
    ['大盤站上20日均線', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月乖離度為正(營收優於股價)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', 'MACD近3日內黃金交叉', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '地量(≤0.5倍均量)', '突破ABC修正下降切線剛形成'],
    ['布林通道高檔(≥80%)', '近3月均價YoY為正', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '相對強弱為正(強於大盤)', '晨星剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '三大法人近3月買超'],
    ['地量(≤0.5倍均量)', '回後買上漲全通過', '近3月均價YoY為正'],
    ['地量(≤0.5倍均量)', '回後買上漲全通過', '三大法人近3月買超'],
    ['三大法人近3月買超', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['相對強弱為負(弱於大盤)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['大盤站上60日均線', 'N字底剛形成', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月均價YoY為負', '夜星剛形成'],
    ['多方力道≥65', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '布林通道高檔(≥80%)', '地量(≤0.5倍均量)'],
    ['多方力道≥80', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '回後買上漲全通過'],
    ['多方力道≥80', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['強勢突破盤', 'MACD近3日內黃金交叉', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', '大盤跌破20日均線', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '近3月均價YoY為正', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '大盤站上20日均線', '回後買上漲全通過'],
    ['大盤站上20日均線', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['大盤站上20日均線', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['大盤站上60日均線', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '大盤站上60日均線', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '爆量(≥1.5倍均量)', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '大盤站上20日均線', '晨星剛形成'],
    ['爆量(≥2倍均量)', '三大法人近3月買超', '母子懷抱(低檔)剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤跌破20日均線', '母子懷抱(高檔)剛形成'],
    ['多方力道≥80', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '夜星剛形成'],
    ['多方力道≥80', '大盤跌破20日均線', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['相對強弱為負(弱於大盤)', '量能區間高檔(≥90百分位)', '母子懷抱(低檔)剛形成'],
    ['地量(≤0.5倍均量)', '型態成形中', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '型態突破確認', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '型態剛形成(剛突破)', '夜星剛形成'],
    ['大盤站上20日均線', '近3月均價YoY為正', '晨星剛形成'],
    ['大盤站上20日均線', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['回後買上漲全通過', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '大盤站上60日均線', '晨星剛形成'],
    ['多方力道≥80', '大盤跌破20日均線', '晨星剛形成'],
    ['強勢突破盤', '相對強弱為正(強於大盤)', '晨星剛形成'],
    ['布林通道高檔(≥80%)', 'KDJ近3日內黃金交叉', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '大盤站上20日均線', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', '突破飆股大量黑K最高點剛形成', 'K線橫盤的突破剛形成'],
    ['爆量(≥1.5倍均量)', '圓弧底剛形成', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '型態成形中', '夜星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '型態突破確認', '夜星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '型態剛形成(剛突破)', '夜星剛形成'],
    ['型態成形中', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['型態突破確認', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['型態剛形成(剛突破)', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['回後買上漲全通過', '近3月均價YoY為正', '母子懷抱(低檔)剛形成'],
    ['回後買上漲全通過', '三大法人近3月賣超', '母子懷抱(低檔)剛形成'],
    ['三大法人近3月買超', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內死亡交叉', '地量(≤0.5倍均量)', '夜星剛形成'],
    ['量能區間高檔(≥90百分位)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['多方力道≥65', '布林通道高檔(≥80%)', '晨星剛形成'],
    ['多方力道≥80', '強勢突破盤', '量能區間低檔(≤10百分位)'],
    ['多方力道≥80', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['強勢突破盤', '近3月均價YoY為正', '晨星剛形成'],
    ['強勢突破盤', '三重底剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '相對強弱為正(強於大盤)', '晨星剛形成'],
    ['大盤跌破20日均線', '大盤跌破60日均線', '夜星剛形成'],
    ['大盤站上60日均線', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['大盤站上60日均線', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月均價YoY為正', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['三大法人近3月買超', '頭肩底剛形成', '圓弧底剛形成'],
    ['三大法人近3月賣超', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', 'N字底剛形成', '晨星剛形成'],
    ['大盤站上20日均線', 'N字底剛形成', '母子懷抱(低檔)剛形成'],
    ['回後買上漲全通過', 'N字底剛形成', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', 'MACD近3日內黃金交叉', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '大盤站上20日均線'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '近3月均價YoY為正'],
    ['強勢突破盤', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', '相對強弱為正(強於大盤)', '夜星剛形成'],
    ['KDJ近3日內死亡交叉', '圓弧底剛形成', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['爆量(≥2倍均量)', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '大盤站上60日均線', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', '近3月均價YoY為正', '夜星剛形成'],
    ['大盤跌破60日均線', '近3月均價YoY為正', '夜星剛形成'],
    ['回後買上漲全通過', '圓弧底剛形成', '晨星剛形成'],
    ['回後買上漲全通過', '突破飆股大量黑K最高點剛形成', '晨星剛形成'],
    ['回後買上漲全通過', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['近3月均價YoY為正', '三大法人近3月買超', '夜星剛形成'],
    ['多方力道≥65', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '晨星剛形成'],
    ['大盤跌破20日均線', '夜星剛形成'],
    ['多方力道≥65', '型態成形中', '晨星剛形成'],
    ['多方力道≥65', '型態突破確認', '晨星剛形成'],
    ['多方力道≥65', '型態剛形成(剛突破)', '晨星剛形成'],
    ['多方力道≥80', 'MACD近3日內黃金交叉', 'N字底剛形成'],
    ['強勢突破盤', 'KDJ近3日內死亡交叉', '地量(≤0.5倍均量)'],
    ['強勢突破盤', '大盤跌破60日均線', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '型態成形中', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '型態突破確認', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '型態剛形成(剛突破)', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['大盤跌破20日均線', '型態成形中', '夜星剛形成'],
    ['大盤跌破20日均線', '型態突破確認', '夜星剛形成'],
    ['大盤跌破20日均線', '型態剛形成(剛突破)', '夜星剛形成'],
    ['近3月均價YoY為正', 'N字底剛形成', '晨星剛形成'],
    ['圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '大盤跌破20日均線', '夜星剛形成'],
    ['布林通道高檔(≥80%)', '地量(≤0.5倍均量)', '量能區間高檔(≥90百分位)'],
    ['布林通道高檔(≥80%)', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '三大法人近3月買超', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內黃金交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['相對強弱為負(弱於大盤)', '大盤跌破60日均線', '夜星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '三大法人近3月買超', '母子懷抱(高檔)剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '圓弧底剛形成', 'K線橫盤的突破剛形成'],
    ['大盤跌破20日均線', '大盤站上60日均線', '夜星剛形成'],
    ['大盤站上60日均線', '近3月均價YoY為正', '晨星剛形成'],
    ['大盤跌破60日均線', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['型態成形中', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['型態突破確認', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['型態剛形成(剛突破)', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月乖離度為正(營收優於股價)', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['多方力道≥65', '爆量(≥2倍均量)', '晨星剛形成'],
    ['多方力道≥65', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(高檔)剛形成'],
    ['多方力道≥80', '量能區間低檔(≤10百分位)', '近3月乖離度為負(股價超前營收)'],
    ['多方力道≥80', '近3月均價YoY為負', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '近3月乖離度為正(營收優於股價)', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '近3月均價YoY為正', '晨星剛形成'],
    ['大盤站上20日均線', '大盤站上60日均線', '晨星剛形成'],
    ['近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '夜星剛形成'],
    ['大盤站上20日均線', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['地量(≤0.5倍均量)', '回後買上漲全通過'],
    ['多方力道≥65', '強勢突破盤', '晨星剛形成'],
    ['多方力道≥65', '大盤跌破60日均線', '夜星剛形成'],
    ['多方力道≥65', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['多方力道≥80', '回後買上漲全通過', 'N字底剛形成'],
    ['布林通道高檔(≥80%)', '近3月均價YoY為正', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內黃金交叉', '地量(≤0.5倍均量)', '回後買上漲全通過'],
    ['相對強弱為正(強於大盤)', '三大法人近3月買超', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '型態成形中', '回後買上漲全通過'],
    ['地量(≤0.5倍均量)', '型態突破確認', '回後買上漲全通過'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月均價YoY為負', '晨星剛形成'],
    ['回後買上漲全通過', '圓弧底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['強勢突破盤', '晨星剛形成'],
    ['多方力道≥65', '布林通道高檔(≥80%)', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '布林通道高檔(≥80%)', '晨星剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '大盤站上60日均線'],
    ['強勢突破盤', '型態成形中', '晨星剛形成'],
    ['強勢突破盤', '型態突破確認', '晨星剛形成'],
    ['強勢突破盤', '型態剛形成(剛突破)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '爆量(≥1.5倍均量)', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['相對強弱為正(強於大盤)', '大盤站上60日均線', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月均價YoY為正', '夜星剛形成'],
    ['大盤站上20日均線', '三大法人近3月買超', '晨星剛形成'],
    ['大盤站上20日均線', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['近3月乖離度為負(股價超前營收)', '近3月均價YoY為負', '晨星剛形成'],
    ['三大法人近3月買超', '母子懷抱(高檔)剛形成', '夜星剛形成'],
    ['三重底剛形成', '圓弧底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['多方力道≥65', 'N字底剛形成', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['地量(≤0.5倍均量)', '晨星剛形成'],
    ['大盤站上20日均線', '晨星剛形成'],
    ['多方力道≥65', 'MACD近3日內黃金交叉', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '量能區間高檔(≥90百分位)'],
    ['強勢突破盤', '大盤站上60日均線', '晨星剛形成'],
    ['布林通道高檔(≥80%)', 'MACD近3日內黃金交叉', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '大盤跌破60日均線', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '型態成形中', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '型態突破確認', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '型態剛形成(剛突破)', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '近3月均價YoY為負', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '近3月均價YoY為負', '夜星剛形成'],
    ['量能區間高檔(≥90百分位)', '圓弧底剛形成', '晨星剛形成'],
    ['大盤站上20日均線', '型態成形中', '晨星剛形成'],
    ['大盤站上20日均線', '型態突破確認', '晨星剛形成'],
    ['大盤站上20日均線', '型態剛形成(剛突破)', '晨星剛形成'],
    ['大盤站上20日均線', '三大法人近3月賣超', '晨星剛形成'],
    ['大盤跌破60日均線', '近3月乖離度為負(股價超前營收)', '夜星剛形成'],
    ['近3月均價YoY為負', '突破飆股大量黑K最高點剛形成', 'K線橫盤的突破剛形成'],
    ['大盤跌破60日均線', '夜星剛形成'],
    ['多方力道≥80', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],
    ['多方力道≥80', '三大法人近3月買超', '夜星剛形成'],
    ['強勢突破盤', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '地量(≤0.5倍均量)', 'N字底剛形成'],
    ['布林通道高檔(≥80%)', '三大法人近3月買超', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', '三大法人近3月買超', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '地量(≤0.5倍均量)', '母子懷抱(高檔)剛形成'],
    ['爆量(≥2倍均量)', '大盤跌破20日均線', '母子懷抱(低檔)剛形成'],
    ['大盤站上60日均線', '三大法人近3月買超', '晨星剛形成'],
    ['大盤跌破60日均線', '型態成形中', '夜星剛形成'],
    ['大盤跌破60日均線', '型態突破確認', '夜星剛形成'],
    ['大盤跌破60日均線', '型態剛形成(剛突破)', '夜星剛形成'],
    ['近3月均價YoY為負', '圓弧底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)'],
    ['多方力道≥80', '強勢突破盤', '量能斜率轉弱(近5日均量<近10日均量20%以上)'],
    ['強勢突破盤', '布林通道高檔(≥80%)', '地量(≤0.5倍均量)'],
    ['強勢突破盤', 'KDJ近3日內黃金交叉', '地量(≤0.5倍均量)'],
    ['布林通道高檔(≥80%)', '大盤站上60日均線', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', '大盤跌破60日均線', '夜星剛形成'],
    ['大盤站上60日均線', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['大盤站上60日均線', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['近3月均價YoY為負', '三重底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['三大法人近3月買超', 'N字底剛形成', '晨星剛形成'],
    ['多方力道≥65', '近3月均價YoY為正', '夜星剛形成'],
    ['多方力道≥65', '三大法人近3月買超', '夜星剛形成'],
    ['多方力道≥80', '地量(≤0.5倍均量)', 'N字底剛形成'],
    ['多方力道≥80', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(高檔)剛形成'],
    ['強勢突破盤', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '相對強弱為負(弱於大盤)', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '大盤跌破20日均線', '夜星剛形成'],
    ['布林通道高檔(≥80%)', '突破ABC修正下降切線剛形成', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '三重底剛形成', '晨星剛形成'],
    ['地量(≤0.5倍均量)', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],
    ['大盤跌破20日均線', '近3月乖離度為正(營收優於股價)', '夜星剛形成'],
    ['近3月均價YoY為正', '圓弧底剛形成', '晨星剛形成'],
    ['大盤站上60日均線', '晨星剛形成'],
    ['多方力道≥65', 'KDJ近3日內黃金交叉', '晨星剛形成'],
    ['多方力道≥65', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['多方力道≥80', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '回後買上漲全通過'],
    ['強勢突破盤', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['MACD近3日內黃金交叉', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '三大法人近3月買超', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '大盤跌破60日均線', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '三大法人近3月買超', '晨星剛形成'],
    ['爆量(≥2倍均量)', '量能斜率轉強(近5日均量>近10日均量20%以上)', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '大盤站上60日均線', '母子懷抱(低檔)剛形成'],
    ['量能區間高檔(≥90百分位)', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤站上20日均線', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['大盤站上20日均線', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['大盤站上60日均線', '型態成形中', '晨星剛形成'],
    ['大盤站上60日均線', '型態突破確認', '晨星剛形成'],
    ['大盤站上60日均線', '型態剛形成(剛突破)', '晨星剛形成'],
    ['大盤跌破60日均線', '近3月均價YoY為負', '突破飆股大量黑K最高點剛形成'],
    ['大盤跌破60日均線', '突破飆股大量黑K最高點剛形成', 'K線橫盤的突破剛形成'],
    ['回後買上漲全通過', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['近3月均價YoY為正', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['多方力道≥65', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['多方力道≥80', '量能斜率轉強(近5日均量>近10日均量20%以上)', '突破ABC修正下降切線剛形成'],
    ['多方力道≥80', '回後買上漲全通過', '複式頭肩底剛形成'],
    ['多方力道≥80', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', 'KDJ近3日內死亡交叉', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '相對強弱為負(弱於大盤)', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '近3月乖離度為負(股價超前營收)', '晨星剛形成'],
    ['爆量(≥2倍均量)', '量能區間高檔(≥90百分位)', '晨星剛形成'],
    ['爆量(≥2倍均量)', '近3月乖離度為正(營收優於股價)', '母子懷抱(低檔)剛形成'],
    ['大盤跌破20日均線', '近3月乖離度為負(股價超前營收)', '夜星剛形成'],
    ['回後買上漲全通過', '突破飆股大量黑K最高點剛形成', '母子懷抱(低檔)剛形成'],
    ['近3月乖離度為正(營收優於股價)', '近3月均價YoY為正', '夜星剛形成'],
    ['三大法人近3月買超', '圓弧底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['三大法人近3月賣超', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '地量(≤0.5倍均量)', '母子懷抱(高檔)剛形成'],
    ['爆量(≥2倍均量)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['近3月乖離度為正(營收優於股價)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['多方力道≥65', 'KDJ近3日內死亡交叉', '晨星剛形成'],
    ['多方力道≥80', 'KDJ近3日內死亡交叉', '突破飆股大量黑K最高點剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)'],
    ['布林通道高檔(≥80%)', '大盤跌破20日均線', '晨星剛形成'],
    ['相對強弱為負(弱於大盤)', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['爆量(≥1.5倍均量)', 'N字底剛形成', '晨星剛形成'],
    ['爆量(≥2倍均量)', '三大法人近3月買超', '夜星剛形成'],
    ['回後買上漲全通過', '近3月均價YoY為正', '晨星剛形成'],
    ['三大法人近3月買超', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['大盤站上60日均線', 'N字底剛形成', '母子懷抱(低檔)剛形成'],
    ['N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['多方力道≥80', '回後買上漲全通過', '三大法人近3月買超'],
    ['強勢突破盤', '回後買上漲全通過', '晨星剛形成'],
    ['布林通道高檔(≥80%)', 'KDJ近3日內死亡交叉', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '大盤站上20日均線', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內黃金交叉', '大盤站上60日均線', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', '量能區間高檔(≥90百分位)', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '大盤站上60日均線', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤站上60日均線', '夜星剛形成'],
    ['型態成形中', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['型態突破確認', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['型態剛形成(剛突破)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['圓弧底剛形成', '突破飆股大量黑K最高點剛形成', 'K線橫盤的突破剛形成'],
    ['突破ABC修正下降切線剛形成', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '夜星剛形成'],
    ['三大法人近3月買超', '夜星剛形成'],
    ['多方力道≥80', 'KDJ近3日內黃金交叉', 'N字底剛形成'],
    ['多方力道≥80', '量能區間高檔(≥90百分位)', 'N字底剛形成'],
    ['強勢突破盤', '三大法人近3月買超', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '相對強弱為正(強於大盤)', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '三大法人近3月買超', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '型態成形中', '夜星剛形成'],
    ['MACD近3日內黃金交叉', '型態突破確認', '夜星剛形成'],
    ['MACD近3日內黃金交叉', '型態剛形成(剛突破)', '夜星剛形成'],
    ['KDJ近3日內黃金交叉', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['爆量(≥1.5倍均量)', '突破ABC修正下降切線剛形成', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '大盤跌破20日均線', '夜星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤跌破60日均線', '夜星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月乖離度為正(營收優於股價)', '晨星剛形成'],
    ['大盤站上60日均線', '三大法人近3月賣超', '晨星剛形成'],
    ['型態成形中', '三大法人近3月買超', '夜星剛形成'],
    ['型態突破確認', '三大法人近3月買超', '夜星剛形成'],
    ['型態剛形成(剛突破)', '三大法人近3月買超', '夜星剛形成'],
    ['近3月均價YoY為正', 'K線橫盤的突破剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '大盤跌破20日均線', '母子懷抱(高檔)剛形成'],
    ['多方力道≥65', '圓弧底剛形成', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '量能區間高檔(≥90百分位)', '晨星剛形成'],
    ['多方力道≥80', '大盤站上20日均線', '回後買上漲全通過'],
    ['多方力道≥80', '大盤跌破20日均線', '母子懷抱(高檔)剛形成'],
    ['多方力道≥80', '三大法人近3月買超', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '突破ABC修正下降切線剛形成'],
    ['MACD近3日內黃金交叉', '量能區間高檔(≥90百分位)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '近3月乖離度為負(股價超前營收)', '夜星剛形成'],
    ['KDJ近3日內黃金交叉', '爆量(≥2倍均量)', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內黃金交叉', '近3月均價YoY為正', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '母子懷抱(高檔)剛形成'],
    ['爆量(≥1.5倍均量)', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['爆量(≥2倍均量)', '大盤站上20日均線', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤站上60日均線', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '大盤站上20日均線', '夜星剛形成'],
    ['大盤跌破60日均線', '近3月乖離度為正(營收優於股價)', '突破飆股大量黑K最高點剛形成'],
    ['近3月乖離度為負(股價超前營收)', '三大法人近3月買超', '晨星剛形成'],
    ['近3月均價YoY為負', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', 'KDJ近3日內黃金交叉', '夜星剛形成'],
    ['多方力道≥65', '三大法人近3月買超', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '回後買上漲全通過', '近3月均價YoY為負'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '量能斜率轉強(近5日均量>近10日均量20%以上)'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '型態成形中'],
    ['布林通道高檔(≥80%)', '爆量(≥2倍均量)', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '量能區間高檔(≥90百分位)', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '型態成形中', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '型態突破確認', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', '型態剛形成(剛突破)', '母子懷抱(低檔)剛形成'],
    ['布林通道高檔(≥80%)', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['MACD近3日內黃金交叉', '量能區間高檔(≥90百分位)', '母子懷抱(高檔)剛形成'],
    ['MACD近3日內黃金交叉', '近3月均價YoY為正', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '近3月均價YoY為正', '晨星剛形成'],
    ['KDJ近3日內黃金交叉', 'KDJ近3日內死亡交叉', '夜星剛形成'],
    ['KDJ近3日內黃金交叉', '量能區間高檔(≥90百分位)', '夜星剛形成'],
    ['KDJ近3日內黃金交叉', '大盤站上20日均線', '晨星剛形成'],
    ['KDJ近3日內死亡交叉', '大盤跌破20日均線', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '近3月均價YoY為正', '晨星剛形成'],
    ['爆量(≥2倍均量)', '近3月均價YoY為負', '母子懷抱(低檔)剛形成'],
    ['量能區間高檔(≥90百分位)', '大盤站上20日均線', '晨星剛形成'],
    ['大盤站上20日均線', '大盤跌破60日均線', '突破飆股大量黑K最高點剛形成'],
    ['回後買上漲全通過', '近3月乖離度為正(營收優於股價)', '突破飆股大量黑K最高點剛形成'],
    ['近3月均價YoY為正', '三大法人近3月買超', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '晨星剛形成'],
    ['多方力道≥65', '近3月乖離度為負(股價超前營收)', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],
    ['多方力道≥80', '三大法人近3月賣超', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '爆量(≥1.5倍均量)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', '夜星剛形成'],
    ['MACD近3日內黃金交叉', '爆量(≥1.5倍均量)', '頭肩底剛形成'],
    ['MACD近3日內黃金交叉', '爆量(≥1.5倍均量)', '夜星剛形成'],
    ['MACD近3日內黃金交叉', '大盤站上60日均線', '母子懷抱(低檔)剛形成'],
    ['相對強弱為正(強於大盤)', '型態成形中', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '型態突破確認', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '型態剛形成(剛突破)', '晨星剛形成'],
    ['相對強弱為正(強於大盤)', '近3月乖離度為負(股價超前營收)', '夜星剛形成'],
    ['爆量(≥1.5倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '圓弧底剛形成'],
    ['地量(≤0.5倍均量)', '大盤站上20日均線', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '大盤跌破20日均線', '母子懷抱(高檔)剛形成'],
    ['地量(≤0.5倍均量)', '大盤跌破60日均線', '晨星剛形成'],
    ['量能斜率轉弱(近5日均量<近10日均量20%以上)', '近3月乖離度為正(營收優於股價)', '夜星剛形成'],
    ['大盤站上20日均線', '三重底剛形成', '母子懷抱(低檔)剛形成'],
    ['大盤跌破60日均線', '頭肩底剛形成', '複式頭肩底剛形成'],
    ['大盤跌破60日均線', '複式頭肩底剛形成', 'K線橫盤的突破剛形成'],
    ['大盤跌破60日均線', '母子懷抱(高檔)剛形成', '夜星剛形成'],
    ['MACD近3日內黃金交叉', '母子懷抱(低檔)剛形成'],
    ['N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['多方力道≥65', '布林通道高檔(≥80%)', '地量(≤0.5倍均量)'],
    ['多方力道≥65', '爆量(≥2倍均量)', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],
    ['多方力道≥65', '大盤跌破60日均線', '晨星剛形成'],
    ['強勢突破盤', '近3月均價YoY為正', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '近3月均價YoY為負', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '型態成形中', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '型態突破確認', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', '型態剛形成(剛突破)', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內黃金交叉', '量能區間高檔(≥90百分位)', '母子懷抱(低檔)剛形成'],
    ['相對強弱為正(強於大盤)', '三重底剛形成', '晨星剛形成'],
    ['相對強弱為負(弱於大盤)', '大盤站上60日均線', '晨星剛形成'],
    ['爆量(≥1.5倍均量)', '大盤站上20日均線', '晨星剛形成'],
    ['爆量(≥2倍均量)', '近3月均價YoY為正', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['量能區間高檔(≥90百分位)', '近3月均價YoY為正', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', '大盤跌破20日均線', '夜星剛形成'],
    ['型態成形中', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['型態突破確認', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['型態剛形成(剛突破)', 'N字底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['回後買上漲全通過', '近3月均價YoY為負', '突破飆股大量黑K最高點剛形成'],
    ['近3月均價YoY為正', '母子懷抱(高檔)剛形成', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '夜星剛形成'],
    ['爆量(≥2倍均量)', '母子懷抱(低檔)剛形成'],
    ['多方力道≥65', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '回後買上漲全通過'],
    ['多方力道≥80', '布林通道高檔(≥80%)', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', 'KDJ近3日內死亡交叉', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', '相對強弱為正(強於大盤)', '回後買上漲全通過'],
    ['多方力道≥80', '爆量(≥1.5倍均量)', 'N字底剛形成'],
    ['多方力道≥80', '大盤站上60日均線', '回後買上漲全通過'],
    ['多方力道≥80', '回後買上漲全通過', '近3月乖離度為正(營收優於股價)'],
    ['多方力道≥80', '近3月乖離度為正(營收優於股價)', '突破飆股大量黑K最高點剛形成'],
    ['強勢突破盤', '相對強弱為正(強於大盤)', '母子懷抱(低檔)剛形成'],
    ['強勢突破盤', '地量(≤0.5倍均量)', '近3月均價YoY為負'],
    ['MACD近3日內黃金交叉', '爆量(≥2倍均量)', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '三大法人近3月賣超', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '型態成形中', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '型態突破確認', '夜星剛形成'],
    ['相對強弱為正(強於大盤)', '型態剛形成(剛突破)', '夜星剛形成'],
    ['爆量(≥1.5倍均量)', '爆量(≥2倍均量)', '母子懷抱(低檔)剛形成'],
    ['爆量(≥1.5倍均量)', '大盤站上60日均線', '母子懷抱(低檔)剛形成'],
    ['爆量(≥1.5倍均量)', '三大法人近3月賣超', '晨星剛形成'],
    ['爆量(≥2倍均量)', '型態成形中', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '型態突破確認', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '型態剛形成(剛突破)', '母子懷抱(低檔)剛形成'],
    ['量能區間高檔(≥90百分位)', '量能斜率轉強(近5日均量>近10日均量20%以上)', '晨星剛形成'],
    ['量能區間高檔(≥90百分位)', '大盤站上60日均線', '晨星剛形成'],
    ['大盤跌破20日均線', '突破ABC修正下降切線剛形成', '突破飆股大量黑K最高點剛形成'],
    ['大盤跌破60日均線', '回後買上漲全通過', '複式頭肩底剛形成'],
    ['近3月乖離度為負(股價超前營收)', '三大法人近3月買超', '夜星剛形成'],
    ['三大法人近3月賣超', '母子懷抱(低檔)剛形成', '晨星剛形成'],
    ['量能斜率轉強(近5日均量>近10日均量20%以上)', 'N字底剛形成', '突破ABC修正下降切線剛形成'],
    ['量能區間低檔(≤10百分位)', '大盤跌破60日均線', '三重底剛形成'],
    ['大盤站上20日均線', '複式頭肩底剛形成', '突破ABC修正下降切線剛形成'],
    ['大盤站上60日均線', '頭肩底剛形成', '突破飆股大量黑K最高點剛形成'],
    ['多方力道≥65', 'KDJ近3日內死亡交叉', '母子懷抱(低檔)剛形成'],
    ['多方力道≥80', 'MACD近3日內黃金交叉', '量能區間低檔(≤10百分位)'],
    ['多方力道≥80', '大盤站上20日均線', 'N字底剛形成'],
    ['多方力道≥80', '三大法人近3月買超', 'N字底剛形成'],
    ['布林通道高檔(≥80%)', '近3月均價YoY為負', '晨星剛形成'],
    ['布林通道高檔(≥80%)', '三大法人近3月賣超', '母子懷抱(低檔)剛形成'],
    ['MACD近3日內黃金交叉', 'KDJ近3日內黃金交叉', '晨星剛形成'],
    ['MACD近3日內黃金交叉', 'KDJ近3日內死亡交叉', '母子懷抱(高檔)剛形成'],
    ['MACD近3日內黃金交叉', '大盤跌破60日均線', '晨星剛形成'],
    ['MACD近3日內黃金交叉', '近3月均價YoY為負', '母子懷抱(低檔)剛形成'],
    ['KDJ近3日內死亡交叉', '大盤站上20日均線', '晨星剛形成'],
    ['爆量(≥1.5倍均量)', '大盤站上20日均線', '母子懷抱(低檔)剛形成'],
    ['爆量(≥2倍均量)', '量能斜率轉弱(近5日均量<近10日均量20%以上)', '回後買上漲全通過'],
    ['地量(≤0.5倍均量)', '大盤跌破60日均線', '夜星剛形成'],
    ['地量(≤0.5倍均量)', '三大法人近3月買超', '母子懷抱(高檔)剛形成'],
    ['量能區間高檔(≥90百分位)', '頭肩底剛形成', '圓弧底剛形成'],
    ['大盤站上20日均線', '回後買上漲全通過', '突破飆股大量黑K最高點剛形成'],
    ['大盤站上20日均線', '回後買上漲全通過', '晨星剛形成'],
    ['大盤跌破20日均線', 'K線橫盤的突破剛形成', '晨星剛形成'],
    ['大盤站上60日均線', '回後買上漲全通過', '晨星剛形成'],
    ['回後買上漲全通過', 'N字底剛形成', '晨星剛形成'],
    ['近3月乖離度為正(營收優於股價)', '突破ABC修正下降切線剛形成', '晨星剛形成'],
    // 2026-09-30 三年回測：20日飆股比例三段都 ≥2 倍基準（舊版選股型 S1 移到這裡）
    ["多方力道≥80", "強勢突破盤", "地量(≤0.5倍均量)"],
    ["多方力道≥80", "量能斜率轉弱(近5日均量<近10日均量20%以上)", "晨星剛形成"],
    ["地量(≤0.5倍均量)", "創52週新高", "近3日向上跳空缺口"],
  ];

  var MOONSHOT_COMBO_STATS = {
    '突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:22, moonshotN:3, pct:13.6, avg:38.9}, 20: {n:22, moonshotN:7, pct:31.8, avg:47.0} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成': { 10: {n:20, moonshotN:2, pct:10.0, avg:55.6}, 20: {n:20, moonshotN:6, pct:30.0, avg:62.1} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 10: {n:76, moonshotN:10, pct:13.2, avg:41.4}, 20: {n:76, moonshotN:19, pct:25.0, avg:63.6} },
    'KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:24, moonshotN:3, pct:12.5, avg:42.8}, 20: {n:24, moonshotN:6, pct:25.0, avg:45.2} },
    '近3月乖離度為負(股價超前營收) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 10: {n:65, moonshotN:8, pct:12.3, avg:42.1}, 20: {n:65, moonshotN:15, pct:23.1, avg:54.4} },
    '多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 10: {n:104, moonshotN:12, pct:11.5, avg:43.1}, 20: {n:104, moonshotN:24, pct:23.1, avg:49.4} },
    '強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 10: {n:31, moonshotN:6, pct:19.4, avg:42.5}, 20: {n:31, moonshotN:7, pct:22.6, avg:46.0} },
    'MACD近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 10: {n:40, moonshotN:4, pct:10.0, avg:43.1}, 20: {n:40, moonshotN:9, pct:22.5, avg:38.6} },
    '突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:27, moonshotN:6, pct:22.2, avg:40.0} },
    '多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成': { 20: {n:55, moonshotN:12, pct:21.8, avg:60.5} },
    '多方力道≥80 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)': { 10: {n:292, moonshotN:30, pct:10.3, avg:47.7}, 20: {n:292, moonshotN:61, pct:20.9, avg:57.7} },
    '三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:29, moonshotN:6, pct:20.7, avg:51.8} },
    '突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:58, moonshotN:12, pct:20.7, avg:46.7} },
    '三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:34, moonshotN:7, pct:20.6, avg:43.4} },
    '近3月乖離度為負(股價超前營收) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:44, moonshotN:9, pct:20.5, avg:52.6} },
    '突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:54, moonshotN:11, pct:20.4, avg:49.0} },
    '地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 10: {n:65, moonshotN:7, pct:10.8, avg:42.6}, 20: {n:65, moonshotN:13, pct:20.0, avg:56.8} },
    '大盤站上20日均線 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:40, moonshotN:8, pct:20.0, avg:43.6} },
    '地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成': { 5: {n:26, moonshotN:3, pct:11.5, avg:35.0}, 10: {n:26, moonshotN:5, pct:19.2, avg:44.6}, 20: {n:26, moonshotN:5, pct:19.2, avg:58.2} },
    '爆量(≥1.5倍均量) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:21, moonshotN:4, pct:19.0, avg:40.5} },
    '多方力道≥65 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 5: {n:32, moonshotN:4, pct:12.5, avg:33.9}, 10: {n:32, moonshotN:6, pct:18.8, avg:42.7}, 20: {n:32, moonshotN:5, pct:15.6, avg:43.9} },
    '布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(低檔)剛形成': { 10: {n:32, moonshotN:4, pct:12.5, avg:46.0}, 20: {n:32, moonshotN:6, pct:18.8, avg:68.0} },
    '三大法人近3月賣超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:48, moonshotN:9, pct:18.8, avg:62.8} },
    '多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成': { 20: {n:65, moonshotN:12, pct:18.5, avg:60.5} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:27, moonshotN:5, pct:18.5, avg:44.3} },
    '突破ABC修正下降切線剛形成 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:27, moonshotN:5, pct:18.5, avg:44.5} },
    '近3月乖離度為負(股價超前營收) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:38, moonshotN:4, pct:10.5, avg:46.8}, 20: {n:38, moonshotN:7, pct:18.4, avg:44.9} },
    '近3月均價YoY為正 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 10: {n:76, moonshotN:8, pct:10.5, avg:42.5}, 20: {n:76, moonshotN:14, pct:18.4, avg:57.9} },
    'KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:33, moonshotN:4, pct:12.1, avg:41.4}, 20: {n:33, moonshotN:6, pct:18.2, avg:45.2} },
    '多方力道≥65 ＋ 三重底剛形成 ＋ 晨星剛形成': { 20: {n:22, moonshotN:4, pct:18.2, avg:54.0} },
    '近3月均價YoY為負 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:33, moonshotN:6, pct:18.2, avg:51.8} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:111, moonshotN:20, pct:18.0, avg:52.8} },
    '大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:111, moonshotN:20, pct:18.0, avg:54.1} },
    '多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成': { 10: {n:67, moonshotN:9, pct:13.4, avg:40.0}, 20: {n:67, moonshotN:12, pct:17.9, avg:65.0} },
    '強勢突破盤 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:39, moonshotN:7, pct:17.9, avg:43.7} },
    '大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:117, moonshotN:21, pct:17.9, avg:54.1} },
    '多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 10: {n:146, moonshotN:16, pct:11.0, avg:41.0}, 20: {n:146, moonshotN:26, pct:17.8, avg:63.1} },
    '爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:130, moonshotN:23, pct:17.7, avg:55.3} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 10: {n:108, moonshotN:12, pct:11.1, avg:40.4}, 20: {n:108, moonshotN:19, pct:17.6, avg:53.6} },
    '大盤站上60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:74, moonshotN:13, pct:17.6, avg:49.1} },
    '地量(≤0.5倍均量) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 10: {n:97, moonshotN:10, pct:10.3, avg:40.4}, 20: {n:97, moonshotN:17, pct:17.5, avg:56.6} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 10: {n:40, moonshotN:4, pct:10.0, avg:47.5}, 20: {n:40, moonshotN:7, pct:17.5, avg:61.0} },
    '近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:40, moonshotN:7, pct:17.5, avg:46.7} },
    '布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤) ＋ 夜星剛形成': { 20: {n:23, moonshotN:4, pct:17.4, avg:41.3} },
    '三大法人近3月賣超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:23, moonshotN:4, pct:17.4, avg:52.8} },
    '突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:145, moonshotN:25, pct:17.2, avg:53.7} },
    '型態成形中 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:145, moonshotN:25, pct:17.2, avg:53.7} },
    '型態突破確認 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:145, moonshotN:25, pct:17.2, avg:53.7} },
    '型態剛形成(剛突破) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:145, moonshotN:25, pct:17.2, avg:53.7} },
    '三大法人近3月賣超 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:70, moonshotN:12, pct:17.1, avg:53.4} },
    '相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成': { 10: {n:94, moonshotN:12, pct:12.8, avg:38.8}, 20: {n:94, moonshotN:16, pct:17.0, avg:59.8} },
    '近3月均價YoY為正 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 10: {n:112, moonshotN:12, pct:10.7, avg:41.4}, 20: {n:112, moonshotN:19, pct:17.0, avg:48.5} },
    '相對強弱為正(強於大盤) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:59, moonshotN:10, pct:16.9, avg:45.0} },
    '強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成': { 10: {n:24, moonshotN:4, pct:16.7, avg:39.3}, 20: {n:24, moonshotN:3, pct:12.5, avg:46.4} },
    '多方力道≥65 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:48, moonshotN:8, pct:16.7, avg:43.5} },
    '布林通道高檔(≥80%) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:138, moonshotN:23, pct:16.7, avg:54.6} },
    'KDJ近3日內黃金交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:42, moonshotN:7, pct:16.7, avg:49.4} },
    '近3月均價YoY為負 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:36, moonshotN:6, pct:16.7, avg:43.6} },
    '三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:96, moonshotN:16, pct:16.7, avg:48.5} },
    '布林通道高檔(≥80%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 10: {n:79, moonshotN:8, pct:10.1, avg:42.6}, 20: {n:79, moonshotN:13, pct:16.5, avg:46.2} },
    '布林通道高檔(≥80%) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:158, moonshotN:26, pct:16.5, avg:48.4} },
    '多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:171, moonshotN:28, pct:16.4, avg:49.6} },
    'KDJ近3日內死亡交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:61, moonshotN:10, pct:16.4, avg:46.3} },
    '相對強弱為正(強於大盤) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:55, moonshotN:9, pct:16.4, avg:43.1} },
    '近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:67, moonshotN:11, pct:16.4, avg:48.3} },
    '布林通道高檔(≥80%) ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:98, moonshotN:16, pct:16.3, avg:52.3} },
    '布林通道高檔(≥80%) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:80, moonshotN:13, pct:16.3, avg:49.8} },
    '爆量(≥1.5倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:80, moonshotN:13, pct:16.3, avg:49.5} },
    '近3月乖離度為負(股價超前營收) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:43, moonshotN:7, pct:16.3, avg:47.4} },
    '三大法人近3月賣超 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:37, moonshotN:5, pct:13.5, avg:38.2}, 20: {n:37, moonshotN:6, pct:16.2, avg:34.4} },
    '強勢突破盤 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:37, moonshotN:6, pct:16.2, avg:43.0} },
    '強勢突破盤 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:37, moonshotN:6, pct:16.2, avg:40.5} },
    '布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 20: {n:179, moonshotN:29, pct:16.2, avg:62.1} },
    '相對強弱為正(強於大盤) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:148, moonshotN:24, pct:16.2, avg:49.7} },
    '突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:87, moonshotN:14, pct:16.1, avg:48.4} },
    '多方力道≥80 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:292, moonshotN:47, pct:16.1, avg:52.3} },
    'MACD近3日內黃金交叉 ＋ 近3月均價YoY為負 ＋ 夜星剛形成': { 20: {n:62, moonshotN:10, pct:16.1, avg:47.0} },
    '相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:137, moonshotN:22, pct:16.1, avg:55.1} },
    '地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:87, moonshotN:14, pct:16.1, avg:52.9} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:211, moonshotN:34, pct:16.1, avg:47.2} },
    '大盤站上20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:118, moonshotN:19, pct:16.1, avg:46.8} },
    '型態成形中 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:87, moonshotN:14, pct:16.1, avg:48.4} },
    '型態突破確認 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:87, moonshotN:14, pct:16.1, avg:48.4} },
    '型態剛形成(剛突破) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:87, moonshotN:14, pct:16.1, avg:48.4} },
    'K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:62, moonshotN:10, pct:16.1, avg:45.0} },
    '強勢突破盤 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 5: {n:25, moonshotN:3, pct:12.0, avg:34.0}, 10: {n:25, moonshotN:4, pct:16.0, avg:43.9}, 20: {n:25, moonshotN:3, pct:12.0, avg:48.5} },
    '大盤跌破60日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:25, moonshotN:4, pct:16.0, avg:42.4} },
    '多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成': { 10: {n:44, moonshotN:7, pct:15.9, avg:42.2}, 20: {n:44, moonshotN:7, pct:15.9, avg:77.3} },
    '多方力道≥65 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:63, moonshotN:10, pct:15.9, avg:47.2} },
    'MACD近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:82, moonshotN:13, pct:15.9, avg:48.8} },
    '三大法人近3月買超 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:63, moonshotN:10, pct:15.9, avg:46.6} },
    '多方力道≥80 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 10: {n:240, moonshotN:25, pct:10.4, avg:43.1}, 20: {n:240, moonshotN:38, pct:15.8, avg:54.3} },
    '多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:203, moonshotN:32, pct:15.8, avg:56.5} },
    '爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:101, moonshotN:16, pct:15.8, avg:55.6} },
    '多方力道≥65 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)': { 20: {n:452, moonshotN:71, pct:15.7, avg:56.6} },
    'KDJ近3日內黃金交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:70, moonshotN:11, pct:15.7, avg:64.1} },
    '多方力道≥80 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成': { 10: {n:154, moonshotN:18, pct:11.7, avg:45.7}, 20: {n:154, moonshotN:24, pct:15.6, avg:54.4} },
    '布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:77, moonshotN:12, pct:15.6, avg:44.0} },
    '布林通道高檔(≥80%) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:64, moonshotN:10, pct:15.6, avg:45.0} },
    '量能區間高檔(≥90百分位) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:128, moonshotN:20, pct:15.6, avg:55.3} },
    '近3月均價YoY為正 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:45, moonshotN:7, pct:15.6, avg:50.1} },
    '爆量(≥2倍均量) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:58, moonshotN:9, pct:15.5, avg:51.7} },
    '地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成': { 20: {n:58, moonshotN:9, pct:15.5, avg:55.8} },
    '近3月乖離度為負(股價超前營收) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:103, moonshotN:16, pct:15.5, avg:50.0} },
    '強勢突破盤 ＋ N字底剛形成 ＋ 晨星剛形成': { 10: {n:26, moonshotN:4, pct:15.4, avg:41.9}, 20: {n:26, moonshotN:3, pct:11.5, avg:68.4} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ N字底剛形成': { 10: {n:26, moonshotN:3, pct:11.5, avg:39.4}, 20: {n:26, moonshotN:4, pct:15.4, avg:59.3} },
    '多方力道≥80 ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 20: {n:52, moonshotN:8, pct:15.4, avg:46.3} },
    '相對強弱為正(強於大盤) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:78, moonshotN:12, pct:15.4, avg:51.1} },
    '爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:52, moonshotN:8, pct:15.4, avg:46.2} },
    '量能區間高檔(≥90百分位) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:26, moonshotN:4, pct:15.4, avg:40.5} },
    '大盤站上20日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:149, moonshotN:23, pct:15.4, avg:49.0} },
    '大盤站上60日均線 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:52, moonshotN:8, pct:15.4, avg:43.6} },
    '回後買上漲全通過 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:39, moonshotN:6, pct:15.4, avg:47.0} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成': { 20: {n:59, moonshotN:9, pct:15.3, avg:61.4} },
    'KDJ近3日內黃金交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:33, moonshotN:5, pct:15.2, avg:45.3} },
    '爆量(≥2倍均量) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:33, moonshotN:5, pct:15.2, avg:44.6} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:66, moonshotN:10, pct:15.2, avg:49.5} },
    '布林通道高檔(≥80%) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:53, moonshotN:8, pct:15.1, avg:44.5} },
    '地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:93, moonshotN:14, pct:15.1, avg:61.7} },
    '相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 10: {n:100, moonshotN:12, pct:12.0, avg:40.3}, 20: {n:100, moonshotN:15, pct:15.0, avg:59.1} },
    '多方力道≥80 ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:334, moonshotN:50, pct:15.0, avg:54.8} },
    '布林通道高檔(≥80%) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:40, moonshotN:6, pct:15.0, avg:47.8} },
    '量能區間高檔(≥90百分位) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:60, moonshotN:9, pct:15.0, avg:44.5} },
    '圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:60, moonshotN:9, pct:15.0, avg:47.7} },
    '多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 10: {n:47, moonshotN:7, pct:14.9, avg:41.6}, 20: {n:47, moonshotN:6, pct:12.8, avg:72.4} },
    '回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 10: {n:47, moonshotN:5, pct:10.6, avg:35.1}, 20: {n:47, moonshotN:7, pct:14.9, avg:47.7} },
    '多方力道≥80 ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成': { 20: {n:397, moonshotN:59, pct:14.9, avg:52.9} },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成': { 20: {n:74, moonshotN:11, pct:14.9, avg:50.8} },
    'MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:54, moonshotN:8, pct:14.8, avg:48.3} },
    '地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 近3月乖離度為負(股價超前營收)': { 20: {n:155, moonshotN:23, pct:14.8, avg:61.3} },
    '近3月乖離度為負(股價超前營收) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:27, moonshotN:4, pct:14.8, avg:40.6} },
    '近3月乖離度為負(股價超前營收) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:81, moonshotN:12, pct:14.8, avg:49.7} },
    '多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 10: {n:75, moonshotN:8, pct:10.7, avg:40.6}, 20: {n:75, moonshotN:11, pct:14.7, avg:56.6} },
    '地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:75, moonshotN:11, pct:14.7, avg:61.5} },
    '大盤跌破20日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:34, moonshotN:5, pct:14.7, avg:51.8} },
    '近3月乖離度為正(營收優於股價) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:34, moonshotN:5, pct:14.7, avg:55.7} },
    '多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成': { 20: {n:157, moonshotN:23, pct:14.6, avg:54.4} },
    '強勢突破盤 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:103, moonshotN:15, pct:14.6, avg:51.8} },
    '相對強弱為正(強於大盤) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:123, moonshotN:18, pct:14.6, avg:46.8} },
    '量能區間高檔(≥90百分位) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:137, moonshotN:20, pct:14.6, avg:47.9} },
    '多方力道≥80 ＋ 晨星剛形成': { 20: {n:408, moonshotN:59, pct:14.5, avg:52.9} },
    '多方力道≥65 ＋ 多方力道≥80 ＋ 晨星剛形成': { 20: {n:408, moonshotN:59, pct:14.5, avg:52.9} },
    '多方力道≥80 ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:408, moonshotN:59, pct:14.5, avg:52.9} },
    '多方力道≥80 ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:408, moonshotN:59, pct:14.5, avg:52.9} },
    '多方力道≥80 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:408, moonshotN:59, pct:14.5, avg:52.9} },
    'MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成': { 20: {n:62, moonshotN:9, pct:14.5, avg:48.6} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:228, moonshotN:33, pct:14.5, avg:51.0} },
    '大盤站上60日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:125, moonshotN:18, pct:14.4, avg:47.2} },
    '三大法人近3月買超 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:104, moonshotN:15, pct:14.4, avg:49.2} },
    '近3月均價YoY為正 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:21, moonshotN:3, pct:14.3, avg:36.2}, 20: {n:21, moonshotN:3, pct:14.3, avg:51.0} },
    '大盤站上60日均線 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:63, moonshotN:8, pct:12.7, avg:41.9}, 20: {n:63, moonshotN:9, pct:14.3, avg:42.1} },
    '多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 10: {n:35, moonshotN:4, pct:11.4, avg:36.2}, 20: {n:35, moonshotN:5, pct:14.3, avg:58.4} },
    '圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:63, moonshotN:9, pct:14.3, avg:43.1} },
    'K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:189, moonshotN:27, pct:14.3, avg:48.1} },
    '多方力道≥80 ＋ 量能區間低檔(≤10百分位) ＋ 型態剛形成(剛突破)': { 20: {n:42, moonshotN:6, pct:14.3, avg:53.9} },
    '多方力道≥80 ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:356, moonshotN:51, pct:14.3, avg:53.9} },
    '強勢突破盤 ＋ 大盤跌破60日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:28, moonshotN:4, pct:14.3, avg:38.1} },
    '布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成': { 20: {n:28, moonshotN:4, pct:14.3, avg:42.4} },
    'MACD近3日內黃金交叉 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成': { 20: {n:56, moonshotN:8, pct:14.3, avg:44.7} },
    'KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:63, moonshotN:9, pct:14.3, avg:56.3} },
    '相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:244, moonshotN:35, pct:14.3, avg:49.9} },
    '爆量(≥2倍均量) ＋ N字底剛形成 ＋ 晨星剛形成': { 20: {n:21, moonshotN:3, pct:14.3, avg:68.4} },
    '量能區間高檔(≥90百分位) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:84, moonshotN:12, pct:14.3, avg:42.7} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:21, moonshotN:3, pct:14.3, avg:40.6} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:273, moonshotN:39, pct:14.3, avg:50.6} },
    '大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:28, moonshotN:4, pct:14.3, avg:51.3} },
    '型態成形中 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:63, moonshotN:9, pct:14.3, avg:43.1} },
    '型態成形中 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:189, moonshotN:27, pct:14.3, avg:48.1} },
    '型態突破確認 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:63, moonshotN:9, pct:14.3, avg:43.1} },
    '型態突破確認 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:189, moonshotN:27, pct:14.3, avg:48.1} },
    '型態剛形成(剛突破) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:63, moonshotN:9, pct:14.3, avg:43.1} },
    '型態剛形成(剛突破) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:189, moonshotN:27, pct:14.3, avg:48.1} },
    '近3月乖離度為正(營收優於股價) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:35, moonshotN:5, pct:14.3, avg:45.1} },
    '近3月均價YoY為負 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:35, moonshotN:5, pct:14.3, avg:48.2} },
    '突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:21, moonshotN:3, pct:14.3, avg:46.5} },
    '近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:71, moonshotN:10, pct:14.1, avg:52.5} },
    '多方力道≥80 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成': { 20: {n:157, moonshotN:22, pct:14.0, avg:48.6} },
    '多方力道≥80 ＋ 爆量(≥2倍均量) ＋ 晨星剛形成': { 20: {n:93, moonshotN:13, pct:14.0, avg:62.2} },
    '大盤站上60日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:164, moonshotN:23, pct:14.0, avg:49.1} },
    '大盤跌破60日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:179, moonshotN:25, pct:14.0, avg:51.5} },
    '強勢突破盤 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:36, moonshotN:5, pct:13.9, avg:42.8} },
    '相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 晨星剛形成': { 20: {n:108, moonshotN:15, pct:13.9, avg:56.4} },
    '大盤跌破20日均線 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:294, moonshotN:41, pct:13.9, avg:51.0} },
    '多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:58, moonshotN:8, pct:13.8, avg:57.3} },
    '多方力道≥80 ＋ 強勢突破盤 ＋ 晨星剛形成': { 20: {n:174, moonshotN:24, pct:13.8, avg:53.4} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 型態剛形成(剛突破)': { 20: {n:116, moonshotN:16, pct:13.8, avg:61.4} },
    '強勢突破盤 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:195, moonshotN:27, pct:13.8, avg:49.6} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:94, moonshotN:13, pct:13.8, avg:56.8} },
    '近3月乖離度為正(營收優於股價) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:87, moonshotN:12, pct:13.8, avg:56.9} },
    '強勢突破盤 ＋ 相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量)': { 20: {n:527, moonshotN:72, pct:13.7, avg:56.6} },
    '相對強弱為正(強於大盤) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:358, moonshotN:49, pct:13.7, avg:53.1} },
    '地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ N字底剛形成': { 20: {n:51, moonshotN:7, pct:13.7, avg:62.4} },
    '近3月均價YoY為正 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 20: {n:51, moonshotN:7, pct:13.7, avg:44.9} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過 ＋ 晨星剛形成': { 10: {n:22, moonshotN:3, pct:13.6, avg:35.6}, 20: {n:22, moonshotN:3, pct:13.6, avg:33.5} },
    '多方力道≥65 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 20: {n:191, moonshotN:26, pct:13.6, avg:63.1} },
    '多方力道≥65 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:346, moonshotN:47, pct:13.6, avg:54.9} },
    '多方力道≥65 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:118, moonshotN:16, pct:13.6, avg:51.5} },
    'MACD近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成': { 20: {n:59, moonshotN:8, pct:13.6, avg:44.9} },
    '相對強弱為負(弱於大盤) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:22, moonshotN:3, pct:13.6, avg:39.3} },
    '爆量(≥2倍均量) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:44, moonshotN:6, pct:13.6, avg:50.5} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:59, moonshotN:8, pct:13.6, avg:44.0} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:118, moonshotN:16, pct:13.6, avg:49.3} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:184, moonshotN:25, pct:13.6, avg:49.0} },
    '回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:22, moonshotN:3, pct:13.6, avg:42.5} },
    '量能區間高檔(≥90百分位) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:37, moonshotN:5, pct:13.5, avg:37.4} },
    '多方力道≥65 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:465, moonshotN:63, pct:13.5, avg:51.4} },
    '相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 20: {n:222, moonshotN:30, pct:13.5, avg:60.9} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成': { 20: {n:104, moonshotN:14, pct:13.5, avg:51.0} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收)': { 20: {n:357, moonshotN:48, pct:13.4, avg:58.0} },
    '多方力道≥80 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:30, moonshotN:4, pct:13.3, avg:41.5}, 20: {n:30, moonshotN:4, pct:13.3, avg:56.4} },
    '近3月均價YoY為負 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:30, moonshotN:4, pct:13.3, avg:37.0} },
    '多方力道≥65 ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:540, moonshotN:72, pct:13.3, avg:52.4} },
    '多方力道≥80 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:83, moonshotN:11, pct:13.3, avg:60.3} },
    '強勢突破盤 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成': { 20: {n:180, moonshotN:24, pct:13.3, avg:50.7} },
    '強勢突破盤 ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 20: {n:128, moonshotN:17, pct:13.3, avg:49.5} },
    '布林通道高檔(≥80%) ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 20: {n:248, moonshotN:33, pct:13.3, avg:48.4} },
    '相對強弱為正(強於大盤) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 20: {n:60, moonshotN:8, pct:13.3, avg:43.4} },
    '爆量(≥1.5倍均量) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:75, moonshotN:10, pct:13.3, avg:45.0} },
    '爆量(≥2倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:30, moonshotN:4, pct:13.3, avg:44.0} },
    '地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過': { 20: {n:188, moonshotN:25, pct:13.3, avg:59.9} },
    '地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 回後買上漲全通過': { 20: {n:249, moonshotN:33, pct:13.3, avg:59.6} },
    '地量(≤0.5倍均量) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成': { 20: {n:98, moonshotN:13, pct:13.3, avg:52.7} },
    '突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:151, moonshotN:20, pct:13.2, avg:46.1} },
    '多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成': { 20: {n:295, moonshotN:39, pct:13.2, avg:51.0} },
    '強勢突破盤 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:76, moonshotN:10, pct:13.2, avg:46.5} },
    '強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:114, moonshotN:15, pct:13.2, avg:53.6} },
    '布林通道高檔(≥80%) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:378, moonshotN:50, pct:13.2, avg:49.2} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:53, moonshotN:7, pct:13.2, avg:48.1} },
    '大盤跌破20日均線 ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:394, moonshotN:52, pct:13.2, avg:52.1} },
    '型態成形中 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:151, moonshotN:20, pct:13.2, avg:46.1} },
    '型態突破確認 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:151, moonshotN:20, pct:13.2, avg:46.1} },
    '型態剛形成(剛突破) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:151, moonshotN:20, pct:13.2, avg:46.1} },
    '地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 夜星剛形成': { 20: {n:122, moonshotN:16, pct:13.1, avg:54.6} },
    '量能區間高檔(≥90百分位) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:183, moonshotN:24, pct:13.1, avg:50.1} },
    '多方力道≥65 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:23, moonshotN:3, pct:13.0, avg:54.8} },
    '多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成': { 20: {n:77, moonshotN:10, pct:13.0, avg:54.0} },
    '多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: {n:92, moonshotN:12, pct:13.0, avg:74.7} },
    '強勢突破盤 ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:315, moonshotN:41, pct:13.0, avg:48.7} },
    '布林通道低檔(≤20%) ＋ MACD近3日內黃金交叉 ＋ 夜星剛形成': { 20: {n:23, moonshotN:3, pct:13.0, avg:50.3} },
    'KDJ近3日內死亡交叉 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:23, moonshotN:3, pct:13.0, avg:54.8} },
    'KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 突破ABC修正下降切線剛形成': { 20: {n:46, moonshotN:6, pct:13.0, avg:45.2} },
    '相對強弱為負(弱於大盤) ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: {n:77, moonshotN:10, pct:13.0, avg:44.7} },
    '爆量(≥2倍均量) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:100, moonshotN:13, pct:13.0, avg:50.2} },
    '近3月均價YoY為負 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:46, moonshotN:6, pct:13.0, avg:42.9} },
    '圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:23, moonshotN:3, pct:13.0, avg:38.5} },
    '多方力道≥65 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:426, moonshotN:55, pct:12.9, avg:53.3} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 型態突破確認': { 20: {n:357, moonshotN:46, pct:12.9, avg:57.7} },
    'MACD近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成': { 20: {n:70, moonshotN:9, pct:12.9, avg:44.2} },
    'MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 夜星剛形成': { 20: {n:31, moonshotN:4, pct:12.9, avg:56.3} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:232, moonshotN:30, pct:12.9, avg:52.2} },
    '大盤站上20日均線 ＋ N字底剛形成 ＋ 晨星剛形成': { 10: {n:39, moonshotN:5, pct:12.8, avg:39.6}, 20: {n:39, moonshotN:4, pct:10.3, avg:57.4} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:352, moonshotN:45, pct:12.8, avg:50.5} },
    'MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 20: {n:78, moonshotN:10, pct:12.8, avg:49.3} },
    'KDJ近3日內黃金交叉 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:86, moonshotN:11, pct:12.8, avg:51.9} },
    '爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:78, moonshotN:10, pct:12.8, avg:46.2} },
    '爆量(≥1.5倍均量) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:156, moonshotN:20, pct:12.8, avg:45.8} },
    '地量(≤0.5倍均量) ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:117, moonshotN:15, pct:12.8, avg:57.4} },
    '地量(≤0.5倍均量) ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:109, moonshotN:14, pct:12.8, avg:54.9} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:352, moonshotN:45, pct:12.8, avg:50.5} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:352, moonshotN:45, pct:12.8, avg:50.5} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:352, moonshotN:45, pct:12.8, avg:50.5} },
    '大盤站上20日均線 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:39, moonshotN:5, pct:12.8, avg:54.3} },
    '近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:39, moonshotN:5, pct:12.8, avg:40.8} },
    '多方力道≥65 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:102, moonshotN:13, pct:12.7, avg:59.3} },
    '布林通道高檔(≥80%) ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成': { 20: {n:126, moonshotN:16, pct:12.7, avg:49.1} },
    '布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 突破ABC修正下降切線剛形成': { 20: {n:55, moonshotN:7, pct:12.7, avg:67.6} },
    '布林通道高檔(≥80%) ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:472, moonshotN:60, pct:12.7, avg:49.9} },
    '量能區間高檔(≥90百分位) ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:71, moonshotN:9, pct:12.7, avg:47.2} },
    '多方力道≥65 ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成': { 20: {n:635, moonshotN:80, pct:12.6, avg:50.8} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 三大法人近3月買超': { 20: {n:475, moonshotN:60, pct:12.6, avg:57.6} },
    '地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 近3月均價YoY為正': { 20: {n:215, moonshotN:27, pct:12.6, avg:61.0} },
    '地量(≤0.5倍均量) ＋ 回後買上漲全通過 ＋ 三大法人近3月買超': { 20: {n:223, moonshotN:28, pct:12.6, avg:57.6} },
    '三大法人近3月買超 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:119, moonshotN:15, pct:12.6, avg:43.8} },
    '相對強弱為負(弱於大盤) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:24, moonshotN:3, pct:12.5, avg:36.4} },
    '大盤站上60日均線 ＋ N字底剛形成 ＋ 晨星剛形成': { 10: {n:40, moonshotN:5, pct:12.5, avg:39.6}, 20: {n:40, moonshotN:4, pct:10.0, avg:57.4} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為負 ＋ 夜星剛形成': { 10: {n:48, moonshotN:5, pct:10.4, avg:44.9}, 20: {n:48, moonshotN:6, pct:12.5, avg:46.4} },
    '多方力道≥65 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:48, moonshotN:6, pct:12.5, avg:50.1} },
    '多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量)': { 20: {n:1734, moonshotN:216, pct:12.5, avg:54.5} },
    '多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過': { 20: {n:439, moonshotN:55, pct:12.5, avg:56.4} },
    '多方力道≥80 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:72, moonshotN:9, pct:12.5, avg:57.9} },
    '強勢突破盤 ＋ MACD近3日內黃金交叉 ＋ 晨星剛形成': { 20: {n:40, moonshotN:5, pct:12.5, avg:42.0} },
    'KDJ近3日內黃金交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:208, moonshotN:26, pct:12.5, avg:44.8} },
    '相對強弱為正(強於大盤) ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:513, moonshotN:64, pct:12.5, avg:55.8} },
    '地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 回後買上漲全通過': { 20: {n:255, moonshotN:32, pct:12.5, avg:53.2} },
    '大盤站上20日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:598, moonshotN:75, pct:12.5, avg:51.8} },
    '大盤站上20日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:72, moonshotN:9, pct:12.5, avg:46.3} },
    '大盤站上60日均線 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:40, moonshotN:5, pct:12.5, avg:54.3} },
    '布林通道高檔(≥80%) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:643, moonshotN:80, pct:12.4, avg:48.7} },
    'MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 晨星剛形成': { 20: {n:121, moonshotN:15, pct:12.4, avg:45.7} },
    '相對強弱為正(強於大盤) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:734, moonshotN:91, pct:12.4, avg:52.2} },
    '爆量(≥2倍均量) ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:129, moonshotN:16, pct:12.4, avg:50.7} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成': { 20: {n:137, moonshotN:17, pct:12.4, avg:56.9} },
    '多方力道≥80 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 10: {n:122, moonshotN:13, pct:10.7, avg:43.2}, 20: {n:122, moonshotN:15, pct:12.3, avg:58.7} },
    '地量(≤0.5倍均量) ＋ 夜星剛形成': { 20: {n:162, moonshotN:20, pct:12.3, avg:55.6} },
    '多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:65, moonshotN:8, pct:12.3, avg:64.1} },
    '多方力道≥80 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:65, moonshotN:8, pct:12.3, avg:68.2} },
    'KDJ近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 20: {n:57, moonshotN:7, pct:12.3, avg:41.3} },
    '相對強弱為負(弱於大盤) ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成': { 20: {n:114, moonshotN:14, pct:12.3, avg:56.6} },
    '地量(≤0.5倍均量) ＋ 型態成形中 ＋ 夜星剛形成': { 20: {n:162, moonshotN:20, pct:12.3, avg:55.6} },
    '地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 夜星剛形成': { 20: {n:162, moonshotN:20, pct:12.3, avg:55.6} },
    '地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成': { 20: {n:162, moonshotN:20, pct:12.3, avg:55.6} },
    '大盤站上20日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:751, moonshotN:92, pct:12.3, avg:52.6} },
    '大盤站上20日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:65, moonshotN:8, pct:12.3, avg:47.1} },
    '回後買上漲全通過 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 20: {n:81, moonshotN:10, pct:12.3, avg:53.4} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 20: {n:180, moonshotN:22, pct:12.2, avg:53.1} },
    'K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:82, moonshotN:10, pct:12.2, avg:45.0} },
    '多方力道≥65 ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:599, moonshotN:73, pct:12.2, avg:52.2} },
    '多方力道≥80 ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 20: {n:74, moonshotN:9, pct:12.2, avg:42.2} },
    '強勢突破盤 ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成': { 20: {n:360, moonshotN:44, pct:12.2, avg:48.8} },
    '布林通道高檔(≥80%) ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成': { 20: {n:353, moonshotN:43, pct:12.2, avg:51.5} },
    '布林通道高檔(≥80%) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成': { 20: {n:74, moonshotN:9, pct:12.2, avg:54.8} },
    '布林通道高檔(≥80%) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:607, moonshotN:74, pct:12.2, avg:48.7} },
    'KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:90, moonshotN:11, pct:12.2, avg:51.3} },
    'KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成': { 20: {n:205, moonshotN:25, pct:12.2, avg:51.6} },
    '爆量(≥1.5倍均量) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:41, moonshotN:5, pct:12.2, avg:44.6} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態成形中 ＋ 夜星剛形成': { 20: {n:180, moonshotN:22, pct:12.2, avg:53.1} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態突破確認 ＋ 夜星剛形成': { 20: {n:180, moonshotN:22, pct:12.2, avg:53.1} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成': { 20: {n:180, moonshotN:22, pct:12.2, avg:53.1} },
    '型態成形中 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:82, moonshotN:10, pct:12.2, avg:45.0} },
    '型態突破確認 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:82, moonshotN:10, pct:12.2, avg:45.0} },
    '型態剛形成(剛突破) ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:82, moonshotN:10, pct:12.2, avg:45.0} },
    '回後買上漲全通過 ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成': { 20: {n:90, moonshotN:11, pct:12.2, avg:48.4} },
    '回後買上漲全通過 ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:49, moonshotN:6, pct:12.2, avg:50.7} },
    '三大法人近3月買超 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:49, moonshotN:6, pct:12.2, avg:44.0} },
    'KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量) ＋ 夜星剛形成': { 10: {n:66, moonshotN:8, pct:12.1, avg:35.4} },
    '量能區間高檔(≥90百分位) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 10: {n:33, moonshotN:4, pct:12.1, avg:35.9} },
    '多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成': { 20: {n:487, moonshotN:59, pct:12.1, avg:49.2} },
    '多方力道≥80 ＋ 強勢突破盤 ＋ 量能區間低檔(≤10百分位)': { 20: {n:33, moonshotN:4, pct:12.1, avg:53.3} },
    '多方力道≥80 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:190, moonshotN:23, pct:12.1, avg:54.6} },
    '強勢突破盤 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:224, moonshotN:27, pct:12.1, avg:50.4} },
    '強勢突破盤 ＋ 三重底剛形成 ＋ 晨星剛形成': { 20: {n:33, moonshotN:4, pct:12.1, avg:54.4} },
    '布林通道高檔(≥80%) ＋ 相對強弱為正(強於大盤) ＋ 晨星剛形成': { 20: {n:638, moonshotN:77, pct:12.1, avg:48.7} },
    '大盤跌破20日均線 ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:256, moonshotN:31, pct:12.1, avg:49.7} },
    '大盤站上60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:727, moonshotN:88, pct:12.1, avg:50.7} },
    '大盤站上60日均線 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:66, moonshotN:8, pct:12.1, avg:47.4} },
    '近3月均價YoY為正 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:99, moonshotN:12, pct:12.1, avg:50.0} },
    '三大法人近3月買超 ＋ 頭肩底剛形成 ＋ 圓弧底剛形成': { 20: {n:33, moonshotN:4, pct:12.1, avg:53.6} },
    '三大法人近3月賣超 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:33, moonshotN:4, pct:12.1, avg:46.7} },
    '多方力道≥80 ＋ N字底剛形成 ＋ 晨星剛形成': { 10: {n:25, moonshotN:3, pct:12.0, avg:42.2} },
    '大盤站上20日均線 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:25, moonshotN:3, pct:12.0, avg:36.2}, 20: {n:25, moonshotN:3, pct:12.0, avg:51.0} },
    '回後買上漲全通過 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:25, moonshotN:3, pct:12.0, avg:36.2}, 20: {n:25, moonshotN:3, pct:12.0, avg:51.0} },
    '強勢突破盤 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成': { 20: {n:25, moonshotN:3, pct:12.0, avg:42.2} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 大盤站上20日均線': { 20: {n:566, moonshotN:68, pct:12.0, avg:55.9} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為正': { 20: {n:460, moonshotN:55, pct:12.0, avg:56.8} },
    '強勢突破盤 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:108, moonshotN:13, pct:12.0, avg:44.6} },
    'KDJ近3日內黃金交叉 ＋ 相對強弱為正(強於大盤) ＋ 夜星剛形成': { 20: {n:217, moonshotN:26, pct:12.0, avg:45.9} },
    'KDJ近3日內死亡交叉 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:25, moonshotN:3, pct:12.0, avg:42.8} },
    'KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:150, moonshotN:18, pct:12.0, avg:47.4} },
    '爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:133, moonshotN:16, pct:12.0, avg:46.1} },
    '地量(≤0.5倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:117, moonshotN:14, pct:12.0, avg:59.7} },
    '量能區間高檔(≥90百分位) ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:274, moonshotN:33, pct:12.0, avg:53.2} },
    '大盤跌破60日均線 ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:225, moonshotN:27, pct:12.0, avg:49.9} },
    '回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:25, moonshotN:3, pct:12.0, avg:54.7} },
    '回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 晨星剛形成': { 20: {n:75, moonshotN:9, pct:12.0, avg:61.3} },
    '回後買上漲全通過 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:108, moonshotN:13, pct:12.0, avg:51.8} },
    '近3月均價YoY為正 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:532, moonshotN:64, pct:12.0, avg:50.5} },
    '多方力道≥65 ＋ 晨星剛形成': { 20: {n:707, moonshotN:84, pct:11.9, avg:51.2} },
    '布林通道高檔(≥80%) ＋ 晨星剛形成': { 20: {n:746, moonshotN:89, pct:11.9, avg:48.5} },
    '大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:578, moonshotN:69, pct:11.9, avg:49.8} },
    '多方力道≥65 ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:707, moonshotN:84, pct:11.9, avg:51.2} },
    '多方力道≥65 ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:707, moonshotN:84, pct:11.9, avg:51.2} },
    '多方力道≥65 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:707, moonshotN:84, pct:11.9, avg:51.2} },
    '多方力道≥80 ＋ MACD近3日內黃金交叉 ＋ N字底剛形成': { 20: {n:84, moonshotN:10, pct:11.9, avg:45.3} },
    '強勢突破盤 ＋ KDJ近3日內死亡交叉 ＋ 地量(≤0.5倍均量)': { 20: {n:109, moonshotN:13, pct:11.9, avg:54.3} },
    '強勢突破盤 ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 20: {n:59, moonshotN:7, pct:11.9, avg:44.6} },
    '布林通道高檔(≥80%) ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:746, moonshotN:89, pct:11.9, avg:48.5} },
    '布林通道高檔(≥80%) ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:746, moonshotN:89, pct:11.9, avg:48.5} },
    '布林通道高檔(≥80%) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:746, moonshotN:89, pct:11.9, avg:48.5} },
    'KDJ近3日內死亡交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:67, moonshotN:8, pct:11.9, avg:50.8} },
    '大盤跌破20日均線 ＋ 型態成形中 ＋ 夜星剛形成': { 20: {n:578, moonshotN:69, pct:11.9, avg:49.8} },
    '大盤跌破20日均線 ＋ 型態突破確認 ＋ 夜星剛形成': { 20: {n:578, moonshotN:69, pct:11.9, avg:49.8} },
    '大盤跌破20日均線 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成': { 20: {n:578, moonshotN:69, pct:11.9, avg:49.8} },
    '近3月均價YoY為正 ＋ N字底剛形成 ＋ 晨星剛形成': { 10: {n:34, moonshotN:4, pct:11.8, avg:39.2} },
    '圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:51, moonshotN:6, pct:11.8, avg:51.8} },
    '多方力道≥65 ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:220, moonshotN:26, pct:11.8, avg:53.0} },
    '布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位)': { 20: {n:186, moonshotN:22, pct:11.8, avg:49.3} },
    '布林通道高檔(≥80%) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:372, moonshotN:44, pct:11.8, avg:49.2} },
    'MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:85, moonshotN:10, pct:11.8, avg:51.5} },
    'KDJ近3日內黃金交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:161, moonshotN:19, pct:11.8, avg:53.0} },
    'KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:34, moonshotN:4, pct:11.8, avg:50.0} },
    '相對強弱為負(弱於大盤) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:127, moonshotN:15, pct:11.8, avg:43.0} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成': { 20: {n:169, moonshotN:20, pct:11.8, avg:56.6} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 圓弧底剛形成 ＋ K線橫盤的突破剛形成': { 20: {n:34, moonshotN:4, pct:11.8, avg:71.8} },
    '大盤跌破20日均線 ＋ 大盤站上60日均線 ＋ 夜星剛形成': { 20: {n:322, moonshotN:38, pct:11.8, avg:49.9} },
    '大盤站上60日均線 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:905, moonshotN:107, pct:11.8, avg:52.0} },
    '大盤跌破60日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:85, moonshotN:10, pct:11.8, avg:42.8} },
    '型態成形中 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:51, moonshotN:6, pct:11.8, avg:51.8} },
    '型態突破確認 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:51, moonshotN:6, pct:11.8, avg:51.8} },
    '型態剛形成(剛突破) ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:51, moonshotN:6, pct:11.8, avg:51.8} },
    '近3月乖離度為正(營收優於股價) ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:76, moonshotN:9, pct:11.8, avg:44.7} },
    '多方力道≥65 ＋ 爆量(≥2倍均量) ＋ 晨星剛形成': { 20: {n:197, moonshotN:23, pct:11.7, avg:55.5} },
    '多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成': { 20: {n:145, moonshotN:17, pct:11.7, avg:57.4} },
    '多方力道≥80 ＋ 量能區間低檔(≤10百分位) ＋ 近3月乖離度為負(股價超前營收)': { 20: {n:315, moonshotN:37, pct:11.7, avg:47.8} },
    '多方力道≥80 ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 20: {n:163, moonshotN:19, pct:11.7, avg:52.0} },
    '布林通道高檔(≥80%) ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 20: {n:315, moonshotN:37, pct:11.7, avg:53.0} },
    '布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:223, moonshotN:26, pct:11.7, avg:47.1} },
    '地量(≤0.5倍均量) ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成': { 20: {n:60, moonshotN:7, pct:11.7, avg:60.9} },
    '地量(≤0.5倍均量) ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:128, moonshotN:15, pct:11.7, avg:52.8} },
    '大盤站上20日均線 ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:1023, moonshotN:120, pct:11.7, avg:51.4} },
    '近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:282, moonshotN:33, pct:11.7, avg:50.3} },
    '地量(≤0.5倍均量) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 10: {n:43, moonshotN:5, pct:11.6, avg:39.1}, 20: {n:43, moonshotN:5, pct:11.6, avg:83.9} },
    'KDJ近3日內死亡交叉 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 夜星剛形成': { 10: {n:69, moonshotN:7, pct:10.1, avg:36.5}, 20: {n:69, moonshotN:8, pct:11.6, avg:48.2} },
    '大盤站上20日均線 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:69, moonshotN:7, pct:10.1, avg:38.5}, 20: {n:69, moonshotN:8, pct:11.6, avg:37.6} },
    '地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 20: {n:301, moonshotN:35, pct:11.6, avg:58.2} },
    '多方力道≥65 ＋ 強勢突破盤 ＋ 晨星剛形成': { 20: {n:292, moonshotN:34, pct:11.6, avg:49.1} },
    '多方力道≥65 ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:121, moonshotN:14, pct:11.6, avg:54.3} },
    '多方力道≥65 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:207, moonshotN:24, pct:11.6, avg:55.0} },
    '多方力道≥80 ＋ 回後買上漲全通過 ＋ N字底剛形成': { 20: {n:638, moonshotN:74, pct:11.6, avg:48.5} },
    '布林通道高檔(≥80%) ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成': { 20: {n:371, moonshotN:43, pct:11.6, avg:51.5} },
    'KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量) ＋ 回後買上漲全通過': { 20: {n:172, moonshotN:20, pct:11.6, avg:66.2} },
    '相對強弱為正(強於大盤) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:490, moonshotN:57, pct:11.6, avg:53.2} },
    '地量(≤0.5倍均量) ＋ 型態成形中 ＋ 回後買上漲全通過': { 20: {n:301, moonshotN:35, pct:11.6, avg:58.2} },
    '地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 回後買上漲全通過': { 20: {n:301, moonshotN:35, pct:11.6, avg:58.2} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 20: {n:112, moonshotN:13, pct:11.6, avg:49.1} },
    '回後買上漲全通過 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:155, moonshotN:18, pct:11.6, avg:57.3} },
    '強勢突破盤 ＋ 晨星剛形成': { 20: {n:391, moonshotN:45, pct:11.5, avg:48.7} },
    '多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成': { 20: {n:391, moonshotN:45, pct:11.5, avg:52.5} },
    '強勢突破盤 ＋ 布林通道高檔(≥80%) ＋ 晨星剛形成': { 20: {n:391, moonshotN:45, pct:11.5, avg:48.7} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 大盤站上60日均線': { 20: {n:583, moonshotN:67, pct:11.5, avg:56.9} },
    '強勢突破盤 ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:391, moonshotN:45, pct:11.5, avg:48.7} },
    '強勢突破盤 ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:391, moonshotN:45, pct:11.5, avg:48.7} },
    '強勢突破盤 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:391, moonshotN:45, pct:11.5, avg:48.7} },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成': { 20: {n:61, moonshotN:7, pct:11.5, avg:46.0} },
    'MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: {n:78, moonshotN:9, pct:11.5, avg:45.8} },
    'MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:61, moonshotN:7, pct:11.5, avg:54.7} },
    'KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:113, moonshotN:13, pct:11.5, avg:62.7} },
    '相對強弱為正(強於大盤) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:855, moonshotN:98, pct:11.5, avg:51.8} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:130, moonshotN:15, pct:11.5, avg:56.6} },
    '大盤站上20日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:695, moonshotN:80, pct:11.5, avg:51.2} },
    '大盤站上20日均線 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:122, moonshotN:14, pct:11.5, avg:59.7} },
    '近3月乖離度為負(股價超前營收) ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 20: {n:104, moonshotN:12, pct:11.5, avg:50.5} },
    '三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成': { 20: {n:209, moonshotN:24, pct:11.5, avg:51.7} },
    '三重底剛形成 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:61, moonshotN:7, pct:11.5, avg:55.0} },
    '多方力道≥65 ＋ N字底剛形成 ＋ 晨星剛形成': { 10: {n:35, moonshotN:4, pct:11.4, avg:41.9} },
    'KDJ近3日內黃金交叉 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:44, moonshotN:5, pct:11.4, avg:38.7} },
    '地量(≤0.5倍均量) ＋ 晨星剛形成': { 20: {n:175, moonshotN:20, pct:11.4, avg:55.4} },
    '大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:1118, moonshotN:128, pct:11.4, avg:51.0} },
    '多方力道≥65 ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成': { 20: {n:44, moonshotN:5, pct:11.4, avg:51.8} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 量能區間高檔(≥90百分位)': { 20: {n:35, moonshotN:4, pct:11.4, avg:54.6} },
    '強勢突破盤 ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:332, moonshotN:38, pct:11.4, avg:49.5} },
    '布林通道高檔(≥80%) ＋ MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成': { 20: {n:79, moonshotN:9, pct:11.4, avg:52.1} },
    '爆量(≥2倍均量) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:35, moonshotN:4, pct:11.4, avg:36.6} },
    '地量(≤0.5倍均量) ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:175, moonshotN:20, pct:11.4, avg:55.4} },
    '地量(≤0.5倍均量) ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:175, moonshotN:20, pct:11.4, avg:55.4} },
    '地量(≤0.5倍均量) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:175, moonshotN:20, pct:11.4, avg:55.4} },
    '地量(≤0.5倍均量) ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 20: {n:44, moonshotN:5, pct:11.4, avg:63.3} },
    '地量(≤0.5倍均量) ＋ 近3月均價YoY為負 ＋ 夜星剛形成': { 20: {n:44, moonshotN:5, pct:11.4, avg:50.0} },
    '量能區間高檔(≥90百分位) ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:44, moonshotN:5, pct:11.4, avg:44.6} },
    '大盤站上20日均線 ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:1118, moonshotN:128, pct:11.4, avg:51.0} },
    '大盤站上20日均線 ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:1118, moonshotN:128, pct:11.4, avg:51.0} },
    '大盤站上20日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:1118, moonshotN:128, pct:11.4, avg:51.0} },
    '大盤站上20日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 20: {n:421, moonshotN:48, pct:11.4, avg:50.6} },
    '大盤跌破60日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成': { 20: {n:167, moonshotN:19, pct:11.4, avg:51.6} },
    '近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成': { 20: {n:370, moonshotN:42, pct:11.4, avg:55.4} },
    '大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:345, moonshotN:39, pct:11.3, avg:50.0} },
    '多方力道≥80 ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 20: {n:177, moonshotN:20, pct:11.3, avg:48.7} },
    '多方力道≥80 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:238, moonshotN:27, pct:11.3, avg:59.5} },
    '強勢突破盤 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 20: {n:124, moonshotN:14, pct:11.3, avg:47.6} },
    '布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ N字底剛形成': { 20: {n:71, moonshotN:8, pct:11.3, avg:65.0} },
    '布林通道高檔(≥80%) ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:497, moonshotN:56, pct:11.3, avg:48.5} },
    'KDJ近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:213, moonshotN:24, pct:11.3, avg:43.4} },
    '相對強弱為正(強於大盤) ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成': { 20: {n:222, moonshotN:25, pct:11.3, avg:51.2} },
    '爆量(≥2倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:62, moonshotN:7, pct:11.3, avg:45.8} },
    '大盤站上60日均線 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:797, moonshotN:90, pct:11.3, avg:49.8} },
    '大盤跌破60日均線 ＋ 型態成形中 ＋ 夜星剛形成': { 20: {n:345, moonshotN:39, pct:11.3, avg:50.0} },
    '大盤跌破60日均線 ＋ 型態突破確認 ＋ 夜星剛形成': { 20: {n:345, moonshotN:39, pct:11.3, avg:50.0} },
    '大盤跌破60日均線 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成': { 20: {n:345, moonshotN:39, pct:11.3, avg:50.0} },
    '近3月均價YoY為負 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:160, moonshotN:18, pct:11.3, avg:52.8} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量)': { 20: {n:676, moonshotN:76, pct:11.2, avg:56.3} },
    '多方力道≥80 ＋ 強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)': { 20: {n:887, moonshotN:99, pct:11.2, avg:52.9} },
    '強勢突破盤 ＋ 布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量)': { 20: {n:676, moonshotN:76, pct:11.2, avg:56.3} },
    '強勢突破盤 ＋ KDJ近3日內黃金交叉 ＋ 地量(≤0.5倍均量)': { 20: {n:233, moonshotN:26, pct:11.2, avg:58.0} },
    '布林通道高檔(≥80%) ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:546, moonshotN:61, pct:11.2, avg:50.9} },
    'MACD近3日內黃金交叉 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:89, moonshotN:10, pct:11.2, avg:47.5} },
    '量能區間高檔(≥90百分位) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:116, moonshotN:13, pct:11.2, avg:50.6} },
    '大盤站上60日均線 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:143, moonshotN:16, pct:11.2, avg:59.1} },
    '大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:402, moonshotN:45, pct:11.2, avg:49.4} },
    '近3月均價YoY為負 ＋ 三重底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:143, moonshotN:16, pct:11.2, avg:54.9} },
    '三大法人近3月買超 ＋ N字底剛形成 ＋ 晨星剛形成': { 10: {n:36, moonshotN:4, pct:11.1, avg:39.2} },
    '多方力道≥65 ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:395, moonshotN:44, pct:11.1, avg:55.6} },
    '多方力道≥65 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:407, moonshotN:45, pct:11.1, avg:53.3} },
    '多方力道≥80 ＋ 地量(≤0.5倍均量) ＋ N字底剛形成': { 20: {n:108, moonshotN:12, pct:11.1, avg:56.3} },
    '多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成': { 20: {n:72, moonshotN:8, pct:11.1, avg:57.8} },
    '強勢突破盤 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:63, moonshotN:7, pct:11.1, avg:46.9} },
    '布林通道高檔(≥80%) ＋ 相對強弱為負(弱於大盤) ＋ 晨星剛形成': { 20: {n:108, moonshotN:12, pct:11.1, avg:47.0} },
    '布林通道高檔(≥80%) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:45, moonshotN:5, pct:11.1, avg:48.8} },
    '布林通道高檔(≥80%) ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:54, moonshotN:6, pct:11.1, avg:51.1} },
    'MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:108, moonshotN:12, pct:11.1, avg:48.2} },
    'MACD近3日內黃金交叉 ＋ 三重底剛形成 ＋ 晨星剛形成': { 20: {n:27, moonshotN:3, pct:11.1, avg:40.9} },
    '地量(≤0.5倍均量) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 20: {n:72, moonshotN:8, pct:11.1, avg:55.1} },
    '大盤跌破20日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成': { 20: {n:217, moonshotN:24, pct:11.1, avg:48.8} },
    '近3月均價YoY為正 ＋ 圓弧底剛形成 ＋ 晨星剛形成': { 20: {n:27, moonshotN:3, pct:11.1, avg:42.2} },
    '大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:1299, moonshotN:143, pct:11.0, avg:50.3} },
    '多方力道≥65 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成': { 20: {n:301, moonshotN:33, pct:11.0, avg:52.1} },
    '多方力道≥65 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:127, moonshotN:14, pct:11.0, avg:62.1} },
    '多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成': { 20: {n:127, moonshotN:14, pct:11.0, avg:57.9} },
    '強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過': { 20: {n:381, moonshotN:42, pct:11.0, avg:58.0} },
    '強勢突破盤 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:136, moonshotN:15, pct:11.0, avg:61.8} },
    'MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 20: {n:91, moonshotN:10, pct:11.0, avg:49.0} },
    'MACD近3日內黃金交叉 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:118, moonshotN:13, pct:11.0, avg:48.1} },
    '相對強弱為正(強於大盤) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:218, moonshotN:24, pct:11.0, avg:54.4} },
    '相對強弱為正(強於大盤) ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:748, moonshotN:82, pct:11.0, avg:51.0} },
    '爆量(≥2倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成': { 20: {n:173, moonshotN:19, pct:11.0, avg:51.5} },
    '爆量(≥2倍均量) ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:191, moonshotN:21, pct:11.0, avg:51.8} },
    '量能區間高檔(≥90百分位) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:309, moonshotN:34, pct:11.0, avg:53.6} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:400, moonshotN:44, pct:11.0, avg:48.7} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:100, moonshotN:11, pct:11.0, avg:49.8} },
    '大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:356, moonshotN:39, pct:11.0, avg:49.5} },
    '大盤站上60日均線 ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:1299, moonshotN:143, pct:11.0, avg:50.3} },
    '大盤站上60日均線 ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:1299, moonshotN:143, pct:11.0, avg:50.3} },
    '大盤站上60日均線 ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:1299, moonshotN:143, pct:11.0, avg:50.3} },
    '大盤跌破60日均線 ＋ 近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:282, moonshotN:31, pct:11.0, avg:51.7} },
    '大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成': { 20: {n:155, moonshotN:17, pct:11.0, avg:53.8} },
    '回後買上漲全通過 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:164, moonshotN:18, pct:11.0, avg:53.4} },
    '近3月均價YoY為正 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:100, moonshotN:11, pct:11.0, avg:65.0} },
    '多方力道≥65 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:338, moonshotN:37, pct:10.9, avg:52.1} },
    '多方力道≥80 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成': { 20: {n:442, moonshotN:48, pct:10.9, avg:55.9} },
    '多方力道≥80 ＋ 回後買上漲全通過 ＋ 複式頭肩底剛形成': { 20: {n:46, moonshotN:5, pct:10.9, avg:39.3} },
    '多方力道≥80 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 20: {n:201, moonshotN:22, pct:10.9, avg:62.1} },
    '強勢突破盤 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成': { 20: {n:221, moonshotN:24, pct:10.9, avg:46.3} },
    'MACD近3日內黃金交叉 ＋ 相對強弱為負(弱於大盤) ＋ 夜星剛形成': { 20: {n:137, moonshotN:15, pct:10.9, avg:49.6} },
    '相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收) ＋ 晨星剛形成': { 20: {n:606, moonshotN:66, pct:10.9, avg:51.8} },
    '爆量(≥2倍均量) ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成': { 20: {n:304, moonshotN:33, pct:10.9, avg:50.4} },
    '爆量(≥2倍均量) ＋ 近3月乖離度為正(營收優於股價) ＋ 母子懷抱(低檔)剛形成': { 20: {n:119, moonshotN:13, pct:10.9, avg:46.2} },
    '大盤跌破20日均線 ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成': { 20: {n:312, moonshotN:34, pct:10.9, avg:51.4} },
    '回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:46, moonshotN:5, pct:10.9, avg:54.8} },
    '近3月乖離度為正(營收優於股價) ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:202, moonshotN:22, pct:10.9, avg:49.4} },
    '三大法人近3月買超 ＋ 圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:201, moonshotN:22, pct:10.9, avg:54.2} },
    '三大法人近3月賣超 ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:46, moonshotN:5, pct:10.9, avg:37.0} },
    '布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量) ＋ 母子懷抱(高檔)剛形成': { 10: {n:37, moonshotN:4, pct:10.8, avg:38.0} },
    '爆量(≥2倍均量) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:37, moonshotN:4, pct:10.8, avg:36.1} },
    '近3月乖離度為正(營收優於股價) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:37, moonshotN:4, pct:10.8, avg:37.0} },
    '多方力道≥65 ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成': { 20: {n:269, moonshotN:29, pct:10.8, avg:48.3} },
    '多方力道≥80 ＋ KDJ近3日內死亡交叉 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:333, moonshotN:36, pct:10.8, avg:56.6} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上)': { 20: {n:295, moonshotN:32, pct:10.8, avg:57.6} },
    '布林通道高檔(≥80%) ＋ 大盤跌破20日均線 ＋ 晨星剛形成': { 20: {n:139, moonshotN:15, pct:10.8, avg:47.4} },
    '相對強弱為負(弱於大盤) ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:223, moonshotN:24, pct:10.8, avg:45.1} },
    '爆量(≥1.5倍均量) ＋ N字底剛形成 ＋ 晨星剛形成': { 20: {n:37, moonshotN:4, pct:10.8, avg:63.7} },
    '爆量(≥2倍均量) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:83, moonshotN:9, pct:10.8, avg:35.9} },
    '回後買上漲全通過 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:213, moonshotN:23, pct:10.8, avg:51.0} },
    '三大法人近3月買超 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:111, moonshotN:12, pct:10.8, avg:65.7} },
    '大盤站上60日均線 ＋ N字底剛形成 ＋ 母子懷抱(低檔)剛形成': { 10: {n:28, moonshotN:3, pct:10.7, avg:36.2}, 20: {n:28, moonshotN:3, pct:10.7, avg:51.0} },
    'N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 20: {n:84, moonshotN:9, pct:10.7, avg:42.1} },
    '多方力道≥80 ＋ 回後買上漲全通過 ＋ 三大法人近3月買超': { 20: {n:3195, moonshotN:342, pct:10.7, avg:51.3} },
    '強勢突破盤 ＋ 回後買上漲全通過 ＋ 晨星剛形成': { 20: {n:177, moonshotN:19, pct:10.7, avg:52.5} },
    '布林通道高檔(≥80%) ＋ KDJ近3日內死亡交叉 ＋ 晨星剛形成': { 20: {n:364, moonshotN:39, pct:10.7, avg:47.4} },
    'MACD近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:140, moonshotN:15, pct:10.7, avg:45.5} },
    'KDJ近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:663, moonshotN:71, pct:10.7, avg:51.3} },
    'KDJ近3日內死亡交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成': { 20: {n:196, moonshotN:21, pct:10.7, avg:46.4} },
    '爆量(≥2倍均量) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:308, moonshotN:33, pct:10.7, avg:49.6} },
    '量能區間高檔(≥90百分位) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:131, moonshotN:14, pct:10.7, avg:58.7} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上60日均線 ＋ 夜星剛形成': { 20: {n:140, moonshotN:15, pct:10.7, avg:49.4} },
    '型態成形中 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 20: {n:84, moonshotN:9, pct:10.7, avg:42.1} },
    '型態突破確認 ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 20: {n:84, moonshotN:9, pct:10.7, avg:42.1} },
    '型態剛形成(剛突破) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 20: {n:84, moonshotN:9, pct:10.7, avg:42.1} },
    '圓弧底剛形成 ＋ 突破飆股大量黑K最高點剛形成 ＋ K線橫盤的突破剛形成': { 20: {n:122, moonshotN:13, pct:10.7, avg:58.8} },
    '突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:56, moonshotN:6, pct:10.7, avg:47.2} },
    'MACD近3日內黃金交叉 ＋ 夜星剛形成': { 20: {n:218, moonshotN:23, pct:10.6, avg:49.7} },
    '三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:784, moonshotN:83, pct:10.6, avg:49.7} },
    '多方力道≥80 ＋ KDJ近3日內黃金交叉 ＋ N字底剛形成': { 20: {n:720, moonshotN:76, pct:10.6, avg:49.4} },
    '多方力道≥80 ＋ 量能區間高檔(≥90百分位) ＋ N字底剛形成': { 20: {n:718, moonshotN:76, pct:10.6, avg:47.4} },
    '強勢突破盤 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:263, moonshotN:28, pct:10.6, avg:48.2} },
    '布林通道高檔(≥80%) ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成': { 20: {n:489, moonshotN:52, pct:10.6, avg:50.0} },
    '布林通道高檔(≥80%) ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:405, moonshotN:43, pct:10.6, avg:50.5} },
    'MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 夜星剛形成': { 20: {n:218, moonshotN:23, pct:10.6, avg:49.7} },
    'MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 夜星剛形成': { 20: {n:218, moonshotN:23, pct:10.6, avg:49.7} },
    'MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 夜星剛形成': { 20: {n:218, moonshotN:23, pct:10.6, avg:49.7} },
    'KDJ近3日內黃金交叉 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:161, moonshotN:17, pct:10.6, avg:56.2} },
    '爆量(≥1.5倍均量) ＋ 突破ABC修正下降切線剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:47, moonshotN:5, pct:10.6, avg:45.2} },
    '爆量(≥2倍均量) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:47, moonshotN:5, pct:10.6, avg:38.0} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:151, moonshotN:16, pct:10.6, avg:44.3} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為正(營收優於股價) ＋ 晨星剛形成': { 20: {n:142, moonshotN:15, pct:10.6, avg:52.8} },
    '大盤站上60日均線 ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 20: {n:500, moonshotN:53, pct:10.6, avg:51.2} },
    '型態成形中 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:784, moonshotN:83, pct:10.6, avg:49.7} },
    '型態突破確認 ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:784, moonshotN:83, pct:10.6, avg:49.7} },
    '型態剛形成(剛突破) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:784, moonshotN:83, pct:10.6, avg:49.7} },
    '近3月均價YoY為正 ＋ K線橫盤的突破剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:47, moonshotN:5, pct:10.6, avg:41.8} },
    '多方力道≥65 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成': { 20: {n:315, moonshotN:33, pct:10.5, avg:54.4} },
    '多方力道≥65 ＋ 圓弧底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:38, moonshotN:4, pct:10.5, avg:38.9} },
    '多方力道≥80 ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成': { 20: {n:275, moonshotN:29, pct:10.5, avg:54.1} },
    '多方力道≥80 ＋ 大盤站上20日均線 ＋ 回後買上漲全通過': { 20: {n:3550, moonshotN:374, pct:10.5, avg:50.5} },
    '多方力道≥80 ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成': { 20: {n:143, moonshotN:15, pct:10.5, avg:59.1} },
    '多方力道≥80 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:295, moonshotN:31, pct:10.5, avg:59.6} },
    '強勢突破盤 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 突破ABC修正下降切線剛形成': { 20: {n:95, moonshotN:10, pct:10.5, avg:57.3} },
    'MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成': { 20: {n:105, moonshotN:11, pct:10.5, avg:49.8} },
    'MACD近3日內黃金交叉 ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成': { 20: {n:124, moonshotN:13, pct:10.5, avg:53.7} },
    'KDJ近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: {n:124, moonshotN:13, pct:10.5, avg:45.4} },
    'KDJ近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 夜星剛形成': { 20: {n:304, moonshotN:32, pct:10.5, avg:44.2} },
    '相對強弱為正(強於大盤) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 母子懷抱(高檔)剛形成': { 20: {n:181, moonshotN:19, pct:10.5, avg:58.5} },
    '爆量(≥1.5倍均量) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:153, moonshotN:16, pct:10.5, avg:57.9} },
    '爆量(≥2倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:276, moonshotN:29, pct:10.5, avg:50.5} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:440, moonshotN:46, pct:10.5, avg:48.7} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 大盤站上20日均線 ＋ 夜星剛形成': { 20: {n:86, moonshotN:9, pct:10.5, avg:47.7} },
    '大盤跌破60日均線 ＋ 近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:315, moonshotN:33, pct:10.5, avg:51.6} },
    '近3月乖離度為負(股價超前營收) ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:602, moonshotN:63, pct:10.5, avg:51.2} },
    '近3月均價YoY為負 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:76, moonshotN:8, pct:10.5, avg:47.1} },
    '布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成': { 20: {n:625, moonshotN:65, pct:10.4, avg:50.1} },
    '多方力道≥65 ＋ KDJ近3日內黃金交叉 ＋ 夜星剛形成': { 20: {n:164, moonshotN:17, pct:10.4, avg:46.9} },
    '多方力道≥65 ＋ 三大法人近3月買超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:560, moonshotN:58, pct:10.4, avg:54.4} },
    '多方力道≥80 ＋ 回後買上漲全通過 ＋ 近3月均價YoY為負': { 20: {n:1578, moonshotN:164, pct:10.4, avg:50.3} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上)': { 20: {n:125, moonshotN:13, pct:10.4, avg:58.9} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 型態成形中': { 20: {n:605, moonshotN:63, pct:10.4, avg:56.6} },
    '布林通道高檔(≥80%) ＋ 爆量(≥2倍均量) ＋ 晨星剛形成': { 20: {n:308, moonshotN:32, pct:10.4, avg:50.5} },
    '布林通道高檔(≥80%) ＋ 量能區間高檔(≥90百分位) ＋ 晨星剛形成': { 20: {n:491, moonshotN:51, pct:10.4, avg:50.8} },
    '布林通道高檔(≥80%) ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成': { 20: {n:625, moonshotN:65, pct:10.4, avg:50.1} },
    '布林通道高檔(≥80%) ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成': { 20: {n:625, moonshotN:65, pct:10.4, avg:50.1} },
    '布林通道高檔(≥80%) ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成': { 20: {n:625, moonshotN:65, pct:10.4, avg:50.1} },
    '布林通道高檔(≥80%) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:163, moonshotN:17, pct:10.4, avg:59.1} },
    'MACD近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(高檔)剛形成': { 20: {n:48, moonshotN:5, pct:10.4, avg:51.6} },
    'MACD近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成': { 20: {n:106, moonshotN:11, pct:10.4, avg:50.1} },
    'MACD近3日內黃金交叉 ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:135, moonshotN:14, pct:10.4, avg:51.6} },
    'KDJ近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 夜星剛形成': { 20: {n:307, moonshotN:32, pct:10.4, avg:43.7} },
    'KDJ近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 夜星剛形成': { 20: {n:135, moonshotN:14, pct:10.4, avg:44.8} },
    'KDJ近3日內黃金交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:550, moonshotN:57, pct:10.4, avg:52.5} },
    'KDJ近3日內死亡交叉 ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:347, moonshotN:36, pct:10.4, avg:49.6} },
    '相對強弱為正(強於大盤) ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:766, moonshotN:80, pct:10.4, avg:52.6} },
    '爆量(≥2倍均量) ＋ 近3月均價YoY為負 ＋ 母子懷抱(低檔)剛形成': { 20: {n:106, moonshotN:11, pct:10.4, avg:48.6} },
    '量能區間高檔(≥90百分位) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:536, moonshotN:56, pct:10.4, avg:51.5} },
    '大盤站上20日均線 ＋ 大盤跌破60日均線 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:240, moonshotN:25, pct:10.4, avg:49.1} },
    '回後買上漲全通過 ＋ 近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:651, moonshotN:68, pct:10.4, avg:53.6} },
    '近3月均價YoY為正 ＋ 三大法人近3月買超 ＋ 晨星剛形成': { 20: {n:814, moonshotN:85, pct:10.4, avg:49.6} },
    '相對強弱為正(強於大盤) ＋ 晨星剛形成': { 20: {n:1168, moonshotN:120, pct:10.3, avg:50.4} },
    '多方力道≥65 ＋ 近3月乖離度為負(股價超前營收) ＋ 母子懷抱(低檔)剛形成': { 20: {n:397, moonshotN:41, pct:10.3, avg:56.6} },
    '多方力道≥80 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:146, moonshotN:15, pct:10.3, avg:54.1} },
    '多方力道≥80 ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 20: {n:116, moonshotN:12, pct:10.3, avg:55.2} },
    '布林通道高檔(≥80%) ＋ 爆量(≥1.5倍均量) ＋ 晨星剛形成': { 20: {n:467, moonshotN:48, pct:10.3, avg:49.3} },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 夜星剛形成': { 20: {n:97, moonshotN:10, pct:10.3, avg:56.5} },
    'MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 頭肩底剛形成': { 20: {n:29, moonshotN:3, pct:10.3, avg:58.5} },
    'MACD近3日內黃金交叉 ＋ 爆量(≥1.5倍均量) ＋ 夜星剛形成': { 20: {n:68, moonshotN:7, pct:10.3, avg:53.4} },
    'MACD近3日內黃金交叉 ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:136, moonshotN:14, pct:10.3, avg:47.9} },
    '相對強弱為正(強於大盤) ＋ 型態成形中 ＋ 晨星剛形成': { 20: {n:1168, moonshotN:120, pct:10.3, avg:50.4} },
    '相對強弱為正(強於大盤) ＋ 型態突破確認 ＋ 晨星剛形成': { 20: {n:1168, moonshotN:120, pct:10.3, avg:50.4} },
    '相對強弱為正(強於大盤) ＋ 型態剛形成(剛突破) ＋ 晨星剛形成': { 20: {n:1168, moonshotN:120, pct:10.3, avg:50.4} },
    '相對強弱為正(強於大盤) ＋ 近3月乖離度為負(股價超前營收) ＋ 夜星剛形成': { 20: {n:418, moonshotN:43, pct:10.3, avg:57.0} },
    '爆量(≥1.5倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 圓弧底剛形成': { 20: {n:97, moonshotN:10, pct:10.3, avg:57.7} },
    '地量(≤0.5倍均量) ＋ 大盤站上20日均線 ＋ 夜星剛形成': { 20: {n:87, moonshotN:9, pct:10.3, avg:48.3} },
    '地量(≤0.5倍均量) ＋ 大盤跌破20日均線 ＋ 母子懷抱(高檔)剛形成': { 20: {n:203, moonshotN:21, pct:10.3, avg:50.1} },
    '地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 20: {n:58, moonshotN:6, pct:10.3, avg:45.3} },
    '量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 近3月乖離度為正(營收優於股價) ＋ 夜星剛形成': { 20: {n:68, moonshotN:7, pct:10.3, avg:59.1} },
    '大盤站上20日均線 ＋ 三重底剛形成 ＋ 母子懷抱(低檔)剛形成': { 20: {n:39, moonshotN:4, pct:10.3, avg:49.2} },
    '大盤跌破60日均線 ＋ 頭肩底剛形成 ＋ 複式頭肩底剛形成': { 20: {n:29, moonshotN:3, pct:10.3, avg:55.0} },
    '大盤跌破60日均線 ＋ 複式頭肩底剛形成 ＋ K線橫盤的突破剛形成': { 20: {n:29, moonshotN:3, pct:10.3, avg:46.4} },
    '大盤跌破60日均線 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成': { 20: {n:117, moonshotN:12, pct:10.3, avg:50.1} },
    'MACD近3日內黃金交叉 ＋ 母子懷抱(低檔)剛形成': { 20: {n:177, moonshotN:18, pct:10.2, avg:46.9} },
    'N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:166, moonshotN:17, pct:10.2, avg:59.1} },
    '多方力道≥65 ＋ 布林通道高檔(≥80%) ＋ 地量(≤0.5倍均量)': { 20: {n:2446, moonshotN:249, pct:10.2, avg:53.9} },
    '多方力道≥65 ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: {n:108, moonshotN:11, pct:10.2, avg:57.4} },
    '多方力道≥65 ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:264, moonshotN:27, pct:10.2, avg:50.4} },
    '多方力道≥65 ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 20: {n:108, moonshotN:11, pct:10.2, avg:44.2} },
    '強勢突破盤 ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成': { 20: {n:137, moonshotN:14, pct:10.2, avg:44.1} },
    '強勢突破盤 ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 20: {n:166, moonshotN:17, pct:10.2, avg:46.5} },
    'MACD近3日內黃金交叉 ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成': { 20: {n:177, moonshotN:18, pct:10.2, avg:46.9} },
    'MACD近3日內黃金交叉 ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成': { 20: {n:177, moonshotN:18, pct:10.2, avg:46.9} },
    'MACD近3日內黃金交叉 ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成': { 20: {n:177, moonshotN:18, pct:10.2, avg:46.9} },
    'KDJ近3日內黃金交叉 ＋ 量能區間高檔(≥90百分位) ＋ 母子懷抱(低檔)剛形成': { 20: {n:187, moonshotN:19, pct:10.2, avg:55.7} },
    '相對強弱為正(強於大盤) ＋ 三重底剛形成 ＋ 晨星剛形成': { 20: {n:49, moonshotN:5, pct:10.2, avg:51.4} },
    '相對強弱為負(弱於大盤) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:443, moonshotN:45, pct:10.2, avg:47.2} },
    '爆量(≥1.5倍均量) ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:462, moonshotN:47, pct:10.2, avg:50.4} },
    '爆量(≥2倍均量) ＋ 近3月均價YoY為正 ＋ 母子懷抱(低檔)剛形成': { 20: {n:127, moonshotN:13, pct:10.2, avg:51.6} },
    '爆量(≥2倍均量) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:118, moonshotN:12, pct:10.2, avg:57.0} },
    '量能區間高檔(≥90百分位) ＋ 近3月均價YoY為正 ＋ 晨星剛形成': { 20: {n:402, moonshotN:41, pct:10.2, avg:51.4} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 大盤跌破20日均線 ＋ 夜星剛形成': { 20: {n:206, moonshotN:21, pct:10.2, avg:50.5} },
    '型態成形中 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:166, moonshotN:17, pct:10.2, avg:59.1} },
    '型態突破確認 ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:166, moonshotN:17, pct:10.2, avg:59.1} },
    '型態剛形成(剛突破) ＋ N字底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:166, moonshotN:17, pct:10.2, avg:59.1} },
    '回後買上漲全通過 ＋ 近3月均價YoY為負 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:539, moonshotN:55, pct:10.2, avg:51.7} },
    '近3月均價YoY為正 ＋ 母子懷抱(高檔)剛形成 ＋ 夜星剛形成': { 20: {n:245, moonshotN:25, pct:10.2, avg:49.1} },
    '相對強弱為正(強於大盤) ＋ 夜星剛形成': { 20: {n:818, moonshotN:83, pct:10.1, avg:55.0} },
    '爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: {n:237, moonshotN:24, pct:10.1, avg:50.2} },
    '多方力道≥65 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過': { 20: {n:661, moonshotN:67, pct:10.1, avg:54.8} },
    '多方力道≥80 ＋ 布林通道高檔(≥80%) ＋ 母子懷抱(低檔)剛形成': { 20: {n:228, moonshotN:23, pct:10.1, avg:57.1} },
    '多方力道≥80 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成': { 20: {n:159, moonshotN:16, pct:10.1, avg:56.1} },
    '多方力道≥80 ＋ 相對強弱為正(強於大盤) ＋ 回後買上漲全通過': { 20: {n:4080, moonshotN:413, pct:10.1, avg:51.3} },
    '多方力道≥80 ＋ 爆量(≥1.5倍均量) ＋ N字底剛形成': { 20: {n:605, moonshotN:61, pct:10.1, avg:48.5} },
    '多方力道≥80 ＋ 大盤站上60日均線 ＋ 回後買上漲全通過': { 20: {n:3662, moonshotN:371, pct:10.1, avg:51.6} },
    '多方力道≥80 ＋ 回後買上漲全通過 ＋ 近3月乖離度為正(營收優於股價)': { 20: {n:2024, moonshotN:205, pct:10.1, avg:50.1} },
    '多方力道≥80 ＋ 近3月乖離度為正(營收優於股價) ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:672, moonshotN:68, pct:10.1, avg:52.9} },
    '強勢突破盤 ＋ 相對強弱為正(強於大盤) ＋ 母子懷抱(低檔)剛形成': { 20: {n:198, moonshotN:20, pct:10.1, avg:48.1} },
    '強勢突破盤 ＋ 地量(≤0.5倍均量) ＋ 近3月均價YoY為負': { 20: {n:208, moonshotN:21, pct:10.1, avg:55.0} },
    'MACD近3日內黃金交叉 ＋ 爆量(≥2倍均量) ＋ 晨星剛形成': { 20: {n:69, moonshotN:7, pct:10.1, avg:37.8} },
    'MACD近3日內黃金交叉 ＋ 三大法人近3月賣超 ＋ 夜星剛形成': { 20: {n:99, moonshotN:10, pct:10.1, avg:51.9} },
    '相對強弱為正(強於大盤) ＋ 型態成形中 ＋ 夜星剛形成': { 20: {n:818, moonshotN:83, pct:10.1, avg:55.0} },
    '相對強弱為正(強於大盤) ＋ 型態突破確認 ＋ 夜星剛形成': { 20: {n:818, moonshotN:83, pct:10.1, avg:55.0} },
    '相對強弱為正(強於大盤) ＋ 型態剛形成(剛突破) ＋ 夜星剛形成': { 20: {n:818, moonshotN:83, pct:10.1, avg:55.0} },
    '爆量(≥1.5倍均量) ＋ 爆量(≥2倍均量) ＋ 母子懷抱(低檔)剛形成': { 20: {n:237, moonshotN:24, pct:10.1, avg:50.2} },
    '爆量(≥1.5倍均量) ＋ 大盤站上60日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:358, moonshotN:36, pct:10.1, avg:49.2} },
    '爆量(≥1.5倍均量) ＋ 三大法人近3月賣超 ＋ 晨星剛形成': { 20: {n:296, moonshotN:30, pct:10.1, avg:50.5} },
    '爆量(≥2倍均量) ＋ 型態成形中 ＋ 母子懷抱(低檔)剛形成': { 20: {n:237, moonshotN:24, pct:10.1, avg:50.2} },
    '爆量(≥2倍均量) ＋ 型態突破確認 ＋ 母子懷抱(低檔)剛形成': { 20: {n:237, moonshotN:24, pct:10.1, avg:50.2} },
    '爆量(≥2倍均量) ＋ 型態剛形成(剛突破) ＋ 母子懷抱(低檔)剛形成': { 20: {n:237, moonshotN:24, pct:10.1, avg:50.2} },
    '量能區間高檔(≥90百分位) ＋ 量能斜率轉強(近5日均量>近10日均量20%以上) ＋ 晨星剛形成': { 20: {n:366, moonshotN:37, pct:10.1, avg:48.9} },
    '量能區間高檔(≥90百分位) ＋ 大盤站上60日均線 ＋ 晨星剛形成': { 20: {n:595, moonshotN:60, pct:10.1, avg:50.3} },
    '大盤跌破20日均線 ＋ 突破ABC修正下降切線剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:99, moonshotN:10, pct:10.1, avg:53.5} },
    '大盤跌破60日均線 ＋ 回後買上漲全通過 ＋ 複式頭肩底剛形成': { 20: {n:69, moonshotN:7, pct:10.1, avg:50.1} },
    '近3月乖離度為負(股價超前營收) ＋ 三大法人近3月買超 ＋ 夜星剛形成': { 20: {n:405, moonshotN:41, pct:10.1, avg:48.9} },
    '三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成 ＋ 晨星剛形成': { 20: {n:218, moonshotN:22, pct:10.1, avg:42.7} },
    '量能斜率轉強(近5日均量>近10日均量20%以上) ＋ N字底剛形成 ＋ 突破ABC修正下降切線剛形成': { 5: {n:30, moonshotN:3, pct:10.0, avg:34.0}, 10: {n:30, moonshotN:3, pct:10.0, avg:46.3} },
    '量能區間低檔(≤10百分位) ＋ 大盤跌破60日均線 ＋ 三重底剛形成': { 10: {n:20, moonshotN:2, pct:10.0, avg:35.1}, 20: {n:20, moonshotN:2, pct:10.0, avg:47.7} },
    '大盤站上20日均線 ＋ 複式頭肩底剛形成 ＋ 突破ABC修正下降切線剛形成': { 10: {n:20, moonshotN:2, pct:10.0, avg:34.0}, 20: {n:20, moonshotN:2, pct:10.0, avg:36.5} },
    '大盤站上60日均線 ＋ 頭肩底剛形成 ＋ 突破飆股大量黑K最高點剛形成': { 10: {n:20, moonshotN:2, pct:10.0, avg:53.1}, 20: {n:20, moonshotN:2, pct:10.0, avg:63.4} },
    '多方力道≥65 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(低檔)剛形成': { 20: {n:320, moonshotN:32, pct:10.0, avg:49.5} },
    '多方力道≥80 ＋ MACD近3日內黃金交叉 ＋ 量能區間低檔(≤10百分位)': { 20: {n:20, moonshotN:2, pct:10.0, avg:31.0} },
    '多方力道≥80 ＋ 大盤站上20日均線 ＋ N字底剛形成': { 20: {n:959, moonshotN:96, pct:10.0, avg:46.7} },
    '多方力道≥80 ＋ 三大法人近3月買超 ＋ N字底剛形成': { 20: {n:906, moonshotN:91, pct:10.0, avg:48.2} },
    '布林通道高檔(≥80%) ＋ 近3月均價YoY為負 ＋ 晨星剛形成': { 20: {n:269, moonshotN:27, pct:10.0, avg:46.3} },
    '布林通道高檔(≥80%) ＋ 三大法人近3月賣超 ＋ 母子懷抱(低檔)剛形成': { 20: {n:220, moonshotN:22, pct:10.0, avg:49.3} },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內黃金交叉 ＋ 晨星剛形成': { 20: {n:110, moonshotN:11, pct:10.0, avg:56.6} },
    'MACD近3日內黃金交叉 ＋ KDJ近3日內死亡交叉 ＋ 母子懷抱(高檔)剛形成': { 20: {n:80, moonshotN:8, pct:10.0, avg:44.3} },
    'MACD近3日內黃金交叉 ＋ 大盤跌破60日均線 ＋ 晨星剛形成': { 20: {n:40, moonshotN:4, pct:10.0, avg:44.0} },
    'MACD近3日內黃金交叉 ＋ 近3月均價YoY為負 ＋ 母子懷抱(低檔)剛形成': { 20: {n:70, moonshotN:7, pct:10.0, avg:41.7} },
    'KDJ近3日內死亡交叉 ＋ 大盤站上20日均線 ＋ 晨星剛形成': { 20: {n:400, moonshotN:40, pct:10.0, avg:47.1} },
    '爆量(≥1.5倍均量) ＋ 大盤站上20日均線 ＋ 母子懷抱(低檔)剛形成': { 20: {n:329, moonshotN:33, pct:10.0, avg:48.2} },
    '爆量(≥2倍均量) ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 回後買上漲全通過': { 20: {n:110, moonshotN:11, pct:10.0, avg:46.4} },
    '地量(≤0.5倍均量) ＋ 大盤跌破60日均線 ＋ 夜星剛形成': { 20: {n:40, moonshotN:4, pct:10.0, avg:59.3} },
    '地量(≤0.5倍均量) ＋ 三大法人近3月買超 ＋ 母子懷抱(高檔)剛形成': { 20: {n:249, moonshotN:25, pct:10.0, avg:55.0} },
    '量能區間高檔(≥90百分位) ＋ 頭肩底剛形成 ＋ 圓弧底剛形成': { 20: {n:40, moonshotN:4, pct:10.0, avg:53.6} },
    '大盤站上20日均線 ＋ 回後買上漲全通過 ＋ 突破飆股大量黑K最高點剛形成': { 20: {n:979, moonshotN:98, pct:10.0, avg:52.4} },
    '大盤站上20日均線 ＋ 回後買上漲全通過 ＋ 晨星剛形成': { 20: {n:270, moonshotN:27, pct:10.0, avg:50.6} },
    '大盤跌破20日均線 ＋ K線橫盤的突破剛形成 ＋ 晨星剛形成': { 20: {n:40, moonshotN:4, pct:10.0, avg:43.0} },
    '大盤站上60日均線 ＋ 回後買上漲全通過 ＋ 晨星剛形成': { 20: {n:279, moonshotN:28, pct:10.0, avg:51.7} },
    '回後買上漲全通過 ＋ N字底剛形成 ＋ 晨星剛形成': { 20: {n:50, moonshotN:5, pct:10.0, avg:57.7} },
    '近3月乖離度為正(營收優於股價) ＋ 突破ABC修正下降切線剛形成 ＋ 晨星剛形成': { 20: {n:60, moonshotN:6, pct:10.0, avg:43.4} },
    "多方力道≥80 ＋ 強勢突破盤 ＋ 地量(≤0.5倍均量)": {"10": {"n": 301, "moonshotN": 33, "pct": 11.0, "avg": 46.4}, "20": {"n": 301, "moonshotN": 63, "pct": 20.9, "avg": 55.5}},
    "多方力道≥80 ＋ 量能斜率轉弱(近5日均量<近10日均量20%以上) ＋ 晨星剛形成": {"10": {"n": 115, "moonshotN": 12, "pct": 10.4, "avg": 44.4}, "20": {"n": 115, "moonshotN": 24, "pct": 20.9, "avg": 49.2}},
    "地量(≤0.5倍均量) ＋ 創52週新高 ＋ 近3日向上跳空缺口": {"10": {"n": 359, "moonshotN": 42, "pct": 11.7, "avg": 44.9}, "20": {"n": 358, "moonshotN": 69, "pct": 19.3, "avg": 57.6}},
  };

  function formatMoonshotEntry(e) {
    return e.n + '筆中' + e.moonshotN + '次飆股(' + e.pct + '%)，平均漲幅+' + e.avg + '%';
  }
  function formatMoonshotStats(statsObj) {
    if (!statsObj) return '';
    return [5, 10, 20].filter(function(h){ return statsObj[h] != null; })
      .map(function(h){ return h + '日：' + formatMoonshotEntry(statsObj[h]); })
      .join('　｜　');
  }

  function pinnedComboStats(rows, combos, targetHorizon, statsFmt, typeLabel) {
    var retCol = 'ret' + targetHorizon + 'd';
    var valid = rows.filter(function(r){ return r[retCol] != null; });
    var flagDefs = buildConditionFlags(rows);
    var flagNames = Object.keys(flagDefs);
    // 同日超額報酬：每筆報酬扣掉同一評估日全部紀錄的平均
    var daySum = {}, dayCnt = {};
    valid.forEach(function(r) { daySum[r.evalDate] = (daySum[r.evalDate] || 0) + r[retCol]; dayCnt[r.evalDate] = (dayCnt[r.evalDate] || 0) + 1; });

    return combos.map(function(combo) {
      var isKdjMorningStar = combo.indexOf('KDJ近3日內黃金交叉') !== -1 && combo.indexOf('晨星剛形成') !== -1;
      var key = combo.join(' ＋ ');
      var label = (isKdjMorningStar ? '⭐ ' : '') + key;
      var recorded = statsFmt ? statsFmt(key) : (formatWinrates(PINNED_COMBO_WINRATES[key]) || null);
      var row = { '條件組合': label, '類型': typeLabel || (isTimingCombo(combo) ? '擇時型' : ''), '樣本數': 0, '歷史記錄': recorded };
      row[targetHorizon+'日平均報酬%'] = null;
      row[targetHorizon+'日勝率%'] = null;
      row['同日超額報酬%'] = null;
      row['t值(同日調整)'] = null;
      row['判讀'] = '';
      var missing = combo.filter(function(name){ return flagNames.indexOf(name) === -1; });
      if (missing.length) {
        row['備註'] = '缺少欄位（可能沒勾選對應的可選資料）：' + missing.join('、');
        return row;
      }
      var matched = valid.filter(function(r) {
        return combo.every(function(name) { return flagDefs[name](r); });
      });
      row['樣本數'] = matched.length;
      if (!matched.length) {
        row['備註'] = '目前回測資料裡沒有符合這個組合的樣本';
        return row;
      }
      var rets = matched.map(function(r){ return r[retCol]; });
      var n = rets.length;
      var avgRet = rets.reduce(function(a,b){return a+b;},0) / n;
      var winRate = rets.filter(function(v){return v>0;}).length / n * 100;
      var exs = matched.map(function(r){ return r[retCol] - daySum[r.evalDate] / dayCnt[r.evalDate]; });
      var exMean = exs.reduce(function(a,b){return a+b;},0) / n;
      var sd = n > 1 ? Math.sqrt(exs.reduce(function(s,v){ return s + (v-exMean)*(v-exMean); }, 0) / (n-1)) : 0;
      var t = sd > 0 ? exMean / (sd / Math.sqrt(n)) : null;
      row[targetHorizon+'日平均報酬%'] = avgRet.toFixed(2);
      row[targetHorizon+'日勝率%'] = winRate.toFixed(1);
      row['同日超額報酬%'] = exMean.toFixed(2);
      row['t值(同日調整)'] = t != null ? t.toFixed(2) : null;
      row['判讀'] = n < 30 ? '樣本不足' : (t != null && t >= 2 ? '✅ 選股有效' : (winRate >= 60 ? '⏱ 擇時（勝率來自大盤）' : '❌ 無效'));
      row['備註'] = n >= 50 ? '' : '⚠️樣本數偏少，僅供參考';
      return row;
    });
  }


  // ── 把陣列of物件轉成HTML表格 ──
  function objArrayToTable(rows, tableId) {
    if (!rows.length) return '<div style="color:#888;font-size:12px">（無資料）</div>';
    var cols = Object.keys(rows[0]);
    var idAttr = tableId ? ' id="'+tableId+'"' : '';
    var thead = '<thead><tr>' + cols.map(function(c){ return '<th style="padding:6px 10px;text-align:left;color:#aaa;font-size:11px;border-bottom:2px solid rgba(255,255,255,.1)">'+c+'</th>'; }).join('') + '</tr></thead>';
    var tbody = '<tbody>' + rows.map(function(r) {
      return '<tr>' + cols.map(function(c) {
        return '<td style="padding:5px 10px;font-size:12px;border-bottom:1px solid rgba(255,255,255,.05)">' + (r[c] == null ? '--' : r[c]) + '</td>';
      }).join('') + '</tr>';
    }).join('') + '</tbody>';
    return '<div style="overflow-x:auto"><table'+idAttr+' style="width:100%;border-collapse:collapse">' + thead + tbody + '</table></div>';
  }

  // ── Render ──
  function tag(text, type) {
    return '<span class="tag '+type+'">'+( type==='bull'?'▲':type==='bear'?'▼':'─' )+' '+text+'</span>';
  }

  function f(n, d) { return n!=null ? (+n).toFixed(d!=null?d:1) : '--'; }
  function fv(n) { return n!=null ? Math.round(n).toLocaleString() : '--'; }

  function renderPBCheck(pb) {
    var verdict = pb.allPass
      ? '<div style="background:rgba(0,200,100,.15);border:1px solid rgba(0,200,100,.4);border-radius:10px;padding:12px 18px;color:#00c864;font-size:15px;font-weight:700;margin-bottom:16px">✅ 符合進場條件（必要條件全部通過）' + (pb.bonusPassed ? '  +成交量增加分' : '') + '</div>'
      : '<div style="background:rgba(255,60,60,.12);border:1px solid rgba(255,60,60,.35);border-radius:10px;padding:12px 18px;color:#ff5555;font-size:15px;font-weight:700;margin-bottom:16px">❌ 不符合進場條件（必要條件 ' + pb.requiredPassed + '/' + pb.requiredTotal + ' 通過）</div>';

    var rows = pb.results.map(function(r) {
      var icon  = r.pass ? '✅' : (r.required ? '❌' : '—');
      var color = r.pass ? '#00c864' : (r.required ? '#ff5555' : '#888');
      var req   = r.required ? '' : ' <span style="font-size:10px;color:#888;background:rgba(255,255,255,.06);padding:1px 6px;border-radius:4px">加分</span>';
      var det   = r.detail ? ' <span style="font-size:11px;color:#888;margin-left:8px">(' + r.detail + ')</span>' : '';
      return '<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid rgba(255,255,255,.05)">'
        + '<span style="font-size:16px;width:22px;text-align:center">' + icon + '</span>'
        + '<span style="font-size:13px;color:' + color + '">' + r.label + req + det + '</span>'
        + '</div>';
    }).join('');

    return verdict
      + '<div style="background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:8px 16px">'
      + rows + '</div>';
  }

  function renderPatternCheck(pt) {
    var verdict = pt.anyJustBroke
      ? '<div style="background:rgba(255,152,0,.15);border:1px solid rgba(255,152,0,.4);border-radius:10px;padding:12px 18px;color:#ff9800;font-size:15px;font-weight:700;margin-bottom:16px">🔥 偵測到剛突破買點（較前一交易日新增）</div>'
      : (pt.anyBreakout
        ? '<div style="background:rgba(0,200,100,.15);border:1px solid rgba(0,200,100,.4);border-radius:10px;padding:12px 18px;color:#00c864;font-size:15px;font-weight:700;margin-bottom:16px">✅ 偵測到型態突破買點</div>'
        : (pt.anyFormed
          ? '<div style="background:rgba(240,165,0,.12);border:1px solid rgba(240,165,0,.35);border-radius:10px;padding:12px 18px;color:#f0a500;font-size:15px;font-weight:700;margin-bottom:16px">🕒 型態成形中，尚未突破確認</div>'
          : '<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:10px;padding:12px 18px;color:#888;font-size:15px;font-weight:700;margin-bottom:16px">－ 目前未偵測到符合的進場型態</div>'));

    var rows = pt.results.map(function(r) {
      var icon  = r.justBroke ? '🔥' : (r.breakout ? '✅' : (r.formed ? '🕒' : '－'));
      var color = r.justBroke ? '#ff9800' : (r.breakout ? '#00c864' : (r.formed ? '#f0a500' : '#666'));
      return '<div style="display:flex;align-items:flex-start;gap:10px;padding:8px 0;border-bottom:1px solid rgba(255,255,255,.05)">'
        + '<span style="font-size:16px;width:22px;text-align:center">' + icon + '</span>'
        + '<div><span style="font-size:13px;color:' + color + ';font-weight:600">' + r.name + (HIGH_CONSOLIDATION_IDS[r.id] ? '（高檔強勢整理：強勢股多為洗盤後續漲）' : '') + '</span>'
        + '<div style="font-size:11px;color:#888;margin-top:2px">' + r.desc + '</div>'
        + (r.detail ? '<div style="font-size:11px;color:#666;margin-top:2px">' + r.detail + '</div>' : '')
        + '</div></div>';
    }).join('');

    return verdict
      + '<div style="background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:8px 16px">'
      + rows + '</div>';
  }

  // ── AI 智能綜合分析（呼叫 OpenAI API）──
  function buildAnalysisPrompt(r) {
    var last = r.data[r.data.length - 1];
    var prev = r.data[r.data.length - 2] || last;
    var chg = last.close - prev.close, chgp = prev.close ? chg / prev.close * 100 : 0;

    var lines = [];
    lines.push('股票：' + r.stockId + ' ' + r.name);
    lines.push('資料日期：' + last.date + '　收盤：' + last.close + '　漲跌：' + (chg>=0?'+':'') + chg.toFixed(2) + ' (' + (chgp>=0?'+':'') + chgp.toFixed(2) + '%)');
    lines.push('成交量：' + last.volume + '　量比(vs MA20量)：' + (last.vm20 ? (last.volume/last.vm20).toFixed(2) : 'N/A') + 'x');
    if (last.macd != null && last.macdSig != null) {
      lines.push('MACD：DIF=' + last.macd.toFixed(2) + '　Signal=' + last.macdSig.toFixed(2) + '　柱狀=' + last.macdHist.toFixed(2) + (last.macd>last.macdSig ? '（偏多）' : '（偏空）'));
    }
    if (last.bbU != null && last.bbL != null) {
      var bbW = last.bbU - last.bbL;
      var bbP = bbW ? (last.close - last.bbL) / bbW * 100 : 50;
      lines.push('布林通道：上軌=' + last.bbU.toFixed(2) + '　下軌=' + last.bbL.toFixed(2) + '　股價位置=' + bbP.toFixed(0) + '%');
    }
    lines.push('');
    lines.push('【DMI多方力道評分】總分 ' + r.total + '/100');
    lines.push('+DI=' + (r.dm.plusDI!=null?r.dm.plusDI.toFixed(1):'N/A') + '　-DI=' + (r.dm.minusDI!=null?r.dm.minusDI.toFixed(1):'N/A') + '　ADX=' + (r.dm.adx!=null?r.dm.adx.toFixed(1):'N/A') + '　ADXR=' + (r.dm.adxr!=null?r.dm.adxr.toFixed(1):'N/A'));
    lines.push('方向性 ' + r.dm.diPts.toFixed(1) + '/50、趨勢強度 ' + r.dm.adxPts.toFixed(1) + '/30、趨勢動能 ' + r.dm.adxrPts.toFixed(1) + '/20');
    var bullSigs = r.dm.sigs.filter(function(s){ return s[1]==='bull'; }).map(function(s){ return s[0]; });
    var bearSigs = r.dm.sigs.filter(function(s){ return s[1]==='bear'; }).map(function(s){ return s[0]; });
    if (bullSigs.length) lines.push('多頭訊號：' + bullSigs.join('、'));
    if (bearSigs.length) lines.push('空頭訊號：' + bearSigs.join('、'));
    lines.push('');
    lines.push('【回後買上漲 8條件核對】必要條件通過 ' + r.pb.requiredPassed + '/' + r.pb.requiredTotal + (r.pb.allPass ? '（全數通過）' : ''));
    lines.push('');
    lines.push('【型態確認，15種進場型態】');
    r.pt.results.forEach(function(p) {
      var status = p.justBroke ? '🔥剛突破（較前一交易日新增）' : (p.breakout ? '✅已突破' : (p.formed ? '🕒成形中未突破' : '－未偵測到'));
      lines.push('・' + p.name + '：' + status + (p.detail ? '（' + p.detail + '）' : ''));
    });

    return lines.join('\n');
  }

  async function runAIAnalysis(r, panel) {
    var btn = panel.querySelector('.ai-analysis-btn');
    var box = panel.querySelector('.ai-result');
    var key = (iOpenAIKey.value || '').trim();
    if (!key) {
      box.style.display = 'block';
      box.style.color = '#ff8a8a';
      box.textContent = '請先在左側輸入 OpenAI API Key，才能使用 AI 智能綜合分析。';
      return;
    }
    var model = iOpenAIModel.value || 'gpt-4o-mini';
    btn.disabled = true;
    var btnOrigText = btn.textContent;
    btn.textContent = '🤖 分析中…';
    box.style.display = 'block';
    box.style.color = '#ddd';
    box.textContent = '正在請 AI 綜合研判技術面數據，請稍候…';

    var prompt = buildAnalysisPrompt(r);
    try {
      var res = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + key },
        body: JSON.stringify({
          model: model,
          temperature: 0.4,
          max_tokens: 900,
          messages: [
            { role: 'system', content: '你是一位精通台股技術分析的資深操盤手，熟悉DMI趨向指標（+DI／-DI／ADX／ADXR）多方力道評分方法論，以及朱家泓《技術分析全攻略》的回後買上漲、頭肩底等進場型態判斷。請根據使用者提供的個股技術數據摘要，用繁體中文給出：1) 整體技術面研判（3-4句） 2) 進場時機與風險提示 3) 綜合建議（積極做多／可考慮／觀望／不建議）。語氣專業、精簡、避免空泛用詞，並提醒這僅為技術面參考，非投資建議。' },
            { role: 'user', content: prompt }
          ]
        })
      });
      var json = await res.json();
      if (!res.ok) {
        var msg = (json.error && json.error.message) ? json.error.message : ('HTTP ' + res.status);
        throw new Error(msg);
      }
      var content = json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content;
      box.textContent = content || '（AI 未回傳有效內容）';
    } catch (e) {
      box.style.color = '#ff8a8a';
      box.textContent = '❌ AI 分析失敗：' + e.message + '\n請確認 API Key 是否正確、額度是否足夠，或稍後再試。';
    } finally {
      btn.disabled = false;
      btn.textContent = btnOrigText;
    }
  }

  function renderSingle(data, name, dm, pb, total, pt) {
    // total passed in
    var vc, vt, vb;
    if(total>=80){vc='#00c864';vt='強力買進訊號';vb='rgba(0,200,100,.15)';}
    else if(total>=65){vc='#2196f3';vt='可考慮進場';vb='rgba(33,150,243,.15)';}
    else if(total>=50){vc='#f0a500';vt='觀望為主';vb='rgba(240,165,0,.15)';}
    else{vc='#ff3c3c';vt='不建議進場';vb='rgba(255,60,60,.15)';}

    var last=data[data.length-1], prev=data[data.length-2]||last;
    var chg=last.close-prev.close, chgp=prev.close?chg/prev.close*100:0;
    var cc=chg>=0?'up':'dn';
    var vr=last.vm20?last.volume/last.vm20:1;
    var vl2=vr>1.2?'放量':vr<0.8?'縮量':'正常';
    var vc2=vr>1.2?'up':vr<0.8?'dn':'flat';

    // Advice
    var act, adv=[], sl, tgt;
    var ma20v=last.ma20||last.close;
    var bbUp=last.bbU||last.close*1.1;
    if(total>=80){
      act='🟢 積極做多';
      adv.push('<strong>'+name+'</strong> 多方力道評分 '+total+' 分，DMI顯示多方力道強勁，建議積極做多。');
      if(dm.tdir==='多頭')adv.push('+DI大於-DI且趨勢轉強，順勢操作，逢低分批佈局。');
      sl='停損設於 <strong>'+(ma20v*0.97).toFixed(1)+'</strong> 元（20MA下方3%）';
      tgt='目標參考 <strong>'+(last.close*((bbUp-last.close)/last.close+1)).toFixed(1)+'</strong> 元（布林上軌）';
    } else if(total>=65){
      act='🔵 可考慮進場';
      adv.push('<strong>'+name+'</strong> 評分 '+total+' 分，DMI偏多，可考慮分批進場。');
      adv.push('建議等待回測均線後再進場，降低風險。');
      sl='停損建議 <strong>'+(ma20v*0.98).toFixed(1)+'</strong> 元（20MA下方2%）';
      tgt='短線目標 <strong>'+(last.close*1.08).toFixed(1)+'</strong> 元（+8%）';
    } else if(total>=50){
      act='🟡 觀望為主';
      adv.push('<strong>'+name+'</strong> 評分 '+total+' 分，DMI訊號混雜或趨勢不明確，建議觀望。');
      adv.push('等待ADX轉強或+DI/-DI方向明確後再行動。');
      sl='暫不建議進場';
      tgt='等待更佳時機';
    } else {
      act='🔴 不適合進場';
      adv.push('<strong>'+name+'</strong> 評分 '+total+' 分，DMI偏空，不建議進場。');
      if(dm.tdir==='空頭')adv.push('目前-DI大於+DI，空方力道較強，切忌逆勢做多，等待趨勢反轉。');
      else adv.push('DMI指標偏弱或資料不足，應持現金等待機會。');
      sl='持倉者建議設停損出場';
      tgt='等待多頭訊號出現';
    }

    var hints=dm.sigs.filter(function(s){return s[1]==='bull';}).map(function(s){return s[0];}).slice(0,5);

    // Dim cards：三個組成多方力道評分的子項目（方向性50分＋趨勢強度30分＋趨勢動能20分）
    var dims=[
      {t:'🧭 方向性 (+DI vs -DI)', score:dm.diPts, max:50, d:'+DI相對-DI的優勢程度'},
      {t:'💪 趨勢強度 (ADX)', score:dm.adxPts, max:30, d:'ADX值，越高趨勢越明確'},
      {t:'🚀 趨勢動能 (ADX vs ADXR)', score:dm.adxrPts, max:20, d:'ADX>ADXR代表趨勢正在轉強'}
    ];
    var dimHtml=dims.map(function(x){
      var pct=x.max?x.score/x.max*100:0, col=pct>=70?'#00c864':pct>=40?'#f0a500':'#ff3c3c';
      return '<div class="dim-card"><div class="dim-row"><div><div class="dim-t">'+x.t+'</div><div class="dim-d">'+x.d+'</div></div><div class="dim-s" style="color:'+col+'">'+x.score.toFixed(1)+'<span>/'+x.max+'</span></div></div><div class="pbar"><div class="pfill" style="background:'+col+';width:'+pct+'%"></div></div></div>';
    }).join('');

    // Signal cols：DMI訊號摘要（取代原本四維度各自的訊號欄）
    var sigCols=[{t:'DMI訊號',s:dm.sigs}];
    var sigHtml=sigCols.map(function(x){
      return '<div><div class="sig-t">'+x.t+'</div>'+(x.s.length?x.s.map(function(s){return tag(s[0],s[1]);}).join(''):tag('無明顯訊號','neu'))+'</div>';
    }).join('');

    // RSI
    var rv=last.rsi, rcolor=rv>70?'#ff5555':rv<30?'#00c864':'#f0a500';
    var rcls=rv>70?'rsi-ob':rv<30?'rsi-os':'rsi-ok';
    var rlbl=rv>70?'⚠️ 超買區（>70），注意回調':rv<30?'💚 超賣區（<30），留意反彈':'位於正常區間（30-70）';

    // KD
    var kdHtml = '<span style="color:#666">資料不足</span>';
    if (last.kdK != null && last.kdD != null) {
      var kdColor = last.kdK > last.kdD ? '#00c864' : '#ff5555';
      var kdLbl = last.kdK > last.kdD ? 'K&gt;D 偏多' : 'K&lt;D 偏空';
      kdHtml = 'K=<span style="font-size:20px;font-weight:700;color:'+kdColor+'">'+last.kdK.toFixed(1)+'</span>'
        + '　D=<span style="font-weight:700">'+last.kdD.toFixed(1)+'</span>'
        + '<div class="rsi-n '+(last.kdK>last.kdD?'rsi-os':'rsi-ob')+'" style="margin-top:8px">'+kdLbl+'</div>';
    }

    // MACD
    var macdHtml = '<span style="color:#666">資料不足</span>';
    if (last.macd != null && last.macdSig != null) {
      var macdBull = last.macd > last.macdSig;
      var macdColor = macdBull ? '#00c864' : '#ff5555';
      var macdCross = '';
      if (prev.macd != null && prev.macdSig != null) {
        if (prev.macd <= prev.macdSig && last.macd > last.macdSig) macdCross = '⚡ 黃金交叉';
        else if (prev.macd >= prev.macdSig && last.macd < last.macdSig) macdCross = '⚡ 死亡交叉';
      }
      var macdLbl = (macdBull ? 'DIF在Signal上方，偏多' : 'DIF在Signal下方，偏空') + (macdCross ? '　' + macdCross : '');
      macdHtml = 'DIF=<span style="font-size:18px;font-weight:700;color:'+macdColor+'">'+last.macd.toFixed(2)+'</span>'
        + '　Signal=<span style="font-weight:700">'+last.macdSig.toFixed(2)+'</span>'
        + '　柱狀=<span style="font-weight:700;color:'+(last.macdHist>=0?'#00c864':'#ff5555')+'">'+last.macdHist.toFixed(2)+'</span>'
        + '<div class="rsi-n '+(macdBull?'rsi-os':'rsi-ob')+'" style="margin-top:8px">'+macdLbl+'</div>';
    }

    // 布林通道
    var bbHtml = '<span style="color:#666">資料不足</span>';
    if (last.bbU != null && last.bbL != null) {
      var bbWidth = last.bbU - last.bbL;
      var bbPos = bbWidth ? (last.close - last.bbL) / bbWidth * 100 : 50;
      var bbColor = bbPos > 80 ? '#ff5555' : bbPos < 20 ? '#00c864' : '#f0a500';
      var bbLbl = bbPos > 80 ? '⚠️ 貼近上軌，注意過熱回檔' : bbPos < 20 ? '💚 貼近下軌，留意反彈' : '位於通道中段';
      // 通道寬度相對股價的比例，用來概略判斷是否進入「窄幅整理」
      var bbWidthPct = last.close ? bbWidth / last.close * 100 : 0;
      if (bbWidthPct < 8) bbLbl += '　🔸通道收窄，留意變盤';
      bbHtml = '上軌=<span style="font-weight:700">'+f(last.bbU)+'</span>　下軌=<span style="font-weight:700">'+f(last.bbL)+'</span>'
        + '<div style="margin-top:4px">股價位置：<span style="font-weight:700;color:'+bbColor+'">'+bbPos.toFixed(0)+'%</span>（通道寬度 '+bbWidthPct.toFixed(1)+'%）</div>'
        + '<div class="rsi-n" style="margin-top:8px;background:rgba(240,165,0,.12);color:'+bbColor+'">'+bbLbl+'</div>';
    }

    // MA table
    var maRows=[['MA5','ma5'],['MA10','ma10'],['MA20','ma20'],['MA60','ma60']].filter(function(x){return last[x[1]]!=null;}).map(function(x){
      var diff=(last.close-last[x[1]])/last[x[1]]*100;
      return '<tr><td>'+x[0]+'</td><td>'+f(last[x[1]])+'</td><td class="'+(diff>0?'tup':'tdn')+'">'+(diff>0?'+':'')+diff.toFixed(2)+'% ('+(diff>0?'上方':'下方')+')</td></tr>';
    }).join('');

    // Raw data
    var rawRows=data.slice(-20).reverse().map(function(d){
      return '<tr><td>'+d.date+'</td><td>'+f(d.open)+'</td><td>'+f(d.high)+'</td><td>'+f(d.low)+'</td><td>'+f(d.close)+'</td><td>'+fv(d.volume)+'</td><td>'+f(d.ma5)+'</td><td>'+f(d.ma20)+'</td><td>'+f(d.ma60)+'</td><td>'+(d.rsi!=null?f(d.rsi):'--')+'</td><td>'+(d.kdK!=null?f(d.kdK):'--')+'</td><td>'+(d.kdD!=null?f(d.kdD):'--')+'</td><td>'+(d.macd!=null?f(d.macd):'--')+'</td><td>'+(d.macdSig!=null?f(d.macdSig):'--')+'</td><td>'+(d.bbU!=null?f(d.bbU):'--')+'</td><td>'+(d.bbL!=null?f(d.bbL):'--')+'</td></tr>';
    }).join('');

    var html =
      '<div style="font-size:20px;font-weight:700;color:#fff;margin-bottom:20px">🏷️ '+name+'</div>'+
      '<div class="metrics">'+
        '<div class="mc"><div class="mc-l">最新收盤</div><div class="mc-v">'+f(last.close)+'</div><div class="mc-d '+cc+'">'+(chg>=0?'+':'')+f(chg)+' ('+(chgp>=0?'+':'')+chgp.toFixed(2)+'%)</div></div>'+
        '<div class="mc"><div class="mc-l">當日成交量</div><div class="mc-v">'+fv(last.volume)+'</div></div>'+
        '<div class="mc"><div class="mc-l">量比 vs MA20</div><div class="mc-v">'+vr.toFixed(2)+'x</div><div class="mc-d '+vc2+'">'+vl2+'</div></div>'+
        '<div class="mc"><div class="mc-l">資料日期</div><div class="mc-v" style="font-size:15px">'+last.date+'</div></div>'+
      '</div>'+
      '<hr class="div">'+
      '<div class="sec">📊 綜合評分</div>'+
      '<div class="score-wrap">'+
        '<div class="score-card"><div class="score-n" style="color:'+vc+'">'+total+'</div><div class="score-l">綜合評分 / 100</div><div class="score-v" style="background:'+vb+';color:'+vc+'">'+vt+'</div></div>'+
        '<div>'+dimHtml+'</div>'+
      '</div>'+
      '<hr class="div">'+
      '<div class="sec">🔔 技術訊號</div>'+
      '<div class="sig-grid">'+sigHtml+'</div>'+
      '<hr class="div">'+
      '<div class="sec">💡 操作建議</div>'+
      '<div class="adv"><div class="adv-a">'+act+'</div><div class="adv-t">'+adv.join('<br>')+'</div><div class="adv-m">🛑 <strong>停損：</strong>'+sl+'<br>🎯 <strong>目標：</strong>'+tgt+(hints.length?'<br>✅ <strong>多頭訊號：</strong>'+hints.join(' · '):'')+'</div></div>'+
      '<button class="btn ai-analysis-btn" type="button" style="margin-top:14px;width:auto;padding:10px 22px">🤖 AI 智能綜合分析</button>'+
      '<div class="ai-result" style="display:none;margin-top:14px;background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px 18px;font-size:13px;line-height:1.8;color:#ddd;white-space:pre-wrap"></div>'+
      '<hr class="div">'+
      '<div class="sec">🎯 回後買上漲 · 進場條件核對</div>'+
      renderPBCheck(pb)+
      '<hr class="div">'+
      '<div class="sec">🔍 型態確認（15種進場型態）</div>'+
      renderPatternCheck(pt)+
      '<hr class="div">'+
      '<div class="sec">📐 指標對照</div>'+
      '<div class="info-grid">'+
        '<div class="ib"><div class="ib-t">RSI 指標</div><div class="rsi-v" style="color:'+rcolor+'">'+f(rv)+'</div><div class="rsi-n '+rcls+'">'+rlbl+'</div></div>'+
        '<div class="ib"><div class="ib-t">KD 指標</div><div style="font-size:15px;margin-top:4px">'+kdHtml+'</div></div>'+
        '<div class="ib"><div class="ib-t">MACD 指標</div><div style="font-size:14px;margin-top:4px">'+macdHtml+'</div></div>'+
        '<div class="ib"><div class="ib-t">布林通道</div><div style="font-size:14px;margin-top:4px">'+bbHtml+'</div></div>'+
        '<div class="ib"><div class="ib-t">均線對照</div><table><thead><tr><th>均線</th><th>數值</th><th>股價偏離</th></tr></thead><tbody>'+maRows+'</tbody></table></div>'+
      '</div>'+
      '<hr class="div">'+
      '<div class="sec">📉 技術分析圖表</div>'+
      '<div class="chart-box"><div id="chart"></div></div>'+
      '<div class="raw-tog" onclick="this.nextElementSibling.classList.toggle(\'open\');this.querySelector(\'span\').textContent=this.nextElementSibling.classList.contains(\'open\')?\'▼\':\'▶\'"><span>▶</span> 📋 原始資料（最近20筆）</div>'+
      '<div class="raw-body"><table><thead><tr><th>日期</th><th>開盤</th><th>最高</th><th>最低</th><th>收盤</th><th>成交量</th><th>MA5</th><th>MA20</th><th>MA60</th><th>RSI</th><th>KD-K</th><th>KD-D</th><th>MACD</th><th>Signal</th><th>BB上軌</th><th>BB下軌</th></tr></thead><tbody>'+rawRows+'</tbody></table></div>';

    var chartId = 'chart_' + name.replace(/\s/g,'_') + '_' + Date.now();
    html = html.replace('id="chart"', 'id="' + chartId + '"');
    return { html: html, chartId: chartId };
  }

  function render(data, name) {
    var dmi = calcDMI(data, 14);
    var dm = scoreDMI(dmi);
    var pb=checkPullbackBuy(data);
    var pt=detectPatterns(data, pb);
    var total=dm.score;
    var res = renderSingle(data, name, dm, pb, total, pt);
    output.innerHTML = res.html;
    drawChart(data, name, res.chartId, pt);
  }

  // 朱家泓「短線轉折波」畫法（依課程圖表2-1-3／2-1-4）：
  // 用收盤價與5日均線的穿越關係取高低點——
  //   收盤價「跌破」5日均線 → 取這段上漲過程的「高點」為一個轉折點
  //   收盤價「突破」5日均線 → 取這段下跌過程的「低點」為一個轉折點
  // 再把這些高低點依序連接成鋸齒狀的轉折波。
  function isSaneBar(d) {
    // 過濾掉資料來源偶爾出現的異常值（例如某天 low 被錯誤回傳為 0 或極小值），
    // 避免單一根爛資料把轉折波拉出一條不合理的長長尖刺
    if (!d) return false;
    if (!(d.high > 0) || !(d.low > 0) || !(d.close > 0)) return false;
    if (!isFinite(d.high) || !isFinite(d.low) || !isFinite(d.close)) return false;
    if (d.ma5 != null && d.ma5 > 0) {
      if (d.low < d.ma5 * 0.4 || d.high > d.ma5 * 2.5) return false;
    }
    return true;
  }

  function buildZigzag(data) {
    var n = data.length;
    // 找到第一個 MA5 已經有值、且資料正常的位置（前4根K棒沒有5日均線可比較）
    var start = 0;
    while (start < n && (data[start].ma5 == null || !isSaneBar(data[start]))) start++;
    if (start >= n - 1) return [];

    var state = data[start].close >= data[start].ma5 ? 'above' : 'below'; // 目前收盤在5日均線上方或下方
    var points = [{
      idx: start,
      price: state === 'above' ? data[start].low : data[start].high,
      type: state === 'above' ? 'L' : 'H'
    }];
    var extremeIdx = start;
    var extremePrice = state === 'above' ? data[start].high : data[start].low;

    for (var i = start + 1; i < n; i++) {
      var d = data[i];
      if (d.ma5 == null || !isSaneBar(d)) continue;
      if (state === 'above') {
        // 收盤在5日均線之上，持續追蹤這段上漲的最高點
        if (d.high > extremePrice) { extremePrice = d.high; extremeIdx = i; }
        if (d.close < d.ma5) {
          // 收盤跌破5日均線 → 確認剛才追蹤到的高點為一個轉折高點
          points.push({idx:extremeIdx, price:extremePrice, type:'H'});
          state = 'below';
          extremeIdx = i; extremePrice = d.low;
        }
      } else {
        // 收盤在5日均線之下，持續追蹤這段下跌的最低點
        if (d.low < extremePrice) { extremePrice = d.low; extremeIdx = i; }
        if (d.close > d.ma5) {
          // 收盤突破5日均線 → 確認剛才追蹤到的低點為一個轉折低點
          points.push({idx:extremeIdx, price:extremePrice, type:'L'});
          state = 'above';
          extremeIdx = i; extremePrice = d.high;
        }
      }
    }

    // 收尾：把目前仍在追蹤中的高/低點畫出來，再接到最新一根K棒的收盤價，確保線一定連到最新資料
    var lastIdx = n - 1;
    if (extremeIdx !== lastIdx) {
      points.push({idx:extremeIdx, price:extremePrice, type: state === 'above' ? 'H' : 'L'});
      points.push({idx:lastIdx, price:data[lastIdx].close, type: state === 'above' ? 'L' : 'H'});
    } else {
      points.push({idx:extremeIdx, price:data[lastIdx].close, type: state === 'above' ? 'H' : 'L'});
    }
    return points;
  }

  function drawChart(data, name, containerId, pt) {
    containerId = containerId || 'chart';
    // 圖表顯示範圍改用完整的 data（已由「分析天數」在抓資料時決定範圍），
    // 不再寫死只看最近120根K棒，讓滑桿調整能真正反映在圖表上。
    var rawTail = data;
    // 過濾掉資料異常的K棒（開高低收有任一項是 0、負值或非數字），
    // 避免圖表出現「沒有K棒的空白位置」卻仍有轉折波或均線的線硬穿過去
    var tail = rawTail.filter(function(d) {
      return d && d.open > 0 && d.high > 0 && d.low > 0 && d.close > 0 &&
        isFinite(d.open) && isFinite(d.high) && isFinite(d.low) && isFinite(d.close);
    });
    var dates = tail.map(function(d){return d.date;});
    var volColors = tail.map(function(d){return d.close>=d.open?'#ef5350':'#26a69a';});
    var histColors = tail.map(function(d){return (d.macdHist||0)>=0?'#ef5350':'#26a69a';});

    var zz = buildZigzag(tail);
    var zzX = zz.map(function(p){ return tail[p.idx].date; });
    var zzY = zz.map(function(p){ return p.price; });

    // 每個「型態確認」如果已成形，就把輔助線（頸線/壓力線/切線/軌道線）畫在圖上，
    // 型態成形但未突破 → 虛線；已經突破 → 實線＋★標註，方便直接在圖上對照型態辨識依據
    var PATTERN_LINE_STYLE = {
      hs:      { color: '#ffd54f', label: '頭肩底頸線' },
      chs:     { color: '#ff8a65', label: '複式頭肩底頸線' },
      nb:      { color: '#81c784', label: 'N字底壓力' },
      tb:      { color: '#4dd0e1', label: '三重底壓力' },
      rb:      { color: '#64b5f6', label: '圓弧底壓力' },
      fb:      { color: '#ba68c8', label: '一字底整理區間高點' },
      abc:     { color: '#ff2ecc', label: 'ABC下降切線起點' },
      channel: { color: '#ffa726', label: '上升軌道線' },
      blackk:  { color: '#e57373', label: '大量黑K高點' },
      kbp:     { color: 'rgba(255,255,255,.85)', label: 'K線橫盤首日高點' }
    };
    // 母子懷抱／晨星／夜星屬於2~3根K線的短線反轉訊號，沒有持續延伸的支撐/壓力線可畫，
    // 改用箭頭標註直接指到型態發生的那根（或那兩根）K棒位置，方便在圖上直接辨識。
    var PATTERN_MARKER_STYLE = {
      harami_bear:  { color: '#ff8a65' },
      harami_bull:  { color: '#81c784' },
      morning_star: { color: '#4dd0e1' },
      evening_star: { color: '#ff5252' }
    };
    var shapes = [];
    var annotations = [];
    var lastIdx = tail.length - 1;
    if (pt && pt.results) {
      pt.results.forEach(function(p) {
        if (!p.formed) return;
        var style = PATTERN_LINE_STYLE[p.id];
        if (style) {
          // 上升軌道線另外多畫一條下緣支撐線（line2），其餘型態只有一條輔助線（line）
          [p.line, p.line2].forEach(function(ln, idx) {
            if (!ln || ln.i1 < 0 || ln.i1 >= tail.length) return;
            var lineEndPrice = ln.p1 + ln.slope * (lastIdx - ln.i1);
            shapes.push({
              type: 'line', xref: 'x', yref: 'y',
              x0: tail[ln.i1].date, y0: ln.p1,
              x1: tail[lastIdx].date, y1: lineEndPrice,
              line: { color: style.color, width: idx === 0 ? 2 : 1.3, dash: p.breakout ? 'solid' : 'dash' }
            });
            if (idx === 0) {
              annotations.push({
                x: tail[ln.i1].date, y: ln.p1, xref: 'x', yref: 'y',
                text: style.label, showarrow: true, arrowhead: 2, arrowcolor: style.color,
                font: { color: style.color, size: 10 }, ax: -10, ay: -30
              });
              if (p.breakout) {
                annotations.push({
                  x: tail[lastIdx].date, y: tail[lastIdx].close, xref: 'x', yref: 'y',
                  text: '★ 突破' + p.name, showarrow: true, arrowhead: 2, arrowcolor: style.color,
                  font: { color: style.color, size: 11 }, ax: 10, ay: -35
                });
              }
            }
          });
          return;
        }
        var mStyle = PATTERN_MARKER_STYLE[p.id];
        if (mStyle && p.marker && p.marker.idx >= 0 && p.marker.idx < tail.length) {
          var m = p.marker;
          annotations.push({
            x: tail[m.idx].date, y: m.price, xref: 'x', yref: 'y',
            text: (p.breakout ? '★ ' : '') + m.label, showarrow: true, arrowhead: 2, arrowcolor: mStyle.color,
            font: { color: mStyle.color, size: 11 },
            ax: 0, ay: m.dir === 'up' ? -32 : 32
          });
        }
      });
    }

    // 分價量表 POC 區（金色帶）＋POC價虛線
    var chartProf = computeVolumeProfile(tail, VP_PARAMS.lookback, VP_PARAMS.bins);
    if (chartProf) {
      var pocX0 = tail[Math.max(0, tail.length - VP_PARAMS.lookback)].date;
      shapes.push({ type: 'rect', xref: 'x', yref: 'y', x0: pocX0, x1: tail[lastIdx].date, y0: chartProf.pocLow, y1: chartProf.pocHigh,
        fillcolor: 'rgba(255,215,0,.10)', line: { width: 0 } });
      shapes.push({ type: 'line', xref: 'x', yref: 'y', x0: pocX0, x1: tail[lastIdx].date, y0: chartProf.pocPrice, y1: chartProf.pocPrice,
        line: { color: '#ffd700', width: 1.2, dash: 'dot' } });
      annotations.push({ x: tail[lastIdx].date, y: chartProf.pocPrice, xref: 'x', yref: 'y', text: 'POC ' + chartProf.pocPrice.toFixed(1),
        showarrow: false, xanchor: 'left', font: { color: '#ffd700', size: 10 } });
    }

    var el = document.getElementById(containerId);
    // 直接量測容器實際可用寬度，不依賴 Plotly 自動偵測（避免容器剛顯示時量到舊的窄寬度）
    var measuredWidth = el && el.parentElement ? el.parentElement.getBoundingClientRect().width : 0;
    var chartWidth = measuredWidth > 100 ? Math.floor(measuredWidth) : undefined;

    var layout = {
      title:{text:name+' 技術分析圖',font:{color:'#fff',size:14}},
      template:'plotly_dark',paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(15,20,35,1)',
      autosize:!chartWidth,height:680,margin:{l:50,r:130,t:45,b:25},
      shapes: shapes, annotations: annotations,
      xaxis:{rangeslider:{visible:false},gridcolor:'rgba(255,255,255,.04)',domain:[0,1]},
      yaxis:{domain:[0.42,1],gridcolor:'rgba(255,255,255,.04)'},
      yaxis2:{domain:[0.22,0.40],gridcolor:'rgba(255,255,255,.04)'},
      yaxis3:{domain:[0,0.20],gridcolor:'rgba(255,255,255,.04)'},
      // 改為右側垂直圖例，避免窄寬度下橫向圖例換行、疊在K線圖上方
      legend:{orientation:'v',x:1.01,y:1,xanchor:'left',bgcolor:'rgba(0,0,0,0)',font:{size:10}},
      showlegend:true
    };
    if (chartWidth) layout.width = chartWidth;

    Plotly.newPlot(containerId, [
      {type:'candlestick',x:dates,open:tail.map(function(d){return d.open;}),high:tail.map(function(d){return d.high;}),low:tail.map(function(d){return d.low;}),close:tail.map(function(d){return d.close;}),name:'K線',increasing:{line:{color:'#ef5350'},fillcolor:'#ef5350'},decreasing:{line:{color:'#26a69a'},fillcolor:'#26a69a'},xaxis:'x',yaxis:'y'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.ma5;}),name:'MA5',line:{color:'#ffeb3b',width:1.2},xaxis:'x',yaxis:'y'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.ma10;}),name:'MA10',line:{color:'#ff9800',width:1.2},xaxis:'x',yaxis:'y'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.ma20;}),name:'MA20',line:{color:'#2196f3',width:1.2},xaxis:'x',yaxis:'y'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.ma60;}),name:'MA60',line:{color:'#9c27b0',width:1.2},xaxis:'x',yaxis:'y'},
      {type:'scatter',x:zzX,y:zzY,name:'轉折波',mode:'lines+markers',line:{color:'#00e5ff',width:1.8},marker:{size:5,color:'#00e5ff'},xaxis:'x',yaxis:'y'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.bbU;}),name:'BB上軌',line:{color:'rgba(100,200,255,.4)',width:1,dash:'dot'},showlegend:false,xaxis:'x',yaxis:'y'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.bbL;}),name:'BB下軌',line:{color:'rgba(100,200,255,.4)',width:1,dash:'dot'},fill:'tonexty',fillcolor:'rgba(100,200,255,.04)',showlegend:false,xaxis:'x',yaxis:'y'},
      {type:'bar',x:dates,y:tail.map(function(d){return d.volume;}),marker:{color:volColors},name:'成交量',opacity:.7,xaxis:'x',yaxis:'y2'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.vm20;}),name:'量MA20',line:{color:'#ff9800',width:1.5},xaxis:'x',yaxis:'y2'},
      {type:'bar',x:dates,y:tail.map(function(d){return d.macdHist;}),marker:{color:histColors},name:'MACD柱',opacity:.8,xaxis:'x',yaxis:'y3'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.macd;}),name:'MACD',line:{color:'#2196f3',width:1.5},xaxis:'x',yaxis:'y3'},
      {type:'scatter',x:dates,y:tail.map(function(d){return d.macdSig;}),name:'Signal',line:{color:'#ff9800',width:1.5},xaxis:'x',yaxis:'y3'}
    ], layout, {responsive:true});

    // 保險：容器寬度可能在畫完之後才穩定下來（字型載入、版面重排等），
    // 多次延遲重新量測並強制套用正確寬度，確保最終一定撐滿容器
    [50, 200, 500].forEach(function(delay) {
      setTimeout(function() {
        var container = document.getElementById(containerId);
        if (!container || !container.parentElement) return;
        var w = Math.floor(container.parentElement.getBoundingClientRect().width);
        if (w > 100 && Math.abs(w - (container._fullLayout ? container._fullLayout.width : 0)) > 20) {
          Plotly.relayout(containerId, { width: w, autosize: false }).catch(function(){});
        }
      }, delay);
    });
  }

  // 視窗尺寸改變（含側邊欄收合、字型載入造成的版面變動）時，重新調整當前顯示的圖表寬度
  window.addEventListener('resize', function() {
    var activePanel = document.querySelector('.tab-panel.active');
    if (!activePanel) return;
    var chartEl = activePanel.querySelector('.chart-box > div');
    if (chartEl && chartEl.id && chartEl.parentElement) {
      var w = Math.floor(chartEl.parentElement.getBoundingClientRect().width);
      if (w > 100) {
        try { Plotly.relayout(chartEl.id, { width: w }); } catch(e) {}
      }
    }
  });

  window.switchTab = switchTab;
})();
</script>
</body>
</html>

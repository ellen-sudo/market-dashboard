function doGet() {
  var html = HtmlService.createHtmlOutput(buildDashboard());
  html.setTitle('Raleigh Market Dashboard');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function buildDashboard() {
  var ss = SpreadsheetApp.openById('1TTvm4cY30u7URVhAmRLX4Z4C9zMiu_OSWt5wNaCFyio');
  var sheet = ss.getSheetByName('Form Responses 1');
  var rows = sheet.getDataRange().getValues();

  // Parse percentage — handles "98.3", "98.30%", 0.983
  function parsePct(val) {
    if (val === null || val === undefined || val === '') return null;
    var s = String(val).replace('%', '').trim();
    var n = parseFloat(s);
    if (isNaN(n)) return null;
    return n > 1 ? n / 100 : n;
  }

  // Deduplicate by month — keep last row per month
  var byMonth = {};
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var issueDate = row[1];
    if (!issueDate) continue;
    var dateObj = (issueDate instanceof Date) ? issueDate : new Date(issueDate);
    if (isNaN(dateObj.getTime())) continue;
    var year = dateObj.getFullYear();
    var month = String(dateObj.getMonth() + 1).padStart(2, '0');
    var key = year + '-' + month;
    byMonth[key] = row; // last row wins
  }

  // Build sorted data array
  var months = Object.keys(byMonth).sort();
  var data = [];
  for (var j = 0; j < months.length; j++) {
    var key = months[j];
    var r = byMonth[key];
    data.push({
      month: key,
      // This year
      nl_ty:  parseFloat(r[4])  || 0,
      np_ty:  parseFloat(r[5])  || 0,
      dom_ty: parseFloat(r[6])  || 0,
      cs_ty:  parseFloat(r[7])  || 0,
      exp_ty: parseFloat(r[8])  || 0,
      pol_ty: parsePct(r[9])    || 0,
      pal_ty: parsePct(r[10])   || 0,
      // Last year
      nl_ly:  parseFloat(r[12]) || 0,
      np_ly:  parseFloat(r[13]) || 0,
      dom_ly: parseFloat(r[14]) || 0,
      cs_ly:  parseFloat(r[15]) || 0,
      exp_ly: parseFloat(r[16]) || 0,
      pol_ly: parsePct(r[17])   || 0,
      pal_ly: parsePct(r[18])   || 0
    });
  }

  var dataJson = JSON.stringify(data);

  var monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];

  // Get latest month label for display
  var latestLabel = '';
  var latestYear = '';
  if (data.length > 0) {
    var parts = data[data.length-1].month.split('-');
    latestLabel = monthNames[parseInt(parts[1])-1].toUpperCase();
    latestYear = parts[0];
  }

  var html = '<!DOCTYPE html>\n'
    + '<html lang="en">\n'
    + '<head>\n'
    + '<meta charset="UTF-8">\n'
    + '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    + '<title>Raleigh Market Dashboard</title>\n'
    + '<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">\n'
    + '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"><\/script>\n'
    + '<style>\n'
    + ':root{--bg:#0F1117;--surface:#181C27;--surface2:#1E2435;--border:rgba(255,255,255,0.07);--gold:#D4A843;--gold-dim:rgba(212,168,67,0.15);--teal:#3ECFB2;--teal-dim:rgba(62,207,178,0.12);--red:#E05C5C;--red-dim:rgba(224,92,92,0.12);--text:#E8E4DC;--muted:#7A8096;--white:#FFFFFF;}\n'
    + '*{margin:0;padding:0;box-sizing:border-box;}\n'
    + 'body{background:var(--bg);font-family:"DM Sans",sans-serif;color:var(--text);min-height:100vh;}\n'
    + '.header{padding:36px 40px 28px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:flex-end;gap:24px;flex-wrap:wrap;}\n'
    + '.header h1{font-family:"Bebas Neue",sans-serif;font-size:48px;letter-spacing:2px;color:var(--white);line-height:1;}\n'
    + '.header h1 span{color:var(--gold);}\n'
    + '.header p{color:var(--muted);font-size:13px;margin-top:6px;}\n'
    + '.badge{background:var(--gold-dim);border:1px solid rgba(212,168,67,0.3);border-radius:4px;padding:8px 16px;text-align:right;}\n'
    + '.badge .mo{font-family:"Bebas Neue",sans-serif;font-size:22px;color:var(--gold);letter-spacing:2px;}\n'
    + '.badge .sub{font-size:11px;color:var(--muted);letter-spacing:1px;text-transform:uppercase;}\n'
    + '.main{padding:32px 40px;max-width:1300px;margin:0 auto;}\n'
    + '.auto-bar{background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:14px 20px;margin-bottom:28px;font-size:13px;color:var(--muted);line-height:1.5;}\n'
    + '.auto-bar strong{color:var(--gold);}\n'
    + '.kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:32px;}\n'
    + '.kpi{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:18px 20px;}\n'
    + '.kpi-label{font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:8px;}\n'
    + '.kpi-value{font-family:"Bebas Neue",sans-serif;font-size:36px;letter-spacing:1px;color:var(--white);line-height:1;}\n'
    + '.kpi-delta{margin-top:6px;font-size:12px;}\n'
    + '.kpi-compare{margin-top:4px;font-size:11px;color:var(--muted);}\n'
    + '.up{color:var(--teal);} .down{color:var(--red);}\n'
    + '.section-label{font-size:11px;text-transform:uppercase;letter-spacing:2px;color:var(--muted);margin-bottom:16px;padding-bottom:8px;border-bottom:1px solid var(--border);}\n'
    + '.charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px;}\n'
    + '.card{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:24px;}\n'
    + '.card-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px;}\n'
    + '.card-title{font-size:14px;font-weight:600;color:var(--white);}\n'
    + '.card-sub{font-size:12px;color:var(--muted);margin-bottom:18px;}\n'
    + '.signal{display:inline-block;font-size:11px;padding:3px 10px;border-radius:3px;font-weight:500;letter-spacing:0.3px;}\n'
    + '.sig-pos{background:var(--teal-dim);color:var(--teal);border:1px solid rgba(62,207,178,0.2);}\n'
    + '.sig-neu{background:var(--gold-dim);color:var(--gold);border:1px solid rgba(212,168,67,0.2);}\n'
    + '.sig-neg{background:var(--red-dim);color:var(--red);border:1px solid rgba(224,92,92,0.2);}\n'
    + '.legend{display:flex;gap:16px;margin-bottom:14px;font-size:11px;color:var(--muted);}\n'
    + '.legend-dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:4px;}\n'
    + '.ratio-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;margin-bottom:32px;}\n'
    + '.ratio-big{font-family:"Bebas Neue",sans-serif;font-size:56px;line-height:1;}\n'
    + '.ratio-label{font-size:12px;color:var(--muted);margin-top:6px;}\n'
    + '.ratio-note{font-size:12px;color:var(--text);margin-top:10px;line-height:1.5;}\n'
    + 'canvas{max-height:200px;}\n'
    + '<\/style>\n'
    + '<\/head>\n'
    + '<body>\n'
    + '<div class="header">\n'
    + '  <div><h1>RALEIGH <span>MARKET<\/span> DASHBOARD<\/h1><p>Triangle MLS — Year-over-Year Analysis<\/p><\/div>\n'
    + '  <div class="badge"><div class="mo">' + latestLabel + ' ' + latestYear + '<\/div><div class="sub">Latest Data<\/div><\/div>\n'
    + '<\/div>\n'
    + '<div class="main">\n'
    + '  <div class="auto-bar">⚡ <strong>Live data<\/strong> — updates automatically each time you submit your Market Update form. Comparing this year vs. same month last year.<\/div>\n'
    + '  <div class="section-label">Key Metrics — This Month vs. Same Month Last Year<\/div>\n'
    + '  <div class="kpi-row" id="kpis"><\/div>\n'
    + '  <div class="section-label">Year-Over-Year Trends<\/div>\n'
    + '  <div class="charts-grid">\n'
    + '    <div class="card"><div class="card-header"><div class="card-title">NEW LISTINGS<\/div><\/div><div class="card-sub">Are more homes coming to market?<\/div><div class="legend"><span><span class="legend-dot" style="background:#D4A843"><\/span>This Year<\/span><span><span class="legend-dot" style="background:#7A8096"><\/span>Last Year<\/span><\/div><canvas id="c_listings"><\/canvas><\/div>\n'
    + '    <div class="card"><div class="card-header"><div class="card-title">NEW PENDINGS<\/div><\/div><div class="card-sub">Are buyers pulling the trigger?<\/div><div class="legend"><span><span class="legend-dot" style="background:#3ECFB2"><\/span>This Year<\/span><span><span class="legend-dot" style="background:#7A8096"><\/span>Last Year<\/span><\/div><canvas id="c_pendings"><\/canvas><\/div>\n'
    + '    <div class="card"><div class="card-header"><div class="card-title">DAYS ON MARKET<\/div><\/div><div class="card-sub">Is the market speeding up or slowing down?<\/div><div class="legend"><span><span class="legend-dot" style="background:#D4A843"><\/span>This Year<\/span><span><span class="legend-dot" style="background:#7A8096"><\/span>Last Year<\/span><\/div><canvas id="c_dom"><\/canvas><\/div>\n'
    + '    <div class="card"><div class="card-header"><div class="card-title">EXPIRED LISTINGS<\/div><\/div><div class="card-sub">Are sellers mispricing?<\/div><div class="legend"><span><span class="legend-dot" style="background:#E05C5C"><\/span>This Year<\/span><span><span class="legend-dot" style="background:#7A8096"><\/span>Last Year<\/span><\/div><canvas id="c_expired"><\/canvas><\/div>\n'
    + '    <div class="card"><div class="card-header"><div class="card-title">% OF LIST PRICE RECEIVED<\/div><\/div><div class="card-sub">Negotiating power benchmark<\/div><div class="legend"><span><span class="legend-dot" style="background:#D4A843"><\/span>This Year<\/span><span><span class="legend-dot" style="background:#7A8096"><\/span>Last Year<\/span><\/div><canvas id="c_pol"><\/canvas><\/div>\n'
    + '    <div class="card"><div class="card-header"><div class="card-title">% CLOSED ABOVE LIST<\/div><\/div><div class="card-sub">How competitive are offers?<\/div><div class="legend"><span><span class="legend-dot" style="background:#3ECFB2"><\/span>This Year<\/span><span><span class="legend-dot" style="background:#7A8096"><\/span>Last Year<\/span><\/div><canvas id="c_pal"><\/canvas><\/div>\n'
    + '  <\/div>\n'
    + '  <div class="section-label">Market Signals<\/div>\n'
    + '  <div class="ratio-row" id="signals"><\/div>\n'
    + '<\/div>\n'
    + '<script>\n'
    + 'var DATA = ' + dataJson + ';\n'
    + '\n'
    + 'function parsePct(v){if(v===null||v===undefined)return 0;return v>1?v/100:v;}\n'
    + '\n'
    + '// Build month labels\n'
    + 'var MN=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];\n'
    + 'var labels=DATA.map(function(d){var p=d.month.split("-");return MN[parseInt(p[1])-1]+" \'"+p[0].slice(2);});\n'
    + '\n'
    + '// Latest data point\n'
    + 'var L=DATA.length>0?DATA[DATA.length-1]:null;\n'
    + '\n'
    + 'function yoyPct(ty,ly){\n'
    + '  if(!ly||ly===0)return null;\n'
    + '  var d=(ty-ly)/ly*100;\n'
    + '  return (d>=0?"+":"")+d.toFixed(1)+"%";\n'
    + '}\n'
    + '\n'
    + 'function yoyClass(ty,ly,lowerIsBetter){\n'
    + '  if(!ly)return "";\n'
    + '  var up=ty>ly;\n'
    + '  var good=lowerIsBetter?!up:up;\n'
    + '  return good?"up":"down";\n'
    + '}\n'
    + '\n'
    + '// KPI Cards\n'
    + 'if(L){\n'
    + '  var kpis=[\n'
    + '    {label:"New Listings",ty:L.nl_ty,ly:L.nl_ly,fmt:function(v){return Math.round(v).toLocaleString();},low:false},\n'
    + '    {label:"New Pendings",ty:L.np_ty,ly:L.np_ly,fmt:function(v){return Math.round(v).toLocaleString();},low:false},\n'
    + '    {label:"Days on Market",ty:L.dom_ty,ly:L.dom_ly,fmt:function(v){return v+" days";},low:true},\n'
    + '    {label:"% of List Received",ty:L.pol_ty,ly:L.pol_ly,fmt:function(v){return (v*100).toFixed(1)+"%";},low:false},\n'
    + '    {label:"% Closed Above List",ty:L.pal_ty,ly:L.pal_ly,fmt:function(v){return (v*100).toFixed(1)+"%";},low:false}\n'
    + '  ];\n'
    + '  var kpiHtml="";\n'
    + '  for(var i=0;i<kpis.length;i++){\n'
    + '    var k=kpis[i];\n'
    + '    var chg=yoyPct(k.ty,k.ly);\n'
    + '    var cls=yoyClass(k.ty,k.ly,k.low);\n'
    + '    kpiHtml+=\'<div class="kpi">\';\n'
    + '    kpiHtml+=\'<div class="kpi-label">\'+k.label+\'<\/div>\';\n'
    + '    kpiHtml+=\'<div class="kpi-value">\'+k.fmt(k.ty)+\'<\/div>\';\n'
    + '    if(chg)kpiHtml+=\'<div class="kpi-delta"><span class="\'+cls+\'">\'+chg+\' vs last year<\/span><\/div>\';\n'
    + '    kpiHtml+=\'<div class="kpi-compare">Last year: \'+k.fmt(k.ly)+\'<\/div>\';\n'
    + '    kpiHtml+=\'<\/div>\';\n'
    + '  }\n'
    + '  document.getElementById("kpis").innerHTML=kpiHtml;\n'
    + '}\n'
    + '\n'
    + '// Chart defaults\n'
    + 'var cd={\n'
    + '  responsive:true,maintainAspectRatio:true,\n'
    + '  plugins:{legend:{display:false},tooltip:{backgroundColor:"#1E2435",titleColor:"#E8E4DC",bodyColor:"#7A8096",borderColor:"rgba(255,255,255,0.07)",borderWidth:1}},\n'
    + '  scales:{x:{ticks:{color:"#7A8096",font:{size:10}},grid:{color:"rgba(255,255,255,0.04)"}},y:{ticks:{color:"#7A8096",font:{size:10}},grid:{color:"rgba(255,255,255,0.04)"}}}\n'
    + '};\n'
    + '\n'
    + 'function pctTicks(cd){\n'
    + '  return Object.assign({},cd,{scales:Object.assign({},cd.scales,{y:Object.assign({},cd.scales.y,{ticks:Object.assign({},cd.scales.y.ticks,{callback:function(v){return v+"%";}})})})});\n'
    + '}\n'
    + '\n'
    + 'function makeBar(id,tyData,lyData,tyColor){\n'
    + '  new Chart(document.getElementById(id),{\n'
    + '    type:"bar",\n'
    + '    data:{labels:labels,datasets:[\n'
    + '      {label:"This Year",data:tyData,backgroundColor:tyColor,borderRadius:2},\n'
    + '      {label:"Last Year",data:lyData,backgroundColor:"rgba(122,128,150,0.3)",borderRadius:2}\n'
    + '    ]},\n'
    + '    options:Object.assign({},cd,{plugins:Object.assign({},cd.plugins,{legend:{display:false}})})\n'
    + '  });\n'
    + '}\n'
    + '\n'
    + 'function makeLine(id,tyData,lyData,tyColor,opts){\n'
    + '  var options=opts||cd;\n'
    + '  new Chart(document.getElementById(id),{\n'
    + '    type:"line",\n'
    + '    data:{labels:labels,datasets:[\n'
    + '      {label:"This Year",data:tyData,borderColor:tyColor,backgroundColor:"rgba(0,0,0,0)",tension:0.4,pointRadius:4,pointBackgroundColor:tyColor,borderWidth:2},\n'
    + '      {label:"Last Year",data:lyData,borderColor:"rgba(122,128,150,0.5)",backgroundColor:"rgba(0,0,0,0)",tension:0.4,pointRadius:3,pointBackgroundColor:"rgba(122,128,150,0.5)",borderWidth:1.5,borderDash:[4,4]}\n'
    + '    ]},\n'
    + '    options:options\n'
    + '  });\n'
    + '}\n'
    + '\n'
    + 'makeBar("c_listings",DATA.map(function(d){return d.nl_ty;}),DATA.map(function(d){return d.nl_ly;}),"rgba(212,168,67,0.7)");\n'
    + 'makeBar("c_pendings",DATA.map(function(d){return d.np_ty;}),DATA.map(function(d){return d.np_ly;}),"rgba(62,207,178,0.7)");\n'
    + 'makeBar("c_dom",DATA.map(function(d){return d.dom_ty;}),DATA.map(function(d){return d.dom_ly;}),"rgba(212,168,67,0.7)");\n'
    + 'makeBar("c_expired",DATA.map(function(d){return d.exp_ty;}),DATA.map(function(d){return d.exp_ly;}),"rgba(224,92,92,0.7)");\n'
    + 'makeLine("c_pol",DATA.map(function(d){return (d.pol_ty*100).toFixed(1);}),DATA.map(function(d){return (d.pol_ly*100).toFixed(1);}),"#D4A843",pctTicks(cd));\n'
    + 'makeLine("c_pal",DATA.map(function(d){return (d.pal_ty*100).toFixed(1);}),DATA.map(function(d){return (d.pal_ly*100).toFixed(1);}),"#3ECFB2",pctTicks(cd));\n'
    + '\n'
    + '// Market Signals\n'
    + 'if(L){\n'
    + '  var ratio=L.nl_ty>0?(L.np_ty/L.nl_ty):0;\n'
    + '  var ratioLY=L.nl_ly>0?(L.np_ly/L.nl_ly):0;\n'
    + '  var ratioSig=ratio>1?"Seller\'s Market":ratio>0.8?"Balanced":"Buyer\'s Market";\n'
    + '  var ratioClass=ratio>1?"sig-pos":ratio>0.8?"sig-neu":"sig-neg";\n'
    + '  var expChg=L.exp_ly>0?((L.exp_ty-L.exp_ly)/L.exp_ly*100).toFixed(0):0;\n'
    + '  var csChg=L.cs_ly>0?((L.cs_ty-L.cs_ly)/L.cs_ly*100).toFixed(0):0;\n'
    + '  var sigHtml="";\n'
    + '  sigHtml+=\'<div class="card">\';\n'
    + '  sigHtml+=\'<div class="ratio-big" style="color:\'+(ratio>1?"var(--teal)":ratio>0.8?"var(--gold)":"var(--red)")+\'">\'+ratio.toFixed(2)+\'<\/div>\';\n'
    + '  sigHtml+=\'<div class="ratio-label">Pending \/ Listing Ratio<\/div>\';\n'
    + '  sigHtml+=\'<div style="margin-top:8px"><span class="signal \'+ratioClass+\'">\'+ratioSig+\'<\/span><\/div>\';\n'
    + '  sigHtml+=\'<div class="ratio-note">Last year this month: \'+ratioLY.toFixed(2)+\'. Above 1.0 means more buyers than available homes.<\/div>\';\n'
    + '  sigHtml+=\'<\/div>\';\n'
    + '  sigHtml+=\'<div class="card">\';\n'
    + '  sigHtml+=\'<div class="ratio-big" style="color:\'+(L.exp_ty>L.exp_ly?"var(--red)":"var(--teal)")+\'">\'+(expChg>0?"+":"")+expChg+\'%<\/div>\';\n'
    + '  sigHtml+=\'<div class="ratio-label">YoY Change in Expired Listings<\/div>\';\n'
    + '  sigHtml+=\'<div style="margin-top:8px"><span class="signal \'+(L.exp_ty>L.exp_ly?"sig-neg":"sig-pos")+\'">\'+( L.exp_ty>L.exp_ly?"Rising expireds":"Improving")+"<\/span><\/div>";\n'
    + '  sigHtml+=\'<div class="ratio-note">This year: \'+L.exp_ty.toLocaleString()+\' expired. Last year: \'+L.exp_ly.toLocaleString()+\'.<\/div>\';\n'
    + '  sigHtml+=\'<\/div>\';\n'
    + '  sigHtml+=\'<div class="card">\';\n'
    + '  sigHtml+=\'<div class="ratio-big" style="color:\'+(L.cs_ty>L.cs_ly?"var(--teal)":"var(--red)")+\'">\'+(csChg>0?"+":"")+csChg+\'%<\/div>\';\n'
    + '  sigHtml+=\'<div class="ratio-label">YoY Change in Closed Sales<\/div>\';\n'
    + '  sigHtml+=\'<div style="margin-top:8px"><span class="signal \'+(L.cs_ty>L.cs_ly?"sig-pos":"sig-neg")+\'">\'+( L.cs_ty>L.cs_ly?"More closings":"Fewer closings")+"<\/span><\/div>";\n'
    + '  sigHtml+=\'<div class="ratio-note">This year: \'+L.cs_ty.toLocaleString()+\' closed. Last year: \'+L.cs_ly.toLocaleString()+\'.<\/div>\';\n'
    + '  sigHtml+=\'<\/div>\';\n'
    + '  document.getElementById("signals").innerHTML=sigHtml;\n'
    + '}\n'
    + '<\/script>\n'
    + '<\/body>\n'
    + '<\/html>';

  return html;
}

<#
.SYNOPSIS
    PingWatch - portable ping surveillance of two IPs with statistics.

.DESCRIPTION
    Pings two targets simultaneously (one internal, one external for checking internet outage during live debug) (default: every 1 s for 1 h 15 min),
    shows live results, logs every reply to CSV and writes a summary with
    loss %, min/avg/max/p95 latency, jitter, longest outage and outage list.
    Works on Windows PowerShell 5.1 and PowerShell 7+. No admin rights needed.
    Press Ctrl+C at any time: the summary is still produced.

.EXAMPLE
    .\PingWatch.ps1 -Target1 192.168.1.1 -Target2 1.1.1.1

.EXAMPLE
    .\PingWatch.ps1 -Target1 192.168.1.1 -Target2 1.1.1.1 -DurationMinutes 75 -IntervalMs 500
#>
param(
    [string]$Target1          = "192.168.1.1",
    [string]$Target2          = "1.1.1.1",
    [int]   $DurationMinutes  = 75,     # 1 h 15 min
    [int]   $IntervalMs       = 1000,   # time between ping rounds
    [int]   $TimeoutMs        = 1000,   # reply timeout per ping
    [int]   $OutageThreshold  = 3,      # consecutive losses counted as an outage
    [string]$OutputDir        = ""
)

if (-not $OutputDir) { $OutputDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path } }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath = Join-Path $OutputDir ("PingWatch_{0}.csv" -f $stamp)
$sumPath = Join-Path $OutputDir ("PingWatch_{0}_summary.txt" -f $stamp)
$repPath = Join-Path $OutputDir ("PingWatch_{0}_report.html" -f $stamp)

# ---------- HTML report template (ASCII only so PS 5.1 reads it in any encoding) ----------
# __CSV__ is replaced by the CSV log, __OUTAGE__ by $OutageThreshold
$reportTemplate = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PingWatch Report</title>
<style>
:root {
  color-scheme: light;
  --page: #f9f9f7; --surface: #fcfcfb; --ink: #0b0b0b; --ink2: #52514e; --muted: #898781;
  --grid: #e1e0d9; --axis: #c3c2b7; --ring: rgba(11,11,11,0.10);
  --s1: #2a78d6; --s2: #eb6834;
  --good: #0ca30c; --serious: #ec835a; --critical: #d03b3b;
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --page: #0d0d0d; --surface: #1a1a19; --ink: #fff; --ink2: #c3c2b7;
    --grid: #2c2c2a; --axis: #383835; --ring: rgba(255,255,255,0.10);
    --s1: #3987e5; --s2: #d95926;
  }
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --page: #0d0d0d; --surface: #1a1a19; --ink: #fff; --ink2: #c3c2b7;
  --grid: #2c2c2a; --axis: #383835; --ring: rgba(255,255,255,0.10);
  --s1: #3987e5; --s2: #d95926;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--page); color: var(--ink); font: 14px/1.45 system-ui, -apple-system, "Segoe UI", sans-serif; }
main { max-width: 1040px; margin: 0 auto; padding: 32px 16px 48px; }
header { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: end; gap: 12px; margin-bottom: 24px; }
h1 { font-size: 24px; margin: 0 0 4px; letter-spacing: -0.01em; }
h2 { font-size: 15px; margin: 0 0 2px; }
.sub { color: var(--ink2); margin: 0; }
.card { background: var(--surface); border: 1px solid var(--ring); border-radius: 12px; padding: 20px; margin-bottom: 16px; }
.cap { color: var(--muted); font-size: 12px; margin: 0 0 14px; }
.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 16px; }
.tile { background: var(--surface); border: 1px solid var(--ring); border-radius: 12px; padding: 16px 20px; }
.tile .who { display: flex; align-items: center; gap: 8px; color: var(--ink2); font-size: 13px; font-weight: 600; }
.tile .dot { width: 10px; height: 10px; border-radius: 50%; }
.tile .big { font-size: 34px; font-weight: 650; margin: 6px 0 2px; letter-spacing: -0.02em; }
.tile .big small { font-size: 15px; color: var(--ink2); font-weight: 500; }
.tile .meta { color: var(--ink2); font-size: 12.5px; display: grid; grid-template-columns: auto auto; gap: 2px 12px; justify-content: start; }
.tile .meta b { color: var(--ink); font-weight: 600; font-variant-numeric: tabular-nums; }
.badge { display: inline-flex; align-items: center; gap: 5px; font-size: 12px; font-weight: 600; padding: 2px 8px; border-radius: 99px; border: 1px solid var(--ring); color: var(--ink); }
.legend { display: flex; flex-wrap: wrap; gap: 16px; font-size: 12.5px; color: var(--ink2); margin-bottom: 10px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; }
.sw { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }
.ln { width: 16px; height: 2px; border-radius: 2px; display: inline-block; }
svg { display: block; width: 100%; height: auto; overflow: visible; }
svg text { fill: var(--muted); font-size: 11px; font-variant-numeric: tabular-nums; }
svg .lbl { fill: var(--ink2); font-size: 12px; font-weight: 600; }
.tip { position: fixed; pointer-events: none; background: var(--surface); color: var(--ink); border: 1px solid var(--ring); border-radius: 8px; padding: 8px 10px; font-size: 12px; box-shadow: 0 4px 16px rgba(0,0,0,.12); opacity: 0; transition: opacity .08s; z-index: 9; white-space: nowrap; }
.tip .t { color: var(--muted); margin-bottom: 4px; }
.tip .r { display: flex; align-items: center; gap: 6px; }
details { margin-top: 4px; }
summary { cursor: pointer; color: var(--ink2); font-size: 13px; }
table { border-collapse: collapse; width: 100%; font-size: 12.5px; margin-top: 10px; font-variant-numeric: tabular-nums; }
th, td { text-align: left; padding: 4px 8px; border-bottom: 1px solid var(--grid); }
th { color: var(--muted); font-weight: 600; }
.tablewrap { max-height: 320px; overflow: auto; }
label.load { cursor: pointer; font-size: 13px; color: var(--ink2); border: 1px solid var(--ring); border-radius: 8px; padding: 6px 12px; background: var(--surface); }
label.load input { display: none; }
footer { color: var(--muted); font-size: 12px; margin-top: 24px; }
</style>
</head>
<body>
<main>
  <header>
    <div>
      <h1>PingWatch report</h1>
      <p class="sub" id="range"></p>
    </div>
    <label class="load">Load another PingWatch CSV&#x2026;<input type="file" accept=".csv" id="file"></label>
  </header>

  <div class="tiles" id="tiles"></div>

  <section class="card">
    <h2>Reply status per ping</h2>
    <p class="cap">One cell per ping. Each row is one target.</p>
    <div class="legend">
      <span><i class="sw" style="background:var(--good)"></i>&#x2713; Success</span>
      <span><i class="sw" style="background:var(--serious)"></i>&#x23f1; Timed out</span>
      <span><i class="sw" style="background:var(--critical)"></i>&#x2715; Unreachable / error</span>
    </div>
    <div id="strip"></div>
  </section>

  <section class="card">
    <h2>Round-trip time</h2>
    <p class="cap">Milliseconds, successful replies only. Gaps are lost pings.</p>
    <div class="legend" id="lineLegend"></div>
    <div id="line"></div>
  </section>

  <section class="card">
    <h2>Outages</h2>
    <p class="cap">Runs of __OUTAGE__ or more consecutive lost pings.</p>
    <div id="outages"></div>
    <details>
      <summary>Show raw log</summary>
      <div class="tablewrap"><table id="raw"></table></div>
    </details>
  </section>

  <footer>Generated from PingWatch.ps1 CSV output.</footer>
</main>
<div class="tip" id="tip"></div>

<script type="text/csv" id="data">
__CSV__
</script>
<script>
const OUTAGE = __OUTAGE__;
const SERIES = ['var(--s1)', 'var(--s2)'];
const tip = document.getElementById('tip');
const NS = 'http://www.w3.org/2000/svg';
const el = (tag, attrs = {}, parent) => {
  const e = document.createElementNS(NS, tag);
  for (const k in attrs) e.setAttribute(k, attrs[k]);
  if (parent) parent.appendChild(e);
  return e;
};
const fmtT = d => d.toTimeString().slice(0, 8);
const statusKind = s => s === 'Success' ? 'good' : s === 'TimedOut' ? 'serious' : 'critical';
const statusIcon = { good: '\u2713', serious: '\u23f1', critical: '\u2715' };

function showTip(ev, html) {
  tip.innerHTML = html;
  tip.style.opacity = 1;
  const w = tip.offsetWidth, h = tip.offsetHeight;
  let x = ev.clientX + 14, y = ev.clientY + 14;
  if (x + w > innerWidth - 8) x = ev.clientX - w - 14;
  if (y + h > innerHeight - 8) y = ev.clientY - h - 14;
  tip.style.left = x + 'px'; tip.style.top = y + 'px';
}
const hideTip = () => tip.style.opacity = 0;

function parse(text) {
  const lines = text.trim().split(/\r?\n/).slice(1);
  const rows = lines.map(l => {
    const [ts, target, status, rtt] = l.split(',');
    return { t: new Date(ts.replace(' ', 'T')), target, status, rtt: rtt === '' || rtt == null ? null : +rtt };
  });
  const targets = [...new Set(rows.map(r => r.target))];
  const times = [...new Set(rows.map(r => +r.t))].sort((a, b) => a - b);
  return { rows, targets, times };
}

function stats(rows) {
  const ok = rows.filter(r => r.status === 'Success');
  const rtts = ok.map(r => r.rtt).sort((a, b) => a - b);
  let jit = 0;
  for (let i = 1; i < ok.length; i++) jit += Math.abs(ok[i].rtt - ok[i - 1].rtt);
  const outs = []; let run = null;
  rows.forEach(r => {
    if (r.status !== 'Success') { run = run || { start: r.t, n: 0 }; run.n++; run.last = r.t; }
    else { if (run && run.n >= OUTAGE) outs.push({ ...run, end: r.t }); run = null; }
  });
  if (run && run.n >= OUTAGE) outs.push({ ...run, end: null });
  return {
    sent: rows.length, recv: ok.length, loss: rows.length ? 100 * (rows.length - ok.length) / rows.length : 0,
    min: rtts[0], max: rtts[rtts.length - 1], avg: rtts.length ? rtts.reduce((a, b) => a + b, 0) / rtts.length : null,
    p95: rtts.length ? rtts[Math.max(0, Math.ceil(0.95 * rtts.length) - 1)] : null,
    jitter: ok.length > 1 ? jit / (ok.length - 1) : null, outs
  };
}

function render(text) {
  const { rows, targets, times } = parse(text);
  const by = Object.fromEntries(targets.map(t => [t, rows.filter(r => r.target === t)]));
  const st = Object.fromEntries(targets.map(t => [t, stats(by[t])]));
  const t0 = new Date(times[0]), t1 = new Date(times[times.length - 1]);
  const durS = Math.round((t1 - t0) / 1000) + 1;
  const both = times.filter(tm => targets.every(t => by[t].some(r => +r.t === tm && r.status !== 'Success'))).length;
  document.getElementById('range').textContent =
    `${t0.toLocaleDateString()} \u00b7 ${fmtT(t0)} \u2192 ${fmtT(t1)} \u00b7 ${Math.floor(durS / 60)} min ${durS % 60} s \u00b7 ${times.length} rounds \u00b7 both targets down in ${both} rounds`;

  // tiles
  const n = v => v == null ? '\u2013' : (Math.round(v * 10) / 10) + ' ms';
  document.getElementById('tiles').innerHTML = targets.map((t, i) => {
    const s = st[t];
    const kind = s.loss === 0 ? 'good' : s.loss >= 50 ? 'critical' : 'serious';
    const label = { good: 'Healthy', serious: 'Degraded', critical: 'Down' }[kind];
    return `<div class="tile">
      <div class="who"><i class="dot" style="background:${SERIES[i]}"></i>${t}
        <span class="badge" style="margin-left:auto"><span style="color:var(--${kind})">${statusIcon[kind]}</span>${label}</span></div>
      <div class="big">${s.loss.toFixed(1)}<small> % loss</small></div>
      <div class="meta">
        <span>Sent / received</span><b>${s.sent} / ${s.recv}</b>
        <span>Min / avg / max</span><b>${s.avg == null ? '\u2013' : `${s.min} / ${s.avg.toFixed(1)} / ${s.max} ms`}</b>
        <span>p95 \u00b7 jitter</span><b>${n(s.p95)} \u00b7 ${n(s.jitter)}</b>
        <span>Outages</span><b>${s.outs.length}</b>
      </div></div>`;
  }).join('');

  // status strip
  const W = 1000, labW = 92, rowH = 22, gap = 6, padB = 22;
  const H = targets.length * (rowH + gap) + padB;
  const strip = document.getElementById('strip'); strip.innerHTML = '';
  const s1 = el('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Reply status per ping' }, strip);
  const cw = (W - labW) / times.length;
  const xOf = tm => labW + times.indexOf(tm) * cw;
  targets.forEach((t, ti) => {
    const y = ti * (rowH + gap);
    el('text', { x: labW - 10, y: y + rowH / 2 + 4, 'text-anchor': 'end', class: 'lbl' }, s1).textContent = t;
    by[t].forEach(r => {
      const k = statusKind(r.status);
      const c = el('rect', { x: xOf(+r.t) + 1, y, width: Math.max(1, cw - 2), height: rowH, rx: Math.min(3, cw / 3), fill: `var(--${k})` }, s1);
      c.addEventListener('mousemove', e => showTip(e, `<div class="t">${fmtT(r.t)}</div><div class="r"><b>${t}</b></div><div class="r"><span style="color:var(--${k})">${statusIcon[k]}</span>${r.status}${r.rtt != null ? ' \u00b7 ' + r.rtt + ' ms' : ''}</div>`));
      c.addEventListener('mouseleave', hideTip);
    });
  });
  const ticks = 6;
  for (let i = 0; i <= ticks; i++) {
    const idx = Math.round(i * (times.length - 1) / ticks);
    el('text', { x: labW + idx * cw + cw / 2, y: H - 4, 'text-anchor': i === 0 ? 'start' : i === ticks ? 'end' : 'middle' }, s1).textContent = fmtT(new Date(times[idx]));
  }

  // line chart
  document.getElementById('lineLegend').innerHTML = targets.map((t, i) =>
    `<span><i class="ln" style="background:${SERIES[i]}"></i>${t}${st[t].recv === 0 ? ' (no replies)' : ''}</span>`).join('');
  const LH = 260, pl = 44, pr = 12, pt = 12, pb = 26;
  const line = document.getElementById('line'); line.innerHTML = '';
  const s2 = el('svg', { viewBox: `0 0 ${W} ${LH}`, role: 'img', 'aria-label': 'Round-trip time chart' }, line);
  const maxR = Math.max(4, ...rows.filter(r => r.rtt != null).map(r => r.rtt));
  const step = [1, 2, 5, 10, 20, 25, 50, 100, 200, 250, 500, 1000].find(s => maxR / s <= 5) || 1000;
  const yMax = Math.ceil(maxR / step) * step;
  const X = tm => pl + (times.length === 1 ? 0 : (times.indexOf(tm) / (times.length - 1)) * (W - pl - pr));
  const Y = v => pt + (1 - v / yMax) * (LH - pt - pb);
  for (let v = 0; v <= yMax; v += step) {
    el('line', { x1: pl, x2: W - pr, y1: Y(v), y2: Y(v), stroke: v === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }, s2);
    el('text', { x: pl - 8, y: Y(v) + 4, 'text-anchor': 'end' }, s2).textContent = v;
  }
  for (let i = 0; i <= ticks; i++) {
    const idx = Math.round(i * (times.length - 1) / ticks);
    el('text', { x: X(times[idx]), y: LH - 6, 'text-anchor': i === 0 ? 'start' : i === ticks ? 'end' : 'middle' }, s2).textContent = fmtT(new Date(times[idx]));
  }
  // lost-ping ticks under baseline
  targets.forEach((t, i) => {
    let d = '', pen = false;
    by[t].forEach(r => {
      if (r.rtt == null) { pen = false; el('rect', { x: X(+r.t) - 1, y: LH - pb + 2 + i * 4, width: 2, height: 3, fill: SERIES[i], opacity: .7 }, s2); return; }
      d += `${pen ? 'L' : 'M'}${X(+r.t).toFixed(1)},${Y(r.rtt).toFixed(1)}`; pen = true;
    });
    if (d) el('path', { d, fill: 'none', stroke: SERIES[i], 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }, s2);
  });
  const cross = el('line', { y1: pt, y2: LH - pb, stroke: 'var(--axis)', 'stroke-width': 1, opacity: 0 }, s2);
  const dots = targets.map((t, i) => el('circle', { r: 4.5, fill: SERIES[i], stroke: 'var(--surface)', 'stroke-width': 2, opacity: 0 }, s2));
  const hit = el('rect', { x: pl, y: 0, width: W - pl - pr, height: LH, fill: 'transparent' }, s2);
  hit.addEventListener('mousemove', e => {
    const b = s2.getBoundingClientRect();
    const sx = (e.clientX - b.left) * W / b.width;
    const idx = Math.max(0, Math.min(times.length - 1, Math.round((sx - pl) / (W - pl - pr) * (times.length - 1))));
    const tm = times[idx];
    cross.setAttribute('x1', X(tm)); cross.setAttribute('x2', X(tm)); cross.setAttribute('opacity', 1);
    let html = `<div class="t">${fmtT(new Date(tm))}</div>`;
    targets.forEach((t, i) => {
      const r = by[t].find(r => +r.t === tm);
      if (r && r.rtt != null) { dots[i].setAttribute('cx', X(tm)); dots[i].setAttribute('cy', Y(r.rtt)); dots[i].setAttribute('opacity', 1); }
      else dots[i].setAttribute('opacity', 0);
      html += `<div class="r"><i class="ln" style="background:${SERIES[i]}"></i>${t}: <b>${r ? (r.rtt != null ? r.rtt + ' ms' : r.status) : '\u2013'}</b></div>`;
    });
    showTip(e, html);
  });
  hit.addEventListener('mouseleave', () => { hideTip(); cross.setAttribute('opacity', 0); dots.forEach(d => d.setAttribute('opacity', 0)); });

  // outages + raw table
  const outs = targets.flatMap(t => st[t].outs.map(o => ({ t, ...o })));
  document.getElementById('outages').innerHTML = outs.length
    ? `<table><tr><th>Target</th><th>Start</th><th>End</th><th>Lost pings</th></tr>${outs.map(o =>
        `<tr><td>${o.t}</td><td>${fmtT(o.start)}</td><td>${o.end ? fmtT(o.end) : '<span style="color:var(--critical)">\u2715</span> still down'}</td><td>${o.n}</td></tr>`).join('')}</table>`
    : '<p class="sub"><span style="color:var(--good)">\u2713</span> No outages recorded.</p>';
  document.getElementById('raw').innerHTML = '<tr><th>Time</th><th>Target</th><th>Status</th><th>RTT (ms)</th></tr>' +
    rows.map(r => `<tr><td>${fmtT(r.t)}</td><td>${r.target}</td><td>${r.status}</td><td>${r.rtt ?? ''}</td></tr>`).join('');
}

render(document.getElementById('data').textContent);
document.getElementById('file').addEventListener('change', e => {
  const f = e.target.files[0]; if (!f) return;
  f.text().then(render);
});
</script>
</body>
</html>
'@

# ---------- statistics helpers ----------
function New-Stat([string]$t) {
    [pscustomobject]@{
        Target      = $t
        Sent        = 0
        Received    = 0
        Rtts        = New-Object System.Collections.Generic.List[double]
        LastRtt     = $null
        JitterSum   = 0.0
        JitterN     = 0
        Streak      = 0
        StreakStart = $null
        MaxStreak   = 0
        Outages     = New-Object System.Collections.Generic.List[object]
    }
}

function Update-Stat($s, [bool]$ok, [double]$rtt, [datetime]$now) {
    $s.Sent++
    if ($ok) {
        $s.Received++
        $s.Rtts.Add($rtt)
        if ($null -ne $s.LastRtt) { $s.JitterSum += [math]::Abs($rtt - $s.LastRtt); $s.JitterN++ }
        $s.LastRtt = $rtt
        if ($s.Streak -ge $OutageThreshold) {
            $s.Outages.Add([pscustomobject]@{ Start = $s.StreakStart; End = $now; Lost = $s.Streak })
        }
        $s.Streak = 0
    } else {
        if ($s.Streak -eq 0) { $s.StreakStart = $now }
        $s.Streak++
        if ($s.Streak -gt $s.MaxStreak) { $s.MaxStreak = $s.Streak }
    }
}

function Get-Summary {
    $L = New-Object System.Collections.Generic.List[string]
    $L.Add("==================== PingWatch summary ====================")
    $L.Add(("Start      : {0}" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss')))
    $L.Add(("End        : {0}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))
    $L.Add(("Duration   : {0:hh\:mm\:ss} (planned {1} min)" -f ((Get-Date) - $startTime), $DurationMinutes))
    $L.Add(("Interval   : {0} ms   Timeout: {1} ms   Outage = {2}+ consecutive losses" -f $IntervalMs, $TimeoutMs, $OutageThreshold))
    $L.Add(("Rounds where BOTH targets failed: {0}  (suggests a local/network-side problem)" -f $script:bothDown))
    $L.Add("")

    foreach ($s in $stats) {
        $lost = $s.Sent - $s.Received
        $loss = if ($s.Sent -gt 0) { 100.0 * $lost / $s.Sent } else { 0 }
        $L.Add(("--- {0} ---" -f $s.Target))
        $L.Add(("Sent / Received / Lost : {0} / {1} / {2}  ({3:N2} % loss)" -f $s.Sent, $s.Received, $lost, $loss))
        if ($s.Rtts.Count -gt 0) {
            $m = $s.Rtts | Measure-Object -Minimum -Maximum -Average
            $sorted = $s.Rtts.ToArray(); [Array]::Sort($sorted)
            $p95 = $sorted[[math]::Max(0, [math]::Ceiling(0.95 * $sorted.Length) - 1)]
            $jit = if ($s.JitterN -gt 0) { $s.JitterSum / $s.JitterN } else { 0 }
            $L.Add(("Latency min/avg/max    : {0} / {1:N1} / {2} ms" -f $m.Minimum, $m.Average, $m.Maximum))
            $L.Add(("Latency p95            : {0} ms" -f $p95))
            $L.Add(("Jitter (avg delta)     : {0:N1} ms" -f $jit))
        } else {
            $L.Add("Latency                : no replies received")
        }
        $L.Add(("Longest loss streak    : {0} pings (~{1:N0} s)" -f $s.MaxStreak, ($s.MaxStreak * $IntervalMs / 1000)))

        $outs = $s.Outages.ToArray()
        if ($s.Streak -ge $OutageThreshold) {
            $outs += [pscustomobject]@{ Start = $s.StreakStart; End = $null; Lost = $s.Streak }
        }
        $L.Add(("Outages                : {0}" -f $outs.Count))
        foreach ($o in $outs) {
            $endTxt = if ($o.End) { $o.End.ToString('HH:mm:ss') } else { 'still down' }
            $L.Add(("   {0} -> {1}   {2} lost" -f $o.Start.ToString('HH:mm:ss'), $endTxt, $o.Lost))
        }
        $L.Add("")
    }
    $L.Add(("Detailed log: {0}" -f $csvPath))
    return $L
}

# ---------- main ----------
$stats     = @((New-Stat $Target1), (New-Stat $Target2))
$pingers   = @((New-Object System.Net.NetworkInformation.Ping), (New-Object System.Net.NetworkInformation.Ping))
$script:bothDown = 0
$startTime = Get-Date
$endTime   = $startTime.AddMinutes($DurationMinutes)

$writer = New-Object System.IO.StreamWriter($csvPath, $false, [System.Text.Encoding]::UTF8)
$writer.WriteLine("Timestamp,Target,Status,RTT_ms")

Write-Host ("PingWatch: {0} and {1} for {2} min (until {3}). Ctrl+C to stop early." -f $Target1, $Target2, $DurationMinutes, $endTime.ToString('HH:mm:ss')) -ForegroundColor Cyan
Write-Host ("Log: {0}" -f $csvPath) -ForegroundColor DarkGray

try {
    while ((Get-Date) -lt $endTime) {
        $sw  = [System.Diagnostics.Stopwatch]::StartNew()
        $now = Get-Date

        # fire both pings at the same time
        $tasks = @()
        for ($i = 0; $i -lt 2; $i++) { $tasks += $pingers[$i].SendPingAsync($stats[$i].Target, $TimeoutMs) }
        try { [System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]$tasks) } catch { }

        $failures = 0
        Write-Host ($now.ToString('HH:mm:ss') + "  ") -NoNewline
        for ($i = 0; $i -lt 2; $i++) {
            $t = $tasks[$i]; $s = $stats[$i]
            $ok = $false; $rtt = 0; $status = 'Error'
            if (-not $t.IsFaulted -and $t.Result) {
                $status = [string]$t.Result.Status
                if ($t.Result.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                    $ok = $true; $rtt = [double]$t.Result.RoundtripTime
                }
            }
            if (-not $ok) { $failures++ }
            Update-Stat $s $ok $rtt $now

            $rttTxt = if ($ok) { $rtt } else { '' }
            $writer.WriteLine(("{0},{1},{2},{3}" -f $now.ToString('yyyy-MM-dd HH:mm:ss.fff'), $s.Target, $status, $rttTxt))

            if ($ok) { Write-Host ("{0,-16} {1,5} ms   " -f $s.Target, $rtt) -NoNewline -ForegroundColor Green }
            else     { Write-Host ("{0,-16} {1,-8}   " -f $s.Target, $status) -NoNewline -ForegroundColor Red }
        }
        Write-Host ""
        if ($failures -eq 2) { $script:bothDown++ }
        $writer.Flush()

        $remaining = $endTime - (Get-Date)
        try {
            $Host.UI.RawUI.WindowTitle = ("PingWatch | {0}: {1:N1}% loss | {2}: {3:N1}% loss | {4:hh\:mm\:ss} left" -f `
                $stats[0].Target, (100.0 * ($stats[0].Sent - $stats[0].Received) / $stats[0].Sent),
                $stats[1].Target, (100.0 * ($stats[1].Sent - $stats[1].Received) / $stats[1].Sent),
                $(if ($remaining.TotalSeconds -gt 0) { $remaining } else { [timespan]::Zero }))
        } catch { }

        $wait = $IntervalMs - $sw.ElapsedMilliseconds
        if ($wait -gt 0) { Start-Sleep -Milliseconds $wait }
    }
}
finally {
    $writer.Close()
    foreach ($p in $pingers) { $p.Dispose() }
    $summary = Get-Summary
    $summary | Set-Content -Path $sumPath -Encoding UTF8
    Write-Host ""
    $summary | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
    Write-Host ("Summary saved to: {0}" -f $sumPath) -ForegroundColor Cyan
    $csvText = [System.IO.File]::ReadAllText($csvPath).Trim()
    $reportTemplate.Replace('__CSV__', $csvText).Replace('__OUTAGE__', [string]$OutageThreshold) | Set-Content -Path $repPath -Encoding UTF8
    Write-Host ("Report saved to : {0}" -f $repPath) -ForegroundColor Cyan
}

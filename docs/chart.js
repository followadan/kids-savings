// 儲蓄累積曲線（每月月底存款）— 純 SVG，無外部函式庫
// renderSavingsChart(container, history, { color, ink, muted, grid, surface })
// history: [{ month: 'YYYY-MM', balance: Number, change: Number }]
(function () {
  const NS = 'http://www.w3.org/2000/svg';
  const fmt = n => Number(n || 0).toLocaleString('zh-TW', { maximumFractionDigits: 0 });
  const signed = n => (n > 0 ? '+' : n < 0 ? '−' : '±') + fmt(Math.abs(n));

  function niceStep(max, count) {
    const raw = max / count;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const f = raw / mag;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * mag;
  }

  function el(tag, attrs, text) {
    const e = document.createElementNS(NS, tag);
    for (const k in attrs) e.setAttribute(k, attrs[k]);
    if (text != null) e.textContent = text;
    return e;
  }

  window.renderSavingsChart = function (box, history, o) {
    o = Object.assign({ color: '#1f8a5b', ink: '#1f2328', muted: '#6b7280', grid: '#ebe5da', surface: '#fff' }, o || {});
    box.innerHTML = '';
    box.style.position = 'relative';
    if (!history || !history.length) {
      box.innerHTML = `<p style="color:${o.muted};margin:0">還沒有紀錄，存第一筆錢後就會出現曲線</p>`;
      return;
    }
    const W = 640, H = 250, L = 58, R = 24, T = 26, B = 42;
    const pw = W - L - R, ph = H - T - B;
    const vals = history.map(h => h.balance);
    const lo = Math.min(0, ...vals);
    const hiRaw = Math.max(...vals, 100);
    const step = Math.max(1, niceStep(hiRaw - lo, 4));
    const yMin = Math.floor(lo / step) * step;
    const yMax = Math.ceil(hiRaw / step) * step || step;
    const n = history.length;
    const X = i => n === 1 ? L + pw / 2 : L + (i / (n - 1)) * pw;
    const Y = v => T + (1 - (v - yMin) / (yMax - yMin)) * ph;

    const svg = el('svg', { viewBox: `0 0 ${W} ${H}`, width: '100%', role: 'img',
      'aria-label': `每月月底累積存款，最新 ${fmt(vals[n - 1])} 元`, style: 'display:block;overflow:visible' });

    // 縱軸：格線 + 刻度（元）
    for (let v = yMin; v <= yMax + 1e-9; v += step) {
      svg.appendChild(el('line', { x1: L, x2: W - R, y1: Y(v), y2: Y(v), stroke: o.grid, 'stroke-width': 1 }));
      svg.appendChild(el('text', { x: L - 8, y: Y(v) + 4, 'text-anchor': 'end', 'font-size': 12, fill: o.muted }, fmt(v)));
    }
    svg.appendChild(el('text', { x: L - 8, y: T - 12, 'text-anchor': 'end', 'font-size': 12, fill: o.muted }, '元'));

    // 橫軸：月份
    const every = Math.ceil(n / 8);
    let lastYear = null;
    history.forEach((h, i) => {
      const yy = Number(h.month.slice(0, 4)), mm = Number(h.month.slice(5, 7));
      if (!(i % every === 0 || i === n - 1)) return;
      svg.appendChild(el('text', { x: X(i), y: H - B + 18, 'text-anchor': 'middle', 'font-size': 12, fill: o.muted }, mm + '月'));
      if (yy !== lastYear) svg.appendChild(el('text', { x: X(i), y: H - B + 34, 'text-anchor': 'middle', 'font-size': 11, fill: o.muted }, String(yy)));
      lastYear = yy;
    });

    // 面積 + 折線
    if (n > 1) {
      const line = history.map((h, i) => `${i ? 'L' : 'M'}${X(i).toFixed(1)} ${Y(h.balance).toFixed(1)}`).join('');
      svg.appendChild(el('path', { d: `${line}L${X(n - 1)} ${Y(Math.max(0, yMin))}L${X(0)} ${Y(Math.max(0, yMin))}Z`, fill: o.color, 'fill-opacity': 0.12 }));
      svg.appendChild(el('path', { d: line, fill: 'none', stroke: o.color, 'stroke-width': 2.5, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
    }
    history.forEach((h, i) => {
      svg.appendChild(el('circle', { cx: X(i), cy: Y(h.balance), r: i === n - 1 ? 5 : 4, fill: o.color, stroke: o.surface, 'stroke-width': 2 }));
    });
    // 只標最新一點
    const lx = X(n - 1), ly = Y(vals[n - 1]);
    svg.appendChild(el('text', { x: Math.min(lx, W - R), y: ly - 12, 'text-anchor': n === 1 ? 'middle' : 'end', 'font-size': 13, 'font-weight': 700, fill: o.ink }, fmt(vals[n - 1]) + ' 元'));

    // 互動：十字線 + 提示框（滑鼠移入或手指點按）
    const cross = el('line', { y1: T, y2: T + ph, stroke: o.muted, 'stroke-width': 1, 'stroke-dasharray': '3 3', visibility: 'hidden' });
    svg.appendChild(cross);
    const tip = document.createElement('div');
    tip.style.cssText = `position:absolute;pointer-events:none;background:${o.ink};color:${o.surface};font-size:13px;line-height:1.4;padding:6px 10px;border-radius:8px;white-space:nowrap;display:none;transform:translate(-50%,-110%);z-index:2`;
    box.appendChild(svg); box.appendChild(tip);
    const slot = n === 1 ? pw : pw / (n - 1);
    history.forEach((h, i) => {
      const hit = el('rect', { x: X(i) - slot / 2, y: T, width: slot, height: ph, fill: 'transparent', style: 'cursor:pointer' });
      const show = () => {
        cross.setAttribute('x1', X(i)); cross.setAttribute('x2', X(i)); cross.setAttribute('visibility', 'visible');
        const r = svg.getBoundingClientRect(), s = r.width / W;
        tip.style.left = Math.max(70, Math.min(r.width - 70, X(i) * s)) + 'px';
        tip.style.top = Y(h.balance) * s + 'px';
        tip.innerHTML = `${h.month.slice(0, 4)} 年 ${Number(h.month.slice(5))} 月底<br><b>${fmt(h.balance)} 元</b>（當月 ${signed(h.change)}）`;
        tip.style.display = 'block';
      };
      const hide = () => { cross.setAttribute('visibility', 'hidden'); tip.style.display = 'none'; };
      hit.addEventListener('mouseenter', show);
      hit.addEventListener('mouseleave', hide);
      hit.addEventListener('touchstart', e => { show(); e.stopPropagation(); }, { passive: true });
      svg.appendChild(hit);
    });
    if (box._outside) document.removeEventListener('touchstart', box._outside);
    box._outside = e => { if (!box.contains(e.target)) { cross.setAttribute('visibility', 'hidden'); tip.style.display = 'none'; } };
    document.addEventListener('touchstart', box._outside, { passive: true });
  };
})();

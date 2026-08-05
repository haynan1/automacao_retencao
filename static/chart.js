/* Gráfico de barras horizontais (série única) — SVG, sem dependências.
   Compartilhado pela tela de resultado e pelo histórico.

   Uso:
     const ctrl = VizChart.mount({ host, seg, tip, data });
     // data = { "Setor": [{label, valor}], "Rubrica": [...], ... }
     ctrl.destroy();  // remove listeners (ex.: ao fechar um modal)
*/
(function () {
  const brl = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  function abrev(v) {
    // Milhares/milhões com 1 decimal só quando não é inteiro, para evitar
    // rótulos de eixo duplicados (ex.: 1500 e 2000 ambos virando "2k").
    const kfmt = (n) => (Number.isInteger(n) ? String(n) : n.toFixed(1).replace('.', ','));
    if (v >= 1e6) return 'R$ ' + kfmt(v / 1e6) + 'M';
    if (v >= 1e3) return 'R$ ' + kfmt(v / 1e3) + 'k';
    return 'R$ ' + Math.round(v);
  }
  function nice(m) {
    if (m <= 0) return 1;
    const p = Math.pow(10, Math.floor(Math.log10(m)));
    return Math.ceil(m / p) * p;
  }

  function svgFor(data, W) {
    const rowH = 30, padTop = 8, padBottom = 30;
    const gutterL = Math.round(Math.min(190, Math.max(96, W * 0.26)));
    const gutterR = 118;
    const plotW = Math.max(60, W - gutterL - gutterR);
    const H = padTop + data.length * rowH + padBottom;
    const max = nice(Math.max.apply(null, data.map((d) => d.valor)));
    const steps = 4;

    let s = `<svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}" role="img" aria-label="Gráfico de barras">`;
    for (let i = 0; i <= steps; i++) {
      const x = (gutterL + plotW * i / steps).toFixed(1);
      s += `<line class="viz-grid" x1="${x}" y1="${padTop}" x2="${x}" y2="${padTop + data.length * rowH}"/>`;
      s += `<text class="viz-axis" x="${x}" y="${H - 10}" text-anchor="middle">${abrev(max * i / steps)}</text>`;
    }
    data.forEach((d, i) => {
      const y = padTop + i * rowH;
      const bw = Math.max(2, plotW * d.valor / max);
      const cy = y + rowH / 2;
      const lbl = d.label.length > 26 ? d.label.slice(0, 25) + '…' : d.label;
      s += `<text class="viz-lbl" x="${gutterL - 10}" y="${cy + 4}" text-anchor="end">${esc(lbl)}</text>`;
      s += `<rect class="viz-bar" x="${gutterL}" y="${y + 6}" width="${bw.toFixed(1)}" height="${rowH - 13}" rx="5" `
         + `data-label="${esc(d.label)}" data-valor="${d.valor}"/>`;
      s += `<text class="viz-val" x="${(gutterL + bw + 8).toFixed(1)}" y="${cy + 4}">${brl.format(d.valor)}</text>`;
    });
    return s + '</svg>';
  }

  function mount(opts) {
    const host = opts.host, seg = opts.seg, tip = opts.tip, data = opts.data || {};
    const keys = Object.keys(data).filter((k) => (data[k] || []).length);
    let atual = keys[0];

    function render(key) {
      if (key === undefined) key = atual;   // re-render da dimensão atual (ex.: ao abrir modal)
      atual = key;
      const d = data[key] || [];
      host.innerHTML = d.length ? svgFor(d, host.clientWidth || 640)
                                : '<p class="text-secondary small mb-0">Sem valores.</p>';
    }

    function onMove(e) {
      if (!tip) return;
      const r = e.target.closest && e.target.closest('rect.viz-bar');
      if (!r) { tip.style.opacity = 0; return; }
      tip.innerHTML = `<strong>${esc(r.dataset.label)}</strong><br>${brl.format(+r.dataset.valor)}`;
      tip.style.opacity = 1;
      tip.style.left = (e.clientX + 14) + 'px';
      tip.style.top = (e.clientY + 14) + 'px';
    }
    function onLeave() { if (tip) tip.style.opacity = 0; }
    function onSeg(e) {
      const b = e.target.closest('button'); if (!b) return;
      seg.querySelectorAll('button').forEach((x) => x.classList.toggle('active', x === b));
      render(b.dataset.key);
    }
    let rt;
    function onResize() { clearTimeout(rt); rt = setTimeout(() => render(atual), 150); }

    if (seg) {
      seg.innerHTML = keys.map((k, i) =>
        `<button type="button" data-key="${esc(k)}" class="${i === 0 ? 'active' : ''}">${esc(k)}</button>`).join('');
      seg.addEventListener('click', onSeg);
    }
    if (tip) {
      host.addEventListener('mousemove', onMove);
      host.addEventListener('mouseleave', onLeave);
    }
    window.addEventListener('resize', onResize);

    if (keys.length) render(atual);
    else host.innerHTML = '<p class="text-secondary small mb-0">Sem valores para exibir.</p>';

    return {
      render,
      destroy() {
        window.removeEventListener('resize', onResize);
        if (seg) seg.removeEventListener('click', onSeg);
        if (tip) { host.removeEventListener('mousemove', onMove); host.removeEventListener('mouseleave', onLeave); }
      },
    };
  }

  window.VizChart = { mount };
})();

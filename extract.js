/* Lee el libro "Análisis económico" de Salva Energy y devuelve todos los datos
   que necesitan la presentación y la cotización. Funciona en navegador y en Node. */
(function (root) {
  function extract(XLSX, wb) {
    const miss = [];
    const sh = n => {
      const s = wb.Sheets[n];
      if (!s) miss.push('hoja "' + n + '"');
      return s || {};
    };
    const v = (n, a) => { const c = sh(n)[a]; return c ? c.v : undefined; };
    const num = (n, a) => { const x = v(n, a); const f = typeof x === 'number' ? x : parseFloat(x); if (!isFinite(f)) { miss.push(n + '!' + a); return 0; } return f; };
    const txt = (n, a) => { const x = v(n, a); return x == null ? '' : String(x).trim(); };
    const cols = (a, b) => { const r = []; let i = XLSX.utils.decode_col(a), j = XLSX.utils.decode_col(b); for (; i <= j; i++) r.push(XLSX.utils.encode_col(i)); return r; };
    const rowNums = (n, c1, c2, r) => cols(c1, c2).map(c => num(n, c + r));
    const colNums = (n, c, r1, r2) => { const o = []; for (let r = r1; r <= r2; r++) o.push(num(n, c + r)); return o; };
    const toDate = x => {
      if (x instanceof Date) return x;
      if (typeof x === 'number') { const d = XLSX.SSF.parse_date_code(x); return new Date(d.y, d.m - 1, d.d); }
      return null;
    };
    const MES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
    const lab = (n, a) => { const d = toDate(v(n, a)); return d ? MES[d.getMonth()] + '-' + String(d.getFullYear()).slice(2) : txt(n, a); };
    const labs = (n, c1, c2, r) => cols(c1, c2).map(c => lab(n, c + r));

    // --- Equipos (Capex completo filas 10-15: los que tienen cantidad > 0)
    let panel = null, inv = null;
    for (let r = 10; r <= 15; r++) {
      const d = txt('Capex completo', 'B' + r), q = num('Capex completo', 'F' + r);
      if (!q) continue;
      if (/PANEL/i.test(d) && !panel) panel = { desc: d, qty: q };
      if (/INVERSOR/i.test(d) && !inv) inv = { desc: d, qty: q };
    }
    panel = panel || { desc: '', qty: num('Cálculos', 'D2') };
    inv = inv || { desc: '', qty: 1 };
    const panelW = num('Cálculos', 'H3');
    const panelBrand = (panel.desc.match(/TRINA|JINKO|JA ?SOLAR|LONGI|CANADIAN|RISEN/i) || [''])[0];
    const invKW = Math.round(num('Cálculos', 'D23') / 1000) || parseFloat((inv.desc.match(/(\d+)\s*KW/i) || [])[1]) || 0;
    const invV = (inv.desc.match(/(\d{3})\s*V/i) || [, '220'])[1];
    const invBrand = /GOODWE/i.test(inv.desc) ? 'GoodWe' : /HUAWEI/i.test(inv.desc) ? 'Huawei' : /FRONIUS/i.test(inv.desc) ? 'Fronius' : /GROWATT/i.test(inv.desc) ? 'Growatt' : /DEYE/i.test(inv.desc) ? 'Deye' : '';
    const cap = s => s ? s.charAt(0) + s.slice(1).toLowerCase() : '';

    // --- Energía tradicional (Datos2)
    const costMonths = labs('Datos2', 'N', 'Z', 5);
    const costVals = rowNums('Datos2', 'N', 'Z', 7);
    const consVals = rowNums('Datos2', 'N', 'Z', 6);

    // --- Liquidación horaria
    const L = 'Liquidación';
    const hourly = { cons: colNums(L, 'D', 6, 29), prod: colNums(L, 'E', 6, 29) };

    // --- Comparativo factura (Resumen)
    const R = 'Resumen';
    const bill = {
      labels: labs(R, 'E', 'Q', 12),
      actual: rowNums(R, 'E', 'Q', 13),
      esfv: rowNums(R, 'E', 'Q', 15),
      futura: cols('F', 'R').map(c => { const x = v(R, c + '14'); return typeof x === 'number' ? x : 0; }),
      ahorro: rowNums(R, 'E', 'Q', 17)
    };

    // --- Valor planta (SIGO)
    const S = 'SIGO';
    const items = [];
    for (let r = 4; r <= 11; r++) {
      const d = txt(S, 'A' + r);
      if (d) items.push({ desc: d, pre: num(S, 'D' + r), iva: num(S, 'E' + r) });
    }
    items.sort((a, b) => b.iva - a.iva);
    const plant = { items, totPre: num(S, 'D12'), totIva: num(S, 'E12') };
    const pays = [];
    for (let r = 17; r <= 20; r++) pays.push({ desc: txt(S, 'C' + r), pct: num(S, 'D' + r), pre: num(S, 'E' + r), iva: num(S, 'F' + r) });

    // --- Indicadores (Capital propio, A O&M, Préstamo)
    const C = 'Capital propio';
    const flowYears = cols('G', 'Y').map(c => v(C, c + '49')).filter(x => x != null);
    const flowVals = cols('G', 'Y').slice(0, flowYears.length).map(c => num(C, c + '52'));
    const net = cols('H', 'Q').map(c => num(C, c + '50'));
    const om = cols('D', 'H').map(c => num('A O&M', c + '3'));
    const P = 'Prestamo 5 años';

    // --- Cotización (SIGO + Capex)
    const q = [
      ['PANELES SOLARES BIFACIALES (SUMINISTRO DE EQUIPOS DE ENERGIA RENOVABLE)', panel.qty, num(S, 'G4')],
      ['INVERSOR (SUMINISTRO DE EQUIPOS DE ENERGIA RENOVABLE)', inv.qty, num(S, 'G5')],
      ['SUMINISTRO DE MATERIALES Y/O EQUIPOS', 1, num(S, 'E5') + num(S, 'E9')],
      ['ESTRUCTURA METALICA', panel.qty, num(S, 'E10')],
      ['DISEÑO E INGENEIRIA', 1, num(S, 'E8')],
      ['SERVICIO DE INSTALACION (COSTO DIRECTO)', 1, num(S, 'J2')],
      ['Administración', 1, num(S, 'J3')],
      ['Imprevistos', 1, num(S, 'J4')],
      ['Utilidad', 1, num(S, 'J5') + num(S, 'J6')],
      ['CERTIFICACION RETIE', 1, num(S, 'E7')],
      ['TRÁMITE ANTE OR', 1, num(S, 'E11')]
    ].map((x, i) => ({ n: i + 1, desc: x[0], qty: x[1], val: x[2] }));
    // El total de la cotización es el de SIGO (E12); los paneles absorben el redondeo, como en la cotización manual
    const qTarget = num(S, 'E12');
    if (qTarget) q[0].val = qTarget - q.slice(1).reduce((s, x) => s + x.val, 0);
    const qTotal = q.reduce((s, x) => s + x.val, 0);
    let qIva = 0;
    for (let r = 5; r <= 11; r++) qIva += num(S, 'E' + r) - num(S, 'D' + r);
    qIva += num(S, 'J6');

    const data = {
      sys: {
        wp: num('Producción', 'B1'),
        panels: panel.qty, panelW, panelBrand: cap(panelBrand),
        invQty: inv.qty, invKW, invV, invBrand, invDesc: inv.desc,
        prod: num(L, 'D2'), cons: num(L, 'D3'),
        imp: num(L, 'F37'), impPct: num(L, 'H37'),
        exc: num(L, 'F32'), excPct: num(L, 'H32'),
        auto: num(L, 'F38'), autoPct: num(L, 'H38'),
        areaReq: num('Cálculos', 'C47'), nic: v('Datos2', 'D1')
      },
      trad: { months: costMonths, cost: costVals, cons: consVals, lastYear: costVals.slice(0, 12).reduce((a, b) => a + b, 0) * 1000 },
      hourly, bill, plant, pays,
      ind: {
        omYear: om[0], om5: om.reduce((a, b) => a + b, 0),
        saveYear: net[0], save10: net.reduce((a, b) => a + b, 0),
        payback: num(C, 'B61'),
        cur5: num(C, 'B64'), cur20: num(C, 'D64'),
        esf5: num(C, 'B65'), esf20: num(C, 'D65'),
        sav5: num(C, 'B66'), sav20: num(C, 'D66'),
        pct5: num(C, 'B67'), pct20: num(C, 'D67'),
        flowYears, flowVals,
        loanPay: num(P, 'H4'), loanRate: num(P, 'D5'), loanYears: num(P, 'D6')
      },
      tax: {
        iva: num('Capex completo', 'P174'),
        renta: num('Capex completo', 'P175'),
        dep: num('Capex completo', 'P176'),
        upme: num('Capex completo', 'K140') * (1 + num('Capex completo', 'G140'))
      },
      quote: { items: q, total: qTotal, iva: qIva, bruto: qTotal - qIva }
    };
    return { data, missing: [...new Set(miss)] };
  }
  if (typeof module !== 'undefined') module.exports = extract; else root.extractSalva = extract;
})(this);

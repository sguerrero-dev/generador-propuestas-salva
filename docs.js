/* Construye las diapositivas (960x540) y la cotización (A4, 794x1123) a partir de los datos del Excel */
(function (root) {
  const C = { green: '#31674A', deep: '#285340', dark: '#166748', lime: '#DDE69D', limeSoft: '#EEF3C9', navy: '#15172C', gold: '#E2C044', grey: '#F2F2F2', ink: '#1d2420' };
  const nf0 = new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf3 = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const nf2 = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const $m = x => '$ ' + nf0.format(Math.round(x));
  const $q = x => nf2.format(x);
  const pct = x => nf0.format(Math.round(x * 100)) + '%';
  const k = x => nf0.format(Math.round(x));
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const fechaLarga = iso => { const [y, m, d] = (iso || '').split('-').map(Number); if (!y) return ''; const M = MESES[m - 1]; return d + ' de ' + M.charAt(0).toUpperCase() + M.slice(1) + ' del ' + y; };

  const IMG = { mk: 'img/mark_green.png', mkL: 'img/mark_lime.png', wmG: 'img/wm_green.png', wmL: 'img/wm_lime.png', wmW: 'img/wm_white.png' };
  const iconPanel = `<svg viewBox="0 0 64 64" width="58" height="58" fill="none" stroke="#9ED39A" stroke-width="2.2"><path d="M14 30h38l-5 18H9z"/><path d="M12 39h37M23 30l-3 18M33 30l-1 18M43 30l1 18"/><path d="M30 48v8M22 56h18"/><circle cx="14" cy="14" r="5"/><path d="M14 3v4M14 21v4M3 14h4M21 14h4M6 6l3 3M19 19l3 3M6 22l3-3M19 9l3-3"/></svg>`;
  const iconTower = `<svg viewBox="0 0 64 64" width="58" height="58" fill="none" stroke="#2c6a4b" stroke-width="2.2"><path d="M32 4L18 60M32 4l14 56M24 30h16M21 42h22M18 60l25-18M46 60L21 42M24 30l19 12M40 30L21 42"/><path d="M10 18h44M14 12h36M10 18l22 12 22-12"/></svg>`;

  function slides(d, f) {
    const s = d.sys, I = d.ind, T = d.tax;
    const S = [];
    const head = (t, extra = '') => `<div class="s-h ${extra}"><h2>${t}</h2></div>`;
    // 1 Portada
    S.push(`<section class="slide img"><img src="img/s_portada.jpg" alt=""></section>`);
    // 2 Título
    S.push(`<section class="slide s2"><div class="s2-l"><img class="lines" src="img/cover_lines.jpg" alt=""><div class="s2-t"><h1>Sistema Solar<br>Fotovoltaico</h1><p>Análisis conceptual</p></div></div>
      <div class="s2-r"><img src="${IMG.wmG}" alt="salva energy"><b>SISTEMA SOLAR</b><b>${esc(f.tipo.toUpperCase())}<br>${esc(f.cliente.toUpperCase())}</b><b>${esc(f.proyecto.toUpperCase())}<br>${esc(f.ciudad.toUpperCase())}</b></div></section>`);
    // 3 Energía tradicional
    S.push(`<section class="slide s3"><img class="mk" src="${IMG.mk}" alt=""><h2 class="big">Situación energía tradicional</h2><p class="sub">Consumo supuesto de energía (basado en cuadro de cargas y tarifa de la zona)</p>
      <div class="ch" style="left:70px;top:132px;width:820px;height:300px"><canvas data-chart="trad"></canvas></div>
      <p class="s3-tot">Costo energía último año: <b>${$m(Math.round(d.trad.lastYear / 1000) * 1000)}</b></p><div class="foot-band"></div></section>`);
    // 4 Requerimientos
    S.push(`<section class="slide s4"><div class="pill">REQUERIMIENTOS</div><ul class="s4-l"><li>Área disponible: <b>${esc(f.areaDisp)} m²</b></li><li>Área requerida: <b>${k(Math.ceil(s.areaReq))} m²</b></li></ul>
      <div class="s4-img">${f.areaImg ? `<img src="${f.areaImg}" alt="Área de instalación">` : `<div class="s4-ph"><div class="s4-grid">${Array.from({ length: s.panels }, () => '<i></i>').join('')}</div><span>${k(s.panels)} paneles · ${k(Math.ceil(s.areaReq))} m² de cubierta</span></div>`}</div>
      <div class="s4-band"><img src="${IMG.wmL}" alt="salva energy"></div></section>`);
    // 5 Tamaño del sistema
    S.push(`<section class="slide s5"><div class="s5-side"></div><h2 class="big wide">Tamaño del sistema</h2>
      <div class="s5-a">${iconPanel}<div><p>${k(s.panels)} paneles ${k(s.panelW)} W${s.panelBrand ? ' ' + esc(s.panelBrand) : ''}</p><p>${k(s.invQty)} inversor on grid ${esc(s.invV)} V ${k(s.invKW)} kW ${esc(s.invBrand)}</p><p>Potencia: <b>${k(s.wp)} Wp</b></p><p>Promedio producción:<br><b>${k(s.prod)} kWh-mes</b></p></div></div>
      <div class="s5-b">${iconTower}<div><p>Promedio consumo red (actual)</p><p><b>${k(s.cons)} kWh-mes</b></p></div></div>
      <div class="s5-chart"><h3>Funcionamiento sistema solar</h3><div class="s5-chips"><span><i style="background:${C.navy}"></i>Energía importada <b>${k(s.imp)} kWh-mes (${pct(s.impPct)})</b></span><span><i style="background:${C.lime}"></i>Excedentes <b>${k(s.exc)} kWh-mes (${pct(s.excPct)})</b></span><span><i style="background:#8DBF7A"></i>Autoconsumo <b>${k(s.auto)} kWh-mes (${pct(s.autoPct)})</b></span></div>
      <div class="ch" style="left:14px;top:96px;width:392px;height:178px"><canvas data-chart="hourly"></canvas></div></div>
      <p class="s5-note">El sistema on grid estará conectado y será<br>dependiente del operador de red</p><img class="wm-br" src="${IMG.wmG}" alt=""></section>`);
    // 6 Factura
    S.push(`<section class="slide s6"><div class="s6-band"><img src="${IMG.mkL}" alt=""></div><h2 class="mid">Así se vería la factura de energía</h2>
      <div class="ch" style="left:60px;top:120px;width:840px;height:300px"><canvas data-chart="bill"></canvas></div>
      <p class="s6-note"><b>Nota:</b> Los datos aquí presentados corresponden a una proyección de la factura de energía con y sin sistema solar fotovoltaico (valores en millones de pesos); la misma podrá variar dependiendo de los hábitos de consumo y el valor de los precios de bolsa y tarifas de energía.</p></section>`);
    // 7 Valor de la planta
    const P = d.plant, pays = d.pays;
    const gantt = [
      ['Orden de compra', [1]], ['Gestión contractual', [1]], [`Anticipo (${pct(pays[0].pct)})`, [2]], ['Diseño', [3, 4, 5]],
      ['Adquisición paneles', [3, 4]], [`Segundo pago (${pct(pays[1].pct)})`, [5]], ['Adquisición inversores', [6, 7]],
      [`Tercer pago (${pct(pays[2].pct)})`, [8]], ['Instalación', [6, 7, 8, 9]], ['Pruebas y puesta en servicio', [10]], [`Cuarto pago (${pct(pays[3].pct)})`, [10]]
    ];
    S.push(`<section class="slide s7"><div class="s7-band"><h2>Valor de la planta solar</h2><img src="${IMG.wmL}" alt=""></div>
      <table class="t7 tdet"><thead><tr><th>Detalle</th><th>PV antes IVA</th><th>PV IVA incluido</th></tr></thead><tbody>${P.items.map(x => `<tr><td>${esc(x.desc)}</td><td>${$m(x.pre)}</td><td>${$m(x.iva)}</td></tr>`).join('')}</tbody><tfoot><tr><td>Total general</td><td>${$m(P.totPre)}</td><td>${$m(P.totIva)}</td></tr></tfoot></table>
      <table class="gantt"><thead><tr><th>Semana</th>${Array.from({ length: 10 }, (_, i) => `<th>${i + 1}</th>`).join('')}</tr></thead><tbody>${gantt.map(g => `<tr><td>${g[0]}</td>${Array.from({ length: 10 }, (_, i) => `<td class="${g[1].includes(i + 1) ? 'on' : ''}"></td>`).join('')}</tr>`).join('')}</tbody></table>
      <table class="t7 tpay"><thead><tr><th>Flujo de caja</th><th>%</th><th>Valor antes IVA</th><th>Valor con IVA</th></tr></thead><tbody>${pays.map(p => `<tr><td>${esc(p.desc)}</td><td>${pct(p.pct)}</td><td>${$m(p.pre)}</td><td>${$m(p.iva)}</td></tr>`).join('')}</tbody><tfoot><tr><td>Total</td><td>${pct(pays.reduce((a, b) => a + b.pct, 0))}</td><td>${$m(pays.reduce((a, b) => a + b.pre, 0))}</td><td>${$m(pays.reduce((a, b) => a + b.iva, 0))}</td></tr></tfoot></table></section>`);
    // 8 Indicadores
    const yrs = x => nf1.format(x) + ' años';
    S.push(`<section class="slide s8"><div class="s8-tag"><h2>Indicadores económicos<br>Capital propio</h2></div>
      <table class="kv" style="left:530px;top:14px;width:410px"><tr><th colspan="2">Costos asumidos por el cliente</th></tr><tr><td>Costos estimados anuales de O&amp;M</td><td>${$m(I.omYear)}</td></tr><tr><td>Costos estimados de O&amp;M en 5 años</td><td>${$m(I.om5)}</td></tr></table>
      <table class="kv" style="left:530px;top:92px;width:410px"><tr><th colspan="2">Ahorros esperados con la planta solar</th></tr><tr><td>Ahorros anuales estimados</td><td>${$m(I.saveYear)}</td></tr><tr><td>Ahorros estimados en 10 años</td><td>${$m(I.save10)}</td></tr><tr><td>Tiempo de retorno de la inversión</td><td><b>${yrs(I.payback)}</b></td></tr></table>
      <table class="vc"><thead><tr><th>Valores corrientes</th><th>5 años</th><th>20 años</th></tr></thead><tbody><tr><td>Situación actual</td><td>${$m(I.cur5)}</td><td>${$m(I.cur20)}</td></tr><tr><td>Costo ESF</td><td>${$m(I.esf5)}</td><td>${$m(I.esf20)}</td></tr><tr><td>Ahorro</td><td>${$m(I.sav5)}</td><td>${$m(I.sav20)}</td></tr><tr><td>% ahorro</td><td>${pct(I.pct5)}</td><td>${pct(I.pct20)}</td></tr></tbody></table>
      <p class="s8-fn">*Ahorro sin descontar beneficios tributarios</p>
      <h3 class="s8-ct">Flujo acumulado (millones de pesos)</h3><div class="ch" style="left:500px;top:250px;width:440px;height:182px"><canvas data-chart="flow"></canvas></div>
      ${f.financ ? `<table class="kv" style="left:530px;top:444px;width:410px"><tr><th colspan="2">Financiamiento</th></tr><tr><td>Cuota fija mensual</td><td>${$m(Math.round(I.loanPay / 1000) * 1000)}</td></tr><tr><td>Condiciones</td><td>${k(I.loanYears)} años, tasa ${nf0.format(I.loanRate * 100)}% E.A.</td></tr></table>` : ''}
      <img class="mk-bl" src="${IMG.mk}" alt=""></section>`);
    // 9 Beneficios tributarios
    S.push(`<section class="slide s9"><div class="s9-tag"><h2>Beneficios<br>tributarios</h2></div><img class="mk" src="${IMG.mk}" alt="">
      <p class="s9-upme">Costo trámite ante la UPME para obtener certificado tributario<br><br>(${$m(T.upme)} IVA incluido)</p>
      ${[['Exclusión del IVA', T.iva, 1], ['Descontar de la declaración de renta hasta 50% de la inversión', T.renta, 2], ['Depreciación acelerada del activo', T.dep, 3]].map((b, i) => `<div class="s9-row" style="top:${48 + i * 128}px"><div class="s9-box b${b[2]}"><p>${b[0]}</p><b>(${$m(b[1])})</b></div></div>`).join('')}</section>`);
    if (f.casos) {
      S.push(`<section class="slide img"><img src="img/s_casos1.jpg" alt=""></section>`);
      S.push(`<section class="slide img"><img src="img/s_casos2.jpg" alt=""></section>`);
    }
    // 12 Consideraciones
    const cons = [
      'La estructura de cubierta deberá contar con una reserva de carga muerta suficiente para la instalación de módulos solares.',
      'Para evitar perforar la cubierta y problemas futuros de impermeabilización, la cubierta deberá contar con elementos de fijación para el anclaje de los paneles.',
      'El cliente deberá entregar el área destinada para la instalación del SSF libre de obstáculos y obstrucciones, limpia, con el sistema de impermeabilización y desagüe finalizado y probado, y con los elementos de fijación instalados de acuerdo con el punto anterior.',
      'Obras civiles no son asumidas por SALVA ENERGY. No incluye sobrestructura ni dados en concreto.',
      'Las cubiertas donde se instalarán los paneles deberán contar con accesos, líneas de vida y puntos de anclaje para un trabajo en alturas seguro.',
      'El cliente deberá suministrar el punto de conexión a internet Ethernet para el monitoreo del sistema fotovoltaico.',
      'La cubierta debe disponer de un Sistema de Protección contra Rayos de acuerdo con lo dispuesto en el Artículo 16 del Reglamento Técnico de Instalaciones Eléctricas – RETIE (Res. 90708 de agosto 30 de 2013 con sus ajustes).',
      'La cubierta/techo/área donde se instalarán los módulos debe contar con un punto hidráulico cercano al sistema para realizar la limpieza con agua.',
      'El cliente deberá proporcionar un lugar seguro para la ubicación de los inversores.',
      `Para obtener los ahorros por medio de la exportación de energía es necesaria la aprobación del operador de red${f.or ? ' (' + esc(f.or) + ')' : ''} para iniciar la exportación.`,
      'Planta solar on-grid (con inversor híbrido compatible con baterías); la planta operará siempre y cuando se encuentre en sincronismo con la red eléctrica. En caso de falla de la red no entrega energía porque no tiene baterías.',
      'No se encuentran incluidos trabajos de independización de la medida de zonas comunes, en caso de requerirse.'
    ];
    S.push(`<section class="slide s12"><h2 class="big wide">Consideraciones</h2><img class="wm-tr" src="${IMG.wmG}" alt=""><ul>${cons.map(c => `<li>${c}</li>`).join('')}</ul><div class="foot-band r"></div></section>`);
    S.push(`<section class="slide img"><img src="img/s_porque.jpg" alt=""></section>`);
    S.push(`<section class="slide img"><img src="img/s_cierre.jpg" alt=""></section>`);
    return S.join('');
  }

  function charts(d, Chart, scope) {
    const made = [];
    const font = { family: 'Open Sans, Arial, sans-serif', size: 10 };
    Chart.defaults.font = font; Chart.defaults.color = '#3c4540'; Chart.defaults.animation = false;
    const labelPlugin = (fmt, sets, opts = {}) => ({
      id: 'lbl', afterDatasetsDraw(ch) {
        const g = ch.ctx; g.save(); g.font = `${opts.bold ? '700 ' : ''}${opts.size || 9}px Open Sans, Arial`; g.fillStyle = opts.color || '#3c4540'; g.textAlign = 'center';
        sets.forEach(si => ch.getDatasetMeta(si).data.forEach((el, i) => { const val = ch.data.datasets[si].data[i]; if (val == null || (opts.skipZero && !val)) return; const y = val < 0 ? el.y + 11 : el.y - 4; g.fillText(fmt(val), el.x, y); }));
        g.restore();
      }
    });
    const grid = { color: 'rgba(0,0,0,.07)' };
    const mk = (name, cfg) => { const cv = scope.querySelector(`canvas[data-chart="${name}"]`); if (!cv) return; cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, devicePixelRatio: 3, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, boxHeight: 8, font } } } }, cfg.options || {}); made.push(new Chart(cv, cfg)); };
    mk('trad', {
      data: { labels: d.trad.months, datasets: [
        { type: 'line', label: 'Consumo actual', data: d.trad.cons, borderColor: C.dark, borderWidth: 2, pointRadius: 0 },
        { type: 'bar', label: 'Costo energía', data: d.trad.cost, backgroundColor: '#E3EDA6', borderColor: '#C9D77A', borderWidth: 1 }] },
      options: { scales: { y: { beginAtZero: true, grid, ticks: { callback: v => '$ ' + nf0.format(v) } }, x: { grid: { display: false } } } },
      plugins: [labelPlugin(v => '$ ' + nf0.format(v), [1])]
    });
    mk('hourly', {
      type: 'line', data: { labels: Array.from({ length: 24 }, (_, i) => i + 1), datasets: [
        { label: 'Producción horaria', data: d.hourly.prod, fill: 'origin', backgroundColor: 'rgba(221,230,157,.85)', borderColor: '#B8C96A', borderWidth: 1.2, pointRadius: 0, tension: .35 },
        { label: 'Consumo horario', data: d.hourly.cons, fill: 'origin', backgroundColor: 'rgba(21,23,44,.88)', borderColor: C.navy, borderWidth: 1, pointRadius: 0, stepped: false, tension: .2 }] },
      options: { scales: { y: { beginAtZero: true, grid, ticks: { font: { size: 8 } }, title: { display: true, text: 'kWh-mes', font: { size: 8 } } }, x: { grid: { display: false }, ticks: { font: { size: 8 } } } }, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, boxHeight: 6, font: { size: 8 } } } } }
    });
    const b = d.bill;
    mk('bill', {
      data: { labels: b.labels, datasets: [
        { type: 'line', label: 'Ahorro', data: b.ahorro, borderColor: C.gold, backgroundColor: C.gold, borderWidth: 2, pointRadius: 2 },
        { type: 'bar', label: 'Factura actual', data: b.actual, backgroundColor: C.green },
        { type: 'bar', label: 'Costo ESFV (O&M)', data: b.esfv, backgroundColor: '#D3E08A' },
        { type: 'bar', label: 'Factura futura', data: b.futura, backgroundColor: C.navy }] },
      options: { scales: { y: { beginAtZero: true, grid, ticks: { callback: v => nf3.format(v) } }, x: { grid: { display: false } } } },
      plugins: [labelPlugin(v => nf3.format(v), [1], { size: 7.5 })]
    });
    mk('flow', {
      type: 'bar', data: { labels: d.ind.flowYears, datasets: [{ label: 'Flujo acumulado', data: d.ind.flowVals.map(v => Math.round(v * 10) / 10), backgroundColor: d.ind.flowVals.map(v => v < 0 ? '#8C9A8F' : '#3C8253') }] },
      options: { layout: { padding: { top: 12 } }, plugins: { legend: { display: false } }, scales: { y: { grid, ticks: { font: { size: 8 } } }, x: { grid: { display: false }, ticks: { font: { size: 7 }, maxRotation: 90, minRotation: 90 } } } },
      plugins: [labelPlugin(v => nf0.format(v), [0], { size: 7 })]
    });
    return made;
  }

  function quote(d, f) {
    const s = d.sys, q = d.quote;
    const kwp = nf1.format(s.wp / 1000);
    const ubic = `${f.proyecto} del municipio de ${f.ciudad}${f.depto ? ', ' + f.depto : ''}`;
    const head = n => `<div class="q-run"><span>Cotización No. ${esc(f.numero)}</span><span>${esc(f.cliente)}</span></div>`;
    const foot = n => `<div class="q-foot"><span>Salva Energy S.A.S. · NIT 901.854.067-6 · www.salva-energy.com</span><span>${n}/3</span></div>`;
    const p1 = `<section class="page">
      <div class="q-top"><div class="q-co"><img src="${IMG.mk}" alt=""><b>SALVA ENERGY SAS</b><span>NIT 901.854.067-6</span><span>Calle 96 # 46-99 Oficina 103</span><span>Tel: (605) 3114351788</span><span>Barranquilla - Colombia</span></div>
      <div class="q-box"><h1>Cotización</h1><p class="q-no">No. ${esc(f.numero)}</p><dl><dt>Para:</dt><dd>${esc(f.para)}</dd><dt>NIT:</dt><dd>${esc(f.nit)}</dd><dt>Fecha:</dt><dd>${esc(f.fecha)}</dd></dl></div></div>
      <p>Barranquilla, ${fechaLarga(f.fecha)}</p><p>Estimado(s) ${esc(f.saludo)},</p>
      <p>Es un gusto para nosotros dirigirnos a usted en nombre de Salva Energy, empresa dedicada a brindar soluciones innovadoras y sostenibles en energía renovable y regulada. Nuestro propósito es acompañar a empresas y organizaciones en la transición hacia un modelo energético más eficiente, confiable y responsable con el medioambiente. En Salva Energy nos mueve la pasión por transformar el futuro a través de soluciones energéticas renovables y de energía regulada que generen valor sostenible para nuestros clientes y comunidades.</p>
      <p>Cada proyecto lo asumimos con un firme compromiso, guiados por la ética y la responsabilidad que exige impulsar un cambio positivo en el sector energético. Nuestro equipo trabaja unido, con dedicación, fomentando el trabajo en equipo y la cercanía hacia nuestros clientes, permitiéndonos entender las necesidades de quienes confían en nosotros, siempre con un profundo sentido de pertenencia hacia la misión de aportar al desarrollo del Caribe colombiano y del país. Conscientes de la importancia de garantizar eficiencia y respaldo energético, ponemos a su disposición nuestra experiencia y un equipo altamente calificado para diseñar y ejecutar proyectos que se adapten a sus necesidades específicas, con la seguridad de que encontrará en nosotros un aliado estratégico para alcanzar sus objetivos de sostenibilidad y competitividad.</p>
      <p>Agradecemos de antemano la atención prestada y quedamos atentos para ampliar cualquier inquietud.</p>
      <p class="q-sig">Marjorie Jiménez<br>CEO &amp; Co-founder<br>SALVA ENERGY<br>www.salva-energy.com<br>Barranquilla, Colombia.</p>
      <p><b>Asunto: Cotización proyecto de sistema solar</b></p>
      <p>SALVA ENERGY pone a consideración la siguiente oferta técnico-económica:</p>
      <h3>1. Alcance técnico</h3><h4>Implementación de sistema fotovoltaico conectado a red</h4>
      <p>La presente propuesta contempla el desarrollo <b>del proyecto fotovoltaico de ${kwp} kWp</b>, ubicado en ${esc(ubic)}, correspondiente al suministro, instalación, legalización y puesta en operación del sistema de generación solar conectado a red, que generará una producción mensual promedio de <b>${k(s.prod)} kWh-mes</b>, y está compuesto por:</p>
      <h4>Ingeniería y diseño del sistema</h4><ul><li>Diseño eléctrico del sistema fotovoltaico.</li><li>Dimensionamiento de equipos según capacidad instalada.</li><li>Diagramas unifilares del sistema.</li><li>Definición de protecciones eléctricas AC y DC.</li><li>Memorias técnicas requeridas para proceso de legalización.</li></ul>
      ${foot(1)}</section>`;
    const p2 = `<section class="page">${head(2)}
      <h4>Suministro de energía renovable</h4><p>Se incluye el suministro de todos los equipos y materiales necesarios para el funcionamiento del sistema on grid (sin baterías, es decir sin respaldo de energía en caso de interrupción con el operador de red), tales como:</p>
      <ul class="hl"><li>${k(s.panels)} paneles solares bifaciales ${k(s.panelW)} W${s.panelBrand ? ', marca ' + esc(s.panelBrand) : ''}</li><li>${k(s.invQty)} inversor${s.invQty > 1 ? 'es' : ''} on grid con smart meter y aplicación, capacidad mínima ${k(s.invKW)} kW${s.invBrand ? ', marca ' + esc(s.invBrand) : ''}</li></ul>
      <h4>Suministro de equipos y materiales</h4><p>Se incluye el suministro de todos los equipos y materiales necesarios y certificados para la implementación del sistema, tales como:</p>
      <ul><li>Estructura de soporte para módulos (para instalación en cubierta de los paneles, rieles, uniones, grapas, pegante especial).</li><li>Cableado solar DC.</li><li>Protecciones eléctricas AC y DC.</li><li>Tableros eléctricos asociados al sistema.</li><li>Canaletas, tuberías.</li></ul>
      <h4>Instalación del sistema</h4><ul><li>Conexionado del sistema en corriente continua (DC) y corriente alterna (AC).</li><li>Montaje de paneles solares en cubierta con elementos SST.</li><li>Integración del sistema al tablero eléctrico de la instalación.</li><li>Pruebas de funcionamiento y puesta en marcha del sistema.</li><li>Ingeniero supervisor a cargo antes, durante y después del proyecto.</li></ul>
      <h4>Certificación eléctrica</h4><ul><li>Gestión y obtención del <b>Certificado RETIE</b> mediante organismo de inspección acreditado.</li><li>Entrega de documentación técnica requerida para la certificación.</li></ul>
      <h4>Trámites ante el operador de red</h4><ul><li>Gestión del proceso de conexión del sistema ante el operador de red${f.or ? ' <b>' + esc(f.or) + '</b>' : ''}.</li><li>Radicación de documentación técnica.</li><li>Seguimiento al proceso hasta la aprobación de conexión.</li></ul>
      <h4>No incluye</h4><ul><li>Acabados estéticos no contemplados.</li><li>Servicios externos al sistema no detallados.</li><li>Conexión de punto de datos.</li><li>Trámite ante la UPME.</li></ul>
      <h3>2. Tiempo de entrega</h3><ul class="two"><li>Diseño: 2-5 días hábiles</li><li>Procura de equipos: 10-30 días hábiles</li><li>Instalación: 5-15 días hábiles</li><li>Pruebas y puesta en servicio: 5-7 días hábiles</li><li>Trámites ante OR: 3-6 meses</li><li>Trámite ante UPME: 4-8 meses</li></ul>
      <h3>3. Garantía de las instalaciones</h3><ul class="two"><li>Estructura de paneles: 10 años</li><li>Materiales e instalación: 2 años</li><li>Paneles solares: 12 años de fabricación, 25 años de producción</li><li>Inversor: 10 años${s.invBrand ? ' (marca ' + esc(s.invBrand) + ')' : ''}</li></ul>
      ${foot(2)}</section>`;
    const p3 = `<section class="page">${head(3)}
      <h3>4. Oferta económica</h3><p>SALVA ENERGY presenta a continuación la oferta económica de los servicios descritos:</p>
      <table class="q-t"><thead><tr><th>Ítem</th><th>Descripción</th><th>Cantidad</th><th>Vr. total</th></tr></thead><tbody>${q.items.map(x => `<tr><td>${x.n}</td><td>${esc(x.desc)}</td><td>${nf2.format(x.qty)}</td><td>${$q(x.val)}</td></tr>`).join('')}</tbody></table>
      <table class="q-tot"><tr><td>Total bruto</td><td>${$q(q.bruto)}</td></tr><tr><td>Subtotal</td><td>${$q(q.bruto)}</td></tr><tr><td>IVA 0%</td><td>0,00</td></tr><tr><td>IVA 19%</td><td>${$q(q.iva)}</td></tr><tr class="g"><td>Total a pagar</td><td>$ ${$q(q.total)}</td></tr></table>
      <h3>5. Condiciones comerciales</h3><p><b>Forma de pago:</b></p><ul>${d.pays.map(p => `<li>${pct(p.pct)} ${esc(p.desc.toLowerCase())} (${$m(p.iva)})</li>`).join('')}</ul>
      <p><b>Validez de la oferta:</b> esta oferta tiene validez de 30 días calendario. En caso de no ser aceptada dentro del periodo de validez, se deberán revisar los valores ofertados de acuerdo con las características del mercado.</p>
      <p><b>Confidencialidad:</b> de acuerdo con el principio de buena fe en la etapa precontractual, la presente propuesta tiene un carácter reservado y confidencial, que impide el suministro a terceros de la información aquí contenida. Su uso se limita exclusivamente al destinatario de la oferta (Artículo 863 del Código de Comercio).</p>
      <p>La presente oferta se presenta el ${fechaLarga(f.fecha)}, y esperamos sea de su total agrado.</p><p>Cordialmente,</p>
      <div class="q-signs"><p><b>Marjorie Katherine Jiménez Guerra</b><br>Gerente de Operaciones<br>Salva Energy S.A.S.<br>mkjimenez@salva-energy.com</p><p><b>${esc(f.asesor)}</b><br>${esc(f.asesorMail)}<br>${esc(f.asesorTel)}</p></div>
      ${foot(3)}</section>`;
    return p1 + p2 + p3;
  }

  root.SalvaDocs = { slides, charts, quote, fmt: { $m, k } };
})(this);

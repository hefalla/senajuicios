/* ── Estado global ─────────────────────────────────────────────── */
let currentTaskId = null;
let currentEventSource = null;
let currentPasos = [];
let ocupado = false;

/* ── Tabs ──────────────────────────────────────────────────────── */
function switchTab(idx) {
  document.querySelectorAll('.tab-btn').forEach((b, i) =>
    b.classList.toggle('active', i === idx));
  document.querySelectorAll('.tab-content').forEach((t, i) =>
    t.classList.toggle('active', i === idx));
  const panelRight = document.querySelector('.panel-right');
  if (panelRight) panelRight.style.display = idx === 3 ? 'none' : '';
  if (idx !== 3) {
    document.getElementById('dashboardPanel').style.display = 'none';
  }
}

/* ── Log ───────────────────────────────────────────────────────── */
function addLog(level, msg) {
  const body = document.getElementById('logBody');
  // Limpiar placeholder la primera vez
  if (body.querySelector('.log-line[style]')) body.innerHTML = '';

  const span = document.createElement('span');
  span.className = `log-line ${level || ''}`;
  const prefix = { ok: '✔ ', err: '✘ ', info: '→ ', titulo: '══ ' }[level] || '';
  span.textContent = prefix + msg;
  body.appendChild(span);
  body.appendChild(document.createElement('br'));
  body.scrollTop = body.scrollHeight;
}

function clearLog() {
  document.getElementById('logBody').innerHTML =
    '<span class="log-line" style="color:rgba(255,255,255,.2); font-style:italic;">Log limpiado.</span>';
}

/* ── Steps ─────────────────────────────────────────────────────── */
function renderSteps(pasos) {
  const list = document.getElementById('stepsList');
  list.innerHTML = '';
  pasos.forEach((label, i) => {
    const div = document.createElement('div');
    div.className = 'step pendiente';
    div.id = `step_${i}`;
    div.innerHTML = `
      <div class="step-icon">○</div>
      <div class="step-label">${label}</div>`;
    list.appendChild(div);
  });
}

function updateStep(index, estado) {
  if (index < 0) {
    // error genérico: marca el último paso en curso como error
    document.querySelectorAll('.step.en-curso').forEach(s => {
      s.className = 'step error';
      s.querySelector('.step-icon').textContent = '✘';
    });
    return;
  }
  const el = document.getElementById(`step_${index}`);
  if (!el) return;
  const icons = { 'en_curso': '◉', 'listo': '✔', 'error': '✘' };
  el.className = `step ${estado.replace('_', '-')}`;
  el.querySelector('.step-icon').textContent = icons[estado] || '○';
}

/* ── Status pill ───────────────────────────────────────────────── */
function setStatus(estado, texto) {
  const pill = document.getElementById('statusPill');
  pill.className = `status-pill ${estado}`;
  pill.textContent = texto;
}

function setSpinner(on) {
  document.getElementById('spinner').classList.toggle('hidden', !on);
}

/* ── Botones y bloqueo ─────────────────────────────────────────── */
const allBtns = ['btn_descargar', 'btn_consultar', 'btn_completo'];

function lockUI() {
  ocupado = true;
  allBtns.forEach(id => {
    const b = document.getElementById(id);
    b.disabled = true;
  });
}

function unlockUI() {
  ocupado = false;
  allBtns.forEach(id => {
    document.getElementById(id).disabled = false;
  });
}

/* ── SSE listener ──────────────────────────────────────────────── */
function conectarSSE(taskId) {
  if (currentEventSource) currentEventSource.close();

  const es = new EventSource(`/api/stream/${taskId}`);
  currentEventSource = es;

  es.onmessage = (e) => {
    const ev = JSON.parse(e.data);

    if (ev.type === '_eof') {
      es.close();
      unlockUI();
      setSpinner(false);
      return;
    }

    if (ev.type === 'log') {
      addLog(ev.level, ev.msg);
      return;
    }

    if (ev.type === 'paso') {
      updateStep(ev.index, ev.estado);
      return;
    }

    if (ev.type === 'archivo') {
      const btn = document.getElementById('btnDownload');
      btn.href = `/api/descargar-archivo/${taskId}`;
      btn.download = ev.nombre;
      const icono = ev.es_zip ? '📦' : '📄';
      btn.textContent = `⬇ ${icono} ${ev.nombre}`;
      btn.classList.add('show');
      // Disparar la descarga automáticamente
      setTimeout(() => btn.click(), 400);
      return;
    }

    if (ev.type === 'fin') {
      const banner = document.getElementById('resultBanner');
      banner.className = `result-banner show ${ev.exito ? 'exito' : 'error'}`;
      document.getElementById('resultIcon').textContent = ev.exito ? '✅' : '❌';
      document.getElementById('resultMsg').textContent = ev.msg;

      document.getElementById('progressTitle').textContent =
        ev.exito ? 'Proceso completado' : 'Proceso detenido';
      setStatus(ev.exito ? 'done' : 'error', ev.exito ? 'LISTO' : 'ERROR');
    }
  };

  es.onerror = () => {
    es.close();
    unlockUI();
    setSpinner(false);
    setStatus('error', 'ERROR');
    addLog('err', 'Conexión con el servidor interrumpida.');
  };
}

/* ── Inicio de tareas ──────────────────────────────────────────── */
function _prepararUI(titulo, pasos) {
  if (ocupado) {
    alert('Ya hay un proceso en curso. Espera a que termine.');
    return false;
  }
  lockUI();

  // Reset panel
  document.getElementById('progressTitle').textContent = titulo;
  document.getElementById('resultBanner').className = 'result-banner';
  document.getElementById('btnDownload').className = 'btn-download';
  setStatus('running', 'EJECUTANDO');
  setSpinner(true);
  renderSteps(pasos);
  addLog('titulo', titulo);
  return true;
}

async function _iniciar(endpoint, payload) {
  try {
    const res = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) {
      addLog('err', data.error || 'Error desconocido');
      unlockUI(); setSpinner(false);
      setStatus('error', 'ERROR');
      return;
    }
    currentTaskId = data.task_id;
    renderSteps(data.pasos);
    conectarSSE(currentTaskId);
  } catch (err) {
    addLog('err', `Error de red: ${err.message}`);
    unlockUI(); setSpinner(false);
    setStatus('error', 'ERROR');
  }
}

function iniciarDescargar() {
  const usuario    = document.getElementById('d_usuario').value.trim();
  const contrasena = document.getElementById('d_contrasena').value.trim();
  const ficha      = document.getElementById('d_ficha').value.trim();
  if (!usuario || !contrasena || !ficha) {
    alert('Completa usuario, contraseña y número de ficha.'); return;
  }
  if (!_prepararUI('Descargando reporte · Ficha ' + ficha, [])) return;
  _iniciar('/api/descargar', { usuario, contrasena, numero_ficha: ficha });
}

function iniciarConsultar() {
  const ficha    = document.getElementById('c_ficha').value.trim();
  if (!ficha) { alert('Ingresa el número de ficha.'); return; }
  const nombre   = document.getElementById('c_nombre').value.trim();
  const apellido = document.getElementById('c_apellido').value.trim();
  const tipo     = document.querySelector('input[name="c_tipo"]:checked').value;
  const activos  = document.getElementById('c_activos').checked;
  if (!_prepararUI('Generando informe · Ficha ' + ficha, [])) return;
  _iniciar('/api/consultar', { numero_ficha: ficha, nombre, apellido, tipo, solo_activos: activos });
}

function iniciarCompleto() {
  const usuario    = document.getElementById('f_usuario').value.trim();
  const contrasena = document.getElementById('f_contrasena').value.trim();
  const ficha      = document.getElementById('f_ficha').value.trim();
  if (!usuario || !contrasena || !ficha) {
    alert('Completa usuario, contraseña y número de ficha.'); return;
  }
  const nombre   = document.getElementById('f_nombre').value.trim();
  const apellido = document.getElementById('f_apellido').value.trim();
  const tipo     = document.querySelector('input[name="f_tipo"]:checked').value;
  const activos  = document.getElementById('f_activos').checked;
  if (!_prepararUI('Proceso completo · Ficha ' + ficha, [])) return;
  _iniciar('/api/completo', { usuario, contrasena, numero_ficha: ficha,
                               nombre, apellido, tipo, solo_activos: activos });
}

/* ── Inicialización ────────────────────────────────────────────── */
let _carpetaSeleccionada = null;

(async () => {
  try {
    const r = await fetch('/api/carpeta');
    const d = await r.json();
    const badge = document.getElementById('carpetaBadge');
    const nombreCarpeta = d.carpeta.replace(/\\/g, '/').split('/').filter(Boolean).pop() || d.carpeta;
    badge.textContent = '📂 ' + nombreCarpeta;
    document.getElementById('carpetaActualDisplay').textContent = d.carpeta;
    _carpetaSeleccionada = d.carpeta;
  } catch { /* silent */ }
})();

/* ── Modal carpeta ─────────────────────────────────────────────── */
function abrirModalCarpeta() {
  document.getElementById('modalCarpeta').classList.add('open');
  document.getElementById('modalMsg').className = 'modal-msg';
  explorarRuta(_carpetaSeleccionada || '');
}

function cerrarModalCarpeta() {
  document.getElementById('modalCarpeta').classList.remove('open');
}

function cerrarModalSiOverlay(e) {
  if (e.target.id === 'modalCarpeta') cerrarModalCarpeta();
}

let _rutaExplorada = '';
let _clickTimer = null;

async function explorarRuta(ruta) {
  const input = document.getElementById('explorerInput');
  if (ruta === undefined) ruta = input.value.trim();
  if (ruta) input.value = ruta;

  const list = document.getElementById('explorerList');
  list.innerHTML = '<div class="explorer-loading">Cargando…</div>';

  try {
    const url = '/api/explorar' + (ruta ? '?ruta=' + encodeURIComponent(ruta) : '');
    const r = await fetch(url);
    const d = await r.json();

    if (d.error) {
      list.innerHTML = `<div class="explorer-loading" style="color:var(--rojo)">${d.error}</div>`;
      return;
    }

    _rutaExplorada = d.ruta_actual || ruta;
    document.getElementById('explorerBar').textContent = _rutaExplorada || '/';
    input.value = _rutaExplorada;

    list.innerHTML = '';
    if (!d.elementos || d.elementos.length === 0) {
      list.innerHTML = '<div class="explorer-loading">Carpeta vacía</div>';
      return;
    }

    d.elementos.forEach(el => {
      const div = document.createElement('div');
      div.className = 'explorer-item';
      if (el.ruta === _carpetaSeleccionada) div.classList.add('selected');

      const icon = el.tipo === 'padre' ? '⬆' : '📁';
      div.innerHTML = `<span class="explorer-icon">${icon}</span><span>${el.nombre}</span>`;
      div.title = el.ruta;

      // Simple click → seleccionar como destino
      div.addEventListener('click', () => {
        clearTimeout(_clickTimer);
        _clickTimer = setTimeout(() => {
          if (el.tipo !== 'padre') {
            _carpetaSeleccionada = el.ruta;
            document.getElementById('explorerInput').value = el.ruta;
            document.querySelectorAll('.explorer-item').forEach(i => i.classList.remove('selected'));
            div.classList.add('selected');
          } else {
            explorarRuta(el.ruta);
          }
        }, 220);
      });

      // Double click → entrar a la carpeta
      div.addEventListener('dblclick', () => {
        clearTimeout(_clickTimer);
        explorarRuta(el.ruta);
      });

      list.appendChild(div);
    });

  } catch (err) {
    list.innerHTML = `<div class="explorer-loading" style="color:var(--rojo)">Error: ${err.message}</div>`;
  }
}

async function guardarCarpeta() {
  const ruta = document.getElementById('explorerInput').value.trim() || _carpetaSeleccionada;
  if (!ruta) {
    mostrarMsgModal('err', 'Selecciona o escribe una ruta primero.');
    return;
  }

  const btn = document.getElementById('btnGuardarCarpeta');
  btn.disabled = true;
  btn.textContent = 'Guardando…';

  try {
    const r = await fetch('/api/carpeta', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ carpeta: ruta }),
    });
    const d = await r.json();

    if (!r.ok) {
      mostrarMsgModal('err', d.error || 'Error al guardar.');
    } else {
      _carpetaSeleccionada = d.carpeta;
      document.getElementById('carpetaBadge').textContent = '📂 ' + d.carpeta;
      document.getElementById('carpetaActualDisplay').textContent = d.carpeta;
      mostrarMsgModal('ok', '✔ Carpeta guardada correctamente.');
      setTimeout(cerrarModalCarpeta, 1200);
    }
  } catch (err) {
    mostrarMsgModal('err', `Error de red: ${err.message}`);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Guardar carpeta';
  }
}

function mostrarMsgModal(tipo, msg) {
  const el = document.getElementById('modalMsg');
  el.className = `modal-msg ${tipo}`;
  el.textContent = msg;
}
/* ── Dashboard ─────────────────────────────────────────────────── */
const DB_ESTADO_COLORS = {
  'EN FORMACION':    '#185FA5',
  'CANCELADO':       '#888780',
  'TRASLADADO':      '#D85A30',
  'RETIRO VOLUNTARIO':'#BA7517',
  'APLAZADO':        '#639922',
  'CONDICIONADO':    '#D4537E',
};
const DB_ESTADO_BADGE = {
  'EN FORMACION':'db-badge-b','CANCELADO':'db-badge-gr','TRASLADADO':'db-badge-gr',
  'RETIRO VOLUNTARIO':'db-badge-a','APLAZADO':'db-badge-a','CONDICIONADO':'db-badge-r',
};
const DB_ESTADO_LABEL = {
  'EN FORMACION':'En formación','CANCELADO':'Cancelado','TRASLADADO':'Trasladado',
  'RETIRO VOLUNTARIO':'Retiro vol.','APLAZADO':'Aplazado','CONDICIONADO':'Condicionado',
};
const DB_JUICIO_COLORS = {'APROBADO':'#1d9e75','POR EVALUAR':'#BA7517','NO APROBADO':'#e24b4a'};

let _dbData      = null;
let _dbFiltro    = 'todos';
let _dbChartEst  = null;
let _dbChartJui  = null;

async function cargarDashboard() {
  const ficha = document.getElementById('db_ficha').value.trim();
  if (!ficha) { alert('Ingresa el número de ficha.'); return; }

  const btn = document.getElementById('btn_dashboard');
  btn.disabled = true;
  btn.innerHTML = '<span class="btn-icon">⏳</span> Cargando…';

  try {
    const r = await fetch(`/api/dashboard/${encodeURIComponent(ficha)}`);
    const d = await r.json();
    if (!r.ok) { alert(d.error || 'Error al cargar el dashboard.'); return; }
    _dbData = d;
    _dbFiltro = 'todos';
    _renderDashboard();
    await _dbCargarConsulta(ficha);
    document.getElementById('dashboardPanel').style.display = 'block';
    document.getElementById('panel-right') && (document.getElementById('panel-right').style.display = 'none');
    document.getElementById('dashboardPanel').scrollIntoView({behavior:'smooth', block:'start'});
  } catch(e) {
    alert('Error de red: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<span class="btn-icon">📊</span> Ver dashboard';
  }
}

function _renderDashboard() {
  const d = _dbData;
  const m = d.ficha_meta;

  document.getElementById('dbTitulo').textContent =
    `Juicios evaluativos — Ficha ${m.ficha}`;
  document.getElementById('dbSubtitulo').textContent =
    `${m.denominacion} · ${m.regional} · ${m.centro} · Reporte: ${m.fecha_reporte}`;

  document.getElementById('dbTotalAp').textContent        = d.total_aprendices;
  document.getElementById('dbActivos').textContent         = d.aprendices_activos;
  document.getElementById('dbRaNoAprob').textContent       = d.ra_no_aprobados;
  document.getElementById('dbCompSinEvaluar').textContent  = d.competencias_sin_evaluar;

  _renderCharts(d);
  _renderFuncionarios(d.ultimos_funcionarios);
  _renderFiltros(d.aprendices);
  _renderTablaAp();
}

function _renderCharts(d) {
  const estLabels = Object.keys(d.estados);
  const estVals   = Object.values(d.estados);
  const estColors = estLabels.map(l => DB_ESTADO_COLORS[l] || '#888780');

  const legEst = document.getElementById('dbLegEstados');
  legEst.innerHTML = estLabels.map((l,i) =>
    `<span style="display:flex;align-items:center;gap:4px;">
      <span style="width:9px;height:9px;border-radius:2px;background:${estColors[i]};flex-shrink:0;display:inline-block;"></span>
      ${DB_ESTADO_LABEL[l]||l} (${estVals[i]})
    </span>`
  ).join('');

  if (_dbChartEst) _dbChartEst.destroy();
  _dbChartEst = new Chart(document.getElementById('dbChartEstados'), {
    type: 'doughnut',
    data: { labels: estLabels, datasets: [{ data: estVals, backgroundColor: estColors, borderWidth: 2, borderColor: 'transparent' }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
  });

  const juiLabels = Object.keys(d.juicios);
  const juiVals   = Object.values(d.juicios);
  const juiColors = juiLabels.map(l => DB_JUICIO_COLORS[l] || '#888780');
  const total     = juiVals.reduce((a,b) => a+b, 0);

  const legJui = document.getElementById('dbLegJuicios');
  legJui.innerHTML = juiLabels.map((l,i) =>
    `<span style="display:flex;align-items:center;gap:4px;">
      <span style="width:9px;height:9px;border-radius:2px;background:${juiColors[i]};flex-shrink:0;display:inline-block;"></span>
      ${l.charAt(0)+l.slice(1).toLowerCase()} (${juiVals[i]} — ${Math.round(juiVals[i]/total*100)}%)
    </span>`
  ).join('');

  if (_dbChartJui) _dbChartJui.destroy();
  _dbChartJui = new Chart(document.getElementById('dbChartJuicios'), {
    type: 'doughnut',
    data: { labels: juiLabels, datasets: [{ data: juiVals, backgroundColor: juiColors, borderWidth: 2, borderColor: 'transparent' }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
  });
}

function _renderFuncionarios(lista) {
  const el = document.getElementById('dbFuncionarios');
  el.innerHTML = lista.map(f =>
    `<div class="db-func-row">
      <div style="flex:1;min-width:0;">
        <div class="db-func-name">${f.nombre}</div>
        <div class="db-func-ra">${f.ra}</div>
      </div>
      <div class="db-func-date">${f.fecha}</div>
    </div>`
  ).join('');
}

function _renderFiltros(aprendices) {
  const estados = ['todos', ...new Set(aprendices.map(a => a.estado))];
  const cont = document.getElementById('dbFiltros');
  cont.innerHTML = estados.map(e =>
    `<button class="db-tab-btn${e===_dbFiltro?' active':''}"
      onclick="dbFiltrar('${e}',this)">
      ${e==='todos'?'Todos':(DB_ESTADO_LABEL[e]||e)}
    </button>`
  ).join('');
}

function dbFiltrar(estado, btn) {
  _dbFiltro = estado;
  document.querySelectorAll('.db-tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  _renderTablaAp();
}

function _juicioBadge(j) {
  if (j==='APROBADO')    return '<span class="db-badge db-badge-g">Aprobado</span>';
  if (j==='NO APROBADO') return '<span class="db-badge db-badge-r">No aprobado</span>';
  return '<span class="db-badge db-badge-gr">Por evaluar</span>';
}

function _renderTablaAp() {
  const data = _dbFiltro === 'todos'
    ? _dbData.aprendices
    : _dbData.aprendices.filter(a => a.estado === _dbFiltro);

  let html = '';
  data.forEach((ap) => {
    const pct = Math.round(ap.aprobados / Math.max(ap.total_ra, 1) * 100);
    const bc  = DB_ESTADO_BADGE[ap.estado] || 'db-badge-gr';
    const el  = DB_ESTADO_LABEL[ap.estado] || ap.estado;

    html += `<tr>
      <td>${ap.nombre}</td>
      <td data-label="Estado"><span class="db-badge ${bc}">${el}</span></td>
      <td data-label="Total" style="text-align:right;">${ap.total_ra}</td>
      <td data-label="Aprobados" style="text-align:right;color:#1d9e75;font-weight:500;">${ap.aprobados}</td>
      <td data-label="No aprob." style="text-align:right;color:${ap.no_aprobados>0?'#e24b4a':'var(--texto-suave)'};">${ap.no_aprobados}</td>
      <td data-label="Progreso">
        <div style="font-size:10px;color:var(--texto-suave);">${pct}%</div>
        <div class="db-pb"><div class="db-pf" style="width:${pct}%"></div></div>
      </td>
    </tr>`;
  });

  document.getElementById('dbTbody').innerHTML = html;
}

/* ── Consulta por competencia y RA ─────────────────────────────── */
let _dbConsultaData = null;
let _dbModoConsulta = 'todos';

function dbSetToggle(m) {
  _dbModoConsulta = m;
  document.getElementById('dbTogTodos').className = 'db-tog-btn' + (m === 'todos' ? ' db-tog-todos' : '');
  document.getElementById('dbTogActivos').className = 'db-tog-btn' + (m === 'activos' ? ' db-tog-activos' : '');
  const ficha = document.getElementById('db_ficha').value.trim();
  if (ficha && _dbConsultaData) _dbCargarConsulta(ficha).then(dbRenderConsulta);
  else dbRenderConsulta();
}

async function _dbCargarConsulta(ficha) {
  const soloActivos = _dbModoConsulta === 'activos';
  const r = await fetch(`/api/consulta_ra/${encodeURIComponent(ficha)}?solo_activos=${soloActivos}`);
  if (!r.ok) return;
  _dbConsultaData = await r.json();
  const selComp = document.getElementById('dbSelComp');
  selComp.innerHTML = '<option value="">— Selecciona una competencia —</option>'
    + _dbConsultaData.competencias.map((c, i) =>
        `<option value="${i}">${c.competencia}</option>`
      ).join('');
  selComp.disabled = false;
  document.getElementById('dbSelRA').disabled = true;
  document.getElementById('dbSelRA').innerHTML = '<option value="">— Selecciona primero una competencia —</option>';
  document.getElementById('dbConsultaPlaceholder').textContent = 'Selecciona una competencia y un resultado de aprendizaje';
  document.getElementById('dbConsultaPlaceholder').style.display = 'block';
  document.getElementById('dbConsultaResult').style.display = 'none';
}

function dbOnCompChange() {
  const ci = document.getElementById('dbSelComp').value;
  const selRA = document.getElementById('dbSelRA');
  document.getElementById('dbConsultaResult').style.display = 'none';
  document.getElementById('dbConsultaPlaceholder').style.display = 'block';
  document.getElementById('dbConsultaPlaceholder').textContent = 'Selecciona un resultado de aprendizaje';
  if (ci === '') { selRA.disabled = true; selRA.innerHTML = '<option value="">— Selecciona primero una competencia —</option>'; return; }
  const comp = _dbConsultaData.competencias[parseInt(ci)];
  selRA.innerHTML = '<option value="">— Selecciona un resultado —</option>'
    + comp.ras.map((r, i) => `<option value="${i}">${r.ra}</option>`).join('');
  selRA.disabled = false;
}

function dbRenderConsulta() {
  const ci = document.getElementById('dbSelComp').value;
  const ri = document.getElementById('dbSelRA').value;
  if (ci === '' || ri === '') return;

  const ORDEN_ESTADO = ['EN FORMACION','CONDICIONADO','APLAZADO','TRASLADADO','RETIRO VOLUNTARIO','CANCELADO'];
  const ACTIVOS = ['EN FORMACION', 'CONDICIONADO'];
  const ESTADO_BADGE = {'EN FORMACION':'db-badge-b','CONDICIONADO':'db-badge-a'};
  const ESTADO_LABEL = {
    'EN FORMACION':'En formación','CONDICIONADO':'Condicionado',
    'APLAZADO':'Aplazado','RETIRO VOLUNTARIO':'Retiro vol.',
    'CANCELADO':'Cancelado','TRASLADADO':'Trasladado',
  };

  const d = _dbConsultaData.competencias[parseInt(ci)].ras[parseInt(ri)];
  const filtrar = a => _dbModoConsulta === 'todos' || ACTIVOS.includes(a.estado);

  const evaluados   = d.evaluados.filter(filtrar);
  const sinEvaluar  = d.sin_evaluar.filter(filtrar);
  const aprobados   = evaluados.filter(a => a.juicio === 'APROBADO').length;
  const noAprobados = evaluados.filter(a => a.juicio === 'NO APROBADO').length;

  document.getElementById('dbCmEval').textContent    = evaluados.length;
  document.getElementById('dbCmAprob').textContent   = aprobados;
  document.getElementById('dbCmNoAprob').textContent = noAprobados;
  document.getElementById('dbCmSinEval').textContent = sinEvaluar.length;

  const eBadge = e => `<span class="db-badge ${ESTADO_BADGE[e]||'db-badge-gr'}">${ESTADO_LABEL[e]||e}</span>`;

  document.getElementById('dbCmTbEval').innerHTML = evaluados.length === 0
    ? `<tr><td colspan="5" style="text-align:center;color:var(--texto-suave);padding:12px;font-size:12px;">Sin registros de evaluación</td></tr>`
    : evaluados.map(a => `<tr>
        <td>${a.nombre}</td>
        <td data-label="Estado">${eBadge(a.estado)}</td>
        <td data-label="Juicio"><span class="db-badge ${a.juicio==='APROBADO'?'db-badge-g':'db-badge-r'}">${a.juicio==='APROBADO'?'Aprobado':'No aprobado'}</span></td>
        <td data-label="Instructor">${a.instructor}</td>
        <td data-label="Fecha">${a.fecha}</td>
      </tr>`).join('');

  const sinEvalSorted = [...sinEvaluar].sort((a,b) =>
    ORDEN_ESTADO.indexOf(a.estado) - ORDEN_ESTADO.indexOf(b.estado));

  document.getElementById('dbCmTbSinEval').innerHTML = sinEvalSorted.length === 0
    ? `<tr><td colspan="2" style="text-align:center;color:var(--texto-suave);padding:12px;font-size:12px;">Todos los aprendices tienen este resultado evaluado.</td></tr>`
    : sinEvalSorted.map(a => `<tr>
        <td>${a.nombre}</td>
        <td data-label="Estado">${eBadge(a.estado)}</td>
      </tr>`).join('');

  // Marcar tabla sin evaluar para layout de 2 columnas en móvil
  const tblSinEval = document.getElementById('dbCmTbSinEval').closest('table');
  if (tblSinEval) tblSinEval.classList.add('tbl-2col');

  document.getElementById('dbConsultaPlaceholder').style.display = 'none';
  document.getElementById('dbConsultaResult').style.display = 'block';
}

function dbSwitchInner(idx) {
  [0, 1, 2].forEach(i => {
    document.getElementById(`dbITab${i}`).classList.toggle('active', i === idx);
    document.getElementById(`dbIPanel${i}`).style.display = i === idx ? '' : 'none';
  });
  const sel = document.getElementById('dbInnerSelect');
  if (sel) sel.value = idx;
}

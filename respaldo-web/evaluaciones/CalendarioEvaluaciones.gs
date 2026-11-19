/************************************************************
 * CALENDARIO DE EVALUACIONES (lee y escribe directo en MAESTRO)
 *
 * Réplica en Apps Script del calendario de pruebas de la web app
 * CalendarioHorariosWeb, pero sin base de datos: todo se lee y se
 * guarda en la hoja MAESTRO.
 *
 * - Cursos: solo filas con CURSO MANDANTE = "SI".
 * - Qué se programa por curso:
 *     HORARIO EVALUACIONES (19:30 o en clases o en ayudantías)
 *         CLASE / AYUDANTIA / LABORATORIO  → en un bloque de ese tipo
 *                                            del horario semanal
 *                                            (columnas Lunes..Viernes)
 *         TARDE                            → 19:30-21:20
 *     CANTIDAD EVALUACIONES (semestrales)  → cuántas
 *     EXAMEN (Sí o No)                     → si lleva examen
 * - Dónde se guarda: en la celda (fila del curso, columna de la fecha)
 *   con el mismo formato que leen las revisiones:
 *     "CLAS 10:30-13:20", "AYUD 13:30-15:20", "LAB/TALLER 8:30-10:20",
 *     "TARDE 19:30-21:20", "EXAM 9:30-11:20"
 *   Si una celda tiene más de una, se separan con ", ".
 *
 * Cómo abrirlo:
 *   a) Menú de la planilla: agregar en onOpen()
 *        .addItem('Calendario de evaluaciones', 'abrirCalendarioEvaluaciones')
 *   b) Como web app: Implementar → Nueva implementación → Aplicación web
 *      (doGet ya está definido abajo).
 *
 * Archivos del proyecto:
 *   - CalendarioEvaluaciones.gs   (este archivo)
 *   - CalendarioEvaluaciones.html (interfaz)
 ************************************************************/

const CEV_CONFIG = {
  // ID del documento externo que contiene la hoja MAESTRO.
  SOURCE_SPREADSHEET_ID: '1cmytVU4wEEMPz-E2Q7lJozkynGJCLpOHcIVRaemh1Pg',
  SOURCE_SHEET_NAME: 'MAESTRO',

  FECHA_INICIO: '2027-03-02',
  FECHA_FIN: '2027-07-09',

  // Período de exámenes: se pinta celeste en el calendario y se avisa si un
  // EXAMEN queda fuera de él.
  EXAMENES_INICIO: '2027-06-24',
  EXAMENES_FIN: '2027-07-09',

  // Salas especiales de evaluación (columnas "Sala especial PRUEB" y
  // "Sala especial EXAM"). Las no exclusivas son etiquetas genéricas que pueden
  // ser varias salas, por lo que no se revisan topes con ellas.
  SALAS_NO_EXCLUSIVAS: ['SALA CON ENCHUFE INDIVIDUAL'],
  // Hoja con capacidades (NOMBRE BANNER / CAPACIDAD). Se busca en la planilla
  // del MAESTRO y, si no está, en la planilla activa (script de revisiones).
  SALAS_ESPECIALES_SHEET: 'SALAS ESPECIALES',

  TITULO: 'Calendario de Evaluaciones',

  FERIADOS_PROP: 'CEV_FERIADOS',

  TARDE_INICIO: '19:30',
  TARDE_FIN: '21:20',
  TARDE_DIAS_PERMITIDOS: [1, 4], // 1 = lunes, 4 = jueves (regla de las revisiones)

  BLOQUES_EXAMEN: [
    ['8:30', '10:20'], ['9:30', '11:20'], ['10:30', '12:20'], ['11:30', '13:20'],
    ['12:30', '14:20'], ['13:30', '15:20'], ['14:30', '16:20'], ['15:30', '17:20'],
    ['16:30', '18:20'], ['17:30', '19:20']
  ],

  // Bloques seguidos del mismo tipo con recreo <= a esto se unen (10:30-11:20 + 11:30-12:20)
  MAX_RECREO_MIN: 10
};

const CEV_MAJORS = [
  { key: 'PLAN COMÚN', label: 'Plan Común', candidates: ['Plan Común', 'Plan Comun', 'PLAN COMÚN', 'PLAN COMUN'] },
  { key: 'ICI', label: 'ICI', candidates: ['ICI'] },
  { key: 'IOC', label: 'IOC', candidates: ['IOC'] },
  { key: 'ICE', label: 'ICE', candidates: ['ICE'] },
  { key: 'ICC', label: 'ICC', candidates: ['ICC'] },
  { key: 'ICA', label: 'ICA', candidates: ['ICA'] },
  { key: 'ICQ', label: 'ICQ', candidates: ['ICQ'] }
];

const CEV_DIAS = [
  { idx: 1, nombre: 'Lunes', corto: 'Lun', headers: ['Lunes'] },
  { idx: 2, nombre: 'Martes', corto: 'Mar', headers: ['Martes'] },
  { idx: 3, nombre: 'Miércoles', corto: 'Mié', headers: ['Miercoles', 'Miércoles'] },
  { idx: 4, nombre: 'Jueves', corto: 'Jue', headers: ['Jueves'] },
  { idx: 5, nombre: 'Viernes', corto: 'Vie', headers: ['Viernes'] }
];

const CEV_TIPO_LABEL = { CLAS: 'CLAS', AYUD: 'AYUD', LABT: 'LAB/TALLER', TARDE: 'TARDE', EXAM: 'EXAM' };

const CEV_TOKEN_RE = /(CLASES|CLASE|CLAS|AYUDANT[IÍ]AS?|AYUD|LAB\s*\/\s*TALLER|LABT|LABORATORIO|LAB|TALLER|TARDE|EXAMEN|EXAM)\s*(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})/gi;

/************************************************************
 * ENTRADAS
 ************************************************************/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('CalendarioEvaluaciones')
    .setTitle(CEV_CONFIG.TITULO)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function abrirCalendarioEvaluaciones() {
  const html = HtmlService.createHtmlOutputFromFile('CalendarioEvaluaciones').setWidth(1600).setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, CEV_CONFIG.TITULO);
}

/************************************************************
 * LECTURA
 ************************************************************/

function cevObtenerDatos() {
  const ctx = cevContexto_();
  const { sh, idx, headerRow, dateCols } = ctx;

  const lastRow = cevUltimaFilaConDatos_(sh, idx.codigo + 1, headerRow);
  const numRows = Math.max(0, lastRow - headerRow);
  const lastCol = Math.max.apply(null, cevColumnasUsadas_(idx, dateCols)) + 1;

  const display = numRows
    ? sh.getRange(headerRow + 1, 1, numRows, lastCol).getDisplayValues()
    : [];

  const courses = [];
  const evals = [];

  display.forEach((row, i) => {
    const sheetRow = headerRow + 1 + i;
    if (cevNorm_(row[idx.cursoMandante]) !== 'SI') return;

    const codigo = cevStr_(row[idx.codigo]);
    const seccion = cevStr_(row[idx.secciones]);
    if (!codigo) return;

    const course = cevArmarCurso_(row, idx, sheetRow);
    courses.push(course);

    dateCols.forEach(dc => {
      const text = cevStr_(row[dc.col]);
      if (!text) return;
      const parsed = cevParsearCelda_(text);
      parsed.entries.forEach(en => {
        evals.push({ courseId: course.id, date: dc.date, type: en.type, inicio: en.inicio, fin: en.fin, token: en.token });
      });
      if (parsed.resto) {
        evals.push({ courseId: course.id, date: dc.date, type: 'OTRO', inicio: '', fin: '', token: parsed.resto });
      }
    });

    if (!seccion) course.issues.push('Fila sin SECCIONES.');
  });

  return {
    ok: true,
    titulo: CEV_CONFIG.TITULO,
    fuente: ctx.ss.getName() + ' › ' + sh.getName(),
    url: ctx.ss.getUrl(),
    cargado: Utilities.formatDate(new Date(), ctx.tz, 'dd/MM/yyyy HH:mm'),
    rango: { inicio: CEV_CONFIG.FECHA_INICIO, fin: CEV_CONFIG.FECHA_FIN },
    fechasConColumna: dateCols.map(d => d.date),
    tarde: { inicio: CEV_CONFIG.TARDE_INICIO, fin: CEV_CONFIG.TARDE_FIN, dias: CEV_CONFIG.TARDE_DIAS_PERMITIDOS },
    bloquesExamen: CEV_CONFIG.BLOQUES_EXAMEN.map(b => ({ inicio: b[0], fin: b[1] })),
    feriados: cevLeerFeriados_(),
    examenes: { inicio: CEV_CONFIG.EXAMENES_INICIO, fin: CEV_CONFIG.EXAMENES_FIN },
    salas: cevSalasEspeciales_(ctx),
    courses,
    evals
  };
}

function cevArmarCurso_(row, idx, sheetRow) {
  const issues = [];

  // Cohortes (carrera + semestre). Acepta 7, "9g", "5hi", "7-8", "7, 8".
  const cohorts = [];
  idx.majors.forEach(m => {
    if (m.col < 0) return;
    cevExtraerSemestres_(row[m.col]).forEach(s => {
      if (!cohorts.some(c => c.major === m.key && c.num === s.num && c.suf === s.suf)) {
        cohorts.push({ major: m.key, label: m.label, num: s.num, suf: s.suf, token: s.token });
      }
    });
  });

  // Horario semanal por tipo
  const bloques = { CLAS: [], AYUD: [], LABT: [] };
  const semanal = [];
  CEV_DIAS.forEach(d => {
    const col = idx.dias[d.idx];
    const text = col >= 0 ? cevStr_(row[col]) : '';
    if (!text) return;
    semanal.push(d.corto + ': ' + text.replace(/\s*\n\s*/g, ', '));
    const porTipo = { CLAS: [], AYUD: [], LABT: [] };
    cevParsearCelda_(text).entries.forEach(en => {
      if (porTipo[en.type]) porTipo[en.type].push({ start: cevMin_(en.inicio), end: cevMin_(en.fin) });
    });
    Object.keys(porTipo).forEach(t => {
      cevUnirIntervalos_(porTipo[t]).forEach(iv => {
        bloques[t].push({ dia: d.nombre, diaIdx: d.idx, corto: d.corto, inicio: cevHora_(iv.start), fin: cevHora_(iv.end) });
      });
    });
  });

  // Reglas de evaluación
  const tipoEvalRaw = idx.horarioEval >= 0 ? cevStr_(row[idx.horarioEval]) : '';
  const tipoEval = cevTipoEvaluacion_(tipoEvalRaw);
  const cantidadRaw = idx.cantidad >= 0 ? cevStr_(row[idx.cantidad]) : '';
  const cantidad = cantidadRaw === '' ? null : (parseInt(cantidadRaw.replace(/[^\d]/g, ''), 10) || 0);
  const examenRaw = idx.examen >= 0 ? cevStr_(row[idx.examen]) : '';
  const examenNorm = cevNorm_(examenRaw);
  const examen = examenNorm === '' ? null : (examenNorm === 'SI' || examenNorm === 'S' || examenNorm === 'YES');

  if (cantidad == null) issues.push('CANTIDAD EVALUACIONES vacía.');
  if (examen == null) issues.push('EXAMEN (Sí o No) vacío.');
  if ((cantidad || 0) > 0 && !tipoEval) {
    issues.push(tipoEvalRaw
      ? `HORARIO EVALUACIONES no reconocido ("${tipoEvalRaw}").`
      : 'Falta HORARIO EVALUACIONES (no se sabe en qué horario evaluar).');
  }
  if ((cantidad || 0) > 0 && ['CLAS', 'AYUD', 'LABT'].includes(tipoEval) && !bloques[tipoEval].length) {
    issues.push(`Evalúa en ${cevTipoNombre_(tipoEval)}, pero no tiene ${CEV_TIPO_LABEL[tipoEval]} en el horario semanal.`);
  }
  if (!cohorts.length) issues.push('Sin semestre en Plan Común / ICI / IOC / ICE / ICC / ICA / ICQ.');

  const rut1 = idx.rut1 >= 0 ? cevRut_(row[idx.rut1]) : '';
  const rut2 = idx.rut2 >= 0 ? cevRut_(row[idx.rut2]) : '';

  return {
    id: String(sheetRow),
    row: sheetRow,
    codigo: cevStr_(row[idx.codigo]),
    seccion: cevStr_(row[idx.secciones]),
    titulo: cevStr_(row[idx.titulo]),
    area: idx.area >= 0 ? cevStr_(row[idx.area]) : '',
    lc: idx.lc >= 0 ? cevStr_(row[idx.lc]) : '',
    cohorts,
    tipoEval,
    tipoEvalRaw,
    cantidad,
    examen,
    examenRaw,
    bloques,
    semanal: semanal.join(' | '),
    profs: [rut1, rut2].filter(Boolean),
    profNombre: idx.nombre1 >= 0 ? cevStr_(row[idx.nombre1]) : '',
    salaPrueb: idx.salaPrueb >= 0 ? cevSala_(row[idx.salaPrueb]) : '',
    salaExam: idx.salaExam >= 0 ? cevSala_(row[idx.salaExam]) : '',
    cupos: idx.cupos >= 0 ? (parseInt(cevStr_(row[idx.cupos]).replace(/[^\d]/g, ''), 10) || 0) : 0,
    issues
  };
}

/************************************************************
 * ESCRITURA
 *
 * items: [{ courseId, codigo, seccion, date, type, inicio, fin }]
 ************************************************************/

function cevAgregar(items) {
  return cevConLock_(ctx => {
    const out = [];
    (items || []).forEach(it => {
      const token = cevValidarToken_(it);
      const loc = cevUbicar_(ctx, it);
      const cell = ctx.sh.getRange(loc.row, loc.col);
      const text = cevStr_(cell.getDisplayValue());
      const ya = cevParsearCelda_(text).entries.some(en => en.token === token);
      if (!ya) cell.setValue(text ? text + ', ' + token : token);
      out.push(cevCeldaResultado_(ctx, loc, it));
    });
    return { ok: true, cells: out };
  });
}

function cevEliminar(item) {
  return cevConLock_(ctx => {
    const loc = cevUbicar_(ctx, item);
    const cell = ctx.sh.getRange(loc.row, loc.col);
    const text = cevStr_(cell.getDisplayValue());
    cell.setValue(cevQuitarToken_(text, item.token));
    return { ok: true, cells: [cevCeldaResultado_(ctx, loc, item)] };
  });
}

/**
 * item: { courseId, codigo, seccion, date, token }   (origen)
 * destino: { date, type, inicio, fin }               (nuevo)
 */
function cevMover(item, destino) {
  return cevConLock_(ctx => {
    const nuevo = Object.assign({}, item, destino);
    const nuevoToken = cevValidarToken_(nuevo);

    const locA = cevUbicar_(ctx, item);
    const cellA = ctx.sh.getRange(locA.row, locA.col);
    cellA.setValue(cevQuitarToken_(cevStr_(cellA.getDisplayValue()), item.token));

    const locB = cevUbicar_(ctx, nuevo);
    const cellB = ctx.sh.getRange(locB.row, locB.col);
    const textB = cevStr_(cellB.getDisplayValue());
    if (!cevParsearCelda_(textB).entries.some(en => en.token === nuevoToken)) {
      cellB.setValue(textB ? textB + ', ' + nuevoToken : nuevoToken);
    }

    return { ok: true, cells: [cevCeldaResultado_(ctx, locA, item), cevCeldaResultado_(ctx, locB, nuevo)] };
  });
}

/** Mueve varias evaluaciones (secciones juntas) en una sola operación. lista: [{ item, destino }] */
function cevMoverVarios(lista) {
  return cevConLock_(ctx => {
    const out = [];
    (lista || []).forEach(x => {
      const nuevo = Object.assign({}, x.item, x.destino);
      const nuevoToken = cevValidarToken_(nuevo);
      const locA = cevUbicar_(ctx, x.item);
      const cellA = ctx.sh.getRange(locA.row, locA.col);
      cellA.setValue(cevQuitarToken_(cevStr_(cellA.getDisplayValue()), x.item.token));
      const locB = cevUbicar_(ctx, nuevo);
      const cellB = ctx.sh.getRange(locB.row, locB.col);
      const textB = cevStr_(cellB.getDisplayValue());
      if (!cevParsearCelda_(textB).entries.some(en => en.token === nuevoToken)) {
        cellB.setValue(textB ? textB + ', ' + nuevoToken : nuevoToken);
      }
      out.push(cevCeldaResultado_(ctx, locA, x.item), cevCeldaResultado_(ctx, locB, nuevo));
    });
    return { ok: true, cells: out };
  });
}

/** Quita varias evaluaciones (secciones juntas) en una sola operación. */
function cevEliminarVarios(items) {
  return cevConLock_(ctx => {
    const out = [];
    (items || []).forEach(item => {
      const loc = cevUbicar_(ctx, item);
      const cell = ctx.sh.getRange(loc.row, loc.col);
      cell.setValue(cevQuitarToken_(cevStr_(cell.getDisplayValue()), item.token));
      out.push(cevCeldaResultado_(ctx, loc, item));
    });
    return { ok: true, cells: out };
  });
}

/**
 * Guarda (o quita) un feriado.
 *   datos = null                                   → quita el feriado
 *   datos = { tipo: 'completo', motivo }           → día completo bloqueado
 *   datos = { tipo: 'parcial', desde, hasta, motivo } → solo ese intervalo bloqueado
 * Devuelve la lista completa de feriados.
 */
function cevGuardarFeriado(date, datos) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error('Fecha inválida.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const list = cevLeerFeriados_().filter(f => f.date !== date);
    if (datos) {
      const tipo = datos.tipo === 'parcial' ? 'parcial' : 'completo';
      const item = { date, tipo, motivo: cevStr_(datos.motivo).slice(0, 120) };
      if (tipo === 'parcial') {
        const d = cevMin_(datos.desde), h = cevMin_(datos.hasta);
        if (d == null || h == null || h <= d) throw new Error('Intervalo inválido: ' + datos.desde + '–' + datos.hasta);
        item.desde = cevHora_(d);
        item.hasta = cevHora_(h);
      }
      list.push(item);
    }
    list.sort((x, y) => x.date.localeCompare(y.date));
    PropertiesService.getScriptProperties().setProperty(CEV_CONFIG.FERIADOS_PROP, JSON.stringify(list));
    return list;
  } finally {
    lock.releaseLock();
  }
}

// Compatibilidad con la versión anterior (alterna feriado completo).
function cevToggleFeriado(date) {
  const existe = cevLeerFeriados_().some(f => f.date === date);
  return cevGuardarFeriado(date, existe ? null : { tipo: 'completo' });
}

function cevConLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ctx = cevContexto_();
    const res = fn(ctx);
    SpreadsheetApp.flush();
    return res;
  } finally {
    lock.releaseLock();
  }
}

function cevCeldaResultado_(ctx, loc, it) {
  const text = cevStr_(ctx.sh.getRange(loc.row, loc.col).getDisplayValue());
  const parsed = cevParsearCelda_(text);
  const entries = parsed.entries.map(en => ({ type: en.type, inicio: en.inicio, fin: en.fin, token: en.token }));
  if (parsed.resto) entries.push({ type: 'OTRO', inicio: '', fin: '', token: parsed.resto });
  return { courseId: String(loc.row), oldCourseId: String(it.courseId), date: it.date, entries };
}

function cevValidarToken_(it) {
  const type = String(it.type || '').toUpperCase();
  if (!CEV_TIPO_LABEL[type]) throw new Error('Tipo de evaluación inválido: ' + it.type);
  if (!(String(it.date) >= CEV_CONFIG.FECHA_INICIO && String(it.date) <= CEV_CONFIG.FECHA_FIN)) {
    throw new Error('La fecha ' + it.date + ' está fuera del rango ' + CEV_CONFIG.FECHA_INICIO + ' a ' + CEV_CONFIG.FECHA_FIN + '.');
  }
  if (!cevStr_(it.inicio) && !cevStr_(it.fin)) return CEV_TIPO_LABEL[type]; // evaluación sin hora
  const ini = cevMin_(it.inicio);
  const fin = cevMin_(it.fin);
  if (ini == null || fin == null || fin <= ini) throw new Error('Horario inválido: ' + it.inicio + '-' + it.fin);
  return CEV_TIPO_LABEL[type] + ' ' + cevHora_(ini) + '-' + cevHora_(fin);
}

/**
 * Busca la fila del curso. Usa el número de fila recibido y verifica
 * CODIGO + SECCIONES; si alguien insertó filas, lo busca de nuevo.
 */
function cevUbicar_(ctx, it) {
  const dc = ctx.dateCols.find(d => d.date === it.date);
  if (!dc) throw new Error('No existe una columna para la fecha ' + it.date + ' en ' + ctx.sh.getName() + '.');

  const codigo = cevNorm_(it.codigo);
  const seccion = cevNorm_(it.seccion);
  const row = Number(it.courseId);

  const coincide = r => {
    const vals = ctx.sh.getRange(r, 1, 1, Math.max(ctx.idx.codigo, ctx.idx.secciones) + 1).getDisplayValues()[0];
    return cevNorm_(vals[ctx.idx.codigo]) === codigo && cevNorm_(vals[ctx.idx.secciones]) === seccion;
  };

  if (row > ctx.headerRow && coincide(row)) return { row, col: dc.col + 1 };

  const last = cevUltimaFilaConDatos_(ctx.sh, ctx.idx.codigo + 1, ctx.headerRow);
  if (last > ctx.headerRow) {
    const n = last - ctx.headerRow;
    const cod = ctx.sh.getRange(ctx.headerRow + 1, ctx.idx.codigo + 1, n, 1).getDisplayValues();
    const sec = ctx.sh.getRange(ctx.headerRow + 1, ctx.idx.secciones + 1, n, 1).getDisplayValues();
    for (let i = 0; i < n; i++) {
      if (cevNorm_(cod[i][0]) === codigo && cevNorm_(sec[i][0]) === seccion) {
        return { row: ctx.headerRow + 1 + i, col: dc.col + 1 };
      }
    }
  }
  throw new Error(`No se encontró ${it.codigo} sec. ${it.seccion} en ${ctx.sh.getName()}. Recarga el calendario.`);
}

/************************************************************
 * CONTEXTO / ENCABEZADOS
 ************************************************************/

function cevContexto_() {
  let id = cevStr_(CEV_CONFIG.SOURCE_SPREADSHEET_ID);
  if (!id && typeof REV_CONFIG !== 'undefined' && REV_CONFIG && REV_CONFIG.SOURCE_SPREADSHEET_ID) {
    id = cevStr_(REV_CONFIG.SOURCE_SPREADSHEET_ID);
  }

  let ss = null;
  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      throw new Error('No se pudo abrir la planilla con ID ' + id + '. Revisa el ID y tus permisos. ' + e.message);
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) throw new Error('Configura CEV_CONFIG.SOURCE_SPREADSHEET_ID con el ID de la planilla del MAESTRO.');

  const sh = ss.getSheetByName(CEV_CONFIG.SOURCE_SHEET_NAME);
  if (!sh) throw new Error(`No existe la hoja "${CEV_CONFIG.SOURCE_SHEET_NAME}" en "${ss.getName()}".`);

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const lastCol = sh.getLastColumn();
  const topRows = Math.min(25, sh.getLastRow());
  const top = sh.getRange(1, 1, topRows, lastCol).getDisplayValues();

  let headerIdx = -1;
  for (let r = 0; r < top.length; r++) {
    if (top[r].some(v => cevNorm_(v) === 'CURSO MANDANTE')) { headerIdx = r; break; }
  }
  if (headerIdx < 0) throw new Error('No se encontró la fila de encabezados (CURSO MANDANTE) en ' + sh.getName() + '.');

  const headerRow = headerIdx + 1;
  const headers = top[headerIdx];
  const headersRaw = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  const idx = cevIndices_(headers);
  const dateCols = cevColumnasFecha_(headersRaw, headers, tz);

  return { ss, sh, tz, headerRow, headers, idx, dateCols };
}

function cevIndices_(headers) {
  const f = (cands, opts) => cevBuscarHeader_(headers, cands, opts);
  const dias = {};
  CEV_DIAS.forEach(d => {
    // Exacto (Lunes/Martes...). Si no, la ÚLTIMA aparición normalizada:
    // la disponibilidad (LUNES) viene antes que el horario (Lunes).
    dias[d.idx] = f(d.headers, { exactCase: true, fallbackLast: true });
  });

  const idx = {
    cursoMandante: f(['CURSO MANDANTE']),
    codigo: f(['CODIGO', 'CÓDIGO']),
    titulo: f(['TITULO', 'TÍTULO']),
    secciones: f(['SECCIONES', 'SECCION', 'SECCIÓN']),
    area: f(['AREA', 'ÁREA']),
    lc: f(['LC', 'LISTA CRUZADA']),
    examen: f(['EXAMEN (Sí o No)', 'EXAMEN (Si o No)']),
    horarioEval: f(['HORARIO EVALUACIONES (19:30 o en clases o en ayudantías)', 'HORARIO EVALUACIONES (19:30 o en clases o en ayudantias)']),
    cantidad: f(['CANTIDAD EVALUACIONES (semestrales)']),
    rut1: f(['RUT PROFESOR 1']),
    rut2: f(['RUT PROFESOR 2']),
    nombre1: f(['NOMBRE PROFESOR BANNER 1', 'NOMBRE PROFESOR 1'], { prefix: true }),
    salaPrueb: f(['Sala especial PRUEB', 'SALA ESPECIAL PRUEBA', 'SALA ESPECIAL PRUEBAS']),
    salaExam: f(['Sala especial EXAM', 'SALA ESPECIAL EXAMEN']),
    cupos: f(['CUPOS']),
    majors: CEV_MAJORS.map(m => ({ key: m.key, label: m.label, col: f(m.candidates) })),
    dias
  };

  const faltan = [];
  [['cursoMandante', 'CURSO MANDANTE'], ['codigo', 'CODIGO'], ['titulo', 'TITULO'], ['secciones', 'SECCIONES'],
   ['examen', 'EXAMEN (Sí o No)'], ['horarioEval', 'HORARIO EVALUACIONES (19:30 o en clases o en ayudantías)'],
   ['cantidad', 'CANTIDAD EVALUACIONES (semestrales)']].forEach(([k, label]) => { if (idx[k] < 0) faltan.push(label); });
  CEV_DIAS.forEach(d => { if (dias[d.idx] < 0) faltan.push(d.headers[0]); });
  if (faltan.length) throw new Error('Faltan columnas en MAESTRO: ' + faltan.join(', '));

  return idx;
}

function cevBuscarHeader_(headers, candidates, opts) {
  opts = opts || {};
  for (const c of candidates) {
    const i = headers.findIndex(h => cevStr_(h) === c);
    if (i >= 0) return i;
  }
  if (opts.exactCase && !opts.fallbackLast) return -1;

  const normC = candidates.map(c => cevNormHeader_(c));
  const matches = [];
  headers.forEach((h, i) => {
    const nh = cevNormHeader_(h);
    if (!nh) return;
    if (normC.some(c => (opts.prefix ? nh.indexOf(c) === 0 : nh === c))) matches.push(i);
  });
  if (!matches.length) return -1;
  return opts.fallbackLast ? matches[matches.length - 1] : matches[0];
}

function cevColumnasFecha_(headersRaw, headersDisplay, tz) {
  const out = [];
  const seen = new Set();
  headersRaw.forEach((h, col) => {
    let key = '';
    if (h instanceof Date && !isNaN(h.getTime())) {
      key = Utilities.formatDate(h, tz, 'yyyy-MM-dd');
    } else {
      key = cevFechaDesdeTexto_(headersDisplay[col]);
    }
    if (!key || key < CEV_CONFIG.FECHA_INICIO || key > CEV_CONFIG.FECHA_FIN) return;
    if (seen.has(key)) return; // si hay columnas duplicadas, usa la primera
    seen.add(key);
    out.push({ date: key, col });
  });
  return out;
}

const CEV_MESES = { ENE: 1, FEB: 2, MAR: 3, ABR: 4, MAY: 5, JUN: 6, JUL: 7, AGO: 8, SEP: 9, SET: 9, OCT: 10, NOV: 11, DIC: 12 };

/**
 * Convierte un encabezado de fecha en texto a "yyyy-MM-dd".
 * Acepta: 2027-03-02 · 02/03/2027 · 2/3/27 · "mar 02 mar" · "02 mar" ·
 *         "2 de marzo" · "02-mar-2027". Si no trae año, usa el año que deja
 *         la fecha dentro del rango configurado.
 */
function cevFechaDesdeTexto_(s) {
  s = cevStr_(s);
  if (!s) return '';
  const p2 = n => ('0' + n).slice(-2);
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + p2(m[2]) + '-' + p2(m[3]);
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?$/);
  if (m) return cevFechaConAnio_(+m[1], +m[2], m[3]);

  const t = cevNorm_(s).replace(/[.,]/g, ' ').replace(/\bDE\b/g, ' ').replace(/\s+/g, ' ').trim();
  // [día semana] DD MES [AAAA]   ej. "MAR 02 MAR", "02 MAR 2027", "2 MARZO"
  m = t.match(/(?:^|\s)(\d{1,2})[\s\-]*([A-Z]{3,})\.?(?:[\s\-]*(\d{2,4}))?$/);
  if (m) {
    const mes = CEV_MESES[m[2].slice(0, 3)];
    if (mes) return cevFechaConAnio_(+m[1], mes, m[3]);
  }
  return '';
}

function cevFechaConAnio_(dia, mes, anioTxt) {
  if (!(dia >= 1 && dia <= 31 && mes >= 1 && mes <= 12)) return '';
  const p2 = n => ('0' + n).slice(-2);
  const arma = y => y + '-' + p2(mes) + '-' + p2(dia);
  if (anioTxt) {
    let y = +anioTxt;
    if (y < 100) y += 2000;
    return arma(y);
  }
  const y0 = +CEV_CONFIG.FECHA_INICIO.slice(0, 4);
  const y1 = +CEV_CONFIG.FECHA_FIN.slice(0, 4);
  for (let y = y0; y <= y1; y++) {
    const f = arma(y);
    if (f >= CEV_CONFIG.FECHA_INICIO && f <= CEV_CONFIG.FECHA_FIN) return f;
  }
  return arma(y0);
}

function cevColumnasUsadas_(idx, dateCols) {
  const cols = [idx.cursoMandante, idx.codigo, idx.titulo, idx.secciones, idx.area, idx.lc, idx.examen,
    idx.horarioEval, idx.cantidad, idx.rut1, idx.rut2, idx.nombre1, idx.salaPrueb, idx.salaExam, idx.cupos];
  idx.majors.forEach(m => cols.push(m.col));
  Object.keys(idx.dias).forEach(k => cols.push(idx.dias[k]));
  dateCols.forEach(d => cols.push(d.col));
  return cols.filter(c => c >= 0);
}

function cevUltimaFilaConDatos_(sh, colNumber, headerRow) {
  const last = sh.getLastRow();
  if (last <= headerRow) return headerRow;
  const vals = sh.getRange(headerRow + 1, colNumber, last - headerRow, 1).getDisplayValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (cevStr_(vals[i][0])) return headerRow + 1 + i;
  }
  return headerRow;
}

/************************************************************
 * PARSEO
 ************************************************************/

function cevParsearCelda_(text) {
  const s = cevStr_(text);
  const entries = [];
  if (!s) return { entries, resto: '' };

  let resto = s;
  const re = new RegExp(CEV_TOKEN_RE.source, 'gi');
  let m;
  while ((m = re.exec(s)) !== null) {
    const type = cevTipoCanonico_(m[1]);
    const ini = cevMin_(m[2]);
    const fin = cevMin_(m[3]);
    if (!type || ini == null || fin == null || fin <= ini) continue;
    entries.push({
      type,
      inicio: cevHora_(ini),
      fin: cevHora_(fin),
      token: CEV_TIPO_LABEL[type] + ' ' + cevHora_(ini) + '-' + cevHora_(fin),
      raw: m[0]
    });
    resto = resto.replace(m[0], ' ');
  }

  // Evaluaciones sin hora (p. ej. "EXAM" o "TARDE", como las exporta la web app
  // cuando no hay horario). Se conservan tal cual para poder quitarlas.
  const reSolo = /(^|[^A-ZÁÉÍÓÚ])(CLASES|CLASE|CLAS|AYUDANT[IÍ]AS?|AYUD|LAB\s*\/\s*TALLER|LABT|LABORATORIO|LAB|TALLER|TARDE|EXAMEN|EXAM)(?![A-ZÁÉÍÓÚ])/i;
  let ms;
  while ((ms = resto.match(reSolo)) !== null) {
    const type = cevTipoCanonico_(ms[2]);
    if (!type) break;
    entries.push({ type, inicio: '', fin: '', token: CEV_TIPO_LABEL[type], raw: ms[2] });
    resto = resto.replace(ms[2], ' ');
  }

  resto = resto.replace(/[\s,;\/|\-]+/g, ' ').trim();
  return { entries, resto };
}

function cevQuitarToken_(text, token) {
  const s = cevStr_(text);
  if (!s) return '';
  const objetivo = cevStr_(token);
  const parsed = cevParsearCelda_(s);
  const i = parsed.entries.findIndex(e => e.token === objetivo);
  let partes;
  if (i >= 0) {
    partes = parsed.entries.filter((_, k) => k !== i).map(e => e.raw.trim());
    if (parsed.resto) partes.push(parsed.resto);
  } else if (parsed.resto && parsed.resto === objetivo) {
    partes = parsed.entries.map(e => e.raw.trim());
  } else {
    return s;
  }
  return partes.filter(Boolean).join(', ');
}

function cevTipoCanonico_(raw) {
  const t = cevNorm_(raw).replace(/\s+/g, '');
  if (t.indexOf('CLAS') === 0) return 'CLAS';
  if (t.indexOf('AYUD') === 0) return 'AYUD';
  if (t.indexOf('LAB') === 0 || t === 'TALLER') return 'LABT';
  if (t === 'TARDE') return 'TARDE';
  if (t.indexOf('EXAM') === 0) return 'EXAM';
  return '';
}

function cevTipoEvaluacion_(raw) {
  const t = cevNorm_(raw);
  if (!t) return '';
  if (t.indexOf('TARDE') >= 0 || t.indexOf('19:30') >= 0) return 'TARDE';
  if (t.indexOf('AYUD') >= 0) return 'AYUD';
  if (t.indexOf('LAB') >= 0 || t.indexOf('TALLER') >= 0) return 'LABT';
  if (t.indexOf('CLAS') >= 0) return 'CLAS';
  return '';
}

function cevTipoNombre_(t) {
  return { CLAS: 'clases', AYUD: 'ayudantía', LABT: 'laboratorio/taller', TARDE: 'tarde (19:30)', EXAM: 'examen' }[t] || t;
}

function cevExtraerSemestres_(value) {
  const s = cevStr_(value);
  if (!s) return [];
  const out = [];
  s.split(/[\s,;\/\-]+|\bY\b|\bU\b/i).forEach(tok => {
    const m = cevStr_(tok).match(/^(\d{1,2})([A-Za-z]*)$/);
    if (!m) return;
    const num = parseInt(m[1], 10);
    if (!num) return;
    out.push({ num, suf: m[2].toUpperCase(), token: m[1] + m[2].toLowerCase() });
  });
  return out;
}

function cevUnirIntervalos_(ivs) {
  if (!ivs.length) return [];
  const sorted = ivs.filter(x => x.start != null && x.end != null).sort((a, b) => a.start - b.start);
  const out = [Object.assign({}, sorted[0])];
  for (let i = 1; i < sorted.length; i++) {
    const last = out[out.length - 1];
    if (sorted[i].start <= last.end + CEV_CONFIG.MAX_RECREO_MIN) last.end = Math.max(last.end, sorted[i].end);
    else out.push(Object.assign({}, sorted[i]));
  }
  return out;
}

/************************************************************
 * SALAS ESPECIALES
 ************************************************************/

function cevSala_(v) {
  const t = cevStr_(v);
  if (!t) return '';
  if (/^(#N\/A|#REF!|#VALUE!|#DIV\/0!|#ERROR!|N\/A|-|—|NO|NINGUNA)$/i.test(t)) return '';
  return t;
}

function cevClaveSala_(v) {
  return cevNorm_(v).replace(/\s+/g, '');
}

/** { noExclusivas: [claves], capacidad: {clave: n}, hojaCapacidad: bool } */
function cevSalasEspeciales_(ctx) {
  const out = {
    noExclusivas: CEV_CONFIG.SALAS_NO_EXCLUSIVAS.map(cevClaveSala_),
    capacidad: {},
    hojaCapacidad: false
  };
  const libros = [ctx.ss];
  try {
    const act = SpreadsheetApp.getActiveSpreadsheet();
    if (act && act.getId && ctx.ss.getId && act.getId() !== ctx.ss.getId()) libros.push(act);
  } catch (e) { /* web app sin planilla activa */ }

  for (const libro of libros) {
    const sh = libro.getSheetByName(CEV_CONFIG.SALAS_ESPECIALES_SHEET);
    if (!sh) continue;
    const vals = sh.getDataRange().getDisplayValues();
    let h = -1, iNom = -1, iCap = -1;
    for (let r = 0; r < Math.min(vals.length, 25); r++) {
      const n = vals[r].map(cevNormHeader_);
      iNom = n.indexOf('NOMBRE BANNER');
      iCap = n.indexOf('CAPACIDAD');
      if (iNom >= 0 && iCap >= 0) { h = r; break; }
    }
    if (h < 0) continue;
    for (let r = h + 1; r < vals.length; r++) {
      const sala = cevStr_(vals[r][iNom]);
      const cap = parseInt(cevStr_(vals[r][iCap]).replace(/[^\d]/g, ''), 10);
      if (sala && cap) out.capacidad[cevClaveSala_(sala)] = cap;
    }
    out.hojaCapacidad = true;
    break;
  }
  return out;
}

/************************************************************
 * FERIADOS
 ************************************************************/

function cevLeerFeriados_() {
  const raw = PropertiesService.getScriptProperties().getProperty(CEV_CONFIG.FERIADOS_PROP);
  if (!raw) return [];
  try {
    const list = JSON.parse(raw);
    if (!Array.isArray(list)) return [];
    return list.map(x => (typeof x === 'string' ? { date: x, tipo: 'completo', motivo: '' } : x))
      .filter(x => x && /^\d{4}-\d{2}-\d{2}$/.test(x.date));
  } catch (e) {
    return [];
  }
}

/************************************************************
 * HELPERS
 ************************************************************/

function cevStr_(v) {
  return v == null ? '' : String(v).trim();
}

function cevNorm_(v) {
  return cevStr_(v).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
}

function cevNormHeader_(v) {
  return cevNorm_(v).replace(/\s+/g, ' ');
}

function cevRut_(v) {
  return cevStr_(v).toUpperCase().replace(/[.\-\s]/g, '');
}

function cevMin_(hhmm) {
  const m = cevStr_(hhmm).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = Number(m[1]);
  const mi = Number(m[2]);
  if (h > 23 || mi > 59) return null;
  return h * 60 + mi;
}

function cevHora_(min) {
  return Math.floor(min / 60) + ':' + ('0' + (min % 60)).slice(-2);
}
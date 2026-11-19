/************************************************************
 * PROGRAMACIÓN DE HORARIOS (lee y escribe directo en MAESTRO)
 *
 * Réplica en Apps Script de la grilla de horarios de la web app
 * CalendarioHorariosWeb, separada del calendario de evaluaciones.
 *
 * - Cursos: solo filas con CURSO MANDANTE = "SI".
 * - Horas a programar: Clases A PROGRAMAR / Ayudantías PROGRAMAR /
 *   Laboratorios o Talleres PROGRAMAR (en bloques de 50 min).
 * - Horario: columnas Lunes..Viernes con el formato que ya usa MAESTRO:
 *     "CLAS 10:30-13:20", "AYUD 13:30-15:20, LAB/TALLER 15:30-17:20"
 *   Bloques de 50 min de 8:30 a 19:20.
 *   Cada cambio en la página reescribe solo la celda de ese curso y día.
 * - Disponibilidad docente: columnas LUNES..VIERNES.
 *
 * Reglas de disponibilidad (validación en la página):
 *   CLASE       → disponibilidad del curso (profesor principal).
 *   AYUDANTÍA   → disponibilidad completa.
 *   LAB/TALLER  → si el curso SOLO tiene LAB/TALLER: disponibilidad del
 *                 profesor 1, o el intervalo común entre profesor 1 y 2.
 *                 si además tiene CLASE y/o AYUDANTÍA: disponibilidad
 *                 completa; si hay RUT PROFESOR LABT, se revisa que ese
 *                 docente no tenga tope con otras horas suyas.
 *
 * Cómo abrirlo:
 *   a) Menú: agregar en onOpen()
 *        .addItem('Programación de horarios', 'abrirProgramacionHorarios')
 *   b) Web app: Implementar → Nueva implementación → Aplicación web.
 *      Proyecto independiente del calendario de evaluaciones (su propio enlace).
 ************************************************************/

const PHR_CONFIG = {
  SOURCE_SPREADSHEET_ID: '1cmytVU4wEEMPz-E2Q7lJozkynGJCLpOHcIVRaemh1Pg',
  SOURCE_SHEET_NAME: 'MAESTRO',
  TITULO: 'Programación de Horarios',


  BLOQUES: [
    ['8:30', '9:20'], ['9:30', '10:20'], ['10:30', '11:20'], ['11:30', '12:20'],
    ['12:30', '13:20'], ['13:30', '14:20'], ['14:30', '15:20'], ['15:30', '16:20'],
    ['16:30', '17:20'], ['17:30', '18:20'], ['18:30', '19:20']
  ],
  BLOQUE_PIVOTE: '12:30',

  // Franjas protegidas (aviso) para Plan Común, 5to-6to y 7mo-8vo.
  PROTEGIDOS: { 2: ['17:30', '18:30'], 3: ['17:30', '18:30'], 5: ['10:30', '11:30', '12:30'] },

  // Salas genéricas que representan varias salas: no se revisan topes.
  SALAS_NO_EXCLUSIVAS: ['SALA CON ENCHUFE INDIVIDUAL'],

  MAX_RECREO_MIN: 10,

  // Aceptaciones de avisos (columna "Revisión" del resumen). Se guardan en las
  // propiedades del proyecto, una por curso-sección, con una huella del horario
  // y de los avisos: si cambian, la aceptación deja de valer y se borra.
  REV_PREFIX: 'PHR_REV_',

  // Profesores JORNADA: alerta si dictan en franjas no ideales.
  // Fuente principal: hoja PROFESORES de la misma planilla del MAESTRO, columna
  // "RUT PROFESOR 1". Si esa hoja tiene una columna de tipo de contrato
  // (encabezado con "JORNADA", "TIPO" o "CONTRATO"), solo se toman las filas cuyo
  // valor contenga "JORNADA"; si no, todos los RUT de la hoja se consideran JORNADA.
  // Respaldo: columna "JORNADA / HONORARIO" de la hoja MAESTRO.
  JORNADAS_SHEET_NAME: 'PROFESORES',
  JORNADAS_RUT_HEADER: 'RUT PROFESOR 1',
  JORNADA_FRANJAS: [
    { dias: [1, 2, 3, 4, 5], desde: '8:30', hasta: '10:20' },
    { dias: [1, 2, 3, 4, 5], desde: '17:30', hasta: '19:20' },
    { dias: [5], desde: '10:30', hasta: '12:20' }
  ]
};

const PHR_MAJORS = [
  { key: 'PLAN COMÚN', label: 'Plan Común', candidates: ['Plan Común', 'Plan Comun', 'PLAN COMÚN', 'PLAN COMUN'] },
  { key: 'ICI', label: 'ICI', candidates: ['ICI'] },
  { key: 'IOC', label: 'IOC', candidates: ['IOC'] },
  { key: 'ICE', label: 'ICE', candidates: ['ICE'] },
  { key: 'ICC', label: 'ICC', candidates: ['ICC'] },
  { key: 'ICA', label: 'ICA', candidates: ['ICA'] },
  { key: 'ICQ', label: 'ICQ', candidates: ['ICQ'] }
];

const PHR_DIAS = [
  { idx: 1, nombre: 'Lunes', horario: ['Lunes'], disp: ['LUNES'] },
  { idx: 2, nombre: 'Martes', horario: ['Martes'], disp: ['MARTES'] },
  { idx: 3, nombre: 'Miércoles', horario: ['Miercoles', 'Miércoles'], disp: ['MIERCOLES', 'MIÉRCOLES'] },
  { idx: 4, nombre: 'Jueves', horario: ['Jueves'], disp: ['JUEVES'] },
  { idx: 5, nombre: 'Viernes', horario: ['Viernes'], disp: ['VIERNES'] }
];

const PHR_TIPO_LABEL = { CLAS: 'CLAS', AYUD: 'AYUD', LABT: 'LAB/TALLER' };
const PHR_TOKEN_RE = /(CLASES|CLASE|CLAS|AYUDANT[IÍ]AS?|AYUD|LAB\s*\/\s*TALLER|LABT|LABORATORIO|LAB|TALLER)\s*(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})/gi;

/************************************************************
 * ENTRADAS
 ************************************************************/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('ProgramacionHorarios')
    .setTitle(PHR_CONFIG.TITULO)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function abrirProgramacionHorarios() {
  const html = HtmlService.createHtmlOutputFromFile('ProgramacionHorarios').setWidth(1600).setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, PHR_CONFIG.TITULO);
}

/************************************************************
 * LECTURA
 ************************************************************/

function phrObtenerDatos() {
  const ctx = phrContexto_();
  const { sh, idx, headerRow } = ctx;

  const lastRow = phrUltimaFila_(sh, idx.codigo + 1, headerRow);
  const n = Math.max(0, lastRow - headerRow);
  const lastCol = Math.max.apply(null, phrColumnasUsadas_(idx)) + 1;
  const rows = n ? sh.getRange(headerRow + 1, 1, n, lastCol).getDisplayValues() : [];

  const courses = [];
  const dispPorRut = {};   // disponibilidad de cada docente (de las filas donde es PROFESOR 1)

  const sinMandante = [];   // filas con código y horas, pero CURSO MANDANTE vacío
  rows.forEach((row, i) => {
    const mand = phrNorm_(row[idx.cursoMandante]);
    if (mand !== 'SI') {
      const cod = phrStr_(row[idx.codigo]);
      const conHoras = [idx.clasesProg, idx.ayudProg, idx.labProg, idx.clasesBase, idx.ayudBase, idx.labBase]
        .some(c => c >= 0 && phrNum_(row[c]) > 0);
      if (cod && !mand && conHoras) sinMandante.push({ codigo: cod, seccion: phrStr_(row[idx.secciones]), titulo: phrStr_(row[idx.titulo]), fila: headerRow + 1 + i });
      return;
    }
    const codigo = phrStr_(row[idx.codigo]);
    if (!codigo) return;
    const c = phrArmarCurso_(row, idx, headerRow + 1 + i);
    courses.push(c);
    if (c.prof1.rut && c.disp && !dispPorRut[c.prof1.rut]) dispPorRut[c.prof1.rut] = c.disp;
  });

  return {
    ok: true,
    titulo: PHR_CONFIG.TITULO,
    fuente: ctx.ss.getName() + ' › ' + sh.getName(),
    url: ctx.ss.getUrl(),
    cargado: Utilities.formatDate(new Date(), ctx.tz, 'dd/MM/yyyy HH:mm'),
    bloques: PHR_CONFIG.BLOQUES.map(b => ({ inicio: b[0], fin: b[1] })),
    pivote: PHR_CONFIG.BLOQUE_PIVOTE,
    protegidos: PHR_CONFIG.PROTEGIDOS,
    salasNoExclusivas: PHR_CONFIG.SALAS_NO_EXCLUSIVAS.map(phrClaveSala_),
    dispPorRut,
    revisiones: phrLeerRevisiones_(),
    jornadaRuts: phrJornadas_(ctx.ss),
    sinMandante,
    jornadaFranjas: PHR_CONFIG.JORNADA_FRANJAS,
    courses
  };
}

function phrArmarCurso_(row, idx, sheetRow) {
  const issues = [];

  const cohorts = [];
  idx.majors.forEach(m => {
    if (m.col < 0) return;
    phrSemestres_(row[m.col]).forEach(s => {
      if (!cohorts.some(c => c.major === m.key && c.num === s.num && c.suf === s.suf)) {
        cohorts.push({ major: m.key, label: m.label, num: s.num, suf: s.suf, token: s.token });
      }
    });
  });

  // Horas a programar: columna "A PROGRAMAR"; si está vacía, se usa Clases / Ayudantías /
  // Laboratorios o Talleres (filas nuevas donde no se copió la fórmula).
  const respaldo = [];
  const horas = (colProg, colBase, nombre) => {
    const t = colProg >= 0 ? phrStr_(row[colProg]) : '';
    if (t !== '') return phrNum_(t);
    const b = colBase >= 0 ? phrStr_(row[colBase]) : '';
    if (b !== '' && phrNum_(b) > 0) { respaldo.push(nombre); return phrNum_(b); }
    return 0;
  };
  const req = {
    CLAS: horas(idx.clasesProg, idx.clasesBase, 'Clases'),
    AYUD: horas(idx.ayudProg, idx.ayudBase, 'Ayudantías'),
    LABT: horas(idx.labProg, idx.labBase, 'Laboratorios o Talleres')
  };
  const soloLab = req.LABT > 0 && req.CLAS === 0 && req.AYUD === 0;

  // Horario semanal (bloques por día y tipo) + texto no reconocido
  const sched = {};
  const otros = {};
  PHR_DIAS.forEach(d => {
    const p = phrParsearCeldaHorario_(phrStr_(row[idx.horario[d.idx]]));
    sched[d.idx] = p.bloques;
    if (p.resto) otros[d.idx] = p.resto;
  });

  // Disponibilidad (columnas LUNES..VIERNES). null = no informada.
  let disp = {};
  let hayDisp = false;
  const dispTexto = {};
  PHR_DIAS.forEach(d => {
    const col = idx.disp[d.idx];
    const t = col >= 0 ? phrStr_(row[col]) : '';
    disp[d.idx] = phrIntervalosDisp_(t);
    if (t) { hayDisp = true; dispTexto[d.idx] = phrResumenIntervalos_(disp[d.idx]); }
  });
  if (!hayDisp) disp = null;

  const prof1 = { rut: idx.rut1 >= 0 ? phrRut_(row[idx.rut1]) : '', nombre: idx.nombre1 >= 0 ? phrStr_(row[idx.nombre1]) : '' };
  const prof2 = { rut: idx.rut2 >= 0 ? phrRut_(row[idx.rut2]) : '', nombre: idx.nombre2 >= 0 ? phrStr_(row[idx.nombre2]) : '' };
  const labtRuts = idx.rutLabt >= 0 ? phrStr_(row[idx.rutLabt]).split(/[;,\n\/]+/).map(phrRut_).filter(Boolean) : [];
  const labtNombre = idx.nombreLabt >= 0 ? phrStr_(row[idx.nombreLabt]) : '';

  if (req.CLAS + req.AYUD + req.LABT === 0) issues.push('Sin horas a programar (CLASES/AYUDANTÍAS/LAB en 0 o vacías).');
  if (respaldo.length) issues.push('"A PROGRAMAR" vacío: se usan las horas de ' + respaldo.join(', ') + '.');
  if (req.CLAS > 0 && !disp) issues.push('Sin disponibilidad docente informada (LUNES..VIERNES): no se valida CLASE.');
  if (soloLab && !disp) issues.push('Curso solo LAB/TALLER sin disponibilidad del profesor 1.');
  if (!cohorts.length) issues.push('Sin semestre en Plan Común / ICI / IOC / ICE / ICC / ICA / ICQ.');
  if (Object.keys(otros).length) issues.push('Texto no reconocido en el horario: ' + Object.keys(otros).map(k => PHR_DIAS[k - 1].nombre + ' "' + otros[k] + '"').join('; '));

  return {
    id: String(sheetRow),
    row: sheetRow,
    codigo: phrStr_(row[idx.codigo]),
    seccion: phrStr_(row[idx.secciones]),
    titulo: phrStr_(row[idx.titulo]),
    area: idx.area >= 0 ? phrStr_(row[idx.area]) : '',
    lc: idx.lc >= 0 ? phrStr_(row[idx.lc]) : '',
    cupos: idx.cupos >= 0 ? phrNum_(row[idx.cupos]) : 0,
    cohorts,
    req,
    soloLab,
    distribucion: idx.distribucion >= 0 ? phrStr_(row[idx.distribucion]) : '',
    salas: {
      CLAS: idx.salaClas >= 0 ? phrSala_(row[idx.salaClas]) : '',
      AYUD: idx.salaAyud >= 0 ? phrSala_(row[idx.salaAyud]) : '',
      LABT: idx.salaLab >= 0 ? phrSala_(row[idx.salaLab]) : ''
    },
    prof1, prof2, labtRuts, labtNombre,
    jornada: idx.jornada >= 0 && phrNorm_(row[idx.jornada]).indexOf('JORNADA') >= 0,
    disp, dispTexto,
    sched, otros,
    issues
  };
}

/************************************************************
 * ESCRITURA
 *
 * ops: [{ courseId, codigo, seccion, dia (1-5), type (CLAS|AYUD|LABT),
 *         bloque (índice 0..11), accion ('add'|'remove') }]
 * Se agrupan por fila y día; cada celda se lee, se modifica y se reescribe.
 ************************************************************/

function phrModificar(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ctx = phrContexto_();
    const grupos = {};
    (ops || []).forEach(op => {
      if (!PHR_TIPO_LABEL[op.type]) throw new Error('Tipo inválido: ' + op.type);
      const dia = Number(op.dia), b = Number(op.bloque);
      if (!(dia >= 1 && dia <= 5)) throw new Error('Día inválido: ' + op.dia);
      if (!(b >= 0 && b < PHR_CONFIG.BLOQUES.length)) throw new Error('Bloque inválido: ' + op.bloque);
      const k = op.courseId + '|' + dia;
      (grupos[k] = grupos[k] || { op, dia, lista: [] }).lista.push({ type: op.type, b, accion: op.accion });
    });

    const cells = [];
    Object.keys(grupos).forEach(k => {
      const g = grupos[k];
      const row = phrUbicarFila_(ctx, g.op);
      const col = ctx.idx.horario[g.dia] + 1;
      const cell = ctx.sh.getRange(row, col);
      const p = phrParsearCeldaHorario_(phrStr_(cell.getDisplayValue()));
      g.lista.forEach(x => {
        const set = new Set(p.bloques[x.type]);
        if (x.accion === 'remove') set.delete(x.b); else set.add(x.b);
        p.bloques[x.type] = Array.from(set).sort((a, b) => a - b);
      });
      cell.setValue(phrArmarCeldaHorario_(p));
      const nuevo = phrParsearCeldaHorario_(phrStr_(cell.getDisplayValue()));
      cells.push({ courseId: String(row), oldCourseId: String(g.op.courseId), dia: g.dia, bloques: nuevo.bloques, resto: nuevo.resto });
    });
    SpreadsheetApp.flush();
    return { ok: true, cells };
  } finally {
    lock.releaseLock();
  }
}

function phrUbicarFila_(ctx, it) {
  const codigo = phrNorm_(it.codigo), seccion = phrNorm_(it.seccion);
  const row = Number(it.courseId);
  const maxc = Math.max(ctx.idx.codigo, ctx.idx.secciones) + 1;
  if (row > ctx.headerRow) {
    const v = ctx.sh.getRange(row, 1, 1, maxc).getDisplayValues()[0];
    if (phrNorm_(v[ctx.idx.codigo]) === codigo && phrNorm_(v[ctx.idx.secciones]) === seccion) return row;
  }
  const last = phrUltimaFila_(ctx.sh, ctx.idx.codigo + 1, ctx.headerRow);
  const n = last - ctx.headerRow;
  if (n > 0) {
    const cod = ctx.sh.getRange(ctx.headerRow + 1, ctx.idx.codigo + 1, n, 1).getDisplayValues();
    const sec = ctx.sh.getRange(ctx.headerRow + 1, ctx.idx.secciones + 1, n, 1).getDisplayValues();
    for (let i = 0; i < n; i++) {
      if (phrNorm_(cod[i][0]) === codigo && phrNorm_(sec[i][0]) === seccion) return ctx.headerRow + 1 + i;
    }
  }
  throw new Error(`No se encontró ${it.codigo} sec. ${it.seccion} en ${ctx.sh.getName()}. Recarga la página.`);
}

/************************************************************
 * JORNADAS
 ************************************************************/

/** RUTs de profesores JORNADA desde la hoja PROFESORES (columna RUT PROFESOR 1). */
function phrJornadas_(ss) {
  try {
    const sh = ss.getSheetByName(PHR_CONFIG.JORNADAS_SHEET_NAME);
    if (!sh) return [];
    const v = sh.getDataRange().getDisplayValues();
    const objetivo = phrNormHeader_(PHR_CONFIG.JORNADAS_RUT_HEADER);
    let h = -1, colRut = -1;
    for (let r = 0; r < Math.min(25, v.length); r++) {
      colRut = v[r].findIndex(x => phrNormHeader_(x) === objetivo);
      if (colRut >= 0) { h = r; break; }
    }
    if (h < 0) return [];
    const hdr = v[h].map(phrNormHeader_);
    const colTipo = hdr.findIndex((x, i) => i !== colRut && (x.indexOf('JORNADA') >= 0 || x.indexOf('TIPO') >= 0 || x.indexOf('CONTRATO') >= 0));
    const out = new Set();
    for (let r = h + 1; r < v.length; r++) {
      const rut = phrRut_(v[r][colRut]);
      if (!rut) continue;
      if (colTipo >= 0 && phrNorm_(v[r][colTipo]).indexOf('JORNADA') < 0) continue;
      out.add(rut);
    }
    return Array.from(out);
  } catch (e) {
    return [];
  }
}

/************************************************************
 * REVISIÓN (aceptar avisos)
 ************************************************************/

function phrLeerRevisiones_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  Object.keys(all).forEach(k => {
    if (k.indexOf(PHR_CONFIG.REV_PREFIX) !== 0) return;
    try { out[k.slice(PHR_CONFIG.REV_PREFIX.length)] = JSON.parse(all[k]); } catch (e) { /* ignora */ }
  });
  return out;
}

/** clave = "CODIGO|SECCION"; datos = { firma, avisos } o null para quitar. */
function phrGuardarRevision(clave, datos) {
  clave = phrStr_(clave);
  if (!clave || clave.length > 80) throw new Error('Clave inválida.');
  const props = PropertiesService.getScriptProperties();
  const k = PHR_CONFIG.REV_PREFIX + clave;
  if (!datos) { props.deleteProperty(k); return null; }
  let usuario = '';
  try { usuario = Session.getActiveUser().getEmail() || ''; } catch (e) { usuario = ''; }
  const obj = {
    firma: phrStr_(datos.firma).slice(0, 40),
    avisos: (datos.avisos || []).slice(0, 30).map(x => phrStr_(x).slice(0, 300)),
    fecha: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    usuario
  };
  props.setProperty(k, JSON.stringify(obj));
  return obj;
}

/** Borra aceptaciones que ya no corresponden (el horario o los avisos cambiaron). */
function phrLimpiarRevisiones(claves) {
  const props = PropertiesService.getScriptProperties();
  (claves || []).forEach(c => props.deleteProperty(PHR_CONFIG.REV_PREFIX + phrStr_(c)));
  return true;
}

/************************************************************
 * CELDAS DE HORARIO
 ************************************************************/

/** "CLAS 10:30-13:20, AYUD 13:30-15:20" → { bloques: {CLAS:[2,3,4], AYUD:[5,6], LABT:[]}, resto } */
function phrParsearCeldaHorario_(text) {
  const bloques = { CLAS: [], AYUD: [], LABT: [] };
  const s = phrStr_(text);
  if (!s) return { bloques, resto: '' };
  let resto = s;
  const re = new RegExp(PHR_TOKEN_RE.source, 'gi');
  let m;
  while ((m = re.exec(s)) !== null) {
    const type = phrTipo_(m[1]);
    const ini = phrMin_(m[2]), fin = phrMin_(m[3]);
    if (!type || ini == null || fin == null || fin <= ini) continue;
    const dentro = [];
    PHR_CONFIG.BLOQUES.forEach((b, i) => {
      if (phrMin_(b[0]) >= ini && phrMin_(b[1]) <= fin) dentro.push(i);
    });
    if (!dentro.length) continue; // rango fuera de la grilla: se deja como texto
    dentro.forEach(i => { if (bloques[type].indexOf(i) < 0) bloques[type].push(i); });
    resto = resto.replace(m[0], ' ');
  }
  Object.keys(bloques).forEach(t => bloques[t].sort((a, b) => a - b));
  resto = resto.replace(/[\s,;\n]+/g, ' ').trim();
  return { bloques, resto };
}

/** Une bloques seguidos del mismo tipo y arma el texto ordenado por hora. */
function phrArmarCeldaHorario_(p) {
  const partes = [];
  Object.keys(p.bloques).forEach(type => {
    const bl = p.bloques[type];
    let ini = null, prev = null;
    bl.concat([null]).forEach(i => {
      if (i !== null && prev !== null && i === prev + 1) { prev = i; return; }
      if (ini !== null) partes.push({ start: ini, txt: PHR_TIPO_LABEL[type] + ' ' + PHR_CONFIG.BLOQUES[ini][0] + '-' + PHR_CONFIG.BLOQUES[prev][1] });
      ini = i; prev = i;
    });
  });
  partes.sort((a, b) => a.start - b.start);
  const txt = partes.map(x => x.txt);
  if (p.resto) txt.push(p.resto);
  return txt.join(', ');
}

/************************************************************
 * CONTEXTO / ENCABEZADOS
 ************************************************************/

function phrContexto_() {
  const id = phrStr_(PHR_CONFIG.SOURCE_SPREADSHEET_ID);
  let ss;
  try {
    ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    throw new Error('No se pudo abrir la planilla con ID ' + id + '. Revisa el ID y tus permisos. ' + e.message);
  }
  if (!ss) throw new Error('Configura PHR_CONFIG.SOURCE_SPREADSHEET_ID.');
  const sh = ss.getSheetByName(PHR_CONFIG.SOURCE_SHEET_NAME);
  if (!sh) throw new Error(`No existe la hoja "${PHR_CONFIG.SOURCE_SHEET_NAME}" en "${ss.getName()}".`);

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const lastCol = sh.getLastColumn();
  const top = sh.getRange(1, 1, Math.min(25, sh.getLastRow()), lastCol).getDisplayValues();
  let h = -1;
  for (let r = 0; r < top.length; r++) {
    if (top[r].some(v => phrNorm_(v) === 'CURSO MANDANTE')) { h = r; break; }
  }
  if (h < 0) throw new Error('No se encontró la fila de encabezados (CURSO MANDANTE).');
  const idx = phrIndices_(top[h]);
  return { ss, sh, tz, headerRow: h + 1, idx };
}

function phrIndices_(headers) {
  const f = (c, o) => phrBuscarHeader_(headers, c, o);
  const horario = {}, disp = {};
  PHR_DIAS.forEach(d => {
    horario[d.idx] = f(d.horario, { exactCase: true, fallback: 'last' });   // Lunes (horario) viene después de LUNES
    disp[d.idx] = f(d.disp, { exactCase: true, fallback: 'first' });       // LUNES (disponibilidad)
  });
  const idx = {
    cursoMandante: f(['CURSO MANDANTE']),
    codigo: f(['CODIGO', 'CÓDIGO']),
    titulo: f(['TITULO', 'TÍTULO']),
    secciones: f(['SECCIONES', 'SECCION', 'SECCIÓN']),
    area: f(['AREA', 'ÁREA']),
    lc: f(['LC', 'LISTA CRUZADA']),
    cupos: f(['CUPOS']),
    clasesProg: f(['Clases A PROGRAMAR', 'CLASES A PROGRAMAR']),
    ayudProg: f(['Ayudantías PROGRAMAR', 'Ayudantias PROGRAMAR']),
    labProg: f(['Laboratorios o Talleres PROGRAMAR']),
    // Respaldo cuando "A PROGRAMAR" está vacío (p. ej. fila nueva sin la fórmula)
    clasesBase: f(['Clases', 'CLASES']),
    ayudBase: f(['Ayudantías', 'Ayudantias', 'AYUDANTÍAS', 'AYUDANTIAS']),
    labBase: f(['Laboratorios o Talleres', 'LABORATORIOS O TALLERES']),
    distribucion: f(['2+1 o 3? (distribución horario de clases)', '2+1 o 3? (distribucion horario de clases)']),
    salaClas: f(['Sala especial CLAS']),
    salaAyud: f(['Sala especial AYUD']),
    salaLab: f(['Sala especial LAB']),
    rut1: f(['RUT PROFESOR 1']),
    nombre1: f(['NOMBRE PROFESOR BANNER 1', 'NOMBRE PROFESOR 1'], { prefix: true }),
    rut2: f(['RUT PROFESOR 2']),
    nombre2: f(['NOMBRE PROFESOR 2'], { prefix: true }),
    rutLabt: f(['RUT PROFESOR LABT']),
    nombreLabt: f(['PROFESOR LABT', 'NOMBRE PROFESOR LABT']),
    jornada: f(['JORNADA / HONORARIO', 'JORNADA/HONORARIO', 'JORNADA']),
    majors: PHR_MAJORS.map(m => ({ key: m.key, label: m.label, col: f(m.candidates) })),
    horario, disp
  };
  const faltan = [];
  [['cursoMandante', 'CURSO MANDANTE'], ['codigo', 'CODIGO'], ['titulo', 'TITULO'], ['secciones', 'SECCIONES'],
   ['clasesProg', 'Clases A PROGRAMAR'], ['ayudProg', 'Ayudantías PROGRAMAR'], ['labProg', 'Laboratorios o Talleres PROGRAMAR']]
    .forEach(([k, l]) => { if (idx[k] < 0) faltan.push(l); });
  PHR_DIAS.forEach(d => { if (horario[d.idx] < 0) faltan.push(d.horario[0]); });
  if (faltan.length) throw new Error('Faltan columnas en MAESTRO: ' + faltan.join(', '));
  PHR_DIAS.forEach(d => { if (disp[d.idx] === horario[d.idx]) disp[d.idx] = -1; });
  return idx;
}

function phrBuscarHeader_(headers, cands, opts) {
  opts = opts || {};
  for (const c of cands) {
    const i = headers.findIndex(h => phrStr_(h) === c);
    if (i >= 0) return i;
  }
  const nc = cands.map(phrNormHeader_);
  const hits = [];
  headers.forEach((h, i) => {
    const nh = phrNormHeader_(h);
    if (nh && nc.some(c => (opts.prefix ? nh.indexOf(c) === 0 : nh === c))) hits.push(i);
  });
  if (!hits.length) return -1;
  return opts.fallback === 'last' ? hits[hits.length - 1] : hits[0];
}

function phrColumnasUsadas_(idx) {
  const cols = [idx.cursoMandante, idx.codigo, idx.titulo, idx.secciones, idx.area, idx.lc, idx.cupos, idx.clasesProg,
    idx.ayudProg, idx.labProg, idx.distribucion, idx.salaClas, idx.salaAyud, idx.salaLab, idx.rut1, idx.nombre1,
    idx.rut2, idx.nombre2, idx.rutLabt, idx.nombreLabt, idx.jornada, idx.clasesBase, idx.ayudBase, idx.labBase];
  idx.majors.forEach(m => cols.push(m.col));
  Object.keys(idx.horario).forEach(k => cols.push(idx.horario[k], idx.disp[k]));
  return cols.filter(c => c >= 0);
}

function phrUltimaFila_(sh, col, headerRow) {
  const last = sh.getLastRow();
  if (last <= headerRow) return headerRow;
  const v = sh.getRange(headerRow + 1, col, last - headerRow, 1).getDisplayValues();
  for (let i = v.length - 1; i >= 0; i--) if (phrStr_(v[i][0])) return headerRow + 1 + i;
  return headerRow;
}

/************************************************************
 * HELPERS
 ************************************************************/

function phrIntervalosDisp_(t) {
  const out = [];
  phrStr_(t).split(/[,;\n]+/).forEach(p => {
    const x = phrStr_(p);
    if (!x || /^[A-Za-z]/.test(x)) return; // "CLAS 8:30-..." no es disponibilidad
    const m = x.match(/(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})/);
    if (!m) return;
    const a = phrMin_(m[1]), b = phrMin_(m[2]);
    if (a != null && b != null && b > a) out.push([a, b]);
  });
  out.sort((x, y) => x[0] - y[0]);
  const merged = [];
  out.forEach(iv => {
    const last = merged[merged.length - 1];
    if (last && iv[0] <= last[1] + PHR_CONFIG.MAX_RECREO_MIN) last[1] = Math.max(last[1], iv[1]);
    else merged.push([iv[0], iv[1]]);
  });
  return merged;
}

function phrResumenIntervalos_(ivs) {
  return ivs.map(iv => phrHora_(iv[0]) + '-' + phrHora_(iv[1])).join(', ');
}

function phrSemestres_(v) {
  const s = phrStr_(v);
  if (!s) return [];
  const out = [];
  s.split(/[\s,;\/\-]+|\bY\b|\bU\b/i).forEach(tok => {
    const m = phrStr_(tok).match(/^(\d{1,2})([A-Za-z]*)$/);
    if (!m || !parseInt(m[1], 10)) return;
    out.push({ num: parseInt(m[1], 10), suf: m[2].toUpperCase(), token: m[1] + m[2].toLowerCase() });
  });
  return out;
}

function phrTipo_(raw) {
  const t = phrNorm_(raw).replace(/\s+/g, '');
  if (t.indexOf('CLAS') === 0) return 'CLAS';
  if (t.indexOf('AYUD') === 0) return 'AYUD';
  if (t.indexOf('LAB') === 0 || t === 'TALLER') return 'LABT';
  return '';
}

function phrSala_(v) {
  const t = phrStr_(v);
  if (!t || /^(#N\/A|#REF!|#VALUE!|#DIV\/0!|#ERROR!|N\/A|-|—|NO|NINGUNA)$/i.test(t)) return '';
  return t;
}
function phrClaveSala_(v) { return phrNorm_(v).replace(/\s+/g, ''); }
function phrNum_(v) { const m = phrStr_(v).match(/\d+/); return m ? parseInt(m[0], 10) : 0; }
function phrStr_(v) { return v == null ? '' : String(v).replace(/ /g, ' ').trim(); }
function phrNorm_(v) { return phrStr_(v).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase(); }
function phrNormHeader_(v) { return phrNorm_(v).replace(/\s+/g, ' '); }
function phrRut_(v) { return phrStr_(v).toUpperCase().replace(/[.\-\s]/g, ''); }
function phrMin_(h) {
  const m = phrStr_(h).match(/^(\d{1,2}):(\d{2})$/);
  if (!m || +m[1] > 23 || +m[2] > 59) return null;
  return (+m[1]) * 60 + (+m[2]);
}
function phrHora_(min) { return Math.floor(min / 60) + ':' + ('0' + (min % 60)).slice(-2); }
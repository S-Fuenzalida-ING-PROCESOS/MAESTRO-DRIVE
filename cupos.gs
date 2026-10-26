
function calcularCupos202610() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_CUP = 'cupospreliminares';
  const SHEET_ESP = 'NRC POR PERIODO ESPEJO';

  const cup = ss.getSheetByName(SHEET_CUP);
  const esp = ss.getSheetByName(SHEET_ESP);
  if (!cup || !esp) throw new Error('Falta alguna hoja: "cupospreliminares" o "NRC POR PERIODO ESPEJO".');

  // --- cupospreliminares ---
  const cupData = cup.getDataRange().getValues();
  if (cupData.length < 2) return;
  const cupHdr = cupData[0];
  const idxCupCodigo = cupHdr.indexOf('CODIGO');
  if (idxCupCodigo === -1) throw new Error('En "cupospreliminares" falta la columna "CODIGO".');

  // Crear/ubicar columna de salida
  let idxOut = cupHdr.indexOf('Cupos 202510');
  if (idxOut === -1) {
    idxOut = cupHdr.length; // agregar al final
    cup.getRange(1, idxOut + 1).setValue('Cupos 202510');
  }

  // --- NRC POR PERIODO ESPEJO ---
  const espData = esp.getDataRange().getValues();
  if (espData.length < 2) {
    const outZeros = Array(cupData.length - 1).fill([0]);
    cup.getRange(2, idxOut + 1, outZeros.length, 1).setValues(outZeros);
    return;
  }
  const espHdr = espData[0];
  const idxEspCodigo = espHdr.indexOf('CODIGO');
  const idxEspNRC    = espHdr.indexOf('NRC');
  const idxEspCupos  = espHdr.indexOf('CUPOS');
  if ([idxEspCodigo, idxEspNRC, idxEspCupos].some(i => i === -1)) {
    throw new Error('En "NRC POR PERIODO ESPEJO" deben existir las columnas: CODIGO, NRC, CUPOS.');
  }

  // mapa[codigo][nrc] = cuposMax (NRC únicos por CODIGO)
  const mapa = Object.create(null);
  for (let i = 1; i < espData.length; i++) {
    const row = espData[i];
    const cod = String(row[idxEspCodigo] ?? '').trim();
    const nrc = String(row[idxEspNRC] ?? '').trim();
    if (!cod || !nrc) continue;

    let cuposNum = Number(row[idxEspCupos]);
    if (!isFinite(cuposNum)) cuposNum = 0;

    if (!mapa[cod]) mapa[cod] = Object.create(null);
    // si el NRC se repite para el CODIGO, conservar el mayor CUPOS
    if (!(nrc in mapa[cod]) || cuposNum > mapa[cod][nrc]) {
      mapa[cod][nrc] = cuposNum;
    }
  }

  // Suma por CODIGO (sumando una sola vez por NRC)
  const sumPorCodigo = Object.create(null);
  for (const cod of Object.keys(mapa)) {
    const byNrc = mapa[cod];
    let s = 0;
    for (const nrc of Object.keys(byNrc)) s += Number(byNrc[nrc]) || 0;
    sumPorCodigo[cod] = s;
  }

  // Escribir salida alineada a cupospreliminares
  const out = [];
  for (let i = 1; i < cupData.length; i++) {
    const cod = String(cupData[i][idxCupCodigo] ?? '').trim();
    const val = cod && (cod in sumPorCodigo) ? sumPorCodigo[cod] : 0;
    out.push([val]);
  }
  cup.getRange(2, idxOut + 1, out.length, 1).setValues(out);
}

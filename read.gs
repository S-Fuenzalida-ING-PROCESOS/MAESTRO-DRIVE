function obtenerNumeroDeColumna(sheet, nombreColumna, filaEncabezados) {
 
  // Verificar si la hoja existe
  if (!sheet) {
    throw new Error(`La hoja "${sheetName}" no existe.`);
  }

  // Obtener los valores de la fila donde están los encabezados
  const encabezados = sheet
    .getRange(filaEncabezados, 1, 1, sheet.getLastColumn()) // Obtiene una fila específica
    .getValues()[0]; // Accede a la primera fila del rango seleccionado
  
  // Buscar el índice del nombre de la columna
  const indice = encabezados.indexOf(nombreColumna);

  // Si no se encuentra el nombre, lanzar un error o devolver un valor específico
  if (indice === -1) {
    throw new Error(`La columna con el nombre "${nombreColumna}" no fue encontrada en la fila ${filaEncabezados}.`);
  }

  // Los índices empiezan desde 0, pero las columnas en Sheets empiezan desde 1
  return indice;
}
function leerInscritosPorPeriodo() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('RAMOS INSCRITOS POR PERIODO');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Desde fila 4 hasta la última fila con datos
  const data = sheet.getRange(6, 1, lastRow - 5, lastCol).getValues();
  
  return data;
}
function verifySheets(hojas_sheets,hoja_existente){ //verifica si la hoja que se va a crear existe, retorna un booleano
  const  hojas = hojas_sheets.getSheets(); // Obtiene todas las hojas del documento
  const  nombresHojas=hojas.map((sheet)=>sheet.getName()) ; // Inicializa un array para almacenar los nombres de las hojas
  return  !nombresHojas.includes(hoja_existente)
}
function obtenerSpreadsheetDeCarpeta(idCarpeta, nombreArchivo) {
  try {
    // Obtener la carpeta por su ID
    const carpeta = DriveApp.getFolderById(idCarpeta);
    
    // Buscar archivos con el nombre especificado
    const archivos = carpeta.getFilesByName(nombreArchivo);
    
    if (archivos.hasNext()) {
      const archivo = archivos.next();
      
      // Verificar que el archivo es un Google Spreadsheet
      if (archivo.getMimeType() === MimeType.GOOGLE_SHEETS) {
        // Retornar el Spreadsheet usando su ID
        return SpreadsheetApp.openById(archivo.getId());
      } else {
        Logger.log("El archivo encontrado no es un Google Spreadsheet.");
        return null;
      }
    } else {
      Logger.log("No se encontró el archivo con el nombre: " + nombreArchivo);
      return null;
    }
  } catch (e) {
    Logger.log("Error: " + e.message);
    return null;
  }
}
function obtenerDatosDesdeHoja(nombreHoja, filaInicio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreHoja);

  if (!sheet) {
    throw new Error(`La hoja "${nombreHoja}" no existe.`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < filaInicio) {
    return [];
  }

  return sheet
    .getRange(
      filaInicio,
      1,
      lastRow - filaInicio + 1,
      lastCol
    )
    .getValues();
}
function leerHojaSinEncabezado(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreHoja);
  
  if (!sheet) {
    throw new Error(`No existe la hoja: ${nombreHoja}`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return []; // solo encabezado o vacía

  // Desde fila 2 (saltamos encabezado)
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}
function normalizarEncabezado(v) {
  return removeAccents(String(v || ""))
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}
function buscarColumnaFlexible(encabezados, nombres) {
  const arr = Array.isArray(nombres) ? nombres : [nombres];

  for (const name of arr) {
    const target = normalizarEncabezado(name);

    for (let i = 0; i < encabezados.length; i++) {
      if (normalizarEncabezado(encabezados[i]) === target) return i;
    }
  }

  return -1;
}
function obtenerIndiceObligatorio(encabezados, nombres) {
  const idx = buscarColumnaFlexible(encabezados, nombres);

  if (idx === -1) {
    throw new Error(`Columna no encontrada: "${(Array.isArray(nombres) ? nombres : [nombres]).join(" / ")}"`);
  }

  return idx;
}

function obtenerDirectorioDocentes_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directorio = {}; // rut (normalizado) -> {nombre, rut, email}

  // 1. HISTORICO PROFESORES: puede tener varias filas por RUT (una por periodo).
  //    Nos quedamos con la fila de mayor PERIODO para cada RUT.
  const hojaHistorico = ss.getSheetByName("HISTORICO PROFESORES");
  if (hojaHistorico) {
    const datos = hojaHistorico.getDataRange().getValues();
    const encabezados = datos.shift();
    const colPeriodo = buscarColumnaFlexible(encabezados, "PERIODO");
    const colRut = buscarColumnaFlexible(encabezados, "RUT PROFESOR 1");
    const colNombre = buscarColumnaFlexible(encabezados, "NOMBRE PROFESOR 1");
    const colEmail = buscarColumnaFlexible(encabezados, "EMAIL PROFESOR 1");

    if (colRut !== -1 && colNombre !== -1) {
      const mejorPeriodoPorRut = {};
      datos.forEach((fila) => {
        const rut = String(fila[colRut] || "").trim();
        if (!rut) return;
        const nombre = String(fila[colNombre] || "").trim();
        if (!nombre) return;
        const periodo = Number(fila[colPeriodo]) || 0;
        if (mejorPeriodoPorRut[rut] === undefined || periodo > mejorPeriodoPorRut[rut]) {
          mejorPeriodoPorRut[rut] = periodo;
          directorio[rut] = {
            nombre: nombre,
            rut: rut,
            email: colEmail !== -1 ? String(fila[colEmail] || "").trim() : ""
          };
        }
      });
    }
  }

  // 2. PROFESORES: roster del periodo actual. Si el RUT ya existía (por HISTORICO),
  //    los datos de PROFESORES tienen prioridad y lo sobrescriben.
  const hojaProfesores = ss.getSheetByName("PROFESORES");
  if (hojaProfesores) {
    const datos = hojaProfesores.getDataRange().getValues();
    const encabezados = datos.shift();
    const colRut = buscarColumnaFlexible(encabezados, "RUT PROFESOR 1");
    const colNombre = buscarColumnaFlexible(encabezados, [
      "NOMBRE PROFESOR BANNER 1 \n(PROFESOR PRINCIPAL SESIÓN 01)",
      "NOMBRE PROFESOR 1"
    ]);
    const colEmail = buscarColumnaFlexible(encabezados, "EMAIL PROFESOR 1");

    if (colRut !== -1 && colNombre !== -1) {
      datos.forEach((fila) => {
        const rut = String(fila[colRut] || "").trim();
        if (!rut) return;
        const nombre = String(fila[colNombre] || "").trim();
        if (!nombre) return;
        directorio[rut] = {
          nombre: nombre,
          rut: rut,
          email: colEmail !== -1 ? String(fila[colEmail] || "").trim() : ""
        };
      });
    }
  }

  return Object.values(directorio);
}

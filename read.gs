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
  // Los datos parten justo después de la fila de encabezados (MATERIA y CURSO),
  // esté donde esté. Antes se leía desde la fila 6 fija y se perdían filas
  // cuando el reporte se pegaba con el encabezado en la fila 1.
  const filaEncabezado = buscarFilaEncabezado_(sheet, ["MATERIA", "CURSO"]);
  if (filaEncabezado === -1) {
    throw new Error('En "RAMOS INSCRITOS POR PERIODO" no se encontró la fila de encabezados (MATERIA, CURSO).');
  }
  // { encabezados, filas }
  return leerTablaConEncabezado_(sheet, filaEncabezado);
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

  // 1) Coincidencia exacta (solo espacios en los extremos). Así "Lunes"
  //    (horario) no se confunde con "LUNES" (disponibilidad) cuando ambas
  //    existen en la misma hoja.
  for (const name of arr) {
    const exacto = String(name == null ? "" : name).trim();
    for (let i = 0; i < encabezados.length; i++) {
      if (String(encabezados[i] == null ? "" : encabezados[i]).trim() === exacto) return i;
    }
  }

  // 2) Coincidencia normalizada (sin tildes, mayúsculas, espacios simples).
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

// ===================== LECTURA DE COLUMNAS POR NOMBRE =====================
// Todas las hojas se leen por el nombre de sus columnas, nunca por su
// posición: así una columna de ayuda agregada (y olvidada) en cualquier hoja
// no altera el resultado de las macros.

// Devuelve un objeto { clave: índice } a partir de { clave: nombre o [nombres] }.
// Lanza un error que lista todas las columnas que faltan.
function columnasPorNombre_(encabezados, definicion, nombreHoja) {
  const indices = {};
  const faltan = [];
  Object.keys(definicion).forEach((clave) => {
    const nombres = definicion[clave];
    const idx = buscarColumnaFlexible(encabezados, nombres);
    if (idx === -1) faltan.push([].concat(nombres)[0]);
    indices[clave] = idx;
  });
  if (faltan.length) {
    throw new Error(`En la hoja "${nombreHoja}" faltan las columnas: ${faltan.join(", ")}.`);
  }
  return indices;
}

// Lee una hoja completa y separa la fila de encabezados del resto.
// filaEncabezado es 1-based (por defecto 1). Con usarDisplay se leen los
// valores tal como se ven en la hoja.
function leerTablaConEncabezado_(hoja, filaEncabezado, usarDisplay) {
  const fila = filaEncabezado || 1;
  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  if (lastRow < fila || lastCol < 1) return { encabezados: [], filas: [] };
  const rango = hoja.getRange(fila, 1, lastRow - fila + 1, lastCol);
  const valores = usarDisplay ? rango.getDisplayValues() : rango.getValues();
  return { encabezados: valores[0], filas: valores.slice(1) };
}

// Busca, en las primeras filas, la fila que contiene todos los encabezados
// indicados. Devuelve el número de fila (1-based) o -1.
function buscarFilaEncabezado_(hoja, nombres, maxFilas) {
  const lastRow = Math.min(hoja.getLastRow(), maxFilas || 15);
  const lastCol = hoja.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return -1;
  const valores = hoja.getRange(1, 1, lastRow, lastCol).getValues();
  for (let r = 0; r < valores.length; r++) {
    if (nombres.every((n) => buscarColumnaFlexible(valores[r], [n]) !== -1)) return r + 1;
  }
  return -1;
}

// Valores de una columna (sin el encabezado), buscada por nombre.
function leerColumnaPorNombre_(hoja, nombres, usarDisplay) {
  const tabla = leerTablaConEncabezado_(hoja, 1, usarDisplay);
  const idx = buscarColumnaFlexible(tabla.encabezados, nombres);
  if (idx === -1) {
    throw new Error(`En la hoja "${hoja.getName()}" falta la columna: ${[].concat(nombres)[0]}.`);
  }
  return tabla.filas.map((fila) => fila[idx]);
}

// RUT de los profesores de la hoja PROFESORES (columna "RUT PROFESOR 1").
function rutsHojaProfesores_(hojaProfesores, usarDisplay) {
  return leerColumnaPorNombre_(hojaProfesores, ["RUT PROFESOR 1", "RUT"], usarDisplay);
}

// Agrega una fila a una hoja de registro ubicando cada valor bajo su
// encabezado. Si falta un encabezado se agrega al final de la fila 1.
function agregarFilaPorEncabezado_(hoja, valoresPorNombre) {
  const lastCol = Math.max(hoja.getLastColumn(), 1);
  const encabezados = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
  const fila = new Array(encabezados.length).fill("");
  Object.keys(valoresPorNombre).forEach((nombre) => {
    let idx = buscarColumnaFlexible(encabezados, [nombre]);
    if (idx === -1) {
      idx = encabezados.length;
      encabezados.push(nombre);
      hoja.getRange(1, idx + 1).setValue(nombre);
      fila.push("");
    }
    fila[idx] = valoresPorNombre[nombre];
  });
  hoja.appendRow(fila);
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

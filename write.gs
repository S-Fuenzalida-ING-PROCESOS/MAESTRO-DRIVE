function recrearHoja(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaExistente = ss.getSheetByName(nombreHoja);

  if (hojaExistente) {
    ss.deleteSheet(hojaExistente);
  }

  return ss.insertSheet(nombreHoja);
}

// ===================== REGISTRO DE RECORDATORIOS =====================
// Lleva constancia de a quién se le ha enviado un recordatorio (o un reenvío
// de formulario) y cuántas veces, para poder auditar envíos y para que
// enviarRecordatoriosNoRespondidos/reenviarFormularioPorRut puedan seguir
// aunque un envío individual falle (correo inválido, tipo de correo faltante
// en TEXTOS CORREO, etc.) sin perder el registro de lo que sí funcionó.
const HOJA_LOG_RECORDATORIOS = "LOG RECORDATORIOS";
const HOJA_RESUMEN_RECORDATORIOS = "RESUMEN RECORDATORIOS";

function obtenerOcrearHojaLogRecordatorios_(ss) {
  let hoja = ss.getSheetByName(HOJA_LOG_RECORDATORIOS);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_LOG_RECORDATORIOS);
    hoja.getRange(1, 1, 1, 7).setValues([["FECHA", "RUT", "NOMBRE", "EMAIL", "FORMULARIO", "RESULTADO", "DETALLE"]]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function obtenerOcrearHojaResumenRecordatorios_(ss) {
  let hoja = ss.getSheetByName(HOJA_RESUMEN_RECORDATORIOS);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_RESUMEN_RECORDATORIOS);
    hoja.getRange(1, 1, 1, 6).setValues([["RUT", "NOMBRE", "EMAIL", "RECORDATORIOS ENVIADOS", "RECORDATORIOS FALLIDOS", "ULTIMO INTENTO"]]);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// Deja una línea en el log (append-only, un registro por intento) y actualiza
// el resumen por RUT. Si para este RUT no existía aún ningún registro (porque
// estas hojas no existían antes de este parche), este intento cuenta como el
// primer recordatorio: no hay forma de reconstruir el historial previo, y no
// hace falta, ya que el requisito es contar hacia adelante desde ahora.
function registrarRecordatorio(rut, nombre, email, formulario, resultado, detalle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLog = obtenerOcrearHojaLogRecordatorios_(ss);
  hojaLog.appendRow([new Date(), rut, nombre || "", email || "", formulario || "", resultado, detalle || ""]);
  actualizarResumenRecordatorio(rut, nombre, email, resultado === "ENVIADO");
}

function actualizarResumenRecordatorio(rut, nombre, email, exito) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = obtenerOcrearHojaResumenRecordatorios_(ss);
  const lastRow = hoja.getLastRow();
  const ahora = new Date();

  if (lastRow >= 2) {
    const data = hoja.getRange(2, 1, lastRow - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(rut).trim()) {
        const filaSheet = i + 2;
        const enviados = Number(data[i][3]) || 0;
        const fallidos = Number(data[i][4]) || 0;
        hoja.getRange(filaSheet, 2, 1, 5).setValues([[
          nombre || data[i][1] || "",
          email || data[i][2] || "",
          exito ? enviados + 1 : enviados,
          exito ? fallidos : fallidos + 1,
          ahora
        ]]);
        return;
      }
    }
  }

  // Primer registro para este RUT.
  hoja.appendRow([rut, nombre || "", email || "", exito ? 1 : 0, exito ? 0 : 1, ahora]);
}

function contarRecordatoriosEnviados(rut) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_RESUMEN_RECORDATORIOS);
  if (!hoja) return 0;
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return 0;
  const data = hoja.getRange(2, 1, lastRow - 1, 4).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(rut).trim()) {
      return Number(data[i][3]) || 0;
    }
  }
  return 0;
}

// ===================== RESPALDO AUTOMÁTICO DE MAESTRO =====================
// Antes de sobrescribir MAESTRO, se guarda una copia en una hoja oculta.
// El respaldo nunca debe impedir el flujo principal: si falla, solo se deja
// constancia en el log de ejecución y se continúa.
function respaldarMaestroAntesDeEscribir(hojaMaestro) {
  try {
    const ss = hojaMaestro.getParent();
    const nombreRespaldo = "MAESTRO_RESPALDO_AUTOMATICO";
    let hojaRespaldo = ss.getSheetByName(nombreRespaldo);
    if (!hojaRespaldo) {
      hojaRespaldo = ss.insertSheet(nombreRespaldo);
      hojaRespaldo.hideSheet();
    }
    hojaRespaldo.clearContents();
    hojaRespaldo.getRange(1, 1).setValue(`Respaldo automático antes de escribir: ${new Date()}`);
    const datos = hojaMaestro.getDataRange().getValues();
    if (datos.length > 0 && datos[0].length > 0) {
      hojaRespaldo.getRange(2, 1, datos.length, datos[0].length).setValues(datos);
    }
  } catch (e) {
    Logger.log("No se pudo respaldar MAESTRO antes de escribir: " + e.message);
  }
}

function escribirDatosYResaltar(hoja, datos) {


  if (!hoja) {
    throw new Error(`No se encontró la hoja con el nombre: ${sheet}`);
  }

  if (!datos || datos.length === 0) {
    throw new Error(`No hay datos para escribir en "${hoja.getName()}". La operación se canceló para no vaciar la hoja.`);
  }

  // Determinar la cantidad de columnas de los datos
  const numColumnasDatos = datos[0].length;

  // FIX (pérdida de datos en MAESTRO): antes se borraba el contenido de la
  // hoja y RECIÉN DESPUÉS se intentaba escribir los datos nuevos. Si "datos"
  // traía filas de distinto ancho (p.ej. por una fusión hecha por posición en
  // vez de por nombre de columna), setValues() lanzaba una excepción con la
  // hoja ya vacía, perdiendo la información. Ahora se valida ANTES de borrar
  // nada: todas las filas deben tener el mismo ancho, y ese ancho debe
  // coincidir con el de la hoja destino.
  const anchoHoja = hoja.getLastColumn();
  for (let i = 0; i < datos.length; i++) {
    if (!Array.isArray(datos[i]) || datos[i].length !== numColumnasDatos) {
      throw new Error(
        `Los datos a escribir en "${hoja.getName()}" no son consistentes: la fila ${i + 1} tiene ${Array.isArray(datos[i]) ? datos[i].length : "?"} columna(s), pero se esperaban ${numColumnasDatos}. No se modificó la hoja.`
      );
    }
  }
  if (numColumnasDatos !== anchoHoja) {
    throw new Error(
      `Los datos a escribir tienen ${numColumnasDatos} columna(s), pero la hoja "${hoja.getName()}" tiene ${anchoHoja}. No se modificó la hoja para evitar perder información.`
    );
  }

  // Respaldo preventivo antes de tocar MAESTRO específicamente.
  if (hoja.getName() === "MAESTRO") {
    respaldarMaestroAntesDeEscribir(hoja);
  }

  // Limpiar contenido y formato solo debajo de los encabezados
  const numFilas = hoja.getLastRow();
  if (numFilas > 1) {
    hoja.getRange(2, 1, numFilas - 1, hoja.getLastColumn()).clearContent().clearFormat();
  }

  // Escribe los datos debajo de los encabezados
  hoja.getRange(2, 1, datos.length, numColumnasDatos).setValues(datos);

  // Encuentra la columna "NRC" de los encabezados
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const columnaNRC = encabezados.indexOf("NRC") + 1;

  if (columnaNRC === 0) {
    throw new Error('No se encontró una columna con el encabezado "NRC".');
  }

  // Recorre los datos y resalta las filas según la condición
  datos.forEach((fila, i) => {
    if (fila[columnaNRC - 1] === "sin NRC asociado") {
      const rangoFila = hoja.getRange(i + 2, 1, 1, numColumnasDatos); // i+2 porque los datos empiezan en la fila 2
      rangoFila.setBackground("yellow");
    }
  });
}
function crear_hoja_nombre(nombre,hojasActuales) {
  
  const hoja = hojasActuales.getSheetByName(nombre)
  if (hoja){
    return hoja
  } 
  if(verifySheets(hojasActuales,nombre)){const nueva_hoja = createSheetByName(hojasActuales,nombre);}
  else {
  deleteSheetByName(hojasActuales,nombre);
  const nueva_hoja = createSheetByName(hojasActuales,nombre);
  }
  
  return hojasActuales.getSheetByName(nombre)

}
function deleteSheetByName(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
    Logger.log("Hoja '" + sheetName + "' eliminada correctamente.");
  } else {
    Logger.log("No se encontró ninguna hoja con el nombre '" + sheetName + "'.");
  }
}
function createSheetByName(spreadsheet, sheetName) {
  const newSheet = spreadsheet.insertSheet(sheetName);
  Logger.log("Nueva hoja '" + sheetName + "' creada correctamente.");
  return newSheet;
}
function escribir_hoja_con_data(informaciones,sheet,encabezado){

informaciones = [encabezado,...informaciones]
sheet.getRange(1, 1, informaciones.length, informaciones[0].length).setValues(informaciones);


}
function escribirEncabezado(hoja, encabezado) {
  

  if (!hoja) {
    throw new Error(`No se encontró la hoja con el nombre: ${sheetName}`);
  }

  // Limpia la primera fila antes de escribir el encabezado
  const numColumnas = hoja.getLastColumn();
  if (numColumnas > 0) {
    hoja.getRange(1, 1, 1, numColumnas).clearContent();
  }

  // Escribe el encabezado en la primera fila
  hoja.getRange(1, 1, 1, encabezado.length).setValues([encabezado]);

  Logger.log("Encabezado escrito correctamente.");
}
function copiarEncabezado(hojaOrigen,hojaDestino) {
  

  if (!hojaOrigen || !hojaDestino) {
    throw new Error(`No se encontraron las hojas especificadas: "${sheetOrigen}" o "${sheetDestino}".`);
  }

  // Obtener la primera fila (encabezado) de la hoja de origen
  const rangoEncabezadoOrigen = hojaOrigen.getRange(1, 1, 1, hojaOrigen.getLastColumn());
  const valoresEncabezado = rangoEncabezadoOrigen.getValues();
  const formatosEncabezado = rangoEncabezadoOrigen.getFontWeights();
  const coloresEncabezado = rangoEncabezadoOrigen.getFontColors();
  const alineacionesEncabezado = rangoEncabezadoOrigen.getHorizontalAlignments();
  const tamaniosEncabezado = rangoEncabezadoOrigen.getFontSizes();

  // Limpiar la primera fila de la hoja de destino
  const numColumnasDestino = hojaDestino.getLastColumn();
  if (numColumnasDestino > 0) {
    hojaDestino.getRange(1, 1, 1, numColumnasDestino).clearContent().clearFormat();
  }

  // Escribir encabezado en la hoja de destino
  const rangoEncabezadoDestino = hojaDestino.getRange(1, 1, 1, valoresEncabezado[0].length);
  rangoEncabezadoDestino.setValues(valoresEncabezado);

  // Aplicar el formato del encabezado original
  rangoEncabezadoDestino.setFontWeights(formatosEncabezado);
  rangoEncabezadoDestino.setFontColors(coloresEncabezado);
  rangoEncabezadoDestino.setHorizontalAlignments(alineacionesEncabezado);
  rangoEncabezadoDestino.setFontSizes(tamaniosEncabezado);
  const coloresFondo = rangoEncabezadoOrigen.getBackgrounds();
  rangoEncabezadoDestino.setBackgrounds(coloresFondo);
  //const numColumnas = hojaOrigen.getLastColumn();
  //for (let col = 1; col <= numColumnas; col++) {
    //const anchoColumna = hojaOrigen.getColumnWidth(col);
    //hojaDestino.setColumnWidth(col, anchoColumna);
  //}
  hojaDestino.setFrozenRows(1)
  
}
function crearSpreadsheetEnCarpeta(nombreArchivo) {
  

  // Buscar la carpeta por su nombre
  const carpeta = DriveApp.getFolderById(id)(idCarpeta);
  

  

  // Crear el nuevo spreadsheet
  const nuevoSpreadsheet = SpreadsheetApp.create(nombreArchivo);

  // Mover el spreadsheet a la carpeta
  const archivo = DriveApp.getFileById(nuevoSpreadsheet.getId());
  carpeta.addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo); // Elimina el archivo de "Mi Unidad"

  
}
function crearSpreadsheetEnCarpetaConRemplazo(nombreArchivo) {
  try {
    const folderId = id_carpeta_archivos_cordinadores;

    // Buscar si existe archivo con ese nombre en la carpeta
    const query = `'${folderId}' in parents and name = '${nombreArchivo}' and trashed = false`;
    const res = Drive.Files.list({ q: query, fields: "files(id, name)" });

    if (res.files && res.files.length > 0) {
      return SpreadsheetApp.openById(res.files[0].id);
    }

    // Crear Spreadsheet
    const ss = SpreadsheetApp.create(nombreArchivo);

    // Mover usando Drive API (forma moderna)
    Drive.Files.update(
      {},
      ss.getId(),
      null,
      {
        addParents: folderId,
        removeParents: "root"
      }
    );

    return ss;

  } catch (e) {
    Logger.log(e);
    throw new Error("No se pudo crear o mover el archivo Google Sheets.");
  }
}

function enviarLinksConPermisos(idCarpeta, listaDatos) {
  try {
    // Obtener la carpeta por su ID
    const carpeta = DriveApp.getFolderById(idCarpeta);

    // Cargar el logo (usa caché si ya se cargó antes en esta ejecución)
    const logoBlob = obtenerLogoFirma_();

    // Iterar por cada elemento en la lista de datos
    listaDatos.forEach(item => {
      const email = item.mail;
      const nombresArchivos = item.archivos;
      let mensaje = `Estimado/a Coordinador/a ${item.name},\n\nJunto con saludar, escribimos para informar que nos encontramos iniciando el proceso de Programación Académica para el primer semestre del 2027. Para ello, le adjuntamos el enlace al archivo que contiene los cursos programados para el semestre correspondientes a su coordinación. Por favor, tener en cuenta lo siguiente:\n
      -Etapa 1: Revisar los cursos, sus cupos, la cantidad de secciones a dictar y las salas especiales requeridas, validando la información detallada y realizando los cambios que sean necesarios. Debe considerar la información correspondiente a los cursos cuyo valor es "SI" en la columna CURSO MANDANTE (columna A), puesto que las columnas con valor NO corresponden a una Lista Cruzada.
      -Etapa 2: Ingresar la información de los profesores que dictarán cada una de las secciones, considerando profesor principal, secundario y/o laboratorio (Ingresar Nombre Completo, RUT y Correo). Información entre las columnas AL:AV de la planilla, según corresponda.\n
      El plazo para entregar esta información es el viernes 2 de octubre.\n
      A su vez, le compartimos la carpeta "INFORMACION DE UTILIDAD" (https://drive.google.com/drive/folders/1zPo4y1OMmjXXIN8RgucXPzeiu1s6Hr8O) la cual contiene distintos archivos que podrían complementar las revisiones y/o asignaciones, en caso de que lo requiera o necesite.\n Si desea crear un nuevo electivo, por favor, descargar una copia del archivo "Catálogo Creación de Curso" disponible en la carpeta “CREACION DE CURSOS” ubicada en "INFORMACION DE UTILIDAD". El archivo con la información del electivo nuevo debe ser subido a la carpeta "CREACION DE CURSOS” indicando, en el nombre del archivo, el título del electivo a crear.\n\nQuedamos atentos a cualquier duda que tenga y de antemano, agradecemos su ayuda.
      \n`;

      nombresArchivos.forEach(nombreArchivo => {
        // Buscar el archivo por nombre dentro de la carpeta
        const archivos = carpeta.getFilesByName(nombreArchivo);
        if (archivos.hasNext()) {
          const archivo = archivos.next();

          // Dar permisos de edición sin el correo de notificación automático de Drive
          Drive.Permissions.create(
            {
              role: "writer",
              type: "user",
              emailAddress: email
            },
            archivo.getId(),
            {
              sendNotificationEmail: false
            }
          );

          // Agregar el link al mensaje
          const link = archivo.getUrl();
          mensaje += `- ${nombreArchivo}: ${link}\n`;

          Logger.log(`Permisos otorgados y link generado para "${nombreArchivo}" a ${email}.`);
        } else {
          Logger.log(`Archivo "${nombreArchivo}" no encontrado en la carpeta.`);
          mensaje += `- ${nombreArchivo}: No encontrado en la carpeta\n`;
        }
      });

      // Enviar el correo con los enlaces, con firma UANDES
      const asunto = "Inicio de Programación Académica - 202710";
      mensaje += `\nSaludos cordiales,`;

      const plainBody = mensaje + '\n\n' + buildPlainSignature_();
      const htmlBody = buildHtmlBodyWithSignature_(mensaje);

      GmailApp.sendEmail(email, asunto, plainBody, {
        htmlBody: htmlBody,
        inlineImages: { logo: logoBlob }
      });
    });
  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    throw new Error("No se pudieron compartir los archivos.");
  }
}
function resaltarFilasCambiadas(hoja, desdeFila, listaBooleana) {
  
  if (!hoja) {
    throw new Error(`No se encontró la hoja "${sheetName}"`);
  }

  // Determina la última fila que debe revisarse según la longitud de la lista booleana
  const hastaFila = desdeFila + listaBooleana.length - 1;

  // Aplica el formato a cada fila basada en la lista booleana
  listaBooleana.forEach((encontrado, index) => {
    const filaActual = desdeFila + index;

    if (encontrado) {
      
    } else {
      hoja.getRange(filaActual, 1, 1, hoja.getLastColumn())
        .setBackground("#16d696"); // Limpia el formato si no está cambiada
    }
  });
}
function escribirHorariosEnHoja(sheet, filaIndice, columnasDias, horario,fila_actual) {
  try {
    // Iterar sobre los días en el diccionario y las columnas especificadas
    ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"].forEach((dia, i) => {

      let valor = horario[dia] || "";;
      const columna = columnasDias[i]; // Obtener la columna correspondiente al día
      if(fila_actual[columna]==""){
         sheet.getRange(filaIndice, columna).setValue(valor); // Escribir en la celda
        
        
      }
      else{

        valor = interseccion_de_bloques(valor,fila_actual)
      }
       // Tomar el valor del diccionario o dejar vacío
     

    });
    return { success: true, message: "Horarios escritos correctamente." };
  } catch (error) {
    return { success: false, message: "Error al escribir en la hoja: " + error.message };
  }
}
function limpiarColumnas(hoja,columnas) {
   // Índices de las columnas a limpiar (1 = Columna A, 3 = Columna C, etc.)
  
  const rango = hoja.getDataRange(); // Obtén todo el rango de datos
  const datos = rango.getValues(); // Obtén los valores como una matriz bidimensional
  
  // Recorre las filas y limpia las columnas especificadas
  for (let i = 1; i < datos.length; i++) {
    columnas.forEach(col => {
      if (col >= 0 && col < datos[i].length) { 
        datos[i][col] = ""; // Limpia el contenido de la celda
      }
    });
  }
  
  // Escribe los datos actualizados en el rango
  rango.setValues(datos);
}
function escribirListaDeListasDesdeFila(datos, nombreHoja, filaInicio) {
  if (!datos || datos.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) {
    throw new Error(`No existe la hoja: ${nombreHoja}`);
  }

  const numFilas = datos.length;
  const numColumnas = Math.max(...datos.map(fila => fila.length));

  // Normalizamos filas más cortas (Sheets requiere matriz rectangular)
  const matriz = datos.map(fila => {
    const copia = fila.slice();
    while (copia.length < numColumnas) copia.push("");
    return copia;
  });

  sheet
    .getRange(filaInicio, 1, numFilas, numColumnas)
    .setValues(matriz);
}
function escribirListaDeListas(hoja, datos, filaInicial, columnaInicial) {
  // Determinar el tamaño del rango (filas y columnas)
  const numFilas = datos.length;
  const numColumnas = datos[0].length;
  
  
  const rango = hoja.getRange(filaInicial, columnaInicial, numFilas, numColumnas);
  rango.setValues(datos); // Escribe los datos
}
function escribirFechasEnFila(listaFechas, listaDiasHabiles, fila, columna) {
    // Abre la hoja activa
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Determina la última columna donde se escribirá y la cantidad de columnas necesarias
    const ultimaColumna = columna + listaFechas.length - 1;

    // Limpia las celdas que están después de la última fecha escrita
    const rangoLimpiar = hoja.getRange(
    fila, // Fila inicial
    columna, // Columna inicial
    hoja.getMaxRows() - fila + 1, // Total de filas restantes (desde `fila` hasta el final)
    hoja.getMaxColumns() - columna + 1 // Total de columnas restantes (desde `columna` hasta el final)
);
    rangoLimpiar.clear({ contentsOnly: true, validationsOnly: true }); // Elimina contenido y listas desplegables

    // Obtiene el rango de toda la fila donde se escribirán las fechas
    const rangoFormato = hoja.getRange(fila, columna, 1, listaFechas.length);

    // Aplica formato a toda la fila
    rangoFormato
        .setNumberFormat("ddd dd mmm") // Formato de fecha
        .setFontWeight("bold")         // Texto en negrita
        .setTextRotation(90);          // Rotación del texto a 90 grados

    // Escribe las fechas y aplica formato condicional a cada celda
    listaFechas.forEach((fecha, index) => {
        const celda = hoja.getRange(fila, columna + index); // Determina la celda horizontalmente

        // Escribe la fecha
        celda.setValue(fecha);

        // Si no es día hábil (fin de semana), aplica formato condicional
        if (!listaDiasHabiles[index]) {
            celda.setBackground("#FF7F7F"); // Fondo rojo suave
            celda.setFontColor("#FFFFFF"); // Texto blanco
        } else {
            celda.setBackground(null); // Fondo normal
            celda.setFontColor(null); // Texto normal
        }
    });
}

function resaltarCambios(cambios, nuevaHoja, rangoInicio) {
  cambios.forEach((cambio) => {
    const { tipoCambio, filaAnterior, filaNueva, indice } = cambio;

    // Calcular la fila de destino asegurándonos de que sea al menos 1
    const filaDestino = Math.max(rangoInicio.fila + indice, 1); // Asegura que la fila sea >= 1
    const columnaDestino = rangoInicio.columna;  // Columna de inicio

    // Obtener el rango a partir de la fila y columna de destino
    const rango = nuevaHoja.getRange(filaDestino, columnaDestino, 1, filaNueva.length);

    // Resaltar según el tipo de cambio
    if (tipoCambio === "insertado") {
      rango.setBackground("#DFF0D8"); // Color verde claro para inserciones
      nuevaHoja.getRange(filaDestino, columnaDestino, 1, filaNueva.length).setValues([filaNueva]);
    } else if (tipoCambio === "eliminado") {
      // Escribir la fila eliminada en la nueva hoja
      rango.setBackground("#F2D7D5"); // Color rojo claro para eliminaciones
      nuevaHoja.getRange(filaDestino, columnaDestino, 1, filaAnterior.length).setValues([filaAnterior]); // Escribir la fila eliminada
    } else if (tipoCambio === "modificado") {
      rango.setBackground("#FCF8E3"); // Color amarillo claro para cambios
      nuevaHoja.getRange(filaDestino, columnaDestino, 1, filaNueva.length).setValues([filaNueva]);
    }
  });
}
function crearListaDesplegableConFormato(hoja, fila, inicioColumna, finColumna,opciones) {
  // Ajustar índices de columnas (de 0-indexado a 1-indexado de Google Sheets)
  const inicioColumnaIndex = inicioColumna + 1;
  const finColumnaIndex = finColumna + 1;

  // Define el rango donde quieres aplicar la lista desplegable
  const rango = hoja.getRange(fila + 1, inicioColumnaIndex, 1, finColumnaIndex - inicioColumnaIndex + 1);

  // Define las opciones para la lista desplegable
   // Cambia según tus opciones

  // Configura la validación de datos
  const validacion = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones)
    .setAllowInvalid(true) // Permite valores no válidos
    .build();

  // Aplica la validación al rango
  rango.setDataValidation(validacion);

  // Configura el formato condicional
  
}
function agregarAlFinal(hoja, datos, enunciado) {
  // Verificar si la lista de listas está vacía
  if (!datos || datos.length === 0) {
    Logger.log("No hay datos para agregar. Operación terminada.");
    return;
  }

  // Verificar si la hoja ya tiene contenido
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila === 0) {
    const columnas = datos[0].length; // Determinar el número de columnas de los datos
    hoja.getRange(1, 1, 1, columnas).setValues([enunciado]);
    hoja.getRange(2, 1, datos.length, datos[0].length).setValues(datos); // Escribir los datos debajo del enunciado
  } else {
    // Agregar los datos al final de lo existente
    const nuevaFilaInicio = ultimaFila + 1; // Determinar dónde comenzar a escribir
    hoja.getRange(nuevaFilaInicio, 1, datos.length, datos[0].length).setValues(datos);
  }
}
function reenviarFormularioPorRut(rutObjetivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMaestro = ss.getSheetByName("MAESTRO");
  const sheetProfesores = ss.getSheetByName("PROFESORES");

  const rut_profesores_jornada = sheetProfesores
    .getDataRange()
    .getDisplayValues()
    .map(fila => fila[1]);

  const colMail1 = obtenerNumeroDeColumna(sheetMaestro, "EMAIL PROFESOR 1", 1);
  const colMail2 = obtenerNumeroDeColumna(sheetMaestro, "EMAIL PROFESOR 2", 1);
  const colRut1  = obtenerNumeroDeColumna(sheetMaestro, "RUT PROFESOR 1", 1);
  const colRut2  = obtenerNumeroDeColumna(sheetMaestro, "RUT PROFESOR 2", 1);

  // FIX (NOMBRE vacío en LOG RECORDATORIOS / RESUMEN RECORDATORIOS): antes se
  // llamaba a registrarRecordatorio(...) con el nombre fijo en "". Se busca la
  // columna de nombre con buscarColumnaFlexible (no obtenerNumeroDeColumna) para
  // que no reviente si la hoja no tiene exactamente ese encabezado (con sus
  // saltos de línea/acentos tal cual) — en ese caso, simplemente se deja "".
  const encabezadosMaestroParaNombre = sheetMaestro.getRange(1, 1, 1, sheetMaestro.getLastColumn()).getValues()[0];
  const colNombre1 = buscarColumnaFlexible(encabezadosMaestroParaNombre, [
    "NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)",
    "NOMBRE PROFESOR BANNER 1"
  ]);
  const colNombre2 = buscarColumnaFlexible(encabezadosMaestroParaNombre, [
    "NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)",
    "NOMBRE PROFESOR 2"
  ]);

  const data = sheetMaestro.getDataRange().getDisplayValues();
  data.shift(); // eliminar encabezado

  // FIX ("Reenviar Rezagados" se detiene con un correo inválido): antes, un
  // solo error de enviarCorreo (correo mal escrito, GmailApp fallando, el
  // tipo de correo faltante en TEXTOS CORREO) lanzaba una excepción sin
  // capturar que cortaba la función a medias. Ahora se captura por intento,
  // queda registro en LOG RECORDATORIOS / RESUMEN RECORDATORIOS, y se informa
  // al final si algo falló, sin dejar de intentar el resto.
  const errores = [];

  data.forEach(fila => {
    const email1 = fila[colMail1];
    const email2 = fila[colMail2];
    const rut1   = fila[colRut1];
    const rut2   = fila[colRut2];
    const nombre1 = colNombre1 !== -1 ? String(fila[colNombre1] || "").trim() : "";
    const nombre2 = colNombre2 !== -1 ? String(fila[colNombre2] || "").trim() : "";

    if (rut1 === rutObjetivo && email1) {
      const esJornada1 = condicionPersonalizada(rut1, rut_profesores_jornada);
      const tipoCorreo1 = esJornada1 ? "REENVIO_PREFERENCIAS" : "REENVIO_DISPONIBILIDAD";
      const link1 = esJornada1 ? linkFormulario2 : linkFormulario1;
      try {
        enviarCorreo(email1, rut1, link1, tipoCorreo1);
        registrarRecordatorio(rut1, nombre1, email1, tipoCorreo1, "ENVIADO", "");
      } catch (e) {
        registrarRecordatorio(rut1, nombre1, email1, tipoCorreo1, "ERROR", e.message);
        errores.push(`RUT ${rut1} (${email1}): ${e.message}`);
      }
    }

    if (rut2 === rutObjetivo && email2) {
      const esJornada2 = condicionPersonalizada(rut2, rut_profesores_jornada);
      const tipoCorreo2 = esJornada2 ? "REENVIO_PREFERENCIAS" : "REENVIO_DISPONIBILIDAD";
      const link2 = esJornada2 ? linkFormulario2 : linkFormulario1;
      try {
        enviarCorreo(email2, rut2, link2, tipoCorreo2);
        registrarRecordatorio(rut2, nombre2, email2, tipoCorreo2, "ENVIADO", "");
      } catch (e) {
        registrarRecordatorio(rut2, nombre2, email2, tipoCorreo2, "ERROR", e.message);
        errores.push(`RUT ${rut2} (${email2}): ${e.message}`);
      }
    }
  });

  if (errores.length > 0) {
    SpreadsheetApp.getUi().alert(
      `Se eliminaron los registros de respuesta, pero el reenvío del formulario falló para:\n\n${errores.join("\n")}\n\nRevisa la hoja "${HOJA_LOG_RECORDATORIOS}" para el detalle.`
    );
  }
}

function escribirResultadoValidacion(hoja, encabezado, mensajes, colores, filaInicio) {
  if (!hoja) {
    throw new Error(`No se encontró la hoja para escribir resultados de validación.`);
  }

  const fila = filaInicio || 2;
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  let col = buscarColumnaFlexible(encabezados, [encabezado]);

  if (col === -1) {
    col = hoja.getLastColumn();
    hoja.getRange(1, col + 1).setValue(encabezado);
  }

  if (mensajes.length === 0) return;

  const filaCol = col + 1;

  hoja.getRange(fila, filaCol, mensajes.length, 1).setValues(mensajes.map(m => [m]));
  hoja.getRange(fila, filaCol, colores.length, 1).setBackgrounds(colores.map(c => [c]));
}

function getLogoBlobFromFolder_(folderId, fileName) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);

  if (!files.hasNext()) {
    throw new Error('No se encontró el archivo de firma "' + fileName + '" en la carpeta indicada.');
  }

  return files.next().getBlob();
}

function escapeHtml_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function textToHtmlPreserveBreaks_(text) {
  return escapeHtml_(text).replace(/\r\n|\n|\r/g, '<br>');
}

function buildHtmlBodyWithSignature_(mailBodyText) {
  const bodyHtml = textToHtmlPreserveBreaks_(mailBodyText);

  const signatureHtml =
    '<br><br>' +
    '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;">' +
    '<tr>' +
    '<td style="vertical-align:top; padding-right:12px;">' +
    '<img src="cid:logo" width="' + LOGO_WIDTH_PX + '" style="display:block;">' +
    '</td>' +
    '<td style="vertical-align:top; font-family:Arial, sans-serif; font-size:12px; line-height:1.35; color:#000;">' +
    '<div><strong>Dirección Curricular</strong></div>' +
    '<div>Facultad de Ingeniería y Ciencias Aplicadas</div>' +
    '<div>Universidad de los Andes</div>' +
    '<div>Av. Mons. Álvaro del Portillo 12.455</div>' +
    '<div>Las Condes, Santiago, Chile</div>' +
    '</td>' +
    '</tr>' +
    '</table>';

  return bodyHtml + signatureHtml;
}

function buildPlainSignature_() {
  return (
    'Dirección Curricular\n' +
    'Facultad de Ingeniería y Ciencias Aplicadas\n' +
    'Universidad de los Andes\n' +
    'Av. Mons. Álvaro del Portillo 12.455\n' +
    'Las Condes, Santiago, Chile'
  );
}

let _logoBlobCache_ = null;
function obtenerLogoFirma_() {
  if (!_logoBlobCache_) {
    _logoBlobCache_ = getLogoBlobFromFolder_(SIGNATURE_LOGO_FOLDER_ID, SIGNATURE_LOGO_FILE_NAME);
  }
  return _logoBlobCache_;
}

var id_hoja_programacion='1cmytVU4wEEMPz-E2Q7lJozkynGJCLpOHcIVRaemh1Pg'//Hoja de programacion para los cursos
var id_archivo_actual="1cmytVU4wEEMPz-E2Q7lJozkynGJCLpOHcIVRaemh1Pg"
var id_carpeta_archivos_cordinadores="1Hl9shX8Ju6dbWR0ZZw_undBNvEy6o2u2"//Cambiar aqui carpeta cordinadores
var id_archivo_cupos_externo = "1NsO_HPCvOBlm5bPN0dsjZpaM9CX_PqlSI5dzoBjdQOE"; // Archivo externo con la programación real de cupos
var nombre_hoja_cupos_externa = "PROGRAMACION 202710"; // Nombre de la hoja dentro de ese archivo — cambiar cada periodo
var personas_para_envio_de_hojas =[
    {name:"Matias Recabarren",mail:"mrecabarren@miuandes.cl",archivos:["COMPUTACION"]},
    {name:"Rafael Delpiano",mail:"rdelpiano@miuandes.cl",archivos:["OBRAS CIVILES"]},
    {name:"Pilar Gazmuri",mail:"mgazmuri@miuandes.cl",archivos:["ELECTRICA"]},
    {name:"Joaquín Carvajal",mail:"jjcarvajal@miuandes.cl",archivos:["MATEMATICA"]},
    {name:"Claudio Telha",mail:"catelha@miuandes.cl",archivos:["INDUSTRIAL"]},
    {name:"Orazio Descalzi",mail:"odescalzi@miuandes.cl",archivos:["FISICA"]},
    {name:"Sichem Guerrero",mail:"sguerrero@miuandes.cl",archivos:["AMBIENTAL","QUIMICA"]},
    {name:"Francisca Sáez",mail:"fdsaez@miuandes.cl",archivos:["FACULTAD"]},
    {name:"Francisca Sáez",mail:"fdsaez@miuandes.cl",archivos:["FACULTAD CEG"]}
    //cambiar aqui las diferentes areas de cordinacion y sus encargados

  ]
var personas_para_quitar_permisos =[
    {name:"Matias Recabarren",mail:"mrecabarren@miuandes.cl",archivos:["COMPUTACION"]},
    {name:"Rafael Delpiano",mail:"rdelpiano@miuandes.cl",archivos:["OBRAS CIVILES"]},
    {name:"Pilar Gazmuri",mail:"mgazmuri@miuandes.cl",archivos:["ELECTRICA"]},
    {name:"Joaquín Carvajal",mail:"jjcarvajal@miuandes.cl",archivos:["MATEMATICA"]},
    {name:"Claudio Telha",mail:"catelha@miuandes.cl",archivos:["INDUSTRIAL"]},
    {name:"Orazio Descalzi",mail:"odescalzi@miuandes.cl",archivos:["FISICA"]},
    {name:"Sichem Guerrero",mail:"sguerrero@miuandes.cl",archivos:["AMBIENTAL","QUIMICA"]},
    {name:"Francisca Sáez",mail:"fdsaez@miuandes.cl",archivos:["FACULTAD"]},
    {name:"Francisca Sáez",mail:"fdsaez@miuandes.cl",archivos:["FACULTAD CEG"]}
    //cambiar aqui las diferentes areas de cordinacion y sus encargados

  ]
// ====== CONFIG FIRMA UANDES ======
const SIGNATURE_LOGO_FOLDER_ID = '142sHFgpwHHZ6XB4y3ogfcAJsVX-kgFSv';
const SIGNATURE_LOGO_FILE_NAME = 'Logo_UANDES_Firma';
const LOGO_WIDTH_PX = 110;

var linkFormulario1 = "https://script.google.com/macros/s/AKfycbxpeKdK8S64pYmz5sGubGt9mBxCGovOEy59V5UFXIWImd9x1bLGg-bdbeVLXouXjMQbNA/exec"; // Cambiar por el link del formulario Honorarios
var linkFormulario2 = "https://script.google.com/macros/s/AKfycbyLywX9NM6ys4FZ2T1eDvua6QOLsR0Pkc7brFy6Amtr2nAei3y1LJikIXICoNiuj-aMJg/exec"; // Cambiar por el link del formulario Jornada
var mensaje_intermedio=""

function onOpen(){
  const ui = SpreadsheetApp.getUi(); //crea referencia a la interfaz de usuario de Spreadsheet

  ui.createMenu("Macros")

    .addSubMenu(
      ui.createMenu('Preparar Hojas')
        .addItem("Agregar secciones", "agregar_secciones")
        .addItem("Crear Calendario","crearCalendario")
        .addItem("Añadir cupos tentativos","anadirCuposTentativos")
        .addItem("Añadir listas cruzadas","anadirListasCruzadas")
        .addItem("Actualizar areas con catalogo antiguo","actualizarCatalogo")
        .addItem("Agregar Columna Horario Sala Especial","agregarDropdownUso")
        .addItem("Replicar info mandante → no mandante","replicarInfoMandante")
        .addItem("Actualizar CUPOS desde archivo externo","actualizarCuposDesdeArchivoExterno")
    )
    .addSubMenu(
      ui.createMenu('Coordinadores')
        .addItem("Crear hoja por area","crear_hoja_por_area")
        .addItem("Validar Cambios","validar_cambios")
        .addItem("Compartir enlaces","compartirYEnviarEnlace")
        .addItem("Quitar permisos de edicion archivos cordinaldores","quitar_permisos")
        .addItem("Extraer datos hojas coordinadores","extraer_datos")
        .addItem("Extraer datos de un área específica", "extraerDatosPorArea")
    )
    .addSubMenu(
      ui.createMenu('Formularios')
        .addItem('Crear Hojas De Respuesta', 'crearHojasDeRespuesta')
        .addItem('Inicializar textos de correo', 'inicializarTextosCorreo')
        .addItem('Eliminar Respuesta y Reenviar', 'borrarFilasPorRutConInput')
        .addItem('Reenviar Rezagados', 'enviarRecordatoriosNoRespondidos')
        .addItem("Envio de formularios Jornada y honorarios","enviarFormulariosCondicional")
        .addItem("Actualizar con datos formulario","actualizar_datos")
    )
    .addSubMenu(
      ui.createMenu('CrearHojas')
        .addItem("Crear Calendario","crearCalendario")
        .addItem("Crear archivo para DPSA","crear_template_dpsa")
        .addItem("Crear archivo para Alumnos","crear_template_HorariosING")
    )
    .addSubMenu(
      ui.createMenu('Alertas')
        .addItem("Generar reporte de alertas","generarReporteAlertas")
    )
    .addSubMenu(
      ui.createMenu('Validación')
        .addItem("Validar DPSA vs Listado_NRC","validar_dpsa")
    )

    .addToUi();

}
function FORZAR_AUTORIZACION() {
  UrlFetchApp.fetch("https://www.google.com");
}
function agregar_secciones(){
const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
const hoja_presupuesto = hojasActuales.getSheetByName('Presupuesto');
const semestre_espejo = Number(pedirParametro())
const data_presupuesto = hoja_presupuesto.getDataRange().getValues();

const columna_semestre = obtenerNumeroDeColumna(hoja_presupuesto,semestre_espejo,1)
const columna_curso = obtenerNumeroDeColumna(hoja_presupuesto,"CURSO",1)
const columna_codigo_curso = obtenerNumeroDeColumna(hoja_presupuesto,"Código curso",1)
const col_plan_estudio = obtenerNumeroDeColumna(hoja_presupuesto,"Plan",1)

data_presupuesto.shift()

const curso_codigo_nombre = obtener_columnas_especificas(data_presupuesto,[columna_codigo_curso,columna_curso,columna_semestre,col_plan_estudio])
//Ahora se obtiene el Curso-Materia
const hoja_catalogo = hojasActuales.getSheetByName('CATALOGO');
const data_catalogo = hoja_catalogo.getDataRange().getValues();
const col_materia = obtenerNumeroDeColumna(hoja_catalogo,"MATERIA",1)
const col_curso_catalogo = obtenerNumeroDeColumna(hoja_catalogo,"CURSO",1)
const col_codigo = obtenerNumeroDeColumna(hoja_catalogo,"CODIGO",1)

data_catalogo.shift()
const curso_materia = obtener_columnas_especificas(data_catalogo,[col_codigo,col_materia,col_curso_catalogo])



//Se obtienen los nrc

const hoja_listado_nrc = hojasActuales.getSheetByName('NRC POR PERIODO ESPEJO')
const data_listado_nrc = hoja_listado_nrc.getDataRange().getValues();
console.log(data_listado_nrc)
const columnaNRC =obtenerNumeroDeColumna(hoja_listado_nrc,"NRC",1)
const columnaCurso = obtenerNumeroDeColumna(hoja_listado_nrc,"CURSO",1)
const columnaMateria = obtenerNumeroDeColumna(hoja_listado_nrc,"MATERIA",1)
const cursosConSecciones = asignarSecciones(curso_codigo_nombre,data_listado_nrc,columnaCurso,columnaMateria,columnaNRC);
const hoja_maestro=hojasActuales.getSheetByName('MAESTRO');
const col_area_maestro = obtenerNumeroDeColumna(hoja_maestro,"AREA",1)
const col_curso_maestro = obtenerNumeroDeColumna(hoja_maestro,"CURSO",1)
const col_materia_maestro = obtenerNumeroDeColumna(hoja_maestro,"MATERIA",1)
const entradas = generar_Entradas(cursosConSecciones,data_listado_nrc,data_catalogo,hoja_catalogo,hoja_listado_nrc,hoja_presupuesto,hoja_maestro,curso_materia,curso_codigo_nombre).map((entrada)=>{

  if (entrada[col_curso_maestro].trim()=="electivos"){
    if(entrada[col_materia_maestro]=="ICC"){
      entrada[col_area_maestro]="COMPUTACION"
    }
    if(entrada[col_materia_maestro]=="ICA"){
      entrada[col_area_maestro]="AMBIENTAL"
    }
    if(entrada[col_materia_maestro]=="ICE"){
      entrada[col_area_maestro]="ELECTRICA"
    }
    if(entrada[col_materia_maestro]=="ICI"){
      entrada[col_area_maestro]="INDUSTRIAL"
    }
    if(entrada[col_materia_maestro]=="IOC"){
      entrada[col_area_maestro]="OBRAS CIVILES"
    }
    if(entrada[col_materia_maestro]=="ICQ"){
      entrada[col_area_maestro]="QUIMICA"
    }
    return entrada
  }
  return entrada

})


escribirDatosYResaltar(hoja_maestro,entradas)






}
function crear_hoja_por_area(){

const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
const hoja_maestro = hojasActuales.getSheetByName('MAESTRO');
const data_maestro = hoja_maestro.getDataRange().getValues();
const encabezado = data_maestro[0]
data_maestro.shift()

const col_area = obtenerNumeroDeColumna(hoja_maestro,"AREA",1)//NO CAMBIAR NOMBRE A LA COLUMNA
const areas = obtener_areas(data_maestro,col_area)
console.log(areas)
areas.forEach((area)=>{

  let entradas=obtener_entradas_area_especifica(area,data_maestro,col_area)
  console.log(entradas)
  let hoja_area = crear_hoja_nombre(area,hojasActuales)
  copiarEncabezado(hoja_maestro,hoja_area)
  escribirDatosYResaltar(hoja_area,entradas)
  let hojas_area_archivo=crearSpreadsheetEnCarpetaConRemplazo(area)
  hoja_area = crear_hoja_nombre(area,hojas_area_archivo)
  copiarEncabezado(hoja_maestro,hoja_area)

  escribirDatosYResaltar(hoja_area,entradas)
  deleteSheetByName(hojas_area_archivo,"Hoja 1")
  const columnasVisibles = [
    "CURSO MANDANTE",
    "AREA",
    "NRC",
    "LC",
    "CODIGO",
    "TITULO",
    "CUPOS",
    "INSCRITOS SEMESTRE ESPEJO ANTERIOR",
    "SECCIONES",
    "Clases A PROGRAMAR",
    "Ayudantías PROGRAMAR",
    "Laboratorios o Talleres PROGRAMAR",
    "Sala especial",
    "RUT PROFESOR 1",
    "NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)",
    "EMAIL PROFESOR 1",
    "RUT PROFESOR 2",
    "NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)",
    "EMAIL PROFESOR 2",
    "RUT PROFESOR LABT",
    "PROFESOR LABT",
    "MAIL PROFESOR LABT",
    "COMENTARIOS COORDINADOR"
  ];
  ocultarColumnasExcepto(hoja_area, columnasVisibles);
  ocultarFilasExcepto(hoja_area, "CURSO MANDANTE", "SI");
})




}
function extraer_datos(){
  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
  const hoja_maestro = hojasActuales.getSheetByName('MAESTRO');
  const data_maestro = hoja_maestro.getDataRange().getValues();
  data_maestro.shift()

  const encabezadosMaestro = hoja_maestro.getRange(1, 1, 1, hoja_maestro.getLastColumn()).getValues()[0];
  const col_area = obtenerNumeroDeColumna(hoja_maestro,"AREA",1)
  const areas = obtener_areas(data_maestro,col_area)
  const advertenciasEncabezado = []

  areas.forEach((area)=>{
    extraerDatosDeArea_(area, hojasActuales, hoja_maestro, encabezadosMaestro, advertenciasEncabezado)
  })

  if (advertenciasEncabezado.length > 0) {
    SpreadsheetApp.getUi().alert("Advertencia: columnas no encontradas al extraer datos\n\n" + advertenciasEncabezado.join("\n"));
  }
}

function extraerDatosPorArea(){
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt(
    "Extraer datos de un área específica",
    "Ingresa el nombre del área/especialidad exactamente como aparece en MAESTRO (por ejemplo: INDUSTRIAL)",
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const areaIngresada = respuesta.getResponseText().trim();
  if (!areaIngresada) {
    ui.alert("Debes ingresar un nombre de área.");
    return;
  }

  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
  const hoja_maestro = hojasActuales.getSheetByName('MAESTRO');
  const data_maestro = hoja_maestro.getDataRange().getValues();
  data_maestro.shift()

  const encabezadosMaestro = hoja_maestro.getRange(1, 1, 1, hoja_maestro.getLastColumn()).getValues()[0];
  const col_area = obtenerNumeroDeColumna(hoja_maestro,"AREA",1)
  const areasDisponibles = obtener_areas(data_maestro,col_area)

  const areaEncontrada = areasDisponibles.find(
    (area) => normalizarEncabezado(area) === normalizarEncabezado(areaIngresada)
  );

  if (!areaEncontrada) {
    ui.alert(`No se encontró el área "${areaIngresada}" en la hoja MAESTRO.\n\nÁreas disponibles:\n` + areasDisponibles.join("\n"));
    return;
  }

  const advertenciasEncabezado = []
  extraerDatosDeArea_(areaEncontrada, hojasActuales, hoja_maestro, encabezadosMaestro, advertenciasEncabezado)

  if (advertenciasEncabezado.length > 0) {
    ui.alert("Advertencia: columnas no encontradas al extraer datos\n\n" + advertenciasEncabezado.join("\n"));
  } else {
    ui.alert(`Datos de "${areaEncontrada}" extraídos correctamente.`);
  }
}

function extraerDatosDeArea_(area, hojasActuales, hoja_maestro, encabezadosMaestro, advertenciasEncabezado){
  let hojas_area_cordinador = obtenerSpreadsheetDeCarpeta(id_carpeta_archivos_cordinadores, area)
  let hoja_cordinador = hojas_area_cordinador.getSheetByName(area);
  let data_cordinador_bruta = hoja_cordinador.getDataRange().getValues();
  const encabezadosCordinador = data_cordinador_bruta.shift();

  const alineacion = alinearFilasPorEncabezado(data_cordinador_bruta, encabezadosCordinador, encabezadosMaestro);
  const data_cordinador = alineacion.filas;
  if (alineacion.columnasNoEncontradas.length > 0) {
    advertenciasEncabezado.push(`• ${area}: no se encontró en el archivo del coordinador la(s) columna(s) [${alineacion.columnasNoEncontradas.join(", ")}] — se dejaron vacías.`);
  }

  let hoja_area = crear_hoja_nombre(area, hojasActuales)
  const filasPrevias = hoja_area.getLastRow() > 1
    ? hoja_area.getRange(2, 1, hoja_area.getLastRow() - 1, hoja_area.getLastColumn()).getValues()
    : [];
  copiarEncabezado(hoja_maestro, hoja_area)
  const col_cupos_actuales = obtenerNumeroDeColumna(hoja_area,"CUPOS",1)
  const col_clave = obtenerNumeroDeColumna(hoja_area,"CODIGO",1)
  const col_seccion = obtenerNumeroDeColumna(hoja_area,"SECCIONES",1)
  compararYPintarCambios(hoja_area, filasPrevias, data_cordinador, 2, [col_clave, col_seccion], col_cupos_actuales)

  const numColumnaHorasClase=obtenerNumeroDeColumna(hoja_area,"Clases A PROGRAMAR",1,)
  const numColumnaHorasAyudantias=obtenerNumeroDeColumna(hoja_area,"Ayudantías PROGRAMAR",1)
  const numColumnaHorasLaboratorio=obtenerNumeroDeColumna(hoja_area,"Laboratorios o Talleres PROGRAMAR",1)
  const numColumnaRutProfesor1=obtenerNumeroDeColumna(hoja_area,"RUT PROFESOR 1",1)
  const numColumnaRutProfesorLab=obtenerNumeroDeColumna(hoja_area,"RUT PROFESOR LABT",1)
  const casosRaros = []

  data_cordinador.forEach((fila) => {
    const codigo  = fila[col_clave]
    const seccion = fila[col_seccion]

    const horasClase  = String(fila[numColumnaHorasClase] || "").trim()
    const horasTaller = String(fila[numColumnaHorasLaboratorio] || "").trim()
    const rutProf1    = String(fila[numColumnaRutProfesor1] || "").trim()
    const rutLab      = String(fila[numColumnaRutProfesorLab] || "").trim()

    const tieneClase  = horasClase !== ""
    const tieneTaller = horasTaller !== ""

    if (tieneClase && tieneTaller && rutLab === "") {
      casosRaros.push(`• ${codigo} / ${seccion} → Clases + Lab sin RUT LAB`)
    }

    if (!tieneClase && tieneTaller && rutProf1 === "") {
      casosRaros.push(`• ${codigo} / ${seccion} → Curso Solo Lab sin RUT PROFESOR 1`)
    }
  })

  if (casosRaros.length > 0) {
    SpreadsheetApp.getUi().alert(`Casos a revisar en ${area}:\n\n` + casosRaros.join("\n"))
  }
}




function mostrarTodasLasHojas() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const usuarioActual = Session.getEffectiveUser().getEmail();
  const hojas = libro.getSheets();

  hojas.forEach((hoja) => hoja.showSheet());

  Logger.log(`Todas las hojas son visibles ahora para: ${usuarioActual}`);
}






function pedirParametro() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt('Introduce el semestre a programar', 'Escribe el valor que deseas usar más tarde:', ui.ButtonSet.OK_CANCEL);

  // Verificar si el usuario presionó "OK" antes de devolver el texto
  if (respuesta.getSelectedButton() == ui.Button.OK) {
    return respuesta.getResponseText(); // Retorna solo el texto ingresado
  } else {
    return null; // Retorna `null` si el usuario canceló
  }
}
function pedirParametro_fechas(mensaje_grande,mensaje_pequeno) {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt(mensaje_grande, mensaje_pequeno, ui.ButtonSet.OK_CANCEL);

  // Verificar si el usuario presionó "OK" antes de devolver el texto
  if (respuesta.getSelectedButton() == ui.Button.OK) {
    return respuesta.getResponseText(); // Retorna solo el texto ingresado
  } else {
    return null; // Retorna `null` si el usuario canceló
  }
}
function compartirYEnviarEnlace() {

    enviarLinksConPermisos(id_carpeta_archivos_cordinadores,personas_para_envio_de_hojas)




}
// Candidatos de nombre para la columna llave de sección: distintos coordinadores
// han guardado esta columna con o sin espacio antes del guion, así que se
// aceptan ambas variantes (igual que en construirMapaColumnasMaestro).
const NOMBRES_COLUMNA_LLAVE_SECCION = [
  "LLAVE Código- sec",
  "LLAVE Código - sec",
  "LLAVE Codigo- sec",
  "LLAVE Codigo - sec"
];

function validar_cambios(){
  const ui = SpreadsheetApp.getUi();
  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
  const hoja_actual = hojasActuales.getActiveSheet();
  const hoja_maestro = hojasActuales.getSheetByName("MAESTRO");

  // ===================== GUARDAS DE SEGURIDAD =====================
  // FIX (pérdida de datos en MAESTRO): la versión anterior fusionaba las
  // filas POR POSICIÓN (actualizarListaOriginal reemplazaba la fila completa
  // de MAESTRO por la fila de la hoja activa en cuanto detectaba una
  // diferencia). Cuando la hoja del coordinador tenía menos columnas que
  // MAESTRO, eso dejaba filas de distinto ancho mezcladas en el mismo
  // arreglo, lo que escribirDatosYResaltar no podía escribir y terminaba
  // borrando MAESTRO sin poder volver a escribirlo. Ahora la fusión es POR
  // NOMBRE DE COLUMNA: solo se copian, columna por columna, los valores de
  // las columnas que existen en ambas hojas, así que el ancho de cada fila de
  // MAESTRO nunca cambia.
  if (!hoja_maestro) {
    ui.alert('No se encontró la hoja "MAESTRO". No se realizaron cambios.');
    return;
  }
  if (hoja_actual.getName() === "MAESTRO") {
    ui.alert('La hoja activa es "MAESTRO". Selecciona la hoja del coordinador con los cambios antes de ejecutar "Validar Cambios".');
    return;
  }

  const encabezados_actual = hoja_actual.getRange(1, 1, 1, hoja_actual.getLastColumn()).getValues()[0];
  const encabezados_maestro = hoja_maestro.getRange(1, 1, 1, hoja_maestro.getLastColumn()).getValues()[0];

  const colCuposActual = buscarColumnaFlexible(encabezados_actual, ["CUPOS"]);
  const colClaveActual = buscarColumnaFlexible(encabezados_actual, NOMBRES_COLUMNA_LLAVE_SECCION);

  if (colCuposActual === -1 || colClaveActual === -1) {
    ui.alert('La hoja activa no tiene las columnas "CUPOS" y/o "LLAVE Código- sec", así que no se puede validar contra MAESTRO. No se modificó MAESTRO.');
    return;
  }

  const colClaveMaestro = buscarColumnaFlexible(encabezados_maestro, NOMBRES_COLUMNA_LLAVE_SECCION);
  if (colClaveMaestro === -1) {
    ui.alert('No se encontró la columna "LLAVE Código- sec" en MAESTRO. No se modificó MAESTRO.');
    return;
  }

  // Mapear, por nombre de encabezado, qué columnas de la hoja activa
  // corresponden a qué columnas de MAESTRO. Solo esas se van a fusionar.
  const mapaColumnas = [];
  encabezados_actual.forEach((nombre, colActual) => {
    if (!nombre) return;
    const colMaestro = buscarColumnaFlexible(encabezados_maestro, [nombre]);
    if (colMaestro !== -1) {
      mapaColumnas.push({ colActual, colMaestro });
    }
  });

  if (mapaColumnas.length === 0) {
    ui.alert("La hoja activa no comparte ninguna columna reconocible con MAESTRO. No se modificó MAESTRO.");
    return;
  }

  const filasActivas = Math.max(hoja_actual.getLastRow() - 1, 0);
  const filasMaestro = Math.max(hoja_maestro.getLastRow() - 1, 0);
  const data_verificada = filasActivas > 0
    ? hoja_actual.getRange(2, 1, filasActivas, hoja_actual.getLastColumn()).getValues()
    : [];
  const data_maestro = filasMaestro > 0
    ? hoja_maestro.getRange(2, 1, filasMaestro, hoja_maestro.getLastColumn()).getValues()
    : [];

  // Indexar MAESTRO por llave de sección para poder ubicar cada fila de la
  // hoja activa sin depender del orden/posición de las filas.
  const indicePorLlave = new Map();
  data_maestro.forEach((fila, i) => {
    const llave = String(fila[colClaveMaestro]).trim();
    if (llave) indicePorLlave.set(llave, i);
  });

  let filasActualizadas = 0;
  const seccionesNoEncontradas = [];

  data_verificada.forEach((filaActual) => {
    const llave = String(filaActual[colClaveActual]).trim();
    if (!llave) return;

    const indiceMaestro = indicePorLlave.get(llave);
    if (indiceMaestro === undefined) {
      seccionesNoEncontradas.push(llave);
      return;
    }

    const filaMaestro = data_maestro[indiceMaestro];
    let huboCambioEnFila = false;

    mapaColumnas.forEach(({ colActual, colMaestro }) => {
      const valorActual = filaActual[colActual];
      const valorMaestro = filaMaestro[colMaestro];
      if (String(valorActual) !== String(valorMaestro)) {
        filaMaestro[colMaestro] = valorActual; // Fusión POR NOMBRE DE COLUMNA, nunca por posición.
        huboCambioEnFila = true;
      }
    });

    if (huboCambioEnFila) filasActualizadas++;
  });

  if (filasActualizadas === 0) {
    ui.alert(
      seccionesNoEncontradas.length > 0
        ? `No se detectaron cambios que aplicar a MAESTRO. Además, ${seccionesNoEncontradas.length} sección(es) de la hoja activa no se encontraron en MAESTRO (revisa la columna "LLAVE Código- sec").`
        : "No se detectaron cambios que aplicar a MAESTRO."
    );
    return;
  }

  // escribirDatosYResaltar valida el ancho de cada fila antes de borrar nada
  // y respalda MAESTRO automáticamente antes de escribir.
  escribirDatosYResaltar(hoja_maestro, data_maestro);

  let mensaje = `Validación completa: se actualizaron ${filasActualizadas} fila(s) en MAESTRO (fusión por nombre de columna).`;
  if (seccionesNoEncontradas.length > 0) {
    mensaje += `\n\n${seccionesNoEncontradas.length} sección(es) de la hoja activa no se encontraron en MAESTRO y se ignoraron:\n${seccionesNoEncontradas.slice(0, 30).join(", ")}${seccionesNoEncontradas.length > 30 ? "…" : ""}`;
  }
  ui.alert(mensaje);
}
function actualizar_datos() {
  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();

  const hoja_maestro = hojasActuales.getSheetByName('MAESTRO');
  const hoja_profesores = hojasActuales.getSheetByName('PROFESORES');
  const hoja_respuestas = hojasActuales.getSheetByName('RESPUESTAS');

  // ---------- RESPUESTAS ----------
  let data_respuestas = hoja_respuestas.getDataRange().getValues();
  data_respuestas.shift(); // ⬅️ eliminar encabezado

  const col_lunes = obtenerNumeroDeColumna(hoja_maestro, "LUNES", 1);
  const col_martes = obtenerNumeroDeColumna(hoja_maestro, "MARTES", 1);
  const col_miercoles = obtenerNumeroDeColumna(hoja_maestro, "MIERCOLES", 1);
  const col_jueves = obtenerNumeroDeColumna(hoja_maestro, "JUEVES", 1);
  const col_viernes = obtenerNumeroDeColumna(hoja_maestro, "VIERNES", 1);
  const col_rut_profesor = obtenerNumeroDeColumna(hoja_maestro, "RUT PROFESOR 1", 1);
  const col_rut_profesor2 = obtenerNumeroDeColumna(hoja_maestro, "RUT PROFESOR 2", 1);
  const col_Metodologia = obtenerNumeroDeColumna(hoja_maestro, "METODOLOGIA", 1);


  const horarios_por_rut = horariosPorRut(data_respuestas);
  const columnasDias = [col_lunes, col_martes, col_miercoles, col_jueves, col_viernes];

  // ---------- PROFESORES ----------
  let rut_profesores_jornada = hoja_profesores
    .getDataRange()
    .getValues();
  rut_profesores_jornada.shift(); // ⬅️ eliminar encabezado
  rut_profesores_jornada = rut_profesores_jornada.map(fila => fila[1]);

  limpiarColumnas(hoja_maestro, columnasDias);

  // ---------- MAESTRO ----------
  let data_maestro = hoja_maestro.getDataRange().getValues();
  data_maestro.shift(); // ⬅️ eliminar encabezado

  const horarios_jornada =
    "8:30-9:20,9:30-10:20,10:30-11:20,11:30-12:20,12:30-13:20," +
    "13:30-14:20,14:30-15:20,15:30-16:20,16:30-17:20," +
    "17:30-18:20,18:30-19:20";

  const dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];

  data_maestro.forEach((fila) => {
    const rut = fila[col_rut_profesor];
    const rut2 = fila[col_rut_profesor2];

    if (condicionPersonalizada(rut, rut_profesores_jornada)) {
      columnasDias.forEach(col => fila[col] = horarios_jornada);
    }

    if (condicionPersonalizada(rut2, rut_profesores_jornada)) {
      columnasDias.forEach(col => fila[col] = horarios_jornada);
    }

    const profesores = horarios_por_rut.filter(
      prof => prof.rut === rut || prof.rut === rut2
    );

    profesores.forEach((profesor) => {
      if (!profesor) return;

      let escrito = verificarCasillasNoVacias(fila, columnasDias);

      columnasDias.forEach((columnaDia, i) => {
        const dia = dias[i];
        if (!escrito) {
          fila[columnaDia] = profesor[dia] || "";
        } else {
          fila[columnaDia] =
            interseccion_de_bloques(profesor[dia], fila[columnaDia]);
        }
      });

      if (verificarCasillastodasVacias(fila, columnasDias) && escrito) {
        columnasDias.forEach(col => fila[col] = "INCOHERENCIA");
      }
    });
  });

  // ---------- PREFERENCIAS ----------
  const hoja_preferencias = hojasActuales.getSheetByName('PREFERENCIAS');
  let data_preferencias = hoja_preferencias.getDataRange().getValues();
  data_preferencias.shift(); // ⬅️ eliminar encabezado

  const col_codigo = obtenerNumeroDeColumna(hoja_maestro, "CODIGO", 1);
  const col_seccion = obtenerNumeroDeColumna(hoja_maestro, "SECCIONES", 1);
  const col_preferencias = obtenerNumeroDeColumna(
    hoja_maestro,
    "2+1 o 3? (distribución horario de clases)",
    1
  );

  data_maestro.forEach(fila => {
    const curso_seccion = data_preferencias.find(
      prof => prof[1] == fila[col_codigo] && prof[2] == fila[col_seccion]
    );
    if (curso_seccion) {
      fila[col_preferencias] = curso_seccion[3];
    }
  });

  // ---------- OTROS ----------
  const hoja_otros = hojasActuales.getSheetByName('OTROS');
  let data_otros = hoja_otros.getDataRange().getValues();
  data_otros.shift(); // ⬅️ eliminar encabezado

  const col_comentarios = obtenerNumeroDeColumna(hoja_maestro, "COMENTARIOS PROFESOR", 1);
  const col_examen = obtenerNumeroDeColumna(hoja_maestro, "EXAMEN (Sí o No)", 1);
  const col_fechas_prueba = obtenerNumeroDeColumna(hoja_maestro, "COMENTARIOS PRUEBA", 1);
  const col_cantidad_evaluaciones = obtenerNumeroDeColumna(
    hoja_maestro,
    "CANTIDAD EVALUACIONES (semestrales)",
    1
  );
  const col_tipo = obtenerNumeroDeColumna(
    hoja_maestro,
    "HORARIO EVALUACIONES (19:30 o en clases o en ayudantías)",
    1
  );
    const col_comentarios_prueba = obtenerNumeroDeColumna(
    hoja_maestro,
    "COMENTARIOS PRUEBA",
    1
  );
  const col_check_respuesta = obtenerNumeroDeColumna(
    hoja_maestro,
    "INGRESO DE HORARIO",
    1
  );
  const col_comentarios2 = obtenerNumeroDeColumna(hoja_maestro, "COMENTARIOS PROFESOR 2", 1);
  const col_fechas_prueba2 = obtenerNumeroDeColumna(hoja_maestro, "COMENTARIOS PRUEBA 2", 1);

  data_maestro.forEach(fila => {
    const cursos_seccion = data_otros.filter(
      prof => prof[1] == fila[col_codigo] && prof[2] == fila[col_seccion]
    );
    cursos_seccion.forEach(curso => {
      const rut = curso[0];
      if (rut == fila[col_rut_profesor]) {
        fila[col_comentarios] = curso[3];
        fila[col_examen] = curso[4];
        fila[col_cantidad_evaluaciones] = curso[5];
        fila[col_tipo] = curso[6];
        fila[col_Metodologia] = curso[7];
        fila[col_fechas_prueba] = curso[8];
      } else if (rut == fila[col_rut_profesor2]) {
        fila[col_comentarios2] = curso[3];
        fila[col_fechas_prueba2] = curso[8];
      }
    });
    if (cursos_seccion.length > 0) {
      fila[col_check_respuesta] = "SI";
    }
  });

  escribirDatosYResaltar(hoja_maestro, data_maestro);
}


function enviarFormulariosCondicional() {

  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = hojasActuales.getSheetByName("MAESTRO");
  const hoja_profesores = hojasActuales.getSheetByName("PROFESORES");
  const ui = SpreadsheetApp.getUi();

  const rut_profesores_jornada = hoja_profesores
    .getDataRange()
    .getDisplayValues()
    .map((fila)=>fila[1]);

  const hojaEnviados = obtenerOcrearHojaFormulariosEnviados_(hojasActuales);
  const rutsYaNotificados = leerRutsFormulariosEnviados_(hojaEnviados);
  const enviados = new Set(rutsYaNotificados);

  const hojaErrores = obtenerOcrearHojaErroresEnvio_(hojasActuales);

  const colMail1 = obtenerNumeroDeColumna(sheet,"EMAIL PROFESOR 1",1);
  const colMail2 = obtenerNumeroDeColumna(sheet,"EMAIL PROFESOR 2",1);
  const colRut1 = obtenerNumeroDeColumna(sheet,"RUT PROFESOR 1",1);
  const colRut2 = obtenerNumeroDeColumna(sheet,"RUT PROFESOR 2",1);

  const data = sheet.getDataRange().getDisplayValues();
  data.shift();

  // 1. Armar la lista de envíos candidatos de esta ejecución (todavía sin enviar nada)
  const enviadosEnEstaCorrida = new Set();
  const candidatos = [];

  data.forEach((fila) => {
    const email1 = fila[colMail1];
    const email2 = fila[colMail2];
    const rut1 = String(fila[colRut1] || "").trim();
    const rut2 = String(fila[colRut2] || "").trim();

    if (email1 && rut1 && !enviados.has(rut1) && !enviadosEnEstaCorrida.has(rut1)) {
      const esJornada1 = condicionPersonalizada(rut1, rut_profesores_jornada);
      candidatos.push({
        rut: rut1,
        email: email1,
        tipoCorreo: esJornada1 ? "PREFERENCIAS" : "DISPONIBILIDAD",
        link: esJornada1 ? linkFormulario2 : linkFormulario1
      });
      enviadosEnEstaCorrida.add(rut1);
    }

    if (email2 && rut2 && !enviados.has(rut2) && !enviadosEnEstaCorrida.has(rut2)) {
      const esJornada2 = condicionPersonalizada(rut2, rut_profesores_jornada);
      candidatos.push({
        rut: rut2,
        email: email2,
        tipoCorreo: esJornada2 ? "PREFERENCIAS" : "DISPONIBILIDAD",
        link: esJornada2 ? linkFormulario2 : linkFormulario1
      });
      enviadosEnEstaCorrida.add(rut2);
    }
  });

  if (candidatos.length === 0) {
    ui.alert("No hay profesores pendientes de recibir el formulario (todos ya fueron notificados, o no tienen RUT/correo registrado).");
    return;
  }

  // 2. Verificación contra PROFESORES + HISTORICO PROFESORES
  const normalizarRut_ = (rut) => String(rut || "").trim().toUpperCase();
  const directorioDocentes = obtenerDirectorioDocentes_();
  const rutsConocidos = new Set(directorioDocentes.map((docente) => normalizarRut_(docente.rut)));

  const rutsNoEncontrados = candidatos
    .map((c) => c.rut)
    .filter((rut) => !rutsConocidos.has(normalizarRut_(rut)));

  if (rutsNoEncontrados.length > 0) {
    const LIMITE_MOSTRAR = 30;
    const listaMostrada = rutsNoEncontrados.slice(0, LIMITE_MOSTRAR).join("\n");
    const extra = rutsNoEncontrados.length > LIMITE_MOSTRAR
      ? `\n... y ${rutsNoEncontrados.length - LIMITE_MOSTRAR} más`
      : "";

    const respuesta = ui.alert(
      "RUT no reconocidos como profesores históricos",
      `${rutsNoEncontrados.length} RUT que recibirían el correo no se encontraron en PROFESORES ni en HISTORICO PROFESORES:\n\n${listaMostrada}${extra}\n\n¿Deseas continuar de todas formas con el envío de correos, o prefieres cancelar para corregir primero?`,
      ui.ButtonSet.YES_NO
    );

    if (respuesta !== ui.Button.YES) {
      ui.alert("Envío cancelado. No se envió ningún correo.");
      return;
    }
  }

  // 3. Envío real, registrando de inmediato cada resultado
  let totalEnviados = 0;
  let totalErrores = 0;

  candidatos.forEach((candidato) => {
    try {
      enviarCorreo(candidato.email, candidato.rut, candidato.link, candidato.tipoCorreo);
      hojaEnviados.appendRow([candidato.rut, new Date()]);
      totalEnviados++;
    } catch (e) {
      hojaErrores.appendRow([candidato.rut, candidato.email, candidato.tipoCorreo, new Date(), e.message]);
      totalErrores++;
    }
  });

  if (totalErrores > 0) {
    ui.alert(
      `Envío de formularios terminado.\n\nEnviados correctamente: ${totalEnviados}\nErrores: ${totalErrores} (revisa la hoja "ERRORES ENVIO FORMULARIOS").\n\nLos que fallaron no quedaron marcados como enviados, así que se reintentarán en la próxima ejecución.`
    );
  } else {
    ui.alert(`Envío de formularios terminado. Correos enviados: ${totalEnviados}.`);
  }

}

function obtenerOcrearHojaErroresEnvio_(spreadsheet) {
  let hoja = spreadsheet.getSheetByName("ERRORES ENVIO FORMULARIOS");
  if (!hoja) {
    hoja = spreadsheet.insertSheet("ERRORES ENVIO FORMULARIOS");
    hoja.getRange(1, 1, 1, 5).setValues([["RUT", "EMAIL", "TIPO CORREO", "FECHA", "ERROR"]]);
  }
  return hoja;
}

function obtenerOcrearHojaFormulariosEnviados_(spreadsheet) {
  let hoja = spreadsheet.getSheetByName("FORMULARIOS ENVIADOS");
  if (!hoja) {
    hoja = spreadsheet.insertSheet("FORMULARIOS ENVIADOS");
    hoja.getRange(1, 1, 1, 2).setValues([["RUT", "FECHA ENVIO"]]);
  }
  return hoja;
}

function leerRutsFormulariosEnviados_(hoja) {
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  const data = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  return data
    .map((fila) => String(fila[0] || "").trim())
    .filter((rut) => rut !== "");
}

// Función para enviar el correo
function enviarCorreo(email, rut, linkFormulario, tipoCorreo) {
  const textoBase = obtenerTextoCorreo_(tipoCorreo);
  const mensaje = String(textoBase.cuerpo).split("{{LINK}}").join(linkFormulario);
  const asunto = textoBase.asunto;

  const logoBlob = obtenerLogoFirma_();
  const plainBody = mensaje + '\n\n' + buildPlainSignature_();
  const htmlBody = buildHtmlBodyWithSignature_(mensaje);

  GmailApp.sendEmail(email, asunto, plainBody, {
    htmlBody: htmlBody,
    inlineImages: { logo: logoBlob }
  });
}

function obtenerTextoCorreo_(tipo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TEXTOS CORREO");
  if (!hoja) {
    throw new Error('No existe la hoja "TEXTOS CORREO". Ejecuta primero "Inicializar textos de correo" desde el menú.');
  }
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(h => String(h).trim().toUpperCase());
  const colTipo = encabezados.indexOf("TIPO");
  const colAsunto = encabezados.indexOf("ASUNTO");
  const colCuerpo = encabezados.indexOf("CUERPO");

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][colTipo]).trim() === tipo) {
      return { asunto: datos[i][colAsunto], cuerpo: datos[i][colCuerpo] };
    }
  }
  throw new Error(`No se encontró el tipo de correo "${tipo}" en la hoja TEXTOS CORREO.`);
}

function inicializarTextosCorreo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("TEXTOS CORREO");
  if (!hoja) {
    hoja = ss.insertSheet("TEXTOS CORREO");
  }
  hoja.clearContents();

  const filas = [
    ["TIPO", "ASUNTO", "CUERPO"],
    ["DISPONIBILIDAD", "FORMULARIO DE DISPONIBILIDAD Y PREFERENCIAS",
      "Estimado profesor/a,\n\nPor favor, complete el siguiente formulario para su disponibilidad del semestre: {{LINK}}\n\nMuchas Gracias.\nArea de Analisis Curricular"],
    ["PREFERENCIAS", "Programación de Cursos Pre-Grado 2026-20 ING",
      "Estimado Profesor/a:\n\nDe cara a la Programación Académica de los cursos de Pre-Grado, el Coordinador de Área nos ha informado que Ud. realizará curso(s) en el primer semestre del 2027.\nLe rogamos que complete la información solicitada en el siguiente formulario.\n\nEl plazo para esto es hasta el día 16 de octubre.\n\nLink al formulario: {{LINK}}.\n\nPor favor tener en consideración los siguientes puntos:\n- El curso se realizará entre marzo y julio de 2027.\n- La modalidad de las clases será 100% presencial durante el próximo semestre.\n\nQuedamos atentos a cualquier duda.\n\nDirección Curricular - Facultad de Ingeniería y Ciencias Aplicadas - Universidad de los Andes"],
    ["RECORDATORIO_DISPONIBILIDAD", "RECORDATORIO: FORMULARIO DE DISPONIBILIDAD Y PREFERENCIAS",
      "Estimado profesor/a,\n\nLe recordamos completar el siguiente formulario:\n{{LINK}}\n\nMuchas gracias.\nÁrea de Análisis Curricular"],
    ["RECORDATORIO_PREFERENCIAS", "RECORDATORIO: FORMULARIO DE PREFERENCIAS",
      "Estimado profesor/a,\n\nLe recordamos completar el siguiente formulario:\n{{LINK}}\n\nMuchas gracias.\nÁrea de Análisis Curricular"],
    ["REENVIO_DISPONIBILIDAD", "FORMULARIO DE DISPONIBILIDAD Y PREFERENCIAS",
      "Estimado profesor/a,\n\nHemos reiniciado su respuesta. Por favor complete el siguiente formulario:\n{{LINK}}\n\nMuchas gracias.\nÁrea de Análisis Curricular"],
    ["REENVIO_PREFERENCIAS", "FORMULARIO DE PREFERENCIAS",
      "Estimado profesor/a,\n\nHemos reiniciado su respuesta. Por favor complete el siguiente formulario:\n{{LINK}}\n\nMuchas gracias.\nÁrea de Análisis Curricular"]
  ];

  hoja.getRange(1, 1, filas.length, 3).setValues(filas);
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 3);

  SpreadsheetApp.getUi().alert('Hoja "TEXTOS CORREO" creada/actualizada. Ya puedes editar ASUNTO y CUERPO directamente ahí — usa {{LINK}} donde quieras que aparezca el link al formulario.');
}
// Ejemplo de una condición personalizada
function condicionPersonalizada(rut,rut_profesores_jornada) {


  // Define tu lógica aquí, por ejemplo:
  // Retorna true si rut1 termina en 'K' o false en caso contrario
  return rut_profesores_jornada.includes(rut) ;
}
function crear_template_dpsa(){
  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
  const hoja_maestro = hojasActuales.getSheetByName("MAESTRO");
  const data_maestro = hoja_maestro.getDataRange().getDisplayValues();
  const encabezado_completo = hoja_maestro.getDataRange().getValues()[0]
  data_maestro.shift()
  const col_lunes= obtenerNumeroDeColumna(hoja_maestro,"Lunes",1)
  const col_martes= obtenerNumeroDeColumna(hoja_maestro,"Martes",1)
  const col_miercoles= obtenerNumeroDeColumna(hoja_maestro,"Miercoles",1)
  const col_jueves= obtenerNumeroDeColumna(hoja_maestro,"Jueves",1)
  const col_viernes= obtenerNumeroDeColumna(hoja_maestro,"Viernes",1)
  const col_dias=[col_lunes,col_martes,col_miercoles,col_jueves,col_viernes]
  const dias=["Lunes","Martes","Miercoles","Jueves","Viernes"]
  const encabezado=["NRC","MATERIA","CURSO","TITULO","SECC.","LISTA CRUZADA","IDENTIFICADOR LIGA","TIPO HORARIO","PARTE PERIODO","CALIFICABLE","CUPOS TOTALES NRC","TIPO DE REUNIÓN","FECHA INICIO","FECHA FIN","DIA","HORA INICIO","HORA FIN","INDICADOR SESION LINEA/PROFESOR","PROFESOR PRINCIPAL","OTROS PROFESORES","LABORATORIO","CARACTERÍSTICA ESPECIAL SALA PRORIDAD 1","CARACTERÍSTICA ESPECIAL SALA PRORIDAD 2","OBSERVACIÓN","METODOLOGIA DE ENSEÑANZA"]
  const fecha_inicio_clases=pedirParametro_fechas("Introduce la fecha de INICIO de clases","dd/mm/aaaa")
  const fecha_fin_clases=pedirParametro_fechas("Introduce la fecha de FIN de clases","dd/mm/aaaa")
  const fecha_inicio_ayud=pedirParametro_fechas("Introduce la fecha de \INICIO de AYUD/LAB","en formato dd/mm/aaaa")
  const fecha_fin_ayud=pedirParametro_fechas("Introduce la fecha de FIN de AYUD/LAB","dd/mm/aaaa")

  const dPSA=agregar_fechas(rellenar_dpsa(data_maestro,col_dias,dias,hoja_maestro,encabezado_completo),fecha_inicio_clases,fecha_fin_clases,fecha_inicio_ayud,fecha_fin_ayud)

  dPSA.unshift(encabezado)
  const hoja_dpsa=crear_hoja_nombre("DPSA",hojasActuales)
  const data_dpsa_anterior=hoja_dpsa.getDataRange().getDisplayValues()
  hoja_dpsa.getDataRange().clear()
  escribirListaDeListas(hoja_dpsa,dPSA,1,1)

  const hoja_ajustes=crear_hoja_nombre("AJUSTES DPSA",hojasActuales)

  const llaves_p=[0,1,2,4,5,7]
  const enunciado_ajustes=["AJUSTES","HORA-DIA CAMBIO","NRC","MATERIA","CURSO","TITULO","SECC.","LISTA CRUZADA","IDENTIFICADOR LIGA","TIPO HORARIO","PARTE PERIODO","CALIFICABLE","CUPOS TOTALES NRC","TIPO DE REUNIÓN","FECHA INICIO","FECHA FIN","DIA","HORA INICIO","HORA FIN","INDICADOR SESION LINEA/PROFESOR","PROFESOR PRINCIPAL","OTROS PROFESORES","LABORATORIO","CARACTERÍSTICA ESPECIAL SALA PRORIDAD 1","CARACTERÍSTICA ESPECIAL SALA PRORIDAD 2","OBSERVACIÓN","METODOLOGIA DE ENSEÑANZA"]
  const cambios=compareLists(normalizarListaDPSA(data_dpsa_anterior,llaves_p,encabezado),normalizarListaDPSA(dPSA,llaves_p,encabezado),llaves_p,enunciado_ajustes).filter((elemento)=>elemento[0]!="")
  if(cambios.length){
    const cols=cambios[0].length
    const validos=cambios.filter(f=>f.length==cols)
    if(validos.length) agregarAlFinal(hoja_ajustes,validos,enunciado_ajustes)
  }
}
function validar_dpsa(){
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja_dpsa = ss.getSheetByName("DPSA");
    const hoja_listado = ss.getSheetByName("LISTADO NRC ACTUAL");
    const hoja_maestro = ss.getSheetByName("MAESTRO");

    if (!hoja_dpsa) throw new Error(`No existe la hoja "DPSA".`);
    if (!hoja_listado) throw new Error(`No existe la hoja "LISTADO NRC ACTUAL".`);
    if (!hoja_maestro) throw new Error(`No existe la hoja "MAESTRO".`);

    const encabezados_dpsa = hoja_dpsa.getDataRange().getValues()[0];
    const encabezados_listado = hoja_listado.getDataRange().getValues()[0];
    const encabezados_maestro = hoja_maestro.getDataRange().getValues()[0];

    const cD = construirMapaColumnasDPSA(encabezados_dpsa);
    const cL = construirMapaColumnasListado(encabezados_listado);
    const cM = construirMapaColumnasMaestro(encabezados_maestro);

    const leerFilas = (hoja) => {
      const filas = hoja.getLastRow() - 1;
      if (filas < 1) return [];
      return hoja.getRange(2, 1, filas, hoja.getLastColumn()).getValues();
    };

    const rows_dpsa = leerFilas(hoja_dpsa);
    const rows_listado = leerFilas(hoja_listado);
    const rows_maestro = leerFilas(hoja_maestro);

    const resultado = validarDPSAVsListado(rows_dpsa, rows_listado, rows_maestro, cD, cL, cM);

    escribirResultadoValidacion(hoja_dpsa, "VALIDACIÓN DPSA vs Listado_NRC", resultado.mensajes, resultado.colores, 2);

    const total = resultado.mensajes.length;
    const ok = resultado.mensajes.filter(m => m === "OK").length;
    const fallas = total - ok;

    ui.alert(`Validación DPSA vs Listado_NRC completada.\n\n${total} filas revisadas.\n${ok} OK · ${fallas} con diferencias.`);
  } catch (e) {
    ui.alert(`Error en validación DPSA:\n${e.message}`);
  }
}
function crear_template_HorariosING(){

  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
  const hoja_maestro = hojasActuales.getSheetByName("MAESTRO");
  const data_maestro = hoja_maestro.getDataRange().getValues();
  const encabezado_fechas=data_maestro[0]
  data_maestro.shift()
  const col_lunes= obtenerNumeroDeColumna(hoja_maestro,"Lunes",1)
  const col_martes= obtenerNumeroDeColumna(hoja_maestro,"Martes",1)
  const col_miercoles= obtenerNumeroDeColumna(hoja_maestro,"Miercoles",1)
  const col_jueves= obtenerNumeroDeColumna(hoja_maestro,"Jueves",1)
  const col_viernes= obtenerNumeroDeColumna(hoja_maestro,"Viernes",1)
  const col_dias=[col_lunes,col_martes,col_miercoles,col_jueves,col_viernes]
  const dias=["Lunes","Martes","Miercoles","Jueves","Viernes"]
  const encabezado=["AREA", "PLAN DE ESTUDIO","NRC", "LISTA CRUZADA", "MATERIA", "CURSO", "SECC."," TITULO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "INICIO", "FIN", "SALA", "TIPO DE REUNIÓN", "PROFESOR"]//largo 18
  const fecha_inicio_clases=pedirParametro_fechas("Introduce la fecha de INICIO de clases","dd/mm/aaaa")
  const fecha_fin_clases=pedirParametro_fechas("Introduce la fecha de FIN de clases","dd/mm/aaaa")
  const fecha_inicio_ayud=pedirParametro_fechas("Introduce la fecha de \INICIO de AYUD/LAB","en formato dd/mm/aaaa")
  const fecha_fin_ayud=pedirParametro_fechas("Introduce la fecha de FIN de AYUD/LAB","dd/mm/aaaa")
  const idSpreadsheet = id_hoja_programacion;
  const hoja_programacion = SpreadsheetApp.openById(idSpreadsheet).getSheetByName('DATOS MAESTRO');
  const data_programacion = hoja_programacion.getDataRange().getDisplayValues();
  const horario_ing=agregar_fechas_HORARIO_ING(rellenar_HORARIO_ING(data_maestro,col_dias,dias,hoja_maestro),fecha_inicio_clases,fecha_fin_clases,fecha_inicio_ayud,fecha_fin_ayud,data_programacion)

  console.log(horario_ing)
  horario_ing.unshift(encabezado)
  const evaluaciones=rellenar_evaluaciones_horarioING(data_maestro,hoja_maestro,encabezado_fechas)
  console.log(evaluaciones)
  const hoja_horario_ing=crear_hoja_nombre("HORARIO ING",hojasActuales)
  hoja_horario_ing.getDataRange().clear()
  escribirListaDeListas(hoja_horario_ing,horario_ing.concat(evaluaciones),1,1)






}
function crearCalendario(){
const fecha_inicio=pedirParametro_fechas("Introduce la fecha de INICIO de programacion","en formato dd/mm/aaaa")
const fecha_fin=pedirParametro_fechas("Introduce la fecha de FIN de programacion","en formato dd/mm/aaaa")
const listaFechas=generarFechasEntre(fecha_inicio, fecha_fin)
const dias_feriados=verificarFinDeSemana(listaFechas)
const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
const hoja_maestro = hojasActuales.getSheetByName("MAESTRO");
const columnainicio=obtenerNumeroDeColumna(hoja_maestro,"OBSERVACION",1)+2




escribirFechasEnFila(listaFechas,dias_feriados,1, columnainicio)






}
function quitar_permisos(){

  removePermissionsFromList(id_carpeta_archivos_cordinadores,personas_para_quitar_permisos)



}



function actualizarCatalogo(){

const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
const hoja_catalogo = hojasActuales.getSheetByName("CATALOGO");
const hoja_catalogo_antiguo = hojasActuales.getSheetByName("CATALOGO ANTIGUO");
const data_catalogo=hoja_catalogo.getDataRange().getDisplayValues()
const data_catalogo_antiguo=hoja_catalogo_antiguo.getDataRange().getDisplayValues()
const col_area_antiguo=obtenerNumeroDeColumna(hoja_catalogo_antiguo,"AREA",1)
const col_area_nuevo=obtenerNumeroDeColumna(hoja_catalogo,"AREA",1)//SE TUVO QUE CAMBIAR EL NOMBRE
const col_codigo_antiguo=obtenerNumeroDeColumna(hoja_catalogo_antiguo,"CODIGO",1)
const col_codigo_nuevo=obtenerNumeroDeColumna(hoja_catalogo,"CODIGO",1)
data_catalogo.forEach((curso)=>{
  let area = data_catalogo_antiguo.find((curso_antiguo)=>
  curso_antiguo[col_codigo_antiguo]==curso[col_codigo_nuevo])
  if(area){
  curso[col_area_nuevo]=area[col_area_antiguo]
  }


})
console.log(data_catalogo)
hoja_catalogo.getDataRange().clearContent()
escribirListaDeListas(hoja_catalogo,data_catalogo,1,1)





}
function anadirCuposTentativos(){
  const data = leerInscritosPorPeriodo();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('RAMOS INSCRITOS POR PERIODO');
  const sheetReprobaciones = spreadsheet.getSheetByName('HISTORICO REPROBACION');
  const sheetMaestro = spreadsheet.getSheetByName('MAESTRO');
  const sheetCatalogo = spreadsheet.getSheetByName('CATALOGO');
  const hojas_homologacion=["ICQ","ICC","ICI","ICE","ICO","ICA"]
  const existencia_discrepancia_codigos = procesarMultiplesHojasHomologas(hojas_homologacion)
  const numColumnaCurso = obtenerNumeroDeColumna(sheet, "CURSO", 1)
  const numColumnaMateria = obtenerNumeroDeColumna(sheet, "MATERIA", 1)
  const numColumnaResquisitos = obtenerNumeroDeColumna(sheetCatalogo, "Requisitos", 1)
  const numColumnaNombre = obtenerNumeroDeColumna(sheetCatalogo, "TITULO", 1)
  const numColumnaCodigo = obtenerNumeroDeColumna(sheetCatalogo, "CODIGO", 1)
  const numColumnaCodigoMaestro = obtenerNumeroDeColumna(sheetMaestro, "CODIGO", 1)
  const conteo = contarPorCodigo(data, numColumnaCurso,numColumnaMateria,existencia_discrepancia_codigos);

  const promediosReprobacion = calcularPromedioReprobacion(sheetReprobaciones, 2, 3)

  const datosCatalogo = obtenerDatosDesdeHoja("CATALOGO", 2)
  const datosMaestro = obtenerDatosDesdeHoja("MAESTRO", 2)
  const cupos = {};
  datosCatalogo.forEach(curso => {
    const codigoCurso = curso[numColumnaCodigo];
    const requisitos = obtenerCodigosDesdeTexto(
      curso[numColumnaResquisitos],
      datosCatalogo,
      numColumnaNombre,
      numColumnaCodigo
    );

    const inscritosCurso = conteo[codigoCurso] || 0;
    let tasaReprobCurso = promediosReprobacion[codigoCurso]; // debe ser 0–1
      if (!tasaReprobCurso&&existencia_discrepancia_codigos[codigoCurso]){
        existencia_discrepancia_codigos[codigoCurso].forEach((lc)=>{
          if (promediosReprobacion[lc[0]]){
            tasaReprobCurso=promediosReprobacion[lc[0]]
          }


        })
      }
      tasaReprobCurso=tasaReprobCurso||0
    const reprobadosCurso = inscritosCurso * tasaReprobCurso;
    let maxAprobados = 0;

    requisitos.forEach(requisito => {
      const inscritos = conteo[requisito] || 0;
      let tasaReprobacion = promediosReprobacion[requisito]; // debe ser 0–1
      if (!tasaReprobacion&&existencia_discrepancia_codigos[requisito]){
        existencia_discrepancia_codigos[requisito].forEach((lc)=>{
          if (promediosReprobacion[lc[0]]){
            tasaReprobacion=promediosReprobacion[lc[0]]
          }


        })
      }
      tasaReprobacion=tasaReprobacion||0
      const aprobados = inscritos * (1 - tasaReprobacion);
      maxAprobados = Math.max(maxAprobados, aprobados);
    });


    cupos[codigoCurso] = Math.ceil(maxAprobados + reprobadosCurso);
  });
  console.log("cupos",cupos)
  const seccionesPorCodigo=contarPorCodigo2(datosMaestro,numColumnaCodigoMaestro)
  console.log("secciones por codigo",seccionesPorCodigo)
  const colCupos = obtenerNumeroDeColumna(sheetMaestro, "CUPOS", 1);

  const filaInicioMaestro = 2;
  const lastRowMaestro = sheetMaestro.getLastRow();

  if (lastRowMaestro >= filaInicioMaestro) {
    const numFilas = lastRowMaestro - filaInicioMaestro + 1;

    // Leer todas las filas del maestro una sola vez
    const maestroData = sheetMaestro
      .getRange(filaInicioMaestro, 1, numFilas, sheetMaestro.getLastColumn())
      .getValues();

    const valoresCupos = [];
    const colores = [];

    maestroData.forEach(row => {
      const codigo = row[numColumnaCodigoMaestro];
      const cuposCurso = cupos[codigo] || 0;
      const secciones = seccionesPorCodigo[codigo] || 0;
      console.log("cantidad de secciones",secciones)
      if (!cuposCurso || !secciones) {
        // si no hay cupos o no hay secciones: dejar vacío y resaltar amarillo
        valoresCupos.push([""]);
        colores.push(["#FFF59D"]);
      } else {
        const cuposPorSeccion = Math.ceil(cuposCurso / secciones);
        valoresCupos.push([cuposPorSeccion]);
        colores.push(["#FFFFFF"]); // fondo blanco (o cambia por el color que uses por defecto)
      }
    });

    // Escribir en bloque la columna "CUPOS TENTATIVOS"
    const rangoCupos = sheetMaestro.getRange(
      filaInicioMaestro,
      colCupos + 1, // obtenerNumeroDeColumna devuelve 0-based
      numFilas,
      1
    );

    rangoCupos.setValues(valoresCupos);
    rangoCupos.setBackgrounds(colores);
  }
}
function anadirListasCruzadas(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCumplimiento = spreadsheet.getSheetByName('CUMPLIMIENTO DE MALLA');
  const sheetHomologacion = spreadsheet.getSheetByName('HOMOLOGACION');
  const sheetMaestro = spreadsheet.getSheetByName('MAESTRO');
  const sheetCatalogo = spreadsheet.getSheetByName('CATALOGO');
  const data_maestro = leerHojaSinEncabezado("MAESTRO")
  const numColumnaCodigoMaestro = obtenerNumeroDeColumna(sheetMaestro, "CODIGO", 1)
  const numColumnaListaCruzadaMaestro = obtenerNumeroDeColumna(sheetMaestro, "LC", 1)
  const numColumnaTituloMaestro = obtenerNumeroDeColumna(sheetMaestro, "TITULO", 1)
  const numColumnaCursoMandanteMaestro = obtenerNumeroDeColumna(sheetMaestro, "CURSO MANDANTE", 1)
  const numColumnaSeccionesMaestro = obtenerNumeroDeColumna(sheetMaestro, "SECCIONES", 1)
  const numColumnaCuposMaestro = obtenerNumeroDeColumna(sheetMaestro, "CUPOS", 1)
  const numColumnaMateriaMaestro = obtenerNumeroDeColumna(sheetMaestro, "MATERIA", 1)
  const numColumnaCursosMaestro = obtenerNumeroDeColumna(sheetMaestro, "CURSO", 1)
  const numColumnaLlaveMaestro = obtenerNumeroDeColumna(sheetMaestro, "LLAVE Código- sec", 1)
  const numColumnaNrcMaestro = obtenerNumeroDeColumna(sheetMaestro, "NRC", 1)
  const hojas_homologacion=["ICQ","ICC","ICI","ICE","ICO","ICA"]
  const existencia_discrepancia_codigos = procesarMultiplesHojasHomologas(hojas_homologacion)
  console.log(JSON.stringify(existencia_discrepancia_codigos, null, 2));

  const dataAvanzeDeMalla = leerHojaSinEncabezado("CUMPLIMIENTO DE MALLA")
  const numColumnareglaCumplimiento = obtenerNumeroDeColumna(sheetCumplimiento, "REGLA", 1)
  const numColumnaNotaCumplimiento = obtenerNumeroDeColumna(sheetCumplimiento, "NOTA", 1)
  const codigos_pendientes = obtenerListaUnicos(dataAvanzeDeMalla,numColumnareglaCumplimiento,numColumnaNotaCumplimiento)
  console.log("codigos pendientes:", codigos_pendientes)
  const listas_cruzadas = []
  let contador_secciones_por_codigo = {}

data_maestro.forEach((seccion)=>{

  let codigo = seccion[numColumnaCodigoMaestro]

  if (!Object.keys(existencia_discrepancia_codigos).includes(codigo)){
    return
  }

  existencia_discrepancia_codigos[codigo].forEach((curso)=>{

    if(!codigos_pendientes.includes(curso[0])){
      return
    }

    // Inicializar contador si no existe
    if(!contador_secciones_por_codigo[codigo]){
      contador_secciones_por_codigo[codigo] = 1
    } else {
      contador_secciones_por_codigo[codigo]++
    }

    let numero_seccion = contador_secciones_por_codigo[codigo]

    let lista_cruzada = [...seccion]

    lista_cruzada[numColumnaCursoMandanteMaestro] = "NO"
    lista_cruzada[numColumnaTituloMaestro] = curso[1]
    lista_cruzada[numColumnaCodigoMaestro] = curso[0]

    lista_cruzada[numColumnaListaCruzadaMaestro] =
      `IN${String(listas_cruzadas.length).padStart(2, '0')}`

    seccion[numColumnaListaCruzadaMaestro] =
      `IN${String(listas_cruzadas.length).padStart(2, '0')}`

    // 👇 Ahora la sección es incremental
    lista_cruzada[numColumnaSeccionesMaestro] = numero_seccion

    lista_cruzada[numColumnaMateriaMaestro] = curso[0].slice(0,3)
    lista_cruzada[numColumnaCursosMaestro] = curso[0].slice(3)

    // 👇 Llave ahora es codigo + numero_seccion
    lista_cruzada[numColumnaLlaveMaestro] =
      `${curso[0]}${numero_seccion}`

    lista_cruzada[numColumnaNrcMaestro] = ""
    lista_cruzada[numColumnaCuposMaestro] = ""

    listas_cruzadas.push(lista_cruzada)
  })
})
  console.log(listas_cruzadas)
  const nueva_data_maestro = data_maestro.concat(listas_cruzadas)
  escribirListaDeListasDesdeFila(nueva_data_maestro,"MAESTRO",2)



}


function crearHojasDeRespuesta(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sheetRespuestas = spreadsheet.getSheetByName('RESPUESTAS')
  const sheetPreferencias = spreadsheet.getSheetByName('PREFERENCIAS')
  const sheetOtros = spreadsheet.getSheetByName('OTROS')
  const sheetEntregados = spreadsheet.getSheetByName('ENTREGADOS')




  escribirEncabezado(recrearHoja("RESPUESTAS"),["Fecha Entrega","Nombre","Rut","Dia","Horario entrega"])
  escribirEncabezado(recrearHoja("PREFERENCIAS"),["RUT","CodigoCurso","Secciones","2+1 o 3 seguidas","Horario entrega"])
  escribirEncabezado(recrearHoja("OTROS"),["RUT","CodigoCurso","Secciones","Comentarios","EXAMEN (Sí o No)","Cantidad de evaluaciones","Horario de evaluaciones","Metodologia","Fechas aproximadas pruebas/examen","Horario entrega"])
  escribirEncabezado(recrearHoja("ENTREGADOS"),["Rut"])




}
function borrarFilasPorRutConInput() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Eliminar registros por RUT",
    "Ingrese el RUT (ej: 12.345.678-9)",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const rut = response.getResponseText().trim();
  if (!rut) return;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const hojas = [
    { sheet: spreadsheet.getSheetByName('RESPUESTAS'), colRut: 3 },
    { sheet: spreadsheet.getSheetByName('PREFERENCIAS'), colRut: 1 },
    { sheet: spreadsheet.getSheetByName('OTROS'), colRut: 1 },
    { sheet: spreadsheet.getSheetByName('ENTREGADOS'), colRut: 1 }
  ];

  hojas.forEach(({ sheet, colRut }) => {
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet
      .getRange(2, colRut, lastRow - 1, 1)
      .getValues();

    for (let i = data.length - 1; i >= 0; i--) {
      const rutCelda = data[i][0];
      if (rutCelda && String(rutCelda).trim() === rut) {
        sheet.deleteRow(i + 2);
      }
    }
  });
  reenviarFormularioPorRut(rut)
  ui.alert(`Se eliminaron todos los registros asociados al RUT: ${rut}`);
}
function enviarRecordatoriosNoRespondidos() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojasRespuestas = [
    { name: "RESPUESTAS", colRut: 3 },
    { name: "PREFERENCIAS", colRut: 1 },
    { name: "OTROS", colRut: 1 },
    { name: "ENTREGADOS", colRut: 1 }
  ];

  const sheetMaestro = ss.getSheetByName("MAESTRO");
  const sheetProfesores = ss.getSheetByName("PROFESORES");

  const rut_profesores_jornada = sheetProfesores
    .getDataRange()
    .getDisplayValues()
    .map(fila => fila[1]);

  // Solo se le recuerda a quien YA fue notificado con el formulario inicial
  // (hoja FORMULARIOS ENVIADOS, la misma que usa enviarFormulariosCondicional).
  // Antes esta función recorría TODO MAESTRO sin exigir esto, así que podía
  // "recordar" un formulario a alguien que nunca lo recibió.
  const hojaEnviados = obtenerOcrearHojaFormulariosEnviados_(ss);
  const rutsNotificados = new Set(leerRutsFormulariosEnviados_(hojaEnviados));

  // SET de profesores que ya respondieron
  const rutsRespondidos = new Set();

  hojasRespuestas.forEach(({ name, colRut }) => {

    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet
      .getRange(2, colRut, lastRow - 1, 1)
      .getDisplayValues();

    data.forEach(fila => {
      const rut = String(fila[0]).trim();
      if (rut) rutsRespondidos.add(rut);
    });

  });

  // Columnas maestro
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
  data.shift();

  const yaEnviados = new Set();
  const exitosos = new Set();
  const fallidos = []; // { rut, email, motivo }
  const sinFormularioInicial = new Set(); // informativo: email en MAESTRO pero nunca se les envió el formulario

  data.forEach(fila => {

    const email1 = fila[colMail1];
    const email2 = fila[colMail2];
    const rut1 = String(fila[colRut1]).trim();
    const rut2 = String(fila[colRut2]).trim();
    const nombre1 = colNombre1 !== -1 ? String(fila[colNombre1] || "").trim() : "";
    const nombre2 = colNombre2 !== -1 ? String(fila[colNombre2] || "").trim() : "";

    [{ rut: rut1, email: email1, nombre: nombre1 }, { rut: rut2, email: email2, nombre: nombre2 }].forEach(({ rut, email, nombre }) => {
      if (!rut || !email) return;
      if (rutsRespondidos.has(rut)) return;
      if (yaEnviados.has(rut)) return;

      if (!rutsNotificados.has(rut)) {
        // Nunca se le envió el formulario inicial: no tiene sentido "recordárselo".
        sinFormularioInicial.add(rut);
        return;
      }

      const esJornada = condicionPersonalizada(rut, rut_profesores_jornada);
      const tipoCorreo = esJornada ? "RECORDATORIO_PREFERENCIAS" : "RECORDATORIO_DISPONIBILIDAD";
      const link = esJornada ? linkFormulario2 : linkFormulario1;

      yaEnviados.add(rut);

      // FIX: antes un solo correo inválido (o cualquier error de enviarCorreo,
      // por ejemplo que falte el TIPO en TEXTOS CORREO) lanzaba una excepción
      // no controlada que cortaba el resto del envío. Ahora se captura por
      // persona, se deja registro en LOG RECORDATORIOS / RESUMEN RECORDATORIOS,
      // y se sigue con el resto.
      try {
        enviarCorreo(email, rut, link, tipoCorreo);
        registrarRecordatorio(rut, nombre, email, tipoCorreo, "ENVIADO", "");
        exitosos.add(rut);
      } catch (e) {
        registrarRecordatorio(rut, nombre, email, tipoCorreo, "ERROR", e.message);
        fallidos.push({ rut, email, motivo: e.message });
      }
    });

  });

  let mensaje = `Recordatorios enviados correctamente a ${exitosos.size} profesor(es).`;

  if (fallidos.length > 0) {
    mensaje += `\n\n${fallidos.length} correo(s) NO se pudieron enviar:\n` +
      fallidos.map(f => `- RUT ${f.rut} (${f.email || "sin correo"}): ${f.motivo}`).join("\n") +
      `\n\nRevisa la hoja "${HOJA_LOG_RECORDATORIOS}" para el detalle.`;
  }

  if (sinFormularioInicial.size > 0) {
    mensaje += `\n\n${sinFormularioInicial.size} profesor(es) con correo en MAESTRO nunca recibieron el formulario inicial (no aparecen en "FORMULARIOS ENVIADOS"), así que no se les envió recordatorio. Usa "Envío de formularios Jornada y honorarios" primero si corresponde.`;
  }

  SpreadsheetApp.getUi().alert(mensaje);

}

function agregarDropdownUso(){

  const nombreColumna = "USO";
  const opciones = ["CLASE","AYUD","LAB"];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MAESTRO");

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  const headers = sheet.getRange(1,1,1,lastColumn).getValues()[0];

  // encontrar índice de la columna
  const col = headers.indexOf(nombreColumna);

  if (col === -1) {
    throw new Error("No se encontró la columna: " + nombreColumna);
  }

  const colSheet = col + 1; // convertir a 1-based para Sheets

  // crear validación
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();

  // aplicar validación a toda la columna (sin encabezado)
  if (lastRow > 1) {
    sheet
      .getRange(2, colSheet, lastRow - 1, 1)
      .setDataValidation(rule);
  }

}



function replicarInfoMandante() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("MAESTRO");
  const data = hoja.getDataRange().getValues();
  const headers = data[0];
  data.shift();

  const colLC = headers.indexOf("LC");
  const colMandante = headers.indexOf("CURSO MANDANTE");
  if (colLC === -1 || colMandante === -1) {
    throw new Error("No se encontraron las columnas LC o CURSO MANDANTE en MAESTRO");
  }

  const nombreNoCopiar = [
    "NRC", "SECCIONES", "CODIGO", "CURSO", "MATERIA", "AREA",
    "LLAVE Código- sec", "LC", "CURSO MANDANTE", "TITULO",
    "CUPOS", "PLAN DE ESTUDIO"
  ];
  const indicesNoCopiar = nombreNoCopiar
    .map(name => headers.indexOf(name))
    .filter(idx => idx !== -1);

  const grupos = {};
  data.forEach((fila, i) => {
    const lc = fila[colLC];
    if (!lc) return;
    if (!grupos[lc]) grupos[lc] = [];
    grupos[lc].push(i);
  });

  Object.values(grupos).forEach(indices => {
    const mandanteIdx = indices.find(i => data[i][colMandante] == "SI");
    if (mandanteIdx === undefined) return;

    indices.forEach(i => {
      if (data[i][colMandante] == "SI") return;
      for (let col = 0; col < data[i].length; col++) {
        if (!indicesNoCopiar.includes(col)) {
          data[i][col] = data[mandanteIdx][col];
        }
      }
    });
  });

  escribirDatosYResaltar(hoja, data);
}

function generarReporteAlertas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("MAESTRO");
  const data = hoja.getDataRange().getValues();
  const headers = data[0];
  data.shift();

  const idx = name => headers.indexOf(name);

  const colCodigo = idx("CODIGO");
  const colSeccion = idx("SECCIONES");
  const colRut1 = idx("RUT PROFESOR 1");
  const colRut2 = idx("RUT PROFESOR 2");
  const colSalaClas = idx("Sala especial CLAS");
  const colSalaAyud = idx("Sala especial AYUD");
  const colSalaLab = idx("Sala especial LAB");
  const colSalaPrueb = idx("Sala especial PRUEB");
  const colSalaExam = idx("Sala especial EXAM");
  const colCupos = idx("CUPOS");
  const colObs = idx("OBSERVACION");

  const diasNombres = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES"];
  const diasCol = diasNombres.map(name => idx(name));

  const fechaHeaders = headers.slice(colObs + 1).filter(h => h instanceof Date);
  const fechaCols = [];
  headers.forEach((h, i) => {
    if (i > colObs && h instanceof Date) fechaCols.push(i);
  });

  const diaSemana = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];

  function parsearBloques(celda, diaNombre) {
    if (!celda || String(celda).trim() === "") return [];
    return String(celda).split(",").map(part => {
      const p = part.trim().split(" ");
      if (p.length < 2) return null;
      const tipo = p[0] === "LAB/TALLER" ? "LABT" : p[0];
      const [inicio, fin] = p[1].split("-");
      return { tipo, inicio, fin, dia: diaNombre };
    }).filter(b => b !== null);
  }

  function parsearEvaluaciones(fila) {
    const evals = [];
    fechaCols.forEach((col, fi) => {
      const celda = fila[col];
      if (!celda || String(celda).trim() === "") return;
      const partes = String(celda).trim().split(/\s+/);
      let tipo, hora;
      if (partes.length === 1) {
        tipo = partes[0];
        hora = "19:30-21:20";
      } else if (partes.length === 2) {
        tipo = partes[0];
        hora = partes[1];
      } else {
        tipo = partes[1];
        hora = partes.slice(2).join(" ");
      }
      if (tipo !== "EXAM" && tipo !== "T" && tipo !== "TARDE") return;
      const diaSemanaNombre = diaSemana[fechaHeaders[fi].getDay()];
      const [inicio, fin] = hora.split("-");
      evals.push({ tipo, inicio, fin, dia_sem: diaSemanaNombre, fecha: fechaHeaders[fi] });
    });
    return evals;
  }

  function parsearPruebasEnBloques(fila) {
    const pruebas = [];
    fechaCols.forEach((col, fi) => {
      const celda = fila[col];
      if (!celda || String(celda).trim() === "") return;
      const partes = String(celda).trim().split(/\s+/);
      let bloque, hora;
      if (partes.length === 1) {
        bloque = partes[0];
        hora = "19:30-21:20";
      } else if (partes.length === 2) {
        bloque = partes[0];
        hora = partes[1];
      } else {
        bloque = partes[1];
        hora = partes.slice(2).join(" ");
      }
      bloque = bloque === "LAB/TALLER" ? "LABT" : bloque;
      const diaSemanaNombre = diaSemana[fechaHeaders[fi].getDay()];
      const [inicio, fin] = hora.split("-");
      pruebas.push({ bloque, inicio, fin, dia_sem: diaSemanaNombre, fecha: fechaHeaders[fi] });
    });
    return pruebas;
  }

  function timeToMin(t) {
    const [h, m] = t.split(":").map(Number);
    return h * 60 + (m || 0);
  }

  function solapan(i1, f1, i2, f2) {
    return timeToMin(i1) < timeToMin(f2) && timeToMin(i2) < timeToMin(f1);
  }

  function mismaFecha(a, b) {
    return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }

  const cursos = data.map((fila) => {
    const bloques = [];
    diasCol.forEach((col, di) => {
      if (col === -1) return;
      bloques.push(...parsearBloques(fila[col], diasNombres[di]));
    });
    return {
      fila,
      codigo: fila[colCodigo],
      seccion: fila[colSeccion],
      rut1: fila[colRut1],
      rut2: fila[colRut2],
      bloques,
      evaluaciones: parsearEvaluaciones(fila),
      pruebas: parsearPruebasEnBloques(fila),
      cupos: fila[colCupos],
      salaClas: fila[colSalaClas],
      salaAyud: fila[colSalaAyud],
      salaLab: fila[colSalaLab],
      salaPrueb: fila[colSalaPrueb],
      salaExam: fila[colSalaExam],
    };
  });

  const alertas = [];

  cursos.forEach(curso => {
    if (!curso.codigo) return;

    const salaPorBloque = {
      CLAS: curso.salaClas,
      AYUD: curso.salaAyud,
      LABT: curso.salaLab,
    };

    curso.pruebas.forEach(prueba => {
      const salaRequerida = prueba.bloque === "EXAM" ? curso.salaExam : curso.salaPrueb;
      if (!salaRequerida || String(salaRequerida).trim() === "") return;

      const salaBloque = salaPorBloque[prueba.bloque];
      if (salaBloque === undefined) return;

      const nombreRequerida = String(salaRequerida).trim();
      const nombreBloque = String(salaBloque || "").trim();
      const colSalaNombre = prueba.bloque === "EXAM" ? "Sala especial EXAM" : "Sala especial PRUEB";

      if (nombreBloque !== nombreRequerida) {
        alertas.push([
          "PRUEBA EN BLOQUE SIN SALA (quiebre DPSA)",
          `${curso.codigo}-${curso.seccion}`,
          curso.rut1 || curso.rut2 || "",
          prueba.dia_sem,
          `${prueba.inicio}-${prueba.fin}`,
          `Prueba programada en ${prueba.bloque} el ${prueba.dia_sem} ${prueba.inicio}-${prueba.fin} necesita sala ${nombreRequerida} (${colSalaNombre}) pero ${prueba.bloque} tiene ${nombreBloque ? "sala " + nombreBloque : "sin sala asignada"} → probable quiebre DPSA`
        ]);
      }
    });
  });

  cursos.forEach(curso => {
    if (!curso.codigo) return;
    const ruts = [curso.rut1, curso.rut2].filter(r => r && String(r).trim() !== "");
    ruts.forEach(rut => {
      const clasDeEste = curso.bloques.filter(b => b.tipo === "CLAS");
      const otrosCursos = cursos.filter(c =>
        c.codigo &&
        (c.rut1 === rut || c.rut2 === rut) &&
        (c.codigo !== curso.codigo || c.seccion !== curso.seccion)
      );

      clasDeEste.forEach(clas => {
        otrosCursos.forEach(otro => {
          if (otro.evaluaciones.length === 0) return;
          const ayudLabt = otro.bloques.filter(b =>
            (b.tipo === "AYUD" || b.tipo === "LABT") &&
            b.dia === clas.dia &&
            solapan(clas.inicio, clas.fin, b.inicio, b.fin)
          );
          ayudLabt.forEach(bloq => {
            alertas.push([
              "TOPE PROFESOR",
              `${otro.codigo}-${otro.seccion}`,
              rut,
              clas.dia,
              `${clas.inicio}-${clas.fin}`,
              `Profesor dicta CLAS en ${curso.codigo}-${curso.seccion} (${clas.dia} ${clas.inicio}-${clas.fin}) y su otro curso tiene ${bloq.tipo} con evaluación en el mismo horario`
            ]);
          });
        });
      });
    });
  });

  const conflictosEvalVistos = new Set();
  cursos.forEach(curso => {
    if (!curso.codigo || curso.evaluaciones.length === 0) return;
    const ruts = [curso.rut1, curso.rut2].filter(r => r && String(r).trim() !== "");
    ruts.forEach(rut => {
      const otrosCursos = cursos.filter(c =>
        c.codigo &&
        (c.rut1 === rut || c.rut2 === rut) &&
        (c.codigo !== curso.codigo || c.seccion !== curso.seccion)
      );
      curso.evaluaciones.forEach(evA => {
        otrosCursos.forEach(otro => {
          otro.evaluaciones.forEach(evB => {
            if (!mismaFecha(evA.fecha, evB.fecha)) return;
            if (!solapan(evA.inicio, evA.fin, evB.inicio, evB.fin)) return;
            const clave = [rut, `${curso.codigo}-${curso.seccion}`, `${otro.codigo}-${otro.seccion}`].sort().join("|");
            if (conflictosEvalVistos.has(clave)) return;
            conflictosEvalVistos.add(clave);
            alertas.push([
              "TOPE PROFESOR",
              `${otro.codigo}-${otro.seccion}`,
              rut,
              evA.dia_sem,
              `${evA.inicio}-${evA.fin}`,
              `Profesor tiene ${evA.tipo} en ${curso.codigo}-${curso.seccion} (${evA.dia_sem} ${evA.inicio}-${evA.fin}) y ${evB.tipo} en ${otro.codigo}-${otro.seccion} el mismo día y horario`
            ]);
          });
        });
      });
    });
  });

  cursos.forEach(curso => {
    if (!curso.codigo || curso.evaluaciones.length === 0) return;
    const ayudDias = curso.bloques.filter(b => b.tipo === "AYUD").map(b => b.dia);
    if (ayudDias.length === 0) return;

    curso.evaluaciones.forEach(ev => {
      if (ayudDias.includes(ev.dia_sem)) {
        alertas.push([
          "TARDE + AYUD mismo día",
          `${curso.codigo}-${curso.seccion}`,
          curso.rut1 || curso.rut2 || "",
          ev.dia_sem,
          `${ev.inicio}-${ev.fin}`,
          `Evaluación y AYUD el mismo día (${ev.dia_sem}) → la ayudantía se pierde`
        ]);
      }
    });
  });

  const hojaSalas = ss.getSheetByName("SALAS ESPECIALES");
  const mapaSalas = {};
  if (hojaSalas) {
    const salasData = hojaSalas.getDataRange().getValues();
    const salasHeaders = salasData[0];
    let colSalaNombre = salasHeaders.indexOf("NOMBRE BANNER");
    if (colSalaNombre === -1) colSalaNombre = salasHeaders.indexOf("SALAS ESPECIALES");
    const colSalaCap = salasHeaders.indexOf("CAPACIDAD");
    if (colSalaNombre !== -1 && colSalaCap !== -1) {
      for (let i = 1; i < salasData.length; i++) {
        const nombre = String(salasData[i][colSalaNombre]).trim();
        const cap = Number(salasData[i][colSalaCap]);
        if (nombre && !isNaN(cap)) {
          mapaSalas[nombre] = cap;
        }
      }
    }
  }

  cursos.forEach(curso => {
    if (!curso.codigo) return;
    const cuposCurso = Number(curso.cupos);
    if (!cuposCurso || isNaN(cuposCurso)) return;

    const salasUsadas = [
      { tipo: "CLAS", sala: curso.salaClas },
      { tipo: "AYUD", sala: curso.salaAyud },
      { tipo: "LAB", sala: curso.salaLab },
      { tipo: "PRUEB", sala: curso.salaPrueb },
      { tipo: "EXAM", sala: curso.salaExam },
    ].filter(s => s.sala && String(s.sala).trim() !== "");

    salasUsadas.forEach(({ tipo, sala }) => {
      const nombreSala = String(sala).trim();
      const capacidad = mapaSalas[nombreSala];
      if (capacidad !== undefined && cuposCurso > capacidad) {
        alertas.push([
          "SOBREPASO CAPACIDAD SALA",
          `${curso.codigo}-${curso.seccion}`,
          curso.rut1 || curso.rut2 || "",
          tipo,
          nombreSala,
          `Curso tiene ${cuposCurso} cupos pero sala "${nombreSala}" tiene capacidad ${capacidad} (sobrepaso de ${cuposCurso - capacidad} cupos)`
        ]);
      }
    });
  });

  const reporteSS = SpreadsheetApp.getActiveSpreadsheet();
  const existente = reporteSS.getSheetByName("REPORTE ALERTAS");
  if (existente) reporteSS.deleteSheet(existente);
  const nuevaHoja = reporteSS.insertSheet("REPORTE ALERTAS");

  const enc = ["TIPO ALERTA", "CURSO", "PROFESOR (RUT)", "DÍA", "HORARIO", "DETALLE"];
  const out = alertas.length ? [enc, ...alertas] : [enc, ["", "", "", "", "", "Sin alertas detectadas"]];
  nuevaHoja.getRange(1, 1, out.length, enc.length).setValues(out);
  nuevaHoja.setFrozenRows(1);
}

function doGet(e) {
  try {
    return router(e, "GET");
  } catch (err) {
    return json({ error: err.message });
  }
}
function doPost(e) {
  try {
    return router(e, "POST");
  } catch (err) {
    return json({ error: err.message });
  }
}

function actualizarCuposDesdeArchivoExterno() {
  const ui = SpreadsheetApp.getUi();

  let hojaExterna;
  try {
    const archivoExterno = SpreadsheetApp.openById(id_archivo_cupos_externo);
    hojaExterna = archivoExterno.getSheetByName(nombre_hoja_cupos_externa);
  } catch (e) {
    ui.alert("No se pudo abrir el archivo externo. Verifica el ID configurado (id_archivo_cupos_externo) y que tengas acceso a él.\n\n" + e.message);
    return;
  }
  if (!hojaExterna) {
    ui.alert(`No se encontró una hoja llamada "${nombre_hoja_cupos_externa}" dentro del archivo externo. Verifica la variable nombre_hoja_cupos_externa.`);
    return;
  }

  const datosExternos = hojaExterna.getDataRange().getValues();
  if (datosExternos.length < 2) {
    ui.alert("La hoja externa no tiene datos.");
    return;
  }
  const encabezadosExternos = datosExternos[0];
  const colLlaveExterna = buscarColumnaFlexible(encabezadosExternos, ["LLAVE Código- sec", "LLAVE Código- sec ", "LLAVE", "LLAVE CODIGO SEC"]);
  const colCuposExterna = buscarColumnaFlexible(encabezadosExternos, ["CUPOS"]);

  if (colLlaveExterna === -1 || colCuposExterna === -1) {
    ui.alert('No se encontraron las columnas "LLAVE Código- sec" y/o "CUPOS" en la hoja externa.');
    return;
  }

  // Mapa LLAVE -> CUPOS desde el archivo externo
  const mapaCupos = new Map();
  for (let i = 1; i < datosExternos.length; i++) {
    const llave = String(datosExternos[i][colLlaveExterna] || "").trim();
    if (llave) mapaCupos.set(llave, datosExternos[i][colCuposExterna]);
  }

  const hojasActuales = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMaestro = hojasActuales.getSheetByName("MAESTRO");
  const datosMaestro = hojaMaestro.getDataRange().getValues();
  const encabezadosMaestro = datosMaestro[0];

  const colLlaveMaestro = obtenerIndiceObligatorio(encabezadosMaestro, ["LLAVE Código- sec", "LLAVE Código- sec ", "LLAVE", "LLAVE CODIGO SEC"]);
  const colCuposMaestro = obtenerIndiceObligatorio(encabezadosMaestro, ["CUPOS"]);

  const nuevosValores = [];
  const colores = [];
  let actualizados = 0;
  let sinCoincidencia = 0;

  for (let i = 1; i < datosMaestro.length; i++) {
    const llave = String(datosMaestro[i][colLlaveMaestro] || "").trim();

    if (llave && mapaCupos.has(llave)) {
      nuevosValores.push([mapaCupos.get(llave)]);
      colores.push(["#FFFFFF"]);
      actualizados++;
    } else {
      nuevosValores.push([datosMaestro[i][colCuposMaestro]]);
      colores.push(["#FFF59D"]);
      if (llave) sinCoincidencia++;
    }
  }

  const rango = hojaMaestro.getRange(2, colCuposMaestro + 1, nuevosValores.length, 1);
  rango.setValues(nuevosValores);
  rango.setBackgrounds(colores);

  ui.alert(
    "Actualización de CUPOS completada",
    `Filas actualizadas: ${actualizados}\nFilas sin coincidencia en el archivo externo (resaltadas en amarillo, valor sin cambios): ${sinCoincidencia}`,
    ui.ButtonSet.OK
  );
}

function obtener_columnas_especificas(data,cols) {
  return data.map((curso)=>{
    let curso_codigo=[]
    cols.forEach((col)=>{
      curso_codigo.push(curso[col])

    })
    return curso_codigo

  })
}

function numeroALetra(n) {
  if (n < 1 || n > 26) return null;
  return String.fromCharCode(64 + n);
}


function cortarColumnas(lista, limiteColumna) {
  
  return lista.map((sublista )=>{
    
      
      
      
      return sublista.slice(0, limiteColumna)}
);
}
const buscarNRCsUnicos = (codigoCurso, nrcs, usados,columnaCurso,columnaMateria,columnaNRC) => {
  console.log("codigo curso:",codigoCurso)
  return nrcs
    .filter((nrc) => `${nrc[columnaMateria]}${nrc[columnaCurso]}` == codigoCurso) // Filtra por código de curso
    .map((nrc) => nrc[columnaNRC]) // Extrae el NRC (posición 3)
    .filter((nrc) => {
      if (!usados.includes(nrc)) { // Si no se ha usado, lo asigna
        usados.push(nrc); // Marca como usado
        return true;
      }
      return false; // Si ya está usado, lo descarta
    });
};
const generarSecciones = (curso, nrcs,columnaCurso,columnaMateria,columnaNRC) => {
  let usados = []
  const [codigoCurso, , numSecciones] = curso;
  const nrcsAsociados = buscarNRCsUnicos(codigoCurso, nrcs, usados,columnaCurso,columnaMateria,columnaNRC);

  // Creamos las secciones usando un map sobre un rango de 0 a numSecciones
  return Array.from({ length: numSecciones }, (_, i) => [
    codigoCurso,
    `${i + 1}`,
    i < nrcsAsociados.length ? nrcsAsociados[i] : "sin NRC asociado",
  ]);
};

const asignarSecciones = (cursos, nrcs,columnaCurso,columnaMateria,columnaNRC) => {
  return cursos.flatMap((curso) => generarSecciones(curso, nrcs,columnaCurso,columnaMateria,columnaNRC));
};


function generar_Entradas(secciones,data_nrc,data_catalogo,hoja_catalogo,hoja_listado_nrc,hoja_presupuesto,hoja_maestro,procesado_presupuesto){
  //Obitene las columnas por nombre
  const col_SCT_Chile_maestro= obtenerNumeroDeColumna(hoja_maestro,"SCT-Chile",1)
  const col_Requisitos_maestro= obtenerNumeroDeColumna(hoja_maestro,"Requisitos",1)
  const col_Habilidades_transversales_maestro= obtenerNumeroDeColumna(hoja_maestro,"Habilidades Transversales",1)
  const col_SCT_Chile_catalogo= obtenerNumeroDeColumna(hoja_catalogo,"SCT-Chile",1)
  const col_Requisitos_catalogo= obtenerNumeroDeColumna(hoja_catalogo,"Requisitos",1)
  const col_codigo_catalogo = obtenerNumeroDeColumna(hoja_catalogo,"CODIGO",1)
  const col_Habilidades_catalogo= obtenerNumeroDeColumna(hoja_catalogo,"Habilidades Transversales",1)
  const col_codigo_maestro=obtenerNumeroDeColumna(hoja_maestro,"CURSO",1)
  const col_materia_maestro=obtenerNumeroDeColumna(hoja_maestro,"MATERIA",1)
  const col_area_catalogo=obtenerNumeroDeColumna(hoja_catalogo,"AREA",1)
  const col_titulo_catalogo=obtenerNumeroDeColumna(hoja_catalogo,"TITULO",1)
  const col_nrc_listado_nrc=obtenerNumeroDeColumna(hoja_listado_nrc,"NRC",1)
  const col_cupos_listado_nrc=obtenerNumeroDeColumna(hoja_listado_nrc,"CUPOS",1)
  console.log("data nrc:",data_nrc)
  return secciones.map((seccion,idx)=>{

    let entrada = new Array(obtenerCantidadDeColumnas(hoja_maestro))
    entrada[0]="SI"
    entrada[7]=seccion[0]//codigo ING1101
    entrada[13]=seccion[1]//seccion 1
    entrada[4]=seccion[2]// NRC 5670
    entrada[col_codigo_maestro]=seccion[0].substring(3)
    entrada[col_materia_maestro]=seccion[0].substring(0, 3)//ING

    let codigo_seccion = seccion[0]+seccion[1]//ING11011
    entrada[3]= codigo_seccion
    let datos_catalogo=data_catalogo.find((curso)=>curso[col_codigo_catalogo]===seccion[0])
    
    //Puede que no encuentre ningun dato asociado en catalogo
    if (datos_catalogo===undefined){
    datos_catalogo = Array(data_catalogo[0].length) 
    }
    
    entrada[1]=datos_catalogo[col_area_catalogo]
    entrada[10]=datos_catalogo[col_titulo_catalogo]
    
    let datos_nrc = data_nrc.find((curso)=>curso[col_nrc_listado_nrc]==seccion[2])//Busca por nrc
    if (datos_nrc===undefined){
    datos_nrc= Array(data_nrc[0].length) 
    }
    entrada[12]=datos_nrc[col_cupos_listado_nrc]//cupos de ese NRC si no encuentra devuelve undefined
    
    let detalles = extraerRango(datos_catalogo,obtenerNumeroDeColumna(hoja_catalogo,"Plan Común",1),obtenerNumeroDeColumna(hoja_catalogo,"Sala especial",1))
    
    entrada = reemplazarParteLista(entrada,16,35,detalles)
    entrada[col_SCT_Chile_maestro]=datos_catalogo[col_SCT_Chile_catalogo]
    entrada[col_Requisitos_maestro]=datos_catalogo[col_Requisitos_catalogo]
    entrada[col_Habilidades_transversales_maestro]=datos_catalogo[col_Habilidades_catalogo]


    return entrada


  })


}
function obtenerCantidadDeColumnas(hoja) {
  // Abre la hoja activa
 
  
  // Obtiene el rango de celdas con datos
  var rango = hoja.getDataRange();
  
  // Obtiene el número de columnas (en el rango con datos)
  var cantidadDeColumnas = rango.getNumColumns();
  
  // Muestra la cantidad de columnas
  return cantidadDeColumnas
}
function obtenerRangoColumnas(listaDeListas,inicio,final) {
  
  return columnasRango = listaDeListas.map(fila => fila.slice(inicio,final));

}
function extraerRango(lista, inicio, fin) {
  if (lista===undefined){
    return new Array(fin-inicio) 
  }
  return lista.slice(inicio, fin + 1); // `fin + 1` porque slice no incluye el último índice
}
function reemplazarParteLista(lista, inicio, fin, sublista) {
  // Elimina desde `inicio` hasta `fin` y agrega `sublista` en su lugar.
  lista.splice(inicio, fin - inicio + 1, ...sublista);
  return lista;
}
function obtener_entradas_area_especifica(area,entradas,col_area){

  const entradas_por_area=entradas.filter((entrada)=>entrada[col_area]==area)
  return entradas_por_area





}
function obtener_areas(data_maestro,col_area){
let revisados=[]
data_maestro.filter((tupla)=>{

  if(revisados.includes(tupla[col_area])||tupla[col_area]===''){
    return false

  }
  revisados.push(tupla[col_area])
  return true


})
return revisados


}
function mostrarHojaYEnviarArchivo(hojaVisible, usuarioPermitido) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const usuarioActual = Session.getEffectiveUser().getEmail();
  const hojas = libro.getSheets();

  if (usuarioActual === usuarioPermitido) {
    // Mostrar solo la hoja permitida
    hojas.forEach((hoja) => {
      
        hoja.showSheet(); // Mostrar la hoja permitida
      
    });

    Logger.log(`Usuario ${usuarioActual} puede ver solo la hoja: ${hojaVisible}`);

    // Asignar permisos de solo lectura al usuario permitido
    libro.addViewer(usuarioPermitido);

    // Enviar correo con enlace al archivo
    const enlaceArchivo = libro.getUrl();
    const asunto = `Acceso al archivo: ${libro.getName()}`;
    const mensaje = `Hola,\n\nSe te ha dado acceso al archivo de Google Sheets para modificar el maestro de tu area con permisos de solo lectura. Puedes accederlo desde el siguiente enlace:\n\n${enlaceArchivo}\n\nSaludos.`;

    GmailApp.sendEmail(usuarioPermitido, asunto, mensaje);

    Logger.log(`Correo enviado a ${usuarioPermitido} con el enlace al archivo.`);
  } else {
    // Ocultar todas las hojas para usuarios no autorizados
    hojas.forEach((hoja) => hoja.hideSheet());
    Logger.log(`Usuario ${usuarioActual} no autorizado. Hojas ocultas.`);
  }


}
function resaltarCambiosEnHoja(hoja, datosOriginales, filaInicio, columnasClave, columnaEliminar, listaCambiada) {
  // Crear un mapa para buscar elementos de manera eficiente en los datos originales
  const generarClave = (fila, columnas) => columnas.map((col) => fila[col]).join("|");
  const mapaOriginal = new Map(
    datosOriginales.map((fila) => [generarClave(fila, columnasClave), fila])
  );

  // Iterar sobre la lista cambiada para pintar directamente en la hoja
  listaCambiada.forEach((filaCambiada, idxCambiado) => {
    const clave = generarClave(filaCambiada, columnasClave);
    const valorEliminar = filaCambiada[columnaEliminar];
    const filaHoja = filaInicio + idxCambiado; // Índice en la hoja para la lista cambiada

    if (valorEliminar === -1) {
      // Destacar en rojo si el elemento es para eliminar
      if (mapaOriginal.has(clave)) {
        console.log("Eliminado:", filaCambiada);
        hoja.getRange(filaHoja, 1, 1, filaCambiada.length).setBackground("red");
      }
    } else if (mapaOriginal.has(clave)) {
      // Si existe en la original, revisar si hay cambios
      const filaOriginal = mapaOriginal.get(clave);
      let tieneCambios = false;

      filaCambiada.forEach((valor, idx) => {
        if (valor !== filaOriginal[idx]) {
          // Destacar celdas con diferencias en azul
          console.log("Diferencia:", filaCambiada, filaOriginal);
          hoja.getRange(filaHoja, idx + 1).setBackground("blue");
          tieneCambios = true;
        }
      });

      if (!tieneCambios) {
        // Si no hay cambios, no se hace nada (puedes personalizar este caso)
      }
    } else {
      // Es un nuevo elemento, destacar toda la fila en verde
      console.log("Nuevo:", filaCambiada);
      hoja.getRange(filaHoja, 1, 1, filaCambiada.length).setBackground("green");
    }
  });
}
/**
 * @deprecated Reemplazada por fusionarCambiosPorColumna(). Esta versión
 * comparaba y reemplazaba FILAS COMPLETAS por posición, asumiendo que
 * listaCambiada (la hoja del coordinador) tiene exactamente las mismas
 * columnas, en el mismo orden, que listaOriginal (MAESTRO). Esa suposición
 * se rompe en la práctica (columnas ocultas borradas, MAESTRO creció después
 * de generar la hoja del coordinador, etc.), y cuando se rompe:
 *   1. filaCambiada (más angosta) reemplazaba a filaOriginal completa,
 *      descartando en silencio las columnas de MAESTRO que no estaban en
 *      esa posición de la hoja del coordinador.
 *   2. La mezcla resultante de filas de distinto ancho hacía fallar
 *      escribirDatosYResaltar() con "El número de columnas de los datos no
 *      coincide con el del intervalo" — y como esa función borra el
 *      contenido de la hoja ANTES de escribir, el error dejaba MAESTRO
 *      vacío. Esto es lo que se reportó como pérdida de datos en MAESTRO al
 *      usar "Validar Cambios".
 * No se eliminó esta función por si algo externo la referenciaba, pero
 * validar_cambios() ya no la usa.
 */
function actualizarListaOriginal(listaOriginal, listaCambiada, columnaClave, columnaEliminar) {
  // Crea un mapa para buscar elementos de manera eficiente en la lista original
  const mapaOriginal = new Map(
    listaOriginal.map((fila) => [fila[columnaClave], fila])
  );

  // Lista para almacenar la lista actualizada
  const listaActualizada = [];

  // Procesa los elementos de la lista cambiada
  listaCambiada.forEach((filaCambiada) => {
    const clave = filaCambiada[columnaClave];
    const valorEliminar = filaCambiada[columnaEliminar];

    if (valorEliminar === -1) {
      // Si el valor en la columna de eliminación es -1, elimina de listaOriginal
      console.log("eliminar",clave)
      mapaOriginal.delete(clave);
    } else if (mapaOriginal.has(clave)) {
      // Si existe en el mapa original y no hay cambios, se mantiene
      const filaOriginal = mapaOriginal.get(clave);

      if (!filaCambiada.every((val, idx) => val === filaOriginal[idx])) {
        // Hay cambios, reemplaza la fila en el mapa
        console.log("actualizar")
        mapaOriginal.set(clave, filaCambiada);
      }
    } else {
      // Es un nuevo elemento, agrégalo al mapa
      console.log("insersion")
      mapaOriginal.set(clave, filaCambiada);
    }
  });

  // Reconstruye la lista a partir del mapa actualizado
  mapaOriginal.forEach((fila) => listaActualizada.push(fila));

  return listaActualizada;
}

/**
 * Fusiona los cambios de una hoja de coordinador hacia MAESTRO alineando
 * POR NOMBRE DE COLUMNA (no por posición). Es la base del fix de
 * validar_cambios(): ya no importa si la hoja del coordinador tiene menos
 * columnas que MAESTRO, más columnas, o en otro orden — solo se sobrescriben
 * en MAESTRO las columnas cuyo nombre de encabezado también existe en la
 * hoja del coordinador; todo lo demás en MAESTRO (incluidas columnas que esa
 * hoja ni siquiera trae) se conserva intacto. El resultado SIEMPRE tiene
 * exactamente el ancho de MAESTRO, así que nunca puede producir las filas de
 * ancho mixto que antes hacían fallar la escritura.
 *
 * @param {Array<Array>} dataOriginal       Filas de MAESTRO (sin encabezado).
 * @param {Array<string>} encabezadosOriginal Encabezados de MAESTRO.
 * @param {number} columnaClaveOriginal     Índice de la columna clave en MAESTRO.
 * @param {Array<Array>} dataCambiada       Filas de la hoja del coordinador (sin encabezado).
 * @param {Array<string>} encabezadosCambiada Encabezados de la hoja del coordinador.
 * @param {number} columnaClaveCambiada     Índice de la columna clave en la hoja del coordinador.
 * @param {number} columnaCupos             Índice de la columna CUPOS en la hoja del coordinador
 *                                           (un valor -1 ahí marca la fila para eliminar de MAESTRO).
 * @returns {{
 *   filas: Array<Array>,          Filas resultantes, listas para escribir en MAESTRO (mismo ancho).
 *   actualizadas: number,         Cuántas filas existentes de MAESTRO cambiaron.
 *   eliminadas: number,           Cuántas filas se eliminaron (CUPOS = -1).
 *   nuevas: number,               Cuántas claves de la hoja del coordinador no existían en MAESTRO.
 *   detalleNuevas: string[],      Esas claves nuevas, para mostrarlas en el resumen.
 *   columnasCoincidentes: number  Cuántas columnas en común se encontraron (0 = probablemente la
 *                                 hoja activa no es la de un coordinador real; no se debe escribir).
 * }}
 */
function fusionarCambiosPorColumna(
  dataOriginal, encabezadosOriginal, columnaClaveOriginal,
  dataCambiada, encabezadosCambiada, columnaClaveCambiada, columnaCupos
) {
  const anchoOriginal = encabezadosOriginal.length;

  // Columnas en común entre ambas hojas: índice en la hoja del coordinador -> índice en MAESTRO.
  const mapaColumnas = [];
  encabezadosCambiada.forEach((nombre, idxCambiada) => {
    if (idxCambiada === columnaClaveCambiada) return; // la clave no se "sobrescribe", se usa para emparejar
    const idxOriginal = buscarColumnaFlexible(encabezadosOriginal, [nombre]);
    if (idxOriginal !== -1) {
      mapaColumnas.push({ idxCambiada, idxOriginal });
    }
  });

  const mapaOriginal = new Map(
    dataOriginal.map((fila) => [fila[columnaClaveOriginal], fila.slice()])
  );

  let actualizadas = 0;
  let eliminadas = 0;
  let nuevas = 0;
  const detalleNuevas = [];

  dataCambiada.forEach((filaCambiada) => {
    const clave = filaCambiada[columnaClaveCambiada];
    if (clave === "" || clave == null) return; // fila vacía en la hoja del coordinador

    const marcaEliminar = filaCambiada[columnaCupos];

    if (marcaEliminar === -1) {
      if (mapaOriginal.has(clave)) {
        mapaOriginal.delete(clave);
        eliminadas++;
      }
      return;
    }

    if (mapaOriginal.has(clave)) {
      const filaOriginal = mapaOriginal.get(clave);
      let cambio = false;

      mapaColumnas.forEach(({ idxCambiada, idxOriginal }) => {
        const nuevoValor = filaCambiada[idxCambiada];
        if (filaOriginal[idxOriginal] !== nuevoValor) {
          filaOriginal[idxOriginal] = nuevoValor;
          cambio = true;
        }
      });

      if (cambio) {
        mapaOriginal.set(clave, filaOriginal);
        actualizadas++;
      }
    } else {
      // Clave que no existe en MAESTRO: se crea una fila del ANCHO DE MAESTRO
      // (nunca del ancho, más angosto, de la hoja del coordinador), llenando
      // solo las columnas que esa hoja efectivamente trae.
      const filaNueva = new Array(anchoOriginal).fill("");
      filaNueva[columnaClaveOriginal] = clave;
      mapaColumnas.forEach(({ idxCambiada, idxOriginal }) => {
        filaNueva[idxOriginal] = filaCambiada[idxCambiada];
      });
      mapaOriginal.set(clave, filaNueva);
      nuevas++;
      detalleNuevas.push(String(clave));
    }
  });

  return {
    filas: Array.from(mapaOriginal.values()),
    actualizadas,
    eliminadas,
    nuevas,
    detalleNuevas,
    columnasCoincidentes: mapaColumnas.length,
  };
}

function horariosPorRut(data){
const horariosPorRut = data.reduce((acc, [fecha, nombre, rut, dia, hora]) => {
   
    if (!acc[rut]) {
        acc[rut] = {
            rut,
            horarios: {
                Lunes: [],
                Martes: [],
                Miércoles: [],
                Jueves: [],
                Viernes: []
            }
        };
    }
    
    
    acc[rut].horarios[dia].push(hora);
    return acc;
}, {});


return Object.values(horariosPorRut).map(profesor => {
    const { rut, horarios } = profesor;
    const horariosConcatenados = Object.keys(horarios).reduce((acc, dia) => {
        
        acc[dia] = horarios[dia].join(",");
        return acc;
    }, {});

    return { rut, ...horariosConcatenados };
});




}
function interseccion_de_bloques(dia1,dia2){
  const lista1=dia1.split(",")
  const lista2=dia2.split(",")
  const set1 = new Set(lista1);
  const set2 = new Set(lista2);


  const interseccion = [...set1].filter(item => set2.has(item));

  return interseccion.join(",")


}
function verificarCasillasNoVacias(lista, indices) {
  return indices.some(indice => lista[indice] !== "");
}
function verificarCasillastodasVacias(lista, indices) {
  return indices.every(indice => lista[indice] === "");
}
function rellenar_dpsa(data_maestro,col_dias,dias,hoja_maestro,encabezado_completo){
  const col_codigo=obtenerNumeroDeColumna(hoja_maestro,"CODIGO",1)
  const col_NRC=obtenerNumeroDeColumna(hoja_maestro,"NRC",1)
  const col_materia=obtenerNumeroDeColumna(hoja_maestro,"MATERIA",1)
  const col_curso=obtenerNumeroDeColumna(hoja_maestro,"CURSO",1)
  const col_seccion=obtenerNumeroDeColumna(hoja_maestro,"SECCIONES",1)
  const col_titulo=obtenerNumeroDeColumna(hoja_maestro,"TITULO",1)
  const col_Lista_cruzada=obtenerNumeroDeColumna(hoja_maestro,"LC",1)
  const col_profesor1=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR 1",1)
  const col_profesor2=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR 2",1)
  const col_profesor_lab=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR LABT",1)
  const col_cupos=obtenerNumeroDeColumna(hoja_maestro,"CUPOS",1)
  const col_mandante=obtenerNumeroDeColumna(hoja_maestro,"CURSO MANDANTE",1)
  const col_conector_liga=obtenerNumeroDeColumna(hoja_maestro,"CONECTOR DE LIGA",1)
  const col_sala_clas=obtenerNumeroDeColumna(hoja_maestro,"Sala especial CLAS",1)
  const col_sala_ayud=obtenerNumeroDeColumna(hoja_maestro,"Sala especial AYUD",1)
  const col_sala_lab=obtenerNumeroDeColumna(hoja_maestro,"Sala especial LAB",1)
  const col_sala_prueb=obtenerNumeroDeColumna(hoja_maestro,"Sala especial PRUEB",1)
  const col_sala_exam=obtenerNumeroDeColumna(hoja_maestro,"Sala especial EXAM",1)
  const col_metodologia=obtenerNumeroDeColumna(hoja_maestro,"METODOLOGIA",1)
  const col_comentarios=obtenerNumeroDeColumna(hoja_maestro,"COMENTARIOS COORDINADOR",1)
  const col_observaciones=obtenerNumeroDeColumna(hoja_maestro,"OBSERVACION",1)

  const dPSA=[]
  const diasSemana = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
  const headersFecha = encabezado_completo.slice(col_observaciones)
  const tiposEspeciales = { ING1400: "CLSS", ING4106: "CLSS", ING6104: "CLSS", ING6100: "OLIN" };

  function sala_especial_por_tipo(tipo,entrada){
    switch(tipo){
      case "CLAS": return entrada[col_sala_clas];
      case "AYUD": return entrada[col_sala_ayud];
      case "LABT": return entrada[col_sala_lab];
      case "PRBA": return entrada[col_sala_prueb];
      case "EXAM": return entrada[col_sala_exam];
      default: return "";
    }
  }

  function base_entrada(entrada,noMandanteLC){
    let e=new Array(25).fill("")
    e[0]=entrada[col_NRC]
    e[1]=entrada[col_materia]
    e[2]=entrada[col_curso]
    e[3]=entrada[col_titulo]
    e[4]=entrada[col_seccion]
    e[5]=entrada[col_Lista_cruzada]
    e[6]=entrada[col_conector_liga]
    e[8]=entrada[col_codigo]=="ING6103"?"Internado":"1"
    e[9]=noMandanteLC?"NO":"SI"
    e[10]=entrada[col_cupos]
    e[7]=tiposEspeciales[entrada[col_codigo]]||""
    e[24]=entrada[col_metodologia]
    return e
  }

  data_maestro.forEach((entrada)=>{
    const codigo=entrada[col_codigo]
    const isMandante=entrada[col_mandante]=="SI"
    const hasLC=entrada[col_Lista_cruzada]
    const noMandanteLC=hasLC&&!isMandante
    let pusoEntrada=false

    // --- CLASS BLOCKS (LUNES-VIERNES) ---
    if(noMandanteLC){
      let nueva_entrada=base_entrada(entrada,true)
      dPSA.push(nueva_entrada)
      pusoEntrada=true
    }else{
      const tieneClas = col_dias.some(col => {
        return String(entrada[col]).split(",").some(b => {
          const t = b.trim().split(" ")[0];
          return t === "CLAS";
        });
      });

      col_dias.forEach((columna,idx)=>{
        const bloques=entrada[columna].split(",")
        bloques.forEach((tipo_hora)=>{
          if(!tipo_hora) return
          pusoEntrada=true

          let nueva_entrada=base_entrada(entrada,noMandanteLC)
          let tipo=tipo_hora.trim().split(" ")[0]
          let hora=tipo_hora.trim().split(" ")[1]
          tipo=tipo=="LAB/TALLER"?"LABT":tipo

          nueva_entrada[7]=codigo=="ING6104"?"PRACTICA":"TEORIA"
          nueva_entrada[11]=tipo
          nueva_entrada[14]=dias[idx].toUpperCase()
          nueva_entrada[15]=hora?hora.split("-")[0]:""
          nueva_entrada[16]=hora?hora.split("-")[1]:""
          nueva_entrada[17]=tipo=="CLAS"?"1":tipo=="AYUD"?"2":"3"
          nueva_entrada[18]=tipo=="CLAS"?entrada[col_profesor1]:tipo=="LABT"&&!tieneClas?entrada[col_profesor1]:""

          if(tipo=="CLAS"){
            let otros=[]
            if(entrada[col_profesor2]) otros.push(entrada[col_profesor2])
            nueva_entrada[19]=otros.join(",")
            if(entrada[col_profesor2]){
              nueva_entrada[23]=entrada[col_comentarios]
            }
          }else if(tipo=="LABT"){
            nueva_entrada[19]=entrada[col_profesor_lab]
            nueva_entrada[20]=entrada[col_sala_lab]
          }

          nueva_entrada[21]=sala_especial_por_tipo(tipo,entrada)
          nueva_entrada[22]=nueva_entrada[21]

          dPSA.push(nueva_entrada)
        })
      })
    }

    // --- TEST ENTRIES (TARDE + EXAM) ---
    if(noMandanteLC) return
    const testData=entrada.slice(col_observaciones)
    testData.forEach((cell,cellIdx)=>{
      if(!cell) return
      pusoEntrada=true
      let partes=cell.trim().split(/\s+/)
      let tipo=""
      let hora=""
      let dia=""

      if(partes.length==1){
        tipo=partes[0]
        hora="19:30-21:20"
      }else if(partes.length==2){
        tipo=partes[0]
        hora=partes[1]
      }else if(partes.length>=3){
        dia=partes[0]
        tipo=partes[1]
        hora=partes.slice(2).join(" ")
      }

      if(tipo!="EXAM"&&tipo!="T"&&tipo!="TARDE") return

      let nueva_entrada=base_entrada(entrada,noMandanteLC)
      nueva_entrada[7]=codigo=="ING6104"?"PRACTICA":"TEORIA"
      nueva_entrada[11]=tipo=="EXAM"?"EXAM":"PRBA"

      if(headersFecha[cellIdx]&&headersFecha[cellIdx] instanceof Date){
        let fechaStr=convertirFechaAString(headersFecha[cellIdx])
        nueva_entrada[12]=fechaStr
        nueva_entrada[13]=fechaStr
        nueva_entrada[14]=dia||diasSemana[headersFecha[cellIdx].getDay()]
      }

      nueva_entrada[15]=hora.split("-")[0]
      nueva_entrada[16]=hora.split("-")[1]

      let testTipo=tipo=="EXAM"?"EXAM":"PRBA"
      nueva_entrada[21]=sala_especial_por_tipo(testTipo,entrada)
      nueva_entrada[22]=nueva_entrada[21]

      dPSA.push(nueva_entrada)
    })

    if((isMandante||tiposEspeciales[codigo])&&!pusoEntrada){
      dPSA.push(base_entrada(entrada,false))
    }
  })

  return dPSA
}
function agregar_fechas(dpsa,fecha_inicio_clases,fecha_fin_clases,fecha_inicio_ayud,fecha_fin_ayud){
  return dpsa.map((entrada)=>{
    if(!entrada[17]){
      // test entries already have dates set
    }else if(entrada[18]){
      entrada[12]=fecha_inicio_clases
      entrada[13]=fecha_fin_clases
    }else{
      entrada[12]=fecha_inicio_ayud
      entrada[13]=fecha_fin_ayud
    }
    return entrada
  })
}
function rellenar_HORARIO_ING(data_maestro,col_dias,dias,hoja_maestro){
  const col_area=obtenerNumeroDeColumna(hoja_maestro,"AREA",1)
  const col_plan_estudio=obtenerNumeroDeColumna(hoja_maestro,"PLAN DE ESTUDIO",1)
  const col_codigo=obtenerNumeroDeColumna(hoja_maestro,"CODIGO",1)
  const col_NRC=obtenerNumeroDeColumna(hoja_maestro,"NRC",1)
  const col_materia=obtenerNumeroDeColumna(hoja_maestro,"MATERIA",1)
  const col_curso=obtenerNumeroDeColumna(hoja_maestro,"CURSO",1)
  const col_seccion=obtenerNumeroDeColumna(hoja_maestro,"SECCIONES",1)
  const col_titulo=obtenerNumeroDeColumna(hoja_maestro,"TITULO",1)
  const col_Lista_cruzada=obtenerNumeroDeColumna(hoja_maestro,"LC",1)
  const col_profesor1=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR 1",1)
  const col_nombre_profesor1=obtenerNumeroDeColumna(hoja_maestro,"NOMBRE PROFESOR BANNER 1 \n(PROFESOR PRINCIPAL SESIÓN 01)",1)
  const col_nombre_profesor2=obtenerNumeroDeColumna(hoja_maestro,"NOMBRE PROFESOR 2\n(2DO PROFESOR - SESIÓN 02)",1)
  const col_nombre_profesorlab=obtenerNumeroDeColumna(hoja_maestro,"PROFESOR LABT ",1)
  const col_mandante=obtenerNumeroDeColumna(hoja_maestro,"CURSO MANDANTE",1)
  const col_profesor2=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR 2",1)
  const col_profesor_lab=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR LABT",1)
  const col_cupos=obtenerNumeroDeColumna(hoja_maestro,"CUPOS",1)
  
  const dPSA=[]
  console.log("numero de columnas dias",col_dias)
  data_maestro.forEach((entrada)=>{
    if(entrada[col_Lista_cruzada]&&entrada[col_mandante]!="SI"){
      let curso_mandante=data_maestro.find((elemento)=>elemento[col_Lista_cruzada]==entrada[col_Lista_cruzada])
      col_dias.forEach((columna)=>{
        entrada[columna]=curso_mandante[columna]
      })
    }
    col_dias.forEach((columna,idx)=>{
      const bloques=entrada[columna].split(",")
      
      bloques.forEach((tipo_hora)=>{
        if(tipo_hora){
        let tipo = tipo_hora.trim().split(" ")[0]
        let hora = tipo_hora.trim().split(" ")[1]
        
        let nueva_entrada=new Array(18).fill("")
        nueva_entrada[0]=entrada[col_area]
        nueva_entrada[1]=entrada[col_plan_estudio]
        nueva_entrada[2]=entrada[col_NRC]
        nueva_entrada[3]=entrada[col_Lista_cruzada]
        nueva_entrada[4]=entrada[col_materia]
        nueva_entrada[5]=entrada[col_curso]
        nueva_entrada[6]=entrada[col_seccion]
        nueva_entrada[7]=entrada[col_titulo]
        if(tipo_hora){
          if(idx==0){
            nueva_entrada[8]=hora
          }
          if(idx==1){
            nueva_entrada[9]=hora
          }
          if(idx==2){
            nueva_entrada[10]=hora
          }
          if(idx==3){
            nueva_entrada[11]=hora
          }
          if(idx==4){
            nueva_entrada[12]=hora
          }
        
        nueva_entrada[13]=""
        nueva_entrada[14]=""
        nueva_entrada[15]=""
        nueva_entrada[16]=tipo=="LAB/TALLER"?"LABT":tipo
        let nombreprofesor=""
        if(tipo=="LAB/TALLER"){
          if(entrada[col_nombre_profesorlab]){
            nombreprofesor=entrada[col_nombre_profesorlab]
          }else{
            nombreprofesor=entrada[col_nombre_profesor1]
            if(entrada[col_nombre_profesor2]!=""){
              nombreprofesor=entrada[col_nombre_profesor1]+"/"+ entrada[col_nombre_profesor2]
            }


          }
        }else{
          nombreprofesor=entrada[col_nombre_profesor1]
          if(entrada[col_nombre_profesor2]!=""){
          nombreprofesor=entrada[col_nombre_profesor1]+"/"+ entrada[col_nombre_profesor2]
          }    
        }
        nueva_entrada[17]=nombreprofesor

        }
        dPSA.push(nueva_entrada)




      
    }
    
    })




    })




  })

  return dPSA
  
}
function rellenar_evaluaciones(data_maestro,hoja_maestro,encabezado){
  const col_codigo=obtenerNumeroDeColumna(hoja_maestro,"CODIGO",1)
  const col_NRC=obtenerNumeroDeColumna(hoja_maestro,"NRC",1)
  const col_materia=obtenerNumeroDeColumna(hoja_maestro,"MATERIA",1)
  const col_curso=obtenerNumeroDeColumna(hoja_maestro,"CURSO",1)
  const col_seccion=obtenerNumeroDeColumna(hoja_maestro,"SECCIONES",1)
  const col_titulo=obtenerNumeroDeColumna(hoja_maestro,"TITULO",1)
  const col_Lista_cruzada=obtenerNumeroDeColumna(hoja_maestro,"LC",1)
  const col_profesor1=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR 1",1)
  const col_profesor2=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR 2",1)
  const col_profesor_lab=obtenerNumeroDeColumna(hoja_maestro,"RUT PROFESOR LABT",1)
  const col_cupos=obtenerNumeroDeColumna(hoja_maestro,"CUPOS",1)
  const col_mandante=obtenerNumeroDeColumna(hoja_maestro,"CURSO MANDANTE",1)
  const col_observaciones=obtenerNumeroDeColumna(hoja_maestro,"OBSERVACION",1)
  const col_area=obtenerNumeroDeColumna(hoja_maestro,"AREA",1)
  const dPSA=[]
  const col_conector_liga=obtenerNumeroDeColumna(hoja_maestro,"CONECTOR DE LIGA",1)
  const diasSemana = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
  const dias = encabezado.slice(col_observaciones)
  data_maestro.forEach((entrada,idx)=>{
    let col_dias=entrada.slice(col_observaciones)
    col_dias.forEach((columna,idx)=>{
      if(columna){
        let salida=columna.trim().split(" ")
        let dia=''
        let tipo=''
        let hora=""
        if (salida.length==2){
        tipo = salida[0]
        hora = salida[1]
        }else if(salida.length==1){
        tipo ="T"
        hora = "19:30-21:20"
        }else{
        dia = salida[0]
        tipo = salida[1]
        hora = salida[2]
        }
        console.log("dia:",dia,"tipo:",tipo,"hora:",hora)
        let nueva_entrada=new Array(25).fill("")
        nueva_entrada[0]=entrada[col_NRC]
        nueva_entrada[1]=entrada[col_materia]
        nueva_entrada[2]=entrada[col_curso]
        nueva_entrada[3]=entrada[col_titulo]
        nueva_entrada[4]=entrada[col_seccion]
        nueva_entrada[5]=entrada[col_Lista_cruzada]
        if((entrada[col_Lista_cruzada]&&entrada[col_mandante]!="SI")){
          entrada[col_Lista_cruzada]=""}
        nueva_entrada[6]=entrada[col_conector_liga]
        nueva_entrada[7]="TEORIA"
        
        nueva_entrada[8]=entrada[col_codigo]=="ING6103"?"i":"1"
        nueva_entrada[9]=""
        nueva_entrada[10]=entrada[col_cupos]
        nueva_entrada[11]=tipo=="EXAM"?"EXAM":"PRBA"
        nueva_entrada[12]=convertirFechaAString(dias[idx])
        nueva_entrada[13]=convertirFechaAString(dias[idx])
        
        nueva_entrada[14]=diasSemana[dias[idx].getDay()]
        nueva_entrada[15]=hora.split("-")[0]
        nueva_entrada[16]=hora.split("-")[1]
        nueva_entrada[17]=''
        nueva_entrada[18]=""
        nueva_entrada[19]=""
        nueva_entrada[20]=""
        nueva_entrada[21]=""
        nueva_entrada[22]=""
        nueva_entrada[23]=""
        nueva_entrada[24]=tipo
        

        dPSA.push(nueva_entrada)




      
    



    }
    })




  })

  return dPSA.filter((entrada)=>entrada[24]=="EXAM"||entrada[24]=="TARDE")
  
}
function agregar_fechas_HORARIO_ING(horario_ing,fecha_inicio_clases,fecha_fin_clases,fecha_inicio_ayud,fecha_fin_ayud,data_programacion){
  console.log(data_programacion)
  return horario_ing.map((entrada)=>{
    if(!entrada[16]){}

    else if (entrada[16]=="CLAS"){
    entrada[13]=fecha_inicio_clases
    entrada[14]=fecha_fin_clases
    }
    else{
      entrada[13]=fecha_inicio_ayud
      entrada[14]=fecha_fin_ayud
      
    }
    return entrada



  })
}
function rellenar_evaluaciones_horarioING(data_maestro,hoja_maestro,encabezado){
  
  const col_NRC=obtenerNumeroDeColumna(hoja_maestro,"NRC",1)
  const col_materia=obtenerNumeroDeColumna(hoja_maestro,"MATERIA",1)
  const col_curso=obtenerNumeroDeColumna(hoja_maestro,"CURSO",1)
  const col_seccion=obtenerNumeroDeColumna(hoja_maestro,"SECCIONES",1)
  const col_titulo=obtenerNumeroDeColumna(hoja_maestro,"TITULO",1)
  const col_Lista_cruzada=obtenerNumeroDeColumna(hoja_maestro,"LC",1)
  const col_area=obtenerNumeroDeColumna(hoja_maestro,"AREA",1)
  const col_plan_estudio=obtenerNumeroDeColumna(hoja_maestro,"PLAN DE ESTUDIO",1)
  const col_mandante=obtenerNumeroDeColumna(hoja_maestro,"CURSO MANDANTE",1)
  const col_observaciones=obtenerNumeroDeColumna(hoja_maestro,"OBSERVACION",1)
  const col_nombre_profesor1=obtenerNumeroDeColumna(hoja_maestro,"NOMBRE PROFESOR BANNER 1 \n(PROFESOR PRINCIPAL SESIÓN 01)",1)
  const col_nombre_profesor2=obtenerNumeroDeColumna(hoja_maestro,"NOMBRE PROFESOR 2\n(2DO PROFESOR - SESIÓN 02)",1)
  const col_nombre_profesorlab=obtenerNumeroDeColumna(hoja_maestro,"PROFESOR LABT ",1)
  
  
  const dPSA=[]
  const diasSemana = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
  const dias = encabezado.slice(col_observaciones)
  data_maestro.forEach((entrada,idx)=>{
    let col_dias=entrada.slice(col_observaciones)
    col_dias.forEach((columna,idx)=>{
      if(columna){
        let salida=columna.trim().split(" ")
        let dia=''
        let tipo=''
        let hora=""
        if (salida.length==2){
        tipo = salida[0]
        hora = salida[1]
        }else if(salida.length==1){
        tipo ="T"
        hora = "19:30-21:20"
        }else{
        dia = salida[0]
        tipo = salida[1]
        hora = salida[2]
        }
        console.log("dia:",dia,"tipo:",tipo,"hora:",hora)
        let nueva_entrada=new Array(18).fill("")
        nueva_entrada[0]=entrada[col_area]
        nueva_entrada[1]=entrada[col_plan_estudio]
        nueva_entrada[2]=entrada[col_NRC]
        nueva_entrada[3]=entrada[col_Lista_cruzada]
        nueva_entrada[4]=entrada[col_materia]
        nueva_entrada[5]=entrada[col_curso]
        nueva_entrada[6]=entrada[col_seccion]
        nueva_entrada[7]=entrada[col_titulo]
        console.log("dia de la semana: ",dias[idx],typeof dias[idx])
        let dia_semana=diasSemana[dias[idx].getDay()]
         if(dia_semana=="LUNES"){
          nueva_entrada[8]=hora
        }if(dia_semana=="MARTES"){
          nueva_entrada[9]=hora
        }if(dia_semana=="MIERCOLES"){
          nueva_entrada[10]=hora
        }if(dia_semana=="JUEVES"){
          nueva_entrada[11]=hora
        }if(dia_semana=="VIERNES"){
          nueva_entrada[12]=hora
        }
        
        nueva_entrada[13]=convertirFechaAString(dias[idx])
        nueva_entrada[14]=convertirFechaAString(dias[idx])
        
        nueva_entrada[15]=""
        nueva_entrada[16]=tipo=="EXAM"?"EXAM":"PRBA"
        let nombreprofesor=""
        if(tipo=="LAB/TALLER"){
          if(entrada[col_nombre_profesorlab]){
            nombreprofesor=entrada[col_nombre_profesorlab]
          }else{
            nombreprofesor=entrada[col_nombre_profesor1]
            if(entrada[col_nombre_profesor2]!=""){
              nombreprofesor=entrada[col_nombre_profesor1]+"/"+ entrada[col_nombre_profesor2]
            }


          }
        }else{
          nombreprofesor=entrada[col_nombre_profesor1]
          if(entrada[col_nombre_profesor2]!=""){
          nombreprofesor=entrada[col_nombre_profesor1]+"/"+ entrada[col_nombre_profesor2]
          }    
        }
        nueva_entrada[17]=nombreprofesor
       
      if(entrada[col_Lista_cruzada]!=""){
        let listas_cruzadas=data_maestro.filter((elementos)=>elementos[col_Lista_cruzada]==entrada[col_Lista_cruzada]&&
        elementos[col_mandante]!="SI")
        listas_cruzadas.forEach((lC)=>{
          let entrada_cruzada=[...nueva_entrada]
          entrada_cruzada[2]=lC[col_NRC]
          entrada_cruzada[6]=lC[col_seccion]
          entrada_cruzada[1]=lC[col_plan_estudio]
          dPSA.push(entrada_cruzada)



        })
      }
      dPSA.push(nueva_entrada)




      
    



    }
    })




  })

  return dPSA
  
}
function generarFechasEntre(fInicial, fFinal) {
    const fechas = [];
    const [diaInicio, mesInicio, añoInicio] = fInicial.split('/').map(Number);
    const [diaFin, mesFin, añoFin] = fFinal.split('/').map(Number);

    const inicio = new Date(añoInicio, mesInicio - 1, diaInicio);
    const fin = new Date(añoFin, mesFin - 1, diaFin);

    while (inicio <= fin) {
        const dia = String(inicio.getDate()).padStart(2, '0');
        const mes = String(inicio.getMonth() + 1).padStart(2, '0');
        const año = inicio.getFullYear();
        fechas.push(`${dia}/${mes}/${año}`);
        inicio.setDate(inicio.getDate() + 1); // Incrementa el día
    }

    return fechas;
}
function verificarFinDeSemana(listaFechas) {
    return listaFechas.map(fecha => {
        const [dia, mes, año] = fecha.split('/').map(Number);
        const diaSemana = new Date(año, mes - 1, dia).getDay(); // Obtiene el día de la semana
        return diaSemana !== 0 && diaSemana !== 6; // true si no es fin de semana
    });
}
function compareLists(oldList, newList, keyColumns,enunciado) {
  // Convert both lists to maps: each key collects all its rows (array bucket)
  function listToMap(list, keyCols) {
    const map = new Map();
    list.forEach(row => {
      const key = keyCols.map(col => row[col]).join('||');
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(row);
    });
    return map;
  }

  function mismasColumnasNoClave(a, b, keyCols) {
    for (let i = 0; i < a.length; i++) {
      if (keyCols.includes(i)) continue;
      if (a[i] != b[i]) return false;
    }
    return true;
  }

  function columnasCambiadas(a, b, keyCols) {
    const changed = [];
    a.forEach((value, colIndex) => {
      if (!keyCols.includes(colIndex) && value != b[colIndex]) {
        changed.push([colIndex + 3, enunciado[colIndex + 2]]);
      }
    });
    return changed;
  }

  const oldMap = listToMap(oldList, keyColumns);
  const newMap = listToMap(newList, keyColumns);

  const differences = [];
  const consumidasNuevas = new Set();

  // Compare buckets: exact match first, then changed, leftovers are removed/added
  oldMap.forEach((oldRows, key) => {
    const newRows = newMap.get(key) || [];
    oldRows.forEach(oldRow => {
      const idxExacta = newRows.findIndex((row, j) =>
        !consumidasNuevas.has(row) && mismasColumnasNoClave(oldRow, row, keyColumns)
      );
      if (idxExacta !== -1) {
        const row = newRows[idxExacta];
        consumidasNuevas.add(row);
        differences.push(['',, ...row]);
        return;
      }

      const idxCambiada = newRows.findIndex((row, j) => !consumidasNuevas.has(row));
      if (idxCambiada !== -1) {
        const newRow = newRows[idxCambiada];
        consumidasNuevas.add(newRow);
        const changedColumns = columnasCambiadas(oldRow, newRow, keyColumns);
        differences.push([`CHANGED ${changedColumns.join('|')}`, getCurrentDateDayAndTime(), ...newRow]);
        return;
      }

      differences.push([`REMOVED `, getCurrentDateDayAndTime(), ...oldRow]);
    });
  });

  newMap.forEach(newRows => {
    newRows.forEach(newRow => {
      if (consumidasNuevas.has(newRow)) return;
      differences.push([`ADDED `, getCurrentDateDayAndTime(), ...newRow]);
    });
  });

  return differences;
}
function normalizarListaDPSA(lista, llavesKey, encabezado) {
  if (!lista || !lista.length) return [];
  const ancho = encabezado.length;
  const esEncabezado = (row) => JSON.stringify(row) === JSON.stringify(encabezado);
  const esVacia = (row) => row.every((c) => c === "" || c === null || c === undefined || String(c).trim() === "");
  const map = new Map();
  lista.forEach((fila) => {
    let row = fila.slice();
    while (row.length < ancho) row.push("");
    if (esEncabezado(row) || esVacia(row)) return;
    const key = llavesKey.map((col) => row[col]).join("||");
    const contenido = JSON.stringify(row);
    if (!map.has(key)) map.set(key, new Set());
    map.get(key).add(contenido);
  });
  const resultado = [];
  map.forEach((contenidos) => {
    contenidos.forEach((contenido) => resultado.push(JSON.parse(contenido)));
  });
  return resultado;
}
function getCurrentDateDayAndTime() {
  const now = new Date();
  const currentDate = now.toISOString().split('T')[0]; // YYYY-MM-DD
  const days = ['DOMINGO', 'LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO'];
  const currentDay = days[now.getDay()];
  const hours = now.getHours().toString().padStart(2, '0');
  const minutes = now.getMinutes().toString().padStart(2, '0');
  const seconds = now.getSeconds().toString().padStart(2, '0');
  const currentTime = `${hours}:${minutes}:${seconds}`;
  return `${currentDate}, ${currentTime}`;
}
function convertirFechaAString(fecha) {
  // Verificar si es una instancia válida de Date
  if (!(fecha instanceof Date) || isNaN(fecha)) {
    throw new Error("Entrada inválida: debe ser una fecha válida.");
  }

  // Obtener día, mes y año
  const dia = fecha.getDate().toString().padStart(2, '0');
  const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
  const anio = fecha.getFullYear();

  // Formatear la fecha como "dd/mm/aaaa"
  return `${dia}/${mes}/${anio}`;
}
function removeSpecificPermissions(fileId, emails) {
  // Obtén el archivo por su ID
  var file = DriveApp.getFileById(fileId);

  // Recorre la lista de correos y elimina permisos
  emails.forEach(function(email) {
    try {
      // Intenta eliminar como editor
      file.removeEditor(email);

      // Intenta eliminar como visualizador

      Logger.log("Permisos eliminados para: " + email);
    } catch (e) {
      Logger.log("No se pudieron eliminar permisos para: " + email + ". Error: " + e.message);
    }
  });
}
function removePermissionsFromList(folderId, lista_datos) {
  // Obtén la carpeta por su ID
  var folder = DriveApp.getFolderById(folderId);

  // Recorre cada elemento de la lista
  lista_datos.forEach(function(entry) {
    var email = entry.mail;
    var archivos = entry.archivos;

    // Para cada archivo en la lista del usuario
    archivos.forEach(function(fileName) {
      // Busca el archivo por su nombre dentro de la carpeta
      var files = folder.getFilesByName(fileName);

      if (files.hasNext()) {
        var file = files.next(); // Obtén el archivo

        try {
          // Quita permisos para este correo
          file.removeEditor(email); // Intenta eliminar como editor
          
          Logger.log("Permisos eliminados para " + email + " en el archivo: " + fileName);
        } catch (e) {
          Logger.log("No se pudo eliminar permisos para " + email + " en el archivo: " + fileName + ". Error: " + e.message);
        }
      } else {
        Logger.log("No se encontró el archivo: " + fileName + " en la carpeta para el usuario: " + entry.name);
      }
    });
  });
}
function parseCustomDate(dateString, year = new Date().getFullYear()) {
  
  const monthMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
  };

  const [day, monthText] = dateString.split('-');
  const month = monthMap[monthText.toLowerCase()];
  
  if (month === undefined) {
    throw new Error(`Invalid month: ${monthText}`);
  }

  return new Date(year, month, parseInt(day, 10));
}
function contarPorCodigo(data, numColumnaCurso,numColumnaMateria,homologaciones) {
  const conteo = {};

  data.forEach(row => {
    const codigo = row[numColumnaMateria]+row[numColumnaCurso];

    if (codigo === "" || codigo == null) return;

    conteo[codigo] = (conteo[codigo] || 0) + 1;
  });
  Object.keys(homologaciones).forEach((codigo)=>{
    console.log(codigo,homologaciones[codigo])
    if (!conteo[codigo]){
      conteo[codigo]=0
    }
    homologaciones[codigo].forEach((lc)=>{
      if (conteo[lc[0]]){
      conteo[codigo]=conteo[codigo]+conteo[lc[0]]}

    })
  })
  return conteo;
}
function contarPorCodigo2(data,numColumnaCodigo) {
  const conteo = {};

  data.forEach(row => {
    const codigo = row[numColumnaCodigo];

    if (codigo === "" || codigo == null) return;

    conteo[codigo] = (conteo[codigo] || 0) + 1;
  });

  return conteo;
}
function parsePercent(raw) {
    if (raw === null || raw === undefined) return null;
    if (typeof raw === 'number' && !isNaN(raw)) return raw;
    var s = String(raw).trim();
    if (s === "" || /^n\/?a$/i.test(s)) return null;
    // quitar %
    s = s.replace("%", "");
    // cambiar coma decimal a punto
    s = s.replace(",", ".");
    var v = parseFloat(s);
    return isNaN(v) ? null : v;
  }
function calcularPromedioReprobacion(sheet, filaEncabezados, filaInicioDatos) {
  // ========================
  // Helpers
  // ========================
  function parsePercent(raw) {
    if (raw === null || raw === undefined) return null;
    if (typeof raw === 'number') return raw;
    const s = String(raw).trim();
    if (s === "" || /^n\/?a$/i.test(s)) return null;
    return parseFloat(s.replace('%', '').replace(',', '.'));
  }

  // ========================
  // Obtener encabezados
  // ========================
  const encabezados = sheet
    .getRange(filaEncabezados, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  // Columnas base
  const colCodigo = obtenerNumeroDeColumna(sheet, "CODIGO", filaEncabezados);
  const colSem10 = obtenerNumeroDeColumna(sheet, "Secciones Sem 10", filaEncabezados);
  const colSem20 = obtenerNumeroDeColumna(sheet, "Secciones Sem 20", filaEncabezados);

  // ========================
  // Detectar columnas REPROBADOS %
  // ========================
  const periodCols = encabezados
    .map((h, i) => ({ header: h, index: i }))
    .filter(h =>
      typeof h.header === "string" &&
      h.header.startsWith("REPROBADOS %")
    )
    .map(h => {
      const period = h.header.replace("REPROBADOS %", "").trim(); // 202310
      return {
        col: h.index,
        period,
        sem: period.slice(-2)
      };
    })
    .sort((a, b) => parseInt(a.period) - parseInt(b.period));

  // ========================
  // Leer datos
  // ========================
  const lastRow = sheet.getLastRow();
  if (lastRow < filaInicioDatos) return {};

  const rows = sheet
    .getRange(
      filaInicioDatos,
      1,
      lastRow - filaInicioDatos + 1,
      sheet.getLastColumn()
    )
    .getValues();

  // ========================
  // Procesamiento
  // ========================
  const resultado = {};

  rows.forEach(row => {
    const codigo = row[colCodigo];
    if (!codigo) return;

    const dicta10 = Boolean(row[colSem10]);
    const dicta20 = Boolean(row[colSem20]);

    const valores = periodCols
      .map(p => ({
        period: p.period,
        sem: p.sem,
        value: parsePercent(row[p.col])
      }))
      .filter(v => v.value !== null);

    if (valores.length === 0) {
      resultado[codigo] = null;
      return;
    }

    const ultimoPeriodo = valores.reduce((a, b) =>
      parseInt(a.period) > parseInt(b.period) ? a : b
    );

    let semestreObjetivo;
    if (dicta10 && !dicta20) {
      semestreObjetivo = "10";
    } else if (dicta20 && !dicta10) {
      semestreObjetivo = "20";
    } else if (dicta10 && dicta20) {
      semestreObjetivo = ultimoPeriodo.sem === "10" ? "20" : "10";
    } else {
      semestreObjetivo = ultimoPeriodo.sem;
    }

    const seleccionados = valores
      .filter(v => v.sem === semestreObjetivo)
      .sort((a, b) => parseInt(b.period) - parseInt(a.period))
      .slice(0, 2);

    if (seleccionados.length === 0) {
      resultado[codigo] = null;
      return;
    }
    console.log(codigo,seleccionados)
    const promedio =
      seleccionados.reduce((s, v) => s + v.value, 0) /
      seleccionados.length;

    resultado[codigo] = Math.round(promedio * 100) / 100;
  });

  return resultado;
}
function obtenerCodigosDesdeTexto(textoCursos, data, colNombre, colCodigo) {
  // si viene vacío, null o no es string → []
  if (!textoCursos || typeof textoCursos !== "string" || textoCursos.trim() === "") {
    return [];
  }

  const cursos = textoCursos
    .split(",")
    .map(c => c.trim())
    .filter(c => c && !/\(\s*p\s*\)/i.test(c));

  const mapa = {};
  data.forEach(row => {
    const nombre = row[colNombre];
    const codigo = row[colCodigo];
    if (nombre && codigo) {
      mapa[String(nombre).trim().toLowerCase()] = codigo;
    }
  });

  return cursos
    .map(nombreCurso => mapa[nombreCurso.toLowerCase()])
    .filter(Boolean);
}
function obtenerListaUnicos(datos, columna,columnaNota) {
  const index = columna;
  const vistos = new Set();

  datos.forEach(fila => {
    const valor = fila[index].slice(-7);
    const nota = fila[columnaNota]

    if (valor !== "" && valor !== null && valor !== undefined && nota=="") {
      vistos.add(valor);
    }
  });

  return Array.from(vistos);
}
function procesarMultiplesHojasHomologas(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const seen = new Set();        // diccionario global de códigos ya vistos
  const resultado = {};          // primary -> [[codigo, nombre], ...]

  for (let i = 0; i < sheetNames.length; i++) {
    const sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 3 || lastCol < 1) continue;

    // desde fila 3
    const datos = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

    for (let r = 0; r < datos.length; r++) {
      const fila = datos[r];
      const primaryRaw = fila[0];
      if (!primaryRaw) continue;

      const primary = String(primaryRaw).trim();

      // si ya fue visto, saltamos la fila completa
      if (seen.has(primary)) continue;

      // lista temporal de cursos nuevos encontrados en esta fila
      const nuevosCursos = [];

      // recorremos pares C/D, E/F, ...
      for (let c = 2; c < lastCol; c += 2) {
        const codeRaw = fila[c];
        if (!codeRaw) continue;

        const code = String(codeRaw).trim();
        if (code === primary) continue;

        if (!seen.has(code)) {
          let nombre = "";
          if (c + 1 < fila.length && fila[c + 1] != null) {
            nombre = String(fila[c + 1]).trim();
          }

          seen.add(code);
          nuevosCursos.push([code, nombre]);
        }
      }

      // solo ahora marcamos el primary como visto
      // y lo guardamos si realmente hay cursos nuevos
      if (nuevosCursos.length > 0) {
        seen.add(primary);
        resultado[primary] = nuevosCursos;
      }
    }
  }

  return resultado;
}

function ocultarColumnasExcepto(hoja, columnasVisibles) {
  const ultimaColumna = hoja.getLastColumn();
  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0];

  encabezados.forEach((nombreColumna, index) => {
    if (!columnasVisibles.includes(nombreColumna)) {
      hoja.hideColumns(index + 1);
    }
  });
}

/* ===== Validación DPSA vs LISTADO NRC ACTUAL ===== */

function construirMapaColumnasDPSA(encabezados) {
  return {
    KEY: buscarColumnaFlexible(encabezados, ["LLAVE CODIGO SEC", "LLAVE CÓDIGO SEC", "LLAVE"]),
    NRC: buscarColumnaFlexible(encabezados, ["NRC"]),
    MATERIA: obtenerIndiceObligatorio(encabezados, ["MATERIA"]),
    CURSO: obtenerIndiceObligatorio(encabezados, ["CURSO"]),
    SECC: obtenerIndiceObligatorio(encabezados, ["SECC.", "SECC", "SECCION", "SECCIÓN"]),
    CUPOS: obtenerIndiceObligatorio(encabezados, ["CUPOS TOTALES NRC", "CUPOS"]),
    TIPO: obtenerIndiceObligatorio(encabezados, ["TIPO DE REUNIÓN", "TIPO DE REUNION", "TIPO"]),
    FINI: obtenerIndiceObligatorio(encabezados, ["FECHA INICIO"]),
    FFIN: obtenerIndiceObligatorio(encabezados, ["FECHA FIN"]),
    DIA: obtenerIndiceObligatorio(encabezados, ["DIA", "DÍA"]),
    HINI: obtenerIndiceObligatorio(encabezados, ["HORA INICIO"]),
    HFIN: obtenerIndiceObligatorio(encabezados, ["HORA FIN"]),
    PROF_PPAL: obtenerIndiceObligatorio(encabezados, ["PROFESOR PRINCIPAL"]),
    LAB: buscarColumnaFlexible(encabezados, ["LABORATORIO", "SALA"]),
    LCRUZ: obtenerIndiceObligatorio(encabezados, ["LISTA CRUZADA"]),
    LIGA: buscarColumnaFlexible(encabezados, ["CONECTOR DE LIGA", "IDENTIFICADOR LIGA"])
  };
}

function construirMapaColumnasListado(encabezados) {
  return {
    KEY: buscarColumnaFlexible(encabezados, ["LLAVE", "LLAVE CODIGO SEC", "LLAVE CÓDIGO SEC"]),
    NRC: buscarColumnaFlexible(encabezados, ["NRC"]),
    MATERIA: obtenerIndiceObligatorio(encabezados, ["MATERIA"]),
    CURSO: obtenerIndiceObligatorio(encabezados, ["CURSO"]),
    SECC: obtenerIndiceObligatorio(encabezados, ["SECC.", "SECC", "SECCION", "SECCIÓN"]),
    CUPOS: obtenerIndiceObligatorio(encabezados, ["CUPOS"]),
    TIPO: obtenerIndiceObligatorio(encabezados, ["TIPO"]),
    INI: obtenerIndiceObligatorio(encabezados, ["INICIO", "FECHA INICIO"]),
    FIN: obtenerIndiceObligatorio(encabezados, ["FIN", "FECHA FIN"]),
    SALA: buscarColumnaFlexible(encabezados, ["SALA", "LABORATORIO"]),
    RUTPROF: obtenerIndiceObligatorio(encabezados, ["RUT PROFESOR", "PROFESOR"]),
    LUN: obtenerIndiceObligatorio(encabezados, ["LUNES"]),
    MAR: obtenerIndiceObligatorio(encabezados, ["MARTES"]),
    MIE: obtenerIndiceObligatorio(encabezados, ["MIERCOLES", "MIÉRCOLES"]),
    JUE: obtenerIndiceObligatorio(encabezados, ["JUEVES"]),
    VIE: obtenerIndiceObligatorio(encabezados, ["VIERNES"]),
    SAB: buscarColumnaFlexible(encabezados, ["SABADO", "SÁBADO"])
  };
}

function construirMapaColumnasMaestro(encabezados) {
  return {
    LLAVE: obtenerIndiceObligatorio(encabezados, ["LLAVE Código- sec ", "LLAVE Código- sec", "LLAVE", "LLAVE CODIGO SEC"]),
    MANDANTE: obtenerIndiceObligatorio(encabezados, ["CURSO MANDANTE", "MANDANTE"]),
    CONECTOR: obtenerIndiceObligatorio(encabezados, ["CONECTOR DE LIGA"]),
    MATERIA: obtenerIndiceObligatorio(encabezados, ["MATERIA"]),
    CURSO: obtenerIndiceObligatorio(encabezados, ["CURSO"]),
    SECC: obtenerIndiceObligatorio(encabezados, ["SECCIONES", "SECC.", "SECC"]),
    LCRUZ: obtenerIndiceObligatorio(encabezados, ["LC"])
  };
}

function construirLlaveSeccion(materia, curso, secc) {
  const mat = normKeyPart(materia);
  const cur = normKeyPart(curso);
  const sec = normSecc(secc);

  if (!mat || !cur || !sec) return "";

  return normKey(`${mat}${cur}${sec}`);
}

function getDpsaKey(row, cD) {
  if (cD.KEY !== -1) {
    const key = normKey(row[cD.KEY]);
    if (key) return key;
  }

  return construirLlaveSeccion(row[cD.MATERIA], row[cD.CURSO], row[cD.SECC]);
}

function indexarListado(rows, cL) {
  const lIndex = new Map();

  rows.forEach(row => {
    const key = getListadoKey(row, cL);
    if (!key) return;

    if (!lIndex.has(key)) lIndex.set(key, []);
    lIndex.get(key).push(row);
  });

  return lIndex;
}

function indexarMaestro(rows, cM) {
  const mIndex = new Map();

  rows.forEach(row => {
    const key = construirLlaveSeccion(row[cM.MATERIA], row[cM.CURSO], row[cM.SECC]);

    if (!key) return;

    if (!mIndex.has(key)) mIndex.set(key, []);
    mIndex.get(key).push(row);
  });

  return mIndex;
}

function getMandanteDesdeMaestro(dRow, cD, mIndex, cM) {
  const key = construirLlaveSeccion(dRow[cD.MATERIA], dRow[cD.CURSO], dRow[cD.SECC]);
  const filas = mIndex.get(key);

  if (!filas || filas.length === 0) {
    return { mandante: "", conector: "" };
  }

  return {
    mandante: normStr(filas[0][cM.MANDANTE]),
    conector: toKey(filas[0][cM.CONECTOR])
  };
}

function getExpectedFromDpsa(row, cD) {
  const tipo = normTipo(row[cD.TIPO]);
  const dia = normDia(row[cD.DIA]);
  const hi = normalizarHora_safe(row[cD.HINI]);
  const hf = normalizarHora_safe(row[cD.HFIN]);
  const rango = (hi && hf) ? `${hi}-${hf}` : "";
  const ini = formatearFecha_safe(row[cD.FINI]);
  const fin = formatearFecha_safe(row[cD.FFIN]);
  const prof = getProfesorDpsa(row, cD);

  return { tipo, dia, rango, ini, fin, prof };
}

function buildScheduleSignature(obj) {
  if (!obj) return "";

  return [
    normTipo(obj.tipo),
    normDia(obj.dia),
    obj.rango || "",
    obj.ini || "",
    obj.fin || ""
  ].join("|");
}

function describeExpected(obj) {
  if (!obj) return "";

  return `${normTipo(obj.tipo)} ${obj.dia || ""} ${obj.rango || ""} ${obj.ini || ""}-${obj.fin || ""}`.trim();
}

function getListadoKey(row, cL) {
  if (cL.KEY !== -1) {
    const key = normKey(row[cL.KEY]);
    if (key) return key;
  }

  return construirLlaveSeccion(row[cL.MATERIA], row[cL.CURSO], row[cL.SECC]);
}

function scoreListadoCandidate(expected, lRow, cL) {
  if (!expected || !lRow) return -1;

  const tipoL = normTipo(lRow[cL.TIPO]);
  const iniL = formatearFecha_safe(lRow[cL.INI]);
  const finL = formatearFecha_safe(lRow[cL.FIN]);
  const diaRangoL = extractListadoDiaRango(lRow, cL, expected.dia);

  let score = 0;

  if (expected.tipo && tipoL === expected.tipo) score += 1000;
  if (expected.ini && eqLoose(expected.ini, iniL)) score += 250;
  if (expected.fin && eqLoose(expected.fin, finL)) score += 250;

  if (expected.dia && diaRangoL.dia === expected.dia) score += 150;
  if (expected.rango && diaRangoL.rango === expected.rango) score += 150;

  if (
    expected.tipo && tipoL === expected.tipo &&
    expected.ini && eqLoose(expected.ini, iniL) &&
    expected.fin && eqLoose(expected.fin, finL) &&
    expected.dia && diaRangoL.dia === expected.dia &&
    expected.rango && diaRangoL.rango === expected.rango
  ) {
    score += 10000;
  }

  return score;
}

function extractListadoDiaRango(lRow, cL, preferredDia) {
  const dayToIndex = getDayToIndex(cL);

  if (preferredDia && dayToIndex[preferredDia] != null) {
    const idx = dayToIndex[preferredDia];
    const parsed = parseRangeFromDayCell(String(lRow[idx] ?? "").trim());
    if (parsed) return { dia: preferredDia, rango: parsed };
  }

  for (const dia in dayToIndex) {
    const idx = dayToIndex[dia];
    if (idx == null || idx === -1) continue;

    const parsed = parseRangeFromDayCell(String(lRow[idx] ?? "").trim());
    if (parsed) return { dia, rango: parsed };
  }

  return { dia: "", rango: "" };
}

function getDayToIndex(cL) {
  const out = {
    "LUNES": cL.LUN,
    "MARTES": cL.MAR,
    "MIERCOLES": cL.MIE,
    "JUEVES": cL.JUE,
    "VIERNES": cL.VIE
  };

  if (cL.SAB !== -1) out["SABADO"] = cL.SAB;

  return out;
}

function getProfesorDpsa(row, cD) {
  return normRut(row[cD.PROF_PPAL]);
}

function encontrarMejorMatchListado(expected, lRows, cL) {
  if (!lRows || lRows.length === 0) return null;

  let candidates = lRows.filter(rw => normTipo(rw[cL.TIPO]) === normTipo(expected.tipo));
  if (candidates.length === 0) candidates = lRows.slice();

  let best = null;
  let bestScore = -1;

  for (const lr of candidates) {
    const score = scoreListadoCandidate(expected, lr, cL);
    if (score > bestScore) {
      bestScore = score;
      best = lr;
    }
  }

  return best || candidates[0] || null;
}

function compareStandard(dRow, lMatch, cD, cL) {
  const diffs = [];

  if (!lMatch) {
    diffs.push("No se encontró fila equivalente en Listado_NRC.");
    return diffs;
  }

  const expected = getExpectedFromDpsa(dRow, cD);

  const tipoL = normTipo(lMatch[cL.TIPO]);
  const iniL = formatearFecha_safe(lMatch[cL.INI]);
  const finL = formatearFecha_safe(lMatch[cL.FIN]);
  const cupL = toInt(lMatch[cL.CUPOS]);
  const profL = normRut(lMatch[cL.RUTPROF]);
  const diaRangoL = extractListadoDiaRango(lMatch, cL, expected.dia);

  const cupD = toInt(dRow[cD.CUPOS]);
  const profD = normRut(getProfesorDpsa(dRow, cD));

  if (expected.tipo && tipoL && expected.tipo !== tipoL) {
    diffs.push(`TIPO difiere (DPSA=${expected.tipo} vs Listado=${tipoL})`);
  }

  if (!eqLoose(expected.ini, iniL)) {
    diffs.push(`FECHA INICIO difiere (DPSA=${show(expected.ini)} vs Listado=${show(iniL)})`);
  }

  if (!eqLoose(expected.fin, finL)) {
    diffs.push(`FECHA FIN difiere (DPSA=${show(expected.fin)} vs Listado=${show(finL)})`);
  }

  if (cupD !== cupL) {
    diffs.push(`CUPOS difieren (DPSA=${safeShow(cupD)} vs Listado=${safeShow(cupL)})`);
  }

  if (profD && profL && profD !== profL) {
    diffs.push(`PROFESOR PRINCIPAL difiere (DPSA=${profD} vs Listado=${profL})`);
  }

  if (expected.dia && expected.rango) {
    if (expected.dia !== diaRangoL.dia || expected.rango !== diaRangoL.rango) {
      diffs.push(`Horario difiere (DPSA=${expected.dia} ${show(expected.rango) || "vacío"} vs Listado=${show(diaRangoL.dia) || "vacío"} ${show(diaRangoL.rango) || "vacío"})`);
    }
  }

  return diffs;
}

function compareListadoAgainstTemplate(lRow, tpl, cL) {
  const diffs = [];

  const tipoL = normTipo(lRow[cL.TIPO]);
  const iniL = formatearFecha_safe(lRow[cL.INI]);
  const finL = formatearFecha_safe(lRow[cL.FIN]);
  const profL = normRut(lRow[cL.RUTPROF]);
  const diaRangoL = extractListadoDiaRango(lRow, cL, tpl.dia);

  if (tpl.tipo && tipoL && normTipo(tpl.tipo) !== tipoL) {
    diffs.push(`TIPO (esperado=${tpl.tipo} vs Listado=${tipoL})`);
  }

  if (tpl.ini && !eqLoose(tpl.ini, iniL)) {
    diffs.push(`FECHA INICIO (esperado=${show(tpl.ini)} vs Listado=${show(iniL)})`);
  }

  if (tpl.fin && !eqLoose(tpl.fin, finL)) {
    diffs.push(`FECHA FIN (esperado=${show(tpl.fin)} vs Listado=${show(finL)})`);
  }

  if (tpl.dia && tpl.rango) {
    if (tpl.dia !== diaRangoL.dia || tpl.rango !== diaRangoL.rango) {
      diffs.push(`FRANJA (esperada ${tpl.dia} ${tpl.rango} vs Listado=${show(diaRangoL.dia) || "vacío"} ${show(diaRangoL.rango) || "vacío"})`);
    }
  }

  const profTpl = normRut(tpl.prof);

  if (profTpl && profL && profTpl !== profL) {
    diffs.push(`PROFESOR PRINCIPAL (esperado=${profTpl} vs Listado=${profL})`);
  }

  return diffs;
}

function evaluarFilaDPSA(dRow, contexto) {
  const RED = contexto.colores.RED;
  const GREEN = contexto.colores.GREEN;
  const cD = contexto.cD;
  const cL = contexto.cL;

  const matchKey = getDpsaKey(dRow, cD);
  const info = getMandanteDesdeMaestro(dRow, cD, contexto.mIndex, contexto.cM);
  const mand = info.mandante;
  const conector = info.conector;

  if (!matchKey) {
    return { mensaje: `Llave vacía en DPSA. No se pudo construir LLAVE CODIGO SEC ni MATERIA+CURSO+SECC.`, color: RED };
  }

  const lRows = contexto.lIndex.get(matchKey);

  if (!lRows || lRows.length === 0) {
    return { mensaje: `No se encontró match en Listado_NRC para llave ${matchKey}.`, color: RED };
  }

  const esLiga = mand !== "SI" && conector;

  if (!esLiga) {
    const expected = getExpectedFromDpsa(dRow, cD);
    const lMatch = encontrarMejorMatchListado(expected, lRows, cL);
    const diffs = compareStandard(dRow, lMatch, cD, cL);

    if (diffs.length === 0) {
      return { mensaje: "OK", color: GREEN };
    }

    return { mensaje: diffs.join(" | "), color: RED };
  }

  const templatesMap = contexto.plantillas.get(conector) || new Map();
  const expectedNo = getExpectedFromDpsa(dRow, cD);
  const preferredKey = buildScheduleSignature(expectedNo);

  let candidateTemplates = [];

  if (preferredKey && templatesMap.has(preferredKey)) {
    candidateTemplates = [templatesMap.get(preferredKey)];
  } else {
    candidateTemplates = Array.from(templatesMap.values());
  }

  if (candidateTemplates.length === 0) {
    return { mensaje: `CONECTOR DE LIGA ${conector}: no se pudo construir plantilla desde MANDANTE=SI.`, color: RED };
  }

  let matched = false;
  const cupD = toInt(dRow[cD.CUPOS]);
  const localDiffs = [];

  for (const tpl of candidateTemplates) {
    const lMatchNo = encontrarMejorMatchListado(tpl, lRows, cL);

    if (!lMatchNo) {
      localDiffs.push(`CONECTOR DE LIGA ${conector}: no se encontró programación en Listado_NRC que siga la plantilla (${describeExpected(tpl)}).`);
      continue;
    }

    const diffsTpl = compareListadoAgainstTemplate(lMatchNo, tpl, cL);
    const cupL = toInt(lMatchNo[cL.CUPOS]);

    if (cupD !== cupL) {
      diffsTpl.push(`CUPOS NO difieren (DPSA=${safeShow(cupD)} vs Listado=${safeShow(cupL)})`);
    }

    if (diffsTpl.length === 0) {
      matched = true;
      break;
    } else {
      localDiffs.push(`CONECTOR DE LIGA ${conector}: ${diffsTpl.join(" | ")}`);
    }
  }

  if (matched) {
    return { mensaje: "OK", color: GREEN };
  }

  return { mensaje: localDiffs.join(" || "), color: RED };
}

function construirPlantillasLC(rowsDPSA, cD, cM, mIndex) {
  const plantillas = new Map();
  const grupos = new Map();

  rowsDPSA.forEach((dRow, idx) => {
    const info = getMandanteDesdeMaestro(dRow, cD, mIndex, cM);
    const conector = info.conector;
    if (!conector) return;

    if (!grupos.has(conector)) grupos.set(conector, { si: [], no: [] });

    const g = grupos.get(conector);
    if (info.mandante === "SI") g.si.push(idx);
    else g.no.push(idx);
  });

  for (const [conector, grp] of grupos.entries()) {
    const baseIdxs = grp.si.length > 0 ? grp.si : grp.no;

    for (const idx of baseIdxs) {
      const row = rowsDPSA[idx];
      const tpl = getExpectedFromDpsa(row, cD);
      const clave = buildScheduleSignature(tpl);

      if (!plantillas.has(conector)) plantillas.set(conector, new Map());

      if (!plantillas.get(conector).has(clave)) {
        plantillas.get(conector).set(clave, tpl);
      }
    }
  }

  return plantillas;
}

function validarDPSAVsListado(rowsDPSA, rowsListado, rowsMaestro, cD, cL, cM) {
  const GREEN = "#C6EFCE";
  const RED = "#FFC7CE";

  const lIndex = indexarListado(rowsListado, cL);
  const mIndex = indexarMaestro(rowsMaestro, cM);
  const plantillas = construirPlantillasLC(rowsDPSA, cD, cM, mIndex);

  const contexto = {
    lIndex,
    mIndex,
    plantillas,
    cD,
    cL,
    cM,
    colores: { GREEN, RED }
  };

  const mensajes = [];
  const colores = [];

  rowsDPSA.forEach(dRow => {
    const { mensaje, color } = evaluarFilaDPSA(dRow, contexto);
    mensajes.push(mensaje);
    colores.push(color);
  });

  return { mensajes, colores };
}

/* ===== Normalizadores y helpers de validación ===== */

function toKey(v) {
  return (v === null || v === undefined) ? "" : String(v).trim();
}

function normStr(v) {
  return removeAccents(String(v === null || v === undefined ? "" : v))
    .trim()
    .toUpperCase();
}

function normKey(v) {
  return removeAccents(String(v === null || v === undefined ? "" : v))
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/-/g, "");
}

function normKeyPart(v) {
  return removeAccents(String(v === null || v === undefined ? "" : v))
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}

function normSecc(v) {
  const s = normKeyPart(v);

  if (!s) return "";

  if (/^\d+$/.test(s)) {
    return String(parseInt(s, 10));
  }

  return s;
}

function normTipo(v) {
  const s = normStr(v)
    .replace(/\s+/g, " ")
    .replace(/-/g, "/");

  if (s === "LAB/TALLER" || s === "LAB TALLER" || s === "LABORATORIO/TALLER") return "LABT";
  if (s === "LAB") return "LABT";
  if (s === "EXAMEN") return "EXAM";

  return s;
}

function eqLoose(a, b) {
  return normStr(a) === normStr(b);
}

function toInt(v) {
  if (v === null || v === undefined || String(v).trim() === "") return NaN;

  const n = Number(String(v).replace(/[^\d\-]/g, ""));

  return Number.isFinite(n) ? n : NaN;
}

function normRut(v) {
  return removeAccents(String(v === null || v === undefined ? "" : v))
    .trim()
    .toUpperCase()
    .replace(/\./g, "")
    .replace(/-/g, "")
    .replace(/\s+/g, "");
}

function normDia(v) {
  return normStr(v);
}

function formatearFecha_safe(fecha) {
  if (fecha === null || fecha === undefined) return "";

  const tz = Session.getScriptTimeZone() || "America/Santiago";

  if (Object.prototype.toString.call(fecha) === "[object Date]" && !isNaN(fecha)) {
    return Utilities.formatDate(fecha, tz, "dd/MM/yyyy");
  }

  if (typeof fecha === "number") {
    const base = new Date(1899, 11, 30);
    const ms = Math.round(fecha * 24 * 60 * 60 * 1000);
    const d = new Date(base.getTime() + ms);
    return Utilities.formatDate(d, tz, "dd/MM/yyyy");
  }

  const s = String(fecha).trim();
  if (!s) return "";

  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);

  if (m) {
    let [, dd, mm, yy] = m;
    dd = dd.padStart(2, "0");
    mm = mm.padStart(2, "0");

    if (yy.length === 2) {
      yy = parseInt(yy, 10) >= 70 ? `19${yy}` : `20${yy}`;
    }

    return `${dd}/${mm}/${yy}`;
  }

  const parsed = new Date(s);

  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, tz, "dd/MM/yyyy");
  }

  return "";
}

function normalizarHora_safe(valor) {
  if (valor === null || valor === undefined || valor === "") return "";

  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor)) {
    const h = valor.getHours().toString().padStart(2, "0");
    const m = valor.getMinutes().toString().padStart(2, "0");
    return `${h}:${m}`;
  }

  if (typeof valor === "number") {
    const totalMin = Math.round(valor * 24 * 60);
    const h = Math.floor(totalMin / 60).toString().padStart(2, "0");
    const min = (totalMin % 60).toString().padStart(2, "0");
    return `${h}:${min}`;
  }

  const s = String(valor).trim();

  const m = s.match(/^(\d{1,2}):(\d{1,2})(?::\d{1,2})?$/);

  if (m) {
    const hh = m[1].padStart(2, "0");
    const mm = m[2].padStart(2, "0");
    return `${hh}:${mm}`;
  }

  const d = new Date(s);

  if (!isNaN(d)) {
    const hh = d.getHours().toString().padStart(2, "0");
    const mm = d.getMinutes().toString().padStart(2, "0");
    return `${hh}:${mm}`;
  }

  return "";
}

function parseRangeFromDayCell(raw) {
  const s = String(raw || "").trim();

  if (!s) return "";

  let m = s.match(/^(\d{1,2}:\d{1,2})\s*-\s*(\d{1,2}:\d{1,2})$/);

  if (m) {
    const hi = normalizarHora_safe(m[1]);
    const hf = normalizarHora_safe(m[2]);

    return (hi && hf) ? `${hi}-${hf}` : "";
  }

  const parts = s.split("-");

  if (parts.length === 2) {
    const hi = normalizarHora_safe(parts[0]);
    const hf = normalizarHora_safe(parts[1]);

    return (hi && hf) ? `${hi}-${hf}` : "";
  }

  const one = normalizarHora_safe(s);

  return one ? `${one}-${one}` : "";
}

function show(v) {
  return (v === null || v === undefined) ? "" : String(v);
}

function safeShow(v) {
  return (v === null || v === undefined || v === "" || Number.isNaN(v)) ? "vacío" : String(v);
}

function router(e, method) {
  const params = e.parameter || {};
  const key = params.key;

  if (key !== API_KEY) {
    return json({ error: "No autorizado" });
  }

  const action = params.action;

  if (!action) {
    return json({ error: "acción requerida" });
  }

  // 👇 Parsear body si es POST
  let body = {};
  if (method === "POST" && e.postData && e.postData.contents) {
    body = JSON.parse(e.postData.contents);
  }

  switch (action) {
    case "ping":
      return json({ ok: true });

    case "maestro.listar":
      return apiMaestroListar(e);

    case "sheet.escribir":
      return apiSheetTest(body);
    
    case "pruebas.fechas":
      return apiPruebasFechas(body);
    
    default:
      return json({ error: "acción no válida" });
  }
}

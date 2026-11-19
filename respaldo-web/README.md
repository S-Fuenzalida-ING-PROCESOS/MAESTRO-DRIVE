# Respaldo web: evaluaciones y horarios (plan B)

Versiones en Apps Script del **calendario de evaluaciones** y de la **programación de horarios**. Leen y escriben directo en la hoja MAESTRO, sin base de datos. Sirven como plan B si la web app de Render (CalendarioHorariosWeb) deja de funcionar.

| Carpeta | Archivos | Qué hace |
|---|---|---|
| `evaluaciones/` | `CalendarioEvaluaciones.gs`, `CalendarioEvaluaciones.html` | Calendario de pruebas y exámenes del semestre |
| `horarios/` | `ProgramacionHorarios.gs`, `ProgramacionHorarios.html` | Grilla semanal de clases, ayudantías y laboratorios |

Versión respaldada: la que estaba implementada al 19-11-2026.

## Importante antes de implementar

- **Cada app va en su propio proyecto de Apps Script.** No se pegan en el proyecto de MAESTRO-DRIVE: `main.gs` ya tiene un `doGet` y chocaría con el de estas apps. Tampoco se juntan las dos en un mismo proyecto, porque cada una define su propio `doGet`.
- El archivo HTML debe llamarse **exactamente** `CalendarioEvaluaciones` o `ProgramacionHorarios` (sin `.html`; Apps Script lo agrega solo).
- Ambas leen la planilla del MAESTRO por ID, en `SOURCE_SPREADSHEET_ID` al inicio del `.gs` (`1cmytVU4wEEMPz-E2Q7lJozkynGJCLpOHcIVRaemh1Pg`). Si el MAESTRO cambia de archivo, actualizar ese ID en los dos `.gs`.
- Las columnas se ubican por nombre de encabezado, no por posición, así que agregar columnas auxiliares en MAESTRO no las afecta.

## Pasos para implementar (cada app por separado)

1. Ir a <https://script.google.com> → **Nuevo proyecto**. Nombrarlo, por ejemplo, "Calendario de Evaluaciones".
2. Reemplazar el contenido de `Código.gs` por el del `.gs` de la carpeta (se puede renombrar el archivo con el mismo nombre).
3. **Archivo → + → HTML**, nombrarlo `CalendarioEvaluaciones` (o `ProgramacionHorarios`) y pegar el contenido del `.html`.
4. Guardar y ejecutar una vez `cevObtenerDatos` (o `phrObtenerDatos`) desde el editor para autorizar el acceso a la planilla.
5. **Implementar → Nueva implementación → Aplicación web**:
   - Ejecutar como: **Yo**.
   - Quién tiene acceso: **Cualquier usuario de Universidad de los Andes** (dominio miuandes.cl).
6. Copiar la URL `/exec` y compartirla con el equipo.

Para publicar cambios posteriores sin cambiar la URL: **Implementar → Gestionar implementaciones → editar (lápiz) → Versión: Nueva versión**.

## Datos que viven en el proyecto, no en la planilla

Se guardan en las **propiedades del script** del proyecto. Un proyecto nuevo parte vacío:

- **Evaluaciones:** feriados y suspensiones parciales (propiedad `CEV_FERIADOS`). Hay que volver a marcarlos en el calendario, o copiar el valor de esa propiedad desde el proyecto antiguo (Configuración del proyecto → Propiedades del script).
- **Horarios:** avisos aceptados en la columna "Revisión" (propiedades `PHR_REV_<CODIGO>|<SECCION>`). Si no se copian, esos avisos vuelven a aparecer como pendientes.

## Opcional: abrir desde el menú de la planilla

Si el proyecto está vinculado a la planilla del MAESTRO, se puede agregar en `onOpen()`:

```js
.addItem('Calendario de evaluaciones', 'abrirCalendarioEvaluaciones')
.addItem('Programación de horarios', 'abrirProgramacionHorarios')
```

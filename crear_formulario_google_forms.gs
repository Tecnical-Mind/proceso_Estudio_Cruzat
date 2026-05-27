/**
 * ============================================================
 * RADIOGRAFÍA DEL HECHO – Creador automático de Google Forms
 * ============================================================
 * INSTRUCCIONES:
 *  1. Abrí Google Apps Script: script.google.com
 *  2. Creá un nuevo proyecto y pegá este código completo
 *  3. Hacé clic en "Ejecutar" → seleccioná "crearFormulario"
 *  4. Autorizá los permisos solicitados
 *  5. El formulario se creará en tu Google Drive automáticamente
 *  6. Los links aparecen en "Registros de ejecución" (abajo del editor)
 * ============================================================
 *
 * BUGS CORREGIDOS:
 *  - Eliminado addFileUploadItem() → no existe en Apps Script
 *  - Eliminado setProgressBar(), setShuffleQuestions(), setAllowResponseEdits()
 *  - Corregidas validaciones de texto (requireTextMatchesPattern en lugar de requireNumberBetween)
 *  - Eliminado SpreadsheetApp.getUi() → falla fuera del contexto de Sheets
 *  - Los links se muestran en Logger (Registros de ejecución)
 * ============================================================
 */

function crearFormulario() {

  // ── Crear formulario base ───────────────────────────────────
  var form = FormApp.create('Radiografía del Hecho – Registro de Accidente');

  form.setDescription(
    'Registro oficial de accidentes y hechos. ' +
    'Completá todos los campos obligatorios (*). ' +
    'Los datos son estrictamente confidenciales y se usan exclusivamente para la gestión del caso.'
  );
  form.setCollectEmail(true);
  form.setLimitOneResponsePerUser(false);
  form.setConfirmationMessage(
    'Formulario recibido correctamente. Nos comunicaremos a la brevedad. ' +
    'Conservá este número de caso para consultas.'
  );

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 1 – DATOS DEL REGISTRO
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 1 — Datos del Registro')
    .setHelpText('Información básica sobre la confección de esta ficha.');

  form.addDateItem()
    .setTitle('Fecha de confección')
    .setRequired(true);

  form.addTextItem()
    .setTitle('N° de expediente / caso')
    .setHelpText('Dejá vacío si no tenés número asignado. Se generará automáticamente.')
    .setRequired(false);

  var estadoCaso = form.addListItem()
    .setTitle('Estado del caso')
    .setRequired(true);
  estadoCaso.setChoices([
    estadoCaso.createChoice('Abierto'),
    estadoCaso.createChoice('En proceso'),
    estadoCaso.createChoice('Cerrado')
  ]);

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 2 – DATOS PERSONALES
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 2 — Datos Personales del Accidentado')
    .setHelpText('Completá con los datos del accidentado/a.');

  form.addTextItem()
    .setTitle('Nombre y apellido completo')
    .setRequired(true);

  // Validación DNI: solo números, 7 u 8 dígitos
  form.addTextItem()
    .setTitle('D.N.I. N°')
    .setHelpText('Solo números, sin puntos ni espacios. Ej: 17183165')
    .setRequired(true)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextMatchesPattern('^[0-9]{7,8}$')
        .setHelpText('Ingresá entre 7 y 8 dígitos numéricos, sin puntos.')
        .build()
    );

  form.addDateItem()
    .setTitle('Fecha de nacimiento')
    .setRequired(true);

  // Validación edad: 1-2 dígitos numéricos
  form.addTextItem()
    .setTitle('Edad (años)')
    .setHelpText('Ingresá solo el número. Ej: 62')
    .setRequired(false)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextMatchesPattern('^[0-9]{1,3}$')
        .setHelpText('Ingresá solo números (sin letras).')
        .build()
    );

  var estadoCivil = form.addListItem()
    .setTitle('Estado civil')
    .setRequired(true);
  estadoCivil.setChoices([
    estadoCivil.createChoice('Soltero/a'),
    estadoCivil.createChoice('Casado/a'),
    estadoCivil.createChoice('Viudo/a'),
    estadoCivil.createChoice('Divorciado/a'),
    estadoCivil.createChoice('Unión convivencial'),
    estadoCivil.createChoice('Otro')
  ]);

  form.addTextItem()
    .setTitle('Nacionalidad')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Nombre del padre (solo si es menor de edad)')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Nombre de la madre (solo si es menor de edad)')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Domicilio completo')
    .setHelpText('Incluí calle, número, manzana, barrio, localidad.')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Localidad / ciudad')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Provincia')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Teléfono fijo')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Teléfono celular')
    .setRequired(true);

  // Validación email
  form.addTextItem()
    .setTitle('Email de contacto')
    .setHelpText('Ej: nombre@correo.com')
    .setRequired(false)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextIsEmail()
        .setHelpText('Ingresá un email válido.')
        .build()
    );

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 3 – GRUPO FAMILIAR
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 3 — Grupo Familiar')
    .setHelpText('Información sobre el entorno familiar del accidentado/a.');

  form.addTextItem()
    .setTitle('Cantidad de hijos')
    .setHelpText('Solo números. Ej: 3')
    .setRequired(false)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextMatchesPattern('^[0-9]{1,2}$')
        .setHelpText('Ingresá un número válido.')
        .build()
    );

  var convive = form.addMultipleChoiceItem()
    .setTitle('¿Convive con familiares?')
    .setRequired(false);
  convive.setChoices([
    convive.createChoice('Sí'),
    convive.createChoice('No')
  ]);

  form.addParagraphTextItem()
    .setTitle('Descripción del grupo familiar')
    .setHelpText('Ej: 5 hijos, 1 de 21 años y 4 casados que viven fuera del hogar.')
    .setRequired(false);

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 4 – DATOS LABORALES
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 4 — Datos Laborales')
    .setHelpText('Situación laboral al momento del accidente.');

  var situacionLaboral = form.addListItem()
    .setTitle('Situación laboral')
    .setRequired(true);
  situacionLaboral.setChoices([
    situacionLaboral.createChoice('Empleado/a en relación de dependencia'),
    situacionLaboral.createChoice('Trabajador/a independiente / monotributista'),
    situacionLaboral.createChoice('Jubilado/a'),
    situacionLaboral.createChoice('Pensionado/a'),
    situacionLaboral.createChoice('Desocupado/a'),
    situacionLaboral.createChoice('Otro')
  ]);

  form.addTextItem()
    .setTitle('Nombre del empleador / empresa')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Domicilio laboral')
    .setRequired(false);

  var bono = form.addMultipleChoiceItem()
    .setTitle('¿Presenta bono de sueldo?')
    .setRequired(true);
  bono.setChoices([
    bono.createChoice('Sí'),
    bono.createChoice('No')
  ]);

  form.addTextItem()
    .setTitle('Antigüedad laboral')
    .setHelpText('Ej: 5 años')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Categoría / cargo laboral')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Sueldo aproximado')
    .setHelpText('Ej: 500000')
    .setRequired(false);

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 5 – DATOS DEL ACCIDENTE
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 5 — Datos del Accidente')
    .setHelpText('Información sobre el hecho en sí.');

  form.addDateItem()
    .setTitle('Fecha del accidente')
    .setRequired(true);

  form.addTimeItem()
    .setTitle('Hora del accidente')
    .setRequired(true);

  var tipoAccidente = form.addListItem()
    .setTitle('Tipo de accidente')
    .setRequired(true);
  tipoAccidente.setChoices([
    tipoAccidente.createChoice('Caída en vía pública'),
    tipoAccidente.createChoice('Accidente laboral'),
    tipoAccidente.createChoice('Accidente de tránsito'),
    tipoAccidente.createChoice('Accidente doméstico'),
    tipoAccidente.createChoice('Agresión / violencia'),
    tipoAccidente.createChoice('Otro')
  ]);

  form.addTextItem()
    .setTitle('Lugar del accidente')
    .setHelpText('Ej: Av. Gral. San Martín S/N')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Dirección exacta / referencia del lugar')
    .setHelpText('Intersección más cercana, punto de referencia, etc.')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Descripción detallada de los hechos')
    .setHelpText('Relatá con el mayor detalle posible cómo ocurrió el accidente.')
    .setRequired(true);

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 6 – LESIONES Y ATENCIÓN MÉDICA
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 6 — Lesiones y Asistencia Médica')
    .setHelpText('Detallá las lesiones sufridas y la atención recibida.');

  var tipoLesion = form.addCheckboxItem()
    .setTitle('Tipo de lesión (podés marcar más de una)')
    .setRequired(true);
  tipoLesion.setChoices([
    tipoLesion.createChoice('Esguince'),
    tipoLesion.createChoice('Fractura'),
    tipoLesion.createChoice('Golpes / contusiones'),
    tipoLesion.createChoice('Cortes / heridas'),
    tipoLesion.createChoice('Lesión muscular'),
    tipoLesion.createChoice('Traumatismo craneal'),
    tipoLesion.createChoice('Quemaduras'),
    tipoLesion.createChoice('Otra')
  ]);

  var zonaAfectada = form.addCheckboxItem()
    .setTitle('Parte del cuerpo afectada (podés marcar más de una)')
    .setRequired(true);
  zonaAfectada.setChoices([
    zonaAfectada.createChoice('Tobillo'),
    zonaAfectada.createChoice('Rodilla'),
    zonaAfectada.createChoice('Cadera'),
    zonaAfectada.createChoice('Brazo / codo'),
    zonaAfectada.createChoice('Mano / muñeca'),
    zonaAfectada.createChoice('Espalda / columna'),
    zonaAfectada.createChoice('Cabeza / cuello'),
    zonaAfectada.createChoice('Pie'),
    zonaAfectada.createChoice('Otra')
  ]);

  form.addTextItem()
    .setTitle('Centro médico donde fue atendido')
    .setRequired(false);

  var realizoRx = form.addMultipleChoiceItem()
    .setTitle('¿Se realizaron radiografías u otros estudios?')
    .setRequired(true);
  realizoRx.setChoices([
    realizoRx.createChoice('Sí'),
    realizoRx.createChoice('No')
  ]);

  form.addParagraphTextItem()
    .setTitle('Diagnóstico médico')
    .setHelpText('Tal como figura en el informe médico.')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Estado actual del accidentado/a')
    .setHelpText('Describí la situación a la fecha de confección de esta ficha.')
    .setRequired(false);

  // ── Nuevos campos de lesiones (segunda hoja) ─────────────────
  form.addParagraphTextItem()
    .setTitle('Descripción de lesiones')
    .setHelpText('Descripción libre de las lesiones. Ej: Ambos tobillos esguinzados.')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Días de internación')
    .setHelpText('Ingresá solo el número. Ej: 3. Si no hubo internación, dejá vacío.')
    .setRequired(false)
    .setValidation(
      FormApp.createTextValidation()
        .requireTextMatchesPattern('^[0-9]{1,4}$')
        .setHelpText('Solo números.')
        .build()
    );

  var tuvoOperacion = form.addMultipleChoiceItem()
    .setTitle('¿Se realizaron operaciones / intervenciones quirúrgicas?')
    .setRequired(false);
  tuvoOperacion.setChoices([
    tuvoOperacion.createChoice('Sí'),
    tuvoOperacion.createChoice('No')
  ]);

  form.addParagraphTextItem()
    .setTitle('Descripción de operaciones / intervenciones')
    .setHelpText('Describí las intervenciones quirúrgicas realizadas, si aplica.')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Tratamientos recibidos o en curso')
    .setHelpText('Ej: Kinesiología, medicación, reposo, yeso, etc.')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Secuelas')
    .setHelpText('Describí las secuelas físicas o psicológicas que persisten hasta la fecha.')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Atención médica recibida')
    .setHelpText('Listá todos los centros médicos, clínicas y profesionales que atendieron al accidentado/a. ' +
                 'Ej: Consultorios Médicos Morales – Centro Médico La Bronce – Ecografía Aguilar – Clínica Santa Clara.')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Seguro del demandado')
    .setHelpText('Nombre de la aseguradora o compañía de seguros del demandado, si se conoce.')
    .setRequired(false);

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 7 – TESTIGOS Y ASISTENCIA
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 7 — Testigos y Asistencia')
    .setHelpText('Información sobre quienes presenciaron o asistieron en el hecho.');

  var testigos = form.addMultipleChoiceItem()
    .setTitle('¿Hubo testigos del accidente?')
    .setRequired(true);
  testigos.setChoices([
    testigos.createChoice('Sí'),
    testigos.createChoice('No')
  ]);

  form.addParagraphTextItem()
    .setTitle('Datos de testigos')
    .setHelpText('Nombre completo, DNI y teléfono de cada testigo (uno por línea).')
    .setRequired(false);

  var policiaIntervino = form.addMultipleChoiceItem()
    .setTitle('¿Intervino personal policial?')
    .setRequired(true);
  policiaIntervino.setChoices([
    policiaIntervino.createChoice('Sí'),
    policiaIntervino.createChoice('No')
  ]);

  form.addParagraphTextItem()
    .setTitle('Observaciones adicionales')
    .setHelpText('Cualquier información relevante no contemplada en los campos anteriores.')
    .setRequired(false);

  // ═══════════════════════════════════════════════════════════
  // SECCIÓN 8 – DOCUMENTACIÓN ADJUNTA
  // Nota: Apps Script NO soporta addFileUploadItem().
  // La carga de archivos se habilita directamente en el editor
  // de Google Forms (formulario → "+" → "Subida de archivos").
  // Acá se agregan campos de texto para que el usuario ingrese
  // el link de Drive de cada documento adjunto.
  // ═══════════════════════════════════════════════════════════
  form.addPageBreakItem()
    .setTitle('SECCIÓN 8 — Documentación Adjunta')
    .setHelpText(
      'INSTRUCCIÓN: Subí tus archivos a Google Drive y pegá aquí el link de cada uno. ' +
      'Asegurate de que el archivo sea accesible ("Cualquier persona con el link puede ver").'
    );

  form.addTextItem()
    .setTitle('Link a fotos del lugar y/o lesiones (Google Drive)')
    .setHelpText('Ej: https://drive.google.com/file/d/...')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Link a estudios médicos / radiografías (Google Drive)')
    .setHelpText('Informes, radiografías, resonancias, etc.')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Link a copia del D.N.I. (Google Drive)')
    .setHelpText('Frente y dorso del DNI del accidentado/a.')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Link a otros documentos (denuncia policial, certificados, etc.)')
    .setHelpText('Denuncias, certificados médicos u otros comprobantes.')
    .setRequired(false);

  // ── Vincular con Google Sheets ──────────────────────────────
  var spreadsheet = SpreadsheetApp.create('Registros – Radiografía del Hecho');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  // ── Mostrar links en los registros de ejecución ─────────────
  var formUrl    = form.getEditUrl();
  var publishUrl = form.getPublishedUrl();
  var sheetUrl   = spreadsheet.getUrl();

  Logger.log('===========================================');
  Logger.log('FORMULARIO CREADO EXITOSAMENTE');
  Logger.log('===========================================');
  Logger.log('Editar formulario : ' + formUrl);
  Logger.log('Link público      : ' + publishUrl);
  Logger.log('Google Sheet      : ' + sheetUrl);
  Logger.log('===========================================');
  Logger.log('PRÓXIMOS PASOS:');
  Logger.log('1. Abrí el formulario con el link de edición');
  Logger.log('2. Agregá campos de "Subida de archivos" manualmente en la Sección 8');
  Logger.log('3. Ejecutá activarTriggerID() para IDs automáticos (reemplazá TU_SPREADSHEET_ID)');
  Logger.log('===========================================');
}

// ─────────────────────────────────────────────────────────────
// FUNCIÓN: Activar trigger de IDs automáticos
// Ejecutá DESPUÉS de crearFormulario().
// Reemplazá TU_SPREADSHEET_ID con el ID del Sheet generado.
// (Está en la URL del Sheet: docs.google.com/spreadsheets/d/ESTE_ID/edit)
// ─────────────────────────────────────────────────────────────
function activarTriggerID() {
  var SPREADSHEET_ID = 'TU_SPREADSHEET_ID'; // ← reemplazá esto

  // Eliminar triggers anteriores del mismo tipo para evitar duplicados
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'agregarIDAutomatico') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('agregarIDAutomatico')
    .forSpreadsheet(SPREADSHEET_ID)
    .onFormSubmit()
    .create();

  Logger.log('Trigger de ID automático activado correctamente.');
}

// ─────────────────────────────────────────────────────────────
// FUNCIÓN: Se ejecuta automáticamente en cada envío del formulario.
// Escribe el ID de caso en la columna A de la fila nueva.
// ─────────────────────────────────────────────────────────────
function agregarIDAutomatico(e) {
  try {
    var sheet  = e.range.getSheet();
    var row    = e.range.getRow();
    var year   = new Date().getFullYear();
    var idCaso = 'EXP-' + year + '-' + String(row - 1).padStart(4, '0');

    sheet.getRange(row, 1).setValue(idCaso);

    // Opcional: enviar email de confirmación al accidentado
    // Descomentá las siguientes líneas para activarlo:
    // var email = e.namedValues['Email de contacto'] ? e.namedValues['Email de contacto'][0] : '';
    // if (email) {
    //   MailApp.sendEmail(
    //     email,
    //     'Registro recibido – Caso ' + idCaso,
    //     'Su caso fue registrado con el número ' + idCaso + '. Nos comunicaremos a la brevedad.'
    //   );
    // }

  } catch (err) {
    Logger.log('Error en agregarIDAutomatico: ' + err.message);
  }
}

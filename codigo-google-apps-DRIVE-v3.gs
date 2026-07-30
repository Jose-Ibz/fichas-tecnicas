// ========================================
// SISTEMA VIAMAR v3.0 - CON GOOGLE DRIVE
// ========================================

// CONFIGURACIÓN
var SHEET_ID = '1CkiLr9gDWJDg2eTpUVX9ZYLRlQUbBp8GNdMBoedU-a8'; // TU SHEET ID
var EMAIL_DESTINO = 'jose@viamar-ibiza.com';
var USERS_SHEET = 'Usuarios';
var INSPECTIONS_SHEET = 'Inspecciones';
var DRIVE_FOLDER_NAME = 'Inspecciones-Viamar'; // Carpeta raíz en Drive
var ANTHROPIC_API_KEY = ''; // Pon aqui tu API Key de Anthropic

// ========== FUNCIONES PRINCIPALES ==========

function doGet(e) {
  try {
    var action = e.parameter.action;
    
    if (action === 'login') {
      return handleLogin(e.parameter.username || e.parameter.user, e.parameter.password || e.parameter.pass);
    } else if (action === 'getHistory') {
      return handleGetHistory(e.parameter.username || e.parameter.user, e.parameter.desde);
    } else if (action === 'getInspection') {
      return handleGetInspection(e.parameter.id);
    } else if (action === 'getLastVarada') {
      return handleGetLastVarada(e.parameter.boat);
    } else if (action === 'getTareas') {
      return handleGetTareas();
    } else if (action === 'loginInventario') {
      return handleLoginInventario(e.parameter.user, e.parameter.pass);
    } else if (action === 'getInventario') {
      return handleGetInventario(e.parameter.embarcacion);
    } else if (action === 'getInventarios') {
      return handleGetInventarios();
    } else if (action === 'getInventarioAccesos') {
      return handleGetInventarioAccesos();
    } else if (action === 'getUsuarios') {
      return handleGetUsuarios();
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Accion no reconocida'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    
    if (action === 'save') {
      return handleSaveInspection(data.data);
    } else if (action === 'uploadPhoto') {
      return handleUploadPhoto(data);
    } else if (action === 'uploadPhotoBatch') {
      return handleUploadPhotoBatch(data);
    } else if (action === 'aiQuery') {
      return handleAIQuery(data.question, data.history);
    } else if (action === 'saveTarea') {
      return handleSaveTarea(data.tarea);
    } else if (action === 'completarTarea') {
      return handleCompletarTarea(data.id);
    } else if (action === 'saveInventario') {
      return handleSaveInventario(data.data);
    } else if (action === 'uploadInventarioPhotoBatch') {
      return handleUploadInventarioPhotoBatch(data);
    } else if (action === 'saveInventarioAcceso') {
      return handleSaveInventarioAcceso(data.acceso);
    } else if (action === 'saveUsuario') {
      return handleSaveUsuario(data.usuario);
    } else if (action === 'toggleUsuario') {
      return handleToggleUsuario(data.username, data.activo);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Accion no reconocida'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== GOOGLE DRIVE - GESTIÓN DE FOTOS ==========

/**
 * Obtener o crear carpeta raíz de inspecciones
 */
function getRootFolder() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    // Crear carpeta raíz
    return DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
}

/**
 * Obtener o crear carpeta de año
 */
function getYearFolder(year) {
  var rootFolder = getRootFolder();
  var folders = rootFolder.getFoldersByName(year.toString());
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return rootFolder.createFolder(year.toString());
  }
}

/**
 * Obtener o crear carpeta de barco
 */
function getBoatFolder(year, boatName) {
  var yearFolder = getYearFolder(year);
  // Limpiar nombre del barco (sin espacios, mayúsculas)
  var cleanBoatName = boatName.toUpperCase().replace(/ /g, '_');
  var folders = yearFolder.getFoldersByName(cleanBoatName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return yearFolder.createFolder(cleanBoatName);
  }
}

/**
 * Obtener o crear carpeta de inspección específica
 */
function getInspectionFolder(year, boatName, month, tipo) {
  var boatFolder = getBoatFolder(year, boatName);
  // Formato: "2026-03-Varada"
  var folderName = year + '-' + padZero(month) + '-' + tipo.charAt(0).toUpperCase() + tipo.slice(1);
  var folders = boatFolder.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return boatFolder.createFolder(folderName);
  }
}

/**
 * Subir foto a Google Drive
 */
function uploadPhotoToDrive(base64Data, inspectionData, photoId) {
  try {
    // Extraer datos
    var boat = inspectionData.embarcacion;
    var tipo = inspectionData.tipo;
    var fecha = new Date(inspectionData.fecha);
    var year = fecha.getFullYear();
    var month = fecha.getMonth() + 1;
    
    // Obtener carpeta de inspección
    var folder = getInspectionFolder(year, boat, month, tipo);
    
    // Limpiar base64 (quitar prefijo data:image/...)
    var base64Clean = base64Data.split(',')[1];
    
    // Determinar extensión
    var extension = 'jpg';
    if (base64Data.indexOf('image/png') > -1) extension = 'png';
    if (base64Data.indexOf('image/jpeg') > -1) extension = 'jpg';
    
    // Nombre del archivo
    var fileName = photoId + '.' + extension;
    
    // Convertir base64 a blob
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Clean),
      'image/' + extension,
      fileName
    );
    
    // Verificar si ya existe el archivo (por si re-sube)
    var existingFiles = folder.getFilesByName(fileName);
    if (existingFiles.hasNext()) {
      // Eliminar archivo antiguo
      existingFiles.next().setTrashed(true);
    }
    
    // Subir archivo
    var file = folder.createFile(blob);
    
    // Hacer el archivo accesible
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Devolver URL
    return {
      success: true,
      url: file.getUrl(),
      fileId: file.getId(),
      fileName: fileName
    };
    
  } catch(error) {
    Logger.log('Error subiendo foto: ' + error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Manejar subida de foto desde cliente
 */
function handleUploadPhoto(data) {
  try {
    var result = uploadPhotoToDrive(
      data.base64Data,
      data.inspectionData,
      data.photoId
    );
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error subiendo foto: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== FUNCIONES DE AUTENTICACIÓN ==========

function handleLogin(username, password) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var userSheet = ss.getSheetByName(USERS_SHEET);
    var data = userSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var dbUser   = String(data[i][0]||'').trim();
      var dbPass   = String(data[i][1]||'').trim();
      var dbActivo = String(data[i][3]||'').trim().toUpperCase();
      var inUser   = String(username||'').trim();
      var inPass   = String(password||'').trim();
      if (dbUser === inUser && dbPass === inPass && dbActivo === 'SI') {
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          user: dbUser,
          role: String(data[i][2]||'').trim()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Usuario o contraseña incorrectos'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error en login: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== FUNCIONES DE INSPECCIONES ==========

function handleSaveInspection(inspData) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var inspSheet = ss.getSheetByName(INSPECTIONS_SHEET);
    
    // Generar ID único
    var id = 'INS_' + new Date().getTime();
    var fecha = new Date(inspData.fecha);
    
    // v3.0: Las fotos ya se suben vía uploadPhotoBatch antes de llamar a save.
    // Aquí solo recogemos las URLs que vienen pre-cargadas en inspData.fotosUrls.
    var fotosUrls = Object.assign({}, inspData.fotosUrls || {});
    
    // Preparar datos para Sheet
    var checksPrevio = inspData.checklistPrevio || {};
    var checksVerif = inspData.verificaciones || {};
    var notas = inspData.notas || {};
    var medidoresBabor = inspData.medidoresBabor || {};
    var medidoresEstribor = inspData.medidoresEstribor || {};
    var medidoresCentral = inspData.medidoresCentral || {};
    
    // Contar items
    var countPrevio = Object.keys(checksPrevio).filter(function(k){ return checksPrevio[k] === true; }).length;
    var countVerif = Object.keys(checksVerif).filter(function(k){ return !k.endsWith('_photo') && checksVerif[k] === true; }).length;
    var countNotas = Object.keys(notas).filter(function(k){ return notas[k]; }).length;
    var countFotos = Object.keys(fotosUrls).length;
    
    // Estado de la inspección
    var estado = inspData.estadoInspeccion || 'ACTIVA';
    var cerrada = 'NO';
    
    // Fila de datos (40 columnas + 2 nuevas)
    var row = [
      id,                                              // A(0) - ID
      Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy'), // B(1) - Fecha
      Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'HH:mm'),      // C(2) - Hora
      inspData.tecnico || '',                          // D(3) - Tecnico
      inspData.embarcacion || '',                      // E(4) - Embarcacion
      inspData.modelo || '',                           // F(5) - Modelo
      inspData.tipo || '',                             // G(6) - Tipo
      inspData.cliente || '',                          // H(7) - Cliente
      inspData.motores || '',                          // I(8) - Motores
      inspData.colas || '',                            // J(9) - Colas
      inspData.motorBabor || '',                       // K(10) - Motor Babor Nº
      inspData.motorEstribor || '',                    // L(11) - Motor Estribor Nº
      inspData.motorCentral || '',                     // M(12) - Motor Central Nº
      inspData.motorAuxiliarNum || '',                 // N(13) - Motor Auxiliar Nº
      inspData.motorAuxiliarModelo || '',              // O(14) - Motor Auxiliar Modelo
      inspData.llaves || '',                           // P(15) - Llaves
      inspData.cabina || '',                           // Q(16) - Cabina
      inspData.amarre || '',                           // R(17) - Amarre
      inspData.pasoHelice || '',                       // S(18) - Paso Helice
      inspData.fechaVarada || '',                      // T(19) - Fecha Varada
      inspData.fechaBotadura || '',                    // U(20) - Fecha Botadura
      inspData.horasBabor || '',                       // V(21) - Horas Babor
      inspData.horasEstribor || '',                    // W(22) - Horas Estribor
      inspData.horasCentral || '',                     // X(23) - Horas Central
      inspData.horasGenerador || '',                   // Y(24) - Horas Generador
      inspData.observaciones || '',                    // Z(25) - Observaciones
      JSON.stringify(checksPrevio),                    // AA(26) - Checklist Previo JSON
      JSON.stringify(checksVerif),                     // AB(27) - Verificaciones JSON
      JSON.stringify(notas),                           // AC(28) - Notas JSON
      JSON.stringify(medidoresBabor),                  // AD(29) - Medidores Babor JSON
      JSON.stringify(medidoresEstribor),               // AE(30) - Medidores Estribor JSON
      JSON.stringify(medidoresCentral),                // AF(31) - Medidores Central JSON
      '{}',                                            // AG(32) - Fotos Reportaje JSON (VACÍO - ahora en Drive)
      countPrevio,                                     // AH(33) - Items Checklist Previo
      countVerif,                                      // AI(34) - Items Verificaciones
      countNotas,                                      // AJ(35) - Total Notas
      countFotos,                                      // AK(36) - Total Fotos
      inspData.varadaImportada || '',                  // AL(37) - Varada Importada
      estado,                                          // AM(38) - Estado
      cerrada,                                         // AN(39) - Cerrada
      JSON.stringify(fotosUrls),                       // AO(40) - Fotos URLs JSON
      '',                                              // AP(41) - Carpeta Drive ID
      JSON.stringify(inspData.seguridad || {})         // AQ(42) - Seguridad JSON
    ];
    
    // Añadir fila
    inspSheet.appendRow(row);
    
    // Si es botadura, cerrar varada anterior
    if (inspData.tipo === 'botadura' && inspData.varadaImportada) {
      closeVarada(inspData.varadaImportada);
    }

    // Si es continuación de varada, cerrar la fila anterior (queda como CERRADA)
    if (inspData.tipo === 'varada' && inspData.esContinuacion && inspData.varadaImportada) {
      closeVarada(inspData.varadaImportada);
    }
    
    // Enviar email (no crítico — si falla no bloquea el guardado)
    try {
      enviarEmailNotificacion(inspData, id, fecha);
    } catch(emailError) {
      Logger.log('Email no enviado (no crítico): ' + emailError.toString());
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      id: id,
      message: 'Inspección guardada correctamente'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error guardando: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// CERRAR VARADA
function closeVarada(varadaId) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var inspSheet = ss.getSheetByName(INSPECTIONS_SHEET);
    var data = inspSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === varadaId) {
        inspSheet.getRange(i + 1, 39).setValue('CERRADA');  // AM(38) Estado
        inspSheet.getRange(i + 1, 40).setValue('SI');       // AN(39) Cerrada
        break;
      }
    }
  } catch(error) {
    Logger.log('Error cerrando varada: ' + error);
  }
}

// ENVIAR EMAIL
function enviarEmailNotificacion(inspData, id, fecha) {
  var estado = inspData.estadoInspeccion || 'PARCIAL';
  var porcentaje = inspData.porcentajeCompleto || 0;
  
  var asunto = '['+estado+'] Inspección '+inspData.tipo.toUpperCase()+': '+inspData.embarcacion+' ('+porcentaje+'%)';
  
  var cuerpo = '<html><body style="font-family:Arial,sans-serif;">';
  cuerpo += '<div style="background:#003d82;color:white;padding:20px;text-align:center;">';
  cuerpo += '<h1>NÁUTICA VIAMAR</h1>';
  cuerpo += '<p>Nueva Inspección Registrada</p>';
  cuerpo += '</div>';
  cuerpo += '<div style="padding:20px;">';
  cuerpo += '<h2 style="color:#003d82;">Detalles</h2>';
  cuerpo += '<table style="width:100%;border-collapse:collapse;">';
  cuerpo += '<tr><td style="padding:8px;border-bottom:1px solid #ddd;"><strong>ID:</strong></td><td style="padding:8px;border-bottom:1px solid #ddd;">' + id + '</td></tr>';
  cuerpo += '<tr><td style="padding:8px;border-bottom:1px solid #ddd;"><strong>Estado:</strong></td><td style="padding:8px;border-bottom:1px solid #ddd;background:'+(estado==='COMPLETADO'?'#d4edda':'#fff3cd')+';color:'+(estado==='COMPLETADO'?'#155724':'#856404')+';font-weight:bold;">' + estado + ' ('+porcentaje+'%)</td></tr>';
  cuerpo += '<tr><td style="padding:8px;border-bottom:1px solid #ddd;"><strong>Tipo:</strong></td><td style="padding:8px;border-bottom:1px solid #ddd;">' + inspData.tipo.toUpperCase() + '</td></tr>';
  cuerpo += '<tr><td style="padding:8px;border-bottom:1px solid #ddd;"><strong>Embarcación:</strong></td><td style="padding:8px;border-bottom:1px solid #ddd;">' + inspData.embarcacion + '</td></tr>';
  cuerpo += '<tr><td style="padding:8px;border-bottom:1px solid #ddd;"><strong>Técnico:</strong></td><td style="padding:8px;border-bottom:1px solid #ddd;">' + inspData.tecnico + '</td></tr>';
  cuerpo += '<tr><td style="padding:8px;border-bottom:1px solid #ddd;"><strong>Fecha:</strong></td><td style="padding:8px;border-bottom:1px solid #ddd;">' + Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '</td></tr>';
  cuerpo += '</table>';
  cuerpo += '</div>';
  cuerpo += '</body></html>';
  
  MailApp.sendEmail({
    to: EMAIL_DESTINO,
    subject: asunto,
    htmlBody: cuerpo
  });
}

// Función llamada por trigger asíncrono para enviar emails pendientes
function enviarEmailPendiente() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    var triggers = ScriptApp.getProjectTriggers();

    Object.keys(all).forEach(function(key) {
      if (key.indexOf('pendingEmail_') === 0) {
        try {
          var d = JSON.parse(all[key]);
          var estado = d.estadoInspeccion || 'PARCIAL';
          var porcentaje = d.porcentajeCompleto || 0;
          var asunto = '['+estado+'] Inspección '+d.tipo.toUpperCase()+': '+d.embarcacion+' ('+porcentaje+'%)';
          var cuerpo = '<html><body style="font-family:Arial,sans-serif;">'
            + '<div style="background:#003d82;color:white;padding:20px;text-align:center;"><h1>NÁUTICA VIAMAR</h1><p>Nueva Inspección Registrada</p></div>'
            + '<div style="padding:20px;"><table style="width:100%;border-collapse:collapse;">'
            + '<tr><td><strong>Estado:</strong></td><td style="font-weight:bold;color:'+(estado==='COMPLETADO'?'#155724':'#856404')+'">' + estado + ' ('+porcentaje+'%)</td></tr>'
            + '<tr><td><strong>Tipo:</strong></td><td>' + d.tipo.toUpperCase() + '</td></tr>'
            + '<tr><td><strong>Embarcación:</strong></td><td>' + d.embarcacion + '</td></tr>'
            + '<tr><td><strong>Técnico:</strong></td><td>' + d.tecnico + '</td></tr>'
            + '<tr><td><strong>Fecha:</strong></td><td>' + d.fechaStr + '</td></tr>'
            + '</table></div></body></html>';
          MailApp.sendEmail({ to: EMAIL_DESTINO, subject: asunto, htmlBody: cuerpo });
          props.deleteProperty(key);
        } catch(e) {
          Logger.log('Error enviando email pendiente ' + key + ': ' + e.toString());
          props.deleteProperty(key); // eliminar igualmente para no acumular
        }
      }
    });

    // Limpiar los triggers de este tipo para no acumularlos
    triggers.forEach(function(t) {
      if (t.getHandlerFunction() === 'enviarEmailPendiente') {
        try { ScriptApp.deleteTrigger(t); } catch(e) {}
      }
    });
  } catch(e) {
    Logger.log('Error en enviarEmailPendiente: ' + e.toString());
  }
}

// OBTENER HISTORIAL
// HELPERS para lectura por nombre de columna (usados por handleGetHistory)
function ghGetVal(row, colMap, name) {
  var idx = colMap[name.toLowerCase()];
  if (idx === undefined || idx >= row.length) return '';
  var v = row[idx];
  return (v === null || v === undefined) ? '' : v;
}
function ghSafeNum(v) {
  if (!v && v !== 0) return '';
  var s = String(v).trim();
  return (s.charAt(0) === '{' || s.charAt(0) === '[') ? '' : s;
}
function ghSafeDate(v) {
  if (!v) return '';
  var s = String(v).trim();
  if (s.charAt(0) === '{' || s.charAt(0) === '[' || s.length < 4) return '';
  return s;
}
function ghSafeText(v) {
  if (!v) return '';
  var s = String(v).trim();
  return (s.charAt(0) === '{' || s.charAt(0) === '[') ? '' : s;
}

function handleGetHistory(username, desde) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var inspSheet = ss.getSheetByName(INSPECTIONS_SHEET);
    var userSheet = ss.getSheetByName(USERS_SHEET);
    var userData = userSheet.getDataRange().getValues();
    var userRole = '';
    for (var u = 1; u < userData.length; u++) {
      if (userData[u][0] === username) { userRole = userData[u][2]; break; }
    }

    var data = inspSheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = {};
    for (var h = 0; h < headers.length; h++) {
      colMap[String(headers[h]||'').trim().toLowerCase()] = h;
    }

    var desdeDate = desde ? new Date(desde) : null;

    var inspections = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      // Filtrar por fecha si se pasa el parámetro "desde"
      if (desdeDate) {
        var fechaIdx = colMap['fecha'];
        if (fechaIdx !== undefined) {
          var rowDate = new Date(data[i][fechaIdx]);
          if (!isNaN(rowDate.getTime()) && rowDate < desdeDate) continue;
        }
      }

      var tecnico = String(ghGetVal(row, colMap, 'Tecnico')||'');
      if (userRole === 'Tecnico' && tecnico !== username) continue;

      var tipo   = String(ghGetVal(row, colMap, 'Tipo')||'').trim();
      var modelo = String(ghGetVal(row, colMap, 'Modelo')||'').trim();
      if (!tipo && (modelo === 'varada' || modelo === 'botadura')) {
        tipo = modelo; modelo = '';
      }

      var estado  = String(ghGetVal(row, colMap, 'Estado')||'').trim();
      var cerrada = String(ghGetVal(row, colMap, 'Cerrada')||'').trim();
      if (!estado) { estado = (cerrada === 'SI') ? 'CERRADA' : (cerrada === 'NO' ? 'ACTIVA' : ''); }

      var verifJson = String(ghGetVal(row, colMap, 'Verificaciones JSON') || ghGetVal(row, colMap, 'Checklist Verificaciones JSON') || '{}');
      var cpJson    = String(ghGetVal(row, colMap, 'Checklist Previo JSON') || '{}');
      var notasJson = String(ghGetVal(row, colMap, 'Notas JSON') || '{}');
      if (!verifJson) verifJson = '{}';
      if (!cpJson)    cpJson    = '{}';
      if (!notasJson) notasJson = '{}';

      inspections.push({
        id:                          String(ghGetVal(row, colMap, 'ID')||''),
        fecha:                       ghGetVal(row, colMap, 'Fecha'),
        hora:                        ghGetVal(row, colMap, 'Hora'),
        tecnico:                     tecnico,
        embarcacion:                 String(ghGetVal(row, colMap, 'Embarcacion')||''),
        modelo:                      modelo,
        tipo:                        tipo,
        cliente:                     ghSafeText(ghGetVal(row, colMap, 'Cliente')),
        motores:                     ghSafeText(ghGetVal(row, colMap, 'Motores')),
        colas:                       ghSafeText(ghGetVal(row, colMap, 'Colas')),
        motorBabor:                  ghSafeText(ghGetVal(row, colMap, 'Motor Babor')),
        motorEstribor:               ghSafeText(ghGetVal(row, colMap, 'Motor Estribor')),
        motorCentral:                ghSafeText(ghGetVal(row, colMap, 'Motor Central')),
        motorAuxiliarNum:            ghSafeText(ghGetVal(row, colMap, 'Motor Aux Nº')),
        motorAuxiliarModelo:         ghSafeText(ghGetVal(row, colMap, 'Motor Aux Modelo')),
        llaves:                      ghSafeText(ghGetVal(row, colMap, 'Llaves')),
        cabina:                      ghSafeText(ghGetVal(row, colMap, 'Cabina')),
        amarre:                      ghSafeText(ghGetVal(row, colMap, 'Amarre')),
        pasoHelice:                  ghSafeNum(ghGetVal(row, colMap, 'Paso Helice')),
        fechaVarada:                 ghSafeDate(ghGetVal(row, colMap, 'Fecha Varada')),
        fechaBotadura:               ghSafeDate(ghGetVal(row, colMap, 'Fecha Botadura')),
        horasBabor:                  ghSafeNum(ghGetVal(row, colMap, 'Horas Babor')),
        horasEstribor:               ghSafeNum(ghGetVal(row, colMap, 'Horas Estribor')),
        horasCentral:                ghSafeNum(ghGetVal(row, colMap, 'Horas Central')),
        horasGenerador:              ghSafeNum(ghGetVal(row, colMap, 'Horas Generador')),
        observaciones:               ghSafeText(ghGetVal(row, colMap, 'Observaciones')),
        checklistPrevioJson:         cpJson,
        checklistVerificacionesJson: verifJson,
        notasJson:                   notasJson,
        medidoresBaborJson:          String(ghGetVal(row, colMap, 'Medidores Babor JSON')||'{}'),
        medidoresEstribOrJson:       String(ghGetVal(row, colMap, 'Medidores Estribor JSON')||'{}'),
        medidoresCentralJson:        String(ghGetVal(row, colMap, 'Medidores Central JSON')||'{}'),
        itemsPrevio:                 ghGetVal(row, colMap, 'Items Previo')||0,
        itemsVerificaciones:         ghGetVal(row, colMap, 'Items Verificaciones')||0,
        totalNotas:                  ghGetVal(row, colMap, 'Total Notas')||0,
        totalFotos:                  ghGetVal(row, colMap, 'Total Fotos')||0,
        varadaImportada:             String(ghGetVal(row, colMap, 'Varada Importada')||''),
        estado:                      estado,
        cerrada:                     cerrada,
        fotosUrlsJson:               String(ghGetVal(row, colMap, 'Fotos URLs JSON')||'{}'),
        carpetaDriveId:              String(ghGetVal(row, colMap, 'Carpeta Drive ID')||''),
        seguridadJson:               String(ghGetVal(row, colMap, 'Seguridad JSON')||'{}')
      });
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      inspections: inspections
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error obteniendo historial: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
// OBTENER INSPECCIÓN POR ID
function handleGetInspection(id) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var inspSheet = ss.getSheetByName(INSPECTIONS_SHEET);
    var data = inspSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        var inspection = {
          id: data[i][0],
          fecha: data[i][1],
          hora: data[i][2],
          tecnico: data[i][3],
          embarcacion: data[i][4],
          modelo: data[i][5],
          tipo: data[i][6],
          cliente: data[i][7],
          motores: data[i][8],
          colas: data[i][9],
          motorBabor: data[i][10],
          motorEstribor: data[i][11],
          motorCentral: data[i][12],
          motorAuxiliarNum: data[i][13],
          motorAuxiliarModelo: data[i][14],
          llaves: data[i][15],
          cabina: data[i][16],
          amarre: data[i][17],
          pasoHelice: data[i][18],
          fechaVarada: data[i][19],
          fechaBotadura: data[i][20],
          horasBabor: data[i][21],
          horasEstribor: data[i][22],
          horasCentral: data[i][23],
          horasGenerador: data[i][24],
          observaciones: data[i][25],
          checklistPrevioJson: data[i][26],
          checklistVerificacionesJson: data[i][27],
          notasJson: data[i][28],
          medidoresBaborJson: data[i][29],
          medidoresEstribOrJson: data[i][30],
          medidoresCentralJson: data[i][31],
          fotosUrlsJson: data[i][40] || '{}',
          estado: data[i][38],
          cerrada: data[i][39],
          seguridadJson: String(data[i][42] || '{}')
        };
        
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          inspection: inspection
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Inspección no encontrada'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error obteniendo inspección: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// OBTENER ÚLTIMA VARADA DE UN BARCO
function handleGetLastVarada(boatName) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var inspSheet = ss.getSheetByName(INSPECTIONS_SHEET);
    var data = inspSheet.getDataRange().getValues();
    var headers = data[0];

    var colMap = {};
    for (var h = 0; h < headers.length; h++) {
      colMap[String(headers[h]||'').trim().toLowerCase()] = h;
    }
    function get(row, name) {
      var idx = colMap[name.toLowerCase()];
      if (idx === undefined || idx >= row.length) return '';
      var v = row[idx];
      return (v === null || v === undefined) ? '' : v;
    }

    var sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);

    var lastVarada = null;
    var lastDate = null;

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (!row[0]) continue;

      var embarcacion = String(get(row,'Embarcacion')||'').trim().toLowerCase();
      var tipo = String(get(row,'Tipo')||'').trim().toLowerCase();
      var modelo = String(get(row,'Modelo')||'').trim().toLowerCase();
      // old format: tipo was in Modelo column
      if (!tipo && (modelo==='varada'||modelo==='botadura')) tipo = modelo;

      var cerrada = String(get(row,'Cerrada')||'').trim().toUpperCase();
      var estado  = String(get(row,'Estado')||'').trim().toUpperCase();

      if (embarcacion !== boatName.toLowerCase().trim()) continue;
      if (tipo !== 'varada') continue;
      // Accept both active and recently closed varadas for import
      // (user might want to import a closed varada too)

      var inspDate = new Date(row[ colMap['fecha'] !== undefined ? colMap['fecha'] : 1 ]);
      if (isNaN(inspDate.getTime())) continue;

      if (!lastDate || inspDate > lastDate) {
        lastVarada = {
          id:                          String(get(row,'ID')||''),
          fecha:                       get(row,'Fecha'),
          tecnico:                     String(get(row,'Tecnico')||''),
          embarcacion:                 String(get(row,'Embarcacion')||''),
          modelo:                      modelo==='varada'?'':String(get(row,'Modelo')||''),
          tipo:                        tipo,
          cliente:                     String(get(row,'Cliente')||''),
          motores:                     String(get(row,'Motores')||''),
          colas:                       String(get(row,'Colas')||''),
          motorBabor:                  String(get(row,'Motor Babor')||''),
          motorEstribor:               String(get(row,'Motor Estribor')||''),
          motorCentral:                String(get(row,'Motor Central')||''),
          motorAuxiliarNum:            String(get(row,'Motor Aux Nº')||''),
          motorAuxiliarModelo:         String(get(row,'Motor Aux Modelo')||''),
          llaves:                      String(get(row,'Llaves')||''),
          cabina:                      String(get(row,'Cabina')||''),
          amarre:                      String(get(row,'Amarre')||''),
          pasoHelice:                  String(get(row,'Paso Helice')||''),
          fechaVarada:                 String(get(row,'Fecha Varada')||''),
          horasBabor:                  String(get(row,'Horas Babor')||''),
          horasEstribor:               String(get(row,'Horas Estribor')||''),
          horasCentral:                String(get(row,'Horas Central')||''),
          horasGenerador:              String(get(row,'Horas Generador')||''),
          observaciones:               String(get(row,'Observaciones')||''),
          checklistPrevioJson:         String(get(row,'Checklist Previo JSON')||'{}'),
          checklistVerificacionesJson: String(get(row,'Verificaciones JSON')||get(row,'Checklist Verificaciones JSON')||'{}'),
          notasJson:                   String(get(row,'Notas JSON')||'{}'),
          medidoresBaborJson:          String(get(row,'Medidores Babor JSON')||'{}'),
          medidoresEstribOrJson:       String(get(row,'Medidores Estribor JSON')||'{}'),
          medidoresCentralJson:        String(get(row,'Medidores Central JSON')||'{}'),
          cerrada:                     cerrada,
          estado:                      estado,
          fotosUrlsJson:               String(get(row,'Fotos URLs JSON')||'{}'),
          seguridadJson:               String(get(row,'Seguridad JSON')||'{}')
        };
        lastDate = inspDate;
      }
    }

    if (lastVarada) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: lastVarada
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'No se encontró varada para: ' + boatName
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
// ========== SUBIDA EN LOTE (todas las fotos en una sola llamada) ==========
function handleUploadPhotoBatch(data) {
  try {
    var photos = data.photos || [];
    var inspectionData = data.inspectionData;
    if (!photos.length) {
      return ContentService.createTextOutput(JSON.stringify({success:true, results:{}}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Crear la carpeta UNA SOLA VEZ para todas las fotos
    var boat  = inspectionData.embarcacion;
    var tipo  = inspectionData.tipo;
    var fecha = new Date(inspectionData.fecha);
    var year  = fecha.getFullYear();
    var month = fecha.getMonth() + 1;
    var folder = getInspectionFolder(year, boat, month, tipo);

    var results = {};
    photos.forEach(function(photo) {
      try {
        var base64Clean = photo.base64Data.split(',')[1];
        var extension = (photo.base64Data.indexOf('image/png') > -1) ? 'png' : 'jpg';
        var fileName  = photo.photoId + '.' + extension;
        var blob = Utilities.newBlob(
          Utilities.base64Decode(base64Clean),
          'image/' + extension,
          fileName
        );
        // Borrar si ya existía
        var existing = folder.getFilesByName(fileName);
        if (existing.hasNext()) existing.next().setTrashed(true);
        // Subir
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        results[photo.photoId] = {success: true, url: file.getUrl(), fileId: file.getId()};
      } catch(photoErr) {
        results[photo.photoId] = {success: false, error: photoErr.toString()};
      }
    });

    return ContentService.createTextOutput(JSON.stringify({success: true, results: results}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error batch fotos: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function padZero(num) {
  return num < 10 ? '0' + num : num.toString();
}

// ========== TRABAJOS PENDIENTES ==========

function getTareasSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Trabajos');
  if (!sh) {
    sh = ss.insertSheet('Trabajos');
    sh.appendRow(['ID','Embarcacion','Descripcion','FechaCreacion','FechaTope','Estado','FechaCompletado','Notas']);
    sh.getRange(1,1,1,8).setFontWeight('bold');
  }
  return sh;
}

function handleGetTareas() {
  try {
    var sh = getTareasSheet();
    var rows = sh.getDataRange().getValues();
    var tareas = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) continue;
      tareas.push({
        id: String(r[0]),
        embarcacion: r[1] || '',
        descripcion: r[2] || '',
        fechaCreacion: r[3] ? Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        fechaTope: r[4] ? Utilities.formatDate(new Date(r[4]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        estado: r[5] || 'PENDIENTE',
        fechaCompletado: r[6] ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        notas: r[7] || ''
      });
    }
    return ContentService.createTextOutput(JSON.stringify({success:true, tareas:tareas}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:e.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSaveTarea(t) {
  try {
    var sh = getTareasSheet();
    var rows = sh.getDataRange().getValues();
    var fechaTope = t.fechaTope ? new Date(t.fechaTope) : '';

    if (t.id) {
      // Editar existente
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(t.id)) {
          sh.getRange(i+1, 2).setValue(t.embarcacion || '');
          sh.getRange(i+1, 3).setValue(t.descripcion || '');
          sh.getRange(i+1, 5).setValue(fechaTope);
          sh.getRange(i+1, 8).setValue(t.notas || '');
          return ContentService.createTextOutput(JSON.stringify({success:true}))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    // Nueva tarea
    var id = 'TAR_' + new Date().getTime();
    sh.appendRow([id, t.embarcacion||'', t.descripcion||'', new Date(), fechaTope, 'PENDIENTE', '', t.notas||'']);
    return ContentService.createTextOutput(JSON.stringify({success:true, id:id}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:e.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleCompletarTarea(id) {
  try {
    var sh = getTareasSheet();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sh.getRange(i+1, 6).setValue('COMPLETADO');
        sh.getRange(i+1, 7).setValue(new Date());
        return ContentService.createTextOutput(JSON.stringify({success:true}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Tarea no encontrada'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:e.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== ASISTENTE IA ==========

function handleAIQuery(question, history) {
  try {
    if (!ANTHROPIC_API_KEY) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'API Key de Anthropic no configurada en el script.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- Construir contexto con todos los datos ---
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var inspSheet = ss.getSheetByName(INSPECTIONS_SHEET);
    var rows = inspSheet.getDataRange().getValues();

    var totalInsp = 0, varadasActivas = 0, varadasTotal = 0, botadurasTotal = 0;
    var inspLines = [];
    var hoy = new Date();

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row[0]) continue;
      totalInsp++;
      var tipo = String(row[6] || '').toLowerCase().trim();
      var estado = String(row[38] || '').trim();
      var cerrada = String(row[39] || '').toUpperCase().trim();
      if (tipo === 'varada') { varadasTotal++; if (cerrada !== 'SI') varadasActivas++; }
      else if (tipo === 'botadura') botadurasTotal++;

      var seg = {};
      try { seg = JSON.parse(String(row[42] || '{}')); } catch(e) {}

      var linea = '- [' + String(row[1] || '') + '] ' +
        String(row[4] || '').toUpperCase() + ' | ' +
        tipo.toUpperCase() + ' | Técnico: ' + String(row[3] || '') +
        ' | Estado: ' + estado;

      if (row[21]) linea += ' | HrBabor: ' + row[21];
      if (row[22]) linea += ' HrEstrib: ' + row[22];
      if (row[23]) linea += ' HrCentral: ' + row[23];
      if (row[5])  linea += ' | Modelo: ' + row[5];
      if (row[7])  linea += ' | Cliente: ' + row[7];
      if (seg.itbFecha) linea += ' | ITB: ' + seg.itbFecha;
      if (seg.pirotecniaFecha) linea += ' | Pirotecnia cad: ' + seg.pirotecniaFecha;
      if (seg.balsaTiene) linea += ' | Balsa: ' + seg.balsaTiene + (seg.balsaFecha ? ' (cad ' + seg.balsaFecha + ')' : '');
      if (seg.radiobalizaTiene) linea += ' | Radiobaliza: ' + seg.radiobalizaTiene + (seg.radiobalizaFecha ? ' (cad ' + seg.radiobalizaFecha + ')' : '');
      if (seg.vhfTiene) linea += ' | VHF: ' + seg.vhfTiene + (seg.vhfBateria ? ' Bat: ' + seg.vhfBateria : '') + (seg.vhfFecha ? ' (cad ' + seg.vhfFecha + ')' : '');
      var obs = String(row[25] || '').trim();
      if (obs) linea += ' | Obs: "' + obs.substring(0, 120) + (obs.length > 120 ? '...' : '') + '"';

      inspLines.push(linea);
    }

    var inspContext =
      'SISTEMA NÁUTICA VIAMAR (Ibiza) — Fecha hoy: ' + Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy') + '\n' +
      'Total inspecciones: ' + totalInsp + ' | Varadas: ' + varadasTotal +
      ' (' + varadasActivas + ' activas) | Botaduras: ' + botadurasTotal + '\n\n' +
      'INSPECCIONES:\n' + inspLines.join('\n');

    // ---- Inventarios (último por barco) ----
    var invContext = '';
    try {
      var invSheet = ss.getSheetByName('Inventarios');
      if (invSheet) {
        var invRows = invSheet.getDataRange().getValues();
        var invPorBarco = {};
        for (var ii = 1; ii < invRows.length; ii++) {
          var ir = invRows[ii];
          if (!ir[0]) continue;
          var bName = String(ir[1]||'');
          if (!invPorBarco[bName] || String(ir[7]) > String(invPorBarco[bName].fechaISO)) {
            invPorBarco[bName] = {fechaISO: ir[7], fechaStr: ir[2], realizadoPor: ir[3], itemsJson: ir[4]||'{}', notas: ir[6]||''};
          }
        }
        var invLines = [];
        Object.keys(invPorBarco).forEach(function(b) {
          var inv = invPorBarco[b];
          var items = {};
          try { items = JSON.parse(inv.itemsJson); } catch(e) {}
          var linea = '- ' + b + ' | Último inv: ' + inv.fechaStr + ' por ' + inv.realizadoPor;
          if (items.bengalas && items.bengalas.caducidad) linea += ' | Bengalas cad: ' + items.bengalas.caducidad;
          if (items.bengalas && items.bengalas.cantidad) linea += ' ('+items.bengalas.cantidad+' ud)';
          if (items.colchonetas && items.colchonetas.ubicacion) linea += ' | Colchonetas: ' + items.colchonetas.ubicacion;
          if (items.lona_fondeo && items.lona_fondeo.ubicacion) linea += ' | Lona fondeo: ' + items.lona_fondeo.ubicacion;
          if (items.chalecos && items.chalecos.cantidad) linea += ' | Chalecos: ' + items.chalecos.cantidad;
          if (items.extintor && items.extintor.cantidad) linea += ' | Extintores: ' + items.extintor.cantidad;
          if (inv.notas) linea += ' | Notas: "' + inv.notas.substring(0, 100) + '"';
          invLines.push(linea);
        });
        if (invLines.length > 0)
          invContext = '\n\nINVENTARIOS (último por barco):\n' + invLines.join('\n');
      }
    } catch(e) { Logger.log('invContext error: '+e); }

    // ---- Incidencias ----
    var incContext = '';
    try {
      var incSheet = ss.getSheetByName('Incidencias');
      if (incSheet) {
        var incRows = incSheet.getDataRange().getValues();
        var incAbiertas = [], incResueltas = 0;
        for (var ki = 1; ki < incRows.length; ki++) {
          var incRow = incRows[ki];
          if (!incRow[0]) continue;
          if (String(incRow[4]||'NO').toUpperCase() === 'SI') { incResueltas++; continue; }
          incAbiertas.push('  ⚠ ' + String(incRow[1]||'') + ': "' + String(incRow[2]||'') + '" ('+String(incRow[3]||'')+', reportado por: '+String(incRow[7]||'')+')');
        }
        incContext = '\n\nINCIDENCIAS: ' + incAbiertas.length + ' abiertas, ' + incResueltas + ' resueltas.\n';
        if (incAbiertas.length > 0) incContext += incAbiertas.join('\n');
        else incContext += 'Sin incidencias abiertas actualmente.';
      }
    } catch(e) { Logger.log('incContext error: '+e); }

    // ---- Trabajos pendientes ----
    var tareasContext = '';
    try {
      var tareasSheet = ss.getSheetByName('Trabajos');
      if (tareasSheet) {
        var tareasRows = tareasSheet.getDataRange().getValues();
        var tareasLines = [];
        for (var ti = 1; ti < tareasRows.length; ti++) {
          var tr = tareasRows[ti];
          if (!tr[0]) continue;
          var estado = String(tr[5]||'').toUpperCase();
          var linea = '  • ['+estado+'] '+String(tr[1]||'-')+': "'+String(tr[2]||'')+'"';
          if (tr[4]) linea += ' (tope: '+Utilities.formatDate(new Date(tr[4]), Session.getScriptTimeZone(), 'dd/MM/yyyy')+')';
          if (tr[7]) linea += ' — Notas: '+String(tr[7]);
          tareasLines.push(linea);
        }
        if (tareasLines.length > 0)
          tareasContext = '\n\nTRABAJOS PENDIENTES:\n' + tareasLines.join('\n');
      }
    } catch(e) { Logger.log('tareasContext error: '+e); }

    var dataContext = inspContext + invContext + incContext + tareasContext;

    var systemPrompt =
      'Eres el asistente inteligente del taller de Náutica Viamar, concesionario oficial Volvo Penta en Ibiza. ' +
      'Tienes acceso completo a: fichas técnicas de inspección de embarcaciones, inventarios de material por barco, incidencias abiertas y trabajos pendientes. ' +
      'Responde siempre en español, de forma clara, concisa y profesional. ' +
      'Si te preguntan por un barco concreto, busca TODOS sus registros (inspecciones e inventarios). ' +
      'Si detectas caducidades próximas o vencidas (ITB, pirotecnia, balsa, radiobaliza, VHF, bengalas), ' +
      'advierte claramente con emojis ⚠️ o 🔴. ' +
      'Cuando des listas, usa formato limpio con guiones o tablas de texto. ' +
      'Si no tienes datos suficientes para responder algo, dilo honestamente.\n\n' +
      dataContext;

    // Construir historial de mensajes para Claude
    var messages = [];
    if (history && Array.isArray(history)) {
      messages = history.slice(-20); // máximo 20 turnos para no exceder tokens
    }
    messages.push({ role: 'user', content: question });

    var payload = {
      model: 'claude-sonnet-4-6',
      max_tokens: 1500,
      system: systemPrompt,
      messages: messages
    };

    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var aiResult = JSON.parse(response.getContentText());

    if (aiResult.content && aiResult.content[0] && aiResult.content[0].text) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        response: aiResult.content[0].text
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Error API Anthropic: ' + JSON.stringify(aiResult)
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error asistente: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// SETUP INICIAL (ejecutar manualmente una vez)
function setupInitialSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Crear hoja Usuarios si no existe
  var userSheet = ss.getSheetByName(USERS_SHEET);
  if (!userSheet) {
    userSheet = ss.insertSheet(USERS_SHEET);
    userSheet.appendRow(['usuario', 'contraseña', 'rol', 'activo']);
    userSheet.appendRow(['admin', 'viamar2025', 'Administrador', 'SI']);
    userSheet.appendRow(['tecnico1', 'viamar123', 'Tecnico', 'SI']);
  }
  
  // Crear hoja Inspecciones si no existe
  var inspSheet = ss.getSheetByName(INSPECTIONS_SHEET);
  if (!inspSheet) {
    inspSheet = ss.insertSheet(INSPECTIONS_SHEET);
    inspSheet.appendRow([
      'ID', 'Fecha', 'Hora', 'Tecnico', 'Embarcacion', 'Modelo', 'Tipo', 'Cliente',
      'Motores', 'Colas', 'Motor Babor Nº', 'Motor Estribor Nº', 'Motor Central Nº',
      'Motor Auxiliar Nº', 'Motor Auxiliar Modelo', 'Llaves', 'Cabina', 'Amarre',
      'Paso Helice', 'Fecha Varada', 'Fecha Botadura', 'Horas Babor', 'Horas Estribor',
      'Horas Central', 'Horas Generador', 'Observaciones', 'Checklist Previo JSON',
      'Checklist Verificaciones JSON', 'Notas JSON', 'Medidores Babor JSON',
      'Medidores Estribor JSON', 'Medidores Central JSON', 'Fotos Reportaje JSON',
      'Items Checklist Previo', 'Items Verificaciones', 'Total Notas', 'Total Fotos',
      'Varada Importada', 'Estado', 'Cerrada', 'Fotos URLs JSON', 'Carpeta Drive ID'
    ]);
  }
  
  Logger.log('Setup completado');
}

// ========== INVENTARIO DE BARCOS v1.0 ==========

var INVENTARIOS_SHEET = 'Inventarios';
var INVENTARIO_ACCESOS_SHEET = 'InventarioAccesos';

function getInventariosSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(INVENTARIOS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(INVENTARIOS_SHEET);
    sh.appendRow(['ID','Embarcacion','Fecha','RealizadoPor','Items_JSON','FotosUrls_JSON','Notas','FechaCreacion']);
    sh.getRange(1,1,1,8).setFontWeight('bold');
  }
  return sh;
}

function getInventarioAccesosSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(INVENTARIO_ACCESOS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(INVENTARIO_ACCESOS_SHEET);
    sh.appendRow(['Embarcacion','Usuario','Password','Activo']);
    sh.getRange(1,1,1,4).setFontWeight('bold');
    sh.appendRow(['DEMO', 'inventario', 'viamar2026', 'SI']);
  }
  return sh;
}

function handleLoginInventario(username, password) {
  try {
    var sh = getInventarioAccesosSheet();
    var data = sh.getDataRange().getValues();
    var inUser = String(username||'').trim();
    var inPass = String(password||'').trim();
    for (var i = 1; i < data.length; i++) {
      var dbUser   = String(data[i][1]||'').trim();
      var dbPass   = String(data[i][2]||'').trim();
      var dbActivo = String(data[i][3]||'').trim().toUpperCase();
      var dbBarco  = String(data[i][0]||'').trim();
      if (dbUser === inUser && dbPass === inPass && dbActivo === 'SI') {
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          embarcacion: dbBarco,
          user: dbUser
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Usuario o contrasena incorrectos'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleGetInventario(embarcacion) {
  try {
    var sh = getInventariosSheet();
    var data = sh.getDataRange().getValues();
    var inventarios = [];
    var barcoUpper = String(embarcacion||'').toUpperCase();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][1]||'').toUpperCase() === barcoUpper) {
        inventarios.push({
          id: data[i][0],
          embarcacion: data[i][1],
          fecha: data[i][2],
          realizadoPor: data[i][3],
          items: data[i][4] ? JSON.parse(data[i][4]) : {},
          fotosUrls: data[i][5] ? JSON.parse(data[i][5]) : {},
          notas: data[i][6],
          fechaCreacion: data[i][7]
        });
      }
    }
    inventarios.sort(function(a,b){ return a.fechaCreacion > b.fechaCreacion ? -1 : 1; });
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      inventarios: inventarios
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleGetInventarios() {
  try {
    var sh = getInventariosSheet();
    var data = sh.getDataRange().getValues();
    var inventarios = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      inventarios.push({
        id: data[i][0],
        embarcacion: data[i][1],
        fecha: data[i][2],
        realizadoPor: data[i][3],
        items: data[i][4] ? JSON.parse(data[i][4]) : {},
        fotosUrls: data[i][5] ? JSON.parse(data[i][5]) : {},
        notas: data[i][6],
        fechaCreacion: data[i][7]
      });
    }
    inventarios.sort(function(a,b){ return a.fechaCreacion > b.fechaCreacion ? -1 : 1; });
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      inventarios: inventarios
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSaveInventario(invData) {
  try {
    var sh = getInventariosSheet();
    var id = invData.id || ('INV_' + new Date().getTime());
    var fotosUrls = Object.assign({}, invData.fotosUrls || {});
    var row = [
      id,
      invData.embarcacion || '',
      invData.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      invData.realizadoPor || '',
      JSON.stringify(invData.items || {}),
      JSON.stringify(fotosUrls),
      invData.notas || '',
      new Date().toISOString()
    ];
    var existing = sh.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] === id) {
        sh.getRange(i+1, 1, 1, row.length).setValues([row]);
        return ContentService.createTextOutput(JSON.stringify({success:true,id:id,updated:true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    sh.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({success:true,id:id,created:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error guardando inventario: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleUploadInventarioPhotoBatch(data) {
  try {
    var photos = data.photos || [];
    var invData = data.inventarioData || {};
    var boat = invData.embarcacion || 'DESCONOCIDO';
    var rootFolder = getRootFolder();
    var invFolders = rootFolder.getFoldersByName('Inventarios');
    var invFolder = invFolders.hasNext() ? invFolders.next() : rootFolder.createFolder('Inventarios');
    var cleanBoat = boat.toUpperCase().replace(/ /g, '_');
    var boatFolders = invFolder.getFoldersByName(cleanBoat);
    var boatFolder = boatFolders.hasNext() ? boatFolders.next() : invFolder.createFolder(cleanBoat);
    var results = {};
    photos.forEach(function(p) {
      try {
        var base64Clean = p.base64Data.split(',')[1];
        var ext = p.base64Data.indexOf('image/png') > -1 ? 'png' : 'jpg';
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Clean), 'image/'+ext, p.photoId+'.'+ext);
        var file = boatFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        results[p.photoId] = {success: true, url: file.getUrl(), fileId: file.getId()};
      } catch(e) {
        results[p.photoId] = {success: false, error: e.toString()};
      }
    });
    return ContentService.createTextOutput(JSON.stringify({success:true,results:results})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error subiendo fotos: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleGetInventarioAccesos() {
  try {
    var sh = getInventarioAccesosSheet();
    var data = sh.getDataRange().getValues();
    var accesos = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      accesos.push({
        embarcacion: data[i][0],
        usuario: data[i][1],
        activo: data[i][3]
      });
    }
    return ContentService.createTextOutput(JSON.stringify({success:true,accesos:accesos})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSaveInventarioAcceso(acceso) {
  try {
    var sh = getInventarioAccesosSheet();
    var data = sh.getDataRange().getValues();
    var newBarco = String(acceso.embarcacion||'').trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]||'').toLowerCase() === newBarco.toLowerCase()) {
        sh.getRange(i+1,1,1,4).setValues([[newBarco, acceso.usuario, acceso.password, acceso.activo||'SI']]);
        return ContentService.createTextOutput(JSON.stringify({success:true,updated:true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    sh.appendRow([newBarco, acceso.usuario, acceso.password, acceso.activo||'SI']);
    return ContentService.createTextOutput(JSON.stringify({success:true,created:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== GESTIÓN DE USUARIOS DEL PANEL ==========

function handleGetUsuarios() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(USERS_SHEET);
    var data = sh.getDataRange().getValues();
    var usuarios = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      usuarios.push({
        usuario: data[i][0],
        rol: data[i][2],
        activo: String(data[i][3]||'').toUpperCase()
      });
    }
    return ContentService.createTextOutput(JSON.stringify({success:true, usuarios:usuarios})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSaveUsuario(u) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(USERS_SHEET);
    var data = sh.getDataRange().getValues();
    var newUser = String(u.usuario||'').trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]||'').toLowerCase() === newUser.toLowerCase()) {
        var pass = u.password && u.password.trim() ? u.password.trim() : data[i][1];
        sh.getRange(i+1,1,1,4).setValues([[newUser, pass, u.rol||'Tecnico', u.activo||'SI']]);
        return ContentService.createTextOutput(JSON.stringify({success:true,updated:true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    if (!u.password || !u.password.trim()) {
      return ContentService.createTextOutput(JSON.stringify({success:false,message:'La contrasena es obligatoria para usuarios nuevos'})).setMimeType(ContentService.MimeType.JSON);
    }
    sh.appendRow([newUser, u.password.trim(), u.rol||'Tecnico', u.activo||'SI']);
    return ContentService.createTextOutput(JSON.stringify({success:true,created:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleToggleUsuario(username, activo) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(USERS_SHEET);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]||'').toLowerCase() === String(username||'').toLowerCase()) {
        sh.getRange(i+1, 4).setValue(activo);
        return ContentService.createTextOutput(JSON.stringify({success:true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false,message:'Usuario no encontrado'})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
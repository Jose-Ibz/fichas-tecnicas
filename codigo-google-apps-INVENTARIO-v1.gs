// ========================================
// SISTEMA VIAMAR - INVENTARIO DE BARCOS v1.0
// ========================================

var SHEET_ID = '1CkiLr9gDWJDg2eTpUVX9ZYLRlQUbBp8GNdMBoedU-a8'; // mismo Sheet que fichas tecnicas
var USERS_SHEET = 'Usuarios';
var INVENTARIOS_SHEET = 'Inventarios';
var INVENTARIO_ACCESOS_SHEET = 'InventarioAccesos';
var BARCOS_SHEET = 'Barcos';
var DRIVE_FOLDER_NAME = 'Inspecciones-Viamar'; // misma carpeta raiz en Drive

// ========== RUTAS GET ==========

function doGet(e) {
  try {
    var action = e.parameter.action;

    if (action === 'login') {
      return handleLogin(e.parameter.user || e.parameter.username, e.parameter.pass || e.parameter.password);
    } else if (action === 'getInventario') {
      return handleGetInventario(e.parameter.embarcacion);
    } else if (action === 'getInventarios') {
      return handleGetInventarios();
    } else if (action === 'getInventarioAccesos') {
      return handleGetInventarioAccesos();
    } else if (action === 'getBarcos') {
      return handleGetBarcos();
    } else if (action === 'getBarcosAll') {
      return handleGetBarcosAll();
    } else if (action === 'getIncidencias') {
      return handleGetIncidencias(e.parameter.resuelta);
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

// ========== RUTAS POST ==========

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    if (action === 'saveInventario') {
      return handleSaveInventario(data.data);
    } else if (action === 'uploadInventarioPhotoBatch') {
      return handleUploadInventarioPhotoBatch(data);
    } else if (action === 'saveInventarioAcceso') {
      return handleSaveInventarioAcceso(data.acceso);
    } else if (action === 'saveBarco') {
      return handleSaveBarco(data.nombre, data.activo);
    } else if (action === 'resolverIncidencia') {
      return handleResolverIncidencia(data.id, data.observaciones);
    } else if (action === 'deleteInventario') {
      return handleDeleteInventario(data.id);
    } else if (action === 'setAntihumedades') {
      return handleSetAntihumedades(data.id, data.valor, data.fecha);
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

// ========== LOGIN (misma hoja Usuarios que fichas tecnicas) ==========

function handleLogin(username, password) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(USERS_SHEET);
    var data = sh.getDataRange().getValues();
    var inUser = String(username||'').trim();
    var inPass = String(password||'').trim();
    for (var i = 1; i < data.length; i++) {
      var dbUser   = String(data[i][0]||'').trim();
      var dbPass   = String(data[i][1]||'').trim();
      var dbActivo = String(data[i][3]||'').trim().toUpperCase();
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
      message: 'Usuario o contrasena incorrectos'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== HOJAS ==========

function getInventariosSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(INVENTARIOS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(INVENTARIOS_SHEET);
    sh.appendRow(['ID','Embarcacion','Fecha','RealizadoPor','Items_JSON','FotosUrls_JSON','Notas','FechaCreacion','Antihumedades']);
    sh.getRange(1,1,1,9).setFontWeight('bold');
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
  }
  return sh;
}

// ========== INVENTARIOS ==========

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
          fecha: fmtFecha(data[i][2]),
          realizadoPor: data[i][3],
          items: data[i][4] ? JSON.parse(data[i][4]) : {},
          fotosUrls: data[i][5] ? JSON.parse(data[i][5]) : {},
          notas: data[i][6],
          fechaCreacion: fmtFecha(data[i][7]),
          antihumedades: data[i][8] === true || data[i][8] === 'TRUE' || data[i][8] === 'true',
          antihumedadesFecha: data[i][9] ? String(data[i][9]) : ''
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

function fmtFecha(v) {
  if (!v) return '';
  try {
    var d = v instanceof Date ? v : new Date(v);
    return isNaN(d.getTime()) ? String(v) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) { return String(v); }
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
        fecha: fmtFecha(data[i][2]),
        realizadoPor: data[i][3],
        items: data[i][4] ? JSON.parse(data[i][4]) : {},
        fotosUrls: data[i][5] ? JSON.parse(data[i][5]) : {},
        notas: data[i][6],
        fechaCreacion: fmtFecha(data[i][7]),
        antihumedades: data[i][8] === true || data[i][8] === 'TRUE' || data[i][8] === 'true',
        antihumedadesFecha: data[i][9] ? String(data[i][9]) : ''
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
      new Date().toISOString(),
      invData.antihumedades === true || invData.antihumedades === 'true' ? true : false,
      invData.antihumedadesFecha || ''
    ];
    var existing = sh.getDataRange().getValues();
    var updated = false;
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] === id) {
        sh.getRange(i+1, 1, 1, row.length).setValues([row]);
        updated = true;
        break;
      }
    }
    if (!updated) sh.appendRow(row);
    saveIncidenciasFromInventario(invData, id);
    return ContentService.createTextOutput(JSON.stringify({success:true, id:id, updated:updated, created:!updated})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error guardando inventario: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== FOTOS ==========

function getRootFolder() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
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

// ========== ACCESOS POR BARCO ==========

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

// ========== BARCOS ==========

function getBarcosSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(BARCOS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BARCOS_SHEET);
    sh.appendRow(['Nombre', 'Activo']);
    sh.getRange(1,1,1,2).setFontWeight('bold');
  }
  return sh;
}

function handleGetBarcos() {
  try {
    var sh = getBarcosSheet();
    var data = sh.getDataRange().getValues();
    var barcos = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (String(data[i][1]||'').toUpperCase() === 'SI') {
        barcos.push(String(data[i][0]).trim());
      }
    }
    barcos.sort();
    return ContentService.createTextOutput(JSON.stringify({success:true, barcos:barcos})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleGetBarcosAll() {
  try {
    var sh = getBarcosSheet();
    var data = sh.getDataRange().getValues();
    var barcos = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      barcos.push({nombre: String(data[i][0]).trim(), activo: String(data[i][1]||'SI').toUpperCase()});
    }
    return ContentService.createTextOutput(JSON.stringify({success:true, barcos:barcos})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSaveBarco(nombre, activo) {
  try {
    var sh = getBarcosSheet();
    var data = sh.getDataRange().getValues();
    var nombreClean = String(nombre||'').trim().toUpperCase();
    if (!nombreClean) return ContentService.createTextOutput(JSON.stringify({success:false,message:'Nombre requerido'})).setMimeType(ContentService.MimeType.JSON);
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]||'').toUpperCase() === nombreClean) {
        sh.getRange(i+1,1,1,2).setValues([[nombreClean, activo||'SI']]);
        return ContentService.createTextOutput(JSON.stringify({success:true,updated:true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    sh.appendRow([nombreClean, activo||'SI']);
    return ContentService.createTextOutput(JSON.stringify({success:true,created:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== INCIDENCIAS ==========

function getIncidenciasSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Incidencias');
  if (!sh) {
    sh = ss.insertSheet('Incidencias');
    sh.appendRow(['ID','Embarcacion','Descripcion','Fecha','Resuelta','Observaciones','FechaCreacion','RealizadoPor']);
    sh.getRange(1,1,1,8).setFontWeight('bold');
  }
  return sh;
}

function saveIncidenciasFromInventario(invData, invId) {
  if (!invData.incidencias || invData.incidencias.length === 0) return;
  var sh = getIncidenciasSheet();
  var existing = sh.getDataRange().getValues();
  var existingIds = {};
  for (var j = 1; j < existing.length; j++) {
    if (existing[j][0]) existingIds[String(existing[j][0])] = true;
  }
  invData.incidencias.forEach(function(inc) {
    if (!inc.descripcion) return;
    var incId = String(inc.id || ('INC_' + new Date().getTime()));
    if (!existingIds[incId]) {
      sh.appendRow([incId, invData.embarcacion||'', inc.descripcion||'', invData.fecha||'', 'NO', '', new Date().toISOString(), invData.realizadoPor||'']);
      existingIds[incId] = true;
    }
  });
}

function handleGetIncidencias(resuelta) {
  try {
    var sh = getIncidenciasSheet();
    var data = sh.getDataRange().getValues();
    var incidencias = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var esResuelta = String(data[i][4]||'NO').toUpperCase();
      if (resuelta && resuelta.toUpperCase() !== esResuelta) continue;
      incidencias.push({
        id: data[i][0],
        embarcacion: data[i][1],
        descripcion: data[i][2],
        fecha: data[i][3],
        resuelta: esResuelta,
        observaciones: data[i][5]||'',
        fechaCreacion: data[i][6],
        realizadoPor: data[i][7]||''
      });
    }
    incidencias.sort(function(a,b){ return a.fechaCreacion > b.fechaCreacion ? -1 : 1; });
    return ContentService.createTextOutput(JSON.stringify({success:true, incidencias:incidencias})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Error: '+error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleResolverIncidencia(id, observaciones) {
  try {
    var sh = getIncidenciasSheet();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sh.getRange(i+1, 5, 1, 2).setValues([['SI', observaciones||'']]);
        return ContentService.createTextOutput(JSON.stringify({success:true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Incidencia no encontrada'})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Error: '+error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteInventario(id) {
  try {
    var sh = getInventariosSheet();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sh.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({success:true})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Inventario no encontrado'})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Error: '+error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSetAntihumedades(id, valor, fecha) {
  try {
    var sh = getInventariosSheet();
    var data = sh.getDataRange().getValues();
    var hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var aplicado = valor === true || valor === 'true';
        sh.getRange(i+1, 9).setValue(aplicado);
        sh.getRange(i+1, 10).setValue(aplicado ? (fecha || hoy) : '');
        return ContentService.createTextOutput(JSON.stringify({success:true, fecha: aplicado ? (fecha || hoy) : ''})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Inventario no encontrado'})).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message:'Error: '+error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== SETUP (ejecutar manualmente una vez) ==========

function autorizarPermisos() {
  SpreadsheetApp.openById(SHEET_ID);
  DriveApp.getRootFolder();
  Logger.log('Permisos autorizados correctamente');
}

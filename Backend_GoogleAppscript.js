/**
 * =====================================================================
 * FUNDACION - BASE DE DATOS (Google Sheets) + API Web App
 * =====================================================================
 *
 * COMO USARLO:
 *  1. Pega TODO este codigo en Apps Script.
 *  2. Guarda (icono de disquete).
 *  3. En el selector de funciones (arriba) elige: inicializarHojas
 *  4. Click en "Ejecutar" y autoriza los permisos que pida.
 *     -> Crea/repara las hojas: Usuarios, Donaciones, Feed, Foro, Metricas.
 *     -> En "Ejecuciones" / "Registros" veras el enlace de tu hoja.
 *  5. Luego: Desplegar > Nueva implementacion > Aplicacion web
 *       Ejecutar como: Yo  |  Acceso: Cualquier persona
 *  6. Copia la URL del deployment y pasamela para conectarla al frontend.
 *
 * NOTA: Funciona tanto si pegas esto desde una hoja existente
 *       (Extensiones > Apps Script) como si es un proyecto nuevo de
 *       Apps Script (en ese caso crea la hoja automaticamente).
 * =====================================================================
 */

// =====================================================================
// CONFIGURACION
// =====================================================================
var CONFIG = {
  HOJA_USUARIOS:   'Usuarios',
  HOJA_DONACIONES: 'Donaciones',
  HOJA_FEED:       'Feed',
  HOJA_FORO:       'Foro',
  HOJA_METRICAS:   'Metricas',
  HOJA_VOLUNTARIOS:'Voluntarios',
  HOJA_NOTIFICACIONES: 'Notificaciones',
  HOJA_ACTIVIDADES:  'Actividades',

  // Roles validos.
  ROLES: ['usuario', 'voluntario', 'lider', 'admin'],
  ROL_POR_DEFECTO: 'usuario',

  ADMIN_INICIAL: {
    nombre:   'Administrador',
    email:    'admin@fundacion.org',
    password: 'admin2026'
  }
};

// La hoja Usuarios tiene 6 columnas; la 6a (Contrasena) NUNCA se expone via doGet.
var COLS_USUARIOS_PUBLICAS = 5;

// =====================================================================
// INICIALIZACION DE HOJAS (ejecutar manualmente una vez)
// =====================================================================
function inicializarHojas() {
  var ss = obtenerLibro_();

  var definiciones = [
    { nombre: CONFIG.HOJA_USUARIOS,       headers: ['UserID', 'Nombre', 'Email', 'Rol', 'FechaRegistro', 'Contraseña'] },
    { nombre: CONFIG.HOJA_DONACIONES,     headers: ['Nombre', 'Email', 'Monto', 'Metodo', 'Fecha', 'Mensaje', 'UserID', 'Estado', 'Comprobante'] },
    { nombre: CONFIG.HOJA_FEED,           headers: ['Titulo', 'Descripcion', 'ImagenURL', 'Fecha', 'Likes', 'LikedBy'] },
    { nombre: CONFIG.HOJA_FORO,           headers: ['Autor', 'Mensaje', 'Fecha', 'UserID', 'AvatarInitial', 'ParentID'] },
    { nombre: CONFIG.HOJA_METRICAS,       headers: ['Clave', 'Valor'] },
    { nombre: CONFIG.HOJA_VOLUNTARIOS,    headers: ['Fecha', 'Nombre', 'Email', 'Telefono', 'Habilidades', 'Disponibilidad', 'Mensaje', 'Estado'] },
    { nombre: CONFIG.HOJA_NOTIFICACIONES, headers: ['NotificacionID', 'UserEmail', 'Mensaje', 'Leido', 'Fecha', 'Tipo', 'Enlace'] },
    { nombre: CONFIG.HOJA_ACTIVIDADES,    headers: ['ActividadID', 'Titulo', 'Descripcion', 'Fecha', 'Asistentes', 'Estado', 'Autor'] }
  ];

  definiciones.forEach(function(def) {
    var hoja = ss.getSheetByName(def.nombre);
    if (!hoja) hoja = ss.insertSheet(def.nombre);

    // Escribe encabezados y estilos SOLO si la fila 1 esta vacia (no pisa datos ni formatos existentes).
    var primera = hoja.getRange(1, 1).getValue();
    if (primera === '' || primera === null) {
      hoja.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      
      // Estilo del encabezado inicial.
      hoja.getRange(1, 1, 1, def.headers.length)
          .setFontWeight('bold')
          .setBackground('#0f766e')
          .setFontColor('#ffffff');
      hoja.setFrozenRows(1);
    }
  });

  // Metricas base que usa el frontend.
  sembrarMetrica_('Visitas', 0);
  sembrarMetrica_('BalanceFinanzas', 0);

  // Crea el admin inicial (si ese correo aun no existe).
  crearAdminInicial_();

  // Limpia la hoja por defecto vacia ("Hoja 1" / "Sheet1") si quedo sin uso.
  ['Hoja 1', 'Hoja1', 'Sheet1'].forEach(function(n) {
    var h = ss.getSheetByName(n);
    if (h && ss.getSheets().length > 1 && h.getDataRange().isBlank()) {
      try { ss.deleteSheet(h); } catch (e) {}
    }
  });

  // Crea la carpeta en Google Drive para subir imagenes (si no existe)
  try {
    obtenerCrearCarpetaDrive_();
  } catch (e) {
    Logger.log('Aviso: No se pudo crear la carpeta en Drive. Asegurese de conceder permisos de Drive.');
  }

  Logger.log('Inicializacion completada.');
  Logger.log('Base de datos: ' + ss.getUrl());
  try { ss.toast('Hojas inicializadas correctamente', 'Listo', 5); } catch (e) {}
}

// =====================================================================
// API - LECTURA (GET)
// =====================================================================
function doGet(e) {
  try {
    var nombreHoja = (e && e.parameter) ? e.parameter.sheet : null;
    if (!nombreHoja) return json_({ error: 'Falta el parametro sheet' });

    var hoja = obtenerHoja_(nombreHoja);
    if (!hoja) return json_({ error: 'Hoja no encontrada: ' + nombreHoja });

    var valores = hoja.getDataRange().getValues();
    valores.shift(); // quitar encabezados

    // Seguridad: nunca devolver la columna Contrasena de Usuarios.
    if (nombreHoja === CONFIG.HOJA_USUARIOS) {
      valores = valores.map(function(fila) { return fila.slice(0, COLS_USUARIOS_PUBLICAS); });
    }

    return json_({ data: valores });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

// =====================================================================
// API - ESCRITURA (POST)
// =====================================================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var accion = body.action;

    // --- Acciones de autenticacion (operan sobre la hoja Usuarios) ---
    if (accion === 'registerUser')     return registrarUsuario_(body);
    if (accion === 'loginUser')        return autenticarUsuario_(body);
    if (accion === 'setRole')          return asignarRol_(body);
    if (accion === 'syncFirebaseUser') return sincronizarFirebaseUser_(body);
    if (accion === 'marcarNotificacionLeida') return marcarNotificacionLeida_(body);

    // --- Subida de imagenes a Drive ---
    if (accion === 'uploadImage')  return subirImagen_(body);

    // --- Asistencia a actividades ---
    if (accion === 'asistir') {
      var hojaAct = obtenerHoja_(CONFIG.HOJA_ACTIVIDADES);
      if (!hojaAct) return json_({ success: false, error: 'Hoja Actividades no encontrada' });
      var actID = body.actividadID;
      
      var datosAct = hojaAct.getDataRange().getValues();
      for (var i = 1; i < datosAct.length; i++) {
        if (String(datosAct[i][0]) === String(actID)) {
           var asisRaw = String(datosAct[i][4] || '');
           var asisArr = [];
           if (asisRaw) {
             try { asisArr = JSON.parse(asisRaw); } catch(e) { asisArr = []; }
           }
           if (!Array.isArray(asisArr)) asisArr = [];
           
           // Validar si ya existe
           for (var j=0; j<asisArr.length; j++) {
             if (asisArr[j].uid === body.userID) {
               return json_({ success: false, error: 'Ya estás registrado para esta actividad.' });
             }
           }
           
           asisArr.push({ uid: body.userID, name: body.userName, email: body.userEmail });
           hojaAct.getRange(i + 1, 5).setValue(JSON.stringify(asisArr));
           
           return json_({ success: true });
        }
      }
      return json_({ success: false, error: 'Actividad no encontrada' });
    }

    // --- Acciones genericas existentes del frontend ---
    var hoja = obtenerHoja_(body.sheet);
    if (!hoja) return json_({ success: false, error: 'Hoja no encontrada: ' + body.sheet });

    if (accion === 'add' || accion === 'register') {
      hoja.appendRow(body.row);
      
      // NUEVO: Enviar correo si es Voluntario
      if (body.sheet === CONFIG.HOJA_VOLUNTARIOS) {
          try { enviarCorreoVoluntario_('registro', body.row); } catch(e) {}
      }
      // NUEVO: Correo si es Donación Confirmada
      if (body.sheet === 'Donaciones' && body.row[7] === 'Confirmada') {
          try { enviarCorreoNotificacion_(body.row[0], body.row[1], 'donacion', body.row[2]); } catch(e) {}
      }

      return json_({ success: true });
    }

    if (accion === 'delete') {
      var fila = body.rowIndex + 2; // +2 por encabezado y base 0
      if (fila <= hoja.getLastRow()) hoja.deleteRow(fila);
      return json_({ success: true });
    }

    if (accion === 'updateRow') {
      var fila = body.rowIndex + 2;
      hoja.getRange(fila, 1, 1, body.row.length).setValues([body.row]);

      // NUEVO: Enviar correo si es Voluntario y fue aceptado
      if (body.sheet === CONFIG.HOJA_VOLUNTARIOS) {
          try {
              if (body.row[7] === 'Aceptado') {
                  enviarCorreoVoluntario_('aceptado', body.row);
              }
          } catch(e) {}
      }
      
      // NUEVO: Correo si la donación pasa a Confirmada
      if (body.sheet === 'Donaciones') {
          try {
              if (body.row[7] === 'Confirmada') {
                  enviarCorreoNotificacion_(body.row[0], body.row[1], 'donacion', body.row[2]);
              }
          } catch(e) {}
      }

      return json_({ success: true });
    }

    if (accion === 'updateLike') {
      var f = body.rowIndex + 2;
      hoja.getRange(f, 5).setValue(body.likes);   // col 5 = Likes
      hoja.getRange(f, 6).setValue(body.likedBy); // col 6 = LikedBy
      return json_({ success: true });
    }

    if (accion === 'increment') {
      var datos = hoja.getDataRange().getValues();
      for (var i = 1; i < datos.length; i++) {
        if (datos[i][0] === body.key) {
          hoja.getRange(i + 1, 2).setValue((parseInt(datos[i][1]) || 0) + 1);
          break;
        }
      }
      return json_({ success: true });
    }

    if (accion === 'updateSettings') {
      var hojaMetricas = obtenerHoja_(CONFIG.HOJA_METRICAS);
      if (!hojaMetricas) return json_({ success: false, error: 'Hoja Metricas no encontrada' });
      var datos = hojaMetricas.getDataRange().getValues();
      var dict = body.settings; // { "Banco": "...", "NIT": "..." }
      
      for (var key in dict) {
        var found = false;
        for (var i = 1; i < datos.length; i++) {
          if (datos[i][0] === key) {
            hojaMetricas.getRange(i + 1, 2).setValue(dict[key]);
            found = true;
            break;
          }
        }
        if (!found) {
          hojaMetricas.appendRow([key, dict[key]]);
        }
      }
      return json_({ success: true });
    }

    return json_({ success: false, error: 'Accion no reconocida: ' + accion });
  } catch (err) {
    return json_({ success: false, error: String(err) });
  }
}

// =====================================================================
// AUTENTICACION (email + contrasena con hash SHA-256)
// =====================================================================

/** Registra un usuario nuevo. body: { nombre, email, password, rol? } */
function registrarUsuario_(body) {
  var hoja = obtenerHoja_(CONFIG.HOJA_USUARIOS);
  if (!hoja) return json_({ success: false, error: 'Ejecuta inicializarHojas primero' });

  var nombre = String(body.nombre || body.name || '').trim();
  var email  = String(body.email || '').trim().toLowerCase();
  var pass   = String(body.password || body.contrasena || '');
  var rol    = normalizarRol_(body.rol || body.role);

  if (!email || !pass) return json_({ success: false, error: 'EMAIL_Y_CONTRASENA_REQUERIDOS' });

  // Verifica email duplicado (col indice 2 = Email).
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][2]).trim().toLowerCase() === email) {
      return json_({ success: false, error: 'EMAIL_EXISTS' });
    }
  }

  var uid = generarUserID_();
  hoja.appendRow([uid, nombre, email, rol, fechaHoy_(), hashPassword_(pass)]);

  try { enviarCorreoNotificacion_(nombre, email, 'bienvenida'); } catch(e) {}

  return json_({ success: true, user: { uid: uid, name: nombre, email: email, rol: rol } });
}

/** Inicia sesion. body: { email, password } */
function autenticarUsuario_(body) {
  var hoja = obtenerHoja_(CONFIG.HOJA_USUARIOS);
  if (!hoja) return json_({ success: false, error: 'Ejecuta inicializarHojas primero' });

  var email = String(body.email || '').trim().toLowerCase();
  var pass  = String(body.password || body.contrasena || '');
  if (!email || !pass) return json_({ success: false, error: 'EMAIL_Y_CONTRASENA_REQUERIDOS' });

  var hash = hashPassword_(pass);
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (String(fila[2]).trim().toLowerCase() === email) {
      if (String(fila[5]) === hash) {
        return json_({ success: true, user: { uid: fila[0], name: fila[1], email: fila[2], rol: fila[3] } });
      }
      return json_({ success: false, error: 'INVALID_CREDENTIALS' });
    }
  }
  return json_({ success: false, error: 'INVALID_CREDENTIALS' });
}

/** Cambia el rol de un usuario (para tu panel admin). body: { email, rol } */
function asignarRol_(body) {
  var hoja = obtenerHoja_(CONFIG.HOJA_USUARIOS);
  if (!hoja) return json_({ success: false, error: 'Ejecuta inicializarHojas primero' });

  var email = String(body.email || '').trim().toLowerCase();
  var rol   = normalizarRol_(body.rol || body.role);

  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][2]).trim().toLowerCase() === email) {
      hoja.getRange(i + 1, 4).setValue(rol); // col 4 = Rol
      return json_({ success: true, email: email, rol: rol });
    }
  }
  return json_({ success: false, error: 'USER_NOT_FOUND' });
}

/** Sincroniza usuario de Google Firebase con la hoja Usuarios. Crea si no existe. body: { email, nombre, uid } */
function sincronizarFirebaseUser_(body) {
  var hoja = obtenerHoja_(CONFIG.HOJA_USUARIOS);
  if (!hoja) return json_({ success: false, error: 'Ejecuta inicializarHojas primero' });

  var email = String(body.email || '').trim().toLowerCase();
  var nombre = String(body.nombre || body.name || email).trim();
  var firebaseUid = String(body.uid || generarUserID_());
  
  if (!email) return json_({ success: false, error: 'EMAIL_REQUIRED' });

  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][2]).trim().toLowerCase() === email) {
      // Existe
      var user = { uid: datos[i][0], name: datos[i][1], email: datos[i][2], rol: datos[i][3] };
      return json_({ success: true, user: user, isNew: false });
    }
  }

  // No existe, crearlo
  var rol = CONFIG.ROL_POR_DEFECTO;
  hoja.appendRow([firebaseUid, nombre, email, rol, fechaHoy_(), 'GOOGLE_AUTH']);
  
  try { enviarCorreoNotificacion_(nombre, email, 'bienvenida'); } catch(e) {}
  
  return json_({ success: true, user: { uid: firebaseUid, name: nombre, email: email, rol: rol }, isNew: true });
}

function marcarNotificacionLeida_(body) {
  var hoja = obtenerHoja_(CONFIG.HOJA_NOTIFICACIONES);
  if (!hoja) return json_({ success: false });
  var id = body.notificacionID;
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(id)) {
      hoja.getRange(i + 1, 4).setValue(true); // Col 4 es 'Leido'
      return json_({ success: true });
    }
  }
  return json_({ success: false });
}

// =====================================================================
// HELPERS
// =====================================================================

/** Devuelve el libro activo (bound) o crea/reutiliza uno si es standalone. */
function obtenerLibro_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  var nuevo = SpreadsheetApp.create('Base de Datos - Fundacion');
  props.setProperty('SPREADSHEET_ID', nuevo.getId());
  Logger.log('Nuevo libro creado: ' + nuevo.getUrl());
  return nuevo;
}

function obtenerHoja_(nombre) {
  return obtenerLibro_().getSheetByName(nombre);
}

function sembrarMetrica_(clave, valor) {
  var hoja = obtenerHoja_(CONFIG.HOJA_METRICAS);
  if (!hoja) return;
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === clave) return; // ya existe
  }
  hoja.appendRow([clave, valor]);
}

function crearAdminInicial_() {
  var hoja = obtenerHoja_(CONFIG.HOJA_USUARIOS);
  if (!hoja) return;

  var email = String(CONFIG.ADMIN_INICIAL.email).trim().toLowerCase();
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][2]).trim().toLowerCase() === email) return; // ya existe
  }
  hoja.appendRow([
    generarUserID_(),
    CONFIG.ADMIN_INICIAL.nombre,
    email,
    'admin',
    fechaHoy_(),
    hashPassword_(CONFIG.ADMIN_INICIAL.password)
  ]);
}

function normalizarRol_(rol) {
  rol = String(rol || '').trim().toLowerCase();
  return CONFIG.ROLES.indexOf(rol) !== -1 ? rol : CONFIG.ROL_POR_DEFECTO;
}

function generarUserID_() {
  return 'usr_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

function fechaHoy_() {
  return new Date().toLocaleDateString('es-CO');
}

/** Hash SHA-256 (hex) para no guardar contrasenas en texto plano. */
function hashPassword_(plain) {
  if (!plain) return '';
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(plain),
    Utilities.Charset.UTF_8
  );
  return raw.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

// =====================================================================
// INTEGRACION CON GOOGLE DRIVE (Subida de Imagenes)
// =====================================================================

function obtenerCrearCarpetaDrive_() {
  var folderName = 'Fundacion_Images';
  var folders = DriveApp.getFoldersByName(folderName);
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return folder;
}

function subirImagen_(body) {
  try {
    var base64Data = body.base64;
    var fileName = body.fileName || ('img_' + new Date().getTime() + '.png');
    var mimeType = body.mimeType || 'image/png';

    if (!base64Data) return json_({ success: false, error: 'No se recibio la imagen' });

    var base64Content = base64Data.split(',')[1] || base64Data;
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, fileName);

    var folder = obtenerCrearCarpetaDrive_();
    var file = folder.createFile(blob);

    // Hacer el archivo publico para que pueda visualizarse en el html
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // En Apps Script nativo, file no tiene getWebContentLink(). 
    // Usamos el ID para crear un link directo para etiquetas <img>
    var fileUrl = 'https://drive.google.com/uc?id=' + file.getId();

    return json_({ success: true, url: fileUrl });
  } catch (err) {
    return json_({ success: false, error: 'Error al subir la imagen: ' + String(err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// ENVÍO DE CORREOS
// =====================================================================
function enviarCorreoVoluntario_(tipo, rowData) {
  var nombre = rowData[1] || 'Voluntario';
  var email = rowData[2];
  
  if (!email || email.indexOf('@') === -1) return;
  
  var asunto = "";
  var mensajeHtml = "";
  
  var baseStyle = "font-family: Arial, sans-serif; color: #333; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #eaeaea; border-radius: 8px;";
  var headerStyle = "text-align: center; border-bottom: 2px solid #0f766e; padding-bottom: 10px; margin-bottom: 20px;";
  var titleStyle = "color: #0f766e; margin: 0; font-size: 24px;";
  var footerStyle = "margin-top: 30px; font-size: 0.85em; color: #777; text-align: center; border-top: 1px solid #eaeaea; padding-top: 15px;";
  
  if (tipo === 'registro') {
    asunto = "Tu inscripción al voluntariado está en revisión - Fundación GAP";
    mensajeHtml = "<div style='" + baseStyle + "'>" +
                  "<div style='" + headerStyle + "'><h2 style='" + titleStyle + "'>¡Gracias por unirte a nuestra causa!</h2></div>" +
                  "<p>Hola <b>" + nombre + "</b>,</p>" +
                  "<p>Tu inscripción al sistema de voluntariado de la fundación ha sido recibida correctamente y <b>está en revisión</b>.</p>" +
                  "<p>Nuestro equipo evaluará tu solicitud y te avisaremos muy pronto cuando hayas sido aceptado.</p>" +
                  "<p>¡Valoramos inmensamente tu disposición para ayudar y construir un futuro lleno de esperanza!</p>" +
                  "<div style='" + footerStyle + "'><i>Por favor no responder a este correo, ya que es un mensaje automático que no necesita respuesta.</i></div>" +
                  "</div>";
  } else if (tipo === 'aceptado') {
    asunto = "¡Has sido aceptado como Voluntario! - Fundación GAP";
    mensajeHtml = "<div style='" + baseStyle + "'>" +
                  "<div style='" + headerStyle + "'><h2 style='" + titleStyle + "'>¡Bienvenido al equipo!</h2></div>" +
                  "<p>Hola <b>" + nombre + "</b>,</p>" +
                  "<p>Nos alegra mucho informarte que <b>has sido aceptado</b> en el sistema de voluntariados de la Fundación GAP.</p>" +
                  "<p>Te invitamos a revisar tu perfil en la fundación. Pronto implementaremos nuevas funciones donde podrás ver tu espacio personal con estadísticas de cuántos voluntariados y donaciones has realizado.</p>" +
                  "<p>¡Gracias por ser parte del cambio y hacer la diferencia!</p>" +
                  "<div style='" + footerStyle + "'><i>Por favor no responder a este correo, ya que es un mensaje automático que no necesita respuesta.</i></div>" +
                  "</div>";
  }
  
  if (asunto && mensajeHtml) {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: mensajeHtml,
      noReply: true
    });
  }
}

function enviarCorreoNotificacion_(nombre, email, tipo, monto) {
  if (!email || email.indexOf('@') === -1) return;
  
  var asunto = "";
  var mensajeHtml = "";
  
  var baseStyle = "font-family: Arial, sans-serif; color: #333; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #eaeaea; border-radius: 8px;";
  var headerStyle = "text-align: center; border-bottom: 2px solid #0f766e; padding-bottom: 10px; margin-bottom: 20px;";
  var titleStyle = "color: #0f766e; margin: 0; font-size: 24px;";
  var footerStyle = "margin-top: 30px; font-size: 0.85em; color: #777; text-align: center; border-top: 1px solid #eaeaea; padding-top: 15px;";
  
  if (tipo === 'bienvenida') {
    asunto = "¡Bienvenido a la Fundación GAP!";
    mensajeHtml = "<div style='" + baseStyle + "'>" +
                  "<div style='" + headerStyle + "'><h2 style='" + titleStyle + "'>¡Bienvenido a nuestra comunidad!</h2></div>" +
                  "<p>Hola <b>" + nombre + "</b>,</p>" +
                  "<p>Nos alegra mucho darte la bienvenida a la sede electrónica de la <b>Fundación GAP</b>. Tu cuenta ha sido creada exitosamente.</p>" +
                  "<p>Nuestra misión es impulsar sueños y proteger derechos. Como parte de esta plataforma, podrás involucrarte en nuestros proyectos, asistir a voluntariados, contribuir con donaciones y ser un miembro activo de nuestra comunidad.</p>" +
                  "<p>¡Gracias por sumarte a nosotros y por hacer la diferencia!</p>" +
                  "<div style='" + footerStyle + "'><i>Por favor no responder a este correo, ya que es un mensaje automático que no necesita respuesta.</i></div>" +
                  "</div>";
  } else if (tipo === 'donacion') {
    asunto = "¡Gracias por tu donación! - Fundación GAP";
    mensajeHtml = "<div style='" + baseStyle + "'>" +
                  "<div style='" + headerStyle + "'><h2 style='" + titleStyle + "'>¡Tu donación ha sido confirmada!</h2></div>" +
                  "<p>Hola <b>" + nombre + "</b>,</p>" +
                  "<p>Queremos expresarte nuestro más sincero agradecimiento. Hemos recibido y confirmado tu generosa donación" + (monto ? " por valor de <b>$" + monto + "</b>" : "") + ".</p>" +
                  "<p>Gracias a tu invaluable apoyo, podemos seguir impulsando proyectos que transforman vidas y protegen los derechos de quienes más lo necesitan. ¡Cada aporte es un pilar fundamental para construir esperanza!</p>" +
                  "<p>¡Gracias por tu enorme corazón solidario y por hacer la diferencia!</p>" +
                  "<div style='" + footerStyle + "'><i>Por favor no responder a este correo, ya que es un mensaje automático que no necesita respuesta.</i></div>" +
                  "</div>";
  }
  
  if (asunto && mensajeHtml) {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: mensajeHtml,
      noReply: true
    });
  }
}
// ============================================================
// CUADRE DE CAJA - SISTEMA PRINCIPAL
// Google Apps Script - Code.gs
// ============================================================

const MASTER_SHEET_NAME = "MAESTRO";
const SHEET_USUARIOS = "Usuarios";
const SHEET_LOCALES = "Locales";
const SHEET_RECURSOS = "Recursos";
const SHEET_ASIGNACIONES = "Asignaciones";
const SHEET_REGISTROS = "Registros_Diarios";
const SHEET_CIERRES = "Cierres_Periodo";
const SHEET_CONFIG = "Configuracion";
const SHEET_TAREAS = "Tareas";
const SHEET_TAREAS_PAGOS = "Tareas_Pagos";
const SHEET_EGRESOS = "Egresos";
const SHEET_TRANSFERENCIAS = "Transferencias";
const SHEET_DEUDAS = "Deudas";
const SHEET_DEUDAS_PAGOS = "Deudas_Pagos";

// ============================================================
// DIAGNÓSTICO — ping simple para verificar conectividad GAS
// ============================================================

function ping(token) {
  // No usa SpreadsheetApp ni PropertiesService — solo verifica que GAS responde
  try {
    const ts = new Date().toISOString();
    if (!token) return { ok: true, ts, msg: 'sin token' };

    // Verificar sesión de forma ligera
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('sesion_' + token);
    return {
      ok: true,
      ts,
      sesionValida: !!raw,
      msg: raw ? 'sesion_ok' : 'sesion_no_encontrada'
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// PUNTO DE ENTRADA WEB APP
// ============================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Cuadre de Caja")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// INICIALIZACIÓN DEL SISTEMA
// ============================================================

function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Crear pestañas si no existen
  const hojas = [
    SHEET_USUARIOS, SHEET_LOCALES, SHEET_RECURSOS,
    SHEET_ASIGNACIONES, SHEET_REGISTROS, SHEET_CIERRES,
    SHEET_CONFIG, SHEET_TAREAS, SHEET_TAREAS_PAGOS,
    SHEET_EGRESOS, SHEET_TRANSFERENCIAS,
    SHEET_DEUDAS, SHEET_DEUDAS_PAGOS
  ];

  hojas.forEach(nombre => {
    if (!ss.getSheetByName(nombre)) {
      ss.insertSheet(nombre);
    }
  });

  // Configurar cabeceras
  _setupUsuarios(ss);
  _setupLocales(ss);
  _setupRecursos(ss);
  _setupAsignaciones(ss);
  _setupRegistros(ss);
  _setupCierres(ss);
  _setupConfig(ss);
  _setupTareas(ss);
  _setupTareasPagos(ss);
  _setupEgresos(ss);
  _setupTransferencias(ss);
  _setupDeudas(ss);
  _setupDeudasPagos(ss);

  // Crear admin por defecto si no existe
  _crearAdminDefault(ss);

  return { ok: true, mensaje: "Sistema inicializado correctamente" };
}

function _setupUsuarios(ss) {
  const h = ss.getSheetByName(SHEET_USUARIOS);
  if (h.getLastRow() === 0) {
    h.appendRow(["ID", "Nombre", "Usuario", "PIN_Hash", "Rol", "Local_ID", "Activo", "Fecha_Creacion"]);
    h.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupLocales(ss) {
  const h = ss.getSheetByName(SHEET_LOCALES);
  if (h.getLastRow() === 0) {
    h.appendRow(["ID", "Nombre", "Activo", "Fecha_Creacion", "Sheet_ID"]);
    h.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupRecursos(ss) {
  const h = ss.getSheetByName(SHEET_RECURSOS);
  if (h.getLastRow() === 0) {
    h.appendRow(["ID", "Nombre", "Tipo", "Activo", "Fecha_Creacion"]);
    h.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupAsignaciones(ss) {
  const h = ss.getSheetByName(SHEET_ASIGNACIONES);
  if (h.getLastRow() === 0) {
    h.appendRow(["ID", "Usuario_ID", "Recurso_ID", "Activo", "Fecha_Asignacion"]);
    h.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupRegistros(ss) {
  const h = ss.getSheetByName(SHEET_REGISTROS);
  if (h.getLastRow() === 0) {
    h.appendRow([
      "ID", "Fecha", "Usuario_ID", "Usuario_Nombre", "Local_ID",
      "Recurso_ID", "Recurso_Nombre", "Tipo", // 'apertura' o 'cierre'
      "Valor", "Ganancia_Dia", "Periodo", "Timestamp"
    ]);
    h.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupCierres(ss) {
  const h = ss.getSheetByName(SHEET_CIERRES);
  if (h.getLastRow() === 0) {
    h.appendRow([
      "ID", "Periodo", "Fecha_Inicio", "Fecha_Fin", "Fecha_Cierre",
      "Usuario_ID", "Usuario_Nombre", "Local_ID",
      "Recurso_ID", "Recurso_Nombre",
      "Saldo_Inicial_Periodo", "Saldo_Final_Periodo",
      "Ganancia_Total_Periodo", "Saldo_Verificado_Admin"
    ]);
    h.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupConfig(ss) {
  const h = ss.getSheetByName(SHEET_CONFIG);
  if (h.getLastRow() === 0) {
    h.appendRow(["Clave", "Valor"]);
    h.appendRow(["version", "1.0"]);
    h.appendRow(["inicializado", new Date().toISOString()]);
    h.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
  }
}

function _crearAdminDefault(ss) {
  const h = ss.getSheetByName(SHEET_USUARIOS);
  const datos = h.getDataRange().getValues();
  const existeAdmin = datos.some(row => row[2] === "admin");

  if (!existeAdmin) {
    const id = _generarID();
    const pinHash = _hashPin("1234");
    h.appendRow([id, "Administrador", "admin", pinHash, "admin", "", true, new Date()]);
  }
}

// ============================================================
// AUTENTICACIÓN
// ============================================================

function login(usuario, pin) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_USUARIOS);
    const datos = h.getDataRange().getValues();
    const pinHash = _hashPin(pin);

    for (let i = 1; i < datos.length; i++) {
      const row = datos[i];
      const id = row[0];
      const nombre = row[1];
      const user = row[2];
      const hashGuardado = row[3];
      const rol = row[4];
      const localId = row[5];
      const activo = row[6];

      if (user === usuario && hashGuardado === pinHash && activo === true) {
        // Crear token de sesión
        const token = _generarToken();
        const props = PropertiesService.getScriptProperties();
        const sesionData = JSON.stringify({
          token, id, nombre, usuario: user, rol, localId,
          expira: Date.now() + (8 * 60 * 60 * 1000) // 8 horas
        });
        props.setProperty("sesion_" + token, sesionData);

        return {
          ok: true,
          token,
          usuario: { id, nombre, usuario: user, rol, localId }
        };
      }
    }

    return { ok: false, error: "Usuario o PIN incorrecto" };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function verificarSesion(token) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty("sesion_" + token);
    if (!raw) return { ok: false, error: "Sesión inválida" };

    const sesion = JSON.parse(raw);
    if (Date.now() > sesion.expira) {
      props.deleteProperty("sesion_" + token);
      return { ok: false, error: "Sesión expirada" };
    }

    return { ok: true, usuario: sesion };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function logout(token) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("sesion_" + token);
  return { ok: true };
}

// ============================================================
// GESTIÓN DE LOCALES (ADMIN)
// ============================================================

function crearLocal(token, nombre) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_LOCALES);
    const id = _generarID();

    h.appendRow([id, nombre.trim(), true, new Date(), ""]);

    return { ok: true, id, nombre: nombre.trim() };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function obtenerLocales(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_LOCALES);
  const datos = h.getDataRange().getValues();

  const locales = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2] === true) {
      locales.push({ id: datos[i][0], nombre: datos[i][1] });
    }
  }

  return { ok: true, locales };
}

function eliminarLocal(token, localId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_LOCALES);
  const datos = h.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === localId) {
      h.getRange(i + 1, 3).setValue(false);
      return { ok: true };
    }
  }

  return { ok: false, error: "Local no encontrado" };
}

// ============================================================
// GESTIÓN DE RECURSOS (ADMIN)
// ============================================================

function crearRecurso(token, nombre, tipo) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_RECURSOS);
    const id = _generarID();

    h.appendRow([id, nombre.trim(), tipo, true, new Date()]);

    return { ok: true, id, nombre: nombre.trim(), tipo };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function obtenerRecursos(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_RECURSOS);
  const datos = h.getDataRange().getValues();

  const recursos = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][3] === true) {
      recursos.push({ id: datos[i][0], nombre: datos[i][1], tipo: datos[i][2] });
    }
  }

  return { ok: true, recursos };
}

function eliminarRecurso(token, recursoId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_RECURSOS);
  const datos = h.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === recursoId) {
      h.getRange(i + 1, 4).setValue(false);
      return { ok: true };
    }
  }

  return { ok: false, error: "Recurso no encontrado" };
}

// ============================================================
// GESTIÓN DE USUARIOS (ADMIN)
// ============================================================

function crearUsuario(token, datos) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_USUARIOS);
    const existentes = h.getDataRange().getValues();

    // Verificar que el usuario no exista
    const existe = existentes.some(row => row[2] === datos.usuario);
    if (existe) return { ok: false, error: "Ese nombre de usuario ya existe" };

    const id = _generarID();
    const pinHash = _hashPin(datos.pin);

    const rolesValidos = ["cajero", "disenador"];
    const rol = rolesValidos.includes(datos.rol) ? datos.rol : "cajero";
    h.appendRow([id, datos.nombre.trim(), datos.usuario.trim(), pinHash, rol, datos.localId || "", true, new Date()]);

    return { ok: true, id };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function obtenerUsuarios(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_USUARIOS);
  const datos = h.getDataRange().getValues();
  const locales = _getLocalesMap(ss);

  const usuarios = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][4] !== "admin" && datos[i][6] === true) {
      usuarios.push({
        id: datos[i][0],
        nombre: datos[i][1],
        usuario: datos[i][2],
        rol: datos[i][4],
        localId: datos[i][5],
        localNombre: locales[datos[i][5]] || "Sin local"
      });
    }
  }

  return { ok: true, usuarios };
}

function eliminarUsuario(token, usuarioId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_USUARIOS);
  const datos = h.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === usuarioId) {
      h.getRange(i + 1, 7).setValue(false);
      return { ok: true };
    }
  }

  return { ok: false, error: "Usuario no encontrado" };
}

function cambiarPin(token, usuarioId, nuevoPIN) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  // Solo admin o el propio usuario puede cambiar el PIN
  if (sesion.usuario.rol !== "admin" && sesion.usuario.id !== usuarioId) {
    return { ok: false, error: "No autorizado" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_USUARIOS);
  const datos = h.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === usuarioId) {
      h.getRange(i + 1, 4).setValue(_hashPin(nuevoPIN));
      return { ok: true };
    }
  }

  return { ok: false, error: "Usuario no encontrado" };
}

// ============================================================
// GESTIÓN DE ASIGNACIONES (ADMIN)
// ============================================================

function asignarRecurso(token, usuarioId, recursoId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_ASIGNACIONES);
    const datos = h.getDataRange().getValues();

    // Verificar que no exista ya
    const existe = datos.some(row => row[1] === usuarioId && row[2] === recursoId && row[3] === true);
    if (existe) return { ok: false, error: "Ya tiene ese recurso asignado" };

    const id = _generarID();
    h.appendRow([id, usuarioId, recursoId, true, new Date()]);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function quitarAsignacion(token, usuarioId, recursoId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(SHEET_ASIGNACIONES);
  const datos = h.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === usuarioId && datos[i][2] === recursoId && datos[i][3] === true) {
      h.getRange(i + 1, 4).setValue(false);
      return { ok: true };
    }
  }

  return { ok: false, error: "Asignación no encontrada" };
}

// Registrar cupos iniciales: genera un registro de apertura con el cupo dado
// para que el cajero empiece el día con ese saldo como referencia
function registrarCuposIniciales(token, usuarioId, cupos) {
  // cupos = [{ recursoId, recursoNombre, cupo }]
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hRegistros = ss.getSheetByName(SHEET_REGISTROS);
    const hUsuarios = ss.getSheetByName(SHEET_USUARIOS);

    const usuarios = hUsuarios.getDataRange().getValues();
    const usuarioRow = usuarios.find(r => r[0] === usuarioId);
    if (!usuarioRow) return { ok: false, error: "Usuario no encontrado" };

    const usuarioNombre = usuarioRow[1];
    const localId = usuarioRow[5];
    const periodo = _getPeriodoActual();
    const timestamp = new Date();

    // Usamos una fecha especial "CUPO_INICIAL" para distinguirlo
    // de una apertura normal del día. Guardamos como apertura del día de hoy
    // solo si no existe ya una apertura hoy para ese recurso.
    const hoy = _fechaHoy();
    const registrosExistentes = hRegistros.getDataRange().getValues();

    cupos.forEach(c => {
      // Verificar si ya hay un registro de cupo inicial para este usuario+recurso hoy
      const yaExiste = registrosExistentes.some(r =>
        r[2] === usuarioId && r[5] === c.recursoId &&
        r[7] === "cupo_inicial"
      );

      if (!yaExiste) {
        const id = _generarID();
        hRegistros.appendRow([
          id, hoy, usuarioId, usuarioNombre, localId,
          c.recursoId, c.recursoNombre, "cupo_inicial",
          c.cupo, 0, periodo, timestamp
        ]);
      } else {
        // Actualizar el cupo existente
        const datos = hRegistros.getDataRange().getValues();
        for (let i = 1; i < datos.length; i++) {
          if (datos[i][2] === usuarioId && datos[i][5] === c.recursoId && datos[i][7] === "cupo_inicial") {
            hRegistros.getRange(i + 1, 9).setValue(c.cupo);
            break;
          }
        }
      }
    });

    return { ok: true, mensaje: "Cupos iniciales registrados" };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function obtenerRecursosDeUsuario(token, usuarioId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  try {
    const uid = usuarioId || sesion.usuario.id;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hAsig = ss.getSheetByName(SHEET_ASIGNACIONES);
    const hRec = ss.getSheetByName(SHEET_RECURSOS);

    if (!hAsig || !hRec) return { ok: true, recursos: [] };

    const asignaciones = hAsig.getDataRange().getValues();
    const recursos = hRec.getDataRange().getValues();

    const recursoMap = {};
    for (let i = 1; i < recursos.length; i++) {
      recursoMap[recursos[i][0]] = { id: recursos[i][0], nombre: recursos[i][1], tipo: recursos[i][2] };
    }

    const resultado = [];
    for (let i = 1; i < asignaciones.length; i++) {
      if (asignaciones[i][1] === uid && asignaciones[i][3] === true) {
        const rec = recursoMap[asignaciones[i][2]];
        if (rec) resultado.push(rec);
      }
    }

    return { ok: true, recursos: resultado };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// REGISTRO DE CAJA DIARIO (CAJERO)
// ============================================================

function registrarApertura(token, valores) {
  // valores = [{ recursoId, recursoNombre, valor }]
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  // Solo cajero o admin pueden registrar
  if (sesion.usuario.rol !== "cajero" && sesion.usuario.rol !== "admin") {
    return { ok: false, error: "Solo cajero o admin pueden registrar aperturas" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_REGISTROS);
    const hoy = _fechaHoy();
    const periodo = _getPeriodoActual();
    const timestamp = new Date();

    // Verificar si ya existe apertura hoy
    const existentes = h.getDataRange().getValues();
    const yaApertura = existentes.some(r =>
      _fechaToStr(r[1]) === hoy && r[2] === sesion.usuario.id && r[7] === "apertura"
    );

    if (yaApertura) {
      return { ok: false, error: "Ya registraste la apertura de hoy" };
    }

    valores.forEach(v => {
      const id = _generarID();
      h.appendRow([
        id, hoy, sesion.usuario.id, sesion.usuario.nombre, sesion.usuario.localId,
        v.recursoId, v.recursoNombre, "apertura",
        v.valor, 0, periodo, timestamp
      ]);
    });

    return { ok: true, mensaje: "Apertura registrada correctamente" };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function registrarCierre(token, valores) {
  // valores = [{ recursoId, recursoNombre, valor }]
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  if (sesion.usuario.rol !== "cajero" && sesion.usuario.rol !== "admin") {
    return { ok: false, error: "Solo cajero o admin pueden registrar cierres" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_REGISTROS);
    const hoy = _fechaHoy();
    const periodo = _getPeriodoActual();
    const timestamp = new Date();

    // Verificar si ya existe cierre hoy
    const existentes = h.getDataRange().getValues();
    const yaCierre = existentes.some(r =>
      _fechaToStr(r[1]) === hoy && r[2] === sesion.usuario.id && r[7] === "cierre"
    );

    if (yaCierre) {
      return { ok: false, error: "Ya registraste el cierre de hoy" };
    }

    // Buscar aperturas de hoy para calcular ganancia
    const aperturasHoy = existentes.filter(r =>
      _fechaToStr(r[1]) === hoy && r[2] === sesion.usuario.id && r[7] === "apertura"
    );

    const aperturaMap = {};
    aperturasHoy.forEach(a => {
      aperturaMap[a[5]] = Number(a[8]);
    });

    valores.forEach(v => {
      const id = _generarID();
      const valorApertura = aperturaMap[v.recursoId] || 0;
      const ganancia = v.valor - valorApertura;

      h.appendRow([
        id, hoy, sesion.usuario.id, sesion.usuario.nombre, sesion.usuario.localId,
        v.recursoId, v.recursoNombre, "cierre",
        v.valor, ganancia, periodo, timestamp
      ]);
    });

    return { ok: true, mensaje: "Cierre registrado correctamente" };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function obtenerEstadoDia(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_REGISTROS);
    const hoy = _fechaHoy();
    const datos = h.getDataRange().getValues();

    const registrosHoy = datos.filter(r =>
      _fechaToStr(r[1]) === hoy && r[2] === sesion.usuario.id
    );

    const tieneApertura = registrosHoy.some(r => r[7] === "apertura");
    const tieneCierre = registrosHoy.some(r => r[7] === "cierre");

    let estado = "sin_apertura";
    if (tieneApertura && !tieneCierre) estado = "abierto";
    if (tieneApertura && tieneCierre) estado = "cerrado";

    // Obtener últimos cierres por recurso (para mostrar valores iniciales)
    const ultimosCierres = _getUltimosCierresPorRecurso(datos, sesion.usuario.id);

    return {
      ok: true,
      fecha: hoy,
      estado,
      ultimosCierres
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function obtenerHistorialCierres30d(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_REGISTROS);
    const datos = h.getDataRange().getValues().slice(1);

    // Últimos 30 días
    const hace30 = new Date();
    hace30.setDate(hace30.getDate() - 30);

    const cierres = datos
      .filter(r => {
        if (r[2] !== sesion.usuario.id) return false;
        if (r[7] !== "cierre") return false;
        const fecha = new Date(r[1]);
        return fecha >= hace30;
      })
      .map(r => ({
        fecha: _fechaToStr(r[1]),
        recursoId: r[5],
        recursoNombre: r[6],
        valor: Number(r[8]),
        ganancia: Number(r[9])
      }))
      .reverse();

    return { ok: true, cierres };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// DASHBOARD ADMIN
// ============================================================

function obtenerDashboardAdmin(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hUsuarios = ss.getSheetByName(SHEET_USUARIOS);
    const hRegistros = ss.getSheetByName(SHEET_REGISTROS);
    const hLocales = ss.getSheetByName(SHEET_LOCALES);

    const usuarios = hUsuarios.getDataRange().getValues().slice(1);
    const registros = hRegistros.getDataRange().getValues().slice(1);
    const localesData = hLocales.getDataRange().getValues().slice(1);

    const hoy = _fechaHoy();
    const periodo = _getPeriodoActual();

    // Mapa de locales
    const localesMap = {};
    localesData.forEach(l => { localesMap[l[0]] = l[1]; });

    // Estado por cajero
    const cajeros = usuarios
      .filter(u => u[4] === "cajero" && u[6] === true)
      .map(u => {
        const uid = u[0];
        const regsHoy = registros.filter(r => _fechaToStr(r[1]) === hoy && r[2] === uid);
        const tieneApertura = regsHoy.some(r => r[7] === "apertura");
        const tieneCierre = regsHoy.some(r => r[7] === "cierre");

        let estado = "sin_apertura";
        if (tieneApertura && !tieneCierre) estado = "abierto";
        if (tieneApertura && tieneCierre) estado = "cerrado";

        // Ganancia del día
        const gananciaDia = regsHoy
          .filter(r => r[7] === "cierre")
          .reduce((acc, r) => acc + Number(r[9]), 0);

        // Ganancia del periodo
        const gananciaPeriodo = registros
          .filter(r => r[2] === uid && r[7] === "cierre" && r[10] === periodo)
          .reduce((acc, r) => acc + Number(r[9]), 0);

        return {
          id: uid,
          nombre: u[1],
          localId: u[5],
          localNombre: localesMap[u[5]] || "Sin local",
          estado,
          gananciaDia,
          gananciaPeriodo
        };
      });

    // Resumen por local
const resumenLocales = localesData
  .filter(l => l[2] === true)
  .map(l => {
    const cajsLocal = cajeros.filter(c => c.localId === l[0]);
    const cajasCerradas = cajsLocal.filter(c => c.estado === "cerrado").length;
    return {
      id:             l[0],
      nombre:         l[1],
      usuariosActivos: cajsLocal.length,
      cajasAbiertas:  cajsLocal.filter(c => c.estado === "abierto").length,
      cajasCerradas,
      totalHoy:       cajsLocal.reduce((a, c) => a + c.gananciaDia, 0),
      totalPeriodo:   cajsLocal.reduce((a, c) => a + c.gananciaPeriodo, 0)
    };
  });

// Ganancia global
const gananciaGlobalPeriodo = cajeros.reduce((a, c) => a + c.gananciaPeriodo, 0);
const gananciaGlobalHoy     = cajeros.reduce((a, c) => a + c.gananciaDia,     0);

return {
  ok: true,
  hoy,           // antes era "fecha: hoy"
  periodo,
  estadoUsuarios: cajeros,       // renombrado
  resumenLocales,                // renombrado
  gananciaGlobalHoy,             // nuevo campo
  gananciaGlobalPeriodo
};
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// CIERRE QUINCENAL (ADMIN)
// ============================================================

function obtenerResumenQuincenal(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  const hRegistros = ss.getSheetByName(SHEET_REGISTROS);
  const hRecursos = ss.getSheetByName(SHEET_RECURSOS);
  const hAsig = ss.getSheetByName(SHEET_ASIGNACIONES);
  const hLocales = ss.getSheetByName(SHEET_LOCALES);

  const usuarios = hUsuarios.getDataRange().getValues().slice(1);
  const registros = hRegistros.getDataRange().getValues();
  const recursos = hRecursos.getDataRange().getValues().slice(1);
  const asignaciones = hAsig.getDataRange().getValues().slice(1);
  const localesData = hLocales.getDataRange().getValues().slice(1);

  const periodo = _getPeriodoActual();
  const localesMap = {};
  localesData.forEach(l => { localesMap[l[0]] = l[1]; });

  const recursoMap = {};
  recursos.forEach(r => { recursoMap[r[0]] = r[1]; });

  // Por cada cajero + recurso asignado
  const resumen = usuarios
    .filter(u => u[4] === "cajero" && u[6] === true)
    .map(u => {
    const userId = u[0];

    // Recursos asignados a este cajero
    const recIds = asignaciones
      .filter(a => a[1] === userId && a[3] === true)
      .map(a => a[2]);

    const detalle = recIds.map(recId => {
      // Ganancia total del periodo para este recurso
      const gananciaTotal = registros.slice(1)
        .filter(r => r[2] === userId && r[5] === recId && r[7] === "cierre" && r[10] === periodo)
        .reduce((acc, r) => acc + Number(r[9]), 0);

      // Saldo actual = último cierre
      const ultimoCierre = registros.slice(1)
        .filter(r => r[2] === userId && r[5] === recId && r[7] === "cierre")
        .sort((a, b) => new Date(b[11]) - new Date(a[11]))[0];

      const saldoActual = ultimoCierre ? Number(ultimoCierre[8]) : 0;

      // Saldo inicial del periodo = primer apertura del periodo
      const primeraApertura = registros
        .filter(r => r[2] === userId && r[5] === recId && r[7] === "apertura" && r[10] === periodo)
        .sort((a, b) => new Date(a[11]) - new Date(b[11]))[0];

      const saldoInicial = primeraApertura ? Number(primeraApertura[8]) : 0;

      return {
        recursoId: recId,
        recursoNombre: recursoMap[recId] || recId,
        saldoInicial,
        saldoActual,
        gananciaTotal
      };
    });

    return {
      usuarioId: userId,
      usuarioNombre: u[1],
      localId: u[5],
      localNombre: localesMap[u[5]] || "Sin local",
      detalle
    };
  });

  const fechas = _getFechasPeriodo(periodo);

  return {
    ok: true,
    periodo,
    fechaInicio: fechas.inicio,
    fechaFin: fechas.fin,
    resumen
  };
}

function ejecutarCierreQuincenal(token, verificaciones) {
  // verificaciones = [{ usuarioId, recursoId, saldoVerificado }]
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  const periodo = _getPeriodoActual();
  const fechas = _getFechasPeriodo(periodo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hCierres = ss.getSheetByName(SHEET_CIERRES);
  const hRegistros = ss.getSheetByName(SHEET_REGISTROS);
  const hUsuarios = ss.getSheetByName(SHEET_USUARIOS);

  const registros = hRegistros.getDataRange().getValues().slice(1);
  const usuarios = hUsuarios.getDataRange().getValues().slice(1);
  const fechaCierre = new Date();

  verificaciones.forEach(v => {
    const regsPeriodo = registros.filter(r =>
      r[2] === v.usuarioId && r[5] === v.recursoId && r[7] === "cierre" && r[10] === periodo
    );

    const gananciaTotal = regsPeriodo.reduce((acc, r) => acc + Number(r[9]), 0);

    const primeraApertura = registros
      .filter(r => r[2] === v.usuarioId && r[5] === v.recursoId && r[7] === "apertura" && r[10] === periodo)
      .sort((a, b) => new Date(a[11]) - new Date(b[11]))[0];

    const saldoInicial = primeraApertura ? Number(primeraApertura[8]) : 0;
    const usuarioNombre = usuarios.find(u => u[0] === v.usuarioId)?.[1] || "";
    const localId = usuarios.find(u => u[0] === v.usuarioId)?.[5] || "";

    const id = _generarID();
    hCierres.appendRow([
      id, periodo, fechas.inicio, fechas.fin, fechaCierre,
      v.usuarioId, usuarioNombre, localId,
      v.recursoId, v.recursoNombre,
      saldoInicial, v.saldoVerificado,
      gananciaTotal, v.saldoVerificado
    ]);
  });

  return {
    ok: true,
    mensaje: `Cierre del periodo ${periodo} ejecutado correctamente`,
    periodo
  };
}

function obtenerHistorialCierres(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_CIERRES);
    if (!h) return { ok: true, cierres: [] };

    const datos = h.getDataRange().getValues().slice(1);

    const cierres = datos.reverse().map(row => ({
      periodo: String(row[1] || ''),
      // ✅ FIX: convertir fechas a string para que GAS pueda serializarlas
      fechaInicio:   _fechaToStr(row[2]),
      fechaFin:      _fechaToStr(row[3]),
      fechaCierre:   _fechaToStr(row[4]),
      usuarioNombre: String(row[6] || ''),
      localId:       String(row[7] || ''),
      recursoNombre: String(row[9] || ''),
      saldoInicial:  Number(row[10]) || 0,
      saldoFinal:    Number(row[11]) || 0,
      gananciaTotal: Number(row[12]) || 0,
      saldoVerificado: Number(row[13]) || 0
    }));

    return { ok: true, cierres };
  } catch(e) {
    return { ok: false, error: "obtenerHistorialCierres: " + e.message };
  }
}

// ============================================================
// GESTIÓN DE TAREAS (TASK MANAGER) - MEJORADO
// ============================================================
// Nota: SHEET_TAREAS y SHEET_TAREAS_PAGOS declarados al inicio del archivo

function _setupTareas(ss) {
  const h = ss.getSheetByName(SHEET_TAREAS);
  if (h && h.getLastRow() === 0) {
    // ACTUALIZADO: añadido Creado_Por_Nombre (col 13)
    h.appendRow(["ID", "Titulo", "Cliente", "Precio", "Estado", "Descripcion",
                 "Fecha_Entrega", "Creado_Por", "Local_ID",
                 "Fecha_Creacion", "Fecha_Actualizacion", "Fecha_Entregado", "Creado_Por_Nombre"]);
    h.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupTareasPagos(ss) {
  const h = ss.getSheetByName(SHEET_TAREAS_PAGOS);
  if (h && h.getLastRow() === 0) {
    h.appendRow(["ID", "Tarea_ID", "Monto", "Fecha_Pago", "Notas", "Creado_Por", "Fecha_Creacion"]);
    h.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

/**
 * Crear tarea (cajero, diseñador y admin pueden crear)
 * MEJORA #1: Diseñador puede crear tareas
 * MEJORA #2: Se guarda el nombre del creador
 * MEJORA #5: Opción de abono inicial
 * @param {string} token - Token de sesión
 * @param {Object} datos - Datos de la tarea
 * @param {string} datos.titulo - Título de la tarea
 * @param {string} datos.cliente - Nombre del cliente
 * @param {number} [datos.precio] - Precio de la tarea
 * @param {string} [datos.fechaEntrega] - Fecha de entrega (YYYY-MM-DD)
 * @param {string} [datos.descripcion] - Descripción del trabajo
 * @param {number} [datos.abonoInicial] - Abono inicial opcional
 * @returns {Object} Resultado de la operación
 */
function crearTarea(token, datos) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  // MEJORA #1: cajero, disenador y admin pueden crear tareas
  const rolesPermitidos = ["cajero", "disenador", "admin"];
  if (!rolesPermitidos.includes(sesion.usuario.rol)) {
    return { ok: false, error: "No tienes permiso para crear tareas" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Crear hoja si no existe
    if (!ss.getSheetByName(SHEET_TAREAS)) {
      ss.insertSheet(SHEET_TAREAS);
      _setupTareas(ss);
    }
    if (!ss.getSheetByName(SHEET_TAREAS_PAGOS)) {
      ss.insertSheet(SHEET_TAREAS_PAGOS);
      _setupTareasPagos(ss);
    }

    const h = ss.getSheetByName(SHEET_TAREAS);
    const id = _generarID();
    const now = new Date();

    // MEJORA #2: Guardar nombre del creador
    h.appendRow([
      id,
      datos.titulo.trim(),
      datos.cliente.trim(),
      Number(datos.precio) || 0,
      "cotizacion",
      datos.descripcion || "",
      datos.fechaEntrega || "",
      sesion.usuario.id,
      sesion.usuario.localId || "",
      now,
      now,
      "", // Fecha_Entregado
      sesion.usuario.nombre // NUEVO: Creado_Por_Nombre
    ]);

    // MEJORA #5: Si hay abono inicial, registrarlo
    if (datos.abonoInicial && Number(datos.abonoInicial) > 0) {
      const hPagos = ss.getSheetByName(SHEET_TAREAS_PAGOS);
      const pagoId = _generarID();
      hPagos.appendRow([
        pagoId,
        id, // Tarea_ID
        Number(datos.abonoInicial),
        _fechaHoy(),
        "Abono inicial",
        sesion.usuario.id,
        now
      ]);
    }

    return { ok: true, id };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

/**
 * Obtener tareas — admin/diseñador ven todas; cajero ve las suyas; archivadas separadas
 * MEJORA #2: Incluye nombre del creador
 * @param {string} token - Token de sesión
 * @param {Object} [filtros] - Filtros opcionales
 * @returns {Object} Lista de tareas
 */
function obtenerTareas(token, filtros) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TAREAS);
    if (!h) return { ok: true, tareas: [] };

    const datos = h.getDataRange().getValues().slice(1);
    const hPagos = ss.getSheetByName(SHEET_TAREAS_PAGOS);
    const pagos = hPagos ? hPagos.getDataRange().getValues().slice(1) : [];
    
    // Obtener mapa de usuarios para nombres
    const hUsuarios = ss.getSheetByName(SHEET_USUARIOS);
    const usuariosData = hUsuarios ? hUsuarios.getDataRange().getValues().slice(1) : [];
    const usuariosMap = {};
    usuariosData.forEach(u => { usuariosMap[u[0]] = u[1]; });
    
    const esAdmin     = sesion.usuario.rol === "admin";
    const esDisenador = sesion.usuario.rol === "disenador";
    const verArchivadas = filtros && filtros.archivadas === true;

    // Umbral de archivado: terminado hace más de 2 días
    const ahora = new Date();
    const DOS_DIAS_MS = 2 * 24 * 60 * 60 * 1000;

    let tareas = datos.filter(r => {
      if (!r[0]) return false;

      // Lógica de archivado
      const esEtregado = r[4] === "entregado";  // archivado 2 días después de entregado
      const fechaEntregado = r[11] ? new Date(r[11]) : null;
      const debeArchivar = esEtregado && fechaEntregado &&
                           (ahora - fechaEntregado) > DOS_DIAS_MS;

      if (verArchivadas) return debeArchivar;
      if (debeArchivar) return false; // excluir de vista normal

      // Filtro de visibilidad por rol
      if (esAdmin || esDisenador) return true;
      return r[7] === sesion.usuario.id; // cajero ve solo las suyas
    });

    if (filtros && filtros.estado) {
      tareas = tareas.filter(r => r[4] === filtros.estado);
    }

    return {
      ok: true,
      tareas: tareas.map(r => {
        const tareaId = r[0];
        const pagosT  = pagos.filter(p => p[1] === tareaId);
        const totalPagado = pagosT.reduce((a, p) => a + Number(p[2]), 0);
        
        // MEJORA #2: Nombre del creador (col 12 = índice 12, o buscar en mapa)
        const creadoPorNombre = r[12] || usuariosMap[r[7]] || "Desconocido";
        
        return {
          id: tareaId,
          titulo:       String(r[1] || ''),
          cliente:      String(r[2] || ''),
          precio:       Number(r[3]),
          estado:       String(r[4] || ''),
          descripcion:  String(r[5] || ''),
          fechaEntrega: _fechaToStr(r[6]),
          creadoPorId:  String(r[7] || ''),
          creadoPorNombre: creadoPorNombre, // NUEVO
          localId:      String(r[8] || ''),
          fechaCreacion: _fechaToStr(r[9]),
          fechaEntregado: _fechaToStr(r[11]),
          totalPagado,
          saldo: Number(r[3]) - totalPagado,
          cantPagos: pagosT.length
        };
      }).reverse()
    };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

/**
 * Actualizar estado de tarea
 * MEJORA #3: Para Drag & Drop
 * @param {string} token - Token de sesión
 * @param {string} tareaId - ID de la tarea
 * @param {string} nuevoEstado - Nuevo estado
 * @returns {Object} Resultado de la operación
 */
function actualizarEstadoTarea(token, tareaId, nuevoEstado) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  const esCajero    = sesion.usuario.rol === "cajero";
  const estadosRestringidos = ["terminado", "entregado", "archivado"];

  // Cajero no puede poner terminado/entregado/archivado
  if (esCajero && estadosRestringidos.includes(nuevoEstado)) {
    return { ok: false, error: "Solo el admin o diseñador puede usar este estado" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TAREAS);
    if (!h) return { ok: false, error: "No existe la hoja de tareas" };

    const datos = h.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === tareaId) {
        if (esCajero && datos[i][7] !== sesion.usuario.id) {
          return { ok: false, error: "No puedes modificar tareas de otro cajero" };
        }
        h.getRange(i + 1, 5).setValue(nuevoEstado);
        h.getRange(i + 1, 11).setValue(new Date());
        // Col 11 = Fecha_Actualizacion, Col 12 = Fecha_Entregado (para auto-archivado)
        if (nuevoEstado === "entregado") {
          h.getRange(i + 1, 12).setValue(new Date());
        }
        // Si se cambia de entregado a otro estado, limpiar fecha entregado
        if (nuevoEstado !== "entregado") {
          h.getRange(i + 1, 12).setValue("");
        }
        return { ok: true };
      }
    }
    return { ok: false, error: "Tarea no encontrada" };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

/**
 * Mover tarea a nuevo estado (para Drag & Drop)
 * MEJORA #3: Función simplificada para D&D
 * @param {string} token - Token de sesión
 * @param {string} tareaId - ID de la tarea
 * @param {string} nuevoEstado - Nuevo estado
 * @returns {Object} Resultado de la operación
 */
function moverTarea(token, tareaId, nuevoEstado) {
  return actualizarEstadoTarea(token, tareaId, nuevoEstado);
}

// Agregar pago a tarea (solo admin)
function agregarPagoTarea(token, tareaId, monto, fecha, notas) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo el admin puede registrar pagos" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEET_TAREAS_PAGOS)) {
      ss.insertSheet(SHEET_TAREAS_PAGOS);
      _setupTareasPagos(ss);
    }
    const h = ss.getSheetByName(SHEET_TAREAS_PAGOS);
    const id = _generarID();
    h.appendRow([id, tareaId, Number(monto), fecha || _fechaHoy(), notas || "", sesion.usuario.id, new Date()]);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Obtener pagos de una tarea
function obtenerPagosTarea(token, tareaId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TAREAS_PAGOS);
    if (!h) return { ok: true, pagos: [] };

    const datos = h.getDataRange().getValues().slice(1);
    const pagos = datos
      .filter(r => r[1] === tareaId)
      .map(r => ({
        id: r[0],
        monto: Number(r[2]),
        fecha: _fechaToStr(r[3]),
        notas: r[4]
      }));

    return { ok: true, pagos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Eliminar tarea (solo admin)
function eliminarTarea(token, tareaId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo el admin puede eliminar tareas" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TAREAS);
    if (!h) return { ok: false, error: "No existe la hoja" };

    const datos = h.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === tareaId) {
        h.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: "Tarea no encontrada" };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

/**
 * MEJORA #6: Obtener estadísticas de tareas para el dashboard admin
 * @param {string} token - Token de sesión
 * @returns {Object} Estadísticas de tareas
 */
function obtenerEstadisticasTareas(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TAREAS);
    if (!h) return { ok: true, stats: { total: 0, porEstado: {}, porVencer: 0, vencidas: 0 } };

    const datos = h.getDataRange().getValues().slice(1).filter(r => r[0]);
    const hPagos = ss.getSheetByName(SHEET_TAREAS_PAGOS);
    const pagos = hPagos ? hPagos.getDataRange().getValues().slice(1) : [];

    const ahora = new Date();
    const TRES_DIAS_MS = 3 * 24 * 60 * 60 * 1000;
    const DOS_DIAS_MS = 2 * 24 * 60 * 60 * 1000;

    // Filtrar solo tareas activas (no archivadas)
    const tareasActivas = datos.filter(r => {
      const esEntregado = r[4] === "entregado";
      const fechaEntregado = r[11] ? new Date(r[11]) : null;
      const debeArchivar = esEntregado && fechaEntregado && (ahora - fechaEntregado) > DOS_DIAS_MS;
      return !debeArchivar;
    });

    // Contar por estado
    const porEstado = {
      cotizacion: 0,
      en_proceso: 0,
      terminado: 0,
      entregado: 0
    };

    let porVencer = 0; // Próximos 3 días
    let vencidas = 0;  // Ya pasó la fecha
    let totalPendienteCobro = 0;

    tareasActivas.forEach(r => {
      const estado = r[4];
      if (porEstado[estado] !== undefined) {
        porEstado[estado]++;
      }

      // Verificar fecha de entrega (solo si no está entregado)
      if (estado !== "entregado" && r[6]) {
        const fechaEntrega = new Date(r[6]);
        const diff = fechaEntrega - ahora;

        if (diff < 0) {
          vencidas++;
        } else if (diff <= TRES_DIAS_MS) {
          porVencer++;
        }
      }

      // Calcular pendiente de cobro
      const tareaId = r[0];
      const precio = Number(r[3]) || 0;
      const pagosT = pagos.filter(p => p[1] === tareaId);
      const totalPagado = pagosT.reduce((a, p) => a + Number(p[2]), 0);
      const saldo = precio - totalPagado;
      if (saldo > 0) {
        totalPendienteCobro += saldo;
      }
    });

    return {
      ok: true,
      stats: {
        total: tareasActivas.length,
        porEstado,
        porVencer,
        vencidas,
        totalPendienteCobro
      }
    };
  } catch(e) {
    return { ok: false, error: "obtenerEstadisticasTareas: " + e.message };
  }
}


// ============================================================
// SETUP NUEVAS HOJAS
// ============================================================

function _setupEgresos(ss) {
  const h = ss.getSheetByName(SHEET_EGRESOS);
  if (h && h.getLastRow() === 0) {
    h.appendRow(["ID","Fecha","Usuario_ID","Usuario_Nombre","Local_ID",
                 "Monto","Concepto","Solicitante","Tipo",
                 "Es_Deuda","Estado","Aprobado_Por","Fecha_Aprobacion","Notas"]);
    h.getRange(1,1,1,14).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupTransferencias(ss) {
  const h = ss.getSheetByName(SHEET_TRANSFERENCIAS);
  if (h && h.getLastRow() === 0) {
    h.appendRow(["ID","Fecha","Local_Origen_ID","Local_Origen_Nombre",
                 "Local_Destino_ID","Local_Destino_Nombre",
                 "Cajero_Origen_ID","Cajero_Origen_Nombre",
                 "Cajero_Destino_ID","Cajero_Destino_Nombre",
                 "Monto","Motivo","Estado","Aprobado_Por",
                 "Fecha_Aprobacion","Fecha_Confirmacion","Notas"]);
    h.getRange(1,1,1,17).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupDeudas(ss) {
  const h = ss.getSheetByName(SHEET_DEUDAS);
  if (h && h.getLastRow() === 0) {
    h.appendRow(["ID","Fecha","Deudor_Nombre","Deudor_Contacto",
                 "Concepto","Monto_Total","Total_Abonado","Saldo_Pendiente",
                 "Estado","Registrado_Por","Notas","Fecha_Vencimiento"]);
    h.getRange(1,1,1,12).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

function _setupDeudasPagos(ss) {
  const h = ss.getSheetByName(SHEET_DEUDAS_PAGOS);
  if (h && h.getLastRow() === 0) {
    h.appendRow(["ID","Deuda_ID","Monto","Fecha_Pago","Metodo",
                 "Notas","Registrado_Por","Timestamp"]);
    h.getRange(1,1,1,8).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
}

// ============================================================
// MÓDULO 1 — EGRESOS DE CAJA
// ============================================================

function registrarEgreso(token, datos) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };
  if (sesion.usuario.rol === "disenador") return { ok: false, error: "No autorizado" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEET_EGRESOS)) { ss.insertSheet(SHEET_EGRESOS); _setupEgresos(ss); }
    const h = ss.getSheetByName(SHEET_EGRESOS);
    const id = _generarID();
    h.appendRow([
      id, new Date(), sesion.usuario.id, sesion.usuario.nombre, sesion.usuario.localId,
      Number(datos.monto), datos.concepto.trim(), datos.solicitante.trim(),
      datos.esDeuda ? "deuda" : "egreso", datos.esDeuda === true,
      "pendiente", "", "", datos.notas || ""
    ]);
    return { ok: true, id, msg: "Registrado. Pendiente de aprobación del administrador." };
  } catch(e) { return { ok: false, error: "registrarEgreso: " + e.message }; }
}

function obtenerEgresosDia(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_EGRESOS);
    if (!h) return { ok: true, egresos: [] };
    const hoy = _fechaHoy();
    const egresos = h.getDataRange().getValues().slice(1)
      .filter(r => _fechaToStr(r[1]) === hoy && r[2] === sesion.usuario.id)
      .map(r => ({
        id: String(r[0]), fecha: _fechaToStr(r[1]), monto: Number(r[5]),
        concepto: String(r[6]), solicitante: String(r[7]), tipo: String(r[8]),
        esDeuda: r[9] === true, estado: String(r[10])
      }));
    return { ok: true, egresos };
  } catch(e) { return { ok: false, error: "obtenerEgresosDia: " + e.message }; }
}

function obtenerEgresosPendientes(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_EGRESOS);
    if (!h) return { ok: true, egresos: [] };
    const egresos = h.getDataRange().getValues().slice(1)
      .filter(r => r[10] === "pendiente")
      .map(r => ({
        id: String(r[0]), fecha: _fechaToStr(r[1]),
        usuarioNombre: String(r[3]), localId: String(r[4]),
        monto: Number(r[5]), concepto: String(r[6]),
        solicitante: String(r[7]), tipo: String(r[8]),
        esDeuda: r[9] === true, estado: String(r[10])
      }));
    return { ok: true, egresos };
  } catch(e) { return { ok: false, error: "obtenerEgresosPendientes: " + e.message }; }
}

function aprobarEgreso(token, egresoId, aprobar, notas) {
  const sesion = verificarSesion(token);
  if (aprobar === undefined) aprobar = true;
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_EGRESOS);
    if (!h) return { ok: false, error: "No existe hoja de egresos" };
    const datos = h.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === egresoId) {
        h.getRange(i+1,11).setValue(aprobar ? "aprobado" : "rechazado");
        h.getRange(i+1,12).setValue(sesion.usuario.nombre);
        h.getRange(i+1,13).setValue(new Date());
        if (notas) h.getRange(i+1,14).setValue(notas);
        return { ok: true };
      }
    }
    return { ok: false, error: "Egreso no encontrado" };
  } catch(e) { return { ok: false, error: "aprobarEgreso: " + e.message }; }
}

// ============================================================
// MÓDULO 2 — TRANSFERENCIAS ENTRE LOCALES
// ============================================================

function solicitarTransferencia(token, datos) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };
  if (sesion.usuario.rol === "disenador") return { ok: false, error: "No autorizado" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEET_TRANSFERENCIAS)) { ss.insertSheet(SHEET_TRANSFERENCIAS); _setupTransferencias(ss); }
    const h = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    const localesMap = _getLocalesMap(ss);
    const id = _generarID();
    h.appendRow([
      id, new Date(),
      sesion.usuario.localId, localesMap[sesion.usuario.localId] || "",
      datos.localDestinoId, localesMap[datos.localDestinoId] || "",
      sesion.usuario.id, sesion.usuario.nombre,
      "", "", // cajero destino se llena al confirmar
      Number(datos.monto), datos.motivo.trim(),
      "pendiente_admin", "", "", "", datos.notas || ""
    ]);
    return { ok: true, id };
  } catch(e) { return { ok: false, error: "solicitarTransferencia: " + e.message }; }
}

function obtenerTransferenciasPendientesAdmin(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    if (!h) return { ok: true, transferencias: [] };
    const transferencias = h.getDataRange().getValues().slice(1)
      .filter(r => r[12] === "pendiente_admin")
      .map(r => ({
        id: String(r[0]), fecha: _fechaToStr(r[1]),
        localOrigen: String(r[3]), localDestino: String(r[5]),
        cajeroOrigen: String(r[7]), monto: Number(r[10]),
        motivo: String(r[11]), estado: String(r[12])
      }));
    return { ok: true, transferencias };
  } catch(e) { return { ok: false, error: "obtenerTransferenciasPendientesAdmin: " + e.message }; }
}

function aprobarTransferencia(token, transId, aprobar, notas) {
  const sesion = verificarSesion(token);
  if (aprobar === undefined) aprobar = true;
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    if (!h) return { ok: false, error: "No existe hoja" };
    const datos = h.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === transId) {
        const nuevoEstado = aprobar ? "aprobada_pendiente_confirmacion" : "rechazada";
        h.getRange(i+1,13).setValue(nuevoEstado);
        h.getRange(i+1,14).setValue(sesion.usuario.nombre);
        h.getRange(i+1,15).setValue(new Date());
        if (notas) h.getRange(i+1,17).setValue(notas);
        return { ok: true };
      }
    }
    return { ok: false, error: "Transferencia no encontrada" };
  } catch(e) { return { ok: false, error: "aprobarTransferencia: " + e.message }; }
}

function obtenerTransferenciasPorConfirmar(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    if (!h) return { ok: true, transferencias: [] };
    const transferencias = h.getDataRange().getValues().slice(1)
      .filter(r => r[12] === "aprobada_pendiente_confirmacion" && r[4] === sesion.usuario.localId)
      .map(r => ({
        id: String(r[0]), fecha: _fechaToStr(r[1]),
        localOrigen: String(r[3]), localDestino: String(r[5]),
        cajeroOrigen: String(r[7]), monto: Number(r[10]),
        motivo: String(r[11]), estado: String(r[12])
      }));
    return { ok: true, transferencias };
  } catch(e) { return { ok: false, error: "obtenerTransferenciasPorConfirmar: " + e.message }; }
}

function confirmarRecepcion(token, transId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    if (!h) return { ok: false, error: "No existe hoja" };
    const datos = h.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === transId && datos[i][4] === sesion.usuario.localId) {
        h.getRange(i+1,9).setValue(sesion.usuario.id);
        h.getRange(i+1,10).setValue(sesion.usuario.nombre);
        h.getRange(i+1,13).setValue("completada");
        h.getRange(i+1,16).setValue(new Date());
        return { ok: true };
      }
    }
    return { ok: false, error: "Transferencia no encontrada o no te corresponde" };
  } catch(e) { return { ok: false, error: "confirmarRecepcion: " + e.message }; }
}

function obtenerTransferenciasDelDia(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    if (!h) return { ok: true, transferencias: [] };
    const hoy = _fechaHoy();
    const transferencias = h.getDataRange().getValues().slice(1)
      .filter(r => _fechaToStr(r[1]) === hoy &&
                   (r[6] === sesion.usuario.id || r[8] === sesion.usuario.id || r[4] === sesion.usuario.localId))
      .map(r => ({
        id: String(r[0]), fecha: _fechaToStr(r[1]),
        localOrigen: String(r[3]), localDestino: String(r[5]),
        cajeroOrigen: String(r[7]), cajeroDestino: String(r[9] || ""),
        monto: Number(r[10]), motivo: String(r[11]), estado: String(r[12]),
        direccion: r[6] === sesion.usuario.id ? "salida" : "entrada"
      }));
    return { ok: true, transferencias };
  } catch(e) { return { ok: false, error: "obtenerTransferenciasDelDia: " + e.message }; }
}

// ============================================================
// MÓDULO 3 — REGISTRO DE DEUDAS (solo admin)
// ============================================================

function registrarDeuda(token, datos) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEET_DEUDAS)) { ss.insertSheet(SHEET_DEUDAS); _setupDeudas(ss); }
    const h = ss.getSheetByName(SHEET_DEUDAS);
    const monto = Number(datos.monto);
    h.appendRow([
      _generarID(), new Date(), datos.deudorNombre.trim(), datos.deudorContacto || "",
      datos.concepto.trim(), monto, 0, monto, "activa",
      sesion.usuario.nombre, datos.notas || "", datos.fechaVencimiento || ""
    ]);
    return { ok: true };
  } catch(e) { return { ok: false, error: "registrarDeuda: " + e.message }; }
}

function obtenerDeudas(token, filtro) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_DEUDAS);
    if (!h) return { ok: true, deudas: [] };
    let deudas = h.getDataRange().getValues().slice(1).filter(r => r[0]);
    if (filtro === "activas") deudas = deudas.filter(r => r[8] !== "saldada");
    if (filtro === "saldadas") deudas = deudas.filter(r => r[8] === "saldada");
    return {
      ok: true,
      deudas: deudas.map(r => ({
        id: String(r[0]), fecha: _fechaToStr(r[1]),
        deudorNombre: String(r[2]), deudorContacto: String(r[3]),
        concepto: String(r[4]), montoTotal: Number(r[5]),
        totalAbonado: Number(r[6]), saldoPendiente: Number(r[7]),
        estado: String(r[8]), registradoPor: String(r[9]),
        notas: String(r[10]), fechaVencimiento: _fechaToStr(r[11])
      })).reverse()
    };
  } catch(e) { return { ok: false, error: "obtenerDeudas: " + e.message }; }
}

function abonarDeuda(token, deudaId, monto, fecha, metodo, notas) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hDeudas = ss.getSheetByName(SHEET_DEUDAS);
    if (!hDeudas) return { ok: false, error: "No existe hoja de deudas" };
    if (!ss.getSheetByName(SHEET_DEUDAS_PAGOS)) { ss.insertSheet(SHEET_DEUDAS_PAGOS); _setupDeudasPagos(ss); }
    ss.getSheetByName(SHEET_DEUDAS_PAGOS).appendRow([
      _generarID(), deudaId, Number(monto), fecha || _fechaHoy(),
      metodo || "efectivo", notas || "", sesion.usuario.nombre, new Date()
    ]);
    const datos = hDeudas.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === deudaId) {
        const nuevoAbonado = Number(datos[i][6]) + Number(monto);
        const nuevoSaldo   = Number(datos[i][5]) - nuevoAbonado;
        const nuevoEstado  = nuevoSaldo <= 0 ? "saldada" : "parcial";
        hDeudas.getRange(i+1,7).setValue(nuevoAbonado);
        hDeudas.getRange(i+1,8).setValue(Math.max(0, nuevoSaldo));
        hDeudas.getRange(i+1,9).setValue(nuevoEstado);
        return { ok: true, nuevoSaldo: Math.max(0, nuevoSaldo), estado: nuevoEstado };
      }
    }
    return { ok: false, error: "Deuda no encontrada" };
  } catch(e) { return { ok: false, error: "abonarDeuda: " + e.message }; }
}

function obtenerPagosDeuda(token, deudaId) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "No autorizado" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(SHEET_DEUDAS_PAGOS);
    if (!h) return { ok: true, pagos: [] };
    const pagos = h.getDataRange().getValues().slice(1)
      .filter(r => r[1] === deudaId)
      .map(r => ({
        id: String(r[0]), monto: Number(r[2]), fecha: _fechaToStr(r[3]),
        metodo: String(r[4]), notas: String(r[5]), registradoPor: String(r[6])
      }));
    return { ok: true, pagos };
  } catch(e) { return { ok: false, error: "obtenerPagosDeuda: " + e.message }; }
}

function _generarID() {
  return Utilities.getUuid().replace(/-/g, "").substring(0, 12);
}

function _generarToken() {
  return Utilities.getUuid().replace(/-/g, "") + Date.now().toString(36);
}

function _hashPin(pin) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pin.toString()
  );
  return bytes.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

// ── FECHA HELPERS ──
// Usamos la zona horaria del Spreadsheet para evitar desfases UTC
function _fechaHoy() {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
}

// Convierte cualquier valor de celda (Date o string) a string "yyyy-MM-dd"
function _fechaToStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(val, tz, "yyyy-MM-dd");
  }
  // Ya es string, normalizar
  return String(val).substring(0, 10);
}

function _getPeriodoActual() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = now.getDate();
  const quincena = d <= 15 ? "Q1" : "Q2";
  return `${y}-${m}-${quincena}`;
}

function _getFechasPeriodo(periodo) {
  // periodo = "2025-01-Q1" o "2025-01-Q2"
  const partes = periodo.split("-");
  const y = parseInt(partes[0]);
  const m = parseInt(partes[1]) - 1;
  const q = partes[2];

  if (q === "Q1") {
    return {
      inicio: `${partes[0]}-${partes[1]}-01`,
      fin: `${partes[0]}-${partes[1]}-15`
    };
  } else {
    const ultimoDia = new Date(y, m + 1, 0).getDate();
    return {
      inicio: `${partes[0]}-${partes[1]}-16`,
      fin: `${partes[0]}-${partes[1]}-${ultimoDia}`
    };
  }
}

function _getUltimosCierresPorRecurso(datos, userId) {
  const mapa = {};

  // First pass: cupo_inicial as base reference
  datos.slice(1)
    .filter(row => row[2] === userId && row[7] === "cupo_inicial")
    .forEach(row => {
      const recId = row[5];
      if (!mapa[recId]) {
        mapa[recId] = {
          recursoId: recId,
          recursoNombre: row[6],
          ultimoValor: Number(row[8]),
          // ✅ FIX: convertir Date a string — GAS no puede serializar Date objects
          ultimaFecha: "cupo_inicial"
        };
      }
    });

  // Second pass: actual closing values override cupo_inicial
  datos.slice(1)
    .filter(row => row[2] === userId && row[7] === "cierre")
    .forEach(row => {
      const recId = row[5];
      // ✅ FIX: _fechaToStr convierte el Date object de Sheets a string "yyyy-MM-dd"
      const fechaStr = _fechaToStr(row[1]);
      if (!mapa[recId]) {
        mapa[recId] = {
          recursoId: recId,
          recursoNombre: row[6],
          ultimoValor: Number(row[8]),
          ultimaFecha: fechaStr   // string, no Date
        };
      } else {
        // Comparación segura entre strings (funciona correctamente en formato yyyy-MM-dd)
        if (mapa[recId].ultimaFecha === "cupo_inicial" || fechaStr > mapa[recId].ultimaFecha) {
          mapa[recId] = {
            recursoId: recId,
            recursoNombre: row[6],
            ultimoValor: Number(row[8]),
            ultimaFecha: fechaStr   // string, no Date
          };
        }
      }
    });

  return mapa;
}

function _getLocalesMap(ss) {
  const h = ss.getSheetByName(SHEET_LOCALES);
  const datos = h.getDataRange().getValues();
  const mapa = {};
  for (let i = 1; i < datos.length; i++) {
    mapa[datos[i][0]] = datos[i][1];
  }
  return mapa;
}




// ============================================================
// ALIASES Y FUNCIONES FALTANTES — CORRECCIÓN DE MISMATCH
// ============================================================

/**
 * Alias de obtenerResumenQuincenal para compatibilidad con el frontend
 * @param {string} token
 */
function obtenerResumenCierreQuincenal(token) {
  return obtenerResumenQuincenal(token);
}

/**
 * Obtiene registros diarios con filtros opcionales (para tab Historial admin)
 * @param {string} token
 * @param {Object} filtros - { usuarioId, fecha, tipo }
 * @returns {{ ok: boolean, registros: Array }}
 */
function obtenerDetalleRegistros(token, filtros) {
  const sesion = verificarSesion(token);
  if (!sesion.ok) return { ok: false, error: "No autorizado" };

  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const h   = ss.getSheetByName(SHEET_REGISTROS);
    if (!h)   return { ok: true, registros: [] };

    const datos = h.getDataRange().getValues().slice(1);
    const f = filtros || {};

    const registros = datos
      .filter(r => {
        if (!r[0]) return false;
        if (f.usuarioId && r[2] !== f.usuarioId) return false;
        if (f.fecha    && _fechaToStr(r[1]) !== f.fecha) return false;
        if (f.tipo     && r[7] !== f.tipo) return false;
        // Si es cajero normal, solo ve los suyos
        if (sesion.usuario.rol === "cajero" && r[2] !== sesion.usuario.id) return false;
        return true;
      })
      .map(r => ({
        fecha:         _fechaToStr(r[1]),
        usuarioNombre: String(r[3] || ""),
        recursoNombre: String(r[6] || ""),
        tipo:          String(r[7] || ""),
        valor:         Number(r[8]) || 0,
        ganancia:      Number(r[9]) || 0,
        periodo:       String(r[10] || "")
      }))
      .reverse();

    return { ok: true, registros };
  } catch(e) {
    return { ok: false, error: "obtenerDetalleRegistros: " + e.message };
  }
}

/**
 * Obtiene egresos Y transferencias pendientes de aprobación (para tab Egresos admin)
 * @param {string} token
 * @returns {{ ok: boolean, egresos: Array, transferencias: Array }}
 */
function obtenerPendientesAprobacion(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Egresos pendientes ──
    const hEgr = ss.getSheetByName(SHEET_EGRESOS);
    const egresos = hEgr
      ? hEgr.getDataRange().getValues().slice(1)
          .filter(r => r[10] === "pendiente")
          .map(r => ({
            id:           String(r[0]),
            fecha:        _fechaToStr(r[1]),
            usuarioNombre: String(r[3] || ""),
            localId:      String(r[4] || ""),
            monto:        Number(r[5]) || 0,
            concepto:     String(r[6] || ""),
            solicitante:  String(r[7] || ""),
            tipo:         String(r[8] || ""),
            esDeuda:      r[9] === true,
            estado:       String(r[10] || "")
          }))
      : [];

    // ── Transferencias pendientes de aprobación admin ──
    const hTrans = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    const transferencias = hTrans
      ? hTrans.getDataRange().getValues().slice(1)
          .filter(r => r[12] === "pendiente_admin")
          .map(r => ({
            id:          String(r[0]),
            fecha:       _fechaToStr(r[1]),
            localOrigen: String(r[3] || ""),
            localDestino: String(r[5] || ""),
            cajeroOrigen: String(r[7] || ""),
            monto:       Number(r[10]) || 0,
            motivo:      String(r[11] || ""),
            estado:      String(r[12] || "")
          }))
      : [];

    return { ok: true, egresos, transferencias };
  } catch(e) {
    return { ok: false, error: "obtenerPendientesAprobacion: " + e.message };
  }
}

/**
 * Rechaza un egreso (wrapper para compatibilidad con frontend)
 * @param {string} token
 * @param {string} egresoId
 */
function rechazarEgreso(token, egresoId) {
  return aprobarEgreso(token, egresoId, false, "");
}

/**
 * Rechaza una transferencia (wrapper para compatibilidad con frontend)
 * @param {string} token
 * @param {string} transId
 */
function rechazarTransferencia(token, transId) {
  return aprobarTransferencia(token, transId, false, "");
}

/**
 * Obtiene historial completo de transferencias para el admin
 * @param {string} token
 * @returns {{ ok: boolean, transferencias: Array }}
 */
function obtenerTransferenciasAdmin(token) {
  const sesion = verificarSesion(token);
  if (!sesion.ok || sesion.usuario.rol !== "admin") return { ok: false, error: "Solo admin" };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h  = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    if (!h)  return { ok: true, transferencias: [] };

    const transferencias = h.getDataRange().getValues().slice(1)
      .filter(r => r[0])
      .map(r => ({
        id:           String(r[0]),
        fecha:        _fechaToStr(r[1]),
        localOrigen:  String(r[3] || ""),
        localDestino: String(r[5] || ""),
        cajeroOrigen: String(r[7] || ""),
        monto:        Number(r[10]) || 0,
        motivo:       String(r[11] || ""),
        estado:       String(r[12] || "")
      }))
      .reverse();

    return { ok: true, transferencias };
  } catch(e) {
    return { ok: false, error: "obtenerTransferenciasAdmin: " + e.message };
  }
}
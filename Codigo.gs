// Servir la interfaz web (Por si la abres directamente)
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Control de Préstamos - Liquida PRO')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// NUEVA FUNCIÓN API: Permite que tu app de GitHub le envíe y pida datos a Google Sheets
function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  var result = {};
  
  try {
    if (action === "obtenerClientes") {
      result = obtenerClientes();
    } else if (action === "agregarCliente") {
      result = agregarCliente(params.nombre, params.telefono, params.direccion);
    } else if (action === "crearPrestamo") {
      result = crearPrestamo(params.idCliente, params.monto);
    } else if (action === "obtenerRutaCobros") {
      result = obtenerRutaCobros(params.fechaFiltroStr);
    } else if (action === "registrarCobro") {
      result = registrarCobro(params.idCobro, params.valor, params.cobrador, params.estado, params.comentario, params.exento);
    } else if (action === "registrarGasto") {
      result = registrarGasto(params.monto, params.concepto, params.cobrador);
    } else if (action === "obtenerMetricasGlobales") {
      result = obtenerMetricasGlobales(params.fechaFiltroStr);
    }
  } catch(err) {
    result = { success: false, error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
                       .setMimeType(ContentService.MimeType.JSON);
}

// --- FUNCIONES INTERNAS DE LA BASE DE DATOS ---
function obtenerClientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DATOS DEL CLIENTE");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var clientes = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== "") {
      clientes.push({ id: data[i][0], nombre: data[i][1], telefono: data[i][2], direccion: data[i][3] });
    }
  }
  return clientes;
}

function agregarCliente(nombre, telefono, direccion) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DATOS DEL CLIENTE");
  var ultimoID = 0;
  if (sheet.getLastRow() > 1) { ultimoID = sheet.getRange(sheet.getLastRow(), 1).getValue(); }
  var nuevoID = ultimoID + 1;
  sheet.appendRow([nuevoID, nombre, telefono, direccion]);
  return { success: true, id: nuevoID, nombre: nombre };
}

function crearPrestamo(idCliente, monto) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDetalle = ss.getSheetByName("DETALLE");
  var sheetPagos = ss.getSheetByName("SEGUIMIENTO DE PAGOS");
  var montoEntregado = parseFloat(monto);
  var interes = 0.20;
  var totalAPagar = montoEntregado * (1 + interes);
  var fechaEntrega = new Date();
  var fechaVencimiento = new Date();
  fechaVencimiento.setMonth(fechaVencimiento.getMonth() + 1);
  
  var diasHabiles = [];
  var tempDate = new Date(fechaEntrega);
  tempDate.setDate(tempDate.getDate() + 1);
  while (tempDate <= fechaVencimiento) {
    if (tempDate.getDay() !== 0) { diasHabiles.push(new Date(tempDate)); }
    tempDate.setDate(tempDate.getDate() + 1);
  }
  var cuotaDiaria = totalAPagar / diasHabiles.length;
  var ultimoIDPrestamo = 0;
  if (sheetDetalle.getLastRow() > 1) { ultimoIDPrestamo = sheetDetalle.getRange(sheetDetalle.getLastRow(), 1).getValue(); }
  var nuevoIDPrestamo = ultimoIDPrestamo + 1;
  
  sheetDetalle.appendRow([nuevoIDPrestamo, idCliente, montoEntregado, interes, totalAPagar, fechaEntrega, fechaVencimiento, "Activo"]);
  
  var ultimoIDCobro = 0;
  if (sheetPagos.getLastRow() > 1) { ultimoIDCobro = sheetPagos.getRange(sheetPagos.getLastRow(), 1).getValue(); }
  var cuotasLote = [];
  for (var i = 0; i < diasHabiles.length; i++) {
    ultimoIDCobro++;
    cuotasLote.push([ultimoIDCobro, nuevoIDPrestamo, diasHabiles[i], 0, "", "Pendiente", "", "No"]);
  }
  if (cuotasLote.length > 0) {
    var filaInicio = sheetPagos.getLastRow() + 1;
    sheetPagos.getRange(filaInicio, 1, cuotasLote.length, cuotasLote[0].length).setValues(cuotasLote);
  }
  return { success: true, message: "Préstamo #" + nuevoIDPrestamo + " creado." };
}

function obtenerRutaCobros(fechaFiltroStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("SEGUIMIENTO DE PAGOS");
  var sheetDetalle = ss.getSheetByName("DETALLE");
  var sheetClientes = ss.getSheetByName("DATOS DEL CLIENTE");
  
  var pagosData = sheetPagos.getDataRange().getValues();
  var detalleData = sheetDetalle.getDataRange().getValues();
  var clientesData = sheetClientes.getDataRange().getValues();
  
  var clientesMap = {};
  for (var i = 1; i < clientesData.length; i++) {
    clientesMap[clientesData[i][0]] = { nombre: clientesData[i][1], telefono: clientesData[i][2], direccion: clientesData[i][3] };
  }
  
  var cuotasContador = {};
  for (var k = 1; k < pagosData.length; k++) {
    var idPrestamo = pagosData[k][1];
    if (idPrestamo) { cuotasContador[idPrestamo] = (cuotasContador[idPrestamo] || 0) + 1; }
  }

  var prestamosMap = {};
  for (var j = 1; j < detalleData.length; j++) {
    var idPrestamo = detalleData[j][0];
    if (idPrestamo !== "") {
      var totalAPagar = parseFloat(detalleData[j][4]) || 0;
      var totalCuotas = cuotasContador[idPrestamo] || 1;
      prestamosMap[idPrestamo] = { idCliente: detalleData[j][1], totalAPagar: totalAPagar, cuotaDiaria: totalAPagar / totalCuotas };
    }
  }
  
  var targetDate = fechaFiltroStr ? new Date(fechaFiltroStr + "T00:00:00") : new Date();
  var targetDateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  var ruta = [];
  for (var k = 1; k < pagosData.length; k++) {
    var fechaPago = new Date(pagosData[k][2]);
    var fechaPagoStr = Utilities.formatDate(fechaPago, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (fechaPagoStr === targetDateStr) {
      var idCobro = pagosData[k][0];
      var idPrestamo = pagosData[k][1];
      var prestamo = prestamosMap[idPrestamo] || {};
      var cliente = clientesMap[prestamo.idCliente] || { nombre: "Desconocido", telefono: "", direccion: "" };
      
      ruta.push({
        idCobro: idCobro,
        idPrestamo: idPrestamo,
        clienteNombre: cliente.nombre,
        clienteTelefono: cliente.telefono,
        clienteDireccion: cliente.direccion,
        valorPagado: pagosData[k][3],
        valorCuota: prestamo.cuotaDiaria || 0,
        cobrador: pagosData[k][4],
        estadoPago: pagosData[k][5],
        comentario: pagosData[k][6],
        exentoMora: pagosData[k][7]
      });
    }
  }
  return { success: true, data: ruta };
}

function registrarCobro(idCobro, valor, cobrador, estado, comentario, exento) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("SEGUIMIENTO DE PAGOS");
  var data = sheetPagos.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == idCobro) {
      var fila = i + 1;
      sheetPagos.getRange(fila, 4).setValue(parseFloat(valor) || 0); 
      sheetPagos.getRange(fila, 5).setValue(cobrador || "");         
      sheetPagos.getRange(fila, 6).setValue(estado);                 
      sheetPagos.getRange(fila, 7).setValue(comentario || "");       
      sheetPagos.getRange(fila, 8).setValue(exento ? "Sí" : "No");   
      return { success: true, message: "Cobro registrado." };
    }
  }
  return { success: false, message: "ID no encontrado." };
}

function registrarGasto(monto, concepto, cobrador) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("GASTOS") || ss.insertSheet("GASTOS");
  if (sheet.getLastRow() === 0) { sheet.appendRow(["ID Gasto", "Fecha", "Concepto", "Monto", "Cobrador"]); }
  var ultimoID = sheet.getLastRow() > 1 ? sheet.getRange(sheet.getLastRow(), 1).getValue() : 0;
  var nuevoID = parseInt(ultimoID || 0) + 1;
  sheet.appendRow([nuevoID, new Date(), concepto, parseFloat(monto) || 0, cobrador]);
  return { success: true, message: "Gasto guardado." };
}

function obtenerMetricasGlobales(fechaFiltroStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDetalle = ss.getSheetByName("DETALLE");
  var sheetPagos = ss.getSheetByName("SEGUIMIENTO DE PAGOS");
  var sheetGastos = ss.getSheetByName("GASTOS");
  var totalCartera = 0, clientesHoy = 0, cobradoHoy = 0, prestadoHoy = 0, gastosHoy = 0;
  
  var targetDate = fechaFiltroStr ? new Date(fechaFiltroStr + "T00:00:00") : new Date();
  var targetDateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  if (sheetDetalle) {
    var detalle = sheetDetalle.getDataRange().getValues();
    for (var i = 1; i < detalle.length; i++) {
      if (detalle[i][7] === "Activo") { totalCartera += parseFloat(detalle[i][4]) || 0; }
      var fechaEntrega = new Date(detalle[i][5]);
      var fechaEntregaStr = Utilities.formatDate(fechaEntrega, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (fechaEntregaStr === targetDateStr) { prestadoHoy += parseFloat(detalle[i][2]) || 0; }
    }
  }
  if (sheetPagos) {
    var pagos = sheetPagos.getDataRange().getValues();
    for (var j = 1; j < pagos.length; j++) {
      var fechaPago = new Date(pagos[j][2]);
      var fechaPagoStr = Utilities.formatDate(fechaPago, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (fechaPagoStr === targetDateStr) {
        clientesHoy++;
        if (pagos[j][5] === "Pagado") { cobradoHoy += parseFloat(pagos[j][3]) || 0; }
      }
    }
  }
  if (sheetGastos) {
    var gastos = sheetGastos.getDataRange().getValues();
    for (var k = 1; k < gastos.length; k++) {
      var fechaGasto = new Date(gastos[k][1]);
      var fechaGastoStr = Utilities.formatDate(fechaGasto, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (fechaGastoStr === targetDateStr) { gastosHoy += parseFloat(gastos[k][3]) || 0; }
    }
  }
  return {
    carteraActiva: totalCartera,
    clientesRuta: clientesHoy,
    cobradoDia: cobradoHoy,
    prestadoDia: prestadoHoy,
    gastosDia: gastosHoy,
    saldoCaja: cobradoHoy - prestadoHoy - gastosHoy
  };
}
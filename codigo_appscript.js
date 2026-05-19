/**
 * =========================================================================================
 * MOTOR INTEGRAL DE ASISTENCIA CLÍNICA Y AUDITORÍA EN LA NUBE
 * Cátedra: Clínica Psicológica y Psicoterapias - Universidad Nacional de Córdoba
 * Autor: Joaquín Gustavo Fernández Farkas
 * =========================================================================================
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Asistencia Clínica')
      .addItem('🛠️ 1. Crear Pestaña de Configuración', 'crearPanelConfiguracion')
      .addItem('🏠 2. Actualizar Menú de Navegación', 'actualizarMenu')
      .addItem('🎨 3. Formatear Hoja General', 'formatearHojaGeneral')
      .addItem('▶️ 4. Clasificar historial acumulado', 'clasificarHistorial')
      .addSeparator()
      .addSubMenu(ui.createMenu('📊 Generar Informes')
          .addItem('📅 Informe por Rango de Clases', 'informeRangoClases')
          .addItem('📆 Informe por Mes Completo', 'informePorMes')
          .addItem('🏆 PLANILLA DE ASISTENCIA 2026 (Anual)', 'informeAnual'))
      .addSeparator()
      .addItem('🟢 5. Activar automatización en tiempo real', 'instalarVigilante')
      .addItem('🛑 6. Apagar automatización', 'apagarVigilante')
      .addToUi();
}

function crearPanelConfiguracion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombreHoja = "⚙️ CONFIGURACIÓN";
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) hoja = ss.insertSheet(nombreHoja, 0);
  
  hoja.getRange("A1:B1").merge().setValue("PANEL DE AUDITORÍA Y SEGURIDAD").setFontWeight("bold").setBackground("#4A154B").setFontColor("#FFF").setHorizontalAlignment("center").setFontSize(14);
  hoja.getRange("A2").setValue("Hora a la que se proyectó el primer QR (HH:MM):").setFontWeight("bold");
  hoja.getRange("B2").setValue("16:30").setBackground("#FFFF00").setHorizontalAlignment("center"); 
  hoja.getRange("A3").setValue("Ventana de Gracia por mal internet (Minutos):").setFontWeight("bold");
  hoja.getRange("B3").setValue("7").setBackground("#FFFF00").setHorizontalAlignment("center");
  hoja.setColumnWidth(1, 350);
  hoja.setColumnWidth(2, 100);
  SpreadsheetApp.getUi().alert("Panel de Configuración creado con éxito.");
}

function auditarToken(marcaTemporal, token) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaConfig = ss.getSheetByName("⚙️ CONFIGURACIÓN");
  if (!hojaConfig) return "⚠️ SIN AUDITAR";
  
  var horaTexto = hojaConfig.getRange("B2").getValue().toString();
  var graciaMinutos = parseInt(hojaConfig.getRange("B3").getValue());
  
  var match = token.match(/QR(\d+)_/i);
  if (!match) return "❌ TOKEN INVÁLIDO";
  
  var numQR = parseInt(match[1]);
  var segundosDesdeInicio = (numQR - 1) * 15;
  
  var partesHora = horaTexto.split(":");
  var fechaAsistencia = new Date(marcaTemporal);
  var horaProyectado = new Date(fechaAsistencia.getFullYear(), fechaAsistencia.getMonth(), fechaAsistencia.getDate(), parseInt(partesHora[0]), parseInt(partesHora[1]), 0);
  horaProyectado.setSeconds(horaProyectado.getSeconds() + segundosDesdeInicio);
  
  var horaLimite = new Date(horaProyectado.getTime());
  horaLimite.setMinutes(horaLimite.getMinutes() + graciaMinutos);
  
  if (fechaAsistencia <= horaLimite) return "🟢 VÁLIDO";
  else return "🔴 SOSPECHOSO (Fuera de tiempo)";
}

function instalarVigilante() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(sheet);
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'procesarNuevaRespuesta') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('procesarNuevaRespuesta').forSpreadsheet(sheet).onFormSubmit().create();
  SpreadsheetApp.getUi().alert("✅ ¡SISTEMA EN TIEMPO REAL ACTIVADO CON ÉXITO!");
}

function apagarVigilante() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(sheet);
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'procesarNuevaRespuesta') ScriptApp.deleteTrigger(triggers[i]);
  }
  SpreadsheetApp.getUi().alert("🛑 SISTEMA EN TIEMRE REAL DESACTIVADO");
}

function reordenarFila(filaOriginal, esEncabezado) {
  var nuevaFila = [];
  for (var i = 0; i < filaOriginal.length; i++) {
    var valor = filaOriginal[i];
    if (esEncabezado) {
      valor = valor.toString().toUpperCase(); 
      if (valor === "MARCA TEMPORAL") valor = "MARCA\nTEMPORAL";
      if (valor === "CORREO ELECTRONICO" || valor === "DIRECCIÓN DE CORREO ELECTRÓNICO") valor = "CORREO\nELECTRONICO";
      if (valor === "PALABRA CLAVE") valor = "PALABRA\nCLAVE";
      if (valor.indexOf("TOKEN") > -1 || valor === "NO TOCAR") valor = "TOKEN DE SEGURIDAD\nDEL SISTEMA";
    }
    nuevaFila.push(valor);
  }
  if (nuevaFila.length >= 6) {
    var temp = nuevaFila[4];
    nuevaFila[4] = nuevaFila[5];
    nuevaFila[5] = temp;
  }
  return nuevaFila;
}

function procesarNuevaRespuesta(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPrincipal = ss.getSheetByName("RESPUESTAS DEL FORMULARIO GRAL");
  if (!hojaPrincipal) return;
  
  var ultimaFila = hojaPrincipal.getLastRow();
  var valoresFila = hojaPrincipal.getRange(ultimaFila, 1, 1, hojaPrincipal.getLastColumn()).getValues()[0];
  var celdaFecha = new Date(valoresFila[0]);
  if (!celdaFecha || isNaN(celdaFecha.getTime())) return; 
  
  var dia = ("0" + celdaFecha.getDate()).slice(-2);
  var mesInt = celdaFecha.getMonth() + 1; 
  var mes = ("0" + mesInt).slice(-2);
  var anio = celdaFecha.getFullYear().toString().slice(-2);
  var anioCompleto = celdaFecha.getFullYear();
  var nombrePestana = "Clase " + dia + "/" + mes + "/" + anio;
  
  var hojaDestino = ss.getSheetByName(nombrePestana);
  var esNueva = false;
  
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet(nombrePestana);
    hojaDestino.getRange("A1").setValue("CLASE " + dia + "/" + mes + "/" + anioCompleto);
    hojaDestino.getRange("A1:D1").merge();
    hojaDestino.getRange("E1").setValue("Hora exacta en que se mostro el QR:");
    hojaDestino.getRange("E1:F1").merge();
    
    var encabezadosOrig = hojaPrincipal.getRange(1, 1, 1, hojaPrincipal.getLastColumn()).getValues()[0];
    var encabezadosListos = reordenarFila(encabezadosOrig, true);
    encabezadosListos.push("RESULTADO DE\nAUDITORÍA");
    hojaDestino.getRange(2, 1, 1, encabezadosListos.length).setValues([encabezadosListos]);
    esNueva = true;
  }
  
  var filaLista = reordenarFila(valoresFila, false);
  var resultadoAuditoria = auditarToken(valoresFila[0], filaLista[5]); 
  filaLista.push(resultadoAuditoria);
  
  hojaDestino.appendRow(filaLista);
  aplicarFormatoEstetico(hojaDestino, mesInt);
  
  if (esNueva) {
    ordenarPestanas(ss);
    actualizarMenu();
  }
}

function clasificarHistorial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPrincipal = ss.getSheetByName("RESPUESTAS DEL FORMULARIO GRAL");
  if (!hojaPrincipal) return;
  
  var datos = hojaPrincipal.getDataRange().getValues();
  if (datos.length < 2) return;
  
  var encabezadosOrig = datos[0];
  var contador = 0;
  
  for (var i = 1; i < datos.length; i++) {
    var filaActual = datos[i];
    var celdaFecha = new Date(filaActual[0]);
    if (!celdaFecha || isNaN(celdaFecha.getTime())) continue; 
    
    var dia = ("0" + celdaFecha.getDate()).slice(-2);
    var mesInt = celdaFecha.getMonth() + 1;
    var mes = ("0" + mesInt).slice(-2);
    var anio = celdaFecha.getFullYear().toString().slice(-2);
    var anioCompleto = celdaFecha.getFullYear();
    var nombrePestana = "Clase " + dia + "/" + mes + "/" + anio;
    
    var hojaDestino = ss.getSheetByName(nombrePestana);
    if (!hojaDestino) {
      hojaDestino = ss.insertSheet(nombrePestana);
      hojaDestino.getRange("A1").setValue("CLASE " + dia + "/" + mes + "/" + anioCompleto);
      hojaDestino.getRange("A1:D1").merge();
      hojaDestino.getRange("E1").setValue("Hora exacta en que se mostro el QR:");
      hojaDestino.getRange("E1:F1").merge();
      
      var encabezadosListos = reordenarFila(encabezadosOrig, true);
      encabezadosListos.push("RESULTADO DE\nAUDITORÍA");
      hojaDestino.getRange(2, 1, 1, encabezadosListos.length).setValues([encabezadosListos]);
    }
    
    var datosDestino = hojaDestino.getDataRange().getValues();
    var yaExiste = false;
    for(var j = 2; j < datosDestino.length; j++){
      if(datosDestino[j][0].toString() === celdaFecha.toString()){ yaExiste = true; break; }
    }
    
    if(!yaExiste){
      var filaLista = reordenarFila(filaActual, false);
      var resultadoAuditoria = auditarToken(filaActual[0], filaLista[5]);
      filaLista.push(resultadoAuditoria);
      
      hojaDestino.appendRow(filaLista);
      contador++;
    }
    aplicarFormatoEstetico(hojaDestino, mesInt);
  }
  
  ordenarPestanas(ss);
  actualizarMenu();
  SpreadsheetApp.getUi().alert("Historial revisado y ordenado cronológicamente con éxito.");
}

function enmascararDNI(dniRaw) {
  if (!dniRaw) return "XX.XXX.000";
  var limpio = dniRaw.toString().replace(/\D/g, "");
  if (limpio.length < 3) return "XX.XXX." + limpio;
  return "XX.XXX." + limpio.slice(-3);
}

function obtenerAlumnosDeHojas(hojasAProcesar) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var alumnosMap = {};
  
  var hojaGral = ss.getSheetByName("RESPUESTAS DEL FORMULARIO GRAL");
  var encabezados = hojaGral.getRange(1, 1, 1, hojaGral.getLastColumn()).getValues()[0];
  
  var idxApellidos = -1, idxNombres = -1, idxDni = -1;
  for (var i = 0; i < encabezados.length; i++) {
    var h = encabezados[i].toString().toUpperCase();
    if (h.indexOf("APELLIDO") > -1) idxApellidos = i;
    if (h.indexOf("NOMBRE") > -1) idxNombres = i;
    if (h.indexOf("DNI") > -1 || h.indexOf("DOCUMENTO") > -1 || h.indexOf("TOCAR") > -1) idxDni = i; 
  }
  
  hojasAProcesar.forEach(function(nombreHoja) {
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 3) return;
    
    for (var f = 2; f < datos.length; f++) {
      var fila = datos[f];
      var nombre = fila[1] ? fila[1].toString().trim() : "";
      var apellido = fila[2] ? fila[2].toString().trim() : "";
      
      var dni = "000";
      if (idxDni > -1) {
        var datosGral = hojaGral.getDataRange().getValues();
        for(var g = 1; g < datosGral.length; g++) {
          if (datosGral[g][idxApellidos].toString().trim() === apellido && datosGral[g][idxNombres].toString().trim() === nombre) {
            dni = datosGral[g][idxDni];
            break;
          }
        }
      }
      
      if (apellido !== "" && nombre !== "") {
        var claveUnica = (apellido + ", " + nombre).toUpperCase();
        alumnosMap[claveUnica] = { apellido: apellido, nombre: nombre, dni: enmascararDNI(dni) };
      }
    }
  });
  
  var listaAlumnos = [];
  for (var key in alumnosMap) listaAlumnos.push(alumnosMap[key]);
  listaAlumnos.sort(function(a, b) { return a.apellido.localeCompare(b.apellido); });
  return listaAlumnos;
}

function construirInformeEstructurado(nombreInforme, tituloPrincipal, listaAlumnos, esAnual) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var hojaVieja = ss.getSheetByName(nombreInforme);
  if (hojaVieja) ss.deleteSheet(hojaVieja);
  var hojaInf = ss.insertSheet(nombreInforme);
  
  var colsCount = esAnual ? 4 : 3;
  
  hojaInf.getRange(1, 1, 1, colsCount).merge().setValue(tituloPrincipal.toUpperCase())
         .setFontSize(14).setFontWeight("bold").setBackground("#4A154B").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  
  var headers = ["APELLIDO", "NOMBRE", "DNI"];
  if (esAnual) headers.push("CONDICIÓN FINAL");
  hojaInf.getRange(2, 1, 1, colsCount).setValues([headers])
         .setFontSize(14).setFontWeight("bold").setBackground("#F3F3F3").setFontColor("#000000").setHorizontalAlignment("center");
         
  var filaActual = 3;
  var ultimaLetra = "";
  
  for (var i = 0; i < listaAlumnos.length; i++) {
    var al = listaAlumnos[i];
    var letraActual = al.apellido.trim().charAt(0).toUpperCase();
    
    if (letraActual !== ultimaLetra) {
      ultimaLetra = letraActual;
      var rangoIndice = hojaInf.getRange(filaActual, 1, 1, colsCount);
      rangoIndice.merge().setValue("APELLIDOS - " + ultimaLetra)
                 .setFontSize(14).setFontWeight("bold").setBackground("#E2E8F0").setFontColor("#2D3748").setHorizontalAlignment("left");
      filaActual++;
    }
    
    hojaInf.getRange(filaActual, 1).setValue(al.apellido);
    hojaInf.getRange(filaActual, 2).setValue(al.nombre);
    hojaInf.getRange(filaActual, 3).setValue(al.dni);
    if (esAnual) {
      hojaInf.getRange(filaActual, 4).setValue("REGULAR"); 
    }
    filaActual++;
  }
  
  var rangoTablaCompleta = hojaInf.getRange(1, 1, filaActual - 1, colsCount);
  rangoTablaCompleta.setFontFamily("Arial").setVerticalAlignment("middle")
                    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  
  if (filaActual > 3) {
    hojaInf.getRange(3, 1, filaActual - 3, colsCount).setFontSize(12).setHorizontalAlignment("left");
  }
  
  if (esAnual && filaActual > 3) {
    var rangoCF = hojaInf.getRange(3, 1, filaActual - 3, colsCount);
    
    var reglaLibre = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$D3="LIBRE"')
        .setBackground("#FBC4C4")
        .setFontColor("#721C24")
        .setRanges([rangoCF])
        .build();
        
    var reglaPromo = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$D3="PROMOCIONAL"')
        .setBackground("#B4E1FF")
        .setFontColor("#004085")
        .setRanges([rangoCF])
        .build();
        
    var reglasActuales = hojaInf.getConditionalFormatRules();
    reglasActuales.push(reglaLibre);
    reglasActuales.push(reglaPromo);
    hojaInf.setConditionalFormatRules(reglasActuales);
  }
  
  for (var col = 1; col <= colsCount; col++) {
    hojaInf.autoResizeColumn(col);
    hojaInf.setColumnWidth(col, hojaInf.getColumnWidth(col) + 30);
  }
  
  hojaInf.setFrozenRows(2);
  ss.setActiveSheet(hojaInf);
  SpreadsheetApp.getUi().alert("¡Informe '" + nombreInforme + "' generado de forma impecable!");
}

function informeRangoClases() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var resultInicio = ui.prompt("📅 INFORME POR RANGO\n\nIngrese fecha inicial de la clase (Ej: 01/05/26):");
  if (resultInicio.getSelectedButton() !== ui.Button.OK) return;
  var inicioStr = "Clase " + resultInicio.getResponseText().trim();
  
  var resultFin = ui.prompt("📅 INFORME POR RANGO\n\nIngrese fecha final de la clase (Ej: 18/05/26):");
  if (resultFin.getSelectedButton() !== ui.Button.OK) return;
  var finStr = "Clase " + resultFin.getResponseText().trim();
  
  var hojas = ss.getSheets();
  var hojasAProcesar = [];
  var incluir = false;
  
  for (var i = 0; i < hojas.length; i++) {
    var name = hojas[i].getName();
    if (name === inicioStr) incluir = true;
    if (incluir && name.indexOf("Clase ") === 0) hojasAProcesar.push(name);
    if (name === finStr) { incluir = false; break; }
  }
  
  if (hojasAProcesar.length === 0) {
    ui.alert("❌ Rango no encontrado. Revise el nombre exacto de las pestañas.");
    return;
  }
  
  var alumnos = obtenerAlumnosDeHojas(hojasAProcesar);
  construirInformeEstructurado("Informe Rango Personalizado", "Asistencia desde " + resultInicio.getResponseText() + " hasta " + resultFin.getResponseText(), alumnos, false);
}

function informePorMes() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var result = ui.prompt("Itemizado por Mes\n\nIngrese número del mes (Ej: 5 para Mayo, 6 para Junio):");
  if (result.getSelectedButton() !== ui.Button.OK) return;
  
  var mesBuscado = ("0" + result.getResponseText().trim()).slice(-2);
  var nombresMeses = {"01":"Enero","02":"Febrero","03":"Marzo","04":"Abril","05":"Mayo","06":"Junio","07":"Julio","08":"Agosto","09":"Septiembre","10":"Octubre","11":"Noviembre","12":"Diciembre"};
  var nombreMesTexto = nombresMeses[mesBuscado] || "Mes " + mesBuscado;
  
  var hojas = ss.getSheets();
  var hojasAProcesar = [];
  
  for (var i = 0; i < hojas.length; i++) {
    var name = hojas[i].getName();
    if (name.indexOf("Clase ") === 0) {
      var partes = name.split("/");
      if (partes.length === 3 && partes[1] === mesBuscado) hojasAProcesar.push(name);
    }
  }
  
  if (hojasAProcesar.length === 0) {
    ui.alert("❌ No hay clases registradas para el mes de " + nombreMesTexto + ".");
    return;
  }
  
  var alumnos = obtenerAlumnosDeHojas(hojasAProcesar);
  construirInformeEstructurado("Informe Meznsual - " + nombreMesTexto, "Consolidado de Asistencia: " + nombreMesTexto, alumnos, false);
}

function informeAnual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = ss.getSheets();
  var hojasAProcesar = [];
  
  for (var i = 0; i < hojas.length; i++) {
    var name = hojas[i].getName();
    if (name.indexOf("Clase ") === 0) hojasAProcesar.push(name);
  }
  
  if (hojasAProcesar.length === 0) {
    SpreadsheetApp.getUi().alert("❌ No se encontraron clases procesadas para compilar el año.");
    return;
  }
  
  var alumnos = obtenerAlumnosDeHojas(hojasAProcesar);
  construirInformeEstructurado("PLANILLA DE ASISTENCIA 2026", "Planilla General Calificatoria - Ciclo Lectivo 2026", alumnos, true);
}

function aplicarFormatoEstetico(hoja, mes) {
  var ultimaFila = hoja.getLastRow();
  var ultimaColumna = hoja.getLastColumn();
  if (ultimaFila === 0 || ultimaColumna === 0) return;

  var colores = { 1: "#B4E1FF", 2: "#FFB3BA", 3: "#BAFFC9", 4: "#FFDFBA", 5: "#FFFF00", 6: "#FF0000", 7: "#CCCCFF", 8: "#E5CCFF", 9: "#FFC3A0", 10: "#FFFFCC", 11: "#DFFF00", 12: "#E0E0E0" };
  var colorFondo = colores[mes] || "#FFFF00";
  var colorTextoEncabezado = (mes === 6) ? "#FFFFFF" : "#000000"; 

  hoja.setTabColor(colorFondo);
  var todo = hoja.getRange(1, 1, ultimaFila, ultimaColumna);
  todo.setFontFamily("Arial").setVerticalAlignment("middle").setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  var fila1 = hoja.getRange("A1:G1");
  fila1.setBackground(colorFondo).setFontColor(colorTextoEncabezado).setFontWeight("bold").setVerticalAlignment("middle");
  hoja.getRange("A1:D1").setFontSize(14).setHorizontalAlignment("center");
  hoja.getRange("E1:F1").setFontSize(11).setHorizontalAlignment("right").setWrap(true);
  hoja.getRange("G1").setBackground("#FFFFFF").setFontColor("#000000").setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  var fila2 = hoja.getRange(2, 1, 1, ultimaColumna);
  fila2.setFontSize(14).setFontWeight("bold").setBackground(colorFondo).setFontColor(colorTextoEncabezado).setHorizontalAlignment("center").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  if (ultimaFila > 2) {
    var datos = hoja.getRange(3, 1, ultimaFila - 2, ultimaColumna);
    datos.setFontSize(12).setFontWeight("normal").setFontColor("#000000").setBackground("#FFFFFF").setHorizontalAlignment("left");
    
    for(var f = 3; f <= ultimaFila; f++) {
      var celdaEstado = hoja.getRange(f, 7);
      var valorEstado = celdaEstado.getValue();
      if(valorEstado.indexOf("🟢") > -1) {
        celdaEstado.setFontColor("#155724").setBackground("#d4edda").setFontWeight("bold");
      } else if (valorEstado.indexOf("🔴") > -1 || valorEstado.indexOf("❌") > -1 || valorEstado.indexOf("⚠️") > -1) {
        celdaEstado.setFontColor("#721c24").setBackground("#f8d7da").setFontWeight("bold");
      }
    }
  }

  for (var col = 1; col <= ultimaColumna; col++) {
    hoja.autoResizeColumn(col);
    var ancho = hoja.getColumnWidth(col);
    hoja.setColumnWidth(col, col === 6 ? Math.max(ancho + 40, 240) : ancho + 40);
  }
  hoja.setFrozenRows(2);
}

function ordenarPestanas(ss) {
  var hojas = ss.getSheets();
  var hojasDeClase = [];
  var cantidadHojasFijas = 0;
  
  for (var i = 0; i < hojas.length; i++) {
    var nombre = hojas[i].getName();
    if (nombre.indexOf("Clase ") === 0) {
      var partes = nombre.replace("Clase ", "").split("/");
      if (partes.length === 3) {
        var fechaTiempo = new Date(2000 + parseInt(partes[2], 10), parseInt(partes[1], 10) - 1, parseInt(partes[0], 10)).getTime();
        hojasDeClase.push({ hoja: hojas[i], tiempo: fechaTiempo });
      }
    } else {
      cantidadHojasFijas++;
    }
  }
  hojasDeClase.sort(function(a, b) { return a.tiempo - b.tiempo; });
  for (var j = 0; j < hojasDeClase.length; j++) {
    ss.setActiveSheet(hojasDeClase[j].hoja);
    ss.moveActiveSheet(cantidadHojasFijas + j + 1); 
  }
  var principal = ss.getSheetByName("RESPUESTAS DEL FORMULARIO GRAL");
  if (principal) ss.setActiveSheet(principal);
}

function formatearHojaGeneral() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("RESPUESTAS DEL FORMULARIO GRAL");
  if (!hoja) return;
  var ultimaFila = hoja.getLastRow();
  var ultimaColumna = hoja.getLastColumn();
  if (ultimaFila === 0 || ultimaColumna === 0) return;

  hoja.getDataRange().setFontFamily("Arial").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(1, 1, 1, ultimaColumna).setFontSize(14).setFontWeight("bold").setBackground("#4A154B").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  if (ultimaFila > 1) hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).setFontSize(12).setBackground("#FFFFFF").setHorizontalAlignment("left");
  for (var col = 1; col <= ultimaColumna; col++) { hoja.autoResizeColumn(col); hoja.setColumnWidth(col, hoja.getColumnWidth(col) + 30); }
  hoja.setFrozenRows(1);
}

function actualizarMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombreMenu = "🏠 MENÚ PRINCIPAL";
  var hojaVieja = ss.getSheetByName(nombreMenu);
  if (hojaVieja) ss.deleteSheet(hojaVieja);
  var hojaMenu = ss.insertSheet(nombreMenu, 0);
  
  var nombresMeses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  var colores = {1: "#B4E1FF", 2: "#FFB3BA", 3: "#BAFFC9", 4: "#FFDFBA", 5: "#FFFF00", 6: "#FF0000", 7: "#CCCCFF", 8: "#E5CCFF", 9: "#FFC3A0", 10: "#FFFFCC", 11: "#DFFF00", 12: "#E0E0E0"};
  
  var hojas = ss.getSheets();
  var clasesPorMes = {};
  for (var i = 0; i < hojas.length; i++) {
    var nombre = hojas[i].getName();
    if (nombre.indexOf("Clase ") === 0) {
      var partes = nombre.replace("Clase ", "").split("/");
      if (partes.length === 3) {
        var mesInt = parseInt(partes[1], 10);
        if (!clasesPorMes[mesInt]) clasesPorMes[mesInt] = [];
        clasesPorMes[mesInt].push({ nombre: nombre, id: hojas[i].getSheetId(), dia: parseInt(partes[0], 10) });
      }
    }
  }
  hojaMenu.setColumnWidth(1, 40); hojaMenu.setColumnWidth(2, 400);
  hojaMenu.getRange("B2").setValue("📚 MENÚ DE NAVEGACIÓN").setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center").setFontFamily("Arial");
  var filaActual = 4;
  for (var m = 1; m <= 12; m++) {
    if (clasesPorMes[m] && clasesPorMes[m].length > 0) {
      clasesPorMes[m].sort(function(a, b) { return a.dia - b.dia; });
      var celdaMes = hojaMenu.getRange(filaActual, 2);
      celdaMes.setValue("📅 " + nombresMeses[m-1]).setFontSize(14).setFontWeight("bold").setBackground(colores[m]).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#000", SpreadsheetApp.BorderStyle.SOLID).setFontFamily("Arial");
      if(m===6) celdaMes.setFontColor("#FFFFFF");
      var filaInicioGrupo = filaActual + 1; filaActual++;
      for (var c = 0; c < clasesPorMes[m].length; c++) {
        var clase = clasesPorMes[m][c];
        hojaMenu.getRange(filaActual, 2).setFormula('=HYPERLINK("#gid=' + clase.id + '", "🔗 Ir a ' + clase.nombre + '")').setFontSize(12).setBackground("#F9F9F9").setBorder(false, true, false, true, false, false, "#000", SpreadsheetApp.BorderStyle.SOLID).setFontFamily("Arial");
        filaActual++;
      }
      hojaMenu.getRange(filaActual - 1, 2).setBorder(false, true, true, true, false, false, "#000", SpreadsheetApp.BorderStyle.SOLID);
      hojaMenu.getRange(filaInicioGrupo, 1, clasesPorMes[m].length, 1).shiftRowGroupDepth(1);
      filaActual++;
    }
  }
  if (filaActual > 4) hojaMenu.collapseAllRowGroups();
  hojaMenu.hideGridlines();
}

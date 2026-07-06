// ============================================
// CONFIGURACIÓN TICS - GEMINI 3 FLASH PREVIEW
// ============================================
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

const INVENTARIO_HEADERS = [
  'ACTIVO/FIJO',
  'DESCRIPCIÓN',
  'MARCA',
  'MODELO',
  'SERIE',
  'CUSTODIO',
  'UBICACIÓN',
  'DIRECCIÓN',
  'ESTADO',
  'OBSERVACIÓN'
];

// Obtiene la hoja principal de inventario y asegura que exista la fila de encabezados.
function getHojaInventario_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Busca explícitamente la hoja llamada 'Inventario'
  let hojaInventario = ss.getSheetByName('Inventario');
  
  // Si no existe, la crea automáticamente para evitar errores
  if (!hojaInventario) {
    hojaInventario = ss.insertSheet('Inventario');
  }
  
  asegurarEncabezadosInventario_(hojaInventario);
  return hojaInventario;
}

// Crea los encabezados una sola vez cuando la hoja de inventario esta vacia.
function asegurarEncabezadosInventario_(sheet) {
  if (sheet.getFrozenRows() > 0) {
    sheet.setFrozenRows(0);
  }

  const firstRow = sheet.getRange(1, 1, 1, INVENTARIO_HEADERS.length).getValues()[0];
  const hojaSinDatos = sheet.getLastRow() === 0 || firstRow.every(cell => cell === '' || cell === null);

  if (!hojaSinDatos) return;

  sheet.getRange(1, 1, 1, INVENTARIO_HEADERS.length).setValues([INVENTARIO_HEADERS]);
  sheet.getRange(1, 1, 1, INVENTARIO_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#0067a8')
    .setFontColor('#ffffff');
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Inventario TICS - Prototipo Web')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Valida el correo ingresado contra la hoja "Accesos".
function validarCorreoListaBlanca(correoIngresado) {
  try {
    if (!correoIngresado) return false;
    const correoBuscado = correoIngresado.trim().toLowerCase();
    if (correoBuscado === '') return false;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaAccesos = ss.getSheetByName('Accesos');
    if (!hojaAccesos) return false;

    const correosAutorizados = hojaAccesos.getRange('A:A').getValues();

    for (let i = 0; i < correosAutorizados.length; i++) {
      const correoCelda = correosAutorizados[i][0].toString().trim().toLowerCase();
      if (correoCelda === '') continue;
      if (correoCelda === correoBuscado) {
        return true;
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

// Limpia el prefijo data:image cuando la imagen llega desde el navegador.
function limpiarBase64Imagen_(base64Image) {
  return (base64Image || '').replace(/^data:image\/(png|jpg|jpeg);base64,/, '');
}

// Envia una o dos imagenes a Gemini en una sola solicitud.
// Imagen 1: codigo DASE o codigo de barras institucional.
// Imagen 2: etiqueta del equipo donde suelen estar marca, modelo y serie.
function processImageAnalysis(imagenes, correoUsuario) {
  if (!validarCorreoListaBlanca(correoUsuario)) return { success: false, error: 'Acceso denegado. Correo no autorizado.' };

  try {
    const imagenCodigoDase = typeof imagenes === 'string'
      ? limpiarBase64Imagen_(imagenes)
      : limpiarBase64Imagen_(imagenes.codigoDase);

    const imagenEtiquetaEquipo = typeof imagenes === 'string'
      ? ''
      : limpiarBase64Imagen_(imagenes.etiquetaEquipo);

    if (!imagenCodigoDase && !imagenEtiquetaEquipo) {
      throw new Error('No se recibieron imagenes para analizar.');
    }

    const prompt = `Analiza la(s) siguiente(s) imagen(es) de inventario. Las fotos pueden contener el codigo institucional DASE y/o la etiqueta tecnica del equipo. 
Extrae UNICAMENTE en JSON plano:

Reglas de extraccion:
- activo_fijo: codigos DASE-XXXXXX. Ignora codigos municipales o numeros sueltos.
- descripcion: tipo de equipo (ej. TODO EN UNO, CPU, MONITOR, TECLADO, MOUSE, LAPTOP, IMPRESORA, ESCANER)
- marca: marca del equipo
- modelo: modelo completo del equipo. Prioriza modelo comercial
- serie: numero de serie (S/N, Serial, Serie, Service Tag). Prioriza Service Tag
- No inventes datos. Si un campo no se ve claro, dejalo vacio.

Responde SOLO el JSON:
{
  "activo_fijo": "",
  "descripcion": "",
  "marca": "",
  "modelo": "",
  "serie": "",
}`;

    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${GEMINI_API_KEY}`;
    const parts = [{ text: prompt }];

    if (imagenCodigoDase) {
      parts.push({ inline_data: { mime_type: 'image/jpeg', data: imagenCodigoDase } });
    }

    if (imagenEtiquetaEquipo) {
      parts.push({ inline_data: { mime_type: 'image/jpeg', data: imagenEtiquetaEquipo } });
    }

    const payload = {
      contents: [{ parts }],
      generationConfig: {
        temperature: 0.1,
        response_mime_type: 'application/json',
        maxOutputTokens: 256,
        thinkingConfig: {
          thinkingLevel: 'minimal'
        }
      }
    };

    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    if (!result.candidates) throw new Error('Fallo en la extracción de datos visuales.');

    let extractedText = result.candidates[0].content.parts[0].text;
    return { success: true, data: JSON.parse(extractedText) };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Devuelve una lista sin repetidos de custodios ya registrados en el inventario.
function getListaCustodios(correoUsuario) {
  if (!validarCorreoListaBlanca(correoUsuario)) return [];
  try {
    const sheet = getHojaInventario_();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const custodios = data.slice(1).map(row => row[5]).filter(String);
    return [...new Set(custodios)].sort();
  } catch (e) { return []; }
}

// Busca todos los equipos asignados exactamente al custodio ingresado.
function buscarEquiposPorCustodio(nombre, correoUsuario) {
  if (!validarCorreoListaBlanca(correoUsuario)) return { success: false, error: 'Acceso denegado.' };
  try {
    const sheet = getHojaInventario_();
    const data = sheet.getDataRange().getValues();
    const nombreBusqueda = nombre.trim().toUpperCase();
    const resultados = data.filter((row, index) => (index !== 0 && row[5].toString().toUpperCase() === nombreBusqueda));
    return { success: true, data: resultados };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// Busca un equipo por ACTIVO/FIJO con coincidencia exacta.
function buscarEquiposPorCodigo(codigo, correoUsuario) {
  if (!validarCorreoListaBlanca(correoUsuario)) return { success: false, error: 'Acceso denegado.' };
  try {
    const sheet = getHojaInventario_();
    const data = sheet.getDataRange().getValues();
    const codigoBusqueda = codigo.trim().toUpperCase();
    const resultados = data.filter((row, index) => {
        if (index === 0) return false;
      const activoFijo = row[0].toString().trim().toUpperCase();
      return activoFijo === codigoBusqueda;
    });
    return { success: true, data: resultados };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// Guarda un nuevo registro o actualiza uno existente. Incluye rastreo especial para equipos sin código ni serie.
function processInventorySave(data, correoUsuario) {
  if (!validarCorreoListaBlanca(correoUsuario)) return { success: false, error: 'Acceso denegado.' };
  let lockAcquired = false;
  const lock = LockService.getScriptLock();

  try {
    data = data || {};
    lock.waitLock(10000);
    lockAcquired = true;

    const sheet = getHojaInventario_();
    const fullData = sheet.getDataRange().getValues();
    
    // Limpiamos los datos entrantes
    const sn = (data.serie || '').toString().trim().toUpperCase();
    const activoFijo = (data.activo_fijo || '').toString().trim().toUpperCase();
    const tipoEquipo = (data.descripcion || '').toString().trim().toUpperCase();
    const marcaEntrada = (data.marca || '').toString().trim().toUpperCase();
    const modeloEntrada = (data.modelo || '').toString().trim().toUpperCase();
    const obsEntrada = (data.observacion || '').toString().trim().toUpperCase();
    
    // Validación backend de campos obligatorios
    if (tipoEquipo === '' || (data.custodio || '').trim() === '' || (data.ubicacion || '').trim() === '' || (data.direccion || '').trim() === '' || (data.estado || '').trim() === '') {
      return { success: false, error: 'Operación bloqueada por el servidor: Faltan campos obligatorios requeridos (*).' };
    }

    // Verifica estrictamente el formato DASE-XXXXXX antes de guardar
    if (activoFijo !== '') {
      const regexDase = /^DASE-\d{6}$/;
      if (!regexDase.test(activoFijo)) {
        return { success: false, error: 'Operación bloqueada por el servidor: El Activo Fijo debe tener el formato DASE- seguido de 6 números.' };
      }
    }

    const rowData = [
      activoFijo,
      tipoEquipo,
      marcaEntrada,
      modeloEntrada,
      sn,
      (data.custodio || '').toString().trim().toUpperCase(),
      (data.ubicacion || '').toString().trim().toUpperCase(),
      (data.direccion || '').toString().trim().toUpperCase(),
      (data.estado || '').toString().trim().toUpperCase(),
      obsEntrada
    ];

    let rowIndex = -1;
    
    // 1. PASO: Identificar si estamos editando un equipo existente
    for (let i = 1; i < fullData.length; i++) {
      const activoCelda = fullData[i][0].toString().trim().toUpperCase();
      const descCelda = fullData[i][1].toString().trim().toUpperCase();
      const marcaCelda = fullData[i][2].toString().trim().toUpperCase();
      const modeloCelda = fullData[i][3].toString().trim().toUpperCase();
      const serieCelda = fullData[i][4].toString().trim().toUpperCase();
      const obsCelda = fullData[i][9].toString().trim().toUpperCase();
      
      // NIVEL 1: Identidad por Activo Fijo (Si tiene)
      if (activoFijo !== '' && activoCelda === activoFijo && descCelda === tipoEquipo) {
        rowIndex = i + 1; 
        break;
      }
      
      // NIVEL 2: Identidad por Serie (Si no tiene Activo Fijo pero sí tiene Serie)
      else if (activoFijo === '' && sn !== '' && serieCelda === sn && descCelda === tipoEquipo) {
        rowIndex = i + 1;
        break;
      }

      // NIVEL 3: EQUIPOS FANTASMA (No tienen Activo Fijo ni Serie)
      else if (activoFijo === '' && sn === '' && activoCelda === '' && serieCelda === '' && descCelda === tipoEquipo) {
        // Al no tener "cédula", usamos la combinación de Marca, Modelo y Observación para encontrarlo
        if (marcaCelda === marcaEntrada && modeloCelda === modeloEntrada && obsCelda === obsEntrada) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    // 2. PASO: Validar que la serie sea estrictamente única en todo el documento
    if (sn !== '') {
      for (let i = 1; i < fullData.length; i++) {
        const serieCelda = fullData[i][4].toString().trim().toUpperCase();
        
        if (serieCelda === sn) {
          // Si la serie ya la tiene OTRA fila distinta a la que vamos a editar, BLOQUEA
          if ((i + 1) !== rowIndex) {
            const activoRobado = fullData[i][0].toString().trim().toUpperCase();
            const descRobado = fullData[i][1].toString().trim().toUpperCase();
            return { success: false, error: `Esta serie ya está en uso por el equipo: ${descRobado} ${activoRobado !== '' ? '('+activoRobado+')' : ''}. La serie no puede repetirse.` };
          }
        }
      }
    }

    // Guardar o Actualizar
    if (rowIndex !== -1) { 
      sheet.getRange(rowIndex, 1, 1, 10).setValues([rowData]); 
    } else { 
      sheet.appendRow(rowData); 
    }

    // Ordenamiento automático
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, 1, lastRow - 1, 10);
      const values = range.getValues();
      const priorUbi = { 'PLANTA BAJA': 1, 'PISO 1': 2, 'PISO 2': 3, 'PISO 3': 4 };
      const priorDes = { 'TODO EN UNO': 1, 'CPU': 2, 'MONITOR': 3, 'TECLADO': 4, 'MOUSE': 5, 'LAPTOP': 6, 'IMPRESORA': 7, 'ESCÁNER': 7 };

      values.sort((a, b) => {
        // 1. Primero ordena por UBICACIÓN (Índice 6)
        const uA = priorUbi[a[6]] || 99, uB = priorUbi[b[6]] || 99;
        if (uA !== uB) return uA - uB;
        
        // 2. Segundo ordena alfabéticamente por DIRECCIÓN (Índice 7)
        if (a[7] !== b[7]) return a[7].localeCompare(b[7]);
        
        // 3. Tercero ordena alfabéticamente por CUSTODIO (Índice 5)
        if (a[5] !== b[5]) return a[5].localeCompare(b[5]);
        
        // 4. Por último ordena por DESCRIPCIÓN del tipo de equipo (Índice 1)
        const dA = priorDes[a[1]] || 99, dB = priorDes[b[1]] || 99;
        return dA - dB;
      });
      range.setValues(values);
    }
    return { success: true };
    } catch (error) {
    return { success: false, error: error.toString() };
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

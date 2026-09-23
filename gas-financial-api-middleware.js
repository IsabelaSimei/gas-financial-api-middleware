// URL base da API e Header de Autenticação (Basic Auth em Base64)
var API_BASE_URL = "https://api.yourdomain.com/api/ChatBox";
var AUTH_HEADER  = "Basic YOUR_BASE64_AUTH_HEADER_HERE"; 

function doPost(e) {
  var requestBody = {};
  try {
    requestBody = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: true, mensaje: "JSON inválido" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = requestBody.action;
  var responseText = '';
  
  // Mapeamento de ações do Chatbot para os parâmetros da API
  if (action === "Valida Usuario") {
    responseText = callChatBoxApi("GET", { cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else if (action === "Informacion de Credito") {
    responseText = callChatBoxApi("GET", { r: 1, cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else if (action === "Informacion de Refinanciamiento") {
    responseText = callChatBoxApi("GET", { r: 2, cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else if (action === "Solicitud de Refinanciamiento") {
    responseText = callChatBoxApi("GET", { r: 1, Prestamo: requestBody.prestamo, cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else if (action === "Informacion de Asesor") {
    responseText = callChatBoxApi("GET", { r: 3, cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else if (action === "Queja Sugerencia") {
    responseText = callChatBoxApi("PUT", { r: 1, Queja: requestBody.queja, cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else if (action === "Morosidad") {
    responseText = callChatBoxApi("GET", { r: 4, cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else if (action === "Consulta Tramite") {
    responseText = callChatBoxApi("GET", { r: 5, cedula: requestBody.cedula, telefono: requestBody.telefono, prov: requestBody.prov });
    
  } else {
    responseText = JSON.stringify({ error: true, mensaje: "Acción no reconocida" });
  }

  // Registra o log no Google Sheets
  saveGSheetResponseFast(action, requestBody.contact_id || "N/A", responseText);

  return ContentService.createTextOutput(responseText)
    .setMimeType(ContentService.MimeType.JSON);
}

function callChatBoxApi(method, params) {
  var queryString = Object.keys(params).map(function(key) {
    if (params[key] !== undefined && params[key] !== null) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    }
    return '';
  }).filter(Boolean).join('&');

  var url = API_BASE_URL + "?" + queryString;
  
  var options = {
    method: method,
    headers: {
      "Authorization": AUTH_HEADER,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    return response.getContentText();
  } catch (err) {
    return JSON.stringify({ error: true, mensaje: "Error de conexión a la API: " + err.message });
  }
}

function saveGSheetResponseFast(action, contact_id, response) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var finalResponse = (typeof response === 'object') ? JSON.stringify(response) : response;
    sheet.appendRow([new Date(), action, contact_id, finalResponse]);
  } catch(e) {
    console.error("Erro ao salvar no GSheet: " + e.toString());
  }
}
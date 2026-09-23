function crearActivadorDiario() {
  // Limpa gatilhos anteriores para evitar duplicidades
  const activadores = ScriptApp.getProjectTriggers();
  for (let activador of activadores) {
    if (activador.getHandlerFunction() === 'procesarLeadsDiarios') {
      ScriptApp.deleteTrigger(activador);
    }
  }

  // Executa de Segunda a Sábado
  const dias = [
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY,
    ScriptApp.WeekDay.SATURDAY
  ];

  dias.forEach(dia => {
    ScriptApp.newTrigger('procesarLeadsDiarios')
      .timeBased()
      .inTimezone("America/Brasilia")
      .onWeekDay(dia)
      .atHour(7)
      .nearMinute(30)
      .create();
  });
}

function asignarAgenciasAuto() {
  const ID_HOJA = "YOUR_SPREADSHEET_ID_HERE";
  const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName("Leads");
  if (!hoja) return;

  const datos = hoja.getDataRange().getValues();
  
  // Índices das colunas na planilha (Base 0)
  const IDX_ZONA = 7;        
  const IDX_AGENCIA = 12;    
  const IDX_CORREO_AG = 20;  
  const IDX_DIRECCION = 21;  

  const correosGerentes = {
    "Zona X": "gerente.x@company.com",
    "Zona A": "gerente.a@company.com",
    "Zona J": "gerente.j@company.com"
  };

  const mapaDirecciones = [
    { keyword: "x", agencia: "Zona X" },
    { keyword: "a", agencia: "Zona A" },
    { keyword: "j", agencia: "Zona J" }
  ];

  let totalActualizados = 0;

  for (let i = 1; i < datos.length; i++) {
    let fila = datos[i];
    let agenciaActual = (fila[IDX_AGENCIA] || "").toString().trim();
    let correoActual = (fila[IDX_CORREO_AG] || "").toString().trim();
    
    const esEmailValido = (email) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

    if (!agenciaActual || agenciaActual === "Sin Agencia" || !esEmailValido(correoActual)) {
      let direccionStr = ((fila[IDX_DIRECCION] || "") + " " + (fila[IDX_ZONA] || "")).toString().toLowerCase();
      let agenciaEncontrada = "";

      for (let mapa of mapaDirecciones) {
        if (direccionStr.includes(mapa.keyword)) {
          agenciaEncontrada = mapa.agencia;
          break;
        }
      }

      if (agenciaEncontrada) {
        let correoGerente = correosGerentes[agenciaEncontrada] || "";
        let numFila = i + 1; 

        hoja.getRange(numFila, IDX_AGENCIA + 1).setValue(agenciaEncontrada);
        hoja.getRange(numFila, IDX_ZONA + 1).setValue(agenciaEncontrada);

        if (correoGerente) {
          hoja.getRange(numFila, IDX_CORREO_AG + 1).setValue(correoGerente);
        }
        totalActualizados++;
      }
    }
  }

  if (totalActualizados > 0) {
    Logger.log(`⚡ Atribuída agência a ${totalActualizados} registros.`);
  }
}

function procesarLeadsDiarios() {
  asignarAgenciasAuto();

  const ID_HOJA = "YOUR_SPREADSHEET_ID_HERE";
  const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName("Leads"); 
  if (!hoja) return;

  const datos = hoja.getDataRange().getValues();
  
  const IDX_NOMBRE = 4;       
  const IDX_TELEFONO = 5;     
  const IDX_DOCUMENTO = 6;    
  const IDX_AGENCIA = 12;     
  const IDX_EMAIL = 13;       
  const IDX_OPCION = 17;      
  const IDX_CORREO_AG = 20;   
  const IDX_ENVIADO = 26;     
  const IDX_FECHA_ENV = 27;   

  const ccGeral = "supervisao@company.com, analista@company.com"; 

  let agenciasLeads = {};
  let filasProcessar = [];

  const estiloTabla = `<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">`;
  const estiloCabecera = `<tr style="background-color: #f2f2f2; font-weight: bold;">`;

  for (let i = 1; i < datos.length; i++) {
    let fila = datos[i];
    let enviado = fila[IDX_ENVIADO];
    
    if (enviado !== "SI" && enviado !== true) {
      let numFila = i + 1;
      let agencia = (fila[IDX_AGENCIA] || "Sem Agencia").toString().trim();
      let correoAgencia = (fila[IDX_CORREO_AG] || "").toString().trim();

      let tdBasico = `<td>${fila[IDX_NOMBRE]}</td><td>${fila[IDX_DOCUMENTO]}</td><td>${fila[IDX_TELEFONO]}</td><td>${fila[IDX_EMAIL]}</td>`;

      if (!agenciasLeads[agencia]) {
        agenciasLeads[agencia] = { correo: correoAgencia, leads: [] };
      }
      agenciasLeads[agencia].leads.push(`<tr>${tdBasico}</tr>`);
      filasProcessar.push(numFila);
    }
  }

  const fechaHoy = new Date();
  const esEmailValido = (email) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

  for (let agencia in agenciasLeads) {
    let data = agenciasLeads[agencia];
    if (data.leads.length > 0 && esEmailValido(data.correo)) {
      
      MailApp.sendEmail({
        to: data.correo,
        cc: ccGeral,
        subject: `Relatório Diário de Leads - ${agencia}`,
        htmlBody: `<h3>Relatório de Novas Oportunidades</h3><p>Agência: <b>${agencia}</b></p>${estiloTabla}<thead>${estiloCabecera}<th>Nome Completo</th><th>Documento</th><th>Telefone</th><th>Email</th></tr></thead><tbody>${data.leads.join("")}</tbody></table>`
      });
    }
  }

  filasProcessar.forEach(numFila => {
    hoja.getRange(numFila, IDX_ENVIADO + 1).setValue("SI");
    hoja.getRange(numFila, IDX_FECHA_ENV + 1).setValue(fechaHoy);
  });

  Logger.log("✅ Processo concluído com sucesso.");
}

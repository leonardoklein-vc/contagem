/**
 * SAMIS - SISTEMA INTEGRADO (Versão Completa: Legado + OCR.space)
 */

// --- CONFIGURAÇÕES GERAIS ---
var ID_PLANILHA = "1vAqNHdlKOfwOSBSNdT7jd5JFvSCEJdK0YN-L9NiNPzk"; // SEU ID
var OCR_API_KEY = "K81464241688957"; // SUA API KEY

// --- MAPA DE OCR (33 ITENS) ---
const MAPA_33_ITENS = [
  {ocr: "ADMINISTRAÇÃO PRESCRIÇÃO MÉDICA", padrao: "ADMINISTRAÇÃO DA PRESCRIÇÃO MÉDICA"},
  {ocr: "BOLETIM DE IDENTIFICAÇÃO", padrao: "BOLETIM DE IDENTIFICAÇÃO"},
  {ocr: "BOMBA DE INFUSÃO OCITOCINA", padrao: "BOMBA DE INFUSÃO OCITOCINA"},
  {ocr: "CHECK LIST", padrao: "CHECK LIST"},
  {ocr: "CONTRA-CHEQUE PRESCRIÇÃO", padrao: "CONTRA-CHEQUE PRESCRIÇÃO"},
  {ocr: "CONTROLE DE DIÁLISE", padrao: "CONTROLE DE DIÁLISE"},
  {ocr: "CONTROLE NEUROLÓGICO", padrao: "CONTROLE NEUROLÓGICO"},
  {ocr: "DECLARAÇÃO DE NASCIDO VIVO", padrao: "DECLARAÇÃO DE NASCIDO VIVO"},
  {ocr: "DECLARAÇÃO DE ÓBITO", padrao: "DECLARAÇÃO DE ÓBITO"},
  {ocr: "DECLARAÇÃO INVOLUNTÁRIA", padrao: "DECLARAÇÃO INVOLUNTÁRIA"},
  {ocr: "DECLARAÇÃO VOLUNTÁRIA", padrao: "DECLARAÇÃO VOLUNTÁRIA"},
  {ocr: "DIU", padrao: "DIU"},
  {ocr: "FICHA DE ANESTESIA", padrao: "FICHA DE ANESTESIA"},
  {ocr: "FICHA DE ATENDIMENTO PRE HOSP SAM", padrao: "FICHA DE ATENDIMENTO PRÉ HOSPITALAR SAMU"},
  {ocr: "IDENTIFICAÇÃO NASCIDO VIVO", padrao: "IDENTIFICAÇÃO DE RECÉM NASCIDO"},
  {ocr: "MOVIMENTAÇÕES DA PRESCRIÇÃO MÉD", padrao: "MOVIMENTAÇÕES DA PRESCRIÇÃO MÉDICA"},
  {ocr: "MOVIMENTAÇÕES DA PRESCRIÇÃO ENFH", padrao: "MOVIMENTAÇÕES DA PRESCRIÇÃO DE ENFERMAGEM"},
  {ocr: "NOTA DE SALA", padrao: "NOTA DE SALA"},
  {ocr: "OPME ORTESE E PROTESES", padrao: "OPME ÓRTESE E PRÓTESES"},
  {ocr: "PRESCRIÇÃO ENFERMAGEM", padrao: "PRESCRIÇÃO ENFERMAGEM"},
  {ocr: "PRESCRIÇÃO HEMODIÁLISE", padrao: "PRESCRIÇÃO HEMODIÁLISE"},
  {ocr: "PRESCRIÇÃO HEMOTERAPIA", padrao: "PRESCRIÇÃO HEMOTERAPIA"},
  {ocr: "PRESCRIÇÃO MÉDICA", padrao: "PRESCRIÇÃO MÉDICA"},
  {ocr: "PRESCRIÇÃO QUIMIOTERAPIA", padrao: "PRESCRIÇÃO QUIMIOTERAPIA"},
  {ocr: "RASTREABILIDADE", padrao: "RASTREABILIDADE"},
  {ocr: "RECUPERAÇÃO PÓS-ANESTÉSICA", padrao: "RECUPERAÇÃO PÓS-ANESTÉSICA"},
  {ocr: "REGISTRO DE ENFERMAGEM", padrao: "REGISTRO DE ENFERMAGEM"},
  {ocr: "TERMO DE DESISTÊNCIA", padrao: "TERMO DE DESISTÊNCIA DE TRATAMENTO A PEDIDO E CONTRA INDICAÇÃO MÉDICA"},
  {ocr: "TERMO DE CONSENTIMENTO", padrao: "TERMO DE CONSENTIMENTO"},
  {ocr: "TERMO DE DECLARAÇÃO DE MORTE ENG", padrao: "TERMO DE DECLARAÇÃO DE MORTE ENCEFÁLICA"},
  {ocr: "TERMO DE DOAÇÃO ORGÃOS", padrao: "TERMO DE DOAÇÃO DE ÓRGÃOS"},
  {ocr: "TERMO DE INTERNAÇÃO", padrao: "TERMO DE INTERNAÇÃO"},
  {ocr: "TERMO DE TRANSFERÊNCIA", padrao: "TERMO DE TRANSFERÊNCIA"}
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Gestão SAMIS')
    .addItem('🖥️ Abrir Portal SAMIS (BI)', 'abrirPainelGestao')
    .addSeparator()
    .addItem('📥 Extrair Dados das Abas', 'abrirPainelExtracao')
    .addItem('📊 Relatórios PDF/Excel (Legado)', 'abrirPainelEstatistica')
    .addToUi();
}

/* ==========================================================================
   1. FUNÇÕES DE INTERFACE
   ========================================================================== */
function abrirPainelGestao() {
  var template = HtmlService.createTemplateFromFile('PainelGestao');
  
  // -------------------------------------------------------------
  // INJEÇÃO DA BIBLIOTECA ESTÚDIO PDF (MODO PRO - LAZY LOAD)
  // -------------------------------------------------------------
  try {
    template.estudioPdfComponent = PDFPRO.getPdfComponent({
      hostAppName: "SAMIS HCPA",
      maxFileSizeMb: 50,
      theme: "modern"
    });
  } catch(e) {
    template.estudioPdfComponent = "<div class='p-5 text-center text-danger'>Erro ao carregar o Estúdio PDF: " + e.message + "</div>";
  }

  var html = template.evaluate()
    .setWidth(1280)
    .setHeight(850)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Contagem de Documentos Prontuários de 2026');
}

function abrirPainelExtracao() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('Extrator').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function abrirPainelEstatistica() {
  var html = HtmlService.createTemplateFromFile('Estatistica').evaluate().setWidth(1200).setHeight(750).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'SAMIS - Relatórios');
}

/* ==========================================================================
   2. API DE DADOS (BACKEND)
   ========================================================================== */

// Ache esta função e modifique o finalzinho dela:
function obterDadosIniciaisPainel() {
  try {
    var ss = (ID_PLANILHA && ID_PLANILHA !== "SEU_ID") ? SpreadsheetApp.openById(ID_PLANILHA) : SpreadsheetApp.getActiveSpreadsheet();
    var pUtil = ss.getSheetByName("Útil");
    if (!pUtil) throw new Error("Aba 'Útil' não encontrada.");

    var locais = pUtil.getRange("B3:B60").getValues().flat().filter(function(s) { return s !== "" && s != null; }).map(String);
    var tipos = pUtil.getRange("C3:C150").getValues().flat().filter(function(s) { return s !== "" && s != null; }).map(String);
    
    // Puxa o ranking que criamos na coluna I
    var vips = pUtil.getRange("I3:I150").getValues().flat().filter(function(s) { return s !== "" && s != null; }).map(String);

    return { 
      locais: locais.sort(), 
      tipos: tipos.sort(),
      rankingVip: vips
    };
  } catch (e) {
    console.error("Erro no obterDadosIniciaisPainel: " + e.message);
    return { erro: "Falha ao carregar dados da aba Útil: " + e.message };
  }
}

function pesquisarRegistros(filtros) {
  try {
    // CORREÇÃO: Usando a forma original de chamar a planilha para não dar erro
    var ss = SpreadsheetApp.openById(ID_PLANILHA);
    var sheet = ss.getSheetByName("Registro de Ocorrências");
    if (!sheet) throw new Error("Aba 'Registro de Ocorrências' não encontrada.");

    var data = sheet.getDataRange().getValues();
    var resultados = [];
    var limite = filtros.limite === "todos" ? data.length : parseInt(filtros.limite, 10);
    
    var fLocal = filtros.local;
    var filtrarLocais = (fLocal && fLocal !== "todos" && Array.isArray(fLocal) && fLocal.length > 0);
    var fTipo = filtros.tipo ? String(filtros.tipo).toLowerCase().trim() : "";
    var fTipoExato = filtros.tipoExato === true; // <-- Verifica se o usuário pediu termo exato
    var fInicio = filtros.dataInicio ? new Date(filtros.dataInicio + "T00:00:00") : null;
    var fFim = filtros.dataFim ? new Date(filtros.dataFim + "T23:59:59") : null;
    var totalItens = 0;

    // Loop reverso (do mais recente para o mais antigo)
    for (var i = data.length - 1; i > 0; i--) {
      if (resultados.length >= limite) break;
      var row = data[i];
      var rDataOriginal = row[0];
      
      var rDataObj = null;
      if (rDataOriginal instanceof Date) {
        rDataObj = rDataOriginal;
      } else if (typeof rDataOriginal === 'string') {
        var p = rDataOriginal.split('/');
        if (p.length === 3) rDataObj = new Date(p[2], p[1] - 1, p[0]);
      }

      var rLocal = String(row[1] || "").trim();
      var rTipo = String(row[2] || "").trim();

      // Filtro de Local
      if (filtrarLocais && fLocal.indexOf(rLocal) === -1) continue;
      
      // Filtro de Tipo (Exato vs Contém)
      if (fTipo !== "") {
        if (fTipoExato) {
          if (rTipo.toLowerCase() !== fTipo) continue; // Precisa ser exatamente igual
        } else {
          if (!rTipo.toLowerCase().includes(fTipo)) continue; // Basta conter o texto (padrão)
        }
      }
      
      // Filtro de Datas
      if (rDataObj) {
        if (fInicio && rDataObj < fInicio) continue;
        if (fFim && rDataObj > fFim) continue;
      }

      var qtd = parseInt(row[3], 10) || 0;
      totalItens += qtd;

      var dtFormatada = rDataOriginal;
      try { if (rDataObj) dtFormatada = Utilities.formatDate(rDataObj, ss.getSpreadsheetTimeZone() || "GMT-3", "dd/MM/yyyy"); } catch(e){}

      resultados.push([dtFormatada, rLocal, rTipo, qtd]);
    }
    return { success: true, dados: resultados, totalItens: totalItens };
  } catch (e) { return { erro: e.message }; }
}

function gerarDadosGraficoDinamico(config) {
  try {
    var ss = (ID_PLANILHA && ID_PLANILHA !== "SEU_ID") ? SpreadsheetApp.openById(ID_PLANILHA) : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Registro de Ocorrências");
    var data = sheet.getDataRange().getValues();
    var stats = {};
    var tz = ss.getSpreadsheetTimeZone();

    var fInicio = config.dataInicio ? new Date(config.dataInicio + "T00:00:00") : null;
    var fFim = config.dataFim ? new Date(config.dataFim + "T23:59:59") : null;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rData = row[0];
      var qtd = parseInt(row[3]) || 0;
      
      if (!(rData instanceof Date)) continue;
      if (fInicio && rData < fInicio) continue;
      if (fFim && rData > fFim) continue;

      var chave = "";
      if (config.eixo === "local") chave = String(row[1] || "Não Informado").trim();
      else if (config.eixo === "tipo") chave = String(row[2] || "Não Informado").trim();
      else if (config.eixo === "tempo") chave = Utilities.formatDate(rData, tz, "MM/yyyy");

      if (chave) stats[chave] = (stats[chave] || 0) + qtd;
    }

    var labels = Object.keys(stats);
    if (labels.length === 0) return { labels: [], data: [], aviso: "Nenhum dado no período" };

    // Ordenação inteligente
    if (config.eixo === "tempo") {
      labels.sort(function(a, b) {
        var pA = a.split('/'); var pB = b.split('/');
        return new Date(pA[1], pA[0]-1) - new Date(pB[1], pB[0]-1);
      });
    } else {
      labels.sort(function(a, b) { return stats[b] - stats[a]; });
    }
    
    if (config.eixo !== "tempo" && labels.length > 20) labels = labels.slice(0, 20);

    return { labels: labels, data: labels.map(function(k) { return stats[k]; }) };
  } catch (e) { 
    return { erro: "Erro no gráfico: " + e.message }; 
  }
}

/* ==========================================================================
   3. FUNÇÕES DE CADASTRO E SALVAMENTO
   ========================================================================== */

function salvarLoteOcorrencias(listaDados) {
  try {
    var ss = SpreadsheetApp.openById(ID_PLANILHA);
    var sheet = ss.getSheetByName("Registro de Ocorrências");
    if (!sheet) throw new Error("Aba 'Registro de Ocorrências' não encontrada.");

    // --- CORREÇÃO DE POSIÇÃO ---
    // Em vez de confiar no getLastRow(), lemos a coluna A inteira
    // e procuramos a última célula que realmente tem texto.
    
    var colunaA = sheet.getRange("A1:A").getValues(); // Pega todos os dados da coluna A
    var ultimaLinhaReal = 0;

    // Varre de baixo para cima (do fim para o começo)
    for (var i = colunaA.length - 1; i >= 0; i--) {
      if (colunaA[i][0] !== "" && colunaA[i][0] != null) {
        ultimaLinhaReal = i + 1; // Encontrou dado! Essa é a última linha usada.
        break;
      }
    }
    
    // A próxima linha será a imediatamente seguinte
    var proximaLinha = ultimaLinhaReal + 1; 
    
    // Segurança: Nunca escreve na linha 1 (cabeçalho)
    if (proximaLinha < 2) proximaLinha = 2;
    // ----------------------------

    var matrizParaColar = listaDados.map(function(item) {
      return [
        new Date(item.data + "T12:00:00"), 
        String(item.local).trim(), 
        String(item.tipo).trim(), 
        parseInt(item.qtd)
      ];
    });

    if (matrizParaColar.length > 0) {
      sheet.getRange(proximaLinha, 1, matrizParaColar.length, 4).setValues(matrizParaColar);
    }
    
    return { success: true, qtd: matrizParaColar.length };
  } catch (e) {
    return { erro: e.message };
  }
}

function cadastrarItemUtil(categoria, valor) {
  var ss = SpreadsheetApp.openById(ID_PLANILHA);
  var sheet = ss.getSheetByName("Útil");
  var col = (categoria === 'local') ? 2 : 3;
  var range = sheet.getRange((categoria === 'local') ? "B3:B60" : "C3:C150");
  var values = range.getValues();
  
  var novoValor = String(valor).trim().toUpperCase();
  
  // Verifica duplicidade
  for (var k = 0; k < values.length; k++) {
    if (values[k][0] && String(values[k][0]).trim().toUpperCase() === novoValor) {
      throw new Error("Este item já está cadastrado!");
    }
  }

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === "") {
      sheet.getRange(3 + i, col).setValue(String(valor).trim());
      return { success: true };
    }
  }
  throw new Error("Limite de linhas atingido para esta categoria.");
}

/* ==========================================================================
   4. FUNÇÕES DO EXTRATOR DE DADOS (LEGADO)
   ========================================================================== */

function obterListaDeAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNames = [];
  var ignoreSheets = ["Relatório Consolidado", "Registro de Ocorrências", "Útil"];
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (ignoreSheets.indexOf(name) === -1) { sheetNames.push(name); }
  });
  return sheetNames;
}

function prepararRelatorio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheetName = "Registro de Ocorrências"; 
  var reportSheet = ss.getSheetByName(reportSheetName);
  if (!reportSheet) { reportSheet = ss.insertSheet(reportSheetName); }
  if (reportSheet.getLastRow() === 0) {
    reportSheet.appendRow(["Data", "Local", "Tipo de Documento", "Quantidade"]);
    reportSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f3f3f3");
    reportSheet.setFrozenRows(1);
  }
  return "Preparado. Dados existentes preservados.";
}

function processarAbaIndividual(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var reportSheet = ss.getSheetByName("Registro de Ocorrências");
  if (!sheet) return { success: false, log: "Erro: Aba não encontrada." };
  try {
    var rangeData = sheet.getRange("A1:AM50").getValues();
    var headerRow = rangeData[0];
    var colunasDatas = [];
    for (var col = 2; col <= 38; col++) {
      var cabecalho = headerRow[col];
      if (cabecalho && String(cabecalho).toUpperCase().trim() !== "" && String(cabecalho).toUpperCase().indexOf("TOTAL") === -1) {
        var dataFormatada = cabecalho;
        if (cabecalho instanceof Date) {
            dataFormatada = Utilities.formatDate(cabecalho, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        }
        colunasDatas.push({ index: col, valor: dataFormatada });
      }
    }
    var registrosEncontrados = [];
    for (var r = 1; r < rangeData.length; r++) {
      var tipoDocumento = String(rangeData[r][0]).trim();
      if (!tipoDocumento || tipoDocumento.toUpperCase().includes("TOTAL")) continue;
      colunasDatas.forEach(function(objData) {
        var qtd = rangeData[r][objData.index];
        if (typeof qtd === 'number' && qtd > 0) {
          registrosEncontrados.push([objData.valor, sheetName, tipoDocumento, qtd]);
        }
      });
    }
    if (registrosEncontrados.length > 0) {
      reportSheet.getRange(reportSheet.getLastRow() + 1, 1, registrosEncontrados.length, 4).setValues(registrosEncontrados);
    }
    return { success: true, count: registrosEncontrados.length, log: "✅ " + sheetName + ": " + registrosEncontrados.length + " registros." };
  } catch (e) { return { success: false, log: "❌ Erro em " + sheetName + ": " + e.message }; }
}

/* ==========================================================================
   5. MOTOR DE ESTATÍSTICA E RELATÓRIOS PDF/EXCEL (LEGADO)
   ========================================================================== */

function processarRelatorio(tipoRelatorio, mes, ano, filtroExtra) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Registro de Ocorrências");
  if (!sheet) throw new Error("Aba 'Registro de Ocorrências' não encontrada. Extraia os dados primeiro.");

  var data = sheet.getDataRange().getValues();
  var matriz = {}, colunasHeader = new Set(), linhasHeader = new Set(), totalGeral = 0;
  var mesInt = parseInt(mes), anoInt = parseInt(ano);

  for (var i = 1; i < data.length; i++) {
    var row = data[i], rowDate = row[0], rowLocal = String(row[1]).trim(), rowTipo = String(row[2]).trim(), rowQtd = parseInt(row[3]) || 0;
    if (!rowLocal || !rowTipo) continue;
    
    var rMes = -1, rAno = -1, rDia = -1;
    if (rowDate instanceof Date) { 
      rMes = rowDate.getMonth() + 1; rAno = rowDate.getFullYear(); rDia = rowDate.getDate(); 
    } else if (typeof rowDate === 'string') {
       var parts = rowDate.split('/');
       if (parts.length >= 2) { rDia = parseInt(parts[0]); rMes = parseInt(parts[1]); rAno = (parts.length===3)?parseInt(parts[2]):anoInt; }
    }

    if (rAno !== anoInt) continue;
    var rowKey = "", colKey = "";
    
    if (tipoRelatorio == 1) { if (rMes !== mesInt || rowLocal !== filtroExtra) continue; colKey = (rDia < 10 ? '0' : '') + rDia + "/" + (rMes < 10 ? '0' : '') + rMes; rowKey = rowTipo; }
    else if (tipoRelatorio == 2) { if (rMes !== mesInt) continue; colKey = rowTipo; rowKey = rowLocal; }
    else if (tipoRelatorio == 3) { colKey = obterNomeMes(rMes); rowKey = rowTipo; }
    else if (tipoRelatorio == 4) { if (rMes !== mesInt) continue; colKey = (rDia < 10 ? '0' : '') + rDia + "/" + (rMes < 10 ? '0' : '') + rMes; rowKey = rowLocal; }
    
    if (rowKey && colKey) {
      if (!matriz[rowKey]) matriz[rowKey] = {};
      matriz[rowKey][colKey] = (matriz[rowKey][colKey] || 0) + rowQtd;
      colunasHeader.add(colKey); linhasHeader.add(rowKey); totalGeral += rowQtd;
    }
  }
  var cols = Array.from(colunasHeader).sort();
  if (tipoRelatorio == 3) {
    var ordemMeses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    cols.sort(function(a,b) { return ordemMeses.indexOf(a) - ordemMeses.indexOf(b); });
  }
  return { matriz: matriz, cols: cols, rows: Array.from(linhasHeader).sort(), totalGeral: totalGeral, titulo: obterTituloRelatorio(tipoRelatorio, filtroExtra), periodo: (tipoRelatorio == 3 ? ano : mes + "/" + ano) };
}

function obterNomeMes(num) { return ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"][num-1] || "Desc"; }
function obterTituloRelatorio(t, e) {
  var m = { 1: "Visão Mensal: Tipos por Dia ("+e+")", 2: "Visão Mensal: Tipos por Local", 3: "Visão Anual Consolidada", 4: "Visão Mensal: Produção Diária por Local" };
  return m[t] || "Relatório SAMIS";
}

function gerarPDFGenerico(tipoRelatorio, mes, ano, filtroExtra) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = processarRelatorio(tipoRelatorio, mes, ano, filtroExtra);
  if (d.rows.length === 0) throw new Error("Sem dados.");
  var html = `<html><head><style>@page{size:A3 landscape;margin:10mm;}body{font-family:Helvetica;font-size:8px;}table{width:100%;border-collapse:collapse;}th,td{border:1px solid #ccc;padding:3px;text-align:center;}th{background:#e8f0fe;}</style></head><body><h2>${d.titulo} - ${d.periodo}</h2><table><thead><tr><th>DESCRIÇÃO</th>`;
  d.cols.forEach(function(c) { html += `<th>${c}</th>`; });
  html += `<th>TOTAL</th></tr></thead><tbody>`;
  d.rows.forEach(function(r) {
    var sl = 0; html += `<tr><td>${r}</td>`;
    d.cols.forEach(function(c) { var v = d.matriz[r][c] || 0; sl += v; html += `<td>${v||'-'}</td>`; });
    html += `<td>${sl}</td></tr>`;
  });
  html += `</tbody></table><div style='text-align:right;font-size:7px;margin-top:10px;'>Gerado em ${Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm")}</div></body></html>`;
  var blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
  blob.setName(d.titulo.replace(/\s/g, "_") + ".pdf");
  return DriveApp.createFile(blob).getUrl();
}

function gerarExcelRelatorio(tipoRelatorio, mes, ano, filtroExtra) {
  var d = processarRelatorio(tipoRelatorio, mes, ano, filtroExtra);
  var ssNew = SpreadsheetApp.create("SAMIS_Export_" + d.titulo);
  var sh = ssNew.getActiveSheet();
  sh.appendRow(["DESCRIÇÃO"].concat(d.cols).concat(["TOTAL"]));
  d.rows.forEach(function(r) {
    var row = [r], sl = 0;
    d.cols.forEach(function(c) { var v = d.matriz[r][c] || 0; sl += v; row.push(v || ""); });
    row.push(sl); sh.appendRow(row);
  });
  SpreadsheetApp.flush();
  return "https://docs.google.com/spreadsheets/d/" + ssNew.getId() + "/export?format=xlsx";
}

function obterLocaisUnicos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUtil = ss.getSheetByName("Útil");
  var locs = [];
  if (sheetUtil) {
    var vals = sheetUtil.getRange("B3:B60").getValues();
    vals.forEach(function(r){ if(r[0]) locs.push(String(r[0]).trim()); });
  }
  return (locs.length > 0) ? locs.sort() : [];
}

function obterDadosGrafico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Registro de Ocorrências");
  var data = sheet.getDataRange().getValues();
  var stats = {};
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0], q = parseInt(data[i][3]) || 0;
    if (d instanceof Date) {
      var k = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "MM/yyyy");
      stats[k] = (stats[k] || 0) + q;
    }
  }
  return stats;
}

/* ==========================================================================
   6. NOVA ENGINE DE OCR (VIA OCR.SPACE API)
   ========================================================================== */

function processarUploadOCR(arquivosBase64) {
  Logger.log("[OCR] Iniciando processamento via OCR.space...");
  
  if (OCR_API_KEY === "COLE_SUA_API_KEY_AQUI") {
    return { success: false, erro: "Configure a API KEY no arquivo Code.gs!" };
  }

  var resultados = [];

  try {
    arquivosBase64.forEach(function(arq) {
      Logger.log("[OCR] Enviando arquivo: " + arq.name);
      
      var base64Full = "data:" + arq.mimeType + ";base64," + arq.data;
      
      var payload = {
        "base64Image": base64Full,
        "apikey": OCR_API_KEY,
        "language": "por",
        "isTable": "true",
        "OCREngine": "2",
        "scale": "true"
      };

      var options = {
        "method": "post",
        "payload": payload,
        "muteHttpExceptions": true
      };

      var response = UrlFetchApp.fetch("https://api.ocr.space/parse/image", options);
      var json = JSON.parse(response.getContentText());

      if (json.IsErroredOnProcessing) {
        throw new Error("Erro na API OCR: " + json.ErrorMessage);
      }

      var textoCompleto = "";
      if (json.ParsedResults && json.ParsedResults.length > 0) {
        json.ParsedResults.forEach(function(page) {
          textoCompleto += page.ParsedText + "\n";
        });
      }
      
      Logger.log("[OCR] Texto extraído (início): " + textoCompleto.substring(0, 100));

      var itens = interpretarTextoFixo(textoCompleto);
      if (itens.length > 0) resultados = resultados.concat(itens);
    });

    return { success: true, dados: resultados };

  } catch (e) {
    Logger.log("[ERRO FATAL OCR] " + e.message);
    return { success: false, erro: "Falha no OCR: " + e.message };
  }
}

function interpretarTextoFixo(texto) {
  var itensEncontrados = [];
  var textoUpper = texto.toUpperCase();

  MAPA_33_ITENS.forEach(function(item) {
    var termoBusca = item.ocr.toUpperCase();
    var indexTermo = textoUpper.indexOf(termoBusca);
    
    if (indexTermo !== -1) {
      // Pega o trecho logo após o nome do documento
      var trechoAposTermo = textoUpper.substring(indexTermo + termoBusca.length, indexTermo + termoBusca.length + 150);
      
      // Limpeza: Troca quebras de linha por espaço
      var trechoLimpo = trechoAposTermo.replace(/[\r\n]+/g, " ");
      
      // Regex que ignora pontos, traços e sujeira e pega o primeiro número
      var matchNumero = trechoLimpo.match(/^[^\d]*(\d+)/);

      if (matchNumero) {
        var qtd = parseInt(matchNumero[1]);
        if (qtd > 0 && qtd < 5000) {
          itensEncontrados.push({
            tipo: item.padrao, 
            qtd: qtd,
            match: true
          });
        }
      }
    }
  });
  
  return itensEncontrados;
}

/* ==========================================================================
   SISTEMA DE LOGS EM TEMPO REAL (CACHE)
   ========================================================================== */
function lerLogsControle(taskId) {
  return CacheService.getUserCache().get(taskId) || "Aguardando servidor...";
}

function registrarLog(taskId, msg, reset) {
  var cache = CacheService.getUserCache();
  var current = reset ? "" : (cache.get(taskId) || "");
  var time = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "GMT-3", "HH:mm:ss");
  var newLog = current + "<div><span style='color:#888'>[" + time + "]</span> " + msg + "</div>";
  cache.put(taskId, newLog, 600); // Guarda por 10 min
}

/* ==========================================================================
   FUNÇÃO: BUSCAR DADOS DE CONTROLES (OTIMIZADA + LIVE LOG)
   ========================================================================== */
function buscarDadosControles(mesAbreviado, taskId) {
  console.log("▶️ INICIANDO BUSCA DE CONTROLES | Mês alvo: " + mesAbreviado);
  registrarLog(taskId, "▶️ INICIANDO BUSCA | Mês alvo: " + mesAbreviado, true);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetUtil = ss.getSheetByName("Útil");
    if (!sheetUtil) throw new Error("Aba 'Útil' não encontrada.");

    var mes = parseInt(mesAbreviado, 10);
    if (isNaN(mes) || mes < 1 || mes > 12) throw new Error("Mês inválido.");
    
    var linhaUrl = mes + 2; 
    var urlPlanilha = sheetUtil.getRange("G" + linhaUrl).getValue();
    console.log("🔗 Endereço lido na célula G" + linhaUrl + ": " + urlPlanilha);
    registrarLog(taskId, "🔗 Lendo endereço na célula G" + linhaUrl);
    
    if (!urlPlanilha || String(urlPlanilha).trim() === "") {
      throw new Error("Nenhum endereço configurado na célula G" + linhaUrl + " para este mês.");
    }

    var idTarget = String(urlPlanilha).trim();
    var idMatch = idTarget.match(/[-\w]{25,}/);
    if (idMatch) idTarget = idMatch[0];
    
    console.log("📂 ID extraído: " + idTarget + " | Tentando abrir planilha alvo...");
    registrarLog(taskId, "📂 Abrindo planilha alvo...");

    var targetSs;
    try { 
      targetSs = SpreadsheetApp.openById(idTarget); 
    } catch(err) { 
      throw new Error("Acesso negado ou ID inválido. Verifique o link na G" + linhaUrl); 
    }
    
    console.log("✅ Planilha alvo aberta com sucesso: " + targetSs.getName());
    registrarLog(taskId, "✅ Planilha acessada: " + targetSs.getName());

    var abas = targetSs.getSheets();
    console.log("📑 Total de abas na planilha alvo: " + abas.length);
    registrarLog(taskId, "📑 Total de abas na planilha alvo: " + abas.length);
    
    var resultados = {};
    var anoAtual = new Date().getFullYear();
    var abasProcessadas = 0;
    var totalRegistrosEncontrados = 0;

    // Loop super rápido (for clássico ao invés de forEach)
    for (var i = 0; i < abas.length; i++) {
      var aba = abas[i];
      var nomeAba = aba.getName();
      
      var matchData = nomeAba.match(/^(\d{2})_(\d{2})/);
      if (!matchData) continue; 

      var dia = matchData[1];
      var mesDaAba = parseInt(matchData[2], 10); 
      if (mesDaAba !== mes) continue; 

      abasProcessadas++;
      console.log("⏳ Lendo aba correspondente: " + nomeAba + "...");
      registrarLog(taskId, "⏳ Lendo aba correspondente: <b>" + nomeAba + "</b>...");

      var dataFormatada = dia + "/" + matchData[2] + "/" + anoAtual;
      var lastRow = aba.getLastRow();
      
      if (lastRow < 1) {
        console.log("⚠️ Aba " + nomeAba + " está vazia. Pulando.");
        continue;
      }
      
      var dadosRS = aba.getRange(1, 18, lastRow, 2).getValues();
      var encontradosNaAba = 0;

      for (var j = 0; j < dadosRS.length; j++) {
        var colR = dadosRS[j][0];
        if (!colR || colR === "") continue;
        
        var textoR = String(colR).trim().toUpperCase();

        if (textoR.indexOf("1409") === 0) {
          var colS = parseFloat(dadosRS[j][1]);
          
          if (!isNaN(colS) && colS > 0) {
            var localEncontrado = "";
            if (textoR.indexOf("-CCA") !== -1) localEncontrado = "CCA- CENT. CIR. AMB.";
            else if (textoR.indexOf("-HEMOD") !== -1) localEncontrado = "HEMODINÂMICA";
            else if (textoR.indexOf("-BLOCO") !== -1) localEncontrado = "BLOCO";

            if (localEncontrado !== "") {
              var chave = dataFormatada + "|" + localEncontrado;
              if (!resultados[chave]) {
                resultados[chave] = { data: dataFormatada, local: localEncontrado, tipo: "NOTA DE SALA", qtd: 0 };
              }
              resultados[chave].qtd += colS;
              encontradosNaAba++;
              totalRegistrosEncontrados++;
            }
          }
        }
      }
      console.log("✔️ Aba " + nomeAba + " finalizada. Registros úteis nela: " + encontradosNaAba);
      registrarLog(taskId, "✔️ Aba finalizada. Registros úteis nela: " + encontradosNaAba);
    }

    var arrayFinal = Object.keys(resultados).map(function(key) { return resultados[key]; });
    
    console.log("🏁 VARREDURA CONCLUÍDA!");
    registrarLog(taskId, "🏁 VARREDURA CONCLUÍDA!");
    
    console.log("📊 Abas do mês lidas: " + abasProcessadas);
    console.log("🧩 Linhas totais encontradas (antes de agrupar): " + totalRegistrosEncontrados);
    console.log("📦 Agrupamentos gerados: " + arrayFinal.length);
    registrarLog(taskId, "📊 Abas lidas: " + abasProcessadas + " | 🧩 Total Linhas: " + totalRegistrosEncontrados + " | 📦 Agrupados: " + arrayFinal.length);

    return { success: true, dados: arrayFinal };

  } catch (e) {
    console.error("❌ ERRO FATAL NA BUSCA: " + e.message);
    registrarLog(taskId, "<span style='color:#ff4a4a'>❌ ERRO: " + e.message + "</span>");
    return { erro: e.message };
  }
}

function FORCAR_AUTORIZACAO() {
  var teste = Drive.Files.list({maxResults: 1});
  console.log("✅ Permissões concedidas com sucesso!");
}

/* ==========================================================================
   7. GERADOR DE RANKING (VIPs SOB DEMANDA)
   ========================================================================== */
function atualizarRankingDeTipos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetRegistros = ss.getSheetByName("Registro de Ocorrências");
    var sheetUtil = ss.getSheetByName("Útil");
    
    if (!sheetRegistros || !sheetUtil) throw new Error("Abas necessárias não encontradas.");

    // 1. Puxa todos os dados de registro (Coluna C = Tipo, Coluna D = Qtd)
    var data = sheetRegistros.getDataRange().getValues();
    var contagem = {};

    // Começa de 1 para pular o cabeçalho
    for (var i = 1; i < data.length; i++) {
      var tipo = String(data[i][2]).trim().toUpperCase();
      var qtd = parseInt(data[i][3]) || 0;
      
      if (tipo !== "" && tipo.indexOf("TOTAL") === -1) {
        contagem[tipo] = (contagem[tipo] || 0) + qtd;
      }
    }

    // 2. Transforma o objeto em um array classificável
    var arrayRanking = [];
    for (var key in contagem) {
      arrayRanking.push({ nome: key, total: contagem[key] });
    }

    // 3. Ordena do maior (mais ocorrências) para o menor
    arrayRanking.sort(function(a, b) {
      return b.total - a.total;
    });

    // 4. Salva o ranking na aba "Útil" (Agora nas colunas I e J)
    // Limpa a área primeiro
    sheetUtil.getRange("I2:J150").clearContent();
    sheetUtil.getRange("I2").setValue("RANKING (VIPs)");
    sheetUtil.getRange("J2").setValue("Vezes Usado");
    sheetUtil.getRange("I2:J2").setFontWeight("bold");

    // Prepara a matriz para salvar de uma vez só (super rápido)
    var matrizSalvamento = [];
    for (var j = 0; j < arrayRanking.length; j++) {
      if (j >= 148) break; // Limite de linhas
      matrizSalvamento.push([arrayRanking[j].nome, arrayRanking[j].total]);
    }

    if (matrizSalvamento.length > 0) {
      // 3 = Linha 3, 9 = Coluna I, length = total de linhas, 2 = total de colunas
      sheetUtil.getRange(3, 9, matrizSalvamento.length, 2).setValues(matrizSalvamento);
    }

    return { success: true, totalProcessado: matrizSalvamento.length };

  } catch (e) {
    return { erro: e.message };
  }
}

/* ==========================================================================
   8. GERADOR DE ETIQUETAS (LOG E INTELIGÊNCIA)
   ========================================================================== */
function obterProximoNumeroEtiqueta(tipo, codigo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log Etiquetas");
    
    // Valores de "Último Impresso" fornecidos (Sugestão será +1)
    const defaults = {
      "Complementar": { "14": 11666, "10": 1765, "03": 8818, "02": 4810 },
      "Convênios": { "03": 2559 }
    };

    let ultimoNoSistema = 0;

    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim() === tipo && String(data[i][2]).trim() === String(codigo)) {
          var rFinal = parseInt(data[i][4], 10) || 0;
          if (rFinal > ultimoNoSistema) ultimoNoSistema = rFinal;
        }
      }
    }

    // Se o que tem no LOG for menor que o "pré-salvo", usa o pré-salvo
    let baseComp = (defaults[tipo] && defaults[tipo][String(codigo)]) ? defaults[tipo][String(codigo)] : 0;
    let valorReferencia = Math.max(ultimoNoSistema, baseComp);

    return valorReferencia + 1;
  } catch(e) {
    return 1;
  }
}

function salvarLogEtiquetas(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log Etiquetas");
    
    if (!sheet) {
      sheet = ss.insertSheet("Log Etiquetas");
      sheet.appendRow(["Data da Impressão", "Tipo", "Código", "Nº Inicial", "Nº Final", "Qtd Impressa"]);
    }
    
    sheet.appendRow([new Date(), dados.tipo, dados.codigo, dados.inicial, dados.final, dados.quantidade]);
    return { success: true };
  } catch(e) {
    return { erro: e.message };
  }
}

// 7. Busca o histórico completo de impressões
function obterHistoricoEtiquetas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log Etiquetas");
    
    if (!sheet) return []; // Se não existir, retorna vazio
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // Se só tiver cabeçalho, retorna vazio
    
    var historico = [];
    
    // Varre a planilha do final para o começo (para os mais recentes ficarem no topo)
    for (var i = data.length - 1; i > 0; i--) {
      // Formata a data (Ex: 15/05/2026 14:30)
      var rawDate = data[i][0];
      var dataFormatada = "";
      if (rawDate instanceof Date) {
        dataFormatada = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      } else {
        dataFormatada = String(rawDate);
      }
      
      historico.push({
        data: dataFormatada,
        tipo: data[i][1],
        codigo: data[i][2],
        inicial: data[i][3],
        final: data[i][4],
        qtd: data[i][5]
      });
    }
    
    return historico;
  } catch(e) {
    return { erro: e.message };
  }
}

// =========================================================================
// PONTE OBRIGATÓRIA DA BIBLIOTECA DE PDF (API ROUTER)
// =========================================================================
function pdfStudioGlobalRouter(action, payload) {
  return PDFPRO.handleApiRequest(action, payload);
}

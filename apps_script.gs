/**
 * Recebe as confirmações do convite e grava na planilha.
 *
 * COMO USAR
 * 1. Crie uma planilha no Google Sheets (ex.: "Confirmações - Ísis Sayuri").
 * 2. Menu Extensões > Apps Script. Apague o conteúdo e cole este código. Salve.
 * 3. Implantar > Nova implantação > tipo "App da Web":
 *      - Executar como: Eu
 *      - Quem tem acesso: Qualquer pessoa
 *    Autorize quando o Google pedir e copie a URL terminada em /exec.
 * 4. Cole essa URL em CONFIG.urlPlanilha dentro do convite.html.
 * (Se editar este código depois, use Implantar > Gerenciar implantações > Nova versão.)
 */
const ABA = 'Confirmações';
const CABECALHO = ['Data/hora', 'Nome', 'Presença', 'Adultos', 'Crianças', 'Total', 'WhatsApp', 'Mensagem'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = (e && e.parameter) || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(CABECALHO);
      sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold').setBackground('#f4a9c9');
      sh.setFrozenRows(1);
    }
    const limpa = v => {            // evita fórmulas injetadas (=, +, -, @)
      v = String(v || '').slice(0, 300);
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    };
    const vai = p.presenca === 'Sim';
    const adultos = vai ? Math.min(parseInt(p.adultos, 10) || 0, 20) : 0;
    const criancas = vai ? Math.min(parseInt(p.criancas, 10) || 0, 20) : 0;
    sh.appendRow([
      new Date(), limpa(p.nome), vai ? 'Sim' : 'Não',
      adultos, criancas, adultos + criancas,
      limpa(p.whatsapp), limpa(p.mensagem)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Abrir a URL /exec no navegador só mostra que está no ar.
function doGet() {
  return ContentService.createTextOutput('Convite: serviço de confirmação ativo.');
}

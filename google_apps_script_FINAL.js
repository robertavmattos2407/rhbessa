/**
 * GOOGLE APPS SCRIPT — CONTABILIDADE BESSA
 * Sistema de Gestão de Pessoas v2.0
 *
 * COMO IMPLANTAR:
 * 1. Abra sua planilha Google Sheets
 * 2. Extensões > Apps Script
 * 3. Apague tudo e cole este código
 * 4. Salve (Ctrl+S)
 * 5. Clique em "Implantar" > "Nova implantação"
 * 6. Tipo: "Aplicativo da Web"
 * 7. Executar como: "Eu"
 * 8. Quem tem acesso: "Qualquer pessoa"
 * 9. Clique em "Implantar" e copie a URL gerada
 * 10. Cole a URL no lugar da constante API nos 3 arquivos HTML
 */

// ─── CONFIGURAÇÃO ───────────────────────────────────────
const ABA_DISC      = 'disc_clima';        // formulário semestral
const ABA_AVALIACAO = 'avaliacao';         // formulário anual
const ABA_GESTOR    = 'gestor_notas';      // notas do gestor

// ─── ROTEADOR PRINCIPAL ─────────────────────────────────
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const action = dados.action || 'saveForm';

    if (action === 'updateGestor') return salvarNotasGestor(dados);
    if (action === 'saveAvaliacao') return salvarAvaliacao(dados);
    return salvarDiscClima(dados);

  } catch(err) {
    return resposta({ ok: false, erro: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'get';
    if (action === 'get') return retornarTodosColaboradores();
    if (action === 'ping') return resposta({ ok: true, status: 'online', data: new Date().toISOString() });
    return resposta({ ok: false, erro: 'Ação desconhecida' });
  } catch(err) {
    return resposta({ ok: false, erro: err.toString() });
  }
}

// ─── GET: retornar colaboradores mesclados para o painel ──
function retornarTodosColaboradores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ler todas as abas
  const discRows     = lerAba(ss, ABA_DISC);
  const avalRows     = lerAba(ss, ABA_AVALIACAO);
  const gestorRows   = lerAba(ss, ABA_GESTOR);

  // Índice por nome normalizado (mais recente por pessoa)
  const mapa = {};

  // 1. Dados DISC+Clima (base)
  discRows.forEach(r => {
    const key = norm(r.nome || '');
    if (!key) return;
    if (!mapa[key] || r.data > (mapa[key].data || '')) {
      mapa[key] = { ...r };
    }
  });

  // 2. Avaliação anual (complementa)
  avalRows.forEach(r => {
    const key = norm(r.nome || '');
    if (!key) return;
    if (!mapa[key]) mapa[key] = {};
    if (r.data > (mapa[key].data_avaliacao || '')) {
      Object.assign(mapa[key], r);
    }
  });

  // 3. Notas do gestor (complementa)
  gestorRows.forEach(r => {
    const key = norm(r.nome || '');
    if (!key) return;
    if (!mapa[key]) mapa[key] = {};
    Object.assign(mapa[key], r);
  });

  const lista = Object.values(mapa);
  return resposta({ ok: true, data: lista, total: lista.length });
}

// ─── POST: salvar DISC + Clima ────────────────────────────
function salvarDiscClima(dados) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = getOuCriarAba(ss, ABA_DISC, [
    'data','nome','setor','nivel','tipo_avaliacao',
    'disc_dominante','disc_D','disc_I','disc_S','disc_C',
    'clima_carga','clima_autonomia','clima_clareza','clima_relacao',
    'clima_reconhec','clima_prazos','clima_equilibrio','clima_psico','clima_media'
  ]);

  const linha = [
    dados.data || new Date().toISOString(),
    dados.nome || '',
    dados.setor || '',
    dados.nivel || '',
    dados.tipo_avaliacao || '',
    dados.perfil_dominante || dados.disc_dominante || '',
    dados.disc_D || 0,
    dados.disc_I || 0,
    dados.disc_S || 0,
    dados.disc_C || 0,
    dados.clima_carga || dados.carga || '',
    dados.clima_autonomia || dados.autonomia || '',
    dados.clima_clareza || dados.clareza || '',
    dados.clima_relacao || dados.relacao || '',
    dados.clima_reconhec || dados.reconhec || '',
    dados.clima_prazos || dados.prazos || '',
    dados.clima_equilibrio || dados.equilibrio || '',
    dados.clima_psico || dados.psico || '',
    dados.clima || ''
  ];

  aba.appendRow(linha);
  return resposta({ ok: true, mensagem: 'DISC e Clima salvos!' });
}

// ─── POST: salvar Avaliação Anual ─────────────────────────
function salvarAvaliacao(dados) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = getOuCriarAba(ss, ABA_AVALIACAO, [
    'data','nome','setor','nivel','tipo_avaliacao',
    'disc_dominante','disc_D','disc_I','disc_S','disc_C',
    'clima_carga','clima_autonomia','clima_clareza','clima_relacao',
    'clima_reconhec','clima_prazos','clima_equilibrio','clima_psico','clima_media',
    // Comportamentais (autoavaliação)
    'c0','c1','c2','c3','c4','c5','c6','c7','c8',
    // Técnicos (t0-t3 para analista, e0-e8 para especialista, g0-g8 para gestor)
    't0','t1','t2','t3',
    'e0','e1','e2','e3','e4','e5','e6','e7','e8',
    'g0','g1','g2','g3','g4','g5','g6','g7','g8',
    // PDI do colaborador
    'pdi1','pdi2','pdi3'
  ]);

  const linha = [
    dados.data || new Date().toISOString(),
    dados.nome || '',
    dados.setor || '',
    dados.nivel || '',
    dados.tipo_avaliacao || '',
    dados.perfil_dominante || '',
    dados.disc_D || 0, dados.disc_I || 0, dados.disc_S || 0, dados.disc_C || 0,
    dados.clima_carga || '', dados.clima_autonomia || '', dados.clima_clareza || '',
    dados.clima_relacao || '', dados.clima_reconhec || '', dados.clima_prazos || '',
    dados.clima_equilibrio || '', dados.clima_psico || '', dados.clima || '',
    // Comportamentais
    ...['c0','c1','c2','c3','c4','c5','c6','c7','c8'].map(k => dados[k] || ''),
    // Técnicos
    ...['t0','t1','t2','t3'].map(k => dados[k] || ''),
    ...['e0','e1','e2','e3','e4','e5','e6','e7','e8'].map(k => dados[k] || ''),
    ...['g0','g1','g2','g3','g4','g5','g6','g7','g8'].map(k => dados[k] || ''),
    // PDI
    dados.pdi1 || '', dados.pdi2 || '', dados.pdi3 || ''
  ];

  aba.appendRow(linha);
  return resposta({ ok: true, mensagem: 'Avaliação anual salva!' });
}

// ─── POST: salvar notas do gestor ─────────────────────────
function salvarNotasGestor(dados) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = getOuCriarAba(ss, ABA_GESTOR, [
    'data_avaliacao','nome','gestor_avaliado','gestor_data','gestor_obs_pdi',
    // Comportamentais gestor
    'gestor_c0','gestor_c1','gestor_c2','gestor_c3','gestor_c4',
    'gestor_c5','gestor_c6','gestor_c7','gestor_c8',
    // Técnicos gestor (analista t0-t3, especialista e0-e8, gestor g0-g8)
    'gestor_t0','gestor_t1','gestor_t2','gestor_t3',
    'gestor_e0','gestor_e1','gestor_e2','gestor_e3','gestor_e4',
    'gestor_e5','gestor_e6','gestor_e7','gestor_e8',
    'gestor_g0','gestor_g1','gestor_g2','gestor_g3','gestor_g4',
    'gestor_g5','gestor_g6','gestor_g7','gestor_g8',
    'tipo_avaliacao'
  ]);

  // Verificar se já existe linha para este colaborador (atualizar em vez de duplicar)
  const valores = aba.getDataRange().getValues();
  const cab     = valores[0] || [];
  const colNome = cab.indexOf('nome');
  const nomeNorm = norm(dados.nome || '');

  // Encontrar linha existente
  let linhaExistente = -1;
  for (let i = 1; i < valores.length; i++) {
    if (norm(String(valores[i][colNome] || '')) === nomeNorm) {
      linhaExistente = i + 1; // +1 porque getRange é 1-based
      break;
    }
  }

  const linha = [
    dados.data || new Date().toISOString(),
    dados.nome || '',
    'SIM',
    dados.gestor_data || new Date().toISOString(),
    dados.gestor_obs_pdi || '',
    ...['c0','c1','c2','c3','c4','c5','c6','c7','c8'].map(k => dados['gestor_'+k] || ''),
    ...['t0','t1','t2','t3'].map(k => dados['gestor_'+k] || ''),
    ...['e0','e1','e2','e3','e4','e5','e6','e7','e8'].map(k => dados['gestor_'+k] || ''),
    ...['g0','g1','g2','g3','g4','g5','g6','g7','g8'].map(k => dados['gestor_'+k] || ''),
    dados.tipo_avaliacao || ''
  ];

  if (linhaExistente > 0) {
    // Atualizar linha existente
    aba.getRange(linhaExistente, 1, 1, linha.length).setValues([linha]);
  } else {
    aba.appendRow(linha);
  }

  return resposta({ ok: true, mensagem: 'Avaliação do gestor salva!' });
}

// ─── UTILITÁRIOS ──────────────────────────────────────────
function norm(str) {
  return String(str || '').toLowerCase().trim()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOuCriarAba(ss, nome, cabecalho) {
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.appendRow(cabecalho);
    // Formatar cabeçalho
    const h = aba.getRange(1, 1, 1, cabecalho.length);
    h.setBackground('#0A1F28');
    h.setFontColor('#FFFFFF');
    h.setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function lerAba(ss, nomeAba) {
  const aba = ss.getSheetByName(nomeAba);
  if (!aba || aba.getLastRow() < 2) return [];

  const valores   = aba.getDataRange().getValues();
  const cabecalho = valores[0].map(c => String(c).trim());
  const resultado = [];

  for (let i = 1; i < valores.length; i++) {
    const obj = {};
    cabecalho.forEach((col, j) => {
      if (col) obj[col] = valores[i][j] !== undefined ? String(valores[i][j]) : '';
    });
    if (obj.nome) resultado.push(obj);
  }

  return resultado;
}

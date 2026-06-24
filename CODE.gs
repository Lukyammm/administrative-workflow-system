// Configuração inicial
function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Cosign - Gerenciamento de Armários');
}

function doPost(e) {
  return handlePost(e);
}

function callAction(action, parametros) {
  var acao = action !== null && action !== undefined ? action.toString().trim() : '';
  if (!acao) {
    return { success: false, error: 'Ação inválida.' };
  }

  var dados = {};
  if (parametros && typeof parametros === 'object') {
    for (var chave in parametros) {
      if (Object.prototype.hasOwnProperty.call(parametros, chave)) {
        var valor = parametros[chave];
        if (valor === undefined || valor === null) {
          dados[chave] = '';
        } else if (typeof valor === 'object') {
          dados[chave] = JSON.stringify(valor);
        } else {
          dados[chave] = valor.toString();
        }
      }
    }
  }
  dados.action = acao;

  var resposta = handlePost({ parameter: dados });
  if (resposta && typeof resposta.getContent === 'function') {
    try {
      return JSON.parse(resposta.getContent());
    } catch (erroParse) {
      return { success: false, error: 'Resposta inválida do servidor.' };
    }
  }
  return resposta;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function normalizarTextoBasico(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }
  var texto = valor.toString().trim().toLowerCase();
  if (typeof texto.normalize === 'function') {
    texto = texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }
  return texto;
}

function normalizarNumeroArmario(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }
  return valor.toString().trim();
}

function normalizarIdentificador(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }
  return valor.toString().trim();
}

function obterChaveNumeroArmario(numero) {
  var numeroNormalizado = normalizarNumeroArmario(numero);
  return numeroNormalizado ? numeroNormalizado : '__sem_numero__';
}

function montarChaveArmarioInterface(tipo, numero, idPlanilha) {
  var tipoNormalizado = normalizarTextoBasico(tipo) || 'geral';
  var numeroNormalizado = normalizarNumeroArmario(numero);
  if (numeroNormalizado) {
    return tipoNormalizado + ':' + numeroNormalizado;
  }
  var idTexto = '';
  if (idPlanilha !== null && idPlanilha !== undefined) {
    idTexto = idPlanilha.toString().trim();
  }
  return tipoNormalizado + ':id-' + (idTexto || Utilities.getUuid());
}

function obterEstruturaPlanilha(sheet) {
  var ultimaColuna = sheet.getLastColumn();
  var cabecalhos = ultimaColuna > 0 ? sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0] : [];
  var mapaIndices = {};

  cabecalhos.forEach(function(cabecalho, indice) {
    var chave = normalizarTextoBasico(cabecalho);
    if (chave && mapaIndices[chave] === undefined) {
      mapaIndices[chave] = indice;
    }
  });

  return {
    ultimaColuna: ultimaColuna,
    mapaIndices: mapaIndices
  };
}

function obterIndiceColuna(estrutura, chave, padrao) {
  if (Array.isArray(chave)) {
    for (var i = 0; i < chave.length; i++) {
      var indiceFlexivel = obterIndiceColuna(estrutura, chave[i], null);
      if (indiceFlexivel !== null && indiceFlexivel !== undefined) {
        return indiceFlexivel;
      }
    }
    return padrao;
  }

  if (chave === null || chave === undefined) {
    return padrao;
  }

  var chaveNormalizada = normalizarTextoBasico(chave);
  if (estrutura.mapaIndices.hasOwnProperty(chaveNormalizada)) {
    return estrutura.mapaIndices[chaveNormalizada];
  }
  return padrao;
}

function obterValorLinha(linha, estrutura, chave, padrao) {
  var indice = obterIndiceColuna(estrutura, chave, null);
  if (indice === null || indice === undefined) {
    return padrao;
  }
  if (indice >= linha.length) {
    return padrao;
  }
  return linha[indice];
}

function definirValorLinha(linha, estrutura, chave, valor) {
  var indice = obterIndiceColuna(estrutura, chave, null);
  if (indice === null || indice === undefined) {
    return;
  }
  while (linha.length < estrutura.ultimaColuna) {
    linha.push('');
  }
  linha[indice] = valor;
}

function obterValorLinhaFlexivel(linha, estrutura, chaves, padrao) {
  if (!Array.isArray(chaves)) {
    chaves = [chaves];
  }

  for (var i = 0; i < chaves.length; i++) {
    var indice = obterIndiceColuna(estrutura, chaves[i], null);
    if (indice !== null && indice !== undefined && indice < linha.length) {
      return linha[indice];
    }
  }

  return padrao;
}

function definirValorLinhaFlexivel(linha, estrutura, chaves, valor) {
  if (!Array.isArray(chaves)) {
    chaves = [chaves];
  }

  for (var i = 0; i < chaves.length; i++) {
    var indice = obterIndiceColuna(estrutura, chaves[i], null);
    if (indice === null || indice === undefined) {
      continue;
    }

    var tamanhoMinimo = Math.max(estrutura.ultimaColuna, indice + 1);
    while (linha.length < tamanhoMinimo) {
      linha.push('');
    }

    linha[indice] = valor;
    return true;
  }

  return false;
}

var CABECALHOS_WHATSAPP = ['whatsapp', 'wpp', 'whats app', 'whatsap', 'zap'];
var CABECALHOS_NOME_VISITANTE = ['nome visitante', 'visitante', 'nome do visitante'];
var CABECALHOS_NOME_ACOMPANHANTE = ['nome acompanhante', 'acompanhante', 'nome do acompanhante', 'responsavel', 'responsável'];
var CABECALHOS_VISITA_ESTENDIDA = ['visita estendida', 'visita extendida', 'visita expandida'];
var CABECALHOS_OBSERVACOES = ['observacoes', 'observações', 'observacao', 'observação', 'obs'];

function converterParaBoolean(valor) {
  if (valor === true || valor === false) {
    return valor;
  }
  if (typeof valor === 'number') {
    return valor !== 0;
  }
  if (typeof valor === 'string') {
    var texto = valor.trim().toLowerCase();
    return texto === 'true' || texto === '1' || texto === 'sim';
  }
  return false;
}

function ehNumeroContingencia(numero) {
  var texto = (numero || '').toString().trim().toLowerCase();
  return texto.indexOf('conting') === 0;
}

function extrairSequenciaContingencia(numero) {
  var texto = (numero || '').toString();
  var match = texto.match(/conting[êe]ncia[-\s]*(\d+)/i);
  if (match && match[1]) {
    var valor = parseInt(match[1], 10);
    return isNaN(valor) ? 0 : valor;
  }
  return 0;
}

function normalizarListaUnidadesParametro(valor) {
  try {
    if (valor === null || valor === undefined) {
      return [];
    }

    if (Array.isArray(valor)) {
      return valor.map(function(item) {
        return item !== null && item !== undefined ? item.toString().trim() : '';
      }).filter(function(item) {
        return item;
    });
    }

    if (typeof valor === 'string') {
      var texto = valor.trim();
      if (!texto) {
        return [];
      }

      if ((texto.charAt(0) === '[' && texto.charAt(texto.length - 1) === ']') ||
          (texto.charAt(0) === '{' && texto.charAt(texto.length - 1) === '}')) {
        try {
          var convertido = JSON.parse(texto);
          if (Array.isArray(convertido)) {
            return normalizarListaUnidadesParametro(convertido);
          }
        } catch (erroJSON) {
          console.error('Falha ao interpretar unidades como JSON:', erroJSON);
        }
      }

      if (texto.indexOf(';') !== -1 || texto.indexOf(',') !== -1) {
        return texto.split(/[;,]/).map(function(item) {
          return item.trim();
        }).filter(function(item) {
          return item;
        });
      }

      return [texto];
    }

    if (typeof valor === 'number' || typeof valor === 'boolean') {
      return [valor.toString()];
    }

    if (typeof valor === 'object') {
      var itens = [];
      for (var chave in valor) {
        if (!valor.hasOwnProperty(chave)) {
          continue;
        }
        var item = valor[chave];
        if (Array.isArray(item)) {
          itens = itens.concat(normalizarListaUnidadesParametro(item));
        } else if (item !== null && item !== undefined) {
          itens.push(item.toString().trim());
        }
      }
      return itens.filter(function(item) {
        return item;
      });
    }
  } catch (erro) {
    console.error('Erro ao normalizar unidades informadas:', erro);
  }

  return [];
}

function obterTimeZoneAplicacao() {
  var timezone = '';
  try {
    timezone = Session.getScriptTimeZone();
  } catch (erroTimezone) {
    console.error('Erro ao obter fuso horário do script:', erroTimezone);
  }

  if (!timezone) {
    timezone = 'America/Sao_Paulo';
  }

  return timezone;
}

function normalizarDataParaTimeZone(data, timezone) {
  if (!(data instanceof Date) || Number.isNaN(data.getTime())) {
    return null;
  }

  var tz = timezone || obterTimeZoneAplicacao();
  var dataTexto = Utilities.formatDate(data, tz, 'yyyy-MM-dd');
  var partes = dataTexto.split('-');
  if (partes.length !== 3) {
    return null;
  }

  var ano = Number(partes[0]);
  var mes = Number(partes[1]);
  var dia = Number(partes[2]);
  if (!Number.isFinite(ano) || !Number.isFinite(mes) || !Number.isFinite(dia)) {
    return null;
  }

  return new Date(ano, mes - 1, dia);
}

function interpretarDataParametroSeguro(valor, timezone) {
  if (valor === null || valor === undefined) {
    return null;
  }

  var texto = valor.toString().trim();
  if (!texto) {
    return null;
  }

  var tz = timezone || obterTimeZoneAplicacao();

  var matchIso = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (matchIso) {
    var anoIso = Number(matchIso[1]);
    var mesIso = Number(matchIso[2]);
    var diaIso = Number(matchIso[3]);
    var dataIso = new Date(anoIso, mesIso - 1, diaIso);
    return normalizarDataParaTimeZone(dataIso, tz);
  }

  var matchBR = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (matchBR) {
    var diaBr = Number(matchBR[1]);
    var mesBr = Number(matchBR[2]);
    var anoBr = Number(matchBR[3]);
    var dataBr = new Date(anoBr, mesBr - 1, diaBr);
    return normalizarDataParaTimeZone(dataBr, tz);
  }

  var tentativa = new Date(texto);
  if (!Number.isNaN(tentativa.getTime())) {
    return normalizarDataParaTimeZone(tentativa, tz);
  }

  return null;
}

function extrairDataValidaDaCelula(valor, exibicao, timezone) {
  var tz = timezone || obterTimeZoneAplicacao();

  if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
    return normalizarDataParaTimeZone(valor, tz);
  }

  if (typeof valor === 'number' && Number.isFinite(valor)) {
    var baseExcel = new Date(Date.UTC(1899, 11, 30));
    var dataSerial = new Date(baseExcel.getTime() + Math.round(valor * 24 * 60 * 60 * 1000));
    var normalizadaSerial = normalizarDataParaTimeZone(dataSerial, tz);
    if (normalizadaSerial) {
      return normalizadaSerial;
    }
  }

  var texto = '';
  if (typeof valor === 'string' && valor.trim()) {
    texto = valor.trim();
  } else if (exibicao !== null && exibicao !== undefined) {
    texto = exibicao.toString().trim();
  }

  if (!texto) {
    return null;
  }

  return interpretarDataParametroSeguro(texto, tz);
}

function obterDataAtualNormalizada(timezone) {
  var agora = new Date();
  return normalizarDataParaTimeZone(agora, timezone || obterTimeZoneAplicacao());
}

function gerarChaveDataComparacao(data, timezone) {
  var normalizada = normalizarDataParaTimeZone(data, timezone || obterTimeZoneAplicacao());
  if (!normalizada) {
    return '';
  }

  return Utilities.formatDate(normalizada, timezone || obterTimeZoneAplicacao(), 'yyyyMMdd');
}

function obterMapasUnidades() {
  var mapas = {
    porId: {},
    porNome: {}
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Unidades');

    if (!sheet) {
      return mapas;
    }

    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) {
      return mapas;
    }

    var dados = sheet.getRange(2, 1, ultimaLinha - 1, 2).getValues();
    dados.forEach(function(row) {
      var id = row[0];
      var nome = row[1];
      if (id === null || id === undefined) {
        return;
      }
      var idTexto = id.toString().trim();
      if (!idTexto) {
        return;
      }
      var nomeTexto = nome !== null && nome !== undefined ? nome.toString().trim() : '';
      mapas.porId[idTexto] = nomeTexto;
      if (nomeTexto) {
        var chaveNome = normalizarTextoBasico(nomeTexto);
        if (!mapas.porNome[chaveNome]) {
          mapas.porNome[chaveNome] = idTexto;
        }
      }
    });
  } catch (erroMapas) {
    console.error('Erro ao obter mapas de unidades:', erroMapas);
  }

  return mapas;
}

function formatarUnidadesParaRegistro(unidadesIds, mapas) {
  if (!Array.isArray(unidadesIds)) {
    return [];
  }

  var resultado = [];
  unidadesIds.forEach(function(unidadeId) {
    if (unidadeId === null || unidadeId === undefined) {
      return;
    }

    var chave = unidadeId.toString().trim();
    if (!chave) {
      return;
    }

    if (normalizarTextoBasico(chave) === 'all') {
      if (resultado.indexOf('Todas as unidades') === -1) {
        resultado.push('Todas as unidades');
      }
      return;
    }

    var nome = mapas && mapas.porId ? mapas.porId[chave] : '';
    if (nome) {
      if (resultado.indexOf(nome) === -1) {
        resultado.push(nome);
      }
    } else {
      if (resultado.indexOf(chave) === -1) {
        resultado.push(chave);
      }
    }
  });

  return resultado;
}

function resolverIdsUnidadesArmazenadas(unidadesValor, mapas) {
  var brutas = normalizarListaUnidadesParametro(unidadesValor);
  if (!brutas.length) {
    return [];
  }

  var ids = [];
  brutas.forEach(function(item) {
    if (item === null || item === undefined) {
      return;
    }

    var textoOriginal = item.toString().trim();
    if (!textoOriginal) {
      return;
    }

    var textoNormalizado = normalizarTextoBasico(textoOriginal);
    if (textoNormalizado === 'all' || textoNormalizado === 'todas as unidades') {
      if (ids.indexOf('all') === -1) {
        ids.push('all');
      }
      return;
    }

    if (mapas && mapas.porId && mapas.porId.hasOwnProperty(textoOriginal)) {
      if (ids.indexOf(textoOriginal) === -1) {
        ids.push(textoOriginal);
      }
      return;
    }

    var separadores = [' - ', '|', ':', '–', ' — '];
    for (var i = 0; i < separadores.length; i++) {
      var sep = separadores[i];
      if (textoOriginal.indexOf(sep) !== -1) {
        var candidato = textoOriginal.split(sep)[0].trim();
        if (candidato) {
          if (mapas && mapas.porId && mapas.porId.hasOwnProperty(candidato)) {
            if (ids.indexOf(candidato) === -1) {
              ids.push(candidato);
            }
            return;
          }
          if (mapas && mapas.porNome) {
            var candidatoNormalizado = normalizarTextoBasico(candidato);
            var idPorNome = mapas.porNome[candidatoNormalizado];
            if (idPorNome && ids.indexOf(idPorNome) === -1) {
              ids.push(idPorNome);
              return;
            }
          }
        }
      }
    }

    if (mapas && mapas.porNome) {
      var idPorNomeDireto = mapas.porNome[textoNormalizado];
      if (idPorNomeDireto && ids.indexOf(idPorNomeDireto) === -1) {
        ids.push(idPorNomeDireto);
        return;
      }
    }

    var matchNumero = textoOriginal.match(/\d+/);
    if (matchNumero && mapas && mapas.porId && mapas.porId.hasOwnProperty(matchNumero[0])) {
      var idNumero = matchNumero[0];
      if (ids.indexOf(idNumero) === -1) {
        ids.push(idNumero);
      }
      return;
    }

    if (ids.indexOf(textoOriginal) === -1) {
      ids.push(textoOriginal);
    }
  });

  return ids;
}

// ID da pasta do Drive para salvar os PDFs - ATUALIZE COM SEU ID
const PASTA_DRIVE_ID = '1nYsGJJUIufxDYVvIanVXCbPx7YuBOYDP';
const PASTA_DRIVE_TEMP_ID = '11M_fFDA9nrOqzIm6zMnaGjYMCQoE4XXs';
const PASTA_DRIVE_FOTOS_ID = '1XKZ6LApUh6RjBddeySCbPFdS6ncDS8DB';
const PASTA_BACKUP_RAIZ_ID = '1ouVHkmVmdJVe6aYVBzBUNYpndk4OKKmu';

// Configuração de cache para leitura dos termos
const TERMOS_CACHE_KEY = 'termos_registrados_cache_v1';
const TERMOS_CACHE_META_KEY = 'termos_registrados_cache_meta_v1';
const TERMOS_CACHE_CHUNK_PREFIX = 'termos_registrados_cache_chunk_v1';
const TERMOS_CACHE_CHUNK_TAMANHO_MAX = 75000; // ~75 KB para manter margem
const TERMOS_CACHE_TTL = 300; // segundos

// Configurações gerais de cache para otimizar leituras
const CACHE_PREFIXO = 'locknac_cache_v1';
const CACHE_TTL_PADRAO = 60; // segundos
const CACHE_TTL_ARMARIOS = 120;
const CACHE_TTL_HISTORICO = 90;
const CACHE_TTL_MOVIMENTACOES = 45;
const CACHE_TTL_INDICE_ARMARIOS = 30;
const LOCK_TIMEOUT_MS = 5000;
const RETRY_MAX_TENTATIVAS = 3;
const RETRY_INTERVALO_MS = 180;
const LIMIAR_LOG_LENTO_MS = 1200;

// Configuração da planilha de liberações externas
const PLANILHA_LIBERACAO_ID = '1UR6ynp6nxbpVMephgKkT8_YDc_ih5bYK565IebfojPI';
const PLANILHA_LIBERACAO_ABA = 'LIBERACAO';
const PLANILHA_LIBERACAO_LINHA_CABECALHO = 10;
const PLANILHA_LIBERACAO_COLUNA_DATA = 4; // Coluna D

// Achados e Perdidos
const PLANILHA_PERTENCES_PERDIDOS_ID = '1OhOr91MpGsRjJXWb1kdLXNI2NCBBvsjidAc4utfgRJ0';
const PLANILHA_PERTENCES_PERDIDOS_ABA = 'PERTENCES PERDIDOS GUARDA-VOLUMES';
const PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO = 3;
const PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL = 2; // Coluna B


// Base clínica externa (Pentaho)
const PLANILHA_ENTRADA_CLINICA_ID = '1ap8BnfjHTpF4KcwuyxGKJEUKddSZnkl3Ldfp3gDAmxg';
const PLANILHA_ENTRADA_CLINICA_ABA = 'PENTAHO_ENTRADA_CLINICA';

function obterAbaEntradaClinicaExterna() {
  var spreadsheet = SpreadsheetApp.openById(PLANILHA_ENTRADA_CLINICA_ID);
  return spreadsheet.getSheetByName(PLANILHA_ENTRADA_CLINICA_ABA);
}

function montarChaveCache() {
  var partes = Array.prototype.slice.call(arguments).filter(function(parte) {
    return parte !== null && parte !== undefined && parte !== '';
  }).map(function(parte) {
    if (typeof parte === 'object') {
      try {
        return JSON.stringify(parte);
      } catch (erro) {
        return '';
      }
    }
    return parte.toString().trim().toLowerCase().replace(/\s+/g, '-');
  });

  if (!partes.length) {
    return CACHE_PREFIXO;
  }

  return CACHE_PREFIXO + ':' + partes.join(':');
}

function executarComLock(chave, callback) {
  var lock = LockService.getScriptLock();
  var lockAdquirido = false;
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
    lockAdquirido = true;
    return callback();
  } finally {
    if (lockAdquirido) {
      try {
        lock.releaseLock();
      } catch (erroLock) {
        // Ignorado propositalmente
      }
    }
  }
}

function ehErroTransitorio(mensagem) {
  var texto = (mensagem || '').toString().toLowerCase();
  if (!texto) {
    return false;
  }
  return texto.indexOf('service invoked too many times') !== -1 ||
    texto.indexOf('rate limit') !== -1 ||
    texto.indexOf('temporarily unavailable') !== -1 ||
    texto.indexOf('internal error') !== -1 ||
    texto.indexOf('timed out') !== -1 ||
    texto.indexOf('exceeded maximum execution time') !== -1;
}

function executarComRetry(callback, opcoes) {
  var maxTentativas = opcoes && opcoes.maxTentativas ? opcoes.maxTentativas : RETRY_MAX_TENTATIVAS;
  var intervaloInicial = opcoes && opcoes.intervaloMs ? opcoes.intervaloMs : RETRY_INTERVALO_MS;
  var ultimaFalha = null;

  for (var tentativa = 1; tentativa <= maxTentativas; tentativa++) {
    try {
      return callback(tentativa);
    } catch (erro) {
      ultimaFalha = erro;
      var mensagem = erro && erro.message ? erro.message : erro;
      var podeTentarNovamente = tentativa < maxTentativas && ehErroTransitorio(mensagem);
      if (!podeTentarNovamente) {
        throw erro;
      }
      Utilities.sleep(Math.min(intervaloInicial * Math.pow(2, tentativa - 1), 1200));
    }
  }

  throw ultimaFalha || new Error('Falha após tentativas de retry.');
}

function executarComCache(chave, ttl, fornecedor) {
  if (!chave) {
    return fornecedor();
  }

  var cache = CacheService.getScriptCache();

  try {
    var armazenado = cache.get(chave);
    if (armazenado) {
      return JSON.parse(armazenado);
    }
  } catch (erroLeitura) {
    try {
      cache.remove(chave);
    } catch (erroRemocao) {
      // Ignorado propositalmente
    }
  }

  var resultado = fornecedor();

  if (resultado && resultado.success) {
    try {
      cache.put(chave, JSON.stringify(resultado), ttl || CACHE_TTL_PADRAO);
    } catch (erroGravacao) {
      // Falhas de cache não devem interromper o fluxo principal
    }
  }

  return resultado;
}

function limparCaches(chaves) {
  if (!chaves) {
    return;
  }

  var lista = Array.isArray(chaves) ? chaves : [chaves];
  if (!lista.length) {
    return;
  }

  var cache = CacheService.getScriptCache();
  lista.forEach(function(chave) {
    if (!chave) {
      return;
    }
    try {
      cache.remove(chave);
    } catch (erroRemocao) {
      // Ignorado propositalmente
    }
  });
}

function limparCacheArmarios() {
  var tipos = ['visitante', 'acompanhante', 'geral'];
  var chaves = [];
  tipos.forEach(function(tipo) {
    chaves.push(montarChaveCache('armarios', tipo));
    chaves.push(montarChaveCache('armarios', tipo, 'com-internacoes'));
    chaves.push(montarChaveCache('armarios', tipo, 'sem-internacoes'));
  });
  limparCaches(chaves);
}

function limparCacheUsuarios() {
  limparCaches(montarChaveCache('usuarios'));
}

function limparCacheHistorico() {
  limparCaches([
    montarChaveCache('historico', 'visitante'),
    montarChaveCache('historico', 'acompanhante')
  ]);
}

function limparCacheCadastroArmarios() {
  limparCaches(montarChaveCache('cadastro-armarios'));
}

function limparCacheUnidades() {
  limparCaches(montarChaveCache('unidades'));
}

function limparCacheMovimentacoes(armarioId, numeroArmario, tipo) {
  var idTexto = armarioId !== undefined && armarioId !== null ? armarioId.toString().trim() : '';
  var numeroTexto = normalizarNumeroArmario(numeroArmario);
  var tipoTexto = tipo ? normalizarTextoBasico(tipo) : '';
  var statusVariantes = ['ativos', 'finalizados'];
  var chaves = [];

  if (idTexto) {
    chaves.push(montarChaveCache('movimentacoes', [idTexto, numeroTexto, tipoTexto].join('|')));
    statusVariantes.forEach(function(status) {
      chaves.push(montarChaveCache('movimentacoes', [idTexto, numeroTexto, tipoTexto, status].join('|')));
    });
    chaves.push(montarChaveCache('movimentacoes', idTexto));
  }

  if (numeroTexto) {
    chaves.push(montarChaveCache('movimentacoes', ['numero', numeroTexto, tipoTexto].join('|')));
    statusVariantes.forEach(function(status) {
      chaves.push(montarChaveCache('movimentacoes', ['numero', numeroTexto, tipoTexto, status].join('|')));
    });
  }

  chaves.push(montarChaveCache('movimentacoes', 'todos'));
  chaves.push(montarChaveCache('movimentacoes', 'todos_finalizados'));
  limparCaches(chaves);
}

function limparCacheIndiceArmarios(sheetName) {
  var nomesAbas = sheetName ? [sheetName] : ['Visitantes', 'Acompanhantes'];
  var chaves = nomesAbas.map(function(nome) {
    return montarChaveCache('indice-armarios', nome);
  });
  limparCaches(chaves);
}

function invalidarCachesArmariosRelacionados(sheetName) {
  limparCacheArmarios();
  limparCacheHistorico();
  limparCacheIndiceArmarios(sheetName);
}

function validarLinhaArmarioEncontrada(armarioData, estrutura, idComparacao, numeroInformado) {
  if (!armarioData || !estrutura) {
    return false;
  }

  var idLinha = obterValorLinha(armarioData, estrutura, 'id', '');
  var numeroLinha = obterValorLinha(armarioData, estrutura, 'numero', '');
  var numeroLinhaNormalizado = normalizarNumeroArmario(numeroLinha);

  if (idComparacao && idLinha !== null && idLinha !== undefined && idLinha.toString().trim() === idComparacao) {
    return true;
  }

  if (numeroInformado && numeroLinhaNormalizado === numeroInformado) {
    return true;
  }

  if (!idComparacao && !numeroInformado) {
    return false;
  }

  return false;
}

function obterIndiceArmarios(sheetName, forcarReconstrucao) {
  var cache = CacheService.getScriptCache();
  var chaveCache = montarChaveCache('indice-armarios', sheetName);

  if (forcarReconstrucao) {
    try {
      cache.remove(chaveCache);
    } catch (erroRemocao) {
      // ignorado
    }
  } else {
    try {
      var armazenado = cache.get(chaveCache);
      if (armazenado) {
        return JSON.parse(armazenado);
      }
    } catch (erroCache) {
      try {
        cache.remove(chaveCache);
      } catch (erroRemocaoCache) {
        // ignorado
      }
    }
  }

  var resultado = { porId: {}, porNumero: {} };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return resultado;
  }

  var estrutura = obterEstruturaPlanilha(sheet);
  var idIndex = obterIndiceColuna(estrutura, 'id', 0);
  var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
  var totalLinhas = sheet.getLastRow();

  if (totalLinhas <= 1) {
    return resultado;
  }

  if (idIndex === null || idIndex === undefined || numeroIndex === null || numeroIndex === undefined) {
    return resultado;
  }

  var ids = sheet.getRange(2, idIndex + 1, totalLinhas - 1, 1).getValues();
  var numeros = sheet.getRange(2, numeroIndex + 1, totalLinhas - 1, 1).getValues();

  for (var i = 0; i < totalLinhas - 1; i++) {
    var linhaPlanilha = i + 2;
    var idValor = ids[i] && ids[i][0] !== null && ids[i][0] !== undefined ? ids[i][0].toString().trim() : '';
    var numeroValor = normalizarNumeroArmario(numeros[i] && numeros[i][0] !== undefined ? numeros[i][0] : '');

    if (idValor) {
      resultado.porId[idValor] = linhaPlanilha;
    }

    if (numeroValor) {
      var chaveNumero = obterChaveNumeroArmario(numeroValor);
      resultado.porNumero[chaveNumero] = linhaPlanilha;
    }
  }

  try {
    cache.put(chaveCache, JSON.stringify(resultado), CACHE_TTL_INDICE_ARMARIOS);
  } catch (erroGravacao) {
    // ignorado
  }

  return resultado;
}

// Inicializar planilha com todas as abas e cabeçalhos
function inicializarPlanilha() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Criar abas se não existirem
    var abas = [
      {
        nome: 'Histórico Visitantes',
        cabecalhos: ['ID', 'Data', 'Número Armário', 'Nome Visitante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Fim', 'Status', 'Tipo', 'Unidade', 'WhatsApp']
      },
      {
        nome: 'Histórico Acompanhantes',
        cabecalhos: ['ID', 'Data', 'Número Armário', 'Nome Acompanhante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Fim', 'Status', 'Tipo', 'Unidade', 'WhatsApp']
      },
      {
        nome: 'Visitantes',
        cabecalhos: ['ID', 'Número', 'Status', 'Nome Visitante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Prevista', 'Data Registro', 'Unidade', 'Termo Aplicado', 'WhatsApp']
      },
      {
        nome: 'Acompanhantes',
        cabecalhos: ['ID', 'Número', 'Status', 'Nome Acompanhante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Data Registro', 'WhatsApp', 'Unidade', 'Termo Aplicado']
      },
      { 
        nome: 'Cadastro Armários', 
        cabecalhos: ['ID', 'Número', 'Tipo', 'Unidade', 'Localização', 'Status', 'Data Cadastro'] 
      },
      { 
        nome: 'Unidades', 
        cabecalhos: ['ID', 'Nome', 'Status', 'Data Cadastro'] 
      },
      {
        nome: 'Usuários',
        cabecalhos: ['ID', 'Nome', 'Email', 'Perfil', 'Acesso Visitantes', 'Acesso Acompanhantes', 'Data Cadastro', 'Status', 'Senha', 'Unidades', 'Precisa Trocar Senha']
      },
      {
        nome: NOME_ABA_SESSOES,
        cabecalhos: ['TokenHash', 'UsuarioId', 'Perfil', 'PrecisaTrocarSenha', 'CriadoEm', 'ExpiraEm']
      },
      { 
        nome: 'LOGS', 
        cabecalhos: ['Data/Hora', 'Usuário', 'Ação', 'Detalhes', 'IP'] 
      },
      { 
        nome: 'Termos de Responsabilidade', 
        cabecalhos: ['ID', 'ArmarioID', 'NumeroArmario', 'Paciente', 'Prontuario', 'Nascimento', 'Setor', 'Leito', 'Consciente', 'Acompanhante', 'Telefone', 'Documento', 'Parentesco', 'Orientacoes', 'Volumes', 'DescricaoVolumes', 'AplicadoEm', 'PDF_URL', 'AssinaturaBase64'] 
      },
      {
        nome: 'Movimentações',
        cabecalhos: ['ID', 'ArmarioID', 'NumeroArmario', 'Tipo', 'Descricao', 'Responsavel', 'Data', 'Hora', 'DataHoraRegistro', 'Status']
      }
    ];
    
    abas.forEach(function(aba) {
      var sheet = ss.getSheetByName(aba.nome);
      if (!sheet) {
        sheet = ss.insertSheet(aba.nome);
        sheet.getRange(1, 1, 1, aba.cabecalhos.length).setValues([aba.cabecalhos]);
        sheet.setFrozenRows(1);
        
        // Formatar cabeçalhos
        var headerRange = sheet.getRange(1, 1, 1, aba.cabecalhos.length);
        headerRange.setBackground('#2c6e8f')
          .setFontColor('white')
          .setFontWeight('bold');
      }
    });
    
    // Migração para instalações já existentes: garante a coluna de troca de senha
    // (a criação acima só escreve cabeçalhos em abas novas).
    garantirColunaPrecisaTrocarSenha();

    // Adicionar alguns dados iniciais de exemplo
    adicionarDadosIniciais();

    registrarLog('SISTEMA', 'Planilha inicializada com sucesso');

    limparCacheArmarios();
    limparCacheUsuarios();
    limparCacheHistorico();
    limparCacheCadastroArmarios();
    limparCacheUnidades();
    limparCacheMovimentacoes();

    return { success: true, message: 'Planilha inicializada com sucesso' };

  } catch (error) {
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Garante a coluna "Precisa Trocar Senha" em instalações já existentes e força a
// rotação da senha padrão (admin123) caso ela ainda esteja em uso.
function garantirColunaPrecisaTrocarSenha() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastColumn() < 1) {
      return;
    }
    var ultimaColuna = sheet.getLastColumn();
    var cabecalhos = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
    var existe = cabecalhos.some(function(cabecalho) {
      return normalizarTextoBasico(cabecalho) === 'precisa trocar senha';
    });
    if (!existe) {
      sheet.getRange(1, ultimaColuna + 1).setValue('Precisa Trocar Senha');
    }

    if (sheet.getLastRow() < 2) {
      return;
    }
    // Marca para troca obrigatória qualquer usuário que ainda use a senha padrão.
    var estrutura = obterEstruturaPlanilha(sheet);
    var indiceSenha = obterIndiceColuna(estrutura, 'senha', null);
    var indiceFlag = obterIndiceColuna(estrutura, ['precisa trocar senha', 'precisatrocarsenha'], null);
    if (indiceSenha === null || indiceSenha === undefined || indiceFlag === null || indiceFlag === undefined) {
      return;
    }
    var totalColunas = estrutura.ultimaColuna || sheet.getLastColumn();
    var valores = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
    for (var i = 0; i < valores.length; i++) {
      var senhaArmazenada = (valores[i][indiceSenha] || '').toString().trim();
      if (senhaArmazenada && validarSenha('admin123', senhaArmazenada)) {
        sheet.getRange(i + 2, indiceFlag + 1).setValue('true');
      }
    }
  } catch (erro) {
    // Migração best-effort; não deve interromper a inicialização.
  }
}

function obterFusoHorarioPadrao() {
  var timezone = '';
  try {
    timezone = Session.getScriptTimeZone();
  } catch (erro) {
    timezone = '';
  }
  return timezone || 'America/Sao_Paulo';
}

function formatarDataPlanilha(valor) {
  if (!valor) {
    return '';
  }
  var timezone = obterFusoHorarioPadrao();
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, timezone, 'dd/MM/yyyy');
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    if (!texto) {
      return '';
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
      var partesData = texto.split('-');
      var ano = parseInt(partesData[0], 10);
      var mes = parseInt(partesData[1], 10) - 1;
      var dia = parseInt(partesData[2], 10);
      if (!isNaN(ano) && !isNaN(mes) && !isNaN(dia)) {
        var dataLocal = new Date(ano, mes, dia);
        return Utilities.formatDate(dataLocal, timezone, 'dd/MM/yyyy');
      }
    }
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
      return texto;
    }
    var textoISO = texto.replace(' ', 'T');
    var data = new Date(textoISO);
    if (!isNaN(data.getTime())) {
      return Utilities.formatDate(data, timezone, 'dd/MM/yyyy');
    }
  }
  return valor;
}

function formatarHorarioPlanilha(valor) {
  if (!valor) {
    return '';
  }
  var timezone = obterFusoHorarioPadrao();
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    var formato = valor.getFullYear() <= 1900 ? 'HH:mm' : 'dd/MM/yyyy HH:mm';
    return Utilities.formatDate(valor, timezone, formato);
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    if (!texto) {
      return '';
    }
    if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(texto)) {
      return texto.slice(0, 5);
    }
    var isoSemFuso = texto.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(?::\d{2})?$/);
    if (isoSemFuso) {
      return isoSemFuso[2];
    }
    var textoISO = texto.replace(' ', 'T');
    var data = new Date(textoISO);
    if (!isNaN(data.getTime())) {
      return Utilities.formatDate(data, timezone, 'HH:mm');
    }
  }
  return valor;
}

function obterDataSomenteDia(valor) {
  if (!valor) {
    return null;
  }
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  }
  if (typeof valor === 'number') {
    var dataNumero = new Date(valor);
    if (!isNaN(dataNumero.getTime())) {
      return new Date(dataNumero.getFullYear(), dataNumero.getMonth(), dataNumero.getDate());
    }
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    if (!texto) {
      return null;
    }
    var matchDMY = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (matchDMY) {
      var dia = parseInt(matchDMY[1], 10);
      var mes = parseInt(matchDMY[2], 10) - 1;
      var ano = parseInt(matchDMY[3], 10);
      if (!isNaN(dia) && !isNaN(mes) && !isNaN(ano)) {
        return new Date(ano, mes, dia);
      }
    }
    var matchISO = texto.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (matchISO) {
      var anoIso = parseInt(matchISO[1], 10);
      var mesIso = parseInt(matchISO[2], 10) - 1;
      var diaIso = parseInt(matchISO[3], 10);
      if (!isNaN(anoIso) && !isNaN(mesIso) && !isNaN(diaIso)) {
        return new Date(anoIso, mesIso, diaIso);
      }
    }
    var textoISO = texto.replace(' ', 'T');
    var dataISO = new Date(textoISO);
    if (!isNaN(dataISO.getTime())) {
      return new Date(dataISO.getFullYear(), dataISO.getMonth(), dataISO.getDate());
    }
  }
  return null;
}

function extrairHorarioComponentes(valor) {
  if (!valor) {
    return null;
  }
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return {
      horas: valor.getHours(),
      minutos: valor.getMinutes(),
      dataBase: valor.getFullYear() > 1900 ? new Date(valor.getFullYear(), valor.getMonth(), valor.getDate()) : null
    };
  }
  if (typeof valor === 'number') {
    var dataNumero = new Date(valor);
    if (!isNaN(dataNumero.getTime())) {
      return {
        horas: dataNumero.getHours(),
        minutos: dataNumero.getMinutes(),
        dataBase: dataNumero.getFullYear() > 1900 ? new Date(dataNumero.getFullYear(), dataNumero.getMonth(), dataNumero.getDate()) : null
      };
    }
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    if (!texto) {
      return null;
    }
    var matchSimples = texto.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (matchSimples) {
      var horasSimples = parseInt(matchSimples[1], 10);
      var minutosSimples = parseInt(matchSimples[2], 10);
      if (!isNaN(horasSimples) && !isNaN(minutosSimples) && horasSimples >= 0 && horasSimples <= 23 && minutosSimples >= 0 && minutosSimples <= 59) {
        return { horas: horasSimples, minutos: minutosSimples, dataBase: null };
      }
    }
    var matchDMYHora = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
    if (matchDMYHora) {
      var diaData = parseInt(matchDMYHora[1], 10);
      var mesData = parseInt(matchDMYHora[2], 10) - 1;
      var anoData = parseInt(matchDMYHora[3], 10);
      var horaData = parseInt(matchDMYHora[4], 10);
      var minutoData = parseInt(matchDMYHora[5], 10);
      if ([diaData, mesData, anoData, horaData, minutoData].every(function(v) { return !isNaN(v); })) {
        return {
          horas: horaData,
          minutos: minutoData,
          dataBase: new Date(anoData, mesData, diaData)
        };
      }
    }
    var matchISOHora = texto.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
    if (matchISOHora) {
      var anoHora = parseInt(matchISOHora[1], 10);
      var mesHora = parseInt(matchISOHora[2], 10) - 1;
      var diaHora = parseInt(matchISOHora[3], 10);
      var horaIso = parseInt(matchISOHora[4], 10);
      var minutoIso = parseInt(matchISOHora[5], 10);
      if ([anoHora, mesHora, diaHora, horaIso, minutoIso].every(function(v) { return !isNaN(v); })) {
        return {
          horas: horaIso,
          minutos: minutoIso,
          dataBase: new Date(anoHora, mesHora, diaHora)
        };
      }
    }
    var textoISO = texto.replace(' ', 'T');
    var dataISO = new Date(textoISO);
    if (!isNaN(dataISO.getTime())) {
      return {
        horas: dataISO.getHours(),
        minutos: dataISO.getMinutes(),
        dataBase: dataISO.getFullYear() > 1900 ? new Date(dataISO.getFullYear(), dataISO.getMonth(), dataISO.getDate()) : null
      };
    }
  }
  return null;
}

function montarDataHoraComDefaults(dataValor, horaValor) {
  var componentes = extrairHorarioComponentes(horaValor);
  if (!componentes) {
    return null;
  }
  var dataBase = componentes.dataBase || obterDataSomenteDia(dataValor) || obterDataSomenteDia(new Date());
  if (!dataBase) {
    return null;
  }
  return new Date(
    dataBase.getFullYear(),
    dataBase.getMonth(),
    dataBase.getDate(),
    componentes.horas,
    componentes.minutos,
    0,
    0
  );
}

function calcularStatusAutomaticoVisitante(statusAtual, dataRegistroValor, horaPrevistaValor, agoraReferencia) {
  var statusNormalizado = normalizarTextoBasico(statusAtual);
  if (statusNormalizado === 'em uso') {
    statusNormalizado = 'em-uso';
  }
  if (['em-uso', 'proximo', 'vencido'].indexOf(statusNormalizado) === -1) {
    return statusAtual;
  }
  var dataHoraPrevista = montarDataHoraComDefaults(dataRegistroValor, horaPrevistaValor);
  if (!dataHoraPrevista) {
    return statusNormalizado || statusAtual;
  }
  var agora = (agoraReferencia instanceof Date && !isNaN(agoraReferencia.getTime())) ? agoraReferencia : new Date();
  var diferencaMinutos = (dataHoraPrevista.getTime() - agora.getTime()) / (1000 * 60);
  if (diferencaMinutos < 0) {
    return 'vencido';
  }
  if (diferencaMinutos <= 10) {
    return 'proximo';
  }
  return 'em-uso';
}

function determinarResponsavelRegistro(valorPreferencial) {
  if (valorPreferencial !== undefined && valorPreferencial !== null) {
    var texto = valorPreferencial.toString().trim();
    if (texto) {
      return texto;
    }
  }

  if (usuarioContextoRequisicao) {
    return usuarioContextoRequisicao;
  }

  try {
    var usuarioAtivo = Session.getActiveUser();
    if (usuarioAtivo && typeof usuarioAtivo.getEmail === 'function') {
      var emailAtivo = usuarioAtivo.getEmail();
      if (emailAtivo) {
        return emailAtivo;
      }
    }
  } catch (erroUsuarioAtivo) {
    // Ignora erro ao obter usuário ativo
  }

  try {
    var usuarioEfetivo = Session.getEffectiveUser();
    if (usuarioEfetivo && typeof usuarioEfetivo.getEmail === 'function') {
      var emailEfetivo = usuarioEfetivo.getEmail();
      if (emailEfetivo) {
        return emailEfetivo;
      }
    }
  } catch (erroUsuarioEfetivo) {
    // Ignora erro ao obter usuário efetivo
  }

  return '';
}

function obterDataValida(valor) {
  if (!valor) {
    return null;
  }
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return new Date(valor.getTime());
  }
  if (typeof valor === 'number' && isFinite(valor)) {
    var dataNumero = new Date(valor);
    return isNaN(dataNumero.getTime()) ? null : dataNumero;
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    if (!texto) {
      return null;
    }
    var dataPorBarras = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (dataPorBarras) {
      var dia = parseInt(dataPorBarras[1], 10);
      var mes = parseInt(dataPorBarras[2], 10) - 1;
      var ano = parseInt(dataPorBarras[3], 10);
      var dataLocal = new Date(ano, mes, dia);
      return isNaN(dataLocal.getTime()) ? null : dataLocal;
    }
    var textoISO = texto.replace(' ', 'T');
    var dataISO = new Date(textoISO);
    if (!isNaN(dataISO.getTime())) {
      return dataISO;
    }
  }
  return null;
}

function formatarDataDDMMAAAA(valor) {
  var dataValida = obterDataValida(valor);
  if (!dataValida) {
    return '';
  }
  try {
    return Utilities.formatDate(dataValida, obterFusoHorarioPadrao(), 'dd/MM/yyyy');
  } catch (erroFormato) {
    return '';
  }
}

function adicionarDias(data, dias) {
  if (!data || isNaN(data.getTime())) {
    return null;
  }
  var novaData = new Date(data.getTime());
  novaData.setDate(novaData.getDate() + dias);
  return novaData;
}

function obterDataHoraAtualFormatada() {
  var agora = new Date();
  var timezone = obterFusoHorarioPadrao();
  return {
    data: agora,
    horaCurta: Utilities.formatDate(agora, timezone, 'HH:mm'),
    dataHoraIso: Utilities.formatDate(agora, timezone, "yyyy-MM-dd'T'HH:mm:ss")
  };
}

function converterParaDataHoraIso(valor, padrao) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, obterFusoHorarioPadrao(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  if (valor && typeof valor === 'string') {
    return valor;
  }
  return padrao !== undefined ? padrao : '';
}

var MAX_TENTATIVAS_LOGIN = 5;
var BLOQUEIO_LOGIN_MINUTOS = 10;

function gerarSaltSenha() {
  return Utilities.getUuid();
}

function bytesParaHex(bytes) {
  return bytes.map(function(byte) {
    var valor = byte < 0 ? byte + 256 : byte;
    return ('0' + valor.toString(16)).slice(-2);
  }).join('');
}

function calcularHashSenha(senha, salt) {
  var material = (salt || '') + senha;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, material, Utilities.Charset.UTF_8);
  return bytesParaHex(digest);
}

function criarHashSenha(senha) {
  var salt = gerarSaltSenha();
  var hash = calcularHashSenha(senha, salt);
  return salt + ':' + hash;
}

function senhaEhHashValido(valor) {
  return valor && valor.toString().indexOf(':') > -1;
}

function validarSenha(senhaInformada, senhaArmazenada) {
  if (!senhaArmazenada) {
    return false;
  }
  var texto = senhaArmazenada.toString().trim();
  if (!texto) {
    return false;
  }
  if (!senhaEhHashValido(texto)) {
    return senhaInformada === texto;
  }
  var partes = texto.split(':');
  if (partes.length < 2) {
    return false;
  }
  var salt = partes[0];
  var hash = partes.slice(1).join(':');
  return calcularHashSenha(senhaInformada, salt) === hash;
}

function obterCacheLogin() {
  return CacheService.getScriptCache();
}

function obterChaveTentativasLogin(login) {
  return 'login_tentativas:' + normalizarTextoBasico(login);
}

function obterChaveBloqueioLogin(login) {
  return 'login_bloqueio:' + normalizarTextoBasico(login);
}

function verificarLoginBloqueado(login) {
  var cache = obterCacheLogin();
  return cache.get(obterChaveBloqueioLogin(login)) === '1';
}

function registrarFalhaLogin(login) {
  var cache = obterCacheLogin();
  var chaveTentativas = obterChaveTentativasLogin(login);
  var tentativas = parseInt(cache.get(chaveTentativas) || '0', 10);
  tentativas += 1;
  cache.put(chaveTentativas, tentativas.toString(), BLOQUEIO_LOGIN_MINUTOS * 60);
  if (tentativas >= MAX_TENTATIVAS_LOGIN) {
    cache.put(obterChaveBloqueioLogin(login), '1', BLOQUEIO_LOGIN_MINUTOS * 60);
  }
  return tentativas;
}

function limparTentativasLogin(login) {
  var cache = obterCacheLogin();
  cache.remove(obterChaveTentativasLogin(login));
  cache.remove(obterChaveBloqueioLogin(login));
}

// ===== Sessões com token (autenticação server-side) =====
var NOME_ABA_SESSOES = 'Sessões';
var SESSAO_TTL_MS = 8 * 60 * 60 * 1000; // 8 horas

function calcularHashToken(token) {
  // Token é um UUID de alta entropia; SHA-256 simples é suficiente para indexar.
  return calcularHashSenha(token, 'sessao');
}

function obterAbaSessoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return null;
  }
  var sheet = ss.getSheetByName(NOME_ABA_SESSOES);
  if (!sheet) {
    sheet = ss.insertSheet(NOME_ABA_SESSOES);
    sheet.getRange(1, 1, 1, 6).setValues([[
      'TokenHash', 'UsuarioId', 'Perfil', 'PrecisaTrocarSenha', 'CriadoEm', 'ExpiraEm'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function criarSessao(usuarioId, perfil, precisaTrocarSenha) {
  var sheet = obterAbaSessoes();
  if (!sheet) {
    return '';
  }
  limparSessoesExpiradas(sheet);
  var token = Utilities.getUuid() + Utilities.getUuid();
  var agora = Date.now();
  sheet.appendRow([
    calcularHashToken(token),
    parseInt(usuarioId, 10) || '',
    (perfil || '').toString(),
    precisaTrocarSenha ? 'true' : 'false',
    new Date(agora).toISOString(),
    new Date(agora + SESSAO_TTL_MS).toISOString()
  ]);
  return token;
}

function resolverSessao(token) {
  if (!token) {
    return null;
  }
  var sheet = obterAbaSessoes();
  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }
  var tokenHash = calcularHashToken(token.toString());
  var valores = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var agora = Date.now();
  for (var i = 0; i < valores.length; i++) {
    if (valores[i][0] === tokenHash) {
      var expiraEm = new Date(valores[i][5]).getTime();
      if (isNaN(expiraEm) || expiraEm <= agora) {
        // Expirada: remove e nega.
        try { sheet.deleteRow(i + 2); } catch (erroDel) {}
        return null;
      }
      // Renova o TTL deslizante apenas quando passou da metade, evitando uma
      // escrita na planilha a cada requisição.
      if ((expiraEm - agora) < (SESSAO_TTL_MS / 2)) {
        var novoExpira = new Date(agora + SESSAO_TTL_MS).toISOString();
        try { sheet.getRange(i + 2, 6).setValue(novoExpira); } catch (erroRenova) {}
      }
      return {
        usuarioId: parseInt(valores[i][1], 10) || null,
        perfil: (valores[i][2] || '').toString(),
        precisaTrocarSenha: normalizarTextoBasico(valores[i][3]) === 'true',
        linha: i + 2
      };
    }
  }
  return null;
}

function encerrarSessao(token) {
  try {
    if (!token) {
      return { success: true };
    }
    var sheet = obterAbaSessoes();
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true };
    }
    var tokenHash = calcularHashToken(token.toString());
    var valores = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = valores.length - 1; i >= 0; i--) {
      if (valores[i][0] === tokenHash) {
        sheet.deleteRow(i + 2);
      }
    }
    return { success: true };
  } catch (error) {
    return { success: true };
  }
}

function atualizarFlagSenhaSessoes(usuarioId, precisaTrocarSenha) {
  try {
    var sheet = obterAbaSessoes();
    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }
    var idAlvo = parseInt(usuarioId, 10);
    var valores = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < valores.length; i++) {
      if ((parseInt(valores[i][1], 10) || null) === idAlvo) {
        sheet.getRange(i + 2, 4).setValue(precisaTrocarSenha ? 'true' : 'false');
      }
    }
  } catch (erro) {
    // Ignora falha de housekeeping.
  }
}

function limparSessoesExpiradas(sheet) {
  try {
    sheet = sheet || obterAbaSessoes();
    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }
    var valores = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    var agora = Date.now();
    for (var i = valores.length - 1; i >= 0; i--) {
      var expiraEm = new Date(valores[i][5]).getTime();
      if (isNaN(expiraEm) || expiraEm <= agora) {
        try { sheet.deleteRow(i + 2); } catch (erroDel) {}
      }
    }
  } catch (erro) {
    // Ignora falha de limpeza.
  }
}

function obterUsuarioPorId(id, estrutura, valores) {
  if (!id) {
    return null;
  }
  var idIndex = obterIndiceColuna(estrutura, 'id', 0);
  for (var i = 0; i < valores.length; i++) {
    if (parseInt(valores[i][idIndex], 10) === id) {
      return { linha: valores[i], indice: i };
    }
  }
  return null;
}

function validarPermissaoAdmin(parametros) {
  // Confia exclusivamente no id resolvido pela sessão (server-side).
  // O usuarioId enviado pelo cliente é ignorado para evitar falsificação.
  var id = parseInt(sessaoUsuarioIdAtual, 10);
  if (!id) {
    return { ok: false, error: 'Acesso negado' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Usuários');
  if (!sheet || sheet.getLastRow() < 2) {
    return { ok: false, error: 'Acesso negado' };
  }
  var estrutura = obterEstruturaPlanilha(sheet);
  var totalColunas = estrutura.ultimaColuna || 10;
  var valores = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
  var usuarioEncontrado = obterUsuarioPorId(id, estrutura, valores);
  if (!usuarioEncontrado) {
    return { ok: false, error: 'Acesso negado' };
  }
  var perfil = obterValorLinha(usuarioEncontrado.linha, estrutura, 'perfil', '');
  var status = obterValorLinha(usuarioEncontrado.linha, estrutura, 'status', '');
  if (normalizarTextoBasico(status) !== 'ativo') {
    return { ok: false, error: 'Acesso negado' };
  }
  if (normalizarTextoBasico(perfil) !== 'admin') {
    return { ok: false, error: 'Acesso negado' };
  }
  return { ok: true };
}

// Adicionar dados iniciais de exemplo
function adicionarDadosIniciais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cadastrar alguns armários físicos
  var cadastroSheet = ss.getSheetByName('Cadastro Armários');
  if (cadastroSheet.getLastRow() === 1) {
    var dataCadastroArmarios = obterDataHoraAtualFormatada().dataHoraIso;
    var armariosIniciais = [
      ['V-01', 'visitante', 'NAC Eletiva', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['V-02', 'visitante', 'NAC Eletiva', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['V-03', 'visitante', 'UIB', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['V-04', 'visitante', 'UIB', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['A-01', 'acompanhante', 'NAC Eletiva', 'Bloco B - 1º Andar', 'ativo', dataCadastroArmarios],
      ['A-02', 'acompanhante', 'UIB', 'Bloco B - 1º Andar', 'ativo', dataCadastroArmarios],
      ['A-03', 'acompanhante', 'UIB', 'Bloco B - 1º Andar', 'ativo', dataCadastroArmarios]
    ];

    armariosIniciais.forEach(function(armario, index) {
      cadastroSheet.getRange(cadastroSheet.getLastRow() + 1, 1, 1, 7)
        .setValues([[index + 1, ...armario]]);
    });

    criarArmariosUso(armariosIniciais.map((armario, index) => [index + 1, ...armario]));
  }

  // Cadastrar usuário admin inicial
  var usuariosSheet = ss.getSheetByName('Usuários');
  if (usuariosSheet.getLastRow() === 1) {
    var dataCadastroUsuario = obterDataHoraAtualFormatada().dataHoraIso;
    var senhaAdmin = criarHashSenha('admin123');
    // Senha inicial provisória: o admin é obrigado a trocá-la no primeiro acesso.
    usuariosSheet.getRange(2, 1, 1, 11)
      .setValues([[1, 'Administrador', 'admin', 'admin', true, true, dataCadastroUsuario, 'ativo', senhaAdmin, 'all', 'true']]);
  }

  // Cadastrar unidades iniciais
  var unidadesSheet = ss.getSheetByName('Unidades');
  if (unidadesSheet && unidadesSheet.getLastRow() === 1) {
    var dataCadastroUnidades = obterDataHoraAtualFormatada().dataHoraIso;
    var unidadesIniciais = [
      [1, 'NAC Eletiva', 'ativa', dataCadastroUnidades],
      [2, 'UIB', 'ativa', dataCadastroUnidades]
    ];
    unidadesSheet.getRange(2, 1, unidadesIniciais.length, 4).setValues(unidadesIniciais);
  }
}

// Função principal para lidar com requisições POST
var usuarioContextoRequisicao = '';
var usuarioContextoRequisicaoId = null;

// Contexto de sessão resolvido server-side (fonte de verdade para autorização).
var sessaoTokenAtual = '';
var sessaoUsuarioIdAtual = null;
var sessaoPerfilAtual = '';
var sessaoPrecisaTrocarSenha = false;

// Ações liberadas sem sessão (login, verificação/inicialização da planilha).
var ACOES_PUBLICAS = {
  autenticarUsuario: true,
  verificarInicializacao: true,
  inicializarPlanilha: true
};

// Ações que exigem perfil administrador.
var ACOES_ADMIN = {
  getUsuarios: true,
  cadastrarUsuario: true,
  atualizarUsuario: true,
  excluirUsuario: true,
  getLogs: true,
  executarBackupSistema: true,
  cadastrarUnidade: true,
  alternarStatusUnidade: true,
  // Cadastro/ocupação de armários é restrito a admin (espelha as travas da UI:
  // "Apenas administradores podem registrar ocupações" e a página de cadastro físico).
  cadastrarArmario: true,
  cadastrarArmarioFisico: true
};

// Ações permitidas mesmo quando o usuário precisa trocar a senha.
var ACOES_SEM_BLOQUEIO_SENHA = {
  trocarSenha: true,
  encerrarSessao: true
};

function respostaJSON(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

function respostaAcessoNegado(motivo, mensagem) {
  return respostaJSON({ success: false, error: mensagem || 'Acesso negado.', motivo: motivo });
}

function resetarContextoSessao() {
  sessaoTokenAtual = '';
  sessaoUsuarioIdAtual = null;
  sessaoPerfilAtual = '';
  sessaoPrecisaTrocarSenha = false;
}

function definirContextoUsuario(parametros) {
  usuarioContextoRequisicao = '';
  usuarioContextoRequisicaoId = null;
  try {
    if (!parametros) {
      return;
    }

    if (parametros.usuarioId !== undefined && parametros.usuarioId !== null) {
      var idTexto = parametros.usuarioId.toString().trim();
      if (idTexto) {
        var idNumero = parseInt(idTexto, 10);
        if (!isNaN(idNumero)) {
          usuarioContextoRequisicaoId = idNumero;
        }
      }
    }

    var camposCandidatos = [
      'usuarioResponsavel',
      'usuario',
      'responsavel',
      'responsavelRegistro',
      'usuarioContexto',
      'usuarioAcao'
    ];

    for (var i = 0; i < camposCandidatos.length; i++) {
      var chave = camposCandidatos[i];
      if (parametros[chave] !== undefined && parametros[chave] !== null) {
        var texto = parametros[chave].toString().trim();
        if (texto) {
          usuarioContextoRequisicao = texto;
          return;
        }
      }
    }
  } catch (erroContexto) {
    usuarioContextoRequisicao = '';
    usuarioContextoRequisicaoId = null;
  }
}

function limparContextoUsuario() {
  usuarioContextoRequisicao = '';
  usuarioContextoRequisicaoId = null;
  resetarContextoSessao();
}

function handlePost(e) {
  var action = e.parameter.action;
  var inicioExecucao = Date.now();
  var acoesComLock = {
    cadastrarArmario: true,
    atualizarHorarioVisitante: true,
    atualizarDadosArmario: true,
    finalizarELiberarArmario: true,
    liberarArmario: true,
    cadastrarUsuario: true,
    atualizarUsuario: true,
    excluirUsuario: true,
    registrarContingencia: true,
    registrarContingenciaTermo: true,
    salvarMovimentacao: true,
    finalizarMovimentacoesArmario: true,
    cadastrarArmarioFisico: true,
    criarArmariosUso: true,
    cadastrarUnidade: true,
    alternarStatusUnidade: true,
    salvarTermoCompleto: true,
    finalizarTermo: true,
    cadastrarPertencePerdido: true,
    atualizarPertencePerdido: true,
    registrarContatoPertence: true,
    excluirPertencePerdido: true
  };

  if (acoesComLock[action] && (!e.parameter._lockApplied || e.parameter._lockApplied !== '1')) {
    var parametrosComLock = {};
    for (var chaveLock in e.parameter) {
      if (Object.prototype.hasOwnProperty.call(e.parameter, chaveLock)) {
        parametrosComLock[chaveLock] = e.parameter[chaveLock];
      }
    }
    parametrosComLock._lockApplied = '1';

    return executarComRetry(function() {
      return executarComLock(action, function() {
        return handlePost({ parameter: parametrosComLock });
      });
    });
  }

  definirContextoUsuario(e && e.parameter);

  try {
    // ===== Gate central de autenticação/autorização =====
    resetarContextoSessao();
    var tokenRequisicao = (e.parameter && e.parameter.token) ? e.parameter.token.toString() : '';
    var sessaoResolvida = null;
    try {
      sessaoResolvida = resolverSessao(tokenRequisicao);
    } catch (erroSessao) {
      sessaoResolvida = null;
    }

    if (sessaoResolvida) {
      sessaoTokenAtual = tokenRequisicao;
      sessaoUsuarioIdAtual = sessaoResolvida.usuarioId;
      sessaoPerfilAtual = sessaoResolvida.perfil;
      sessaoPrecisaTrocarSenha = sessaoResolvida.precisaTrocarSenha;
      // O id confiável da sessão sobrescreve qualquer usuarioId enviado pelo cliente.
      usuarioContextoRequisicaoId = sessaoResolvida.usuarioId;
    }

    if (!ACOES_PUBLICAS[action]) {
      if (!sessaoResolvida) {
        return respostaAcessoNegado('SESSAO_INVALIDA', 'Sessão inválida ou expirada. Faça login novamente.');
      }
      if (sessaoPrecisaTrocarSenha && !ACOES_SEM_BLOQUEIO_SENHA[action]) {
        return respostaAcessoNegado('SENHA_EXPIRADA', 'É necessário trocar a senha antes de continuar.');
      }
      if (ACOES_ADMIN[action] && normalizarTextoBasico(sessaoPerfilAtual) !== 'admin') {
        return respostaAcessoNegado('SEM_PERMISSAO', 'Acesso restrito a administradores.');
      }
    }

    switch(action) {
      case 'getArmarios':
        return ContentService.createTextOutput(JSON.stringify(getArmarios(
          e.parameter.tipo,
          e.parameter.incluirInternacoes,
          e.parameter.incluirTermos
        )))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastrarArmario':
        return ContentService.createTextOutput(JSON.stringify(cadastrarArmario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'atualizarHorarioVisitante':
        return ContentService.createTextOutput(JSON.stringify(atualizarHorarioVisitante(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'atualizarDadosArmario':
        return ContentService.createTextOutput(JSON.stringify(atualizarDadosArmario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'finalizarELiberarArmario':
        return ContentService.createTextOutput(JSON.stringify(finalizarELiberarArmario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'liberarArmario':
        return ContentService.createTextOutput(JSON.stringify(liberarArmario(
          e.parameter.id,
          e.parameter.tipo,
          e.parameter.numero,
          e.parameter.usuarioResponsavel
        )))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getUsuarios':
        return ContentService.createTextOutput(JSON.stringify(getUsuarios()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastrarUsuario':
        return ContentService.createTextOutput(JSON.stringify(cadastrarUsuario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'atualizarUsuario':
        return ContentService.createTextOutput(JSON.stringify(atualizarUsuario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'excluirUsuario':
        return ContentService.createTextOutput(JSON.stringify(excluirUsuario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'autenticarUsuario':
        return ContentService.createTextOutput(JSON.stringify(autenticarUsuario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'trocarSenha':
        return ContentService.createTextOutput(JSON.stringify(trocarSenha(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'encerrarSessao':
        return ContentService.createTextOutput(JSON.stringify(encerrarSessao(sessaoTokenAtual || (e.parameter && e.parameter.token))))
          .setMimeType(ContentService.MimeType.JSON);

      case 'registrarContingencia':
        return ContentService.createTextOutput(JSON.stringify(registrarContingencia(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'registrarContingenciaTermo':
        return ContentService.createTextOutput(JSON.stringify(registrarContingenciaTermo(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getLogs':
        return ContentService.createTextOutput(JSON.stringify(getLogs()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getNotificacoes':
        return ContentService.createTextOutput(JSON.stringify(getNotificacoes()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getEstatisticasDashboard':
        return ContentService.createTextOutput(JSON.stringify(getEstatisticasDashboard(e.parameter.tipoUsuario)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getDescarte':
        return ContentService.createTextOutput(JSON.stringify(getItensDescarte()))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getHistorico':
        return ContentService.createTextOutput(JSON.stringify(getHistorico(e.parameter.tipo)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getPlanilhaLiberacao':
        return ContentService.createTextOutput(JSON.stringify(getPlanilhaLiberacao(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getCadastroArmarios':
        return ContentService.createTextOutput(JSON.stringify(getCadastroArmarios()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastrarArmarioFisico':
        return ContentService.createTextOutput(JSON.stringify(cadastrarArmarioFisico(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getUnidades':
        return ContentService.createTextOutput(JSON.stringify(getUnidades()))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getSetores':
        return ContentService.createTextOutput(JSON.stringify(getSetores()))
          .setMimeType(ContentService.MimeType.JSON);

      case 'buscarPacienteBaseVitae':
        return ContentService.createTextOutput(JSON.stringify(buscarPacienteBaseVitae(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'cadastrarUnidade':
        return ContentService.createTextOutput(JSON.stringify(cadastrarUnidade(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'alternarStatusUnidade':
        return ContentService.createTextOutput(JSON.stringify(alternarStatusUnidade(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'salvarTermoCompleto':
        return ContentService.createTextOutput(JSON.stringify(salvarTermoCompleto(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'finalizarTermo':
        return ContentService.createTextOutput(JSON.stringify(finalizarTermo(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'gerarTermoPDFTemporario':
        return ContentService.createTextOutput(JSON.stringify(gerarTermoPDFTemporario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getTermo':
        return ContentService.createTextOutput(JSON.stringify(getTermo(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'excluirArquivoTemporario':
        return ContentService.createTextOutput(JSON.stringify(excluirArquivoTemporario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getMovimentacoesResumo':
        return ContentService.createTextOutput(JSON.stringify(getMovimentacoesResumo(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getMovimentacoes':
        return ContentService.createTextOutput(JSON.stringify(getMovimentacoes(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getRegistrosImagens':
        return ContentService.createTextOutput(JSON.stringify(getRegistrosImagens(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'salvarMovimentacao':
        return ContentService.createTextOutput(JSON.stringify(salvarMovimentacao(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'verificarInicializacao':
        return ContentService.createTextOutput(JSON.stringify(verificarInicializacao()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'inicializarPlanilha':
        return ContentService.createTextOutput(JSON.stringify(inicializarPlanilha()))
          .setMimeType(ContentService.MimeType.JSON);

      case 'getPertencesPerdidos':
        return ContentService.createTextOutput(JSON.stringify(listarPertencesPerdidos()))
          .setMimeType(ContentService.MimeType.JSON);

      case 'cadastrarPertencePerdido':
        return ContentService.createTextOutput(JSON.stringify(cadastrarPertencePerdido(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'atualizarPertencePerdido':
        return ContentService.createTextOutput(JSON.stringify(atualizarPertencePerdido(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'registrarContatoPertence':
        return ContentService.createTextOutput(JSON.stringify(registrarContatoPertence(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'excluirPertencePerdido':
        return ContentService.createTextOutput(JSON.stringify(excluirPertencePerdido(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'executarBackupSistema':
        return ContentService.createTextOutput(JSON.stringify(executarBackupSistema(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      default:
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Ação não reconhecida: ' + action }))
          .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    registrarLog('ERRO', `Erro em handlePost (${action}): ${error.toString()}`);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Erro ao processar a solicitação.' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    var duracaoMs = Date.now() - inicioExecucao;
    if (duracaoMs >= LIMIAR_LOG_LENTO_MS) {
      console.log('Ação lenta detectada:', action, '| duração(ms):', duracaoMs);
    }
    limparContextoUsuario();
  }
}

// Funções para Armários
function obterMapaInternacoesBaseVitae() {
  var chaveCache = montarChaveCache('base-vitae', 'internacoes');
  var resultado = executarComCache(chaveCache, CACHE_TTL_PADRAO, function() {
    try {
      var sheet = obterAbaEntradaClinicaExterna();
      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: {} };
      }

      var estrutura = obterEstruturaPlanilha(sheet);
      var totalLinhas = sheet.getLastRow() - 1;
      var totalColunas = estrutura.ultimaColuna || sheet.getLastColumn();
      var dados = sheet.getRange(2, 1, totalLinhas, totalColunas).getValues();

      var prontuarioIndex = obterIndiceColuna(estrutura, ['prontuario'], 0);
      var entradaIndex = obterIndiceColuna(estrutura, ['entrada', 'data entrada'], 12);
      var saidaIndex = obterIndiceColuna(estrutura, ['saida', 'saída'], 13);
      var destinoIndex = obterIndiceColuna(estrutura, ['destino'], 16);

      var mapa = {};

      dados.forEach(function(linha, idx) {
        var prontuarioValor = obterValorLinhaFlexivel(linha, estrutura, ['prontuario'], linha[prontuarioIndex]);
        var prontuario = normalizarIdentificador(prontuarioValor);
        if (!prontuario) {
          return;
        }

        var destinoBruto = destinoIndex > -1 ? linha[destinoIndex] : '';
        var destinoNormalizado = normalizarTextoBasico(destinoBruto);
        var entradaData = entradaIndex > -1 ? obterDataValida(linha[entradaIndex]) : null;
        var saidaData = saidaIndex > -1 ? obterDataValida(linha[saidaIndex]) : null;
        var dataReferencia = saidaData || entradaData || new Date(0);

        var registroAtual = mapa[prontuario];
        if (registroAtual && registroAtual.dataReferencia && registroAtual.dataReferencia.getTime() > dataReferencia.getTime()) {
          return;
        }

        var internado = destinoNormalizado === 'encontra-se internado' || destinoNormalizado === 'encontrase internado';
        var ultimaAlta = registroAtual && registroAtual.ultimaAlta ? registroAtual.ultimaAlta : null;

        if (!internado && saidaData) {
          if (!ultimaAlta || saidaData.getTime() > ultimaAlta.getTime()) {
            ultimaAlta = saidaData;
          }
        }

        mapa[prontuario] = {
          internadoAtual: internado,
          destinoAtual: destinoBruto || '',
          dataReferencia: dataReferencia,
          ultimaAlta: ultimaAlta
        };
      });

      return { success: true, data: mapa };
    } catch (erro) {
      registrarLog('ERRO', 'Falha ao mapear PENTAHO_ENTRADA_CLINICA: ' + erro.toString());
      return { success: true, data: {} };
    }
  });

  if (resultado && resultado.success && resultado.data) {
    return resultado.data;
  }
  return {};
}

function obterMapaPacientesBaseVitae() {
  var chaveCache = montarChaveCache('base-vitae', 'pacientes-detalhes');
  var resultado = executarComCache(chaveCache, CACHE_TTL_PADRAO, function() {
    try {
      var sheet = obterAbaEntradaClinicaExterna();
      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: {} };
      }

      var estrutura = obterEstruturaPlanilha(sheet);
      var totalLinhas = sheet.getLastRow() - 1;
      var totalColunas = estrutura.ultimaColuna || sheet.getLastColumn();
      var dados = sheet.getRange(2, 1, totalLinhas, totalColunas).getValues();

      var prontuarioIndex = obterIndiceColuna(estrutura, ['prontuario'], 0);
      var nomeIndex = obterIndiceColuna(estrutura, ['paciente', 'nome paciente', 'nome do paciente', 'nome'], 1);
      var setorIndex = obterIndiceColuna(estrutura, ['setor internacao', 'setor internação', 'setor'], 6);
      var enfermariaIndex = obterIndiceColuna(estrutura, ['enfermaria'], 9);
      var leitoIndex = obterIndiceColuna(estrutura, ['leito'], 10);
      var entradaIndex = obterIndiceColuna(estrutura, ['entrada', 'data entrada'], 12);
      var saidaIndex = obterIndiceColuna(estrutura, ['saida', 'saída'], 13);

      var mapa = {};

      dados.forEach(function(linha) {
        var prontuarioValor = obterValorLinhaFlexivel(linha, estrutura, ['prontuario'], linha[prontuarioIndex]);
        var prontuario = normalizarIdentificador(prontuarioValor);
        if (!prontuario) {
          return;
        }

        var entradaData = entradaIndex > -1 ? obterDataValida(linha[entradaIndex]) : null;
        var saidaData = saidaIndex > -1 ? obterDataValida(linha[saidaIndex]) : null;
        var dataReferencia = saidaData || entradaData || new Date(0);
        var dataReferenciaNumero = dataReferencia instanceof Date ? dataReferencia.getTime() : 0;

        var registroAtual = mapa[prontuario];
        if (registroAtual && registroAtual.dataReferencia !== undefined && registroAtual.dataReferencia > dataReferenciaNumero) {
          return;
        }

        mapa[prontuario] = {
          nome: obterValorLinhaFlexivel(linha, estrutura, ['paciente', 'nome paciente', 'nome do paciente', 'nome'], linha[nomeIndex]) || '',
          setor: obterValorLinhaFlexivel(linha, estrutura, ['setor internacao', 'setor internação', 'setor'], linha[setorIndex]) || '',
          enfermaria: obterValorLinhaFlexivel(linha, estrutura, ['enfermaria'], linha[enfermariaIndex]) || '',
          leito: obterValorLinhaFlexivel(linha, estrutura, ['leito'], linha[leitoIndex]) || '',
          dataReferencia: dataReferenciaNumero
        };
      });

      return { success: true, data: mapa };
    } catch (erro) {
      registrarLog('ERRO', 'Falha ao mapear PENTAHO_ENTRADA_CLINICA para busca de pacientes: ' + erro.toString());
      return { success: true, data: {} };
    }
  });

  if (resultado && resultado.success && resultado.data) {
    return resultado.data;
  }
  return {};
}

function obterMapaNascimentosPacientes() {
  var chaveCache = montarChaveCache('paciente-por-setor', 'nascimentos');
  return executarComCache(chaveCache, CACHE_TTL_PADRAO, function() {
    try {
      var sheet = obterAbaEntradaClinicaExterna();
      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: {} };
      }

      var totalColunas = Math.max(sheet.getLastColumn(), 8);
      var dadosSheet = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
      var mapa = {};

      dadosSheet.forEach(function(linha) {
        var prontuarioLinha = normalizarIdentificador(linha[2]);
        if (!prontuarioLinha) {
          return;
        }

        var prontuarioLinhaSemZeros = prontuarioLinha.replace(/^0+/, '') || prontuarioLinha;
        var valorNascimento = linha[7];
        var nascimentoFormatado = formatarDataDDMMAAAA(valorNascimento) || (valorNascimento ? valorNascimento.toString().trim() : '');

        if (nascimentoFormatado) {
          mapa[prontuarioLinha] = nascimentoFormatado;
          if (prontuarioLinhaSemZeros !== prontuarioLinha) {
            mapa[prontuarioLinhaSemZeros] = nascimentoFormatado;
          }
        }
      });

      return { success: true, data: mapa };
    } catch (erroMapa) {
      registrarLog('ERRO', 'Falha ao montar mapa de nascimentos: ' + erroMapa.toString());
      return { success: true, data: {} };
    }
  });
}

function buscarNascimentoPorProntuario(prontuarioEntrada, prontuarioSemZeros) {
  try {
    var mapaResultado = obterMapaNascimentosPacientes();
    if (!mapaResultado || !mapaResultado.success) {
      return '';
    }

    var mapa = mapaResultado.data || {};
    return mapa[prontuarioEntrada] || mapa[prontuarioSemZeros] || '';
  } catch (erro) {
    registrarLog('ERRO', 'Falha ao buscar nascimento na aba PENTAHO_ENTRADA_CLINICA: ' + erro.toString());
    return '';
  }
}

function obterMapaPacientesBaseVitae() {
  var chaveCache = montarChaveCache('base-vitae', 'pacientes');
  return executarComCache(chaveCache, CACHE_TTL_PADRAO, function() {
    try {
      var sheet = obterAbaEntradaClinicaExterna();
      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: {} };
      }

      var estrutura = obterEstruturaPlanilha(sheet);
      var totalLinhas = sheet.getLastRow() - 1;
      var totalColunas = estrutura.ultimaColuna || Math.max(sheet.getLastColumn(), 13);
      var dadosSheet = sheet.getRange(2, 1, totalLinhas, totalColunas).getValues();

      var prontuarioIndex = obterIndiceColuna(estrutura, ['prontuario'], 0);
      var nomeIndex = obterIndiceColuna(estrutura, ['nome'], 1);
      var setorIndex = obterIndiceColuna(estrutura, ['setor'], 6);
      var leitoAIndex = obterIndiceColuna(estrutura, ['parte leito a', 'leito a'], 9);
      var leitoBIndex = obterIndiceColuna(estrutura, ['parte leito b', 'leito b'], 10);
      var referenciaIndex = obterIndiceColuna(estrutura, ['data referencia', 'referencia', 'referência'], 12);

      var mapa = {};

      dadosSheet.forEach(function(linha) {
        var prontuarioValor = obterValorLinhaFlexivel(linha, estrutura, ['prontuario'], linha[prontuarioIndex]);
        var prontuarioLinha = normalizarIdentificador(prontuarioValor);
        if (!prontuarioLinha) {
          return;
        }

        var prontuarioSemZeros = prontuarioLinha.replace(/^0+/, '') || prontuarioLinha;
        var nome = linha[nomeIndex] ? linha[nomeIndex].toString().trim() : '';
        var setor = linha[setorIndex] ? linha[setorIndex].toString().trim() : '';
        var leitoA = linha[leitoAIndex] ? linha[leitoAIndex].toString().trim() : '';
        var leitoB = linha[leitoBIndex] ? linha[leitoBIndex].toString().trim() : '';
        var dataReferencia = obterDataValida(linha[referenciaIndex]);
        var timestampReferencia = dataReferencia ? dataReferencia.getTime() : 0;

        var registroAtual = mapa[prontuarioLinha];
        if (registroAtual && registroAtual.dataReferencia > timestampReferencia) {
          return;
        }

        var registro = {
          prontuario: prontuarioLinha,
          nome: nome,
          leito: [leitoA, leitoB].filter(Boolean).join(' - '),
          setor: setor,
          dataReferencia: timestampReferencia
        };

        mapa[prontuarioLinha] = registro;
        if (prontuarioSemZeros !== prontuarioLinha) {
          mapa[prontuarioSemZeros] = registro;
        }
      });

      return { success: true, data: mapa };
    } catch (erroMapa) {
      registrarLog('ERRO', 'Falha ao montar mapa da PENTAHO_ENTRADA_CLINICA: ' + erroMapa.toString());
      return { success: true, data: {} };
    }
  });
}

function buscarPacienteBaseVitae(dados) {
  try {
    var prontuarioEntrada = dados && dados.prontuario ? normalizarIdentificador(dados.prontuario) : '';
    if (!prontuarioEntrada) {
      return { success: false, error: 'Prontuário não informado' };
    }

    var prontuarioSemZeros = prontuarioEntrada.replace(/^0+/, '') || prontuarioEntrada;
    var registroMaisRecente = obterRegistroBaseVitaePorProntuario(prontuarioEntrada, prontuarioSemZeros);

    if (!registroMaisRecente) {
      return { success: false, error: 'Prontuário não localizado. Insira Manualmente os dados' };
    }

    var nascimento = buscarNascimentoPorProntuario(prontuarioEntrada, prontuarioSemZeros);
    var setorInternacao = registroMaisRecente.setor || '';

    return {
      success: true,
      data: {
        prontuario: registroMaisRecente.prontuario || prontuarioEntrada,
        nome: registroMaisRecente.nome || '',
        leito: registroMaisRecente.leito || '',
        setor: setorInternacao,
        nascimento: nascimento
      }
    };
  } catch (erroBusca) {
    registrarLog('ERRO', 'Falha ao buscar paciente na PENTAHO_ENTRADA_CLINICA: ' + erroBusca.toString());
    return { success: false, error: 'Erro ao buscar paciente na PENTAHO_ENTRADA_CLINICA' };
  }
}

function obterRegistroBaseVitaePorProntuario(prontuarioEntrada, prontuarioSemZeros) {
  try {
    var sheet = obterAbaEntradaClinicaExterna();
    if (!sheet || sheet.getLastRow() < 2) {
      return null;
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var prontuarioIndex = obterIndiceColuna(estrutura, ['prontuario'], 0);
    if (prontuarioIndex === null || prontuarioIndex === undefined) {
      return null;
    }

    var totalLinhas = sheet.getLastRow() - 1;
    var totalColunas = estrutura.ultimaColuna || sheet.getLastColumn();
    var colunaProntuario = sheet.getRange(2, prontuarioIndex + 1, totalLinhas, 1);

    var encontrarMatches = function(valor) {
      if (!valor) {
        return [];
      }
      return colunaProntuario
        .createTextFinder(valor)
        .matchEntireCell(true)
        .findAll();
    };

    var matches = encontrarMatches(prontuarioEntrada);
    if (!matches.length && prontuarioSemZeros && prontuarioSemZeros !== prontuarioEntrada) {
      matches = encontrarMatches(prontuarioSemZeros);
    }

    if (!matches.length) {
      return null;
    }

    var nomeIndex = obterIndiceColuna(estrutura, ['nome'], 1);
    var setorIndex = obterIndiceColuna(estrutura, ['setor'], 6);
    var leitoAIndex = obterIndiceColuna(estrutura, ['parte leito a', 'leito a'], 9);
    var leitoBIndex = obterIndiceColuna(estrutura, ['parte leito b', 'leito b'], 10);
    var referenciaIndex = obterIndiceColuna(estrutura, ['data referencia', 'referencia', 'referência'], 12);

    var registroMaisRecente = null;
    var referenciaMaisRecente = 0;

    matches.forEach(function(match) {
      var row = match.getRow();
      var linha = sheet.getRange(row, 1, 1, totalColunas).getValues()[0];
      var prontuarioValor = obterValorLinhaFlexivel(linha, estrutura, ['prontuario'], linha[prontuarioIndex]);
      var prontuarioLinha = normalizarIdentificador(prontuarioValor);
      if (!prontuarioLinha) {
        return;
      }

      var dataReferencia = obterDataValida(linha[referenciaIndex]);
      var timestampReferencia = dataReferencia ? dataReferencia.getTime() : 0;
      if (registroMaisRecente && referenciaMaisRecente > timestampReferencia) {
        return;
      }

      var leitoA = linha[leitoAIndex] ? linha[leitoAIndex].toString().trim() : '';
      var leitoB = linha[leitoBIndex] ? linha[leitoBIndex].toString().trim() : '';

      registroMaisRecente = {
        prontuario: prontuarioLinha,
        nome: linha[nomeIndex] ? linha[nomeIndex].toString().trim() : '',
        leito: [leitoA, leitoB].filter(Boolean).join(' - '),
        setor: linha[setorIndex] ? linha[setorIndex].toString().trim() : '',
        dataReferencia: timestampReferencia
      };
      referenciaMaisRecente = timestampReferencia;
    });

    if (!registroMaisRecente) {
      return null;
    }

    if (prontuarioSemZeros && registroMaisRecente.prontuario && registroMaisRecente.prontuario !== prontuarioSemZeros) {
      var prontuarioSemZerosLinha = registroMaisRecente.prontuario.replace(/^0+/, '') || registroMaisRecente.prontuario;
      if (prontuarioSemZerosLinha === prontuarioSemZeros) {
        registroMaisRecente.prontuario = prontuarioSemZerosLinha;
      }
    }

    return registroMaisRecente;
  } catch (erro) {
    registrarLog('ERRO', 'Falha ao buscar prontuário na Base Vitae: ' + erro.toString());
    return null;
  }
}

function getArmarios(tipo, incluirInternacoes, incluirTermos) {
  var tipoNormalizadoOriginal = normalizarTextoBasico(tipo);
  if (!tipoNormalizadoOriginal) {
    tipoNormalizadoOriginal = 'geral';
  }

  var incluirInternacoesNormalizado = converterParaBoolean(incluirInternacoes);
  var incluirTermosNormalizado = incluirTermos === undefined ? true : converterParaBoolean(incluirTermos);
  var chaveCacheTipo = tipoNormalizadoOriginal;
  if (chaveCacheTipo === 'admin' || chaveCacheTipo === 'ambos' || chaveCacheTipo === 'todos') {
    chaveCacheTipo = 'geral';
  }

  var chaveCache = montarChaveCache(
    'armarios',
    chaveCacheTipo,
    incluirInternacoesNormalizado ? 'com-internacoes' : 'sem-internacoes',
    incluirTermosNormalizado ? 'com-termos' : 'sem-termos'
  );

  return executarComCache(chaveCache, CACHE_TTL_ARMARIOS, function() {
    try {
      var mapaInternacoes = incluirInternacoesNormalizado ? obterMapaInternacoesBaseVitae() : null;
      var tipoNormalizado = tipoNormalizadoOriginal;
      var incluirTermos = tipoNormalizado === 'acompanhante' || tipoNormalizado === 'admin' ||
        tipoNormalizado === 'ambos' || tipoNormalizado === 'todos' || tipoNormalizado === 'geral';
      var termosMap = {};

      if (incluirTermos && incluirTermosNormalizado) {
        var termosInfo = obterTermosRegistrados();
        termosInfo.termos.forEach(function(termo) {
          if (!termo) {
            return;
          }

          var chaveId = '';
          if (termo.armarioId !== null && termo.armarioId !== undefined) {
            chaveId = termo.armarioId.toString().trim();
          }

          if (!chaveId) {
            return;
          }

          if (!termosMap[chaveId]) {
            termosMap[chaveId] = {};
          }

          var numeroChave = obterChaveNumeroArmario(termo.numeroArmario);
          var termoAtual = termosMap[chaveId][numeroChave];
          var termoFinalizado = termo.finalizado;
          if (termoFinalizado === undefined) {
            var statusTermo = normalizarTextoBasico(termo.status);
            termoFinalizado = Boolean(termo.pdfUrl || (termo.assinaturas && termo.assinaturas.finalizadoEm) || statusTermo === 'finalizado');
          }

          if (!termoAtual) {
            termosMap[chaveId][numeroChave] = termo;
            return;
          }

          var atualFinalizado = termoAtual.finalizado;
          if (atualFinalizado === undefined) {
            var statusAtual = normalizarTextoBasico(termoAtual.status);
            atualFinalizado = Boolean(termoAtual.pdfUrl || (termoAtual.assinaturas && termoAtual.assinaturas.finalizadoEm) || statusAtual === 'finalizado');
          }

          if (!termoFinalizado && atualFinalizado) {
            termosMap[chaveId][numeroChave] = termo;
          } else if (termoFinalizado === atualFinalizado) {
            var idAtualNumero = Number(termoAtual.id) || 0;
            var idNovoNumero = Number(termo.id) || 0;
            if (idNovoNumero > idAtualNumero) {
              termosMap[chaveId][numeroChave] = termo;
            }
          }
        });
      }

      if (tipoNormalizado === 'admin' || tipoNormalizado === 'ambos' || tipoNormalizado === 'todos' || tipoNormalizado === 'geral') {
        var visitantes = getArmariosFromSheet('Visitantes', 'visitante', null, mapaInternacoes);
        var acompanhantes = getArmariosFromSheet('Acompanhantes', 'acompanhante', termosMap, mapaInternacoes);
        return { success: true, data: visitantes.concat(acompanhantes) };
      }

      var sheetName = tipoNormalizado === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
      var mapa = tipoNormalizado === 'acompanhante' ? termosMap : null;
      return { success: true, data: getArmariosFromSheet(sheetName, tipoNormalizado, mapa, mapaInternacoes) };
    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar armários: ${error.toString()}`);
      return { success: false, error: 'Erro ao processar a solicitação.' };
    }
  });
}

function getArmariosFromSheet(sheetName, tipo, termosMap, mapaInternacoes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var isVisitante = sheetName === 'Visitantes';
  var estrutura = obterEstruturaPlanilha(sheet);
  if (isVisitante) {
    estrutura = garantirColunaVisitaEstendida(sheet, estrutura);
  }
  estrutura = garantirColunaProntuario(sheet, estrutura);
  if (!isVisitante) {
    estrutura = garantirColunaObservacoesAcompanhantes(sheet, estrutura);
    estrutura = garantirColunasFotoContingencia(sheet, estrutura);
  }
  var totalLinhas = sheet.getLastRow() - 1;
  var totalColunas = estrutura.ultimaColuna || (isVisitante ? 14 : 12);
  var dados = sheet.getRange(2, 1, totalLinhas, totalColunas).getValues();
  var armarios = [];

  var idIndex = obterIndiceColuna(estrutura, 'id', 0);
  var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
  var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
  var nomeChaves = isVisitante ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE;
  var nomeIndex = obterIndiceColuna(estrutura, nomeChaves, 3);
  var pacienteIndex = obterIndiceColuna(estrutura, 'nome paciente', 4);
  var leitoIndex = obterIndiceColuna(estrutura, 'leito', 5);
  var volumesIndex = obterIndiceColuna(estrutura, 'volumes', 6);
  var horaInicioIndex = obterIndiceColuna(estrutura, 'hora inicio', 7);
  var horaPrevistaIndex = isVisitante ? obterIndiceColuna(estrutura, 'hora prevista', 8) : -1;
  var dataRegistroIndex = obterIndiceColuna(estrutura, 'data registro', isVisitante ? 9 : 8);
  var unidadeIndex = obterIndiceColuna(estrutura, 'unidade', null);
  if (unidadeIndex === null || unidadeIndex === undefined) {
    unidadeIndex = isVisitante ? 10 : 10;
  }
  var termoIndex = obterIndiceColuna(estrutura, 'termo aplicado', null);
  if (termoIndex === null || termoIndex === undefined) {
    termoIndex = isVisitante ? 11 : 11;
  }
  var whatsappIndex = obterIndiceColuna(estrutura, CABECALHOS_WHATSAPP, null);
  if (whatsappIndex === null || whatsappIndex === undefined) {
    whatsappIndex = isVisitante ? 12 : 9;
  }
  var observacoesIndex = obterIndiceColuna(estrutura, CABECALHOS_OBSERVACOES, null);
  var visitaEstendidaIndex = isVisitante
    ? obterIndiceColuna(estrutura, CABECALHOS_VISITA_ESTENDIDA, null)
    : -1;
  var fotoContingenciaUrlIndex = !isVisitante
    ? obterIndiceColuna(estrutura, ['foto contingencia url', 'foto contingência url'], null)
    : -1;
  var fotoContingenciaIdIndex = !isVisitante
    ? obterIndiceColuna(estrutura, ['foto contingencia id', 'foto contingência id'], null)
    : -1;

  var houveAtualizacaoStatus = false;
  var agoraReferencia = new Date();

  for (var i = 0; i < dados.length; i++) {
    var row = dados[i];
    var idPlanilha = row[idIndex];
    if (!idPlanilha && idPlanilha !== 0) {
      continue;
    }

    var statusValor = row[statusIndex];
    var statusNormalizado = normalizarTextoBasico(statusValor);
    var status;
    switch (statusNormalizado) {
      case 'em-uso':
      case 'em uso':
        status = 'em-uso';
        break;
      case 'proximo':
        status = 'proximo';
        break;
      case 'vencido':
        status = 'vencido';
        break;
      case 'livre':
        status = 'livre';
        break;
      default:
        status = statusNormalizado || 'livre';
        break;
    }

    if (isVisitante && statusIndex > -1) {
      var dataRegistroBruta = dataRegistroIndex > -1 ? row[dataRegistroIndex] : null;
      var horaPrevistaBruta = horaPrevistaIndex > -1 ? row[horaPrevistaIndex] : null;
      var novoStatus = calcularStatusAutomaticoVisitante(status, dataRegistroBruta, horaPrevistaBruta, agoraReferencia);
      if (novoStatus && novoStatus !== status) {
        status = novoStatus;
        row[statusIndex] = novoStatus;
        houveAtualizacaoStatus = true;
      }
    }

    var numeroBruto = row[numeroIndex] || '';
    var numeroNormalizado = normalizarNumeroArmario(numeroBruto);
    var idInterface = montarChaveArmarioInterface(tipo, numeroNormalizado, idPlanilha);
    var ehContingencia = ehNumeroContingencia(numeroNormalizado) || status === 'contingencia';
    var prontuarioNormalizado = normalizarIdentificador(obterValorLinha(row, estrutura, 'prontuario', ''));
    var infoInternacao = prontuarioNormalizado && mapaInternacoes ? mapaInternacoes[prontuarioNormalizado] : null;

    var armario = {
      id: idInterface,
      idPlanilha: idPlanilha,
      numero: numeroNormalizado,
      status: status,
      nomeVisitante: obterValorLinha(row, estrutura, nomeChaves, row[nomeIndex] || ''),
      nomePaciente: row[pacienteIndex] || '',
      prontuario: prontuarioNormalizado,
      leito: row[leitoIndex] || '',
      volumes: row[volumesIndex] || 0,
      horaInicio: formatarHorarioPlanilha(row[horaInicioIndex]),
      tipo: tipo,
      unidade: unidadeIndex !== null && unidadeIndex !== undefined ? (row[unidadeIndex] || '') : '',
      termoAplicado: termoIndex !== null && termoIndex !== undefined ? converterParaBoolean(row[termoIndex]) : false,
      whatsapp: whatsappIndex !== null && whatsappIndex !== undefined ? (row[whatsappIndex] || '') : '',
      observacoes: observacoesIndex !== null && observacoesIndex !== undefined ? (row[observacoesIndex] || '') : '',
      fotoContingenciaUrl: fotoContingenciaUrlIndex > -1 ? (row[fotoContingenciaUrlIndex] || '') : '',
      fotoContingenciaId: fotoContingenciaIdIndex > -1 ? (row[fotoContingenciaIdIndex] || '') : '',
      ehContingencia: ehContingencia,
      pacienteDeAlta: infoInternacao ? !infoInternacao.internadoAtual : false,
      destinoAtual: infoInternacao ? (infoInternacao.destinoAtual || '') : '',
      dataAltaReferencia: infoInternacao && infoInternacao.ultimaAlta ? formatarDataPlanilha(infoInternacao.ultimaAlta) : '',
      dataAltaIso: infoInternacao && infoInternacao.ultimaAlta ? infoInternacao.ultimaAlta.toISOString() : ''
    };

    if (isVisitante) {
      armario.horaPrevista = horaPrevistaIndex > -1 ? formatarHorarioPlanilha(row[horaPrevistaIndex]) : '';
      armario.dataRegistro = dataRegistroIndex > -1 ? formatarDataPlanilha(row[dataRegistroIndex]) : '';
      armario.visitaEstendida = visitaEstendidaIndex > -1
        ? converterParaBoolean(row[visitaEstendidaIndex])
        : false;
    } else {
      armario.dataRegistro = dataRegistroIndex > -1 ? formatarDataPlanilha(row[dataRegistroIndex]) : '';
      armario.visitaEstendida = false;
    }

    var volumesNumero = parseInt(armario.volumes, 10);
    armario.volumes = isNaN(volumesNumero) ? 0 : volumesNumero;

    if (tipo === 'acompanhante') {
      var termosPorId = null;
      if (termosMap) {
        var chaveId = idPlanilha !== null && idPlanilha !== undefined ? idPlanilha.toString().trim() : '';
        termosPorId = chaveId ? termosMap[chaveId] : null;
      }

      var termoRelacionado = null;
      if (termosPorId) {
        var chaveNumero = obterChaveNumeroArmario(numeroNormalizado);
        termoRelacionado = termosPorId[chaveNumero] || null;
        if (!termoRelacionado && chaveNumero !== '__sem_numero__') {
          termoRelacionado = termosPorId['__sem_numero__'] || null;
        }
      }
      if (termoRelacionado) {
        var statusTermoNormalizado = normalizarTextoBasico(termoRelacionado.status);
        var termoFinalizado = statusTermoNormalizado === 'finalizado';

        if (!termoFinalizado && (termoRelacionado.pdfUrl || (termoRelacionado.assinaturas && termoRelacionado.assinaturas.finalizadoEm))) {
          termoFinalizado = true;
          statusTermoNormalizado = 'finalizado';
        }

        var possuiTermo = Boolean(termoRelacionado);
        var termoEmAndamento = possuiTermo && !termoFinalizado;
        var statusDescricao = termoRelacionado.status || '';

        if (!statusDescricao) {
          statusDescricao = termoFinalizado ? 'Finalizado' : (possuiTermo ? 'Em andamento' : '');
        } else if (statusTermoNormalizado === 'finalizado') {
          statusDescricao = 'Finalizado';
        } else if (statusTermoNormalizado === 'em andamento') {
          statusDescricao = 'Em andamento';
        }

        var termoStatus = termoFinalizado ? 'finalizado' : (termoEmAndamento ? 'em andamento' : 'pendente');

        armario.termoAplicado = termoEmAndamento;
        armario.termoFinalizado = termoFinalizado;
        armario.termoStatus = termoStatus;
        armario.termoInfo = {
          id: termoRelacionado.id,
          aplicadoEm: termoRelacionado.aplicadoEm,
          finalizadoEm: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.finalizadoEm : '',
          pdfUrl: termoRelacionado.pdfUrl || '',
          responsavel: termoRelacionado.acompanhante,
          metodoFinal: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.metodoFinal : '',
          cpfFinal: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.cpfFinal : '',
          status: statusDescricao
        };
      } else {
        armario.termoAplicado = false;
        armario.termoFinalizado = false;
        armario.termoStatus = 'pendente';
        armario.termoInfo = null;
      }
    } else {
      armario.termoFinalizado = false;
      armario.termoStatus = 'pendente';
      armario.termoInfo = null;
    }

    armarios.push(armario);
  }

  if (houveAtualizacaoStatus && statusIndex > -1) {
    var statusAtualizados = dados.map(function(linha) {
      return [linha[statusIndex] || ''];
    });
    sheet.getRange(2, statusIndex + 1, totalLinhas, 1).setValues(statusAtualizados);
    limparCacheArmarios();
  }

  return armarios;
}

function cadastrarArmario(armarioData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = armarioData.tipo === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
    var sheet = ss.getSheetByName(sheetName);
    var historicoSheet = ss.getSheetByName(
      armarioData.tipo === 'acompanhante' ? 'Histórico Acompanhantes' : 'Histórico Visitantes'
    );

    if (!sheet || !historicoSheet) {
      return { success: false, error: 'Abas não encontradas' };
    }

    garantirEstruturaHistorico(historicoSheet);

    var totalLinhas = sheet.getLastRow();
    if (totalLinhas < 2) {
      return { success: false, error: 'Nenhum armário cadastrado' };
    }

  var estrutura = obterEstruturaPlanilha(sheet);
  if (sheetName === 'Visitantes') {
    estrutura = garantirColunaVisitaEstendida(sheet, estrutura);
  }
  estrutura = garantirColunaProntuario(sheet, estrutura);
  if (sheetName === 'Acompanhantes') {
    estrutura = garantirColunaObservacoesAcompanhantes(sheet, estrutura);
  }
  var totalColunas = estrutura.ultimaColuna || (sheetName === 'Visitantes' ? 14 : 12);
  var linhaPlanilha = -1;
  var linhaAtual = null;
    var idParametroBruto = armarioData.idPlanilha !== undefined && armarioData.idPlanilha !== ''
      ? armarioData.idPlanilha
      : armarioData.id;
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);

    if (totalLinhas > 1 && idParametroBruto !== undefined && idParametroBruto !== null && idParametroBruto !== '') {
      var idTexto = idParametroBruto.toString().trim();
      if (idTexto) {
        var intervaloId = sheet.getRange(2, idIndex + 1, totalLinhas - 1, 1);
        var idFinder = intervaloId.createTextFinder(idTexto).matchEntireCell(true);
        var idEncontrado = idFinder ? idFinder.findNext() : null;
        if (idEncontrado) {
          linhaPlanilha = idEncontrado.getRow();
          linhaAtual = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
        }
      }
    }

    if ((linhaPlanilha === -1 || !linhaAtual) && armarioData.numero && totalLinhas > 1) {
      var numeroInformado = armarioData.numero.toString().trim();
      if (numeroInformado) {
        var intervaloNumero = sheet.getRange(2, numeroIndex + 1, totalLinhas - 1, 1);
        var numeroFinder = intervaloNumero.createTextFinder(numeroInformado).matchEntireCell(true);
        var correspondencias = numeroFinder ? numeroFinder.findAll() : [];
        for (var j = 0; j < correspondencias.length; j++) {
          var linhaCandidata = correspondencias[j].getRow();
          var valoresLinha = sheet.getRange(linhaCandidata, 1, 1, totalColunas).getValues()[0];
          var statusLinha = normalizarTextoBasico(
            obterValorLinha(valoresLinha, estrutura, 'status', valoresLinha[statusIndex])
          );
          if (statusLinha === 'livre') {
            linhaPlanilha = linhaCandidata;
            linhaAtual = valoresLinha;
            break;
          }
        }
      }
    }

    if (linhaPlanilha === -1 || !linhaAtual) {
      return { success: false, error: 'Armário não encontrado' };
    }

    var statusAtual = normalizarTextoBasico(linhaAtual[statusIndex]);
    if (statusAtual !== 'livre') {
      return { success: false, error: 'Armário já está em uso' };
    }

    var dataHoraAtual = obterDataHoraAtualFormatada();
    var responsavelRegistro = determinarResponsavelRegistro(armarioData.usuarioResponsavel);
    var horaInicio = dataHoraAtual.horaCurta;
    var dataRegistro = dataHoraAtual.dataHoraIso;
  var volumes = parseInt(armarioData.volumes, 10);
  if (isNaN(volumes) || volumes < 0) {
    volumes = 0;
  }
  var whatsapp = armarioData.whatsapp !== null && armarioData.whatsapp !== undefined
    ? armarioData.whatsapp.toString().trim()
    : '';
  var observacoes = armarioData.observacoes ? armarioData.observacoes.toString().trim() : '';
  var numeroArmario = linhaAtual[numeroIndex];
  var unidadeAtual = obterValorLinha(linhaAtual, estrutura, 'unidade', '');
  var novaLinha = linhaAtual.slice();
    while (novaLinha.length < totalColunas) {
      novaLinha.push('');
    }
    var nomeChavesCadastro = sheetName === 'Visitantes' ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE;

    definirValorLinha(novaLinha, estrutura, 'status', 'em-uso');
    definirValorLinha(novaLinha, estrutura, nomeChavesCadastro, armarioData.nomeVisitante);
    definirValorLinha(novaLinha, estrutura, 'nome paciente', armarioData.nomePaciente);
    definirValorLinha(novaLinha, estrutura, 'prontuario', armarioData.prontuario || '');
    definirValorLinha(novaLinha, estrutura, 'leito', armarioData.leito);
    definirValorLinha(novaLinha, estrutura, 'volumes', volumes);
    definirValorLinha(novaLinha, estrutura, 'hora inicio', horaInicio);
    if (sheetName === 'Visitantes') {
      var visitaEstendida = converterParaBoolean(armarioData.visitaEstendida);
      definirValorLinha(novaLinha, estrutura, 'hora prevista', visitaEstendida ? '' : (armarioData.horaPrevista || ''));
      definirValorLinha(novaLinha, estrutura, CABECALHOS_VISITA_ESTENDIDA, visitaEstendida);
    } else {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', '');
    }
    definirValorLinha(novaLinha, estrutura, 'data registro', dataRegistro);
    definirValorLinha(novaLinha, estrutura, 'unidade', unidadeAtual);
    definirValorLinha(novaLinha, estrutura, CABECALHOS_WHATSAPP, whatsapp);
    definirValorLinha(novaLinha, estrutura, 'termo aplicado', false);
    definirValorLinha(novaLinha, estrutura, CABECALHOS_OBSERVACOES, observacoes);

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([novaLinha]);

    var historicoLastRow = historicoSheet.getLastRow();
    var ultimoHistoricoId = historicoLastRow > 1
      ? Number(historicoSheet.getRange(historicoLastRow, 1).getValue()) || 0
      : 0;
    var historicoId = ultimoHistoricoId + 1;
    var proximaLinhaHistorico = historicoLastRow + 1;

    var dataHistorico = dataHoraAtual.data;

    var historicoLinha = [
      historicoId,
      dataHistorico,
      numeroArmario,
      armarioData.nomeVisitante,
      armarioData.nomePaciente,
      armarioData.leito,
      volumes,
      horaInicio,
      '',
      'EM USO',
      armarioData.tipo,
      unidadeAtual,
      whatsapp,
      responsavelRegistro,
      observacoes
    ];

    historicoSheet.getRange(proximaLinhaHistorico, 1, 1, historicoLinha.length).setValues([historicoLinha]);

    registrarLog('CADASTRO', `Armário ${numeroArmario} cadastrado para ${armarioData.nomeVisitante}`);

    invalidarCachesArmariosRelacionados(sheetName);

    return {
      success: true,
      message: 'Armário cadastrado com sucesso',
      id: linhaAtual[idIndex]
    };

  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar armário: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function gerarProximoNumeroContingencia(sheet, estrutura, numeroIndex, statusIndex) {
  var totalLinhas = sheet.getLastRow();
  if (totalLinhas <= 1) {
    return 'Contingência-01';
  }

  var totalColunas = estrutura.ultimaColuna || sheet.getLastColumn();
  var dados = sheet.getRange(2, 1, totalLinhas - 1, totalColunas).getValues();
  var maiorSequencia = 0;

  dados.forEach(function(linha) {
    var numeroAtual = obterValorLinha(linha, estrutura, 'numero', linha[numeroIndex] || '');
    if (ehNumeroContingencia(numeroAtual)) {
      maiorSequencia = Math.max(maiorSequencia, extrairSequenciaContingencia(numeroAtual));
    }
  });

  var proximaSequencia = maiorSequencia + 1;
  return 'Contingência-' + String(proximaSequencia).padStart(2, '0');
}

function registrarContingencia(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Acompanhantes');
    var historicoSheet = ss.getSheetByName('Histórico Acompanhantes');

    if (!sheet || !historicoSheet) {
      return { success: false, error: 'Abas não encontradas' };
    }

    garantirEstruturaHistorico(historicoSheet);

    var estrutura = garantirColunaProntuario(sheet, obterEstruturaPlanilha(sheet));
    estrutura = garantirColunaObservacoesAcompanhantes(sheet, estrutura);
    estrutura = garantirColunasFotoContingencia(sheet, estrutura);
    var totalColunas = estrutura.ultimaColuna || 12;
    var totalLinhas = sheet.getLastRow();
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);

    var linhasDisponiveis = totalLinhas > 1 ? totalLinhas - 1 : 0;
    var ids = linhasDisponiveis > 0
      ? sheet.getRange(2, idIndex + 1, linhasDisponiveis, 1).getValues()
      : [];
    var numeros = linhasDisponiveis > 0
      ? sheet.getRange(2, numeroIndex + 1, linhasDisponiveis, 1).getValues()
      : [];
    var status = linhasDisponiveis > 0
      ? sheet.getRange(2, statusIndex + 1, linhasDisponiveis, 1).getValues()
      : [];

    var linhaDisponivel = -1;
    var numeroDisponivel = '';
    var maiorId = 0;
    var maiorSequencia = 0;

    for (var i = 0; i < linhasDisponiveis; i++) {
      var idLinha = Number(ids[i][0]) || 0;
      if (idLinha > maiorId) {
        maiorId = idLinha;
      }

      var numeroAtual = numeros[i][0] || '';
      var statusAtual = normalizarTextoBasico(status[i][0] || '');

      if (ehNumeroContingencia(numeroAtual)) {
        maiorSequencia = Math.max(maiorSequencia, extrairSequenciaContingencia(numeroAtual));
        if (linhaDisponivel === -1 && statusAtual === 'livre') {
          linhaDisponivel = i;
          numeroDisponivel = numeroAtual;
        }
      }
    }

    var numeroContingencia = numeroDisponivel || ('Contingência-' + String(maiorSequencia + 1).padStart(2, '0'));
    var linhaPlanilha = linhaDisponivel > -1 ? linhaDisponivel + 2 : totalLinhas + 1;
    var idGerado = linhaDisponivel > -1 ? ids[linhaDisponivel][0] : maiorId + 1;
    var dataHoraAtual = obterDataHoraAtualFormatada();
    var responsavel = determinarResponsavelRegistro(dados.usuarioResponsavel);
    var nomeChavesCadastro = CABECALHOS_NOME_ACOMPANHANTE;
    var volumes = parseInt(dados.volumes, 10);
    volumes = isNaN(volumes) || volumes < 0 ? 0 : volumes;
    var observacoes = dados.observacoes ? dados.observacoes.toString().trim() : '';

    var linhaBase = linhaDisponivel > -1
      ? sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0]
      : new Array(totalColunas).fill('');

    var fotoBase64 = (dados.fotoBase64 || '').toString().trim();
    var fotoMime = (dados.fotoMime || '').toString().trim() || 'image/jpeg';
    var fotoNome = dados.fotoNome || 'contingencia';

    if (!fotoBase64) {
      return { success: false, error: 'A foto da contingência é obrigatória.' };
    }

    var nomeArquivoFoto = gerarNomeArquivoEvidencia('contingencia', numeroContingencia);
    var fotoRegistrada;
    try {
      fotoRegistrada = salvarImagemBase64EmPasta(fotoBase64, fotoMime, nomeArquivoFoto, PASTA_DRIVE_FOTOS_ID);
    } catch (erroFoto) {
      registrarLog('ERRO', 'Falha ao salvar foto de contingência: ' + erroFoto.toString());
      return { success: false, error: 'Falha ao salvar a foto da contingência. Verifique as permissões do Drive.' };
    }
    if (!fotoRegistrada || !fotoRegistrada.url) {
      return { success: false, error: 'Falha ao salvar a foto da contingência. Verifique as permissões do Drive.' };
    }

    definirValorLinha(linhaBase, estrutura, 'id', idGerado);
    definirValorLinha(linhaBase, estrutura, 'numero', numeroContingencia);
    definirValorLinha(linhaBase, estrutura, 'status', 'contingencia');
    definirValorLinha(linhaBase, estrutura, nomeChavesCadastro, dados.nomeAcompanhante || dados.nomeVisitante || '');
    definirValorLinha(linhaBase, estrutura, 'nome paciente', dados.nomePaciente || '');
    definirValorLinha(linhaBase, estrutura, 'prontuario', dados.prontuario || '');
    definirValorLinha(linhaBase, estrutura, 'leito', dados.leito || '');
    definirValorLinha(linhaBase, estrutura, 'volumes', volumes);
    definirValorLinha(linhaBase, estrutura, 'hora inicio', dataHoraAtual.horaCurta);
    definirValorLinha(linhaBase, estrutura, 'hora prevista', '');
    definirValorLinha(linhaBase, estrutura, 'data registro', dataHoraAtual.dataHoraIso);
    definirValorLinha(linhaBase, estrutura, 'unidade', dados.unidade || '');
    definirValorLinha(linhaBase, estrutura, CABECALHOS_WHATSAPP, dados.whatsapp || '');
    definirValorLinha(linhaBase, estrutura, 'termo aplicado', false);
    definirValorLinha(linhaBase, estrutura, CABECALHOS_OBSERVACOES, observacoes);
    definirValorLinha(linhaBase, estrutura, ['foto contingencia url', 'foto contingência url'], fotoRegistrada.url || '');
    definirValorLinha(linhaBase, estrutura, ['foto contingencia id', 'foto contingência id'], fotoRegistrada.id || '');
    definirValorLinha(linhaBase, estrutura, ['foto contingencia nome', 'foto contingência nome'], fotoRegistrada.nome || fotoNome);

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([linhaBase]);

    var historicoLastRow = historicoSheet.getLastRow();
    var ultimoHistoricoId = historicoLastRow > 1
      ? Number(historicoSheet.getRange(historicoLastRow, 1).getValue()) || 0
      : 0;
    var historicoId = ultimoHistoricoId + 1;
    var proximaLinhaHistorico = historicoLastRow + 1;

    var historicoLinha = [
      historicoId,
      dataHoraAtual.data,
      numeroContingencia,
      dados.nomeAcompanhante || dados.nomeVisitante || '',
      dados.nomePaciente || '',
      dados.leito || '',
      volumes,
      dataHoraAtual.horaCurta,
      '',
      'CONTINGENCIA',
      'acompanhante',
      dados.unidade || '',
      dados.whatsapp || '',
      responsavel,
      observacoes
    ];

    historicoSheet.getRange(proximaLinhaHistorico, 1, 1, historicoLinha.length).setValues([historicoLinha]);

    invalidarCachesArmariosRelacionados('Acompanhantes');

    return {
      success: true,
      data: {
        id: idGerado,
        numero: numeroContingencia,
        status: 'contingencia',
        tipo: 'acompanhante',
        nomeVisitante: dados.nomeAcompanhante || dados.nomeVisitante || '',
        nomePaciente: dados.nomePaciente || '',
        prontuario: dados.prontuario || '',
        leito: dados.leito || '',
        volumes: volumes,
        horaInicio: dataHoraAtual.horaCurta,
        horaPrevista: '',
        dataRegistro: dataHoraAtual.dataHoraIso,
        unidade: dados.unidade || '',
        whatsapp: dados.whatsapp || '',
        observacoes: observacoes,
        visitaEstendida: false,
        termoAplicado: false,
        termoFinalizado: false,
        termoStatus: 'pendente',
        fotoContingenciaUrl: fotoRegistrada.url || '',
        fotoContingenciaId: fotoRegistrada.id || '',
        ehContingencia: true
      }
    };

  } catch (error) {
    registrarLog('ERRO', `Erro ao registrar contingência: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function registrarContingenciaTermo(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Acompanhantes');

    if (!sheet) {
      return { success: false, error: 'Aba "Acompanhantes" não encontrada' };
    }

    var estrutura = garantirColunaProntuario(sheet, obterEstruturaPlanilha(sheet));
    estrutura = garantirColunaObservacoesAcompanhantes(sheet, estrutura);
    estrutura = garantirColunasFotoContingencia(sheet, estrutura);
    var totalColunas = estrutura.ultimaColuna || 12;
    var totalLinhas = sheet.getLastRow();
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);

    var linhasPlanilha = totalLinhas > 1
      ? sheet.getRange(2, 1, totalLinhas - 1, totalColunas).getValues()
      : [];

    var linhaDisponivel = -1;
    var numeroDisponivel = '';
    var maiorId = 0;

    linhasPlanilha.forEach(function(linha, indice) {
      var idLinha = Number(linha[idIndex]) || 0;
      if (idLinha > maiorId) {
        maiorId = idLinha;
      }

      var numeroAtual = obterValorLinha(linha, estrutura, 'numero', linha[numeroIndex] || '');
      var statusAtual = normalizarTextoBasico(obterValorLinha(linha, estrutura, 'status', linha[statusIndex] || ''));

      if (linhaDisponivel === -1 && ehNumeroContingencia(numeroAtual) && statusAtual === 'livre') {
        linhaDisponivel = indice;
        numeroDisponivel = numeroAtual;
      }
    });

    var numeroContingencia = numeroDisponivel || gerarProximoNumeroContingencia(sheet, estrutura, numeroIndex, statusIndex);
    var linhaPlanilha = linhaDisponivel > -1 ? linhaDisponivel + 2 : totalLinhas + 1;
    var idGerado = linhaDisponivel > -1 ? linhasPlanilha[linhaDisponivel][idIndex] : maiorId + 1;
    var dataHoraAtual = obterDataHoraAtualFormatada();

    var linhaBase = linhaDisponivel > -1
      ? sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0]
      : new Array(totalColunas).fill('');

    definirValorLinha(linhaBase, estrutura, 'id', idGerado);
    definirValorLinha(linhaBase, estrutura, 'numero', numeroContingencia);
    definirValorLinha(linhaBase, estrutura, 'status', 'contingencia');
    definirValorLinha(linhaBase, estrutura, CABECALHOS_NOME_ACOMPANHANTE, '');
    definirValorLinha(linhaBase, estrutura, 'nome paciente', '');
    definirValorLinha(linhaBase, estrutura, 'prontuario', '');
    definirValorLinha(linhaBase, estrutura, 'leito', '');
    definirValorLinha(linhaBase, estrutura, 'volumes', 0);
    definirValorLinha(linhaBase, estrutura, 'hora inicio', dataHoraAtual.horaCurta);
    definirValorLinha(linhaBase, estrutura, 'hora prevista', '');
    definirValorLinha(linhaBase, estrutura, 'data registro', dataHoraAtual.dataHoraIso);
    definirValorLinha(linhaBase, estrutura, 'unidade', dados.unidade || '');
    definirValorLinha(linhaBase, estrutura, CABECALHOS_WHATSAPP, '');
    definirValorLinha(linhaBase, estrutura, 'termo aplicado', false);
    definirValorLinha(linhaBase, estrutura, CABECALHOS_OBSERVACOES, '');

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([linhaBase]);

    invalidarCachesArmariosRelacionados('Acompanhantes');

    return {
      success: true,
      data: {
        id: idGerado,
        numero: numeroContingencia,
        status: 'contingencia',
        tipo: 'acompanhante',
        nomeVisitante: '',
        nomePaciente: '',
        prontuario: '',
        leito: '',
        volumes: 0,
        horaInicio: dataHoraAtual.horaCurta,
        horaPrevista: '',
        dataRegistro: dataHoraAtual.dataHoraIso,
        unidade: dados.unidade || '',
        whatsapp: '',
        observacoes: '',
        visitaEstendida: false,
        termoAplicado: false,
        termoFinalizado: false,
        termoStatus: 'pendente',
        ehContingencia: true
      }
    };
  } catch (error) {
    registrarLog('ERRO', 'Erro ao registrar contingência para termo: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function atualizarHorarioVisitante(parametros) {
  try {
    var idInformado = normalizarIdentificador(parametros.id || parametros.idPlanilha);
    var numeroInformado = normalizarNumeroArmario(parametros.numero);
    var horaPrevista = parametros.horaPrevista ? parametros.horaPrevista.toString().trim() : '';
    var visitaEstendida = converterParaBoolean(parametros.visitaEstendida);
    var usuarioResponsavel = determinarResponsavelRegistro(parametros.usuarioResponsavel);

    if (!idInformado && !numeroInformado) {
      return { success: false, error: 'Armário não informado para atualização' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Visitantes');

    if (!sheet) {
      return { success: false, error: 'Aba de Visitantes não encontrada' };
    }

    var estrutura = garantirColunaVisitaEstendida(sheet, obterEstruturaPlanilha(sheet));
    var totalLinhas = sheet.getLastRow();

    if (totalLinhas <= 1) {
      return { success: false, error: 'Nenhum armário cadastrado' };
    }

    var totalColunas = estrutura.ultimaColuna || sheet.getLastColumn();
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
    var horaPrevistaIndex = obterIndiceColuna(estrutura, 'hora prevista', 8);
    var visitaEstendidaIndex = obterIndiceColuna(estrutura, CABECALHOS_VISITA_ESTENDIDA, null);

    if (horaPrevistaIndex === null || visitaEstendidaIndex === null) {
      return { success: false, error: 'Colunas de horário previstas não encontradas' };
    }

    var linhaPlanilha = -1;
    var armarioData = null;

    if (idInformado) {
      var intervaloId = sheet.getRange(2, idIndex + 1, totalLinhas - 1, 1);
      var idFinder = intervaloId.createTextFinder(idInformado).matchEntireCell(true);
      var idEncontrado = idFinder ? idFinder.findNext() : null;
      if (idEncontrado) {
        linhaPlanilha = idEncontrado.getRow();
        armarioData = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
      }
    }

    if (linhaPlanilha === -1 && numeroInformado) {
      var intervaloNumero = sheet.getRange(2, numeroIndex + 1, totalLinhas - 1, 1);
      var numeroFinder = intervaloNumero.createTextFinder(numeroInformado).matchEntireCell(true);
      var numeroEncontrado = numeroFinder ? numeroFinder.findNext() : null;
      if (numeroEncontrado) {
        linhaPlanilha = numeroEncontrado.getRow();
        armarioData = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
      }
    }

    if (linhaPlanilha === -1 || !armarioData) {
      return { success: false, error: 'Armário não encontrado' };
    }

    var statusAtual = normalizarTextoBasico(armarioData[statusIndex]);
    var statusPermitidos = ['em-uso', 'proximo', 'vencido'];
    if (statusPermitidos.indexOf(statusAtual) === -1) {
      return { success: false, error: 'Armário não está em uso para atualização de horário' };
    }

    var novaHoraPrevista = visitaEstendida ? '' : horaPrevista;
    armarioData[horaPrevistaIndex] = novaHoraPrevista;
    armarioData[visitaEstendidaIndex] = visitaEstendida;

    var dataRegistroIndex = obterIndiceColuna(estrutura, 'data registro', 9);
    var dataRegistroValor = (dataRegistroIndex !== null && dataRegistroIndex !== undefined && dataRegistroIndex < armarioData.length)
      ? armarioData[dataRegistroIndex]
      : null;
    var novoStatus = calcularStatusAutomaticoVisitante('em-uso', dataRegistroValor, novaHoraPrevista, new Date());
    if (novoStatus) {
      armarioData[statusIndex] = novoStatus;
    }

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([armarioData]);

    invalidarCachesArmariosRelacionados('Visitantes');

    var numeroArmario = armarioData[numeroIndex] || numeroInformado || '';
    var detalheHorario = visitaEstendida ? 'visita estendida' : (novaHoraPrevista || '-');
    registrarLog('ATUALIZACAO', `Horário do armário ${numeroArmario} atualizado para ${detalheHorario} por ${usuarioResponsavel}`);

    return {
      success: true,
      horaPrevista: novaHoraPrevista,
      visitaEstendida: visitaEstendida
    };
  } catch (error) {
    registrarLog('ERRO', `Erro ao atualizar horário do armário: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function atualizarDadosArmario(parametros) {
  try {
    var tipoNormalizado = normalizarTextoBasico(parametros.tipo);
    if (!tipoNormalizado) {
      tipoNormalizado = 'visitante';
    }

    var sheetName = tipoNormalizado === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: false, error: 'Aba de armários não encontrada' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    if (sheetName === 'Visitantes') {
      estrutura = garantirColunaVisitaEstendida(sheet, estrutura);
    }
    estrutura = garantirColunaProntuario(sheet, estrutura);
    if (sheetName === 'Acompanhantes') {
      estrutura = garantirColunaObservacoesAcompanhantes(sheet, estrutura);
    }

    var totalLinhas = sheet.getLastRow();
    if (totalLinhas <= 1) {
      return { success: false, error: 'Nenhum armário cadastrado' };
    }

    var totalColunas = estrutura.ultimaColuna || sheet.getLastColumn();
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
    var nomeIndex = obterIndiceColuna(estrutura, sheetName === 'Visitantes' ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE, 3);
    var pacienteIndex = obterIndiceColuna(estrutura, 'nome paciente', 4);
    var leitoIndex = obterIndiceColuna(estrutura, 'leito', 5);
    var volumesIndex = obterIndiceColuna(estrutura, 'volumes', 6);
    var horaInicioIndex = obterIndiceColuna(estrutura, 'hora inicio', 7);
    var horaPrevistaIndex = sheetName === 'Visitantes' ? obterIndiceColuna(estrutura, 'hora prevista', 8) : -1;
    var dataRegistroIndex = sheetName === 'Visitantes' ? obterIndiceColuna(estrutura, 'data registro', 9) : obterIndiceColuna(estrutura, 'data registro', 8);
    var unidadeIndex = obterIndiceColuna(estrutura, 'unidade', sheetName === 'Visitantes' ? 10 : 10);
    var termoIndex = obterIndiceColuna(estrutura, 'termo aplicado', sheetName === 'Visitantes' ? 11 : 11);
    var whatsappIndex = obterIndiceColuna(estrutura, CABECALHOS_WHATSAPP, sheetName === 'Visitantes' ? 12 : 9);
    var observacoesIndex = obterIndiceColuna(estrutura, CABECALHOS_OBSERVACOES, null);
    var visitaEstendidaIndex = sheetName === 'Visitantes' ? obterIndiceColuna(estrutura, CABECALHOS_VISITA_ESTENDIDA, null) : -1;

    var idInformado = normalizarIdentificador(parametros.id || parametros.idPlanilha);
    var numeroInformado = normalizarNumeroArmario(parametros.numero);

    if (!idInformado && !numeroInformado) {
      return { success: false, error: 'Armário não informado para atualização' };
    }

    var linhaPlanilha = -1;
    var armarioData = null;

    if (idInformado) {
      var intervaloId = sheet.getRange(2, idIndex + 1, totalLinhas - 1, 1);
      var idFinder = intervaloId.createTextFinder(idInformado).matchEntireCell(true);
      var idEncontrado = idFinder ? idFinder.findNext() : null;
      if (idEncontrado) {
        linhaPlanilha = idEncontrado.getRow();
        armarioData = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
      }
    }

    if (linhaPlanilha === -1 && numeroInformado) {
      var intervaloNumero = sheet.getRange(2, numeroIndex + 1, totalLinhas - 1, 1);
      var numeroFinder = intervaloNumero.createTextFinder(numeroInformado).matchEntireCell(true);
      var numeroEncontrado = numeroFinder ? numeroFinder.findNext() : null;
      if (numeroEncontrado) {
        linhaPlanilha = numeroEncontrado.getRow();
        armarioData = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
      }
    }

    if (linhaPlanilha === -1 || !armarioData) {
      return { success: false, error: 'Armário não encontrado' };
    }

    var statusAtual = normalizarTextoBasico(armarioData[statusIndex]);
    if (statusAtual === 'livre') {
      return { success: false, error: 'Armário está livre. Cadastre um uso antes de editar.' };
    }

    var nomeVisitante = parametros.nomeVisitante ? parametros.nomeVisitante.toString().trim() : '';
    var nomePaciente = parametros.nomePaciente ? parametros.nomePaciente.toString().trim() : '';
    var leito = parametros.leito ? parametros.leito.toString().trim() : '';
    var whatsapp = parametros.whatsapp ? parametros.whatsapp.toString().trim() : '';
    var volumes = parseInt(parametros.volumes, 10);
    if (isNaN(volumes) || volumes < 0) {
      volumes = 0;
    }
    var prontuario = normalizarIdentificador(parametros.prontuario);
    var observacoes = parametros.observacoes ? parametros.observacoes.toString().trim() : '';
    var visitaEstendida = sheetName === 'Visitantes' ? converterParaBoolean(parametros.visitaEstendida) : false;
    var horaPrevista = sheetName === 'Visitantes' ? (visitaEstendida ? '' : (parametros.horaPrevista ? parametros.horaPrevista.toString().trim() : '')) : '';

    definirValorLinha(armarioData, estrutura, sheetName === 'Visitantes' ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE, nomeVisitante);
    definirValorLinha(armarioData, estrutura, 'nome paciente', nomePaciente);
    definirValorLinha(armarioData, estrutura, 'prontuario', prontuario);
    definirValorLinha(armarioData, estrutura, 'leito', leito);
    definirValorLinha(armarioData, estrutura, 'volumes', volumes);
    definirValorLinha(armarioData, estrutura, CABECALHOS_WHATSAPP, whatsapp);
    definirValorLinha(armarioData, estrutura, CABECALHOS_OBSERVACOES, observacoes);

    if (sheetName === 'Visitantes') {
      if (horaPrevistaIndex === null || horaPrevistaIndex === undefined || visitaEstendidaIndex === null || visitaEstendidaIndex === undefined) {
        return { success: false, error: 'Colunas de horário previstas não encontradas' };
      }
      armarioData[horaPrevistaIndex] = horaPrevista;
      armarioData[visitaEstendidaIndex] = visitaEstendida;

      var dataRegistroValor = (dataRegistroIndex !== null && dataRegistroIndex !== undefined && dataRegistroIndex < armarioData.length)
        ? armarioData[dataRegistroIndex]
        : null;
      var novoStatus = calcularStatusAutomaticoVisitante(statusAtual, dataRegistroValor, horaPrevista, new Date());
      if (novoStatus) {
        armarioData[statusIndex] = novoStatus;
        statusAtual = novoStatus;
      }
    }

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([armarioData]);

    var numeroArmario = obterValorLinha(armarioData, estrutura, 'numero', armarioData[numeroIndex] || '');
    var unidadeAtual = unidadeIndex !== null && unidadeIndex !== undefined ? (armarioData[unidadeIndex] || '') : '';
    var horaInicioFormatada = formatarHorarioPlanilha(armarioData[horaInicioIndex]);
    var horaPrevistaFormatada = horaPrevistaIndex > -1 ? formatarHorarioPlanilha(armarioData[horaPrevistaIndex]) : '';
    var dataRegistroValorFormatado = dataRegistroIndex !== null && dataRegistroIndex !== undefined ? armarioData[dataRegistroIndex] : '';

    registrarLog('ATUALIZACAO', 'Dados do armário ' + numeroArmario + ' (' + sheetName + ') atualizados');
    invalidarCachesArmariosRelacionados(sheetName);

    return {
      success: true,
      data: {
        id: armarioData[idIndex],
        numero: numeroArmario,
        status: statusAtual,
        tipo: tipoNormalizado,
        nomeVisitante: nomeVisitante,
        nomePaciente: nomePaciente,
        prontuario: prontuario,
        leito: leito,
        volumes: volumes,
        whatsapp: whatsapp,
        observacoes: observacoes,
        horaInicio: horaInicioFormatada,
        horaPrevista: horaPrevistaFormatada,
        visitaEstendida: visitaEstendida,
        dataRegistro: dataRegistroValorFormatado,
        unidade: unidadeAtual,
        termoAplicado: termoIndex !== null && termoIndex !== undefined ? converterParaBoolean(armarioData[termoIndex]) : false
      }
    };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao atualizar dados do armário: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function liberarArmario(id, tipo, numero, usuarioResponsavel) {
  try {
    var tipoNormalizado = normalizarTextoBasico(tipo);
    var ehAcompanhante = tipoNormalizado === 'acompanhante';
    var idComparacao = id !== null && id !== undefined ? id.toString().trim() : '';
    var numeroInformado = normalizarNumeroArmario(numero);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = ehAcompanhante ? 'Acompanhantes' : 'Visitantes';
    var sheet = ss.getSheetByName(sheetName);
    var historicoSheet = ss.getSheetByName(
      ehAcompanhante ? 'Histórico Acompanhantes' : 'Histórico Visitantes'
    );

    if (!sheet || !historicoSheet) {
      return { success: false, error: 'Abas não encontradas' };
    }

    garantirEstruturaHistorico(historicoSheet);

    var responsavelRegistro = determinarResponsavelRegistro(usuarioResponsavel);

    // Encontrar o armário na aba atual
    var estrutura = obterEstruturaPlanilha(sheet);
    if (!ehAcompanhante) {
      estrutura = garantirColunaVisitaEstendida(sheet, estrutura);
    }
    estrutura = garantirColunaProntuario(sheet, estrutura);
    if (ehAcompanhante) {
      estrutura = garantirColunaObservacoesAcompanhantes(sheet, estrutura);
    }
    var totalColunas = estrutura.ultimaColuna || (sheetName === 'Visitantes' ? 14 : 12);
    var totalLinhas = sheet.getLastRow();
    if (totalLinhas <= 1) {
      return { success: false, error: 'Nenhum armário cadastrado' };
    }

    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var linhaPlanilha = -1;
    var armarioData = null;

    if (totalLinhas > 1) {
      var indiceArmarios = obterIndiceArmarios(sheetName);
      var chaveNumeroInformado = numeroInformado ? obterChaveNumeroArmario(numeroInformado) : '';

      if (idComparacao && indiceArmarios.porId[idComparacao]) {
        linhaPlanilha = indiceArmarios.porId[idComparacao];
      }

      if (linhaPlanilha === -1 && chaveNumeroInformado && indiceArmarios.porNumero[chaveNumeroInformado]) {
        linhaPlanilha = indiceArmarios.porNumero[chaveNumeroInformado];
      }

      if (linhaPlanilha !== -1) {
        armarioData = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
        if (!validarLinhaArmarioEncontrada(armarioData, estrutura, idComparacao, numeroInformado)) {
          linhaPlanilha = -1;
          armarioData = null;
        }
      }

      if (linhaPlanilha === -1 || !armarioData) {
        var indiceAtualizado = obterIndiceArmarios(sheetName, true);

        if (idComparacao && indiceAtualizado.porId[idComparacao]) {
          linhaPlanilha = indiceAtualizado.porId[idComparacao];
        }

        if (linhaPlanilha === -1 && chaveNumeroInformado && indiceAtualizado.porNumero[chaveNumeroInformado]) {
          linhaPlanilha = indiceAtualizado.porNumero[chaveNumeroInformado];
        }

        if (linhaPlanilha !== -1) {
          armarioData = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
          if (!validarLinhaArmarioEncontrada(armarioData, estrutura, idComparacao, numeroInformado)) {
            linhaPlanilha = -1;
            armarioData = null;
          }
        }
      }
    }

    if (linhaPlanilha === -1 || !armarioData) {
      return { success: false, error: 'Armário não encontrado' };
    }

    var statusPadrao = (statusIndex !== null && statusIndex !== undefined && statusIndex < armarioData.length)
      ? armarioData[statusIndex]
      : '';
    var statusAtual = normalizarTextoBasico(
      obterValorLinha(armarioData, estrutura, 'status', statusPadrao)
    );
    if (statusAtual === 'livre') {
      return { success: false, error: 'Armário já está livre' };
    }

    // Limpar dados do armário (deixar apenas número e status livre)
    var unidadeAtual = obterValorLinha(armarioData, estrutura, 'unidade', '');
    var novaLinha = armarioData.slice();
    while (novaLinha.length < totalColunas) {
      novaLinha.push('');
    }

    var nomeColuna = sheetName === 'Visitantes' ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE;
    var definirCampoComFallback = function(chave, valor, indiceFallback) {
      definirValorLinha(novaLinha, estrutura, chave, valor);
      var indiceExiste = obterIndiceColuna(estrutura, chave, null);
      var indiceDestino = indiceExiste;
      if ((indiceDestino === null || indiceDestino === undefined) &&
          indiceFallback !== null && indiceFallback !== undefined &&
          indiceFallback >= 0) {
        indiceDestino = indiceFallback;
      }
      if (indiceDestino !== null && indiceDestino !== undefined && indiceDestino >= 0) {
        while (novaLinha.length <= indiceDestino) {
          novaLinha.push('');
        }
        novaLinha[indiceDestino] = valor;
      }
    };

    definirCampoComFallback('status', 'livre', statusIndex);
    definirValorLinha(novaLinha, estrutura, nomeColuna, '');
    definirValorLinha(novaLinha, estrutura, 'nome paciente', '');
    definirValorLinha(novaLinha, estrutura, 'leito', '');
    definirValorLinha(novaLinha, estrutura, 'volumes', '');
    definirValorLinha(novaLinha, estrutura, 'hora inicio', '');
    if (sheetName === 'Visitantes') {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', '');
      definirValorLinha(novaLinha, estrutura, CABECALHOS_VISITA_ESTENDIDA, false);
    }
    var dataHoraAtual = obterDataHoraAtualFormatada();
    definirValorLinha(novaLinha, estrutura, 'data registro', dataHoraAtual.dataHoraIso);
    definirValorLinha(novaLinha, estrutura, CABECALHOS_WHATSAPP, '');
    definirValorLinha(novaLinha, estrutura, 'unidade', unidadeAtual);
    definirValorLinha(novaLinha, estrutura, 'termo aplicado', false);
    definirValorLinha(novaLinha, estrutura, CABECALHOS_OBSERVACOES, '');

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([novaLinha]);

    // Atualizar histórico - encontrar a entrada mais recente deste armário
    var historicoLastRow = historicoSheet.getLastRow();
    var numeroArmario = obterValorLinha(armarioData, estrutura, 'numero', armarioData[numeroIndex]);
    numeroArmario = numeroArmario ? numeroArmario.toString().trim() : '';
    if (!numeroArmario) {
      numeroArmario = numeroInformado;
    }

    if (historicoLastRow > 1 && numeroArmario) {
      var totalLinhasHistorico = historicoLastRow - 1;
      var numerosHistorico = historicoSheet.getRange(2, 3, totalLinhasHistorico, 1).getValues().flat();
      var statusHistorico = historicoSheet.getRange(2, 10, totalLinhasHistorico, 1).getValues().flat();
      var linhaHistorico = -1;

      for (var i = numerosHistorico.length - 1; i >= 0; i--) {
        if (numerosHistorico[i] === numeroArmario && statusHistorico[i] === 'EM USO') {
          linhaHistorico = i + 2;
          break;
        }
      }

      if (linhaHistorico !== -1) {
        var horaFim = dataHoraAtual.horaCurta;
        var intervaloAtualizacao = historicoSheet.getRange(linhaHistorico, 9, 1, 6);
        var dadosAtualizacao = intervaloAtualizacao.getValues()[0];
        var valoresAtualizados = [horaFim, 'FINALIZADO', responsavelRegistro || dadosAtualizacao[5]];

        dadosAtualizacao[0] = valoresAtualizados[0];
        dadosAtualizacao[1] = valoresAtualizados[1];
        dadosAtualizacao[5] = valoresAtualizados[2];

        intervaloAtualizacao.setValues([dadosAtualizacao]);
      }
    }

    registrarLog('LIBERAÇÃO', `Armário ${numeroArmario} liberado`);

    SpreadsheetApp.flush();
    invalidarCachesArmariosRelacionados(sheetName);

    return { success: true, message: 'Armário liberado com sucesso' };

  } catch (error) {
    registrarLog('ERRO', `Erro ao liberar armário: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Funções para Usuários
function getUsuarios() {
  return executarComCache(montarChaveCache('usuarios'), CACHE_TTL_PADRAO, function() {
    try {
      var permissao = validarPermissaoAdmin({ usuarioId: usuarioContextoRequisicaoId });
      if (!permissao.ok) {
        return { success: false, error: permissao.error };
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Usuários');

      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [] };
      }

      var estrutura = obterEstruturaPlanilha(sheet);
      var totalColunas = estrutura.ultimaColuna || 10;
      var mapasUnidades = obterMapasUnidades();
      var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
      var usuarios = [];

      dados.forEach(function(linha) {
        var id = obterValorLinha(linha, estrutura, 'id', linha[0]);
        if (!id && id !== 0) {
          return;
        }

        var perfilValor = obterValorLinha(linha, estrutura, 'perfil', 'usuario');
        var perfil = perfilValor ? perfilValor.toString().trim().toLowerCase() : 'usuario';
        var unidadesBrutas = obterValorLinhaFlexivel(linha, estrutura, ['unidades', 'unidade', 'acesso unidades'], '');
        var unidades = resolverIdsUnidadesArmazenadas(unidadesBrutas, mapasUnidades);
        var unidadesUnicas = [];
        unidades.forEach(function(unidade) {
          if (unidadesUnicas.indexOf(unidade) === -1) {
            unidadesUnicas.push(unidade);
          }
        });

        usuarios.push({
          id: id,
          nome: obterValorLinha(linha, estrutura, 'nome', ''),
          email: obterValorLinha(linha, estrutura, 'email', ''),
          perfil: perfil,
          acessoVisitantes: converterParaBoolean(obterValorLinha(linha, estrutura, 'acesso visitantes', false)),
          acessoAcompanhantes: converterParaBoolean(obterValorLinha(linha, estrutura, 'acesso acompanhantes', false)),
          dataCadastro: obterValorLinha(linha, estrutura, 'data cadastro', ''),
          status: obterValorLinha(linha, estrutura, 'status', ''),
          unidades: unidadesUnicas
        });
      });

      return { success: true, data: usuarios };

    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar usuários: ${error.toString()}`);
      return { success: false, error: 'Erro ao processar a solicitação.' };
    }
  });
}

function cadastrarUsuario(dados) {
  try {
    var permissao = validarPermissaoAdmin(dados);
    if (!permissao.ok) {
      return { success: false, error: permissao.error };
    }

    var nome = (dados.nome || '').toString().trim();
    var email = (dados.email || '').toString().trim();
    var perfil = (dados.perfil || '').toString().trim().toLowerCase();
    var senha = (dados.senha || '').toString().trim();
    var unidadesLista = normalizarListaUnidadesParametro(dados.unidades);
    var unidadesUnicas = [];
    var incluiTodas = false;

    unidadesLista.forEach(function(unidade) {
      var chave = unidade.toString().trim();
      if (!chave) {
        return;
      }
      if (normalizarTextoBasico(chave) === 'all') {
        incluiTodas = true;
        return;
      }
      if (unidadesUnicas.indexOf(chave) === -1) {
        unidadesUnicas.push(chave);
      }
    });

    if (incluiTodas || (perfil === 'admin' && unidadesUnicas.length === 0)) {
      unidadesUnicas = ['all'];
    }

    if (!nome || !email || !perfil) {
      return { success: false, error: 'Nome, matrícula e perfil são obrigatórios' };
    }

    if (!senha) {
      return { success: false, error: 'Informe uma senha para o usuário' };
    }

    if (unidadesUnicas.length === 0 && perfil !== 'admin') {
      return { success: false, error: 'Informe ao menos uma unidade de acesso' };
    }

    var mapasUnidades = obterMapasUnidades();
    var unidadesFormatadas = formatarUnidadesParaRegistro(unidadesUnicas, mapasUnidades);
    var unidadesTexto = unidadesFormatadas.join('; ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');

    if (!sheet) {
      return { success: false, error: 'Aba de usuários não encontrada' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var ultimaLinha = sheet.getLastRow();
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var proximoId = 1;

    if (ultimaLinha >= 2) {
      var idRange = sheet.getRange(2, idIndex + 1, ultimaLinha - 1, 1).getValues().flat();
      var idsNumericos = idRange.map(function(valor) {
        var numero = parseInt(valor, 10);
        return isNaN(numero) ? null : numero;
      }).filter(function(valor) {
        return valor !== null;
      });
      var ultimoId = idsNumericos.length > 0 ? Math.max.apply(null, idsNumericos) : 0;
      proximoId = ultimoId + 1;
    }

    var acessoVisitantes = converterParaBoolean(dados.acessoVisitantes);
    var acessoAcompanhantes = converterParaBoolean(dados.acessoAcompanhantes);
    var dataCadastro = obterDataHoraAtualFormatada().dataHoraIso;

    var novaLinha = new Array(totalColunas);
    for (var i = 0; i < totalColunas; i++) {
      novaLinha[i] = '';
    }

    definirValorLinha(novaLinha, estrutura, 'id', proximoId);
    definirValorLinha(novaLinha, estrutura, 'nome', nome);
    definirValorLinha(novaLinha, estrutura, 'email', email);
    definirValorLinha(novaLinha, estrutura, 'perfil', perfil);
    definirValorLinha(novaLinha, estrutura, 'acesso visitantes', acessoVisitantes);
    definirValorLinha(novaLinha, estrutura, 'acesso acompanhantes', acessoAcompanhantes);
    definirValorLinha(novaLinha, estrutura, 'data cadastro', dataCadastro);
    definirValorLinha(novaLinha, estrutura, 'status', 'ativo');
    definirValorLinha(novaLinha, estrutura, 'senha', criarHashSenha(senha));
    if (!definirValorLinhaFlexivel(novaLinha, estrutura, ['unidades', 'unidade', 'acesso unidades'], unidadesTexto)) {
      definirValorLinha(novaLinha, estrutura, 'unidades', unidadesTexto);
    }
    // Novo usuário deve trocar a senha definida pelo admin no primeiro acesso.
    definirValorLinhaFlexivel(novaLinha, estrutura, ['precisa trocar senha', 'precisatrocarsenha'], 'true');

    sheet.getRange(ultimaLinha + 1, 1, 1, totalColunas).setValues([novaLinha]);

    registrarLog('CADASTRO USUARIO', `Usuário ${nome} cadastrado`);

    limparCacheUsuarios();

    return {
      success: true,
      message: 'Usuário cadastrado com sucesso',
      id: proximoId,
      usuario: {
        id: proximoId,
        nome: nome,
        email: email,
        perfil: perfil,
        acessoVisitantes: acessoVisitantes,
        acessoAcompanhantes: acessoAcompanhantes,
        dataCadastro: dataCadastro,
        status: 'ativo',
        unidades: unidadesUnicas.slice()
      }
    };

  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar usuário: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function atualizarUsuario(dados) {
  try {
    var permissao = validarPermissaoAdmin(dados);
    if (!permissao.ok) {
      return { success: false, error: permissao.error };
    }

    var id = parseInt(dados.id, 10);
    if (!id) {
      return { success: false, error: 'ID do usuário inválido' };
    }

    var nome = (dados.nome || '').toString().trim();
    var email = (dados.email || '').toString().trim();
    var perfil = (dados.perfil || '').toString().trim().toLowerCase();
    var senha = (dados.senha || '').toString().trim();
    var status = (dados.status || '').toString().trim().toLowerCase();
    var acessoVisitantes = converterParaBoolean(dados.acessoVisitantes);
    var acessoAcompanhantes = converterParaBoolean(dados.acessoAcompanhantes);
    var unidadesLista = normalizarListaUnidadesParametro(dados.unidades);
    var unidadesUnicas = [];
    var incluiTodas = false;

    unidadesLista.forEach(function(unidade) {
      var chave = unidade.toString().trim();
      if (!chave) {
        return;
      }
      if (normalizarTextoBasico(chave) === 'all') {
        incluiTodas = true;
        return;
      }
      if (unidadesUnicas.indexOf(chave) === -1) {
        unidadesUnicas.push(chave);
      }
    });

    if (incluiTodas || (perfil === 'admin' && unidadesUnicas.length === 0)) {
      unidadesUnicas = ['all'];
    }

    if (!nome || !email || !perfil) {
      return { success: false, error: 'Nome, matrícula e perfil são obrigatórios' };
    }

    if (!senha) {
      return { success: false, error: 'Informe a senha do usuário' };
    }

    if (unidadesUnicas.length === 0 && perfil !== 'admin') {
      return { success: false, error: 'Informe ao menos uma unidade de acesso' };
    }

    var mapasUnidades = obterMapasUnidades();
    var unidadesFormatadas = formatarUnidadesParaRegistro(unidadesUnicas, mapasUnidades);
    var unidadesTexto = unidadesFormatadas.join('; ');

    if (!status || ['ativo', 'inativo'].indexOf(status) === -1) {
      status = 'ativo';
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var ultimaLinha = sheet.getLastRow();
    var faixa = sheet.getRange(2, 1, ultimaLinha - 1, totalColunas);
    var valores = faixa.getValues();
    var encontrado = false;

    for (var i = 0; i < valores.length; i++) {
      var valorId = valores[i][idIndex];
      if (parseInt(valorId, 10) === id) {
        definirValorLinha(valores[i], estrutura, 'nome', nome);
        definirValorLinha(valores[i], estrutura, 'email', email);
        definirValorLinha(valores[i], estrutura, 'perfil', perfil);
        definirValorLinha(valores[i], estrutura, 'acesso visitantes', acessoVisitantes);
        definirValorLinha(valores[i], estrutura, 'acesso acompanhantes', acessoAcompanhantes);
        definirValorLinha(valores[i], estrutura, 'status', status);
        definirValorLinha(valores[i], estrutura, 'senha', criarHashSenha(senha));
        if (!definirValorLinhaFlexivel(valores[i], estrutura, ['unidades', 'unidade', 'acesso unidades'], unidadesTexto)) {
          definirValorLinha(valores[i], estrutura, 'unidades', unidadesTexto);
        }
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    faixa.setValues(valores);

    registrarLog('ATUALIZAR USUARIO', 'Usuário ' + nome + ' atualizado');

    limparCacheUsuarios();

    return {
      success: true,
      message: 'Usuário atualizado com sucesso',
      usuario: {
        id: id,
        nome: nome,
        email: email,
        perfil: perfil,
        acessoVisitantes: acessoVisitantes,
        acessoAcompanhantes: acessoAcompanhantes,
        status: status,
        unidades: unidadesUnicas.slice()
      }
    };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao atualizar usuário: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function excluirUsuario(dados) {
  try {
    var permissao = validarPermissaoAdmin(dados);
    if (!permissao.ok) {
      return { success: false, error: permissao.error };
    }

    var id = parseInt(dados.id, 10);
    if (!id) {
      return { success: false, error: 'ID inválido' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var ultimaLinha = sheet.getLastRow();
    var valores = sheet.getRange(2, 1, ultimaLinha - 1, totalColunas).getValues();
    var linhaExcluir = -1;
    var nomeUsuario = '';

    for (var i = 0; i < valores.length; i++) {
      var valorId = valores[i][idIndex];
      if (parseInt(valorId, 10) === id) {
        linhaExcluir = i + 2;
        nomeUsuario = obterValorLinha(valores[i], estrutura, 'nome', '');
        break;
      }
    }

    if (linhaExcluir === -1) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    sheet.deleteRow(linhaExcluir);
    registrarLog('EXCLUIR USUARIO', 'Usuário ' + nomeUsuario + ' removido');

    limparCacheUsuarios();

    return { success: true };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao excluir usuário: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function autenticarUsuario(dados) {
  try {
    var login = (dados.usuario || dados.matricula || dados.email || dados.login || '').toString().trim();
    var senhaInformada = (dados.senha || '').toString().trim();

    if (!login || !senhaInformada) {
      return { success: false, error: 'Informe usuário e senha' };
    }

    if (verificarLoginBloqueado(login)) {
      return { success: false, error: 'Usuário bloqueado temporariamente. Tente novamente mais tarde.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Nenhum usuário cadastrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var mapasUnidades = obterMapasUnidades();
    var dadosUsuarios = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
    var alvoNormalizado = normalizarTextoBasico(login);
    var linhaUsuario = null;
    var indiceLinhaUsuario = -1;

    for (var i = 0; i < dadosUsuarios.length; i++) {
      var linha = dadosUsuarios[i];
      var identificadores = [];

      ['usuario', 'nome', 'matricula', 'email'].forEach(function(chave) {
        var valor = obterValorLinha(linha, estrutura, chave, '');
        if (valor !== null && valor !== undefined) {
          var texto = valor.toString().trim();
          if (texto) {
            identificadores.push(texto);
          }
        }
      });

      var encontrou = identificadores.some(function(valor) {
        return normalizarTextoBasico(valor) === alvoNormalizado;
      });

      if (encontrou) {
        linhaUsuario = linha;
        indiceLinhaUsuario = i;
        break;
      }
    }

    if (!linhaUsuario) {
      registrarFalhaLogin(login);
      return { success: false, error: 'Usuário não encontrado' };
    }

    var status = obterValorLinha(linhaUsuario, estrutura, 'status', '');
    if (normalizarTextoBasico(status) !== 'ativo') {
      registrarFalhaLogin(login);
      return { success: false, error: 'Usuário inativo' };
    }

    var senhaArmazenada = obterValorLinha(linhaUsuario, estrutura, 'senha', '');
    if (senhaArmazenada !== null && senhaArmazenada !== undefined) {
      senhaArmazenada = senhaArmazenada.toString().trim();
    } else {
      senhaArmazenada = '';
    }

    if (!validarSenha(senhaInformada, senhaArmazenada)) {
      registrarFalhaLogin(login);
      return { success: false, error: 'Senha incorreta' };
    }

    if (!senhaEhHashValido(senhaArmazenada)) {
      var indiceSenha = obterIndiceColuna(estrutura, 'senha', null);
      if (indiceSenha !== null && indiceSenha !== undefined) {
        var linhaAtualizacao = indiceLinhaUsuario + 2;
        sheet.getRange(linhaAtualizacao, indiceSenha + 1).setValue(criarHashSenha(senhaInformada));
      }
    }

    limparTentativasLogin(login);

    var unidadesTexto = obterValorLinhaFlexivel(linhaUsuario, estrutura, ['unidades', 'unidade', 'acesso unidades'], '');
    var unidadesLista = resolverIdsUnidadesArmazenadas(unidadesTexto, mapasUnidades);

    var precisaTrocarSenha = converterParaBoolean(
      obterValorLinhaFlexivel(linhaUsuario, estrutura, ['precisa trocar senha', 'precisatrocarsenha'], false)
    );

    var perfilUsuario = obterValorLinha(linhaUsuario, estrutura, 'perfil', '');
    var idUsuario = obterValorLinha(linhaUsuario, estrutura, 'id', '');

    var usuarioEncontrado = {
      id: idUsuario,
      nome: obterValorLinha(linhaUsuario, estrutura, 'nome', ''),
      email: obterValorLinha(linhaUsuario, estrutura, 'email', ''),
      usuario: obterValorLinha(linhaUsuario, estrutura, 'usuario', login) || login,
      matricula: obterValorLinha(linhaUsuario, estrutura, 'matricula', ''),
      perfil: perfilUsuario,
      acessoVisitantes: converterParaBoolean(obterValorLinha(linhaUsuario, estrutura, 'acesso visitantes', false)),
      acessoAcompanhantes: converterParaBoolean(obterValorLinha(linhaUsuario, estrutura, 'acesso acompanhantes', false)),
      unidades: unidadesLista,
      status: status,
      precisaTrocarSenha: precisaTrocarSenha
    };

    var token = criarSessao(idUsuario, perfilUsuario, precisaTrocarSenha);

    registrarLog('LOGIN', 'Usuário ' + usuarioEncontrado.nome + ' autenticado');

    return {
      success: true,
      usuario: usuarioEncontrado,
      token: token,
      precisaTrocarSenha: precisaTrocarSenha,
      linha: indiceLinhaUsuario + 2
    };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao autenticar usuário: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function trocarSenha(dados) {
  try {
    var usuarioId = parseInt(sessaoUsuarioIdAtual, 10);
    if (!usuarioId) {
      return { success: false, error: 'Sessão inválida. Faça login novamente.' };
    }

    var senhaAtual = (dados.senhaAtual || '').toString();
    var novaSenha = (dados.novaSenha || '').toString();

    if (!novaSenha || novaSenha.trim().length < 6) {
      return { success: false, error: 'A nova senha deve ter ao menos 6 caracteres.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Usuário não encontrado.' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 11;
    var valores = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
    var usuario = obterUsuarioPorId(usuarioId, estrutura, valores);
    if (!usuario) {
      return { success: false, error: 'Usuário não encontrado.' };
    }

    var senhaArmazenada = (obterValorLinha(usuario.linha, estrutura, 'senha', '') || '').toString().trim();
    if (!validarSenha(senhaAtual, senhaArmazenada)) {
      return { success: false, error: 'Senha atual incorreta.' };
    }

    var linhaPlanilha = usuario.indice + 2;
    var indiceSenha = obterIndiceColuna(estrutura, 'senha', null);
    if (indiceSenha === null || indiceSenha === undefined) {
      return { success: false, error: 'Não foi possível atualizar a senha.' };
    }
    sheet.getRange(linhaPlanilha, indiceSenha + 1).setValue(criarHashSenha(novaSenha.trim()));

    var indiceFlag = obterIndiceColuna(estrutura, ['precisa trocar senha', 'precisatrocarsenha'], null);
    if (indiceFlag !== null && indiceFlag !== undefined) {
      sheet.getRange(linhaPlanilha, indiceFlag + 1).setValue('false');
    }

    atualizarFlagSenhaSessoes(usuarioId, false);
    sessaoPrecisaTrocarSenha = false;

    registrarLog('TROCA SENHA', 'Usuário id ' + usuarioId + ' trocou a senha');
    return { success: true };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao trocar senha: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Funções para Histórico
function getHistorico(tipo) {
  var tipoNormalizado = normalizarTextoBasico(tipo) === 'acompanhante' ? 'acompanhante' : 'visitante';
  var chaveCache = montarChaveCache('historico', tipoNormalizado);

  return executarComCache(chaveCache, CACHE_TTL_HISTORICO, function() {
    try {
      if (tipoNormalizado === 'acompanhante') {
        return montarHistoricoTermosResponsabilidade();
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetName = tipoNormalizado === 'acompanhante' ? 'Histórico Acompanhantes' : 'Histórico Visitantes';
      var sheet = ss.getSheetByName(sheetName);

      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [] };
      }

      garantirEstruturaHistorico(sheet);

      var totalLinhasDados = sheet.getLastRow() - 1;
      var totalColunasDados = Math.max(sheet.getLastColumn(), 15);
      var data = sheet.getRange(2, 1, totalLinhasDados, totalColunasDados).getValues();
      var dadosBackup = obterLinhasSheetAtualEBackups(sheetName, {
        tipoArquivo: 'geral',
        incluirPlanilhaAtual: false,
        colunasMinimas: 15
      });
      if (dadosBackup.length) {
        data = data.concat(dadosBackup);
      }
      var historico = [];

      data.forEach(function(row) {
        if (row[0]) {
          historico.push({
            id: row[0],
            data: formatarDataPlanilha(row[1]),
            armario: row[2],
            nome: row[3],
            paciente: row[4],
            leito: row[5],
            volumes: row[6],
            horaInicio: formatarHorarioPlanilha(row[7]),
            horaFim: formatarHorarioPlanilha(row[8]),
            status: row[9],
            tipo: row[10],
            unidade: row[11],
            whatsapp: row[12] || '',
            usuario: row[13] || '',
            observacoes: row[14] || ''
          });
        }
      });

      return { success: true, data: historico.reverse() }; // Mais recentes primeiro

    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar histórico: ${error.toString()}`);
      return { success: false, error: 'Erro ao processar a solicitação.' };
    }
  });
}

function montarHistoricoTermosResponsabilidade() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Termos de Responsabilidade');

  if (!sheet && !existeArquivoBackupDisponivel('termos')) {
    return { success: true, data: [] };
  }

  var data = [];
  if (sheet && sheet.getLastRow() >= 2) {
    var totalLinhasDados = sheet.getLastRow() - 1;
    var totalColunasDados = Math.max(sheet.getLastColumn(), 21);
    data = sheet.getRange(2, 1, totalLinhasDados, totalColunasDados).getValues();
  }

  var dadosBackup = obterLinhasSheetAtualEBackups('Termos de Responsabilidade', {
    tipoArquivo: 'termos',
    incluirPlanilhaAtual: false,
    colunasMinimas: 21
  });
  if (dadosBackup.length) {
    data = data.concat(dadosBackup);
  }

  if (!data.length) {
    return { success: true, data: [] };
  }

  var historico = [];

  data.forEach(function(row) {
    if (!row[0]) {
      return;
    }

    var assinaturas = normalizarAssinaturas(row[18]);
    var dataAplicacao = converterParaDataHoraIso(row[16], '');
    var finalizadoEm = converterParaDataHoraIso(assinaturas.finalizadoEm, '');

    historico.push({
      id: row[0],
      data: formatarDataPlanilha(dataAplicacao),
      armario: row[2] || '',
      nome: row[9] || '',
      paciente: row[3] || '',
      leito: row[7] || '',
      volumes: row[15] || '',
      horaInicio: formatarHorarioPlanilha(dataAplicacao),
      horaFim: formatarHorarioPlanilha(finalizadoEm),
      status: row[19] || (finalizadoEm ? 'Finalizado' : 'Em andamento'),
      tipo: 'Acompanhante',
      unidade: row[6] || '',
      whatsapp: row[10] || '',
      usuario: assinaturas.responsavelFinalizacao || '',
      observacoes: row[20] || ''
    });
  });

  return { success: true, data: historico.reverse() };
}

function getPlanilhaLiberacao(parametros) {
  var timezone = obterTimeZoneAplicacao();

  try {
    var params = parametros || {};
    var pacienteFiltroTexto = params.paciente !== undefined && params.paciente !== null
      ? params.paciente.toString().trim()
      : '';
    var prontuarioFiltroTexto = params.prontuario !== undefined && params.prontuario !== null
      ? params.prontuario.toString().trim()
      : '';

    var dataInicioParametro = params.dataInicio !== undefined ? params.dataInicio : params.data;
    var dataFimParametro = params.dataFim !== undefined ? params.dataFim : params.data;

    var dataInicio = interpretarDataParametroSeguro(dataInicioParametro, timezone);
    var dataFim = interpretarDataParametroSeguro(dataFimParametro, timezone);

    var pacienteFiltroNormalizado = normalizarTextoBasico(pacienteFiltroTexto);
    var prontuarioFiltroNormalizado = normalizarTextoBasico(prontuarioFiltroTexto);
    var tokensPacienteFiltro = pacienteFiltroNormalizado
      ? pacienteFiltroNormalizado.split(/\s+/).filter(function(token) { return token; })
      : [];
    var tokensProntuarioFiltro = prontuarioFiltroNormalizado
      ? prontuarioFiltroNormalizado.split(/\s+/).filter(function(token) { return token; })
      : [];

    var filtroTextoAtivo = tokensPacienteFiltro.length > 0 || tokensProntuarioFiltro.length > 0;
    var aplicarFiltroData = !filtroTextoAtivo;

    if (aplicarFiltroData) {
      if (dataInicio && !dataFim) {
        dataFim = dataInicio;
      }
      if (dataFim && !dataInicio) {
        dataInicio = dataFim;
      }

      if (!dataInicio && !dataFim) {
        var dataPadrao = obterDataAtualNormalizada(timezone);
        dataInicio = dataPadrao;
        dataFim = dataPadrao;
      }

      if (!dataInicio || !dataFim) {
        return { success: false, error: 'Informe um período válido para realizar a busca.' };
      }

      if (dataInicio.getTime() > dataFim.getTime()) {
        var temp = dataInicio;
        dataInicio = dataFim;
        dataFim = temp;
      }
    }

    var chaveInicio = aplicarFiltroData ? gerarChaveDataComparacao(dataInicio, timezone) : null;
    var chaveFim = aplicarFiltroData ? gerarChaveDataComparacao(dataFim, timezone) : null;

    if (aplicarFiltroData && (!chaveInicio || !chaveFim)) {
      return { success: false, error: 'Informe um período válido para realizar a busca.' };
    }

    function contemTodosTokens(textoNormalizado, tokens) {
      if (!tokens.length) {
        return true;
      }
      if (!textoNormalizado) {
        return false;
      }
      for (var i = 0; i < tokens.length; i++) {
        if (textoNormalizado.indexOf(tokens[i]) === -1) {
          return false;
        }
      }
      return true;
    }

    var spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(PLANILHA_LIBERACAO_ID);
    } catch (erroAcesso) {
      registrarLog('ERRO', 'Falha ao acessar planilha de liberações: ' + erroAcesso.toString());
      return { success: false, error: 'Não foi possível acessar a planilha de liberações.' };
    }

    if (!spreadsheet) {
      return { success: false, error: 'Planilha de liberações não encontrada.' };
    }

    var sheet = spreadsheet.getSheetByName(PLANILHA_LIBERACAO_ABA);
    if (!sheet) {
      return { success: false, error: 'Aba de liberações não encontrada.' };
    }

    var totalColunas = Math.max(sheet.getLastColumn(), 7);
    var cabecalhos = sheet.getRange(PLANILHA_LIBERACAO_LINHA_CABECALHO, 1, 1, totalColunas).getDisplayValues()[0] || [];
    cabecalhos = cabecalhos.map(function(valor) {
      return valor === null || valor === undefined ? '' : valor;
    });
    var cabecalhosNormalizados = cabecalhos.map(function(valor) {
      return normalizarTextoBasico(valor);
    });

    function encontrarIndiceCabecalhoFlexivel(lista, termos) {
      if (!Array.isArray(lista)) {
        return -1;
      }
      var candidatos = Array.isArray(termos) ? termos : [termos];
      for (var i = 0; i < lista.length; i++) {
        var cabecalhoNormalizado = lista[i] || '';
        for (var j = 0; j < candidatos.length; j++) {
          var termoNormalizado = normalizarTextoBasico(candidatos[j]);
          if (!termoNormalizado) {
            continue;
          }
          if (cabecalhoNormalizado === termoNormalizado || cabecalhoNormalizado.indexOf(termoNormalizado) !== -1) {
            return i;
          }
        }
      }
      return -1;
    }

    var indicePaciente = encontrarIndiceCabecalhoFlexivel(cabecalhosNormalizados, ['paciente']);
    var indiceProntuario = encontrarIndiceCabecalhoFlexivel(cabecalhosNormalizados, ['prontuario']);

    var totalLinhasDados = Math.max(sheet.getLastRow() - PLANILHA_LIBERACAO_LINHA_CABECALHO, 0);
    var linhasFiltradas = [];
    var totalPlanilha = totalLinhasDados;

    if (totalLinhasDados > 0) {
      var rangeDados = sheet.getRange(PLANILHA_LIBERACAO_LINHA_CABECALHO + 1, 1, totalLinhasDados, totalColunas);
      var valoresBrutos = rangeDados.getValues();
      var valoresFormatados = rangeDados.getDisplayValues();

      for (var i = 0; i < valoresBrutos.length; i++) {
        var linhaBruta = valoresBrutos[i];
        var linhaExibicao = valoresFormatados[i];
        var possuiConteudo = linhaExibicao.some(function(celula) {
          return celula !== null && celula !== undefined && String(celula).trim() !== '';
        });

        if (!possuiConteudo) {
          continue;
        }

        var valorData = linhaBruta[PLANILHA_LIBERACAO_COLUNA_DATA - 1];
        var exibicaoData = linhaExibicao[PLANILHA_LIBERACAO_COLUNA_DATA - 1];
        var dataLinha = extrairDataValidaDaCelula(valorData, exibicaoData, timezone);
        var chaveLinha = gerarChaveDataComparacao(dataLinha, timezone);

        if (aplicarFiltroData) {
          if (!chaveLinha || chaveLinha < chaveInicio || chaveLinha > chaveFim) {
            continue;
          }
        }

        var valorPaciente = indicePaciente >= 0 && indicePaciente < linhaExibicao.length
          ? linhaExibicao[indicePaciente]
          : '';
        var valorProntuario = indiceProntuario >= 0 && indiceProntuario < linhaExibicao.length
          ? linhaExibicao[indiceProntuario]
          : '';

        var pacienteNormalizadoLinha = normalizarTextoBasico(valorPaciente);
        if (!contemTodosTokens(pacienteNormalizadoLinha, tokensPacienteFiltro)) {
          continue;
        }

        var prontuarioNormalizadoLinha = normalizarTextoBasico(valorProntuario);
        if (!contemTodosTokens(prontuarioNormalizadoLinha, tokensProntuarioFiltro)) {
          continue;
        }

        linhasFiltradas.push(linhaExibicao.map(function(celula) {
          return celula === null || celula === undefined ? '' : celula;
        }));
      }
    }

    return {
      success: true,
      columns: cabecalhos,
      rows: linhasFiltradas,
      filtro: {
        dataInicio: dataInicio ? Utilities.formatDate(dataInicio, timezone, 'yyyy-MM-dd') : '',
        dataFim: dataFim ? Utilities.formatDate(dataFim, timezone, 'yyyy-MM-dd') : '',
        dataInicioFormatada: dataInicio ? Utilities.formatDate(dataInicio, timezone, 'dd/MM/yyyy') : '',
        dataFimFormatada: dataFim ? Utilities.formatDate(dataFim, timezone, 'dd/MM/yyyy') : '',
        paciente: pacienteFiltroTexto,
        prontuario: prontuarioFiltroTexto
      },
      totalPlanilha: totalPlanilha,
      totalFiltrado: linhasFiltradas.length
    };

  } catch (erro) {
    registrarLog('ERRO', 'Erro ao buscar dados de liberação: ' + erro.toString());
    return { success: false, error: 'Erro ao buscar dados de liberação.' };
  }
}

// Funções para Cadastro de Armários Físicos
function getCadastroArmarios() {
  return executarComCache(montarChaveCache('cadastro-armarios'), CACHE_TTL_PADRAO, function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Cadastro Armários');

      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [] };
      }

      var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
      var armarios = [];

      data.forEach(function(row) {
        if (row[0]) {
          armarios.push({
            id: row[0],
            numero: row[1],
            tipo: row[2],
            unidade: row[3],
            localizacao: row[4],
            status: row[5],
            dataCadastro: row[6]
          });
        }
      });

      return { success: true, data: armarios };

    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar cadastro de armários: ${error.toString()}`);
      return { success: false, error: 'Erro ao processar a solicitação.' };
    }
  });
}

function cadastrarArmarioFisico(armarioData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Cadastro Armários');
    
    if (!sheet) {
      return { success: false, error: 'Aba de cadastro não encontrada' };
    }
    
    var totalLinhas = Math.max(sheet.getLastRow()-1, 0);
    var todosNumeros = totalLinhas > 0 ? sheet.getRange(2, 2, totalLinhas, 1).getValues().flat().filter(String) : [];

    var quantidade = parseInt(armarioData.quantidade || 1, 10);
    if (isNaN(quantidade) || quantidade < 1) {
      quantidade = 1;
    }

    var prefixo = armarioData.prefixo || '';
    var numeroInicial = parseInt(armarioData.numeroInicial || 1, 10);
    if (isNaN(numeroInicial) || numeroInicial < 1) {
      numeroInicial = 1;
    }

    var novosArmarios = [];

    if (quantidade === 1 && armarioData.numero) {
      if (todosNumeros.indexOf(armarioData.numero) !== -1) {
        return { success: false, error: 'Número de armário já existe' };
      }
      novosArmarios.push(armarioData.numero);
    } else {
      for (var i = 0; i < quantidade; i++) {
        var numeroGerado = prefixo ? prefixo + '-' + String(numeroInicial + i).padStart(3, '0') : String(numeroInicial + i);
        if (todosNumeros.indexOf(numeroGerado) !== -1 || novosArmarios.indexOf(numeroGerado) !== -1) {
          return { success: false, error: 'Não foi possível gerar numeração sem conflitos. Ajuste o prefixo ou número inicial.' };
        }
        novosArmarios.push(numeroGerado);
      }
    }

    var lastRow = sheet.getLastRow();
    var ultimoId = lastRow > 1 ? Math.max.apply(null, sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) : 0;

    var dataCadastro = obterDataHoraAtualFormatada().dataHoraIso;
    var linhas = novosArmarios.map(function(numero, index) {
      return [
        ultimoId + index + 1,
        numero,
        armarioData.tipo,
        armarioData.unidade,
        armarioData.localizacao,
        'ativo',
        dataCadastro
      ];
    });

    if (linhas.length > 0) {
      sheet.getRange(lastRow + 1, 1, linhas.length, 7).setValues(linhas);

      // Também criar nas abas de uso
      criarArmariosUso(linhas);
    }

    registrarLog('CADASTRO', `Armários físicos cadastrados: ${novosArmarios.join(', ')}`);

    limparCacheCadastroArmarios();
    limparCacheArmarios();
    limparCacheIndiceArmarios();

    return {
      success: true,
      message: 'Armários físicos cadastrados com sucesso',
      ids: linhas.map(function(linha) { return linha[0]; }),
      numeros: novosArmarios
    };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar armário físico: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function criarArmariosUso(armarios) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    armarios.forEach(function(armario) {
      var sheetName = armario[2] === 'visitante' ? 'Visitantes' : 'Acompanhantes';
      var sheet = ss.getSheetByName(sheetName);

      if (sheet) {
        var lastRow = sheet.getLastRow();
        var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;

        var estrutura = obterEstruturaPlanilha(sheet);
        if (armario[2] === 'visitante') {
          estrutura = garantirColunaVisitaEstendida(sheet, estrutura);
        }

        var totalColunas = estrutura.ultimaColuna || (armario[2] === 'visitante' ? 14 : 12);
        var novaLinha = new Array(totalColunas).fill('');
        var dataRegistro = obterDataHoraAtualFormatada().dataHoraIso;
        var nomeChavesCadastro = armario[2] === 'visitante' ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE;

        definirValorLinha(novaLinha, estrutura, 'id', novoId);
        definirValorLinha(novaLinha, estrutura, 'numero', armario[1]);
        definirValorLinha(novaLinha, estrutura, 'status', 'livre');
        definirValorLinha(novaLinha, estrutura, nomeChavesCadastro, '');
        definirValorLinha(novaLinha, estrutura, 'nome paciente', '');
        definirValorLinha(novaLinha, estrutura, 'leito', '');
        definirValorLinha(novaLinha, estrutura, 'volumes', 0);
        definirValorLinha(novaLinha, estrutura, 'hora inicio', '');
        if (armario[2] === 'visitante') {
          definirValorLinha(novaLinha, estrutura, 'hora prevista', '');
          definirValorLinha(novaLinha, estrutura, CABECALHOS_VISITA_ESTENDIDA, false);
        }
        definirValorLinha(novaLinha, estrutura, 'data registro', dataRegistro);
        definirValorLinha(novaLinha, estrutura, 'unidade', armario[3]);
        definirValorLinha(novaLinha, estrutura, CABECALHOS_WHATSAPP, '');
        definirValorLinha(novaLinha, estrutura, 'termo aplicado', false);

        sheet.getRange(lastRow + 1, 1, 1, totalColunas).setValues([novaLinha]);
      }
    });
    
  } catch (error) {
    console.error('Erro ao criar armários de uso:', error);
  }
}

// Funções para Unidades
function getUnidades() {
  return executarComCache(montarChaveCache('unidades'), CACHE_TTL_PADRAO, function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Unidades');

      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [] };
      }

      var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
      var unidades = [];

      data.forEach(function(row) {
        if (row[0]) {
          unidades.push({
            id: row[0],
            nome: row[1],
            status: row[2],
            dataCadastro: row[3]
          });
        }
      });

      return { success: true, data: unidades };

    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar unidades: ${error.toString()}`);
      return { success: false, error: 'Erro ao processar a solicitação.' };
    }
  });
}

function getSetores() {
  return executarComCache(montarChaveCache('setores'), CACHE_TTL_PADRAO, function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Cadastro');

      if (!sheet) {
        return { success: true, data: [] };
      }

      var ultimaLinha = sheet.getLastRow();
      if (ultimaLinha < 2) {
        return { success: true, data: [] };
      }

      var valores = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      var setoresMapeados = {};
      var setores = [];

      for (var i = 0; i < valores.length; i++) {
        var bruto = valores[i][0];
        if (bruto === null || bruto === undefined) {
          continue;
        }
        var texto = bruto.toString().trim();
        if (!texto) {
          continue;
        }
        var chave = normalizarTextoBasico(texto);
        if (!chave) {
          continue;
        }
        if (setoresMapeados[chave]) {
          continue;
        }
        setoresMapeados[chave] = true;
        setores.push(texto);
      }

      setores.sort(function(a, b) {
        return a.localeCompare(b, 'pt-BR', { sensitivity: 'base' });
      });

      return { success: true, data: setores };

    } catch (error) {
      registrarLog('ERRO', 'Erro ao buscar setores: ' + error.toString());
      return { success: false, error: 'Erro ao processar a solicitação.' };
    }
  });
}

function cadastrarUnidade(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Unidades');

    if (!sheet) {
      return { success: false, error: 'Aba de unidades não encontrada' };
    }
    
    // Verificar se unidade já existe
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
    var unidadeExistente = data.find(row => row[1].toLowerCase() === dados.nome.toLowerCase());
    
    if (unidadeExistente) {
      return { success: false, error: 'Unidade já cadastrada' };
    }
    
    var lastRow = sheet.getLastRow();
    var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
    
    var dataCadastro = obterDataHoraAtualFormatada().dataHoraIso;

    var novaLinha = [
      novoId,
      dados.nome,
      'ativa',
      dataCadastro
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, 4).setValues([novaLinha]);

    registrarLog('CADASTRO UNIDADE', `Unidade ${dados.nome} cadastrada`);

    limparCacheUnidades();

    return { success: true, message: 'Unidade cadastrada com sucesso', id: novoId };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar unidade: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function alternarStatusUnidade(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Unidades');
    
    if (!sheet) {
      return { success: false, error: 'Aba de unidades não encontrada' };
    }
    
    var data = sheet.getDataRange().getValues();
    var unidadeIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === dados.nome) {
        unidadeIndex = i;
        break;
      }
    }
    
    if (unidadeIndex === -1) {
      return { success: false, error: 'Unidade não encontrada' };
    }
    
    var novoStatus = data[unidadeIndex][2] === 'ativa' ? 'inativa' : 'ativa';
    sheet.getRange(unidadeIndex + 1, 3).setValue(novoStatus);

    registrarLog('ALTERAÇÃO UNIDADE', `Status da unidade ${dados.nome} alterado para ${novoStatus}`);

    limparCacheUnidades();

    return { success: true, message: `Unidade ${novoStatus === 'ativa' ? 'ativada' : 'desativada'} com sucesso` };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao alternar status da unidade: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Funções para Termos de Responsabilidade
function obterChavesCacheTermos(cache) {
  var chaves = [TERMOS_CACHE_KEY, TERMOS_CACHE_META_KEY];
  if (!cache) {
    return chaves;
  }

  try {
    var metaJson = cache.get(TERMOS_CACHE_META_KEY);
    if (metaJson) {
      try {
        var meta = JSON.parse(metaJson);
        if (meta && Array.isArray(meta.chaves)) {
          meta.chaves.forEach(function(chave) {
            if (chave && chaves.indexOf(chave) === -1) {
              chaves.push(chave);
            }
          });
        }
      } catch (erroParseMeta) {
        registrarLog('AVISO_CACHE', 'Falha ao interpretar metadados do cache de termos: ' + erroParseMeta);
      }
    }
  } catch (erroLeituraMeta) {
    registrarLog('AVISO_CACHE', 'Falha ao consultar metadados do cache de termos: ' + erroLeituraMeta);
  }

  return chaves;
}

function invalidarCacheTermosInterno(cache) {
  if (!cache) {
    return;
  }

  var chaves = obterChavesCacheTermos(cache);
  var chavesVistas = {};

  chaves.forEach(function(chave) {
    if (!chave || chavesVistas[chave]) {
      return;
    }
    chavesVistas[chave] = true;
    try {
      cache.remove(chave);
    } catch (erroRemocao) {
      registrarLog('AVISO_CACHE', 'Falha ao remover chave de cache de termos (' + chave + '): ' + erroRemocao);
    }
  });
}

function carregarTermosDoCache(cache) {
  if (!cache) {
    return { sucesso: false, termos: [] };
  }

  try {
    var metaJson = cache.get(TERMOS_CACHE_META_KEY);
    if (metaJson) {
      var termos = [];
      var meta = JSON.parse(metaJson);
      if (meta && Array.isArray(meta.chaves)) {
        var dadosIncompletos = false;
        for (var i = 0; i < meta.chaves.length; i++) {
          var chave = meta.chaves[i];
          if (!chave) {
            continue;
          }
          var chunkJson = cache.get(chave);
          if (!chunkJson) {
            dadosIncompletos = true;
            break;
          }
          try {
            var chunk = JSON.parse(chunkJson);
            if (Array.isArray(chunk) && chunk.length) {
              Array.prototype.push.apply(termos, chunk);
            }
          } catch (erroParseChunk) {
            dadosIncompletos = true;
            registrarLog('AVISO_CACHE', 'Falha ao interpretar bloco do cache de termos: ' + erroParseChunk);
            break;
          }
        }

        if (!dadosIncompletos) {
          return { sucesso: true, termos: termos };
        }

        invalidarCacheTermosInterno(cache);
        return { sucesso: false, termos: [] };
      }

      if (meta && Array.isArray(meta.chaves) && !meta.chaves.length) {
        return { sucesso: true, termos: [] };
      }
    }
  } catch (erroCacheParticionado) {
    registrarLog('AVISO_CACHE', 'Falha ao ler cache particionado de termos: ' + erroCacheParticionado);
    invalidarCacheTermosInterno(cache);
    return { sucesso: false, termos: [] };
  }

  try {
    var dadosCache = cache.get(TERMOS_CACHE_KEY);
    if (dadosCache) {
      var termosCache = JSON.parse(dadosCache);
      if (Array.isArray(termosCache)) {
        return { sucesso: true, termos: termosCache };
      }
    }
  } catch (erroCacheLegado) {
    registrarLog('AVISO_CACHE', 'Falha ao ler cache de termos: ' + erroCacheLegado);
    invalidarCacheTermosInterno(cache);
  }

  return { sucesso: false, termos: [] };
}

function armazenarTermosNoCache(cache, termos) {
  if (!cache) {
    return;
  }

  invalidarCacheTermosInterno(cache);

  if (!Array.isArray(termos) || !termos.length) {
    try {
      cache.put(TERMOS_CACHE_META_KEY, JSON.stringify({ chaves: [] }), TERMOS_CACHE_TTL);
    } catch (erroMetaVazio) {
      registrarLog('AVISO_CACHE', 'Falha ao armazenar metadados vazios do cache de termos: ' + erroMetaVazio);
    }
    try {
      cache.remove(TERMOS_CACHE_KEY);
    } catch (erroRemocaoLegado) {
      registrarLog('AVISO_CACHE', 'Falha ao remover cache legado de termos: ' + erroRemocaoLegado);
    }
    return;
  }

  var chunkChaves = [];
  var chunkAtual = [];
  var chunkTamanho = 2; // Para os colchetes do array
  var chunkIndice = 0;
  var cacheValido = true;

  function salvarChunkAtual() {
    if (!chunkAtual.length) {
      return true;
    }
    var chaveChunk = TERMOS_CACHE_CHUNK_PREFIX + '_' + chunkIndice++;
    try {
      cache.put(chaveChunk, JSON.stringify(chunkAtual), TERMOS_CACHE_TTL);
      chunkChaves.push(chaveChunk);
      return true;
    } catch (erroGravacao) {
      registrarLog('AVISO_CACHE', 'Falha ao armazenar bloco do cache de termos: ' + erroGravacao);
      return false;
    }
  }

  for (var i = 0; i < termos.length; i++) {
    var termo = termos[i];
    var termoJson;
    try {
      termoJson = JSON.stringify(termo);
    } catch (erroSerializacao) {
      registrarLog('AVISO_CACHE', 'Falha ao serializar termo para cache: ' + erroSerializacao);
      cacheValido = false;
      break;
    }

    if (!termoJson || termoJson.length + 2 > TERMOS_CACHE_CHUNK_TAMANHO_MAX) {
      registrarLog('AVISO_CACHE', 'Termo excede limite de tamanho para cache: ' + (termo && termo.id ? termo.id : 'sem id'));
      cacheValido = false;
      break;
    }

    var separador = chunkAtual.length ? 1 : 0;
    if (chunkTamanho + termoJson.length + separador > TERMOS_CACHE_CHUNK_TAMANHO_MAX) {
      if (!salvarChunkAtual()) {
        cacheValido = false;
        break;
      }
      chunkAtual = [];
      chunkTamanho = 2;
    }

    chunkAtual.push(termo);
    chunkTamanho += termoJson.length + separador;
  }

  if (cacheValido && chunkAtual.length) {
    if (!salvarChunkAtual()) {
      cacheValido = false;
    }
  }

  if (!cacheValido) {
    chunkChaves.forEach(function(chave) {
      try {
        cache.remove(chave);
      } catch (erroRemoverChunk) {
        registrarLog('AVISO_CACHE', 'Falha ao remover bloco inválido do cache de termos (' + chave + '): ' + erroRemoverChunk);
      }
    });
    return;
  }

  try {
    cache.put(TERMOS_CACHE_META_KEY, JSON.stringify({ chaves: chunkChaves }), TERMOS_CACHE_TTL);
    try {
      cache.remove(TERMOS_CACHE_KEY);
    } catch (erroRemocaoLegadoFinal) {
      registrarLog('AVISO_CACHE', 'Falha ao remover cache legado de termos após atualização: ' + erroRemocaoLegadoFinal);
    }
  } catch (erroMeta) {
    registrarLog('AVISO_CACHE', 'Falha ao armazenar metadados do cache de termos: ' + erroMeta);
    chunkChaves.forEach(function(chave) {
      try {
        cache.remove(chave);
      } catch (erroRemoverChunkFinal) {
        registrarLog('AVISO_CACHE', 'Falha ao remover bloco de cache após erro de metadados (' + chave + '): ' + erroRemoverChunkFinal);
      }
    });
  }
}

function limparCacheTermos() {
  try {
    invalidarCacheTermosInterno(CacheService.getScriptCache());
  } catch (erroCache) {
    registrarLog('AVISO_CACHE', 'Falha ao limpar cache de termos: ' + erroCache);
  }
}

function salvarTermoCompleto(dadosTermo) {
  try {
    var orientacoes = dadosTermo.orientacoes;
    if (typeof orientacoes === 'string' && orientacoes !== '') {
      try {
        orientacoes = JSON.parse(orientacoes);
      } catch (erroOrientacoes) {
        orientacoes = orientacoes.split(',').map(function(item) { return item.trim(); }).filter(String);
      }
    }
    if (!Array.isArray(orientacoes)) {
      orientacoes = [];
    }

    var volumes = dadosTermo.volumes;
    if (typeof volumes === 'string' && volumes !== '') {
      try {
        volumes = JSON.parse(volumes);
      } catch (erroVolumes) {
        volumes = [];
      }
    }
    if (!Array.isArray(volumes)) {
      volumes = [];
    }
    volumes = volumes.map(function(item) {
      if (typeof item === 'string') {
        return { quantidade: 0, descricao: item };
      }
      var quantidadeNumero = Number(item.quantidade);
      return {
        quantidade: isNaN(quantidadeNumero) ? 0 : quantidadeNumero,
        descricao: item && item.descricao ? String(item.descricao) : '',
        fotoBase64: item && item.fotoBase64 ? String(item.fotoBase64) : '',
        fotoMime: item && item.fotoMime ? String(item.fotoMime) : '',
        fotoUrl: item && item.fotoUrl ? String(item.fotoUrl) : ''
      };
    }).filter(function(item) {
      return item.quantidade > 0 && item.descricao;
    });

    var descricaoVolumes = dadosTermo.descricaoVolumes;
    if (!descricaoVolumes) {
      descricaoVolumes = volumes.map(function(item) {
        return item.quantidade + 'x ' + item.descricao;
      }).join('; ');
    }

    var totalVolumes = volumes.reduce(function(total, volume) {
      return total + (Number(volume.quantidade) || 0);
    }, 0);

    dadosTermo.orientacoes = orientacoes;
    dadosTermo.volumes = volumes;
    dadosTermo.descricaoVolumes = descricaoVolumes;
    var numeroInformado = normalizarNumeroArmario(dadosTermo.numeroArmario);

    // 1. Salvar na aba "Termos de Responsabilidade"
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Termos de Responsabilidade');

    if (!sheet) {
      throw new Error('Aba "Termos de Responsabilidade" não encontrada');
    }

    var dadosExistentes = sheet.getDataRange().getValues();
    var linhaExistente = -1;
    var termoId = null;
    var aplicadoEmAtual = obterDataHoraAtualFormatada().dataHoraIso;
    var aplicadoEm = aplicadoEmAtual;

    for (var i = dadosExistentes.length - 1; i >= 1; i--) {
      var idLinha = dadosExistentes[i][1];
      if (String(idLinha) !== String(dadosTermo.armarioId)) {
        continue;
      }

      var numeroLinha = dadosExistentes[i][2] ? dadosExistentes[i][2].toString().trim() : '';
      if (numeroInformado && normalizarNumeroArmario(numeroLinha) !== numeroInformado) {
        continue;
      }

      var assinaturasExistentes = normalizarAssinaturas(dadosExistentes[i][18]);
      var statusLinha = normalizarTextoBasico(dadosExistentes[i][19]);
      var finalizado = Boolean(dadosExistentes[i][17] || statusLinha === 'finalizado' || (assinaturasExistentes && assinaturasExistentes.finalizadoEm));

      if (!finalizado) {
        linhaExistente = i + 1;
        termoId = dadosExistentes[i][0];
        aplicadoEm = converterParaDataHoraIso(dadosExistentes[i][16], aplicadoEmAtual);
        break;
      }
    }

    if (linhaExistente === -1) {
      var lastRow = sheet.getLastRow();
      termoId = lastRow > 1 ? Math.max.apply(null, sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()) + 1 : 1;
      linhaExistente = lastRow + 1;
    }

    var valorAtualAssinatura = '';
    var statusTermo = 'Em andamento';
    if (linhaExistente <= dadosExistentes.length && linhaExistente - 1 >= 0) {
      var linhaAtual = dadosExistentes[linhaExistente - 1];
      if (linhaAtual) {
        if (linhaAtual.length > 18) {
          valorAtualAssinatura = linhaAtual[18];
        }
        if (linhaAtual.length > 19) {
          var statusExistente = linhaAtual[19];
          if (statusExistente && statusExistente.toString().trim()) {
            statusTermo = statusExistente;
          }
        }
      }
    }

    if (normalizarTextoBasico(statusTermo) !== 'finalizado') {
      statusTermo = 'Em andamento';
    }

    var assinaturasInfo = normalizarAssinaturas(valorAtualAssinatura);
    assinaturasInfo.inicial = dadosTermo.assinaturaBase64 || assinaturasInfo.inicial || '';
    assinaturasInfo.mimeInicial = dadosTermo.assinaturaMime || assinaturasInfo.mimeInicial || (assinaturasInfo.inicial ? 'image/png' : '');
    if (assinaturasInfo.inicial && assinaturasInfo.inicial.length > 49000) {
      return { success: false, error: 'A assinatura está muito grande. Limpe o campo e refaça o desenho em tamanho menor.' };
    }

    var numeroArmarioOficial = numeroInformado;

    var sheetAcompanhantes = ss.getSheetByName('Acompanhantes');
    var dadosAcompanhantes = [];
    var estruturaAcompanhantes = null;
    var linhaAcompanhante = -1;

    if (sheetAcompanhantes) {
      estruturaAcompanhantes = obterEstruturaPlanilha(sheetAcompanhantes);
      dadosAcompanhantes = sheetAcompanhantes.getDataRange().getValues();
      for (var indiceA = 1; indiceA < dadosAcompanhantes.length; indiceA++) {
        var linha = dadosAcompanhantes[indiceA];
        if (linha && String(linha[0]).trim() === String(dadosTermo.armarioId).trim()) {
          linhaAcompanhante = indiceA;
          if (linha.length > 1 && linha[1]) {
            numeroArmarioOficial = linha[1];
          } else if (!numeroArmarioOficial) {
            numeroArmarioOficial = dadosTermo.armarioId;
          }
          break;
        }
      }
    }

    if (linhaAcompanhante === -1) {
      var sheetVisitantes = ss.getSheetByName('Visitantes');
      if (sheetVisitantes) {
        var dadosVisitantes = sheetVisitantes.getDataRange().getValues();
        for (var indiceV = 1; indiceV < dadosVisitantes.length; indiceV++) {
          var linhaVisitante = dadosVisitantes[indiceV];
          if (linhaVisitante && String(linhaVisitante[0]).trim() === String(dadosTermo.armarioId).trim()) {
            if (linhaVisitante.length > 1 && linhaVisitante[1]) {
              numeroArmarioOficial = linhaVisitante[1];
            } else if (!numeroArmarioOficial) {
              numeroArmarioOficial = dadosTermo.armarioId;
            }
            break;
          }
        }
      }
    }

    if (!numeroArmarioOficial) {
      numeroArmarioOficial = dadosTermo.armarioId || '';
    }

    dadosTermo.numeroArmario = numeroArmarioOficial;

    volumes = volumes.map(function(item, indice) {
      var volumeAtual = {
        quantidade: item.quantidade,
        descricao: item.descricao,
        fotoUrl: item.fotoUrl || ''
      };

      if (item.fotoBase64) {
        var nomeEvidencia = gerarNomeArquivoEvidencia('registro_entrada', numeroArmarioOficial) + '_v' + (indice + 1);
        var arquivoFoto = salvarImagemBase64EmPasta(item.fotoBase64, item.fotoMime || 'image/jpeg', nomeEvidencia, PASTA_DRIVE_FOTOS_ID);
        if (arquivoFoto) {
          volumeAtual.fotoUrl = arquivoFoto.url;
          volumeAtual.fotoId = arquivoFoto.id;
          volumeAtual.fotoNome = arquivoFoto.nome;
        }
      }

      if (volumeAtual.fotoUrl) {
        registrarRegistroImagem({
          armarioId: dadosTermo.armarioId,
          numeroArmario: numeroArmarioOficial,
          tipo: 'termo',
          contexto: 'Aplicação do termo',
          titulo: volumeAtual.descricao ? 'Volume ' + (indice + 1) + ' - ' + volumeAtual.descricao : 'Volume ' + (indice + 1),
          detalhe: 'Quantidade: ' + (volumeAtual.quantidade || '-'),
          responsavel: dadosTermo.acompanhante || dadosTermo.paciente || '',
          dataHora: aplicadoEm,
          fotoUrl: volumeAtual.fotoUrl,
          fotoId: volumeAtual.fotoId || '',
          fotoNome: volumeAtual.fotoNome || ''
        });
      }

      return volumeAtual;
    });

    var fotosPendentes = volumes.some(function(item) { return !item.fotoUrl; });
    if (fotosPendentes) {
      return { success: false, error: 'Inclua a foto de cada volume antes de salvar o termo.' };
    }

    var linhaDados = [
      termoId,
      dadosTermo.armarioId,
      numeroArmarioOficial,
      dadosTermo.paciente,
      dadosTermo.prontuario,
      dadosTermo.nascimento,
      dadosTermo.setor,
      dadosTermo.leito,
      dadosTermo.consciente,
      dadosTermo.acompanhante,
      dadosTermo.telefone || '',
      dadosTermo.documento || '',
      dadosTermo.parentesco || '',
      orientacoes.join(','),
      JSON.stringify(volumes),
      descricaoVolumes,
      aplicadoEm,
      '',
      JSON.stringify(assinaturasInfo),
      statusTermo
    ];

    sheet.getRange(linhaExistente, 1, 1, linhaDados.length).setValues([linhaDados]);
    SpreadsheetApp.flush();

    // Confirmação rápida de persistência para evitar falsos positivos no front-end
    var linhaGravada = sheet.getRange(linhaExistente, 1, 1, 3).getValues();
    var idGravado = linhaGravada && linhaGravada[0] ? linhaGravada[0][0] : '';
    var armarioGravado = linhaGravada && linhaGravada[0] ? linhaGravada[0][1] : '';
    if (String(idGravado) !== String(termoId) || String(armarioGravado) !== String(dadosTermo.armarioId || '')) {
      throw new Error('Falha ao confirmar o salvamento do termo. Tente novamente.');
    }

    // 2. Atualizar status do armário na aba "Acompanhantes"
    var cadastroArmario = dadosTermo.cadastroArmario;
    if (typeof cadastroArmario === 'string' && cadastroArmario) {
      try {
        cadastroArmario = JSON.parse(cadastroArmario);
      } catch (erroCadastro) {
        cadastroArmario = null;
      }
    }

    var cadastroArmarioValido = cadastroArmario && String(cadastroArmario.id) === String(dadosTermo.armarioId)
      ? cadastroArmario
      : null;

    if (linhaAcompanhante > -1 && sheetAcompanhantes && estruturaAcompanhantes) {
      var totalColunasAcompanhantes = estruturaAcompanhantes.ultimaColuna || 12;
      var linhaAtualizada = dadosAcompanhantes[linhaAcompanhante] ? dadosAcompanhantes[linhaAcompanhante].slice() : [];

      while (linhaAtualizada.length < totalColunasAcompanhantes) {
        linhaAtualizada.push('');
      }

      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'volumes', totalVolumes);
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'termo aplicado', true);

      var dadosCadastroAcompanhante = cadastroArmarioValido || {
        nomeVisitante: dadosTermo.acompanhante || '',
        nomePaciente: dadosTermo.paciente || '',
        prontuario: dadosTermo.prontuario || '',
        leito: dadosTermo.leito || '',
        whatsapp: dadosTermo.telefone || ''
      };

      if (!dadosCadastroAcompanhante.prontuario) {
        dadosCadastroAcompanhante.prontuario = dadosTermo.prontuario || '';
      }

      var statusAtual = normalizarTextoBasico(obterValorLinha(linhaAtualizada, estruturaAcompanhantes, 'status', ''));
      if (statusAtual && statusAtual !== 'livre' && statusAtual !== 'contingencia') {
        throw new Error('Armário já está em uso. Atualize a lista e tente novamente.');
      }

      var dataHoraAtualCadastro = obterDataHoraAtualFormatada();
      var horaInicioCadastro = dataHoraAtualCadastro.horaCurta;
      var dataRegistroCadastro = dataHoraAtualCadastro.dataHoraIso;
      var unidadeAtual = obterValorLinha(linhaAtualizada, estruturaAcompanhantes, 'unidade', '');
      var whatsappCadastro = dadosCadastroAcompanhante.whatsapp ? dadosCadastroAcompanhante.whatsapp.toString().trim() : '';
      var nomeColunaCadastro = CABECALHOS_NOME_ACOMPANHANTE;

      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'status', 'em-uso');
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, nomeColunaCadastro, dadosCadastroAcompanhante.nomeVisitante || dadosTermo.acompanhante || '');
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'nome paciente', dadosCadastroAcompanhante.nomePaciente || dadosTermo.paciente || '');
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'prontuario', dadosCadastroAcompanhante.prontuario || dadosTermo.prontuario || '');
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'leito', dadosCadastroAcompanhante.leito || dadosTermo.leito || '');
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'hora inicio', horaInicioCadastro);
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'hora prevista', '');
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'data registro', dataRegistroCadastro);
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, CABECALHOS_WHATSAPP, whatsappCadastro);

      // Registrar histórico de uso
      var historicoSheet = ss.getSheetByName('Histórico Acompanhantes');
      if (historicoSheet) {
        var historicoLastRow = historicoSheet.getLastRow();
        var historicoId = historicoLastRow > 1
          ? Math.max.apply(null, historicoSheet.getRange(2, 1, historicoLastRow - 1, 1).getValues().flat()) + 1
          : 1;

        var historicoLinha = [
          historicoId,
          dataRegistroCadastro,
          numeroArmarioOficial,
          dadosCadastroAcompanhante.nomeVisitante || dadosTermo.acompanhante || '',
          dadosCadastroAcompanhante.nomePaciente || dadosTermo.paciente || '',
          dadosCadastroAcompanhante.leito || dadosTermo.leito || '',
          totalVolumes,
          horaInicioCadastro,
          '',
          'EM USO',
          'acompanhante',
          unidadeAtual,
          whatsappCadastro
        ];

        historicoSheet.getRange(historicoLastRow + 1, 1, 1, historicoLinha.length).setValues([historicoLinha]);
      }

      sheetAcompanhantes.getRange(linhaAcompanhante + 1, 1, 1, totalColunasAcompanhantes).setValues([linhaAtualizada]);
    }

    limparCacheTermos();
    invalidarCachesArmariosRelacionados('Acompanhantes');

    registrarLog('TERMO_APLICADO', `Termo inicial registrado para armário ${dadosTermo.numeroArmario}`);

    return {
      success: true,
      message: 'Termo registrado. Finalize na liberação para gerar o PDF.',
      termoId: termoId
    };

  } catch (error) {
    registrarLog('ERRO_TERMO', `Erro ao salvar termo: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function normalizarAssinaturas(valor) {
  var info = {
    inicial: '',
    final: '',
    mimeInicial: '',
    mimeFinal: '',
    metodoFinal: '',
    cpfFinal: '',
    finalizadoEm: '',
    responsavelFinalizacao: '',
    fotoEntregaUrl: '',
    fotoEntregaId: '',
    fotoEntregaNome: ''
  };

  if (!valor) {
    return info;
  }

  if (typeof valor === 'string') {
    try {
      var json = JSON.parse(valor);
      info.inicial = json.inicial || '';
      info.final = json.final || '';
      info.mimeInicial = json.mimeInicial || '';
      info.mimeFinal = json.mimeFinal || '';
      info.metodoFinal = json.metodoFinal || '';
      info.cpfFinal = json.cpfFinal || '';
      info.finalizadoEm = json.finalizadoEm || '';
      info.responsavelFinalizacao = json.responsavelFinalizacao || '';
      info.fotoEntregaUrl = json.fotoEntregaUrl || '';
      info.fotoEntregaId = json.fotoEntregaId || '';
      info.fotoEntregaNome = json.fotoEntregaNome || '';
      return info;
    } catch (erro) {
      info.inicial = valor;
      return info;
    }
  }

  if (typeof valor === 'object') {
    info.inicial = valor.inicial || '';
    info.final = valor.final || '';
    info.mimeInicial = valor.mimeInicial || '';
    info.mimeFinal = valor.mimeFinal || '';
    info.metodoFinal = valor.metodoFinal || '';
    info.cpfFinal = valor.cpfFinal || '';
    info.finalizadoEm = valor.finalizadoEm || '';
    info.responsavelFinalizacao = valor.responsavelFinalizacao || '';
    info.fotoEntregaUrl = valor.fotoEntregaUrl || '';
    info.fotoEntregaId = valor.fotoEntregaId || '';
    info.fotoEntregaNome = valor.fotoEntregaNome || '';
  }

  if (!info.mimeInicial && info.inicial) {
    info.mimeInicial = 'image/png';
  }

  if (!info.mimeFinal && info.final) {
    info.mimeFinal = 'image/png';
  }

  return info;
}

function montarChaveIndiceTermo(armarioId, numeroArmario) {
  var armarioTexto = armarioId !== undefined && armarioId !== null ? armarioId.toString().trim() : '';
  var numeroTexto = numeroArmario !== undefined && numeroArmario !== null ? numeroArmario.toString().trim() : '';
  return armarioTexto + '|' + numeroTexto;
}

function atualizarIndiceTermo(indice, chave, termo) {
  if (!chave) {
    return;
  }

  var atual = indice[chave] || {
    ultimo: null,
    ultimoNaoFinalizado: null,
    ultimoNaoFinalizadoSemPdf: null,
    ultimoFinalizado: null
  };

  atual.ultimo = termo;

  if (termo && termo.finalizado) {
    atual.ultimoFinalizado = termo;
  } else if (termo) {
    atual.ultimoNaoFinalizado = termo;
    if (!termo.pdfUrl) {
      atual.ultimoNaoFinalizadoSemPdf = termo;
    }
  }

  indice[chave] = atual;
}

function construirIndiceTermosPorArmario(termos) {
  var indice = {};
  if (!Array.isArray(termos)) {
    return indice;
  }

  termos.forEach(function(termo) {
    if (!termo || termo.armarioId === undefined || termo.armarioId === null) {
      return;
    }

    var armarioIdTexto = termo.armarioId.toString().trim();
    var numeroNormalizado = normalizarNumeroArmario(termo.numeroArmario);
    var chaveEspecifica = montarChaveIndiceTermo(armarioIdTexto, numeroNormalizado || '__sem_numero__');
    var chaveGenerica = montarChaveIndiceTermo(armarioIdTexto, '*');

    atualizarIndiceTermo(indice, chaveEspecifica, termo);
    atualizarIndiceTermo(indice, chaveGenerica, termo);
  });

  return indice;
}

function obterTermosRegistrados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Termos de Responsabilidade');

  if (!sheet || sheet.getLastRow() < 2) {
    return { sheet: sheet, termos: [], indiceArmario: {} };
  }

  var cache = CacheService.getScriptCache();
  var resultadoCache = carregarTermosDoCache(cache);

  if (resultadoCache && resultadoCache.sucesso) {
    var indiceCache = construirIndiceTermosPorArmario(resultadoCache.termos);
    return { sheet: sheet, termos: resultadoCache.termos, indiceArmario: indiceCache };
  }

  var data = sheet.getDataRange().getValues();
  var termos = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) {
      continue;
    }

    var assinaturas = normalizarAssinaturas(data[i][18]);
    var orientacoes = data[i][13] ? data[i][13].split(',').filter(String) : [];
    var volumes = [];

    if (data[i][14]) {
      try {
        volumes = JSON.parse(data[i][14]);
      } catch (erroVolume) {
        volumes = [];
      }
    }

    var numeroArmarioValor = data[i][2] ? data[i][2].toString().trim() : '';
    var statusBruto = data[i][19] || '';
    var statusNormalizado = normalizarTextoBasico(statusBruto);
    var possuiPdf = Boolean(data[i][17]);
    var finalizado = Boolean(possuiPdf || statusNormalizado === 'finalizado' || (assinaturas && assinaturas.finalizadoEm));

    termos.push({
      linha: i + 1,
      id: data[i][0],
      armarioId: data[i][1],
      numeroArmario: numeroArmarioValor,
      paciente: data[i][3],
      prontuario: data[i][4],
      nascimento: data[i][5],
      setor: data[i][6],
      leito: data[i][7],
      consciente: data[i][8],
      acompanhante: data[i][9],
      telefone: data[i][10],
      documento: data[i][11],
      parentesco: data[i][12],
      orientacoes: orientacoes,
      volumes: Array.isArray(volumes) ? volumes : [],
      descricaoVolumes: data[i][15],
      aplicadoEm: data[i][16],
      pdfUrl: data[i][17],
      assinaturas: assinaturas,
      status: statusBruto,
      statusNormalizado: statusNormalizado,
      finalizado: finalizado,
      possuiPdf: possuiPdf
    });
  }

  armazenarTermosNoCache(cache, termos);
  var indice = construirIndiceTermosPorArmario(termos);

  return { sheet: sheet, termos: termos, indiceArmario: indice };
}

function getTermo(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Termos de Responsabilidade');

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Termo não encontrado' };
    }
    
    var data = sheet.getDataRange().getValues();
    var termo = null;
    var termoFinalizadoMaisRecente = null;
    var armarioIdInformado = dados.armarioId !== null && dados.armarioId !== undefined
      ? dados.armarioId.toString().trim()
      : '';
    var numeroInformado = normalizarNumeroArmario(dados.numeroArmario);
    var incluirFinalizados = converterParaBoolean(dados.incluirFinalizados);

    for (var i = data.length - 1; i >= 1; i--) {
      var idLinha = data[i][1];
      var idLinhaTexto = idLinha !== null && idLinha !== undefined ? idLinha.toString().trim() : '';
      if (armarioIdInformado && idLinhaTexto !== armarioIdInformado) {
        continue;
      }

      var numeroLinha = data[i][2] ? data[i][2].toString().trim() : '';
      var numeroLinhaNormalizado = normalizarNumeroArmario(numeroLinha);
      if (numeroInformado && numeroLinhaNormalizado !== numeroInformado) {
        continue;
      }

      var assinaturas = normalizarAssinaturas(data[i][18]);
      var orientacoesSalvas = data[i][13] ? data[i][13].split(',').filter(String) : [];
      var volumesSalvos = [];
      if (data[i][14]) {
        try {
          volumesSalvos = JSON.parse(data[i][14]);
        } catch (erroVolumes) {
          volumesSalvos = [];
        }
      }

      var statusBruto = data[i][19] || '';
      var statusNormalizado = normalizarTextoBasico(statusBruto);
      var possuiPdf = Boolean(data[i][17]);
      var termoFinalizado = Boolean(possuiPdf || statusNormalizado === 'finalizado' || (assinaturas && assinaturas.finalizadoEm));

      var termoAtual = {
        id: data[i][0],
        armarioId: data[i][1],
        numeroArmario: numeroLinha,
        paciente: data[i][3],
        prontuario: data[i][4],
        nascimento: data[i][5],
        setor: data[i][6],
        leito: data[i][7],
        consciente: data[i][8],
        acompanhante: data[i][9],
        telefone: data[i][10],
        documento: data[i][11],
        parentesco: data[i][12],
        orientacoes: orientacoesSalvas,
        volumes: Array.isArray(volumesSalvos) ? volumesSalvos : [],
        descricaoVolumes: data[i][15],
        aplicadoEm: data[i][16],
        pdfUrl: data[i][17],
        assinaturaBase64: assinaturas.inicial,
        assinaturaMimeInicial: assinaturas.mimeInicial || 'image/png',
        assinaturaMimeFinal: assinaturas.mimeFinal || 'image/png',
        assinaturas: assinaturas,
        finalizadoEm: assinaturas.finalizadoEm,
        metodoFinal: assinaturas.metodoFinal,
        cpfConfirmacao: assinaturas.cpfFinal,
        status: statusBruto,
        statusNormalizado: statusNormalizado,
        finalizado: termoFinalizado
      };

      if (termoFinalizado) {
        if (incluirFinalizados && !termoFinalizadoMaisRecente) {
          termoFinalizadoMaisRecente = termoAtual;
        }
        if (!incluirFinalizados) {
          continue;
        }
      }

      termo = termoAtual;
      if (!termoFinalizado) {
        break;
      }
    }

    if (!termo && incluirFinalizados && termoFinalizadoMaisRecente) {
      termo = termoFinalizadoMaisRecente;
    }

    if (!termo) {
      return { success: false, error: 'Termo não encontrado' };
    }

    return { success: true, data: termo };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar termo: ${error.toString()}`);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function finalizarTermo(dados) {
  try {
    var armarioId = parseInt(dados.armarioId, 10);
    if (!armarioId) {
      return { success: false, error: 'ID do armário inválido' };
    }

    var metodo = (dados.metodo || 'assinatura').toString();
    var confirmacao = dados.confirmacao || '';
    var assinaturaFinal = dados.assinaturaFinal || '';
    var fotoEntregaBase64 = dados.fotoEntregaBase64 || '';
    var fotoEntregaMime = dados.fotoEntregaMime || 'image/jpeg';
    var numeroInformado = normalizarNumeroArmario(dados.numeroArmario);
    var tipoTermo = dados && dados.tipo ? dados.tipo.toString() : '';

    var termosInfo = obterTermosRegistrados();
    if (!termosInfo.sheet) {
      return { success: false, error: 'Aba "Termos de Responsabilidade" não encontrada' };
    }

    var termoEncontrado = null;
    var termoFinalizadoMaisRecente = null;
    var chaveIndice = montarChaveIndiceTermo(armarioId.toString().trim(), numeroInformado ? numeroInformado : '*');
    var indiceArmario = termosInfo.indiceArmario || {};
    var indiceSelecionado = indiceArmario[chaveIndice];

    if (indiceSelecionado) {
      termoFinalizadoMaisRecente = indiceSelecionado.ultimoFinalizado || null;
      termoEncontrado = indiceSelecionado.ultimoNaoFinalizadoSemPdf
        || indiceSelecionado.ultimoNaoFinalizado
        || indiceSelecionado.ultimo
        || null;
    }

    if (!termoEncontrado) {
      var chaveGenerica = montarChaveIndiceTermo(armarioId.toString().trim(), '*');
      var indiceGenerico = indiceArmario[chaveGenerica];
      if (indiceGenerico) {
        termoFinalizadoMaisRecente = termoFinalizadoMaisRecente || indiceGenerico.ultimoFinalizado || null;
        termoEncontrado = indiceGenerico.ultimoNaoFinalizadoSemPdf
          || indiceGenerico.ultimoNaoFinalizado
          || indiceGenerico.ultimo
          || null;
      }
    }

    if (!termoEncontrado) {
      for (var i = termosInfo.termos.length - 1; i >= 0; i--) {
        var termoAtual = termosInfo.termos[i];
        if (!termoAtual) {
          continue;
        }

        if (termoAtual.armarioId != armarioId) {
          continue;
        }

        var numeroTermo = normalizarNumeroArmario(termoAtual.numeroArmario);
        if (numeroInformado && numeroTermo !== numeroInformado) {
          continue;
        }

        var statusNormalizado = normalizarTextoBasico(termoAtual.status || termoAtual.statusNormalizado || '');
        var finalizado = termoAtual.finalizado;
        if (finalizado === undefined) {
          finalizado = Boolean(termoAtual.pdfUrl || (termoAtual.assinaturas && termoAtual.assinaturas.finalizadoEm) || statusNormalizado === 'finalizado');
        }

        if (finalizado) {
          if (!termoFinalizadoMaisRecente) {
            termoFinalizadoMaisRecente = termoAtual;
          }
          continue;
        }

        if (!termoAtual.pdfUrl) {
          termoEncontrado = termoAtual;
          break;
        }
      }
    }

    if (!termoEncontrado && termoFinalizadoMaisRecente) {
      termoEncontrado = termoFinalizadoMaisRecente;
    }

    if (!termoEncontrado) {
      return { success: false, error: 'Termo não localizado para este armário' };
    }

    var assinaturas = termoEncontrado.assinaturas || normalizarAssinaturas('');
    var finalizacaoInfo = obterDataHoraAtualFormatada();
    var finalizacaoIso = finalizacaoInfo.dataHoraIso;
    assinaturas.mimeInicial = assinaturas.mimeInicial || (termoEncontrado.assinaturas && termoEncontrado.assinaturas.mimeInicial) || 'image/png';
    assinaturas.mimeFinal = assinaturas.mimeFinal || '';
    assinaturas.metodoFinal = metodo;
    assinaturas.cpfFinal = metodo === 'cpf' ? confirmacao : '';
    assinaturas.finalizadoEm = finalizacaoIso;
    assinaturas.final = metodo === 'assinatura' ? assinaturaFinal : '';
    if (metodo === 'assinatura') {
      assinaturas.mimeFinal = dados.assinaturaMimeFinal || assinaturas.mimeFinal || 'image/png';
    }
    if (assinaturas.final && assinaturas.final.length > 49000) {
      return { success: false, error: 'A assinatura de encerramento está muito grande. Peça para refazer utilizando traços menores.' };
    }
    var responsavelFinalizacao = determinarResponsavelRegistro(dados.usuarioResponsavel);
    assinaturas.responsavelFinalizacao = responsavelFinalizacao;

    var movimentacoesResultado = getMovimentacoes({ armarioId: armarioId, numeroArmario: numeroInformado });
    var movimentacoes = [];
    if (movimentacoesResultado && movimentacoesResultado.success && Array.isArray(movimentacoesResultado.data)) {
      movimentacoes = movimentacoesResultado.data;
    } else if (movimentacoesResultado && movimentacoesResultado.success) {
      registrarLog('AVISO_TERMO', 'Dados de movimentações inválidos ao finalizar termo do armário ' + termoEncontrado.numeroArmario);
    } else if (movimentacoesResultado && !movimentacoesResultado.success) {
      registrarLog('AVISO_TERMO', 'Movimentações indisponíveis ao finalizar termo do armário ' + termoEncontrado.numeroArmario + ': ' + (movimentacoesResultado.error || 'dados inválidos'));
    }

    var numeroParaRegistro = numeroInformado || termoEncontrado.numeroArmario || 'armario';

    if (metodo === 'assinatura') {
      if (!fotoEntregaBase64 && !assinaturas.fotoEntregaUrl) {
        return { success: false, error: 'A foto da entrega é obrigatória para finalizar o termo.' };
      }

      if (fotoEntregaBase64) {
        var nomeFotoSaida = gerarNomeArquivoEvidencia('registro_saida', numeroParaRegistro);
        var arquivoSaida = salvarImagemBase64EmPasta(fotoEntregaBase64, fotoEntregaMime, nomeFotoSaida, PASTA_DRIVE_FOTOS_ID);
        if (arquivoSaida) {
          assinaturas.fotoEntregaUrl = arquivoSaida.url;
          assinaturas.fotoEntregaId = arquivoSaida.id;
          assinaturas.fotoEntregaNome = arquivoSaida.nome;
        }
      }

      if (assinaturas.fotoEntregaUrl) {
        registrarRegistroImagem({
          armarioId: armarioId,
          numeroArmario: numeroInformado || termoEncontrado.numeroArmario,
          tipo: 'entrega',
          contexto: 'Finalização do termo',
          titulo: 'Entrega do termo',
          detalhe: '',
          responsavel: termoEncontrado.acompanhante || termoEncontrado.paciente || '',
          dataHora: finalizacaoIso,
          fotoUrl: assinaturas.fotoEntregaUrl,
          fotoId: assinaturas.fotoEntregaId || '',
          fotoNome: assinaturas.fotoEntregaNome || ''
        });
      }
    }

    var dadosPDF = {
      numeroArmario: termoEncontrado.numeroArmario,
      paciente: termoEncontrado.paciente,
      prontuario: termoEncontrado.prontuario,
      nascimento: termoEncontrado.nascimento,
      setor: termoEncontrado.setor,
      leito: termoEncontrado.leito,
      consciente: termoEncontrado.consciente,
      acompanhante: termoEncontrado.acompanhante,
      telefone: termoEncontrado.telefone,
      documento: termoEncontrado.documento,
      parentesco: termoEncontrado.parentesco,
      orientacoes: termoEncontrado.orientacoes,
      volumes: termoEncontrado.volumes,
      descricaoVolumes: termoEncontrado.descricaoVolumes,
      aplicadoEm: termoEncontrado.aplicadoEm,
      finalizadoEm: finalizacaoIso,
      assinaturaInicial: assinaturas.inicial,
      assinaturaFinal: assinaturas.final,
      assinaturaMimeInicial: assinaturas.mimeInicial || 'image/png',
      assinaturaMimeFinal: assinaturas.mimeFinal || (assinaturas.final ? 'image/png' : ''),
      metodoFinal: assinaturas.metodoFinal,
      cpfFinal: assinaturas.cpfFinal,
      responsavelFinalizacao: assinaturas.responsavelFinalizacao,
      movimentacoes: movimentacoes
    };

    var resultadoPDF = gerarESalvarTermoPDF(dadosPDF);
    if (!resultadoPDF.success) {
      throw new Error(resultadoPDF.error || 'Falha ao gerar PDF');
    }

    termosInfo.sheet.getRange(termoEncontrado.linha, 18, 1, 3).setValues([[
      resultadoPDF.pdfUrl,
      JSON.stringify(assinaturas),
      'Finalizado'
    ]]);

    termoEncontrado.status = 'Finalizado';

    finalizarMovimentacoesArmario(armarioId, numeroInformado, tipoTermo);

    limparCacheTermos();
    limparCacheArmarios();

    registrarLog('TERMO_FINALIZADO', 'Termo finalizado para armário ' + termoEncontrado.numeroArmario);

    return {
      success: true,
      pdfUrl: resultadoPDF.pdfUrl,
      finalizadoEm: finalizacaoIso
    };

  } catch (error) {
    registrarLog('ERRO_TERMO', 'Erro ao finalizar termo: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function ehErroTermoNaoLocalizado(mensagemErro) {
  if (!mensagemErro) {
    return false;
  }
  var texto = normalizarTextoBasico(mensagemErro);
  return texto.indexOf('termo nao localizado') !== -1;
}

function finalizarELiberarArmario(dados) {
  try {
    var finalizacao = finalizarTermo(dados) || { success: false, error: 'Falha desconhecida na finalização' };
    var finalizacaoNaoLocalizada = !finalizacao.success && ehErroTermoNaoLocalizado(finalizacao.error);

    if (!finalizacao.success && !finalizacaoNaoLocalizada) {
      return { success: false, etapa: 'finalizacao', error: finalizacao.error || 'Erro ao finalizar termo' };
    }

    var liberacao = liberarArmario(dados.armarioId, dados.tipo || 'acompanhante', dados.numeroArmario, dados.usuarioResponsavel);
    if (!liberacao || !liberacao.success) {
      return {
        success: false,
        etapa: 'liberacao',
        error: liberacao && liberacao.error ? liberacao.error : 'Erro ao liberar armário',
        finalizacao: finalizacao
      };
    }

    return {
      success: true,
      finalizacao: finalizacao,
      liberacao: liberacao,
      termoNaoLocalizado: finalizacaoNaoLocalizada
    };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao finalizar e liberar armário: ' + error.toString());
    return { success: false, etapa: 'finalizacao-liberacao', error: 'Erro ao processar a solicitação.' };
  }
}

// Função para gerar e salvar PDF
function obterPastaSeguraDrive(pastaIdPreferida) {
  var pastaPreferida = pastaIdPreferida && pastaIdPreferida.toString().trim()
    ? pastaIdPreferida.toString().trim()
    : '';
  var pastaPadrao = PASTA_DRIVE_ID && PASTA_DRIVE_ID.toString().trim()
    ? PASTA_DRIVE_ID.toString().trim()
    : '';

  if (pastaPreferida) {
    try {
      return DriveApp.getFolderById(pastaPreferida);
    } catch (error) {
      console.warn('Pasta preferida inválida ou inacessível:', error);
    }
  }

  if (pastaPadrao) {
    try {
      return DriveApp.getFolderById(pastaPadrao);
    } catch (erroPadrao) {
      console.warn('Pasta padrão inválida ou inacessível:', erroPadrao);
    }
  }

  throw new Error('Não foi possível acessar a pasta configurada no Drive. Revise os IDs das pastas e as permissões do WebApp.');
}

function gerarNomeArquivoEvidencia(prefixo, numeroArmario) {
  var numero = numeroArmario || 'sem-numero';
  var timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'ddMMyyyy_HHmmss');
  return prefixo + '_' + numero + '_' + timestamp;
}

function salvarImagemBase64EmPasta(base64, mimePadrao, nomeArquivo, pastaId) {
  if (!base64) {
    return null;
  }

  var mime = (mimePadrao || '').toString().trim() || 'image/jpeg';
  var conteudoBase64 = base64;
  var padrao = /^data:([^;]+);base64,/i;
  var extensao = mime.indexOf('png') !== -1 ? '.png' : '.jpg';

  if (padrao.test(base64)) {
    var match = base64.match(padrao);
    if (match && match[1]) {
      mime = match[1];
      extensao = mime.indexOf('png') !== -1 ? '.png' : extensao;
    }
    conteudoBase64 = base64.replace(padrao, '');
  }

  var pasta = obterPastaSeguraDrive(pastaId || PASTA_DRIVE_FOTOS_ID);
  var blob = Utilities.newBlob(Utilities.base64Decode(conteudoBase64), mime, nomeArquivo + extensao);
  var arquivo = pasta.createFile(blob).setName(nomeArquivo + extensao);

  try {
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (erroCompartilhamento) {
    registrarLog('AVISO_DRIVE', 'Arquivo salvo sem compartilhamento público: ' + erroCompartilhamento.toString());
  }

  return {
    id: arquivo.getId(),
    url: 'https://drive.google.com/uc?export=view&id=' + arquivo.getId(),
    nome: arquivo.getName(),
    mime: mime,
    pastaId: pasta.getId()
  };
}

function gerarESalvarTermoPDF(dadosTermo, opcoes) {
  try {
    var configuracoes = opcoes || {};
    var pastaDestino = obterPastaSeguraDrive(configuracoes.pastaId);

    var htmlContent = criarHTMLTermo(dadosTermo);
    var htmlOutput = HtmlService
      .createHtmlOutput(htmlContent)
      .setWidth(800)
      .setHeight(1200);

    var blob = htmlOutput.getBlob().getAs('application/pdf');

    var prefixoArquivo = configuracoes.prefixoArquivo || 'Termo_Responsabilidade_';
    var nomeArquivo = prefixoArquivo + dadosTermo.numeroArmario + '_' +
                     Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'ddMMyyyy_HHmmss') + '.pdf';

    var arquivoPDF = pastaDestino.createFile(blob).setName(nomeArquivo);

    try {
      arquivoPDF.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (erroCompartilhamentoPdf) {
      registrarLog('AVISO_DRIVE', 'PDF salvo sem compartilhamento público: ' + erroCompartilhamentoPdf.toString());
    }

    var fileId = arquivoPDF.getId();
    var previewUrl = 'https://drive.google.com/file/d/' + fileId + '/preview';
    var downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;

    return {
      success: true,
      pdfUrl: previewUrl,
      downloadUrl: downloadUrl,
      fileId: fileId,
      pastaId: pastaDestino.getId()
    };

  } catch (error) {
    console.error('Erro ao gerar PDF:', error);
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function montarFonteAssinatura(base64, mimePadrao) {
  if (!base64) {
    return '';
  }
  if (typeof base64 === 'string' && base64.indexOf('data:') === 0) {
    return base64;
  }
  var tipo = mimePadrao && mimePadrao.toString().trim() ? mimePadrao : 'image/png';
  return 'data:' + tipo + ';base64,' + base64;
}

function criarHTMLTermo(dadosTermo) {
  var hospitalNome = 'Hospital Universitário do Ceará';
  var orientacoesPredefinidas = [
    'Os meus pertences ficam sob guarda e responsabilidade da unidade; os itens que estou levando para o leito permanecem sob minha responsabilidade.',
    'Em caso de piora clínica, o serviço social/NAC entrará em contato com a família para recolher os pertences;',
    'Após 15 dias da alta ou transferência, os itens não retirados poderão ser descartados conforme as normas vigentes.'
  ];
  var orientacoes = [];
  if (Array.isArray(dadosTermo.orientacoes) && dadosTermo.orientacoes.length) {
    orientacoes = dadosTermo.orientacoes.map(function(item) {
      switch (item) {
        case 'ori1':
          return orientacoesPredefinidas[0];
        case 'ori2':
          return orientacoesPredefinidas[1];
        case 'ori3':
          return orientacoesPredefinidas[2];
        default:
          return item;
      }
    }).filter(function(texto) { return texto && texto.trim(); });
  }
  if (!orientacoes.length) {
    orientacoes = orientacoesPredefinidas;
  }

  var volumesLista = Array.isArray(dadosTermo.volumes) ? dadosTermo.volumes : [];
  var movimentacoesLista = Array.isArray(dadosTermo.movimentacoes) ? dadosTermo.movimentacoes : [];

  function formatarDataHoraCompleta(data) {
    if (!data) return 'Não informada';
    try {
      var date = new Date(data);
      if (isNaN(date.getTime())) {
        return data;
      }
      return Utilities.formatDate(date, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
    } catch (erro) {
      return data;
    }
  }

  function formatarMovimentacao(mov) {
    var data = mov.data ? new Date(mov.data) : null;
    var hora = mov.hora ? new Date('1970-01-01T' + mov.hora + 'Z') : null;
    var dataFormatada = data && !isNaN(data.getTime())
      ? Utilities.formatDate(data, 'America/Sao_Paulo', 'dd/MM/yyyy')
      : (mov.data || '');
    var horaFormatada = mov.hora
      ? mov.hora
      : (hora && !isNaN(hora.getTime())
          ? Utilities.formatDate(hora, 'America/Sao_Paulo', 'HH:mm')
          : '');
    var assinaturaSrc = montarFonteAssinatura(mov.assinatura, mov.assinaturaMime || 'image/png');
    var assinaturaHtml = assinaturaSrc
      ? '<img src="' + assinaturaSrc + '" class="assinatura-img" alt="Assinatura do responsável" />'
      : (mov.responsavel || '');
    return {
      data: dataFormatada,
      hora: horaFormatada,
      tipo: (mov.tipo || '').toString().toUpperCase(),
      descricao: mov.descricao || '',
      responsavel: mov.responsavel || '',
      assinaturaHtml: assinaturaHtml
    };
  }

  var movimentosNormalizados = movimentacoesLista.map(formatarMovimentacao);
  while (movimentosNormalizados.length < 8) {
    movimentosNormalizados.push({ data: '', hora: '', tipo: '', descricao: '', responsavel: '', assinaturaHtml: '' });
  }
  var dataDevolucaoTexto = dadosTermo.finalizadoEm
    ? formatarDataHoraCompleta(dadosTermo.finalizadoEm)
    : '__________________________';
  var conferenteTexto = (dadosTermo.responsavelFinalizacao || '').toString().trim();
  if (!conferenteTexto) {
    conferenteTexto = '__________________________';
  }

  var assinaturaInicialSrc = montarFonteAssinatura(dadosTermo.assinaturaInicial, dadosTermo.assinaturaMimeInicial);
  var assinaturaInicialHtml = assinaturaInicialSrc
    ? '<img src="' + assinaturaInicialSrc + '" class="assinatura-img" alt="Assinatura inicial" />'
    : '<div class="assinatura-linha">Assinatura não registrada digitalmente.</div>';

  var assinaturaFinalHtml = '';
  if (dadosTermo.metodoFinal === 'cpf' && dadosTermo.cpfFinal) {
    assinaturaFinalHtml = '<div class="assinatura-linha">Confirmação por CPF: ' + dadosTermo.cpfFinal + '</div>';
  } else if (dadosTermo.metodoFinal === 'manual') {
    assinaturaFinalHtml = '<div class="assinatura-linha">Finalização manual registrada no sistema.</div>';
  } else if (dadosTermo.assinaturaFinal) {
    var assinaturaFinalSrc = montarFonteAssinatura(dadosTermo.assinaturaFinal, dadosTermo.assinaturaMimeFinal);
    assinaturaFinalHtml = assinaturaFinalSrc
      ? '<img src="' + assinaturaFinalSrc + '" class="assinatura-img" alt="Assinatura final" />'
      : '<div class="assinatura-linha">Assinatura final não registrada.</div>';
  } else {
    assinaturaFinalHtml = '<div class="assinatura-linha">Assinatura final não registrada.</div>';
  }

  var partes = [];
  partes.push('<!DOCTYPE html>');
  partes.push('<html>');
  partes.push('<head>');
  partes.push('<base target="_top">');
  partes.push('<style>');
  partes.push('  body { font-family: Arial, sans-serif; margin: 24px; color: #0b1324; }');
  partes.push('  h1, h2, h3 { margin: 0; }');
  partes.push('  .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0b1324; padding-bottom: 12px; margin-bottom: 16px; }');
  partes.push('  .header h1 { font-size: 20px; text-transform: uppercase; }');
  partes.push('  .info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 6px 24px; margin-bottom: 16px; font-size: 13px; }');
  partes.push('  .section-title { font-weight: bold; text-transform: uppercase; font-size: 13px; margin: 18px 0 8px; }');
  partes.push('  .orientacoes { font-size: 12px; margin-left: 18px; }');
  partes.push('  .volumes-table { width: 100%; border-collapse: collapse; margin-top: 6px; font-size: 12px; }');
  partes.push('  .volumes-table th, .volumes-table td { border: 1px solid #0b1324; padding: 6px; text-align: left; }');
  partes.push('  .assinatura-box { margin-top: 20px; text-align: center; }');
  partes.push('  .assinatura-img { max-width: 260px; max-height: 120px; border: 1px solid #d0d7e2; padding: 6px; }');
  partes.push('  .assinatura-linha { border-bottom: 1px solid #0b1324; display: inline-block; padding: 4px 16px; min-width: 240px; font-size: 12px; }');
  partes.push('  .footer { margin-top: 18px; font-size: 10px; text-align: center; color: #3d4a63; }');
  partes.push('  .page-break { page-break-before: always; }');
  partes.push('  .mov-table { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 6px; }');
  partes.push('  .mov-table th, .mov-table td { border: 1px solid #0b1324; padding: 5px; vertical-align: top; }');
  partes.push('  .mov-table th { background: #eef3fb; }');
  partes.push('  .devolucao-box { border: 1px solid #0b1324; padding: 10px; margin-top: 10px; min-height: 70px; }');
  partes.push('  .observacoes { border: 1px solid #0b1324; min-height: 90px; margin-top: 12px; padding: 8px; font-size: 12px; }');
  partes.push('  .label { font-weight: bold; }');
  partes.push('</style>');
  partes.push('</head>');
  partes.push('<body>');
  partes.push('<div class="header">');
  partes.push('  <div>');
  partes.push('    <h1>Termo de Responsabilidade</h1>');
  partes.push('    <h3>' + hospitalNome + '</h3>');
  partes.push('  </div>');
  partes.push('  <div style="text-align:right;font-size:12px;">');
  partes.push('    <div><strong>Nº do Armário:</strong> ' + (dadosTermo.numeroArmario || '') + '</div>');
  partes.push('    <div><strong>Data de início:</strong> ' + formatarDataParaHTML(dadosTermo.aplicadoEm) + '</div>');
  partes.push('  </div>');
  partes.push('</div>');
  partes.push('<div class="section-title">Dados do Paciente</div>');
  partes.push('<div class="info-grid">');
  partes.push('  <div><span class="label">Nome:</span> ' + (dadosTermo.paciente || '') + '</div>');
  partes.push('  <div><span class="label">Prontuário:</span> ' + (dadosTermo.prontuario || '') + '</div>');
  partes.push('  <div><span class="label">Data de nascimento:</span> ' + formatarDataParaHTML(dadosTermo.nascimento) + '</div>');
  partes.push('  <div><span class="label">Setor/Leito:</span> ' + (dadosTermo.setor || '') + ' ' + (dadosTermo.leito || '') + '</div>');
  partes.push('  <div><span class="label">Paciente consciente/orientado:</span> ' + (dadosTermo.consciente || '') + '</div>');
  partes.push('</div>');
  partes.push('<div class="section-title">Responsável pelo Armário</div>');
  partes.push('<div class="info-grid">');
  partes.push('  <div><span class="label">Nome:</span> ' + (dadosTermo.acompanhante || '') + '</div>');
  partes.push('  <div><span class="label">Documento:</span> ' + (dadosTermo.documento || 'Não informado') + '</div>');
  partes.push('  <div><span class="label">Telefone:</span> ' + (dadosTermo.telefone || 'Não informado') + '</div>');
  partes.push('  <div><span class="label">Parentesco:</span> ' + (dadosTermo.parentesco || 'Não informado') + '</div>');
  partes.push('</div>');
  partes.push('<div class="section-title">Orientações repassadas</div>');
  partes.push('<ul class="orientacoes">');
  orientacoes.forEach(function(item) {
    partes.push('<li>' + item + '</li>');
  });
  partes.push('</ul>');
  partes.push('<div class="section-title">Volumes armazenados</div>');
  partes.push('<table class="volumes-table">');
  partes.push('  <thead><tr><th style="width:20%">Quantidade</th><th>Descrição</th></tr></thead>');
  partes.push('  <tbody>');
  if (volumesLista.length) {
    volumesLista.forEach(function(volume) {
      partes.push('<tr><td>' + (volume.quantidade || '') + '</td><td>' + (volume.descricao || '') + '</td></tr>');
    });
  } else {
    partes.push('<tr><td colspan="2">Sem volumes informados.</td></tr>');
  }
  partes.push('  </tbody>');
  partes.push('</table>');
  partes.push('<div class="assinatura-box">');
  partes.push('  <div class="section-title">Assinatura do responsável - Etapa inicial</div>');
  partes.push(assinaturaInicialHtml);
  partes.push('  <div style="margin-top:6px; font-size:12px;">' + (dadosTermo.acompanhante || '') + '</div>');
  partes.push('</div>');
  partes.push('<div class="footer">Primeira etapa concluída em ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') + '.</div>');
  partes.push('<div class="page-break"></div>');
  partes.push('<div class="section-title">Movimentações registradas</div>');
  partes.push('<table class="mov-table">');
  partes.push('  <thead><tr><th style="width:16%">Data</th><th style="width:12%">Hora</th><th style="width:18%">Tipo</th><th>Descrição</th><th style="width:22%">Assinatura do responsável</th></tr></thead>');
  partes.push('  <tbody>');
  movimentosNormalizados.forEach(function(mov) {
    partes.push('<tr><td>' + (mov.data || '') + '</td><td>' + (mov.hora || '') + '</td><td>' + (mov.tipo || '') + '</td><td>' + (mov.descricao || '') + '</td><td>' + (mov.assinaturaHtml || '') + '</td></tr>');
  });
  partes.push('  </tbody>');
  partes.push('</table>');
  partes.push('<div class="section-title">Devolução de pertences</div>');
  partes.push('<div class="devolucao-box">Data: ' + dataDevolucaoTexto + ' &nbsp;&nbsp; Conferente: ' + conferenteTexto + '</div>');
  partes.push('<div class="section-title">Observações complementares</div>');
  partes.push('<div class="observacoes"></div>');
  partes.push('<div class="assinatura-box">');
  partes.push('  <div class="section-title">Assinatura de encerramento</div>');
  partes.push(assinaturaFinalHtml);
  partes.push('  <div style="margin-top:6px; font-size:12px;">' + (dadosTermo.acompanhante || '') + '</div>');
  partes.push('  <div style="margin-top:4px; font-size:11px;">Encerrado em: ' + formatarDataHoraCompleta(dadosTermo.finalizadoEm) + '</div>');
  partes.push('</div>');
  partes.push('<div class="footer">Documento gerado automaticamente em ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') + ' - ' + hospitalNome + '.</div>');
  partes.push('</body>');
  partes.push('</html>');

  return partes.join('');
}


function formatarDataParaHTML(data) {
  if (!data) return 'Não informada';
  try {
    var date = new Date(data);
    return Utilities.formatDate(date, 'America/Sao_Paulo', 'dd/MM/yyyy');
  } catch (error) {
    return data;
  }
}

function montarDadosTermoParaPDF(termo, movimentacoes) {
  var assinaturas = termo.assinaturas || {};
  var statusNormalizado = normalizarTextoBasico(termo.status);
  var finalizadoEm = assinaturas.finalizadoEm || termo.finalizadoEm ||
    (statusNormalizado === 'finalizado' ? termo.aplicadoEm : '');

  return {
    numeroArmario: termo.numeroArmario || termo.armarioId || '',
    paciente: termo.paciente || '',
    prontuario: termo.prontuario || '',
    nascimento: termo.nascimento || '',
    setor: termo.setor || '',
    leito: termo.leito || '',
    consciente: termo.consciente || '',
    acompanhante: termo.acompanhante || '',
    telefone: termo.telefone || '',
    documento: termo.documento || '',
    parentesco: termo.parentesco || '',
    orientacoes: termo.orientacoes || [],
    volumes: termo.volumes || [],
    descricaoVolumes: termo.descricaoVolumes || '',
    aplicadoEm: termo.aplicadoEm || '',
    finalizadoEm: finalizadoEm || '',
    assinaturaInicial: assinaturas.inicial || '',
    assinaturaFinal: assinaturas.final || '',
    assinaturaMimeInicial: assinaturas.mimeInicial || 'image/png',
    assinaturaMimeFinal: assinaturas.mimeFinal || (assinaturas.final ? 'image/png' : ''),
    metodoFinal: assinaturas.metodoFinal || '',
    cpfFinal: assinaturas.cpfFinal || '',
    responsavelFinalizacao: assinaturas.responsavelFinalizacao || '',
    movimentacoes: movimentacoes || []
  };
}

function gerarTermoPDFTemporario(dados) {
  try {
    var parametros = dados || {};
    var termoResposta = getTermo({
      armarioId: parametros.armarioId,
      numeroArmario: parametros.numeroArmario,
      incluirFinalizados: true
    });

    if (!termoResposta.success || !termoResposta.data) {
      return { success: false, error: termoResposta.error || 'Termo não encontrado' };
    }

    var termo = termoResposta.data;
    var movResposta = getMovimentacoes({ armarioId: termo.armarioId, numeroArmario: termo.numeroArmario, tipo: 'acompanhante' });
    var movimentacoes = movResposta && movResposta.success && Array.isArray(movResposta.data) ? movResposta.data : [];

    var dadosPDF = montarDadosTermoParaPDF(termo, movimentacoes);
    var resultadoPDF = gerarESalvarTermoPDF(dadosPDF, {
      pastaId: PASTA_DRIVE_TEMP_ID,
      prefixoArquivo: 'Termo_Temporario_'
    });

    if (!resultadoPDF.success) {
      return { success: false, error: resultadoPDF.error || 'Falha ao gerar PDF temporário' };
    }

    registrarLog('TERMO_PDF_TEMP', 'PDF temporário gerado para armário ' + (termo.numeroArmario || termo.armarioId || ''));

    return { success: true, data: { pdfUrl: resultadoPDF.pdfUrl, downloadUrl: resultadoPDF.downloadUrl, fileId: resultadoPDF.fileId } };

  } catch (error) {
    registrarLog('ERRO_TERMO', 'Erro ao gerar PDF temporário: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function excluirArquivoTemporario(dados) {
  var fileId = dados && dados.fileId ? dados.fileId.toString().trim() : '';
  if (!fileId) {
    return { success: false, error: 'ID do arquivo não informado' };
  }

  try {
    var arquivo = DriveApp.getFileById(fileId);
    var pastaTemporaria = obterPastaSeguraDrive(PASTA_DRIVE_TEMP_ID);
    var pertencePastaTemporaria = false;

    try {
      var pastas = arquivo.getParents();
      while (pastas.hasNext()) {
        var pasta = pastas.next();
        if (pasta.getId() === pastaTemporaria.getId()) {
          pertencePastaTemporaria = true;
          break;
        }
      }
    } catch (erroPastas) {
      console.warn('Não foi possível verificar pastas do arquivo temporário.', erroPastas);
    }

    if (pertencePastaTemporaria) {
      try {
        pastaTemporaria.removeFile(arquivo);
      } catch (erroRemocao) {
        console.warn('Falha ao remover arquivo da pasta temporária, enviando para lixeira.', erroRemocao);
      }
    }

    arquivo.setTrashed(true);
    registrarLog('TERMO_PDF_TEMP_REMOVIDO', 'PDF temporário removido: ' + fileId);
    return { success: true, removido: true };

  } catch (error) {
    registrarLog('ERRO_TERMO', 'Erro ao excluir PDF temporário: ' + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Funções para Movimentações
function garantirEstruturaMovimentacoes(sheet) {
  var colunaStatus = 10;
  var colunaItens = 11;
  var colunaVolume = 12;
  var colunaAssinatura = 13;
  var colunaAssinaturaMime = 14;
  var colunaFotoUrl = 15;
  var colunaFotoId = 16;

  if (!sheet) {
    return {
      colunaStatus: colunaStatus,
      colunaItens: colunaItens,
      colunaVolume: colunaVolume,
      colunaFotoUrl: colunaFotoUrl,
      colunaFotoId: colunaFotoId,
      ultimaColuna: colunaFotoId
    };
  }

  var minimoColunas = Math.max(colunaItens, colunaStatus, colunaVolume, colunaAssinaturaMime, colunaFotoId);
  var totalColunas = sheet.getLastColumn();
  if (totalColunas < minimoColunas) {
    sheet.insertColumnsAfter(totalColunas, minimoColunas - totalColunas);
    totalColunas = sheet.getLastColumn();
  }

  var cabecalhos = sheet.getRange(1, 1, 1, Math.max(totalColunas, minimoColunas)).getValues()[0];
  if (!cabecalhos[colunaStatus - 1]) {
    sheet.getRange(1, colunaStatus).setValue('Status');
  }
  if (!cabecalhos[colunaItens - 1]) {
    sheet.getRange(1, colunaItens).setValue('Itens');
  }
  if (!cabecalhos[colunaVolume - 1]) {
    sheet.getRange(1, colunaVolume).setValue('Volume');
  }
  if (!cabecalhos[colunaAssinatura - 1]) {
    sheet.getRange(1, colunaAssinatura).setValue('Assinatura');
  }
  if (!cabecalhos[colunaAssinaturaMime - 1]) {
    sheet.getRange(1, colunaAssinaturaMime).setValue('Assinatura MIME');
  }
  if (!cabecalhos[colunaFotoUrl - 1]) {
    sheet.getRange(1, colunaFotoUrl).setValue('Foto URL');
  }
  if (!cabecalhos[colunaFotoId - 1]) {
    sheet.getRange(1, colunaFotoId).setValue('Foto ID');
  }

  return {
    colunaStatus: colunaStatus,
    colunaItens: colunaItens,
    colunaVolume: colunaVolume,
    colunaAssinatura: colunaAssinatura,
    colunaAssinaturaMime: colunaAssinaturaMime,
    colunaFotoUrl: colunaFotoUrl,
    colunaFotoId: colunaFotoId,
    ultimaColuna: Math.max(totalColunas, minimoColunas)
  };
}

function garantirEstruturaRegistroImagens(sheet) {
  var totalColunasMinimas = 12;
  var colunaFotoUrl = 10;
  var colunaFotoId = 11;
  var colunaFotoNome = 12;

  if (!sheet) {
    return {
      totalColunas: totalColunasMinimas,
      colunaFotoUrl: colunaFotoUrl,
      colunaFotoId: colunaFotoId,
      colunaFotoNome: colunaFotoNome
    };
  }

  var cabecalhos = ['ID', 'Armário ID', 'Número Armário', 'Tipo', 'Contexto', 'Título', 'Detalhe', 'Responsável', 'Data/Hora', 'Foto URL', 'Foto ID', 'Foto Nome'];
  var totalColunas = sheet.getLastColumn();
  if (totalColunas < totalColunasMinimas) {
    if (totalColunas === 0) {
      sheet.insertColumns(1, totalColunasMinimas);
    } else {
      sheet.insertColumnsAfter(totalColunas, totalColunasMinimas - totalColunas);
    }
    totalColunas = sheet.getLastColumn();
  }

  var linhaCabecalho = sheet.getRange(1, 1, 1, Math.max(totalColunas, totalColunasMinimas)).getValues()[0];
  for (var i = 0; i < cabecalhos.length; i++) {
    if (!linhaCabecalho[i]) {
      sheet.getRange(1, i + 1).setValue(cabecalhos[i]);
    }
  }

  return {
    totalColunas: Math.max(totalColunas, totalColunasMinimas),
    colunaFotoUrl: colunaFotoUrl,
    colunaFotoId: colunaFotoId,
    colunaFotoNome: colunaFotoNome
  };
}

function registrarRegistroImagem(registro) {
  try {
    if (!registro || !registro.fotoUrl) {
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Registro de Imagens');
    if (!sheet) {
      sheet = ss.insertSheet('Registro de Imagens');
    }

    var estrutura = garantirEstruturaRegistroImagens(sheet);
    var ultimaLinha = sheet.getLastRow();
    var urlNormalizada = registro.fotoUrl.toString().trim();
    var totalRegistros = Math.max(ultimaLinha - 1, 0);

    if (totalRegistros > 0) {
      var urlsExistentes = sheet.getRange(2, estrutura.colunaFotoUrl, totalRegistros, 1).getValues().flat();
      var possuiUrl = urlsExistentes.some(function(valor) {
        return valor && valor.toString().trim() === urlNormalizada;
      });
      if (possuiUrl) {
        return;
      }
    }

    var novoId = ultimaLinha > 1
      ? Math.max.apply(null, sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues().flat().map(function(valor) {
          var numero = Number(valor);
          return Number.isFinite(numero) ? numero : 0;
        })) + 1
      : 1;

    var dataHoraRegistro = registro.dataHora || obterDataHoraAtualFormatada().dataHoraIso;
    var novaLinha = new Array(estrutura.totalColunas).fill('');
    novaLinha[0] = novoId;
    novaLinha[1] = registro.armarioId || '';
    novaLinha[2] = registro.numeroArmario || '';
    novaLinha[3] = registro.tipo || '';
    novaLinha[4] = registro.contexto || '';
    novaLinha[5] = registro.titulo || '';
    novaLinha[6] = registro.detalhe || '';
    novaLinha[7] = registro.responsavel || '';
    novaLinha[8] = dataHoraRegistro;
    novaLinha[estrutura.colunaFotoUrl - 1] = urlNormalizada;
    novaLinha[estrutura.colunaFotoId - 1] = registro.fotoId || '';
    novaLinha[estrutura.colunaFotoNome - 1] = registro.fotoNome || '';

    sheet.getRange(ultimaLinha + 1, 1, 1, estrutura.totalColunas).setValues([novaLinha]);
  } catch (erroRegistroImagem) {
    registrarLog('ERRO_IMAGEM', 'Falha ao registrar imagem: ' + erroRegistroImagem.toString());
  }
}

function construirRegistrosImagensLegado(dados) {
  var imagens = [];
  var parametros = dados || {};

  try {
    var termoResposta = getTermo({
      armarioId: parametros.armarioId,
      numeroArmario: parametros.numeroArmario,
      incluirFinalizados: true
    });

    if (termoResposta && termoResposta.success && termoResposta.data) {
      var termo = termoResposta.data;
      var volumes = Array.isArray(termo.volumes) ? termo.volumes : [];
      var aplicadoEm = termo.aplicadoEm ? converterParaDataHoraIso(termo.aplicadoEm, '') : '';
      volumes.forEach(function(volume, indice) {
        if (!volume.fotoUrl) return;
        imagens.push({
          tipo: 'termo',
          contexto: 'Aplicação do termo',
          titulo: volume.descricao ? 'Volume ' + (indice + 1) + ' - ' + volume.descricao : 'Volume ' + (indice + 1),
          detalhe: 'Quantidade: ' + (volume.quantidade || '-'),
          responsavel: termo.acompanhante || termo.paciente || '',
          dataHora: aplicadoEm,
          fotoUrl: volume.fotoUrl,
          fotoId: volume.fotoId || '',
          fotoNome: volume.fotoNome || ''
        });
      });

      if (termo.assinaturas && termo.assinaturas.fotoEntregaUrl) {
        imagens.push({
          tipo: 'entrega',
          contexto: 'Finalização do termo',
          titulo: 'Entrega do termo',
          detalhe: '',
          responsavel: termo.acompanhante || termo.paciente || '',
          dataHora: termo.assinaturas.finalizadoEm || '',
          fotoUrl: termo.assinaturas.fotoEntregaUrl,
          fotoId: termo.assinaturas.fotoEntregaId || '',
          fotoNome: termo.assinaturas.fotoEntregaNome || ''
        });
      }
    }
  } catch (erroTermoImagem) {
    registrarLog('AVISO_IMAGEM', 'Erro ao coletar imagens do termo: ' + erroTermoImagem.toString());
  }

  try {
    var movResposta = getMovimentacoes({
      armarioId: parametros.armarioId,
      numeroArmario: parametros.numeroArmario,
      incluirFinalizados: true
    });

    if (movResposta && movResposta.success && Array.isArray(movResposta.data)) {
      movResposta.data.forEach(function(mov) {
        if (!mov.fotoUrl) return;
        var itensMov = [];
        if (Array.isArray(mov.itens)) {
          itensMov = mov.itens;
        } else if (mov.itens && typeof mov.itens === 'string') {
          try {
            var itensParseados = JSON.parse(mov.itens);
            if (Array.isArray(itensParseados)) {
              itensMov = itensParseados;
            }
          } catch (erroItensMov) {
            itensMov = [];
          }
        }
        var descricaoItens = itensMov.length
          ? 'Itens: ' + itensMov.map(function(item) { return Number(item.quantidade) + 'x ' + item.descricao; }).join('; ')
          : '';
        var dataHora = mov.data && mov.hora
          ? mov.data + ' ' + mov.hora
          : mov.dataHoraRegistro || '';

        imagens.push({
          tipo: 'movimentacao',
          contexto: formatarTituloMovimento(normalizarTextoBasico(mov.tipo)) || 'Movimentação',
          titulo: mov.descricao || 'Registro da movimentação',
          detalhe: descricaoItens,
          responsavel: mov.responsavel || '',
          dataHora: dataHora,
          fotoUrl: mov.fotoUrl,
          fotoId: mov.fotoId || '',
          fotoNome: mov.fotoNome || ''
        });
      });
    }
  } catch (erroMovImagem) {
    registrarLog('AVISO_IMAGEM', 'Erro ao coletar imagens de movimentações: ' + erroMovImagem.toString());
  }

  return imagens;
}

function getRegistrosImagens(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Registro de Imagens');
    var registros = [];
    var possuiRegistrosArmario = false;
    var numeroFiltro = normalizarNumeroArmario(dados && dados.numeroArmario ? dados.numeroArmario : '');
    var armarioIdTexto = dados && dados.armarioId !== null && dados.armarioId !== undefined ? dados.armarioId.toString().trim() : '';

    var valoresConsolidados = [];
    if (sheet && sheet.getLastRow() >= 2) {
      var estrutura = garantirEstruturaRegistroImagens(sheet);
      var totalLinhas = sheet.getLastRow() - 1;
      var valoresAtuais = sheet.getRange(2, 1, totalLinhas, estrutura.totalColunas).getValues();
      valoresConsolidados = valoresConsolidados.concat(valoresAtuais);
    }
    var valoresBackup = obterLinhasSheetAtualEBackups('Registro de Imagens', {
      tipoArquivo: 'geral',
      incluirPlanilhaAtual: false,
      colunasMinimas: 11
    });
    if (valoresBackup.length) {
      valoresConsolidados = valoresConsolidados.concat(valoresBackup);
    }

    if (valoresConsolidados.length) {
      valoresConsolidados.forEach(function(linha) {
        var numeroLinha = normalizarNumeroArmario(linha[2]);
        if (armarioIdTexto && String(linha[1]).trim() !== armarioIdTexto) {
          return;
        }
        if (!armarioIdTexto && numeroFiltro && numeroLinha !== numeroFiltro) {
          return;
        }

        possuiRegistrosArmario = true;

        var url = linha[9];
        if (!url) {
          return;
        }

        registros.push({
          id: linha[0],
          armarioId: linha[1],
          numeroArmario: linha[2],
          tipo: linha[3],
          contexto: linha[4],
          titulo: linha[5],
          detalhe: linha[6],
          responsavel: linha[7],
          dataHora: linha[8],
          fotoUrl: url,
          fotoId: linha[10] || '',
          fotoNome: linha[11] || ''
        });
      });
    }

    if (!registros.length) {
      registros = construirRegistrosImagensLegado({ armarioId: armarioIdTexto, numeroArmario: numeroFiltro });
    }

    if (!possuiRegistrosArmario && registros.length) {
      try {
        registros.forEach(function(item) { registrarRegistroImagem(item); });
        sheet = ss.getSheetByName('Registro de Imagens');
      } catch (erroPersistencia) {
        registrarLog('AVISO_IMAGEM', 'Falha ao persistir registros de imagens recuperados: ' + erroPersistencia.toString());
      }
    }

    registros.sort(function(a, b) {
      var dataA = a.dataHora ? a.dataHora.toString() : '';
      var dataB = b.dataHora ? b.dataHora.toString() : '';
      return dataB.localeCompare(dataA);
    });

    return { success: true, data: registros };
  } catch (erroRegistros) {
    registrarLog('ERRO_IMAGEM', 'Erro ao buscar registros de imagens: ' + erroRegistros.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function existeArquivoBackupDisponivel(tipoArquivo) {
  return listarArquivosBackup(tipoArquivo).length > 0;
}

function obterLinhasSheetAtualEBackups(nomeAba, opcoes) {
  var config = opcoes || {};
  var tipoArquivo = config.tipoArquivo || 'geral';
  var incluirPlanilhaAtual = config.incluirPlanilhaAtual !== false;
  var colunasMinimas = config.colunasMinimas || 1;
  var linhas = [];

  if (incluirPlanilhaAtual) {
    try {
      var ssAtual = SpreadsheetApp.getActiveSpreadsheet();
      var abaAtual = ssAtual.getSheetByName(nomeAba);
      if (abaAtual && abaAtual.getLastRow() >= 2) {
        var colunasAtual = Math.max(abaAtual.getLastColumn(), colunasMinimas);
        linhas = linhas.concat(abaAtual.getRange(2, 1, abaAtual.getLastRow() - 1, colunasAtual).getValues());
      }
    } catch (erroAtual) {
      registrarLog('AVISO_BACKUP', 'Falha ao ler aba atual "' + nomeAba + '": ' + erroAtual.toString());
    }
  }

  var arquivos = listarArquivosBackup(tipoArquivo);
  arquivos.forEach(function(arquivo) {
    try {
      var planilha = SpreadsheetApp.open(arquivo);
      var aba = planilha.getSheetByName(nomeAba);
      if (!aba || aba.getLastRow() < 2) {
        return;
      }
      var colunas = Math.max(aba.getLastColumn(), colunasMinimas);
      var dados = aba.getRange(2, 1, aba.getLastRow() - 1, colunas).getValues();
      if (dados.length) {
        linhas = linhas.concat(dados);
      }
    } catch (erroArquivo) {
      registrarLog('AVISO_BACKUP', 'Falha ao ler backup "' + arquivo.getName() + '" (' + nomeAba + '): ' + erroArquivo.toString());
    }
  });

  return linhas;
}

function listarArquivosBackup(tipoArquivo) {
  var raiz;
  try {
    raiz = DriveApp.getFolderById(PASTA_BACKUP_RAIZ_ID);
  } catch (erroPasta) {
    registrarLog('AVISO_BACKUP', 'Pasta de backup não acessível: ' + erroPasta.toString());
    return [];
  }

  var prefixo = tipoArquivo === 'termos' ? 'BACKUP-TERMOS-' : (tipoArquivo === 'logs' ? 'BACKUP-LOGS-' : 'BACKUP-GERAL-');
  var arquivos = [];

  if (tipoArquivo === 'termos' || tipoArquivo === 'logs') {
    var nomePastaTopo = tipoArquivo === 'termos' ? 'TERMOS' : 'LOGS';
    var pastasTopo = raiz.getFoldersByName(nomePastaTopo);
    if (!pastasTopo.hasNext()) {
      return [];
    }
    arquivos = coletarArquivosBackupEmPastas(pastasTopo.next(), prefixo);
  } else {
    arquivos = coletarArquivosBackupEmPastas(raiz, prefixo);
  }

  arquivos.sort(function(a, b) {
    return (a.getName() || '').localeCompare(b.getName() || '');
  });

  return arquivos;
}

function coletarArquivosBackupEmPastas(pastaRaiz, prefixo) {
  var arquivos = [];
  if (!pastaRaiz) {
    return arquivos;
  }

  var arquivosDiretos = pastaRaiz.getFiles();
  while (arquivosDiretos.hasNext()) {
    var arquivo = arquivosDiretos.next();
    var nome = (arquivo.getName() || '').toString();
    if (nome.indexOf(prefixo) === 0) {
      arquivos.push(arquivo);
    }
  }

  var subpastas = pastaRaiz.getFolders();
  while (subpastas.hasNext()) {
    var subpasta = subpastas.next();
    arquivos = arquivos.concat(coletarArquivosBackupEmPastas(subpasta, prefixo));
  }

  return arquivos;
}

function garantirEstruturaHistorico(sheet) {
  if (!sheet) {
    return 13;
  }
  var minimoColunas = 15;
  var totalColunas = sheet.getLastColumn();
  if (totalColunas < minimoColunas) {
    sheet.insertColumnsAfter(totalColunas, minimoColunas - totalColunas);
    totalColunas = sheet.getLastColumn();
  }
  var cabecalhos = sheet.getRange(1, 1, 1, Math.max(totalColunas, minimoColunas)).getValues()[0];
  if (!cabecalhos[13]) {
    sheet.getRange(1, 14).setValue('Usuário');
  }
  if (!cabecalhos[14]) {
    sheet.getRange(1, 15).setValue('Observações');
  }
  return Math.max(totalColunas, minimoColunas);
}

function formatarTituloMovimento(tipo) {
  var mapa = {
    entrada: 'Entrada de pertences',
    saida: 'Saída de pertences',
    'saída': 'Saída de pertences',
    conferencia: 'Conferência',
    'conferência': 'Conferência'
  };
  var chave = normalizarTextoBasico(tipo);
  return mapa[chave] || tipo || '';
}

function garantirColunaVisitaEstendida(sheet, estrutura) {
  if (!sheet) {
    return estrutura;
  }

  var estruturaAtual = estrutura || obterEstruturaPlanilha(sheet);
  var indiceExistente = obterIndiceColuna(estruturaAtual, CABECALHOS_VISITA_ESTENDIDA, null);

  if (indiceExistente !== null && indiceExistente !== undefined) {
    return estruturaAtual;
  }

  var indiceHoraPrevista = obterIndiceColuna(estruturaAtual, 'hora prevista', null);
  if (indiceHoraPrevista !== null && indiceHoraPrevista !== undefined) {
    sheet.insertColumnAfter(indiceHoraPrevista + 1);
    sheet.getRange(1, indiceHoraPrevista + 2).setValue('Visita Estendida');
  } else {
    var ultimaColuna = estruturaAtual.ultimaColuna || sheet.getLastColumn();
    sheet.insertColumnAfter(ultimaColuna);
    sheet.getRange(1, ultimaColuna + 1).setValue('Visita Estendida');
  }

  return obterEstruturaPlanilha(sheet);
}

function garantirColunaProntuario(sheet, estrutura) {
  if (!sheet) {
    return estrutura;
  }

  var estruturaAtual = estrutura || obterEstruturaPlanilha(sheet);
  var indiceProntuario = obterIndiceColuna(estruturaAtual, 'prontuario', null);

  if (indiceProntuario !== null && indiceProntuario !== undefined) {
    return estruturaAtual;
  }

  var ultimaColuna = estruturaAtual.ultimaColuna || sheet.getLastColumn();
  sheet.insertColumnAfter(ultimaColuna);
  sheet.getRange(1, ultimaColuna + 1).setValue('Prontuário');

  return obterEstruturaPlanilha(sheet);
}

function garantirColunaObservacoesAcompanhantes(sheet, estrutura) {
  if (!sheet) {
    return estrutura;
  }

  var estruturaAtual = estrutura || obterEstruturaPlanilha(sheet);
  var indiceObservacoes = obterIndiceColuna(estruturaAtual, CABECALHOS_OBSERVACOES, null);

  if (indiceObservacoes !== null && indiceObservacoes !== undefined) {
    return estruturaAtual;
  }

  var indiceReferencia = obterIndiceColuna(estruturaAtual, 'termo aplicado', null);
  var ultimaColuna = estruturaAtual.ultimaColuna || sheet.getLastColumn();
  if (indiceReferencia !== null && indiceReferencia !== undefined) {
    sheet.insertColumnAfter(indiceReferencia + 1);
    sheet.getRange(1, indiceReferencia + 2).setValue('Observações');
  } else {
    sheet.insertColumnAfter(ultimaColuna);
    sheet.getRange(1, ultimaColuna + 1).setValue('Observações');
  }

  return obterEstruturaPlanilha(sheet);
}

function garantirColunasFotoContingencia(sheet, estrutura) {
  if (!sheet) {
    return estrutura;
  }

  var estruturaAtual = estrutura || obterEstruturaPlanilha(sheet);
  var indiceUrl = obterIndiceColuna(estruturaAtual, ['foto contingencia url', 'foto contingência url'], null);
  var indiceId = obterIndiceColuna(estruturaAtual, ['foto contingencia id', 'foto contingência id'], null);
  var indiceNome = obterIndiceColuna(estruturaAtual, ['foto contingencia nome', 'foto contingência nome'], null);

  if (indiceUrl !== null && indiceId !== null && indiceNome !== null) {
    return estruturaAtual;
  }

  var ultimaColuna = estruturaAtual.ultimaColuna || sheet.getLastColumn();
  var novasColunas = ['Foto Contingência URL', 'Foto Contingência ID', 'Foto Contingência Nome'];

  novasColunas.forEach(function(nomeColuna, index) {
    if (obterIndiceColuna(estruturaAtual, nomeColuna, null) === null) {
      sheet.insertColumnAfter(ultimaColuna + index);
      sheet.getRange(1, ultimaColuna + index + 1).setValue(nomeColuna);
    }
  });

  return obterEstruturaPlanilha(sheet);
}

function getMovimentacoes(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Movimentações');
  var estruturaMov = garantirEstruturaMovimentacoes(sheet);
  var colunaStatus = estruturaMov.colunaStatus;
  var colunaItens = estruturaMov.colunaItens;
  var colunaVolume = estruturaMov.colunaVolume;
  var colunaAssinatura = estruturaMov.colunaAssinatura;
  var colunaAssinaturaMime = estruturaMov.colunaAssinaturaMime;
  var colunaFotoUrl = estruturaMov.colunaFotoUrl;
  var colunaFotoId = estruturaMov.colunaFotoId;
  var possuiArmario = dados && dados.armarioId !== undefined && dados.armarioId !== null && dados.armarioId !== '';
  var armarioId = possuiArmario ? dados.armarioId : null;
  var armarioIdTexto = possuiArmario && armarioId !== null && armarioId !== undefined ? armarioId.toString().trim() : '';
  var numeroInformado = normalizarNumeroArmario(dados ? dados.numeroArmario : '');
  var tipoInformado = dados && dados.tipo ? normalizarTextoBasico(dados.tipo) : '';
  var tiposMovimentacaoValidos = ['entrada', 'saida', 'saída', 'conferencia', 'conferência'];
  var deveFiltrarPorTipoMovimentacao = tiposMovimentacaoValidos.indexOf(tipoInformado) !== -1;
  var incluirFinalizados = converterParaBoolean(dados && dados.incluirFinalizados);
  var possuiIdentificacao = armarioIdTexto || numeroInformado;
  var chaveIdentificacao = armarioIdTexto
    ? [armarioIdTexto, numeroInformado, tipoInformado, incluirFinalizados ? 'finalizados' : 'ativos'].join('|')
    : numeroInformado
      ? ['numero', numeroInformado, tipoInformado, incluirFinalizados ? 'finalizados' : 'ativos'].join('|')
      : incluirFinalizados ? 'todos_finalizados' : 'todos';
  var chaveCache = montarChaveCache('movimentacoes', chaveIdentificacao);

  return executarComCache(chaveCache, CACHE_TTL_MOVIMENTACOES, function() {
    try {
      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [] };
      }

      var totalLinhas = sheet.getLastRow() - 1;
      if (totalLinhas <= 0 || !possuiIdentificacao) {
        return { success: true, data: [] };
      }

      var largura = Math.max(
        estruturaMov.ultimaColuna,
        sheet.getLastColumn(),
        colunaStatus,
        colunaItens,
        colunaVolume,
        colunaAssinatura,
        colunaAssinaturaMime,
        colunaFotoUrl,
        colunaFotoId
      );
      var dadosMovimentacoes = sheet.getRange(2, 1, totalLinhas, largura).getValues();
      var movimentacoes = [];
      for (var j = 0; j < dadosMovimentacoes.length; j++) {
        var linhaDados = dadosMovimentacoes[j];
        var idLinha = linhaDados[1];
        var idLinhaTexto = idLinha !== null && idLinha !== undefined ? idLinha.toString().trim() : '';
        if (armarioIdTexto && idLinhaTexto !== armarioIdTexto) {
          continue;
        }

        var numeroLinhaNormalizado = normalizarNumeroArmario(linhaDados[2]);
        if (numeroInformado && numeroLinhaNormalizado !== numeroInformado) {
          continue;
        }

        var tipoLinha = linhaDados[3];
        if (deveFiltrarPorTipoMovimentacao && normalizarTextoBasico(tipoLinha) !== tipoInformado) {
          continue;
        }

        var statusLinha = colunaStatus ? linhaDados[colunaStatus - 1] : '';
        var statusNormalizado = normalizarTextoBasico(statusLinha);
        if (!incluirFinalizados && statusNormalizado === 'finalizado') {
          continue;
        }

        var itensValor = colunaItens && colunaItens <= linhaDados.length ? linhaDados[colunaItens - 1] : '';
        var volumeValor = colunaVolume && colunaVolume <= linhaDados.length ? linhaDados[colunaVolume - 1] : '';
        var itens = [];
        if (Array.isArray(itensValor)) {
          itens = itensValor;
        } else if (typeof itensValor === 'string' && itensValor.trim()) {
          try {
            var itensParse = JSON.parse(itensValor);
            if (Array.isArray(itensParse)) {
              itens = itensParse;
            }
          } catch (erroItens) {
            itens = [];
          }
        }

        if ((!itens || !itens.length) && linhaDados[4]) {
          var volumeNormalizado = Number(volumeValor);
          var volumeItens = Number.isFinite(volumeNormalizado) && volumeNormalizado > 0 ? volumeNormalizado : 1;
          itens = [{ quantidade: volumeItens, descricao: linhaDados[4].toString() }];
        }

        movimentacoes.push({
          id: linhaDados[0],
          armarioId: linhaDados[1],
          numeroArmario: linhaDados[2],
          tipo: linhaDados[3],
          descricao: linhaDados[4],
          responsavel: linhaDados[5],
          data: formatarDataPlanilha(linhaDados[6]),
          hora: formatarHorarioPlanilha(linhaDados[7]),
          dataHoraRegistro: converterParaDataHoraIso(linhaDados[8], ''),
          status: statusLinha || '',
          assinatura: colunaAssinatura && colunaAssinatura <= linhaDados.length ? linhaDados[colunaAssinatura - 1] : '',
          assinaturaMime: colunaAssinaturaMime && colunaAssinaturaMime <= linhaDados.length
            ? linhaDados[colunaAssinaturaMime - 1]
            : '',
          itens: itens,
          fotoUrl: colunaFotoUrl && colunaFotoUrl <= linhaDados.length ? linhaDados[colunaFotoUrl - 1] : '',
          fotoId: colunaFotoId && colunaFotoId <= linhaDados.length ? linhaDados[colunaFotoId - 1] : ''
        });
      }

      return { success: true, data: movimentacoes };

    } catch (error) {
      registrarLog('ERRO', 'Erro ao buscar movimentações: ' + error.toString());
      return { success: false, error: 'Erro ao processar a solicitação.' };
    }
  });
}

function getMovimentacoesResumo(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Movimentações');
  var estruturaMov = garantirEstruturaMovimentacoes(sheet);
  var colunaStatus = estruturaMov.colunaStatus;

  var listaArmarios = dados && dados.armarios ? dados.armarios : [];
  if (typeof listaArmarios === 'string') {
    try {
      listaArmarios = JSON.parse(listaArmarios);
    } catch (erroParse) {
      listaArmarios = [];
    }
  }

  if (!sheet || sheet.getLastRow() < 2 || !Array.isArray(listaArmarios) || !listaArmarios.length) {
    return { success: true, data: {} };
  }

  var chavesEsperadas = {};
  var numerosParaChaves = {};

  listaArmarios.forEach(function(item) {
    var idNormalizado = normalizarIdentificador(item.id || item.armarioId);
    var numeroNormalizado = normalizarNumeroArmario(item.numero || item.numeroArmario);
    var chavePrincipal = idNormalizado || numeroNormalizado;

    if (!chavePrincipal) {
      return;
    }

    chavesEsperadas[chavePrincipal] = true;

    if (numeroNormalizado) {
      if (!numerosParaChaves[numeroNormalizado]) {
        numerosParaChaves[numeroNormalizado] = [];
      }
      numerosParaChaves[numeroNormalizado].push(chavePrincipal);
    }
  });

  if (Object.keys(chavesEsperadas).length === 0) {
    return { success: true, data: {} };
  }

  var totalLinhas = sheet.getLastRow() - 1;
  if (totalLinhas <= 0) {
    return { success: true, data: {} };
  }

  var largura = Math.max(estruturaMov.ultimaColuna, sheet.getLastColumn(), estruturaMov.colunaItens, estruturaMov.colunaVolume, estruturaMov.colunaFotoId);
  var linhas = sheet.getRange(2, 1, totalLinhas, largura).getValues();
  var contagens = {};

  for (var i = 0; i < linhas.length; i++) {
    var linha = linhas[i];
    var statusLinha = colunaStatus && colunaStatus <= linha.length ? linha[colunaStatus - 1] : '';
    var statusNormalizado = normalizarTextoBasico(statusLinha);
    if (statusNormalizado === 'finalizado') {
      continue;
    }

    var idLinha = normalizarIdentificador(linha[1]);
    var numeroLinha = normalizarNumeroArmario(linha[2]);
    var chave = '';

    if (idLinha && chavesEsperadas[idLinha]) {
      chave = idLinha;
    } else if (numeroLinha && numerosParaChaves[numeroLinha] && numerosParaChaves[numeroLinha].length) {
      chave = numerosParaChaves[numeroLinha][0];
    }

    if (!chave) {
      continue;
    }

    contagens[chave] = (contagens[chave] || 0) + 1;
  }

  return { success: true, data: contagens };
}

function salvarMovimentacao(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Movimentações');

    if (!sheet) {
      return { success: false, error: 'Aba de movimentações não encontrada' };
    }

  var estruturaMov = garantirEstruturaMovimentacoes(sheet);
  var colunaStatus = estruturaMov.colunaStatus;
  var colunaItens = estruturaMov.colunaItens;
  var colunaVolume = estruturaMov.colunaVolume;
  var colunaAssinatura = estruturaMov.colunaAssinatura;
  var colunaAssinaturaMime = estruturaMov.colunaAssinaturaMime;
  var colunaFotoUrl = estruturaMov.colunaFotoUrl;
  var colunaFotoId = estruturaMov.colunaFotoId;
  var larguraMovimentacao = Math.max(
    colunaItens,
    estruturaMov.ultimaColuna,
    sheet.getLastColumn(),
    colunaVolume,
    colunaAssinaturaMime,
    colunaFotoId
  );

    // Buscar número do armário
    var tipoArmarioNormalizado = normalizarTextoBasico(dados.tipoArmario);
    var tipoNormalizado = normalizarTextoBasico(dados.tipo);
    var numeroArmario = normalizarNumeroArmario(dados.numeroArmario);
    if (!numeroArmario) {
      var nomeSheetArmario = tipoNormalizado === 'visitante' ? 'Visitantes' : 'Acompanhantes';
      var armarioSheet = ss.getSheetByName(nomeSheetArmario);
      if (armarioSheet) {
        var estruturaArmario = obterEstruturaPlanilha(armarioSheet);
        var totalLinhasArmario = armarioSheet.getLastRow();
        if (totalLinhasArmario > 1) {
          if (nomeSheetArmario === 'Visitantes') {
            estruturaArmario = garantirColunaVisitaEstendida(armarioSheet, estruturaArmario);
          }
          var dadosArmario = armarioSheet.getRange(2, 1, totalLinhasArmario - 1, estruturaArmario.ultimaColuna || (nomeSheetArmario === 'Visitantes' ? 14 : 12)).getValues();
          for (var i = 0; i < dadosArmario.length; i++) {
            var linha = dadosArmario[i];
            if (String(linha[0]) === String(dados.armarioId)) {
              var numeroLinha = obterValorLinha(linha, estruturaArmario, 'numero', linha[1]);
              numeroArmario = numeroLinha ? numeroLinha.toString().trim() : '';
              break;
            }
          }
        }
      }
    }

    var lastRow = sheet.getLastRow();
    var novoId = lastRow > 1 ? Math.max.apply(null, sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()) + 1 : 1;

    var registroAtual = obterDataHoraAtualFormatada();
    var dataMovimentacao = formatarDataPlanilha(dados.data);
    var horaMovimentacao = formatarHorarioPlanilha(dados.hora);
    var registroMovimento = registroAtual.dataHoraIso;
  var assinaturaBase64 = (dados.assinaturaBase64 || '').toString().trim();
  var assinaturaMime = (dados.assinaturaMime || '').toString().trim() || 'image/png';
  var fotoBase64 = (dados.fotoBase64 || '').toString().trim();
  var fotoMime = (dados.fotoMime || '').toString().trim() || 'image/jpeg';
  var fotoUrl = (dados.fotoUrl || '').toString().trim();
  var fotoId = '';
  var fotoNome = '';
  var primeiraFotoItem = { base64: '', mime: '' };

  if (!assinaturaBase64) {
    return { success: false, error: 'A assinatura do responsável é obrigatória.' };
  }

    var padraoPrefixo = /^data:([^;]+);base64,/i;
    if (padraoPrefixo.test(assinaturaBase64)) {
      var match = assinaturaBase64.match(padraoPrefixo);
      if (match && match[1]) {
        assinaturaMime = match[1];
      }
      assinaturaBase64 = assinaturaBase64.replace(padraoPrefixo, '');
    }

  var itensSerializados = '';
  var itensNormalizados = [];
  var itensInformados = dados.itens;

    if (typeof itensInformados === 'string' && itensInformados) {
      try {
        itensInformados = JSON.parse(itensInformados);
      } catch (erroParseItens) {
        itensInformados = [];
      }
    }

      if (Array.isArray(itensInformados) && itensInformados.length) {
        try {
          itensNormalizados = itensInformados.map(function(item) {
            var quantidadeNumero = Number(item.quantidade);
            var fotoItemBase64 = item && item.fotoBase64 ? item.fotoBase64.toString().trim() : '';
            var fotoItemMime = item && item.fotoMime ? item.fotoMime.toString().trim() : '';

            if (!primeiraFotoItem.base64 && fotoItemBase64) {
              primeiraFotoItem.base64 = fotoItemBase64;
              primeiraFotoItem.mime = fotoItemMime || 'image/jpeg';
            }

            return {
              quantidade: Number.isFinite(quantidadeNumero) ? quantidadeNumero : 0,
              descricao: item && item.descricao ? item.descricao.toString() : ''
            };
          }).filter(function(item) {
            return item.quantidade > 0 && item.descricao;
          });
        } catch (erroItens) {
          itensNormalizados = [];
        }
      }

    var volumeTotal = 0;
    if (itensNormalizados.length) {
      itensSerializados = JSON.stringify(itensNormalizados);
      volumeTotal = itensNormalizados.reduce(function(soma, item) { return soma + (Number(item.quantidade) || 0); }, 0);
    }

    if (!itensSerializados && dados.descricao) {
      var volumeFallback = Number(dados.volume);
      var volumeConsiderado = Number.isFinite(volumeFallback) && volumeFallback > 0 ? volumeFallback : 1;
      itensNormalizados = [{ quantidade: volumeConsiderado, descricao: dados.descricao.toString() }];
      itensSerializados = JSON.stringify(itensNormalizados);
      volumeTotal = volumeConsiderado;
    }

  if (!fotoBase64 && primeiraFotoItem.base64) {
    fotoBase64 = primeiraFotoItem.base64;
    fotoMime = primeiraFotoItem.mime || fotoMime;
  }

  if (!fotoBase64 && !fotoUrl) {
    return { success: false, error: 'A foto da movimentação é obrigatória.' };
  }

  if (!volumeTotal) {
    var volumeInformado = Number(dados.volume);
    volumeTotal = Number.isFinite(volumeInformado) && volumeInformado > 0 ? volumeInformado : 1;
  }

  if (fotoBase64) {
      var nomeMov = gerarNomeArquivoEvidencia(tipoNormalizado === 'saida' || tipoNormalizado === 'saída' ? 'registro_mov_saida' : 'registro_mov_entrada', numeroArmario || dados.numeroArmario);
      var arquivoMov = salvarImagemBase64EmPasta(fotoBase64, fotoMime, nomeMov, PASTA_DRIVE_FOTOS_ID);
      if (arquivoMov) {
        fotoUrl = arquivoMov.url;
        fotoMime = arquivoMov.mime || fotoMime;
        fotoId = arquivoMov.id;
        fotoNome = arquivoMov.nome || '';
      }
    }

    var fotoIdValor = fotoUrl ? (fotoId || '') : '';

  if (fotoUrl) {
    registrarRegistroImagem({
      armarioId: dados.armarioId,
      numeroArmario: numeroArmario,
      tipo: 'movimentacao',
      contexto: formatarTituloMovimento(tipoNormalizado) || 'Movimentação',
      titulo: dados.descricao || 'Registro da movimentação',
      detalhe: itensNormalizados.length ? 'Itens: ' + itensNormalizados.map(function(item) { return Number(item.quantidade) + 'x ' + item.descricao; }).join('; ') : '',
      responsavel: dados.responsavel || '',
      dataHora: registroMovimento,
      fotoUrl: fotoUrl,
      fotoId: fotoIdValor,
      fotoNome: fotoNome
    });
  }

  var novaLinha = new Array(larguraMovimentacao).fill('');
  novaLinha[0] = novoId;
  novaLinha[1] = dados.armarioId;
  novaLinha[2] = numeroArmario;
    novaLinha[3] = dados.tipo;
    novaLinha[4] = dados.descricao;
    novaLinha[5] = dados.responsavel;
    novaLinha[6] = dataMovimentacao;
    novaLinha[7] = horaMovimentacao;
    novaLinha[8] = registroMovimento;
    novaLinha[colunaStatus - 1] = 'ativo';
    if (colunaItens) {
      novaLinha[colunaItens - 1] = itensSerializados;
    }
  if (colunaVolume) {
    novaLinha[colunaVolume - 1] = volumeTotal || '';
  }
  if (colunaAssinatura) {
    novaLinha[colunaAssinatura - 1] = assinaturaBase64;
  }
  if (colunaAssinaturaMime) {
    novaLinha[colunaAssinaturaMime - 1] = assinaturaMime;
  }
  if (colunaFotoUrl) {
    novaLinha[colunaFotoUrl - 1] = fotoUrl || '';
  }
  if (colunaFotoId) {
    novaLinha[colunaFotoId - 1] = fotoIdValor;
  }

    sheet.getRange(lastRow + 1, 1, 1, novaLinha.length).setValues([novaLinha]);

    registrarLog('MOVIMENTAÇÃO', "Movimentação registrada para armário " + numeroArmario);

    limparCacheMovimentacoes(dados.armarioId, numeroArmario, tipoArmarioNormalizado || tipoNormalizado);

    return { success: true, message: 'Movimentação registrada com sucesso', id: novoId };

  } catch (error) {
    registrarLog('ERRO', "Erro ao salvar movimentação: " + error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function finalizarMovimentacoesArmario(armarioId, numeroArmario, tipo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Movimentações');

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    var estruturaMov = garantirEstruturaMovimentacoes(sheet);
    var colunaStatus = estruturaMov.colunaStatus;
    var totalLinhas = sheet.getLastRow() - 1;
    if (totalLinhas <= 0) {
      return;
    }

    var idTexto = armarioId !== undefined && armarioId !== null ? armarioId.toString().trim() : '';
    var numeroNormalizado = normalizarNumeroArmario(numeroArmario);
    var valoresId = sheet.getRange(2, 2, totalLinhas, 1).getValues();
    var valoresNumero = sheet.getRange(2, 3, totalLinhas, 1).getValues();
    var valoresStatus = sheet.getRange(2, colunaStatus, totalLinhas, 1).getValues();
    var houveAtualizacao = false;

    for (var i = 0; i < totalLinhas; i++) {
      var idLinha = valoresId[i][0];
      var idLinhaTexto = idLinha !== null && idLinha !== undefined ? idLinha.toString().trim() : '';
      if (idLinhaTexto !== idTexto) {
        continue;
      }

      if (numeroNormalizado && normalizarNumeroArmario(valoresNumero[i][0]) !== numeroNormalizado) {
        continue;
      }

      var statusAtual = normalizarTextoBasico(valoresStatus[i][0]);
      if (statusAtual === 'finalizado') {
        continue;
      }

      valoresStatus[i][0] = 'finalizado';
      houveAtualizacao = true;
    }

    if (houveAtualizacao) {
      sheet.getRange(2, colunaStatus, totalLinhas, 1).setValues(valoresStatus);
      limparCacheMovimentacoes(armarioId, numeroArmario, tipo);
    }

  } catch (error) {
    registrarLog('AVISO_MOVIMENTACAO', 'Falha ao finalizar movimentações: ' + error.toString());
  }
}

// Funções para LOGS
// Achados e Perdidos - Pertences em guarda-volume
function obterSheetPertencesPerdidos() {
  var spreadsheet = SpreadsheetApp.openById(PLANILHA_PERTENCES_PERDIDOS_ID);
  var sheet = spreadsheet.getSheetByName(PLANILHA_PERTENCES_PERDIDOS_ABA);
  if (!sheet) {
    throw new Error('Aba de pertences perdidos não encontrada.');
  }
  return sheet;
}

function garantirEstruturaPertences(sheet) {
  var headersObrigatorios = [
    'DONO DO PERTENCE',
    'DATA DA GUARDA',
    'DESCRIÇÃO',
    'TELEFONE PARA CONTATO',
    'FOI ENCONTRADO?',
    'HISTÓRICO DE CONTATO'
  ];

  var totalColunasExistentes = Math.max(sheet.getLastColumn() - PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL + 1, headersObrigatorios.length);
  var cabecalhosRange = sheet.getRange(PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO, PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL, 1, totalColunasExistentes);
  var cabecalhos = cabecalhosRange.getValues()[0];

  var alterado = false;
  headersObrigatorios.forEach(function(nome, indice) {
    if (!cabecalhos[indice]) {
      cabecalhos[indice] = nome;
      alterado = true;
    }
  });

  if (alterado) {
    cabecalhosRange.setValues([cabecalhos]);
  }

  var estrutura = {
    ultimaColuna: totalColunasExistentes,
    mapaIndices: {}
  };

  cabecalhos.forEach(function(cabecalho, indice) {
    var chave = normalizarTextoBasico(cabecalho);
    if (chave && estrutura.mapaIndices[chave] === undefined) {
      estrutura.mapaIndices[chave] = indice;
    }
  });

  return estrutura;
}

function listarPertencesPerdidos() {
  try {
    var sheet = obterSheetPertencesPerdidos();
    var estrutura = garantirEstruturaPertences(sheet);

    var totalLinhasDados = Math.max(sheet.getLastRow() - PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO, 0);
    if (totalLinhasDados <= 0) {
      return { success: true, dados: [] };
    }

    var rangeValores = sheet.getRange(
      PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO + 1,
      PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL,
      totalLinhasDados,
      estrutura.ultimaColuna
    );

    var valores = rangeValores.getValues();
    var exibicoes = rangeValores.getDisplayValues();

    var itens = valores.map(function(linha, indice) {
      var linhaExibicao = exibicoes[indice] || [];
      var linhaPlanilha = PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO + 1 + indice;

      var dono = obterValorLinhaFlexivel(linha, estrutura, ['dono do pertence', 'dono'], '');
      var dataGuardaValor = obterValorLinhaFlexivel(linha, estrutura, ['data da guarda', 'data'], '');
      var dataGuardaTexto = obterValorLinhaFlexivel(linhaExibicao, estrutura, ['data da guarda', 'data'], '');
      var descricao = obterValorLinhaFlexivel(linha, estrutura, ['descrição', 'descricao'], '');
      var telefone = obterValorLinhaFlexivel(linha, estrutura, ['telefone para contato', 'telefone'], '');
      var encontradoTexto = obterValorLinhaFlexivel(linhaExibicao, estrutura, ['foi encontrado?', 'encontrado'], '');
      var historicoContato = obterValorLinhaFlexivel(linha, estrutura, ['histórico de contato', 'historico de contato', 'historico'], '');

      var encontradoNormalizado = normalizarTextoBasico(encontradoTexto);
      var encontrado = encontradoNormalizado === 'sim' || encontradoNormalizado === 'true';

      return {
        idLinha: linhaPlanilha,
        dono: dono,
        dataGuarda: dataGuardaValor,
        dataGuardaTexto: dataGuardaTexto,
        descricao: descricao,
        telefone: telefone,
        encontrado: encontrado,
        encontradoTexto: encontrado ? 'Sim' : 'Não',
        historicoContato: historicoContato || ''
      };
    });

    return { success: true, dados: itens };
  } catch (erro) {
    registrarLog('ERRO', 'Erro ao listar pertences perdidos: ' + erro.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function cadastrarPertencePerdido(parametros) {
  try {
    var sheet = obterSheetPertencesPerdidos();
    var estrutura = garantirEstruturaPertences(sheet);

    var dono = (parametros.dono || '').toString().trim();
    var descricao = (parametros.descricao || '').toString().trim();
    var telefone = (parametros.telefone || '').toString().trim();
    var dataEntrada = parametros.dataGuarda ? new Date(parametros.dataGuarda) : '';
    if (dataEntrada && isNaN(dataEntrada.getTime())) {
      dataEntrada = '';
    }

    var novoHistorico = (parametros.historicoContato || '').toString().trim();

    var linha = new Array(estrutura.ultimaColuna).fill('');
    definirValorLinhaFlexivel(linha, estrutura, ['dono do pertence', 'dono'], dono);
    definirValorLinhaFlexivel(linha, estrutura, ['data da guarda', 'data'], dataEntrada);
    definirValorLinhaFlexivel(linha, estrutura, ['descrição', 'descricao'], descricao);
    definirValorLinhaFlexivel(linha, estrutura, ['telefone para contato', 'telefone'], telefone);
    definirValorLinhaFlexivel(linha, estrutura, ['foi encontrado?', 'encontrado'], 'Não');
    if (novoHistorico) {
      definirValorLinhaFlexivel(linha, estrutura, ['histórico de contato', 'historico'], novoHistorico);
    }

    sheet.getRange(sheet.getLastRow() + 1, PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL, 1, estrutura.ultimaColuna)
      .setValues([linha]);

    registrarLog('ACHADOS', 'Novo pertence registrado para ' + (dono || 'sem nome'));
    return { success: true };
  } catch (erro) {
    registrarLog('ERRO', 'Erro ao cadastrar pertence perdido: ' + erro.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function atualizarPertencePerdido(parametros) {
  try {
    var idLinha = parseInt(parametros.idLinha || parametros.id || parametros.linha, 10);
    if (!idLinha || idLinha < PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO) {
      return { success: false, error: 'Linha inválida' };
    }

    var sheet = obterSheetPertencesPerdidos();
    var estrutura = garantirEstruturaPertences(sheet);

    var linhaAtual = sheet.getRange(idLinha, PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL, 1, estrutura.ultimaColuna).getValues()[0] || [];

    if (parametros.dono !== undefined) {
      definirValorLinhaFlexivel(linhaAtual, estrutura, ['dono do pertence', 'dono'], parametros.dono);
    }
    if (parametros.descricao !== undefined) {
      definirValorLinhaFlexivel(linhaAtual, estrutura, ['descrição', 'descricao'], parametros.descricao);
    }
    if (parametros.telefone !== undefined) {
      definirValorLinhaFlexivel(linhaAtual, estrutura, ['telefone para contato', 'telefone'], parametros.telefone);
    }
    if (parametros.dataGuarda !== undefined) {
      var dataAtualizada = parametros.dataGuarda ? new Date(parametros.dataGuarda) : '';
      if (dataAtualizada && isNaN(dataAtualizada.getTime())) {
        dataAtualizada = '';
      }
      definirValorLinhaFlexivel(linhaAtual, estrutura, ['data da guarda', 'data'], dataAtualizada);
    }
    if (parametros.encontrado !== undefined) {
      var encontradoValor = converterParaBoolean(parametros.encontrado) ? 'Sim' : 'Não';
      definirValorLinhaFlexivel(linhaAtual, estrutura, ['foi encontrado?', 'encontrado'], encontradoValor);
    }
    if (parametros.historicoContato !== undefined) {
      definirValorLinhaFlexivel(linhaAtual, estrutura, ['histórico de contato', 'historico'], parametros.historicoContato);
    }

    sheet.getRange(idLinha, PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL, 1, estrutura.ultimaColuna).setValues([linhaAtual]);
    registrarLog('ACHADOS', 'Pertence atualizado na linha ' + idLinha);
    return { success: true };
  } catch (erro) {
    registrarLog('ERRO', 'Erro ao atualizar pertence perdido: ' + erro.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function registrarContatoPertence(parametros) {
  try {
    var idLinha = parseInt(parametros.idLinha || parametros.id || parametros.linha, 10);
    if (!idLinha || idLinha < PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO) {
      return { success: false, error: 'Linha inválida' };
    }

    var anotacao = (parametros.anotacao || '').toString().trim();
    var responsavel = (parametros.usuarioResponsavel || parametros.responsavel || '').toString().trim() || 'Equipe';

    var sheet = obterSheetPertencesPerdidos();
    var estrutura = garantirEstruturaPertences(sheet);
    var range = sheet.getRange(idLinha, PLANILHA_PERTENCES_PERDIDOS_COLUNA_INICIAL, 1, estrutura.ultimaColuna);
    var linhaAtual = range.getValues()[0] || [];

    var historicoAtual = obterValorLinhaFlexivel(linhaAtual, estrutura, ['histórico de contato', 'historico de contato', 'historico'], '');
    var dataTexto = Utilities.formatDate(new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm');
    var entrada = dataTexto + ' - ' + responsavel + (anotacao ? ': ' + anotacao : ': ligação registrada');
    var novoHistorico = historicoAtual ? historicoAtual + '\n' + entrada : entrada;

    definirValorLinhaFlexivel(linhaAtual, estrutura, ['histórico de contato', 'historico de contato', 'historico'], novoHistorico);
    range.setValues([linhaAtual]);

    registrarLog('ACHADOS', 'Contato registrado para pertence na linha ' + idLinha + ' por ' + responsavel);
    return { success: true, historicoContato: novoHistorico };
  } catch (erro) {
    registrarLog('ERRO', 'Erro ao registrar contato de pertence: ' + erro.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function excluirPertencePerdido(parametros) {
  try {
    var idLinha = parseInt(parametros.idLinha || parametros.id || parametros.linha, 10);
    if (!idLinha || idLinha <= PLANILHA_PERTENCES_PERDIDOS_LINHA_CABECALHO) {
      return { success: false, error: 'Linha inválida' };
    }

    var sheet = obterSheetPertencesPerdidos();
    var ultimaLinha = sheet.getLastRow();
    if (idLinha > ultimaLinha) {
      return { success: false, error: 'Registro não encontrado' };
    }

    sheet.deleteRow(idLinha);

    registrarLog('ACHADOS', 'Pertence excluído na linha ' + idLinha);
    return { success: true };
  } catch (erro) {
    registrarLog('ERRO', 'Erro ao excluir pertence perdido: ' + erro.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function registrarLog(acao, detalhes) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return;
    }

    var sheet = ss.getSheetByName('LOGS');
    if (!sheet) {
      try {
        sheet = ss.insertSheet('LOGS');
      } catch (erroCriacao) {
        console.error('Não foi possível criar aba de LOGS:', erroCriacao);
        return;
      }
    }

    if (sheet.getLastColumn() < 5) {
      sheet.insertColumns(sheet.getLastColumn() + 1, 5 - sheet.getLastColumn());
    }

    var cabecalhos = sheet.getRange(1, 1, 1, 5).getValues()[0];
    if (!cabecalhos[0]) {
      sheet.getRange(1, 1, 1, 5).setValues([[
        'Data/Hora',
        'Usuário',
        'Ação',
        'Detalhes',
        'IP'
      ]]);
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      lastRow = 1;
    }

    var dataLog = obterDataHoraAtualFormatada().dataHoraIso;
    var usuarioLog = determinarResponsavelRegistro(usuarioContextoRequisicao) || 'desconhecido';

    var novaLinha = [
      dataLog,
      usuarioLog,
      acao || '',
      detalhes || '',
      ''
    ];

    sheet.getRange(lastRow + 1, 1, 1, 5).setValues([novaLinha]);

  } catch (error) {
    console.error('Falha ao registrar log:', error);
  }
}

function getLogs() {
  try {
    var permissao = validarPermissaoAdmin({ usuarioId: usuarioContextoRequisicaoId });
    if (!permissao.ok) {
      return { success: false, error: permissao.error };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LOGS');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
    var logs = [];
    
    data.forEach(function(row) {
      if (row[0]) {
        logs.push({
          dataHora: row[0],
          usuario: row[1],
          acao: row[2],
          detalhes: row[3],
          ip: row[4]
        });
      }
    });
    
    return { success: true, data: logs.reverse() }; // Mais recentes primeiro
    
  } catch (error) {
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

function obterOuCriarSubpastaDrive(pastaPai, nomePasta) {
  var nomeLimpo = (nomePasta || '').toString().trim();
  if (!nomeLimpo) {
    return pastaPai;
  }

  var pastas = pastaPai.getFoldersByName(nomeLimpo);
  if (pastas.hasNext()) {
    return pastas.next();
  }
  return pastaPai.createFolder(nomeLimpo);
}

function obterOuCriarArquivoBackup(pasta, nomeArquivo) {
  var arquivos = pasta.getFilesByName(nomeArquivo);
  if (arquivos.hasNext()) {
    return SpreadsheetApp.open(arquivos.next());
  }

  var arquivo = SpreadsheetApp.create(nomeArquivo);
  var fileDrive = DriveApp.getFileById(arquivo.getId());
  pasta.addFile(fileDrive);
  try {
    DriveApp.getRootFolder().removeFile(fileDrive);
  } catch (erroRemocao) {
    console.log('Arquivo de backup mantido no Meu Drive por permissão: ' + erroRemocao);
  }

  return arquivo;
}

function obterOuCriarAbaDestino(planilha, nomeAba, cabecalhos) {
  var aba = planilha.getSheetByName(nomeAba);
  if (!aba) {
    aba = planilha.insertSheet(nomeAba);
  }

  if (cabecalhos && cabecalhos.length) {
    var maxColunasDestino = aba.getMaxColumns();
    if (maxColunasDestino < cabecalhos.length) {
      aba.insertColumnsAfter(maxColunasDestino, cabecalhos.length - maxColunasDestino);
    }

    var cabecalhosDestino = aba.getRange(1, 1, 1, cabecalhos.length).getValues()[0];
    var cabecalhosIguais = true;
    for (var i = 0; i < cabecalhos.length; i++) {
      if ((cabecalhosDestino[i] || '').toString() !== (cabecalhos[i] || '').toString()) {
        cabecalhosIguais = false;
        break;
      }
    }

    if (aba.getLastRow() === 0 || !cabecalhosIguais) {
      aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    }
  }

  return aba;
}

function copiarDepoisLimparAba(abaOrigem, abaDestino) {
  if (!abaOrigem || !abaDestino) {
    return 0;
  }

  var ultimaLinha = abaOrigem.getLastRow();
  var ultimaColuna = abaOrigem.getLastColumn();
  if (ultimaLinha < 2 || ultimaColuna < 1) {
    return 0;
  }

  var quantidade = ultimaLinha - 1;
  var dados = abaOrigem.getRange(2, 1, quantidade, ultimaColuna).getValues();
  if (!dados.length) {
    return 0;
  }

  var maxColunasDestino = abaDestino.getMaxColumns();
  if (maxColunasDestino < ultimaColuna) {
    abaDestino.insertColumnsAfter(maxColunasDestino, ultimaColuna - maxColunasDestino);
  }

  var destinoLinha = Math.max(abaDestino.getLastRow(), 1) + 1;
  abaDestino.getRange(destinoLinha, 1, dados.length, ultimaColuna).setValues(dados);
  abaOrigem.getRange(2, 1, quantidade, ultimaColuna).clearContent();

  return dados.length;
}

function executarBackupSistema() {
  try {
    var permissao = validarPermissaoAdmin({ usuarioId: usuarioContextoRequisicaoId });
    if (!permissao.ok) {
      return { success: false, error: permissao.error };
    }

    var raiz = DriveApp.getFolderById(PASTA_BACKUP_RAIZ_ID);
    var agora = new Date();
    var timezone = Session.getScriptTimeZone() || 'America/Fortaleza';
    var ano = Utilities.formatDate(agora, timezone, 'yyyy');
    var mesNumero = Utilities.formatDate(agora, timezone, 'MM');
    var mesAno = Utilities.formatDate(agora, timezone, 'yyyy-MM');
    var timestamp = Utilities.formatDate(agora, timezone, "yyyy-MM-dd'T'HH:mm:ss");

    var pastaAno = obterOuCriarSubpastaDrive(raiz, ano);
    var pastaMesGeral = obterOuCriarSubpastaDrive(pastaAno, mesNumero + '-' + mesAno);
    var pastaLogs = obterOuCriarSubpastaDrive(raiz, 'LOGS');
    var pastaTermos = obterOuCriarSubpastaDrive(raiz, 'TERMOS');
    var pastaLogsAno = obterOuCriarSubpastaDrive(pastaLogs, ano);
    var pastaLogsMes = obterOuCriarSubpastaDrive(pastaLogsAno, mesNumero + '-' + mesAno);
    var pastaTermosAno = obterOuCriarSubpastaDrive(pastaTermos, ano);
    var pastaTermosMes = obterOuCriarSubpastaDrive(pastaTermosAno, mesNumero + '-' + mesAno);

    var arquivoBackupGeral = obterOuCriarArquivoBackup(pastaMesGeral, 'BACKUP-GERAL-' + mesAno);
    var arquivoBackupLogs = obterOuCriarArquivoBackup(pastaLogsMes, 'BACKUP-LOGS-' + mesAno);
    var arquivoBackupTermos = obterOuCriarArquivoBackup(pastaTermosMes, 'BACKUP-TERMOS-' + mesAno);

    var ssPrincipal = SpreadsheetApp.getActiveSpreadsheet();
    var relatorio = [];
    var totalRegistros = 0;

    var mapeamento = [
      { origem: 'Histórico Visitantes', arquivo: arquivoBackupGeral },
      { origem: 'Histórico Acompanhantes', arquivo: arquivoBackupGeral },
      { origem: 'Registro de Imagens', arquivo: arquivoBackupGeral },
      { origem: 'LOGS', arquivo: arquivoBackupLogs },
      { origem: 'Termos de Responsabilidade', arquivo: arquivoBackupTermos }
    ];

    mapeamento.forEach(function(item) {
      try {
        var abaOrigem = ssPrincipal.getSheetByName(item.origem);
        if (!abaOrigem) {
          relatorio.push({ aba: item.origem, movidos: 0, observacao: 'Aba não encontrada.' });
          return;
        }

        var ultimaColuna = abaOrigem.getLastColumn();
        if (ultimaColuna < 1) {
          relatorio.push({ aba: item.origem, movidos: 0, observacao: 'Aba sem colunas.' });
          return;
        }

        var cabecalhos = abaOrigem.getRange(1, 1, 1, ultimaColuna).getValues()[0];
        var abaDestino = obterOuCriarAbaDestino(item.arquivo, item.origem, cabecalhos);
        var movidos = copiarDepoisLimparAba(abaOrigem, abaDestino);
        totalRegistros += movidos;
        relatorio.push({ aba: item.origem, movidos: movidos });
      } catch (erroAba) {
        relatorio.push({ aba: item.origem, movidos: 0, observacao: 'Erro: ' + erroAba.toString() });
      }
    });

    registrarLog('BACKUP_EXECUTADO', 'Backup executado em ' + timestamp + '. Registros movidos: ' + totalRegistros);
    invalidarCachesArmariosRelacionados();

    return {
      success: true,
      totalRegistrosMovidos: totalRegistros,
      timestamp: timestamp,
      relatorio: relatorio
    };
  } catch (error) {
    registrarLog('BACKUP_ERRO', error.toString());
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Funções para notificações
function getNotificacoes() {
  try {
    var agora = new Date();
    var notificacoes = [];

    var armariosVisitantes = getArmarios('visitante', false, false);
    var armariosAcompanhantes = getArmarios('acompanhante', false, false);

    // Verificar armários vencidos e próximos do vencimento (visitantes e acompanhantes)
    [armariosVisitantes, armariosAcompanhantes].forEach(function(resultado) {
      if (resultado && resultado.success) {
        resultado.data.forEach(function(armario) {
          if (armario.status === 'em-uso' && armario.horaPrevista) {
            try {
              var hoje = new Date().toISOString().split('T')[0];
              var horaPrevista = new Date(hoje + 'T' + armario.horaPrevista + ':00');
              var diferencaMinutos = (horaPrevista - agora) / (1000 * 60);

              if (diferencaMinutos < 0) {
                notificacoes.push({
                  tipo: 'danger',
                  titulo: `Armário ${armario.numero} vencido`,
                  tempo: `Há ${Math.abs(Math.round(diferencaMinutos))} minutos`
                });
              } else if (diferencaMinutos <= 10) {
                notificacoes.push({
                  tipo: 'warning',
                  titulo: `Armário ${armario.numero} próximo do horário`,
                  tempo: `Há ${Math.round(diferencaMinutos)} minutos`
                });
              }
            } catch (e) {
              // Ignora erro de parsing de data
            }
          }
        });
      }
    });

    // Aviso para movimentação de contingência quando houver armários livres
    if (armariosAcompanhantes && armariosAcompanhantes.success) {
      var contigencias = armariosAcompanhantes.data.filter(function(item) {
        return normalizarTextoBasico(item.status) === 'contingencia';
      });
      var livres = armariosAcompanhantes.data.filter(function(item) {
        return normalizarTextoBasico(item.status) === 'livre';
      });

      if (contigencias.length && livres.length) {
        notificacoes.unshift({
          tipo: 'warning',
          titulo: 'Há armários livres para mover contingências',
          tempo: `${livres.length} livre(s) para ${contigencias.length} contingência(s)`
        });
      }
    }

    return { success: true, data: notificacoes };

  } catch (error) {
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Função para obter estatísticas do Monitor
function getEstatisticasDashboard(tipoUsuario) {
  try {
    var estatisticas = {
      livres: 0,
      emUso: 0,
      proximo: 0,
      vencidos: 0
    };

    var tipos = [];

    // Definir quais tipos de armário o usuário pode ver
    var perfil = normalizarTextoBasico(tipoUsuario);
    if (!perfil) {
      perfil = 'geral';
    }

    if (perfil === 'admin' || perfil === 'ambos' || perfil === 'geral' || perfil === 'todos') {
      tipos = ['visitante', 'acompanhante'];
    } else if (perfil === 'visitante') {
      tipos = ['visitante'];
    } else if (perfil === 'acompanhante') {
      tipos = ['acompanhante'];
    }

    var agora = new Date();

    tipos.forEach(function(tipo) {
      var armarios = getArmarios(tipo, false, false);
      if (armarios.success) {
        armarios.data.forEach(function(armario) {
          if (armario.status === 'livre') {
            estatisticas.livres++;
          } else if (armario.status === 'em-uso') {
            if (armario.horaPrevista) {
              try {
                var hoje = new Date().toISOString().split('T')[0];
                var horaPrevista = new Date(hoje + 'T' + armario.horaPrevista + ':00');
                var diferencaMinutos = (horaPrevista - agora) / (1000 * 60);
                
                if (diferencaMinutos < 0) {
                  estatisticas.vencidos++;
                } else if (diferencaMinutos <= 10) {
                  estatisticas.proximo++;
                } else {
                  estatisticas.emUso++;
                }
              } catch (e) {
                estatisticas.emUso++;
              }
            } else {
              estatisticas.emUso++;
            }
          }
        });
      }
    });
    
    return { success: true, data: estatisticas };
    
  } catch (error) {
    return { success: false, error: 'Erro ao processar a solicitação.' };
  }
}

// Função para verificar se o sistema está inicializado
function verificarInicializacao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = [
    'Histórico Visitantes', 
    'Histórico Acompanhantes', 
    'Visitantes', 
    'Acompanhantes', 
    'Cadastro Armários', 
    'Unidades', 
    'Usuários', 
    'LOGS',
    'Termos de Responsabilidade',
    'Movimentações'
  ];
  
  for (var i = 0; i < abas.length; i++) {
    if (!ss.getSheetByName(abas[i])) {
      return { success: true, inicializado: false };
    }
  }
  
  return { success: true, inicializado: true };
}

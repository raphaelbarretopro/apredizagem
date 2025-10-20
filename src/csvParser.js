/*
  src/csvParser.js
  -----------------
  Wrapper simples em torno do PapaParse com configurações apropriadas para o
  CSV de "Ativos" usado pelo sistema. Expondo `AppCSV.parseAtivos` que
  recebe o texto do CSV e retorna um array de objetos (header-aware).

  Observações:
  - O delimitador usado neste projeto é ponto-e-vírgula (`;`) — comum em
    CSVs gerados em ambientes PT-BR.
  - O módulo assume que a biblioteca PapaParse foi carregada no index.html.
*/
// src/csvParser.js
// Este arquivo contém um wrapper mínimo para o PapaParse, com foco
// no CSV "Ativos" usado pelo sistema. Abaixo comentamos cada linha
// para facilitar o entendimento do fluxo.

(function(global){
  // IIFE (Immediately Invoked Function Expression)
  // - Objetivo: encapsular o código e expor apenas o que precisamos no
  //   escopo global, evitando poluir o namespace.

  // parseAtivos
  // - Função pública que recebe o texto do CSV e retorna um array de
  //   objetos onde cada objeto representa uma linha, indexado pelos
  //   nomes das colunas (header-aware).
  // Linha por linha explicada:
  // -----------------------------------------
  // declaração da função que recebe o conteúdo do CSV como string
  function parseAtivos(csvText) {
    // Verifica se a biblioteca PapaParse está disponível no ambiente
    // (ela é carregada via CDN no index.html). Se não estiver, lançamos
    // um erro claro para ajudar no diagnóstico.
    if (typeof Papa === 'undefined') throw new Error('PapaParse não encontrado');

    // Chamada ao Papa.parse para transformar o CSV em objetos JS.
    // - csvText: a string completa do CSV lida do arquivo.
    // - { header: true }: instrui o PapaParse a usar a primeira linha
    //   como cabeçalho e mapear cada linha subsequente como objeto.
    // - skipEmptyLines: ignora linhas vazias para evitar objetos
    //   desnecessários no resultado.
    // - delimiter: ';' é o separador utilizado nos CSVs gerados em
    //   ambientes PT-BR (ponto-e-vírgula). Manter isto evita parsing
    //   incorreto quando o CSV não usa vírgula.
    const data = Papa.parse(csvText, { header: true, skipEmptyLines: true, delimiter: ';' }).data;

    // Retorna apenas a propriedade `data` do resultado do PapaParse,
    // que é o array de objetos (linhas do CSV mapeadas pelos headers).
    return data;
  }

  // Exporta o parser no objeto `AppCSV` no escopo global para que
  // outras partes do aplicativo (ex: UI/processor) possam chamar
  // `AppCSV.parseAtivos(csvText)` sem depender de módulos.
  // A construção `global.AppCSV = global.AppCSV || {}` preserva qualquer
  // outro conteúdo já exposto em `AppCSV`.
  global.AppCSV = global.AppCSV || {};
  global.AppCSV.parseAtivos = parseAtivos;

// Fecha a IIFE passando o objeto global `window` como parâmetro.
})(window);

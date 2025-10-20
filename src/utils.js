/*
  src/utils.js
  -----------------
  Coleção de funções utilitárias puras (sem efeitos colaterais) usadas
  por outros módulos do projeto. Estas funções cuidam de tarefas comuns
  como normalizar o RA (registro acadêmico), sanitizar nomes de arquivos
  e limpar textos vindos de arquivos CSV/XLS que podem ter problemas de
  codificação.

  Comentários neste arquivo explicam o propósito e o comportamento de cada
  função. As funções expostas ficam em `window.AppUtils` para serem
  acessíveis globalmente pelo código sem uso de módulos ES.
*/
(function(global){
  // normalizeRa
  // - Entrada: `raw` pode ser string, número ou qualquer valor que represente o RA
  // - Saída: número inteiro correspondente ao RA sem zeros à esquerda, ou null se inválido
  // Objetivo: padronizar o RA para comparação cruzada entre CSV e planilha.
  function normalizeRa(raw) {
    // Se o valor for undefined ou null, não há RA válido
    if (raw === undefined || raw === null) return null;
    // Converte para string, remove qualquer caractere não numérico, remove zeros à esquerda e trim
    const s = String(raw).replace(/\D/g, '').replace(/^0+/, '').trim();
    // Se, depois de limpar, a string estiver vazia, não há RA válido
    if (!s) return null;
    // Converte a string para inteiro
    const n = parseInt(s, 10);
    // Se parseInt falhar, retorna null, senão retorna o número
    return Number.isNaN(n) ? null : n;
  }

  // sanitizeFilename
  // - Entrada: string que será transformada em nome seguro para arquivo
  // - Saída: string sanitizada, sem caracteres especiais, com espaços substituídos por '_' e tamanho limitado
  function sanitizeFilename(s) {
    // Valores nulos retornam string vazia (evita erros downstream)
    if (s === undefined || s === null) return '';
    // Passo a passo:
    // 1) trim -> remove espaços nas extremidades
    // 2) replace(/\s+/g, '_') -> substitui espaços por underline
    // 3) replace(/[^a-zA-Z0-9_\-\.]/g, '') -> remove caracteres não permitidos
    // 4) substring(0,60) -> limita o comprimento a 60 caracteres
    return String(s)
      .trim()
      .replace(/\s+/g, '_')
      .replace(/[^a-zA-Z0-9_\-\.]/g, '')
      .substring(0, 60);
  }

  // extractDayFromHeader
  // - Recebe uma célula (objeto SheetJS) e tenta extrair o número do dia
  // - Retorna string vazia se não for possível extrair
  function extractDayFromHeader(cell) {
    // Se a célula for nula/indefinida, retorna string vazia
    if (!cell) return '';
    // Preferimos a propriedade `w` (display text) se disponível; caso contrário usamos `v` (raw value)
    const text = (cell.w !== undefined) ? String(cell.w) : String(cell.v);
    // Procura o primeiro grupo de 1 ou 2 dígitos no texto
    const m = text.match(/(\d{1,2})/);
    // Se encontrar, converte para inteiro e retorna como string (sem zeros à esquerda)
    return m ? String(parseInt(m[1], 10)) : '';
  }

  // sanitizeText
  // - Entrada: qualquer valor (geralmente string) vindo de CSV/XLS
  // - Saída: string 'limpa' pronta para exibição ou escrita em planilha
  // Passos realizados:
  // 1) Normaliza Unicode para NFC (reduz variação de composições de acentos)
  // 2) Remove o caractere de substituição '�' (aparece quando houve falha de encoding)
  // 3) Remove caracteres de controle indesejáveis
  // 4) Colapsa múltiplos espaços em um único espaço e faz trim
  function sanitizeText(raw) {
    if (raw === undefined || raw === null) return '';
    try {
      // Garante que temos uma string para trabalhar
      let s = String(raw);
      // Normaliza Unicode (quando disponível) para reduzir problemas com acentuação combinada
      if (typeof s.normalize === 'function') s = s.normalize('NFC');
      // Remove o caractere de substituição '�' resultante de decodificação incorreta
      s = s.replace(/�/g, '');
      // Remove caracteres de controle (exceto newlines e tabs) que podem quebrar arquivos
      s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
      // Colapsa múltiplos espaços e trima
      s = s.replace(/\s+/g, ' ').trim();
      return s;
    } catch (e) {
      // Se qualquer coisa falhar, retornamos a versão string simples do valor
      return String(raw);
    }
  }

  // Exporta as funções como AppUtils no escopo global (compatibilidade com app sem bundler)
  global.AppUtils = global.AppUtils || {};
  global.AppUtils.normalizeRa = normalizeRa;
  global.AppUtils.sanitizeFilename = sanitizeFilename;
  global.AppUtils.extractDayFromHeader = extractDayFromHeader;
  global.AppUtils.sanitizeText = sanitizeText;
})(window);

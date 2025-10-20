// src/exporter.js - encapsula a criação do workbook e download (ExcelJS)
(function(global){
  // Definição da função pública buildAndDownloadWorkbook.
  // Recebe um objeto com várias opções e retorna o nome do ficheiro
  // gerado (após iniciar o download).
  async function buildAndDownloadWorkbook({ saida, perEmpresaReport, monthYearText, turmaFiltro, empresaFiltro, companyText, debug }) {
    // Cria um novo workbook em memória usando ExcelJS.
    // ExcelJS permite montar planilhas com estilos, mesclagens e depois
    // gerar um buffer que podemos transformar em Blob para download.
    const wb = new ExcelJS.Workbook();

    // Cria/seleciona uma worksheet chamada 'Relatório'
    const ws = wb.addWorksheet('Relatório');

    // Define as colunas com cabeçalhos diferentes conforme o modo
    // perEmpresaReport (quando a intenção é gerar relatório por empresa).
    if (perEmpresaReport) {
      // Quando for por empresa, incluímos a coluna 'EMPRESA'
      ws.columns = [
        { header: 'TURMA', key: 'turma' },
        { header: 'ALUNO', key: 'aluno' },
        { header: 'EMPRESA', key: 'empresa' },
        { header: 'PRÁTICA', key: 'pratica' },
        { header: 'CURSO', key: 'curso' },
        { header: 'Faltas Justificadas (dias)', key: 'faltasJustDays' },
        { header: 'Nº Faltas Justificadas', key: 'faltasJustCount' },
        { header: 'Faltas Não Justificadas (dias)', key: 'faltasNaoJustDays' },
        { header: 'Nº Faltas Não Justificadas', key: 'faltasNaoJustCount' },
        { header: 'Atrasos (dias)', key: 'atrasosDays' },
        { header: 'Nº Horas de Atraso', key: 'horasAtraso' },
        { header: 'Total Horas de ausência no curso', key: 'totalAusenciaHoras' }
      ];
    } else {
      // Modo padrão sem coluna de empresa
      ws.columns = [
        { header: 'TURMA', key: 'turma' },
        { header: 'ALUNO', key: 'aluno' },
        { header: 'PRÁTICA', key: 'pratica' },
        { header: 'CURSO', key: 'curso' },
        { header: 'Faltas Justificadas (dias)', key: 'faltasJustDays' },
        { header: 'Nº Faltas Justificadas', key: 'faltasJustCount' },
        { header: 'Faltas Não Justificadas (dias)', key: 'faltasNaoJustDays' },
        { header: 'Nº Faltas Não Justificadas', key: 'faltasNaoJustCount' },
        { header: 'Atrasos (dias)', key: 'atrasosDays' },
        { header: 'Nº Horas de Atraso', key: 'horasAtraso' },
        { header: 'Total Horas de ausência no curso', key: 'totalAusenciaHoras' }
      ];
    }

    // Ajuste de larguras das colunas para uma aparência mais legível.
    // Definimos duas matrizes de larguras dependendo se há coluna 'EMPRESA'.
    const colWidths = perEmpresaReport
      ? [24, 60, 30, 40, 28, 30, 12, 30, 12, 22, 14, 20]
      : [24, 70, 40, 28, 30, 12, 30, 12, 22, 14, 20];

    // Aplica as larguras às colunas já definidas. Se faltar alguma
    // largura, usamos 15 como padrão.
    ws.columns.forEach((col, idx) => { col.width = colWidths[idx] || 15; });

    // Construção de cabeçalho visual do relatório (linhas superiores)
    // Mesclamos células e colocamos títulos grandes para identificação.
    const lastCol = ws.columns.length; // número de colunas para mesclar

    // Linha 1: nome da unidade
    ws.mergeCells(1, 1, 1, lastCol);
    const topCell = ws.getCell(1, 1);
    topCell.value = 'SENAI - MARACANÃ';
    topCell.alignment = { vertical: 'middle', horizontal: 'center' };
    topCell.font = { bold: true, size: 16 };

    // Linha 2: programa
    ws.mergeCells(2, 1, 2, lastCol);
    const progCell = ws.getCell(2, 1);
    progCell.value = 'PROGRAMA DE APRENDIZAGEM INDUSTRIAL';
    progCell.alignment = { vertical: 'middle', horizontal: 'center' };
    progCell.font = { bold: true, size: 12 };

    // Linha 3: título que inclui o mês/ano ou empresa (quando aplicável)
    ws.mergeCells(3, 1, 3, lastCol);
    const titleCell = ws.getCell(3, 1);
    if (perEmpresaReport) {
      titleCell.value = `Relatório de Frequência - Aprendizes - ${monthYearText || '***MÊS/ANO***'}`;
    } else {
      titleCell.value = `Relatório de Frequência - Aprendizes - ${monthYearText || '***MÊS/ANO***'} - Empresa: ${companyText || '***EMPRESA***'}`;
    }
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    titleCell.font = { bold: true, size: 14 };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };

    // Linha 4: cabeçalho da tabela com os nomes das colunas. Aqui aplicamos
    // estilo básico (bold, alinhamento central e cor de fundo) e altura.
    const headerRow = ws.getRow(4);
    headerRow.values = ws.columns.map(c => c.header);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    headerRow.height = 40;
    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });

    // Insere as linhas de dados já preparadas (saida = array de objetos).
    // ExcelJS irá casar as chaves do objeto com os `key` das colunas.
    ws.addRows(saida);

    // Formatação das linhas de dados: alternância de cor, alinhamento
    // por coluna e bordas. Percorremos do índice 5 (primeira linha de
    // dados) até a última linha existente.
    const firstDataRow = 5;
    const lastDataRow = ws.lastRow.number;
    for (let r = firstDataRow; r <= lastDataRow; r++) {
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // cor de fundo alternada para facilitar leitura
        if ((r - firstDataRow) % 2 === 0) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCF2F2' } };
        }
        // coluna 5 (índice 5) contém os dias de faltas — alinhar à esquerda
        if (colNumber === 5) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
        // bordas finas em todas as células para um visual tabular
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
      // altura de linha padrão para legibilidade
      row.height = 20;
    }

    // Construção do nome do ficheiro seguindo a regra:
    // relatorio_frequencia_MES-ANO_CODTURMA[_NOMEEMPRESA].xlsx
    // Usamos AppUtils.sanitizeFilename quando disponível para garantir
    // que o nome seja seguro para o sistema de ficheiros.
    const sanitize = (s) => (typeof AppUtils !== 'undefined' && AppUtils.sanitizeFilename) ? AppUtils.sanitizeFilename(s) : String(s||'').trim().replace(/\s+/g,'_').replace(/[^a-zA-Z0-9_\-\.]/g,'').substring(0,60);
    let mesAno = '';
    if (monthYearText && typeof monthYearText === 'string') mesAno = monthYearText.replace('/', '-');
    const turmaPart = sanitize(turmaFiltro || 'SEM_TURMA');
    const mesAnoPart = sanitize(mesAno || 'SEM_MES-ANO');
    const includeEmpresa = !!(empresaFiltro && String(empresaFiltro).trim());
    const filename = includeEmpresa
      ? `relatorio_frequencia_${mesAnoPart}_${turmaPart}_${sanitize(companyText)}.xlsx`
      : `relatorio_frequencia_${mesAnoPart}_${turmaPart}.xlsx`;

    // Log do nome final quando em modo debug para ajudar no diagnóstico
    if (debug) console.log('[DEBUG] Exporter final filename:', filename);

    // Geração do buffer do workbook (xlsx) e criação de Blob para
    // iniciar o download no navegador.
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);

    // Cria um elemento <a> temporário para disparar o download do arquivo
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    // Retornamos o nome do ficheiro gerado para quem chamou a função
    return filename;
  }

  // Expondo a função no objeto global AppExporter para uso externo
  global.AppExporter = global.AppExporter || {};
  global.AppExporter.buildAndDownloadWorkbook = buildAndDownloadWorkbook;

})(window);

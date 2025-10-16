// app.js - Processamento de arquivos no navegador (VERSÃO FINAL COM LÓGICA DE MAPEAMENTO CORRIGIDA)

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsText(file, 'utf-8');
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

let ativosDataGlobal = [];

// Atualiza empresas conforme turma selecionada
function atualizarEmpresasPorTurma(turmaSelecionada) {
  let empresas = [];
  if (turmaSelecionada) {
    empresas = [...new Set(ativosDataGlobal.filter(a => a.CODTURMA && a.CODTURMA.trim() === turmaSelecionada)
      .map(a => a.NOMEEMPRESA_NOVO ? a.NOMEEMPRESA_NOVO.trim() : '').filter(Boolean))].sort();
  } else {
    empresas = [...new Set(ativosDataGlobal.map(a => a.NOMEEMPRESA_NOVO ? a.NOMEEMPRESA_NOVO.trim() : '').filter(Boolean))].sort();
  }
  const empresaSelect = document.getElementById('empresaSelect');
  empresaSelect.innerHTML = '<option value="">Todas</option>' + empresas.map(e => `<option value="${e}">${e}</option>`).join('');
}

document.getElementById('ativosCsv').addEventListener('change', async function (e) {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const ativosText = await readFileAsText(file);
    ativosDataGlobal = Papa.parse(ativosText, { header: true, skipEmptyLines: true, delimiter: ';' }).data;
    const turmas = [...new Set(ativosDataGlobal
      .map(a => a.CODTURMA ? a.CODTURMA.trim() : '')
      .filter(Boolean)
      .filter(t => t.toUpperCase().startsWith('APR'))
    )].sort();
    const turmaSelect = document.getElementById('turmaSelect');
    turmaSelect.innerHTML = '<option value="">Todas</option>' + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
    atualizarEmpresasPorTurma('');
    document.getElementById('output').innerHTML = `<p>Foram encontradas um total de ${turmas.length} turmas no Programa de Aprendizagem Industrial.</p>`;
  } catch (error) {
    console.error("Erro ao ler o arquivo CSV:", error);
    document.getElementById('output').innerHTML = `<p style="color:red"><b>Erro ao processar o arquivo CSV:</b> ${error.message}</p>`;
  }
});

document.getElementById('turmaSelect').addEventListener('change', function (e) {
  const turmaSelecionada = e.target.value.trim();
  atualizarEmpresasPorTurma(turmaSelecionada);
});

document.getElementById('freqForm').addEventListener('submit', async function (e) {
  e.preventDefault();
  const ativosFile = document.getElementById('ativosCsv').files[0];
  const freqFile = document.getElementById('freqXls').files[0];
  const turmaFiltro = document.getElementById('turmaSelect').value.trim();
  const empresaFiltro = document.getElementById('empresaSelect').value.trim();

  if (!ativosFile || !freqFile) {
    alert('Selecione ambos os arquivos.');
    return;
  }
  document.getElementById('output').innerHTML = 'Processando arquivos...';

  try {
    const ativosText = await readFileAsText(ativosFile);
    let ativosFiltrados = Papa.parse(ativosText, { header: true, skipEmptyLines: true, delimiter: ';' }).data;

    if (turmaFiltro) {
      ativosFiltrados = ativosFiltrados.filter(a => a.CODTURMA && a.CODTURMA.trim() === turmaFiltro);
    }
    if (empresaFiltro) {
      ativosFiltrados = ativosFiltrados.filter(a => a.NOMEEMPRESA_NOVO && a.NOMEEMPRESA_NOVO.trim() === empresaFiltro);
    }

    if (ativosFiltrados.length === 0) {
      document.getElementById('output').innerHTML = '<p style="color:red; font-weight: bold;">Nenhum aluno encontrado após a aplicação dos filtros.</p>';
      return;
    }

    const alunosMap = {};
    ativosFiltrados.forEach(a => {
      let raNum = null;
      if (a.REGISTRO_ACADEMICO) {
        const raStr = String(a.REGISTRO_ACADEMICO).replace(/\D/g, '').replace(/^0+/, '').trim();
        if (raStr) raNum = parseInt(raStr, 10);
      }
      if (raNum) {
        // incluir NOMEEMPRESA_NOVO e TIPO_PRATICA para poder emitir relatório por empresa/prática quando solicitado
        alunosMap[raNum] = {
          ALUNO: a.ALUNO,
          CODTURMA: a.CODTURMA,
          DTINICIAL: a.DTINICIAL,
          DTFINAL: a.DTFINAL,
          EMPRESA: a.NOMEEMPRESA_NOVO,
          PRATICA: a.TIPO_PRATICA,
          CURSO: a.MATRIZ
        };
      }
    });

    // **CORREÇÃO DEFINITIVA 1: Criar uma lista dos RAs que precisamos encontrar**
    const rasFromCsv = new Set(Object.keys(alunosMap).map(ra => parseInt(ra, 10)));

    const freqBuffer = await readFileAsArrayBuffer(freqFile);
    const workbook = XLSX.read(freqBuffer, { type: 'array', cellStyles: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const frequencias = {};
    const range = XLSX.utils.decode_range(sheet['!ref']);

    // **CORREÇÃO: Os RAs estão na COLUNA B (índice 1), começando da LINHA 15 (índice 14)**
    const linhaAlunos = {}; // Mapeia linha -> RA
    const raCol = 1; // Coluna B
    const nameCol = 2; // Coluna C
    const firstStudentRow = 14; // Linha 15 no Excel (índice 14)

    for (let R = firstStudentRow; R <= range.e.r; ++R) {
      const raCellAddress = XLSX.utils.encode_cell({ c: raCol, r: R });
      const raCell = sheet[raCellAddress];

      if (raCell && raCell.v) {
        const raStr = String(raCell.v).replace(/\D/g, '').replace(/^0+/, '').trim();
        if (raStr) {
          const raNum = parseInt(raStr, 10);
          const nameCellAddress = XLSX.utils.encode_cell({ c: nameCol, r: R });
          const nameCell = sheet[nameCellAddress];
          const studentName = (nameCell && nameCell.v) ? String(nameCell.v).trim() : '';

          if (rasFromCsv.has(raNum)) {
            // inicializa contadores e listas de dias
            frequencias[raNum] = {
              nome: studentName,
              faltasJust: 0,
              faltasNaoJust: 0,
              atrasos: 0,
              faltasJustDays: [],        // dias (número do dia) das faltas justificadas
              faltasNaoJustDays: [],     // dias das faltas não justificadas
              atrasosDays: []            // dias com atrasos
            };
            linhaAlunos[R] = raNum;
          }
        }
      }
    }

    // **CORREÇÃO: Agora iteramos pelas LINHAS de cada aluno, não por colunas**
    // As colunas representam as DATAS (começando por volta da coluna E - índice 4)
    const dateRow = 13; // Linha 14 no Excel contém os cabeçalhos com as datas
    const firstDateCol = 4; // Coluna E (índice 4) é onde começam as presenças
    let linhasProcessadas = 0;

    // Identifica colunas de datas válidas
    const dateHeaderRow = 13; // Linha 14 no Excel (índice 13)
    let validDateCols = [];
    for (let C = firstDateCol; C <= range.e.c; ++C) {
      const headerCellAddress = XLSX.utils.encode_cell({ c: C, r: dateHeaderRow });
      const headerCell = sheet[headerCellAddress];
      if (headerCell && headerCell.v) {
        const headerText = String(headerCell.v).trim().toUpperCase();
        if (headerText !== 'AULAS VS PRESENÇA' && headerText !== 'FALTAS NO MÊS') {
          validDateCols.push(C);
        }
      }
    }

    // Mapear coluna->dia (extraindo o dia do cabeçalho) para usar ao registrar dias
    const dateColToDay = {};
    const extractDayFromHeader = (cell) => {
      if (!cell) return '';
      // Prefer the formatted text if available
      const text = (cell.w !== undefined) ? String(cell.w) : String(cell.v);
      // Buscar primeiro número (dia) no texto, normalmente formatado como DD or DD/MM/YYYY
      const m = text.match(/(\d{1,2})/);
      return m ? String(parseInt(m[1], 10)) : '';
    };
    for (const C of validDateCols) {
      const headerCellAddress = XLSX.utils.encode_cell({ c: C, r: dateHeaderRow });
      const headerCell = sheet[headerCellAddress];
      dateColToDay[C] = extractDayFromHeader(headerCell);
    }

      // Extrair Mês/Ano: PRIORITÁRIO -> célula mesclada AK5:AN5 (linha 5). Se vazio, usar pesquisas anteriores (PERÍODO COMP etc.)
      let monthYearText = '';
      try {
        // Colunas AK..AN correspondem a índices 36..39 (A=0)
        const mergedRow = 4; // linha 5 -> índice 4
        const mergedCols = [36, 37, 38, 39];
        let mergedText = '';
        for (const mc of mergedCols) {
          const mAddr = XLSX.utils.encode_cell({ c: mc, r: mergedRow });
          const mCell = sheet[mAddr];
          if (mCell && (mCell.w !== undefined || mCell.v !== undefined)) {
            mergedText += (mCell.w !== undefined) ? String(mCell.w) : String(mCell.v);
            // add a space to separate parts if multiple cells
            mergedText += ' ';
          }
        }
        mergedText = mergedText.trim();
        if (mergedText) {
          // tentar extrair MM/YYYY ou dd/mm/yyyy a partir do texto
          const mMatch1 = mergedText.match(/(\d{1,2}\/\d{4})/);
          const mMatch2 = mergedText.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
          if (mMatch1) {
            const mm = String(mMatch1[1].split('/')[0]).padStart(2, '0');
            const yy = mMatch1[1].split('/')[1];
            monthYearText = `${mm}/${yy}`;
          } else if (mMatch2) {
            const parts = mMatch2[1].split('/');
            const mm = String(parts[1]).padStart(2, '0');
            const yy = parts[2];
            monthYearText = `${mm}/${yy}`;
          } else {
            // fallback: if text contains a month name and year, try to map month names (pt-BR)
            const months = {
              'janeiro': '01','fevereiro':'02','março':'03','marco':'03','abril':'04','maio':'05','junho':'06',
              'julho':'07','agosto':'08','setembro':'09','outubro':'10','novembro':'11','dezembro':'12'
            };
            const lower = mergedText.toLowerCase();
            const yearMatch = lower.match(/(20\d{2}|19\d{2})/);
            if (yearMatch) {
              const y = yearMatch[1];
              for (const [name, num] of Object.entries(months)) {
                if (lower.indexOf(name) !== -1) {
                  monthYearText = `${num}/${y}`;
                  break;
                }
              }
            }
            // if still empty, just assign the raw trimmed text as last resort
            if (!monthYearText) monthYearText = mergedText;
          }
        }
      } catch (e) {
        // se algo falhar, manter monthYearText vazio e permitir os fallbacks existentes
        monthYearText = monthYearText || '';
      }
      if (!monthYearText) {
        outerLoop:
        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ c: c, r: r });
          const cell = sheet[addr];
          if (cell && cell.v && /PER[IÍ]ODO\s*COMP/i.test(String(cell.v))) {
            const leftAddr = XLSX.utils.encode_cell({ c: c - 1, r: r });
            const rightAddr = XLSX.utils.encode_cell({ c: c + 1, r: r });
            const leftCell = sheet[leftAddr];
            const rightCell = sheet[rightAddr];
            const extractText = (cellObj) => {
              if (!cellObj || (cellObj.v === undefined || cellObj.v === null)) return '';
              // se for número (data serial do Excel), tentar converter
              try {
                if (cellObj.t === 'n' && typeof XLSX !== 'undefined' && XLSX.SSF && typeof XLSX.SSF.parse_date_code === 'function') {
                  const d = XLSX.SSF.parse_date_code(cellObj.v);
                  if (d && d.y) {
                    const dd = String(d.d).padStart(2, '0');
                    const mm = String(d.m).padStart(2, '0');
                    const yyyy = String(d.y);
                    return `${dd}/${mm}/${yyyy}`; // formato dd/mm/yyyy
                  }
                }
              } catch (e) {
                // ignore parse errors
              }
              const text = (cellObj.w !== undefined) ? String(cellObj.w) : String(cellObj.v);
              // procura por padrões tipo DD/MM/YYYY ou MM/YYYY
              const m2 = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
              if (m2) return m2[1];
              const m = text.match(/(\d{1,2}\/\d{4})/);
              if (m) return m[1];
              return text.trim();
            };
            // priorizar célula à esquerda ou direita que contenha uma data/range
            const rawText = extractText(leftCell) || extractText(rightCell) || ((String(cell.v).split(':')[1] || '').trim());
            // tentar extrair a primeira data no formato dd/mm/yyyy
            let my = '';
            if (rawText) {
              const mDate = rawText.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
              if (mDate) {
                const day = mDate[1];
                const month = mDate[2].padStart(2, '0');
                const year = mDate[3];
                my = `${month}/${year}`;
              } else {
                // tentar formato MM/YYYY
                const m2 = rawText.match(/(\d{1,2})\/(\d{4})/);
                if (m2) {
                  const month = m2[1].padStart(2, '0');
                  const year = m2[2];
                  my = `${month}/${year}`;
                } else {
                  my = rawText.trim();
                }
              }
            }
            monthYearText = my;
            break outerLoop;
          }
          }
        }
      }

  // se ainda não encontrou, tentar localizar qualquer range de datas no sheet e extrair mês/ano da primeira
      if (!monthYearText) {
        outerLoop2:
        for (let r2 = range.s.r; r2 <= range.e.r; r2++) {
          for (let c2 = range.s.c; c2 <= range.e.c; c2++) {
            const addr2 = XLSX.utils.encode_cell({ c: c2, r: r2 });
            const cell2 = sheet[addr2];
            if (!cell2 || (cell2.v === undefined || cell2.v === null)) continue;
            // texto formatado preferencialmente
            const txt2 = (cell2.w !== undefined) ? String(cell2.w) : String(cell2.v);
            const rangeMatch = txt2.match(/(\d{1,2}\/\d{1,2}\/\d{4})\s*-\s*(\d{1,2}\/\d{1,2}\/\d{4})/);
            if (rangeMatch) {
              const mDate = rangeMatch[1].match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
              if (mDate) {
                const month = String(mDate[2]).padStart(2, '0');
                const year = mDate[3];
                monthYearText = `${month}/${year}`;
                break outerLoop2;
              }
            }
          }
        }
      }
      // novo fallback: tentar extrair a partir dos cabeçalhos das colunas de data válidas
      if (!monthYearText && Array.isArray(validDateCols) && validDateCols.length) {
        for (const C of validDateCols) {
          const headerCellAddress2 = XLSX.utils.encode_cell({ c: C, r: dateHeaderRow });
          const headerCell2 = sheet[headerCellAddress2];
          if (!headerCell2 || (headerCell2.v === undefined || headerCell2.v === null)) continue;
          const txtH = (headerCell2.w !== undefined) ? String(headerCell2.w) : String(headerCell2.v);
          const mDateH = txtH.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
          if (mDateH) {
            const month = String(mDateH[2]).padStart(2, '0');
            const year = mDateH[3];
            monthYearText = `${month}/${year}`;
            break;
          }
          const m2H = txtH.match(/(\d{1,2})\/(\d{4})/);
          if (m2H) {
            const month = String(m2H[1]).padStart(2, '0');
            const year = m2H[2];
            monthYearText = `${month}/${year}`;
            break;
          }
        }
      }

      // Determinar nome da empresa para o título: preferir o filtro selecionado, senão pegar do CSV filtrado
      let companyText = (empresaFiltro && empresaFiltro.length) ? empresaFiltro : '';
      if (!companyText && Array.isArray(ativosFiltrados) && ativosFiltrados.length) {
        const first = ativosFiltrados.find(x => x.NOMEEMPRESA_NOVO && String(x.NOMEEMPRESA_NOVO).trim());
        if (first && first.NOMEEMPRESA_NOVO) companyText = String(first.NOMEEMPRESA_NOVO).trim();
      }

    for (const R_str in linhaAlunos) {
      const R = parseInt(R_str, 10);
      const raDoAluno = linhaAlunos[R];
      // Percorre apenas as colunas de datas válidas
      for (const C of validDateCols) {
        const freqCellAddress = XLSX.utils.encode_cell({ c: C, r: R });
        const freqCell = sheet[freqCellAddress];
        if (freqCell && freqCell.v) {
          const lowerVal = String(freqCell.v).toLowerCase();
          // Validação robusta para cor de fundo (bgColor)
          let isJustified = false;
          if (freqCell.s && freqCell.s.bgColor) {
            const bg = freqCell.s.bgColor;
            if (bg.rgb || bg.theme !== undefined || bg.indexed !== undefined) {
              isJustified = true;
            }
          }
          // obter dia desta coluna (se disponível)
          const dayForThisCol = dateColToDay[C] || '';
          if (lowerVal === 'f') {
            if (isJustified) {
              frequencias[raDoAluno].faltasJust += 1;
              if (dayForThisCol) frequencias[raDoAluno].faltasJustDays.push(dayForThisCol);
            } else {
              frequencias[raDoAluno].faltasNaoJust += 1;
              if (dayForThisCol) frequencias[raDoAluno].faltasNaoJustDays.push(dayForThisCol);
            }
          } else if (lowerVal === '3') {
            if (!isJustified) {
              frequencias[raDoAluno].atrasos += 1;
              if (dayForThisCol) frequencias[raDoAluno].atrasosDays.push(dayForThisCol);
            }
          } else if (lowerVal === '2') {
            if (!isJustified) {
              frequencias[raDoAluno].atrasos += 2;
              if (dayForThisCol) frequencias[raDoAluno].atrasosDays.push(dayForThisCol);
            }
          } else if (lowerVal === '1') {
            if (!isJustified) {
              frequencias[raDoAluno].atrasos += 3;
              if (dayForThisCol) frequencias[raDoAluno].atrasosDays.push(dayForThisCol);
            }
          }
        }
      }
      linhasProcessadas++;
    }
    // #008000
    // #000000
    // #FFFFFF
    // #339966
    // #FFCC00
    // definir se devemos gerar relatório por empresa (quando turma selecionada e todas as empresas)
    const perEmpresaReport = (turmaFiltro && turmaFiltro.length > 0) && (!empresaFiltro || empresaFiltro.length === 0);

    const saida = [];
    for (const raStr in alunosMap) {
      const ra = parseInt(raStr, 10);
      const alunoInfo = alunosMap[ra];
      const freqInfo = frequencias[ra] || { faltasJust: 0, faltasNaoJust: 0, atrasos: 0 };
  // calcular total de horas de ausência: atrasos (já em horas) + (faltas não justificadas + faltas justificadas) * 4
  const totalAusenciaHoras = (freqInfo.atrasos || 0) + ((freqInfo.faltasNaoJust || 0) + (freqInfo.faltasJust || 0)) * 4;
    const rowObj = {
        mes: (alunoInfo.DTINICIAL || '').split('/')[1] || '',
        inicio: alunoInfo.DTINICIAL,
        termino: alunoInfo.DTFINAL,
        turma: alunoInfo.CODTURMA,
        aluno: alunoInfo.ALUNO,
    pratica: alunoInfo.PRATICA ? String(alunoInfo.PRATICA).trim() : '',
    curso: alunoInfo.CURSO ? String(alunoInfo.CURSO).trim() : '',
        faltasJustCount: freqInfo.faltasJust,
        faltasJustDays: (freqInfo.faltasJustDays && freqInfo.faltasJustDays.length) ? freqInfo.faltasJustDays.join(', ') : '',
        faltasNaoJustCount: freqInfo.faltasNaoJust,
        faltasNaoJustDays: (freqInfo.faltasNaoJustDays && freqInfo.faltasNaoJustDays.length) ? freqInfo.faltasNaoJustDays.join(', ') : '',
        horasAtraso: freqInfo.atrasos,
        atrasosDays: (freqInfo.atrasosDays && freqInfo.atrasosDays.length) ? freqInfo.atrasosDays.join(', ') : '',
        totalAusenciaHoras: totalAusenciaHoras
    };
    // se relatório por empresa, adicione o campo EMPRESA vindo do CSV e use para ordenação posteriormente
    if (perEmpresaReport) {
      rowObj.empresa = alunoInfo.EMPRESA ? String(alunoInfo.EMPRESA).trim() : '';
    }
    saida.push(rowObj);
    }

    // ordenar por empresa quando for relatório por empresa
    if (perEmpresaReport) {
      saida.sort((a, b) => {
        const A = (a.empresa || '').toUpperCase();
        const B = (b.empresa || '').toUpperCase();
        if (A < B) return -1;
        if (A > B) return 1;
        return 0;
      });
    }

    if (saida.length === 0) {
      document.getElementById('output').innerHTML = '<p style="color:red"><b>Nenhum dado cruzado encontrado após a filtragem.</b></p>';
      return;
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Relatório');

    // definir colunas (keys e headers)
    if (perEmpresaReport) {
      ws.columns = [
       // { header: 'MÊS', key: 'mes' },
        //{ header: 'INÍCIO', key: 'inicio' },
       // { header: 'TÉRMINO', key: 'termino' },
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
      ws.columns = [
        //{ header: 'MÊS', key: 'mes' },
       // { header: 'INÍCIO', key: 'inicio' },
        //{ header: 'TÉRMINO', key: 'termino' },
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

  // larguras aproximadas (aumentadas para evitar corte de texto)
  // se relatório por empresa, adiciona largura para a coluna 'Empresa'
  // Ordem: MÊS, INÍCIO, TÉRMINO, TURMA, ALUNO, [EMPRESA?], FaltasJustDays, NºFaltasJust, FaltasNaoJustDays, NºFaltasNaoJust, Atrasos(dias), NºHorasAtraso, TotalHorasAusencia
  const colWidths = perEmpresaReport
    ? [24, 60, 30, 40, 28, 30, 12, 30, 12, 22, 14, 20] // added CURSO width
    : [24, 70, 40, 28, 30, 12, 30, 12, 22, 14, 20]; // added CURSO width
    ws.columns.forEach((col, idx) => {
      col.width = colWidths[idx] || 15;
    });

    // Cabeçalhos fixos do relatório
    const lastCol = ws.columns.length;
    // Linha 1: SENAI - MARACANÃ (centralizado)
    ws.mergeCells(1, 1, 1, lastCol);
    const topCell = ws.getCell(1, 1);
    topCell.value = 'SENAI - MARACANÃ';
    topCell.alignment = { vertical: 'middle', horizontal: 'center' };
    topCell.font = { bold: true, size: 16 };
    ws.getRow(1).height = 20;

    // Linha 2: PROGRAMA DE APRENDIZAGEM INDUSTRIAL (centralizado)
    ws.mergeCells(2, 1, 2, lastCol);
    const progCell = ws.getCell(2, 1);
    progCell.value = 'PROGRAMA DE APRENDIZAGEM INDUSTRIAL';
    progCell.alignment = { vertical: 'middle', horizontal: 'center' };
    progCell.font = { bold: true, size: 12 };
    ws.getRow(2).height = 18;

    // Linha 3: Título do relatório (mês/empresa) - manter estilo amarelo original
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
    ws.getRow(3).height = 22;

    // Cabeçalho de colunas em linha 4
    const headerRow = ws.getRow(4);
    headerRow.values = ws.columns.map(c => c.header);
  headerRow.font = { bold: true };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  headerRow.height = 40; // aumentado para evitar corte de texto nas colunas do cabeçalho
    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    });

    // Adicionar dados começando na linha 5
    ws.addRows(saida);

    // Aplicar estilo nas linhas de dados: cor de preenchimento para linha de estudante e bordas
    const firstDataRow = 5;
    const lastDataRow = ws.lastRow.number;
    for (let r = firstDataRow; r <= lastDataRow; r++) {
      const row = ws.getRow(r);
      // preencher com cor clara para as linhas que contêm aluno (alternar se desejar)
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // alternar cor de fundo para facilitar leitura (linhas pares com leve azul)
        if ((r - firstDataRow) % 2 === 0) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCF2F2' } };
        }
        // alinhamentos
        if (colNumber === 5) { // coluna ALUNO
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
        // bordas em todas as células
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });
      row.height = 20;
    }

    wb.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      // Construir nome dinâmico: relatorio_frequencia_MES-ANO_CODTURMA_NOMEEMPRESA_NOVO
      try {
        const sanitize = s => String(s || '').replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-\.]/g, '').substring(0, 60);
        // monthYearText -> esperamos 'MM/YYYY' ou 'MM/YYYY' -> transformar para 'MM-YYYY'
        let mesAno = '';
        if (monthYearText && typeof monthYearText === 'string') {
          mesAno = monthYearText.replace('/', '-');
        }
        const turmaPart = sanitize(turmaFiltro || 'SEM_TURMA');
        const mesAnoPart = sanitize(mesAno || 'SEM_MES-ANO');
        // Se o usuário escolheu 'Todas' as empresas (empresaFiltro vazio), OMITIR o nome da empresa no filename
        const includeEmpresa = !!(empresaFiltro && String(empresaFiltro).trim());
        const filename = includeEmpresa
          ? `relatorio_frequencia_${mesAnoPart}_${turmaPart}_${sanitize(companyText)}.xlsx`
          : `relatorio_frequencia_${mesAnoPart}_${turmaPart}.xlsx`;
        a.download = filename;
      } catch (e) {
        a.download = 'relatorio_frequencia.xlsx';
      }
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      document.getElementById('output').innerHTML = '<p style="color:green"><b>Relatório gerado com sucesso!</b></p>';
    });

  } catch (error) {
    console.error("Ocorreu um erro CRÍTICO durante o processamento:", error);
    document.getElementById('output').innerHTML = `<p style="color:red"><b>Ocorreu um erro:</b> ${error.message}. Verifique o console do navegador para mais detalhes.</p>`;
  }
});
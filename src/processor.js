// src/processor.js - encapsula a lógica de varredura da planilha e contagem de faltas/atrasos
  (function(global){
    // processSheetFrequency
    // - sheet: objeto SheetJS (uma aba) já lida via XLSX.read
    // - ativosFiltrados: array de objetos vindos do CSV (após filtros)
    // - turmaFiltro / empresaFiltro: filtros usados (informativos)
    // - options: objeto opcional para sobrescrever colunas/linhas padrão
    // Retorna: { alunosMap, frequencias, monthYearText, validDateCols, dateColToDay }
    // Comentário didático:
    // A função cruza o CSV com a planilha por RA (registro acadêmico).
    // Passos principais:
    // 1) Monta um mapa de alunos a partir do CSV (alunosMap)
    // 2) Varre as linhas da planilha procurando RAs que existam no CSV
    // 3) Identifica as colunas de data válidas (cabeçalho) e cria um mapa
    //    coluna -> dia (para exibir os dias de falta)
    // 4) Para cada célula de frequência decide: falta justificada, falta
    //    não justificada ou atrasos (com pesos: 3->1h,2->2h,1->3h)
    async function processSheetFrequency({ sheet, ativosFiltrados, turmaFiltro, empresaFiltro, options }) {
      // Garantir que options exista para evitar checagens nulas adiante
      options = options || {};

      // decode_range devolve um objeto com s (start) e e (end) com índices de linhas/colunas
      const range = XLSX.utils.decode_range(sheet['!ref']);

      // Constrói mapa de alunos do CSV: chave é o RA numérico
      const alunosMap = {};
      (ativosFiltrados || []).forEach(a => {
        let raNum = null;
        // Se o registro acadêmico estiver presente no objeto do CSV
        if (a.REGISTRO_ACADEMICO) {
          // Preferimos usar AppUtils.normalizeRa se disponível (normaliza zeros e caracteres),
          // caso contrário aplicamos remoção de não-dígitos e parseInt simples.
          raNum = (typeof AppUtils !== 'undefined' && AppUtils.normalizeRa) ? AppUtils.normalizeRa(a.REGISTRO_ACADEMICO) : (function(raw){ const s=String(raw).replace(/\D/g,'').replace(/^0+/,'').trim(); return s?parseInt(s,10):null; })(a.REGISTRO_ACADEMICO);
        }
        if (raNum) {
          // Armazenamos somente campos relevantes para o relatório
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

      // Conjunto de RAs presentes no CSV para busca eficiente (O(1) por busca)
      const rasFromCsv = new Set(Object.keys(alunosMap).map(r => parseInt(r,10)));

      // localizar linhas de alunos na planilha
      // Valores padrão (podem ser sobrescritos via options)
      const firstStudentRow = (options && options.firstStudentRow !== undefined) ? options.firstStudentRow : 14; // índice (0-based)
      const raCol = (options && options.raCol !== undefined) ? options.raCol : 1; // coluna onde se espera o RA (B)
      const nameCol = (options && options.nameCol !== undefined) ? options.nameCol : 2; // coluna do nome do aluno (C)

      // linhaAlunos map: chave = número da linha na planilha, valor = RA
      const linhaAlunos = {};
      // frequencias: mapa RA -> objeto com contadores
      const frequencias = {};

      // Varre as linhas da planilha a partir da linha estimada onde os alunos começam
      for (let R = firstStudentRow; R <= range.e.r; ++R) {
        // Monta o endereço da célula que contém o RA (usando encode_cell do SheetJS)
        const raCellAddress = XLSX.utils.encode_cell({ c: raCol, r: R });
        const raCell = sheet[raCellAddress];
        // Se a célula existe e tem valor, tentamos extrair o RA numérico
        if (raCell && raCell.v) {
          const raNum = (typeof AppUtils !== 'undefined' && AppUtils.normalizeRa) ? AppUtils.normalizeRa(raCell.v) : (function(raw){ const s=String(raw).replace(/\D/g,'').replace(/^0+/,'').trim(); return s?parseInt(s,10):null; })(raCell.v);
          // Se o RA extraído existe e está presente no CSV, registramos a linha como aluno conhecido
          if (raNum && rasFromCsv.has(raNum)) {
            const nameCellAddress = XLSX.utils.encode_cell({ c: nameCol, r: R });
            const nameCell = sheet[nameCellAddress];
            const studentName = (nameCell && nameCell.v) ? String(nameCell.v).trim() : '';
            linhaAlunos[R] = raNum;
            frequencias[raNum] = {
              nome: studentName,
              faltasJust: 0,
              faltasNaoJust: 0,
              atrasos: 0,
              faltasJustDays: [],
              faltasNaoJustDays: [],
              atrasosDays: []
            };
          }
        }
      }

      // identificar colunas de data válidas (usa helper se disponível)
      // Padrão: começamos a procurar em C=4 (coluna E) e cabeçalho na linha 13
      const firstDateCol = (options && options.firstDateCol !== undefined) ? options.firstDateCol : 4;
      const dateHeaderRow = (options && options.dateHeaderRow !== undefined) ? options.dateHeaderRow : 13;
      let validDateCols = [];
      if (typeof AppXLSX !== 'undefined' && AppXLSX.findValidDateCols) {
        // Se o helper AppXLSX estiver disponível, delegamos a ele
        validDateCols = AppXLSX.findValidDateCols(sheet, { firstDateCol, dateHeaderRow });
      } else {
        // Fallback simples: consideramos todas as células não vazias na linha de cabeçalho
        // exceto colunas com textos específicos que identificamos como não-datas
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
      }

      // mapear coluna->dia (ex: coluna 4 -> '12') usando helper quando disponível
      const dateColToDay = {};
      if (typeof AppXLSX !== 'undefined' && AppXLSX.buildDateColToDay) {
        Object.assign(dateColToDay, AppXLSX.buildDateColToDay(sheet, validDateCols, { dateHeaderRow }));
      } else {
        // Função local para extrair o dia do texto do cabeçalho
        const extractDayFromHeader = (cell) => {
          if (typeof AppUtils !== 'undefined' && AppUtils.extractDayFromHeader) return AppUtils.extractDayFromHeader(cell);
          if (!cell) return '';
          const text = (cell.w !== undefined) ? String(cell.w) : String(cell.v);
          const m = text.match(/(\d{1,2})/);
          return m ? String(parseInt(m[1], 10)) : '';
        };
        for (const C of validDateCols) {
          const headerCellAddress = XLSX.utils.encode_cell({ c: C, r: dateHeaderRow });
          const headerCell = sheet[headerCellAddress];
          dateColToDay[C] = extractDayFromHeader(headerCell);
        }
      }

      // extrair mes/ano (usa helper se houver)
      let monthYearText = '';
      if (typeof AppXLSX !== 'undefined' && AppXLSX.extractMonthYear) {
        monthYearText = AppXLSX.extractMonthYear(sheet, { firstDateCol, dateHeaderRow });
      } else {
        // fallback simplificado: tenta ler células mescladas em uma linha específica
        try {
          const mergedRow = 4; // linha 5 (0-based index)
          const mergedCols = [36,37,38,39];
          let mergedText = '';
          for (const mc of mergedCols) {
            const mAddr = XLSX.utils.encode_cell({ c: mc, r: mergedRow });
            const mCell = sheet[mAddr];
            if (mCell && (mCell.w !== undefined || mCell.v !== undefined)) {
              mergedText += (mCell.w !== undefined) ? String(mCell.w) : String(mCell.v);
              mergedText += ' ';
            }
          }
          mergedText = mergedText.trim();
          if (mergedText) {
            const mMatch1 = mergedText.match(/(\d{1,2}\/\d{4})/);
            if (mMatch1) {
              const mm = String(mMatch1[1].split('/')[0]).padStart(2,'0');
              const yy = mMatch1[1].split('/')[1];
              monthYearText = `${mm}/${yy}`;
            }
          }
        } catch (e) {
          monthYearText = '';
        }
      }

      // percorrer cada aluno (linhas) e contar nas colunas de data
      for (const R_str in linhaAlunos) {
        const R = parseInt(R_str, 10);
        const raDoAluno = linhaAlunos[R];
        for (const C of validDateCols) {
          const freqCellAddress = XLSX.utils.encode_cell({ c: C, r: R });
          const freqCell = sheet[freqCellAddress];
          // apenas processa quando a célula existe e tem valor
          if (freqCell && freqCell.v) {
            const lowerVal = String(freqCell.v).toLowerCase();
            let isJustified = false;
            // heurística para detectar célula justificada: presença de estilo/ cor de fundo
            if (freqCell.s && freqCell.s.bgColor) {
              const bg = freqCell.s.bgColor;
              // se houver alguma das propriedades de cor (rgb/theme/indexed), consideramos justificado
              if (bg.rgb || bg.theme !== undefined || bg.indexed !== undefined) {
                isJustified = true;
              }
            }
            const dayForThisCol = dateColToDay[C] || '';
            // Lógica de contagem:
            // - 'f' = falta (justificada se isJustified)
            // - '3','2','1' = códigos de atraso com pesos distintos
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
                // código '3' soma 1 hora (convenção do projeto)
                frequencias[raDoAluno].atrasos += 1;
                if (dayForThisCol) frequencias[raDoAluno].atrasosDays.push(dayForThisCol);
              }
            } else if (lowerVal === '2') {
              if (!isJustified) {
                // código '2' soma 2 horas
                frequencias[raDoAluno].atrasos += 2;
                if (dayForThisCol) frequencias[raDoAluno].atrasosDays.push(dayForThisCol);
              }
            } else if (lowerVal === '1') {
              if (!isJustified) {
                // código '1' soma 3 horas
                frequencias[raDoAluno].atrasos += 3;
                if (dayForThisCol) frequencias[raDoAluno].atrasosDays.push(dayForThisCol);
              }
            }
          }
        }
      }

      // Retorna os dados agregados para o chamador
      return {
        alunosMap,
        frequencias,
        monthYearText,
        validDateCols,
        dateColToDay
      };
    }

    // Expõe a função como parte do namespace global AppProcessor
    global.AppProcessor = global.AppProcessor || {};
    global.AppProcessor.processSheetFrequency = processSheetFrequency;

  })(window);

// src/xlsxReader.js - helpers para ler e interpretar planilhas de frequência (SheetJS/XLSX)
  (function(global){
    // findStudentRows
    // - Objetivo: localizar as linhas da planilha que correspondem a
    //   alunos presentes no CSV, extraindo o RA e o nome.
    function findStudentRows(sheet, options) {
      // Permite sobrescrever colunas/linhas padrão via options
      const raCol = (options && options.raCol !== undefined) ? options.raCol : 1;
      const nameCol = (options && options.nameCol !== undefined) ? options.nameCol : 2;
      const startRow = (options && options.startRow !== undefined) ? options.startRow : 14;
      // decode_range usa o '!ref' para obter o range usado na planilha
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const linhaAlunos = {};
      const frequenciasInit = {};

      // Varre da linha de início até o final do range
      for (let R = startRow; R <= range.e.r; ++R) {
        // Monta endereço da célula do RA e lê
        const raCellAddress = XLSX.utils.encode_cell({ c: raCol, r: R });
        const raCell = sheet[raCellAddress];
        if (raCell && raCell.v) {
          // Remove não-dígitos e zeros à esquerda para obter RA numérico
          const raStr = String(raCell.v).replace(/\D/g, '').replace(/^0+/, '').trim();
          if (raStr) {
            const raNum = parseInt(raStr, 10);
            // Lê o nome do aluno na coluna apropriada
            const nameCellAddress = XLSX.utils.encode_cell({ c: nameCol, r: R });
            const nameCell = sheet[nameCellAddress];
            const studentName = (nameCell && nameCell.v) ? String(nameCell.v).trim() : '';
            // Inicializa estrutura de frequências para este RA
            frequenciasInit[raNum] = { nome: studentName, faltasJust:0, faltasNaoJust:0, atrasos:0, faltasJustDays:[], faltasNaoJustDays:[], atrasosDays:[] };
            // Mapeia a linha para o RA para uso posterior na varredura
            linhaAlunos[R] = raNum;
          }
        }
      }
      // Retorna duas estruturas: mapeamento linha->RA e estruturas iniciais de frequência
      return { linhaAlunos, frequenciasInit };
    }

    // findValidDateCols
    // - Objetivo: identificar quais colunas correspondem a dias do mês
    function findValidDateCols(sheet, options) {
      const firstDateCol = (options && options.firstDateCol !== undefined) ? options.firstDateCol : 4;
      const dateHeaderRow = (options && options.dateHeaderRow !== undefined) ? options.dateHeaderRow : 13;
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const validDateCols = [];

      // Varre da coluna inicial até o fim do range e coleta colunas cujo
      // cabeçalho não seja um texto que indicamos como não-data.
      for (let C = firstDateCol; C <= range.e.c; ++C) {
        const headerCellAddress = XLSX.utils.encode_cell({ c: C, r: dateHeaderRow });
        const headerCell = sheet[headerCellAddress];
        if (headerCell && headerCell.v) {
          const headerText = String(headerCell.v).trim().toUpperCase();
          // Exclui cabeçalhos conhecidos que não representam datas
          if (headerText !== 'AULAS VS PRESENÇA' && headerText !== 'FALTAS NO MÊS') {
            validDateCols.push(C);
          }
        }
      }
      return validDateCols;
    }

    // buildDateColToDay
    // - Objetivo: criar um mapa coluna -> dia (ex: coluna 4 -> '12')
    function buildDateColToDay(sheet, validDateCols, options) {
      const dateHeaderRow = (options && options.dateHeaderRow !== undefined) ? options.dateHeaderRow : 13;
      const map = {};
      for (const C of validDateCols) {
        const headerCellAddress = XLSX.utils.encode_cell({ c: C, r: dateHeaderRow });
        const headerCell = sheet[headerCellAddress];
        // Preferimos o texto formatado (w) quando disponível
        const text = (headerCell && headerCell.w !== undefined) ? String(headerCell.w) : (headerCell && headerCell.v !== undefined) ? String(headerCell.v) : '';
        const m = text.match(/(\d{1,2})/);
        map[C] = m ? String(parseInt(m[1], 10)) : '';
      }
      return map;
    }

    // extractMonthYear
    // - Objetivo: tentar extrair o mês/ano do cabeçalho da planilha utilizando
    //   várias heurísticas (células mescladas conhecidas, textos próximos a "PERÍODO",
    //   intervalos de datas, ou cabeçalhos de coluna com data completa).
    function extractMonthYear(sheet, options) {
      let monthYearText = '';
      try {
        // Heurística 1: ler blocos mesclados em uma linha conhecida (linha 5 -> índice 4)
        const mergedRow = 4; // linha 5 -> índice 4
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
          // Procura padrões como 'MM/YYYY' ou 'DD/MM/YYYY'
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
            // Tenta detectar nomes de mês em texto livre (ex: 'Agosto de 2024')
            const months = { 'janeiro':'01','fevereiro':'02','março':'03','marco':'03','abril':'04','maio':'05','junho':'06','julho':'07','agosto':'08','setembro':'09','outubro':'10','novembro':'11','dezembro':'12' };
            const lower = mergedText.toLowerCase();
            const yearMatch = lower.match(/(20\d{2}|19\d{2})/);
            if (yearMatch) {
              const y = yearMatch[1];
              for (const [name, num] of Object.entries(months)) {
                if (lower.indexOf(name) !== -1) { monthYearText = `${num}/${y}`; break; }
              }
            }
            if (!monthYearText) monthYearText = mergedText; // fallback para texto bruto
          }
        }
      } catch (e) {
        monthYearText = monthYearText || '';
      }

      // Heurística 2: procurar uma célula que contenha 'PERÍODO COMP' e extrair datas próximas
      if (!monthYearText) {
        const range = XLSX.utils.decode_range(sheet['!ref']);
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
              // Função auxiliar para extrair texto, datas ou valores numéricos formatados
              const extractText = (cellObj) => {
                if (!cellObj || (cellObj.v === undefined || cellObj.v === null)) return '';
                try {
                  // Se a célula for numérica e o SheetJS tiver parse_date_code, converte para dd/mm/yyyy
                  if (cellObj.t === 'n' && typeof XLSX !== 'undefined' && XLSX.SSF && typeof XLSX.SSF.parse_date_code === 'function') {
                    const d = XLSX.SSF.parse_date_code(cellObj.v);
                    if (d && d.y) {
                      const dd = String(d.d).padStart(2, '0');
                      const mm = String(d.m).padStart(2, '0');
                      const yyyy = String(d.y);
                      return `${dd}/${mm}/${yyyy}`;
                    }
                  }
                } catch (e) {}
                const text = (cellObj.w !== undefined) ? String(cellObj.w) : String(cellObj.v);
                const m2 = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
                if (m2) return m2[1];
                const m = text.match(/(\d{1,2}\/\d{4})/);
                if (m) return m[1];
                return text.trim();
              };
              const rawText = extractText(leftCell) || extractText(rightCell) || ((String(cell.v).split(':')[1] || '').trim());
              let my = '';
              if (rawText) {
                const mDate = rawText.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                if (mDate) {
                  const day = mDate[1];
                  const month = mDate[2].padStart(2, '0');
                  const year = mDate[3];
                  my = `${month}/${year}`;
                } else {
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

      // Heurística 3: procurar intervalos do tipo 'dd/mm/yyyy - dd/mm/yyyy' em qualquer célula
      if (!monthYearText) {
        const range2 = XLSX.utils.decode_range(sheet['!ref']);
        outerLoop2:
        for (let r2 = range2.s.r; r2 <= range2.e.r; r2++) {
          for (let c2 = range2.s.c; c2 <= range2.e.c; c2++) {
            const addr2 = XLSX.utils.encode_cell({ c: c2, r: r2 });
            const cell2 = sheet[addr2];
            if (!cell2 || (cell2.v === undefined || cell2.v === null)) continue;
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

      // Heurística 4: olhar a linha de cabeçalho das datas em busca de um formato dd/mm/yyyy ou mm/yyyy
      if (!monthYearText) {
        const range3 = XLSX.utils.decode_range(sheet['!ref']);
        const dateHeaderRow = (options && options.dateHeaderRow !== undefined) ? options.dateHeaderRow : 13;
        for (let C = (options && options.firstDateCol !== undefined ? options.firstDateCol : 4); C <= range3.e.c; ++C) {
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

      // Retorna o resultado, que pode ser string vazia se nenhuma heurística funcionar
      return monthYearText;
    }

    // Export
    global.AppXLSX = global.AppXLSX || {};
    global.AppXLSX.findStudentRows = findStudentRows;
    global.AppXLSX.findValidDateCols = findValidDateCols;
    global.AppXLSX.buildDateColToDay = buildDateColToDay;
    global.AppXLSX.extractMonthYear = extractMonthYear;

  })(window);
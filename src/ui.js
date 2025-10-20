// src/ui.js - responsabilidade pela UI e orquestração (substitui app.js)
  (function(global){
    // Variável de debug que ativa logs adicionais quando true
    const APP_DEBUG = false;

    // Fallbacks para leitura de arquivos quando AppIO não estiver presente.
    // Estas funções usam diretamente FileReader sem heurísticas de encoding.
    function readFileAsTextFallback(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsText(file, 'utf-8');
      });
    }
    function readFileAsArrayBufferFallback(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
      });
    }

    // Armazena temporariamente os dados lidos do CSV de Ativos para uso
    // por funções auxiliares (ex: preencher selects de turma/empresa)
    let ativosDataGlobal = [];

    // Função que atualiza o select de empresas com base na turma selecionada
    function atualizarEmpresasPorTurma(turmaSelecionada) {
      let empresas = [];
      if (turmaSelecionada) {
        // Filtra apenas os ativos daquela turma e extrai nomes de empresa únicos
        empresas = [...new Set(ativosDataGlobal.filter(a => a.CODTURMA && a.CODTURMA.trim() === turmaSelecionada)
          .map(a => a.NOMEEMPRESA_NOVO ? a.NOMEEMPRESA_NOVO.trim() : '').filter(Boolean))].sort();
      } else {
        // Sem turma selecionada, lista todas as empresas existentes no CSV
        empresas = [...new Set(ativosDataGlobal.map(a => a.NOMEEMPRESA_NOVO ? a.NOMEEMPRESA_NOVO.trim() : '').filter(Boolean))].sort();
      }
      const empresaSelect = document.getElementById('empresaSelect');
      if (empresaSelect) empresaSelect.innerHTML = '<option value="">Todas</option>' + empresas.map(e => `<option value="${e}">${e}</option>`).join('');
    }

    // Hook up event listeners quando o DOM estiver pronto
    document.addEventListener('DOMContentLoaded', () => {
      // Referências aos elementos do HTML necessários
      const ativosInput = document.getElementById('ativosCsv');
      const turmaSelect = document.getElementById('turmaSelect');
      const empresaSelect = document.getElementById('empresaSelect');
      const freqForm = document.getElementById('freqForm');
      const outputDiv = document.getElementById('output');

      // Quando o usuário seleciona o CSV de "Ativos" carregamos e
      // populamos as listas de turma/empresa para filtros.
      if (ativosInput) {
        ativosInput.addEventListener('change', async function (e) {
          // Pega o primeiro arquivo selecionado (se houver múltiplos)
          const file = e.target.files[0];
          if (!file) return;
          try {
            // Leitura com heurística de encoding quando AppIO estiver disponível
            const ativosText = (typeof AppIO !== 'undefined' && AppIO.readFileAsText) ? await AppIO.readFileAsText(file) : await readFileAsTextFallback(file);
            // Parser: converte CSV em array de objetos (header-aware)
            ativosDataGlobal = (typeof AppCSV !== 'undefined' && AppCSV.parseAtivos) ? AppCSV.parseAtivos(ativosText) : Papa.parse(ativosText, { header: true, skipEmptyLines: true, delimiter: ';' }).data;

            // Extraímos as turmas que começam com 'APR' (convenção usada no projeto)
            const turmas = [...new Set(ativosDataGlobal
              .map(a => a.CODTURMA ? a.CODTURMA.trim() : '')
              .filter(Boolean)
              .filter(t => t.toUpperCase().startsWith('APR'))
            )].sort();

            // Popula select de turmas com opção 'Todas' + cada turma encontrada
            if (turmaSelect) turmaSelect.innerHTML = '<option value="">Todas</option>' + turmas.map(t => `<option value="${t}">${t}</option>`).join('');
            // Atualiza lista de empresas com base na seleção (inicialmente vazio -> todas)
            atualizarEmpresasPorTurma('');
            if (outputDiv) outputDiv.innerHTML = `<p>Foram encontradas um total de ${turmas.length} turmas no Programa de Aprendizagem Industrial.</p>`;
          } catch (error) {
            // Em caso de erro na leitura/parsing, mostramos mensagem e logamos no console
            console.error('Erro ao ler o arquivo CSV:', error);
            if (outputDiv) outputDiv.innerHTML = `<p style="color:red"><b>Erro ao processar o arquivo CSV:</b> ${error.message}</p>`;
          }
        });
      }

      // Quando o usuário troca a turma selecionada, atualizamos as empresas
      if (turmaSelect) {
        turmaSelect.addEventListener('change', function (e) {
          const turmaSelecionada = e.target.value.trim();
          atualizarEmpresasPorTurma(turmaSelecionada);
        });
      }

      // Handler principal do formulário: quando o usuário clica em Gerar/Enviar
      if (freqForm) {
        freqForm.addEventListener('submit', async function (e) {
          // Prevenir submissão padrão de formulário
          e.preventDefault();

          // Ler arquivos selecionados nos inputs (novamente para garantir acesso)
          const ativosFile = document.getElementById('ativosCsv').files[0];
          const freqFile = document.getElementById('freqXls').files[0];
          const turmaFiltro = (turmaSelect && turmaSelect.value) ? turmaSelect.value.trim() : '';
          const empresaFiltro = (empresaSelect && empresaSelect.value) ? empresaSelect.value.trim() : '';

          // Validação rápida: ambos os arquivos são obrigatórios
          if (!ativosFile || !freqFile) {
            alert('Selecione ambos os arquivos.');
            return;
          }
          if (outputDiv) outputDiv.innerHTML = 'Processando arquivos...';

          // Fluxo principal quando o formulário é enviado pelo usuário.
          // Passos:
          // 1) Ler e parsear CSV dos Ativos
          // 2) Ler planilha de frequência (XLS/XLSX)
          // 3) Chamar AppProcessor para cruzar e contar
          // 4) Construir 'saida' e delegar para AppExporter
          try {
            // Leitura do CSV (com heurística de encoding via AppIO quando disponível)
            const ativosText = (typeof AppIO !== 'undefined' && AppIO.readFileAsText) ? await AppIO.readFileAsText(ativosFile) : await readFileAsTextFallback(ativosFile);
            let ativosFiltrados = (typeof AppCSV !== 'undefined' && AppCSV.parseAtivos) ? AppCSV.parseAtivos(ativosText) : Papa.parse(ativosText, { header: true, skipEmptyLines: true, delimiter: ';' }).data;

            // Aplicação dos filtros selecionados pelo usuário
            if (turmaFiltro) ativosFiltrados = ativosFiltrados.filter(a => a.CODTURMA && a.CODTURMA.trim() === turmaFiltro);
            if (empresaFiltro) ativosFiltrados = ativosFiltrados.filter(a => a.NOMEEMPRESA_NOVO && a.NOMEEMPRESA_NOVO.trim() === empresaFiltro);

            // Se não sobrar nenhum aluno após filtros, informamos e abortamos
            if (ativosFiltrados.length === 0) {
              if (outputDiv) outputDiv.innerHTML = '<p style="color:red; font-weight: bold;">Nenhum aluno encontrado após a aplicação dos filtros.</p>';
              return;
            }

            // Leitura da planilha de frequência como ArrayBuffer (para o SheetJS)
            const freqBuffer = (typeof AppIO !== 'undefined' && AppIO.readFileAsArrayBuffer) ? await AppIO.readFileAsArrayBuffer(freqFile) : await readFileAsArrayBufferFallback(freqFile);
            const workbook = XLSX.read(freqBuffer, { type: 'array', cellStyles: true });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Preferimos usar AppProcessor (módulo que centraliza a lógica de varredura)
            let processamentoResult = null;
            if (typeof AppProcessor !== 'undefined' && AppProcessor.processSheetFrequency) {
              try {
                processamentoResult = await AppProcessor.processSheetFrequency({ sheet, ativosFiltrados, turmaFiltro, empresaFiltro, options: {} });
              } catch (err) {
                // Se o AppProcessor falhar por algum motivo, caímos em fallback (não implementado aqui)
                console.warn('AppProcessor failed, using fallback inline', err);
                processamentoResult = null;
              }
            }

            // Variáveis que receberão o resultado do processamento
            let alunosMap = {};
            let frequencias = {};
            let monthYearText = '';
            let validDateCols = [];
            let dateColToDay = {};

            // Se o processamento ocorreu, extraímos os mapas necessários
            if (processamentoResult) {
              alunosMap = processamentoResult.alunosMap || {};
              frequencias = processamentoResult.frequencias || {};
              monthYearText = processamentoResult.monthYearText || '';
              validDateCols = processamentoResult.validDateCols || [];
              dateColToDay = processamentoResult.dateColToDay || {};
            } else {
              // Deve ser raro — o AppProcessor existe para evitar lógica duplicada.
              throw new Error('Processamento inline não implementado no UI; habilite AppProcessor');
            }

            // companyText resolution: prefer filtro de empresa, senão pega a primeira encontrada
            let companyText = (empresaFiltro && empresaFiltro.length) ? empresaFiltro : '';
            if (!companyText && Array.isArray(ativosFiltrados) && ativosFiltrados.length) {
              const first = ativosFiltrados.find(x => x.NOMEEMPRESA_NOVO && String(x.NOMEEMPRESA_NOVO).trim());
              if (first && first.NOMEEMPRESA_NOVO) companyText = String(first.NOMEEMPRESA_NOVO).trim();
            }

            // build output rows: transformamos os mapas em um array `saida` de objetos
            const saida = [];
            for (const raStr in alunosMap) {
              const ra = parseInt(raStr, 10);
              const alunoInfo = alunosMap[ra];
              const freqInfo = frequencias[ra] || { faltasJust: 0, faltasNaoJust: 0, atrasos: 0 };
              // total de horas de ausência = horas de atraso + faltas * 4 (convenção)
              const totalAusenciaHoras = (freqInfo.atrasos || 0) + ((freqInfo.faltasNaoJust || 0) + (freqInfo.faltasJust || 0)) * 4;

              // sanitize textual fields to avoid broken characters and layout issues
              const sanitize = (v) => (typeof AppUtils !== 'undefined' && AppUtils.sanitizeText) ? AppUtils.sanitizeText(v) : (v ? String(v).trim() : '');
              const alunoSan = sanitize(alunoInfo.ALUNO);
              const praticaSan = sanitize(alunoInfo.PRATICA);
              const cursoSan = sanitize(alunoInfo.CURSO);
              const turmaSan = sanitize(alunoInfo.CODTURMA);

              const rowObj = {
                mes: (alunoInfo.DTINICIAL || '').split('/')[1] || '',
                inicio: alunoInfo.DTINICIAL,
                termino: alunoInfo.DTFINAL,
                turma: turmaSan,
                aluno: alunoSan,
                pratica: praticaSan,
                curso: cursoSan,
                faltasJustCount: freqInfo.faltasJust,
                faltasJustDays: (freqInfo.faltasJustDays && freqInfo.faltasJustDays.length) ? freqInfo.faltasJustDays.join(', ') : '',
                faltasNaoJustCount: freqInfo.faltasNaoJust,
                faltasNaoJustDays: (freqInfo.faltasNaoJustDays && freqInfo.faltasNaoJustDays.length) ? freqInfo.faltasNaoJustDays.join(', ') : '',
                horasAtraso: freqInfo.atrasos,
                atrasosDays: (freqInfo.atrasosDays && freqInfo.atrasosDays.length) ? freqInfo.atrasosDays.join(', ') : '',
                totalAusenciaHoras: totalAusenciaHoras
              };
              saida.push(rowObj);
            }

            // Se for relatório por empresa (turma selecionada, empresa não), ordenamos por empresa
            const perEmpresaReport = (turmaFiltro && turmaFiltro.length > 0) && (!empresaFiltro || empresaFiltro.length === 0);
            if (perEmpresaReport) {
              saida.sort((a, b) => { const A = (a.empresa || '').toUpperCase(); const B = (b.empresa || '').toUpperCase(); if (A < B) return -1; if (A > B) return 1; return 0; });
            }

            if (saida.length === 0) {
              if (outputDiv) outputDiv.innerHTML = '<p style="color:red"><b>Nenhum dado cruzado encontrado após a filtragem.</b></p>';
              return;
            }

            // Delegamos a exportação para AppExporter, que monta o XLSX e aciona o download
            const exporterOptions = { saida, perEmpresaReport, monthYearText, turmaFiltro, empresaFiltro, companyText, debug: APP_DEBUG };
            if (typeof AppExporter !== 'undefined' && AppExporter.buildAndDownloadWorkbook) {
              try {
                const finalName = await AppExporter.buildAndDownloadWorkbook(exporterOptions);
                if (APP_DEBUG) console.log('[DEBUG] Final filename (exporter):', finalName);
                if (outputDiv) outputDiv.innerHTML = '<p style="color:green"><b>Relatório gerado com sucesso!</b></p>';
              } catch (ex) {
                console.warn('Exporter failed:', ex);
                if (outputDiv) outputDiv.innerHTML = '<p style="color:red"><b>Erro ao exportar o relatório. Verifique o console.</b></p>';
              }
            } else {
              // Mensagem de diagnóstico avançado quando o exportador não estiver carregado
              console.error('AppExporter não disponível; verifique se src/exporter.js foi carregado.');
              try {
                console.log('window.AppExporter ===', window.AppExporter);
                // Tenta encontrar a tag <script> responsável por carregar o exportador
                const scripts = Array.from(document.getElementsByTagName('script'));
                const exporterScript = scripts.find(s => s.src && (s.src.endsWith('/src/exporter.js') || s.src.indexOf('src/exporter.js') !== -1));
                if (exporterScript) {
                  console.log('Encontrado tag <script> para exporter:', exporterScript.src, exporterScript);
                } else {
                  console.log('Nenhuma tag <script> encontrada com src exportador (procure por src/exporter.js)');
                }
                // Tenta buscar entradas de recurso na API de performance para confirmar o fetch
                try {
                  const perf = performance.getEntriesByType('resource') || [];
                  const exportEntry = perf.find(p => p.name && p.name.indexOf('src/exporter.js') !== -1);
                  console.log('Performance resource entry for exporter (if any):', exportEntry || 'nenhum');
                } catch (pe) {
                  console.log('Performance lookup falhou:', pe);
                }
              } catch (diagErr) {
                console.log('Erro ao executar diagnósticos do exporter:', diagErr);
              }
              if (outputDiv) outputDiv.innerHTML = '<p style="color:red"><b>Exportador não carregado.</b> Verifique o Console (F12) e a aba Network para `src/exporter.js`, depois recarregue a página.</p>';
            }

          } catch (error) {
            // Captura qualquer erro do fluxo principal e informa o usuário
            console.error('Ocorreu um erro durante o processamento:', error);
            if (outputDiv) outputDiv.innerHTML = `<p style="color:red"><b>Ocorreu um erro:</b> ${error.message}. Verifique o console do navegador para mais detalhes.</p>`;
          }
        });
      }
    });

    // expose for debugging if needed
    global.AppUI = global.AppUI || {};
    global.AppUI.VERSION = '1.0.0-modular';

  })(window);

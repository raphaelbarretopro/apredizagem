# Processador de Frequência — Aprendizagem (documentação didática)

Este projeto é uma ferramenta cliente (rodando 100% no navegador) para
cruzar as informações do CSV de "Ativos" com as planilhas mensais de
frequência (arquivos .xls / .xlsx), calcular faltas e horas de atraso e
gerar um relatório Excel formatado.

Sumário rápido
- Objetivo: gerar relatórios de frequência por Turma e (opcionalmente) por
	Empresa, lendo os arquivos localmente no browser e exportando um .xlsx.
- Arquitetura: pequenos módulos JS em `src/` carregados diretamente por
	`index.html` (sem bundler). As bibliotecas externas são carregadas via CDN.

Como abrir e usar
1. Abra `index.html` no navegador (duplo clique funciona; para evitar
	 restrições de permissão em alguns browsers, você pode servir com um
	 servidor local simples: `python -m http.server` ou `npx http-server`).
2. No formulário selecione:
	 - `Arquivo Ativos` (CSV) — arquivo gerado pelo sistema de cadastro de
		 aprendizes. Deve usar `;` como separador. O parser usado é o PapaParse.
	 - `Arquivo Frequência` (.xls/.xlsx) — planilha mensal com presença.
3. Opcionalmente filtre por `Turma` e/ou `Empresa`.
4. Clique em "Processar e Gerar Relatório" — o download do relatório será
	 iniciado automaticamente.

Descrição detalhada dos arquivos

- `index.html`
	- Front-end simples com o formulário de upload e selects para filtrar
		`Turma` e `Empresa` que são populados dinamicamente a partir do CSV.
	- Carrega as bibliotecas externas (SheetJS / PapaParse / ExcelJS) e os
		módulos em `src/` numa ordem pensada para satisfazer dependências.

- `src/utils.js`
	- Funções utilitárias puras usadas pelo restante do código.
	- Funções principais:
		- `normalizeRa` — padroniza o RA removendo caracteres não-numéricos
			e zeros à esquerda.
		- `sanitizeFilename` — produz nomes de arquivo seguros para download.
		- `extractDayFromHeader` — tenta extrair o dia (número) de um texto de
			cabeçalho de coluna (ex.: "12" ou "12/08").
		- `sanitizeText` — limpa texto com problemas de encoding e caracteres
			invisíveis.
	- Testes ideais: unitários para `normalizeRa` e `extractDayFromHeader`.

- `src/io.js`
	- Abstrai leitura de arquivos pelo FileReader.
	- Fornece duas funções:
		- `readFileAsText(file)` — lê como texto; tenta UTF-8 e faz fallback para
			Windows-1252 (com TextDecoder) se detectar caracteres de substituição.
		- `readFileAsArrayBuffer(file)` — lê em ArrayBuffer (usado pelo SheetJS).
	- Importante: fallback de encoding reduz problemas com CSVs gerados no
		Windows/Excel que vêm em CP1252.

- `src/csvParser.js`
	- Pequeno wrapper em torno do PapaParse para parse do CSV de Ativos.
	- Opções fixas: `header:true`, `skipEmptyLines:true`, `delimiter:';'`.

- `src/xlsxReader.js`
	- Coleção de heurísticas para extrair informação útil das planilhas de
		frequência (formatos variados). As funções tentam ser tolerantes a
		diferentes layouts.
	- Funções principais:
		- `findStudentRows(sheet, options)` — varre linhas a partir de uma linha
			de início e identifica RA e nome, inicializando estruturas de
			frequência.
		- `findValidDateCols(sheet, options)` — determina quais colunas são dias
			do mês (filtrando textos comuns que não representam dias).
		- `buildDateColToDay(sheet, validDateCols, options)` — mapeia coluna -> dia.
		- `extractMonthYear(sheet, options)` — heurísticas para extrair mês/ano
			do cabeçalho (blocos mesclados, células próximas a "PERÍODO COMP",
			intervalos de datas, ou cabeçalhos de data).
	- Observação: planilhas muito diferentes podem exigir ajuste de parâmetros
		(ex.: `firstDateCol`, `dateHeaderRow`).

- `src/processor.js`
	- Responsável pelo núcleo do problema: cruza os dados do CSV com a
		planilha por RA e computa:
		- Nº de Faltas Justificadas (lista de dias)
		- Nº de Faltas Não Justificadas (lista de dias)
		- Nº de Horas de Atraso (soma de horas mapeadas por códigos)
	- Regras implementadas (resumido):
		- Células vazias → presença (não conta como falta)
		- Células com 'F' (ou similar) → falta (classifica-se como justificada
			ou não dependendo de estilo/heurística)
		- Códigos '3','2','1' usados para representar atrasos de 1/2/3 horas
			(a lógica mapeia conforme definido no código).

- `src/exporter.js`
	- Gera o arquivo `.xlsx` final com ExcelJS e aplica estilos básicos
		(cabeçalho mesclado, larguras de coluna, preenchimento alternado).
	- Nome do arquivo segue a convenção:
		`relatorio_frequencia_<MM-YYYY>_<CODTURMA>[_<NOMEEMPRESA>].xlsx` (omitindo
		o segmento empresa se o filtro for "Todas").

- `src/ui.js`
	- Faz a ligação entre DOM e lógica:
		- Popula selects de Turma/Empresa a partir do CSV de Ativos.
		- Lida com eventos do formulário, leitura dos arquivos e fluxo de
			processamento (chama `AppProcessor` e `AppExporter`).
		- Mostra mensagens de progresso/erro em `#output`.
	- Variável útil: `APP_DEBUG` pode ser ativada para mensagens de
		diagnóstico no console.

Comportamento esperado e limitações
- A detecção de justificativas (falta justificada) depende de estilos de
	célula na planilha (por cor). Nem sempre .xls antigos preservam estilos
	quando lidos pelo SheetJS, portanto essa heurística pode falhar.
- Layouts diferentes (colunas deslocadas, linhas de cabeçalho em posições
	diferentes) podem exigir ajustes manuais nas opções (ver funções em
	`src/xlsxReader.js` e parâmetros em `src/processor.js`).
- Problemas de encoding no CSV são tratados com um fallback para Windows-1252
	em `src/io.js`, mas arquivos muito corrompidos podem ainda apresentar
	caracteres estranhos — usar `AppUtils.sanitizeText` ajuda.

Dicas de depuração rápida
- Abra o DevTools (F12) e ative `APP_DEBUG = true` em `src/ui.js` para ver
	logs com etapas internas (detecção de RA, colunas de data, contagens por RA).
- Verifique alguns RAs manualmente no CSV e compare com as linhas do
	relatório para confirmar o mapeamento.
- Se as turmas não aparecem no select, verifique se o CSV foi carregado
	corretamente e se o delimitador é `;`.

Contribuição e histórico
- O código foi refatorado de uma única entrada (`app.js`) para módulos em
	`src/` para facilitar manutenção. O arquivo `app.js` antigo foi removido
	do repositório porque sua lógica foi substituída por `src/ui.js`.

Conteúdo do diretório (breve)
- `index.html` — UI e carregamento das bibliotecas/CDNs
- `estilo.css` — estilos da UI
- `src/`:
	- `utils.js`, `io.js`, `csvParser.js`, `xlsxReader.js`, `processor.js`, `exporter.js`, `ui.js`

Se quiser, eu posso:
- adicionar exemplos de CSV e XLSX (pequenos fixtures) para testes locais;
- gerar testes unitários para `AppUtils` (normalizeRa, sanitizeText);
- atualizar o README com instruções para rodar um servidor local e commits
	sugeridos (git) para versionamento.

-----
README atualizado com explicações por arquivo — se quiser que eu adicione
um exemplo mínimo de CSV e planilha para testar o fluxo, diga e eu crio.

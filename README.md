# Processador de Frequência — Aprendizagem

Breve README e checklist para validação rápida.

Arquivos principais:
- `index.html` — interface web (aberta no navegador).
- `src/ui.js` — orquestra leitura/parse/contagem/export e ligações da UI (substitui o antigo `app.js`).
- `src/utils.js` — funções puras (normalizeRa, sanitizeFilename, extractDayFromHeader).
- `src/csvParser.js` — wrapper PapaParse.
- `src/xlsxReader.js` — heurísticas para localizar colunas/datas/mês-ano.
- `src/processor.js` — processamento da planilha (contagem de faltas/atrasos).
- `src/exporter.js` — construção do workbook e download (ExcelJS).

Como usar:
1. Abra `index.html` no navegador (duplo clique ou via um servidor local simples).
2. Selecione o CSV de Ativos e o arquivo de Frequência (.xls ou .xlsx).
3. Escolha filtros (Turma / Empresa) e clique em "Processar e Gerar Relatório".
4. O relatório será gerado e transferido ao seu computador.

Checklist de validação rápida:
- [ ] Verificar que as turmas aparecem no combobox após carregar o CSV.
- [ ] Selecionar uma turma específica e verificar lista de empresas atualizada.
- [ ] Gerar relatório com empresa específica e verificar que o nome da empresa aparece no título do arquivo e no conteúdo.
- [ ] Gerar relatório com empresa = "Todas" e verificar que o nome da empresa NÃO aparece no nome do arquivo.
- [ ] Conferir alguns RAs do CSV com as linhas do relatório para garantir que as contagens batem.
- [ ] Abrir devtools (F12) e verificar mensagens de DEBUG quando `APP_DEBUG` = true.

Notas:
- A detecção de justificativas depende de estilos de célula (`cell.s.bgColor`) presentes no arquivo lido. Em alguns `.xls` antigos os estilos podem não ser preservados; nesses casos justificativas por cor podem não ser detectadas.
- Se houver variações fortes no layout das planilhas (colunas deslocadas), ajuste os parâmetros em `src/processor.js` (por exemplo `firstStudentRow`, `raCol`, `firstDateCol`, `dateHeaderRow`).

Próximos passos recomendados:
- Adicionar `src/ui.js` para separar DOM/UI da lógica.
- Adicionar testes unitários para `AppUtils` e funções de extração de datas.
- Sincronizar mudanças com `projeto/web/app.js` se existir outra entrada.

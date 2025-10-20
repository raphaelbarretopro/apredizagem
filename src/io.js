/*
  src/io.js
  -----------------
  Funções utilitárias responsáveis por ler arquivos no navegador usando
  a API FileReader. Expondo os helpers como AppIO para os demais módulos
  consumirem. Este arquivo contém uma heurística simples para lidar com
  CSVs que não estejam em UTF-8: quando detectamos o caractere de
  substituição '�' o código tenta reler o arquivo e decodificá-lo como
  Windows-1252 (cp1252) usando TextDecoder, que está disponível na maioria
  dos navegadores modernos.

  Observação: o fallback tenta melhorar a experiência com arquivos gerados
  no Windows; ele não substitui uma solução completa de detecção de encoding
  (que exigiria bibliotecas adicionais ou endpoint server-side).
*/
  (function(global){
    // IIFE para encapsular e exportar somente o que for necessário
    // (evita poluir o escopo global diretamente).

    // readFileAsText
    // - Parâmetro: `file` é um objeto File obtido por um <input type="file">.
    // - Retorno: Promise que resolve para uma string com o conteúdo do arquivo.
    // - Comportamento: tenta ler como UTF-8; se detectar sinais de má
    //   decodificação (caractere de substituição '�') tenta reler como
    //   windows-1252 usando TextDecoder.
    function readFileAsText(file) {
      // Retornamos uma Promise para permitir await/then pelo chamador.
      return new Promise((resolve, reject) => {
        // FileReader para leitura textual inicial (UTF-8)
        const reader = new FileReader();

        // onload será chamado quando a leitura em UTF-8 terminar
        reader.onload = e => {
          // resultado da leitura (string) em UTF-8
          let text = e.target.result;

          // Heurística: se a string contém o caractere de substituição '�'
          // é provável que tenhamos problemas de encoding. Nesse caso,
          // e se TextDecoder estiver disponível, tentamos reler como
          // ArrayBuffer e decodificar como windows-1252 (CP1252).
          if (text.indexOf('�') !== -1 && typeof TextDecoder !== 'undefined') {
            try {
              // Criamos outro FileReader para ler binário (ArrayBuffer)
              const fr = new FileReader();
              fr.onload = ev => {
                // Convertendo o resultado para Uint8Array para o TextDecoder
                const array = new Uint8Array(ev.target.result);
                try {
                  // TextDecoder com cp1252 pode recuperar acentos comuns do Windows
                  const decoder = new TextDecoder('windows-1252');
                  const decoded = decoder.decode(array);
                  // Retornamos o texto decodificado com CP1252
                  resolve(decoded);
                } catch (dErr) {
                  // Se o TextDecoder falhar por algum motivo, mantemos
                  // o texto original em UTF-8 como fallback.
                  resolve(text);
                }
              };

              // Se houver erro lendo como ArrayBuffer, mantemos o texto original
              fr.onerror = () => resolve(text);
              // Inicia a leitura binária do arquivo
              fr.readAsArrayBuffer(file);
              return; // já lidamos com a resolução dentro do fr.onload
            } catch (ex) {
              // Qualquer exceção inesperada aqui não deve interromper o fluxo;
              // apenas continuamos e retornamos o texto original abaixo.
            }
          }

          // Se não encontramos '�' ou não foi possível decodificar em CP1252,
          // retornamos o texto lido originalmente em UTF-8.
          resolve(text);
        };

        // Em caso de erro no FileReader principal, rejeitamos a Promise
        reader.onerror = reject;

        // Inicia a leitura do arquivo como texto em UTF-8
        reader.readAsText(file, 'utf-8');
      });
    }

    // readFileAsArrayBuffer - devolve o ArrayBuffer do arquivo
    // - Útil quando a API consumidora (ex: SheetJS) precisa do binário
    //   do arquivo para parsear XLS/XLSX corretamente.
    function readFileAsArrayBuffer(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        // onload devolve o resultado como ArrayBuffer
        reader.onload = e => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
      });
    }

    // Exporta os helpers como AppIO no escopo global para que outras
    // partes do aplicativo (UI, processor) possam usá-los sem módulos.
    global.AppIO = global.AppIO || {};
    global.AppIO.readFileAsText = readFileAsText;
    global.AppIO.readFileAsArrayBuffer = readFileAsArrayBuffer;
  })(window);

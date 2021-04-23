# jogo_da_velha_inteligente
Desenvolvimento de jogo da velha inteligente em VBA para Excel, usando algoritmo de MinMax.

O algoritmo foi desenvolvido utilizando uma mescla de VBA para Excel e os recursos de planilha.

O conteúdo dos arquivos é o seguinte:
  - Jogo da Velha Inteligente.xlsm: é o arquivo final de Excel em que o jogo da velha foi implementado. Baixe para jogar! Não se esqueça de habilitar os recursos de macro.
  - gatilho_open_file: contém o trigger que é acionado sempre que o arquivo é aberto. Exibe mensagens de inicialização e carrega as fórmulas a serem utilizadas na planilha.
  - gatilho_worksheet_change: contém o gatilho que é acionado sempre que é realizada uma alteração no tabuleiro do jogo.
  - nivel_facil: contém a macro que define o nível fácil (MinMax com árvore de profundidade 2).
  - nivel_medio: contém a macro que define o nível médio (MinMax com árvore de profundidade 3).
  - nivel_dificil: contém a macro que define o nível difícil (MinMax com árvore de profundidade 6).
  - outras_macros: contém as macros de limpeza de tabuleiro e solicitação para máquina jogar primeiro.

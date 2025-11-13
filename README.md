# Dashboard de Notas Fiscais

Aplicação web estática (HTML/CSS/JS) que lê um arquivo Excel de notas fiscais e apresenta indicadores e gráficos interativos.

## Como usar

1. Abra o arquivo `index.html` em qualquer navegador moderno (Chrome, Edge, Firefox).
2. A página tentará carregar automaticamente o arquivo padrão `TESTE222222 - NOTA FISCAL.xlsx` que deve estar na mesma pasta.
3. Caso deseje usar outro arquivo, clique em **Importar outro arquivo** e selecione um Excel (`.xlsx` ou `.xls`).

## Funcionalidades

- **Indicadores principais**: quantidade total de registros, soma e média da métrica selecionada.
- **Controles de análise**:
  - Escolha de coluna categórica para agrupamento.
  - Escolha de coluna numérica para métricas.
  - Escolha de coluna de data (opcional) para séries temporais.
  - Filtro textual por categoria.
  - Intervalo de datas (início/fim).
- **Gráficos interativos**:
  - Gráfico de barras por categoria.
  - Gráfico de linha por período.
- **Tabela detalhada** com rolagem e cabeçalho fixo.
- **Exportação CSV** dos dados filtrados.

## Dependências

CDNs utilizados (já referenciados em `index.html`):

- [SheetJS](https://docs.sheetjs.com/) para leitura do Excel.
- [Chart.js](https://www.chartjs.org/) para visualizações.

## Pré-requisitos

- Navegador com suporte a ES2020.
- O arquivo Excel deve possuir cabeçalhos na primeira linha.

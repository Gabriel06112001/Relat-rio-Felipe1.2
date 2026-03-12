# Relatório CAP – Semana 12

Relatório interativo gerado a partir da planilha `CAP_Semana_12.xlsx`.

## Estrutura

```
├── data/                  # Planilhas de entrada (.xlsx)
├── output/                # Relatórios gerados (.html)
├── src/
│   └── gerar_relatorio.js # Script gerador
├── node_modules/          # Dependências (não versionado)
├── .gitignore
├── package.json
└── README.md
```

## Como usar

### 1. Instalar dependências (apenas na primeira vez)
```bash
npm install
```

### 2. Atualizar a planilha
Substitua o arquivo em `data/CAP_Semana_12.xlsx` pela versão mais recente.

### 3. Gerar o relatório
```bash
npm start
```

O relatório será gerado em `output/relatorio_CAP_Semana12.html`.

## Tecnologias

- **Node.js** — geração do HTML
- **xlsx** — leitura da planilha Excel
- **Chart.js 4.4.4** — gráficos interativos
- **ag-Grid Community 31.3.2** — tabela de dados
- **SheetJS (CDN)** — exportação para `.xlsx` no navegador

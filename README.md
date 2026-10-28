# SICOF · Controle de Liquidação & Restos a Pagar (BCMS)
**Batalhão Central de Manutenção e Suprimento (BCMS) · UGs 160329 (OGU) / 167329 (FEx)**  
**Base de Apoio Logístico (Ba Ap Log)**  

Painel executivo web (Single Page Application estática e autocontida) para **auditoria, rastreamento e controle rigoroso da liquidação da despesa pública e restos a pagar**, construído em estrita consonância com o Plano de Contas Aplicado ao Setor Público (PCASP) e o Sistema Integrado de Administração Financeira (SIAFI).

- **URL Pública (GitHub Pages):** `https://DeCampos603.github.io/dashboard-liquidacao-bcms/`
- **Ambiente Operacional:** Exercício Financeiro 2026
- **Automação CI/CD:** Atualização diária programada via GitHub Actions (`.github/workflows/deploy-pages.yml`)

---

## 🏛️ 1. Módulos e Contas do PCASP Monitoradas

O painel confronta em tempo real a execução das contas da **Classe 6 (Controles da Execução do Orçamento)**:

| Módulo | Código Contábil | Título Oficial da Conta | Função no Painel |
| :--- | :---: | :--- | :--- |
| **Exercício Corrente** | `622110000` | Crédito Disponível | Saldo orçamentário remanescente para novos empenhos |
| **Exercício Corrente** | `622130100` | Crédito Empenhado a Liquidar | Dotação comprometida aguardando atesto e faturamento |
| **Exercício Corrente** | `622920101` | **Empenhos a Liquidar** | **Saldo analítico de cada NE com atesto pendente** |
| **Exercício Corrente** | `622920103` | Empenhos Liquidados a Pagar | Despesa liquidada aguardando emissão de Ordem Bancária |
| **Exercício Corrente** | `622920104` | Empenhos Pagos | Total pago no exercício via OB bancária |
| **Restos a Pagar** | `631100000` | **RPNP a Liquidar (Inscritos)** | **RP Não Processados em risco de bloqueio / cancelamento** |
| **Restos a Pagar** | `631200000` | RPNP em Liquidação | Materiais/serviços de anos anteriores em conferência |
| **Restos a Pagar** | `631300000` | RPNP Liquidado a Pagar | Passivo exigível de exercícios anteriores pendente de OB |
| **Restos a Pagar** | `632100000` | RPP a Pagar | Restos a Pagar Processados |

---

## 🚀 2. Recursos e Diferenciais de Engenharia

- **Arquitetura Autocontida (Zero Runtime Overhead):** O arquivo `site/index.html` contém todo o layout, estilos, scripts e base estruturada injetada inline. Funciona tanto no GitHub Pages quanto com duplo clique local (`file://`).
- **Design Tokens & Acessibilidade WCAG 2.2 AAA:** Contraste reforçado (> 7:1), modo escuro executivo e modo claro com alternância instantânea a 60fps GPU.
- **Matriz de Aging de Liquidação:** Cálculo automático dos dias de atraso entre a data de emissão da NE e a data atual:
  - 🟢 **0-30 dias:** Prazo regular de faturamento/entrega.
  - 🟡 **31-60 dias:** Atenção / cobrança preventiva.
  - 🟠 **61-90 dias:** Alerta / notificação ao fornecedor.
  - 🔴 **> 90 dias:** Risco crítico de inexecução contratual.
- **Gaveta Modal de Auditoria:** Drill-down completo para visualização dos itens de compra, descrição do objeto, processo licitatório e Nota de Crédito de origem.
- **Exportação para Excel (.xls / .csv):** Geração e download imediato da planilha de auditoria diretamente pelo navegador.

---

## 🛠️ 3. Como Publicar no GitHub Pages (Passo a Passo)

### Passo 1: Criar o Repositório no GitHub
1. Acesse: https://github.com/new
2. Nome do repositório: **`dashboard-liquidacao-bcms`**
3. Visibilidade: **Public**
4. Clique em **Create repository** (sem marcar README ou .gitignore, pois já criamos localmente).

### Passo 2: Inicializar o Git e Enviar os Arquivos
No PowerShell, dentro da pasta `dashboard-liquidacao-bcms`:
```powershell
cd "g:\Meu Drive\REPO\dashboard-liquidacao-bcms"
git init
git add .
git commit -m "feat: lancamento do dashboard de liquidacao e restos a pagar do BCMS"
git branch -M main
git remote add origin https://github.com/DeCampos603/dashboard-liquidacao-bcms.git
git push -u origin main
```

### Passo 3: Ativar o GitHub Pages via GitHub Actions
1. No seu repositório no GitHub, clique na aba **Settings**.
2. No menu lateral esquerdo, selecione **Pages**.
3. Na seção **Build and deployment**:
   - Em **Source**, selecione: **`GitHub Actions`**.
4. Pronto! O workflow `.github/workflows/deploy-pages.yml` será disparado automaticamente e publicará o site na URL:
   `https://DeCampos603.github.io/dashboard-liquidacao-bcms/`

---

## 📊 4. Estrutura de Arquivos

```
dashboard-liquidacao-bcms/
├── .github/
│   └── workflows/
│       └── deploy-pages.yml         # Pipeline CI/CD diário p/ GitHub Pages
├── data/
│   └── liquidacao_data.json         # Base estruturada extraída do SIAFI (421 NEs + RAP)
├── site/
│   └── index.html                  # Aplicação SPA estática autocontida (HTML+CSS+JS)
├── scripts/
│   └── extrair_liquidacao_bcms.py   # Extrator ETL de planilhas e fechamento contábil
├── gerar_dashboard.py              # Script orquestrador de build e injeção de dados
├── requirements.txt                # Dependências Python (openpyxl)
├── .gitignore                      # Regras de exclusão do Git
└── README.md                       # Documentação técnica e guia operacional
```

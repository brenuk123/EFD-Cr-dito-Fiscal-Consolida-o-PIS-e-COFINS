# Sistema de Consolidação e Conferência Fiscal — PIS e COFINS (EFD Contribuições)

Aplicação web especializada para automação do fechamento mensal de créditos de PIS e COFINS a partir da planilha padrão de apuração (`EFD 8_2026` e competências equivalentes). O sistema inspeciona diretamente a estrutura XML e a paleta de estilos/temas do arquivo `.xlsx` para identificar registros marcados em azul nas 8 abas obrigatórias, unifica os dados, audita cálculos de alíquotas, cruza totais com a aba `Resumo`, permite ajustes manuais de conferência e exporta uma pasta de trabalho Excel completa e auditável.

## User Review & Critical Decisions

> [!IMPORTANT]
> As decisões abaixo refletem as preferências confirmadas para o fluxo de auditoria, persistência e exportação. Revise o plano e clique em **Proceed** para iniciarmos a construção.

- **Decisão Confirmada 1 — Persistência de Perfis e Histórico Local**: O sistema salvará localmente (`localStorage` / `IndexedDB`) tanto os **perfis de configuração** (mapeamento de seções por aba, colunas de base/PIS/COFINS, alíquotas, tolerância em R$ e paleta de cores azuis reconhecidas) quanto o **histórico de fechamentos mensais** consolidados para consulta rápida e comparação entre competências.
- **Decisão Confirmada 2 — Revisão Manual na Conferência**: Na tela de conferência dos registros, o analista fiscal poderá **incluir ou excluir linhas manualmente** (mesmo linhas que originalmente não estavam azuis ou que foram marcadas por engano) e **marcar divergências de cálculo como "Justificada / Revisada"** com observação opcional, atualizando os totalizadores em tempo real.
- **Decisão Confirmada 3 — Exportação Excel Expandida com Auditoria**: O arquivo `.xlsx` exportado incluirá, além das abas previstas (`Resumo por aba`, `Consolidado` e uma planilha individual para cada categoria com registros), duas abas dedicadas de auditoria: **`Divergências`** (registros com diferença de PIS/COFINS acima da tolerância e seu status de justificativa) e **`Duplicados`** (itens repetidos na mesma aba ou documentos presentes em múltiplas abas).

---

## 1. Overview & Core Concept

- **O Que Faz**: Lê arquivos `.xlsx` de fechamento mensal de PIS/COFINS preservando metadados de preenchimento de célula (cores RGB diretas e cores de tema com *tint*), extrai linhas destacadas em azul nas oito abas-alvo (`Sucata 1`, `Uso e consumo`, `Nfse`, `04 e 05 Outros`, `Locação`, `Imobilizado`, `Frete` e `Resumo`), padroniza os valores de Base de Cálculo, PIS e COFINS, reconcilia automaticamente com a linha "Total" da aba `Resumo`, aponta divergências de alíquota e notas duplicadas, e gera o Excel consolidado final.
- **Público-Alvo / Persona**: Analistas fiscais, contadores e auditores tributários responsáveis pela apuração mensal da EFD Contribuições.
- **Valor Principal**: Elimina a cópia manual de linhas coloridas entre dezenas de blocos de planilhas heterogêneas, reduzindo o risco de omissão de créditos legítimos, captura em duplicidade ou erros de fórmula antes da transmissão da obrigação acessória.

---

## 2. User Experience & Visual Design

### Fluxos Principais do Usuário
1. **Importação e Diagnóstico Inicial (Passo 1)**:
   - O usuário carrega a planilha `.xlsx` do mês (ou carrega uma planilha de demonstração fiel à estrutura `EFD 8_2026` para teste imediato).
   - O motor descompacta o XML nativo do Excel, identifica as 8 abas esperadas, contabiliza o total de linhas e extrai todas as cores de preenchimento presentes nas células, pré-selecionando automaticamente os tons dentro do espectro azul (matiz $190^\circ\text{–}260^\circ$).
2. **Critérios de Identificação e Parametrização de Abas (Passo 2)**:
   - **Seletor Interativo de Cores**: Amostras visuais com código hexadecimal e contagem de células permitem ativar ou desativar quais cores contam como "azul" em um clique.
   - **Perfis de Configuração Salvos**: Barra de perfis permitindo salvar, renomear ou restaurar regras customizadas (ex.: *"Padrão EFD 2026"*, *"Regra Especial Filial"*).
   - **Matriz de Seções por Aba**: Configuração detalhada de linha de cabeçalho, primeira e última linha de dados, coluna inicial/final, linha de rótulo de bloco, cabeçalhos manuais (separados por `;`), nomes das colunas de Base/PIS/COFINS e alíquotas esperadas (%). Já vem pré-configurada para a estrutura multi-seção das abas `04 e 05 Outros` (seção superior com alíquotas 1,65% / 7,60% e seção inferior com cabeçalhos manuais e alíquotas 0,65% / 4,00%) e `Resumo` (blocos lado a lado nas colunas `A:C` e `D:F`).
3. **Painel de Conferência Fiscal e Revisão Manual (Passo 3)**:
   - **Indicadores Executivos (KPIs)**: Total de registros ativos, Base de Cálculo acumulada (R$), PIS apurado (R$), COFINS apurado (R$), Crédito Total do Mês (R$), contagem de divergências pendentes/revisadas e alertas de duplicidade.
   - **Tabela de Consolidação por Categoria & Reconciliação com Resumo**: Visão lado a lado do total apurado por aba e quadro comparativo entre a linha `"Total"` oficial da aba `Resumo` versus a soma das linhas azuis capturadas na própria aba `Resumo`, destacando qualquer diferença acima da tolerância configurada (padrão `R$ 0,02`).
   - **Grade de Auditoria de Registros**: Tabela de alta densidade com filtros combinados (por aba de origem, busca textual livre, status de cálculo, duplicados e linhas incluídas/excluídas manualmente). Permite:
     - Alternar a inclusão/exclusão individual de linhas no consolidado final (ou inspecionar todas as linhas da aba original para incluir manualmente uma linha esquecida sem cor azul).
     - Marcar uma divergência de cálculo como **"Justificada / Revisada"** com nota de auditoria, separando divergências críticas de exceções fiscais conhecidas.
4. **Histórico de Fechamentos e Exportação Excel (Passo 4)**:
   - Salvamento do fechamento atual no histórico local da aplicação com carimbo de data/hora, nome do arquivo e resumo dos créditos.
   - Exportação instantânea de arquivo `.xlsx` formatado (cabeçalhos congelados, autofiltro, máscara contábil `#,##0.00`) contendo: `Resumo por aba`, `Consolidado`, `Divergências`, `Duplicados` e uma aba individual para cada categoria.

### Identidade Visual e Tema
- **Direção Estética**: *Precision Tax & Financial Ledger Workspace* — interface corporativa limpa, de alta densidade informacional, inspirada em plataformas modernas de controladoria e auditoria contábil.
- **Paleta de Cores (Regra 60-30-10)**:
  - **60% Canvas Neutro Dominante**: Fundo principal em cinza-gelo frio de baixa saturação (`#F8FAFC` / `slate-50`) com superfícies de trabalho em branco puro (`#FFFFFF`).
  - **30% Superfícies Estruturais e Bordas**: Divisórias precisas de `1px solid #E2E8F0` (`slate-200`), cabeçalhos de tabela em `#F1F5F9` (`slate-100`) e tipografia primária em `#0F172A` (`slate-900`) com metadados secundários em `#475569` (`slate-600`).
  - **10% Orçamento de Acento e Semântica**: Azul institucional profundo (`#1E40AF` / `#1F4E79`) para ações primárias e seleção de abas; estados semânticos sempre acompanhados de rótulo textual explícito e ícone funcional:
    - Conformidade / OK: Verde esmeralda (`#15803D`)
    - Divergência Revisada / Alerta: Âmbar (`#B45309`)
    - Divergência Pendente / Duplicidade: Vermelho carmesim (`#B91C1C`)
- **Tipografia e Hierarquia (Regra 2+1 Fontes)**:
  - **Display & Títulos Estruturais**: `Plus Jakarta Sans` (SemiBold/Bold 600–700) para cabeçalhos de seção e navegação.
  - **Corpo & Controles de Interface**: `Public Sans` (Regular/Medium 400–500) para leitura nítida em controles densos de 12px–14px.
  - **Dados Numéricos, Alíquotas, Linhas e Códigos Hex**: `JetBrains Mono` com `font-variant-numeric: tabular-nums` obrigatório em todas as colunas financeiras (Base, PIS, COFINS, Diferenças), números de linha do Excel e contadores para alinhamento vertical perfeito de casas decimais.
- **Arquitetura de Componentes e Layout**:
  - **Top Bar Contract**: Barra superior única com 3 zonas separadas por `gap-8`: (1) Marca em linha única (`EFD Crédito Fiscal`), (2) Links de navegação diretos entre as visões principais (`Importação`, `Critérios e Regras`, `Conferência Fiscal`, `Histórico Mensal`), e (3) Ação primária contextual (`Exportar Consolidado .xlsx`).
  - **Disciplina Zero-Pill**: Metadados estáticos (aba de origem, linha do Excel, contagens) apresentados como texto limpo com separadores tipográficos (`·`) ou colunas tabulares alinhadas; controles de filtro funcionam como botões segmentados interativos.
  - **Elevação Única**: Superfícies planas delimitadas por bordas sutis de `1px` (`border-slate-200`), sem cartões aninhados ou barras laterais grossas coloridas.

---

## 3. Key Product Decisions & Trade-Offs

- **Decisão 1 — Leitura Direta do XML Interno do `.xlsx` (via JSZip) + Geração via ExcelJS**
  - *Abordagem Escolhida*: Descompactar o arquivo `.xlsx` diretamente no navegador lendo `xl/workbook.xml`, `xl/styles.xml`, `xl/theme/theme1.xml`, `xl/sharedStrings.xml` e `xl/worksheets/sheet*.xml` para extração exata das cores de preenchimento (incluindo cálculo de *HSL tint* sobre cores de tema do Office), combinada ao `ExcelJS` para gerar a planilha exportada estilizada.
  - *Por Quê*: Bibliotecas genéricas de planilha frequentemente descartam ou distorcem cores baseadas em `theme` + `tint` do Excel ou travam a thread principal em planilhas grandes. O parser XML dedicado garante fidelidade de 100% à lógica validada no protótipo do usuário e altíssima velocidade.
  - *Alternativas Consideradas*: Usar apenas SheetJS Community (que não expõe estilos de preenchimento na versão gratuita) ou processar no backend (desnecessário e mais lento para arquivos locais).

- **Decisão 2 — Mecanismo Não-Destrutivo de Inclusão/Exclusão Manual e Justificativa de Divergências**
  - *Abordagem Escolhida*: Manter um mapa de estado de auditoria (`manualOverrides`) indexado por `aba + linha` que registra se uma linha foi incluída manualmente, excluída manualmente do cálculo ou se sua divergência foi justificada pelo analista.
  - *Por Quê*: Permite que o usuário reexecute a consolidação ou troque filtros de cor sem perder as justificativas e revisões manuais feitas naquela sessão, além de refletir instantaneamente nos cartões de totalização e no Excel exportado.

- **Decisão 3 — Gerador Integrado de Planilha Exemplo (`EFD 8_2026 Demo`)**
  - *Abordagem Escolhida*: Disponibilizar na tela de importação, ao lado do seletor de arquivo `.xlsx`, um botão para carregar ou baixar uma planilha `.xlsx` real de demonstração com as 8 abas, linhas azuis, seções especiais (`04 e 05 Outros` e `Resumo`), casos de divergência e documentos duplicados.
  - *Por Quê*: Permite testar e validar todo o fluxo de ponta a ponta imediatamente, mesmo quando o analista estiver em um computador sem o arquivo mensal em mãos.

---

## 4. Technical Architecture & Data Strategy

### Diagrama de Arquitetura e Fluxo de Dados

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                        TOP BAR (3-Zone Navigation)                          │
│  [EFD Crédito Fiscal]     [Importação | Critérios | Conferência | Histórico]│
└──────────────────────────────────────┬──────────────────────────────────────┘
                                       │
       ┌───────────────────────────────┼───────────────────────────────┐
       ▼                               ▼                               ▼
┌──────────────────────┐   ┌───────────────────────┐   ┌──────────────────────┐
│ 1. Motor de Leitura  │   │ 2. Gestor de Regras   │   │ 3. Persistência Local│
│    XML (.xlsx/JSZip) │   │    & Cores Azuis      │   │    (localStorage)    │
├──────────────────────┤   ├───────────────────────┤   ├──────────────────────┤
│ • workbook & rels    │   │ • Seletor Hex/HSL     │   │ • Perfis de Abas/    │
│ • styles & theme1    │   │ • Seções por Aba      │   │   Seções e Alíquotas │
│   (RGB + Tint HSL)   │   │   (04/05 Outros,      │   │ • Histórico de       │
│ • sharedStrings      │   │    Resumo, Padrão)    │   │   Fechamentos        │
│ • Linhas & Células   │   │ • Tolerância & Ocultas│   │ • Overrides Manuais  │
└──────────┬───────────┘   └───────────┬───────────┘   └──────────┬───────────┘
           │                           │                          │
           └───────────────────────────┼──────────────────────────┘
                                       ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                   4. Motor de Consolidação & Auditoria Fiscal               │
├─────────────────────────────────────────────────────────────────────────────┤
│ • Filtragem de linhas: (Célula com Cor Azul Ativa OU Inclusão Manual)       │
│   E NÃO (Exclusão Manual)                                                   │
│ • Padronização de Base, PIS e COFINS + Cálculo de PIS/COFINS Esperado       │
│ • Diagnóstico de Status: OK | Divergente | Justificado | Sem Dados          │
│ • Detecção de Duplicidade: Chave de Item (Doc+Série+Forn+Loja+Item)         │
│   e Cruzamento entre Abas (mesmo documento em múltiplas categorias)         │
│ • Reconciliação com a linha "Total" da aba Resumo                           │
└──────────────────┬───────────────────────────────────────┬──────────────────┘
                   ▼                                       ▼
┌──────────────────────────────────────┐   ┌──────────────────────────────────┐
│ 5. Painel Interativo de Conferência  │   │ 6. Exportador Excel (ExcelJS)    │
├──────────────────────────────────────┤   ├──────────────────────────────────┤
│ • KPIs Totais & Resumo por Categoria │   │ • Aba 1: Resumo por aba + Check  │
│ • Comparativo Aba Resumo vs Capturado│   │ • Aba 2: Consolidado Geral       │
│ • Grade com Busca, Filtros, Inclusão/│   │ • Aba 3: Divergências & Notas    │
│   Exclusão Manual e Justificativa    │   │ • Aba 4: Duplicados              │
│ • Explorador de Todas as Linhas da   │   │ • Abas 5..N: Categorias Ativas   │
│   Aba para Inclusão Avulsa           │   │   (Sucata 1, Uso e consumo, etc.)│
└──────────────────────────────────────┘   └──────────────────────────────────┘
```

### Modelo de Dados e Estado
- **`ParsedWorkbook`**: Estrutura em memória contendo as 8 abas mapeadas (`SH`), dicionário de frequência de cores hexadecimais encontradas (`COLORS`), conjunto de cores classificadas como azuis (`BLUE`) e metadados do arquivo carregado.
- **`TabSectionConfig`**: Lista de regras por aba contendo `h` (linha de cabeçalho), `a` (1ª linha de dados), `b` (última linha), `c1`/`c2` (intervalo de colunas), `lab` (linha de rótulo de bloco), `man` (cabeçalhos manuais separados por `;`), mapeamento dos campos `base`, `pis`, `cof` e alíquotas `rp` (%) e `rc` (%).
- **`ManualAuditState`**:
  - `excludedRowKeys`: Conjunto de chaves `aba|linha` desmarcadas manualmente pelo usuário para não compor o crédito.
  - `includedRowKeys`: Conjunto de chaves `aba|linha` incluídas manualmente pelo usuário mesmo sem preenchimento azul original.
  - `justifiedMap`: Mapa `aba|linha -> { note: string, timestamp: string }` para divergências revisadas/justificadas.
- **`ConsolidatedRecord`**: Registro normalizado com `aba`, `linha`, `d` (dicionário de colunas originais), `base`, `pis`, `cof`, `ep` (PIS esperado), `ec` (COFINS esperado), `dp` (diferença PIS), `dc` (diferença COFINS), `st` (`OK` | `Divergente` | `Justificado` | `Sem dados`), `dup` (booleano de duplicidade de item), `mx` (outras abas onde o mesmo documento aparece) e flags de revisão manual (`isManualIncluded`, `isExcluded`, `justificationNote`).
- **`SavedProfile` & `ClosingHistoryEntry`**: Registros persistidos em `localStorage` para reutilização de regras mensais e consulta de fechamentos passados.

### Mapeamento de Interações e Estado
- **Seleção de Cores (`ColorChipSelector`)**: Clicar em uma cor alterna sua presença no `Set` de cores azuis e recalcula imediatamente os registros consolidados e KPIs.
- **Editor de Seções (`SectionConfigGrid`)**: Alterar qualquer parâmetro de linha/coluna/alíquota, adicionar `+ seção` ou remover `×` atualiza a configuração ativa e permite salvar em um Perfil nomeado.
- **Revisão na Grade (`RecordsAuditTable`)**:
  - **Incluir/Excluir Linha**: Checkbox/botão por linha alterna a participação do registro nos totalizadores gerais, na reconciliação e na exportação (com opção de visualizar ou ocultar linhas excluídas e modal rápido para inspecionar linhas não-azuis da planilha caso o analista queira incluir alguma manualmente).
  - **Justificar Divergência**: Botão em registros divergentes abre um popover/input enxuto para registrar a justificativa fiscal (ex.: *"Alíquota diferenciada / monofásico"*), alterando o status para `Justificado` e abatendo do contador de divergências pendentes.
- **Exportação (`ExcelExporter`)**: Gera o buffer `.xlsx` via `ExcelJS` respeitando apenas os registros ativos (ou listando o status e a justificativa nas abas de `Consolidado` e `Divergências`) e dispara o download imediato pelo navegador.

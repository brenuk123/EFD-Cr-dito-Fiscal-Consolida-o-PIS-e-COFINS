export const EFD_TABS = [
  'Sucata 1',
  'Uso e consumo',
  'Nfse',
  '04 e 05 Outros',
  'Locação',
  'Imobilizado',
  'Frete',
  'Resumo',
] as const;

export type EfdTabName = (typeof EFD_TABS)[number] | string;

export interface TabSectionConfig {
  id: string;
  h: number | string; // Linha de cabeçalho (padrão EFD 8_2026: 3)
  a: number | string; // 1ª linha de dados (padrão EFD 8_2026: 4)
  b: number | string; // Última linha (opcional)
  c1: string;         // Coluna de (ex: A ou C)
  c2: string;         // Coluna até (ex: J, G, C, F)
  lab: number | string; // Linha rótulo do bloco (ex: Resumo linha 2)
  man: string;        // Cabeçalhos manuais separados por ;
  base: string;       // Campo Base de Cálculo
  pis: string;        // Campo PIS
  cof: string;        // Campo COFINS
  rp: number | string; // Alíquota PIS (%)
  rc: number | string; // Alíquota COFINS (%)
}

export type TabConfigMap = Record<string, TabSectionConfig[]>;

export interface ParsedCell {
  v: string | number;
  bg: string | null;
}

export interface ParsedRow {
  hidden: boolean;
  cells: Map<number, ParsedCell>;
}

export interface ParsedSheet {
  name: string;
  rows: Map<number, ParsedRow>;
}

export interface ParsedWorkbook {
  fileName: string;
  fileSize: number;
  loadedAt: string;
  sheets: Record<string, ParsedSheet>;
  allSheetNames: string[];
  colorCounts: Record<string, number>;
  defaultBlueColors: string[];
  importLogs: TabImportDiag[];
}

export interface TabImportDiag {
  tabName: string;
  found: boolean;
  rowCount: number;
  blueCellCount: number;
}

export type CalculationStatus = 'OK' | 'Divergente' | 'Justificado' | 'Sem dados';

export interface ConsolidatedRecord {
  key: string; // `${aba}|${linha}|${sectionIdx}|${c1}`
  aba: string;
  linha: number;
  sectionIndex: number;
  d: Record<string, string | number>;
  base: number;
  pis: number;
  cof: number;
  ep: number;
  ec: number;
  dp: number;
  dc: number;
  rawStatus: 'OK' | 'Divergente' | 'Sem dados';
  st: CalculationStatus;
  dup: boolean;
  mx: string[];
  dk: string;
  ik: string;
  hasBlueCell: boolean;
  matchedBlueColors: string[];
  isManualIncluded: boolean;
  isExcluded: boolean;
  justificationNote: string;
  justifiedAt?: string;
}

export interface RawSheetCandidateRow {
  key: string;
  aba: string;
  linha: number;
  sectionIndex: number;
  hasBlueCell: boolean;
  matchedBlueColors: string[];
  isCurrentlyIncluded: boolean;
  isManualIncluded: boolean;
  isExcluded: boolean;
  previewText: string;
  base: number;
  pis: number;
  cof: number;
  d: Record<string, string | number>;
}

export interface JustificationEntry {
  note: string;
  updatedAt: string;
}

export interface ManualAuditState {
  includedRowKeys: string[];
  excludedRowKeys: string[];
  justifications: Record<string, JustificationEntry>;
}

export interface TabAggregation {
  aba: string;
  n: number;
  excludedCount: number;
  manualIncludedCount: number;
  base: number;
  pis: number;
  cof: number;
  tot: number;
  div: number;
  justified: number;
  dup: number;
}

export interface ResumoTotalComparisonRow {
  item: string;
  resumoTotalLine: number;
  resumoBlueRows: number;
  diff: number;
  isOk: boolean;
}

export interface ResumoComparisonBlock {
  bloco: string;
  colRange: string;
  baseResumo: number | null;
  pisResumo: number | null;
  cofResumo: number | null;
  baseAzuis: number;
  pisAzuis: number;
  cofAzuis: number;
  diffBase: number | null;
  diffPis: number | null;
  diffCof: number | null;
  status: 'OK' | 'Diferença' | 'Sem linha Total';
}

export interface SavedProfile {
  id: string;
  name: string;
  updatedAt: string;
  tolerance: number;
  ignoreHidden: boolean;
  blueColors: string[];
  tabConfig: TabConfigMap;
}

export interface ClosingHistoryEntry {
  id: string;
  competencia: string;
  fileName: string;
  savedAt: string;
  totalRecords: number;
  totalBase: number;
  totalPis: number;
  totalCofins: number;
  totalCredit: number;
  pendingDivergences: number;
  justifiedDivergences: number;
  duplicateCount: number;
  byTab: TabAggregation[];
  notes?: string;
}

export const STD_SECTION_DEFAULT: Omit<TabSectionConfig, 'id'> = {
  h: 3,
  a: 4,
  b: '',
  c1: 'A',
  c2: '',
  lab: '',
  man: '',
  base: 'Base Pis/Cofins',
  pis: 'Pis',
  cof: 'Cofins',
  rp: 1.65,
  rc: 7.6,
};

/**
 * Exact default configuration matching the EFD 8_2026 structure
 */
export function createDefaultTabConfig(): TabConfigMap {
  const cfg: TabConfigMap = {};
  EFD_TABS.forEach((t, idx) => {
    cfg[t] = [
      {
        id: `std-${idx}-0`,
        ...STD_SECTION_DEFAULT,
      },
    ];
  });

  cfg['04 e 05 Outros'] = [
    {
      id: 'outros-0',
      ...STD_SECTION_DEFAULT,
      b: 36,
      c1: 'C',
      c2: 'J',
      base: 'B.C',
    },
    {
      id: 'outros-1',
      ...STD_SECTION_DEFAULT,
      h: '',
      a: 39,
      b: 70,
      c1: 'C',
      c2: 'G',
      man: 'Conta;Descrição;Valor;Pis;Cofins',
      base: 'Valor',
      pis: 'Pis',
      cof: 'Cofins',
      rp: 0.65,
      rc: 4.0,
    },
  ];

  cfg['Resumo'] = [
    {
      id: 'resumo-0',
      ...STD_SECTION_DEFAULT,
      b: 5,
      c1: 'A',
      c2: 'C',
      lab: 2,
      base: 'Soma de Vl Mês',
      pis: 'Soma de PIS',
      cof: 'Soma de COFINS',
    },
    {
      id: 'resumo-1',
      ...STD_SECTION_DEFAULT,
      b: 5,
      c1: 'D',
      c2: 'F',
      lab: 2,
      base: 'Soma de Vl Mês',
      pis: 'Soma de PIS',
      cof: 'Soma de COFINS',
    },
  ];

  return cfg;
}

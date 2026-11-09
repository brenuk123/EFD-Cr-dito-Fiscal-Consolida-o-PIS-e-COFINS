import { ConsolidatedRecord } from '../types/efd';
import { norm } from './excelParser';

export interface AccountingMapEntry {
  ctaLancamento: string;
  ctaCredito: string;
  descricao: string;
}

export interface AccountingPivotAccountRow {
  rowLabel: string; // Ex: '11040400001' ou 'DIESEL S10'
  ctaLancamento: string;
  ctaCredito: string;
  descricaoMapa: string;
  somaPis: number;
  somaCofins: number;
  historicoPis: string;
  historicoCofins: string;
  isMapped: boolean;
}

export interface AccountingPivotFilialGroup {
  filial: string;
  subtotalPis: number;
  subtotalCofins: number;
  accounts: AccountingPivotAccountRow[];
}

export interface AccountingCategoryBlock {
  blockCode: string; // Ex: '02 - Uso e Consumo'
  sourceTab: string; // Ex: 'Uso e consumo'
  ctaDebitoPis: string; // '11070100002'
  ctaDebitoCofins: string; // '11070100003'
  filialGroups: AccountingPivotFilialGroup[];
  totalPis: number;
  totalCofins: number;
}

export const DEFAULT_CTA_DEBITO_PIS = '11070100002';
export const DEFAULT_CTA_DEBITO_COFINS = '11070100003';

export const CATEGORY_BLOCK_DEFS: Array<{
  blockCode: string;
  sourceTab: string;
}> = [
  { blockCode: '02 - Uso e Consumo', sourceTab: 'Uso e consumo' },
  { blockCode: '03 - NFSe', sourceTab: 'Nfse' },
  { blockCode: '06 - Locação', sourceTab: 'Locação' },
  { blockCode: '10 - Imobilizado', sourceTab: 'Imobilizado' },
  { blockCode: '14 - Frete', sourceTab: 'Frete' },
  { blockCode: 'XX - Sucatas', sourceTab: 'Sucata 1' },
  { blockCode: 'XX - Outros / Diesel', sourceTab: '04 e 05 Outros' },
];

export const DEFAULT_ACCOUNTING_MAP: AccountingMapEntry[] = [
  { ctaLancamento: '11040100001', ctaCredito: '41010200001', descricao: 'MU01 - EPI' },
  { ctaLancamento: '11040100002', ctaCredito: '41010200001', descricao: 'MU01 - EPI' },
  { ctaLancamento: '11040100003', ctaCredito: '41010200001', descricao: 'MU01 - EPI' },
  { ctaLancamento: '11040100004', ctaCredito: '41010200001', descricao: 'MU01 - EPI' },
  { ctaLancamento: '11040100005', ctaCredito: '41010200002', descricao: 'MU01 - EPI' },
  { ctaLancamento: '11040200001', ctaCredito: '41010200002', descricao: 'MU02 - UNIFORMES' },
  { ctaLancamento: '11040200002', ctaCredito: '41010200002', descricao: 'MU02 - UNIFORMES' },
  { ctaLancamento: '11040200003', ctaCredito: '41010200002', descricao: 'MU02 - UNIFORMES' },
  { ctaLancamento: '11040300001', ctaCredito: '41010200003', descricao: 'MU03 - MATERIAL DE LIMPEZA' },
  { ctaLancamento: '11040400001', ctaCredito: '41010200004', descricao: 'MU04 - OLEOS  E GRAXAS' },
  { ctaLancamento: '11040400002', ctaCredito: '41010200004', descricao: 'MU04 - OLEOS  E GRAXAS' },
  { ctaLancamento: '11040400003', ctaCredito: '41010200004', descricao: 'MU04 - OLEOS  E GRAXAS' },
  { ctaLancamento: '11040400004', ctaCredito: '41010200004', descricao: 'MU04 - OLEOS  E GRAXAS' },
  { ctaLancamento: '11040400005', ctaCredito: '41010200004', descricao: 'MU04 - OLEOS  E GRAXAS' },
  { ctaLancamento: '11040400006', ctaCredito: '41010200004', descricao: 'MU04 - OLEOS  E GRAXAS' },
  { ctaLancamento: '11040500001', ctaCredito: '41010200005', descricao: 'MU06 - PECAS' },
  { ctaLancamento: '11040500002', ctaCredito: '41010200005', descricao: 'MU06 - PECAS' },
  { ctaLancamento: '11040500003', ctaCredito: '41010200005', descricao: 'MU06 - PECAS' },
  { ctaLancamento: '11040500004', ctaCredito: '41010200005', descricao: 'MU06 - PECAS' },
  { ctaLancamento: '11040500005', ctaCredito: '41010200005', descricao: 'MU06 - PECAS' },
  { ctaLancamento: '11040600001', ctaCredito: '41010200006', descricao: 'MU07 - GASES' },
  { ctaLancamento: '11040700001', ctaCredito: '41010200019', descricao: 'MU17 - USO E CONSUMO' },
  { ctaLancamento: '11040800001', ctaCredito: '41010200008', descricao: 'MU09 - MATERIAL ELETRICO' },
  { ctaLancamento: '11040900003', ctaCredito: '41010200008', descricao: 'MU09 - MATERIAL ELETRICO' },
  { ctaLancamento: '11041000001', ctaCredito: '41010200010', descricao: 'MU11 - USO E CONSUMO - MATERIAIS DE CONS' },
  { ctaLancamento: '11041100002', ctaCredito: '41010200010', descricao: 'MU11 - USO E CONSUMO - MATERIAIS DE CONS' },
  { ctaLancamento: '11041200001', ctaCredito: '41010200018', descricao: 'MU05 - FILTROS' },
  { ctaLancamento: '11041200002', ctaCredito: '41010200018', descricao: 'MU05 - FILTROS' },
  { ctaLancamento: '11041200003', ctaCredito: '41010200018', descricao: 'MU05 - FILTROS' },
  { ctaLancamento: '11041300001', ctaCredito: '41010200014', descricao: 'MU14 - FERRAMENTAS' },
  { ctaLancamento: '11041400001', ctaCredito: '41010200014', descricao: 'MU14 - FERRAMENTAS' },
  { ctaLancamento: '11041500001', ctaCredito: '41010200015', descricao: 'MU15 - MATERIAL DE INFORMATICA' },
  { ctaLancamento: '11041600001', ctaCredito: '41010200016', descricao: 'MU16 - COMBUSTIVEL' },
  { ctaLancamento: '11041700001', ctaCredito: '41010200019', descricao: 'MU17 - USO E CONSUMO' },
  { ctaLancamento: '11041800001', ctaCredito: '41010200017', descricao: 'MU18 - PERFIS E CHAPAS METAIS' },
  { ctaLancamento: '11041900001', ctaCredito: '41010200013', descricao: 'MU19 - SOLDAS E ABRASIVOS' },
  { ctaLancamento: '12050100001', ctaCredito: '12050100001', descricao: 'AI01 - AT - EQUIPAMENTO' },
  { ctaLancamento: '12050100002', ctaCredito: '12050100002', descricao: 'AI02 - AT - CACAMBAS' },
  { ctaLancamento: '12050100003', ctaCredito: '12050100003', descricao: 'AI03 - AT - TI' },
  { ctaLancamento: '12050100004', ctaCredito: '12050100004', descricao: 'AI04 - AT - VEICULOS' },
  { ctaLancamento: '12050100005', ctaCredito: '12050100005', descricao: 'AI05 - AT - MOVEIS E UTENS' },
  { ctaLancamento: '12050100007', ctaCredito: '12050100007', descricao: 'AI07 - AT - INSTALACOES' },
  { ctaLancamento: '41010200001', ctaCredito: '41010200001', descricao: 'MU01 - EPI' },
  { ctaLancamento: '42010200025', ctaCredito: '42040400001', descricao: 'MANUT PREVENTIVA BALANCA' },
  { ctaLancamento: '42040300001', ctaCredito: '42040300001', descricao: 'SV03 - FRETES DE VENDAS' },
  { ctaLancamento: '42040300002', ctaCredito: '42040300002', descricao: 'SV03 - FRETES DE COMPRAS' },
  { ctaLancamento: '42040300005', ctaCredito: '42040300005', descricao: 'SV03 - FRETES DE EQUIPAMENTOS' },
  { ctaLancamento: '42040400001', ctaCredito: '42040400001', descricao: 'SV05 -  MANUTENCAO DE EQUIPAMENTOS' },
  { ctaLancamento: '42040400002', ctaCredito: '42040400002', descricao: 'SV05 - MANUTENCAO DE VEICULOS' },
  { ctaLancamento: '42040400005', ctaCredito: '42040400005', descricao: 'SV05 - MANUTENCAO PREDIAL' },
  { ctaLancamento: '42040400008', ctaCredito: '42040400008', descricao: 'SERVICO DE USINAGEM' },
  { ctaLancamento: '42040500001', ctaCredito: '42040500001', descricao: 'SV06 - SERVICOS ADMINISTRATIVOS' },
  { ctaLancamento: '42040500002', ctaCredito: '42040500002', descricao: 'SV06 - SERVICOS IMPORTACAO' },
  { ctaLancamento: '42040600001', ctaCredito: '42040600001', descricao: 'AL01 - ALUGUEIS DE EQUIPAMENTOS' },
  { ctaLancamento: '42040600002', ctaCredito: '42040600002', descricao: 'AL02 - ALUGUEIS DE IMOVEIS' },
  { ctaLancamento: '42040600004', ctaCredito: '42040600004', descricao: 'AL04 - ALUGUEL DE VEICULOS' },
  { ctaLancamento: '42010200039', ctaCredito: '42010200039', descricao: 'SERV DESTINACAO RESIDUOS INDUSTRIAIS' },
  { ctaLancamento: '42040500020', ctaCredito: '42040500020', descricao: 'SV06 - ELABORACAO DE PROJETOS' },
  { ctaLancamento: '42040500008', ctaCredito: '42040500008', descricao: 'SV06 - CONSULTORIAS' },
  { ctaLancamento: '42040200080', ctaCredito: '42040200080', descricao: 'MU05 - FILTRO' },
  { ctaLancamento: '11030100002', ctaCredito: '41010100010', descricao: 'CPV - CUSTOS DOS PRODUTOS VENDIDOS' },
  { ctaLancamento: '11030100004', ctaCredito: '41010100010', descricao: 'CPV - CUSTOS DOS PRODUTOS VENDIDOS' },
  { ctaLancamento: '11030100003', ctaCredito: '41010100010', descricao: 'CPV - CUSTOS DOS PRODUTOS VENDIDOS' },
  { ctaLancamento: '11030100012', ctaCredito: '41010100010', descricao: 'CPV - CUSTOS DOS PRODUTOS VENDIDOS' },
  { ctaLancamento: '11030100005', ctaCredito: '41010100011', descricao: 'CPV - CUSTOS DOS PRODUTOS VENDIDOS' },
];

const ACCOUNTING_MAP_STORAGE_KEY = 'efd_pis_cofins_accounting_map_v1';

export function loadAccountingMap(): AccountingMapEntry[] {
  try {
    const raw = localStorage.getItem(ACCOUNTING_MAP_STORAGE_KEY);
    if (!raw) return DEFAULT_ACCOUNTING_MAP;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) && parsed.length > 0
      ? parsed
      : DEFAULT_ACCOUNTING_MAP;
  } catch {
    return DEFAULT_ACCOUNTING_MAP;
  }
}

export function saveAccountingMapToStorage(entries: AccountingMapEntry[]): void {
  try {
    localStorage.setItem(ACCOUNTING_MAP_STORAGE_KEY, JSON.stringify(entries));
  } catch (err) {
    console.error('Erro ao salvar mapa contábil:', err);
  }
}

/**
 * Infers MM/YYYY from the workbook file name (e.g. "EFD 8_2026.xlsx" -> "08/2026")
 */
export function inferCompetenciaFromFileName(fileName: string): string {
  const m = /(\d{1,2})[_\-/](\d{4})/.exec(fileName || '');
  if (m) {
    const month = m[1].padStart(2, '0');
    return `${month}/${m[2]}`;
  }
  return '08/2026';
}

function extractFilial(record: ConsolidatedRecord): string {
  const d = record.d;
  const candidateKeys = [
    'filial',
    'cod filial',
    'estabelecimento',
    'estab',
    'loja',
  ];
  for (const ck of candidateKeys) {
    if (d[ck] !== undefined && String(d[ck]).trim() !== '') {
      const raw = String(d[ck]).trim();
      // If it's a 2-digit code like '01', normalize to '0101' if appropriate or keep as-is
      if (/^\d{4}$/.test(raw)) return raw;
      if (raw === '01') return '0101';
      if (raw === '02') return '0102';
      if (raw === '03') return '0103';
      return raw;
    }
  }
  return '0101';
}

function extractAccountAndRowLabel(
  record: ConsolidatedRecord,
  mapLookup: Map<string, AccountingMapEntry>
): { rowLabel: string; ctaLancamento: string } {
  const d = record.d;

  // 1. Check explicit account column names
  const accountColKeys = [
    'conta',
    'cta lancamento',
    'conta contabil',
    'cta contabil',
    'contabil',
    'c contabil',
    'conta debito',
    'cod conta',
  ];

  for (const ak of accountColKeys) {
    for (const [k, v] of Object.entries(d)) {
      if (k === ak || k.includes(ak)) {
        const cleanVal = String(v ?? '')
          .trim()
          .replace(/\D/g, '');
        if (cleanVal.length >= 8) {
          return { rowLabel: cleanVal, ctaLancamento: cleanVal };
        }
      }
    }
  }

  // 2. Scan all values in row for an 11-digit code or a code present in mapLookup
  for (const val of Object.values(d)) {
    const s = String(val ?? '').trim();
    if (mapLookup.has(s)) {
      return { rowLabel: s, ctaLancamento: s };
    }
    if (/^\d{11}$/.test(s)) {
      return { rowLabel: s, ctaLancamento: s };
    }
  }

  // 3. Check if description or product is Diesel / Fuel (like in XX - Diesel block of the user's screenshot)
  const allText = Object.values(d)
    .map((v) => String(v ?? '').toUpperCase())
    .join(' ');
  const descCol = String(
    d['descricao'] || d['descricao do item'] || d['produto'] || ''
  )
    .trim()
    .toUpperCase();

  if (allText.includes('DIESEL') || allText.includes('COMBUSTIVEL')) {
    return {
      rowLabel: descCol || '11041600001',
      ctaLancamento: '11041600001',
    };
  }

  // 4. Fallback defaults by tab category if a sheet doesn't have an explicit account column
  const fallbackByTab: Record<string, string> = {
    'Sucata 1': '11030100002',
    'Uso e consumo': '11041700001',
    Nfse: '42040400001',
    Locação: '42040600001',
    Imobilizado: '12050100001',
    Frete: '42040300002',
    '04 e 05 Outros': '11041600001',
  };

  const fallbackAcc = fallbackByTab[record.aba] || 'N/I';
  const rawContaVal = String(d['conta'] ?? '').trim();
  return {
    rowLabel: rawContaVal || fallbackAcc,
    ctaLancamento: rawContaVal && mapLookup.has(rawContaVal) ? rawContaVal : fallbackAcc,
  };
}

export function buildAccountingDynamics(params: {
  allRecords: ConsolidatedRecord[];
  accountingMap: AccountingMapEntry[];
  competencia: string;
  ctaDebitoPis?: string;
  ctaDebitoCofins?: string;
}): AccountingCategoryBlock[] {
  const {
    allRecords,
    accountingMap,
    competencia,
    ctaDebitoPis = DEFAULT_CTA_DEBITO_PIS,
    ctaDebitoCofins = DEFAULT_CTA_DEBITO_COFINS,
  } = params;

  const activeRecords = allRecords.filter((r) => !r.isExcluded);

  // Build lookup dictionary (first match wins, preserving exact order)
  const mapLookup = new Map<string, AccountingMapEntry>();
  accountingMap.forEach((entry) => {
    const key = entry.ctaLancamento.trim();
    if (key && !mapLookup.has(key)) {
      mapLookup.set(key, entry);
    }
  });

  const cleanComp = competencia.trim() || '08/2026';

  return CATEGORY_BLOCK_DEFS.map(({ blockCode, sourceTab }) => {
    const tabRecords = activeRecords.filter((r) => norm(r.aba) === norm(sourceTab));

    // Group by Filial -> RowLabel|CtaLancamento
    const filialMap = new Map<
      string,
      Map<
        string,
        {
          rowLabel: string;
          ctaLancamento: string;
          somaPis: number;
          somaCofins: number;
        }
      >
    >();

    tabRecords.forEach((rec) => {
      const filial = extractFilial(rec);
      const { rowLabel, ctaLancamento } = extractAccountAndRowLabel(rec, mapLookup);
      const groupKey = `${rowLabel}|${ctaLancamento}`;

      if (!filialMap.has(filial)) {
        filialMap.set(filial, new Map());
      }
      const accMap = filialMap.get(filial)!;
      const existing = accMap.get(groupKey) || {
        rowLabel,
        ctaLancamento,
        somaPis: 0,
        somaCofins: 0,
      };
      existing.somaPis += rec.pis || 0;
      existing.somaCofins += rec.cof || 0;
      accMap.set(groupKey, existing);
    });

    const sortedFiliais = [...filialMap.keys()].sort((a, b) => a.localeCompare(b));

    let totalPis = 0;
    let totalCofins = 0;

    const filialGroups: AccountingPivotFilialGroup[] = sortedFiliais.map((filial) => {
      const accMap = filialMap.get(filial)!;
      let subtotalPis = 0;
      let subtotalCofins = 0;

      const accounts: AccountingPivotAccountRow[] = [...accMap.values()].map(
        (item) => {
          subtotalPis += item.somaPis;
          subtotalCofins += item.somaCofins;

          const mapped = mapLookup.get(item.ctaLancamento);
          const ctaCredito = mapped ? mapped.ctaCredito : 'NÃO MAPEADA';
          const descricaoMapa = mapped ? mapped.descricao : 'CONTA NÃO MAPEADA';

          const historicoPis = `CRED PIS AQUIS. ${descricaoMapa} - ${cleanComp}`;
          const historicoCofins = `CRED COFINS AQUIS. ${descricaoMapa} - ${cleanComp}`;

          return {
            rowLabel: item.rowLabel,
            ctaLancamento: item.ctaLancamento,
            ctaCredito,
            descricaoMapa,
            somaPis: item.somaPis,
            somaCofins: item.somaCofins,
            historicoPis,
            historicoCofins,
            isMapped: Boolean(mapped),
          };
        }
      );

      totalPis += subtotalPis;
      totalCofins += subtotalCofins;

      return {
        filial,
        subtotalPis,
        subtotalCofins,
        accounts,
      };
    });

    return {
      blockCode,
      sourceTab,
      ctaDebitoPis,
      ctaDebitoCofins,
      filialGroups,
      totalPis,
      totalCofins,
    };
  });
}

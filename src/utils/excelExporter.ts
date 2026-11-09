import ExcelJS from 'exceljs';
import {
  ConsolidatedRecord,
  EFD_TABS,
  ResumoComparisonBlock,
  ResumoTotalComparisonRow,
  TabAggregation,
} from '../types/efd';
import {
  AccountingCategoryBlock,
  AccountingMapEntry,
} from './accountingDynamics';
import { norm } from './excelParser';

const AUDIT_COLS = [
  'Base de cálculo (padronizada)',
  'PIS (planilha)',
  'COFINS (planilha)',
  'PIS esperado',
  'COFINS esperado',
  'Diferença PIS',
  'Diferença COFINS',
  'Status do cálculo',
  'Justificativa de auditoria',
  'Origem da seleção',
  'Registro duplicado',
  'Mesmo documento também em',
];

function buildAuditValues(o: ConsolidatedRecord) {
  return [
    o.base,
    o.pis,
    o.cof,
    o.ep,
    o.ec,
    o.dp,
    o.dc,
    o.st,
    o.justificationNote || '',
    o.isManualIncluded
      ? 'Incluída manualmente'
      : `Cor azul (${o.matchedBlueColors.map((c) => '#' + c).join(', ')})`,
    o.dup ? 'Sim' : 'Não',
    o.mx.join(', '),
  ];
}

function populateDinamicaContabilSheet(
  wb: ExcelJS.Workbook,
  accountingBlocks: AccountingCategoryBlock[],
  accountingMap: AccountingMapEntry[]
) {
  const ws = wb.addWorksheet('Dinâmica Contábil');
  ws.views = [{ state: 'frozen', ySplit: 4 }];

  const blackFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF000000' },
  };

  const filialBlueFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFB4C6E7' },
  };

  const rightHeaderFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF1F5F9' },
  };

  const whiteBoldFont: Partial<ExcelJS.Font> = {
    bold: true,
    color: { argb: 'FFFFFFFF' },
    size: 10,
  };

  const darkBoldFont: Partial<ExcelJS.Font> = {
    bold: true,
    color: { argb: 'FF0F172A' },
    size: 10,
  };

  const mutedAccountFont: Partial<ExcelJS.Font> = {
    color: { argb: 'FF475569' },
    size: 9,
  };

  // Render only blocks that have records (or all 7 standard blocks if user prefers full template)
  const blocksToRender = accountingBlocks.filter(
    (b) => b.filialGroups.length > 0
  );
  const activeBlocks =
    blocksToRender.length > 0 ? blocksToRender : accountingBlocks;

  activeBlocks.forEach((block, blkIdx) => {
    // Each block occupies 7 columns + 1 spacer column = 8 columns stride
    const sc = 1 + blkIdx * 8;

    // Set column widths for this block
    ws.getColumn(sc).width = 18; // Rótulos de Linha
    ws.getColumn(sc + 1).width = 14; // Soma de pis
    ws.getColumn(sc + 2).width = 15; // Soma de cofins
    ws.getColumn(sc + 3).width = 15; // Cta Lançamento
    ws.getColumn(sc + 4).width = 15; // Cta Crédito
    ws.getColumn(sc + 5).width = 52; // Histórico PIS
    ws.getColumn(sc + 6).width = 54; // Histórico COFINS
    ws.getColumn(sc + 7).width = 4; // Spacer column

    // Row 1: Category Title (e.g. "02 - Uso e Consumo")
    ws.mergeCells(1, sc, 1, sc + 2);
    const titleCell = ws.getRow(1).getCell(sc);
    titleCell.value = block.blockCode;
    titleCell.font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    for (let c = sc; c <= sc + 2; c++) {
      ws.getRow(1).getCell(c).border = {
        bottom: { style: 'thin', color: { argb: 'FF64748B' } },
      };
    }

    // Row 2: "Cta Débito" above Soma de pis / Soma de cofins
    ws.mergeCells(2, sc + 1, 2, sc + 2);
    const debitoLabelCell = ws.getRow(2).getCell(sc + 1);
    debitoLabelCell.value = 'Cta Débito';
    debitoLabelCell.font = darkBoldFont;
    debitoLabelCell.alignment = { horizontal: 'center' };

    // Row 3: Debit Account Numbers (11070100002 | 11070100003)
    const debPisCell = ws.getRow(3).getCell(sc + 1);
    debPisCell.value = block.ctaDebitoPis;
    debPisCell.font = { size: 9, bold: true, color: { argb: 'FF1E293B' } };
    debPisCell.alignment = { horizontal: 'center' };

    const debCofCell = ws.getRow(3).getCell(sc + 2);
    debCofCell.value = block.ctaDebitoCofins;
    debCofCell.font = { size: 9, bold: true, color: { argb: 'FF1E293B' } };
    debCofCell.alignment = { horizontal: 'center' };

    // Row 4: Table Headers
    const r4 = ws.getRow(4);
    const pivotHeaders = ['Rótulos de Linha', 'Soma de pis', 'Soma de cofins'];
    pivotHeaders.forEach((hText, i) => {
      const cell = r4.getCell(sc + i);
      cell.value = hText;
      cell.fill = blackFill;
      cell.font = whiteBoldFont;
      cell.alignment = {
        horizontal: i === 0 ? 'left' : 'right',
        vertical: 'middle',
      };
    });

    const rightHeaders = [
      'Cta Lançamento',
      'Cta Crédito',
      'Histórico PIS',
      'Histórico COFINS',
    ];
    rightHeaders.forEach((hText, i) => {
      const cell = r4.getCell(sc + 3 + i);
      cell.value = hText;
      cell.fill = rightHeaderFill;
      cell.font = darkBoldFont;
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      };
    });

    // Data Rows starting at Row 5
    let currentRow = 5;

    block.filialGroups.forEach((fg) => {
      // Filial Subtotal Row
      const fRow = ws.getRow(currentRow++);
      const fCell0 = fRow.getCell(sc);
      const fCell1 = fRow.getCell(sc + 1);
      const fCell2 = fRow.getCell(sc + 2);

      fCell0.value = fg.filial;
      fCell0.font = darkBoldFont;
      fCell0.fill = filialBlueFill;

      fCell1.value = fg.subtotalPis;
      fCell1.numFmt = '#,##0.00';
      fCell1.font = darkBoldFont;
      fCell1.fill = filialBlueFill;

      fCell2.value = fg.subtotalCofins;
      fCell2.numFmt = '#,##0.00';
      fCell2.font = darkBoldFont;
      fCell2.fill = filialBlueFill;

      // Account Rows inside Filial
      fg.accounts.forEach((acc) => {
        const aRow = ws.getRow(currentRow++);
        const cLabel = aRow.getCell(sc);
        const cPis = aRow.getCell(sc + 1);
        const cCof = aRow.getCell(sc + 2);
        const cLanc = aRow.getCell(sc + 3);
        const cCred = aRow.getCell(sc + 4);
        const cHistPis = aRow.getCell(sc + 5);
        const cHistCof = aRow.getCell(sc + 6);

        cLabel.value = `   ${acc.rowLabel}`;
        cLabel.font = { size: 10, color: { argb: 'FF0F172A' } };

        cPis.value = acc.somaPis;
        cPis.numFmt = '#,##0.00';
        cPis.font = { size: 10, color: { argb: 'FF0F172A' } };

        cCof.value = acc.somaCofins;
        cCof.numFmt = '#,##0.00';
        cCof.font = { size: 10, color: { argb: 'FF0F172A' } };

        cLanc.value = acc.ctaLancamento;
        cLanc.font = mutedAccountFont;

        cCred.value = acc.ctaCredito;
        cCred.font = acc.isMapped
          ? mutedAccountFont
          : { size: 9, bold: true, color: { argb: 'FFB91C1C' } };

        cHistPis.value = acc.historicoPis;
        cHistPis.font = mutedAccountFont;

        cHistCof.value = acc.historicoCofins;
        cHistCof.font = mutedAccountFont;

        // Subtle borders on the pivot columns
        [cLabel, cPis, cCof].forEach((cell) => {
          cell.border = {
            bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
            right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          };
        });
      });
    });

    // Bottom Total Geral Row for this block
    const totRow = ws.getRow(currentRow);
    const tCell0 = totRow.getCell(sc);
    const tCell1 = totRow.getCell(sc + 1);
    const tCell2 = totRow.getCell(sc + 2);

    tCell0.value = 'Total Geral';
    tCell0.fill = blackFill;
    tCell0.font = whiteBoldFont;

    tCell1.value = block.totalPis;
    tCell1.numFmt = '#,##0.00';
    tCell1.fill = blackFill;
    tCell1.font = whiteBoldFont;

    tCell2.value = block.totalCofins;
    tCell2.numFmt = '#,##0.00';
    tCell2.fill = blackFill;
    tCell2.font = whiteBoldFont;
  });

  // Also add the "Mapa Contábil" reference sheet
  const wsMapa = wb.addWorksheet('Mapa Contábil');
  wsMapa.addRow(['Cta Lançamento', 'Cta Crédito', 'Descrição']);
  const hdr = wsMapa.getRow(1);
  hdr.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  hdr.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1F4E79' },
  };
  accountingMap.forEach((m) => {
    wsMapa.addRow([m.ctaLancamento, m.ctaCredito, m.descricao]);
  });
  wsMapa.views = [{ state: 'frozen', ySplit: 1 }];
  wsMapa.getColumn(1).width = 18;
  wsMapa.getColumn(2).width = 18;
  wsMapa.getColumn(3).width = 44;
}

export async function exportConsolidatedExcel(params: {
  sourceFileName: string;
  allRecords: ConsolidatedRecord[];
  usedColumns: Array<[string, string]>;
  tabAggregations: TabAggregation[];
  resumoTotalRows: ResumoTotalComparisonRow[] | null;
  resumoBlocks?: ResumoComparisonBlock[];
  accountingBlocks?: AccountingCategoryBlock[];
  accountingMap?: AccountingMapEntry[];
}): Promise<void> {
  const {
    sourceFileName,
    allRecords,
    usedColumns,
    tabAggregations,
    resumoTotalRows,
    accountingBlocks,
    accountingMap,
  } = params;

  const activeRecords = allRecords.filter((r) => !r.isExcluded);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'EFD Crédito Fiscal — Consolidação PIS e COFINS';
  wb.created = new Date();

  const createStyledSheet = (
    name: string,
    headers: string[],
    rows: Array<Array<string | number | null | undefined>>
  ) => {
    const ws = wb.addWorksheet(name.slice(0, 31));
    ws.addRow(headers);
    rows.forEach((r) => ws.addRow(r));

    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E79' },
    };

    ws.views = [{ state: 'frozen', ySplit: 1 }];
    if (headers.length > 0) {
      ws.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: headers.length },
      };
    }

    headers.forEach((x, i) => {
      const col = ws.getColumn(i + 1);
      col.width = Math.min(40, Math.max(12, String(x).length + 4));
      col.numFmt = '#,##0.00';
    });
    if (ws.columnCount >= 2) {
      ws.getColumn(2).numFmt = '0';
    }
    return ws;
  };

  const makeRecordLine = (cols: Array<[string, string]>, o: ConsolidatedRecord) => [
    o.aba,
    o.linha,
    ...cols.map(([k]) => o.d[k] ?? ''),
    ...buildAuditValues(o),
  ];

  // =========================================================================
  // 1. Aba: Resumo por aba
  // =========================================================================
  const grandTotal = tabAggregations.reduce(
    (s, a) => ({
      n: s.n + a.n,
      manInc: s.manInc + a.manualIncludedCount,
      manExc: s.manExc + a.excludedCount,
      b: s.b + a.base,
      p: s.p + a.pis,
      c: s.c + a.cof,
      v: s.v + a.div,
      j: s.j + a.justified,
      d: s.d + a.dup,
    }),
    { n: 0, manInc: 0, manExc: 0, b: 0, p: 0, c: 0, v: 0, j: 0, d: 0 }
  );

  const wsResumo = createStyledSheet(
    'Resumo por aba',
    [
      'Aba',
      'Registros ativos',
      'Base (R$)',
      'PIS (R$)',
      'COFINS (R$)',
      'Total de créditos (R$)',
      'Divergências pendentes',
      'Divergências justificadas',
      'Duplicados',
      'Incluídos manualmente',
      'Excluídos manualmente',
    ],
    [
      ...tabAggregations.map((a) => [
        a.aba,
        a.n,
        a.base,
        a.pis,
        a.cof,
        a.tot,
        a.div,
        a.justified,
        a.dup,
        a.manualIncludedCount,
        a.excludedCount,
      ]),
      [
        'Total',
        grandTotal.n,
        grandTotal.b,
        grandTotal.p,
        grandTotal.c,
        grandTotal.p + grandTotal.c,
        grandTotal.v,
        grandTotal.j,
        grandTotal.d,
        grandTotal.manInc,
        grandTotal.manExc,
      ],
    ]
  );

  wsResumo.getColumn(1).width = 22;
  [2, 7, 8, 9, 10, 11].forEach((cIdx) => {
    wsResumo.getColumn(cIdx).numFmt = '0';
  });
  wsResumo.getRow(wsResumo.rowCount).font = { bold: true };

  if (resumoTotalRows && resumoTotalRows.length > 0) {
    wsResumo.addRow([]);
    const subHdr = wsResumo.addRow([
      'Comparação com a aba Resumo',
      'Linha Total (Resumo)',
      'Registros azuis (Resumo)',
      'Diferença',
      'Status',
    ]);
    subHdr.font = { bold: true };
    resumoTotalRows.forEach((row) => {
      const added = wsResumo.addRow([
        row.item,
        row.resumoTotalLine,
        row.resumoBlueRows,
        row.diff,
        row.isOk ? 'OK' : 'Divergente',
      ]);
      added.getCell(2).numFmt = '#,##0.00';
      added.getCell(3).numFmt = '#,##0.00';
      added.getCell(4).numFmt = '#,##0.00';
    });
  }

  // =========================================================================
  // 2. Aba: Dinâmica Contábil (Side-by-Side Pivot + De-Para Cta Crédito & Históricos)
  // =========================================================================
  if (accountingBlocks && accountingMap) {
    populateDinamicaContabilSheet(wb, accountingBlocks, accountingMap);
  }

  // =========================================================================
  // 3. Aba: Consolidado
  // =========================================================================
  createStyledSheet(
    'Consolidado',
    ['Aba de origem', 'Linha no Excel', ...usedColumns.map((c) => c[1]), ...AUDIT_COLS],
    activeRecords.map((o) => makeRecordLine(usedColumns, o))
  );

  // =========================================================================
  // 4. Aba: Divergências (Auditoria)
  // =========================================================================
  const divRecords = activeRecords.filter(
    (o) => o.rawStatus === 'Divergente' || o.st === 'Divergente' || o.st === 'Justificado'
  );
  const wsDiv = createStyledSheet(
    'Divergências',
    [
      'Aba de origem',
      'Linha no Excel',
      'Status atual',
      'Justificativa / Revisão',
      'Documento',
      'Fornecedor / Descrição',
      'Base de cálculo (R$)',
      'PIS planilha (R$)',
      'PIS esperado (R$)',
      'Diferença PIS (R$)',
      'COFINS planilha (R$)',
      'COFINS esperado (R$)',
      'Diferença COFINS (R$)',
    ],
    divRecords.map((o) => {
      const doc =
        o.d[norm('Documento')] ||
        o.d[norm('Numero Doc')] ||
        o.d[norm('Conta')] ||
        '—';
      const desc =
        o.d[norm('Razão Social')] ||
        o.d[norm('Forn/Cliente')] ||
        o.d[norm('Descrição')] ||
        o.d[norm('Descrição do Item')] ||
        '—';
      return [
        o.aba,
        o.linha,
        o.st,
        o.justificationNote || 'Pendente de justificativa',
        doc,
        desc,
        o.base,
        o.pis,
        o.ep,
        o.dp,
        o.cof,
        o.ec,
        o.dc,
      ];
    })
  );
  wsDiv.getColumn(2).numFmt = '0';

  // =========================================================================
  // 5. Aba: Duplicados (Auditoria)
  // =========================================================================
  const dupRecords = activeRecords.filter((o) => o.dup || o.mx.length > 0);
  const wsDup = createStyledSheet(
    'Duplicados',
    [
      'Aba de origem',
      'Linha no Excel',
      'Duplicado na aba',
      'Mesmo documento em outra aba',
      'Chave Documento (Doc|Série|Forn|Loja)',
      'Chave Item NF',
      'Base de cálculo (R$)',
      'PIS (R$)',
      'COFINS (R$)',
      'Status cálculo',
    ],
    dupRecords.map((o) => [
      o.aba,
      o.linha,
      o.dup ? 'Sim' : 'Não',
      o.mx.join(', ') || '—',
      o.dk || '—',
      o.ik || '—',
      o.base,
      o.pis,
      o.cof,
      o.st,
    ])
  );
  wsDup.getColumn(2).numFmt = '0';

  // =========================================================================
  // 6..N. Abas individuais por categoria
  // =========================================================================
  for (const t of EFD_TABS) {
    const rs = activeRecords.filter((o) => o.aba === t);
    if (!rs.length) continue;
    const cols = usedColumns.filter(([k]) => rs.some((o) => k in o.d));
    createStyledSheet(
      t,
      ['Aba de origem', 'Linha no Excel', ...cols.map((c) => c[1]), ...AUDIT_COLS],
      rs.map((o) => makeRecordLine(cols, o))
    );
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const cleanName = sourceFileName.replace(/\.xlsx$/i, '');
  a.href = url;
  a.download = `Consolidado_PIS_COFINS_${cleanName}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

import ExcelJS from 'exceljs';

/**
 * Generates an authentic EFD 8_2026 (.xlsx) workbook in memory matching the exact
 * row/column coordinates of EFD 8_2026 and including the 'Conta Contábil' (Cta Lançamento)
 * and 'Filial' (0101, 0102, 0103) codes from the user's accounting map.
 */
export async function generateDemoEfdWorkbookBuffer(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'EFD Contribuições 8_2026';
  wb.created = new Date(2026, 7, 31);

  const blueFillPrimary: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFBDD7EE' }, // #BDD7EE - Azul claro padrão de seleção fiscal
  };

  const blueFillSecondary: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF8EA9DB' }, // #8EA9DB - Azul médio secundário
  };

  const grayHeaderFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9D9D9' },
  };

  const yellowNoteFill: ExcelJS.Fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFF2CC' }, // Linha amarela (não azul) para demonstrar exclusão automática / inclusão manual
  };

  const styleRange = (
    row: ExcelJS.Row,
    c1: number,
    c2: number,
    fill: ExcelJS.Fill | null,
    bold = false
  ) => {
    for (let c = c1; c <= c2; c++) {
      const cell = row.getCell(c);
      if (fill) cell.fill = fill;
      if (bold) cell.font = { bold: true };
    }
  };

  const stdHeaders = [
    'Filial',
    'Documento',
    'Serie',
    'Forn/Cliente',
    'Loja',
    'Razão Social',
    'Item NF',
    'Produto',
    'Descrição do Item',
    'CFOP',
    'Conta Contábil',
    'Base Pis/Cofins',
    'Pis',
    'Cofins',
  ];

  // =========================================================================
  // 1. Sucata 1 (XX - Sucatas)
  // =========================================================================
  const wsSucata = wb.addWorksheet('Sucata 1');
  wsSucata.getRow(1).values = ['EFD CONTRIBUIÇÕES - COMPETÊNCIA 08/2026 - AQUISIÇÃO DE SUCATA'];
  wsSucata.getRow(3).values = stdHeaders;
  styleRange(wsSucata.getRow(3), 1, stdHeaders.length, grayHeaderFill, true);

  const sucataData = [
    {
      vals: ['0101', '00045812', '1', 'FORN001', '01', 'METAIS RECICLAVEIS SUL LTDA', '001', 'SUC-01', 'SUCATA DE FERRO MISTA INDUSTRIAL', '1101', '11030100002', 42500.0, 701.25, 3230.0],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', '00045890', '1', 'FORN002', '01', 'COMERCIO DE APARAS ALFA SA', '001', 'SUC-02', 'SUCATA DE ACO INOX 304', '1101', '11030100004', 18400.0, 303.6, 1398.4],
      fill: blueFillPrimary,
    },
    {
      vals: ['0102', '00045901', '1', 'FORN005', '01', 'SUCATAS NORTE LTDA', '001', 'SUC-04', 'SUCATA DE ESTAMPARIA PESADA', '1101', '11030100003', 15000.0, 247.5, 1140.0],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', '00045915', '1', 'FORN003', '01', 'RECICLAGEM VALE MEI (SIMPLES)', '001', 'SUC-03', 'APARAS SEM DIREITO A CREDITO', '1101', '11030100002', 6200.0, 0, 0],
      fill: yellowNoteFill,
    },
  ];
  sucataData.forEach((item, idx) => {
    const r = wsSucata.getRow(4 + idx);
    r.values = item.vals;
    styleRange(r, 1, stdHeaders.length, item.fill);
  });

  // =========================================================================
  // 2. Uso e consumo (02 - Uso e Consumo)
  // =========================================================================
  const wsUso = wb.addWorksheet('Uso e consumo');
  wsUso.getRow(1).values = ['EFD CONTRIBUIÇÕES - COMPETÊNCIA 08/2026 - INSUMOS E USO/CONSUMO'];
  wsUso.getRow(3).values = stdHeaders;
  styleRange(wsUso.getRow(3), 1, stdHeaders.length, grayHeaderFill, true);

  const usoData = [
    {
      vals: ['0101', '00192011', '2', 'FORN010', '01', 'SUPRIMENTOS INDUSTRIAIS BRASIL', '001', 'EPI-104', 'LUVA DE PROTECAO TERMICA FORNO', '1556', '11040100001', 12800.0, 211.2, 972.8],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', '00192020', '2', 'FORN012', '01', 'LUBRIFICANTES BRASIL SA', '001', 'LUB-01', 'OLEO HIDRAULICO INDUSTRIAL 68', '1556', '11040400001', 24000.0, 396.0, 1824.0],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', '00192032', '1', 'FORN015', '01', 'PECAS E ROLAMENTOS SUL LTDA', '001', 'PEC-88', 'ROLAMENTO DE ESFERAS EIXO PRINCIPAL', '1556', '11040500001', 19500.0, 321.75, 1482.0],
      fill: blueFillPrimary,
    },
    {
      // Divergência intencional de cálculo: Base 8500 -> PIS esperado 140,25 (informado 125,00)
      vals: ['0101', '00192045', '2', 'FORN011', '01', 'QUIMICA INDUSTRIAL DELTA LTDA', '001', 'INS-209', 'DESENGRAXANTE LINHA DE USINAGEM', '1556', '11041200001', 8500.0, 125.0, 646.0],
      fill: blueFillSecondary,
    },
    {
      // Registro duplicado (mesmo Documento + Série + Forn + Loja + Item NF)
      vals: ['0101', '00192011', '2', 'FORN010', '01', 'SUPRIMENTOS INDUSTRIAIS BRASIL', '001', 'EPI-104', 'LUVA DE PROTECAO TERMICA FORNO', '1556', '11040100001', 12800.0, 211.2, 972.8],
      fill: blueFillPrimary,
    },
    {
      // Linha não-azul (material de expediente administrativo)
      vals: ['0101', '00192199', '1', 'FORN019', '01', 'PAPELARIA CENTRAL LTDA', '001', 'ADM-012', 'MATERIAL DE ESCRITORIO ADMINISTRATIVO', '1556', '11041700001', 1450.0, 0, 0],
      fill: null,
    },
  ];
  usoData.forEach((item, idx) => {
    const r = wsUso.getRow(4 + idx);
    r.values = item.vals;
    styleRange(r, 1, stdHeaders.length, item.fill);
  });

  // =========================================================================
  // 3. Nfse (03 - NFSe)
  // =========================================================================
  const wsNfse = wb.addWorksheet('Nfse');
  wsNfse.getRow(1).values = ['EFD CONTRIBUIÇÕES - COMPETÊNCIA 08/2026 - SERVIÇOS TOMADOS (NFSE)'];
  wsNfse.getRow(3).values = stdHeaders;
  styleRange(wsNfse.getRow(3), 1, stdHeaders.length, grayHeaderFill, true);

  const nfseData = [
    {
      vals: ['0101', '20260844', 'RPS', 'SERV004', '01', 'MANUTENCAO DE CALDEIRAS TECH LTDA', '001', 'SRV-MNT', 'MANUTENCAO DE EQUIPAMENTOS USINAGEM', '1933', '42040400001', 34000.0, 561.0, 2584.0],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', '20260891', 'RPS', 'SERV008', '01', 'OFICINA PESADA DIESEL LTDA', '001', 'SRV-VEI', 'MANUTENCAO DE VEICULOS FROTA', '1933', '42040400002', 9600.0, 158.4, 729.6],
      fill: blueFillPrimary,
    },
    {
      vals: ['0102', '20260912', 'RPS', 'SERV012', '01', 'BALANCAS PRECISAO INDUSTRIAL', '001', 'SRV-BAL', 'MANUT PREVENTIVA BALANCA RODOVIARIA', '1933', '42010200025', 4000.0, 66.0, 304.0],
      fill: blueFillPrimary,
    },
  ];
  nfseData.forEach((item, idx) => {
    const r = wsNfse.getRow(4 + idx);
    r.values = item.vals;
    styleRange(r, 1, stdHeaders.length, item.fill);
  });

  // =========================================================================
  // 4. 04 e 05 Outros (XX - Outros / Diesel)
  // Seção 1: h=3, a=4, b=36, c1=C (col 3), c2=J (col 10), base='B.C', pis='Pis', cof='Cofins' (1.65% / 7.6%)
  // Seção 2: h='', a=39, b=70, c1=C (col 3), c2=G (col 7), man='Conta;Descrição;Valor;Pis;Cofins', base='Valor' (0.65% / 4.0%)
  // =========================================================================
  const wsOutros = wb.addWorksheet('04 e 05 Outros');
  wsOutros.getRow(1).values = ['', '', 'APURAÇÃO BLOCO 04 E 05 - OUTROS CRÉDITOS E COMBUSTÍVEIS (08/2026)'];

  const outrosSec1Headers = [
    '',
    '',
    'Filial',
    'Documento',
    'Serie',
    'Conta',
    'Descrição',
    'B.C',
    'Pis',
    'Cofins',
  ];
  wsOutros.getRow(3).values = outrosSec1Headers;
  styleRange(wsOutros.getRow(3), 3, 10, grayHeaderFill, true);

  const rOut4 = wsOutros.getRow(4);
  rOut4.values = ['', '', '0101', '88410', '1', '11041600001', 'DIESEL S10 FROTA INDUSTRIAL', 68000.0, 1122.0, 5168.0];
  styleRange(rOut4, 3, 10, blueFillPrimary);

  const rOut5 = wsOutros.getRow(5);
  rOut5.values = ['', '', '0101', '88411', '1', '11040600001', 'GAS COMBUSTIVEL FORNO DE FUSAO', 24500.0, 404.25, 1862.0];
  styleRange(rOut5, 3, 10, blueFillPrimary);

  const rOut6 = wsOutros.getRow(6);
  rOut6.values = ['', '', '0101', '88499', '1', '42040500001', 'TELEFONIA CORPORATIVA (NAO CREDITA)', 3200.0, 0, 0];
  styleRange(rOut6, 3, 10, null);

  // Seção 2 (linhas 39..70, colunas C a G: Conta;Descrição;Valor;Pis;Cofins, alíquotas 0.65% e 4%)
  wsOutros.getRow(38).values = ['', '', 'BLOCO SECUNDÁRIO - ALÍQUOTAS DIFERENCIADAS (0,65% e 4,00%)'];
  const rOut39 = wsOutros.getRow(39);
  rOut39.values = ['', '', '11041600001', 'OLEO DIESEL B S10 GERADOR', 15000.0, 97.5, 600.0];
  styleRange(rOut39, 3, 7, blueFillSecondary);

  const rOut40 = wsOutros.getRow(40);
  rOut40.values = ['', '', '11041800001', 'PERFIS E CHAPAS METAIS ESPECIAIS', 8000.0, 52.0, 320.0];
  styleRange(rOut40, 3, 7, blueFillSecondary);

  // =========================================================================
  // 5. Locação (06 - Locação)
  // =========================================================================
  const wsLoc = wb.addWorksheet('Locação');
  wsLoc.getRow(1).values = ['EFD CONTRIBUIÇÕES - COMPETÊNCIA 08/2026 - LOCAÇÃO DE PRÉDIOS E MÁQUINAS'];
  wsLoc.getRow(3).values = stdHeaders;
  styleRange(wsLoc.getRow(3), 1, stdHeaders.length, grayHeaderFill, true);

  const locData = [
    {
      vals: ['0101', 'LOC0826', 'U', 'LOC001', '01', 'LOGISTICA FROTA PESADA SA', '001', 'LOC-VEI', 'ALUGUEL DE VEICULOS OPERACIONAIS', '1949', '42040600004', 45000.0, 742.5, 3420.0],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', 'LOC0827', 'U', 'LOC002', '01', 'RENTAL EMPILHADEIRAS PESADAS LTDA', '001', 'LOC-EMP', 'ALUGUEIS DE EQUIPAMENTOS INDUSTRIAIS', '1949', '42040600001', 19200.0, 316.8, 1459.2],
      fill: blueFillPrimary,
    },
  ];
  locData.forEach((item, idx) => {
    const r = wsLoc.getRow(4 + idx);
    r.values = item.vals;
    styleRange(r, 1, stdHeaders.length, item.fill);
  });

  // =========================================================================
  // 6. Imobilizado (10 - Imobilizado)
  // =========================================================================
  const wsImob = wb.addWorksheet('Imobilizado');
  wsImob.getRow(1).values = ['EFD CONTRIBUIÇÕES - COMPETÊNCIA 08/2026 - CRÉDITO SOBRE ATIVO IMOBILIZADO'];
  wsImob.getRow(3).values = stdHeaders;
  styleRange(wsImob.getRow(3), 1, stdHeaders.length, grayHeaderFill, true);

  const imobData = [
    {
      vals: ['0101', '00912044', '1', 'MAQ001', '01', 'ROMI MAQUINAS OPERATRIZES SA', '001', 'CNC-500', 'CENTRO DE USINAGEM VERTICAL (PARCELA 12/48)', '1551', '12050100001', 28000.0, 462.0, 2128.0],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', '00912088', '1', 'MAQ002', '01', 'WEG EQUIPAMENTOS ELETRICOS SA', '001', 'SUB-200', 'TRANSFORMADOR DE FORCA INDUSTRIAL (06/48)', '1551', '12050100007', 16000.0, 264.0, 1216.0],
      fill: blueFillPrimary,
    },
  ];
  imobData.forEach((item, idx) => {
    const r = wsImob.getRow(4 + idx);
    r.values = item.vals;
    styleRange(r, 1, stdHeaders.length, item.fill);
  });

  // =========================================================================
  // 7. Frete (14 - Frete)
  // =========================================================================
  const wsFrete = wb.addWorksheet('Frete');
  wsFrete.getRow(1).values = ['EFD CONTRIBUIÇÕES - COMPETÊNCIA 08/2026 - FRETES SOBRE COMPRAS E VENDAS'];
  wsFrete.getRow(3).values = stdHeaders;
  styleRange(wsFrete.getRow(3), 1, stdHeaders.length, grayHeaderFill, true);

  const freteData = [
    {
      vals: ['0101', '308812', '1', 'TRAN01', '01', 'RODONAVES TRANSPORTES E ENCOMENDAS', '001', 'FRT-CMP', 'FRETE SOBRE COMPRAS DE INSUMOS', '1352', '42040300002', 21400.0, 353.1, 1626.4],
      fill: blueFillPrimary,
    },
    {
      vals: ['0101', '308899', '1', 'TRAN02', '01', 'EXPRESSO MERCURIO CARGAS LTDA', '001', 'FRT-VND', 'FRETE OPERACAO DE VENDA CIF', '2352', '42040300001', 14600.0, 240.9, 1109.6],
      fill: blueFillPrimary,
    },
  ];
  freteData.forEach((item, idx) => {
    const r = wsFrete.getRow(4 + idx);
    r.values = item.vals;
    styleRange(r, 1, stdHeaders.length, item.fill);
  });

  // =========================================================================
  // 8. Resumo
  // =========================================================================
  const wsResumo = wb.addWorksheet('Resumo');
  wsResumo.getRow(1).values = ['QUADRO RESUMO DE APURAÇÃO MENSAL DE CRÉDITOS - 08/2026'];
  wsResumo.getRow(2).values = [
    'Insumos, Serviços e Operações Internas',
    '',
    '',
    'Imobilizado, Locações e Fretes',
    '',
    '',
  ];
  styleRange(wsResumo.getRow(2), 1, 6, grayHeaderFill, true);

  wsResumo.getRow(3).values = [
    'Soma de Vl Mês',
    'Soma de PIS',
    'Soma de COFINS',
    'Soma de Vl Mês',
    'Soma de PIS',
    'Soma de COFINS',
  ];
  styleRange(wsResumo.getRow(3), 1, 6, grayHeaderFill, true);

  const res4 = wsResumo.getRow(4);
  res4.values = [199500.0, 3291.75, 15162.0, 64200.0, 1059.3, 4879.2];
  styleRange(res4, 1, 6, blueFillPrimary);

  const res5 = wsResumo.getRow(5);
  res5.values = [115500.0, 1905.75, 8778.0, 80000.0, 1320.0, 6080.0];
  styleRange(res5, 1, 6, blueFillPrimary);

  // Total reconciliation block on rows 8 and 9
  const res8 = wsResumo.getRow(8);
  res8.values = ['Total', 'PIS', 'COFINS'];
  styleRange(res8, 1, 3, grayHeaderFill, true);

  const res9 = wsResumo.getRow(9);
  res9.values = [459200.0, 7576.8, 34899.2];
  styleRange(res9, 1, 3, null, true);

  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}

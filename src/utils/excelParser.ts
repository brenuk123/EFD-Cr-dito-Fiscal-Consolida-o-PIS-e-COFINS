import JSZip from 'jszip';
import {
  ConsolidatedRecord,
  EFD_TABS,
  ManualAuditState,
  ParsedCell,
  ParsedRow,
  ParsedSheet,
  ParsedWorkbook,
  RawSheetCandidateRow,
  ResumoComparisonBlock,
  ResumoTotalComparisonRow,
  TabAggregation,
  TabConfigMap,
  TabImportDiag,
  TabSectionConfig,
} from '../types/efd';

const RX: Record<string, RegExp> = {};
const attr = (rawAttrs: string, key: string): string | null => {
  const r = RX[key] || (RX[key] = new RegExp('(?:^|\\s)' + key + '="([^"]*)"'));
  const m = r.exec(rawAttrs);
  return m ? m[1] : null;
};

const unescapeXml = (s: string): string =>
  s.indexOf('&') < 0
    ? s
    : s.replace(/&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g, (_m, e) => {
        const map: Record<string, string> = {
          amp: '&',
          lt: '<',
          gt: '>',
          quot: '"',
          apos: "'",
        };
        if (map[e]) return map[e];
        const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : +e.slice(1);
        return String.fromCodePoint(code);
      });

export function norm(s: unknown): string {
  return String(s ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function c2n(col: string): number {
  return [...(col || '').toUpperCase().replace(/[^A-Z]/g, '')].reduce(
    (n, ch) => n * 26 + ch.charCodeAt(0) - 64,
    0
  );
}

export function n2c(num: number): string {
  let n = num;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - 1 - r) / 26;
  }
  return s;
}

export function parseNum(x: unknown): number {
  if (typeof x === 'number') return Number.isFinite(x) ? x : 0;
  if (x == null || x === '') return 0;
  const raw = String(x).trim();
  const normalized = raw.includes(',')
    ? raw.replace(/\./g, '').replace(',', '.')
    : raw;
  const val = +normalized;
  return Number.isNaN(val) ? 0 : val;
}

export function formatCurrency(n: number | null | undefined): string {
  const val = typeof n === 'number' && Number.isFinite(n) ? n : 0;
  return val.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export function applyTint(hexInput: string, t: number): string {
  const hex = hexInput.replace('#', '').slice(-6);
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const mx = Math.max(r, g, b);
  const mn = Math.min(r, g, b);
  const d = mx - mn;
  let h = 0;
  let s = 0;
  let l = (mx + mn) / 2;
  if (d) {
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h =
      mx === r
        ? ((g - b) / d + 6) % 6
        : mx === g
        ? (b - r) / d + 2
        : (r - g) / d + 4;
    h /= 6;
  }
  l = t < 0 ? l * (1 + t) : l * (1 - t) + t;
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (xIn: number) => {
    const x = ((xIn % 1) + 1) % 1;
    const val =
      x < 1 / 6
        ? p + (q - p) * 6 * x
        : x < 0.5
        ? q
        : x < 2 / 3
        ? p + (q - p) * (2 / 3 - x) * 6
        : p;
    return Math.round(val * 255)
      .toString(16)
      .padStart(2, '0');
  };
  return (f(h + 1 / 3) + f(h) + f(h - 1 / 3)).toUpperCase();
}

export function isBlueHex(hexInput: string | null | undefined): boolean {
  if (!hexInput) return false;
  const hStr = hexInput.replace('#', '').slice(-6);
  if (hStr.length < 6) return false;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hStr.slice(i, i + 2), 16) / 255);
  const mx = Math.max(r, g, b);
  const d = mx - Math.min(r, g, b);
  if (!d || d / mx < 0.06) return false;
  const H =
    60 *
    (mx === r
      ? ((g - b) / d + 6) % 6
      : mx === g
      ? (b - r) / d + 2
      : (r - g) / d + 4);
  return H >= 190 && H <= 260;
}

export async function parseExcelWorkbook(
  file: File | { name: string; size: number; arrayBuffer: () => Promise<ArrayBuffer> }
): Promise<ParsedWorkbook> {
  const buffer = await file.arrayBuffer();
  const Z = await JSZip.loadAsync(buffer);
  const rd = async (p: string) => (Z.file(p) ? Z.file(p)!.async('string') : '');

  const wbx = await rd('xl/workbook.xml');
  const rel = await rd('xl/_rels/workbook.xml.rels');
  if (!wbx) {
    throw new Error('Arquivo Excel (.xlsx) inválido: xl/workbook.xml não encontrado.');
  }

  const rm: Record<string, string> = {};
  for (const m of rel.matchAll(/<Relationship ([^>]*?)\/>/g)) {
    const id = attr(m[1], 'Id');
    const target = attr(m[1], 'Target');
    if (id && target) rm[id] = target;
  }

  const names = [...wbx.matchAll(/<sheet ([^>]*?)\/>/g)].map((m) => {
    const rId = attr(m[1], 'r:id') || attr(m[1], 'id') || '';
    return {
      n: unescapeXml(attr(m[1], 'name') || ''),
      p: 'xl/' + (rm[rId] || '').replace(/^\/?(xl\/)?/, ''),
    };
  });

  const thm = await rd('xl/theme/theme1.xml');
  const sty = await rd('xl/styles.xml');

  const TH = [
    ...((thm.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/) || [''])[0]).matchAll(
      /<a:(?:dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>([\s\S]*?)<\/a:\w+>/g
    ),
  ].map((m) =>
    ((m[1].match(/(?:lastClr|val)="([0-9A-Fa-f]{6})"/) || [, '000000'])[1]).toUpperCase()
  );

  const fl = [
    ...((sty.match(/<fills[\s\S]*?<\/fills>/) || [''])[0]).matchAll(
      /<fill>([\s\S]*?)<\/fill>/g
    ),
  ].map((m) => {
    if (!/patternType="solid"/.test(m[1])) return null;
    const c = (/<fgColor ([^>]*?)\/>/.exec(m[1]) || [])[1];
    if (!c) return null;
    const rgb = attr(c, 'rgb');
    const th = attr(c, 'theme');
    if (rgb) return rgb.slice(-6).toUpperCase();
    if (th != null) {
      const thIdx = [1, 0, 3, 2][+th] ?? +th;
      return applyTint(TH[thIdx] || 'FFFFFF', parseFloat(attr(c, 'tint') || '0'));
    }
    return null;
  });

  const XC = (
    ((sty.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [''])[0]).match(
      /<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g
    ) || []
  )
    .map((x) => {
      const fid = attr(x, 'fillId');
      return fid != null ? fl[+fid] || null : null;
    })
    .map((c) => (c === 'FFFFFF' ? null : c));

  const ss = await rd('xl/sharedStrings.xml');
  const SST = [...ss.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) =>
    unescapeXml(
      [...m[1].replace(/<rPh[\s\S]*?<\/rPh>/g, '').matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)]
        .map((x) => x[1])
        .join('')
    )
  );

  const colorCounts: Record<string, number> = {};
  const sheets: Record<string, ParsedSheet> = {};
  const importLogs: TabImportDiag[] = [];

  const parseSheetFast = (sheetName: string, x: string): ParsedSheet => {
    const rows = new Map<number, ParsedRow>();
    let sheetBlueCount = 0;

    for (const m of x.matchAll(/<row ([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const r = +(attr(m[1], 'r') || 0);
      if (!r) continue;
      const hidden = /hidden="(?:1|true)"/.test(m[1]);
      const cells = new Map<number, ParsedCell>();
      let c = 0;

      if (m[2]) {
        for (const k of m[2].matchAll(/<c ([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
          const a = k[1];
          const rr = attr(a, 'r');
          c = rr ? c2n(rr.replace(/\d+/g, '')) : c + 1;
          const s = attr(a, 's');
          const col = s != null ? XC[+s] || null : null;
          if (col) {
            colorCounts[col] = (colorCounts[col] || 0) + 1;
            if (isBlueHex(col)) sheetBlueCount++;
          }

          let val: string | number = '';
          if (k[2] && c <= 150) {
            const t = attr(a, 't');
            const vv = /<v>([\s\S]*?)<\/v>/.exec(k[2]);
            const isNode = /<is>([\s\S]*?)<\/is>/.exec(k[2]);
            if (t === 'inlineStr' && isNode) {
              val = unescapeXml(
                [...isNode[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)]
                  .map((x) => x[1])
                  .join('')
              );
            } else if (vv) {
              val =
                t === 's'
                  ? SST[+vv[1]] ?? ''
                  : t === 'str' || t === 'e' || t === 'b'
                  ? unescapeXml(vv[1])
                  : +vv[1];
            }
          }
          if (col || (val !== '' && val !== undefined)) {
            cells.set(c, { v: val, bg: col });
          }
        }
      }
      rows.set(r, { hidden, cells });
    }

    return { name: sheetName, rows };
  };

  for (const t of EFD_TABS) {
    const s = names.find((x) => norm(x.n) === norm(t));
    if (!s || !Z.file(s.p)) {
      importLogs.push({
        tabName: t,
        found: false,
        rowCount: 0,
        blueCellCount: 0,
      });
      continue;
    }
    const xmlStr = await Z.file(s.p)!.async('string');
    const parsed = parseSheetFast(t, xmlStr);
    sheets[t] = parsed;

    let blueCellsInSheet = 0;
    parsed.rows.forEach((row) => {
      row.cells.forEach((cell) => {
        if (cell.bg && isBlueHex(cell.bg)) blueCellsInSheet++;
      });
    });

    importLogs.push({
      tabName: t,
      found: true,
      rowCount: parsed.rows.size,
      blueCellCount: blueCellsInSheet,
    });
  }

  const defaultBlueColors = Object.keys(colorCounts).filter(isBlueHex);

  return {
    fileName: file.name,
    fileSize: file.size,
    loadedAt: new Date().toISOString(),
    sheets,
    allSheetNames: names.map((x) => x.n),
    colorCounts,
    defaultBlueColors,
    importLogs,
  };
}

/**
 * Helper to resolve headers and column bounds for a section.
 * Automatically falls back if header row `sec.h` is empty (e.g. checks row 1, 2, 3)
 * so both strict EFD 8_2026 files and custom spreadsheets work seamlessly.
 */
function resolveSectionHeaders(
  sheet: ParsedSheet,
  sec: TabSectionConfig
): {
  c1: number;
  c2: number;
  names: string[];
  nk: string[];
  effectiveFirstDataRow: number;
} | null {
  const c1 = c2n(sec.c1 || 'A');
  let c2 = sec.c2 ? c2n(sec.c2) : 0;
  let names: string[] = [];
  let effectiveFirstDataRow = +sec.a || 4;

  if (sec.man && sec.man.trim()) {
    names = sec.man.split(';').map((s) => s.trim());
    if (!c2) c2 = c1 + names.length - 1;
  } else {
    let headerRowNum = +sec.h;
    let hr = headerRowNum ? sheet.rows.get(headerRowNum) : undefined;

    // Check if hr actually has non-empty values in >= c1
    const hasHeaderValues = (row?: ParsedRow) => {
      if (!row) return false;
      for (const [colIdx, cell] of row.cells.entries()) {
        if (colIdx >= c1 && String(cell.v ?? '').trim() !== '') return true;
      }
      return false;
    };

    if (!hasHeaderValues(hr)) {
      // Fallback: check rows 3, 2, 1
      for (const candidateH of [3, 2, 1]) {
        const candRow = sheet.rows.get(candidateH);
        if (hasHeaderValues(candRow)) {
          hr = candRow;
          headerRowNum = candidateH;
          if (+sec.a <= headerRowNum) {
            effectiveFirstDataRow = headerRowNum + 1;
          }
          break;
        }
      }
    }

    if (!hr) return null;

    if (!c2) {
      const nonEmptyCols = [...hr.cells.entries()]
        .filter(([k, cell]) => k >= c1 && String(cell.v ?? '').trim() !== '')
        .map(([k]) => k);
      c2 = Math.max(c1, ...nonEmptyCols);
    }

    const seen: Record<string, number> = {};
    for (let c = c1; c <= c2; c++) {
      const rawVal = hr.cells.get(c)?.v;
      const n = String(rawVal ?? '').trim() || 'Coluna ' + n2c(c);
      seen[n] = (seen[n] || 0) + 1;
      names.push(seen[n] > 1 ? `${n} (${seen[n]})` : n);
    }
  }

  return {
    c1,
    c2,
    names,
    nk: names.map(norm),
    effectiveFirstDataRow,
  };
}

function findColumnIndex(nk: string[], targetField: string, fallbacks: string[]): number {
  const candidates = [targetField, ...fallbacks].filter(Boolean).map(norm);
  for (const cand of candidates) {
    const exact = nk.indexOf(cand);
    if (exact >= 0) return exact;
  }
  for (const cand of candidates) {
    const partial = nk.findIndex((colNorm) => colNorm.includes(cand) || cand.includes(colNorm));
    if (partial >= 0) return partial;
  }
  return -1;
}

export function buildConsolidation(
  workbook: ParsedWorkbook | null,
  tabConfig: TabConfigMap,
  activeBlueColors: Set<string>,
  tolerance: number,
  ignoreHidden: boolean,
  manualState: ManualAuditState
): {
  records: ConsolidatedRecord[];
  usedColumns: Array<[string, string]>;
} {
  if (!workbook) return { records: [], usedColumns: [] };

  const includedSet = new Set(manualState.includedRowKeys);
  const excludedSet = new Set(manualState.excludedRowKeys);
  const cm = new Map<string, string>();
  const out: ConsolidatedRecord[] = [];

  for (const tab of EFD_TABS) {
    const S = workbook.sheets[tab];
    if (!S) continue;

    const sections = tabConfig[tab] || [];
    sections.forEach((sec, secIdx) => {
      const resolved = resolveSectionHeaders(S, sec);
      if (!resolved) return;

      const { c1, c2, names, nk, effectiveFirstDataRow } = resolved;

      const ib = findColumnIndex(nk, sec.base, [
        'Base Pis/Cofins',
        'Base Calculo',
        'Vlr Contabil',
        'B.C',
        'Valor',
        'Soma de Vl Mês',
        'Base',
      ]);
      const ip = findColumnIndex(nk, sec.pis, [
        'Pis',
        'PIS',
        'Vlr PIS',
        'Soma de PIS',
      ]);
      const ic = findColumnIndex(nk, sec.cof, [
        'Cofins',
        'COFINS',
        'Vlr COFINS',
        'Soma de COFINS',
      ]);

      const maxRowKey = S.rows.size > 0 ? Math.max(...S.rows.keys()) : effectiveFirstDataRow;
      const last = +sec.b || maxRowKey;
      const rp = parseNum(sec.rp);
      const rc = parseNum(sec.rc);
      const startRow = Math.min(+sec.a || effectiveFirstDataRow, effectiveFirstDataRow);

      for (let r = +sec.a || startRow; r <= last; r++) {
        const row = S.rows.get(r);
        if (!row) continue;
        if (ignoreHidden && row.hidden) continue;

        const rowKey = `${tab}|${r}|${secIdx}|${c1}`;
        const matchedBlueColors: string[] = [];
        let hasNonEmptyValue = false;

        for (let c = c1; c <= c2; c++) {
          const cell = row.cells.get(c);
          if (!cell) continue;
          if (cell.v !== '' && cell.v !== undefined) hasNonEmptyValue = true;
          if (cell.bg && activeBlueColors.has(cell.bg)) {
            if (!matchedBlueColors.includes(cell.bg)) {
              matchedBlueColors.push(cell.bg);
            }
          }
        }

        const hasBlueCell = matchedBlueColors.length > 0;
        const isManualIncluded = includedSet.has(rowKey);
        const isExcluded = excludedSet.has(rowKey);

        if ((!hasBlueCell && !isManualIncluded) || !hasNonEmptyValue) continue;

        const d: Record<string, string | number> = {};
        names.forEach((n, i) => {
          const cellVal = row.cells.get(c1 + i)?.v;
          if (cellVal !== undefined && cellVal !== '') {
            d[nk[i]] = cellVal;
          }
          if (!cm.has(nk[i])) {
            cm.set(nk[i], n);
          }
        });

        const L = sec.lab ? S.rows.get(+sec.lab)?.cells.get(c1)?.v ?? '' : '';
        if (L !== '') {
          d['rotulo do bloco'] = L;
          cm.set('rotulo do bloco', 'Rótulo do bloco');
        }

        const nm = (i: number) => (i < 0 ? 0 : parseNum(row.cells.get(c1 + i)?.v));
        const base = nm(ib);
        const pis = nm(ip);
        const cof = nm(ic);

        const ep = (base * rp) / 100;
        const ec = (base * rc) / 100;
        const dp = pis - ep;
        const dc = cof - ec;

        const rawStatus: 'OK' | 'Divergente' | 'Sem dados' =
          ib < 0 || ip < 0 || ic < 0
            ? 'Sem dados'
            : Math.abs(dp) > tolerance || Math.abs(dc) > tolerance
            ? 'Divergente'
            : 'OK';

        const justification = manualState.justifications[rowKey];
        const st =
          rawStatus === 'Divergente' && justification ? 'Justificado' : rawStatus;

        const g = (keysToTry: string[]) => {
          for (const k of keysToTry) {
            const v = d[norm(k)];
            if (v != null && String(v).trim() !== '') return String(v).trim();
          }
          return '';
        };

        const doc = g(['Documento', 'Numero Doc', 'Nota', 'NF']);
        const serie = g(['Serie', 'Série']);
        const forn = g(['Forn/Cliente', 'Cod Forn/Cli', 'Fornecedor', 'Razao Social']);
        const loja = g(['Loja']);
        const itemNf = g(['Item NF', 'Num Item', 'Cod Produto']);

        const dk = doc ? [doc, serie, forn, loja].join('|') : '';
        const ik = dk ? dk + '|' + itemNf : '';

        out.push({
          key: rowKey,
          aba: tab,
          linha: r,
          sectionIndex: secIdx,
          d,
          base,
          pis,
          cof,
          ep,
          ec,
          dp,
          dc,
          rawStatus,
          st,
          dup: false,
          mx: [],
          dk,
          ik,
          hasBlueCell,
          matchedBlueColors,
          isManualIncluded,
          isExcluded,
          justificationNote: justification?.note || '',
          justifiedAt: justification?.updatedAt,
        });
      }
    });
  }

  // Evaluate duplicates across active (non-excluded) records
  const activeRecords = out.filter((o) => !o.isExcluded);
  const ci: Record<string, number> = {};
  const da: Record<string, Set<string>> = {};

  activeRecords.forEach((o) => {
    if (o.ik) ci[o.ik] = (ci[o.ik] || 0) + 1;
    if (o.dk) {
      if (!da[o.dk]) da[o.dk] = new Set();
      da[o.dk].add(o.aba);
    }
  });

  activeRecords.forEach((o) => {
    o.dup = Boolean(o.ik && ci[o.ik] > 1);
    o.mx = o.dk ? [...da[o.dk]].filter((t) => t !== o.aba) : [];
  });

  const allCols = [...cm.entries()];
  const usedColumns = allCols.filter(([k]) => out.some((o) => k in o.d));

  return { records: out, usedColumns };
}

/**
 * Extracts all candidate rows from a sheet (including non-blue rows)
 * so the user can manually include or exclude any row in the table.
 */
export function getAllCandidateRowsForSheet(
  workbook: ParsedWorkbook | null,
  tabName: string,
  tabConfig: TabConfigMap,
  activeBlueColors: Set<string>,
  ignoreHidden: boolean,
  manualState: ManualAuditState
): RawSheetCandidateRow[] {
  if (!workbook || !workbook.sheets[tabName]) return [];
  const S = workbook.sheets[tabName];
  const sections = tabConfig[tabName] || [];
  const includedSet = new Set(manualState.includedRowKeys);
  const excludedSet = new Set(manualState.excludedRowKeys);

  const candidates: RawSheetCandidateRow[] = [];

  sections.forEach((sec, secIdx) => {
    const resolved = resolveSectionHeaders(S, sec);
    if (!resolved) return;

    const { c1, c2, names, nk, effectiveFirstDataRow } = resolved;
    const ib = findColumnIndex(nk, sec.base, [
      'Base Pis/Cofins',
      'Base Calculo',
      'Vlr Contabil',
      'B.C',
      'Valor',
      'Soma de Vl Mês',
      'Base',
    ]);
    const ip = findColumnIndex(nk, sec.pis, ['Pis', 'PIS', 'Vlr PIS', 'Soma de PIS']);
    const ic = findColumnIndex(nk, sec.cof, [
      'Cofins',
      'COFINS',
      'Vlr COFINS',
      'Soma de COFINS',
    ]);

    const maxRowKey = S.rows.size > 0 ? Math.max(...S.rows.keys()) : effectiveFirstDataRow;
    const last = +sec.b || maxRowKey;

    for (let r = +sec.a || effectiveFirstDataRow; r <= last; r++) {
      const row = S.rows.get(r);
      if (!row) continue;
      if (ignoreHidden && row.hidden) continue;

      const matchedBlueColors: string[] = [];
      let hasNonEmptyValue = false;
      const d: Record<string, string | number> = {};

      names.forEach((n, i) => {
        const cell = row.cells.get(c1 + i);
        if (cell) {
          if (cell.v !== undefined && cell.v !== '') {
            d[n] = cell.v;
            hasNonEmptyValue = true;
          }
          if (cell.bg && activeBlueColors.has(cell.bg)) {
            if (!matchedBlueColors.includes(cell.bg)) {
              matchedBlueColors.push(cell.bg);
            }
          }
        }
      });

      if (!hasNonEmptyValue) continue;

      const rowKey = `${tabName}|${r}|${secIdx}|${c1}`;
      const hasBlueCell = matchedBlueColors.length > 0;
      const isManualIncluded = includedSet.has(rowKey);
      const isExcluded = excludedSet.has(rowKey);
      const isCurrentlyIncluded = (hasBlueCell || isManualIncluded) && !isExcluded;

      const nm = (i: number) => (i < 0 ? 0 : parseNum(row.cells.get(c1 + i)?.v));
      const base = nm(ib);
      const pis = nm(ip);
      const cof = nm(ic);

      const previewText = Object.entries(d)
        .slice(0, 5)
        .map(([k, v]) => `${k}: ${v}`)
        .join(' · ');

      candidates.push({
        key: rowKey,
        aba: tabName,
        linha: r,
        sectionIndex: secIdx,
        hasBlueCell,
        matchedBlueColors,
        isCurrentlyIncluded,
        isManualIncluded,
        isExcluded,
        previewText,
        base,
        pis,
        cof,
        d,
      });
    }
  });

  return candidates;
}

export function computeTabAggregations(
  allRecords: ConsolidatedRecord[]
): TabAggregation[] {
  return EFD_TABS.map((aba) => {
    const tabRecords = allRecords.filter((r) => r.aba === aba);
    const active = tabRecords.filter((r) => !r.isExcluded);
    const excludedCount = tabRecords.filter((r) => r.isExcluded).length;
    const manualIncludedCount = active.filter((r) => r.isManualIncluded).length;

    const base = active.reduce((acc, x) => acc + x.base, 0);
    const pis = active.reduce((acc, x) => acc + x.pis, 0);
    const cof = active.reduce((acc, x) => acc + x.cof, 0);

    return {
      aba,
      n: active.length,
      excludedCount,
      manualIncludedCount,
      base,
      pis,
      cof,
      tot: pis + cof,
      div: active.filter((x) => x.st === 'Divergente').length,
      justified: active.filter((x) => x.st === 'Justificado').length,
      dup: active.filter((x) => x.dup).length,
    };
  });
}

/**
 * Exact implementation of `resumoTot()` from the user's EFD 8_2026 specification:
 * Looks for the row in `Resumo` that has cells "Total", "PIS", "COFINS"
 * and reads the values from `r + 1` (the row immediately below).
 */
export function extractResumoOfficialTotals(
  workbook: ParsedWorkbook | null
): { total: number; pis: number; cofins: number } | null {
  if (!workbook || !workbook.sheets['Resumo']) return null;
  const S = workbook.sheets['Resumo'];

  for (const [r, row] of S.rows.entries()) {
    const entries = [...row.cells.entries()];
    const nx = S.rows.get(r + 1);
    if (!nx || !entries.some(([, cell]) => norm(cell.v) === 'total')) continue;

    const cs = entries.filter(([, cell]) =>
      /^(total|pis|cofins)$/.test(norm(cell.v))
    );
    if (cs.length >= 3) {
      const o: Record<string, number> = { total: 0, pis: 0, cofins: 0 };
      cs.slice(0, 3).forEach(([c, cell]) => {
        o[norm(cell.v)] = parseNum(nx.cells.get(c)?.v);
      });
      return {
        total: o['total'] ?? 0,
        pis: o['pis'] ?? 0,
        cofins: o['cofins'] ?? 0,
      };
    }
  }
  return null;
}

export function computeResumoTotalComparisonRows(
  workbook: ParsedWorkbook | null,
  tabAggregations: TabAggregation[],
  tolerance: number
): ResumoTotalComparisonRow[] | null {
  const rr = extractResumoOfficialTotals(workbook);
  const resumoAgg = tabAggregations.find((t) => t.aba === 'Resumo');
  if (!rr || !resumoAgg) return null;

  const items: Array<[string, number, number]> = [
    ['Valor do mês (base)', rr.total, resumoAgg.base],
    ['PIS', rr.pis, resumoAgg.pis],
    ['COFINS', rr.cofins, resumoAgg.cof],
  ];

  return items.map(([item, a, b]) => ({
    item,
    resumoTotalLine: a,
    resumoBlueRows: b,
    diff: a - b,
    isOk: Math.abs(a - b) <= tolerance,
  }));
}

export function computeResumoBlockBreakdown(
  workbook: ParsedWorkbook | null,
  tabConfig: TabConfigMap,
  allRecords: ConsolidatedRecord[]
): ResumoComparisonBlock[] {
  if (!workbook || !workbook.sheets['Resumo']) return [];
  const sheet = workbook.sheets['Resumo'];
  const sections = tabConfig['Resumo'] || [];
  const activeResumoRecords = allRecords.filter(
    (r) => r.aba === 'Resumo' && !r.isExcluded
  );

  return sections.map((sec, idx) => {
    const c1 = c2n(sec.c1 || 'A');
    const colRange = `${sec.c1 || 'A'}:${sec.c2 || 'C'}`;
    const labelVal = sec.lab ? sheet.rows.get(+sec.lab)?.cells.get(c1)?.v : '';
    const label =
      labelVal != null && String(labelVal).trim() !== ''
        ? String(labelVal).trim()
        : `Bloco ${colRange}`;

    const blockRecords = activeResumoRecords.filter((x) => x.sectionIndex === idx);
    const baseAzuis = blockRecords.reduce((acc, x) => acc + x.base, 0);
    const pisAzuis = blockRecords.reduce((acc, x) => acc + x.pis, 0);
    const cofAzuis = blockRecords.reduce((acc, x) => acc + x.cof, 0);

    return {
      bloco: label,
      colRange,
      baseResumo: null,
      pisResumo: null,
      cofResumo: null,
      baseAzuis,
      pisAzuis,
      cofAzuis,
      diffBase: null,
      diffPis: null,
      diffCof: null,
      status: 'OK',
    };
  });
}

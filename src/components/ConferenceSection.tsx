import React, { useMemo, useState } from 'react';
import {
  ConsolidatedRecord,
  EFD_TABS,
  ParsedWorkbook,
  ResumoComparisonBlock,
  ResumoTotalComparisonRow,
  TabAggregation,
} from '../types/efd';
import { formatCurrency, norm } from '../utils/excelParser';
import {
  AlertTriangle,
  Archive,
  CheckCircle2,
  Download,
  FileCheck2,
  ListPlus,
  Search,
  Undo2,
  XCircle,
} from 'lucide-react';

interface ConferenceSectionProps {
  workbook: ParsedWorkbook | null;
  allRecords: ConsolidatedRecord[];
  usedColumns: Array<[string, string]>;
  tabAggregations: TabAggregation[];
  resumoTotalRows: ResumoTotalComparisonRow[] | null;
  resumoBlocks: ResumoComparisonBlock[];
  onToggleRowInclusion: (
    rowKey: string,
    hasBlueCell: boolean,
    currentlyIncluded: boolean
  ) => void;
  onSaveJustification: (rowKey: string, note: string) => void;
  onClearJustification: (rowKey: string) => void;
  onOpenManualRowModal: (tabName?: string) => void;
  onExportExcel: () => void;
  onSaveToHistory: (competencia: string, notes: string) => void;
  isExporting: boolean;
}

type RecordFilterMode =
  | 'active'
  | 'divergent'
  | 'justified'
  | 'duplicate'
  | 'manual_included'
  | 'excluded';

const QUICK_JUSTIFICATIONS = [
  'Alíquota diferenciada / regime especial',
  'Diferença de arredondamento de centavos na NF',
  'Crédito parcial proporcional ao rateio',
  'Conferido com documento fiscal original',
];

export const ConferenceSection: React.FC<ConferenceSectionProps> = ({
  workbook,
  allRecords,
  usedColumns,
  tabAggregations,
  resumoTotalRows,
  resumoBlocks,
  onToggleRowInclusion,
  onSaveJustification,
  onClearJustification,
  onOpenManualRowModal,
  onExportExcel,
  onSaveToHistory,
  isExporting,
}) => {
  const [selectedTab, setSelectedTab] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterMode, setFilterMode] = useState<RecordFilterMode>('active');

  // Justification inline editing state
  const [editingRowKey, setEditingRowKey] = useState<string | null>(null);
  const [justificationDraft, setJustificationDraft] = useState<string>('');

  // Save to history inline form state
  const [showSaveHistoryForm, setShowSaveHistoryForm] = useState(false);
  const [competenciaInput, setCompetenciaInput] = useState('08/2026');
  const [historyNotesInput, setHistoryNotesInput] = useState('');
  const [savedToast, setSavedToast] = useState(false);

  const grandTotals = useMemo(() => {
    return tabAggregations.reduce(
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
  }, [tabAggregations]);

  const filteredRecords = useMemo(() => {
    const q = norm(searchQuery);
    return allRecords.filter((o) => {
      if (selectedTab && o.aba !== selectedTab) return false;

      if (filterMode === 'excluded') {
        if (!o.isExcluded) return false;
      } else {
        if (o.isExcluded) return false;
        if (filterMode === 'divergent' && o.st !== 'Divergente') return false;
        if (filterMode === 'justified' && o.st !== 'Justificado') return false;
        if (filterMode === 'duplicate' && !o.dup && o.mx.length === 0)
          return false;
        if (filterMode === 'manual_included' && !o.isManualIncluded)
          return false;
      }

      if (q) {
        const hay = norm(
          `${o.aba} ${o.linha} ${Object.values(o.d).join(' ')} ${o.justificationNote}`
        );
        if (!hay.includes(q)) return false;
      }

      return true;
    });
  }, [allRecords, selectedTab, filterMode, searchQuery]);

  const handleOpenJustify = (rec: ConsolidatedRecord) => {
    setEditingRowKey(rec.key);
    setJustificationDraft(
      rec.justificationNote || 'Alíquota diferenciada / regime especial'
    );
  };

  const handleConfirmJustify = (rowKey: string) => {
    if (!justificationDraft.trim()) return;
    onSaveJustification(rowKey, justificationDraft.trim());
    setEditingRowKey(null);
    setJustificationDraft('');
  };

  const handleSaveClosing = (e: React.FormEvent) => {
    e.preventDefault();
    if (!competenciaInput.trim()) return;
    onSaveToHistory(competenciaInput.trim(), historyNotesInput.trim());
    setShowSaveHistoryForm(false);
    setHistoryNotesInput('');
    setSavedToast(true);
    setTimeout(() => setSavedToast(false), 3500);
  };

  const formatCellValue = (v: string | number | undefined) => {
    if (v === undefined || v === '') return '';
    if (typeof v === 'number') {
      return v.toLocaleString('pt-BR', { maximumFractionDigits: 4 });
    }
    return String(v);
  };

  return (
    <section className="space-y-6">
      {/* Executive KPI Strip */}
      <div className="bg-white border border-slate-200 rounded-lg p-6">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 pb-5 border-b border-slate-200">
          <div>
            <h2 className="text-lg font-bold text-slate-900 tracking-tight">
              03. Painel de Conferência Fiscal e Consolidação de Créditos
            </h2>
            <p className="text-sm text-slate-600 mt-0.5">
              Totalização automática por categoria, auditoria de alíquotas de PIS e COFINS, verificação de duplicidades e confronto com a aba Resumo.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <button
              type="button"
              onClick={() => onOpenManualRowModal(selectedTab || undefined)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-800 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
            >
              <ListPlus className="w-4 h-4 text-blue-900" />
              Incluir / Excluir Linhas da Planilha
            </button>

            <button
              type="button"
              onClick={() => setShowSaveHistoryForm((prev) => !prev)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-800 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
            >
              <Archive className="w-4 h-4 text-slate-700" />
              Salvar no Histórico Mensal
            </button>

            <button
              type="button"
              onClick={onExportExcel}
              disabled={isExporting || grandTotals.n === 0}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-blue-900 rounded-md hover:bg-blue-800 disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
            >
              <Download className="w-4 h-4" />
              {isExporting ? 'Gerando Excel...' : 'Exportar Excel Consolidado'}
            </button>
          </div>
        </div>

        {/* Save to History Drawer */}
        {showSaveHistoryForm && (
          <form
            onSubmit={handleSaveClosing}
            className="mt-4 p-4 bg-slate-50 border border-slate-200 rounded-md flex flex-wrap items-end gap-4"
          >
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Competência (Mês/Ano)
              </label>
              <input
                type="text"
                value={competenciaInput}
                onChange={(e) => setCompetenciaInput(e.target.value)}
                placeholder="Ex: 08/2026"
                className="text-xs font-mono tabular-nums px-3 py-1.5 bg-white border border-slate-300 rounded-md w-36 focus:outline-none focus:ring-2 focus:ring-blue-800"
                required
              />
            </div>
            <div className="flex-1 min-w-[220px]">
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Observações do Fechamento (opcional)
              </label>
              <input
                type="text"
                value={historyNotesInput}
                onChange={(e) => setHistoryNotesInput(e.target.value)}
                placeholder="Ex: Fechamento revisado após conciliação da aba Resumo"
                className="text-xs px-3 py-1.5 bg-white border border-slate-300 rounded-md w-full focus:outline-none focus:ring-2 focus:ring-blue-800"
              />
            </div>
            <div className="flex items-center gap-2">
              <button
                type="submit"
                className="px-4 py-1.5 text-xs font-semibold text-white bg-slate-900 rounded-md hover:bg-slate-800 transition-colors cursor-pointer whitespace-nowrap shrink-0"
              >
                Confirmar Salvamento
              </button>
              <button
                type="button"
                onClick={() => setShowSaveHistoryForm(false)}
                className="px-3 py-1.5 text-xs text-slate-600 hover:text-slate-900 cursor-pointer"
              >
                Cancelar
              </button>
            </div>
          </form>
        )}

        {savedToast && (
          <div className="mt-4 px-4 py-2.5 bg-emerald-50 border border-emerald-200 rounded-md text-xs text-emerald-900 font-medium flex items-center justify-between">
            <span>
              Fechamento da competência salvo com sucesso no Histórico Mensal.
            </span>
          </div>
        )}

        {/* 7-Metric Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-7 gap-4 pt-5">
          <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-md">
            <div className="text-xs text-slate-600">Registros Ativos</div>
            <div className="text-xl font-bold font-mono tabular-nums text-slate-900 mt-1">
              {grandTotals.n.toLocaleString('pt-BR')}
            </div>
            <div className="text-[11px] text-slate-500 mt-1">
              {grandTotals.manInc > 0 || grandTotals.manExc > 0
                ? `+${grandTotals.manInc} incl. · -${grandTotals.manExc} excl.`
                : 'Linhas azuis ativas'}
            </div>
          </div>

          <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-md">
            <div className="text-xs text-slate-600">Base de Cálculo (R$)</div>
            <div className="text-lg font-bold font-mono tabular-nums text-slate-900 mt-1">
              {formatCurrency(grandTotals.b)}
            </div>
            <div className="text-[11px] text-slate-500 mt-1">
              Soma padronizada
            </div>
          </div>

          <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-md">
            <div className="text-xs text-slate-600">Crédito PIS (R$)</div>
            <div className="text-lg font-bold font-mono tabular-nums text-slate-900 mt-1">
              {formatCurrency(grandTotals.p)}
            </div>
            <div className="text-[11px] text-slate-500 mt-1">
              Apurado nas abas
            </div>
          </div>

          <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-md">
            <div className="text-xs text-slate-600">Crédito COFINS (R$)</div>
            <div className="text-lg font-bold font-mono tabular-nums text-slate-900 mt-1">
              {formatCurrency(grandTotals.c)}
            </div>
            <div className="text-[11px] text-slate-500 mt-1">
              Apurado nas abas
            </div>
          </div>

          <div className="p-3.5 bg-blue-950 text-white border border-blue-900 rounded-md">
            <div className="text-xs text-blue-200">Total de Créditos (R$)</div>
            <div className="text-lg font-bold font-mono tabular-nums text-white mt-1">
              {formatCurrency(grandTotals.p + grandTotals.c)}
            </div>
            <div className="text-[11px] text-blue-200 mt-1">
              PIS + COFINS do mês
            </div>
          </div>

          <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-md">
            <div className="text-xs text-slate-600">Divergências Cálculo</div>
            <div
              className={`text-xl font-bold font-mono tabular-nums mt-1 ${
                grandTotals.v > 0 ? 'text-red-700' : 'text-emerald-700'
              }`}
            >
              {grandTotals.v}
            </div>
            <div className="text-[11px] text-slate-500 mt-1">
              {grandTotals.j > 0
                ? `${grandTotals.j} justificada(s)`
                : 'Pendentes de revisão'}
            </div>
          </div>

          <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-md">
            <div className="text-xs text-slate-600">Itens Duplicados</div>
            <div
              className={`text-xl font-bold font-mono tabular-nums mt-1 ${
                grandTotals.d > 0 ? 'text-amber-700' : 'text-emerald-700'
              }`}
            >
              {grandTotals.d}
            </div>
            <div className="text-[11px] text-slate-500 mt-1">
              Chave NF + Item
            </div>
          </div>
        </div>
      </div>

      {/* Category Summary & Resumo Tab Reconciliation Split */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Category Breakdown Table (7 cols) */}
        <div className="lg:col-span-7 bg-white border border-slate-200 rounded-lg p-6">
          <div className="flex items-center justify-between mb-3">
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Totalização de Créditos por Categoria (Aba)
              </h3>
              <p className="text-xs text-slate-600">
                Clique em uma aba para filtrar os registros na grade de auditoria abaixo.
              </p>
            </div>
            {selectedTab && (
              <button
                type="button"
                onClick={() => setSelectedTab('')}
                className="text-xs font-medium text-blue-900 hover:underline cursor-pointer"
              >
                Mostrar todas as abas
              </button>
            )}
          </div>

          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                  <th className="py-2.5 px-3">Aba</th>
                  <th className="py-2.5 px-2.5 text-right">Registros</th>
                  <th className="py-2.5 px-2.5 text-right">Base (R$)</th>
                  <th className="py-2.5 px-2.5 text-right">PIS (R$)</th>
                  <th className="py-2.5 px-2.5 text-right">COFINS (R$)</th>
                  <th className="py-2.5 px-2.5 text-right">Total (R$)</th>
                  <th className="py-2.5 px-2.5 text-right">Diverg.</th>
                  <th className="py-2.5 px-2.5 text-right">Duplic.</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {tabAggregations
                  .filter((a) => !workbook || workbook.sheets[a.aba])
                  .map((a) => {
                    const isSelected = selectedTab === a.aba;
                    return (
                      <tr
                        key={a.aba}
                        onClick={() =>
                          setSelectedTab(isSelected ? '' : a.aba)
                        }
                        className={`cursor-pointer transition-colors ${
                          isSelected
                            ? 'bg-blue-50/90 font-semibold'
                            : 'hover:bg-slate-50'
                        }`}
                      >
                        <td className="py-2 px-3 text-slate-900 whitespace-nowrap">
                          <span>{a.aba}</span>
                          {(a.manualIncludedCount > 0 || a.excludedCount > 0) && (
                            <span className="ml-1.5 text-[11px] font-normal text-slate-500">
                              · ajustada
                            </span>
                          )}
                        </td>
                        <td className="py-2 px-2.5 text-right font-mono tabular-nums">
                          {a.n}
                        </td>
                        <td className="py-2 px-2.5 text-right font-mono tabular-nums">
                          {formatCurrency(a.base)}
                        </td>
                        <td className="py-2 px-2.5 text-right font-mono tabular-nums">
                          {formatCurrency(a.pis)}
                        </td>
                        <td className="py-2 px-2.5 text-right font-mono tabular-nums">
                          {formatCurrency(a.cof)}
                        </td>
                        <td className="py-2 px-2.5 text-right font-mono tabular-nums font-medium text-slate-900">
                          {formatCurrency(a.tot)}
                        </td>
                        <td
                          className={`py-2 px-2.5 text-right font-mono tabular-nums ${
                            a.div > 0
                              ? 'text-red-700 font-semibold'
                              : 'text-slate-600'
                          }`}
                        >
                          {a.div}
                          {a.justified > 0 && (
                            <span className="text-amber-700 font-normal ml-1">
                              ({a.justified} rev.)
                            </span>
                          )}
                        </td>
                        <td
                          className={`py-2 px-2.5 text-right font-mono tabular-nums ${
                            a.dup > 0
                              ? 'text-amber-700 font-semibold'
                              : 'text-slate-600'
                          }`}
                        >
                          {a.dup}
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
              <tfoot>
                <tr className="bg-slate-100 border-t-2 border-slate-300 font-bold text-slate-900">
                  <td className="py-2.5 px-3">Total Geral</td>
                  <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                    {grandTotals.n}
                  </td>
                  <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                    {formatCurrency(grandTotals.b)}
                  </td>
                  <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                    {formatCurrency(grandTotals.p)}
                  </td>
                  <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                    {formatCurrency(grandTotals.c)}
                  </td>
                  <td className="py-2.5 px-2.5 text-right font-mono tabular-nums text-blue-950">
                    {formatCurrency(grandTotals.p + grandTotals.c)}
                  </td>
                  <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                    {grandTotals.v}
                  </td>
                  <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                    {grandTotals.d}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>

        {/* Comparison with Resumo Tab (5 cols) */}
        <div className="lg:col-span-5 bg-white border border-slate-200 rounded-lg p-6 flex flex-col justify-between space-y-4">
          <div>
            <h3 className="text-sm font-bold text-slate-900">
              Comparação com a Aba Resumo
            </h3>
            <p className="text-xs text-slate-600 mb-3">
              Confronto direto entre a linha <span className="font-mono">"Total"</span> localizada na aba Resumo e a soma das linhas azuis capturadas na aba Resumo.
            </p>

            {resumoTotalRows ? (
              <div className="overflow-x-auto border border-slate-200 rounded-md">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                      <th className="py-2.5 px-3">Item</th>
                      <th className="py-2.5 px-2.5 text-right">
                        Linha "Total" (Resumo)
                      </th>
                      <th className="py-2.5 px-2.5 text-right">
                        Azuis (Resumo)
                      </th>
                      <th className="py-2.5 px-2.5 text-right">Diferença</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {resumoTotalRows.map((row) => (
                      <tr key={row.item} className="hover:bg-slate-50">
                        <td className="py-2.5 px-3 font-medium text-slate-900">
                          {row.item}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                          {formatCurrency(row.resumoTotalLine)}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono tabular-nums">
                          {formatCurrency(row.resumoBlueRows)}
                        </td>
                        <td
                          className={`py-2.5 px-2.5 text-right font-mono tabular-nums font-semibold ${
                            row.isOk ? 'text-emerald-700' : 'text-red-700'
                          }`}
                        >
                          {formatCurrency(row.diff)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-md text-xs text-slate-600">
                Linha "Total" não localizada na aba Resumo (ou aba Resumo ausente no arquivo).
              </div>
            )}
          </div>

          {/* Breakdown of Resumo Blocks */}
          {resumoBlocks.length > 0 && (
            <div className="pt-3 border-t border-slate-200">
              <div className="text-xs font-semibold text-slate-800 mb-2">
                Detalhamento por Bloco Lido na Aba Resumo
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {resumoBlocks.map((blk, i) => (
                  <div
                    key={i}
                    className="p-3 bg-slate-50 border border-slate-200 rounded-md text-xs"
                  >
                    <div className="font-semibold text-slate-900 truncate" title={blk.bloco}>
                      {blk.bloco}
                    </div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      Colunas {blk.colRange}
                    </div>
                    <div className="mt-2 space-y-1 font-mono tabular-nums text-[11px]">
                      <div className="flex justify-between">
                        <span className="text-slate-600">Base:</span>
                        <span className="font-medium text-slate-900">
                          R$ {formatCurrency(blk.baseAzuis)}
                        </span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-slate-600">PIS:</span>
                        <span className="font-medium text-slate-900">
                          R$ {formatCurrency(blk.pisAzuis)}
                        </span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-slate-600">COFINS:</span>
                        <span className="font-medium text-slate-900">
                          R$ {formatCurrency(blk.cofAzuis)}
                        </span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Detailed Audit & Manual Review Table */}
      <div className="bg-white border border-slate-200 rounded-lg p-6 space-y-4">
        <div className="flex flex-col xl:flex-row xl:items-center xl:justify-between gap-4">
          <div>
            <h3 className="text-sm font-bold text-slate-900">
              Registros Consolidados e Revisão Manual de Auditoria
            </h3>
            <p className="text-xs text-slate-600 mt-0.5">
              Exibindo{' '}
              <span className="font-mono tabular-nums font-semibold text-slate-900">
                {filteredRecords.length}
              </span>{' '}
              registro(s)
              {filteredRecords.length > 350
                ? ' (mostrando os primeiros 350 em tela; a exportação Excel inclui todos os registros)'
                : ''}
              . Você pode justificar divergências ou excluir/incluir linhas diretamente.
            </p>
          </div>

          {/* Filter Controls Bar */}
          <div className="flex flex-wrap items-center gap-2.5">
            <select
              value={selectedTab}
              onChange={(e) => setSelectedTab(e.target.value)}
              className="text-xs font-medium bg-slate-50 border border-slate-300 rounded-md px-3 py-1.5 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-800"
            >
              <option value="">Todas as abas</option>
              {EFD_TABS.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>

            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Buscar NF, fornecedor, conta..."
                className="pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-300 rounded-md w-52 focus:outline-none focus:ring-2 focus:ring-blue-800"
              />
            </div>

            {/* Interactive Segmented Filter Buttons */}
            <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-md">
              {(
                [
                  ['active', `Ativos (${grandTotals.n})`],
                  ['divergent', `Divergentes (${grandTotals.v})`],
                  ['justified', `Justificados (${grandTotals.j})`],
                  ['duplicate', `Duplicados (${grandTotals.d})`],
                  ['manual_included', `Incluídos Man. (${grandTotals.manInc})`],
                  ['excluded', `Excluídos (${grandTotals.manExc})`],
                ] as Array<[RecordFilterMode, string]>
              ).map(([mode, label]) => (
                <button
                  key={mode}
                  type="button"
                  onClick={() => setFilterMode(mode)}
                  className={`px-2.5 py-1 text-xs font-medium rounded transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                    filterMode === mode
                      ? 'bg-white text-slate-900 shadow-2xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {label}
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* Records Grid */}
        <div className="overflow-auto max-h-[540px] border border-slate-200 rounded-md">
          <table className="w-full border-collapse text-left text-xs">
            <thead className="sticky top-0 z-10">
              <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                <th className="py-2.5 px-2.5 whitespace-nowrap">Revisão Linha</th>
                <th className="py-2.5 px-2.5 whitespace-nowrap">Aba de origem</th>
                <th className="py-2.5 px-2.5 text-right whitespace-nowrap">
                  Linha Excel
                </th>
                <th className="py-2.5 px-2.5 text-right whitespace-nowrap">
                  Base (R$)
                </th>
                <th className="py-2.5 px-2.5 text-right whitespace-nowrap">
                  PIS (R$)
                </th>
                <th className="py-2.5 px-2.5 text-right whitespace-nowrap">
                  COFINS (R$)
                </th>
                <th className="py-2.5 px-2.5 whitespace-nowrap">
                  Status & Auditoria Fiscal
                </th>
                <th className="py-2.5 px-2.5 whitespace-nowrap">Duplicado</th>
                {usedColumns.map(([k, label]) => (
                  <th key={k} className="py-2.5 px-2.5 whitespace-nowrap">
                    {label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filteredRecords.length === 0 ? (
                <tr>
                  <td
                    colSpan={8 + usedColumns.length}
                    className="py-10 text-center text-sm text-slate-500"
                  >
                    Nenhum registro encontrado para os filtros selecionados.
                  </td>
                </tr>
              ) : (
                filteredRecords.slice(0, 350).map((o) => {
                  const isEditingThis = editingRowKey === o.key;
                  return (
                    <React.Fragment key={o.key}>
                      <tr
                        className={`hover:bg-slate-50/90 ${
                          o.isExcluded ? 'opacity-60 bg-slate-50' : ''
                        }`}
                      >
                        {/* Manual Include/Exclude Action */}
                        <td className="py-2 px-2.5 whitespace-nowrap">
                          <button
                            type="button"
                            onClick={() =>
                              onToggleRowInclusion(
                                o.key,
                                o.hasBlueCell,
                                !o.isExcluded
                              )
                            }
                            className={`inline-flex items-center gap-1 px-2 py-1 rounded text-[11px] font-medium border transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                              o.isExcluded
                                ? 'bg-slate-100 text-slate-700 border-slate-300 hover:bg-emerald-50 hover:text-emerald-800 hover:border-emerald-300'
                                : 'bg-white text-slate-700 border-slate-200 hover:bg-red-50 hover:text-red-700 hover:border-red-200'
                            }`}
                            title={
                              o.isExcluded
                                ? 'Restaurar esta linha na apuração'
                                : 'Excluir manualmente esta linha da apuração'
                            }
                          >
                            {o.isExcluded ? (
                              <>
                                <Undo2 className="w-3 h-3" />
                                Restaurar
                              </>
                            ) : (
                              <>
                                <XCircle className="w-3 h-3 text-slate-400" />
                                Excluir
                              </>
                            )}
                          </button>
                        </td>

                        <td className="py-2 px-2.5 font-medium text-slate-900 whitespace-nowrap">
                          <span>{o.aba}</span>
                          {o.isManualIncluded && (
                            <span className="ml-1.5 text-[11px] text-blue-800 font-normal">
                              · manual
                            </span>
                          )}
                        </td>

                        <td className="py-2 px-2.5 text-right font-mono tabular-nums text-slate-700 whitespace-nowrap">
                          {o.linha}
                        </td>

                        <td className="py-2 px-2.5 text-right font-mono tabular-nums font-medium text-slate-900 whitespace-nowrap">
                          {formatCurrency(o.base)}
                        </td>

                        <td className="py-2 px-2.5 text-right font-mono tabular-nums text-slate-900 whitespace-nowrap">
                          <div>{formatCurrency(o.pis)}</div>
                          {o.rawStatus === 'Divergente' && (
                            <div className="text-[10px] text-slate-500">
                              Esp: {formatCurrency(o.ep)}
                            </div>
                          )}
                        </td>

                        <td className="py-2 px-2.5 text-right font-mono tabular-nums text-slate-900 whitespace-nowrap">
                          <div>{formatCurrency(o.cof)}</div>
                          {o.rawStatus === 'Divergente' && (
                            <div className="text-[10px] text-slate-500">
                              Esp: {formatCurrency(o.ec)}
                            </div>
                          )}
                        </td>

                        {/* Status & Justification Column */}
                        <td className="py-2 px-2.5 whitespace-nowrap">
                          <div className="flex items-center gap-2">
                            {o.st === 'OK' && (
                              <span className="inline-flex items-center gap-1 text-emerald-700 font-medium">
                                <CheckCircle2 className="w-3.5 h-3.5" />
                                OK
                              </span>
                            )}

                            {o.st === 'Divergente' && (
                              <>
                                <span className="inline-flex items-center gap-1 text-red-700 font-semibold">
                                  <AlertTriangle className="w-3.5 h-3.5" />
                                  Divergente
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleOpenJustify(o)}
                                  className="px-2 py-0.5 text-[11px] font-medium text-blue-900 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded transition-colors cursor-pointer"
                                >
                                  Justificar
                                </button>
                              </>
                            )}

                            {o.st === 'Justificado' && (
                              <div className="flex items-center gap-1.5">
                                <span className="inline-flex items-center gap-1 text-amber-800 font-semibold">
                                  <FileCheck2 className="w-3.5 h-3.5" />
                                  Justificado
                                </span>
                                <span
                                  className="text-slate-600 max-w-[160px] truncate"
                                  title={o.justificationNote}
                                >
                                  · {o.justificationNote}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleOpenJustify(o)}
                                  className="text-[11px] text-blue-900 hover:underline cursor-pointer"
                                >
                                  Editar
                                </button>
                                <button
                                  type="button"
                                  onClick={() => onClearJustification(o.key)}
                                  className="text-[11px] text-red-700 hover:underline cursor-pointer"
                                >
                                  Remover
                                </button>
                              </div>
                            )}

                            {o.st === 'Sem dados' && (
                              <span className="text-slate-500">Sem dados</span>
                            )}
                          </div>
                        </td>

                        {/* Duplicate Status */}
                        <td className="py-2 px-2.5 whitespace-nowrap">
                          {o.dup ? (
                            <span className="text-amber-800 font-semibold">
                              Sim (Item repetido)
                            </span>
                          ) : o.mx.length > 0 ? (
                            <span className="text-amber-700">
                              Doc em: {o.mx.join(', ')}
                            </span>
                          ) : (
                            <span className="text-slate-500">Não</span>
                          )}
                        </td>

                        {/* Dynamic Spreadsheet Columns */}
                        {usedColumns.map(([k]) => {
                          const cellVal = o.d[k];
                          const isNum = typeof cellVal === 'number';
                          return (
                            <td
                              key={k}
                              className={`py-2 px-2.5 whitespace-nowrap text-slate-700 ${
                                isNum ? 'text-right font-mono tabular-nums' : ''
                              }`}
                            >
                              {formatCellValue(cellVal)}
                            </td>
                          );
                        })}
                      </tr>

                      {/* Inline Justification Row Editor */}
                      {isEditingThis && (
                        <tr className="bg-blue-50/60 border-b border-blue-200">
                          <td colSpan={8 + usedColumns.length} className="p-3">
                            <div className="flex flex-wrap items-center gap-3">
                              <span className="text-xs font-semibold text-slate-900">
                                Justificar divergência ({o.aba} · Linha {o.linha} · Dif. PIS:{' '}
                                {formatCurrency(o.dp)} / Dif. COFINS: {formatCurrency(o.dc)}):
                              </span>
                              <input
                                type="text"
                                value={justificationDraft}
                                onChange={(e) =>
                                  setJustificationDraft(e.target.value)
                                }
                                placeholder="Informe o motivo fiscal ou selecione abaixo..."
                                className="text-xs px-3 py-1.5 bg-white border border-slate-300 rounded-md w-80 focus:outline-none focus:ring-2 focus:ring-blue-800"
                                autoFocus
                              />
                              <div className="flex flex-wrap items-center gap-1.5">
                                {QUICK_JUSTIFICATIONS.map((preset) => (
                                  <button
                                    key={preset}
                                    type="button"
                                    onClick={() => setJustificationDraft(preset)}
                                    className="px-2 py-1 text-[11px] bg-white border border-slate-200 rounded hover:bg-slate-100 text-slate-700 cursor-pointer"
                                  >
                                    {preset}
                                  </button>
                                ))}
                              </div>
                              <div className="flex items-center gap-2 ml-auto">
                                <button
                                  type="button"
                                  onClick={() => handleConfirmJustify(o.key)}
                                  className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-900 rounded hover:bg-blue-800 cursor-pointer"
                                >
                                  Salvar Justificativa
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setEditingRowKey(null)}
                                  className="px-2.5 py-1.5 text-xs text-slate-600 hover:text-slate-900 cursor-pointer"
                                >
                                  Cancelar
                                </button>
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
};

import React, { useState } from 'react';
import {
  AccountingCategoryBlock,
  AccountingMapEntry,
  DEFAULT_ACCOUNTING_MAP,
} from '../utils/accountingDynamics';
import { formatCurrency } from '../utils/excelParser';
import {
  BookOpen,
  Download,
  Plus,
  RotateCcw,
  Trash2,
} from 'lucide-react';

interface AccountingDynamicsSectionProps {
  blocks: AccountingCategoryBlock[];
  accountingMap: AccountingMapEntry[];
  onChangeAccountingMap: (next: AccountingMapEntry[]) => void;
  competencia: string;
  onChangeCompetencia: (val: string) => void;
  ctaDebitoPis: string;
  onChangeCtaDebitoPis: (val: string) => void;
  ctaDebitoCofins: string;
  onChangeCtaDebitoCofins: (val: string) => void;
  onExportExcel: () => void;
  isExporting: boolean;
}

export const AccountingDynamicsSection: React.FC<
  AccountingDynamicsSectionProps
> = ({
  blocks,
  accountingMap,
  onChangeAccountingMap,
  competencia,
  onChangeCompetencia,
  ctaDebitoPis,
  onChangeCtaDebitoPis,
  ctaDebitoCofins,
  onChangeCtaDebitoCofins,
  onExportExcel,
  isExporting,
}) => {
  const [showMapEditor, setShowMapEditor] = useState(false);
  const [newLanc, setNewLanc] = useState('');
  const [newCred, setNewCred] = useState('');
  const [newDesc, setNewDesc] = useState('');

  const activeBlocks = blocks.filter((b) => b.filialGroups.length > 0);

  const handleAddMapping = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newLanc.trim() || !newCred.trim() || !newDesc.trim()) return;
    const next = [
      {
        ctaLancamento: newLanc.trim(),
        ctaCredito: newCred.trim(),
        descricao: newDesc.trim().toUpperCase(),
      },
      ...accountingMap,
    ];
    onChangeAccountingMap(next);
    setNewLanc('');
    setNewCred('');
    setNewDesc('');
  };

  const handleDeleteMapping = (idx: number) => {
    const next = [...accountingMap];
    next.splice(idx, 1);
    onChangeAccountingMap(next);
  };

  const handleResetMap = () => {
    onChangeAccountingMap(DEFAULT_ACCOUNTING_MAP);
  };

  return (
    <section className="bg-white border border-slate-200 rounded-lg p-6 space-y-6">
      {/* Header & Controls */}
      <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 pb-5 border-b border-slate-200">
        <div>
          <h2 className="text-lg font-bold text-slate-900 tracking-tight">
            04. Dinâmica de Contabilização de Créditos (De-Para Contábil)
          </h2>
          <p className="text-sm text-slate-600 mt-0.5">
            Pré-visualização exata da guia <span className="font-mono font-semibold text-slate-800">Dinâmica Contábil</span> que será gerada no Excel exportado, agrupada por categoria, filial e conta de lançamento com histórico padronizado.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <button
            type="button"
            onClick={() => setShowMapEditor((prev) => !prev)}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-800 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            <BookOpen className="w-4 h-4 text-blue-900" />
            {showMapEditor
              ? 'Ocultar Mapa Contábil'
              : `Gerenciar Mapa Contábil (${accountingMap.length} contas)`}
          </button>

          <button
            type="button"
            onClick={onExportExcel}
            disabled={isExporting || activeBlocks.length === 0}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-blue-900 rounded-md hover:bg-blue-800 disabled:opacity-50 transition-colors cursor-pointer whitespace-nowrap shrink-0"
          >
            <Download className="w-4 h-4" />
            {isExporting
              ? 'Exportando...'
              : 'Exportar Excel com Dinâmica Contábil'}
          </button>
        </div>
      </div>

      {/* Accounting Parameters Bar */}
      <div className="flex flex-wrap items-center gap-6 p-4 bg-slate-50 border border-slate-200 rounded-md text-xs">
        <label className="flex items-center gap-2">
          <span className="font-semibold text-slate-700">
            Competência no Histórico:
          </span>
          <input
            type="text"
            value={competencia}
            onChange={(e) => onChangeCompetencia(e.target.value)}
            placeholder="08/2026"
            className="w-24 font-mono tabular-nums px-2.5 py-1.5 bg-white border border-slate-300 rounded text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-800"
          />
        </label>

        <label className="flex items-center gap-2">
          <span className="font-semibold text-slate-700">
            Cta Débito PIS:
          </span>
          <input
            type="text"
            value={ctaDebitoPis}
            onChange={(e) => onChangeCtaDebitoPis(e.target.value)}
            className="w-32 font-mono tabular-nums px-2.5 py-1.5 bg-white border border-slate-300 rounded text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-800"
          />
        </label>

        <label className="flex items-center gap-2">
          <span className="font-semibold text-slate-700">
            Cta Débito COFINS:
          </span>
          <input
            type="text"
            value={ctaDebitoCofins}
            onChange={(e) => onChangeCtaDebitoCofins(e.target.value)}
            className="w-32 font-mono tabular-nums px-2.5 py-1.5 bg-white border border-slate-300 rounded text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-800"
          />
        </label>
      </div>

      {/* Collapsible Accounting Map Editor */}
      {showMapEditor && (
        <div className="p-4 bg-slate-50 border border-slate-200 rounded-md space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Mapa De-Para Contábil (Cta Lançamento → Cta Crédito → Descrição)
              </h3>
              <p className="text-xs text-slate-600">
                Adicione novas contas contábeis caso surjam novos itens no fechamento mensal.
              </p>
            </div>
            <button
              type="button"
              onClick={handleResetMap}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-100 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Restaurar Mapa Padrão
            </button>
          </div>

          <form
            onSubmit={handleAddMapping}
            className="flex flex-wrap items-end gap-3 pt-1"
          >
            <div>
              <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                Cta Lançamento
              </label>
              <input
                type="text"
                value={newLanc}
                onChange={(e) => setNewLanc(e.target.value)}
                placeholder="Ex: 11040100001"
                className="text-xs font-mono tabular-nums px-2.5 py-1.5 bg-white border border-slate-300 rounded w-36"
                required
              />
            </div>
            <div>
              <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                Cta Crédito
              </label>
              <input
                type="text"
                value={newCred}
                onChange={(e) => setNewCred(e.target.value)}
                placeholder="Ex: 41010200001"
                className="text-xs font-mono tabular-nums px-2.5 py-1.5 bg-white border border-slate-300 rounded w-36"
                required
              />
            </div>
            <div className="flex-1 min-w-[220px]">
              <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                Descrição para Histórico
              </label>
              <input
                type="text"
                value={newDesc}
                onChange={(e) => setNewDesc(e.target.value)}
                placeholder="Ex: MU01 - EPI"
                className="text-xs px-2.5 py-1.5 bg-white border border-slate-300 rounded w-full"
                required
              />
            </div>
            <button
              type="submit"
              className="inline-flex items-center gap-1 px-3.5 py-1.5 text-xs font-semibold text-white bg-blue-900 rounded hover:bg-blue-800 cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              Adicionar Conta
            </button>
          </form>

          <div className="max-h-60 overflow-auto border border-slate-200 rounded bg-white">
            <table className="w-full border-collapse text-left text-xs">
              <thead className="sticky top-0 bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                <tr>
                  <th className="py-2 px-3">Cta Lançamento</th>
                  <th className="py-2 px-3">Cta Crédito</th>
                  <th className="py-2 px-3">Descrição</th>
                  <th className="py-2 px-3 text-right">Ação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {accountingMap.map((m, idx) => (
                  <tr
                    key={`${m.ctaLancamento}-${idx}`}
                    className="hover:bg-slate-50"
                  >
                    <td className="py-1.5 px-3 font-mono tabular-nums font-medium text-slate-900">
                      {m.ctaLancamento}
                    </td>
                    <td className="py-1.5 px-3 font-mono tabular-nums text-slate-700">
                      {m.ctaCredito}
                    </td>
                    <td className="py-1.5 px-3 text-slate-800">{m.descricao}</td>
                    <td className="py-1.5 px-3 text-right">
                      <button
                        type="button"
                        onClick={() => handleDeleteMapping(idx)}
                        className="text-slate-400 hover:text-red-700 cursor-pointer"
                        title="Remover linha do mapa"
                      >
                        <Trash2 className="w-3.5 h-3.5 inline" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Side-by-Side / Stacked Preview of Category Pivot Tables */}
      {activeBlocks.length === 0 ? (
        <div className="p-8 bg-slate-50 border border-slate-200 rounded-md text-center text-xs text-slate-500">
          Nenhum registro consolidado ativo para gerar a dinâmica contábil.
        </div>
      ) : (
        <div className="space-y-6">
          {activeBlocks.map((block) => (
            <div
              key={block.blockCode}
              className="border border-slate-200 rounded-md overflow-hidden"
            >
              {/* Block Top Header */}
              <div className="px-4 py-2.5 bg-slate-100 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
                <div className="flex items-center gap-3">
                  <span className="font-bold text-sm text-slate-900">
                    {block.blockCode}
                  </span>
                  <span className="text-xs text-slate-500">
                    · Aba de origem: {block.sourceTab}
                  </span>
                </div>
                <div className="flex items-center gap-4 text-xs font-mono tabular-nums text-slate-700">
                  <span>
                    Cta Débito PIS: <strong>{block.ctaDebitoPis}</strong>
                  </span>
                  <span>·</span>
                  <span>
                    Cta Débito COFINS: <strong>{block.ctaDebitoCofins}</strong>
                  </span>
                </div>
              </div>

              {/* Pivot + De-Para Table */}
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-300">
                      <th className="py-2 px-3 bg-slate-950 text-white font-semibold whitespace-nowrap">
                        Rótulos de Linha
                      </th>
                      <th className="py-2 px-3 bg-slate-950 text-white font-semibold text-right whitespace-nowrap">
                        Soma de pis ({block.ctaDebitoPis})
                      </th>
                      <th className="py-2 px-3 bg-slate-950 text-white font-semibold text-right whitespace-nowrap">
                        Soma de cofins ({block.ctaDebitoCofins})
                      </th>
                      <th className="py-2 px-3 bg-slate-100 text-slate-700 font-semibold whitespace-nowrap">
                        Cta Lançamento
                      </th>
                      <th className="py-2 px-3 bg-slate-100 text-slate-700 font-semibold whitespace-nowrap">
                        Cta Crédito
                      </th>
                      <th className="py-2 px-3 bg-slate-100 text-slate-700 font-semibold whitespace-nowrap">
                        Histórico PIS
                      </th>
                      <th className="py-2 px-3 bg-slate-100 text-slate-700 font-semibold whitespace-nowrap">
                        Histórico COFINS
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {block.filialGroups.map((fg) => (
                      <React.Fragment key={fg.filial}>
                        {/* Filial Subtotal Row */}
                        <tr className="bg-blue-100/80 font-bold text-slate-900">
                          <td className="py-1.5 px-3 font-mono tabular-nums">
                            {fg.filial}
                          </td>
                          <td className="py-1.5 px-3 text-right font-mono tabular-nums">
                            {formatCurrency(fg.subtotalPis)}
                          </td>
                          <td className="py-1.5 px-3 text-right font-mono tabular-nums">
                            {formatCurrency(fg.subtotalCofins)}
                          </td>
                          <td colSpan={4} className="bg-white" />
                        </tr>

                        {/* Account Rows */}
                        {fg.accounts.map((acc) => (
                          <tr
                            key={`${fg.filial}-${acc.rowLabel}-${acc.ctaLancamento}`}
                            className="hover:bg-slate-50"
                          >
                            <td className="py-1.5 px-3 pl-6 font-mono tabular-nums text-slate-900 whitespace-nowrap">
                              {acc.rowLabel}
                            </td>
                            <td className="py-1.5 px-3 text-right font-mono tabular-nums text-slate-900 whitespace-nowrap">
                              {formatCurrency(acc.somaPis)}
                            </td>
                            <td className="py-1.5 px-3 text-right font-mono tabular-nums text-slate-900 whitespace-nowrap">
                              {formatCurrency(acc.somaCofins)}
                            </td>
                            <td className="py-1.5 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                              {acc.ctaLancamento}
                            </td>
                            <td
                              className={`py-1.5 px-3 font-mono tabular-nums whitespace-nowrap ${
                                acc.isMapped
                                  ? 'text-slate-600'
                                  : 'text-red-700 font-semibold'
                              }`}
                            >
                              {acc.ctaCredito}
                            </td>
                            <td className="py-1.5 px-3 font-mono text-[11px] text-slate-600 whitespace-nowrap">
                              {acc.historicoPis}
                            </td>
                            <td className="py-1.5 px-3 font-mono text-[11px] text-slate-600 whitespace-nowrap">
                              {acc.historicoCofins}
                            </td>
                          </tr>
                        ))}
                      </React.Fragment>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-slate-950 text-white font-bold">
                      <td className="py-2 px-3">Total Geral</td>
                      <td className="py-2 px-3 text-right font-mono tabular-nums">
                        {formatCurrency(block.totalPis)}
                      </td>
                      <td className="py-2 px-3 text-right font-mono tabular-nums">
                        {formatCurrency(block.totalCofins)}
                      </td>
                      <td colSpan={4} className="bg-white" />
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          ))}
        </div>
      )}
    </section>
  );
};

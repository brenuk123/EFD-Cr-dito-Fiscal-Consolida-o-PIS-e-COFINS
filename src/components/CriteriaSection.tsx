import React, { useState } from 'react';
import {
  EFD_TABS,
  ParsedWorkbook,
  SavedProfile,
  STD_SECTION_DEFAULT,
  TabConfigMap,
  TabSectionConfig,
} from '../types/efd';
import { isBlueHex } from '../utils/excelParser';
import {
  Check,
  Plus,
  RotateCcw,
  Save,
  Sliders,
  Trash2,
} from 'lucide-react';

interface CriteriaSectionProps {
  workbook: ParsedWorkbook | null;
  blueColors: Set<string>;
  onToggleBlueColor: (hex: string) => void;
  onResetAutoBlueColors: () => void;
  toleranceInput: string;
  onChangeToleranceInput: (val: string) => void;
  ignoreHidden: boolean;
  onChangeIgnoreHidden: (val: boolean) => void;
  tabConfig: TabConfigMap;
  onChangeTabConfig: (next: TabConfigMap) => void;
  onResetDefaultConfig: () => void;
  profiles: SavedProfile[];
  activeProfileId: string;
  onSelectProfile: (profile: SavedProfile) => void;
  onSaveProfile: (name: string) => void;
  onDeleteProfile: (id: string) => void;
  onRunConsolidation: () => void;
}

const SECTION_FIELDS: Array<{
  key: keyof Omit<TabSectionConfig, 'id'>;
  label: string;
  wide?: boolean;
  placeholder?: string;
}> = [
  { key: 'h', label: 'Linha cabeçalho', placeholder: '3' },
  { key: 'a', label: '1ª linha dados', placeholder: '4' },
  { key: 'b', label: 'Última linha', placeholder: 'Fim' },
  { key: 'c1', label: 'Col. de', placeholder: 'A' },
  { key: 'c2', label: 'Col. até', placeholder: 'Auto' },
  { key: 'lab', label: 'Linha rótulo', placeholder: '—' },
  {
    key: 'man',
    label: 'Cabeçalhos manuais (sep. por ;)',
    wide: true,
    placeholder: 'Conta;Descrição;Valor;Pis;Cofins',
  },
  { key: 'base', label: 'Campo Base de Cálculo', wide: true },
  { key: 'pis', label: 'Campo PIS', wide: true },
  { key: 'cof', label: 'Campo COFINS', wide: true },
  { key: 'rp', label: 'Alíq. PIS (%)' },
  { key: 'rc', label: 'Alíq. COFINS (%)' },
];

export const CriteriaSection: React.FC<CriteriaSectionProps> = ({
  workbook,
  blueColors,
  onToggleBlueColor,
  onResetAutoBlueColors,
  toleranceInput,
  onChangeToleranceInput,
  ignoreHidden,
  onChangeIgnoreHidden,
  tabConfig,
  onChangeTabConfig,
  onResetDefaultConfig,
  profiles,
  activeProfileId,
  onSelectProfile,
  onSaveProfile,
  onDeleteProfile,
  onRunConsolidation,
}) => {
  const [newProfileName, setNewProfileName] = useState('');
  const [showNewProfileInput, setShowNewProfileInput] = useState(false);

  const sortedColorEntries = workbook
    ? Object.entries(workbook.colorCounts).sort((a, b) => b[1] - a[1])
    : [];

  const handleFieldChange = (
    tab: string,
    secIndex: number,
    field: keyof Omit<TabSectionConfig, 'id'>,
    value: string
  ) => {
    const next = { ...tabConfig };
    const list = [...(next[tab] || [])];
    list[secIndex] = {
      ...list[secIndex],
      [field]: value,
    };
    next[tab] = list;
    onChangeTabConfig(next);
  };

  const handleAddSection = (tab: string) => {
    const next = { ...tabConfig };
    const list = [...(next[tab] || [])];
    list.push({
      id: `${tab}-${Date.now()}`,
      ...STD_SECTION_DEFAULT,
    });
    next[tab] = list;
    onChangeTabConfig(next);
  };

  const handleRemoveSection = (tab: string, secIndex: number) => {
    const next = { ...tabConfig };
    const list = [...(next[tab] || [])];
    if (list.length <= 1) return;
    list.splice(secIndex, 1);
    next[tab] = list;
    onChangeTabConfig(next);
  };

  const handleCreateProfile = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newProfileName.trim()) return;
    onSaveProfile(newProfileName.trim());
    setNewProfileName('');
    setShowNewProfileInput(false);
  };

  return (
    <section className="bg-white border border-slate-200 rounded-lg p-6 space-y-6">
      {/* Header & Profile Bar */}
      <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 pb-5 border-b border-slate-200">
        <div>
          <h2 className="text-lg font-bold text-slate-900 tracking-tight">
            02. Critérios de Identificação e Parametrização de Abas
          </h2>
          <p className="text-sm text-slate-600 mt-1">
            Defina quais cores de preenchimento são consideradas azuis e ajuste o mapeamento de colunas e alíquotas para cada aba (especialmente <span className="font-medium text-slate-800">04 e 05 Outros</span> e <span className="font-medium text-slate-800">Resumo</span>).
          </p>
        </div>

        {/* Saved Profiles Control */}
        <div className="flex flex-wrap items-center gap-2">
          <label className="text-xs font-medium text-slate-600 mr-1">
            Perfil de Regras:
          </label>
          <select
            value={activeProfileId}
            onChange={(e) => {
              const found = profiles.find((p) => p.id === e.target.value);
              if (found) onSelectProfile(found);
            }}
            className="text-xs font-medium bg-slate-50 border border-slate-300 rounded-md px-3 py-2 text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-800"
          >
            {profiles.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
              </option>
            ))}
          </select>

          {!showNewProfileInput ? (
            <button
              type="button"
              onClick={() => setShowNewProfileInput(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Save className="w-3.5 h-3.5" />
              Salvar Perfil
            </button>
          ) : (
            <form onSubmit={handleCreateProfile} className="flex items-center gap-1.5">
              <input
                type="text"
                value={newProfileName}
                onChange={(e) => setNewProfileName(e.target.value)}
                placeholder="Nome do perfil..."
                className="text-xs border border-slate-300 rounded-md px-2.5 py-1.5 w-40 focus:outline-none focus:ring-2 focus:ring-blue-800"
                autoFocus
              />
              <button
                type="submit"
                className="px-2.5 py-1.5 text-xs font-medium text-white bg-blue-900 rounded-md hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                Confirmar
              </button>
              <button
                type="button"
                onClick={() => setShowNewProfileInput(false)}
                className="px-2 py-1.5 text-xs text-slate-600 hover:text-slate-900 cursor-pointer"
              >
                Cancelar
              </button>
            </form>
          )}

          {activeProfileId !== 'default-efd-8-2026' && (
            <button
              type="button"
              onClick={() => onDeleteProfile(activeProfileId)}
              className="inline-flex items-center gap-1 px-2.5 py-2 text-xs font-medium text-red-700 bg-white border border-slate-200 rounded-md hover:bg-red-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              title="Excluir perfil selecionado"
            >
              <Trash2 className="w-3.5 h-3.5" />
              Excluir
            </button>
          )}

          <button
            type="button"
            onClick={onResetDefaultConfig}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            title="Restaurar configuração padrão EFD 8_2026"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Padrão EFD 8_2026
          </button>
        </div>
      </div>

      {/* Color Selection Palette */}
      <div className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div>
            <h3 className="text-sm font-semibold text-slate-900">
              Paleta de Preenchimento Encontrada no Arquivo
            </h3>
            <p className="text-xs text-slate-600">
              Clique nas amostras abaixo para ativar ou desativar quais cores identificam linhas de crédito de PIS e COFINS.
            </p>
          </div>
          {sortedColorEntries.length > 0 && (
            <button
              type="button"
              onClick={onResetAutoBlueColors}
              className="text-xs font-medium text-blue-900 hover:underline cursor-pointer whitespace-nowrap shrink-0"
            >
              Redefinir para tons azuis automáticos (HSL 190°–260°)
            </button>
          )}
        </div>

        {sortedColorEntries.length === 0 ? (
          <div className="p-4 bg-slate-50 border border-slate-200 rounded-md text-xs text-slate-600">
            {workbook
              ? 'Nenhuma cor de preenchimento sólido foi encontrada nas células das 8 abas.'
              : 'Importe uma planilha Excel (.xlsx) acima para inspecionar automaticamente as cores de preenchimento das células.'}
          </div>
        ) : (
          <div className="flex flex-wrap gap-2 pt-1">
            {sortedColorEntries.map(([hex, count]) => {
              const isSelected = blueColors.has(hex);
              const autoBlue = isBlueHex(hex);
              return (
                <button
                  key={hex}
                  type="button"
                  onClick={() => onToggleBlueColor(hex)}
                  className={`inline-flex items-center gap-2.5 px-3 py-2 rounded-md border text-xs transition-all cursor-pointer whitespace-nowrap shrink-0 ${
                    isSelected
                      ? 'border-blue-900 bg-blue-50/80 text-slate-900 font-semibold shadow-xs'
                      : 'border-slate-200 bg-white text-slate-600 hover:border-slate-300 hover:text-slate-900'
                  }`}
                >
                  <span
                    className="w-4 h-4 rounded-xs border border-slate-400 shrink-0"
                    style={{ backgroundColor: `#${hex}` }}
                  />
                  <span className="font-mono tabular-nums">#{hex}</span>
                  <span aria-hidden="true" className="text-slate-400">
                    ·
                  </span>
                  <span className="font-mono tabular-nums text-slate-500">
                    {count.toLocaleString('pt-BR')} cél.
                  </span>
                  {autoBlue && (
                    <span className="text-[11px] text-blue-800 font-normal">
                      (Azul)
                    </span>
                  )}
                  {isSelected && (
                    <Check className="w-3.5 h-3.5 text-blue-900 shrink-0" />
                  )}
                </button>
              );
            })}
          </div>
        )}
      </div>

      {/* Global Parameters Bar */}
      <div className="flex flex-wrap items-center justify-between gap-6 pt-4 border-t border-slate-200">
        <div className="flex flex-wrap items-center gap-6">
          <label className="flex items-center gap-2.5 text-sm text-slate-700">
            <span className="font-medium">Tolerância de conferência (R$):</span>
            <input
              type="text"
              value={toleranceInput}
              onChange={(e) => onChangeToleranceInput(e.target.value)}
              className="w-24 font-mono tabular-nums text-sm px-2.5 py-1.5 border border-slate-300 rounded-md bg-white text-right focus:outline-none focus:ring-2 focus:ring-blue-800"
            />
          </label>

          <label className="inline-flex items-center gap-2 text-sm text-slate-700 cursor-pointer select-none">
            <input
              type="checkbox"
              checked={ignoreHidden}
              onChange={(e) => onChangeIgnoreHidden(e.target.checked)}
              className="w-4 h-4 rounded border-slate-300 text-blue-900 focus:ring-blue-800"
            />
            <span>Ignorar linhas ocultas no Excel</span>
          </label>
        </div>

        <button
          type="button"
          onClick={onRunConsolidation}
          className="inline-flex items-center gap-2 px-5 py-2.5 text-xs font-semibold text-white bg-blue-900 rounded-md hover:bg-blue-800 transition-colors cursor-pointer whitespace-nowrap shrink-0"
        >
          <Sliders className="w-4 h-4" />
          Recalcular e Consolidar Agora
        </button>
      </div>

      {/* Section Rules per Tab */}
      <div className="pt-2">
        <div className="mb-3">
          <h3 className="text-sm font-semibold text-slate-900">
            Mapeamento de Seções, Colunas e Alíquotas por Aba
          </h3>
          <p className="text-xs text-slate-600 mt-0.5">
            Uma linha é capturada se qualquer célula entre <span className="font-mono">Col. de</span> e <span className="font-mono">Col. até</span> possuir cor marcada como azul. As abas <span className="font-semibold">04 e 05 Outros</span> e <span className="font-semibold">Resumo</span> possuem múltiplas seções independentes.
          </p>
        </div>

        <div className="overflow-x-auto border border-slate-200 rounded-md">
          <table className="w-full border-collapse text-left text-xs">
            <thead>
              <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                <th className="py-2.5 px-3 whitespace-nowrap">Aba</th>
                {SECTION_FIELDS.map((f) => (
                  <th key={f.key} className="py-2.5 px-2 whitespace-nowrap">
                    {f.label}
                  </th>
                ))}
                <th className="py-2.5 px-3 text-right whitespace-nowrap">
                  Seções
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {EFD_TABS.map((tab) => {
                const sections = tabConfig[tab] || [];
                const isSheetMissing = workbook && !workbook.sheets[tab];
                return sections.map((sec, idx) => (
                  <tr
                    key={sec.id || `${tab}-${idx}`}
                    className={`hover:bg-slate-50/80 ${
                      isSheetMissing ? 'opacity-60 bg-slate-50/40' : ''
                    }`}
                  >
                    <td className="py-2 px-3 font-medium text-slate-900 whitespace-nowrap">
                      {idx === 0 ? (
                        <div className="flex items-center gap-1.5">
                          <span>{tab}</span>
                          {sections.length > 1 && (
                            <span className="text-[11px] font-mono tabular-nums text-slate-500">
                              · {sections.length} seções
                            </span>
                          )}
                        </div>
                      ) : (
                        <span className="text-slate-400 pl-3">↳ Seção {idx + 1}</span>
                      )}
                    </td>

                    {SECTION_FIELDS.map((field) => {
                      const isNumeric = ['h', 'a', 'b', 'lab', 'rp', 'rc'].includes(
                        field.key
                      );
                      return (
                        <td key={field.key} className="py-1.5 px-1.5">
                          <input
                            type="text"
                            value={String(sec[field.key] ?? '')}
                            placeholder={field.placeholder}
                            onChange={(e) =>
                              handleFieldChange(
                                tab,
                                idx,
                                field.key,
                                e.target.value
                              )
                            }
                            className={`px-2 py-1 text-xs border border-slate-200 rounded bg-white text-slate-900 focus:outline-none focus:border-blue-800 focus:ring-1 focus:ring-blue-800 ${
                              field.wide ? 'w-36' : 'w-14'
                            } ${isNumeric ? 'font-mono tabular-nums text-right' : ''}`}
                          />
                        </td>
                      );
                    })}

                    <td className="py-1.5 px-3 text-right whitespace-nowrap">
                      <div className="inline-flex items-center gap-1">
                        {sections.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleRemoveSection(tab, idx)}
                            className="px-2 py-1 text-xs font-medium text-red-700 hover:bg-red-50 rounded border border-transparent hover:border-red-200 transition-colors cursor-pointer"
                            title="Remover esta seção"
                          >
                            ×
                          </button>
                        )}
                        {idx === sections.length - 1 && (
                          <button
                            type="button"
                            onClick={() => handleAddSection(tab)}
                            className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-blue-900 bg-slate-100 hover:bg-blue-50 rounded transition-colors cursor-pointer whitespace-nowrap shrink-0"
                            title="Adicionar nova seção de leitura nesta aba"
                          >
                            <Plus className="w-3 h-3" />
                            seção
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                ));
              })}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
};

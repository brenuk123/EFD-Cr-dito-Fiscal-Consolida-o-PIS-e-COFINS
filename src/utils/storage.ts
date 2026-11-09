import {
  ClosingHistoryEntry,
  SavedProfile,
  createDefaultTabConfig,
} from '../types/efd';

const PROFILES_KEY = 'efd_pis_cofins_profiles_v1';
const HISTORY_KEY = 'efd_pis_cofins_history_v1';

export function loadSavedProfiles(): SavedProfile[] {
  try {
    const raw = localStorage.getItem(PROFILES_KEY);
    if (!raw) return getBuiltInProfiles();
    const parsed = JSON.parse(raw) as SavedProfile[];
    return Array.isArray(parsed) && parsed.length > 0
      ? parsed
      : getBuiltInProfiles();
  } catch {
    return getBuiltInProfiles();
  }
}

export function saveProfilesToStorage(profiles: SavedProfile[]): void {
  try {
    localStorage.setItem(PROFILES_KEY, JSON.stringify(profiles));
  } catch (err) {
    console.error('Erro ao salvar perfis:', err);
  }
}

export function getBuiltInProfiles(): SavedProfile[] {
  return [
    {
      id: 'default-efd-8-2026',
      name: 'Padrão EFD 8_2026 (Original)',
      updatedAt: new Date().toISOString(),
      tolerance: 0.02,
      ignoreHidden: false,
      blueColors: ['BDD7EE', '8EA9DB', '9BC2E6', '4472C4', '5B9BD5', 'D9E1F2'],
      tabConfig: createDefaultTabConfig(),
    },
  ];
}

export function loadClosingHistory(): ClosingHistoryEntry[] {
  try {
    const raw = localStorage.getItem(HISTORY_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as ClosingHistoryEntry[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function saveClosingHistoryToStorage(entries: ClosingHistoryEntry[]): void {
  try {
    localStorage.setItem(HISTORY_KEY, JSON.stringify(entries));
  } catch (err) {
    console.error('Erro ao salvar histórico mensal:', err);
  }
}

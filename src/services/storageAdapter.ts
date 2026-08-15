import { AppData, emptyAppData } from "@/types/appData";

export const STORAGE_KEY = "SISTEMA_HORAS_EXTRAS_DATA";

/**
 * Backend de persistência abstraído.
 * Hoje: localStorage. Futuro: File System local via Electron (IPC),
 * bastando implementar outro Backend e trocar `activeBackend`.
 */
interface StorageBackend {
  read(): Promise<string | null>;
  write(payload: string): Promise<void>;
  writeSync(payload: string): void;
}

const localStorageBackend: StorageBackend = {
  async read() {
    try {
      return localStorage.getItem(STORAGE_KEY);
    } catch {
      return null;
    }
  },
  async write(payload) {
    localStorage.setItem(STORAGE_KEY, payload);
  },
  writeSync(payload) {
    try {
      localStorage.setItem(STORAGE_KEY, payload);
    } catch {
      /* ignora falhas em contexto de fechamento */
    }
  },
};

// Detecta ponte Electron (quando existir), senão usa localStorage.
const electronBridge = (globalThis as any)?.electronAPI;
const electronBackend: StorageBackend | null = electronBridge
  ? {
      read: () => electronBridge.readAppData(),
      write: (payload: string) => electronBridge.writeAppData(payload),
      writeSync: (payload: string) => {
        electronBridge.writeAppDataSync?.(payload) ?? localStorageBackend.writeSync(payload);
      },
    }
  : null;

const activeBackend: StorageBackend = electronBackend ?? localStorageBackend;

const normalize = (raw: unknown): AppData => {
  const base = emptyAppData();
  if (!raw || typeof raw !== "object") return base;
  const d = raw as Partial<AppData>;
  return {
    lancamentos: Array.isArray(d.lancamentos) ? d.lancamentos : [],
    relatorios: Array.isArray(d.relatorios) ? d.relatorios : [],
    configuracoes: d.configuracoes && typeof d.configuracoes === "object" ? d.configuracoes : {},
    ultimaAtualizacao: typeof d.ultimaAtualizacao === "string" ? d.ultimaAtualizacao : base.ultimaAtualizacao,
  };
};

export async function loadAppData(): Promise<AppData> {
  const raw = await activeBackend.read();
  if (!raw) return emptyAppData();
  try {
    return normalize(JSON.parse(raw));
  } catch {
    return emptyAppData();
  }
}

export async function saveAppData(data: AppData): Promise<AppData> {
  const payload: AppData = { ...data, ultimaAtualizacao: new Date().toISOString() };
  await activeBackend.write(JSON.stringify(payload));
  return payload;
}

/** Salvamento de emergência (fechamento de aba / troca de visibilidade). */
export function saveAppDataSync(data: AppData): AppData {
  const payload: AppData = { ...data, ultimaAtualizacao: new Date().toISOString() };
  activeBackend.writeSync(JSON.stringify(payload));
  return payload;
}

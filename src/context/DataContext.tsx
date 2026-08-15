import React, { createContext, useCallback, useContext, useEffect, useRef, useState } from "react";
import { AppData, Lancamento, Relatorio, emptyAppData } from "@/types/appData";
import { loadAppData, saveAppData, saveAppDataSync } from "@/services/storageAdapter";

interface DataContextValue {
  data: AppData;
  isLoaded: boolean;
  isSaving: boolean;
  hasUnsavedChanges: boolean;
  salvarDados: () => Promise<void>;
  adicionarLancamento: (l: Omit<Lancamento, "id" | "timestamp"> & Partial<Pick<Lancamento, "id" | "timestamp">>) => Lancamento;
  atualizarLancamento: (id: string, patch: Partial<Lancamento>) => void;
  removerLancamento: (id: string) => void;
  salvarRelatorio: (r: Omit<Relatorio, "id"> & Partial<Pick<Relatorio, "id">>) => Relatorio;
  atualizarConfiguracao: (chave: string, valor: any) => void;
}

const DataContext = createContext<DataContextValue | null>(null);

const novoId = () =>
  typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`;

export const DataProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [data, setData] = useState<AppData>(emptyAppData);
  const [isLoaded, setIsLoaded] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);

  // Refs para uso dentro de listeners globais (sempre com o valor atual).
  const dataRef = useRef(data);
  const dirtyRef = useRef(hasUnsavedChanges);
  dataRef.current = data;
  dirtyRef.current = hasUnsavedChanges;

  useEffect(() => {
    let cancelled = false;
    loadAppData().then((loaded) => {
      if (cancelled) return;
      setData(loaded);
      setIsLoaded(true);
    });
    return () => {
      cancelled = true;
    };
  }, []);

  const mutate = useCallback((fn: (prev: AppData) => AppData) => {
    setData((prev) => fn(prev));
    setHasUnsavedChanges(true);
  }, []);

  const salvarDados = useCallback(async () => {
    setIsSaving(true);
    try {
      const saved = await saveAppData(dataRef.current);
      setData((prev) => ({ ...prev, ultimaAtualizacao: saved.ultimaAtualizacao }));
      setHasUnsavedChanges(false);
    } finally {
      setIsSaving(false);
    }
  }, []);

  const adicionarLancamento: DataContextValue["adicionarLancamento"] = useCallback(
    (entrada) => {
      const lancamento: Lancamento = {
        id: entrada.id ?? novoId(),
        timestamp: entrada.timestamp ?? Date.now(),
        data: entrada.data,
        servidor: entrada.servidor,
        horasExtras: entrada.horasExtras,
        observacao: entrada.observacao,
        status: entrada.status,
      };
      mutate((prev) => ({ ...prev, lancamentos: [lancamento, ...prev.lancamentos] }));
      return lancamento;
    },
    [mutate]
  );

  const atualizarLancamento = useCallback(
    (id: string, patch: Partial<Lancamento>) => {
      mutate((prev) => ({
        ...prev,
        lancamentos: prev.lancamentos.map((l) => (l.id === id ? { ...l, ...patch, id: l.id } : l)),
      }));
    },
    [mutate]
  );

  const removerLancamento = useCallback(
    (id: string) => {
      mutate((prev) => ({ ...prev, lancamentos: prev.lancamentos.filter((l) => l.id !== id) }));
    },
    [mutate]
  );

  const salvarRelatorio: DataContextValue["salvarRelatorio"] = useCallback(
    (entrada) => {
      const relatorio: Relatorio = { ...entrada, id: entrada.id ?? novoId() };
      mutate((prev) => ({
        ...prev,
        relatorios: [relatorio, ...prev.relatorios.filter((r) => r.id !== relatorio.id)],
      }));
      return relatorio;
    },
    [mutate]
  );

  const atualizarConfiguracao = useCallback(
    (chave: string, valor: any) => {
      mutate((prev) => ({ ...prev, configuracoes: { ...prev.configuracoes, [chave]: valor } }));
    },
    [mutate]
  );

  // Auto-save de emergência + atalho Ctrl/Cmd+S
  useEffect(() => {
    const emergencia = () => {
      if (!dirtyRef.current) return;
      saveAppDataSync(dataRef.current);
      dirtyRef.current = false;
      setHasUnsavedChanges(false);
    };

    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (!dirtyRef.current) return;
      emergencia();
      e.preventDefault();
      e.returnValue = "";
    };

    const onVisibility = () => {
      if (document.visibilityState === "hidden") emergencia();
    };

    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
        e.preventDefault();
        void salvarDados();
      }
    };

    window.addEventListener("beforeunload", onBeforeUnload);
    window.addEventListener("visibilitychange", onVisibility);
    document.addEventListener("visibilitychange", onVisibility);
    window.addEventListener("pagehide", emergencia);
    window.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener("beforeunload", onBeforeUnload);
      window.removeEventListener("visibilitychange", onVisibility);
      document.removeEventListener("visibilitychange", onVisibility);
      window.removeEventListener("pagehide", emergencia);
      window.removeEventListener("keydown", onKeyDown);
    };
  }, [salvarDados]);

  return (
    <DataContext.Provider
      value={{
        data,
        isLoaded,
        isSaving,
        hasUnsavedChanges,
        salvarDados,
        adicionarLancamento,
        atualizarLancamento,
        removerLancamento,
        salvarRelatorio,
        atualizarConfiguracao,
      }}
    >
      {children}
    </DataContext.Provider>
  );
};

export const useData = () => {
  const ctx = useContext(DataContext);
  if (!ctx) throw new Error("useData deve ser usado dentro de <DataProvider>");
  return ctx;
};

export interface Lancamento {
  id: string;
  data: string;
  servidor: string;
  horasExtras: number;
  observacao?: string;
  status: string;
  timestamp: number;
}

export interface Relatorio {
  id: string;
  titulo: string;
  dataGeracao: string;
  arquivoNome: string;
  conteudo: any;
}

export interface AppData {
  lancamentos: Lancamento[];
  relatorios: Relatorio[];
  configuracoes: Record<string, any>;
  ultimaAtualizacao: string;
}

export const emptyAppData = (): AppData => ({
  lancamentos: [],
  relatorios: [],
  configuracoes: {},
  ultimaAtualizacao: new Date().toISOString(),
});

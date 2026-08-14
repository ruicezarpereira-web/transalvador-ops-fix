import { useEffect, useMemo, useState } from "react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { toast } from "@/hooks/use-toast";
import ServerCombobox, { Servidor } from "@/components/ServerCombobox";
import MultiServerSelect from "@/components/MultiServerSelect";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { drawHeader, drawInfoBox, drawSignatures, drawFooters, tableTheme } from "@/lib/pdf";

// Tipos
type FuncaoID =
  | "coordenador_geral" | "coordenador_setorial" | "supervisor" | "agente_fiscalizacao"
  | "guarda_civil" | "agente_operacoes" | "assistente_tecnico" | "motorista"
  | "coordenador" | "supervisor1" | "supervisor2" | "apoio_adm";

type DiaTrabalho = { data: string; horas: number; funcao: FuncaoID };

type OperacaoTipo = "ordinaria" | "reveillon" | "carnaval";

type Lancamento = {
  id: string;
  operacao: OperacaoTipo;
  nomeOperacao: string;
  periodo: { inicio: string; fim: string };
  servidor: Servidor;
  dias: DiaTrabalho[];
  createdAt: string;
};

type ValoresOperacao = { [key in FuncaoID]?: number };

const valoresDefault = {
  ordinaria: {
    coordenador: 20.5,
    supervisor: 15.5,
    agente_fiscalizacao: 12.0,
    apoio_adm: 10.0,
  } as ValoresOperacao,
  reveillon: {
    coordenador: 26.22,
    supervisor1: 25.07,
    supervisor2: 23.85,
    agente_fiscalizacao: 22.79,
    apoio_adm: 10.0,
  } as ValoresOperacao,
  carnaval: {
    coordenador_geral: 45.01,
    coordenador_setorial: 33.76,
    supervisor: 33.76,
    agente_fiscalizacao: 30.01,
    guarda_civil: 30.01,
    agente_operacoes: 21.93,
    assistente_tecnico: 21.93,
    motorista: 21.93,
  } as ValoresOperacao,
};

const alimentacaoDefault = {
  ordinaria: { proporcional: true, valorHora: 2.0, minimoHoras: 8 },
  reveillon: { 12: 13.68, proporcional: true },
  carnaval: {
    1: 3.05, 2: 6.09, 3: 9.14, 4: 12.18, 5: 15.23, 6: 18.27, 7: 21.32, 8: 35.92,
    9: 39.01, 10: 42.1, 11: 43.08, 12: 47.0, 13: 50.05, 14: 53.1, 15: 56.15,
    16: 57.13, 17: 58.11, 18: 59.08, 19: 59.09, 20: 62.14, 21: 65.19, 22: 68.24,
    23: 71.29, 24: 74.64,
  },
} as any;

type OpItem = { id: string; label: string; inicio: string; fim: string; tipo: OperacaoTipo };

// Utils de datas
const pad2 = (n: number) => n.toString().padStart(2, "0");
const toISO = (d: Date) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const toBR = (iso: string) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const fmtBRL = (n: number) => new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(n);
const clampToPeriodo = (dateIso: string, inicio: string, fim: string) =>
  dateIso < inicio ? inicio : dateIso > fim ? fim : dateIso;

const easterSunday = (year: number) => {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
};
const addDays = (d: Date, delta: number) => {
  const nd = new Date(d);
  nd.setDate(nd.getDate() + delta);
  return nd;
};
const carnavalPeriodo = (year: number) => {
  const pascoa = easterSunday(year);
  const terca = addDays(pascoa, -47);
  const sexta = addDays(terca, -4);
  return { inicio: toISO(sexta), fim: toISO(terca) };
};

const buildOpcoes = (year: number): OpItem[] => {
  const ops: OpItem[] = [];
  ops.push({
    id: `ordinaria-01-20-jan-${year}`,
    label: `01/01 a 20/01/${year}`,
    inicio: `${year}-01-01`,
    fim: `${year}-01-20`,
    tipo: "ordinaria",
  });
  for (let m = 0; m <= 10; m++) {
    const start = new Date(year, m, 21);
    const end = new Date(year, m + 1, 20);
    ops.push({
      id: `ordinaria-21-${pad2(m + 1)}-20-${pad2(m + 2)}-${year}`,
      label: `21/${pad2(m + 1)} a 20/${pad2(m + 2)}/${year}`,
      inicio: toISO(start),
      fim: toISO(end),
      tipo: "ordinaria",
    });
  }
  ops.push({
    id: `ordinaria-21-31-dez-${year}`,
    label: `21/12 a 31/12/${year}`,
    inicio: `${year}-12-21`,
    fim: `${year}-12-31`,
    tipo: "ordinaria",
  });
  ops.push({
    id: `reveillon-${year}`,
    label: `Reveillon ${year}`,
    inicio: `${year}-12-24`,
    fim: `${year + 1}-01-01`,
    tipo: "reveillon",
  });
  const car = carnavalPeriodo(year);
  ops.push({ id: `carnaval-${year}`, label: `Carnaval ${year}`, inicio: car.inicio, fim: car.fim, tipo: "carnaval" });
  return ops;
};

// Persistência
const load = <T,>(key: string, fallback: T): T => {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
};
const save = (key: string, value: unknown) => localStorage.setItem(key, JSON.stringify(value));

// Acesso único (usuário único + senha master)
const USUARIO_MASTER = "RCPPJ";
const SENHA_MASTER_PADRAO = "ruicpj@123";
const SECURITY_QUESTION = "My birthday?";
const SECURITY_ANSWER = "27 de Setembro";

// Normalização de cabeçalhos do Excel
const normalize = (s: any) =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]/g, "")
    .toUpperCase();

const excelDateToISO = (v: any): string => {
  if (v === undefined || v === null || v === "") return "";
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${pad2(d.m)}-${pad2(d.d)}`;
    return "";
  }
  const s = String(v).trim();
  const br = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$/);
  if (br) {
    const y = br[3].length === 2 ? `20${br[3]}` : br[3];
    return `${y}-${pad2(Number(br[2]))}-${pad2(Number(br[1]))}`;
  }
  return s;
};

const Index = () => {
  useEffect(() => {
    document.title = "GEOPS - Gerador de Operações Especiais Segep";
  }, []);

  const [servidores, setServidores] = useState<Servidor[]>(() => load<Servidor[]>("servidores", []));
  const [lancamentos, setLancamentos] = useState<Lancamento[]>(() => load<Lancamento[]>("lancamentos", []));
  const [editingId, setEditingId] = useState<string | null>(null);
  const [tab, setTab] = useState<string>("lancamentos");

  // Estado de Lançamento
  const initialYear = new Date().getFullYear();
  const initialOps = buildOpcoes(initialYear);
  const [ano, setAno] = useState<number>(initialYear);
  const [operacaoId, setOperacaoId] = useState<string>(initialOps[0]?.id ?? "");
  const [periodo, setPeriodo] = useState({ inicio: initialOps[0]?.inicio ?? "", fim: initialOps[0]?.fim ?? "" });
  const [periodoCustomizado, setPeriodoCustomizado] = useState({ inicio: "", fim: "" });
  const [matriculaSelecionada, setMatriculaSelecionada] = useState<string | undefined>(undefined);
  const [dias, setDias] = useState<DiaTrabalho[]>([{ data: initialOps[0]?.inicio ?? "", horas: 8, funcao: "coordenador" }]);

  // Lançamento por data (lote)
  const [loteDatas, setLoteDatas] = useState<string[]>([]);
  const [loteNovaData, setLoteNovaData] = useState<string>("");
  const [loteHoras, setLoteHoras] = useState<number>(8);
  const [loteFuncao, setLoteFuncao] = useState<FuncaoID>("agente_fiscalizacao");
  const [loteMatriculas, setLoteMatriculas] = useState<string[]>([]);
  const [loteConflitos, setLoteConflitos] = useState<string[]>([]);

  // Configurações
  const [senhaMaster, setSenhaMaster] = useState<string>(() => load<string>("senhaMaster", SENHA_MASTER_PADRAO));
  const [contatos, setContatos] = useState<{ telefone1: string; telefone2: string }>(() =>
    load("contatosSetor", { telefone1: "", telefone2: "" })
  );
  const [gestor, setGestor] = useState<{ nome: string; cargo: string }>(() =>
    load("gestorSetor", { nome: "", cargo: "Gestor da Operação" })
  );
  const [valores, setValores] = useState(() => load("valoresOperacoes", valoresDefault));
  const [alimentacao, setAlimentacao] = useState(() => load("alimentacaoOperacoes", alimentacaoDefault));

  // Login
  const [loginSenha, setLoginSenha] = useState("");
  const [novaSenha, setNovaSenha] = useState("");
  const [confirmarSenha, setConfirmarSenha] = useState("");
  const [respostaSeguranca, setRespostaSeguranca] = useState("");
  const [novaSenhaSeguranca, setNovaSenhaSeguranca] = useState("");
  const [sistemaDesbloqueado, setSistemaDesbloqueado] = useState<boolean>(() => load<boolean>("sistemaDesbloqueado", false));

  const opcoesOperacao = useMemo(() => buildOpcoes(ano), [ano]);
  const selectedOp = useMemo(() => opcoesOperacao.find((o) => o.id === operacaoId), [opcoesOperacao, operacaoId]);

  const anosDisponiveis = useMemo(() => {
    const cy = new Date().getFullYear();
    const startYear = Math.max(2025, cy);
    return Array.from({ length: 15 }, (_, i) => startYear + i);
  }, []);

  useEffect(() => {
    if (!opcoesOperacao.find((o) => o.id === operacaoId)) {
      const first = opcoesOperacao[0];
      if (first) {
        setOperacaoId(first.id);
        setPeriodo({ inicio: first.inicio, fim: first.fim });
        setDias([{ data: first.inicio, horas: 8, funcao: "coordenador" }]);
      }
    }
  }, [opcoesOperacao, operacaoId]);

  // Importação Excel
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const formatarCPF = (cpf?: string) => {
    if (!cpf) return "";
    const s = cpf.toString().replace(/\D/g, "");
    return s.length === 11 ? s.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") : String(cpf);
  };

  const importarExcel = async () => {
    if (!excelFile) {
      toast({ title: "Selecione um arquivo Excel", variant: "destructive" });
      return;
    }
    try {
      const data = await excelFile.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const matriz: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });

      // Localiza a linha de cabeçalho (a planilha tem título mesclado na linha 1)
      const headerIdx = matriz.findIndex((row) =>
        row.some((c) => normalize(c) === "NOME") && row.some((c) => normalize(c).startsWith("MATRICULA"))
      );
      if (headerIdx < 0) {
        toast({ title: "Cabeçalho não encontrado", description: "A planilha deve conter as colunas NOME e MATRÍCULA.", variant: "destructive" });
        return;
      }

      const header = matriz[headerIdx].map((c) => normalize(c));
      const idx = (...alvos: string[]) => header.findIndex((h) => alvos.some((a) => h === a || h.startsWith(a)));

      const iNome = idx("NOME");
      const iNasc = idx("NASC", "DATADENASCIMENTO");
      const iSexo = idx("SEXO");
      const iMat = idx("MATRICULA");
      const iRU = idx("RUNICO", "REGISTROUNICO");
      const iRG = idx("RG");
      const iCPF = idx("CPF");
      const iAdm = idx("ADMISSAO");
      // Coluna CARGO/FUNÇÃO é intencionalmente ignorada.

      const lista: Servidor[] = [];
      const vistos = new Set<string>();
      for (let r = headerIdx + 1; r < matriz.length; r++) {
        const row = matriz[r] || [];
        const nome = String(row[iNome] ?? "").trim();
        const matricula = String(row[iMat] ?? "").trim();
        if (!nome || !matricula) continue;
        if (vistos.has(matricula)) continue;
        vistos.add(matricula);
        lista.push({
          nome,
          matricula,
          nascimento: iNasc >= 0 ? excelDateToISO(row[iNasc]) : "",
          sexo: iSexo >= 0 ? String(row[iSexo] ?? "").trim() : "",
          registroUnico: iRU >= 0 ? String(row[iRU] ?? "").trim() : "",
          rg: iRG >= 0 ? String(row[iRG] ?? "").trim() : "",
          cpf: iCPF >= 0 ? String(row[iCPF] ?? "").trim() : "",
          admissao: iAdm >= 0 ? excelDateToISO(row[iAdm]) : "",
        });
      }

      if (lista.length === 0) {
        toast({ title: "Nenhum servidor encontrado", variant: "destructive" });
        return;
      }
      lista.sort((a, b) => a.nome.localeCompare(b.nome));
      setServidores(lista);
      toast({ title: `${lista.length} servidores importados` });
    } catch (e) {
      console.error(e);
      toast({ title: "Erro ao importar", variant: "destructive" });
    }
  };

  // Cálculos
  const valorHora = (f: FuncaoID, tipo: OperacaoTipo) => (valores as any)[tipo]?.[f] || 0;

  const calcAlimentacao = (tipo: OperacaoTipo, horas: number) => {
    if (tipo === "ordinaria") {
      const cfg = (alimentacao as any).ordinaria || {};
      const min = cfg.minimoHoras ?? 8;
      const vh = cfg.valorHora ?? 2;
      return horas >= min ? horas * vh : 0;
    }
    if (tipo === "reveillon") {
      const v12 = (alimentacao as any).reveillon?.[12] ?? 13.68;
      return horas >= 12 ? v12 : (v12 * horas) / 12;
    }
    return (alimentacao as any).carnaval?.[horas] || 0;
  };

  const funcaoLabel = (f: FuncaoID) => {
    const labels: Record<FuncaoID, string> = {
      coordenador_geral: "Coordenador Geral",
      coordenador_setorial: "Coordenador Setorial",
      supervisor: "Supervisor",
      agente_fiscalizacao: "Agente de Fiscalização",
      guarda_civil: "Guarda Civil Municipal",
      agente_operacoes: "Agente de Operações",
      assistente_tecnico: "Assistente Técnico Administrativo",
      motorista: "Motorista",
      coordenador: "Coordenador",
      supervisor1: "Supervisor 1",
      supervisor2: "Supervisor 2",
      apoio_adm: "Apoio Administrativo",
    };
    return labels[f] || f;
  };

  const getFuncoesDisponiveisParaOperacao = (tipo: OperacaoTipo): FuncaoID[] => {
    switch (tipo) {
      case "carnaval":
        return ["coordenador_geral", "coordenador_setorial", "supervisor", "agente_fiscalizacao", "guarda_civil", "agente_operacoes", "assistente_tecnico", "motorista"];
      case "reveillon":
        return ["coordenador", "supervisor1", "supervisor2", "agente_fiscalizacao", "apoio_adm"];
      case "ordinaria":
      default:
        return ["coordenador", "supervisor", "agente_fiscalizacao", "apoio_adm"];
    }
  };

  // Mantém a função do lote válida para a operação
  useEffect(() => {
    const disponiveis = getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria");
    if (!disponiveis.includes(loteFuncao)) setLoteFuncao(disponiveis[0]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedOp?.tipo]);

  // Persistência
  useEffect(() => save("servidores", servidores), [servidores]);
  useEffect(() => save("lancamentos", lancamentos), [lancamentos]);
  useEffect(() => save("senhaMaster", senhaMaster), [senhaMaster]);
  useEffect(() => save("contatosSetor", contatos), [contatos]);
  useEffect(() => save("gestorSetor", gestor), [gestor]);
  useEffect(() => save("valoresOperacoes", valores), [valores]);
  useEffect(() => save("alimentacaoOperacoes", alimentacao), [alimentacao]);
  useEffect(() => save("sistemaDesbloqueado", sistemaDesbloqueado), [sistemaDesbloqueado]);

  const realizarLogin = () => {
    if (loginSenha === senhaMaster) {
      setSistemaDesbloqueado(true);
      setLoginSenha("");
      toast({ title: "Acesso liberado", description: `Bem-vindo, ${USUARIO_MASTER}.` });
      return;
    }
    toast({ title: "Senha master inválida", variant: "destructive" });
  };

  const alterarSenha = () => {
    if (novaSenha !== confirmarSenha) {
      toast({ title: "Senhas não coincidem", variant: "destructive" });
      return;
    }
    if (novaSenha.length < 6) {
      toast({ title: "Senha deve ter pelo menos 6 caracteres", variant: "destructive" });
      return;
    }
    save("senhaMaster", novaSenha);
    setSenhaMaster(novaSenha);
    setNovaSenha("");
    setConfirmarSenha("");
    toast({ title: "Senha master alterada" });
  };

  const redefinirSenha = () => {
    if (respostaSeguranca.trim().toLowerCase() !== SECURITY_ANSWER.toLowerCase()) {
      toast({ title: "Resposta de segurança incorreta", variant: "destructive" });
      return;
    }
    if (novaSenhaSeguranca.length < 6) {
      toast({ title: "Nova senha deve ter pelo menos 6 caracteres", variant: "destructive" });
      return;
    }
    save("senhaMaster", novaSenhaSeguranca);
    setSenhaMaster(novaSenhaSeguranca);
    setRespostaSeguranca("");
    setNovaSenhaSeguranca("");
    toast({ title: "Senha master redefinida" });
  };

  const logout = () => {
    setSistemaDesbloqueado(false);
    toast({ title: "Sessão encerrada" });
  };

  const salvarAlteracoes = () => {
    save("valoresOperacoes", valores);
    save("alimentacaoOperacoes", alimentacao);
    save("contatosSetor", contatos);
    save("gestorSetor", gestor);
    toast({ title: "Alterações salvas com sucesso" });
  };

  const limparTudo = () => {
    setServidores([]);
    setLancamentos([]);
    localStorage.removeItem("servidores");
    localStorage.removeItem("lancamentos");
    toast({ title: "Dados limpos" });
  };

  // CRUD de dias (lançamento por servidor)
  const adicionarDia = () => {
    const funcoesDisponiveis = getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria");
    setDias((d) => [...d, { data: periodo.inicio, horas: 8, funcao: funcoesDisponiveis[0] }]);
  };
  const removerDia = (index: number) => setDias((d) => d.filter((_, i) => i !== index));
  const atualizarDia = (index: number, patch: Partial<DiaTrabalho>) =>
    setDias((d) => d.map((dia, i) => (i === index ? { ...dia, ...patch } : dia)));

  const salvarLancamento = async () => {
    if (!matriculaSelecionada) {
      toast({ title: "Selecione um servidor válido", variant: "destructive" });
      return;
    }
    const srv = servidores.find((s) => s.matricula === matriculaSelecionada);
    if (!srv) {
      toast({ title: "Servidor não encontrado", variant: "destructive" });
      return;
    }
    const diasValidos = dias.filter((d) => d.data && d.horas > 0);
    if (diasValidos.length === 0) {
      toast({ title: "Adicione pelo menos um dia", variant: "destructive" });
      return;
    }
    const datas = diasValidos.map((d) => d.data);
    if (new Set(datas).size !== datas.length) {
      toast({ title: "Datas repetidas", description: "Remova dias duplicados no lançamento.", variant: "destructive" });
      return;
    }
    const datasSet = new Set(datas);
    const conflito = lancamentos.some((l) => {
      if (editingId && l.id === editingId) return false;
      if (l.servidor.matricula !== srv.matricula) return false;
      return l.dias.some((d) => datasSet.has(d.data));
    });
    if (conflito) {
      toast({ title: "Conflito de datas", description: "Já existe lançamento para este servidor em uma das datas.", variant: "destructive" });
      return;
    }
    const foraPeriodo = diasValidos.find((d) => d.data < periodo.inicio || d.data > periodo.fim);
    if (foraPeriodo) {
      toast({ title: "Data fora do período", variant: "destructive" });
      return;
    }

    const novo: Lancamento = {
      id: editingId ?? crypto.randomUUID(),
      operacao: selectedOp?.tipo ?? "ordinaria",
      nomeOperacao: selectedOp?.label ?? "",
      periodo,
      servidor: srv,
      dias: diasValidos,
      createdAt: new Date().toISOString(),
    };

    setLancamentos((prev) => (editingId ? prev.map((l) => (l.id === editingId ? novo : l)) : [novo, ...prev]));
    const eraEdicao = !!editingId;
    setEditingId(null);
    setDias([{ data: periodo.inicio, horas: 8, funcao: getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria")[0] }]);
    toast({ title: eraEdicao ? "Lançamento atualizado" : "Frequência registrada no histórico" });
  };

  // Lançamento por data (vários servidores)
  const adicionarLoteData = () => {
    if (!loteNovaData) return;
    if (loteNovaData < periodo.inicio || loteNovaData > periodo.fim) {
      toast({ title: "Data fora do período da operação", variant: "destructive" });
      return;
    }
    if (loteDatas.includes(loteNovaData)) {
      toast({ title: "Data já adicionada", variant: "destructive" });
      return;
    }
    setLoteDatas((d) => [...d, loteNovaData].sort());
    setLoteNovaData("");
  };

  const salvarLancamentoPorData = () => {
    setLoteConflitos([]);
    if (loteDatas.length === 0) {
      toast({ title: "Adicione pelo menos uma data", variant: "destructive" });
      return;
    }
    if (loteMatriculas.length === 0) {
      toast({ title: "Selecione pelo menos um servidor", variant: "destructive" });
      return;
    }
    if (!loteHoras || loteHoras <= 0) {
      toast({ title: "Informe a carga horária", variant: "destructive" });
      return;
    }

    const tipo = selectedOp?.tipo ?? "ordinaria";
    const nomeOperacao = selectedOp?.label ?? "";
    const conflitos: string[] = [];
    const bloqueados: string[] = [];
    let criados = 0;
    let atualizados = 0;

    setLancamentos((prev) => {
      const next = [...prev];
      loteMatriculas.forEach((mat) => {
        const srv = servidores.find((s) => s.matricula === mat);
        if (!srv) return;

        const datasExistentes = new Set(
          next.filter((l) => l.servidor.matricula === mat).flatMap((l) => l.dias.map((d) => d.data))
        );
        const datasConflitantes = loteDatas.filter((data) => datasExistentes.has(data));
        if (datasConflitantes.length > 0) {
          // Servidor com data(s) já lançada(s): NÃO é lançado e o erro é apontado.
          conflitos.push(
            `${srv.nome} (${srv.matricula}) — já possui lançamento em ${datasConflitantes.map(toBR).join(", ")}`
          );
          bloqueados.push(mat);
          return;
        }

        const novosDias: DiaTrabalho[] = loteDatas.map((data) => ({ data, horas: loteHoras, funcao: loteFuncao }));
        const existenteIdx = next.findIndex(
          (l) => l.servidor.matricula === mat && l.nomeOperacao === nomeOperacao && l.periodo.inicio === periodo.inicio && l.periodo.fim === periodo.fim
        );
        if (existenteIdx >= 0) {
          const alvo = next[existenteIdx];
          next[existenteIdx] = {
            ...alvo,
            dias: [...alvo.dias, ...novosDias].sort((a, b) => a.data.localeCompare(b.data)),
          };
          atualizados++;
        } else {
          next.unshift({
            id: crypto.randomUUID(),
            operacao: tipo,
            nomeOperacao,
            periodo,
            servidor: srv,
            dias: novosDias,
            createdAt: new Date().toISOString(),
          });
          criados++;
        }
      });
      return next;
    });

    setLoteConflitos(conflitos);
    if (conflitos.length > 0) {
      toast({
        title: `${conflitos.length} servidor(es) não lançado(s)`,
        description: conflitos.slice(0, 3).join(" | ") + (conflitos.length > 3 ? " ..." : ""),
        variant: "destructive",
      });
    }
    if (criados + atualizados > 0) {
      toast({
        title: "Lançamento por data concluído",
        description: `${criados} novo(s), ${atualizados} atualizado(s)`,
      });
    }
    // Mantém selecionados apenas os que falharam, para correção
    setLoteMatriculas(bloqueados);
  };


  // Consolidação
  const [filtroOperacao, setFiltroOperacao] = useState<string>("todos");
  const consolidado = useMemo(() => {
    const map = new Map<string, {
      matricula: string; nome: string; cpf: string; coordenador: number; supervisor: number; agente: number; apoio: number;
      totalHoras: number; valorHoras: number; alimentacao: number; transporte: number; totalGeral: number;
    }>();

    lancamentos.forEach((l) => {
      if (filtroOperacao !== "todos" && l.nomeOperacao !== filtroOperacao) return;
      const key = l.servidor.matricula;
      if (!map.has(key)) {
        map.set(key, {
          matricula: key, nome: l.servidor.nome, cpf: l.servidor.cpf || "",
          coordenador: 0, supervisor: 0, agente: 0, apoio: 0,
          totalHoras: 0, valorHoras: 0, alimentacao: 0, transporte: 0, totalGeral: 0,
        });
      }
      const row = map.get(key)!;
      l.dias.forEach((d) => {
        const v = valorHora(d.funcao, l.operacao);
        if (d.funcao.includes("coordenador")) row.coordenador += d.horas;
        else if (d.funcao.includes("supervisor")) row.supervisor += d.horas;
        else if (d.funcao.includes("agente") || d.funcao.includes("guarda") || d.funcao.includes("fiscaliza")) row.agente += d.horas;
        else row.apoio += d.horas;
        row.totalHoras += d.horas;
        row.valorHoras += d.horas * v;
        row.alimentacao += calcAlimentacao(l.operacao, d.horas);
      });
      if (l.operacao === "carnaval") row.transporte += 20;
      row.totalGeral = row.valorHoras + row.alimentacao + row.transporte;
    });

    return Array.from(map.values());
  }, [lancamentos, filtroOperacao, valores, alimentacao]);

  const nomesOperacoesConsolidadas = useMemo(
    () => Array.from(new Set(lancamentos.map((l) => l.nomeOperacao))),
    [lancamentos]
  );

  const contatoRodape = [contatos.telefone1, contatos.telefone2].filter(Boolean).join(" • ");

  // Totais de um lançamento
  const totaisLancamento = (l: Lancamento) => {
    let horas = 0, valor = 0, alim = 0;
    l.dias.forEach((d) => {
      horas += d.horas;
      valor += d.horas * valorHora(d.funcao, l.operacao);
      alim += calcAlimentacao(l.operacao, d.horas);
    });
    const transporte = l.operacao === "carnaval" ? 20 : 0;
    return { horas, valor, alim, transporte, total: valor + alim + transporte };
  };

  // PDF — Extrato do lançamento registrado
  const gerarExtratoLancamentoPDF = (l: Lancamento) => {
    const doc = new jsPDF();
    let y = drawHeader(doc, "Extrato de Frequência Registrada", l.nomeOperacao);

    y = drawInfoBox(doc, y, [
      ["Servidor", l.servidor.nome],
      ["Matrícula", l.servidor.matricula],
      ["CPF", formatarCPF(l.servidor.cpf)],
      ["Registro Único", l.servidor.registroUnico || "—"],
      ["Operação", l.nomeOperacao],
      ["Período", `${toBR(l.periodo.inicio)} a ${toBR(l.periodo.fim)}`],
      ["Registrado em", new Date(l.createdAt).toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" })],
      ["Dias lançados", String(l.dias.length)],
    ]);

    const t = totaisLancamento(l);
    const rows = [...l.dias]
      .sort((a, b) => a.data.localeCompare(b.data))
      .map((d) => [
        toBR(d.data),
        `${d.horas}h`,
        funcaoLabel(d.funcao),
        fmtBRL(valorHora(d.funcao, l.operacao)),
        fmtBRL(d.horas * valorHora(d.funcao, l.operacao)),
        fmtBRL(calcAlimentacao(l.operacao, d.horas)),
      ]);

    autoTable(doc, {
      startY: y,
      head: [["Data", "Horas", "Função na operação", "Valor/Hora", "Valor Horas", "Alimentação"]],
      body: rows,
      foot: [["TOTAL", `${t.horas}h`, "", "", fmtBRL(t.valor), fmtBRL(t.alim)]],
      margin: { left: 14, right: 14 },
      columnStyles: {
        0: { halign: "center", cellWidth: 26 },
        1: { halign: "center", cellWidth: 18 },
        3: { halign: "right", cellWidth: 26 },
        4: { halign: "right", cellWidth: 28 },
        5: { halign: "right", cellWidth: 28 },
      },
      ...(tableTheme as any),
    });

    let fy = ((doc as any).lastAutoTable?.finalY ?? y) + 10;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    doc.setTextColor(26, 45, 84);
    doc.text(`TOTAL A RECEBER: ${fmtBRL(t.total)}`, 14, fy);
    doc.setTextColor(0, 0, 0);

    fy += 22;
    drawSignatures(doc, fy, [
      { titulo: l.servidor.nome, sub: `Matrícula ${l.servidor.matricula} — Servidor` },
      { titulo: gestor.nome || "Gestor responsável", sub: gestor.cargo || "Assinatura e carimbo" },
    ]);

    drawFooters(doc, contatoRodape);
    doc.save(`Extrato_${l.servidor.matricula}.pdf`);
    toast({ title: "Extrato gerado" });
  };

  // PDF — Relatório do histórico de lançamentos
  const gerarRelatorioHistoricoPDF = () => {
    const lista = lancamentos.filter((l) => filtroHistorico === "todos" || l.nomeOperacao === filtroHistorico);
    if (lista.length === 0) {
      toast({ title: "Nenhum lançamento para exportar", variant: "destructive" });
      return;
    }
    const doc = new jsPDF({ orientation: "landscape" });
    let y = drawHeader(
      doc,
      "Relatório de Frequências Registradas",
      filtroHistorico === "todos" ? "Todas as operações" : filtroHistorico
    );

    const ordenada = [...lista].sort((a, b) => a.servidor.nome.localeCompare(b.servidor.nome));
    const acc = ordenada.reduce(
      (a, l) => {
        const t = totaisLancamento(l);
        a.h += t.horas; a.v += t.valor; a.al += t.alim; a.tr += t.transporte; a.g += t.total;
        return a;
      },
      { h: 0, v: 0, al: 0, tr: 0, g: 0 }
    );

    y = drawInfoBox(doc, y, [
      ["Lançamentos", String(ordenada.length)],
      ["Total de horas", acc.h.toFixed(2)],
      ["Emissão", new Date().toLocaleDateString("pt-BR")],
      ["Total geral", fmtBRL(acc.g)],
    ]);

    const rows = ordenada.map((l, i) => {
      const t = totaisLancamento(l);
      const datas = [...l.dias].sort((a, b) => a.data.localeCompare(b.data)).map((d) => toBR(d.data)).join(", ");
      return [
        String(i + 1),
        l.servidor.matricula,
        l.servidor.nome,
        formatarCPF(l.servidor.cpf),
        l.nomeOperacao,
        datas,
        t.horas.toFixed(2),
        fmtBRL(t.valor),
        fmtBRL(t.alim),
        fmtBRL(t.total),
      ];
    });

    autoTable(doc, {
      startY: y,
      head: [["#", "Matrícula", "Nome do servidor", "CPF", "Operação", "Datas lançadas", "Horas", "Valor Horas", "Alimentação", "Total"]],
      body: rows,
      foot: [["", "", "TOTAL GERAL", "", "", "", acc.h.toFixed(2), fmtBRL(acc.v), fmtBRL(acc.al), fmtBRL(acc.g)]],
      margin: { left: 14, right: 14 },
      columnStyles: {
        0: { halign: "center", cellWidth: 8 },
        1: { halign: "center" },
        3: { halign: "center" },
        5: { cellWidth: 70 },
        6: { halign: "center" },
        7: { halign: "right" }, 8: { halign: "right" }, 9: { halign: "right" },
      },
      ...(tableTheme as any),
      bodyStyles: { ...(tableTheme.bodyStyles as any), fontSize: 7.5 },
      headStyles: { ...(tableTheme.headStyles as any), fontSize: 7.5 },
      footStyles: { ...(tableTheme.footStyles as any), fontSize: 7.5 },
    });

    const fy = ((doc as any).lastAutoTable?.finalY ?? y) + 20;
    drawSignatures(doc, fy, [
      { titulo: gestor.nome || "Gestor responsável", sub: gestor.cargo || "Assinatura e carimbo" },
      { titulo: "Data: ____/____/________", sub: "Conferência" },
    ]);

    drawFooters(doc, contatoRodape);
    doc.save("Relatorio_Frequencias_Registradas.pdf");
    toast({ title: "Relatório gerado" });
  };


  // PDF — Planilha consolidada
  const gerarPlanilhaPDF = () => {
    if (consolidado.length === 0) {
      toast({ title: "Nenhum dado para exportar", variant: "destructive" });
      return;
    }
    const doc = new jsPDF({ orientation: "landscape" });
    const subtitulo = filtroOperacao !== "todos" ? filtroOperacao : "Todas as operações";
    let y = drawHeader(doc, "Planilha Consolidada de Operações Especiais", subtitulo);

    const totals = consolidado.reduce(
      (acc, r) => {
        acc.h += r.totalHoras; acc.v += r.valorHoras; acc.a += r.alimentacao; acc.t += r.transporte; acc.g += r.totalGeral;
        return acc;
      },
      { h: 0, v: 0, a: 0, t: 0, g: 0 }
    );

    y = drawInfoBox(doc, y, [
      ["Servidores", String(consolidado.length)],
      ["Total de horas", totals.h.toFixed(2)],
      ["Emissão", new Date().toLocaleDateString("pt-BR")],
      ["Total geral", fmtBRL(totals.g)],
    ]);

    const rows = [...consolidado]
      .sort((a, b) => a.nome.localeCompare(b.nome))
      .map((r, i) => [
        String(i + 1),
        r.matricula,
        r.nome,
        formatarCPF(r.cpf),
        r.coordenador.toFixed(2),
        r.supervisor.toFixed(2),
        r.agente.toFixed(2),
        r.apoio.toFixed(2),
        r.totalHoras.toFixed(2),
        fmtBRL(r.valorHoras),
        fmtBRL(r.alimentacao),
        r.transporte > 0 ? fmtBRL(r.transporte) : "—",
        fmtBRL(r.totalGeral),
      ]);

    autoTable(doc, {
      startY: y,
      head: [["#", "Matrícula", "Nome do servidor", "CPF", "Coord. (h)", "Super. (h)", "Agente (h)", "Apoio (h)", "Total (h)", "Valor Horas", "Alimentação", "Transporte", "Total Geral"]],
      body: rows,
      foot: [[
        "", "", "TOTAL GERAL", "", "", "", "", "",
        totals.h.toFixed(2), fmtBRL(totals.v), fmtBRL(totals.a), fmtBRL(totals.t), fmtBRL(totals.g),
      ]],
      margin: { left: 14, right: 14 },
      columnStyles: {
        0: { halign: "center", cellWidth: 8 },
        1: { halign: "center" },
        3: { halign: "center" },
        4: { halign: "center" }, 5: { halign: "center" }, 6: { halign: "center" }, 7: { halign: "center" }, 8: { halign: "center" },
        9: { halign: "right" }, 10: { halign: "right" }, 11: { halign: "right" }, 12: { halign: "right" },
      },
      ...(tableTheme as any),
      bodyStyles: { ...(tableTheme.bodyStyles as any), fontSize: 7.5 },
      headStyles: { ...(tableTheme.headStyles as any), fontSize: 7.5 },
      footStyles: { ...(tableTheme.footStyles as any), fontSize: 7.5 },
    });

    const fy = ((doc as any).lastAutoTable?.finalY ?? y) + 20;
    drawSignatures(doc, fy, [
      { titulo: gestor.nome || "Gestor responsável", sub: gestor.cargo || "Assinatura e carimbo" },
      { titulo: "Data: ____/____/________", sub: "Conferência" },
    ]);

    drawFooters(doc, contatoRodape);
    const nomeArquivo = filtroOperacao !== "todos" ? `Planilha_${filtroOperacao.replace(/\s+/g, "_")}.pdf` : "Planilha_Consolidada_Todas.pdf";
    doc.save(nomeArquivo);
    toast({ title: "Planilha gerada" });
  };

  const exportarPlanilhaExcel = () => {
    if (consolidado.length === 0) {
      toast({ title: "Nenhum dado para exportar", variant: "destructive" });
      return;
    }
    const headers = ["Matrícula", "Nome", "CPF", "Coord. (h)", "Super. (h)", "Agente (h)", "Apoio (h)", "Total Horas", "Valor Horas", "Alimentação", "Transporte", "Total Geral"];
    const rows = [...consolidado]
      .sort((a, b) => a.nome.localeCompare(b.nome))
      .map((r) => [
        r.matricula, r.nome, formatarCPF(r.cpf),
        Number(r.coordenador.toFixed(2)), Number(r.supervisor.toFixed(2)), Number(r.agente.toFixed(2)), Number(r.apoio.toFixed(2)),
        Number(r.totalHoras.toFixed(2)), Number(r.valorHoras.toFixed(2)), Number(r.alimentacao.toFixed(2)),
        Number(r.transporte.toFixed(2)), Number(r.totalGeral.toFixed(2)),
      ]);
    const totals = consolidado.reduce(
      (acc, r) => { acc.h += r.totalHoras; acc.v += r.valorHoras; acc.a += r.alimentacao; acc.t += r.transporte; acc.g += r.totalGeral; return acc; },
      { h: 0, v: 0, a: 0, t: 0, g: 0 }
    );
    const ws = XLSX.utils.aoa_to_sheet([
      headers,
      ...rows,
      ["TOTAL GERAL", "", "", "", "", "", "", Number(totals.h.toFixed(2)), Number(totals.v.toFixed(2)), Number(totals.a.toFixed(2)), Number(totals.t.toFixed(2)), Number(totals.g.toFixed(2))],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Consolidado");
    const nomeArquivo = filtroOperacao !== "todos" ? `Planilha_${filtroOperacao.replace(/\s+/g, "_")}.xlsx` : "Planilha_Consolidada_Todas.xlsx";
    XLSX.writeFile(wb, nomeArquivo);
    toast({ title: "Planilha Excel gerada" });
  };

  const zerarLancamentos = () => {
    setLancamentos([]);
    localStorage.removeItem("lancamentos");
    toast({ title: "Lançamentos zerados", description: "Banco de servidores preservado." });
  };

  const excluirLancamento = (id: string) => {
    setLancamentos((prev) => prev.filter((l) => l.id !== id));
    toast({ title: "Lançamento excluído" });
  };

  const editarLancamento = (l: Lancamento) => {
    setEditingId(l.id);
    const y = new Date(l.periodo.inicio).getFullYear();
    setAno(y);
    const op = buildOpcoes(y).find((o) => o.inicio === l.periodo.inicio && o.fim === l.periodo.fim && o.tipo === l.operacao);
    if (op) setOperacaoId(op.id);
    setPeriodo(l.periodo);
    setMatriculaSelecionada(l.servidor.matricula);
    setDias(l.dias);
    setTab("lancamentos");
    toast({ title: "Editando lançamento", description: `${l.servidor.nome} - ${l.nomeOperacao}` });
  };

  // Tela de login (usuário único + senha master)
  if (!sistemaDesbloqueado) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-background to-card/60">
        <Card className="w-full max-w-md">
          <CardHeader className="text-center">
            <CardTitle className="text-2xl mb-2">GEOPS</CardTitle>
            <p className="text-muted-foreground text-sm">Sistema de Geração e Controle de Operações Especiais</p>
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              <div className="space-y-2">
                <Label>Usuário</Label>
                <Input value={USUARIO_MASTER} readOnly disabled />
              </div>
              <div className="space-y-2">
                <Label>Senha master</Label>
                <Input
                  type="password"
                  value={loginSenha}
                  onChange={(e) => setLoginSenha(e.target.value)}
                  placeholder="Digite a senha master"
                  onKeyDown={(e) => e.key === "Enter" && realizarLogin()}
                />
              </div>
              <Button onClick={realizarLogin} className="w-full">Entrar</Button>

              <hr className="my-4" />

              <div className="space-y-3">
                <h4 className="font-medium">Recuperar acesso</h4>
                <div className="space-y-2">
                  <Label>Pergunta: {SECURITY_QUESTION}</Label>
                  <Input value={respostaSeguranca} onChange={(e) => setRespostaSeguranca(e.target.value)} placeholder="Digite a resposta" />
                </div>
                <div className="space-y-2">
                  <Label>Nova senha master</Label>
                  <Input type="password" value={novaSenhaSeguranca} onChange={(e) => setNovaSenhaSeguranca(e.target.value)} placeholder="Digite a nova senha" />
                </div>
                <Button onClick={redefinirSenha} variant="secondary" className="w-full">Redefinir senha master</Button>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>
    );
  }

  const funcoesDaOperacao = getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria");

  return (
    <div className="min-h-screen">
      <header className="border-b bg-gradient-to-br from-background to-card/60">
        <div className="container py-8">
          <div className="flex justify-between items-start gap-4">
            <div>
              <h1 className="text-3xl font-bold tracking-tight">GEOPS - Gerador de Operações Especiais Segep</h1>
              <p className="text-muted-foreground mt-1">Sistema de Geração e Controle de Operações Especiais.</p>
            </div>
            <div className="flex items-center gap-3">
              <div className="text-right">
                <div className="font-medium">{USUARIO_MASTER}</div>
                <div className="text-sm text-muted-foreground">Acesso master</div>
              </div>
              <Button onClick={logout} variant="outline" size="sm">Sair</Button>
            </div>
          </div>
        </div>
      </header>

      <main className="container py-8">
        <Tabs value={tab} onValueChange={setTab} className="w-full">
          <TabsList className="grid w-full grid-cols-5">
            <TabsTrigger value="lancamentos">Por Servidor</TabsTrigger>
            <TabsTrigger value="por-data">Por Data</TabsTrigger>
            <TabsTrigger value="planilha">Planilha</TabsTrigger>
            <TabsTrigger value="logs">Histórico</TabsTrigger>
            <TabsTrigger value="rh">Banco de Dados e Configurações</TabsTrigger>
          </TabsList>

          {/* Lançamento por servidor */}
          <TabsContent value="lancamentos" className="mt-6">
            <div className="grid gap-6 md:grid-cols-3">
              <Card className="md:col-span-1">
                <CardHeader><CardTitle>Configurações</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  <div className="space-y-2">
                    <Label>Ano</Label>
                    <Select value={String(ano)} onValueChange={(v) => setAno(Number(v))}>
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent className="z-50">
                        {anosDisponiveis.map((y) => (<SelectItem key={y} value={String(y)}>{y}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label>Operação</Label>
                    <Select
                      value={operacaoId}
                      onValueChange={(v) => {
                        setOperacaoId(v);
                        const op = opcoesOperacao.find((o) => o.id === v);
                        if (op) {
                          setPeriodo({ inicio: op.inicio, fim: op.fim });
                          setLoteDatas([]);
                          setDias((d) => (d.length === 0 ? [{ data: op.inicio, horas: 8, funcao: getFuncoesDisponiveisParaOperacao(op.tipo)[0] }] : d));
                        }
                      }}
                    >
                      <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
                      <SelectContent className="z-50">
                        {opcoesOperacao.map((op) => (<SelectItem key={op.id} value={op.id}>{op.label}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>

                  {selectedOp?.tipo === "carnaval" && (
                    <div className="space-y-2">
                      <Label>Período Customizado (apenas Carnaval)</Label>
                      <div className="grid grid-cols-2 gap-2">
                        <Input
                          type="date"
                          value={periodoCustomizado.inicio}
                          onChange={(e) => {
                            setPeriodoCustomizado((prev) => ({ ...prev, inicio: e.target.value }));
                            if (e.target.value && periodoCustomizado.fim) setPeriodo({ inicio: e.target.value, fim: periodoCustomizado.fim });
                          }}
                        />
                        <Input
                          type="date"
                          value={periodoCustomizado.fim}
                          onChange={(e) => {
                            setPeriodoCustomizado((prev) => ({ ...prev, fim: e.target.value }));
                            if (periodoCustomizado.inicio && e.target.value) setPeriodo({ inicio: periodoCustomizado.inicio, fim: e.target.value });
                          }}
                        />
                      </div>
                    </div>
                  )}

                  <div className="space-y-2">
                    <Label>Período</Label>
                    <div className="text-sm text-muted-foreground">{toBR(periodo.inicio)} a {toBR(periodo.fim)}</div>
                  </div>

                  <div className="space-y-2">
                    <Label>Servidor</Label>
                    <ServerCombobox
                      servidores={servidores}
                      value={matriculaSelecionada}
                      onChange={setMatriculaSelecionada}
                      placeholder="Buscar servidor..."
                    />
                    {matriculaSelecionada && (
                      <div className="text-sm text-muted-foreground">
                        Selecionado: {servidores.find((s) => s.matricula === matriculaSelecionada)?.nome}
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>

              <Card className="md:col-span-2">
                <CardHeader><CardTitle>Lançamento de Horas</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  {dias.map((d, i) => (
                    <div key={i} className="flex items-center gap-2 border rounded-md p-3 flex-wrap">
                      <Input
                        type="date"
                        value={d.data}
                        onChange={(e) => {
                          const v = e.target.value;
                          if (v < periodo.inicio || v > periodo.fim) {
                            toast({ title: "Data fora do período", variant: "destructive" });
                            atualizarDia(i, { data: clampToPeriodo(v, periodo.inicio, periodo.fim) });
                          } else {
                            atualizarDia(i, { data: v });
                          }
                        }}
                        className="max-w-[190px]"
                      />
                      <Input
                        type="number"
                        min={1}
                        max={selectedOp?.tipo === "carnaval" ? 24 : 12}
                        value={d.horas}
                        onChange={(e) => {
                          const limit = selectedOp?.tipo === "carnaval" ? 24 : 12;
                          const n = Math.max(1, Math.min(Number(e.target.value || 0), limit));
                          atualizarDia(i, { horas: n });
                        }}
                        className="w-20"
                      />
                      <Select value={d.funcao} onValueChange={(v: FuncaoID) => atualizarDia(i, { funcao: v })}>
                        <SelectTrigger className="w-56"><SelectValue placeholder="Função" /></SelectTrigger>
                        <SelectContent className="z-50">
                          {funcoesDaOperacao.map((funcao) => (<SelectItem key={funcao} value={funcao}>{funcaoLabel(funcao)}</SelectItem>))}
                        </SelectContent>
                      </Select>
                      <div className="text-sm whitespace-nowrap">Valor/h: {fmtBRL(valorHora(d.funcao, selectedOp?.tipo ?? "ordinaria"))}</div>
                      <div className="ms-auto flex gap-2">
                        <Button variant="secondary" onClick={() => atualizarDia(i, { horas: d.horas + 1 })}>+1h</Button>
                        <Button variant="destructive" onClick={() => removerDia(i)}>Remover</Button>
                      </div>
                    </div>
                  ))}

                  <div className="flex gap-2 flex-wrap">
                    <Button onClick={adicionarDia}>+ Adicionar dia</Button>
                    <Button onClick={salvarLancamento}>{editingId ? "Salvar alterações" : "Registrar frequência"}</Button>
                    {editingId && (
                      <Button
                        variant="secondary"
                        onClick={() => {
                          setEditingId(null);
                          setDias([{ data: periodo.inicio, horas: 8, funcao: getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria")[0] }]);
                          toast({ title: "Edição cancelada" });
                        }}
                      >
                        Cancelar edição
                      </Button>
                    )}
                  </div>

                </CardContent>
              </Card>
            </div>
          </TabsContent>

          {/* Lançamento por data */}
          <TabsContent value="por-data" className="mt-6">
            <div className="grid gap-6 md:grid-cols-3">
              <Card className="md:col-span-1">
                <CardHeader><CardTitle>Datas e carga horária</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  <div className="space-y-2">
                    <Label>Operação</Label>
                    <Select
                      value={operacaoId}
                      onValueChange={(v) => {
                        setOperacaoId(v);
                        const op = opcoesOperacao.find((o) => o.id === v);
                        if (op) { setPeriodo({ inicio: op.inicio, fim: op.fim }); setLoteDatas([]); }
                      }}
                    >
                      <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
                      <SelectContent className="z-50">
                        {opcoesOperacao.map((op) => (<SelectItem key={op.id} value={op.id}>{op.label}</SelectItem>))}
                      </SelectContent>
                    </Select>
                    <div className="text-xs text-muted-foreground">Período: {toBR(periodo.inicio)} a {toBR(periodo.fim)}</div>
                  </div>

                  <div className="space-y-2">
                    <Label>Adicionar data</Label>
                    <div className="flex gap-2">
                      <Input type="date" value={loteNovaData} min={periodo.inicio} max={periodo.fim} onChange={(e) => setLoteNovaData(e.target.value)} />
                      <Button onClick={adicionarLoteData}>+</Button>
                    </div>
                    <div className="flex flex-wrap gap-2 pt-1">
                      {loteDatas.length === 0 && <span className="text-xs text-muted-foreground">Nenhuma data adicionada.</span>}
                      {loteDatas.map((d) => (
                        <Badge key={d} variant="secondary" className="cursor-pointer" onClick={() => setLoteDatas((prev) => prev.filter((x) => x !== d))}>
                          {toBR(d)} ✕
                        </Badge>
                      ))}
                    </div>
                  </div>

                  <div className="space-y-2">
                    <Label>Carga horária por dia</Label>
                    <Input
                      type="number"
                      min={1}
                      max={selectedOp?.tipo === "carnaval" ? 24 : 12}
                      value={loteHoras}
                      onChange={(e) => {
                        const limit = selectedOp?.tipo === "carnaval" ? 24 : 12;
                        setLoteHoras(Math.max(1, Math.min(Number(e.target.value || 0), limit)));
                      }}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label>Função na operação</Label>
                    <Select value={loteFuncao} onValueChange={(v: FuncaoID) => setLoteFuncao(v)}>
                      <SelectTrigger><SelectValue placeholder="Função" /></SelectTrigger>
                      <SelectContent className="z-50">
                        {funcoesDaOperacao.map((f) => (<SelectItem key={f} value={f}>{funcaoLabel(f)}</SelectItem>))}
                      </SelectContent>
                    </Select>
                    <div className="text-xs text-muted-foreground">Valor/h: {fmtBRL(valorHora(loteFuncao, selectedOp?.tipo ?? "ordinaria"))}</div>
                  </div>
                </CardContent>
              </Card>

              <Card className="md:col-span-2">
                <CardHeader><CardTitle>Servidores do dia</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  <MultiServerSelect servidores={servidores} selecionados={loteMatriculas} onChange={setLoteMatriculas} />
                  {loteConflitos.length > 0 && (
                    <div className="rounded-md border border-destructive/40 bg-destructive/10 p-3 space-y-1">
                      <div className="text-sm font-medium text-destructive">
                        {loteConflitos.length} servidor(es) NÃO foram lançados (datas já lançadas):
                      </div>
                      <ul className="list-disc pl-5 text-xs text-destructive space-y-0.5">
                        {loteConflitos.map((c) => (<li key={c}>{c}</li>))}
                      </ul>
                    </div>
                  )}
                  <div className="text-xs text-muted-foreground">
                    Serão criados {loteDatas.length} dia(s) de {loteHoras}h para cada servidor selecionado. Servidores que já possuam lançamento em qualquer uma das datas não são lançados e o sistema aponta o erro.
                  </div>

                  <Button onClick={salvarLancamentoPorData}>Lançar para os servidores selecionados</Button>
                </CardContent>
              </Card>
            </div>
          </TabsContent>

          {/* Planilha */}
          <TabsContent value="planilha" className="mt-6">
            <Card>
              <CardHeader><CardTitle>Resumo de Horas Extras</CardTitle></CardHeader>
              <CardContent>
                <div className="grid md:grid-cols-3 gap-3 mb-4">
                  <div className="space-y-2 md:col-span-1">
                    <Label>Operação</Label>
                    <Select value={filtroOperacao} onValueChange={setFiltroOperacao}>
                      <SelectTrigger><SelectValue placeholder="Todas" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todos">Todas as Operações</SelectItem>
                        {nomesOperacoesConsolidadas.map((n) => (<SelectItem key={n} value={n}>{n}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>
                </div>

                <div className="flex gap-2 mb-4 flex-wrap">
                  <Button onClick={gerarPlanilhaPDF}>Exportar PDF</Button>
                  <Button variant="secondary" onClick={exportarPlanilhaExcel}>Exportar Excel</Button>
                  <Button variant="destructive" onClick={zerarLancamentos}>Zerar Dados</Button>
                </div>

                <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Matrícula</TableHead>
                        <TableHead>Nome</TableHead>
                        <TableHead>CPF</TableHead>
                        <TableHead>Coord. (h)</TableHead>
                        <TableHead>Super. (h)</TableHead>
                        <TableHead>Agente (h)</TableHead>
                        <TableHead>Apoio (h)</TableHead>
                        <TableHead>Total Horas</TableHead>
                        <TableHead>Valor Horas</TableHead>
                        <TableHead>Alimentação</TableHead>
                        <TableHead>Transporte</TableHead>
                        <TableHead>Total Geral</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {consolidado.length === 0 && (
                        <TableRow><TableCell colSpan={12} className="text-center text-muted-foreground">Nenhum lançamento.</TableCell></TableRow>
                      )}
                      {[...consolidado].sort((a, b) => a.nome.localeCompare(b.nome)).map((r) => (
                        <TableRow key={r.matricula}>
                          <TableCell>{r.matricula}</TableCell>
                          <TableCell>{r.nome}</TableCell>
                          <TableCell>{formatarCPF(r.cpf)}</TableCell>
                          <TableCell>{r.coordenador.toFixed(2)}</TableCell>
                          <TableCell>{r.supervisor.toFixed(2)}</TableCell>
                          <TableCell>{r.agente.toFixed(2)}</TableCell>
                          <TableCell>{r.apoio.toFixed(2)}</TableCell>
                          <TableCell>{r.totalHoras.toFixed(2)}</TableCell>
                          <TableCell>{fmtBRL(r.valorHoras)}</TableCell>
                          <TableCell>{fmtBRL(r.alimentacao)}</TableCell>
                          <TableCell>{r.transporte > 0 ? fmtBRL(r.transporte) : "-"}</TableCell>
                          <TableCell>{fmtBRL(r.totalGeral)}</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Histórico */}
          <TabsContent value="logs" className="mt-6">
            <Card>
              <CardHeader><CardTitle>Histórico de Lançamentos</CardTitle></CardHeader>
              <CardContent>
                <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Data</TableHead>
                        <TableHead>Servidor</TableHead>
                        <TableHead>Operação</TableHead>
                        <TableHead>Período</TableHead>
                        <TableHead>Dias</TableHead>
                        <TableHead>Ações</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {lancamentos.length === 0 && (
                        <TableRow><TableCell colSpan={6} className="text-center text-muted-foreground">Sem lançamentos.</TableCell></TableRow>
                      )}
                      {[...lancamentos].sort((a, b) => a.servidor.nome.localeCompare(b.servidor.nome)).map((l) => (
                        <TableRow key={l.id}>
                          <TableCell>{new Date(l.createdAt).toLocaleDateString("pt-BR")}</TableCell>
                          <TableCell>{l.servidor.nome} ({l.servidor.matricula})</TableCell>
                          <TableCell>{l.nomeOperacao}</TableCell>
                          <TableCell>{toBR(l.periodo.inicio)} a {toBR(l.periodo.fim)}</TableCell>
                          <TableCell>{l.dias.length}</TableCell>
                          <TableCell className="flex gap-2">
                            <Button size="sm" onClick={() => editarLancamento(l)}>Editar</Button>
                            <Button size="sm" variant="secondary" onClick={() => gerarFrequenciaPDF(l)}>PDF</Button>
                            <Button size="sm" variant="destructive" onClick={() => excluirLancamento(l.id)}>Excluir</Button>
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Banco de Dados / Configurações */}
          <TabsContent value="rh" className="mt-6">
            <div className="grid gap-6 md:grid-cols-2">
              <Card>
                <CardHeader><CardTitle>Acesso</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  <div className="text-sm text-muted-foreground">Usuário único: <span className="font-medium text-foreground">{USUARIO_MASTER}</span></div>
                  <div className="space-y-3">
                    <h4 className="font-medium">Alterar senha master</h4>
                    <div className="space-y-2">
                      <Label>Nova senha</Label>
                      <Input type="password" value={novaSenha} onChange={(e) => setNovaSenha(e.target.value)} placeholder="Digite a nova senha" />
                    </div>
                    <div className="space-y-2">
                      <Label>Confirmar senha</Label>
                      <Input type="password" value={confirmarSenha} onChange={(e) => setConfirmarSenha(e.target.value)} placeholder="Confirme a nova senha" />
                    </div>
                    <Button onClick={alterarSenha} variant="secondary" className="w-full">Alterar senha master</Button>
                  </div>
                  <Button onClick={logout} variant="destructive" className="w-full">Sair</Button>
                </CardContent>
              </Card>

              <Card>
                <CardHeader><CardTitle>Identificação nos documentos</CardTitle></CardHeader>
                <CardContent className="space-y-3">
                  <div className="space-y-1">
                    <Label>Nome do gestor (assinatura)</Label>
                    <Input value={gestor.nome} onChange={(e) => setGestor((g) => ({ ...g, nome: e.target.value }))} placeholder="Nome completo do gestor" />
                  </div>
                  <div className="space-y-1">
                    <Label>Cargo do gestor</Label>
                    <Input value={gestor.cargo} onChange={(e) => setGestor((g) => ({ ...g, cargo: e.target.value }))} placeholder="Ex.: Gerente de Operações Especiais" />
                  </div>
                  <div className="grid md:grid-cols-2 gap-2">
                    <div className="space-y-1">
                      <Label>Telefone principal</Label>
                      <Input value={contatos.telefone1} onChange={(e) => setContatos((c) => ({ ...c, telefone1: e.target.value }))} placeholder="(71) 9 0000-0000" />
                    </div>
                    <div className="space-y-1">
                      <Label>Telefone secundário</Label>
                      <Input value={contatos.telefone2} onChange={(e) => setContatos((c) => ({ ...c, telefone2: e.target.value }))} placeholder="(71) 9 0000-0000" />
                    </div>
                  </div>
                  <Button onClick={salvarAlteracoes}>Salvar alterações</Button>
                </CardContent>
              </Card>

              <Card className="md:col-span-2">
                <CardHeader><CardTitle>Valores por Operação</CardTitle></CardHeader>
                <CardContent>
                  <div className="grid md:grid-cols-3 gap-6">
                    {(["ordinaria", "reveillon", "carnaval"] as OperacaoTipo[]).map((tipo) => (
                      <div key={tipo}>
                        <h4 className="font-medium mb-3 capitalize">Operação {tipo}</h4>
                        <div className="grid grid-cols-2 gap-2">
                          {Object.entries((valores as any)[tipo]).map(([funcao, valor]) => (
                            <div key={funcao} className="space-y-1">
                              <Label className="text-xs">{funcaoLabel(funcao as FuncaoID)}</Label>
                              <Input
                                type="number"
                                step="0.01"
                                value={valor as number}
                                onChange={(e) => setValores((prev: any) => ({ ...prev, [tipo]: { ...prev[tipo], [funcao]: Number(e.target.value) } }))}
                                className="text-xs"
                              />
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="mt-6 grid md:grid-cols-3 gap-6">
                    <div>
                      <h5 className="text-sm font-medium mb-2">Alimentação — Ordinária</h5>
                      <Label className="text-xs">Valor por hora (mínimo 8h)</Label>
                      <Input
                        type="number"
                        step="0.01"
                        value={(alimentacao as any).ordinaria?.valorHora ?? 2}
                        onChange={(e) => setAlimentacao((prev: any) => ({ ...prev, ordinaria: { ...prev.ordinaria, valorHora: Number(e.target.value) } }))}
                        className="text-xs"
                      />
                    </div>
                    <div>
                      <h5 className="text-sm font-medium mb-2">Alimentação — Reveillon</h5>
                      <Label className="text-xs">Valor para 12h</Label>
                      <Input
                        type="number"
                        step="0.01"
                        value={(alimentacao as any).reveillon?.[12] ?? 13.68}
                        onChange={(e) => setAlimentacao((prev: any) => ({ ...prev, reveillon: { ...prev.reveillon, 12: Number(e.target.value) } }))}
                        className="text-xs"
                      />
                    </div>
                    <div>
                      <h5 className="text-sm font-medium mb-2">Alimentação — Carnaval (1h a 24h)</h5>
                      <div className="grid grid-cols-3 gap-2 max-h-60 overflow-y-auto">
                        {Object.entries((alimentacao as any).carnaval || {})
                          .sort(([a], [b]) => Number(a) - Number(b))
                          .map(([horas, valor]) => (
                            <div key={horas} className="space-y-1">
                              <Label className="text-xs">{horas}h</Label>
                              <Input
                                type="number"
                                step="0.01"
                                value={valor as number}
                                onChange={(e) => setAlimentacao((prev: any) => ({ ...prev, carnaval: { ...prev.carnaval, [horas]: Number(e.target.value) } }))}
                                className="text-xs"
                              />
                            </div>
                          ))}
                      </div>
                    </div>
                  </div>
                  <Button className="mt-4" onClick={salvarAlteracoes}>Salvar valores</Button>
                </CardContent>
              </Card>

              <Card className="md:col-span-2">
                <CardHeader><CardTitle>Servidores</CardTitle></CardHeader>
                <CardContent>
                  <div className="flex gap-2 mb-2 flex-wrap items-center">
                    <Input type="file" accept=".xlsx,.xls" onChange={(e) => setExcelFile(e.target.files?.[0] ?? null)} className="max-w-xs" />
                    <Button onClick={importarExcel}>Importar Excel</Button>
                    <Button variant="secondary" onClick={limparTudo}>Limpar tudo</Button>
                    <Badge variant="secondary">{servidores.length} servidor(es)</Badge>
                  </div>
                  <div className="text-xs text-muted-foreground mb-4">
                    Colunas esperadas: NOME, NASC., SEXO, MATRÍCULA, R. ÚNICO, R.G, C.P.F, ADMISSÃO. A coluna CARGO/FUNÇÃO é ignorada na importação.
                  </div>

                  <div className="rounded-md border overflow-auto max-h-[520px]" style={{ boxShadow: "var(--shadow-elevated)" }}>
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>Nome</TableHead>
                          <TableHead>Nasc.</TableHead>
                          <TableHead>Sexo</TableHead>
                          <TableHead>Matrícula</TableHead>
                          <TableHead>R. Único</TableHead>
                          <TableHead>R.G</TableHead>
                          <TableHead>C.P.F</TableHead>
                          <TableHead>Admissão</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {servidores.length === 0 && (
                          <TableRow><TableCell colSpan={8} className="text-center text-muted-foreground">Nenhum servidor cadastrado.</TableCell></TableRow>
                        )}
                        {servidores.map((s) => (
                          <TableRow key={s.matricula}>
                            <TableCell>{s.nome}</TableCell>
                            <TableCell>{toBR(s.nascimento || "")}</TableCell>
                            <TableCell>{s.sexo}</TableCell>
                            <TableCell>{s.matricula}</TableCell>
                            <TableCell>{s.registroUnico}</TableCell>
                            <TableCell>{s.rg}</TableCell>
                            <TableCell>{formatarCPF(s.cpf)}</TableCell>
                            <TableCell>{toBR(s.admissao || "")}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                </CardContent>
              </Card>

              <Card className="md:col-span-2">
                <CardHeader><CardTitle>Sobre</CardTitle></CardHeader>
                <CardContent className="space-y-1 text-sm">
                  <p>GEOPS - Gerador de Operações Especiais Segep.</p>
                  <p>Desenvolvido por Rui Cezar Pereira da Paixão Junior</p>
                  <p>Salvador, 2025.</p>
                </CardContent>
              </Card>
            </div>
          </TabsContent>
        </Tabs>
      </main>
    </div>
  );
};

export default Index;

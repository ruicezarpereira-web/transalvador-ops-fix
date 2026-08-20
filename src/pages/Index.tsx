import React, { useEffect, useMemo, useState } from "react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Checkbox } from "@/components/ui/checkbox";
import { toast } from "@/hooks/use-toast";
import type { Servidor } from "@/components/ServerCombobox";
import MultiServerSelect from "@/components/MultiServerSelect";
import Header from "@/components/Header";
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

type OperacaoTipo = "ordinaria" | "reveillon" | "carnaval" | "extraordinaria";

type Operacao = { id: string; nome: string; tipo: OperacaoTipo; periodo: { inicio: string; fim: string } };

type Lancamento = {
  id: string;
  operacaoId?: string;
  operacao: OperacaoTipo;
  nomeOperacao: string;
  periodo: { inicio: string; fim: string };
  servidor: Servidor;
  dias: DiaTrabalho[];
  createdAt: string;
};

type ValoresOperacao = { [key in FuncaoID]?: number };

type SystemUser = {
  id: string;
  nome: string;
  usuario: string;
  senha: string;
  mustChangePassword: boolean;
};

const valoresDefault = {
  ordinaria: {
    coordenador: 20.5,
    supervisor: 15.5,
    agente_fiscalizacao: 12.0,
    apoio_adm: 10.0,
  } as ValoresOperacao,
  extraordinaria: {
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
  extraordinaria: { proporcional: true, valorHora: 2.0, minimoHoras: 8 },
  reveillon: { 12: 13.68, proporcional: true },
  carnaval: {
    1: 3.05, 2: 6.09, 3: 9.14, 4: 12.18, 5: 15.23, 6: 18.27, 7: 21.32, 8: 35.92,
    9: 39.01, 10: 42.1, 11: 43.08, 12: 47.0, 13: 50.05, 14: 53.1, 15: 56.15,
    16: 57.13, 17: 58.11, 18: 59.08, 19: 59.09, 20: 62.14, 21: 65.19, 22: 68.24,
    23: 71.29, 24: 74.64,
  },
} as any;

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
const SENHA_MASTER_PADRAO = "Segep@Transalvador2026";

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

  // Operações (cadastro manual — nome, tipo, período)
  const [operacoes, setOperacoes] = useState<Operacao[]>(() => load<Operacao[]>("operacoes", []));
  const [novaOpNome, setNovaOpNome] = useState("");
  const [novaOpTipo, setNovaOpTipo] = useState<OperacaoTipo>("ordinaria");
  const [novaOpInicio, setNovaOpInicio] = useState("");
  const [novaOpFim, setNovaOpFim] = useState("");
  const [editandoOpId, setEditandoOpId] = useState<string | null>(null);

  // Estado de Lançamento
  const initialYear = new Date().getFullYear();
  const [ano, setAno] = useState<number>(initialYear);
  const [operacaoId, setOperacaoId] = useState<string>("");
  const [buscaOperacao, setBuscaOperacao] = useState("");
  const [periodo, setPeriodo] = useState({ inicio: "", fim: "" });
  const [periodoCustomizado, setPeriodoCustomizado] = useState({ inicio: "", fim: "" });
  const [matriculasSelecionadas, setMatriculasSelecionadas] = useState<string[]>([]);
  const [dias, setDias] = useState<DiaTrabalho[]>([{ data: "", horas: 8, funcao: "coordenador" }]);

  // Lançamento por data (lote)
  const [loteDias, setLoteDias] = useState<DiaTrabalho[]>([]);
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
  const [loginUsuario, setLoginUsuario] = useState<string>(USUARIO_MASTER);
  const [loginSenha, setLoginSenha] = useState("");
  const [novaSenha, setNovaSenha] = useState("");
  const [confirmarSenha, setConfirmarSenha] = useState("");
  const [sistemaDesbloqueado, setSistemaDesbloqueado] = useState<boolean>(() => load<boolean>("sistemaDesbloqueado", false));
  const [sessaoAtual, setSessaoAtual] = useState<{ usuario: string; role: "master" | "comum"; nome: string } | null>(() =>
    load("sessaoAtual", null)
  );

  // Usuários comuns (gestão de usuários — apenas Master)
  const [usuariosComuns, setUsuariosComuns] = useState<SystemUser[]>(() => load<SystemUser[]>("usuariosComuns", []));
  const [novoUsuarioNome, setNovoUsuarioNome] = useState("");
  const [novoUsuarioLogin, setNovoUsuarioLogin] = useState("");

  // Pergunta e resposta de segurança (configurável — nunca fica fixa no código)
  const [secQuestion, setSecQuestion] = useState<string>(() => load("secQuestion", ""));
  const [secAnswer, setSecAnswer] = useState<string>(() => load("secAnswer", ""));
  const [novaSecQuestion, setNovaSecQuestion] = useState("");
  const [novaSecAnswer, setNovaSecAnswer] = useState("");
  const [mostrarRecuperacao, setMostrarRecuperacao] = useState(false);
  const [respostaTentativa, setRespostaTentativa] = useState("");
  const [novaSenhaRecuperada, setNovaSenhaRecuperada] = useState("");

  // Troca de senha — usuário comum (exige senha atual)
  const [senhaAtualComum, setSenhaAtualComum] = useState("");
  const [novaSenhaComum, setNovaSenhaComum] = useState("");
  const [confirmarSenhaComum, setConfirmarSenhaComum] = useState("");

  // Primeiro acesso após reset (senha padrão "123456")
  const [primeiraTrocaSenha, setPrimeiraTrocaSenha] = useState("");
  const [primeiraTrocaSenhaConfirmar, setPrimeiraTrocaSenhaConfirmar] = useState("");

  // Modo escuro
  const [darkMode, setDarkMode] = useState<boolean>(() => load<boolean>("darkMode", false));

  // Relatórios (Relatório Geral por servidor/operação)
  const [relFuncao, setRelFuncao] = useState<string>("todas");
  const [relOperacao, setRelOperacao] = useState<string>("todas");
  const [relAno, setRelAno] = useState<string>("todos");
  const [relDataIni, setRelDataIni] = useState("");
  const [relDataFim, setRelDataFim] = useState("");
  const [relServidor, setRelServidor] = useState<string>("todos");
  const [relIncluirDetalhes, setRelIncluirDetalhes] = useState(false);
  const [relTipoOperacao, setRelTipoOperacao] = useState<string>("todos");
  type SortDir = "asc" | "desc";
  const [relSort, setRelSort] = useState<{ campo: string; dir: SortDir }>({ campo: "nome", dir: "asc" });
  const [gerenciarSort, setGerenciarSort] = useState<{ campo: string; dir: SortDir }>({ campo: "nome", dir: "asc" });

  const opcoesOperacao = useMemo(() => {
    const q = buscaOperacao.trim().toLowerCase();
    return operacoes
      .filter((o) => new Date(o.periodo.inicio + "T00:00:00").getFullYear() === ano)
      .filter((o) => !q || o.nome.toLowerCase().includes(q))
      .map((o) => ({ id: o.id, label: o.nome, inicio: o.periodo.inicio, fim: o.periodo.fim, tipo: o.tipo }))
      .sort((a, b) => a.inicio.localeCompare(b.inicio));
  }, [operacoes, ano, buscaOperacao]);
  const selectedOp = useMemo(() => opcoesOperacao.find((o) => o.id === operacaoId), [opcoesOperacao, operacaoId]);

  const anosDisponiveis = useMemo(() => {
    const cy = new Date().getFullYear();
    const startYear = Math.max(2025, cy);
    return Array.from({ length: 15 }, (_, i) => startYear + i);
  }, []);

  useEffect(() => {
    if (operacaoId && opcoesOperacao.find((o) => o.id === operacaoId)) return;
    const first = opcoesOperacao[0];
    if (first) {
      setOperacaoId(first.id);
      setPeriodo({ inicio: first.inicio, fim: first.fim });
      setDias([{ data: first.inicio, horas: 8, funcao: "coordenador" }]);
    } else {
      setOperacaoId("");
      setPeriodo({ inicio: "", fim: "" });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [opcoesOperacao]);

  // CRUD de Operações (Configurações — Master)
  const salvarOperacao = () => {
    if (!novaOpNome.trim() || !novaOpInicio || !novaOpFim) {
      toast({ title: "Preencha nome, tipo e período", variant: "destructive" });
      return;
    }
    if (novaOpFim < novaOpInicio) {
      toast({ title: "Data final não pode ser antes da inicial", variant: "destructive" });
      return;
    }
    const op: Operacao = {
      id: editandoOpId ?? crypto.randomUUID(),
      nome: novaOpNome.trim(),
      tipo: novaOpTipo,
      periodo: { inicio: novaOpInicio, fim: novaOpFim },
    };
    setOperacoes((prev) => (editandoOpId ? prev.map((o) => (o.id === editandoOpId ? op : o)) : [...prev, op]));
    setNovaOpNome(""); setNovaOpTipo("ordinaria"); setNovaOpInicio(""); setNovaOpFim(""); setEditandoOpId(null);
    toast({ title: editandoOpId ? "Operação atualizada" : "Operação criada" });
  };

  const editarOperacao = (o: Operacao) => {
    setEditandoOpId(o.id);
    setNovaOpNome(o.nome);
    setNovaOpTipo(o.tipo);
    setNovaOpInicio(o.periodo.inicio);
    setNovaOpFim(o.periodo.fim);
  };

  const cancelarEdicaoOperacao = () => {
    setEditandoOpId(null);
    setNovaOpNome(""); setNovaOpTipo("ordinaria"); setNovaOpInicio(""); setNovaOpFim("");
  };

  const excluirOperacao = (id: string) => {
    setOperacoes((prev) => prev.filter((o) => o.id !== id));
    toast({ title: "Operação excluída" });
  };

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
    if (tipo === "ordinaria" || tipo === "extraordinaria") {
      const cfg = (alimentacao as any)[tipo] || {};
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

  const resumoLoteDias = useMemo(() => {
    const tipo = selectedOp?.tipo ?? "ordinaria";
    return loteDias.reduce(
      (acc, d) => {
        const vh = d.horas * valorHora(d.funcao, tipo);
        const al = calcAlimentacao(tipo, d.horas);
        return { horas: acc.horas + d.horas, valorHoras: acc.valorHoras + vh, alimentacao: acc.alimentacao + al, total: acc.total + vh + al };
      },
      { horas: 0, valorHoras: 0, alimentacao: 0, total: 0 }
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loteDias, selectedOp, valores, alimentacao]);

  const resumoDias = useMemo(() => {
    const tipo = selectedOp?.tipo ?? "ordinaria";
    return dias.reduce(
      (acc, d) => {
        const vh = d.horas * valorHora(d.funcao, tipo);
        const al = calcAlimentacao(tipo, d.horas);
        return { horas: acc.horas + d.horas, valorHoras: acc.valorHoras + vh, alimentacao: acc.alimentacao + al, total: acc.total + vh + al };
      },
      { horas: 0, valorHoras: 0, alimentacao: 0, total: 0 }
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dias, selectedOp, valores, alimentacao]);

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
      case "extraordinaria":
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
  useEffect(() => save("operacoes", operacoes), [operacoes]);
  useEffect(() => save("senhaMaster", senhaMaster), [senhaMaster]);
  useEffect(() => save("contatosSetor", contatos), [contatos]);
  useEffect(() => save("gestorSetor", gestor), [gestor]);
  useEffect(() => save("valoresOperacoes", valores), [valores]);
  useEffect(() => save("alimentacaoOperacoes", alimentacao), [alimentacao]);
  useEffect(() => save("sistemaDesbloqueado", sistemaDesbloqueado), [sistemaDesbloqueado]);
  useEffect(() => save("sessaoAtual", sessaoAtual), [sessaoAtual]);
  useEffect(() => save("usuariosComuns", usuariosComuns), [usuariosComuns]);
  useEffect(() => save("secQuestion", secQuestion), [secQuestion]);
  useEffect(() => save("secAnswer", secAnswer), [secAnswer]);
  useEffect(() => save("darkMode", darkMode), [darkMode]);
  useEffect(() => {
    document.documentElement.classList.toggle("dark", darkMode);
  }, [darkMode]);

  const realizarLogin = () => {
    if (loginUsuario === USUARIO_MASTER) {
      if (loginSenha === senhaMaster) {
        setSessaoAtual({ usuario: USUARIO_MASTER, role: "master", nome: "RCPPJ" });
        setSistemaDesbloqueado(true);
        setLoginSenha("");
        toast({ title: "Acesso liberado", description: `Bem-vindo, ${USUARIO_MASTER}.` });
        return;
      }
      toast({ title: "Senha master inválida", variant: "destructive" });
      return;
    }
    const usuario = usuariosComuns.find((u) => u.usuario.toLowerCase() === loginUsuario.trim().toLowerCase());
    if (usuario && usuario.senha === loginSenha) {
      setSessaoAtual({ usuario: usuario.usuario, role: "comum", nome: usuario.nome });
      setSistemaDesbloqueado(true);
      setLoginSenha("");
      toast({ title: "Acesso liberado", description: `Bem-vindo, ${usuario.nome}.` });
      return;
    }
    toast({ title: "Usuário ou senha inválidos", variant: "destructive" });
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

  // Pergunta/resposta de segurança — definidas pelo próprio Master, nunca fixas no código
  const definirPerguntaSeguranca = () => {
    if (!novaSecQuestion.trim() || !novaSecAnswer.trim()) {
      toast({ title: "Preencha a pergunta e a resposta", variant: "destructive" });
      return;
    }
    setSecQuestion(novaSecQuestion.trim());
    setSecAnswer(novaSecAnswer.trim());
    setNovaSecQuestion("");
    setNovaSecAnswer("");
    toast({ title: "Pergunta de segurança definida" });
  };

  const redefinirSenhaComPergunta = () => {
    if (!secAnswer) {
      toast({ title: "Nenhuma pergunta de segurança configurada", variant: "destructive" });
      return;
    }
    if (respostaTentativa.trim().toLowerCase() !== secAnswer.toLowerCase()) {
      toast({ title: "Resposta de segurança incorreta", variant: "destructive" });
      return;
    }
    if (novaSenhaRecuperada.length < 6) {
      toast({ title: "Nova senha deve ter pelo menos 6 caracteres", variant: "destructive" });
      return;
    }
    save("senhaMaster", novaSenhaRecuperada);
    setSenhaMaster(novaSenhaRecuperada);
    setRespostaTentativa("");
    setNovaSenhaRecuperada("");
    setMostrarRecuperacao(false);
    toast({ title: "Senha master redefinida" });
  };

  // Gestão de usuários (apenas Master)
  const criarUsuarioComum = () => {
    if (!novoUsuarioNome.trim() || !novoUsuarioLogin.trim()) {
      toast({ title: "Preencha nome e usuário de login", variant: "destructive" });
      return;
    }
    const loginNormalizado = novoUsuarioLogin.trim();
    if (loginNormalizado.toLowerCase() === USUARIO_MASTER.toLowerCase() || usuariosComuns.some((u) => u.usuario.toLowerCase() === loginNormalizado.toLowerCase())) {
      toast({ title: "Já existe um usuário com esse login", variant: "destructive" });
      return;
    }
    const novo: SystemUser = {
      id: crypto.randomUUID(),
      nome: novoUsuarioNome.trim(),
      usuario: loginNormalizado,
      senha: "123456",
      mustChangePassword: true,
    };
    setUsuariosComuns((prev) => [...prev, novo]);
    setNovoUsuarioNome("");
    setNovoUsuarioLogin("");
    toast({ title: "Usuário criado", description: `Senha padrão: 123456` });
  };

  const resetarSenhaUsuario = (id: string) => {
    setUsuariosComuns((prev) => prev.map((u) => (u.id === id ? { ...u, senha: "123456", mustChangePassword: true } : u)));
    toast({ title: "Senha redefinida para 123456" });
  };

  const excluirUsuarioComum = (id: string) => {
    setUsuariosComuns((prev) => prev.filter((u) => u.id !== id));
    toast({ title: "Usuário excluído" });
  };

  // Troca de senha do usuário comum (exige confirmação da senha atual)
  const alterarSenhaUsuarioComum = () => {
    if (!sessaoAtual || sessaoAtual.role !== "comum") return;
    const usuarioAtual = usuariosComuns.find((u) => u.usuario === sessaoAtual.usuario);
    if (!usuarioAtual || usuarioAtual.senha !== senhaAtualComum) {
      toast({ title: "Senha atual incorreta", variant: "destructive" });
      return;
    }
    if (novaSenhaComum !== confirmarSenhaComum) {
      toast({ title: "Senhas não coincidem", variant: "destructive" });
      return;
    }
    if (novaSenhaComum.length < 6) {
      toast({ title: "Senha deve ter pelo menos 6 caracteres", variant: "destructive" });
      return;
    }
    setUsuariosComuns((prev) => prev.map((u) => (u.id === usuarioAtual.id ? { ...u, senha: novaSenhaComum, mustChangePassword: false } : u)));
    setSenhaAtualComum("");
    setNovaSenhaComum("");
    setConfirmarSenhaComum("");
    toast({ title: "Senha alterada com sucesso" });
  };

  // Troca obrigatória após reset pelo Master (login com senha padrão 123456)
  const confirmarPrimeiraTroca = () => {
    if (!sessaoAtual || sessaoAtual.role !== "comum") return;
    const usuarioAtual = usuariosComuns.find((u) => u.usuario === sessaoAtual.usuario);
    if (!usuarioAtual) return;
    if (primeiraTrocaSenha !== primeiraTrocaSenhaConfirmar) {
      toast({ title: "Senhas não coincidem", variant: "destructive" });
      return;
    }
    if (primeiraTrocaSenha.length < 6) {
      toast({ title: "Senha deve ter pelo menos 6 caracteres", variant: "destructive" });
      return;
    }
    setUsuariosComuns((prev) => prev.map((u) => (u.id === usuarioAtual.id ? { ...u, senha: primeiraTrocaSenha, mustChangePassword: false } : u)));
    setPrimeiraTrocaSenha("");
    setPrimeiraTrocaSenhaConfirmar("");
    toast({ title: "Senha definida com sucesso" });
  };

  const logout = () => {
    setSistemaDesbloqueado(false);
    setSessaoAtual(null);
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

  // Soma de horas já lançadas para um servidor numa data específica (considerando outros lançamentos)
  const horasJaLancadasNoDia = (matricula: string, data: string, excluirLancamentoId?: string) =>
    lancamentos
      .filter((l) => l.servidor.matricula === matricula && (!excluirLancamentoId || l.id !== excluirLancamentoId))
      .flatMap((l) => l.dias)
      .filter((d) => d.data === data)
      .reduce((sum, d) => sum + d.horas, 0);

  // CRUD de dias (lançamento por servidor)
  const adicionarDia = () => {
    const funcoesDisponiveis = getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria");
    setDias((d) => [...d, { data: periodo.inicio, horas: 8, funcao: funcoesDisponiveis[0] }]);
  };
  const removerDia = (index: number) => setDias((d) => d.filter((_, i) => i !== index));
  const atualizarDia = (index: number, patch: Partial<DiaTrabalho>) =>
    setDias((d) => d.map((dia, i) => (i === index ? { ...dia, ...patch } : dia)));

  const salvarLancamento = async () => {
    if (matriculasSelecionadas.length === 0) {
      toast({ title: "Selecione ao menos um servidor", variant: "destructive" });
      return;
    }
    if (editingId && matriculasSelecionadas.length > 1) {
      toast({ title: "Ao editar um lançamento, selecione apenas um servidor", variant: "destructive" });
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
    const foraPeriodo = diasValidos.find((d) => d.data < periodo.inicio || d.data > periodo.fim);
    if (foraPeriodo) {
      toast({ title: "Data fora do período", variant: "destructive" });
      return;
    }

    const conflitos: string[] = [];
    const bloqueados: string[] = [];
    let criados = 0;
    let atualizados = 0;

    setLancamentos((prev) => {
      const next = [...prev];
      matriculasSelecionadas.forEach((mat) => {
        const srv = servidores.find((s) => s.matricula === mat);
        if (!srv) return;

        const datasSet = new Set(datas);
        const conflito = next.some((l) => {
          if (editingId && l.id === editingId) return false;
          if (l.servidor.matricula !== srv.matricula) return false;
          return l.dias.some((d) => datasSet.has(d.data));
        });
        if (conflito) {
          conflitos.push(`${srv.nome} (${srv.matricula}) — já possui lançamento em uma dessas datas`);
          bloqueados.push(mat);
          return;
        }

        const diaExcede24h = diasValidos.find(
          (d) => horasJaLancadasNoDia(srv.matricula, d.data, editingId ?? undefined) + d.horas > 24
        );
        if (diaExcede24h) {
          conflitos.push(`${srv.nome} (${srv.matricula}) — ultrapassaria 24h em ${toBR(diaExcede24h.data)}`);
          bloqueados.push(mat);
          return;
        }

        const novo: Lancamento = {
          id: editingId ?? crypto.randomUUID(),
          operacaoId: selectedOp?.id,
          operacao: selectedOp?.tipo ?? "ordinaria",
          nomeOperacao: selectedOp?.label ?? "",
          periodo,
          servidor: srv,
          dias: diasValidos,
          createdAt: new Date().toISOString(),
        };

        const idx = next.findIndex((l) => l.id === novo.id);
        if (editingId && idx >= 0) {
          next[idx] = novo;
          atualizados++;
        } else {
          next.unshift(novo);
          criados++;
        }
      });
      return next;
    });

    if (conflitos.length > 0) {
      toast({
        title: `${conflitos.length} servidor(es) não lançado(s)`,
        description: conflitos.slice(0, 3).join(" | ") + (conflitos.length > 3 ? " ..." : ""),
        variant: "destructive",
      });
    }
    if (criados + atualizados > 0) {
      toast({ title: "Lançamento efetuado com sucesso", description: `${criados} novo(s), ${atualizados} atualizado(s)` });
    }

    const eraEdicao = !!editingId;
    setEditingId(null);
    setDias([{ data: periodo.inicio, horas: 8, funcao: getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria")[0] }]);
    setMatriculasSelecionadas(eraEdicao ? [] : bloqueados);
  };

  // Lançamento por data (vários servidores) — cada data com horas/função próprias
  const adicionarLoteData = () => {
    if (!loteNovaData) return;
    if (loteNovaData < periodo.inicio || loteNovaData > periodo.fim) {
      toast({ title: "Data fora do período da operação", variant: "destructive" });
      return;
    }
    if (loteDias.some((d) => d.data === loteNovaData)) {
      toast({ title: "Data já adicionada", variant: "destructive" });
      return;
    }
    setLoteDias((d) => [...d, { data: loteNovaData, horas: loteHoras, funcao: loteFuncao }].sort((a, b) => a.data.localeCompare(b.data)));
    setLoteNovaData("");
  };

  const atualizarLoteDia = (index: number, patch: Partial<DiaTrabalho>) =>
    setLoteDias((prev) => prev.map((d, i) => (i === index ? { ...d, ...patch } : d)));

  const removerLoteDia = (index: number) => setLoteDias((prev) => prev.filter((_, i) => i !== index));

  const salvarLancamentoPorData = () => {
    setLoteConflitos([]);
    if (loteDias.length === 0) {
      toast({ title: "Adicione pelo menos uma data", variant: "destructive" });
      return;
    }
    if (loteMatriculas.length === 0) {
      toast({ title: "Selecione pelo menos um servidor", variant: "destructive" });
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
        const datasConflitantes = loteDias.filter((d) => datasExistentes.has(d.data)).map((d) => d.data);
        if (datasConflitantes.length > 0) {
          // Servidor com data(s) já lançada(s): NÃO é lançado e o erro é apontado.
          conflitos.push(
            `${srv.nome} (${srv.matricula}) — já possui lançamento em ${datasConflitantes.map(toBR).join(", ")}`
          );
          bloqueados.push(mat);
          return;
        }

        const datasExcedendo24h = loteDias.filter((d) => horasJaLancadasNoDia(mat, d.data) + d.horas > 24).map((d) => d.data);
        if (datasExcedendo24h.length > 0) {
          conflitos.push(
            `${srv.nome} (${srv.matricula}) — ultrapassaria 24h em ${datasExcedendo24h.map(toBR).join(", ")}`
          );
          bloqueados.push(mat);
          return;
        }

        const novosDias: DiaTrabalho[] = loteDias.map((d) => ({ data: d.data, horas: d.horas, funcao: d.funcao }));
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
            operacaoId: selectedOp?.id,
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
        title: "Lançamento efetuado com sucesso",
        description: `${criados} novo(s), ${atualizados} atualizado(s)`,
      });
    }
    // Mantém selecionados apenas os que falharam, para correção
    setLoteMatriculas(bloqueados);
  };


  // Consolidação
  const [filtroHistorico, setFiltroHistorico] = useState<string>("todos");
  const [filtroTipoHistorico, setFiltroTipoHistorico] = useState<string>("todos");
  const [detalheId, setDetalheId] = useState<string | null>(null);

  const nomesOperacoesConsolidadas = useMemo(
    () => Array.from(new Set(lancamentos.map((l) => l.nomeOperacao))),
    [lancamentos]
  );

  // Módulo de Relatórios / Gerenciar Lançamentos — dados achatados por dia trabalhado
  const diasFlat = useMemo(
    () =>
      lancamentos.flatMap((l) =>
        l.dias.map((d) => ({
          lancamentoId: l.id,
          matricula: l.servidor.matricula,
          nome: l.servidor.nome,
          cpf: l.servidor.cpf || "",
          nomeOperacao: l.nomeOperacao,
          operacaoTipo: l.operacao,
          data: d.data,
          ano: d.data.slice(0, 4),
          funcao: d.funcao,
          horas: d.horas,
          valor: d.horas * valorHora(d.funcao, l.operacao),
          alimentacao: calcAlimentacao(l.operacao, d.horas),
        }))
      ),
    [lancamentos, valores, alimentacao]
  );

  const funcoesUsadas = useMemo(
    () => Array.from(new Set(diasFlat.map((d) => d.funcao))),
    [diasFlat]
  );
  const anosRelatorio = useMemo(
    () => Array.from(new Set(diasFlat.map((d) => d.ano))).sort(),
    [diasFlat]
  );

  const relatorioFiltrado = useMemo(
    () =>
      diasFlat.filter(
        (d) =>
          (relFuncao === "todas" || d.funcao === relFuncao) &&
          (relOperacao === "todas" || d.nomeOperacao === relOperacao) &&
          (relTipoOperacao === "todos" || d.operacaoTipo === relTipoOperacao) &&
          (relAno === "todos" || d.ano === relAno) &&
          (!relDataIni || d.data >= relDataIni) &&
          (!relDataFim || d.data <= relDataFim) &&
          (relServidor === "todos" || d.matricula === relServidor)
      ),
    [diasFlat, relFuncao, relOperacao, relTipoOperacao, relAno, relDataIni, relDataFim, relServidor]
  );

  const relatorioGrupos = useMemo(() => {
    const map = new Map<string, {
      matricula: string; nome: string; cpf: string; nomeOperacao: string;
      dias: { data: string; funcao: FuncaoID; horas: number; valor: number; alimentacao: number }[];
    }>();
    relatorioFiltrado.forEach((d) => {
      const key = `${d.matricula}::${d.nomeOperacao}`;
      if (!map.has(key)) map.set(key, { matricula: d.matricula, nome: d.nome, cpf: d.cpf, nomeOperacao: d.nomeOperacao, dias: [] });
      map.get(key)!.dias.push({ data: d.data, funcao: d.funcao, horas: d.horas, valor: d.valor, alimentacao: d.alimentacao });
    });
    return Array.from(map.values())
      .map((g) => {
        const dias = [...g.dias].sort((a, b) => a.data.localeCompare(b.data));
        return {
          ...g,
          dias,
          diasTrabalhados: dias.length,
          horas: dias.reduce((s, d) => s + d.horas, 0),
          valor: dias.reduce((s, d) => s + d.valor, 0),
          alimentacao: dias.reduce((s, d) => s + d.alimentacao, 0),
        };
      })
      .sort((a, b) => a.nome.localeCompare(b.nome) || a.nomeOperacao.localeCompare(b.nomeOperacao));
  }, [relatorioFiltrado]);

  const compareOrdenavel = (a: any, b: any) =>
    typeof a === "number" && typeof b === "number" ? a - b : String(a).localeCompare(String(b), "pt-BR");

  const toggleSort = (setSort: (fn: (prev: { campo: string; dir: SortDir }) => { campo: string; dir: SortDir }) => void, campo: string) =>
    setSort((prev) => (prev.campo === campo ? { campo, dir: prev.dir === "asc" ? "desc" : "asc" } : { campo, dir: "asc" }));

  const getRelSortValue = (r: (typeof relatorioGrupos)[number], campo: string) => {
    switch (campo) {
      case "matricula": return r.matricula;
      case "nomeOperacao": return r.nomeOperacao;
      case "diasTrabalhados": return r.diasTrabalhados;
      case "horas": return r.horas;
      case "valor": return r.valor;
      case "alimentacao": return r.alimentacao;
      case "total": return r.valor + r.alimentacao;
      case "nome": default: return r.nome;
    }
  };

  const relatorioGruposOrdenado = useMemo(() => {
    const arr = [...relatorioGrupos];
    arr.sort((a, b) => {
      const c = compareOrdenavel(getRelSortValue(a, relSort.campo), getRelSortValue(b, relSort.campo));
      return relSort.dir === "asc" ? c : -c;
    });
    return arr;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [relatorioGrupos, relSort]);

  // Gerenciar Lançamentos — busca e seleção múltipla (1 linha por dia trabalhado)
  const [buscaGerenciarNome, setBuscaGerenciarNome] = useState("");
  const [buscaGerenciarData, setBuscaGerenciarData] = useState("");
  const [buscaGerenciarDataIni, setBuscaGerenciarDataIni] = useState("");
  const [buscaGerenciarDataFim, setBuscaGerenciarDataFim] = useState("");
  const [linhasSelecionadas, setLinhasSelecionadas] = useState<Set<string>>(new Set());

  const gerenciarLinhas = useMemo(() => {
    const q = buscaGerenciarNome.trim().toLowerCase();
    return diasFlat
      .filter((d) => filtroHistorico === "todos" || d.nomeOperacao === filtroHistorico)
      .filter((d) => filtroTipoHistorico === "todos" || d.operacaoTipo === filtroTipoHistorico)
      .filter((d) => !q || d.nome.toLowerCase().includes(q))
      .filter((d) => !buscaGerenciarData || d.data === buscaGerenciarData)
      .filter((d) => !buscaGerenciarDataIni || d.data >= buscaGerenciarDataIni)
      .filter((d) => !buscaGerenciarDataFim || d.data <= buscaGerenciarDataFim);
  }, [diasFlat, filtroHistorico, filtroTipoHistorico, buscaGerenciarNome, buscaGerenciarData, buscaGerenciarDataIni, buscaGerenciarDataFim]);

  const getGerenciarSortValue = (d: (typeof gerenciarLinhas)[number], campo: string) => {
    switch (campo) {
      case "data": return d.data;
      case "funcao": return funcaoLabel(d.funcao);
      case "nomeOperacao": return d.nomeOperacao;
      case "horas": return d.horas;
      case "total": return d.valor + d.alimentacao;
      case "nome": default: return d.nome;
    }
  };

  const gerenciarLinhasOrdenadas = useMemo(() => {
    const arr = [...gerenciarLinhas];
    arr.sort((a, b) => {
      const c = compareOrdenavel(getGerenciarSortValue(a, gerenciarSort.campo), getGerenciarSortValue(b, gerenciarSort.campo));
      return gerenciarSort.dir === "asc" ? c : -c;
    });
    return arr;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [gerenciarLinhas, gerenciarSort]);

  const linhaKey = (lancamentoId: string, data: string) => `${lancamentoId}::${data}`;

  const excluirDiaLancamento = (lancamentoId: string, data: string) => {
    setLancamentos((prev) =>
      prev
        .map((l) => (l.id === lancamentoId ? { ...l, dias: l.dias.filter((d) => d.data !== data) } : l))
        .filter((l) => l.dias.length > 0)
    );
    setLinhasSelecionadas((prev) => {
      const next = new Set(prev);
      next.delete(linhaKey(lancamentoId, data));
      return next;
    });
    toast({ title: "Lançamento excluído" });
  };

  const excluirLinhasSelecionadas = () => {
    if (linhasSelecionadas.size === 0) return;
    const porLancamento = new Map<string, Set<string>>();
    linhasSelecionadas.forEach((key) => {
      const [lid, data] = key.split("::");
      if (!porLancamento.has(lid)) porLancamento.set(lid, new Set());
      porLancamento.get(lid)!.add(data);
    });
    setLancamentos((prev) =>
      prev
        .map((l) => {
          const datasRemover = porLancamento.get(l.id);
          if (!datasRemover) return l;
          return { ...l, dias: l.dias.filter((d) => !datasRemover.has(d.data)) };
        })
        .filter((l) => l.dias.length > 0)
    );
    toast({ title: `${linhasSelecionadas.size} lançamento(s) excluído(s)` });
    setLinhasSelecionadas(new Set());
  };

  const toggleLinhaSelecionada = (key: string) => {
    setLinhasSelecionadas((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  };

  const exportarGerenciarExcel = () => {
    if (gerenciarLinhasOrdenadas.length === 0) {
      toast({ title: "Nenhum lançamento para exportar", variant: "destructive" });
      return;
    }
    const headers = ["Servidor", "Matrícula", "CPF", "Operação", "Data Trabalhada", "Função", "Horas", "Valor Horas", "Alimentação", "Total"];
    const rows = gerenciarLinhasOrdenadas.map((d) => [
      d.nome, d.matricula, formatarCPF(d.cpf), d.nomeOperacao, toBR(d.data), funcaoLabel(d.funcao),
      d.horas, Number(d.valor.toFixed(2)), Number(d.alimentacao.toFixed(2)), Number((d.valor + d.alimentacao).toFixed(2)),
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lancamentos");
    XLSX.writeFile(wb, "Gerenciar_Lancamentos.xlsx");
    toast({ title: "Excel gerado" });
  };

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

  // PDF — Relatório Geral (por servidor/operação), com modo detalhado opcional
  const gerarRelatorioModuloPDF = () => {
    if (relatorioGruposOrdenado.length === 0) {
      toast({ title: "Nenhum dado para o filtro selecionado", variant: "destructive" });
      return;
    }
    const doc = new jsPDF({ orientation: "landscape" });
    const tipoLabels: Record<string, string> = { ordinaria: "Ordinária", extraordinaria: "Extraordinária", reveillon: "Reveillon", carnaval: "Carnaval" };
    const subtitulo = [
      relOperacao === "todas" ? "Todas as operações" : relOperacao,
      relTipoOperacao === "todos" ? "" : `Tipo: ${tipoLabels[relTipoOperacao] ?? relTipoOperacao}`,
      relAno === "todos" ? "Todos os anos" : relAno,
      relFuncao === "todas" ? "" : funcaoLabel(relFuncao as FuncaoID),
    ].filter(Boolean).join(" • ");

    let y = drawHeader(doc, "Relatório Geral de Servidores", subtitulo);

    const acc = relatorioGruposOrdenado.reduce((a, r) => ({ h: a.h + r.horas, v: a.v + r.valor, al: a.al + r.alimentacao }), { h: 0, v: 0, al: 0 });

    if (!relIncluirDetalhes) {
      const rows = relatorioGruposOrdenado.map((r) => [r.matricula, r.nome, r.nomeOperacao, String(r.diasTrabalhados), r.horas.toFixed(2), fmtBRL(r.valor), fmtBRL(r.alimentacao), fmtBRL(r.valor + r.alimentacao)]);
      autoTable(doc, {
        startY: y,
        head: [["Matrícula", "Nome", "Operação", "Dias Trabalhados", "Horas", "Valor Horas", "Alimentação", "Total"]],
        body: rows,
        foot: [["", "", "TOTAL GERAL", "", acc.h.toFixed(2), fmtBRL(acc.v), fmtBRL(acc.al), fmtBRL(acc.v + acc.al)]],
        margin: { left: 14, right: 14 },
        ...(tableTheme as any),
      });
    } else {
      // Modo detalhado: para cada servidor/operação, um bloco com as datas trabalhadas individualmente
      relatorioGruposOrdenado.forEach((g) => {
        const espacoNecessario = 20 + g.dias.length * 7;
        if (y + espacoNecessario > 190) {
          doc.addPage();
          y = 20;
        }
        doc.setFont("helvetica", "bold");
        doc.setFontSize(9.5);
        doc.setTextColor(26, 45, 84);
        doc.text(`${g.nome}  —  Matrícula ${g.matricula}  —  ${g.nomeOperacao}`, 14, y);
        doc.setTextColor(0, 0, 0);
        y += 4;

        const rows = g.dias.map((d) => [toBR(d.data), `${d.horas}h`, funcaoLabel(d.funcao), fmtBRL(d.valor), fmtBRL(d.alimentacao), fmtBRL(d.valor + d.alimentacao)]);
        autoTable(doc, {
          startY: y,
          head: [["Data Trabalhada", "Horas", "Função", "Valor Horas", "Alimentação", "Total do dia"]],
          body: rows,
          foot: [["", `${g.horas}h`, "", fmtBRL(g.valor), fmtBRL(g.alimentacao), fmtBRL(g.valor + g.alimentacao)]],
          margin: { left: 14, right: 14 },
          ...(tableTheme as any),
          bodyStyles: { ...(tableTheme.bodyStyles as any), fontSize: 8 },
          headStyles: { ...(tableTheme.headStyles as any), fontSize: 8 },
          footStyles: { ...(tableTheme.footStyles as any), fontSize: 8 },
        });
        y = ((doc as any).lastAutoTable?.finalY ?? y) + 10;
      });

      if (y > 180) { doc.addPage(); y = 20; }
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);
      doc.text(`TOTAL GERAL: ${acc.h.toFixed(2)}h — ${fmtBRL(acc.v + acc.al)}`, 14, y + 6);
    }

    drawFooters(doc, contatoRodape);
    doc.save("Relatorio_Geral_Servidores.pdf");
    toast({ title: "Relatório gerado" });
  };

  // Excel — Relatório Geral (por servidor/operação), com modo detalhado opcional
  const exportarRelatorioExcel = () => {
    if (relatorioGruposOrdenado.length === 0) {
      toast({ title: "Nenhum dado para o filtro selecionado", variant: "destructive" });
      return;
    }
    const wb = XLSX.utils.book_new();
    const tipoLabels: Record<string, string> = { ordinaria: "Ordinária", extraordinaria: "Extraordinária", reveillon: "Reveillon", carnaval: "Carnaval" };
    const filtrosAplicados: (string | number)[][] = [
      ["Filtros aplicados"],
      ["Operação", relOperacao === "todas" ? "Todas" : relOperacao],
      ["Tipo de Operação", relTipoOperacao === "todos" ? "Todos" : (tipoLabels[relTipoOperacao] ?? relTipoOperacao)],
      ["Ano", relAno === "todos" ? "Todos" : relAno],
      ["Função", relFuncao === "todas" ? "Todas" : funcaoLabel(relFuncao as FuncaoID)],
      ["Servidor", relServidor === "todos" ? "Todos" : (servidores.find((s) => s.matricula === relServidor)?.nome ?? relServidor)],
      ["Período (data trabalhada)", `${relDataIni ? toBR(relDataIni) : "—"} a ${relDataFim ? toBR(relDataFim) : "—"}`],
      [],
    ];

    if (!relIncluirDetalhes) {
      const headers = ["Matrícula", "Nome", "Operação", "Dias Trabalhados", "Horas", "Valor Horas", "Alimentação", "Total"];
      const rows = relatorioGruposOrdenado.map((r) => [
        r.matricula, r.nome, r.nomeOperacao, r.diasTrabalhados,
        Number(r.horas.toFixed(2)), Number(r.valor.toFixed(2)), Number(r.alimentacao.toFixed(2)), Number((r.valor + r.alimentacao).toFixed(2)),
      ]);
      const ws = XLSX.utils.aoa_to_sheet([...filtrosAplicados, headers, ...rows]);
      XLSX.utils.book_append_sheet(wb, ws, "Relatorio Geral");
    } else {
      const headers = ["Matrícula", "Nome", "Operação", "Data Trabalhada", "Função", "Horas", "Valor Horas", "Alimentação", "Total"];
      const rows: (string | number)[][] = [];
      relatorioGruposOrdenado.forEach((g) => {
        g.dias.forEach((d) => {
          rows.push([g.matricula, g.nome, g.nomeOperacao, toBR(d.data), funcaoLabel(d.funcao), d.horas, Number(d.valor.toFixed(2)), Number(d.alimentacao.toFixed(2)), Number((d.valor + d.alimentacao).toFixed(2))]);
        });
      });
      const ws = XLSX.utils.aoa_to_sheet([...filtrosAplicados, headers, ...rows]);
      XLSX.utils.book_append_sheet(wb, ws, "Relatorio Detalhado");
    }

    XLSX.writeFile(wb, "Relatorio_Geral_Servidores.xlsx");
    toast({ title: "Excel gerado" });
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
    const y = new Date(l.periodo.inicio + "T00:00:00").getFullYear();
    setAno(y);
    const op = l.operacaoId
      ? operacoes.find((o) => o.id === l.operacaoId)
      : operacoes.find((o) => o.periodo.inicio === l.periodo.inicio && o.periodo.fim === l.periodo.fim && o.tipo === l.operacao);
    if (op) setOperacaoId(op.id);
    setPeriodo(l.periodo);
    setMatriculasSelecionadas([l.servidor.matricula]);
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
                <Input
                  value={loginUsuario}
                  onChange={(e) => setLoginUsuario(e.target.value)}
                  placeholder="Usuário"
                />
              </div>
              <div className="space-y-2">
                <Label>Senha</Label>
                <Input
                  type="password"
                  value={loginSenha}
                  onChange={(e) => setLoginSenha(e.target.value)}
                  placeholder="Digite sua senha"
                  onKeyDown={(e) => e.key === "Enter" && realizarLogin()}
                />
              </div>
              <Button onClick={realizarLogin} className="w-full">Entrar</Button>

              {loginUsuario === USUARIO_MASTER && (
                <div className="text-center">
                  <button
                    type="button"
                    className="text-xs text-muted-foreground underline"
                    onClick={() => setMostrarRecuperacao((v) => !v)}
                  >
                    Esqueci minha senha
                  </button>
                </div>
              )}

              {loginUsuario === USUARIO_MASTER && mostrarRecuperacao && (
                <div className="space-y-3 border-t pt-4">
                  {!secAnswer ? (
                    <p className="text-xs text-muted-foreground">
                      Nenhuma pergunta de segurança foi configurada ainda. Entre com a senha master e configure uma em Configurações.
                    </p>
                  ) : (
                    <>
                      <h4 className="font-medium text-sm">Recuperar acesso</h4>
                      <div className="space-y-2">
                        <Label>Pergunta: {secQuestion}</Label>
                        <Input value={respostaTentativa} onChange={(e) => setRespostaTentativa(e.target.value)} placeholder="Digite a resposta" />
                      </div>
                      <div className="space-y-2">
                        <Label>Nova senha master</Label>
                        <Input type="password" value={novaSenhaRecuperada} onChange={(e) => setNovaSenhaRecuperada(e.target.value)} placeholder="Digite a nova senha" />
                      </div>
                      <Button onClick={redefinirSenhaComPergunta} variant="secondary" className="w-full">Redefinir senha master</Button>
                    </>
                  )}
                </div>
              )}
            </div>
          </CardContent>
        </Card>
      </div>
    );
  }

  // Troca obrigatória de senha após reset pelo Master (usuário comum logado com "123456")
  if (sessaoAtual?.role === "comum" && usuariosComuns.find((u) => u.usuario === sessaoAtual.usuario)?.mustChangePassword) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-background to-card/60">
        <Card className="w-full max-w-md">
          <CardHeader className="text-center">
            <CardTitle className="text-2xl mb-2">Defina sua senha</CardTitle>
            <p className="text-muted-foreground text-sm">Sua senha foi redefinida pelo administrador. Escolha uma nova senha para continuar.</p>
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              <div className="space-y-2">
                <Label>Nova senha</Label>
                <Input type="password" value={primeiraTrocaSenha} onChange={(e) => setPrimeiraTrocaSenha(e.target.value)} placeholder="Digite a nova senha" />
              </div>
              <div className="space-y-2">
                <Label>Confirmar nova senha</Label>
                <Input type="password" value={primeiraTrocaSenhaConfirmar} onChange={(e) => setPrimeiraTrocaSenhaConfirmar(e.target.value)} placeholder="Confirme a nova senha" />
              </div>
              <Button onClick={confirmarPrimeiraTroca} className="w-full">Definir senha e continuar</Button>
              <Button onClick={logout} variant="outline" className="w-full">Sair</Button>
            </div>
          </CardContent>
        </Card>
      </div>
    );
  }

  const funcoesDaOperacao = getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria");

  return (
    <div className="min-h-screen">
      <Header
        onLogout={logout}
        darkMode={darkMode}
        onToggleDarkMode={() => setDarkMode((v) => !v)}
        usuarioLogado={sessaoAtual?.nome}
      />

      <main className="container py-8">
        <Tabs value={tab} onValueChange={setTab} className="w-full">
          <TabsList className="flex w-full flex-wrap h-auto items-stretch gap-1.5 p-1.5">
            <TabsTrigger value="lancamentos" className="flex-1 min-w-[140px] px-3 py-2 whitespace-nowrap">Lançar por Servidor</TabsTrigger>
            <TabsTrigger value="por-data" className="flex-1 min-w-[140px] px-3 py-2 whitespace-nowrap">Lançar por Data</TabsTrigger>
            <TabsTrigger value="logs" className="flex-1 min-w-[140px] px-3 py-2 whitespace-nowrap">Gerenciar Lançamentos</TabsTrigger>
            <TabsTrigger value="relatorios" className="flex-1 min-w-[140px] px-3 py-2 whitespace-nowrap">Relatórios</TabsTrigger>
            <TabsTrigger value="rh" className="flex-1 min-w-[140px] px-3 py-2 whitespace-nowrap">Configurações</TabsTrigger>
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
                    <Label>Buscar operação</Label>
                    <Input value={buscaOperacao} onChange={(e) => setBuscaOperacao(e.target.value)} placeholder="Buscar por nome..." />
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
                          setLoteDias([]);
                          setDias((d) => (d.length === 0 ? [{ data: op.inicio, horas: 8, funcao: getFuncoesDisponiveisParaOperacao(op.tipo)[0] }] : d));
                        }
                      }}
                    >
                      <SelectTrigger><SelectValue placeholder={opcoesOperacao.length === 0 ? "Nenhuma operação cadastrada" : "Selecione"} /></SelectTrigger>
                      <SelectContent className="z-50">
                        {opcoesOperacao.map((op) => (<SelectItem key={op.id} value={op.id}>{op.label}</SelectItem>))}
                      </SelectContent>
                    </Select>
                    {opcoesOperacao.length === 0 && (
                      <div className="text-xs text-muted-foreground">Nenhuma operação cadastrada para {ano}. Cadastre em Configurações → Gerenciar Operações.</div>
                    )}
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
                    <Label>Servidor(es)</Label>
                    <MultiServerSelect servidores={servidores} selecionados={matriculasSelecionadas} onChange={setMatriculasSelecionadas} />
                  </div>
                </CardContent>
              </Card>

              <Card className="md:col-span-2">
                <CardHeader><CardTitle>Lançamento de Horas</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  {dias.map((d, i) => {
                    const tipoAtual = selectedOp?.tipo ?? "ordinaria";
                    const valorDiaHoras = d.horas * valorHora(d.funcao, tipoAtual);
                    const alimentacaoDia = calcAlimentacao(tipoAtual, d.horas);
                    return (
                    <div key={i} className="border rounded-md p-3 space-y-2">
                    <div className="flex items-center gap-2 flex-wrap">
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
                    <div className="text-xs text-muted-foreground pl-1">
                      {d.horas}h — {funcaoLabel(d.funcao)}: {fmtBRL(valorDiaHoras)} (horas) + {fmtBRL(alimentacaoDia)} (alimentação) = <span className="font-medium text-foreground">{fmtBRL(valorDiaHoras + alimentacaoDia)}</span> (total do dia)
                    </div>
                    </div>
                    );
                  })}

                  {dias.length > 0 && (
                    <div className="rounded-md border bg-muted/40 p-3 grid grid-cols-2 md:grid-cols-4 gap-2 text-sm">
                      <div><span className="text-muted-foreground">Total horas: </span><span className="font-medium">{resumoDias.horas}h</span></div>
                      <div><span className="text-muted-foreground">Valor horas: </span><span className="font-medium">{fmtBRL(resumoDias.valorHoras)}</span></div>
                      <div><span className="text-muted-foreground">Alimentação: </span><span className="font-medium">{fmtBRL(resumoDias.alimentacao)}</span></div>
                      <div><span className="text-muted-foreground">Total geral: </span><span className="font-semibold">{fmtBRL(resumoDias.total)}</span></div>
                    </div>
                  )}

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
                    <Label>Ano</Label>
                    <Select value={String(ano)} onValueChange={(v) => setAno(Number(v))}>
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent className="z-50">
                        {anosDisponiveis.map((y) => (<SelectItem key={y} value={String(y)}>{y}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label>Buscar operação</Label>
                    <Input value={buscaOperacao} onChange={(e) => setBuscaOperacao(e.target.value)} placeholder="Buscar por nome..." />
                  </div>

                  <div className="space-y-2">
                    <Label>Operação</Label>
                    <Select
                      value={operacaoId}
                      onValueChange={(v) => {
                        setOperacaoId(v);
                        const op = opcoesOperacao.find((o) => o.id === v);
                        if (op) { setPeriodo({ inicio: op.inicio, fim: op.fim }); setLoteDias([]); }
                      }}
                    >
                      <SelectTrigger><SelectValue placeholder={opcoesOperacao.length === 0 ? "Nenhuma operação cadastrada" : "Selecione"} /></SelectTrigger>
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
                    <div className="text-xs text-muted-foreground">A carga horária e função abaixo são usadas como padrão para cada nova data — depois de adicionada, cada dia pode ser ajustado individualmente na lista ao lado.</div>
                  </div>

                  <div className="space-y-2">
                    <Label>Carga horária padrão</Label>
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
                    <Label>Função padrão</Label>
                    <Select value={loteFuncao} onValueChange={(v: FuncaoID) => setLoteFuncao(v)}>
                      <SelectTrigger><SelectValue placeholder="Função" /></SelectTrigger>
                      <SelectContent className="z-50">
                        {funcoesDaOperacao.map((f) => (<SelectItem key={f} value={f}>{funcaoLabel(f)}</SelectItem>))}
                      </SelectContent>
                    </Select>
                    <div className="text-xs text-muted-foreground">Valor/h: {fmtBRL(valorHora(loteFuncao, selectedOp?.tipo ?? "ordinaria"))}</div>
                  </div>

                  <div className="space-y-2">
                    <Label>Dias adicionados a este lançamento</Label>
                    {loteDias.length === 0 && <div className="text-xs text-muted-foreground">Nenhuma data adicionada.</div>}
                    {loteDias.map((d, i) => {
                      const tipoAtual = selectedOp?.tipo ?? "ordinaria";
                      const valorDiaHoras = d.horas * valorHora(d.funcao, tipoAtual);
                      const alimentacaoDia = calcAlimentacao(tipoAtual, d.horas);
                      return (
                        <div key={d.data} className="border rounded-md p-3 space-y-2">
                          <div className="flex items-center gap-2 flex-wrap">
                            <div className="text-sm font-medium min-w-[90px]">{toBR(d.data)}</div>
                            <Input
                              type="number"
                              min={1}
                              max={selectedOp?.tipo === "carnaval" ? 24 : 12}
                              value={d.horas}
                              onChange={(e) => {
                                const limit = selectedOp?.tipo === "carnaval" ? 24 : 12;
                                atualizarLoteDia(i, { horas: Math.max(1, Math.min(Number(e.target.value || 0), limit)) });
                              }}
                              className="w-20"
                            />
                            <Select value={d.funcao} onValueChange={(v: FuncaoID) => atualizarLoteDia(i, { funcao: v })}>
                              <SelectTrigger className="w-48"><SelectValue placeholder="Função" /></SelectTrigger>
                              <SelectContent className="z-50">
                                {funcoesDaOperacao.map((funcao) => (<SelectItem key={funcao} value={funcao}>{funcaoLabel(funcao)}</SelectItem>))}
                              </SelectContent>
                            </Select>
                            <Button variant="destructive" size="sm" className="ms-auto" onClick={() => removerLoteDia(i)}>Remover</Button>
                          </div>
                          <div className="text-xs text-muted-foreground pl-1">
                            {d.horas}h — {funcaoLabel(d.funcao)}: {fmtBRL(valorDiaHoras)} (horas) + {fmtBRL(alimentacaoDia)} (alimentação) = <span className="font-medium text-foreground">{fmtBRL(valorDiaHoras + alimentacaoDia)}</span> (total do dia)
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  {loteDias.length > 0 && (
                    <div className="rounded-md border bg-muted/40 p-3 grid grid-cols-2 gap-2 text-sm">
                      <div><span className="text-muted-foreground">Total horas: </span><span className="font-medium">{resumoLoteDias.horas}h</span></div>
                      <div><span className="text-muted-foreground">Valor horas: </span><span className="font-medium">{fmtBRL(resumoLoteDias.valorHoras)}</span></div>
                      <div><span className="text-muted-foreground">Alimentação: </span><span className="font-medium">{fmtBRL(resumoLoteDias.alimentacao)}</span></div>
                      <div><span className="text-muted-foreground">Total geral (por servidor): </span><span className="font-semibold">{fmtBRL(resumoLoteDias.total)}</span></div>
                    </div>
                  )}
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
                    Serão criados {loteDias.length} dia(s), nas horas/função definidas em cada linha, para cada servidor selecionado. Servidores que já possuam lançamento em qualquer uma das datas (ou que ultrapassem 24h/dia) não são lançados e o sistema aponta o erro.
                  </div>

                  <Button onClick={salvarLancamentoPorData}>Lançar para os servidores selecionados</Button>
                </CardContent>
              </Card>
            </div>
          </TabsContent>

          {/* Gerenciar Lançamentos */}
          <TabsContent value="logs" className="mt-6">
            <Card>
              <CardHeader><CardTitle>Gerenciar Lançamentos</CardTitle></CardHeader>
              <CardContent>
                <div className="grid md:grid-cols-6 gap-3 mb-4">
                  <div className="space-y-1">
                    <Label className="text-xs">Operação</Label>
                    <Select value={filtroHistorico} onValueChange={setFiltroHistorico}>
                      <SelectTrigger><SelectValue placeholder="Todas" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todos">Todas as Operações</SelectItem>
                        {nomesOperacoesConsolidadas.map((n) => (<SelectItem key={n} value={n}>{n}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Tipo de Operação</Label>
                    <Select value={filtroTipoHistorico} onValueChange={setFiltroTipoHistorico}>
                      <SelectTrigger><SelectValue placeholder="Todos" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todos">Todos os tipos</SelectItem>
                        <SelectItem value="ordinaria">Ordinária</SelectItem>
                        <SelectItem value="extraordinaria">Extraordinária</SelectItem>
                        <SelectItem value="reveillon">Reveillon</SelectItem>
                        <SelectItem value="carnaval">Carnaval</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Nome do servidor</Label>
                    <Input value={buscaGerenciarNome} onChange={(e) => setBuscaGerenciarNome(e.target.value)} placeholder="Buscar por nome..." />
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Data trabalhada</Label>
                    <Input type="date" value={buscaGerenciarData} onChange={(e) => setBuscaGerenciarData(e.target.value)} />
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Intervalo (início)</Label>
                    <Input type="date" value={buscaGerenciarDataIni} onChange={(e) => setBuscaGerenciarDataIni(e.target.value)} />
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Intervalo (fim)</Label>
                    <Input type="date" value={buscaGerenciarDataFim} onChange={(e) => setBuscaGerenciarDataFim(e.target.value)} />
                  </div>
                </div>

                <div className="flex gap-2 mb-4 flex-wrap items-center">
                  <Button variant="secondary" onClick={exportarGerenciarExcel}>Exportar Excel</Button>
                  <Button
                    variant="destructive"
                    disabled={linhasSelecionadas.size === 0}
                    onClick={excluirLinhasSelecionadas}
                  >
                    Excluir selecionados {linhasSelecionadas.size > 0 ? `(${linhasSelecionadas.size})` : ""}
                  </Button>
                  {(buscaGerenciarNome || buscaGerenciarData || buscaGerenciarDataIni || buscaGerenciarDataFim || filtroHistorico !== "todos" || filtroTipoHistorico !== "todos") && (
                    <Button
                      variant="outline"
                      onClick={() => { setBuscaGerenciarNome(""); setBuscaGerenciarData(""); setBuscaGerenciarDataIni(""); setBuscaGerenciarDataFim(""); setFiltroHistorico("todos"); setFiltroTipoHistorico("todos"); }}
                    >
                      Limpar filtros
                    </Button>
                  )}
                </div>

                <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead className="w-10">
                          <Checkbox
                            checked={gerenciarLinhasOrdenadas.length > 0 && gerenciarLinhasOrdenadas.every((d) => linhasSelecionadas.has(linhaKey(d.lancamentoId, d.data)))}
                            onCheckedChange={(checked) => {
                              if (checked) {
                                setLinhasSelecionadas(new Set(gerenciarLinhasOrdenadas.map((d) => linhaKey(d.lancamentoId, d.data))));
                              } else {
                                setLinhasSelecionadas(new Set());
                              }
                            }}
                          />
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setGerenciarSort, "nome")}>
                          Servidor {gerenciarSort.campo === "nome" ? (gerenciarSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setGerenciarSort, "data")}>
                          Data Trabalhada {gerenciarSort.campo === "data" ? (gerenciarSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setGerenciarSort, "funcao")}>
                          Função {gerenciarSort.campo === "funcao" ? (gerenciarSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setGerenciarSort, "nomeOperacao")}>
                          Operação {gerenciarSort.campo === "nomeOperacao" ? (gerenciarSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setGerenciarSort, "horas")}>
                          Horas {gerenciarSort.campo === "horas" ? (gerenciarSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setGerenciarSort, "total")}>
                          Total {gerenciarSort.campo === "total" ? (gerenciarSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead>Ações</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {gerenciarLinhasOrdenadas.length === 0 && (
                        <TableRow><TableCell colSpan={8} className="text-center text-muted-foreground">Nenhum lançamento encontrado.</TableCell></TableRow>
                      )}
                      {gerenciarLinhasOrdenadas.map((d) => {
                        const key = linhaKey(d.lancamentoId, d.data);
                        const l = lancamentos.find((x) => x.id === d.lancamentoId);
                        return (
                          <TableRow key={key}>
                            <TableCell>
                              <Checkbox checked={linhasSelecionadas.has(key)} onCheckedChange={() => toggleLinhaSelecionada(key)} />
                            </TableCell>
                            <TableCell>{d.nome} ({d.matricula})</TableCell>
                            <TableCell>{toBR(d.data)}</TableCell>
                            <TableCell>{funcaoLabel(d.funcao)}</TableCell>
                            <TableCell>{d.nomeOperacao}</TableCell>
                            <TableCell>{d.horas}h</TableCell>
                            <TableCell>{fmtBRL(d.valor + d.alimentacao)}</TableCell>
                            <TableCell className="flex gap-2 flex-wrap">
                              {l && <Button size="sm" onClick={() => editarLancamento(l)}>Editar</Button>}
                              <Button size="sm" variant="destructive" onClick={() => excluirDiaLancamento(d.lancamentoId, d.data)}>Excluir</Button>
                            </TableCell>
                          </TableRow>
                        );
                      })}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Relatórios de Operações e Servidores */}
          <TabsContent value="relatorios" className="mt-6">
            <Card>
              <CardHeader><CardTitle>Relatório Geral</CardTitle></CardHeader>
              <CardContent>
                <div className="grid md:grid-cols-5 gap-3 mb-4">
                  <div className="space-y-1">
                    <Label className="text-xs">Servidor</Label>
                    <Select value={relServidor} onValueChange={setRelServidor}>
                      <SelectTrigger><SelectValue placeholder="Todos" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todos">Todos os servidores</SelectItem>
                        {[...servidores].sort((a, b) => a.nome.localeCompare(b.nome)).map((s) => (
                          <SelectItem key={s.matricula} value={s.matricula}>{s.nome}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Função</Label>
                    <Select value={relFuncao} onValueChange={setRelFuncao}>
                      <SelectTrigger><SelectValue placeholder="Todas" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todas">Todas as funções</SelectItem>
                        {funcoesUsadas.map((f) => (<SelectItem key={f} value={f}>{funcaoLabel(f)}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Operação</Label>
                    <Select value={relOperacao} onValueChange={setRelOperacao}>
                      <SelectTrigger><SelectValue placeholder="Todas" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todas">Todas</SelectItem>
                        {nomesOperacoesConsolidadas.map((n) => (<SelectItem key={n} value={n}>{n}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Tipo de Operação</Label>
                    <Select value={relTipoOperacao} onValueChange={setRelTipoOperacao}>
                      <SelectTrigger><SelectValue placeholder="Todos" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todos">Todos os tipos</SelectItem>
                        <SelectItem value="ordinaria">Ordinária</SelectItem>
                        <SelectItem value="extraordinaria">Extraordinária</SelectItem>
                        <SelectItem value="reveillon">Reveillon</SelectItem>
                        <SelectItem value="carnaval">Carnaval</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-1">
                    <Label className="text-xs">Ano</Label>
                    <Select value={relAno} onValueChange={setRelAno}>
                      <SelectTrigger><SelectValue placeholder="Todos" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todos">Todos os anos</SelectItem>
                        {anosRelatorio.map((a) => (<SelectItem key={a} value={a}>{a}</SelectItem>))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div className="space-y-1">
                      <Label className="text-xs">Data trabalhada (início)</Label>
                      <Input type="date" value={relDataIni} onChange={(e) => setRelDataIni(e.target.value)} />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Data trabalhada (fim)</Label>
                      <Input type="date" value={relDataFim} onChange={(e) => setRelDataFim(e.target.value)} />
                    </div>
                  </div>
                </div>

                <div className="flex gap-3 mb-4 flex-wrap items-center">
                  <Button onClick={gerarRelatorioModuloPDF}>Exportar PDF</Button>
                  <Button variant="secondary" onClick={exportarRelatorioExcel}>Exportar Excel</Button>
                  <label className="flex items-center gap-2 text-sm cursor-pointer">
                    <Checkbox checked={relIncluirDetalhes} onCheckedChange={(v) => setRelIncluirDetalhes(!!v)} />
                    Incluir datas trabalhadas detalhadas por servidor
                  </label>
                  <Button
                    variant="outline"
                    onClick={() => { setRelFuncao("todas"); setRelOperacao("todas"); setRelTipoOperacao("todos"); setRelAno("todos"); setRelDataIni(""); setRelDataFim(""); setRelServidor("todos"); }}
                  >
                    Limpar filtros (Geral)
                  </Button>
                </div>

                <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "matricula")}>
                          Matrícula {relSort.campo === "matricula" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "nome")}>
                          Nome {relSort.campo === "nome" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "nomeOperacao")}>
                          Operação {relSort.campo === "nomeOperacao" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "diasTrabalhados")}>
                          Dias Trabalhados {relSort.campo === "diasTrabalhados" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "horas")}>
                          Horas {relSort.campo === "horas" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "valor")}>
                          Valor Horas {relSort.campo === "valor" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "alimentacao")}>
                          Alimentação {relSort.campo === "alimentacao" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                        <TableHead className="cursor-pointer select-none" onClick={() => toggleSort(setRelSort, "total")}>
                          Total {relSort.campo === "total" ? (relSort.dir === "asc" ? "▲" : "▼") : ""}
                        </TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {relatorioGruposOrdenado.length === 0 && (
                        <TableRow><TableCell colSpan={8} className="text-center text-muted-foreground">Nenhum dado para o filtro selecionado.</TableCell></TableRow>
                      )}
                      {relatorioGruposOrdenado.map((r) => (
                        <TableRow key={`${r.matricula}::${r.nomeOperacao}`}>
                          <TableCell>{r.matricula}</TableCell>
                          <TableCell>{r.nome}</TableCell>
                          <TableCell>{r.nomeOperacao}</TableCell>
                          <TableCell>{r.diasTrabalhados}</TableCell>
                          <TableCell>{r.horas.toFixed(2)}</TableCell>
                          <TableCell>{fmtBRL(r.valor)}</TableCell>
                          <TableCell>{fmtBRL(r.alimentacao)}</TableCell>
                          <TableCell>{fmtBRL(r.valor + r.alimentacao)}</TableCell>
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
            {sessaoAtual?.role === "comum" ? (
              <div className="grid gap-6 max-w-xl">
                <Card>
                  <CardHeader><CardTitle>Minha conta</CardTitle></CardHeader>
                  <CardContent className="space-y-4">
                    <div className="text-sm text-muted-foreground">Usuário: <span className="font-medium text-foreground">{sessaoAtual.nome}</span></div>
                    <div className="space-y-3">
                      <h4 className="font-medium">Alterar minha senha</h4>
                      <div className="space-y-2">
                        <Label>Senha atual</Label>
                        <Input type="password" value={senhaAtualComum} onChange={(e) => setSenhaAtualComum(e.target.value)} placeholder="Digite a senha atual" />
                      </div>
                      <div className="space-y-2">
                        <Label>Nova senha</Label>
                        <Input type="password" value={novaSenhaComum} onChange={(e) => setNovaSenhaComum(e.target.value)} placeholder="Digite a nova senha" />
                      </div>
                      <div className="space-y-2">
                        <Label>Confirmar nova senha</Label>
                        <Input type="password" value={confirmarSenhaComum} onChange={(e) => setConfirmarSenhaComum(e.target.value)} placeholder="Confirme a nova senha" />
                      </div>
                      <Button onClick={alterarSenhaUsuarioComum} variant="secondary" className="w-full">Alterar senha</Button>
                    </div>
                    <Button onClick={logout} variant="destructive" className="w-full">Sair</Button>
                  </CardContent>
                </Card>
              </div>
            ) : (
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
                  <div className="space-y-3 border-t pt-4">
                    <h4 className="font-medium">Pergunta de segurança (recuperação de senha)</h4>
                    {secQuestion && (
                      <div className="text-xs text-muted-foreground">Pergunta atual: <span className="text-foreground">{secQuestion}</span></div>
                    )}
                    <div className="space-y-2">
                      <Label>Nova pergunta</Label>
                      <Input value={novaSecQuestion} onChange={(e) => setNovaSecQuestion(e.target.value)} placeholder="Ex.: Qual o nome do seu primeiro animal de estimação?" />
                    </div>
                    <div className="space-y-2">
                      <Label>Nova resposta</Label>
                      <Input value={novaSecAnswer} onChange={(e) => setNovaSecAnswer(e.target.value)} placeholder="Resposta" />
                    </div>
                    <Button onClick={definirPerguntaSeguranca} variant="secondary" className="w-full">Salvar pergunta de segurança</Button>
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
                    {(["ordinaria", "extraordinaria", "reveillon", "carnaval"] as OperacaoTipo[]).map((tipo) => (
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

                  <div className="mt-6 grid md:grid-cols-4 gap-6">
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
                      <h5 className="text-sm font-medium mb-2">Alimentação — Extraordinária</h5>
                      <Label className="text-xs">Valor por hora (mínimo 8h)</Label>
                      <Input
                        type="number"
                        step="0.01"
                        value={(alimentacao as any).extraordinaria?.valorHora ?? 2}
                        onChange={(e) => setAlimentacao((prev: any) => ({ ...prev, extraordinaria: { ...prev.extraordinaria, valorHora: Number(e.target.value) } }))}
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
                <CardHeader><CardTitle>Gerenciar Operações</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  <div className="grid md:grid-cols-5 gap-2 items-end">
                    <div className="space-y-1 md:col-span-2">
                      <Label className="text-xs">Nome</Label>
                      <Input value={novaOpNome} onChange={(e) => setNovaOpNome(e.target.value)} placeholder="Ex.: Operação 21/05 a 20/06/2026" />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Tipo</Label>
                      <Select value={novaOpTipo} onValueChange={(v: OperacaoTipo) => setNovaOpTipo(v)}>
                        <SelectTrigger><SelectValue /></SelectTrigger>
                        <SelectContent className="z-50">
                          <SelectItem value="ordinaria">Ordinária</SelectItem>
                          <SelectItem value="extraordinaria">Extraordinária</SelectItem>
                          <SelectItem value="reveillon">Reveillon</SelectItem>
                          <SelectItem value="carnaval">Carnaval</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Início</Label>
                      <Input type="date" value={novaOpInicio} onChange={(e) => setNovaOpInicio(e.target.value)} />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Fim</Label>
                      <Input type="date" value={novaOpFim} onChange={(e) => setNovaOpFim(e.target.value)} />
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <Button onClick={salvarOperacao}>{editandoOpId ? "Salvar alterações" : "+ Criar operação"}</Button>
                    {editandoOpId && <Button variant="outline" onClick={cancelarEdicaoOperacao}>Cancelar edição</Button>}
                  </div>

                  <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>Nome</TableHead>
                          <TableHead>Tipo</TableHead>
                          <TableHead>Período</TableHead>
                          <TableHead>Ações</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {operacoes.length === 0 && (
                          <TableRow><TableCell colSpan={4} className="text-center text-muted-foreground">Nenhuma operação cadastrada.</TableCell></TableRow>
                        )}
                        {[...operacoes].sort((a, b) => b.periodo.inicio.localeCompare(a.periodo.inicio)).map((o) => (
                          <TableRow key={o.id}>
                            <TableCell>{o.nome}</TableCell>
                            <TableCell className="capitalize">{o.tipo}</TableCell>
                            <TableCell>{toBR(o.periodo.inicio)} a {toBR(o.periodo.fim)}</TableCell>
                            <TableCell className="flex gap-2 flex-wrap">
                              <Button size="sm" variant="secondary" onClick={() => editarOperacao(o)}>Editar</Button>
                              <Button size="sm" variant="destructive" onClick={() => excluirOperacao(o.id)}>Excluir</Button>
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
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
                <CardHeader><CardTitle>Gerenciamento de Usuários</CardTitle></CardHeader>
                <CardContent className="space-y-4">
                  <div className="flex gap-2 flex-wrap items-end">
                    <div className="space-y-1">
                      <Label className="text-xs">Nome completo</Label>
                      <Input value={novoUsuarioNome} onChange={(e) => setNovoUsuarioNome(e.target.value)} placeholder="Nome do novo usuário" className="w-56" />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Usuário (login)</Label>
                      <Input value={novoUsuarioLogin} onChange={(e) => setNovoUsuarioLogin(e.target.value)} placeholder="login.usuario" className="w-48" />
                    </div>
                    <Button onClick={criarUsuarioComum}>+ Criar usuário</Button>
                  </div>
                  <div className="text-xs text-muted-foreground">Todo usuário novo é criado com a senha padrão "123456" e deve trocá-la no primeiro acesso.</div>

                  <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>Nome</TableHead>
                          <TableHead>Usuário</TableHead>
                          <TableHead>Status</TableHead>
                          <TableHead>Ações</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {usuariosComuns.length === 0 && (
                          <TableRow><TableCell colSpan={4} className="text-center text-muted-foreground">Nenhum usuário comum cadastrado.</TableCell></TableRow>
                        )}
                        {usuariosComuns.map((u) => (
                          <TableRow key={u.id}>
                            <TableCell>{u.nome}</TableCell>
                            <TableCell>{u.usuario}</TableCell>
                            <TableCell>
                              {u.mustChangePassword ? (
                                <Badge variant="outline" className="bg-amber-100 text-amber-900 border-amber-300">Aguardando 1º acesso</Badge>
                              ) : (
                                <Badge variant="outline" className="bg-emerald-100 text-emerald-900 border-emerald-300">Ativo</Badge>
                              )}
                            </TableCell>
                            <TableCell className="flex gap-2 flex-wrap">
                              <Button size="sm" variant="secondary" onClick={() => resetarSenhaUsuario(u.id)}>Resetar senha</Button>
                              <Button size="sm" variant="destructive" onClick={() => excluirUsuarioComum(u.id)}>Excluir</Button>
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                </CardContent>
              </Card>

              <Card className="md:col-span-2 border-destructive/40">
                <CardHeader><CardTitle className="text-destructive">Zona de Perigo</CardTitle></CardHeader>
                <CardContent className="space-y-3">
                  <p className="text-sm text-muted-foreground">
                    Remove todos os lançamentos de horas do sistema. O cadastro de servidores é preservado. Essa ação não pode ser desfeita.
                  </p>
                  <Button variant="destructive" onClick={zerarLancamentos}>Zerar Dados (todos os lançamentos)</Button>
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
            )}
          </TabsContent>
        </Tabs>
      </main>
    </div>
  );
};

export default Index;

import { useEffect, useMemo, useState } from "react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { toast } from "@/hooks/use-toast";
import ServerCombobox, { Servidor } from "@/components/ServerCombobox";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

// Tipos
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

type FuncaoID = 
  | "coordenador_geral" | "coordenador_setorial" | "supervisor" | "agente_fiscalizacao" 
  | "guarda_civil" | "agente_operacoes" | "assistente_tecnico" | "motorista"
  | "coordenador" | "supervisor1" | "supervisor2" | "apoio_adm";

// Valores por operação
type ValoresOperacao = {
  [key in FuncaoID]?: number;
};

type AlimentacaoConfig = {
  [key: number]: number; // horas -> valor
};

const valoresDefault = {
  ordinaria: {
    coordenador: 20.50,
    supervisor: 15.50,
    agente_fiscalizacao: 12.00,
    apoio_adm: 10.00,
  } as ValoresOperacao,
  reveillon: {
    coordenador: 26.22,
    supervisor1: 25.07,
    supervisor2: 23.85,
    agente_fiscalizacao: 22.79,
    apoio_adm: 10.00,
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
    9: 39.01, 10: 42.10, 11: 43.08, 12: 47.00, 13: 50.05, 14: 53.10, 15: 56.15, 
    16: 57.13, 17: 58.11, 18: 59.08, 19: 59.09, 20: 62.14, 21: 65.19, 22: 68.24, 
    23: 71.29, 24: 74.64
  },
} as any;

// Operações dinâmicas por ano
type OpItem = { id: string; label: string; inicio: string; fim: string; tipo: OperacaoTipo };

// Utils de datas
const pad2 = (n: number) => n.toString().padStart(2, "0");
const toISO = (d: Date) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;

// Formatações BR
const toBR = (iso: string) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};
const fmtBRL = (n: number) => new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(n);
const clampToPeriodo = (dateIso: string, inicio: string, fim: string) => (dateIso < inicio ? inicio : dateIso > fim ? fim : dateIso);

// Páscoa (Meeus/Jones/Butcher)
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
  const month = Math.floor((h + l - 7 * m + 114) / 31); // 3=March, 4=April
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
  // 01/01 a 20/01
  ops.push({
    id: `ordinaria-01-20-jan-${year}`,
    label: `01/01 a 20/01/${year}`,
    inicio: `${year}-01-01`,
    fim: `${year}-01-20`,
    tipo: "ordinaria",
  });
  // 21/01 a 20/02 ... 21/11 a 20/12
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
  // 21/12 a 31/12
  ops.push({
    id: `ordinaria-21-31-dez-${year}`,
    label: `21/12 a 31/12/${year}`,
    inicio: `${year}-12-21`,
    fim: `${year}-12-31`,
    tipo: "ordinaria",
  });
  // Reveillon
  ops.push({
    id: `reveillon-${year}`,
    label: `Reveillon ${year}`,
    inicio: `${year}-12-24`,
    fim: `${year + 1}-01-01`,
    tipo: "reveillon",
  });
  // Carnaval
  const car = carnavalPeriodo(year);
  ops.push({
    id: `carnaval-${year}`,
    label: `Carnaval ${year}`,
    inicio: car.inicio,
    fim: car.fim,
    tipo: "carnaval",
  });
  return ops;
};

// Helpers de persistência
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

const Index = () => {
  // SEO básico
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

  // Administração (Banco de Dados) e Valores
  const [adminUser, setAdminUser] = useState<string>(() => load<string>("adminUser", "RCPPJ"));
  const [adminPass, setAdminPass] = useState<string>(() => load<string>("adminPass", "ruicpj@123"));
  const [adminLogged, setAdminLogged] = useState<boolean>(() => load<boolean>("adminLogged", false));
  const [contatos, setContatos] = useState<{ telefone1: string; telefone2: string }>(() => load("contatosSetor", { telefone1: "", telefone2: "" }));
  const [valores, setValores] = useState(() => load("valoresOperacoes", valoresDefault));
  const [alimentacao, setAlimentacao] = useState(() => load("alimentacaoOperacoes", alimentacaoDefault));
  
  // Sistema de perfis de usuário
  type PerfilUsuario = {
    id: string;
    nomeSetor: string;
    nomeGestor: string;
    matricula: string;
    login: string;
    senha: string;
    ativo: boolean;
  };
  
  const [perfisUsuario, setPerfisUsuario] = useState<PerfilUsuario[]>(() => load("perfisUsuario", []));
  const [usuarioLogado, setUsuarioLogado] = useState<PerfilUsuario | null>(() => load("usuarioLogado", null));
  
  // Estado para criação de novo perfil
  const [novoPerfil, setNovoPerfil] = useState({
    nomeSetor: "",
    nomeGestor: "",
    matricula: "",
    login: "",
    senha: ""
  });
  
  // Estado para login de usuário regular
  const [loginRegular, setLoginRegular] = useState({ login: "", senha: "" });
  
  // Estado para alterar nome de usuário
  const [novoUsuario, setNovoUsuario] = useState("");
  
  const SECURITY_QUESTION = "My birthday?";
  const SECURITY_ANSWER = "27 de Setembro";

  // Estados do sistema de login unificado
  const [loginUsuario, setLoginUsuario] = useState("");
  const [loginSenha, setLoginSenha] = useState("");
  const [novaSenha, setNovaSenha] = useState("");
  const [confirmarSenha, setConfirmarSenha] = useState("");
  const [respostaSeguranca, setRespostaSeguranca] = useState("");
  const [novaSenhaSeguranca, setNovaSenhaSeguranca] = useState("");
  const [sistemaDesbloqueado, setSistemaDesbloqueado] = useState<boolean>(() => load<boolean>("sistemaDesbloqueado", false));
  const [usuarioTipo, setUsuarioTipo] = useState<"admin" | "regular" | null>(() => load<string>("usuarioTipo", null) as "admin" | "regular" | null);
  
  // Opções de operação por ano
  const opcoesOperacao = useMemo(() => buildOpcoes(ano), [ano]);
  const selectedOp = useMemo(() => opcoesOperacao.find((o) => o.id === operacaoId), [opcoesOperacao, operacaoId]);
  
  // Anos disponíveis - apenas a partir de 2025
  const anosDisponiveis = useMemo(() => {
    const cy = new Date().getFullYear();
    const startYear = Math.max(2025, cy);
    return Array.from({ length: 15 }, (_, i) => startYear + i);
  }, []);
  
  // Ajustar seleção quando o ano muda
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
  const getValue = (obj: any, keys: string[]) => {
    for (const k of keys) {
      if (obj[k] !== undefined) return obj[k];
    }
    return "";
  };
  const formatarCPF = (cpf?: string) => {
    if (!cpf) return "";
    const s = cpf.toString().replace(/\D/g, "");
    return s.length === 11 ? s.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") : cpf;
  };
  
  const importarExcel = async () => {
    if (!adminLogged) {
      toast({ title: "Acesso restrito", description: "Somente administrador pode importar a base.", variant: "destructive" });
      return;
    }
    if (!excelFile) {
      toast({ title: "Selecione um arquivo Excel", variant: "destructive" });
      return;
    }
    try {
      const data = await excelFile.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json: any[] = XLSX.utils.sheet_to_json(ws);
      const lista: Servidor[] = json
        .map((r) => ({
          matricula: getValue(r, ["Matrícula", "matricula", "MATRICULA", "MATRÍCULA"]),
          nome: getValue(r, ["Nome", "nome", "NOME"]),
          cpf: getValue(r, ["CPF", "cpf"]),
          cargo: getValue(r, ["Cargo", "cargo", "CARGO"]),
        }))
        .filter((s) => s.matricula && s.nome);
      if (lista.length === 0) {
        toast({ title: "Nenhum servidor encontrado", variant: "destructive" });
        return;
      }
      setServidores(lista);
      toast({ title: `${lista.length} servidores importados` });
    } catch (e) {
      console.error(e);
      toast({ title: "Erro ao importar", variant: "destructive" });
    }
  };

  // Helpers de cálculo e recursos
  const valorHora = (f: FuncaoID, tipo: OperacaoTipo) => {
    return valores[tipo]?.[f] || 0;
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

  // Persistência
  useEffect(() => save("servidores", servidores), [servidores]);
  useEffect(() => save("lancamentos", lancamentos), [lancamentos]);
  useEffect(() => save("adminUser", adminUser), [adminUser]);
  useEffect(() => save("adminPass", adminPass), [adminPass]);
  useEffect(() => save("adminLogged", adminLogged), [adminLogged]);
  useEffect(() => save("contatosSetor", contatos), [contatos]);
  useEffect(() => save("valoresOperacoes", valores), [valores]);
  useEffect(() => save("alimentacaoOperacoes", alimentacao), [alimentacao]);
  useEffect(() => save("perfisUsuario", perfisUsuario), [perfisUsuario]);
  useEffect(() => save("usuarioLogado", usuarioLogado), [usuarioLogado]);
  useEffect(() => save("sistemaDesbloqueado", sistemaDesbloqueado), [sistemaDesbloqueado]);
  useEffect(() => save("usuarioTipo", usuarioTipo), [usuarioTipo]);

  // Sistema de login unificado
  const realizarLogin = () => {
    // Verificar se é login de administrador
    if (loginUsuario === adminUser && loginSenha === adminPass) {
      setAdminLogged(true);
      setSistemaDesbloqueado(true);
      setUsuarioTipo("admin");
      setLoginUsuario("");
      setLoginSenha("");
      toast({ title: "Login de administrador realizado com sucesso" });
      return;
    }
    
    // Verificar se é login de usuário regular
    const perfil = perfisUsuario.find(p => 
      p.login === loginUsuario && 
      p.senha === loginSenha && 
      p.ativo
    );
    
    if (perfil) {
      setUsuarioLogado(perfil);
      setSistemaDesbloqueado(true);
      setUsuarioTipo("regular");
      setLoginUsuario("");
      setLoginSenha("");
      toast({ title: `Bem-vindo, ${perfil.nomeSetor}` });
      return;
    }
    
    toast({ title: "Credenciais inválidas", variant: "destructive" });
  };

  const alterarSenha = () => {
    if (!adminLogged) {
      toast({ title: "Faça login primeiro", variant: "destructive" });
      return;
    }
    if (novaSenha !== confirmarSenha) {
      toast({ title: "Senhas não coincidem", variant: "destructive" });
      return;
    }
    if (novaSenha.length < 6) {
      toast({ title: "Senha deve ter pelo menos 6 caracteres", variant: "destructive" });
      return;
    }
    setAdminPass(novaSenha);
    setNovaSenha("");
    setConfirmarSenha("");
    toast({ title: "Senha alterada com sucesso" });
  };

  const redefinirSenha = () => {
    if (respostaSeguranca !== SECURITY_ANSWER) {
      toast({ title: "Resposta de segurança incorreta", variant: "destructive" });
      return;
    }
    if (novaSenhaSeguranca.length < 6) {
      toast({ title: "Nova senha deve ter pelo menos 6 caracteres", variant: "destructive" });
      return;
    }
    setAdminPass(novaSenhaSeguranca);
    setRespostaSeguranca("");
    setNovaSenhaSeguranca("");
    toast({ title: "Senha redefinida com sucesso" });
  };

  const logout = () => {
    setAdminLogged(false);
    setUsuarioLogado(null);
    setSistemaDesbloqueado(false);
    setUsuarioTipo(null);
    toast({ title: "Logout realizado" });
  };

  // Funções do sistema de perfis
  const criarPerfil = () => {
    if (!adminLogged) {
      toast({ title: "Acesso restrito", description: "Somente administrador pode criar perfis.", variant: "destructive" });
      return;
    }
    
    if (!novoPerfil.nomeSetor || !novoPerfil.nomeGestor || !novoPerfil.matricula || !novoPerfil.login || !novoPerfil.senha) {
      toast({ title: "Preencha todos os campos", variant: "destructive" });
      return;
    }
    
    if (perfisUsuario.some(p => p.login === novoPerfil.login)) {
      toast({ title: "Login já existe", description: "Escolha outro nome de usuário.", variant: "destructive" });
      return;
    }
    
    const perfil: PerfilUsuario = {
      id: crypto.randomUUID(),
      ...novoPerfil,
      ativo: true
    };
    
    setPerfisUsuario(prev => [...prev, perfil]);
    setNovoPerfil({ nomeSetor: "", nomeGestor: "", matricula: "", login: "", senha: "" });
    toast({ title: "Perfil criado com sucesso" });
  };

  const togglePerfilAtivo = (id: string) => {
    setPerfisUsuario(prev => prev.map(p => 
      p.id === id ? { ...p, ativo: !p.ativo } : p
    ));
  };

  const excluirPerfil = (id: string) => {
    setPerfisUsuario(prev => prev.filter(p => p.id !== id));
    toast({ title: "Perfil excluído" });
  };

  // Remover funções duplicadas pois agora o login é unificado

  const alterarNomeUsuario = () => {
    if (!adminLogged) {
      toast({ title: "Faça login primeiro", variant: "destructive" });
      return;
    }
    
    if (!novoUsuario.trim()) {
      toast({ title: "Digite um nome de usuário válido", variant: "destructive" });
      return;
    }
    
    setAdminUser(novoUsuario);
    setNovoUsuario("");
    toast({ title: "Nome de usuário alterado com sucesso" });
  };

  const salvarAlteracoes = () => {
    if (!adminLogged) {
      toast({ title: "Acesso restrito", variant: "destructive" });
      return;
    }
    
    // Força salvamento no localStorage
    save("valoresOperacoes", valores);
    save("alimentacaoOperacoes", alimentacao);
    save("perfisUsuario", perfisUsuario);
    save("adminUser", adminUser);
    save("adminPass", adminPass);
    save("contatosSetor", contatos);
    
    toast({ title: "Alterações salvas com sucesso", description: "Dados salvos no arquivo HTML" });
  };

  // Dados de teste
  const carregarDadosTeste = () => {
    if (!adminLogged) {
      toast({ title: "Acesso restrito", description: "Somente administrador pode carregar dados.", variant: "destructive" });
      return;
    }
    const demoServidores: Servidor[] = [
      { matricula: "001", nome: "João Silva", cpf: "11111111111", cargo: "Coordenador" },
      { matricula: "002", nome: "Maria Souza", cpf: "22222222222", cargo: "Supervisor" },
      { matricula: "003", nome: "Carlos Oliveira", cpf: "33333333333", cargo: "Agente de Trânsito" },
    ];
    setServidores(demoServidores);
    toast({ title: "Base carregada", description: "Servidores de teste adicionados." });
  };

  const limparTudo = () => {
    if (!adminLogged) {
      toast({ title: "Acesso restrito", description: "Somente administrador pode limpar a base.", variant: "destructive" });
      return;
    }
    setServidores([]);
    setLancamentos([]);
    localStorage.removeItem("servidores");
    localStorage.removeItem("lancamentos");
    toast({ title: "Dados limpos" });
  };

  // CRUD de dias
  const adicionarDia = () => {
    const funcoesDisponiveis = getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria");
    setDias((d) => [...d, { data: periodo.inicio, horas: 8, funcao: funcoesDisponiveis[0] }]);
  };
  
  const removerDia = (index: number) => setDias((d) => d.filter((_, i) => i !== index));
  
  const atualizarDia = (index: number, patch: Partial<DiaTrabalho>) =>
    setDias((d) => d.map((dia, i) => (i === index ? { ...dia, ...patch } : dia)));

  // Salvar lançamento com validações e geração de PDF
  const salvarLancamento = async () => {
    if (!matriculaSelecionada) {
      toast({ title: "Selecione um servidor válido", description: "Use a busca para escolher um servidor.", variant: "destructive" });
      return;
    }
    const srv = servidores.find((s) => s.matricula === matriculaSelecionada);
    if (!srv) {
      toast({ title: "Servidor não encontrado", description: "A matrícula selecionada não existe na base.", variant: "destructive" });
      return;
    }

    const diasValidos = dias.filter((d) => d.data && d.horas > 0);
    if (diasValidos.length === 0) {
      toast({ title: "Adicione pelo menos um dia", variant: "destructive" });
      return;
    }

    // Não permitir a mesma data repetida (dentro do formulário)
    const datas = diasValidos.map((d) => d.data);
    const hasDup = new Set(datas).size !== datas.length;
    if (hasDup) {
      toast({ title: "Datas repetidas", description: "Remova dias duplicados no lançamento.", variant: "destructive" });
      return;
    }

    // Bloquear conflito com lançamentos existentes para o mesmo servidor (qualquer operação)
    const datasSet = new Set(datas);
    const conflito = lancamentos.some((l) => {
      if (editingId && l.id === editingId) return false;
      if (l.servidor.matricula !== srv.matricula) return false;
      return l.dias.some((d) => datasSet.has(d.data));
    });
    if (conflito) {
      toast({ title: "Conflito de datas", description: "Já existe lançamento para este servidor em uma das datas informadas.", variant: "destructive" });
      return;
    }

    // Restringir datas ao período da operação selecionada
    const foraPeriodo = diasValidos.find((d) => d.data < periodo.inicio || d.data > periodo.fim);
    if (foraPeriodo) {
      toast({ title: "Data fora do período", description: "Há dias fora do período da operação selecionada.", variant: "destructive" });
      return;
    }

    const nomeOperacao = selectedOp?.label ?? "";
    const novo: Lancamento = {
      id: editingId ?? crypto.randomUUID(),
      operacao: selectedOp?.tipo ?? "ordinaria",
      nomeOperacao,
      periodo,
      servidor: srv,
      dias: diasValidos,
      createdAt: new Date().toISOString(),
    };

    setLancamentos((prev) => {
      if (editingId) return prev.map((l) => (l.id === editingId ? novo : l));
      return [novo, ...prev];
    });

    // Gerar PDF automaticamente
    await gerarFrequenciaPDF();

    // reset básico
    setEditingId(null);
    setDias([{ data: periodo.inicio, horas: 8, funcao: getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria")[0] }]);
    toast({ title: "Lançamento salvo e PDF gerado" });
  };

  // Consolidação
  const [filtroOperacao, setFiltroOperacao] = useState<string>("todos");
  const consolidado = useMemo(() => {
    const map = new Map<string, {
      matricula: string; nome: string; cpf: string; coordenador: number; supervisor: number; agente: number; apoio: number; totalHoras: number; valorHoras: number; alimentacao: number; transporte: number; totalGeral: number;
    }>();

    lancamentos.forEach((l) => {
      if (filtroOperacao !== "todos" && l.nomeOperacao !== filtroOperacao) return;

      const key = l.servidor.matricula;
      if (!map.has(key)) {
        map.set(key, { 
          matricula: key, 
          nome: l.servidor.nome, 
          cpf: l.servidor.cpf || "",
          coordenador: 0, 
          supervisor: 0, 
          agente: 0, 
          apoio: 0, 
          totalHoras: 0, 
          valorHoras: 0, 
          alimentacao: 0, 
          transporte: 0, 
          totalGeral: 0 
        });
      }
      const row = map.get(key)!;

      l.dias.forEach((d) => {
        const v = valorHora(d.funcao, l.operacao);
        // Mapear funções para colunas adequadas
        if (d.funcao.includes("coordenador")) row.coordenador += d.horas;
        else if (d.funcao.includes("supervisor")) row.supervisor += d.horas;
        else if (d.funcao.includes("agente") || d.funcao.includes("guarda") || d.funcao.includes("fiscaliza")) row.agente += d.horas;
        else row.apoio += d.horas;
        
        row.totalHoras += d.horas;
        row.valorHoras += d.horas * v;
        
        // Cálculo de alimentação por operação
        if (l.operacao === "ordinaria") {
          if (d.horas >= 8) row.alimentacao += 2 * d.horas;
        } else if (l.operacao === "reveillon") {
          if (d.horas >= 12) row.alimentacao += 13.68;
          else row.alimentacao += (13.68 * d.horas) / 12;
        } else if (l.operacao === "carnaval") {
          const valorAlimentacao = (alimentacao as any).carnaval?.[d.horas] || 0;
          row.alimentacao += valorAlimentacao;
        }
      });

      if (l.operacao === "carnaval") row.transporte += 20; // exemplo por lançamento
      row.totalGeral = row.valorHoras + row.alimentacao + row.transporte;
    });

    return Array.from(map.values());
  }, [lancamentos, filtroOperacao, valores, alimentacao]);

  const nomesOperacoesConsolidadas = useMemo(() => {
    return Array.from(new Set(lancamentos.map((l) => l.nomeOperacao)));
  }, [lancamentos]);

  // PDFs
  const gerarFrequenciaPDF = async () => {
    if (!matriculaSelecionada) {
      toast({ title: "Selecione um servidor válido", variant: "destructive" });
      return;
    }
    const srv = servidores.find((s) => s.matricula === matriculaSelecionada);
    if (!srv) {
      toast({ title: "Servidor não encontrado", variant: "destructive" });
      return;
    }

    const doc = new jsPDF();

    // Cabeçalho simples sem logo
    doc.setFontSize(16);
    doc.text("TRANSALVADOR", 105, 15, { align: "center" } as any);
    doc.setFontSize(12);
    doc.text("Superintendência de Trânsito de Salvador", 105, 22, { align: "center" } as any);
    doc.setFontSize(14);
    doc.text("Frequência de Operação Especial", 105, 32, { align: "center" } as any);

    const nomeOp = selectedOp?.label ?? "";
    doc.setFontSize(10);
    doc.text(`Operação: ${nomeOp}`, 20, 45);
    doc.text(`Período: ${toBR(periodo.inicio)} a ${toBR(periodo.fim)}`, 20, 52);
    doc.text(`Servidor: ${srv.nome}`, 20, 59);
    doc.text(`Matrícula: ${srv.matricula}`, 20, 66);
    doc.text(`CPF: ${formatarCPF(srv.cpf)}`, 20, 73);

    const ordered = [...dias].filter((d) => d.data && d.horas > 0).sort((a, b) => a.data.localeCompare(b.data));
    if (ordered.length === 0) {
      toast({ title: "Adicione pelo menos um dia", variant: "destructive" });
      return;
    }

    const rows = ordered.map((d) => {
      const vHora = valorHora(d.funcao, selectedOp?.tipo ?? "ordinaria");
      const valorTotal = d.horas * vHora;
      let valorAlimentacao = 0;
      
      // Cálculo alimentação por tipo de operação
      if (selectedOp?.tipo === "ordinaria") {
        valorAlimentacao = d.horas >= 8 ? d.horas * 2 : 0;
      } else if (selectedOp?.tipo === "reveillon") {
        valorAlimentacao = d.horas >= 12 ? 13.68 : (13.68 * d.horas) / 12;
      } else if (selectedOp?.tipo === "carnaval") {
        valorAlimentacao = (alimentacao as any).carnaval?.[d.horas] || 0;
      }
      
      return [
        toBR(d.data),
        d.horas,
        funcaoLabel(d.funcao),
        fmtBRL(vHora),
        fmtBRL(valorTotal),
        fmtBRL(valorAlimentacao),
        "",
      ];
    });

    autoTable(doc, {
      startY: 85,
      head: [["Data", "Horas", "Função", "Valor/Hora", "Valor Total", "Alimentação", "Assinatura"]],
      body: rows,
      styles: { fontSize: 10 },
    });

    const finalY = ((doc as any).lastAutoTable?.finalY ?? 85) + 15;
    doc.text("________________________________________", 40, finalY);
    doc.text("Assinatura do Servidor", 40, finalY + 7);
    doc.text("________________________________________", 120, finalY);
    doc.text("Carimbo e Assinatura do Gestor", 120, finalY + 7);

    doc.save(`Frequencia_${srv.matricula}.pdf`);
    toast({ title: "Frequência gerada" });
  };

  const gerarPlanilhaPDF = async () => {
    if (consolidado.length === 0) {
      toast({ title: "Nenhum dado para exportar", variant: "destructive" });
      return;
    }
    const doc = new jsPDF({ orientation: "landscape" });

    // Cabeçalho simples sem logo
    doc.setFontSize(16);
    doc.text("TRANSALVADOR", 148, 15, { align: "center" } as any);
    doc.setFontSize(12);
    doc.text("Superintendência de Trânsito de Salvador", 148, 22, { align: "center" } as any);
    doc.setFontSize(14);
    doc.text("Planilha Consolidada - Operações Especiais", 148, 32, { align: "center" } as any);

    if (filtroOperacao !== "todos") {
      doc.setFontSize(12);
      doc.text(`Operação: ${filtroOperacao}`, 20, 42);
    }

    const headers = [
      "Matrícula",
      "Nome",
      "CPF",
      "Coord. (h)",
      "Super. (h)",
      "Agente (h)",
      "Apoio (h)",
      "Total Horas",
      "Valor Horas",
      "Alimentação",
      "Transporte",
      "Total Geral",
    ];

    const rows = [...consolidado]
      .sort((a, b) => a.nome.localeCompare(b.nome))
      .map((r) => [
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
        r.transporte > 0 ? fmtBRL(r.transporte) : '-',
        fmtBRL(r.totalGeral),
      ]);

    autoTable(doc, {
      startY: 50,
      head: [headers],
      body: rows,
      styles: { fontSize: 7 },
      margin: { top: 50 },
    });

    const totals = consolidado.reduce(
      (acc, r) => {
        acc.h += r.totalHoras;
        acc.v += r.valorHoras;
        acc.a += r.alimentacao;
        acc.t += r.transporte;
        acc.g += r.totalGeral;
        return acc;
      },
      { h: 0, v: 0, a: 0, t: 0, g: 0 }
    );

    const finalY = ((doc as any).lastAutoTable?.finalY ?? 50) + 10;
    doc.setFontSize(10);
    doc.text(
      `TOTAL GERAL: Horas: ${totals.h.toFixed(2)} | Valor Horas: R$ ${totals.v.toFixed(2)} | Alimentação: R$ ${totals.a.toFixed(2)} | Transporte: R$ ${totals.t.toFixed(2)} | Total: R$ ${totals.g.toFixed(2)}`,
      20,
      finalY
    );
    
    // Campo para assinatura do chefe de setor
    const signatureY = finalY + 20;
    doc.text("________________________________________", 20, signatureY);
    doc.text("Assinatura do Chefe de Setor", 20, signatureY + 7);
    doc.text("________________________________________", 120, signatureY);
    doc.text("Data: ___/___/_____", 120, signatureY + 7);

    const nomeArquivo = filtroOperacao !== "todos" ? `Planilha_${filtroOperacao.replace(/\s+/g, "_")}.pdf` : "Planilha_Consolidada_Todas.pdf";
    doc.save(nomeArquivo);
    toast({ title: "Planilha gerada" });
  };

  const exportarPlanilhaExcel = () => {
    if (consolidado.length === 0) {
      toast({ title: "Nenhum dado para exportar", variant: "destructive" });
      return;
    }
    const headers = [
      "Matrícula",
      "Nome",
      "CPF", 
      "Coord. (h)",
      "Super. (h)",
      "Agente (h)",
      "Apoio (h)",
      "Total Horas",
      "Valor Horas",
      "Alimentação",
      "Transporte",
      "Total Geral",
    ];
    const rows = [...consolidado]
      .sort((a, b) => a.nome.localeCompare(b.nome))
      .map((r) => [
        r.matricula,
        r.nome,
        formatarCPF(r.cpf),
        Number(r.coordenador.toFixed(2)),
        Number(r.supervisor.toFixed(2)),
        Number(r.agente.toFixed(2)),
        Number(r.apoio.toFixed(2)),
        Number(r.totalHoras.toFixed(2)),
        Number(r.valorHoras.toFixed(2)),
        Number(r.alimentacao.toFixed(2)),
        Number(r.transporte.toFixed(2)),
        Number(r.totalGeral.toFixed(2)),
      ]);
    
    const totals = consolidado.reduce(
      (acc, r) => {
        acc.h += r.totalHoras;
        acc.v += r.valorHoras;
        acc.a += r.alimentacao;
        acc.t += r.transporte;
        acc.g += r.totalGeral;
        return acc;
      },
      { h: 0, v: 0, a: 0, t: 0, g: 0 }
    );

    const ws = XLSX.utils.aoa_to_sheet([
      headers,
      ...rows,
      [
        "TOTAL GERAL", "", "", "", "", "", "",
        Number(totals.h.toFixed(2)),
        Number(totals.v.toFixed(2)),
        Number(totals.a.toFixed(2)),
        Number(totals.t.toFixed(2)),
        Number(totals.g.toFixed(2)),
      ],
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

  // Se o sistema não estiver desbloqueado, mostrar apenas tela de login
  if (!sistemaDesbloqueado) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-background to-card/60">
        <Card className="w-full max-w-md">
          <CardHeader className="text-center">
            <CardTitle className="text-2xl mb-2">GEOPS - Sistema Bloqueado</CardTitle>
            <p className="text-muted-foreground">Sistema de Geração e Controle de Operações Especiais</p>
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              <div className="space-y-2">
                <Label>Usuário</Label>
                <Input
                  value={loginUsuario}
                  onChange={(e) => setLoginUsuario(e.target.value)}
                  placeholder="Digite o usuário"
                />
              </div>
              <div className="space-y-2">
                <Label>Senha</Label>
                <Input
                  type="password"
                  value={loginSenha}
                  onChange={(e) => setLoginSenha(e.target.value)}
                  placeholder="Digite a senha"
                  onKeyDown={(e) => e.key === "Enter" && realizarLogin()}
                />
              </div>
              <Button onClick={realizarLogin} className="w-full">
                Entrar no Sistema
              </Button>
              
              <hr className="my-4" />
              
              <div className="space-y-3">
                <h4 className="font-medium">Redefinir Senha de Administrador</h4>
                <div className="space-y-2">
                  <Label>Pergunta: {SECURITY_QUESTION}</Label>
                  <Input
                    value={respostaSeguranca}
                    onChange={(e) => setRespostaSeguranca(e.target.value)}
                    placeholder="Digite a resposta"
                  />
                </div>
                <div className="space-y-2">
                  <Label>Nova Senha</Label>
                  <Input
                    type="password"
                    value={novaSenhaSeguranca}
                    onChange={(e) => setNovaSenhaSeguranca(e.target.value)}
                    placeholder="Digite a nova senha"
                  />
                </div>
                <Button onClick={redefinirSenha} variant="secondary" className="w-full">
                  Redefinir Senha
                </Button>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="min-h-screen">
      <header className="border-b bg-gradient-to-br from-background to-card/60">
        <div className="container py-8">
          <div className="flex justify-between items-start">
            <div>
              <h1 className="text-3xl font-bold tracking-tight">GEOPS - Gerador de Operações Especiais Segep</h1>
              <p className="text-muted-foreground mt-1">Sistema de Geração e Controle de Operações Especiais.</p>
            </div>
            
            {/* Informações do usuário logado */}
            <div className="flex items-center gap-4">
              {usuarioTipo === "admin" ? (
                <div className="flex items-center gap-2">
                  <div className="text-right">
                    <div className="font-medium text-green-600">Administrador</div>
                    <div className="text-sm text-muted-foreground">{adminUser}</div>
                  </div>
                  <Button onClick={logout} variant="outline" size="sm">
                    Logout
                  </Button>
                </div>
              ) : usuarioLogado ? (
                <div className="flex items-center gap-2">
                  <div className="text-right">
                    <div className="font-medium">{usuarioLogado.nomeSetor}</div>
                    <div className="text-sm text-muted-foreground">{usuarioLogado.nomeGestor}</div>
                  </div>
                  <Button onClick={logout} variant="outline" size="sm">
                    Logout
                  </Button>
                </div>
               ) : null}
            </div>
          </div>
        </div>
      </header>

      <main className="container py-8">
        <Tabs value={tab} onValueChange={setTab} className="w-full">
          <TabsList className="grid w-full" style={{ gridTemplateColumns: usuarioTipo === "admin" ? "repeat(4, 1fr)" : "repeat(3, 1fr)" }}>
            {usuarioTipo === "admin" && <TabsTrigger value="rh">Banco de Dados</TabsTrigger>}
            <TabsTrigger value="lancamentos">Lançamentos</TabsTrigger>
            <TabsTrigger value="planilha">Planilha</TabsTrigger>
            <TabsTrigger value="logs">Histórico de Lançamentos</TabsTrigger>
          </TabsList>

          {/* Banco de Dados - Apenas para administradores */}
          {usuarioTipo === "admin" && (
          <TabsContent value="rh" className="mt-6">
            <div className="grid gap-6 md:grid-cols-2">
              
              {/* Painel de Acesso */}
              <Card>
                <CardHeader>
                  <CardTitle>Painel de Acesso</CardTitle>
                </CardHeader>
                <CardContent>
                  {!adminLogged ? (
                    <div className="space-y-4">
                      <div className="space-y-2">
                        <Label>Usuário</Label>
                        <Input
                          value={loginUsuario}
                          onChange={(e) => setLoginUsuario(e.target.value)}
                          placeholder="Digite o usuário"
                        />
                      </div>
                      <div className="space-y-2">
                        <Label>Senha</Label>
                        <Input
                          type="password"
                          value={loginSenha}
                          onChange={(e) => setLoginSenha(e.target.value)}
                          placeholder="Digite a senha"
                          onKeyDown={(e) => e.key === "Enter" && realizarLogin()}
                        />
                      </div>
                      <Button onClick={realizarLogin} className="w-full">
                        Entrar
                      </Button>
                      
                      <hr className="my-4" />
                      
                      <div className="space-y-3">
                        <h4 className="font-medium">Redefinir Senha</h4>
                        <div className="space-y-2">
                          <Label>Pergunta: {SECURITY_QUESTION}</Label>
                          <Input
                            value={respostaSeguranca}
                            onChange={(e) => setRespostaSeguranca(e.target.value)}
                            placeholder="Digite a resposta"
                          />
                        </div>
                        <div className="space-y-2">
                          <Label>Nova Senha</Label>
                          <Input
                            type="password"
                            value={novaSenhaSeguranca}
                            onChange={(e) => setNovaSenhaSeguranca(e.target.value)}
                            placeholder="Digite a nova senha"
                          />
                        </div>
                        <Button onClick={redefinirSenha} variant="secondary" className="w-full">
                          Redefinir Senha
                        </Button>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      <div className="text-green-600 font-medium">
                        ✓ Logado como: {adminUser}
                      </div>
                      
                       <div className="space-y-3">
                         <h4 className="font-medium">Alterar Nome de Usuário</h4>
                         <div className="space-y-2">
                           <Label>Novo Nome de Usuário</Label>
                           <Input
                             value={novoUsuario}
                             onChange={(e) => setNovoUsuario(e.target.value)}
                             placeholder="Digite o novo nome de usuário"
                           />
                         </div>
                         <Button onClick={alterarNomeUsuario} variant="secondary" className="w-full">
                           Alterar Nome de Usuário
                         </Button>
                       </div>

                       <div className="space-y-3">
                         <h4 className="font-medium">Alterar Senha</h4>
                         <div className="space-y-2">
                           <Label>Nova Senha</Label>
                           <Input
                             type="password"
                             value={novaSenha}
                             onChange={(e) => setNovaSenha(e.target.value)}
                             placeholder="Digite a nova senha"
                           />
                         </div>
                         <div className="space-y-2">
                           <Label>Confirmar Senha</Label>
                           <Input
                             type="password"
                             value={confirmarSenha}
                             onChange={(e) => setConfirmarSenha(e.target.value)}
                             placeholder="Confirme a nova senha"
                           />
                         </div>
                         <Button onClick={alterarSenha} variant="secondary" className="w-full">
                           Alterar Senha
                         </Button>
                       </div>
                      
                       <Button onClick={salvarAlteracoes} className="w-full">
                         Salvar Alterações
                       </Button>
                       
                       <Button onClick={logout} variant="destructive" className="w-full">
                         Logout
                       </Button>
                    </div>
                  )}
                </CardContent>
              </Card>

              {/* Configuração de Valores */}
              <Card>
                <CardHeader>
                  <CardTitle>Valores por Operação</CardTitle>
                </CardHeader>
                <CardContent>
                  {!adminLogged ? (
                    <div className="text-sm text-muted-foreground">
                      Faça login para configurar os valores por operação.
                    </div>
                  ) : (
                    <div className="space-y-6">
                      <div>
                        <h4 className="font-medium mb-3">Operação Ordinária</h4>
                        <div className="grid grid-cols-2 gap-2">
                          {Object.entries(valores.ordinaria).map(([funcao, valor]) => (
                            <div key={funcao} className="space-y-1">
                              <Label className="text-xs">{funcaoLabel(funcao as FuncaoID)}</Label>
                              <Input
                                type="number"
                                step="0.01"
                                value={valor}
                                onChange={(e) => setValores(prev => ({
                                  ...prev,
                                  ordinaria: { ...prev.ordinaria, [funcao]: Number(e.target.value) }
                                }))}
                                className="text-xs"
                              />
                            </div>
                          ))}
                        </div>
                      </div>
                      
                      <div>
                        <h4 className="font-medium mb-3">Operação Reveillon</h4>
                        <div className="grid grid-cols-2 gap-2">
                          {Object.entries(valores.reveillon).map(([funcao, valor]) => (
                            <div key={funcao} className="space-y-1">
                              <Label className="text-xs">{funcaoLabel(funcao as FuncaoID)}</Label>
                              <Input
                                type="number"
                                step="0.01"
                                value={valor}
                                onChange={(e) => setValores(prev => ({
                                  ...prev,
                                  reveillon: { ...prev.reveillon, [funcao]: Number(e.target.value) }
                                }))}
                                className="text-xs"
                              />
                            </div>
                          ))}
                        </div>
                      </div>
                      
                      <div>
                        <h4 className="font-medium mb-3">Operação Carnaval</h4>
                        <div className="grid grid-cols-2 gap-2">
                          {Object.entries(valores.carnaval).map(([funcao, valor]) => (
                            <div key={funcao} className="space-y-1">
                              <Label className="text-xs">{funcaoLabel(funcao as FuncaoID)}</Label>
                              <Input
                                type="number"
                                step="0.01"
                                value={valor}
                                onChange={(e) => setValores(prev => ({
                                  ...prev,
                                  carnaval: { ...prev.carnaval, [funcao]: Number(e.target.value) }
                                }))}
                                className="text-xs"
                              />
                            </div>
                          ))}
                        </div>
                       </div>
                       
                       <div>
                         <h4 className="font-medium mb-3">Valores de Alimentação</h4>
                         
                         <div className="space-y-4">
                           <div>
                             <h5 className="text-sm font-medium mb-2">Operação Ordinária</h5>
                             <div className="space-y-2">
                               <div>
                                 <Label className="text-xs">Valor por hora (mínimo 8h)</Label>
                                 <Input
                                   type="number"
                                   step="0.01"
                                   value={(alimentacao.ordinaria as any)?.valorHora || 2}
                                   onChange={(e) => setAlimentacao(prev => ({
                                     ...prev,
                                     ordinaria: { ...prev.ordinaria, valorHora: Number(e.target.value) }
                                   }))}
                                   className="text-xs"
                                 />
                               </div>
                             </div>
                           </div>

                           <div>
                             <h5 className="text-sm font-medium mb-2">Operação Reveillon</h5>
                             <div>
                               <Label className="text-xs">Valor para 12h</Label>
                               <Input
                                 type="number"
                                 step="0.01"
                                 value={(alimentacao.reveillon as any)?.[12] || 13.68}
                                 onChange={(e) => setAlimentacao(prev => ({
                                   ...prev,
                                   reveillon: { ...prev.reveillon, 12: Number(e.target.value) }
                                 }))}
                                 className="text-xs"
                               />
                             </div>
                           </div>

                            <div>
                              <h5 className="text-sm font-medium mb-2">Operação Carnaval</h5>
                              <div className="grid grid-cols-3 gap-2 max-h-60 overflow-y-auto">
                                {Object.entries((alimentacao.carnaval as any) || {})
                                  .sort(([a], [b]) => Number(a) - Number(b))
                                  .map(([horas, valor]) => (
                                  <div key={horas} className="space-y-1">
                                    <Label className="text-xs">{horas}h</Label>
                                    <Input
                                      type="number"
                                      step="0.01"
                                      value={valor as number}
                                      onChange={(e) => setAlimentacao(prev => ({
                                        ...prev,
                                        carnaval: { ...prev.carnaval, [horas]: Number(e.target.value) }
                                      }))}
                                      className="text-xs"
                                    />
                                  </div>
                                ))}
                              </div>
                            </div>
                         </div>
                       </div>
                     </div>
                   )}
                 </CardContent>
                </Card>


              {/* Servidores */}
              <Card>
                <CardHeader>
                  <CardTitle>Servidores</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="flex gap-2 mb-4 flex-wrap items-center">
                    <Input type="file" accept=".xlsx,.xls" onChange={(e) => setExcelFile(e.target.files?.[0] ?? null)} className="max-w-xs" disabled={!adminLogged} />
                    <Button onClick={importarExcel} disabled={!adminLogged}>Importar Excel</Button>
                    <Button onClick={carregarDadosTeste} disabled={!adminLogged}>Carregar dados de teste</Button>
                    <Button variant="secondary" onClick={limparTudo} disabled={!adminLogged}>Limpar tudo</Button>
                  </div>
                  {!adminLogged && (
                    <div className="text-xs text-muted-foreground mb-3">Apenas o administrador pode alterar os dados do banco.</div>
                  )}

                  <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>Matrícula</TableHead>
                          <TableHead>Nome</TableHead>
                          <TableHead>CPF</TableHead>
                          <TableHead>Cargo</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {servidores.length === 0 && (
                          <TableRow>
                            <TableCell colSpan={4} className="text-center text-muted-foreground">Nenhum servidor cadastrado.</TableCell>
                          </TableRow>
                        )}
                        {servidores.map((s) => (
                          <TableRow key={s.matricula}>
                            <TableCell>{s.matricula}</TableCell>
                            <TableCell>{s.nome}</TableCell>
                            <TableCell>{formatarCPF(s.cpf)}</TableCell>
                            <TableCell>{s.cargo}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                </CardContent>
               </Card>

              {/* Sistema de Perfis de Usuário */}
              <Card>
                <CardHeader>
                  <CardTitle>Sistema de Perfis de Usuário</CardTitle>
                </CardHeader>
                <CardContent>
                  {!adminLogged ? (
                    <div className="text-sm text-muted-foreground">
                      Faça login como administrador para gerenciar perfis de usuário.
                    </div>
                  ) : (
                    <div className="space-y-6">
                      <div>
                        <h4 className="font-medium mb-3">Criar Novo Perfil</h4>
                        <div className="grid grid-cols-2 gap-4">
                          <div className="space-y-2">
                            <Label>Nome do Setor</Label>
                            <Input
                              value={novoPerfil.nomeSetor}
                              onChange={(e) => setNovoPerfil(prev => ({ ...prev, nomeSetor: e.target.value }))}
                              placeholder="Ex: SEGEP, GEOP, etc."
                            />
                          </div>
                          <div className="space-y-2">
                            <Label>Nome do Gestor</Label>
                            <Input
                              value={novoPerfil.nomeGestor}
                              onChange={(e) => setNovoPerfil(prev => ({ ...prev, nomeGestor: e.target.value }))}
                              placeholder="Nome completo do gestor"
                            />
                          </div>
                          <div className="space-y-2">
                            <Label>Matrícula</Label>
                            <Input
                              value={novoPerfil.matricula}
                              onChange={(e) => setNovoPerfil(prev => ({ ...prev, matricula: e.target.value }))}
                              placeholder="Matrícula do gestor"
                            />
                          </div>
                          <div className="space-y-2">
                            <Label>Login</Label>
                            <Input
                              value={novoPerfil.login}
                              onChange={(e) => setNovoPerfil(prev => ({ ...prev, login: e.target.value }))}
                              placeholder="Nome de usuário para login"
                            />
                          </div>
                          <div className="space-y-2">
                            <Label>Senha</Label>
                            <Input
                              type="password"
                              value={novoPerfil.senha}
                              onChange={(e) => setNovoPerfil(prev => ({ ...prev, senha: e.target.value }))}
                              placeholder="Senha de acesso"
                            />
                          </div>
                        </div>
                        <Button onClick={criarPerfil} className="mt-4">Criar Perfil</Button>
                      </div>

                      <div>
                        <h4 className="font-medium mb-3">Perfis Existentes</h4>
                        <div className="space-y-2">
                          {perfisUsuario.length === 0 ? (
                            <div className="text-sm text-muted-foreground">Nenhum perfil criado ainda.</div>
                          ) : (
                            perfisUsuario.map((perfil) => (
                              <div key={perfil.id} className="flex items-center justify-between p-3 border rounded">
                                <div className="flex-1">
                                  <div className="font-medium">{perfil.nomeSetor} - {perfil.nomeGestor}</div>
                                  <div className="text-sm text-muted-foreground">
                                    Login: {perfil.login} | Matrícula: {perfil.matricula}
                                  </div>
                                </div>
                                <div className="flex gap-2">
                                  <Button
                                    size="sm"
                                    variant={perfil.ativo ? "default" : "secondary"}
                                    onClick={() => togglePerfilAtivo(perfil.id)}
                                  >
                                    {perfil.ativo ? "Ativo" : "Inativo"}
                                  </Button>
                                  <Button
                                    size="sm"
                                    variant="destructive"
                                    onClick={() => excluirPerfil(perfil.id)}
                                  >
                                    Excluir
                                  </Button>
                                </div>
                              </div>
                            ))
                          )}
                        </div>
                      </div>
                    </div>
                  )}
                </CardContent>
              </Card>

              {/* Ajuda */}
              <Card>
                <CardHeader>
                  <CardTitle>Ajuda</CardTitle>
                </CardHeader>
                <CardContent>
                  <ol className="list-decimal list-inside space-y-1 text-sm text-muted-foreground">
                    <li>Escolha Ano e Operação na aba Lançamentos (período é definido automaticamente).</li>
                    <li>Selecione o servidor e adicione os dias, horas e função de cada dia.</li>
                    <li>Salve o lançamento para gerar a Frequência em PDF (para assinatura).</li>
                    <li>Na aba Planilha, exporte o consolidado em PDF ou Excel.</li>
                    <li>Use a aba Logs para editar ou excluir lançamentos; a planilha é atualizada automaticamente.</li>
                  </ol>
                  <div className="text-xs text-muted-foreground mt-3">
                    Observações: não é permitido dois lançamentos na mesma data para o mesmo servidor; o servidor pode ter funções diferentes em dias distintos da mesma operação.
                  </div>
                </CardContent>
              </Card>

              {/* Sobre */}
              <Card>
                <CardHeader>
                  <CardTitle>Sobre</CardTitle>
                </CardHeader>
                <CardContent className="space-y-3">
                  <p className="text-sm">GEOPS - Gerador de Operações Especiais Segep.</p>
                  <p className="text-sm">Desenvolvido por Rui Cezar Pereira da Paixão Junior</p>
                  <p className="text-sm">Salvador, 2025.</p>

                  <div className="grid md:grid-cols-2 gap-2">
                    <div className="space-y-1">
                      <Label>Telefone principal</Label>
                      <Input value={contatos.telefone1} onChange={(e)=>setContatos((c)=>({...c, telefone1: e.target.value}))} placeholder="(71) 9 0000-0000" disabled={!adminLogged} />
                    </div>
                    <div className="space-y-1">
                      <Label>Telefone secundário</Label>
                      <Input value={contatos.telefone2} onChange={(e)=>setContatos((c)=>({...c, telefone2: e.target.value}))} placeholder="(71) 9 0000-0000" disabled={!adminLogged} />
                    </div>
                  </div>
                  <Button onClick={()=>toast({ title: "Contatos atualizados" })} disabled={!adminLogged}>Salvar Contatos</Button>
                  {!adminLogged && (
                    <div className="text-xs text-muted-foreground">Faça login no painel de acesso para editar os telefones.</div>
                  )}
                </CardContent>
                </Card>
            </div>
          </TabsContent>
          )}

          {/* Lançamentos */}
          <TabsContent value="lancamentos" className="mt-6">
            <div className="grid gap-6 md:grid-cols-3">
              <Card className="md:col-span-1">
                <CardHeader>
                  <CardTitle>Configurações</CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="space-y-2">
                    <Label>Ano</Label>
                    <Select value={String(ano)} onValueChange={(v) => setAno(Number(v))}>
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent className="z-50">
                        {anosDisponiveis.map((y) => (
                          <SelectItem key={y} value={String(y)}>{y}</SelectItem>
                        ))}
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
                          setDias((d) => (d.length === 0 ? [{ data: op.inicio, horas: 8, funcao: getFuncoesDisponiveisParaOperacao(op.tipo)[0] }] : d));
                        }
                      }}
                    >
                      <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
                      <SelectContent className="z-50">
                        {opcoesOperacao.map((op) => (
                          <SelectItem key={op.id} value={op.id}>{op.label}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>

                  {selectedOp?.tipo === "carnaval" && (
                    <div className="space-y-2">
                      <Label>Período Customizado (apenas para Carnaval)</Label>
                      <div className="grid grid-cols-2 gap-2">
                        <Input
                          type="date"
                          value={periodoCustomizado.inicio}
                          onChange={(e) => {
                            setPeriodoCustomizado(prev => ({ ...prev, inicio: e.target.value }));
                            if (e.target.value && periodoCustomizado.fim) {
                              setPeriodo({ inicio: e.target.value, fim: periodoCustomizado.fim });
                            }
                          }}
                          placeholder="Início"
                        />
                        <Input
                          type="date"
                          value={periodoCustomizado.fim}
                          onChange={(e) => {
                            setPeriodoCustomizado(prev => ({ ...prev, fim: e.target.value }));
                            if (periodoCustomizado.inicio && e.target.value) {
                              setPeriodo({ inicio: periodoCustomizado.inicio, fim: e.target.value });
                            }
                          }}
                          placeholder="Fim"
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
                <CardHeader>
                  <CardTitle>Lançamento de Horas</CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  {dias.map((d, i) => (
                    <div key={i} className="flex items-center gap-2 border rounded-md p-3">
                      <Input type="date" value={d.data} onChange={(e) => {
                        const v = e.target.value;
                        if (v < periodo.inicio || v > periodo.fim) {
                          toast({ title: "Data fora do período", variant: "destructive" });
                          atualizarDia(i, { data: clampToPeriodo(v, periodo.inicio, periodo.fim) });
                        } else {
                          atualizarDia(i, { data: v });
                        }
                      }} className="max-w-[220px]" />
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
                        className="w-24"
                      />
                      <Select value={d.funcao} onValueChange={(v: FuncaoID) => atualizarDia(i, { funcao: v })}>
                        <SelectTrigger className="w-56"><SelectValue placeholder="Função" /></SelectTrigger>
                        <SelectContent className="z-50">
                          {getFuncoesDisponiveisParaOperacao(selectedOp?.tipo || "ordinaria").map((funcao) => (
                            <SelectItem key={funcao} value={funcao}>{funcaoLabel(funcao)}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                      <div className="text-sm whitespace-nowrap">
                        Valor/h: R$ {valorHora(d.funcao, selectedOp?.tipo ?? "ordinaria").toFixed(2)}
                      </div>
                      <div className="ms-auto flex gap-2">
                        <Button variant="secondary" onClick={() => atualizarDia(i, { horas: d.horas + 1 })}>+1h</Button>
                        <Button variant="destructive" onClick={() => removerDia(i)}>Remover</Button>
                      </div>
                    </div>
                  ))}

                  <div className="flex gap-2">
                    <Button onClick={adicionarDia}>+ Adicionar dia</Button>
                    <Button onClick={salvarLancamento}>Salvar Lançamento</Button>
                  </div>
                </CardContent>
              </Card>
            </div>
          </TabsContent>

          {/* Planilha */}
          <TabsContent value="planilha" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Resumo de Horas Extras</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="grid md:grid-cols-3 gap-3 mb-4">
                  <div className="space-y-2 md:col-span-1">
                    <Label>Operação</Label>
                    <Select value={filtroOperacao} onValueChange={setFiltroOperacao}>
                      <SelectTrigger><SelectValue placeholder="Todas" /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="todos">Todas as Operações</SelectItem>
                        {nomesOperacoesConsolidadas.map((n) => (
                          <SelectItem key={n} value={n}>{n}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                </div>

                <div className="flex gap-2 mb-4">
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
                        <TableRow>
                          <TableCell colSpan={12} className="text-center text-muted-foreground">Nenhum lançamento.</TableCell>
                        </TableRow>
                      )}
                      {consolidado.map((r) => (
                        <TableRow key={r.matricula}>
                          <TableCell>{r.matricula}</TableCell>
                          <TableCell>{r.nome}</TableCell>
                          <TableCell>{formatarCPF(r.cpf)}</TableCell>
                          <TableCell>{r.coordenador.toFixed(2)}</TableCell>
                          <TableCell>{r.supervisor.toFixed(2)}</TableCell>
                          <TableCell>{r.agente.toFixed(2)}</TableCell>
                          <TableCell>{r.apoio.toFixed(2)}</TableCell>
                          <TableCell>{r.totalHoras.toFixed(2)}</TableCell>
                          <TableCell>R$ {r.valorHoras.toFixed(2)}</TableCell>
                          <TableCell>R$ {r.alimentacao.toFixed(2)}</TableCell>
                          <TableCell>{r.transporte > 0 ? `R$ ${r.transporte.toFixed(2)}` : '-'}</TableCell>
                          <TableCell>R$ {r.totalGeral.toFixed(2)}</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

               {/* Histórico de Lançamentos */}
          <TabsContent value="logs" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Histórico de Lançamentos</CardTitle>
              </CardHeader>
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
                        <TableRow>
                          <TableCell colSpan={6} className="text-center text-muted-foreground">Sem lançamentos.</TableCell>
                        </TableRow>
                      )}
                      {[...lancamentos].sort((a,b)=>a.servidor.nome.localeCompare(b.servidor.nome)).map((l) => (
                        <TableRow key={l.id}>
                          <TableCell>{new Date(l.createdAt).toLocaleDateString('pt-BR')}</TableCell>
                          <TableCell>{l.servidor.nome} ({l.servidor.matricula})</TableCell>
                          <TableCell>{l.nomeOperacao}</TableCell>
                          <TableCell>{toBR(l.periodo.inicio)} a {toBR(l.periodo.fim)}</TableCell>
                          <TableCell>{l.dias.length}</TableCell>
                          <TableCell className="flex gap-2">
                            <Button size="sm" onClick={() => editarLancamento(l)}>Editar</Button>
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
        </Tabs>
      </main>
    </div>
  );
};

export default Index;
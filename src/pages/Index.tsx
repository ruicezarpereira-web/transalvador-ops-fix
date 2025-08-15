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
    4: 12.18, 6: 18.27, 7: 21.32, 8: 35.92, 
    11: 43.08, 12: 47.00, 19: 59.09, 24: 74.64
  },
} as any;

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

// Storage functions
const save = (key: string, data: any) => {
  try {
    localStorage.setItem(`geops_${key}`, JSON.stringify(data));
  } catch (e) {
    console.error("Erro ao salvar dados:", e);
  }
};

const load = (key: string, defaultValue: any): any => {
  try {
    const item = localStorage.getItem(`geops_${key}`);
    return item ? JSON.parse(item) : defaultValue;
  } catch (e) {
    console.error("Erro ao carregar dados:", e);
    return defaultValue;
  }
};

const Index = () => {
  // Estados de autenticação global
  const [sistemaLogado, setSistemaLogado] = useState<boolean>(() => load("sistemaLogado", false));
  const [tipoUsuarioLogado, setTipoUsuarioLogado] = useState<'admin' | 'usuario' | null>(() => load("tipoUsuarioLogado", null) as 'admin' | 'usuario' | null);
  const [loginForm, setLoginForm] = useState({ usuario: "", senha: "" });

  // Estados de administração
  const [adminUser, setAdminUser] = useState<string>(() => load("adminUser", "RCPPJ"));
  const [adminPass, setAdminPass] = useState<string>(() => load("adminPass", "ruicpj@123"));
  const [adminLogged, setAdminLogged] = useState<boolean>(() => load("adminLogged", false));
  
  // Estados de perfis de usuário
  const [perfisUsuario, setPerfisUsuario] = useState<PerfilUsuario[]>(() => load("perfisUsuario", []));
  const [usuarioLogado, setUsuarioLogado] = useState<PerfilUsuario | null>(() => load("usuarioLogado", null));

  // Estados básicos do sistema
  const [servidores, setServidores] = useState<Servidor[]>(() => load("servidores", []));
  const [lancamentos, setLancamentos] = useState<Lancamento[]>(() => load("lancamentos", []));
  const [filtroOperacao, setFiltroOperacao] = useState<string>("todos");

  // Função de login unificado
  const realizarLoginUnificado = () => {
    // Verificar se é admin
    if (loginForm.usuario === adminUser && loginForm.senha === adminPass) {
      setSistemaLogado(true);
      setTipoUsuarioLogado('admin');
      setAdminLogged(true);
      setLoginForm({ usuario: "", senha: "" });
      toast({ title: "Login de administrador realizado com sucesso" });
      return;
    }
    
    // Verificar se é usuário cadastrado
    const usuarioEncontrado = perfisUsuario.find(p => 
      p.login === loginForm.usuario && p.senha === loginForm.senha && p.ativo
    );
    
    if (usuarioEncontrado) {
      setSistemaLogado(true);
      setTipoUsuarioLogado('usuario');
      setUsuarioLogado(usuarioEncontrado);
      setLoginForm({ usuario: "", senha: "" });
      toast({ title: `Bem-vindo, ${usuarioEncontrado.nomeGestor}!` });
      return;
    }
    
    toast({ title: "Credenciais inválidas", variant: "destructive" });
  };
  
  const logout = () => {
    setSistemaLogado(false);
    setTipoUsuarioLogado(null);
    setAdminLogged(false);
    setUsuarioLogado(null);
    setLoginForm({ usuario: "", senha: "" });
    toast({ title: "Logout realizado com sucesso" });
  };

  // Persistir estados de login
  useEffect(() => save("sistemaLogado", sistemaLogado), [sistemaLogado]);
  useEffect(() => save("tipoUsuarioLogado", tipoUsuarioLogado), [tipoUsuarioLogado]);
  useEffect(() => save("adminUser", adminUser), [adminUser]);
  useEffect(() => save("adminPass", adminPass), [adminPass]);
  useEffect(() => save("adminLogged", adminLogged), [adminLogged]);
  useEffect(() => save("perfisUsuario", perfisUsuario), [perfisUsuario]);
  useEffect(() => save("usuarioLogado", usuarioLogado), [usuarioLogado]);

  // Se não estiver logado, mostrar apenas tela de login
  if (!sistemaLogado) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-background to-muted flex items-center justify-center p-4">
        <Card className="w-full max-w-md">
          <CardHeader className="text-center">
            <CardTitle className="text-2xl font-bold text-primary">
              GEOPS - Gerador de Operações Especiais Segep
            </CardTitle>
            <p className="text-muted-foreground">
              Sistema de Geração e Controle de Operações Especiais.
            </p>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="space-y-2">
              <Label htmlFor="login-usuario">Usuário</Label>
              <Input
                id="login-usuario"
                type="text"
                value={loginForm.usuario}
                onChange={(e) => setLoginForm(prev => ({ ...prev, usuario: e.target.value }))}
                placeholder="Digite seu usuário"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="login-senha">Senha</Label>
              <Input
                id="login-senha"
                type="password"
                value={loginForm.senha}
                onChange={(e) => setLoginForm(prev => ({ ...prev, senha: e.target.value }))}
                placeholder="Digite sua senha"
                onKeyPress={(e) => e.key === 'Enter' && realizarLoginUnificado()}
              />
            </div>
            <Button 
              onClick={realizarLoginUnificado} 
              className="w-full"
              disabled={!loginForm.usuario || !loginForm.senha}
            >
              Entrar no Sistema
            </Button>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-background to-muted">
      <header className="border-b bg-card/50 backdrop-blur supports-[backdrop-filter]:bg-card/50">
        <div className="container mx-auto px-4 py-4 flex justify-between items-center">
          <div>
            <h1 className="text-2xl font-bold text-primary">
              GEOPS - Gerador de Operações Especiais Segep
            </h1>
            <p className="text-sm text-muted-foreground">
              Sistema de Geração e Controle de Operações Especiais.
            </p>
          </div>
          <div className="flex items-center gap-4">
            {usuarioLogado && (
              <div className="text-sm text-muted-foreground">
                <div>{usuarioLogado.nomeSetor}</div>
                <div>Gestor: {usuarioLogado.nomeGestor}</div>
              </div>
            )}
            {tipoUsuarioLogado === 'admin' && (
              <div className="text-sm text-muted-foreground">
                Administrador do Sistema
              </div>
            )}
            <Button variant="outline" size="sm" onClick={logout}>
              Sair
            </Button>
          </div>
        </div>
      </header>

      <main className="container mx-auto px-4 py-6">
        <Tabs defaultValue="configuracao" className="space-y-4">
          <TabsList className="grid w-full grid-cols-5">
            <TabsTrigger value="configuracao">Configuração</TabsTrigger>
            <TabsTrigger value="lancamento">Lançamento</TabsTrigger>
            <TabsTrigger value="planilha">Planilha</TabsTrigger>
            <TabsTrigger value="logs">Logs</TabsTrigger>
            {tipoUsuarioLogado === 'admin' && <TabsTrigger value="banco-dados">Banco de Dados</TabsTrigger>}
          </TabsList>

          {/* Configuração */}
          <TabsContent value="configuracao" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Configuração de Operação</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <p>Sistema configurado e funcionando!</p>
                  <p className="text-sm text-muted-foreground">
                    {tipoUsuarioLogado === 'admin' ? 'Modo Administrador' : 'Modo Usuário'}
                  </p>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Lançamento */}
          <TabsContent value="lancamento" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Lançamentos</CardTitle>
              </CardHeader>
              <CardContent>
                <p>Módulo de lançamentos em desenvolvimento...</p>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Planilha */}
          <TabsContent value="planilha" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Planilhas</CardTitle>
              </CardHeader>
              <CardContent>
                <p>Módulo de planilhas em desenvolvimento...</p>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Logs */}
          <TabsContent value="logs" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Logs</CardTitle>
              </CardHeader>
              <CardContent>
                <p>Módulo de logs em desenvolvimento...</p>
              </CardContent>
            </Card>
          </TabsContent>

          {/* Banco de Dados - apenas para admin */}
          {tipoUsuarioLogado === 'admin' && (
            <TabsContent value="banco-dados" className="mt-6">
              <Card>
                <CardHeader>
                  <CardTitle>Banco de Dados (Administrador)</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="space-y-4">
                    <p>Área restrita do administrador</p>
                    <div className="grid gap-4">
                      <div>
                        <Label>Usuário Admin</Label>
                        <Input
                          value={adminUser}
                          onChange={(e) => setAdminUser(e.target.value)}
                          placeholder="Nome de usuário admin"
                        />
                      </div>
                      <div>
                        <Label>Senha Admin</Label>
                        <Input
                          type="password"
                          value={adminPass}
                          onChange={(e) => setAdminPass(e.target.value)}
                          placeholder="Nova senha admin"
                        />
                      </div>
                    </div>
                  </div>
                </CardContent>
              </Card>
            </TabsContent>
          )}
        </Tabs>
      </main>
    </div>
  );
};

export default Index;
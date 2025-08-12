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

// Tipos
type DiaTrabalho = { data: string; horas: number };

type Lancamento = {
  id: string;
  operacao: OperacaoID;
  nomeOperacao: string;
  periodo: { inicio: string; fim: string };
  servidor: Servidor;
  funcao: FuncaoID;
  dias: DiaTrabalho[];
  createdAt: string;
};

type OperacaoID = "padrao" | "dezembro" | "janeiro" | "reveillon" | "carnaval";

type FuncaoID = "coordenador" | "supervisor" | "agente" | "apoio";

const valoresFuncao: Record<FuncaoID, number> = {
  coordenador: 20.5,
  supervisor: 15.5,
  agente: 12,
  apoio: 10,
};

const operacoes: { id: OperacaoID; label: string }[] = [
  { id: "padrao", label: "Padrão (21/mês - 20/mês+1)" },
  { id: "dezembro", label: "Dezembro (21-31/12)" },
  { id: "janeiro", label: "Janeiro (01-20/01)" },
  { id: "reveillon", label: "Reveillon" },
  { id: "carnaval", label: "Carnaval" },
];

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
    document.title = "Sistema de Operações Especiais - TRANSALVADOR";
  }, []);

  const [servidores, setServidores] = useState<Servidor[]>(() => load<Servidor[]>("servidores", []));
  const [lancamentos, setLancamentos] = useState<Lancamento[]>(() => load<Lancamento[]>("lancamentos", []));

  // Estado de Lançamento
  const [operacao, setOperacao] = useState<OperacaoID>("reveillon");
  const [periodo, setPeriodo] = useState({ inicio: "2024-12-24", fim: "2025-01-01" });
  const [funcao, setFuncao] = useState<FuncaoID>("coordenador");
  const [matriculaSelecionada, setMatriculaSelecionada] = useState<string | undefined>(undefined);
  const [dias, setDias] = useState<DiaTrabalho[]>([{ data: "2024-12-24", horas: 8 }] );

  // Valor por hora considerando a operação
  const valorHoraAtual = useMemo(() => {
    let v = valoresFuncao[funcao];
    if (operacao === "reveillon") v *= 2;
    return v;
  }, [funcao, operacao]);

  // Persistência
  useEffect(() => save("servidores", servidores), [servidores]);
  useEffect(() => save("lancamentos", lancamentos), [lancamentos]);

  // Dados de teste
  const carregarDadosTeste = () => {
    const demoServidores: Servidor[] = [
      { matricula: "001", nome: "João Silva", cpf: "11111111111", cargo: "Coordenador" },
      { matricula: "002", nome: "Maria Souza", cpf: "22222222222", cargo: "Supervisor" },
      { matricula: "003", nome: "Carlos Oliveira", cpf: "33333333333", cargo: "Agente de Trânsito" },
    ];
    setServidores(demoServidores);
    toast({ title: "Base carregada", description: "Servidores de teste adicionados." });
  };

  const limparTudo = () => {
    setServidores([]);
    setLancamentos([]);
    localStorage.removeItem("servidores");
    localStorage.removeItem("lancamentos");
    toast({ title: "Dados limpos" });
  };

  // CRUD de dias
  const adicionarDia = () => setDias((d) => [...d, { data: periodo.inicio, horas: 8 }]);
  const removerDia = (index: number) => setDias((d) => d.filter((_, i) => i !== index));
  const atualizarDia = (index: number, patch: Partial<DiaTrabalho>) =>
    setDias((d) => d.map((dia, i) => (i === index ? { ...dia, ...patch } : dia)));

  // Salvar lançamento (com correção da validação do servidor)
  const salvarLancamento = () => {
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

    const nomeOperacao = operacoes.find((o) => o.id === operacao)?.label ?? operacao;

    const novo: Lancamento = {
      id: crypto.randomUUID(),
      operacao,
      nomeOperacao,
      periodo,
      servidor: srv,
      funcao,
      dias: diasValidos,
      createdAt: new Date().toISOString(),
    };

    setLancamentos((prev) => [novo, ...prev]);
    // reset básico
    setDias([{ data: periodo.inicio, horas: 8 }]);
    toast({ title: "Lançamento salvo" });
  };

  // Consolidação
  const [filtroOperacao, setFiltroOperacao] = useState<string>("todos");
  const consolidado = useMemo(() => {
    const map = new Map<string, {
      matricula: string; nome: string; coordenador: number; supervisor: number; agente: number; apoio: number; totalHoras: number; valorHoras: number; alimentacao: number; transporte: number; totalGeral: number;
    }>();

    lancamentos.forEach((l) => {
      if (filtroOperacao !== "todos" && l.nomeOperacao !== filtroOperacao) return;

      const horas = l.dias.reduce((acc, d) => acc + d.horas, 0);
      let valor = valoresFuncao[l.funcao];
      if (l.operacao === "reveillon") valor *= 2;

      const key = l.servidor.matricula;
      if (!map.has(key)) {
        map.set(key, { matricula: key, nome: l.servidor.nome, coordenador: 0, supervisor: 0, agente: 0, apoio: 0, totalHoras: 0, valorHoras: 0, alimentacao: 0, transporte: 0, totalGeral: 0 });
      }
      const row = map.get(key)!;
      row[l.funcao] += horas as number;
      row.totalHoras += horas;
      row.valorHoras += horas * valor;
      l.dias.forEach((d) => { if (d.horas >= 8) row.alimentacao += 2 * d.horas; });
      if (l.operacao === "carnaval") row.transporte += 20; // exemplo
      row.totalGeral = row.valorHoras + row.alimentacao + row.transporte;
    });

    return Array.from(map.values());
  }, [lancamentos, filtroOperacao]);

  const nomesOperacoesConsolidadas = useMemo(() => {
    return Array.from(new Set(lancamentos.map((l) => l.nomeOperacao)));
  }, [lancamentos]);

  return (
    <div className="min-h-screen">
      <header className="border-b bg-gradient-to-br from-background to-card/60">
        <div className="container py-8">
          <h1 className="text-3xl font-bold tracking-tight">Sistema de Operações Especiais – TRANSALVADOR</h1>
          <p className="text-muted-foreground mt-1">Lançamentos rápidos, seleção inteligente de servidores e consolidação clara.</p>
        </div>
      </header>

      <main className="container py-8">
        <Tabs defaultValue="lancamentos" className="w-full">
          <TabsList className="grid grid-cols-3 w-full">
            <TabsTrigger value="rh">Banco de Dados (RH)</TabsTrigger>
            <TabsTrigger value="lancamentos">Lançamentos</TabsTrigger>
            <TabsTrigger value="planilha">Planilha</TabsTrigger>
          </TabsList>

          {/* RH */}
          <TabsContent value="rh" className="mt-6">
            <div className="grid gap-6 md:grid-cols-2">
              <Card>
                <CardHeader>
                  <CardTitle>Servidores</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="flex gap-2 mb-4 flex-wrap">
                    <Button onClick={carregarDadosTeste}>Carregar dados de teste</Button>
                    <Button variant="secondary" onClick={limparTudo}>Limpar tudo</Button>
                  </div>

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
                            <TableCell>{s.cpf}</TableCell>
                            <TableCell>{s.cargo}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                </CardContent>
              </Card>

              <Card>
                <CardHeader>
                  <CardTitle>Ajuda</CardTitle>
                </CardHeader>
                <CardContent>
                  <p className="text-sm text-muted-foreground">Use o botão de teste para popular a base. A importação por Excel poderá ser adicionada depois.</p>
                </CardContent>
              </Card>
            </div>
          </TabsContent>

          {/* Lançamentos */}
          <TabsContent value="lancamentos" className="mt-6">
            <div className="grid gap-6 md:grid-cols-3">
              <Card className="md:col-span-1">
                <CardHeader>
                  <CardTitle>Configurações</CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="space-y-2">
                    <Label>Operação</Label>
                    <Select value={operacao} onValueChange={(v: OperacaoID) => setOperacao(v)}>
                      <SelectTrigger><SelectValue placeholder="Selecione" /></SelectTrigger>
                      <SelectContent className="z-50">
                        {operacoes.map((op) => (
                          <SelectItem key={op.id} value={op.id}>{op.label}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label>Período</Label>
                    <div className="flex gap-2">
                      <Input type="date" value={periodo.inicio} onChange={(e) => setPeriodo((p) => ({ ...p, inicio: e.target.value }))} />
                      <Input type="date" value={periodo.fim} onChange={(e) => setPeriodo((p) => ({ ...p, fim: e.target.value }))} />
                    </div>
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

                  <div className="space-y-2">
                    <Label>Função na Operação</Label>
                    <Select value={funcao} onValueChange={(v: FuncaoID) => setFuncao(v)}>
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent className="z-50">
                        <SelectItem value="coordenador">Coordenador</SelectItem>
                        <SelectItem value="supervisor">Supervisor</SelectItem>
                        <SelectItem value="agente">Agente de Trânsito</SelectItem>
                        <SelectItem value="apoio">Apoio Administrativo</SelectItem>
                      </SelectContent>
                    </Select>
                    <div className="text-sm font-medium">Valor por hora: R$ {valorHoraAtual.toFixed(2)}</div>
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
                      <Input type="date" value={d.data} onChange={(e) => atualizarDia(i, { data: e.target.value })} className="max-w-[220px]" />
                      <Input
                        type="number"
                        min={1}
                        max={operacao === "carnaval" ? 24 : 12}
                        value={d.horas}
                        onChange={(e) => {
                          const n = Math.max(1, Math.min(Number(e.target.value || 0), operacao === "carnaval" ? 24 : 12));
                          atualizarDia(i, { horas: n });
                        }}
                        className="w-24"
                      />
                      <span className="text-sm">horas</span>
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

                <div className="rounded-md border overflow-auto" style={{ boxShadow: "var(--shadow-elevated)" }}>
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Matrícula</TableHead>
                        <TableHead>Nome</TableHead>
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
                          <TableCell colSpan={11} className="text-center text-muted-foreground">Nenhum lançamento.</TableCell>
                        </TableRow>
                      )}
                      {consolidado.map((r) => (
                        <TableRow key={r.matricula}>
                          <TableCell>{r.matricula}</TableCell>
                          <TableCell>{r.nome}</TableCell>
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
        </Tabs>
      </main>
    </div>
  );
};

export default Index;

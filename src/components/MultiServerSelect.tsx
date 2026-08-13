import * as React from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Badge } from "@/components/ui/badge";
import type { Servidor } from "@/components/ServerCombobox";

interface Props {
  servidores: Servidor[];
  selecionados: string[];
  onChange: (matriculas: string[]) => void;
}

const MultiServerSelect: React.FC<Props> = ({ servidores, selecionados, onChange }) => {
  const [busca, setBusca] = React.useState("");

  const filtrados = React.useMemo(() => {
    const q = busca.trim().toLowerCase();
    const base = [...servidores].sort((a, b) => a.nome.localeCompare(b.nome));
    if (!q) return base;
    return base.filter((s) =>
      `${s.nome} ${s.matricula} ${s.cpf ?? ""}`.toLowerCase().includes(q)
    );
  }, [servidores, busca]);

  const toggle = (matricula: string) => {
    onChange(
      selecionados.includes(matricula)
        ? selecionados.filter((m) => m !== matricula)
        : [...selecionados, matricula]
    );
  };

  return (
    <div className="space-y-2">
      <div className="flex gap-2 items-center">
        <Input
          value={busca}
          onChange={(e) => setBusca(e.target.value)}
          placeholder="Buscar por nome, matrícula ou CPF..."
        />
        <Button
          type="button"
          variant="secondary"
          onClick={() => onChange(filtrados.map((s) => s.matricula))}
        >
          Selecionar todos
        </Button>
        <Button type="button" variant="outline" onClick={() => onChange([])}>
          Limpar
        </Button>
      </div>

      <div className="flex items-center gap-2">
        <Badge variant="secondary">{selecionados.length} selecionado(s)</Badge>
        <span className="text-xs text-muted-foreground">{filtrados.length} servidor(es) listado(s)</span>
      </div>

      <div className="max-h-72 overflow-auto rounded-md border divide-y">
        {filtrados.length === 0 && (
          <div className="p-3 text-sm text-muted-foreground">Nenhum servidor encontrado.</div>
        )}
        {filtrados.map((s) => (
          <label
            key={s.matricula}
            className="flex items-center gap-3 p-2 cursor-pointer hover:bg-muted/50"
          >
            <Checkbox
              checked={selecionados.includes(s.matricula)}
              onCheckedChange={() => toggle(s.matricula)}
            />
            <div className="flex flex-col">
              <span className="text-sm font-medium">{s.nome}</span>
              <span className="text-xs text-muted-foreground">
                {s.matricula}
                {s.cpf ? ` • ${s.cpf}` : ""}
              </span>
            </div>
          </label>
        ))}
      </div>
    </div>
  );
};

export default MultiServerSelect;

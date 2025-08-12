import * as React from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import {
  Command,
  CommandEmpty,
  CommandGroup,
  CommandInput,
  CommandItem,
  CommandList,
} from "@/components/ui/command";
import { cn } from "@/lib/utils";

export type Servidor = {
  matricula: string;
  nome: string;
  cpf?: string;
  cargo?: string;
};

interface ServerComboboxProps {
  servidores: Servidor[];
  value?: string; // matricula selecionada
  onChange: (matricula?: string) => void;
  placeholder?: string;
}

export const ServerCombobox: React.FC<ServerComboboxProps> = ({
  servidores,
  value,
  onChange,
  placeholder = "Selecione um servidor",
}) => {
  const [open, setOpen] = React.useState(false);
  const selected = servidores.find((s) => s.matricula === value);

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button
          variant="outline"
          role="combobox"
          aria-expanded={open}
          className={cn("w-full justify-between", !selected && "text-muted-foreground")}
        >
          {selected ? `${selected.nome} (${selected.matricula})` : placeholder}
        </Button>
      </PopoverTrigger>
      <PopoverContent className="p-0 w-[var(--radix-popover-trigger-width)] z-50 bg-popover">
        <Command filter={(value, search) => {
          // custom filter: busca por nome, matrícula e cpf parcial
          const srv = servidores.find((s) => `${s.nome} (${s.matricula})` === value);
          const target = `${srv?.nome ?? ""} ${srv?.matricula ?? ""} ${srv?.cpf ?? ""}`.toLowerCase();
          return target.includes(search.toLowerCase()) ? 1 : 0;
        }}>
          <CommandInput placeholder="Buscar por nome, matrícula ou CPF..." />
          <CommandList>
            <CommandEmpty>Nenhum servidor encontrado.</CommandEmpty>
            <CommandGroup>
              {servidores.map((s) => (
                <CommandItem
                  key={s.matricula}
                  value={`${s.nome} (${s.matricula})`}
                  className="cursor-pointer"
                  onSelect={() => {
                    onChange(s.matricula);
                    setOpen(false);
                  }}
                >
                  <div className="flex flex-col">
                    <span className="font-medium">{s.nome}</span>
                    <span className="text-xs text-muted-foreground">{s.matricula}{s.cpf ? ` • ${s.cpf}` : ""}</span>
                  </div>
                </CommandItem>
              ))}
            </CommandGroup>
          </CommandList>
        </Command>
      </PopoverContent>
    </Popover>
  );
};

export default ServerCombobox;

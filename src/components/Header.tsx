import React from "react";
import { useData } from "@/context/DataContext";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { toast } from "@/components/ui/sonner";
import {
  AlertCircle,
  CheckCircle2,
  Save,
  Loader2,
  LogOut,
  Shield,
} from "lucide-react";

interface HeaderProps {
  onLogout?: () => void;
}

export const Header: React.FC<HeaderProps> = ({ onLogout }) => {
  const { hasUnsavedChanges, isSaving, salvarDados } = useData();

  const handleSave = async () => {
    try {
      await salvarDados();
      toast.success("Dados salvos com sucesso!", {
        description:
          "Todos os lançamentos e relatórios foram gravados localmente.",
      });
    } catch {
      toast.error("Erro ao salvar dados", {
        description: "Não foi possível gravar os dados localmente.",
      });
    }
  };

  return (
    <header className="border-b bg-gradient-to-br from-background to-card/60">
      <div className="container py-6 md:py-8">
        <div className="flex flex-col md:flex-row md:items-start md:justify-between gap-4">
          <div>
            <h1 className="text-2xl md:text-3xl font-bold tracking-tight">
              GEOPS - Gerador de Operações Especiais Segep
            </h1>
            <p className="text-muted-foreground mt-1 text-sm md:text-base">
              Sistema de Geração e Controle de Operações Especiais.
            </p>
          </div>

          <div className="flex items-center gap-2 md:gap-3 shrink-0">
            {hasUnsavedChanges ? (
              <Badge
                variant="outline"
                className="bg-amber-100 text-amber-900 border-amber-300 hover:bg-amber-100 gap-1.5 px-2.5 py-1"
              >
                <AlertCircle
                  className="h-3.5 w-3.5 animate-pulse"
                  aria-hidden="true"
                />
                <span className="hidden sm:inline">Alterações pendentes</span>
                <span className="sm:hidden">Pendentes</span>
              </Badge>
            ) : (
              <Badge
                variant="outline"
                className="bg-emerald-100 text-emerald-900 border-emerald-300 hover:bg-emerald-100 gap-1.5 px-2.5 py-1"
              >
                <CheckCircle2 className="h-3.5 w-3.5" aria-hidden="true" />
                <span className="hidden sm:inline">Dados salvos</span>
                <span className="sm:hidden">Salvos</span>
              </Badge>
            )}

            <TooltipProvider delayDuration={300}>
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    onClick={handleSave}
                    disabled={isSaving || !hasUnsavedChanges}
                    size="sm"
                    className="gap-2"
                  >
                    {isSaving ? (
                      <Loader2
                        className="h-4 w-4 animate-spin"
                        aria-hidden="true"
                      />
                    ) : (
                      <Save className="h-4 w-4" aria-hidden="true" />
                    )}
                    <span className="hidden md:inline">
                      {isSaving ? "Salvando..." : "Salvar Dados"}
                    </span>
                    <span className="sr-only md:hidden">
                      Salvar Dados (Ctrl+S)
                    </span>
                  </Button>
                </TooltipTrigger>
                <TooltipContent side="bottom">
                  <p>Salvar Dados (Ctrl+S)</p>
                </TooltipContent>
              </Tooltip>
            </TooltipProvider>

            {onLogout && (
              <Button
                onClick={onLogout}
                variant="outline"
                size="sm"
                className="gap-2"
              >
                <LogOut className="h-4 w-4" aria-hidden="true" />
                <span className="hidden md:inline">Sair</span>
              </Button>
            )}

            <div className="hidden lg:flex items-center gap-2 text-right pl-2 border-l">
              <Shield className="h-4 w-4 text-muted-foreground" />
              <div>
                <div className="text-sm font-medium">RCPPJ</div>
                <div className="text-xs text-muted-foreground">Acesso master</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </header>
  );
};

export default Header;

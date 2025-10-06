import { useState } from "react";
import { AlertCircle, CheckCircle2, Download, ArrowLeft } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { toast } from "sonner";

interface ErrorDetail {
  row: number;
  col: string;
  type: string;
  value: string;
  proposed?: string;
  reason?: string;
}

interface AnalysisViewProps {
  fileName: string;
  errors: ErrorDetail[];
  onDownload: () => void;
  onBack: () => void;
}

export const AnalysisView = ({ fileName, errors, onDownload, onBack }: AnalysisViewProps) => {
  const [selectedFixes, setSelectedFixes] = useState<Set<number>>(new Set());

  const errorCounts = errors.reduce((acc, error) => {
    acc[error.type] = (acc[error.type] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const toggleFix = (index: number) => {
    const newSelected = new Set(selectedFixes);
    if (newSelected.has(index)) {
      newSelected.delete(index);
    } else {
      newSelected.add(index);
    }
    setSelectedFixes(newSelected);
  };

  const selectAll = () => {
    setSelectedFixes(new Set(errors.map((_, i) => i)));
    toast.success("Todas as correções selecionadas");
  };

  const handleApplyFixes = () => {
    if (selectedFixes.size === 0) {
      toast.error("Selecione pelo menos uma correção");
      return;
    }
    toast.success(`${selectedFixes.size} correções aplicadas com sucesso!`);
    setTimeout(() => onDownload(), 1000);
  };

  return (
    <section className="min-h-screen bg-gradient-to-br from-background via-muted/10 to-background px-4 py-16">
      <div className="container max-w-6xl mx-auto">
        <div className="mb-8">
          <Button variant="ghost" onClick={onBack} className="mb-4">
            <ArrowLeft className="w-4 h-4 mr-2" />
            Voltar
          </Button>
          
          <div className="flex items-start justify-between">
            <div>
              <h1 className="text-4xl font-bold mb-2">Análise Completa</h1>
              <p className="text-muted-foreground">{fileName}</p>
            </div>
            
            <div className="flex gap-2">
              <Button variant="outline" onClick={selectAll}>
                Selecionar Todas
              </Button>
              <Button 
                className="bg-gradient-success hover:opacity-90"
                onClick={handleApplyFixes}
                disabled={selectedFixes.size === 0}
              >
                <CheckCircle2 className="w-4 h-4 mr-2" />
                Aplicar {selectedFixes.size > 0 && `(${selectedFixes.size})`}
              </Button>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
          <Card className="p-6 shadow-card">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-muted-foreground">Total de Erros</p>
                <p className="text-3xl font-bold">{errors.length}</p>
              </div>
              <AlertCircle className="w-8 h-8 text-destructive" />
            </div>
          </Card>

          {Object.entries(errorCounts).map(([type, count]) => (
            <Card key={type} className="p-6 shadow-card">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm text-muted-foreground">#{type}</p>
                  <p className="text-3xl font-bold">{count}</p>
                </div>
                <Badge variant="destructive">{type}</Badge>
              </div>
            </Card>
          ))}
        </div>

        <Tabs defaultValue="all" className="space-y-6">
          <TabsList className="bg-muted/50">
            <TabsTrigger value="all">Todos ({errors.length})</TabsTrigger>
            {Object.entries(errorCounts).map(([type, count]) => (
              <TabsTrigger key={type} value={type}>
                #{type} ({count})
              </TabsTrigger>
            ))}
          </TabsList>

          <TabsContent value="all" className="space-y-4">
            {errors.map((error, index) => (
              <Card 
                key={index} 
                className={`p-6 shadow-card transition-all cursor-pointer ${
                  selectedFixes.has(index) ? 'ring-2 ring-primary bg-primary/5' : ''
                }`}
                onClick={() => error.proposed && toggleFix(index)}
              >
                <div className="flex items-start justify-between gap-4">
                  <div className="flex-1 space-y-3">
                    <div className="flex items-center gap-3">
                      <Badge variant="destructive">#{error.type}</Badge>
                      <span className="text-sm text-muted-foreground">
                        Célula: {error.col}{error.row}
                      </span>
                    </div>
                    
                    <div className="space-y-2">
                      <div>
                        <p className="text-sm font-medium text-muted-foreground mb-1">Valor Atual:</p>
                        <code className="block p-3 bg-destructive/10 rounded-md text-sm font-mono">
                          {error.value}
                        </code>
                      </div>
                      
                      {error.proposed && (
                        <>
                          <div>
                            <p className="text-sm font-medium text-muted-foreground mb-1">Correção Proposta:</p>
                            <code className="block p-3 bg-accent/10 rounded-md text-sm font-mono">
                              {error.proposed}
                            </code>
                          </div>
                          
                          {error.reason && (
                            <div className="flex items-start gap-2 p-3 bg-muted/50 rounded-md">
                              <AlertCircle className="w-4 h-4 text-muted-foreground mt-0.5" />
                              <p className="text-sm text-muted-foreground">{error.reason}</p>
                            </div>
                          )}
                        </>
                      )}
                    </div>
                  </div>
                  
                  {error.proposed && (
                    <div className="flex items-center">
                      <input
                        type="checkbox"
                        checked={selectedFixes.has(index)}
                        onChange={() => toggleFix(index)}
                        className="w-5 h-5 rounded border-2 border-primary text-primary focus:ring-primary"
                        onClick={(e) => e.stopPropagation()}
                      />
                    </div>
                  )}
                </div>
              </Card>
            ))}
          </TabsContent>

          {Object.keys(errorCounts).map((type) => (
            <TabsContent key={type} value={type} className="space-y-4">
              {errors
                .filter((error) => error.type === type)
                .map((error, index) => (
                  <Card 
                    key={index} 
                    className={`p-6 shadow-card transition-all cursor-pointer ${
                      selectedFixes.has(errors.indexOf(error)) ? 'ring-2 ring-primary bg-primary/5' : ''
                    }`}
                    onClick={() => error.proposed && toggleFix(errors.indexOf(error))}
                  >
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex-1 space-y-3">
                        <div className="flex items-center gap-3">
                          <Badge variant="destructive">#{error.type}</Badge>
                          <span className="text-sm text-muted-foreground">
                            Célula: {error.col}{error.row}
                          </span>
                        </div>
                        
                        <div className="space-y-2">
                          <div>
                            <p className="text-sm font-medium text-muted-foreground mb-1">Valor Atual:</p>
                            <code className="block p-3 bg-destructive/10 rounded-md text-sm font-mono">
                              {error.value}
                            </code>
                          </div>
                          
                          {error.proposed && (
                            <>
                              <div>
                                <p className="text-sm font-medium text-muted-foreground mb-1">Correção Proposta:</p>
                                <code className="block p-3 bg-accent/10 rounded-md text-sm font-mono">
                                  {error.proposed}
                                </code>
                              </div>
                              
                              {error.reason && (
                                <div className="flex items-start gap-2 p-3 bg-muted/50 rounded-md">
                                  <AlertCircle className="w-4 h-4 text-muted-foreground mt-0.5" />
                                  <p className="text-sm text-muted-foreground">{error.reason}</p>
                                </div>
                              )}
                            </>
                          )}
                        </div>
                      </div>
                      
                      {error.proposed && (
                        <div className="flex items-center">
                          <input
                            type="checkbox"
                            checked={selectedFixes.has(errors.indexOf(error))}
                            onChange={() => toggleFix(errors.indexOf(error))}
                            className="w-5 h-5 rounded border-2 border-primary text-primary focus:ring-primary"
                            onClick={(e) => e.stopPropagation()}
                          />
                        </div>
                      )}
                    </div>
                  </Card>
                ))}
            </TabsContent>
          ))}
        </Tabs>

        <div className="mt-8 flex justify-end gap-4">
          <Button variant="outline" onClick={onBack}>
            Cancelar
          </Button>
          <Button 
            className="bg-gradient-success hover:opacity-90"
            onClick={handleApplyFixes}
            disabled={selectedFixes.size === 0}
          >
            <Download className="w-4 h-4 mr-2" />
            Aplicar e Baixar ({selectedFixes.size})
          </Button>
        </div>
      </div>
    </section>
  );
};

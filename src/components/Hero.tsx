import { Upload, CheckCircle2, Download, Zap } from "lucide-react";
import { Button } from "@/components/ui/button";

interface HeroProps {
  onGetStarted: () => void;
}

export const Hero = ({ onGetStarted }: HeroProps) => {
  return (
    <section className="relative min-h-screen flex items-center justify-center overflow-hidden bg-gradient-to-br from-background via-primary/5 to-accent/5">
      <div className="absolute inset-0 bg-grid-pattern opacity-5"></div>
      
      <div className="container px-4 py-16 mx-auto relative z-10">
        <div className="max-w-4xl mx-auto text-center space-y-8 animate-fade-in">
          <div className="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-primary/10 text-primary text-sm font-medium mb-4">
            <Zap className="w-4 h-4" />
            <span>Correção Automática de Fórmulas Excel</span>
          </div>
          
          <h1 className="text-5xl md:text-7xl font-bold tracking-tight">
            Corrija erros em suas
            <span className="block mt-2 bg-gradient-primary bg-clip-text text-transparent">
              planilhas automaticamente
            </span>
          </h1>
          
          <p className="text-xl md:text-2xl text-muted-foreground max-w-2xl mx-auto">
            Detecte e corrija erros como #REF!, #VALOR!, #NOME? em segundos. 
            Economize horas de trabalho manual com IA.
          </p>
          
          <div className="flex flex-col sm:flex-row gap-4 justify-center pt-4">
            <Button 
              size="lg" 
              className="text-lg px-8 py-6 bg-gradient-primary hover:opacity-90 shadow-glow"
              onClick={onGetStarted}
            >
              <Upload className="w-5 h-5 mr-2" />
              Começar Agora - Grátis
            </Button>
            <Button 
              size="lg" 
              variant="outline"
              className="text-lg px-8 py-6"
            >
              Ver Demo
            </Button>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8 pt-16 max-w-3xl mx-auto">
            <div className="flex flex-col items-center gap-3 p-6 rounded-xl bg-card/50 backdrop-blur border shadow-card">
              <div className="w-12 h-12 rounded-full bg-primary/10 flex items-center justify-center">
                <Upload className="w-6 h-6 text-primary" />
              </div>
              <h3 className="font-semibold text-lg">Upload Simples</h3>
              <p className="text-sm text-muted-foreground text-center">
                Faça upload do seu arquivo .xlsx
              </p>
            </div>
            
            <div className="flex flex-col items-center gap-3 p-6 rounded-xl bg-card/50 backdrop-blur border shadow-card">
              <div className="w-12 h-12 rounded-full bg-accent/10 flex items-center justify-center">
                <CheckCircle2 className="w-6 h-6 text-accent" />
              </div>
              <h3 className="font-semibold text-lg">Análise Inteligente</h3>
              <p className="text-sm text-muted-foreground text-center">
                IA detecta e propõe correções
              </p>
            </div>
            
            <div className="flex flex-col items-center gap-3 p-6 rounded-xl bg-card/50 backdrop-blur border shadow-card">
              <div className="w-12 h-12 rounded-full bg-primary/10 flex items-center justify-center">
                <Download className="w-6 h-6 text-primary" />
              </div>
              <h3 className="font-semibold text-lg">Download Corrigido</h3>
              <p className="text-sm text-muted-foreground text-center">
                Baixe sua planilha sem erros
              </p>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};

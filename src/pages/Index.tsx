import { useState } from "react";
import { Hero } from "@/components/Hero";
import { FileUpload } from "@/components/FileUpload";
import { AnalysisView } from "@/components/AnalysisView";
import { SheetSelector } from "@/components/SheetSelector";
import { analyzeExcelFile, applyFixes } from "@/utils/excelAnalyzer";
import { getSheetNames } from "@/utils/excelUtils";
import { toast } from "sonner";
import { saveAs } from "file-saver";
import { Button } from "@/components/ui/button";

type View = "hero" | "upload" | "sheet-selection" | "analysis";

interface ErrorDetail {
  row: number;
  col: string;
  type: string;
  value: string;
  proposed?: string;
  reason?: string;
  sheet?: string;
  severity?: 'critical' | 'warning' | 'info';
}

const Index = () => {
  const [currentView, setCurrentView] = useState<View>("hero");
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [errors, setErrors] = useState<ErrorDetail[]>([]);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);

  const handleGetStarted = () => {
    setCurrentView("upload");
  };

  const handleFileSelect = async (file: File) => {
    setSelectedFile(file);
    
    try {
      const sheets = await getSheetNames(file);
      setAvailableSheets(sheets);
      
      if (sheets.length === 1) {
        // Se só tem uma aba, analisa direto
        setSelectedSheets(sheets);
        await performAnalysis(file, sheets);
      } else {
        // Se tem múltiplas abas, mostra seletor
        setSelectedSheets(sheets); // Seleciona todas por padrão
        setCurrentView("sheet-selection");
      }
    } catch (error) {
      console.error("Erro ao ler arquivo:", error);
      toast.error("Erro ao ler o arquivo. Verifique o formato e tente novamente.");
    }
  };

  const handleAnalyzeSheets = async () => {
    if (!selectedFile || selectedSheets.length === 0) {
      toast.error("Selecione pelo menos uma aba para analisar");
      return;
    }
    
    await performAnalysis(selectedFile, selectedSheets);
  };

  const performAnalysis = async (file: File, sheets: string[]) => {
    setIsAnalyzing(true);
    
    try {
      const detectedErrors = await analyzeExcelFile(file, sheets);
      
      if (detectedErrors.length === 0) {
        toast.success("Nenhum erro detectado! Sua planilha está perfeita! 🎉");
        setCurrentView("upload");
        setSelectedFile(null);
      } else {
        setErrors(detectedErrors);
        setCurrentView("analysis");
        toast.success(`${detectedErrors.length} erro(s) detectado(s) em ${sheets.length} aba(s)!`);
      }
    } catch (error) {
      console.error("Erro ao analisar arquivo:", error);
      toast.error("Erro ao analisar o arquivo. Verifique o formato e tente novamente.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleDownload = async () => {
    if (!selectedFile) return;

    try {
      // Coletar índices dos erros que têm propostas
      const fixableIndices = errors
        .map((error, index) => error.proposed ? index : -1)
        .filter(index => index !== -1);

      const selectedFixes = new Set(fixableIndices);
      
      const fixedBlob = await applyFixes(selectedFile, selectedFixes, errors);
      const fileName = selectedFile.name.replace(/\.xlsx?$/i, '_corrigido.xlsx');
      
      saveAs(fixedBlob, fileName);
      toast.success("Arquivo corrigido baixado com sucesso! 🎉");
      
      // Reset após download
      setTimeout(() => {
        setCurrentView("upload");
        setSelectedFile(null);
        setErrors([]);
      }, 2000);
    } catch (error) {
      console.error("Erro ao aplicar correções:", error);
      toast.error("Erro ao gerar arquivo corrigido. Tente novamente.");
    }
  };

  const handleBack = () => {
    setCurrentView("upload");
    setSelectedFile(null);
    setErrors([]);
  };

  if (isAnalyzing) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-background">
        <div className="text-center space-y-4">
          <div className="w-16 h-16 border-4 border-primary border-t-transparent rounded-full animate-spin mx-auto"></div>
          <h2 className="text-2xl font-bold">Analisando sua planilha...</h2>
          <p className="text-muted-foreground">Isso pode levar alguns segundos</p>
        </div>
      </div>
    );
  }

  return (
    <>
      {currentView === "hero" && <Hero onGetStarted={handleGetStarted} />}
      {currentView === "upload" && <FileUpload onFileSelect={handleFileSelect} />}
      {currentView === "sheet-selection" && (
        <section className="min-h-screen flex items-center justify-center bg-gradient-to-br from-background via-muted/20 to-background px-4 py-16">
          <div className="w-full max-w-3xl space-y-6">
            <div className="text-center space-y-2 mb-8">
              <h2 className="text-3xl font-bold">Selecione as Abas</h2>
              <p className="text-muted-foreground">
                Escolha quais abas deseja analisar em busca de erros
              </p>
            </div>
            
            <SheetSelector
              sheets={availableSheets}
              selectedSheets={selectedSheets}
              onSelectionChange={setSelectedSheets}
            />
            
            <div className="flex gap-4 justify-center">
              <Button variant="outline" onClick={handleBack}>
                Voltar
              </Button>
              <Button 
                size="lg"
                className="bg-gradient-primary hover:opacity-90"
                onClick={handleAnalyzeSheets}
                disabled={selectedSheets.length === 0}
              >
                Analisar {selectedSheets.length > 0 && `(${selectedSheets.length} aba${selectedSheets.length > 1 ? 's' : ''})`}
              </Button>
            </div>
          </div>
        </section>
      )}
      {currentView === "analysis" && selectedFile && (
        <AnalysisView
          fileName={selectedFile.name}
          errors={errors}
          onDownload={handleDownload}
          onBack={handleBack}
        />
      )}
    </>
  );
};

export default Index;

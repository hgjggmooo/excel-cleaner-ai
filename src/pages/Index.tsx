import { useState } from "react";
import { Hero } from "@/components/Hero";
import { FileUpload } from "@/components/FileUpload";
import { AnalysisView } from "@/components/AnalysisView";
import { analyzeExcelFile, applyFixes } from "@/utils/excelAnalyzer";
import { toast } from "sonner";
import { saveAs } from "file-saver";

type View = "hero" | "upload" | "analysis";

interface ErrorDetail {
  row: number;
  col: string;
  type: string;
  value: string;
  proposed?: string;
  reason?: string;
}

const Index = () => {
  const [currentView, setCurrentView] = useState<View>("hero");
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [errors, setErrors] = useState<ErrorDetail[]>([]);
  const [isAnalyzing, setIsAnalyzing] = useState(false);

  const handleGetStarted = () => {
    setCurrentView("upload");
  };

  const handleFileSelect = async (file: File) => {
    setSelectedFile(file);
    setIsAnalyzing(true);
    
    try {
      const detectedErrors = await analyzeExcelFile(file);
      
      if (detectedErrors.length === 0) {
        toast.success("Nenhum erro detectado! Sua planilha está perfeita! 🎉");
        setCurrentView("upload");
        setSelectedFile(null);
      } else {
        setErrors(detectedErrors);
        setCurrentView("analysis");
        toast.success(`${detectedErrors.length} erro(s) detectado(s)!`);
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

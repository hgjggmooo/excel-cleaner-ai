import { useCallback, useState } from "react";
import { useDropzone } from "react-dropzone";
import { Upload, FileSpreadsheet, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";

interface FileUploadProps {
  onFileSelect: (file: File) => void;
}

export const FileUpload = ({ onFileSelect }: FileUploadProps) => {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      const file = acceptedFiles[0];
      setSelectedFile(file);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false,
  });

  const handleAnalyze = () => {
    if (selectedFile) {
      onFileSelect(selectedFile);
    }
  };

  const handleRemove = () => {
    setSelectedFile(null);
  };

  return (
    <section className="min-h-screen flex items-center justify-center bg-gradient-to-br from-background via-muted/20 to-background px-4 py-16">
      <Card className="w-full max-w-2xl p-8 shadow-glow">
        <div className="space-y-6">
          <div className="text-center space-y-2">
            <h2 className="text-3xl font-bold">Upload sua Planilha</h2>
            <p className="text-muted-foreground">
              Suportamos arquivos .xlsx e .xls até 10MB
            </p>
          </div>

          <div
            {...getRootProps()}
            className={`
              border-2 border-dashed rounded-xl p-12 text-center cursor-pointer
              transition-all duration-300
              ${isDragActive 
                ? 'border-primary bg-primary/5 scale-105' 
                : 'border-border hover:border-primary/50 hover:bg-muted/50'
              }
            `}
          >
            <input {...getInputProps()} />
            
            {!selectedFile ? (
              <div className="space-y-4">
                <div className="w-16 h-16 mx-auto rounded-full bg-primary/10 flex items-center justify-center">
                  <Upload className="w-8 h-8 text-primary" />
                </div>
                
                {isDragActive ? (
                  <p className="text-lg font-medium text-primary">
                    Solte o arquivo aqui...
                  </p>
                ) : (
                  <>
                    <p className="text-lg font-medium">
                      Arraste e solte seu arquivo aqui
                    </p>
                    <p className="text-sm text-muted-foreground">
                      ou clique para selecionar
                    </p>
                  </>
                )}
              </div>
            ) : (
              <div className="space-y-4">
                <div className="w-16 h-16 mx-auto rounded-full bg-accent/10 flex items-center justify-center">
                  <FileSpreadsheet className="w-8 h-8 text-accent" />
                </div>
                
                <div className="space-y-2">
                  <p className="text-lg font-medium">{selectedFile.name}</p>
                  <p className="text-sm text-muted-foreground">
                    {(selectedFile.size / 1024 / 1024).toFixed(2)} MB
                  </p>
                </div>
                
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={(e) => {
                    e.stopPropagation();
                    handleRemove();
                  }}
                  className="mx-auto"
                >
                  <X className="w-4 h-4 mr-2" />
                  Remover
                </Button>
              </div>
            )}
          </div>

          {selectedFile && (
            <Button
              size="lg"
              className="w-full bg-gradient-primary hover:opacity-90"
              onClick={handleAnalyze}
            >
              Analisar Planilha
            </Button>
          )}

          <div className="grid grid-cols-3 gap-4 pt-4 text-center text-sm text-muted-foreground">
            <div>
              <p className="font-semibold text-foreground">Seguro</p>
              <p>100% privado</p>
            </div>
            <div>
              <p className="font-semibold text-foreground">Rápido</p>
              <p>Análise em segundos</p>
            </div>
            <div>
              <p className="font-semibold text-foreground">Preciso</p>
              <p>IA avançada</p>
            </div>
          </div>
        </div>
      </Card>
    </section>
  );
};

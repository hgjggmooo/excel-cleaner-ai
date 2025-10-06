import { useState } from "react";
import { Checkbox } from "@/components/ui/checkbox";
import { Card } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Sheet } from "lucide-react";

interface SheetSelectorProps {
  sheets: string[];
  selectedSheets: string[];
  onSelectionChange: (sheets: string[]) => void;
}

export const SheetSelector = ({ sheets, selectedSheets, onSelectionChange }: SheetSelectorProps) => {
  const handleToggle = (sheetName: string) => {
    if (selectedSheets.includes(sheetName)) {
      onSelectionChange(selectedSheets.filter(s => s !== sheetName));
    } else {
      onSelectionChange([...selectedSheets, sheetName]);
    }
  };

  const handleSelectAll = () => {
    if (selectedSheets.length === sheets.length) {
      onSelectionChange([]);
    } else {
      onSelectionChange(sheets);
    }
  };

  return (
    <Card className="p-6 shadow-card">
      <div className="space-y-4">
        <div className="flex items-center justify-between">
          <h3 className="text-lg font-semibold">Selecione as Abas para Analisar</h3>
          <button
            onClick={handleSelectAll}
            className="text-sm text-primary hover:underline"
          >
            {selectedSheets.length === sheets.length ? 'Desmarcar Todas' : 'Selecionar Todas'}
          </button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          {sheets.map((sheetName) => (
            <div
              key={sheetName}
              className="flex items-center space-x-3 p-3 rounded-lg border hover:bg-muted/50 transition-colors cursor-pointer"
              onClick={() => handleToggle(sheetName)}
            >
              <Checkbox
                id={sheetName}
                checked={selectedSheets.includes(sheetName)}
                onCheckedChange={() => handleToggle(sheetName)}
                onClick={(e) => e.stopPropagation()}
              />
              <Label
                htmlFor={sheetName}
                className="flex items-center gap-2 cursor-pointer flex-1"
              >
                <Sheet className="w-4 h-4 text-muted-foreground" />
                <span className="font-medium">{sheetName}</span>
              </Label>
            </div>
          ))}
        </div>

        <p className="text-sm text-muted-foreground">
          {selectedSheets.length} de {sheets.length} aba(s) selecionada(s)
        </p>
      </div>
    </Card>
  );
};
import * as XLSX from 'xlsx';

export const getSheetNames = (file: File): Promise<string[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        resolve(workbook.SheetNames);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('Erro ao ler o arquivo'));
    reader.readAsBinaryString(file);
  });
};

export const getSheetStatistics = (file: File): Promise<{
  sheetName: string;
  rowCount: number;
  colCount: number;
  formulaCount: number;
}[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary', cellFormula: true });
        
        const stats = workbook.SheetNames.map(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
          
          let formulaCount = 0;
          for (let R = range.s.r; R <= range.e.r; R++) {
            for (let C = range.s.c; C <= range.e.c; C++) {
              const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: C })];
              if (cell && cell.f) formulaCount++;
            }
          }
          
          return {
            sheetName,
            rowCount: range.e.r - range.s.r + 1,
            colCount: range.e.c - range.s.c + 1,
            formulaCount,
          };
        });
        
        resolve(stats);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('Erro ao ler o arquivo'));
    reader.readAsBinaryString(file);
  });
};
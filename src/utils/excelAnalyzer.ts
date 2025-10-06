import * as XLSX from 'xlsx';

interface ErrorDetail {
  row: number;
  col: string;
  type: string;
  value: string;
  proposed?: string;
  reason?: string;
}

const ERROR_PATTERNS = {
  REF: /#REF!/gi,
  VALUE: /#VALUE!/gi,
  NAME: /#NAME\?/gi,
  DIV: /#DIV\/0!/gi,
  NULL: /#NULL!/gi,
  NA: /#N\/A/gi,
};

const VLOOKUP_REGEX = /VLOOKUP\s*\(/gi;
const PROCV_REGEX = /PROCV\s*\(/gi;

export const analyzeExcelFile = async (file: File): Promise<ErrorDetail[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary', cellFormula: true });
        const errors: ErrorDetail[] = [];

        // Analisar primeira planilha por padrão
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Percorrer todas as células
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        for (let R = range.s.r; R <= range.e.r; R++) {
          for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];

            if (!cell) continue;

            const cellValue = cell.v?.toString() || '';
            const cellFormula = cell.f || '';
            
            // Detectar erros no valor
            for (const [errorType, pattern] of Object.entries(ERROR_PATTERNS)) {
              if (pattern.test(cellValue) || pattern.test(cellFormula)) {
                const error: ErrorDetail = {
                  row: R + 1,
                  col: XLSX.utils.encode_col(C),
                  type: errorType,
                  value: cellFormula ? `=${cellFormula}` : cellValue,
                };

                // Propor correção baseada no tipo de erro
                const proposal = proposeFixForError(error, cellFormula, worksheet, range);
                if (proposal) {
                  error.proposed = proposal.formula;
                  error.reason = proposal.reason;
                }

                errors.push(error);
                break;
              }
            }
          }
        }

        resolve(errors);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('Erro ao ler o arquivo'));
    reader.readAsBinaryString(file);
  });
};

const proposeFixForError = (
  error: ErrorDetail,
  formula: string,
  worksheet: XLSX.WorkSheet,
  range: XLSX.Range
): { formula: string; reason: string } | null => {
  if (!formula) return null;

  const upperFormula = formula.toUpperCase();

  // Correção para #REF! em VLOOKUP/PROCV
  if (error.type === 'REF' && (VLOOKUP_REGEX.test(upperFormula) || PROCV_REGEX.test(upperFormula))) {
    const lastCol = XLSX.utils.encode_col(range.e.c);
    const lastRow = range.e.r + 1;
    const suggestedRange = `A1:${lastCol}${lastRow}`;
    
    // Tentar extrair argumentos da fórmula
    const match = formula.match(/\(([^)]+)\)/);
    if (match) {
      const args = match[1].split(',').map(arg => arg.trim());
      if (args.length >= 3) {
        const proposed = `VLOOKUP(${args[0]}, ${suggestedRange}, ${args[2]}, FALSE)`;
        return {
          formula: `=${proposed}`,
          reason: 'Substituído o intervalo quebrado (#REF!) por um range dinâmico baseado nos dados disponíveis.',
        };
      }
    }
  }

  // Correção para #VALUE! - envolver com IFERROR
  if (error.type === 'VALUE') {
    return {
      formula: `=IFERROR(${formula}, 0)`,
      reason: 'Encapsulado com IFERROR para tratar o erro #VALUE! e retornar 0 em caso de falha.',
    };
  }

  // Correção para #NAME? - possível nome de função errado
  if (error.type === 'NAME') {
    // Tentar corrigir nomes comuns
    let fixedFormula = formula;
    const commonFixes: Record<string, string> = {
      'SOMA': 'SUM',
      'SE': 'IF',
      'CONT.SE': 'COUNTIF',
      'SOMASE': 'SUMIF',
    };

    for (const [wrong, correct] of Object.entries(commonFixes)) {
      const regex = new RegExp(`\\b${wrong}\\b`, 'gi');
      if (regex.test(fixedFormula)) {
        fixedFormula = fixedFormula.replace(regex, correct);
        return {
          formula: `=${fixedFormula}`,
          reason: `Nome de função corrigido de ${wrong} para ${correct}.`,
        };
      }
    }
  }

  // Correção para #DIV/0! - envolver com IFERROR
  if (error.type === 'DIV') {
    return {
      formula: `=IFERROR(${formula}, "N/A")`,
      reason: 'Encapsulado com IFERROR para evitar divisão por zero.',
    };
  }

  return null;
};

export const applyFixes = (
  file: File,
  selectedFixes: Set<number>,
  errors: ErrorDetail[]
): Promise<Blob> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary', cellFormula: true });
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Aplicar as correções selecionadas
        selectedFixes.forEach((index) => {
          const error = errors[index];
          if (error.proposed) {
            const cellAddress = `${error.col}${error.row}`;
            const cell = worksheet[cellAddress];
            
            if (cell) {
              // Remover o '=' inicial se existir
              const formula = error.proposed.startsWith('=') 
                ? error.proposed.substring(1) 
                : error.proposed;
              
              cell.f = formula;
              delete cell.v; // Remover valor para forçar recálculo
            }
          }
        });

        // Gerar novo arquivo
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
        
        // Converter para Blob
        const buf = new ArrayBuffer(wbout.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < wbout.length; i++) {
          view[i] = wbout.charCodeAt(i) & 0xFF;
        }
        
        resolve(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('Erro ao processar o arquivo'));
    reader.readAsBinaryString(file);
  });
};

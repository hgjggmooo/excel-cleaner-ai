import * as XLSX from 'xlsx';

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

interface ValidationError extends ErrorDetail {
  validationType?: 'type' | 'merged' | 'format' | 'cross-sheet';
}

const ERROR_PATTERNS = {
  REF: /#REF!/gi,
  VALUE: /#VALUE!/gi,
  NAME: /#NAME\?/gi,
  DIV: /#DIV\/0!/gi,
  NULL: /#NULL!/gi,
  NA: /#N\/A/gi,
};

// Padrões para detecção de tipos de dados
const NUMBER_PATTERN = /^-?\d+\.?\d*$/;
const DATE_PATTERN = /^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/;
const CURRENCY_PATTERN = /^[R$€£¥]\s*-?\d+[.,]?\d*$/;

const VLOOKUP_REGEX = /VLOOKUP\s*\(/gi;
const PROCV_REGEX = /PROCV\s*\(/gi;

export const analyzeExcelFile = async (file: File, selectedSheets?: string[]): Promise<ErrorDetail[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary', cellFormula: true, cellNF: true });
        const errors: ErrorDetail[] = [];

        // Determinar quais sheets analisar
        const sheetsToAnalyze = selectedSheets && selectedSheets.length > 0 
          ? selectedSheets 
          : workbook.SheetNames;

        console.log('📊 Analisando sheets:', sheetsToAnalyze);

        // Analisar cada sheet selecionado
        sheetsToAnalyze.forEach(sheetName => {
          if (!workbook.Sheets[sheetName]) return;
          
          const worksheet = workbook.Sheets[sheetName];
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
          
          console.log(`🔍 Analisando "${sheetName}" - Range: ${XLSX.utils.encode_range(range)}`);
          
          // Detectar células mescladas
          const mergedCells = worksheet['!merges'] || [];
          
          let cellsChecked = 0;
          let formulaCells = 0;
          
          // Análise de fórmulas e erros
          for (let R = range.s.r; R <= range.e.r; R++) {
            for (let C = range.s.c; C <= range.e.c; C++) {
              const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
              const cell = worksheet[cellAddress];

              if (!cell) continue;
              
              cellsChecked++;

              const cellValue = cell.v?.toString() || '';
              const cellFormula = cell.f || '';
              const cellType = cell.t; // tipo: 'n' (number), 's' (string), 'e' (error), etc.
              const cellError = cell.w; // valor formatado
              
              if (cellFormula) formulaCells++;
              
              // CRÍTICO: Verificar se a célula tem tipo 'e' (error)
              if (cellType === 'e') {
                console.log(`❌ Erro detectado em ${cellAddress}:`, cellValue, cellFormula);
                
                let errorType = 'VALUE';
                if (cellValue.includes('REF')) errorType = 'REF';
                else if (cellValue.includes('DIV')) errorType = 'DIV';
                else if (cellValue.includes('NAME')) errorType = 'NAME';
                else if (cellValue.includes('NULL')) errorType = 'NULL';
                else if (cellValue.includes('N/A')) errorType = 'NA';
                
                const error: ErrorDetail = {
                  row: R + 1,
                  col: XLSX.utils.encode_col(C),
                  type: errorType,
                  value: cellFormula ? `=${cellFormula}` : cellValue,
                  sheet: sheetName,
                  severity: errorType === 'REF' || errorType === 'DIV' ? 'critical' : 'warning',
                };

                const proposal = proposeFixForError(error, cellFormula, worksheet, range, workbook);
                if (proposal) {
                  error.proposed = proposal.formula;
                  error.reason = proposal.reason;
                }

                errors.push(error);
                continue;
              }
              
              // 1. Detectar erros de fórmula via padrões de texto
              for (const [errorType, pattern] of Object.entries(ERROR_PATTERNS)) {
                if (pattern.test(cellValue) || pattern.test(cellFormula)) {
                  console.log(`⚠️ Erro de padrão detectado em ${cellAddress}:`, errorType);
                  
                  const error: ErrorDetail = {
                    row: R + 1,
                    col: XLSX.utils.encode_col(C),
                    type: errorType,
                    value: cellFormula ? `=${cellFormula}` : cellValue,
                    sheet: sheetName,
                    severity: errorType === 'REF' || errorType === 'DIV' ? 'critical' : 'warning',
                  };

                  const proposal = proposeFixForError(error, cellFormula, worksheet, range, workbook);
                  if (proposal) {
                    error.proposed = proposal.formula;
                    error.reason = proposal.reason;
                  }

                  errors.push(error);
                  break;
                }
              }
              
              // 2. Validação de tipos de dados
              if (cellValue && !cellFormula) {
                const typeError = validateDataType(cell, R, C, sheetName, worksheet, range);
                if (typeError) errors.push(typeError);
              }
            }
          }
          
          console.log(`✅ "${sheetName}": ${cellsChecked} células verificadas, ${formulaCells} fórmulas`);
          
          // 3. Validar células mescladas problemáticas
          mergedCells.forEach(merge => {
            const mergeError = validateMergedCell(merge, sheetName, worksheet);
            if (mergeError) errors.push(mergeError);
          });
          
          // 4. Detectar referências cruzadas quebradas entre sheets
          const crossSheetErrors = validateCrossSheetReferences(worksheet, sheetName, workbook, range);
          errors.push(...crossSheetErrors);
        });

        console.log(`🎯 Total de erros detectados: ${errors.length}`);
        resolve(errors);
      } catch (error) {
        console.error('💥 Erro ao analisar:', error);
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('Erro ao ler o arquivo'));
    reader.readAsBinaryString(file);
  });
};

// Nova função: Validar tipo de dados
const validateDataType = (
  cell: XLSX.CellObject,
  row: number,
  col: number,
  sheetName: string,
  worksheet: XLSX.WorkSheet,
  range: XLSX.Range
): ErrorDetail | null => {
  const cellValue = cell.v?.toString() || '';
  const colLetter = XLSX.utils.encode_col(col);
  
  // Detectar coluna de números com texto
  let numericCount = 0;
  let textCount = 0;
  
  for (let R = range.s.r; R <= Math.min(range.s.r + 20, range.e.r); R++) {
    const checkCell = worksheet[XLSX.utils.encode_cell({ r: R, c: col })];
    if (!checkCell || checkCell.f) continue;
    
    const val = checkCell.v?.toString() || '';
    if (NUMBER_PATTERN.test(val)) numericCount++;
    else if (val.length > 0) textCount++;
  }
  
  // Se coluna é majoritariamente numérica mas tem texto
  if (numericCount > 5 && textCount < numericCount / 3) {
    if (!NUMBER_PATTERN.test(cellValue) && cellValue.length > 0) {
      return {
        row: row + 1,
        col: colLetter,
        type: 'TYPE_MISMATCH',
        value: cellValue,
        sheet: sheetName,
        severity: 'warning',
        reason: 'Texto detectado em coluna numérica. Isso pode causar erros em cálculos.',
      };
    }
  }
  
  return null;
};

// Nova função: Validar células mescladas
const validateMergedCell = (
  merge: XLSX.Range,
  sheetName: string,
  worksheet: XLSX.WorkSheet
): ErrorDetail | null => {
  const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
  const endCell = XLSX.utils.encode_cell({ r: merge.e.r, c: merge.e.c });
  
  // Verificar se célula mesclada contém fórmulas (pode causar problemas)
  const cell = worksheet[startCell];
  if (cell && cell.f) {
    return {
      row: merge.s.r + 1,
      col: XLSX.utils.encode_col(merge.s.c),
      type: 'MERGED_FORMULA',
      value: `Mesclado ${startCell}:${endCell}`,
      sheet: sheetName,
      severity: 'warning',
      reason: 'Célula mesclada contém fórmula, o que pode causar comportamento inesperado.',
    };
  }
  
  return null;
};

// Nova função: Validar referências entre sheets
const validateCrossSheetReferences = (
  worksheet: XLSX.WorkSheet,
  sheetName: string,
  workbook: XLSX.WorkBook,
  range: XLSX.Range
): ErrorDetail[] => {
  const errors: ErrorDetail[] = [];
  const sheetRefPattern = /['"]?([^'"!]+)['"]?!/g;
  
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = worksheet[cellAddress];
      
      if (!cell || !cell.f) continue;
      
      const formula = cell.f;
      let match;
      
      while ((match = sheetRefPattern.exec(formula)) !== null) {
        const referencedSheet = match[1];
        
        // Verificar se o sheet referenciado existe
        if (!workbook.SheetNames.includes(referencedSheet)) {
          errors.push({
            row: R + 1,
            col: XLSX.utils.encode_col(C),
            type: 'CROSS_SHEET_REF',
            value: `=${formula}`,
            sheet: sheetName,
            severity: 'critical',
            reason: `Referência a sheet inexistente: "${referencedSheet}"`,
            proposed: `=${formula.replace(match[0], `${workbook.SheetNames[0]}!`)}`,
          });
        }
      }
    }
  }
  
  return errors;
};

const proposeFixForError = (
  error: ErrorDetail,
  formula: string,
  worksheet: XLSX.WorkSheet,
  range: XLSX.Range,
  workbook?: XLSX.WorkBook
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
        
        // Aplicar as correções selecionadas em todas as sheets
        selectedFixes.forEach((index) => {
          const error = errors[index];
          if (error.proposed) {
            // Determinar qual sheet usar (default para a primeira se não especificado)
            const sheetName = error.sheet || workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            if (!worksheet) {
              console.warn(`Sheet "${sheetName}" não encontrada. Pulando correção.`);
              return;
            }
            
            const cellAddress = `${error.col}${error.row}`;
            const cell = worksheet[cellAddress];
            
            if (cell) {
              // Remover o '=' inicial se existir
              const formula = error.proposed.startsWith('=') 
                ? error.proposed.substring(1) 
                : error.proposed;
              
              cell.f = formula;
              delete cell.v; // Remover valor para forçar recálculo
            } else {
              // Criar célula se não existir
              worksheet[cellAddress] = {
                f: error.proposed.startsWith('=') 
                  ? error.proposed.substring(1) 
                  : error.proposed,
                t: 'n'
              };
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

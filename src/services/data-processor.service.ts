
import { Injectable, signal } from '@angular/core';
import ExcelJS from 'exceljs';

declare const XLSX: any;

export interface ProcessingState {
  stage: string;
  progress: number;
  logs: string[];
}

@Injectable({
  providedIn: 'root'
})
export class DataProcessorService {
  private _state = signal<ProcessingState>({ stage: '空闲', progress: 0, logs: [] });
  state = this._state.asReadonly();

  // --- High Contrast Semantic Color Palette (Hex without #) ---
  private readonly COLORS = {
    HEADER_FIXED: "FF334155",    // Slate 700 (Dark Grey)
    HEADER_TOTAL: "FF1D4ED8",    // Blue 700 (Strong Blue)
    HEADER_SMALL4: "FF0F766E",   // Teal 700 (Strong Teal)
    HEADER_OTHER: "FF7C3AED",    // Violet 600 (Strong Purple)
    HEADER_SUMMARY: "FF0F172A",  // Slate 900 (Black-ish)
    
    ROW_ODD: "FFFFFFFF",         // White
    ROW_EVEN: "FFF8FAFC",        // Slate 50 (Very light grey)
    
    // Highlight colors
    ROW_HIGHLIGHT_GRADE: "FFDBEAFE", // Blue 100 (Light Blue for Year Grade)
    ROW_HIGHLIGHT_TOTAL: "FFF3E8FF", // Purple 100 (Light Purple for Total Score rows)
    
    BORDER: "FF94A3B8"           // Slate 400 (Visible Border)
  };

  reset() {
    this._state.set({ stage: '空闲', progress: 0, logs: [] });
  }

  private log(message: string) {
    this._state.update(s => ({
      ...s,
      logs: [...s.logs, `[${new Date().toLocaleTimeString()}] ${message}`]
    }));
  }

  private setProgress(stage: string, progress: number) {
    this._state.update(s => ({ ...s, stage, progress }));
  }

  // Helper to read file as Base64 string
  private readFileAsBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        resolve(reader.result as string);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async process(jsonFile: File, excelFile: File | null, watermarkFile: File | null) {
    try {
      this.setProgress('正在初始化', 0);
      this.log('启动处理引擎...');

      // 1. Read JSON
      this.setProgress('读取数据中', 10);
      const jsonText = await jsonFile.text();
      let rawData: any[] = JSON.parse(jsonText);
      if (!Array.isArray(rawData) || rawData.length === 0) {
        throw new Error('无效的 JSON 文件：必须是非空数组。');
      }
      this.log(`从 JSON 加载了 ${rawData.length} 条记录。`);

      // 2. Read Knowledge Base
      let kbData: any[] = [];
      if (excelFile) {
        this.log('正在读取知识库 Excel...');
        const buffer = await excelFile.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        kbData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        this.log(`从知识库加载了 ${kbData.length} 条记录。`);
      }

      // 3. Identification Phase
      this.setProgress('语义分析中', 20);
      const idField = this.inferIdField(rawData);
      this.log(`检测到唯一标识字段 (ID): "${idField}"`);

      // 4. Merge Data
      this.setProgress('数据合并与清洗', 30);
      const mergedData = this.mergeData(rawData, kbData, idField);

      // 5. Subject Detection
      this.setProgress('识别学科', 40);
      const { subjects, small4Subjects, otherSubjects } = this.identifySubjects(mergedData, idField);
      this.log(`检测到学科: ${subjects.join(', ')}`);
      
      // 6. Data Calculation
      this.setProgress('计算成绩', 50);
      const processedData = await this.processInChunks(mergedData, subjects, small4Subjects);

      // 7. Process Watermark (If exists)
      let watermarkBuffer: ArrayBuffer | null = null;
      let watermarkExtension: 'png' | 'jpeg' = 'png';
      if (watermarkFile) {
        this.setProgress('处理背景水印', 70);
        this.log('正在注入背景水印并调整透明度...');
        watermarkBuffer = await watermarkFile.arrayBuffer();
        if (watermarkFile.type === 'image/jpeg' || watermarkFile.name.endsWith('.jpg') || watermarkFile.name.endsWith('.jpeg')) {
          watermarkExtension = 'jpeg';
        }
      }

      // 8. Generate Excel Workbook using ExcelJS
      this.setProgress('生成工作簿', 80);
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'ScoreMaster Agent';
      workbook.lastModifiedBy = 'ScoreMaster Agent';
      workbook.created = new Date();
      workbook.modified = new Date();

      // -- Sheet 1: Year Grade Total Rank --
      this.log('生成工作表: 全年级总分排行');
      this.buildGradeTotalSheetExcelJS(workbook, processedData, subjects.length >= 2, subjects.length === 1 && small4Subjects.length > 0, watermarkBuffer, watermarkExtension);

      // -- Sheet 2..N: Subject Ranks --
      if (subjects.length >= 2) {
        for (const subj of subjects) {
          const isSmall4 = small4Subjects.includes(subj);
          this.buildSubjectSheetExcelJS(workbook, processedData, subj, isSmall4, watermarkBuffer, watermarkExtension);
        }
      }

      // -- Sheet: Class Ranks (Complex Styling) --
      const classes = [...new Set(processedData.map(d => d.class))].sort();
      this.log(`正在为 ${classes.length} 个班级生成分班表...`);
      for (const cls of classes) {
        const classData = processedData.filter(d => d.class === cls);
        this.buildClassSheetExcelJS(workbook, classData, subjects, small4Subjects, otherSubjects, cls, watermarkBuffer, watermarkExtension);
      }

      // -- Sheet: Summary --
      this.log('生成多维统计报告...');
      this.buildDetailedSummarySheetExcelJS(workbook, processedData, subjects, small4Subjects, watermarkBuffer, watermarkExtension);

      // 9. Output
      this.setProgress('正在完成', 95);
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      
      this.setProgress('完成', 100);
      return url;

    } catch (e: any) {
      this.log(`错误: ${e.message}`);
      this.setProgress('错误', 0);
      throw e;
    }
  }

  // --- Logic Helpers ---

  private inferIdField(data: any[]): string {
    if (data.length === 0) return '';
    const keys = Object.keys(data[0]);
    const commonIdPatterns = [/考号/i, /学号/i, /ID/i, /^id$/i, /student_id/i, /uid/i, /no/i, /code/i];
    for (const pattern of commonIdPatterns) {
      const found = keys.find(k => pattern.test(k));
      if (found) return found;
    }
    for (const key of keys) {
      const values = data.map(item => item[key]);
      const uniqueValues = new Set(values);
      if (uniqueValues.size === data.length) {
        const type = typeof values[0];
        if (type === 'string' || type === 'number') return key;
      }
    }
    return keys[0];
  }

  private identifySubjects(data: any[], idField: string) {
    const record = data[0];
    const keys = Object.keys(record);
    const excludeCombos = /(理化|史政|生地|政史|化生|物化|地政|文综|理综)/i;
    const excludeKeywords = /(总分|total|score_sum|备注|状态|排名|序号|id|考号|学号|姓名|班级|年级|_raw|class|grade|name)/i;

    const subjects = keys.filter(key => {
      if (key === idField) return false;
      const val = record[key];
      const num = Number(val);
      if (isNaN(num) || val === '' || val === null) return false;
      if (excludeKeywords.test(key)) return false;
      if (excludeCombos.test(key)) return false;
      return true;
    });

    const small4Keywords = ['政治', '地理', '化学', '生物', 'politics', 'geography', 'chemistry', 'biology'];
    const small4Subjects = subjects.filter(s => small4Keywords.some(kw => s.toLowerCase().includes(kw)));
    const otherSubjects = subjects.filter(s => !small4Subjects.includes(s));

    return { subjects, small4Subjects, otherSubjects };
  }

  private mergeData(raw: any[], kb: any[], rawId: string): any[] {
    let kbId = '';
    if (kb.length > 0) kbId = this.inferIdField(kb);
    const kbMap = new Map();
    if (kbId) kb.forEach(row => kbMap.set(String(row[kbId]), row));

    const namePattern = /(姓名|name|xm)/i;
    const classPattern = /(班级|class|bj|bjmc)/i;
    const gradePattern = /(年级|grade|nj)/i;

    return raw.map(row => {
      const idVal = String(row[rawId]);
      const kbRow = kbMap.get(idVal) || {};
      const combined = { ...row, ...kbRow };
      const keys = Object.keys(combined);

      const nameKey = keys.find(k => namePattern.test(k));
      const classKey = keys.find(k => classPattern.test(k));
      const gradeKey = keys.find(k => gradePattern.test(k));

      return {
        _raw: row,
        id: idVal,
        name: combined[nameKey || ''] || row['姓名'] || idVal,
        class: combined[classKey || ''] || row['班级'] || '未分类',
        grade: combined[gradeKey || ''] || row['年级'] || '未分类',
        ...row 
      };
    });
  }

  private async processInChunks(data: any[], subjects: string[], small4: string[], skipScoring: boolean = false) {
    const result = [...data];
    const CHUNK_SIZE = 5000;
    
    const stats: any = {};
    if (!skipScoring) {
      small4.forEach(subj => {
        const values = data.map(d => Number(d[subj]) || 0);
        stats[subj] = { min: Math.min(...values), max: Math.max(...values) };
      });
    }

    for (let i = 0; i < result.length; i += CHUNK_SIZE) {
      if (i > 0) await new Promise(r => setTimeout(r, 0));
      const end = Math.min(i + CHUNK_SIZE, result.length);

      for (let j = i; j < end; j++) {
        const row = result[j];
        const calculatedTotal = subjects.reduce((acc, s) => acc + (Number(row[s]) || 0), 0);
        const totalKeys = /(总分|total|score_sum)/i;
        const providedTotalKey = Object.keys(row).find(k => totalKeys.test(k));
        let rawTotal = calculatedTotal;
        if (providedTotalKey) {
           const pt = Number(row[providedTotalKey]);
           if (Math.abs(pt - calculatedTotal) <= 2) rawTotal = pt;
        }
        row._rawTotal = rawTotal;

        let assignedTotal = rawTotal;
        
        if (!skipScoring) {
          const rawSmall4Sum = small4.reduce((acc, s) => acc + (Number(row[s]) || 0), 0);
          assignedTotal -= rawSmall4Sum;

          small4.forEach(subj => {
            const s = Number(row[subj]) || 0;
            const { min, max } = stats[subj];
            let assigned = s;
            if (max !== min) assigned = Math.round(40 + ((s - min) / (max - min)) * 60);
            row[`assigned_${subj}`] = assigned;
            assignedTotal += assigned;
          });
        } else {
          // Skip scoring: assigned values = raw values
          small4.forEach(subj => {
            const s = Number(row[subj]) || 0;
            row[`assigned_${subj}`] = s;
          });
          // assignedTotal remains equal to rawTotal
        }
        
        row._assignedTotal = assignedTotal;
      }
    }

    this.calcRanks(result, '_rawTotal', 'yearRank_raw');
    this.calcRanks(result, '_assignedTotal', 'yearRank_assigned');
    subjects.forEach(s => {
      this.calcRanks(result, s, `yearRank_${s}`);
      if (small4.includes(s)) this.calcRanks(result, `assigned_${s}`, `yearRank_assigned_${s}`);
    });

    const byClass: any = {};
    result.forEach(r => {
      if (!byClass[r.class]) byClass[r.class] = [];
      byClass[r.class].push(r);
    });

    Object.values(byClass).forEach((rows: any) => {
      this.calcRanks(rows, '_rawTotal', 'classRank_raw');
      this.calcRanks(rows, '_assignedTotal', 'classRank_assigned');
      subjects.forEach(s => {
        this.calcRanks(rows, s, `classRank_${s}`);
        if (small4.includes(s)) this.calcRanks(rows, `assigned_${s}`, `classRank_assigned_${s}`);
      });
    });

    return result;
  }

  private calcRanks(data: any[], key: string, rankKey: string) {
    data.sort((a, b) => (Number(b[key]) || 0) - (Number(a[key]) || 0));
    let rank = 1;
    for (let i = 0; i < data.length; i++) {
      if (i > 0 && (Number(data[i][key]) || 0) < (Number(data[i - 1][key]) || 0)) {
        rank = i + 1;
      }
      data[i][rankKey] = rank;
    }
  }

  // --- Styling Helpers ---

  // Modified: Accepts null for fillHex to support transparency
  private setStyle(cell: any, fillHex: string | null, fontColorHex: string, isBold: boolean, isHeader: boolean) {
    cell.s = {
      font: { 
        name: 'Microsoft YaHei', 
        sz: isHeader ? 11 : 10, 
        bold: isBold, 
        color: { rgb: fontColorHex } 
      },
      alignment: { 
        vertical: "center", 
        horizontal: "center", 
        wrapText: true 
      },
      border: {
        top: { style: "thin", color: { rgb: this.COLORS.BORDER } },
        bottom: { style: "thin", color: { rgb: this.COLORS.BORDER } },
        left: { style: "thin", color: { rgb: this.COLORS.BORDER } },
        right: { style: "thin", color: { rgb: this.COLORS.BORDER } }
      }
    };
    
    // Only apply fill if color is provided. If null (watermark mode), leave it undefined (transparent)
    if (fillHex) {
       cell.s.fill = { fgColor: { rgb: fillHex } };
    }
  }

  // Modified: accepts hasWatermark to determine if we should skip row coloring
  private applyTableBodyStyle(ws: any, rowCount: number, colCount: number, startRow: number = 1, hasWatermark: boolean = false) {
    // If watermark exists, we MUST NOT set any fill color (bg) to allow the image to show through.
    // However, we still want to apply borders and alignment.
    
    for (let r = startRow; r < rowCount + startRow; r++) {
       // If watermark exists, force null (transparent). Otherwise use alternating colors.
       const fill = hasWatermark ? null : (r % 2 === 0 ? this.COLORS.ROW_EVEN : this.COLORS.ROW_ODD);
       
       for (let c = 0; c < colCount; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          if (!ws[addr]) continue;
          
          // Apply style with or without fill
          this.setStyle(ws[addr], fill, "000000", false, false);
       }
    }
  }

  private applyTableHeadStyle(ws: any, row: number, colStart: number, colEnd: number, colorHex: string) {
    for (let c = colStart; c <= colEnd; c++) {
      const addr = XLSX.utils.encode_cell({ r: row, c });
      if (!ws[addr]) continue;
      this.setStyle(ws[addr], colorHex, "FFFFFF", true, true);
    }
  }

  // --- ExcelJS Builders ---

  private buildGradeTotalSheetExcelJS(workbook: ExcelJS.Workbook, data: any[], multiSubject: boolean, singleSmall4: boolean, watermarkBuffer: ArrayBuffer | null, watermarkExtension: 'png' | 'jpeg') {
    const sheet = workbook.addWorksheet('全年级总分排行', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    data.sort((a, b) => a.yearRank_raw - b.yearRank_raw);

    const header = multiSubject 
      ? ['序号', '班级', '姓名', '原始总分', '', '', '', '序号', '班级', '姓名', '赋分总分']
      : singleSmall4 
        ? ['序号', '班级', '姓名', '原始成绩', '', '', '', '序号', '班级', '姓名', '赋分成绩']
        : ['序号', '班级', '姓名', '原始成绩'];
    
    sheet.addRow(header);

    const rawList = [...data].sort((a, b) => a.yearRank_raw - b.yearRank_raw);
    const assignedList = [...data].sort((a, b) => a.yearRank_assigned - b.yearRank_assigned);

    for (let i = 0; i < data.length; i++) {
      const r = rawList[i];
      const row = [i + 1, r.class, r.name, r._rawTotal];
      if (multiSubject || singleSmall4) {
        const a = assignedList[i];
        row.push('', '', '', i + 1, a.class, a.name, a._assignedTotal);
      }
      sheet.addRow(row);
    }

    // Set Column Widths
    sheet.columns = [
      { width: 8 }, { width: 14 }, { width: 14 }, { width: 12 }, 
      { width: 3 }, { width: 3 }, { width: 3 }, 
      { width: 8 }, { width: 14 }, { width: 14 }, { width: 12 }
    ];

    // Styling
    this.applySheetStylesExcelJS(sheet, header.length, this.COLORS.HEADER_TOTAL, watermarkBuffer);

    // Watermark
    if (watermarkBuffer) {
       const imageId = workbook.addImage({
         buffer: watermarkBuffer,
         extension: watermarkExtension,
       });
       sheet.addBackgroundImage(imageId);
    }
  }

  private buildSubjectSheetExcelJS(workbook: ExcelJS.Workbook, data: any[], subject: string, isSmall4: boolean, watermarkBuffer: ArrayBuffer | null, watermarkExtension: 'png' | 'jpeg') {
    const sheet = workbook.addWorksheet(`${subject}成绩排行`, {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    const rawKey = subject;
    const assignedKey = `assigned_${subject}`;

    const header = isSmall4
       ? ['序号', '班级', '姓名', '原始成绩', '', '', '', '序号', '班级', '姓名', '赋分成绩']
       : ['序号', '班级', '姓名', '原始成绩'];
    sheet.addRow(header);

    const rawList = [...data].sort((a, b) => (b[rawKey]||0) - (a[rawKey]||0) || String(a.id).localeCompare(String(b.id)));
    const assignedList = isSmall4 ? [...data].sort((a, b) => (b[assignedKey]||0) - (a[assignedKey]||0) || String(a.id).localeCompare(String(b.id))) : [];

    for(let i=0; i<data.length; i++) {
       const r = rawList[i];
       const row = [i+1, r.class, r.name, r[rawKey]];
       if(isSmall4) {
          const a = assignedList[i];
          row.push('', '', '', i+1, a.class, a.name, a[assignedKey]);
       }
       sheet.addRow(row);
    }

    // Set Column Widths
    sheet.columns = [
      { width: 8 }, { width: 14 }, { width: 14 }, { width: 12 }, 
      { width: 3 }, { width: 3 }, { width: 3 }, 
      { width: 8 }, { width: 14 }, { width: 14 }, { width: 12 }
    ];

    const color = isSmall4 ? this.COLORS.HEADER_SMALL4 : this.COLORS.HEADER_OTHER;
    this.applySheetStylesExcelJS(sheet, header.length, color, watermarkBuffer);

    if (watermarkBuffer) {
       const imageId = workbook.addImage({
         buffer: watermarkBuffer,
         extension: watermarkExtension,
       });
       sheet.addBackgroundImage(imageId);
    }
  }

  private buildClassSheetExcelJS(workbook: ExcelJS.Workbook, classData: any[], subjects: string[], small4: string[], otherSubjects: string[], className: string, watermarkBuffer: ArrayBuffer | null, watermarkExtension: 'png' | 'jpeg') {
    const sheet = workbook.addWorksheet(`${className}总分排行`, {
      views: [{ state: 'frozen', ySplit: 2 }]
    });

    classData.sort((a, b) => a.classRank_raw - b.classRank_raw);

    // Header 1
    const row1 = ['序号', '班级', '姓名', '总分汇总', '', '', '', '', '', ''];
    small4.forEach(s => {
      row1.push(s, '', '', '', '', '', '');
    });
    otherSubjects.forEach(s => {
      row1.push(s, '', '');
    });
    sheet.addRow(row1);

    // Header 2
    const row2 = ['', '', '', '原始分', '班排', '年排', '', '赋分', '班排', '年排'];
    small4.forEach(() => {
      row2.push('原始分', '班排', '年排', '', '赋分', '班排', '年排');
    });
    otherSubjects.forEach(() => {
      row2.push('得分', '班排', '年排');
    });
    sheet.addRow(row2);

    // Merges
    // Fixed cols
    sheet.mergeCells(1, 1, 2, 1);
    sheet.mergeCells(1, 2, 2, 2);
    sheet.mergeCells(1, 3, 2, 3);
    
    let col = 4;
    // Total
    sheet.mergeCells(1, col, 1, col + 6);
    col += 7;
    // Small4
    small4.forEach(() => {
      sheet.mergeCells(1, col, 1, col + 6);
      col += 7;
    });
    // Others
    otherSubjects.forEach(() => {
      sheet.mergeCells(1, col, 1, col + 2);
      col += 3;
    });

    // Data
    for(let i=0; i<classData.length; i++) {
      const r = classData[i];
      const row: any[] = [i+1, r.class, r.name];
      row.push(r._rawTotal, r.classRank_raw, r.yearRank_raw, '', r._assignedTotal, r.classRank_assigned, r.yearRank_assigned);
      for(const s of small4) {
         row.push(r[s], r[`classRank_${s}`], r[`yearRank_${s}`], '', r[`assigned_${s}`], r[`classRank_assigned_${s}`], r[`yearRank_assigned_${s}`]);
      }
      for(const s of otherSubjects) {
        row.push(r[s], r[`classRank_${s}`], r[`yearRank_${s}`]);
      }
      sheet.addRow(row);
    }

    // Styling
    // Colors map for headers
    // Fixed: 1-3
    // Total: 4-10
    // ...
    
    // We can iterate cells to apply styles
    // Or just apply to all and then override headers
    
    // Apply body styles first
    const rowCount = classData.length;
    const colCount = col; // last col index + 1 roughly

    // Basic styling for all cells
    sheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
         cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
         cell.border = {
           top: { style: 'thin', color: { argb: this.COLORS.BORDER } },
           left: { style: 'thin', color: { argb: this.COLORS.BORDER } },
           bottom: { style: 'thin', color: { argb: this.COLORS.BORDER } },
           right: { style: 'thin', color: { argb: this.COLORS.BORDER } }
         };
         cell.font = { name: 'Microsoft YaHei', size: 10 };

         if (rowNumber > 2) {
             // Body rows
             if (!watermarkBuffer) {
                const fillArgb = (rowNumber - 2) % 2 === 0 ? this.COLORS.ROW_EVEN : this.COLORS.ROW_ODD;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } };
             }
         } else {
             // Header rows
             cell.font = { name: 'Microsoft YaHei', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
             // Determine header color based on column index
             const c = Number(cell.col);
             let headerColor = this.COLORS.HEADER_FIXED;
             
             if (c > 3 && c <= 10) headerColor = this.COLORS.HEADER_TOTAL;
             else if (c > 10) {
                // Calculate offset from 11
                const offset = c - 11;
                const small4BlockSize = 7;
                const totalSmall4Cols = small4.length * small4BlockSize;
                
                if (offset < totalSmall4Cols) {
                   headerColor = this.COLORS.HEADER_SMALL4;
                } else {
                   headerColor = this.COLORS.HEADER_OTHER;
                }
             } else {
               headerColor = this.COLORS.HEADER_FIXED;
             }
             
             cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor } };
         }
      });
    });

    if (watermarkBuffer) {
       const imageId = workbook.addImage({
         buffer: watermarkBuffer,
         extension: watermarkExtension,
       });
       sheet.addBackgroundImage(imageId);
    }
  }

  private buildDetailedSummarySheetExcelJS(workbook: ExcelJS.Workbook, data: any[], subjects: string[], small4: string[], watermarkBuffer: ArrayBuffer | null, watermarkExtension: 'png' | 'jpeg') {
    const sheet = workbook.addWorksheet('成绩深度分析', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    const headers = ['统计群体', '项目', '参考人数', '平均分', '中位数', '最高分', '最低分', '高分名单 (Top 5)', '低分名单 (Bottom 5)'];
    sheet.addRow(headers);

    const classes = [...new Set(data.map(d => d.class))].sort();
    const groups = ['全年级', ...classes];

    const calcMedian = (values: number[]) => {
      if (values.length === 0) return 0;
      values.sort((a, b) => a - b);
      const half = Math.floor(values.length / 2);
      if (values.length % 2) return values[half];
      return (values[half - 1] + values[half]) / 2.0;
    };

    const getHolders = (dataset: any[], key: string, val: number, isMin = false) => {
      const matches = dataset.filter(d => Number(d[key]) === val);
      const count = matches.length;
      const names = matches.map(d => d.name);
      if (count === 0) return '-';
      if (count > 5) return `${names[0]}, ${names[1]} 等 ${count} 人`;
      return names.join('、');
    };

    for(const grp of groups) {
      const rows = grp === '全年级' ? data : data.filter(d => d.class === grp);
      const count = rows.length;
      
      // 1. Raw Total
      const rawTotals = rows.map(r => r._rawTotal);
      sheet.addRow([
        grp, '原始总分', count,
        (rawTotals.reduce((a,b)=>a+b,0)/count).toFixed(1),
        calcMedian(rawTotals).toFixed(1),
        Math.max(...rawTotals), Math.min(...rawTotals),
        getHolders(rows, '_rawTotal', Math.max(...rawTotals)),
        getHolders(rows, '_rawTotal', Math.min(...rawTotals), true)
      ]);

      // 2. Assigned Total
      if (small4.length > 0) {
        const assTotals = rows.map(r => r._assignedTotal);
        sheet.addRow([
          grp, '赋分总分', count,
          (assTotals.reduce((a,b)=>a+b,0)/count).toFixed(1),
          calcMedian(assTotals).toFixed(1),
          Math.max(...assTotals), Math.min(...assTotals),
          getHolders(rows, '_assignedTotal', Math.max(...assTotals)),
          getHolders(rows, '_assignedTotal', Math.min(...assTotals), true)
        ]);
      }

      // 3. Subjects
      for(const s of subjects) {
        const vals = rows.map(r => Number(r[s]) || 0);
        sheet.addRow([
          grp, s, count,
          (vals.reduce((a,b)=>a+b,0)/count).toFixed(1),
          calcMedian(vals).toFixed(1),
          Math.max(...vals), Math.min(...vals),
          getHolders(rows, s, Math.max(...vals)),
          getHolders(rows, s, Math.min(...vals), true)
        ]);
      }
    }

    // Column widths
    sheet.columns = [
      { width: 15 }, { width: 15 }, { width: 10 }, { width: 10 }, 
      { width: 10 }, { width: 10 }, { width: 10 }, { width: 40 }, { width: 40 }
    ];

    // Styling
    sheet.eachRow((row, rowNumber) => {
      const isHeader = rowNumber === 1;
      const rowValues = row.values as any[];
      const groupName = rowValues[1]; // ExcelJS row values are 1-based index? No, actually array. row.values[1] is col 1.
      const itemName = rowValues[2];

      row.eachCell((cell) => {
         cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
         cell.border = {
           top: { style: 'thin', color: { argb: this.COLORS.BORDER } },
           left: { style: 'thin', color: { argb: this.COLORS.BORDER } },
           bottom: { style: 'thin', color: { argb: this.COLORS.BORDER } },
           right: { style: 'thin', color: { argb: this.COLORS.BORDER } }
         };
         cell.font = { name: 'Microsoft YaHei', size: isHeader ? 11 : 10, bold: isHeader };

         if (isHeader) {
            cell.font.color = { argb: 'FFFFFFFF' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: this.COLORS.HEADER_SUMMARY } };
         } else {
            // Body
            let fillArgb: string | null = null;
            if (!watermarkBuffer) {
               fillArgb = (rowNumber - 1) % 2 === 0 ? this.COLORS.ROW_EVEN : this.COLORS.ROW_ODD;
               
               // Highlighting
               if (groupName === '全年级') fillArgb = this.COLORS.ROW_HIGHLIGHT_GRADE;
               else if (String(itemName).includes('总分')) fillArgb = this.COLORS.ROW_HIGHLIGHT_TOTAL;
               
               if (groupName === '全年级' || String(itemName).includes('总分')) cell.font.bold = true;
            } else {
               if (groupName === '全年级' || String(itemName).includes('总分')) cell.font.bold = true;
            }

            if (fillArgb) {
               cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } };
            }
         }
      });
    });

    if (watermarkBuffer) {
       const imageId = workbook.addImage({
         buffer: watermarkBuffer,
         extension: watermarkExtension,
       });
       sheet.addBackgroundImage(imageId);
    }
  }

  private applySheetStylesExcelJS(sheet: ExcelJS.Worksheet, headerCols: number, headerColor: string, watermarkBuffer: ArrayBuffer | null) {
     sheet.eachRow((row, rowNumber) => {
        const isHeader = rowNumber === 1;
        row.eachCell((cell) => {
           cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
           cell.border = {
             top: { style: 'thin', color: { argb: this.COLORS.BORDER } },
             left: { style: 'thin', color: { argb: this.COLORS.BORDER } },
             bottom: { style: 'thin', color: { argb: this.COLORS.BORDER } },
             right: { style: 'thin', color: { argb: this.COLORS.BORDER } }
           };
           cell.font = { 
             name: 'Microsoft YaHei', 
             size: isHeader ? 11 : 10, 
             bold: isHeader,
             color: { argb: isHeader ? 'FFFFFFFF' : 'FF000000' }
           };

           if (isHeader) {
             cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor } };
           } else {
             if (!watermarkBuffer) {
               const fillArgb = (rowNumber - 1) % 2 === 0 ? this.COLORS.ROW_EVEN : this.COLORS.ROW_ODD;
               cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } };
             }
           }
        });
     });
  }
}

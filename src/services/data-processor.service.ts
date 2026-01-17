
import { Injectable, signal } from '@angular/core';

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
    HEADER_FIXED: "334155",    // Slate 700 (Dark Grey)
    HEADER_TOTAL: "1D4ED8",    // Blue 700 (Strong Blue)
    HEADER_SMALL4: "0F766E",   // Teal 700 (Strong Teal)
    HEADER_OTHER: "7C3AED",    // Violet 600 (Strong Purple)
    HEADER_SUMMARY: "0F172A",  // Slate 900 (Black-ish)
    
    ROW_ODD: "FFFFFF",         // White
    ROW_EVEN: "F8FAFC",        // Slate 50 (Very light grey)
    
    // Highlight colors
    ROW_HIGHLIGHT_GRADE: "DBEAFE", // Blue 100 (Light Blue for Year Grade)
    ROW_HIGHLIGHT_TOTAL: "F3E8FF", // Purple 100 (Light Purple for Total Score rows)
    
    BORDER: "94A3B8"           // Slate 400 (Visible Border)
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
      let watermarkBase64: string | null = null;
      if (watermarkFile) {
        this.setProgress('处理背景水印', 70);
        this.log('正在注入背景水印并调整透明度...');
        watermarkBase64 = await this.readFileAsBase64(watermarkFile);
      }

      // 8. Generate Excel Workbook
      this.setProgress('生成工作簿', 80);
      const wb = XLSX.utils.book_new();

      // -- Sheet 1: Year Grade Total Rank --
      this.log('生成工作表: 全年级总分排行');
      const gradeTotalSheet = this.buildGradeTotalSheet(processedData, subjects.length >= 2, subjects.length === 1 && small4Subjects.length > 0, watermarkBase64);
      XLSX.utils.book_append_sheet(wb, gradeTotalSheet, '全年级总分排行');

      // -- Sheet 2..N: Subject Ranks --
      if (subjects.length >= 2) {
        for (const subj of subjects) {
          const isSmall4 = small4Subjects.includes(subj);
          const sheet = this.buildSubjectSheet(processedData, subj, isSmall4, watermarkBase64);
          XLSX.utils.book_append_sheet(wb, sheet, `${subj}成绩排行`);
        }
      }

      // -- Sheet: Class Ranks (Complex Styling) --
      const classes = [...new Set(processedData.map(d => d.class))].sort();
      this.log(`正在为 ${classes.length} 个班级生成分班表...`);
      for (const cls of classes) {
        const classData = processedData.filter(d => d.class === cls);
        const classSheet = this.buildClassSheet(classData, subjects, small4Subjects, otherSubjects, watermarkBase64);
        XLSX.utils.book_append_sheet(wb, classSheet, `${cls}总分排行`);
      }

      // -- Sheet: Summary --
      this.log('生成多维统计报告...');
      const summarySheet = this.buildDetailedSummarySheet(processedData, subjects, small4Subjects, watermarkBase64);
      XLSX.utils.book_append_sheet(wb, summarySheet, '成绩深度分析');

      // 9. Output
      this.setProgress('正在完成', 95);
      const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbOut], { type: 'application/octet-stream' });
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

  private async processInChunks(data: any[], subjects: string[], small4: string[]) {
    const result = [...data];
    const CHUNK_SIZE = 5000;
    
    const stats: any = {};
    small4.forEach(subj => {
      const values = data.map(d => Number(d[subj]) || 0);
      stats[subj] = { min: Math.min(...values), max: Math.max(...values) };
    });

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
    for (let r = startRow; r < rowCount + startRow; r++) {
       // If watermark exists, use null (transparent) fill, otherwise use alternating colors
       const fill = hasWatermark ? null : (r % 2 === 0 ? this.COLORS.ROW_EVEN : this.COLORS.ROW_ODD);
       for (let c = 0; c < colCount; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          if (!ws[addr]) continue;
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

  // --- Builders ---

  private buildGradeTotalSheet(data: any[], multiSubject: boolean, singleSmall4: boolean, watermarkBase64: string | null) {
    data.sort((a, b) => a.yearRank_raw - b.yearRank_raw);

    const aoa = [];
    const header = multiSubject 
      ? ['序号', '班级', '姓名', '原始总分', '', '', '', '序号', '班级', '姓名', '赋分总分']
      : singleSmall4 
        ? ['序号', '班级', '姓名', '原始成绩', '', '', '', '序号', '班级', '姓名', '赋分成绩']
        : ['序号', '班级', '姓名', '原始成绩'];
    aoa.push(header);

    const rawList = [...data].sort((a, b) => a.yearRank_raw - b.yearRank_raw);
    const assignedList = [...data].sort((a, b) => a.yearRank_assigned - b.yearRank_assigned);

    for (let i = 0; i < data.length; i++) {
      const r = rawList[i];
      const row = [i + 1, r.class, r.name, r._rawTotal];
      if (multiSubject || singleSmall4) {
        const a = assignedList[i];
        row.push('', '', '', i + 1, a.class, a.name, a._assignedTotal);
      }
      aoa.push(row);
    }
    
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    
    // Attempt to attach watermark (Library dependent, but data is ready)
    if (watermarkBase64) {
      ws['!backgroundImage'] = watermarkBase64; 
    }

    // Apply Styling
    // Header (Keep colors)
    this.applyTableHeadStyle(ws, 0, 0, header.length - 1, this.COLORS.HEADER_TOTAL);
    // Body (Remove colors if watermark exists)
    this.applyTableBodyStyle(ws, aoa.length - 1, header.length, 1, !!watermarkBase64);

    ws['!cols'] = [{wch:6}, {wch:12}, {wch:12}, {wch:10}, {wch:2}, {wch:2}, {wch:2}, {wch:6}, {wch:12}, {wch:12}, {wch:10}];
    return ws;
  }

  private buildSubjectSheet(data: any[], subject: string, isSmall4: boolean, watermarkBase64: string | null) {
     const aoa = [];
     const rawKey = subject;
     const assignedKey = `assigned_${subject}`;

     const header = isSmall4
        ? ['序号', '班级', '姓名', '原始成绩', '', '', '', '序号', '班级', '姓名', '赋分成绩']
        : ['序号', '班级', '姓名', '原始成绩'];
     aoa.push(header);

     const rawList = [...data].sort((a, b) => (b[rawKey]||0) - (a[rawKey]||0) || String(a.id).localeCompare(String(b.id)));
     const assignedList = isSmall4 ? [...data].sort((a, b) => (b[assignedKey]||0) - (a[assignedKey]||0) || String(a.id).localeCompare(String(b.id))) : [];

     for(let i=0; i<data.length; i++) {
        const r = rawList[i];
        const row = [i+1, r.class, r.name, r[rawKey]];
        if(isSmall4) {
           const a = assignedList[i];
           row.push('', '', '', i+1, a.class, a.name, a[assignedKey]);
        }
        aoa.push(row);
     }
     
     const ws = XLSX.utils.aoa_to_sheet(aoa);

     if (watermarkBase64) {
        ws['!backgroundImage'] = watermarkBase64;
     }
     
     // Styling
     const color = isSmall4 ? this.COLORS.HEADER_SMALL4 : this.COLORS.HEADER_OTHER;
     this.applyTableHeadStyle(ws, 0, 0, header.length - 1, color);
     this.applyTableBodyStyle(ws, aoa.length - 1, header.length, 1, !!watermarkBase64);

     ws['!cols'] = [{wch:6}, {wch:12}, {wch:12}, {wch:10}, {wch:2}, {wch:2}, {wch:2}, {wch:6}, {wch:12}, {wch:12}, {wch:10}];
     return ws;
  }

  private buildClassSheet(classData: any[], subjects: string[], small4: string[], otherSubjects: string[], watermarkBase64: string | null) {
    classData.sort((a, b) => a.classRank_raw - b.classRank_raw);
    const aoa: any[][] = [[], []]; 
    const merges: any[] = [];
    
    // -- Structure Definition --
    // Store column ranges for styling: { start, end, color }
    const styleZones: any[] = [];
    let colIdx = 0;

    // 1. Fixed Columns
    const fixedHeaders = ['序号', '班级', '姓名'];
    fixedHeaders.forEach((h, i) => {
      aoa[0].push(h); aoa[1].push(''); 
      merges.push({ s: { r: 0, c: i }, e: { r: 1, c: i } });
    });
    styleZones.push({ s: 0, e: 2, c: this.COLORS.HEADER_FIXED });
    colIdx = 3;

    // 2. Total Block
    aoa[0].push('总分汇总');
    for(let k=0; k<6; k++) aoa[0].push('');
    aoa[1].push('原始分', '班排', '年排', '', '赋分', '班排', '年排');
    merges.push({ s: { r: 0, c: colIdx }, e: { r: 0, c: colIdx + 6 } });
    styleZones.push({ s: colIdx, e: colIdx + 6, c: this.COLORS.HEADER_TOTAL });
    colIdx += 7;

    // 3. Small 4
    for (const s of small4) {
      aoa[0].push(s);
      for(let k=0; k<6; k++) aoa[0].push('');
      aoa[1].push('原始分', '班排', '年排', '', '赋分', '班排', '年排');
      merges.push({ s: { r: 0, c: colIdx }, e: { r: 0, c: colIdx + 6 } });
      styleZones.push({ s: colIdx, e: colIdx + 6, c: this.COLORS.HEADER_SMALL4 });
      colIdx += 7;
    }

    // 4. Others
    for (const s of otherSubjects) {
      aoa[0].push(s);
      for(let k=0; k<2; k++) aoa[0].push('');
      aoa[1].push('得分', '班排', '年排');
      merges.push({ s: { r: 0, c: colIdx }, e: { r: 0, c: colIdx + 2 } });
      styleZones.push({ s: colIdx, e: colIdx + 2, c: this.COLORS.HEADER_OTHER });
      colIdx += 3;
    }

    // -- Data Rows --
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
      aoa.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges;
    ws['!cols'] = new Array(colIdx).fill({wch: 10});
    
    if (watermarkBase64) {
        ws['!backgroundImage'] = watermarkBase64;
    }

    // -- Apply Styling --
    // Body
    this.applyTableBodyStyle(ws, aoa.length - 2, colIdx, 2, !!watermarkBase64);

    // Headers (Semantic Coloring) - Keep these always
    styleZones.forEach(zone => {
      this.applyTableHeadStyle(ws, 0, zone.s, zone.e, zone.c);
      this.applyTableHeadStyle(ws, 1, zone.s, zone.e, zone.c);
    });

    return ws;
  }

  private buildDetailedSummarySheet(data: any[], subjects: string[], small4: string[], watermarkBase64: string | null) {
    const aoa = [];
    const headers = ['统计群体', '项目', '参考人数', '平均分', '中位数', '最高分', '最低分', '高分名单 (Top 5)', '低分名单 (Bottom 5)'];
    aoa.push(headers);

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
      aoa.push([
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
        aoa.push([
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
        aoa.push([
          grp, s, count,
          (vals.reduce((a,b)=>a+b,0)/count).toFixed(1),
          calcMedian(vals).toFixed(1),
          Math.max(...vals), Math.min(...vals),
          getHolders(rows, s, Math.max(...vals)),
          getHolders(rows, s, Math.min(...vals), true)
        ]);
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{wch:15}, {wch:15}, {wch:10}, {wch:10}, {wch:10}, {wch:10}, {wch:10}, {wch:40}, {wch:40}];
    
    if (watermarkBase64) {
        ws['!backgroundImage'] = watermarkBase64;
    }

    // Custom Styling for Summary
    this.applyTableHeadStyle(ws, 0, 0, headers.length - 1, this.COLORS.HEADER_SUMMARY);
    
    // Body Logic (Highlighting)
    for (let r = 1; r < aoa.length; r++) {
       const rowData = aoa[r];
       const groupName = rowData[0];
       const itemName = rowData[1];
       
       // If watermark exists, default to transparent (null).
       // If no watermark, apply alternating colors
       let fill = watermarkBase64 ? null : (r % 2 === 0 ? this.COLORS.ROW_EVEN : this.COLORS.ROW_ODD);
       let bold = false;

       // Highlight Whole Grade
       if (groupName === '全年级') {
         if (!watermarkBase64) fill = this.COLORS.ROW_HIGHLIGHT_GRADE;
         bold = true;
       } 
       // Highlight Totals in class sections
       else if (String(itemName).includes('总分')) {
         if (!watermarkBase64) fill = this.COLORS.ROW_HIGHLIGHT_TOTAL;
         bold = true;
       }

       for (let c = 0; c < headers.length; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          if (!ws[addr]) continue;
          this.setStyle(ws[addr], fill, "000000", bold, false);
       }
    }

    return ws;
  }
}

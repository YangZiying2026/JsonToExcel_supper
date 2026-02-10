
import { Component, signal, inject, ViewChild, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataProcessorService } from './services/data-processor.service';
import { SafeUrl, SafeHtml, DomSanitizer } from '@angular/platform-browser';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './app.component.html',
})
export class AppComponent {
  dataService = inject(DataProcessorService);
  sanitizer: DomSanitizer = inject(DomSanitizer);

  // Optional because they are conditionally rendered via @if
  @ViewChild('jsonInput') jsonInput?: ElementRef<HTMLInputElement>;
  @ViewChild('excelInput') excelInput?: ElementRef<HTMLInputElement>;
  @ViewChild('watermarkInput') watermarkInput?: ElementRef<HTMLInputElement>;

  jsonFile = signal<File | null>(null);
  excelFile = signal<File | null>(null);
  watermarkFile = signal<File | null>(null);
  outputFilename = signal<string>('');
  downloadUrl = signal<SafeUrl | null>(null);
  
  // 分科统计模式开关
  isSubjectStatsMode = signal<boolean>(false);
  
  // 高级设置折叠状态
  isAdvancedSettingsOpen = signal<boolean>(false);

  // Markdown State
  mdFiles = signal<{title: string, filename: string}[]>([]);
  currentMdContent = signal<SafeHtml>('');
  currentMdIndex = signal<number>(0);

  constructor() {
    this.loadMdList();
  }

  async loadMdList() {
    try {
      // Use absolute path /markdown/... to ensure correct routing
      const res = await fetch('/markdown/list.json');
      if (res.ok) {
        const list = await res.json();
        this.mdFiles.set(list);
        if (list.length > 0) {
          this.selectMd(0);
        }
      } else {
        console.error('Markdown list not found. Status:', res.status);
        this.currentMdContent.set(this.sanitizer.bypassSecurityTrustHtml(
          `<div class="text-center p-4 text-slate-500">
             <p class="font-bold mb-2">无法加载文档配置</p>
             <p class="text-xs">如果您刚刚添加了此功能，请尝试<span class="text-red-500 font-bold">重启开发服务器</span>以应用静态资源配置。</p>
           </div>`
        ));
      }
    } catch (e) {
      console.error('Failed to load markdown list', e);
    }
  }

  async selectMd(index: number) {
    this.currentMdIndex.set(index);
    const file = this.mdFiles()[index];
    if (!file) return;
    try {
      const res = await fetch(`/markdown/${file.filename}`);
      if (res.ok) {
        const text = await res.text();
        const { marked } = await import('marked');
        const html = await marked.parse(text);
        this.currentMdContent.set(this.sanitizer.bypassSecurityTrustHtml(html as string));
      } else {
        this.currentMdContent.set(this.sanitizer.bypassSecurityTrustHtml(`<p class="text-red-500">无法加载文档内容 (Error ${res.status})</p>`));
      }
    } catch (e) {
      console.error('Failed to load markdown file', e);
      this.currentMdContent.set(this.sanitizer.bypassSecurityTrustHtml(`<p class="text-red-500">加载出错</p>`));
    }
  }

  // Helper to visualize progress step
  getStageIndex(stage: string): number {
    if (stage.includes('初始化') || stage.includes('读取')) return 0;
    if (stage.includes('语义') || stage.includes('合并') || stage.includes('识别')) return 1;
    if (stage.includes('计算')) return 2;
    if (stage.includes('生成') || stage.includes('完成')) return 3;
    return 0;
  }

  onJsonSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      const file = input.files[0];
      // 20MB limit
      if (file.size > 20 * 1024 * 1024) {
        alert('文件过大，请上传小于 20MB 的 JSON 文件。');
        input.value = '';
        return;
      }
      this.jsonFile.set(file);
    }
  }

  onExcelSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      const file = input.files[0];
      // 20MB limit
      if (file.size > 20 * 1024 * 1024) {
        alert('文件过大，请上传小于 20MB 的 Excel 文件。');
        input.value = '';
        return;
      }
      this.excelFile.set(file);
    }
  }

  onWatermarkSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      const file = input.files[0];
      // 10MB limit for images
      if (file.size > 10 * 1024 * 1024) {
        alert('图片过大，请上传小于 10MB 的图片文件。');
        input.value = '';
        return;
      }
      this.watermarkFile.set(file);
    }
  }

  // 切换分科统计模式
  toggleSubjectStatsMode() {
    this.isSubjectStatsMode.set(!this.isSubjectStatsMode());
  }

  toggleAdvancedSettings() {
    this.isAdvancedSettingsOpen.set(!this.isAdvancedSettingsOpen());
  }

  onFilenameInput(event: Event) {
    const input = event.target as HTMLInputElement;
    this.outputFilename.set(input.value.trim());
  }

  get finalFilename(): string {
    const name = this.outputFilename();
    if (!name) return 'ScoreMaster_Report.xlsx';
    return name.endsWith('.xlsx') ? name : `${name}.xlsx`;
  }

  async startProcessing() {
    if (!this.jsonFile()) return;
    try {
      const url = await this.dataService.process(
        this.jsonFile()!, 
        this.excelFile(), 
        this.watermarkFile(),
        this.isSubjectStatsMode()
      );
      this.downloadUrl.set(this.sanitizer.bypassSecurityTrustUrl(url));
    } catch (e) {
      console.error(e);
      // State is already handled in service
    }
  }

  reset() {
    // 1. Reset service internal state
    this.dataService.reset();
    
    // 2. Reset component local state
    this.jsonFile.set(null);
    this.excelFile.set(null);
    this.watermarkFile.set(null);
    this.outputFilename.set('');
    this.downloadUrl.set(null);

    // Note: We do NOT need to manually clear the input values here.
    // Because the inputs are inside an @if block that toggles based on state,
    // they are removed from DOM and re-created fresh (empty) when we return to 'Idle'.
    // Manual clearing caused a crash because ViewChildren are undefined in 'Done' state.
  }
}


import { Component, signal, inject, ViewChild, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DataProcessorService } from './services/data-processor.service';
import { SafeUrl, DomSanitizer } from '@angular/platform-browser';

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
  downloadUrl = signal<SafeUrl | null>(null);

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
      this.jsonFile.set(input.files[0]);
    }
  }

  onExcelSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.excelFile.set(input.files[0]);
    }
  }

  onWatermarkSelect(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.watermarkFile.set(input.files[0]);
    }
  }

  async startProcessing() {
    if (!this.jsonFile()) return;
    try {
      const url = await this.dataService.process(
        this.jsonFile()!, 
        this.excelFile(), 
        this.watermarkFile()
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
    this.downloadUrl.set(null);

    // Note: We do NOT need to manually clear the input values here.
    // Because the inputs are inside an @if block that toggles based on state,
    // they are removed from DOM and re-created fresh (empty) when we return to 'Idle'.
    // Manual clearing caused a crash because ViewChildren are undefined in 'Done' state.
  }
}

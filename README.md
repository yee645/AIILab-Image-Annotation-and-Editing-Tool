# Azure Kinect Annotator

Azure Kinect MKV 影片標記與剪輯工具，支援逐幀瀏覽、時間區間標記、CSV 匯出，以及透過 ffmpeg 直接剪輯影片。

## 功能

- 逐幀瀏覽 `.mkv` 影片（支援前進/後退 1 幀或 30 幀）
- 標記時間區間（start / end），暫存多段標記
- 將標記區間匯出為 CSV
- 擷取當前畫面為 PNG
- 根據標記區間直接剪輯影片（ffmpeg stream copy，不重新編碼）
- 深色 / 淺色主題切換

## 系統需求

- Windows 10 / 11
- [ffmpeg](https://www.gyan.dev/ffmpeg/builds/)（匯出剪輯功能需要，需加入系統 PATH）

## 使用方式

### 執行檔

直接執行 `dist/AIILab-Image-Annotation-and-Editing-Tool.exe`。

### 從原始碼執行

```bash
pip install opencv-python pillow
python annotator.py
```

## 快捷鍵

| 按鍵 | 功能 |
|------|------|
| A / Left | 後退 1 幀 |
| D / Right | 前進 1 幀 |
| W | 後退 30 幀 |
| S | 前進 30 幀 |
| Shift+A / Shift+Left | 跳到影片開頭 |
| Shift+D / Shift+Right | 跳到影片結尾 |
| J | 標記 start_us |
| K | 標記 end_us |
| Enter | 儲存區間到暫存清單 |
| P | 擷取目前畫面 |
| Ctrl+S | 立即將暫存標記寫入 CSV |
| Shift+S | 顯示目前所有已暫存的標記區間 |
| Ctrl+E | 將暫存區間匯出為剪輯影片 |
| Ctrl+Z | 復原上一筆標記區間 |
| N | 寫入 CSV 並載入下一部影片 |
| Backspace | 回到上一部影片 |
| Ctrl+Q / Esc | 退出 |

## 授權

MIT License

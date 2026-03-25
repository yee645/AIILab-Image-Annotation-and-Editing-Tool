# AIILab Image Annotation and Editing Tool

Azure Kinect MKV 影片標記與剪輯工具，支援逐幀瀏覽、時間區間標記、CSV 匯出，以及透過 mkvmerge 裁剪影片並完整保留所有軌道（Color、Depth、IR、IMU、Calibration）。

## 功能

- 逐幀瀏覽 `.mkv` 影片（支援前進/後退 1 幀或 30 幀）
- 標記時間區間（start / end），暫存多段標記
- 暫存標記跨影片保留，每段記錄來源影片與時間點
- 將標記區間匯出為 CSV
- 擷取當前畫面為 PNG
- Export：根據 CSV 或暫存標記裁剪當前影片（mkvmerge，保留所有軌道）
- Export All：一次匯出 CSV 中所有影片的標記區間
- 可自訂 CSV、擷取圖片、匯出影片的命名格式
- 深色 / 淺色主題切換

## 系統需求

- Windows 10 / 11
- [MKVToolNix](https://mkvtoolnix.download/)（匯出剪輯功能需要 mkvmerge）

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
| Shift+S | 顯示目前所有已暫存的標記區間（含來源影片） |
| Ctrl+E | 匯出剪輯影片（優先 CSV，無 CSV 則用暫存標記） |
| Ctrl+Z | 復原上一筆標記區間 |
| N | 載入下一部影片（暫存標記保留） |
| Backspace | 回到上一部影片 |
| Ctrl+Q / Esc | 退出 |

## 設定

透過 Settings 對話框可設定：

| 項目 | 預設值 | 可用變數 |
|------|--------|----------|
| CSV Filename | `annotation.csv` | `{filename}` |
| Image Filename | `{filename}_{frame}_{us}.png` | `{filename}`, `{frame}`, `{us}` |
| Export Filename | `{filename}_seg{segment}.mkv` | `{filename}`, `{segment}` |

## 授權

MIT License

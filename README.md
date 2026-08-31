# AegisOps 架構圖完全重繪版 PPTX

此 repository 產出一份 **16:9 單頁 PowerPoint (`.pptx`)**，以可編輯元件重建架構圖（不是整張 raster 背景圖）。

## 安裝

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 產生 PowerPoint

```bash
python generate_pptx.py
```

輸出檔案：

- `output/aegisops_mvp_redraw.pptx`

## 驗證

```bash
python validate_pptx.py --file output/aegisops_mvp_redraw.pptx
```

驗證會檢查：

1. 檔案存在
2. 可被 ZIP / OpenXML 解析
3. 含大量可編輯 shape 與 text run
4. 非僅由單一全投影片 raster image 構成（以圖片物件數量限制做基本防呆）

## 可編輯性設計

- 主要標題、副標、badge、區塊文字皆為 editable text boxes 或 shape text frame。
- Containers/cards/rounded rectangles/roadmap blocks/connector arrows 皆為 PowerPoint shapes。
- 右側重點區 icon 使用 vector-like editable shape 近似（可直接替換）。
- 內容模型、樣式常數、shape rendering、export pipeline 分離：
  - `pptx_redraw/models.py`
  - `pptx_redraw/theme.py`
  - `pptx_redraw/renderer.py`
  - `pptx_redraw/pipeline.py`

## 已知限制

- 因來源為 raster 參考圖，個別 icon 採用風格一致的可編輯近似圖示，非逐像素還原。
- 字型會依開啟環境可用字型而有些微差異。

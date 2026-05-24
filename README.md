# 活動支出追蹤系統

一個 Mobile-first 的活動支出及收據管理網站，可部署到 GitHub Pages / Vercel。

## 功能

- 建立活動檔案
- 記錄活動預算、已使用金額、剩餘金額
- 批量輸入支出紀錄
- 支援雙引號批量輸入，例如 `"大型活動佈置材料"`
- 支出搜尋、分類篩選、日期排序
- 匯出 Excel
- 匯出 PDF
- 密碼保護高風險操作
- localStorage 本機儲存
- JSON 備份及還原
- Mobile-first 手機底部導航

## 安裝方法

```bash
npm install
npm run dev
```

瀏覽器開啟：

```bash
http://localhost:5173
```

## 建置

```bash
npm run build
```

建置後會產生 `dist` 資料夾。

## 部署到 Vercel

1. 將所有檔案上載到 GitHub repo
2. 到 Vercel 新增 Project
3. 選擇你的 GitHub repo
4. Framework 選 Vite
5. Build Command 使用：
   ```bash
   npm run build
   ```
6. Output Directory 使用：
   ```bash
   dist
   ```

## 部署到 GitHub Pages

可使用 Vercel 較簡單。如要用 GitHub Pages，建議安裝 gh-pages 或使用 GitHub Actions。

## 注意

資料儲存在瀏覽器 localStorage，不同裝置不會自動同步。請定期使用「一鍵備份 JSON」。

# 模擬版

`index.html` 是打包好的模擬版，直接用瀏覽器打開就能玩，不會連到任何 Google 服務。

- `mocks.js`：用純 JavaScript 假造 Google 日曆、試算表、寄信等服務
- `wrapper.html`：外框（分頁切換、示範資料）
- `build.js`：把 `../apps-script/` 的程式和這裡的檔案打包成 `index.html`

改完 `apps-script/` 之後重新打包：

```bash
node build.js
```

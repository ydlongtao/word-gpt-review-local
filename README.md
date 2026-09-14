# Word GPT 本地审阅

在 Word 侧边栏中让 GPT 润色选区或正文，预览逐字差异，再将真正改变的片段写成 Word 原生修订。未修改的文字保持原样，可以在 Word 中逐条接受或拒绝修改。

## 在这台 Mac 上使用

1. 使用 Node.js 22.9 或更新版本，运行 `npm install`。
2. 双击 `安装Word插件.command`，或运行 `npm run setup`。它会生成并信任 Microsoft Office 本地开发证书、创建 `.env`、将插件清单放入 Word 的 `wef` 文件夹。macOS 可能要求系统授权。
3. 编辑项目里的 `.env`，填写模型配置。不要把真实密钥提交到 Git，也不必发送到聊天中。
4. 运行 `npm start`，或双击 `启动审阅服务.command`。使用期间保持终端运行。
5. 在 Word 打开文档，进入「主页 → 加载项」，选择 **GPT 本地审阅**。如未出现，请先保存文档，再退出并重新打开 Word。
6. 选择要修改的文字，填写要求，点击「生成修改建议」。也可以选择「文档正文」。
7. 查看绿色新增、红色删除和普通色未改动的预览，然后点击「以修订方式写入 Word」。
8. 在「审阅 → 所有标记」查看，在 Word 中逐条接受或拒绝。显示颜色由 Word 的审阅设置决定，不一定与插件预览相同。

浏览器可打开 https://localhost:3443 查看示例；实际读取和写入文档必须在 Word 内执行。

## 模型配置

云端 OpenAI 兼容接口：

```dotenv
MODEL_BASE_URL=https://api.openai.com/v1
MODEL_NAME=填写你的账户实际可用的模型名
MODEL_API_KEY=填写你的API密钥
PORT=3443
```

本地 Ollama 的 OpenAI 兼容接口：

```dotenv
MODEL_BASE_URL=http://localhost:11434/v1
MODEL_NAME=填写已下载的模型名称
MODEL_API_KEY=ollama
PORT=3443
```

需要模型支持 Chat Completions 和按指令输出 JSON。插件没有内置密钥，也不会自动读取其他应用的登录凭据。ChatGPT 的网页登录状态不能直接替代模型 API 配置。修改配置后重启服务。默认端口为 3443；如需更改，必须同步修改 manifest.xml 中的全部地址，再运行 setup。

## 处理范围与保护

- 支持普通正文、标题、列表的文本润色。每次最多 80 段、12000 字符。正文范围不包含页眉、页脚、脚注和文本框。
- 段落结构保持不变；不支持让模型合并段落、改变表格结构或添加图片。
- 含已有修订、批注、超链接、域、内容控件、嵌入对象的范围会被拒绝。段内手动换行和特殊分隔符也暂不处理。表格等复杂排版尚未在真实 Word 中验证，请优先使用普通正文。
- 未改动片段不被替换，因此原格式保持不动。新增片段的格式由 Word 在插入位置继承；不能承诺复杂混合字体的完全保真。
- 每次建议绑定生成时的 Word 范围。写入前检查文本和复杂内容；原文发生变化就停止。生成和写入期间请避免同时编辑该区域，当前不支持多人协同写入的事务隔离。
- 写入期间开启 TrackAll，完成后恢复之前的修订模式。修订作者是当前 Word 用户名，不能通过此 API 固定为“GPT”。本插件不会自动接受或拒绝你原有的修订。
- Word API 不提供此批次的原子事务；宿主中途失败时可能保留部分修改。页面会明确提示检查审阅窗格或撤销，并禁用直接重试。

## 本地的含义

侧边栏与代理服务在本机运行；服务仅监听 127.0.0.1。API Key 只由本地服务读取，网页不会收到它。模型请求不写入日志；`.env` 和依赖目录被 Git 忽略。

选择云端模型时，选定文字和修改要求会发送到配置的供应商。选择本地模型时，推理留在本机，但 Office.js 仍从微软 CDN 加载，因此这不是完整离线插件。

敏感接口检查 localhost Host、同源 Origin 和进程随机令牌。HTTPS 不代替模型供应商自身的数据处理约定。

## Windows 和网页版

插件清单使用跨平台 Office.js，要求 WordApi 1.4。Mac 安装脚本只配置 Mac，Windows / 网页版需按微软说明旁加载清单；尚未在这些平台实测：

- [Mac 加载说明](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)
- [Windows 加载说明](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- [网页版加载说明](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)

## 验证与已知限制

`npm test` 的 20 项测试检查字符差异、Unicode 定位、重复文本、返回结构、请求校验和错误处理，并用模拟宿主检查修订开关恢复与过期原文保护。`npm run validate` 检查 Office 插件清单。`node scripts/check-ui.mjs` 在本地 Chrome 检查页面与预览，并截图到 `artifacts/preview.png`。

自动测试不能证明真实 Word 宿主的范围定位和修订行为，也没有在缺少用户模型配置时调用实际 GPT。首次使用请用一份测试文档，检查：普通文字、混合加粗、中英文、重复词、emoji、接受全部得到新文、拒绝全部恢复原文，以及生成期间修改原文后应拒绝写入。

## 文件

- `public/word.mjs`：读取和绑定段落、定位差异、写入原生修订。
- `public/engine.mjs`：差异计算、结果校验与字符映射。
- `model.mjs` / `server.mjs`：模型调用和本地 HTTPS 服务。
- `manifest.xml`：Word 插件清单。
- `scripts/setup.mjs`：证书与 Mac 旁加载配置。

## 卸载

停止服务，删除 `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/word-gpt-review.xml`，保存并重启 Word。开发证书可能被其他 Office 开发项目共享；仅在不再需要时运行 `npx office-addin-dev-certs uninstall`。

## 实现依据

参考产品明确使用 Word Track changes：[GPT for Word 功能说明](https://gptforwork.com/docs/gpt-for-word)。本实现先计算差异，再利用 [Word 搜索范围](https://learn.microsoft.com/zh-cn/office/dev/add-ins/word/search-option-guidance) 定位片段，配合 [ChangeTrackingMode](https://learn.microsoft.com/en-us/javascript/api/word/word.changetrackingmode?view=word-js-1.4) 写入修订。

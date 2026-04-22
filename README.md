# Office Agent Local (Word)

本项目是一个本地运行的 Word 边栏 Agent 起步实现，目标是贴近 Copilot 体验：
- 插件在 Word 任务窗格内使用
- Agent 在用户机器本地运行
- 模型调用直连用户配置的 OpenAI 兼容接口
- 知识库上传、索引和检索默认本地存储

## 目录
- apps/word-addin: Word 插件前端（Task Pane）
- apps/local-agent: 本地 agent 服务（API + 检索 + 模型调用）

## 已实现能力（起步版）
- 本地 agent 健康检查: GET /health
- 模型配置保存: POST /v1/provider/config
- 知识库文件上传: POST /v1/kb/upload
- 对话接口: POST /v1/chat
- Word 侧边栏 UI:
  - 配置 base_url / api_key / model
  - 上传知识库文件
  - 对话并展示回复
  - 将最后回复插入当前光标位置
  - 通过同源 `/api` 代理访问本地 agent，降低桌面 WebView 网络拦截概率

## 快速开始
1. 安装依赖
```bash
npm install
```

2. 启动本地 agent
```bash
npm run dev:agent
```
默认监听 http://127.0.0.1:8787

3. 启动插件前端
```bash
npm run dev:addin
```
默认地址 https://localhost:3001
首次运行如果浏览器或 Word 提示证书不受信任，需要先安装并信任本机开发证书。

如果你已经看到“加载项不可用”，通常按下面顺序修复：
- 关闭 Word
- 运行 `npx office-addin-dev-certs install` 并完成证书信任
- 重新运行 `npm run dev:addin`
- 再次侧载 `apps/word-addin/manifest.xml`
- 重新打开 Word

4. 在 Word 侧载加载项
- 使用 Office 加载项侧载方式加载 manifest 文件。
- 在 Microsoft 365 桌面版 Word 中，优先使用“我的加载项 / 上传我的加载项”导入：
  - 选择 `apps/word-addin/manifest.xml`
  - 按提示完成加载
- 如果你的 Word 版本支持开发者侧载，也可以通过开发者工具直接加载 manifest。
- 加载完成后，在 Word 功能区点击“打开 Agent”即可显示边栏。
- 如果看不到按钮，先确认本地前端已启动在 `https://localhost:3001`，本地 agent 已启动在 `http://127.0.0.1:8787`。

## 网络错误排障（针对“无法加载此加载项”）
1. 确认前端页面可打开
- 在浏览器访问 `https://localhost:3001`，如果打不开，先重启 `npm run dev:addin`。

2. 确认本地 agent 可访问
- 在终端访问 `http://127.0.0.1:8787/health`，应返回 `ok: true`。

3. 确认证书已信任
- 运行 `npx office-addin-dev-certs install`。
- 完成后重启 Word 和前端开发服务。

4. 重新侧载清单
- 关闭 Word。
- 重新侧载 `apps/word-addin/manifest.xml`。
- 重新打开 Word 测试。

## 使用流程
1. 在边栏“模型配置”里填写：
- Base URL（例如 https://api.openai.com/v1）
- API Key
- Model

2. 上传本地文件构建知识库
- 当前起步版按文本读取并分块索引

3. 输入指令并发送
- 会自动附带当前文档和选区文本（有长度限制）

4. 点击“插入最后回复到光标”
- 将助手回复写回 Word 当前选区

## 高可用性下一步（建议）
- provider key 改为系统安全存储（Windows Credential Manager）
- 知识库改成真正向量检索（embedding + 向量库）
- 增加编辑计划 JSON + diff 预览 + 可撤销操作历史
- 支持断路器、超时、重试和熔断日志
- 增加离线健康自检与版本兼容检查

## 注意事项
- 当前是起步实现，知识库上传默认按纯文本处理。
- Office Add-in 在不同 Office 版本上有兼容差异，建议先在 Microsoft 365 桌面版验证。
- 出于本地化目标，项目默认不经过云端中转服务。

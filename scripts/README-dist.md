# Office Agent Local — Word 智能文档助手

## 快速开始

### 1. 安装
双击 `install.bat`，按提示完成安装。

### 2. 启动
双击 `start.bat` 启动本地 Agent 服务。

### 3. 使用
打开 Microsoft Word，在「开始」选项卡中点击 **「打开助手」** 按钮，侧边栏会打开。

首次使用需在侧边栏「设置」中配置：
- **API 地址**：你的 OpenAI 兼容接口地址（如 `https://api.openai.com/v1`）
- **API Key**：你的密钥
- **模型**：选择或输入模型名称

## 系统要求

- Windows 10/11
- Microsoft Word 2016 或更高版本（Microsoft 365 最佳）
- Node.js 18+（如未安装，请从 https://nodejs.org 下载）

## 功能说明

| 功能 | 说明 |
|------|------|
| 💬 对话 | 在侧边栏与 AI 对话，询问文档相关问题 |
| 📝 文档操作 | AI 可读取、插入、替换、删除文档内容 |
| 🎨 格式控制 | 支持段落样式、字体、加粗、斜体等格式操作 |
| 📚 知识库 | 上传文件作为知识库，AI 可检索引用 |
| 🔄 多轮迭代 | AI 可自主执行多步操作，支持感知→操作循环 |

## 卸载

双击 `uninstall.bat` 即可卸载。

## 目录结构

```
├── manifest.xml        ← Word Add-in 清单
├── install.bat         ← 安装脚本
├── start.bat           ← 启动脚本
├── uninstall.bat       ← 卸载脚本
└── server/
    ├── server.js       ← 主程序
    ├── package.json
    ├── node_modules/   ← 运行时依赖
    ├── public/         ← 前端界面
    └── data/           ← 运行时数据（配置、知识库）
```

## 隐私说明

- 所有数据存储在本地，不会上传到任何第三方服务器
- API 调用直接从你的电脑发往你配置的模型服务
- 知识库文件仅保存在本地 `server/data/` 目录

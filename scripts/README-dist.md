# Office Agent Local — Word 智能文档助手（dist 版）

## 快速开始

### 1) 安装
双击 `install.bat`。

安装脚本会尝试自动注册 `manifest.xml` 到 Word。  
如果自动导入失败或 Word 中看不到加载项，请手动操作：

1. 打开 Word：`文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 受信任的加载项目录`
2. 添加**当前 dist 目录**（即包含 `manifest.xml` 的目录）
3. 重启 Word

### 2) 启动
双击 `start.bat` 启动本地 Agent 服务（默认 `http://127.0.0.1:8787`）。

### 3) 使用
打开 Microsoft Word，在「开始」选项卡点击 **「Word Agent / 打开助手」**。

首次使用请在侧边栏配置：
- **API 地址**（如 `https://api.openai.com/v1`）
- **API Key**
- **模型**

## 主要能力
- 本地对话与知识库检索
- 文档读写与格式化工具调用
- 长程 Autopilot 多轮执行（可配置最大迭代次数）
- 操作计划预览 + 人工确认执行

## 系统要求
- Windows 10/11
- Microsoft Word 2016+（推荐 Microsoft 365 桌面版）
- Node.js 18+

## 卸载
双击 `uninstall.bat`。

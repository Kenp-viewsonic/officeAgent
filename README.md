# Office Agent Local (Word)

本项目是本地运行的 Word 任务窗格 Agent，贴近 Copilot 使用体验：
- 插件在 Word 内运行
- Agent 与知识库都在本地机器
- 直连用户配置的 OpenAI 兼容模型服务

## 目录
- `apps/word-addin`：Word 插件前端（Task Pane）
- `apps/local-agent`：本地 Agent 服务（API / 推理 / 检索）
- `scripts`：打包、安装、启动脚本

## 已实现能力（当前版本）
- 模型配置与预设管理（保存/加载/删除）
- 知识库上传、列表、清空、导入导出
- 聊天会话管理（新建/切换/删除/重试）
- 文档工具链（读取、插入、替换、删除、格式操作）
- **长程 Autopilot（Agent Loop）**
  - `smart_action` 模式下自动进入多轮迭代执行
  - 支持 `maxIterations` 控制最大迭代次数（默认 10）
  - 通过 `sessionId + /v1/chat/agent-continue` 持续推进任务
  - 每轮基于工具执行结果判断是否继续，完成后 `task_complete` 收束
  - 遇到需要确认的操作计划时，前端展示计划预览并支持“确认/取消”

## 快速开始
从release下载最新版本，解压后双击 `dist/install.bat` 注册 Manifest，然后双击 `dist/start.bat` 启动本地服务，打开 Word，在“开始”选项卡点击 **Word Agent / 打开助手**。
部分版本office**无法全自动注册**，进行如下操作：
- 打开 `dist`文件夹， `右键 -> 属性 -> 共享(若无共享选项则先至网络与共享中心，启用文件共享) -> 高级共享 -> 共享此文件夹`复制弹窗中显示的网络路径
- 打开 Word：`文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 受信任的加载项目录`
- 手动添加 包含 `manifest.xml` 的 `dist` 目录的 **网络路径**
- 关闭并重启 Word 
- 开始 -> 加载项 -> 更多加载项 -> 共享文件夹 -> Word Agent
在菜单中填入API配置即可使用
## 开发模式（源码运行）
1. 安装依赖
```bash
npm install
```
2. 启动本地 Agent
```bash
npm run dev:agent
```
默认地址：`http://127.0.0.1:8787`

3. 启动插件前端
```bash
npm run dev:addin
```
默认地址：`https://localhost:3001`

4. 在 Word 侧载 `apps/word-addin/manifest.xml`

## 生产分发模式（dist 打包产物，推荐给最终用户）
1. 生成打包产物
```bash
npm run package
```
输出目录：`dist/`

2. 在 `dist` 中执行安装与启动
- 双击 `dist/install.bat`（注册 Manifest）
- 双击 `dist/start.bat`（启动本地服务）
- 打开 Word，在“开始”选项卡点击 **Word Agent / 打开助手**

3. 如果 `install.bat` 没有顺利导入到 Word（未看到加载项）
- 打开 Word：`文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 受信任的加载项目录`
- 手动添加 **包含 `manifest.xml` 的 `dist` 目录**
- 关闭并重启 Word 后重试

## 说明
- 当前建议在 Microsoft 365 桌面版 Word 使用。
- 所有数据默认本地存储，不经过云端中转。

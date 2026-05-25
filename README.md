# AI 文本转 Word

一个基于 WinUI 3 / Windows App SDK 的 Windows 桌面工具，用来把从 ChatGPT、Claude、Copilot 等 AI 对话中复制出来的 Markdown 文本整理成干净的 Word 文档。

## 特性

- 粘贴即整理，适合 AI 对话、提示词、教程、方案文档等内容。
- 支持常见 Markdown 结构：标题、段落、加粗、斜体、行内代码、列表、引用、分割线、代码块。
- 导出 `.docx` Word 文档，并提供字体、字号、行距、页边距、引用块和列表密度设置。
- 自动读取 Windows 系统已安装字体，字体框支持搜索。
- 自动保存导出设置，下次打开继续使用。
- 提供示例文本，方便第一次打开时快速查看效果。
- 本地处理，不上传文本内容。

## 截图

项目还在早期阶段，界面会继续打磨。当前版本主界面采用 Fluent 风格、Mica 背景和简化标题栏。

## 安装

当前推荐从 GitHub Releases 下载最新的 `.msix` 安装包。

如果是本地构建的测试包，需要信任本地测试证书后再安装。

## 开发环境

- Windows 10/11
- Visual Studio，安装 Windows App SDK / WinUI 3 相关工作负载
- .NET SDK

运行测试：

```powershell
dotnet test AiTextToWord.slnx
```

构建：

```powershell
dotnet build AiTextToWord.slnx
```

生成 x64 MSIX 包需要使用 Visual Studio / MSBuild 的 Windows App SDK 打包工具链。

## 路线图

- PDF 导出
- 更接近 Word 页面效果的预览
- 表格和任务列表支持
- 代码块高亮
- 更多文档模板
- 更完善的 GitHub Actions 打包流程

## 许可证

MIT

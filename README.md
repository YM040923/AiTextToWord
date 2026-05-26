# AI 文本转 Word

把从 ChatGPT、Claude、Copilot 等 AI 对话里复制出来的 Markdown 文本，整理成干净、可交付的 Word 或 PDF 文档。

这是一个基于 **WinUI 3 / Windows App SDK** 的 Windows 桌面应用，界面遵循 Fluent Design，文本转换在本地完成，默认不会上传你的文档内容。

## 功能特性

- 粘贴即整理：适合 AI 对话、提示词、教程、方案、学习笔记等内容。
- Markdown 结构识别：支持标题、段落、加粗、斜体、行内代码、代码块、引用、分割线、无序列表和有序列表。
- Word / PDF 导出：支持 `.docx` 和 `.pdf`。
- 简单好用的导出设置：字体、正文字号、标题字号、行距、页边距、引用样式、列表密度。
- 系统字体读取：自动读取 Windows 已安装字体，并支持字体搜索。
- 设置页：把导出设置、AI 设置助手和基础偏好集中到独立页面。
- 导出后操作：导出完成后可直接打开文件或打开所在文件夹。
- 本地优先：核心文本整理和导出流程在本机完成。

## 安装

推荐从 [GitHub Releases](https://github.com/YM040923/AiTextToWord/releases/latest) 下载最新版。

当前发布包使用本地测试证书签名。首次安装时，请下载同一个 Release 里的三个文件，并放在同一文件夹：

- `AiTextToWord.App_*_x64.msix`
- `AiTextToWord.App_*_x64.cer`
- `Install-AiTextToWord.ps1`

然后在该文件夹打开 PowerShell，运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-AiTextToWord.ps1
```

脚本会自动请求管理员权限，把测试证书加入受信任证书存储，然后安装并启动应用。

更多安装说明见 [docs/INSTALL.md](docs/INSTALL.md)。

## 使用

1. 打开应用。
2. 把 AI 回复、Markdown 笔记或提示词内容粘贴到左侧输入区。
3. 在右侧确认整理预览。
4. 进入“设置”页面调整字体、字号、行距、边距等格式。
5. 点击“导出 Word”或选择 PDF 导出。

## 开发

环境要求：

- Windows 10/11
- Visual Studio 2026 或兼容版本
- Windows App SDK / WinUI 3 相关工作负载
- .NET SDK

恢复依赖：

```powershell
dotnet restore AiTextToWord.slnx
```

运行测试：

```powershell
dotnet test AiTextToWord.slnx
```

构建：

```powershell
dotnet build AiTextToWord.slnx --configuration Release
```

生成 MSIX 需要使用 Visual Studio / MSBuild 的 Windows App SDK 打包工具链。

## 路线图

见 [docs/ROADMAP.md](docs/ROADMAP.md)。

## 反馈与贡献

欢迎通过 GitHub Issues 反馈问题或建议。提交 Issue 前，建议先看一下是否已有相同问题。

- Bug 反馈：请提供 Windows 版本、应用版本、复现步骤和原始文本片段。
- 功能建议：请描述使用场景，而不仅是控件名称。

## 许可证

[MIT](LICENSE)

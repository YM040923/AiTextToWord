# 贡献指南

感谢你愿意改进 AI 文本转 Word。这个项目还在早期阶段，最需要的是稳定性反馈、真实文本样本和清晰的复现步骤。

## 提交 Issue

Bug 反馈请尽量包含：

- 应用版本。
- Windows 版本。
- 复现步骤。
- 预期结果和实际结果。
- 可公开的原始文本片段。
- 如果是导出问题，请说明 Word、PDF 还是两者都有问题。

功能建议请描述你的使用场景，例如“我经常把课程笔记导出给学生，需要固定标题层级和页边距”，这比只说“增加一个按钮”更容易转成可实现设计。

## 本地开发

```powershell
dotnet restore AiTextToWord.slnx
dotnet test AiTextToWord.slnx
dotnet build AiTextToWord.slnx --configuration Release
```

## 代码风格

- 保持功能简单，不为少数场景提前堆复杂配置。
- 优先增加核心转换和导出逻辑的测试。
- 不把用户文本上传到外部服务，除非用户明确配置并触发相关功能。
- 修改 UI 时保持简体中文文案。

## Pull Request

提交 PR 前请确认：

- 测试通过。
- README 或文档已同步更新。
- 没有提交本地生成的安装包、临时文件或个人配置。

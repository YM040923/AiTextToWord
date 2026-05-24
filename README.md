# AI Text to Word

AI Text to Word is a Windows desktop app for turning text copied from AI chat tools into a clean Word document.

## Features

- Paste AI chat text into a WinUI 3 desktop app.
- Clean common copy-instruction wrappers.
- Preview headings, paragraphs, lists, quotes, dividers, and code blocks.
- Export a readable `.docx` file.

## First Release Scope

Supported Markdown-like structures:

- Headings `#`, `##`, `###`
- Paragraphs
- Unordered and ordered lists
- Blockquotes
- Fenced code blocks
- Horizontal dividers

Not supported in the first release:

- PDF export
- Rich text editing
- Templates
- Tables
- Images
- Footnotes

## Build

Requirements:

- Windows
- Visual Studio with Windows App SDK / WinUI 3 workload
- .NET SDK compatible with the installed Windows App SDK template

Run tests:

```powershell
dotnet test AiTextToWord.slnx
```

Build:

```powershell
dotnet build AiTextToWord.slnx
```

Build from Visual Studio if the WinUI project requires Visual Studio tooling on your machine.

## License

MIT

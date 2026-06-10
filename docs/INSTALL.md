# 安装说明

## 推荐方式

从 GitHub Releases 下载最新版：

https://github.com/YM040923/AiTextToWord/releases/latest

同一个 Release 里通常会提供：

- `AiTextToWord.App_*_x64.msix`
- `AiTextToWord.App_*_x64.cer`
- `Install-AiTextToWord.ps1`

把三个文件放在同一个文件夹后，在该文件夹打开 PowerShell，运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-AiTextToWord.ps1
```

安装脚本会做三件事：

1. 请求管理员权限。
2. 把 `.cer` 测试证书导入 `LocalMachine\Root` 和 `LocalMachine\TrustedPeople`。
3. 安装 `.msix` 并启动应用。

## 为什么需要证书

当前项目还处于个人开源早期阶段，安装包使用本地测试证书签名。Windows 安装 MSIX 时必须验证签名证书，所以首次安装需要信任随包提供的 `.cer` 文件。

未来如果项目使用正式代码签名证书或上架 Microsoft Store，这一步就可以取消。

## 手动安装

如果不使用脚本，可以用管理员 PowerShell 运行：

```powershell
Import-Certificate -FilePath ".\AiTextToWord.App_1.0.0.18_x64.cer" -CertStoreLocation Cert:\LocalMachine\Root
Import-Certificate -FilePath ".\AiTextToWord.App_1.0.0.18_x64.cer" -CertStoreLocation Cert:\LocalMachine\TrustedPeople
Add-AppxPackage -Path ".\AiTextToWord.App_1.0.0.18_x64.msix" -ForceUpdateFromAnyVersion
```

如果文件名版本不同，请替换为你下载到的实际文件名。

## 卸载

可以在 Windows 设置中卸载：

`设置` -> `应用` -> `已安装的应用` -> `AI 文本转 Word`

也可以用 PowerShell：

```powershell
Get-AppxPackage -Name YM040923.AiTextToWord | Remove-AppxPackage
```

## 常见问题

### 提示 0x800B0109 或证书不受信任

说明 `.cer` 没有加入受信任根证书。请用管理员 PowerShell 重新运行安装脚本，或手动导入证书。

### 双击 MSIX 无法安装

请先运行安装脚本。直接双击 MSIX 时，Windows 不一定会自动信任本地测试证书。

### 安装后开始菜单找不到

可以尝试在开始菜单搜索 `AI 文本转 Word`。如果仍找不到，重新运行：

```powershell
Add-AppxPackage -Path ".\AiTextToWord.App_1.0.0.18_x64.msix" -ForceUpdateFromAnyVersion
```

### Windows 阻止运行脚本

在脚本所在目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-AiTextToWord.ps1
```

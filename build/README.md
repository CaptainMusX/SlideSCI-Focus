# 构建 SlideSCI Focus VSTO 安装包

在仓库根目录运行：

```powershell
pwsh -File .\build\Build-Installer.ps1
```

脚本会自动定位带 OfficeTools targets 的 Visual Studio，安装并锁定 LaTeX 的 Node 依赖，清理构建输出，然后生成 ClickOnce 兼容发布目录：

- `artifacts\SlideSCI-Focus-<版本>\setup.exe`
- `artifacts\SlideSCI-Focus-<版本>\CaptainMusX.SlideSCI.Focus.vsto`
- `artifacts\SlideSCI-Focus-<版本>\Application Files\`
- `artifacts\SlideSCI-Focus-<版本>\SHA256SUMS.txt`

默认不把 MathJax 的 `node_modules` 放进发布目录，以保持与原作者 Release 一致；如需离线 LaTeX SVG 运行时，可显式运行 `-BundleLatexRuntime`。

`Build-Installer.ps1` 会通过 MSBuild 的 `RestorePackagesConfig=true` 自动还原 `packages.config` 中的 NuGet 依赖；本地 `packages` 缓存不需要提交到仓库。

本地没有配置正式代码签名证书时，脚本会在当前用户的 `%LOCALAPPDATA%\SlideSCI-Focus\build-signing` 中生成开发用自签名证书。它适合本机测试，正式分发前应替换为受信任的发布证书，并保留同一签名身份用于后续更新。

## Inno Setup 安装包

在已生成 Release 输出后运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build\Build-InnoInstaller.ps1
```

该安装器参考原作者 Advanced Installer 的直接布局：把 `bin\Release` 中的 `CaptainMusX.SlideSCI.Focus.vsto`、`CaptainMusX.SlideSCI.Focus.dll.manifest`、程序集和 `latex-converter` 文件放在用户选择的目录，排除 `app.publish`、`Application Files` 和 `node_modules`，并按应用清单中的 `keyName="CaptainMusX.SlideSCI.Focus"` 注册加载项。构建时会从已签名应用清单提取公钥，安装时只为当前清单 URL 与该公钥写入用户级 VSTO Inclusion 信任，卸载时同步移除。安装前只清理本 fork 旧的 VSTO URL、ClickOnce 元数据和注册项，不会删除原作者的 `SlideSCI` 加载项；不会运行旧的 ClickOnce `setup.exe`，因此不会复用已删除的 D 盘清单。卸载入口由 Inno Setup 提供。

安装前清理已安装版本（脚本会先备份注册项并保护仓库目录）：

```powershell
powershell -ExecutionPolicy Bypass -File .\build\Clean-SlideSCI-Focus-Install.ps1
```

如果 Office 仍指向已删除的旧安装目录，请关闭 PowerPoint，运行清理脚本后重新执行最新 Inno 安装包。不要再使用 VSTOInstaller 安装 `.vsto` 文件，否则会重新进入 ClickOnce 部署路径。

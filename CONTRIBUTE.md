## 如何开发一个PPT插件

我是如何开发PPT插件的分享：如何开发一个PPT插件：使用VSTO开发

## 插件更新版本

每一轮实际改进都必须完成一次可追溯发布闭环，不能只保留本地代码：

1. 更新 `CHANGELOG.md`，并递增版本；同步 `AssemblyInfo.cs`、`SlideSCI.csproj` 的 `ApplicationVersion` 和 Inno 配置。
2. 运行 Release 构建、核心回归和 Inno 安装包构建：

   ```powershell
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\build\Build-Installer.ps1 -Configuration Release
   powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-CoreRegression.ps1
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\build\Build-InnoInstaller.ps1
   ```

3. 检查新安装包的绝对路径、版本、SHA-256 和签名性质。
4. 选择性暂存本轮文件，提交后将 commit 推送到 GitHub（默认先核对 `focus` 远程和当前分支）。
5. 交付时报告 GitHub commit/推送结果、新安装包路径和验证结果；真实 PowerPoint 安装、冷启动和人工界面验收单独说明。

完整规则见根目录 [`AGENTS.md`](AGENTS.md)，构建命令和安装器细节见 [`build/README.md`](build/README.md)。

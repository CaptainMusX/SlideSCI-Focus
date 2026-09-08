# SlideSCI Focus agent development rules

本文件是本仓库中代理执行开发、修复和优化任务时的必读规则。

## 每轮改进的交付门槛

凡是本轮修改了受 Git 跟踪的源代码、项目配置、测试、文档、构建脚本或发布元数据，都算一轮改进。每轮改进在交付前必须完成以下闭环：

1. 先检查 `git status --short --branch`、当前分支和 GitHub 远程；保留用户已有的未提交修改，不使用 `git add -A` 把无关文件混入提交。
2. 递增发布版本，并保持 `SlideSCI/Properties/AssemblyInfo.cs`、`SlideSCI/SlideSCI.csproj` 的 `ApplicationVersion`、`CHANGELOG.md` 和 Inno 配置中的版本一致。
3. 完成与本轮改动相称的回归检查；至少运行 `tests/Run-CoreRegression.ps1`（它要求先有 Release DLL）。
4. 构建新的 Release 和安装包：

   ```powershell
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\build\Build-Installer.ps1 -Configuration Release
   pwsh -NoProfile -ExecutionPolicy Bypass -File .\build\Build-InnoInstaller.ps1
   ```

   安装包必须来自当前版本的 `SlideSCI/bin/Release`，不能使用旧的 `artifacts` 输出或 ClickOnce 中间目录。

5. 检查安装包确实存在，记录版本、绝对路径、文件大小、SHA-256 和签名性质；构建成功不等于已在真实 PowerPoint 文档中完成安装、冷启动或人工 UI 验收。
6. 只暂存本轮及用户明确纳入本轮的文件，提交可追溯的 commit，然后将该 commit 推送到 GitHub。当前仓库通常使用 `focus` 远程；每轮仍必须先核对远程和分支，不能凭记忆推送到错误仓库。
7. 最终交付必须同时报告：commit 和 GitHub 推送结果、安装包绝对路径与 SHA-256、运行过的检查，以及尚未完成的真实 Office/设备验收。若推送或打包失败，不能把本轮标记为完成，应继续排查或明确报告阻塞原因。

只有纯只读审查、没有文件改动的回合不需要创建新版本或推送。任何实际改进回合都不能只留下本地修改或只给出构建命令而不生成新的安装包。

## 发布安全

- 使用选择性 `git add <path>`；先查看 staged diff，再提交。
- 不提交私钥、用户配置、临时文件、`packages` 缓存或整个 `artifacts` 输出目录。
- 本地开发证书可以用于本机验证，但交付时必须说明它不是公共信任链证书。
- 构建/回归测试、安装包检查、真实 PowerPoint COM 行为、安装卸载、冷启动和用户实际 UI 操作分别报告，不能相互替代。

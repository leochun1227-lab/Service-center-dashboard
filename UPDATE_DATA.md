# 更新和发布看板

双击 `RUN_UPDATE_DATA.bat`：拉取 C4C → 补充 SAP 发票数据 → 更新状态历史 → 生成看板 → 自动提交网站文件 → 推送 GitHub。

Render 服务的 Settings → Auto-Deploy 设为 **On Commit**，连接本仓库的 `main` 分支。以后每次推送都会触发部署，无需手动点击 Deploy。`render.yaml` 也已声明这个配置；如果服务不是通过 Blueprint 管理，请在 Render 设置一次。

数据更新成功、发布失败时，双击 `RUN_PUBLISH_DATA.bat` 重试发布，无需再次拉取数据。日志位于项目的 `outputs` 目录。发布成功代表已推送 GitHub，网站上线仍需等待 Render 构建完成。

## 换到另一台 Windows 电脑

1. 安装 Python 3.10 或更新版本，以及 Git 或 GitHub Desktop。
2. 在 GitHub Desktop 登录有仓库推送权限的账号，将仓库克隆到任意目录，切换到 `main`。配置 Git 的提交姓名和邮箱。
3. 安装与 Python 位数一致的 SAP HANA ODBC 驱动，确保新电脑可访问 C4C 和 SAP HANA 所在网络。
4. 双击 `RUN_UPDATE_DATA.bat`。首次运行自动创建项目内的 `.venv`，并按 `requirements.txt` 安装依赖，需要联网。
5. 首次运行按提示输入 C4C 账号密码、SAP HANA 连接串；状态历史如果使用独立账号，也在这次配置时输入。之后双击更新 BAT 会自动读取，不再逐次询问。

连接信息使用 Windows DPAPI 加密，保存在当前 Windows 用户的 `%LOCALAPPDATA%\ServiceCenterDashboard\connections.dpapi`，不会进入 Git 仓库或更新日志。同一 Windows 账号重新打开终端、重启电脑或把项目克隆到其他目录后仍可读取。环境变量可临时覆盖已保存的值，不会覆盖加密文件。

换电脑或换 Windows 账号需要各配置一次。只有账号密码或连接信息变更时，才手动运行 `SETUP_CONNECTIONS.bat` 修改；留空可保留原值。日常更新不会因为连接失败而反复弹出输入框。

GitHub 使用 Git Credential Manager 已保存的登录，日常发布已关闭交互式登录提示。每台电脑首次完成 Git 登录即可；凭据过期或权限被撤销时需重新登录。Render 自动部署不需要在每次更新时输入 Render 密码。

所有工作簿、网页、日志和备份都以 BAT 所在的项目目录为基准，不依赖固定用户名或盘符。新电脑不需要安装 Codex。`.venv` 不提交到 Git，也不要从旧电脑复制；若移动整个项目目录，请先删除复制过来的 `.venv`，下次运行会重新创建。

自动提交只包含 Render 使用的网站文件：`overview.html`、`dashboard-data.js`、两个 `assets` 文件和 `render.yaml`。Python 脚本等代码修改仍按正常开发流程提交。其他暂存的文件不会混入自动数据提交。

如果 GitHub 的 `main` 有新提交，发布会停止并保留本地数据。先在 GitHub Desktop 同步和处理冲突，再双击 `RUN_PUBLISH_DATA.bat`；脚本不会强制覆盖远程。

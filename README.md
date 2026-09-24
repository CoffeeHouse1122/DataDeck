# DataDeck

Windows 桌面应用，使用 Electron、Vue 3 和 TypeScript。

## 环境准备

使用 Windows x64、Node.js 22 和 Git，在项目根目录执行：

```powershell
npm ci
```

## 启动开发服务

```powershell
npm run dev
```

启动桌面窗口，支持界面热更新和主进程代码监视。终端按 `Ctrl+C` 停止开发服务。

## 构建与打包

| 命令 | 用途 | 输出目录 |
| --- | --- | --- |
| `npm run build` | 类型检查并编译 | `dist-electron/` |
| `npm run preview` | 运行已编译版本，需先 build | — |
| `npm run pack:win` | 构建并生成 Windows 免安装目录包 | `release/win-unpacked/` |
| `npm run dist:win` | 构建并生成 Windows NSIS 安装包 | `release/` |

本地打包命令不会发布。安装包名称为 `DataDeck-版本号-Setup.exe`，当前使用无签名发布。

## 发布到 GitHub

推送新的 `v*` 标签会触发 [Release DataDeck](https://github.com/CoffeeHouse1122/DataDeck/actions/workflows/release.yml) 工作流，自动检查、构建并发布 Windows x64 安装包及更新文件。

先提交本次要发布的代码并确认工作区干净。以下以 **0.1.5** 为例，后续发布需统一替换为新的版本号：

```powershell
npm version 0.1.5 --no-git-tag-version
git diff -- package.json package-lock.json
git status --short
git add -- package.json package-lock.json
git commit -m "发布 0.1.5 版本"
git tag -a v0.1.5 -m "发布 0.1.5 版本"
git push origin master
git push origin v0.1.5
```

按顺序执行，任一步失败应先处理再继续。标签必须与 `package.json` 版本一致，不覆盖已发布的标签。

工作流成功后，在 [Releases](https://github.com/CoffeeHouse1122/DataDeck/releases) 确认以下三个文件齐全：

- `DataDeck-版本号-Setup.exe`
- `DataDeck-版本号-Setup.exe.blockmap`
- `latest.yml`

未完成发布的已有标签可在 Actions → Release DataDeck → Run workflow 中填写 `release_tag` 重试。

## 针对性检查

```powershell
npm run test:updater     # 更新流程
npm run test:resources   # 打包资源隔离
npm run verify:release  # 检查 release/ 中完整的安装包、更新元数据和解包目录
```

依赖目录、构建产物、安装包、本地资料和真实环境配置不提交到 Git。

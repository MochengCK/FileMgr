# FileMgr

FileMgr 是一个基于 **Tauri 2 + Vue 3** 的桌面文件管理器，目标是在 Windows 平台上提供轻量、流畅又现代的文件浏览体验。

## 源码发布范围

本仓库以“**仅后端**”形式发布（Rust 后端 + Tauri 壳工程）。前端源码与 Node/Vite 工程文件为有意不包含。

## 功能特性

- 多标签页文件浏览，支持在同一窗口中打开多个路径  
- 类资源管理器的文件列表视图，支持排序、隐藏文件、系统文件显示等  
- 上下文右键菜单，与系统行为尽量保持一致  
- 侧边栏快速访问常用路径（此电脑、下载、图片等）  
- 内置设置页面，可调节：
  - 主题模式与强调色
  - 窗口特效
  - 列表行距、背景显示
  - 搜索行为（是否需要回车触发等）
- 使用 Rust 编写的后端，负责文件系统访问与性能敏感操作

> 具体界面与交互以实际版本为准。

## 运行环境

- 操作系统：Windows 10/11（桌面环境）  
- Rust：stable 工具链（用于构建 `rust-backend/` 与 `src-tauri/`）  
- Node.js：建议 18+（仅在你本地拥有前端源码时需要）  

## 后端开发（Rust）

构建：

```bash
cd rust-backend
cargo build
```

运行（如果后端工程提供可运行二进制）：

```bash
cd rust-backend
cargo run
```

测试：

```bash
cd rust-backend
cargo test
```

## 桌面应用（Tauri）开发与构建

桌面应用依赖前端资源。如果你本地拥有前端源码，可在仓库根目录执行：

```bash
npm install
npm run dev
```

构建发布版安装包：

```bash
npm run build
```

> 首次构建前，请确保已按 Tauri 官方文档准备好 Rust 工具链与 Windows 平台依赖。

## 目录结构概览

- `rust-backend/`：Rust 后端工程（文件操作等逻辑）  
- `src-tauri/`：Tauri 配置与壳工程  
- `LICENSE`：Apache License 2.0  

前端源码（如 `src/`、`package.json`、`vite.config.js`、`dist/`）不包含在本仓库中。

（实际内容可能会随版本演进略有变化）

## 许可证

本项目在 [LICENSE](./LICENSE) 中声明的 **Apache License 2.0** 下发布。


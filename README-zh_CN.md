# FileMgr

FileMgr 是一个基于 **Tauri 2 + Vue 3** 的桌面文件管理器，目标是在 Windows 平台上提供轻量、流畅又现代的文件浏览体验。

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
- Node.js：建议 18+  
- 包管理：npm（仓库内已包含 `package-lock.json`）  

## 前端开发命令

在项目根目录 `e:\FileMgr` 下执行：

```bash
npm install
```

### 仅运行前端（浏览器预览）

```bash
npm run web:dev
```

构建前端静态资源：

```bash
npm run web:build
```

预览构建结果：

```bash
npm run web:preview
```

## 桌面应用（Tauri）开发与构建

启动 Tauri 开发模式（会同时启动前端和桌面壳）：

```bash
npm run dev
```

构建发布版安装包：

```bash
npm run build
```

> 首次构建前，请确保已按 Tauri 官方文档准备好 Rust 工具链与 Windows 平台依赖。

## 目录结构概览

- `src/`  
  - `App.vue`：主界面与应用逻辑入口  
  - `main.js`：前端启动入口  
  - `app/`：应用内部模块（如上下文菜单、侧边栏等）  
- `dist/`：前端构建输出目录  
- `rust-backend/`：Rust 后端工程（文件操作等逻辑）  
- `src-tauri/`（如存在）：Tauri 配置与壳工程  
- `package.json`：前端依赖与 npm 脚本  
- `Cargo.toml`：Rust 后端依赖与构建配置  

（实际内容可能会随版本演进略有变化）

## 许可证

本项目在 [LICENSE](./LICENSE) 中声明的 **GNU General Public License v3.0 (GPL-3.0)** 下发布。  
你可以在遵守该许可证条款的前提下自由使用、修改和分发本项目。


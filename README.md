<div align="center">
  <table width="100%">
    <tr>
      <td align="right"><a href="./README-zh_CN.md">简体中文</a></td>
    </tr>
  </table>
</div>

# FileMgr

FileMgr is a desktop file manager built with **Tauri 2** and **Vue 3**, aiming to provide a lightweight, smooth, and modern file browsing experience on Windows.

## Features

- Multi‑tab file browsing, allowing multiple locations in one window  
- Explorer‑like file list view with sorting, hidden/system file toggle, and more  
- Context menus that closely mimic native behavior  
- Sidebar with quick access to common locations (This PC, Downloads, Pictures, etc.)  
- Built‑in Settings page for tuning:
  - Theme mode and accent color
  - Window visual effects
  - List row gap and background appearance
  - Search behavior (e.g., whether Enter is required)
- Rust backend for filesystem access and performance‑sensitive operations

> UI and interactions may evolve; refer to the running app for the most up‑to‑date behavior.

## Requirements

- OS: Windows 10/11 (desktop)  
- Node.js: 18+ recommended  
- Package manager: npm (a `package-lock.json` is included)  

## Frontend Commands

Run the following in the project root `e:\FileMgr`:

```bash
npm install
```

### Dev server (web only)

Start the Vite dev server:

```bash
npm run web:dev
```

Build static frontend assets:

```bash
npm run web:build
```

Preview the built bundle:

```bash
npm run web:preview
```

## Desktop App (Tauri) Dev & Build

Start Tauri dev mode (launches both frontend and desktop shell):

```bash
npm run dev
```

Create a production build/installer:

```bash
npm run build
```

> Before building, ensure that the Rust toolchain and Windows dependencies required by Tauri are properly installed, according to the official Tauri documentation.

## Project Structure (Overview)

- `src/`  
  - `App.vue`: Main UI and application logic  
  - `main.js`: Frontend entry point  
  - `app/`: Internal modules (context menu, sidebar navigation, etc.)  
- `dist/`: Built frontend assets  
- `rust-backend/`: Rust backend crate (filesystem and related logic)  
- `src-tauri/` (if present): Tauri configuration and shell project  
- `package.json`: Frontend dependencies and npm scripts  
- `Cargo.toml`: Rust backend dependencies and build configuration  

Actual layout may evolve as the project grows.

## License

This project is released under the **GNU General Public License v3.0 (GPL‑3.0)** as stated in [LICENSE](./LICENSE).  
You are free to use, modify, and redistribute it under the terms of that license.


use serde::{Deserialize, Serialize};
use std::{
    collections::{HashMap, HashSet, VecDeque},
    fs,
    io,
    path::PathBuf,
    sync::{Mutex, OnceLock},
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};
use tauri::{Emitter, Manager};
use walkdir::WalkDir;
#[cfg(target_os = "windows")]
use window_vibrancy::{apply_acrylic, apply_blur, apply_mica, clear_acrylic, clear_blur, clear_mica};
use std::sync::atomic::{AtomicU64, Ordering};
#[cfg(windows)]
mod win_app_id;

#[cfg(windows)]
static ICON_CACHE: OnceLock<Mutex<IconCache>> = OnceLock::new();

#[cfg(windows)]
static SEARCH_LATEST_REQUEST_ID: OnceLock<Mutex<HashMap<String, u64>>> = OnceLock::new();

#[cfg(windows)]
static THUMB_CACHE: OnceLock<Mutex<ThumbCache>> = OnceLock::new();

#[cfg(windows)]
static GALLERY_LATEST_REQUEST_ID: OnceLock<Mutex<HashMap<String, u64>>> = OnceLock::new();

#[cfg(windows)]
static NATIVE_MENU_IO_SEQ: AtomicU64 = AtomicU64::new(1);

static FOLDER_SIZE_CACHE: OnceLock<Mutex<HashMap<String, FolderSizeCacheEntry>>> = OnceLock::new();

struct FolderSizeCacheEntry {
    size: u64,
    dir_modified_ms: u128,
    at: Instant,
}

fn folder_size_cache() -> &'static Mutex<HashMap<String, FolderSizeCacheEntry>> {
    FOLDER_SIZE_CACHE.get_or_init(|| Mutex::new(HashMap::new()))
}

static FOLDER_SIZE_REQ_SEQ: AtomicU64 = AtomicU64::new(1);
static FOLDER_SIZE_LATEST_REQUEST_ID: OnceLock<Mutex<HashMap<String, u64>>> = OnceLock::new();

fn folder_size_latest_map() -> &'static Mutex<HashMap<String, u64>> {
    FOLDER_SIZE_LATEST_REQUEST_ID.get_or_init(|| Mutex::new(HashMap::new()))
}

static DIR_STATS_LATEST_REQUEST_ID: OnceLock<Mutex<HashMap<String, u64>>> = OnceLock::new();

fn dir_stats_latest_map() -> &'static Mutex<HashMap<String, u64>> {
    DIR_STATS_LATEST_REQUEST_ID.get_or_init(|| Mutex::new(HashMap::new()))
}

#[cfg(windows)]
use std::cell::RefCell;

#[cfg(windows)]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct IconCacheKey {
    path: String,
    size: u32,
}

#[cfg(windows)]
struct IconCache {
    cap: usize,
    map: HashMap<IconCacheKey, Option<String>>,
    order: VecDeque<IconCacheKey>,
}

#[cfg(windows)]
impl IconCache {
    fn new(cap: usize) -> Self {
        Self {
            cap: cap.max(64),
            map: HashMap::new(),
            order: VecDeque::new(),
        }
    }

    fn touch(&mut self, key: &IconCacheKey) {
        self.order.retain(|k| k != key);
        self.order.push_back(key.clone());
    }

    fn get(&mut self, key: &IconCacheKey) -> Option<Option<String>> {
        let v = self.map.get(key).cloned();
        if v.is_some() {
            self.touch(key);
        }
        v
    }

    fn insert(&mut self, key: IconCacheKey, value: Option<String>) {
        self.touch(&key);
        self.map.insert(key, value);
        while self.map.len() > self.cap {
            let k = match self.order.pop_front() {
                Some(v) => v,
                None => break,
            };
            self.map.remove(&k);
        }
    }
}

#[cfg(windows)]
fn icon_cache() -> &'static Mutex<IconCache> {
    ICON_CACHE.get_or_init(|| Mutex::new(IconCache::new(1500)))
}

#[cfg(windows)]
fn search_latest_map() -> &'static Mutex<HashMap<String, u64>> {
    SEARCH_LATEST_REQUEST_ID.get_or_init(|| Mutex::new(HashMap::new()))
}

#[cfg(windows)]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct ThumbCacheKey {
    path: String,
    size: u32,
}

#[cfg(windows)]
struct ThumbCache {
    cap: usize,
    map: HashMap<ThumbCacheKey, Option<String>>,
    order: VecDeque<ThumbCacheKey>,
}

#[cfg(windows)]
impl ThumbCache {
    fn new(cap: usize) -> Self {
        Self {
            cap: cap.max(64),
            map: HashMap::new(),
            order: VecDeque::new(),
        }
    }

    fn touch(&mut self, key: &ThumbCacheKey) {
        self.order.retain(|k| k != key);
        self.order.push_back(key.clone());
    }

    fn get(&mut self, key: &ThumbCacheKey) -> Option<Option<String>> {
        let v = self.map.get(key).cloned();
        if v.is_some() {
            self.touch(key);
        }
        v
    }

    fn insert(&mut self, key: ThumbCacheKey, value: Option<String>) {
        self.touch(&key);
        self.map.insert(key, value);
        while self.map.len() > self.cap {
            let k = match self.order.pop_front() {
                Some(v) => v,
                None => break,
            };
            self.map.remove(&k);
        }
    }
}

#[cfg(windows)]
fn thumb_cache() -> &'static Mutex<ThumbCache> {
    THUMB_CACHE.get_or_init(|| Mutex::new(ThumbCache::new(600)))
}

#[cfg(windows)]
fn gallery_latest_map() -> &'static Mutex<HashMap<String, u64>> {
    GALLERY_LATEST_REQUEST_ID.get_or_init(|| Mutex::new(HashMap::new()))
}


#[cfg(windows)]
use std::os::windows::fs::MetadataExt;

#[cfg(windows)]
use base64::Engine as _;

#[cfg(windows)]
use windows::Win32::{
    Foundation::{FILETIME, HWND, LPARAM, LRESULT, POINT, PROPERTYKEY, WPARAM},
    System::Com::{
        CoCreateInstance, CoInitializeEx, CoTaskMemAlloc, CoTaskMemFree, CoUninitialize, CLSCTX_INPROC_SERVER,
        COINIT_APARTMENTTHREADED,
    },
    System::Com::StructuredStorage::{PropVariantClear, PROPVARIANT},
    Graphics::Gdi::{
        CreateCompatibleDC, DeleteDC, DeleteObject, GetDIBits, GetObjectW, SelectObject, BITMAP,
        BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HBITMAP, HGDIOBJ,
    },
    Storage::FileSystem::{GetDiskFreeSpaceExW, GetDriveTypeW, GetVolumeInformationW, FILE_FLAGS_AND_ATTRIBUTES},
    Storage::FileSystem::FILE_ATTRIBUTE_NORMAL,
    UI::{
        Controls::{IImageList, TaskDialogIndirect, TASKDIALOGCONFIG, TASKDIALOG_BUTTON, TDF_ALLOW_DIALOG_CANCELLATION, TDF_USE_HICON_MAIN},
        Shell::{
            SHBindToParent, SHGetDesktopFolder, SHGetFileInfoW, SHGetImageList, SHParseDisplayName, StrRetToBufW, IContextMenu, IContextMenu2,
            IContextMenu3, IEnumIDList, IShellFolder, CMF_NORMAL, CMINVOKECOMMANDINFOEX, CMIC_MASK_PTINVOKE,
            SHCONTF_FOLDERS,
            SHCONTF_NONFOLDERS, SHFILEINFOW, SHGDN_FORPARSING, SHGDN_NORMAL, SHGFI_ICON, SHGFI_LARGEICON, SHGFI_PIDL,
            SHGFI_SMALLICON, SHGFI_ICONLOCATION, SHGFI_DISPLAYNAME, SHGFI_SYSICONINDEX, SHGFI_USEFILEATTRIBUTES, SHGFI_TYPENAME,
            SHCreateItemFromIDList, IShellItem2,
            Common::{ITEMIDLIST, STRRET},
            DestinationList, EnumerableObjectCollection, ShellLink, ICustomDestinationList, IShellLinkW,
            SHObjectProperties, SHOP_FILEPATH, ShellExecuteW, FO_DELETE, FOF_ALLOWUNDO, FOF_NOCONFIRMATION, FOF_SILENT, SHFileOperationW,
            SHFILEOPSTRUCTW, SHEmptyRecycleBinW, SHERB_NOCONFIRMATION, SHERB_NOPROGRESSUI, SHERB_NOSOUND,
        },
        Shell::Common::{IObjectArray, IObjectCollection},
        WindowsAndMessaging::{
            CallWindowProcW, CreatePopupMenu, DestroyIcon, DestroyMenu, GetCursorPos, GetIconInfo, GetMenuItemCount, GetMenuItemID,
            MessageBoxW, IDOK, MB_DEFBUTTON2, MB_ICONWARNING, MB_OKCANCEL,
            PostMessageW, SetForegroundWindow,
            SetWindowLongPtrW, TrackPopupMenuEx, GWLP_WNDPROC, HMENU, ICONINFO, TPM_RETURNCMD, TPM_RIGHTBUTTON, WM_DRAWITEM,
            WM_INITMENUPOPUP, WM_MEASUREITEM, WM_MENUCHAR, WM_NULL, HICON,
            SW_SHOW,
        },
    },
};
#[cfg(windows)]
use windows::core::{Interface, PCSTR, PCWSTR, PSTR, PWSTR};
#[cfg(windows)]
use windows::Win32::UI::Shell::PropertiesSystem::{IPropertyStore, PSGetPropertyKeyFromName};
#[cfg(windows)]
use windows::Win32::UI::Shell::SetCurrentProcessExplicitAppUserModelID;

#[derive(Debug, Deserialize)]
struct QuickAccessPathParams {
    path: String,
}

#[derive(Debug, Deserialize)]
struct ShowQuickAccessContextMenuParams {
    path: String,
}

#[derive(Debug, Serialize, Clone)]
struct QuickAccessEntry {
    path: String,
    label: String,
    pinned: bool,
}

#[derive(Debug, Deserialize)]
struct ListDirParams {
    path: String,
    #[serde(default)]
    show_hidden: bool,
    #[serde(default)]
    show_system: bool,
}

#[derive(Debug, Deserialize)]
struct FolderSizeParams {
    path: String,
    #[serde(default)]
    show_hidden: bool,
    #[serde(default)]
    show_system: bool,
}

#[derive(Debug, Deserialize)]
struct DirStatsParams {
    path: String,
    #[serde(default)]
    recursive: bool,
    #[serde(default)]
    show_hidden: bool,
    #[serde(default)]
    show_system: bool,
}

#[derive(Debug, Deserialize)]
struct DirStatsStreamParams {
    request_id: u64,
    path: String,
    #[serde(default)]
    recursive: bool,
    #[serde(default)]
    show_hidden: bool,
    #[serde(default)]
    show_system: bool,
}

#[derive(Debug, Deserialize)]
struct DirStatsCancelParams {
    path: String,
}

#[derive(Debug, Deserialize)]
struct SearchDirParams {
    base_path: String,
    query: String,
    request_id: u64,
    #[serde(default)]
    scope: String,
    #[serde(default)]
    show_hidden: bool,
    #[serde(default)]
    show_system: bool,
    #[serde(default)]
    full_text: bool,
}

#[derive(Debug, Deserialize)]
struct SetJumpListParams {
    recent: Vec<String>,
}

#[derive(Debug, Deserialize)]
struct GetIconParams {
    path: String,
    #[serde(default)]
    size: Option<u32>,
}

#[derive(Debug, Deserialize)]
struct GetStockIconParams {
    id: u32,
    #[serde(default)]
    size: Option<u32>,
}

#[derive(Debug, Deserialize)]
struct GetNewItemIconParams {
    name: String,
    is_dir: bool,
    #[serde(default)]
    size: Option<u32>,
}

#[derive(Debug, Deserialize)]
struct ConfirmMessageBoxParams {
    title: String,
    message: String,
}

#[derive(Debug, Deserialize)]
struct ConfirmTaskDialogParams {
    title: String,
    instruction: String,
    content: String,
    ok_text: String,
    cancel_text: String,
    #[serde(default)]
    icon_path: Option<String>,
    #[serde(default)]
    icon_name: Option<String>,
    #[serde(default)]
    icon_is_dir: Option<bool>,
    #[serde(default)]
    width: Option<u32>,
}

#[derive(Debug, Deserialize)]
struct GetBasicFileInfoParams {
    path: String,
}

#[derive(Debug, Serialize)]
struct BasicFileInfo {
    is_dir: bool,
    type_name: Option<String>,
    size_bytes: Option<u64>,
}

#[derive(Debug, Deserialize)]
struct GetIconsBatchParams {
    paths: Vec<String>,
    #[serde(default)]
    size: Option<u32>,
}

#[derive(Debug, Serialize)]
struct IconBatchItem {
    path: String,
    icon_png_base64: Option<String>,
}

#[derive(Debug, Serialize)]
struct RootItemDetailed {
    path: String,
    label: String,
    icon_png_base64: Option<String>,
    total_bytes: Option<u64>,
    free_bytes: Option<u64>,
}

#[derive(Debug, Serialize)]
struct ShellEntryItem {
    name: String,
    path: String,
    is_dir: bool,
    size: Option<u64>,
    modified_ms: Option<u128>,
    original_location: Option<String>,
    deleted_ms: Option<u128>,
    item_type: Option<String>,
}

#[derive(Debug, Serialize, Clone)]
struct DirEntryItem {
    name: String,
    path: String,
    is_dir: bool,
    size: Option<u64>,
    modified_ms: Option<u128>,
}

#[derive(Debug, Serialize, Clone)]
struct SearchResultItem {
    name: String,
    path: String,
    is_dir: bool,
    size: Option<u64>,
    modified_ms: Option<u128>,
    snippet: Option<String>,
}

#[derive(Debug, Serialize, Clone)]
struct DirStatsResult {
    path: String,
    items: u64,
    files: u64,
    folders: u64,
    files_bytes: u64,
}

fn system_time_to_ms(t: SystemTime) -> Option<u128> {
    t.duration_since(UNIX_EPOCH).ok().map(|d| d.as_millis())
}

#[tauri::command]
fn list_dir(params: ListDirParams) -> Result<Vec<DirEntryItem>, String> {
    let mut items = Vec::new();
    let dir_path = PathBuf::from(params.path);
    for entry in fs::read_dir(&dir_path).map_err(|e| e.to_string())? {
        let entry = entry.map_err(|e| e.to_string())?;
        let full_path = entry.path();
        let mut file_name = entry.file_name().to_string_lossy().to_string();
        let metadata = entry.metadata().ok();
        #[cfg(windows)]
        {
            if !params.show_system {
                let is_system = metadata
                    .as_ref()
                    .map(|m| (m.file_attributes() & 0x4) != 0)
                    .unwrap_or(false);
                if is_system {
                    continue;
                }
            }
            if !params.show_hidden {
                let is_hidden = metadata
                    .as_ref()
                    .map(|m| (m.file_attributes() & 0x2) != 0)
                    .unwrap_or(false);
                if is_hidden {
                    continue;
                }
            }

            unsafe {
                let p = full_path.to_string_lossy().to_string();
                let wide = to_wide_null(&p);
                let mut info = SHFILEINFOW::default();
                let flags = SHGFI_DISPLAYNAME;
                let ret = SHGetFileInfoW(
                    PCWSTR(wide.as_ptr()),
                    FILE_FLAGS_AND_ATTRIBUTES(0),
                    Some(&mut info),
                    std::mem::size_of::<SHFILEINFOW>() as u32,
                    flags,
                );
                if ret != 0 {
                    let disp = utf16_z_to_string(&info.szDisplayName);
                    if !disp.trim().is_empty() {
                        file_name = disp;
                    }
                }
            }
        }
        #[cfg(not(windows))]
        {
            if !params.show_hidden && file_name.starts_with('.') {
                continue;
            }
        }
        let is_dir = metadata.as_ref().map(|m| m.is_dir()).unwrap_or(false);
        let size = metadata
            .as_ref()
            .map(|m| if m.is_file() { Some(m.len()) } else { None })
            .flatten();
        let modified_ms = metadata
            .as_ref()
            .and_then(|m| m.modified().ok())
            .and_then(system_time_to_ms);
        items.push(DirEntryItem {
            name: file_name,
            path: full_path.to_string_lossy().to_string(),
            is_dir,
            size,
            modified_ms,
        });
    }
    Ok(items)
}

#[tauri::command]
async fn dir_stats(params: DirStatsParams) -> Result<DirStatsResult, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        let base = PathBuf::from(&p);
        let m = fs::metadata(&base).map_err(|e| e.to_string())?;
        if !m.is_dir() {
            return Err("不是文件夹".to_string());
        }
        let show_hidden = params.show_hidden;
        let show_system = params.show_system;
        let recursive = params.recursive;

        let mut items: u64 = 0;
        let mut files: u64 = 0;
        let mut folders: u64 = 0;
        let mut files_bytes: u64 = 0;

        if !recursive {
            for entry in fs::read_dir(&base).map_err(|e| e.to_string())? {
                let entry = entry.map_err(|e| e.to_string())?;
                let metadata = entry.metadata().ok();
                #[cfg(windows)]
                {
                    let attrs = metadata.as_ref().map(|m| m.file_attributes()).unwrap_or(0);
                    if (attrs & 0x400) != 0 {
                        continue;
                    }
                    if !show_system && (attrs & 0x4) != 0 {
                        continue;
                    }
                    if !show_hidden && (attrs & 0x2) != 0 {
                        continue;
                    }
                }
                #[cfg(not(windows))]
                {
                    let file_name = entry.file_name().to_string_lossy().to_string();
                    if !show_hidden && file_name.starts_with('.') {
                        continue;
                    }
                }
                let is_dir = metadata.as_ref().map(|m| m.is_dir()).unwrap_or(false);
                items += 1;
                if is_dir {
                    folders += 1;
                    continue;
                }
                files += 1;
                if let Some(len) = metadata.as_ref().map(|m| if m.is_file() { Some(m.len()) } else { None }).flatten() {
                    files_bytes = files_bytes.saturating_add(len);
                }
            }
            return Ok(DirStatsResult {
                path: p,
                items,
                files,
                folders,
                files_bytes,
            });
        }

        #[cfg(windows)]
        fn should_skip_entry(meta: Option<&std::fs::Metadata>, show_hidden: bool, show_system: bool) -> bool {
            let attrs = meta.map(|m| m.file_attributes()).unwrap_or(0);
            if (attrs & 0x400) != 0 {
                return true;
            }
            if !show_system && (attrs & 0x4) != 0 {
                return true;
            }
            if !show_hidden && (attrs & 0x2) != 0 {
                return true;
            }
            false
        }

        let walker = WalkDir::new(&base)
            .min_depth(1)
            .follow_links(false)
            .into_iter()
            .filter_entry(|e| {
                #[cfg(windows)]
                {
                    let meta = e.metadata().ok();
                    if e.file_type().is_symlink() {
                        return false;
                    }
                    return !should_skip_entry(meta.as_ref(), show_hidden, show_system);
                }
                #[cfg(not(windows))]
                {
                    let name = e.file_name().to_string_lossy();
                    if !show_hidden && name.starts_with('.') {
                        return false;
                    }
                    true
                }
            });

        for entry in walker {
            let entry = match entry {
                Ok(v) => v,
                Err(_) => continue,
            };
            if entry.file_type().is_symlink() {
                continue;
            }
            let meta = entry.metadata().ok();
            #[cfg(windows)]
            {
                if should_skip_entry(meta.as_ref(), show_hidden, show_system) {
                    continue;
                }
            }
            #[cfg(not(windows))]
            {
                let name = entry.file_name().to_string_lossy();
                if !show_hidden && name.starts_with('.') {
                    continue;
                }
            }

            items += 1;
            if entry.file_type().is_dir() {
                folders += 1;
                continue;
            }
            if entry.file_type().is_file() {
                files += 1;
                if let Some(len) = meta.as_ref().map(|m| m.len()) {
                    files_bytes = files_bytes.saturating_add(len);
                }
            }
        }

        Ok(DirStatsResult {
            path: p,
            items,
            files,
            folders,
            files_bytes,
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

#[derive(Serialize, Clone)]
struct DirStatsProgressPayload {
    request_id: u64,
    path: String,
    stats: DirStatsResult,
    done: bool,
}

#[tauri::command]
async fn dir_stats_cancel(params: DirStatsCancelParams) -> Result<(), String> {
    let p = params.path.trim().to_string();
    if p.is_empty() {
        return Ok(());
    }
    let key = p.to_lowercase();
    if let Ok(mut map) = dir_stats_latest_map().lock() {
        map.insert(key, 0);
    }
    Ok(())
}

#[tauri::command]
async fn dir_stats_stream(window: tauri::Window, params: DirStatsStreamParams) -> Result<DirStatsResult, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let rid = params.request_id;
        if rid == 0 {
            return Err("request_id 不能为空".to_string());
        }
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        let base = PathBuf::from(&p);
        let m = fs::metadata(&base).map_err(|e| e.to_string())?;
        if !m.is_dir() {
            return Err("不是文件夹".to_string());
        }
        let show_hidden = params.show_hidden;
        let show_system = params.show_system;
        let recursive = params.recursive;

        let key = p.to_lowercase();
        if let Ok(mut map) = dir_stats_latest_map().lock() {
            map.insert(key.clone(), rid);
        }

        let mut items: u64 = 0;
        let mut files: u64 = 0;
        let mut folders: u64 = 0;
        let mut files_bytes: u64 = 0;

        let emit = |done: bool,
                    window: &tauri::Window,
                    p: &str,
                    rid: u64,
                    items: u64,
                    files: u64,
                    folders: u64,
                    files_bytes: u64| {
            let _ = window.emit(
                "dir_stats_progress",
                DirStatsProgressPayload {
                    request_id: rid,
                    path: p.to_string(),
                    stats: DirStatsResult {
                        path: p.to_string(),
                        items,
                        files,
                        folders,
                        files_bytes,
                    },
                    done,
                },
            );
        };

        let is_cancelled = |key: &str, rid: u64| -> bool {
            if let Ok(map) = dir_stats_latest_map().lock() {
                return map.get(key).copied().unwrap_or(0) != rid;
            }
            false
        };

        if !recursive {
            for entry in fs::read_dir(&base).map_err(|e| e.to_string())? {
                if is_cancelled(&key, rid) {
                    return Ok(DirStatsResult {
                        path: p,
                        items,
                        files,
                        folders,
                        files_bytes,
                    });
                }
                let entry = entry.map_err(|e| e.to_string())?;
                let metadata = entry.metadata().ok();
                #[cfg(windows)]
                {
                    let attrs = metadata.as_ref().map(|m| m.file_attributes()).unwrap_or(0);
                    if (attrs & 0x400) != 0 {
                        continue;
                    }
                    if !show_system && (attrs & 0x4) != 0 {
                        continue;
                    }
                    if !show_hidden && (attrs & 0x2) != 0 {
                        continue;
                    }
                }
                #[cfg(not(windows))]
                {
                    let file_name = entry.file_name().to_string_lossy().to_string();
                    if !show_hidden && file_name.starts_with('.') {
                        continue;
                    }
                }
                let is_dir = metadata.as_ref().map(|m| m.is_dir()).unwrap_or(false);
                items += 1;
                if is_dir {
                    folders += 1;
                    continue;
                }
                files += 1;
                if let Some(len) = metadata
                    .as_ref()
                    .map(|m| if m.is_file() { Some(m.len()) } else { None })
                    .flatten()
                {
                    files_bytes = files_bytes.saturating_add(len);
                }
            }
            emit(true, &window, &p, rid, items, files, folders, files_bytes);
            return Ok(DirStatsResult {
                path: p,
                items,
                files,
                folders,
                files_bytes,
            });
        }

        #[cfg(windows)]
        fn should_skip_meta(meta: &std::fs::Metadata, show_hidden: bool, show_system: bool) -> bool {
            let attrs = meta.file_attributes();
            if (attrs & 0x400) != 0 {
                return true;
            }
            if !show_system && (attrs & 0x4) != 0 {
                return true;
            }
            if !show_hidden && (attrs & 0x2) != 0 {
                return true;
            }
            false
        }

        let mut last_emit = Instant::now();
        let mut it = WalkDir::new(base).follow_links(false).into_iter();
        while let Some(next) = it.next() {
            if is_cancelled(&key, rid) {
                return Ok(DirStatsResult {
                    path: p,
                    items,
                    files,
                    folders,
                    files_bytes,
                });
            }
            let entry = match next {
                Ok(v) => v,
                Err(_) => continue,
            };
            if entry.depth() == 0 {
                continue;
            }
            let ft = entry.file_type();
            if ft.is_symlink() {
                continue;
            }
            let meta = entry.metadata().ok();
            #[cfg(windows)]
            {
                if let Some(md) = meta.as_ref() {
                    if should_skip_meta(md, show_hidden, show_system) {
                        if ft.is_dir() {
                            it.skip_current_dir();
                        }
                        continue;
                    }
                }
            }
            #[cfg(not(windows))]
            {
                if !show_hidden {
                    let name = entry.file_name().to_string_lossy();
                    if name.starts_with('.') {
                        if ft.is_dir() {
                            it.skip_current_dir();
                        }
                        continue;
                    }
                }
            }

            items += 1;
            if ft.is_dir() {
                folders += 1;
            } else if ft.is_file() {
                files += 1;
                if let Some(len) = meta.as_ref().map(|m| m.len()) {
                    files_bytes = files_bytes.saturating_add(len);
                }
            }

            if last_emit.elapsed() >= Duration::from_millis(200) {
                emit(false, &window, &p, rid, items, files, folders, files_bytes);
                last_emit = Instant::now();
            }
        }

        emit(true, &window, &p, rid, items, files, folders, files_bytes);
        Ok(DirStatsResult {
            path: p,
            items,
            files,
            folders,
            files_bytes,
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

#[derive(Serialize, Clone)]
struct FolderSizeProgressPayload {
    request_id: u64,
    path: String,
    size: u64,
    done: bool,
}

#[tauri::command]
async fn folder_size(params: FolderSizeParams) -> Result<u64, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        let show_hidden = params.show_hidden;
        let show_system = params.show_system;
        let base = PathBuf::from(&p);
        let m = fs::metadata(&base).map_err(|e| e.to_string())?;
        if !m.is_dir() {
            return Err("不是文件夹".to_string());
        }
        let dir_modified_ms = m
            .modified()
            .ok()
            .and_then(system_time_to_ms)
            .unwrap_or(0);

        let cache_key = format!(
            "{}|{}|{}",
            p.to_lowercase(),
            if show_hidden { 1 } else { 0 },
            if show_system { 1 } else { 0 }
        );
        let ttl = Duration::from_secs(10 * 60);
        if let Ok(cache) = folder_size_cache().lock() {
            if let Some(hit) = cache.get(&cache_key) {
                if hit.dir_modified_ms == dir_modified_ms && hit.at.elapsed() < ttl {
                    return Ok(hit.size);
                }
            }
        }

        let mut total: u64 = 0;
        let mut it = WalkDir::new(base).follow_links(false).into_iter();
        while let Some(next) = it.next() {
            let entry = match next {
                Ok(v) => v,
                Err(_) => continue,
            };
            if entry.depth() == 0 {
                continue;
            }
            let ft = entry.file_type();
            if ft.is_dir() {
                #[cfg(windows)]
                {
                    if !show_system || !show_hidden {
                        if let Ok(md) = entry.metadata() {
                            let attrs = md.file_attributes();
                            if (!show_system && (attrs & 0x4) != 0) || (!show_hidden && (attrs & 0x2) != 0) {
                                it.skip_current_dir();
                                continue;
                            }
                        }
                    }
                }
                #[cfg(not(windows))]
                {
                    if !show_hidden {
                        let name = entry.file_name().to_string_lossy();
                        if name.starts_with('.') {
                            it.skip_current_dir();
                            continue;
                        }
                    }
                }
                continue;
            }
            if !ft.is_file() {
                continue;
            }
            let md = match entry.metadata() {
                Ok(v) => v,
                Err(_) => continue,
            };
            #[cfg(windows)]
            {
                if !show_system || !show_hidden {
                    let attrs = md.file_attributes();
                    if !show_system && (attrs & 0x4) != 0 {
                        continue;
                    }
                    if !show_hidden && (attrs & 0x2) != 0 {
                        continue;
                    }
                }
            }
            #[cfg(not(windows))]
            {
                if !show_hidden {
                    let name = entry.file_name().to_string_lossy();
                    if name.starts_with('.') {
                        continue;
                    }
                }
            }
            total = total.saturating_add(md.len());
        }

        if let Ok(mut cache) = folder_size_cache().lock() {
            cache.insert(
                cache_key,
                FolderSizeCacheEntry {
                    size: total,
                    dir_modified_ms,
                    at: Instant::now(),
                },
            );
        }
        Ok(total)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn folder_size_stream(window: tauri::Window, params: FolderSizeParams) -> Result<u64, String> {
    let request_id = FOLDER_SIZE_REQ_SEQ.fetch_add(1, Ordering::Relaxed);
    let window_label = window.label().to_string();
    tauri::async_runtime::spawn_blocking(move || {
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        let show_hidden = params.show_hidden;
        let show_system = params.show_system;
        let base = PathBuf::from(&p);
        let m = fs::metadata(&base).map_err(|e| e.to_string())?;
        if !m.is_dir() {
            return Err("不是文件夹".to_string());
        }
        let dir_modified_ms = m
            .modified()
            .ok()
            .and_then(system_time_to_ms)
            .unwrap_or(0);

        let cache_key = format!(
            "{}|{}|{}",
            p.to_lowercase(),
            if show_hidden { 1 } else { 0 },
            if show_system { 1 } else { 0 }
        );
        let latest_key = format!("{window_label}::{cache_key}");
        if let Ok(mut m) = folder_size_latest_map().lock() {
            m.insert(latest_key.clone(), request_id);
        }

        let ttl = Duration::from_secs(10 * 60);
        if let Ok(cache) = folder_size_cache().lock() {
            if let Some(hit) = cache.get(&cache_key) {
                if hit.dir_modified_ms == dir_modified_ms && hit.at.elapsed() < ttl {
                    let _ = window.emit(
                        "folder_size_progress",
                        FolderSizeProgressPayload {
                            request_id,
                            path: p,
                            size: hit.size,
                            done: true,
                        },
                    );
                    return Ok(request_id);
                }
            }
        }

        let mut total: u64 = 0;
        let mut last_emit = Instant::now();
        let mut last_emitted_total: u64 = 0;
        let mut it = WalkDir::new(base).follow_links(false).into_iter();
        while let Some(next) = it.next() {
            let entry = match next {
                Ok(v) => v,
                Err(_) => continue,
            };
            if entry.depth() == 0 {
                continue;
            }
            let ft = entry.file_type();
            if ft.is_dir() {
                #[cfg(windows)]
                {
                    if !show_system || !show_hidden {
                        if let Ok(md) = entry.metadata() {
                            let attrs = md.file_attributes();
                            if (!show_system && (attrs & 0x4) != 0) || (!show_hidden && (attrs & 0x2) != 0) {
                                it.skip_current_dir();
                                continue;
                            }
                        }
                    }
                }
                #[cfg(not(windows))]
                {
                    if !show_hidden {
                        let name = entry.file_name().to_string_lossy();
                        if name.starts_with('.') {
                            it.skip_current_dir();
                            continue;
                        }
                    }
                }
                continue;
            }
            if !ft.is_file() {
                continue;
            }
            let md = match entry.metadata() {
                Ok(v) => v,
                Err(_) => continue,
            };
            #[cfg(windows)]
            {
                if !show_system || !show_hidden {
                    let attrs = md.file_attributes();
                    if !show_system && (attrs & 0x4) != 0 {
                        continue;
                    }
                    if !show_hidden && (attrs & 0x2) != 0 {
                        continue;
                    }
                }
            }
            #[cfg(not(windows))]
            {
                if !show_hidden {
                    let name = entry.file_name().to_string_lossy();
                    if name.starts_with('.') {
                        continue;
                    }
                }
            }
            total = total.saturating_add(md.len());

            if total != last_emitted_total && last_emit.elapsed() >= Duration::from_millis(120) {
                let should_cancel = folder_size_latest_map()
                    .lock()
                    .ok()
                    .and_then(|m| m.get(&latest_key).cloned())
                    .map(|v| v != request_id)
                    .unwrap_or(false);
                if should_cancel {
                    return Ok(request_id);
                }
                let _ = window.emit(
                    "folder_size_progress",
                    FolderSizeProgressPayload {
                        request_id,
                        path: p.clone(),
                        size: total,
                        done: false,
                    },
                );
                last_emit = Instant::now();
                last_emitted_total = total;
            }
        }

        let should_cancel = folder_size_latest_map()
            .lock()
            .ok()
            .and_then(|m| m.get(&latest_key).cloned())
            .map(|v| v != request_id)
            .unwrap_or(false);
        if should_cancel {
            return Ok(request_id);
        }

        if let Ok(mut cache) = folder_size_cache().lock() {
            cache.insert(
                cache_key,
                FolderSizeCacheEntry {
                    size: total,
                    dir_modified_ms,
                    at: Instant::now(),
                },
            );
        }

        let _ = window.emit(
            "folder_size_progress",
            FolderSizeProgressPayload {
                request_id,
                path: p,
                size: total,
                done: true,
            },
        );
        Ok(request_id)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn search_dir(window: tauri::Window, params: SearchDirParams) -> Result<Vec<SearchResultItem>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        #[cfg(windows)]
        let window_label = window.label().to_string();
        let base_path = params.base_path.trim().to_string();
        let base = PathBuf::from(&base_path);
        let q = params.query.trim().to_string();
        let q_lower = q.to_lowercase();
        if q.is_empty() {
            return Ok(Vec::<SearchResultItem>::new());
        }
        let request_id = params.request_id;
        let scope = params.scope.trim().to_string();
        let show_hidden = params.show_hidden;
        let show_system = params.show_system;
        let full_text = params.full_text;
        #[cfg(windows)]
        let cancel_key = if scope.is_empty() {
            window_label.clone()
        } else {
            format!("{window_label}::{scope}")
        };

        #[derive(Serialize, Clone)]
        struct SearchProgressPayload {
            request_id: u64,
            scanned: u64,
            matched: u64,
            done: bool,
        }

        #[derive(Serialize, Clone)]
        struct SearchResultBatchPayload {
            request_id: u64,
            results: Vec<SearchResultItem>,
        }

        #[cfg(windows)]
        {
            if let Ok(mut map) = search_latest_map().lock() {
                map.insert(cancel_key.clone(), request_id);
            }
        }

        let q_is_ascii = q_lower.is_ascii();

        let mut batch: Vec<SearchResultItem> = Vec::new();
        let mut scanned: u64 = 0;
        let mut matched: u64 = 0;
        let mut last_emit = std::time::Instant::now();

        fn clamp_snippet(s: &str, start: usize, end: usize) -> String {
            let mut a = start.min(s.len());
            let mut b = end.min(s.len());
            while a > 0 && !s.is_char_boundary(a) {
                a -= 1;
            }
            while b < s.len() && !s.is_char_boundary(b) {
                b += 1;
            }
            let mut out = String::new();
            if a > 0 {
                out.push('…');
            }
            out.push_str(&s[a..b]);
            if b < s.len() {
                out.push('…');
            }
            out
        }

        fn normalize_snippet(s: String) -> String {
            let compact = s
                .replace('\r', " ")
                .replace('\n', " ")
                .replace('\t', " ")
                .split_whitespace()
                .collect::<Vec<_>>()
                .join(" ");
            compact.trim().to_string()
        }

        fn should_skip_content_by_ext(path: &PathBuf) -> bool {
            let ext = path
                .extension()
                .map(|x| x.to_string_lossy().to_string().to_lowercase())
                .unwrap_or_default();
            if ext.is_empty() {
                return false;
            }
            matches!(
                ext.as_str(),
                "exe"
                    | "dll"
                    | "sys"
                    | "msi"
                    | "zip"
                    | "rar"
                    | "7z"
                    | "gz"
                    | "bz2"
                    | "xz"
                    | "png"
                    | "jpg"
                    | "jpeg"
                    | "gif"
                    | "bmp"
                    | "webp"
                    | "ico"
                    | "pdf"
                    | "mp3"
                    | "wav"
                    | "flac"
                    | "mp4"
                    | "mkv"
                    | "mov"
                    | "avi"
                    | "wmv"
            )
        }

        fn find_content_snippet(
            path: &PathBuf,
            q: &str,
            q_lower: &str,
            q_is_ascii: bool,
        ) -> Option<String> {
            const MAX_BYTES: u64 = 2 * 1024 * 1024;
            let meta = fs::metadata(path).ok()?;
            if !meta.is_file() {
                return None;
            }
            if meta.len() > MAX_BYTES {
                return None;
            }
            if should_skip_content_by_ext(path) {
                return None;
            }

            let bytes = fs::read(path).ok()?;
            if bytes.is_empty() {
                return None;
            }
            let hay = String::from_utf8_lossy(&bytes);
            if hay.is_empty() {
                return None;
            }

            let pos = if q_is_ascii {
                let low = hay.to_ascii_lowercase();
                low.find(q_lower)
            } else {
                hay.find(q)
            }?;

            let q_len = q.len().max(1);
            let start = pos.saturating_sub(48);
            let end = (pos + q_len).saturating_add(96);
            let snippet = clamp_snippet(&hay, start, end);
            let snippet = normalize_snippet(snippet);
            if snippet.is_empty() {
                None
            } else {
                Some(snippet)
            }
        }

        
        let walker = WalkDir::new(base).follow_links(false).into_iter().filter_entry(move |e| {
            if e.depth() == 0 {
                return true;
            }
            #[cfg(not(windows))]
            {
                if !show_hidden {
                    let name = e.file_name().to_string_lossy();
                    if name.starts_with('.') {
                        return false;
                    }
                }
            }
            #[cfg(windows)]
            {
                if !show_system || !show_hidden {
                    if let Ok(m) = e.metadata() {
                        let attrs = m.file_attributes();
                        if !show_system && (attrs & 0x4) != 0 {
                            return false;
                        }
                        if !show_hidden && (attrs & 0x2) != 0 {
                            return false;
                        }
                    }
                }
            }
            true
        });
        for entry in walker {
            let entry = match entry {
                Ok(v) => v,
                Err(_) => continue,
            };
            if entry.depth() == 0 {
                continue;
            }
            scanned = scanned.saturating_add(1);

            #[cfg(windows)]
            {
                if scanned % 256 == 0 {
                    let latest = search_latest_map()
                        .lock()
                        .ok()
                        .and_then(|m| m.get(&cancel_key).copied())
                        .unwrap_or(request_id);
                    if latest != request_id {
                        let _ = window.emit(
                            "search_progress",
                            SearchProgressPayload {
                                request_id,
                                scanned,
                                matched,
                                done: true,
                            },
                        );
                        if let Ok(mut map) = search_latest_map().lock() {
                            if map.get(&cancel_key).copied() == Some(request_id) {
                                map.remove(&cancel_key);
                            }
                        }
                        return Ok(Vec::<SearchResultItem>::new());
                    }
                }
            }

            if last_emit.elapsed() >= std::time::Duration::from_millis(120) {
                let _ = window.emit(
                    "search_progress",
                    SearchProgressPayload {
                        request_id,
                        scanned,
                        matched,
                        done: false,
                    },
                );
                last_emit = std::time::Instant::now();
            }
            let name = entry.file_name().to_string_lossy().to_string();
            #[cfg(not(windows))]
            {
                if !params.show_hidden && name.starts_with('.') {
                    continue;
                }
            }
            let name_match = if q_is_ascii {
                name.to_ascii_lowercase().contains(&q_lower)
            } else {
                name.to_lowercase().contains(&q_lower)
            };
            let path = entry.path().to_string_lossy().to_string();
            if path.trim().is_empty() {
                continue;
            }

            let is_dir = entry.file_type().is_dir();
            let metadata = entry.metadata().ok();
            #[cfg(windows)]
            {
                if !params.show_system {
                    let is_system = metadata
                        .as_ref()
                        .map(|m| (m.file_attributes() & 0x4) != 0)
                        .unwrap_or(false);
                    if is_system {
                        continue;
                    }
                }
                if !params.show_hidden {
                    let is_hidden = metadata
                        .as_ref()
                        .map(|m| (m.file_attributes() & 0x2) != 0)
                        .unwrap_or(false);
                    if is_hidden {
                        continue;
                    }
                }
            }
            let size = metadata
                .as_ref()
                .map(|m| if m.is_file() { Some(m.len()) } else { None })
                .flatten();
            let modified_ms = metadata
                .as_ref()
                .and_then(|m| m.modified().ok())
                .and_then(system_time_to_ms);

            let mut snippet: Option<String> = None;
            if !name_match && full_text && !is_dir {
                snippet = find_content_snippet(&entry.path().to_path_buf(), &q, &q_lower, q_is_ascii);
                if snippet.is_none() {
                    continue;
                }
            } else if !name_match {
                continue;
            }

            let item = SearchResultItem {
                name,
                path,
                is_dir,
                size,
                modified_ms,
                snippet,
            };
            batch.push(item);
            matched = matched.saturating_add(1);

            if batch.len() >= 120 {
                let payload = SearchResultBatchPayload {
                    request_id,
                    results: std::mem::take(&mut batch),
                };
                let _ = window.emit("search_result_batch", payload);
            }
        }
        if !batch.is_empty() {
            let payload = SearchResultBatchPayload {
                request_id,
                results: std::mem::take(&mut batch),
            };
            let _ = window.emit("search_result_batch", payload);
        }
        let _ = window.emit(
            "search_progress",
            SearchProgressPayload {
                request_id,
                scanned,
                matched,
                done: true,
            },
        );
        #[cfg(windows)]
        if let Ok(mut map) = search_latest_map().lock() {
            if map.get(&cancel_key).copied() == Some(request_id) {
                map.remove(&cancel_key);
            }
        }
        Ok(Vec::<SearchResultItem>::new())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
fn list_roots() -> Vec<String> {
    #[cfg(windows)]
    {
        let mut roots = Vec::new();
        for i in 0..26u8 {
            let letter = (b'A' + i) as char;
            let path = format!("{letter}:\\");
            if std::path::Path::new(&path).exists() {
                roots.push(path);
            }
        }
        return roots;
    }
    #[cfg(not(windows))]
    vec!["/".to_string()]
}

#[tauri::command]
async fn set_jump_list(params: SetJumpListParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        let recent = params.recent;
        return tauri::async_runtime::spawn_blocking(move || set_jump_list_windows(recent))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
async fn list_quick_access() -> Result<Vec<QuickAccessEntry>, String> {
    #[cfg(windows)]
    {
        return tauri::async_runtime::spawn_blocking(move || list_quick_access_windows())
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(windows))]
    {
        Ok(vec![])
    }
}

#[tauri::command]
async fn is_in_quick_access(params: QuickAccessPathParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        let p = params.path.trim().to_string();
        return tauri::async_runtime::spawn_blocking(move || is_in_quick_access_windows(&p))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
async fn pin_to_quick_access(window: tauri::Window, params: QuickAccessPathParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        return tauri::async_runtime::spawn_blocking(move || pin_to_quick_access_windows(window, &p))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
async fn remove_from_quick_access(window: tauri::Window, params: QuickAccessPathParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        return tauri::async_runtime::spawn_blocking(move || remove_from_quick_access_windows(window, &p))
            .await
            .map_err(|e| e.to_string())?;
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        let _ = params;
        Ok(false)
    }
}

#[cfg(windows)]
fn to_wide_null(s: &str) -> Vec<u16> {
    use std::os::windows::ffi::OsStrExt;
    std::ffi::OsStr::new(s).encode_wide().chain(std::iter::once(0)).collect()
}

#[cfg(windows)]
fn utf16_z_to_string(buf: &[u16]) -> String {
    let end = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
    String::from_utf16_lossy(&buf[..end])
}

#[cfg(windows)]
fn drive_label_for_root(path: &str) -> String {
    let letter = path.chars().next().unwrap_or('C').to_ascii_uppercase();
    let mut vol = [0u16; 260];
    let wide = to_wide_null(path);
    let ok = unsafe {
        GetVolumeInformationW(
            PCWSTR(wide.as_ptr()),
            Some(&mut vol),
            None,
            None,
            None,
            None,
        )
        .is_ok()
    };
    let vol_name = if ok { utf16_z_to_string(&vol) } else { String::new() };
    if !vol_name.trim().is_empty() {
        return format!("{vol_name} ({letter}:)");
    }

    const DRIVE_UNKNOWN: u32 = 0;
    const DRIVE_NO_ROOT_DIR: u32 = 1;
    const DRIVE_REMOVABLE: u32 = 2;
    const DRIVE_FIXED: u32 = 3;
    const DRIVE_REMOTE: u32 = 4;
    const DRIVE_CDROM: u32 = 5;
    const DRIVE_RAMDISK: u32 = 6;

    let drive_type = unsafe { GetDriveTypeW(PCWSTR(wide.as_ptr())) };
    let base = match drive_type {
        DRIVE_REMOVABLE => "可移动磁盘",
        DRIVE_REMOTE => "网络驱动器",
        DRIVE_CDROM => "DVD 驱动器",
        DRIVE_RAMDISK => "RAM 磁盘",
        DRIVE_FIXED => "本地磁盘",
        DRIVE_UNKNOWN | DRIVE_NO_ROOT_DIR => "磁盘",
        _ => "磁盘",
    };
    format!("{base} ({letter}:)")
}

#[cfg(windows)]
fn drive_space_for_root(path: &str) -> Option<(u64, u64)> {
    let wide = to_wide_null(path);
    let mut free_bytes: u64 = 0;
    let mut total_bytes: u64 = 0;
    let mut total_free_bytes: u64 = 0;
    let ok = unsafe {
        GetDiskFreeSpaceExW(
            PCWSTR(wide.as_ptr()),
            Some(&mut free_bytes),
            Some(&mut total_bytes),
            Some(&mut total_free_bytes),
        )
        .is_ok()
    };
    if !ok || total_bytes == 0 {
        return None;
    }
    Some((total_bytes, total_free_bytes))
}

#[cfg(windows)]
fn normalize_compare_path_win(p: &str) -> String {
    p.trim().replace('/', "\\").to_ascii_lowercase()
}

#[cfg(windows)]
fn quick_access_namespace() -> &'static str {
    "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}"
}

#[cfg(windows)]
unsafe fn context_menu_has_verb(cm: &IContextMenu, verb: &str) -> bool {
    let hmenu: HMENU = match CreatePopupMenu() {
        Ok(v) => v,
        Err(_) => return false,
    };

    let id_first: u32 = 1;
    let id_last: u32 = 0x7fff;
    if cm.QueryContextMenu(hmenu, 0, id_first, id_last, CMF_NORMAL).is_err() {
        let _ = DestroyMenu(hmenu);
        return false;
    }

    let count = GetMenuItemCount(Some(hmenu));
    if count <= 0 {
        let _ = DestroyMenu(hmenu);
        return false;
    }

    let target = verb.to_ascii_lowercase();
    let mut found = false;
    for i in 0..count {
        let cmd_id = GetMenuItemID(hmenu, i);
        if cmd_id == u32::MAX || cmd_id == 0 {
            continue;
        }
        let id = cmd_id as u32;
        if id < id_first {
            continue;
        }
        let verb_offset = id.saturating_sub(id_first) as usize;
        let mut buf = [0u8; 260];
        if cm
            .GetCommandString(verb_offset, 0x00000000, None, PSTR(buf.as_mut_ptr()), buf.len() as u32)
            .is_ok()
        {
            if let Ok(s) = std::ffi::CStr::from_bytes_until_nul(&buf).map(|x| x.to_string_lossy().to_string()) {
                if s.to_ascii_lowercase() == target {
                    found = true;
                    break;
                }
            }
        }
    }

    let _ = DestroyMenu(hmenu);
    found
}

#[cfg(windows)]
unsafe fn quick_access_child_has_verb(folder: &IShellFolder, child: *const ITEMIDLIST, verb: &str) -> bool {
    let hwnd = HWND(std::ptr::null_mut());
    let child_ptrs = [child];
    let cm: IContextMenu = match folder.GetUIObjectOf(hwnd, child_ptrs.as_slice(), None) {
        Ok(v) => v,
        Err(_) => return false,
    };
    context_menu_has_verb(&cm, verb)
}

#[cfg(windows)]
fn list_quick_access_windows() -> Result<Vec<QuickAccessEntry>, String> {
    let _com = com_init();
    unsafe {
        let desktop = SHGetDesktopFolder().map_err(|e| e.to_string())?;
        let wide = to_wide_null(quick_access_namespace());
        let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
        SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None).map_err(|e| e.to_string())?;
        if pidl.is_null() {
            return Ok(vec![]);
        }
        let pidl_guard = PidlGuard(pidl);
        let folder: IShellFolder = desktop
            .BindToObject::<_, IShellFolder>(pidl_guard.0, None)
            .map_err(|e| e.to_string())?;

        let mut enum_list: Option<IEnumIDList> = None;
        let flags = SHCONTF_FOLDERS.0 as u32;
        let hr = folder.EnumObjects(HWND(std::ptr::null_mut()), flags, &mut enum_list);
        if hr.0 < 0 {
            return Ok(vec![]);
        }
        let enum_list = match enum_list {
            Some(x) => x,
            None => return Ok(vec![]),
        };

        let mut seen = std::collections::HashSet::<String>::new();
        let mut out: Vec<QuickAccessEntry> = Vec::new();
        loop {
            if out.len() >= 24 {
                break;
            }
            let mut fetched: u32 = 0;
            let mut rgelt: [*mut ITEMIDLIST; 1] = [std::ptr::null_mut()];
            let hr = enum_list.Next(&mut rgelt, Some(&mut fetched as *mut u32));
            let child_pidl = rgelt[0];
            if hr.0 < 0 || fetched == 0 || child_pidl.is_null() {
                break;
            }
            let child_guard = PidlGuard(child_pidl);

            let mut disp: STRRET = std::mem::zeroed();
            if folder.GetDisplayNameOf(child_guard.0, SHGDN_FORPARSING, &mut disp).is_err() {
                continue;
            }
            let path = strret_to_string(&mut disp, child_guard.0).unwrap_or_default();
            if path.trim().is_empty() {
                continue;
            }
            let key = normalize_compare_path_win(&path);
            if key.is_empty() || seen.contains(&key) {
                continue;
            }
            seen.insert(key);

            let mut disp2: STRRET = std::mem::zeroed();
            let label = if folder.GetDisplayNameOf(child_guard.0, SHGDN_NORMAL, &mut disp2).is_ok() {
                let s = strret_to_string(&mut disp2, child_guard.0).unwrap_or_default();
                if s.trim().is_empty() {
                    jump_list_title_for_path(&path)
                } else {
                    s
                }
            } else {
                jump_list_title_for_path(&path)
            };
            let pinned = quick_access_child_has_verb(&folder, child_guard.0, "unpinfromhome");

            out.push(QuickAccessEntry {
                path,
                label,
                pinned,
            });
        }
        Ok(out)
    }
}

#[cfg(windows)]
fn is_in_quick_access_windows(path: &str) -> Result<bool, String> {
    let key = normalize_compare_path_win(path);
    if key.is_empty() {
        return Ok(false);
    }
    let list = list_quick_access_windows()?;
    Ok(list.into_iter().any(|x| normalize_compare_path_win(&x.path) == key))
}

#[cfg(windows)]
unsafe fn invoke_context_menu_verb(cm: &IContextMenu, hwnd: HWND, verb: &str) -> Result<(), String> {
    let c_verb = std::ffi::CString::new(verb).map_err(|_| "verb".to_string())?;
    let mut invoke: CMINVOKECOMMANDINFOEX = std::mem::zeroed();
    invoke.cbSize = std::mem::size_of::<CMINVOKECOMMANDINFOEX>() as u32;
    invoke.fMask = 0;
    invoke.hwnd = hwnd;
    invoke.lpVerb = PCSTR(c_verb.as_ptr() as *const u8);
    invoke.nShow = 1;
    cm.InvokeCommand(&mut invoke as *mut _ as *mut _)
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg(windows)]
fn pin_to_quick_access_windows(window: tauri::Window, path: &str) -> Result<bool, String> {
    let _com = com_init();
    unsafe {
        let hwnd = window.hwnd().map_err(|e| e.to_string())?;
        if let Ok((cm, _, _)) = shell_context_menu_for_paths(hwnd, &[path.to_string()]) {
            if invoke_context_menu_verb(&cm, hwnd, "pintohome").is_ok() {
                return Ok(true);
            }
        }
        let op = to_wide_null("pintohome");
        let file = to_wide_null(path);
        let res = ShellExecuteW(
            Some(hwnd),
            PCWSTR(op.as_ptr()),
            PCWSTR(file.as_ptr()),
            PCWSTR(std::ptr::null()),
            PCWSTR(std::ptr::null()),
            SW_SHOW,
        );
        Ok(res.0 as isize > 32)
    }
}

#[cfg(windows)]
fn remove_from_quick_access_windows(window: tauri::Window, path: &str) -> Result<bool, String> {
    let _com = com_init();
    unsafe {
        let hwnd = window.hwnd().map_err(|e| e.to_string())?;
        if let Ok((cm, _, _)) = shell_context_menu_for_paths(hwnd, &[path.to_string()]) {
            if invoke_context_menu_verb(&cm, hwnd, "unpinfromhome").is_ok() {
                return Ok(true);
            }
            if invoke_context_menu_verb(&cm, hwnd, "removefromhome").is_ok() {
                return Ok(true);
            }
        }

        let desktop = SHGetDesktopFolder().map_err(|e| e.to_string())?;
        let wide = to_wide_null(quick_access_namespace());
        let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
        SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None).map_err(|e| e.to_string())?;
        if pidl.is_null() {
            return Ok(false);
        }
        let pidl_guard = PidlGuard(pidl);
        let qa_folder: IShellFolder = desktop
            .BindToObject::<_, IShellFolder>(pidl_guard.0, None)
            .map_err(|e| e.to_string())?;

        let mut enum_list: Option<IEnumIDList> = None;
        let flags = SHCONTF_FOLDERS.0 as u32;
        let hr = qa_folder.EnumObjects(HWND(std::ptr::null_mut()), flags, &mut enum_list);
        if hr.0 < 0 {
            return Ok(false);
        }
        let enum_list = match enum_list {
            Some(x) => x,
            None => return Ok(false),
        };

        let target = normalize_compare_path_win(path);
        loop {
            let mut fetched: u32 = 0;
            let mut rgelt: [*mut ITEMIDLIST; 1] = [std::ptr::null_mut()];
            let hr = enum_list.Next(&mut rgelt, Some(&mut fetched as *mut u32));
            let child_pidl = rgelt[0];
            if hr.0 < 0 || fetched == 0 || child_pidl.is_null() {
                break;
            }
            let child_guard = PidlGuard(child_pidl);

            let mut disp: STRRET = std::mem::zeroed();
            if qa_folder.GetDisplayNameOf(child_guard.0, SHGDN_FORPARSING, &mut disp).is_err() {
                continue;
            }
            let p = strret_to_string(&mut disp, child_guard.0).unwrap_or_default();
            if normalize_compare_path_win(&p) != target {
                continue;
            }

            let child_ptrs = [child_guard.0 as *const ITEMIDLIST];
            let cm = qa_folder
                .GetUIObjectOf::<IContextMenu>(hwnd, child_ptrs.as_slice(), None)
                .map_err(|e| e.to_string())?;
            if invoke_context_menu_verb(&cm, hwnd, "unpinfromhome").is_ok() {
                return Ok(true);
            }
            if invoke_context_menu_verb(&cm, hwnd, "removefromhome").is_ok() {
                return Ok(true);
            }
        }
        Ok(false)
    }
}

#[cfg(windows)]
fn jump_list_title_for_path(p: &str) -> String {
    let s = p.trim();
    if s.is_empty() {
        return String::new();
    }
    if s == "\\" || s == "/" {
        return s.to_string();
    }
    if s.len() == 3 && s.as_bytes()[1] == b':' && (s.as_bytes()[2] == b'\\' || s.as_bytes()[2] == b'/') {
        let letter = s.chars().next().unwrap_or('C').to_ascii_uppercase();
        return format!("{letter}:\\");
    }
    let trimmed = s.trim_end_matches(&['\\', '/'][..]);
    let part = trimmed
        .rsplit_once('\\')
        .map(|x| x.1)
        .or_else(|| trimmed.rsplit_once('/').map(|x| x.1))
        .unwrap_or(trimmed);
    if part.is_empty() {
        s.to_string()
    } else {
        part.to_string()
    }
}

#[cfg(windows)]
fn set_jump_list_windows(recent: Vec<String>) -> Result<bool, String> {
    let _com = com_init();
    let exe = std::env::current_exe().map_err(|e| e.to_string())?;
    let exe = exe.to_string_lossy().to_string();
    let exe_wide = to_wide_null(&exe);
    let cat_recent_wide = to_wide_null("最近");
    let cat_pinned_wide = to_wide_null("固定");

    unsafe {
        let icon_location_for_target = |target: &str| -> Option<(Vec<u16>, i32)> {
            let t = target.trim();
            if t.is_empty() {
                return None;
            }
            let mut info = SHFILEINFOW::default();
            if t.starts_with("shell:") || t.starts_with("::{") {
                let wide = to_wide_null(t);
                let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
                if SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None).is_err() || pidl.is_null() {
                    return None;
                }
                let pidl_guard = PidlGuard(pidl);
                let flags = SHGFI_ICONLOCATION | SHGFI_PIDL;
                let ret = SHGetFileInfoW(
                    PCWSTR(pidl_guard.0 as *const u16),
                    FILE_FLAGS_AND_ATTRIBUTES(0),
                    Some(&mut info),
                    std::mem::size_of::<SHFILEINFOW>() as u32,
                    flags,
                );
                if ret == 0 {
                    return None;
                }
                let loc = utf16_z_to_string(&info.szDisplayName);
                if loc.trim().is_empty() {
                    return None;
                }
                return Some((to_wide_null(&loc), info.iIcon));
            }

            let wide = to_wide_null(t);
            let flags = SHGFI_ICONLOCATION;
            let ret = SHGetFileInfoW(
                PCWSTR(wide.as_ptr()),
                FILE_FLAGS_AND_ATTRIBUTES(0),
                Some(&mut info),
                std::mem::size_of::<SHFILEINFOW>() as u32,
                flags,
            );
            if ret == 0 {
                return None;
            }
            let loc = utf16_z_to_string(&info.szDisplayName);
            if loc.trim().is_empty() {
                return None;
            }
            Some((to_wide_null(&loc), info.iIcon))
        };

        let list: ICustomDestinationList =
            CoCreateInstance(&DestinationList, None, CLSCTX_INPROC_SERVER).map_err(|e: windows::core::Error| e.to_string())?;
        let app_id = to_wide_null(win_app_id::app_user_model_id());
        let _ = list.SetAppID(PCWSTR(app_id.as_ptr()));
        let mut _max = 0u32;
        let _removed: IObjectArray = list.BeginList(&mut _max).map_err(|e: windows::core::Error| e.to_string())?;

        let build_link_for_path = |p: &str| -> Result<IShellLinkW, String> {
            let args = format!("--path \"{}\"", p.replace('"', ""));
            let args_wide = to_wide_null(&args);
            let link: IShellLinkW =
                CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER).map_err(|e: windows::core::Error| e.to_string())?;
            link.SetPath(PCWSTR(exe_wide.as_ptr())).map_err(|e: windows::core::Error| e.to_string())?;
            link.SetArguments(PCWSTR(args_wide.as_ptr()))
                .map_err(|e: windows::core::Error| e.to_string())?;

            {
                let store: IPropertyStore = link.cast().map_err(|e: windows::core::Error| e.to_string())?;
                let mut key = PROPERTYKEY::default();
                let key_name = to_wide_null("System.AppUserModel.ID");
                if PSGetPropertyKeyFromName(PCWSTR(key_name.as_ptr()), &mut key).is_ok() {
                    let id_wide = to_wide_null(win_app_id::app_user_model_id());
                    let bytes = id_wide.len().saturating_mul(2);
                    let mem = CoTaskMemAlloc(bytes);
                    if !mem.is_null() {
                        std::ptr::copy_nonoverlapping(id_wide.as_ptr(), mem as *mut u16, id_wide.len());
                        let mut pv = PROPVARIANT::default();
                        let pv0 = &mut *pv.Anonymous.Anonymous;
                        pv0.vt = std::mem::transmute(31u16);
                        pv0.Anonymous.pwszVal = PWSTR(mem as *mut u16);
                        let _ = store.SetValue(&key, &pv);
                        let _ = store.Commit();
                        let _ = PropVariantClear(&mut pv);
                    }
                }
            }

            if let Some((icon_file, icon_index)) = icon_location_for_target(p) {
                let _ = link.SetIconLocation(PCWSTR(icon_file.as_ptr()), icon_index);
            }

            let title = jump_list_title_for_path(p);
            if !title.is_empty() {
                let store: IPropertyStore = link.cast().map_err(|e: windows::core::Error| e.to_string())?;
                let mut key = PROPERTYKEY::default();
                let key_name = to_wide_null("System.Title");
                PSGetPropertyKeyFromName(PCWSTR(key_name.as_ptr()), &mut key)
                    .map_err(|e: windows::core::Error| e.to_string())?;

                let title_wide = to_wide_null(&title);
                let bytes = title_wide.len().saturating_mul(2);
                let mem = CoTaskMemAlloc(bytes);
                if !mem.is_null() {
                    std::ptr::copy_nonoverlapping(title_wide.as_ptr(), mem as *mut u16, title_wide.len());
                    let mut pv = PROPVARIANT::default();
                    let pv0 = &mut *pv.Anonymous.Anonymous;
                    pv0.vt = std::mem::transmute(31u16);
                    pv0.Anonymous.pwszVal = PWSTR(mem as *mut u16);
                    store.SetValue(&key, &pv).map_err(|e: windows::core::Error| e.to_string())?;
                    store.Commit().map_err(|e: windows::core::Error| e.to_string())?;
                    let _ = PropVariantClear(&mut pv);
                }
            }

            Ok(link)
        };

        let pinned_paths = match list_quick_access_windows() {
            Ok(v) => v,
            Err(_) => vec![],
        };
        let mut pinned_added = 0usize;
        let pinned_collection: IObjectCollection =
            CoCreateInstance(&EnumerableObjectCollection, None, CLSCTX_INPROC_SERVER).map_err(|e: windows::core::Error| e.to_string())?;
        let mut fallback_candidates: Vec<String> = Vec::new();
        for it in pinned_paths.into_iter() {
            if pinned_added >= 12 {
                break;
            }
            let p = it.path.trim().to_string();
            if p.is_empty() {
                continue;
            }
            if p.starts_with("shell:") {
                continue;
            }
            if it.pinned {
                let link = build_link_for_path(&p)?;
                pinned_collection
                    .AddObject(&link)
                    .map_err(|e: windows::core::Error| e.to_string())?;
                pinned_added += 1;
            } else if fallback_candidates.len() < 6 {
                fallback_candidates.push(p);
            }
        }
        if pinned_added == 0 && !fallback_candidates.is_empty() {
            for p in fallback_candidates.into_iter() {
                if pinned_added >= 12 {
                    break;
                }
                if !std::path::Path::new(&p).exists() {
                    continue;
                }
                let link = build_link_for_path(&p)?;
                pinned_collection
                    .AddObject(&link)
                    .map_err(|e: windows::core::Error| e.to_string())?;
                pinned_added += 1;
            }
        }
        if pinned_added > 0 {
            let array: IObjectArray = pinned_collection.cast().map_err(|e: windows::core::Error| e.to_string())?;
            list.AppendCategory(PCWSTR(cat_pinned_wide.as_ptr()), &array)
                .map_err(|e: windows::core::Error| e.to_string())?;
        }

        let collection: IObjectCollection =
            CoCreateInstance(&EnumerableObjectCollection, None, CLSCTX_INPROC_SERVER).map_err(|e: windows::core::Error| e.to_string())?;

        let mut added = 0usize;
        for raw in recent.into_iter() {
            if added >= 12 {
                break;
            }
            let p = raw.trim().to_string();
            if p.is_empty() {
                continue;
            }
            if p.starts_with("shell:") {
                continue;
            }
            if !std::path::Path::new(&p).exists() {
                continue;
            }

            let link = build_link_for_path(&p)?;
            collection.AddObject(&link).map_err(|e: windows::core::Error| e.to_string())?;
            added += 1;
        }

        let array: IObjectArray = collection.cast().map_err(|e: windows::core::Error| e.to_string())?;
        list.AppendCategory(PCWSTR(cat_recent_wide.as_ptr()), &array)
            .map_err(|e: windows::core::Error| e.to_string())?;
        list.CommitList().map_err(|e: windows::core::Error| e.to_string())?;
        Ok(true)
    }
}

#[cfg(windows)]
fn icon_png_base64_for_path(path: &str, size: Option<u32>) -> io::Result<Option<String>> {
    unsafe {
        let mut info = SHFILEINFOW::default();
        let wide = to_wide_null(path);
        let flags = SHGFI_ICON
            | if size.unwrap_or(16) <= 16 {
                SHGFI_SMALLICON
            } else {
                SHGFI_LARGEICON
            };

        let ret = SHGetFileInfoW(
            PCWSTR(wide.as_ptr()),
            FILE_FLAGS_AND_ATTRIBUTES(0),
            Some(&mut info),
            std::mem::size_of::<SHFILEINFOW>() as u32,
            flags,
        );
        if ret == 0 || info.hIcon.0 == std::ptr::null_mut() {
            return Ok(None);
        }

        return icon_png_base64_from_hicon(info.hIcon);
    }
}

#[cfg(windows)]
fn icon_png_base64_for_new_item(name: &str, is_dir: bool, size: Option<u32>) -> io::Result<Option<String>> {
    let mut safe = name.trim().to_string();
    if safe.is_empty() {
        safe = if is_dir { "folder".to_string() } else { "file".to_string() };
    }
    safe = safe
        .chars()
        .map(|c| match c {
            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' => '_',
            _ => c,
        })
        .collect();
    while safe.ends_with('.') || safe.ends_with(' ') {
        safe.pop();
    }
    if safe.is_empty() {
        safe = if is_dir { "folder".to_string() } else { "file".to_string() };
    }
    let dummy = if is_dir {
        "folder".to_string()
    } else {
        let idx = safe.rfind('.');
        if let Some(i) = idx {
            if i > 0 && i + 1 < safe.len() {
                let ext = safe[i + 1..].trim();
                if !ext.is_empty() {
                    format!("file.{ext}")
                } else {
                    "file".to_string()
                }
            } else {
                "file".to_string()
            }
        } else {
            "file".to_string()
        }
    };
    unsafe {
        let mut info = SHFILEINFOW::default();
        let wide = to_wide_null(&dummy);
        let flags = SHGFI_ICON
            | SHGFI_USEFILEATTRIBUTES
            | if size.unwrap_or(16) <= 16 {
                SHGFI_SMALLICON
            } else {
                SHGFI_LARGEICON
            };
        let attrs = if is_dir {
            windows::Win32::Storage::FileSystem::FILE_ATTRIBUTE_DIRECTORY
        } else {
            FILE_ATTRIBUTE_NORMAL
        };
        let ret = SHGetFileInfoW(
            PCWSTR(wide.as_ptr()),
            attrs,
            Some(&mut info),
            std::mem::size_of::<SHFILEINFOW>() as u32,
            flags,
        );
        if ret == 0 || info.hIcon.0 == std::ptr::null_mut() {
            return Ok(None);
        }
        icon_png_base64_from_hicon(info.hIcon)
    }
}

#[cfg(windows)]
fn icon_png_base64_from_hicon(hicon: HICON) -> io::Result<Option<String>> {
    unsafe {
        let mut iconinfo = ICONINFO::default();
        if GetIconInfo(hicon, &mut iconinfo).is_err() {
            let _ = DestroyIcon(hicon);
            return Ok(None);
        }

        let hbitmap: HBITMAP = iconinfo.hbmColor;
        if hbitmap.0.is_null() {
            if !iconinfo.hbmMask.0.is_null() {
                let _ = DeleteObject(HGDIOBJ(iconinfo.hbmMask.0));
            }
            let _ = DestroyIcon(hicon);
            return Ok(None);
        }

        let mut bmp: BITMAP = std::mem::zeroed();
        if GetObjectW(hbitmap.into(), std::mem::size_of::<BITMAP>() as i32, Some(&mut bmp as *mut _ as *mut _)) == 0 {
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            if !iconinfo.hbmMask.0.is_null() {
                let _ = DeleteObject(HGDIOBJ(iconinfo.hbmMask.0));
            }
            let _ = DestroyIcon(hicon);
            return Ok(None);
        }

        let width = bmp.bmWidth;
        let height = bmp.bmHeight;
        if width <= 0 || height <= 0 {
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            if !iconinfo.hbmMask.0.is_null() {
                let _ = DeleteObject(HGDIOBJ(iconinfo.hbmMask.0));
            }
            let _ = DestroyIcon(hicon);
            return Ok(None);
        }

        let mut bi = BITMAPINFO::default();
        bi.bmiHeader = BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: width,
            biHeight: -height,
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0 as u32,
            ..Default::default()
        };

        let dc = CreateCompatibleDC(None);
        if dc.0.is_null() {
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            if !iconinfo.hbmMask.0.is_null() {
                let _ = DeleteObject(HGDIOBJ(iconinfo.hbmMask.0));
            }
            let _ = DestroyIcon(hicon);
            return Ok(None);
        }

        let old = SelectObject(dc, HGDIOBJ(hbitmap.0));
        if old.0.is_null() {
            let _ = DeleteDC(dc);
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            if !iconinfo.hbmMask.0.is_null() {
                let _ = DeleteObject(HGDIOBJ(iconinfo.hbmMask.0));
            }
            let _ = DestroyIcon(hicon);
            return Ok(None);
        }

        let mut buf = vec![0u8; (width as usize) * (height as usize) * 4];
        let got = GetDIBits(
            dc,
            hbitmap,
            0,
            height as u32,
            Some(buf.as_mut_ptr() as *mut _),
            &mut bi,
            DIB_RGB_COLORS,
        );

        let _ = SelectObject(dc, old);
        let _ = DeleteDC(dc);
        let _ = DeleteObject(HGDIOBJ(hbitmap.0));
        if !iconinfo.hbmMask.0.is_null() {
            let _ = DeleteObject(HGDIOBJ(iconinfo.hbmMask.0));
        }
        let _ = DestroyIcon(hicon);

        if got == 0 {
            return Ok(None);
        }

        for px in buf.chunks_exact_mut(4) {
            let b = px[0];
            let g = px[1];
            let r = px[2];
            let a = px[3];
            px[0] = r;
            px[1] = g;
            px[2] = b;
            px[3] = a;
        }

        let img = image::RgbaImage::from_raw(width as u32, height as u32, buf)
            .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "icon buffer"))?;
        let mut png = Vec::new();
        img.write_to(&mut std::io::Cursor::new(&mut png), image::ImageFormat::Png)
            .map_err(|e| io::Error::new(io::ErrorKind::Other, e.to_string()))?;
        Ok(Some(base64::engine::general_purpose::STANDARD.encode(png)))
    }
}

#[cfg(windows)]
struct ComGuard(bool);

#[cfg(windows)]
impl Drop for ComGuard {
    fn drop(&mut self) {
        if self.0 {
            unsafe { CoUninitialize() };
        }
    }
}

#[cfg(windows)]
fn com_init() -> ComGuard {
    let ok = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_ok() };
    ComGuard(ok)
}

#[cfg(windows)]
struct PidlGuard(*mut ITEMIDLIST);

#[cfg(windows)]
impl Drop for PidlGuard {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe { CoTaskMemFree(Some(self.0 as _)) };
        }
    }
}

#[cfg(windows)]
fn strret_to_string(ret: &mut STRRET, pidl: *const ITEMIDLIST) -> io::Result<String> {
    unsafe {
        let mut buf = vec![0u16; 4096];
        StrRetToBufW(ret, Some(pidl), buf.as_mut_slice())
            .map_err(|e| io::Error::new(io::ErrorKind::Other, e.to_string()))?;
        Ok(String::from_utf16_lossy(&buf).trim_end_matches('\0').to_string())
    }
}

#[cfg(windows)]
fn icon_png_base64_for_shell_parsing_name(name: &str, size: Option<u32>) -> io::Result<Option<String>> {
    let _com = com_init();
    unsafe {
        let wide = to_wide_null(name);
        let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
        let hr = SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None);
        if hr.is_err() || pidl.is_null() {
            return Ok(None);
        }
        let _pidl_guard = PidlGuard(pidl);

        let mut info = SHFILEINFOW::default();
        let flags = SHGFI_ICON
            | SHGFI_PIDL
            | if size.unwrap_or(16) <= 16 {
                SHGFI_SMALLICON
            } else {
                SHGFI_LARGEICON
            };
        let ret = SHGetFileInfoW(
            PCWSTR(pidl as *const u16),
            FILE_FLAGS_AND_ATTRIBUTES(0),
            Some(&mut info),
            std::mem::size_of::<SHFILEINFOW>() as u32,
            flags,
        );
        if ret == 0 || info.hIcon.0 == std::ptr::null_mut() {
            return Ok(None);
        }
        icon_png_base64_from_hicon(info.hIcon)
    }
}

#[cfg(windows)]
fn icon_png_base64_for_any(path: &str, size: Option<u32>) -> io::Result<Option<String>> {
    if let Some(v) = icon_png_base64_for_path(path, size)? {
        return Ok(Some(v));
    }
    icon_png_base64_for_shell_parsing_name(path, size)
}

#[cfg(windows)]
fn icon_png_base64_for_any_jumbo(path: &str, size: u32) -> io::Result<Option<String>> {
    let _com = com_init();
    unsafe {
        let wide = to_wide_null(path);
        let mut info = SHFILEINFOW::default();
        let flags = SHGFI_SYSICONINDEX;
        let ret = SHGetFileInfoW(
            PCWSTR(wide.as_ptr()),
            FILE_FLAGS_AND_ATTRIBUTES(0),
            Some(&mut info),
            std::mem::size_of::<SHFILEINFOW>() as u32,
            flags,
        );
        if ret == 0 {
            let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
            let hr = SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None);
            if hr.is_ok() && !pidl.is_null() {
                let _pidl_guard = PidlGuard(pidl);
                let mut info_pidl = SHFILEINFOW::default();
                let flags_pidl = SHGFI_SYSICONINDEX | SHGFI_PIDL;
                let ret_pidl = SHGetFileInfoW(
                    PCWSTR(pidl as *const u16),
                    FILE_FLAGS_AND_ATTRIBUTES(0),
                    Some(&mut info_pidl),
                    std::mem::size_of::<SHFILEINFOW>() as u32,
                    flags_pidl,
                );
                if ret_pidl != 0 {
                    info = info_pidl;
                } else {
                    let mut info2 = SHFILEINFOW::default();
                    let flags2 = SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES;
                    let ret2 = SHGetFileInfoW(
                        PCWSTR(wide.as_ptr()),
                        FILE_ATTRIBUTE_NORMAL,
                        Some(&mut info2),
                        std::mem::size_of::<SHFILEINFOW>() as u32,
                        flags2,
                    );
                    if ret2 == 0 {
                        return Ok(None);
                    }
                    info = info2;
                }
            } else {
            let mut info2 = SHFILEINFOW::default();
            let flags2 = SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES;
            let ret2 = SHGetFileInfoW(
                PCWSTR(wide.as_ptr()),
                FILE_ATTRIBUTE_NORMAL,
                Some(&mut info2),
                std::mem::size_of::<SHFILEINFOW>() as u32,
                flags2,
            );
            if ret2 == 0 {
                return Ok(None);
            }
            info = info2;
            }
        }

        let list_kind: i32 = if size >= 96 { 4 } else if size >= 48 { 2 } else { 1 };
        let image_list: IImageList = match SHGetImageList(list_kind) {
            Ok(v) => v,
            Err(_) => return Ok(None),
        };
        let hicon = match image_list.GetIcon(info.iIcon, 1) {
            Ok(v) => v,
            Err(_) => return Ok(None),
        };
        if hicon.0.is_null() {
            return Ok(None);
        }
        icon_png_base64_from_hicon(hicon)
    }
}

#[cfg(windows)]
fn icon_png_base64_for_drag_target(path: &str, size: u32) -> io::Result<Option<String>> {
    if size >= 48 {
        if let Some(v) = icon_png_base64_for_any_jumbo(path, size)? {
            return Ok(Some(v));
        }
    }
    icon_png_base64_for_any(path, Some(size))
}

#[cfg(windows)]
fn filetime_to_unix_ms(ft: FILETIME) -> Option<u128> {
    let raw = ((ft.dwHighDateTime as u64) << 32) | (ft.dwLowDateTime as u64);
    if raw == 0 {
        return None;
    }
    let ms_since_1601 = raw / 10_000;
    let ms_since_1970 = ms_since_1601.checked_sub(11_644_473_600_000)?;
    Some(ms_since_1970 as u128)
}

#[cfg(windows)]
unsafe fn shell_item2_from_parsing_name(parsing_name: &str) -> io::Result<Option<IShellItem2>> {
    let wide = to_wide_null(parsing_name);
    let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
    let hr = SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None);
    if hr.is_err() || pidl.is_null() {
        return Ok(None);
    }
    let pidl_guard = PidlGuard(pidl);

    match SHCreateItemFromIDList::<IShellItem2>(pidl_guard.0) {
        Ok(v) => Ok(Some(v)),
        Err(_) => Ok(None),
    }
}

#[cfg(windows)]
unsafe fn shell_item2_get_string(item: &IShellItem2, key: *const PROPERTYKEY) -> Option<String> {
    let out: PWSTR = item.GetString(key).ok()?;
    if out.is_null() {
        return None;
    }
    let s = out.to_string().ok();
    CoTaskMemFree(Some(out.0 as _));
    s
}

#[cfg(windows)]
unsafe fn shell_item2_get_u64(item: &IShellItem2, key: *const PROPERTYKEY) -> Option<u64> {
    item.GetUInt64(key).ok()
}

#[cfg(windows)]
unsafe fn shell_item2_get_u32(item: &IShellItem2, key: *const PROPERTYKEY) -> Option<u32> {
    item.GetUInt32(key).ok()
}

#[cfg(windows)]
fn property_key_from_name(name: &str) -> Option<PROPERTYKEY> {
    unsafe {
        let wide = to_wide_null(name);
        let mut key = PROPERTYKEY::default();
        if PSGetPropertyKeyFromName(PCWSTR(wide.as_ptr()), &mut key).is_ok() {
            Some(key)
        } else {
            None
        }
    }
}

#[cfg(windows)]
unsafe fn shell_item2_get_filetime_ms(item: &IShellItem2, key: *const PROPERTYKEY) -> Option<u128> {
    let ft = item.GetFileTime(key).ok()?;
    filetime_to_unix_ms(ft)
}

#[cfg(windows)]
fn list_shell_folder_impl(shell_path: &str) -> io::Result<Vec<ShellEntryItem>> {
    let _com = com_init();
    unsafe {
        let is_recycle_bin = shell_path.eq_ignore_ascii_case("shell:RecycleBinFolder")
            || shell_path.to_ascii_lowercase().contains("recyclebinfolder");
        let desktop = SHGetDesktopFolder().map_err(|e| io::Error::new(io::ErrorKind::Other, e.to_string()))?;

        let wide = to_wide_null(shell_path);
        let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
        let hr = SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None);
        if hr.is_err() || pidl.is_null() {
            return Ok(vec![]);
        }
        let pidl_guard = PidlGuard(pidl);

        let folder: IShellFolder = match desktop.BindToObject::<_, IShellFolder>(pidl_guard.0, None) {
            Ok(v) => v,
            Err(_) => return Ok(vec![]),
        };

        let mut enum_list: Option<IEnumIDList> = None;
        let flags = (SHCONTF_FOLDERS.0 | SHCONTF_NONFOLDERS.0) as u32;
        let hr = folder.EnumObjects(HWND(std::ptr::null_mut()), flags, &mut enum_list);
        if hr.0 < 0 {
            return Ok(vec![]);
        }
        let enum_list = match enum_list {
            Some(x) => x,
            None => return Ok(vec![]),
        };

        let key_size = if is_recycle_bin {
            let wide = to_wide_null("System.Size");
            let mut key = PROPERTYKEY::default();
            let _ = PSGetPropertyKeyFromName(PCWSTR(wide.as_ptr()), &mut key);
            Some(key)
        } else {
            None
        };
        let key_modified = if is_recycle_bin {
            let wide = to_wide_null("System.DateModified");
            let mut key = PROPERTYKEY::default();
            let _ = PSGetPropertyKeyFromName(PCWSTR(wide.as_ptr()), &mut key);
            Some(key)
        } else {
            None
        };
        let key_deleted = if is_recycle_bin {
            let wide = to_wide_null("System.Recycle.DateDeleted");
            let mut key = PROPERTYKEY::default();
            let _ = PSGetPropertyKeyFromName(PCWSTR(wide.as_ptr()), &mut key);
            Some(key)
        } else {
            None
        };
        let key_deleted_from = if is_recycle_bin {
            let wide = to_wide_null("System.Recycle.DeletedFrom");
            let mut key = PROPERTYKEY::default();
            let _ = PSGetPropertyKeyFromName(PCWSTR(wide.as_ptr()), &mut key);
            Some(key)
        } else {
            None
        };
        let key_item_type = if is_recycle_bin {
            let wide = to_wide_null("System.ItemTypeText");
            let mut key = PROPERTYKEY::default();
            let _ = PSGetPropertyKeyFromName(PCWSTR(wide.as_ptr()), &mut key);
            Some(key)
        } else {
            None
        };
        let key_item_name_display = if is_recycle_bin {
            let wide = to_wide_null("System.ItemNameDisplay");
            let mut key = PROPERTYKEY::default();
            let _ = PSGetPropertyKeyFromName(PCWSTR(wide.as_ptr()), &mut key);
            Some(key)
        } else {
            None
        };

        let mut out = Vec::new();
        loop {
            let mut fetched: u32 = 0;
            let mut rgelt: [*mut ITEMIDLIST; 1] = [std::ptr::null_mut()];
            let hr = enum_list.Next(&mut rgelt, Some(&mut fetched as *mut u32));
            let child_pidl = rgelt[0];
            if hr.0 < 0 || fetched == 0 || child_pidl.is_null() {
                break;
            }
            let child_guard = PidlGuard(child_pidl);

            let is_dir = true;

            let mut disp: STRRET = std::mem::zeroed();
            if folder.GetDisplayNameOf(child_guard.0, SHGDN_NORMAL, &mut disp).is_err() {
                continue;
            }
            let mut name = strret_to_string(&mut disp, child_guard.0)?;
            if name.trim().is_empty() {
                continue;
            }

            let mut parsing: STRRET = std::mem::zeroed();
            if folder
                .GetDisplayNameOf(child_guard.0, SHGDN_FORPARSING, &mut parsing)
                .is_err()
            {
                continue;
            }
            let path = strret_to_string(&mut parsing, child_guard.0)?;
            if path.trim().is_empty() {
                continue;
            }

            let mut size: Option<u64> = None;
            let mut modified_ms: Option<u128> = None;
            let mut original_location: Option<String> = None;
            let mut deleted_ms: Option<u128> = None;
            let mut item_type: Option<String> = None;

            if is_recycle_bin {
                if let Some(item2) = shell_item2_from_parsing_name(&path)? {
                    if let Some(k) = key_item_name_display.as_ref() {
                        if let Some(v) = shell_item2_get_string(&item2, k) {
                            let v = v.trim().to_string();
                            if !v.is_empty() {
                                name = v;
                            }
                        }
                    }
                    if let Some(k) = key_size.as_ref() {
                        size = shell_item2_get_u64(&item2, k);
                    }
                    if let Some(k) = key_modified.as_ref() {
                        modified_ms = shell_item2_get_filetime_ms(&item2, k);
                    }
                    if let Some(k) = key_deleted_from.as_ref() {
                        original_location = shell_item2_get_string(&item2, k);
                    }
                    if let Some(k) = key_deleted.as_ref() {
                        deleted_ms = shell_item2_get_filetime_ms(&item2, k);
                    }
                    if let Some(k) = key_item_type.as_ref() {
                        item_type = shell_item2_get_string(&item2, k);
                    }
                }
            }

            out.push(ShellEntryItem {
                name,
                path,
                is_dir,
                size,
                modified_ms,
                original_location,
                deleted_ms,
                item_type,
            });
        }

        Ok(out)
    }
}

#[tauri::command]
fn list_roots_detailed() -> Result<Vec<RootItemDetailed>, String> {
    #[cfg(windows)]
    {
        let mut roots = Vec::new();
        for i in 0..26u8 {
            let letter = (b'A' + i) as char;
            let path = format!("{letter}:\\");
            if !std::path::Path::new(&path).exists() {
                continue;
            }
            let label = drive_label_for_root(&path);
            let icon_png_base64 = icon_png_base64_for_path(&path, Some(16)).ok().flatten();
            let (total_bytes, free_bytes) = drive_space_for_root(&path)
                .map(|(t, f)| (Some(t), Some(f)))
                .unwrap_or((None, None));
            roots.push(RootItemDetailed {
                path,
                label,
                icon_png_base64,
                total_bytes,
                free_bytes,
            });
        }
        return Ok(roots);
    }
    #[cfg(not(windows))]
    Ok(vec![RootItemDetailed {
        path: "/".to_string(),
        label: "/".to_string(),
        icon_png_base64: None,
        total_bytes: None,
        free_bytes: None,
    }])
}

#[tauri::command]
fn get_icon_png_base64(params: GetIconParams) -> Result<Option<String>, String> {
    #[cfg(windows)]
    {
        let size = params.size.unwrap_or(64);
        let path = params.path.trim().to_string();
        if is_dynamic_icon_path(&path) {
            let effective_size = size.max(128);
            if let Ok(Some(v)) = icon_png_base64_for_any_jumbo(&path, effective_size) {
                    return Ok(Some(v));
            }
            return icon_png_base64_for_any(&path, Some(effective_size)).map_err(|e| e.to_string());
        }
        let key = IconCacheKey {
            path,
            size,
        };
        if let Ok(mut cache) = icon_cache().lock() {
            if let Some(v) = cache.get(&key) {
                return Ok(v);
            }
        }
        let v = if size >= 48 {
            match icon_png_base64_for_any_jumbo(&key.path, size) {
                Ok(Some(v)) => Some(v),
                _ => icon_png_base64_for_any(&key.path, Some(size)).map_err(|e| e.to_string())?,
            }
        } else {
            icon_png_base64_for_any(&key.path, Some(size)).map_err(|e| e.to_string())?
        };
        if let Ok(mut cache) = icon_cache().lock() {
            cache.insert(key, v.clone());
        }
        return Ok(v);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(None)
    }
}

#[tauri::command]
fn get_stock_icon_png_base64(params: GetStockIconParams) -> Result<Option<String>, String> {
    #[cfg(windows)]
    unsafe {
        use windows::Win32::UI::Shell::{
            SHGetStockIconInfo, SHGSI_ICON, SHGSI_LARGEICON, SHGSI_SMALLICON, SHSTOCKICONID,
            SHSTOCKICONINFO,
        };

        let size = params.size.unwrap_or(16);
        let key = IconCacheKey {
            path: format!("__stock_icon__:{:08x}", params.id),
            size,
        };
        if let Ok(mut cache) = icon_cache().lock() {
            if let Some(v) = cache.get(&key) {
                return Ok(v);
            }
        }

        let mut sii = SHSTOCKICONINFO::default();
        sii.cbSize = std::mem::size_of::<SHSTOCKICONINFO>() as u32;
        let flags = SHGSI_ICON | if size <= 16 { SHGSI_SMALLICON } else { SHGSI_LARGEICON };
        let hr = SHGetStockIconInfo(SHSTOCKICONID(params.id as i32), flags, &mut sii);
        let v = if hr.is_err() || sii.hIcon.0.is_null() {
            None
        } else {
            icon_png_base64_from_hicon(sii.hIcon).map_err(|e| e.to_string())?
        };
        if let Ok(mut cache) = icon_cache().lock() {
            cache.insert(key, v.clone());
        }
        Ok(v)
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(None)
    }
}

#[tauri::command]
fn get_new_item_icon_png_base64(params: GetNewItemIconParams) -> Result<Option<String>, String> {
    #[cfg(windows)]
    {
        let size = params.size.unwrap_or(16);
        let name = params.name.trim().to_string();
        let is_dir = params.is_dir;
        let key = IconCacheKey {
            path: format!("__new_item_icon__:{}:{}", if is_dir { "dir" } else { "file" }, name.to_ascii_lowercase()),
            size,
        };
        if let Ok(mut cache) = icon_cache().lock() {
            if let Some(v) = cache.get(&key) {
                return Ok(v);
            }
        }
        let v = icon_png_base64_for_new_item(&name, is_dir, Some(size)).map_err(|e| e.to_string())?;
        if let Ok(mut cache) = icon_cache().lock() {
            cache.insert(key, v.clone());
        }
        Ok(v)
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(None)
    }
}

#[tauri::command]
fn confirm_message_box(params: ConfirmMessageBoxParams) -> Result<bool, String> {
    #[cfg(windows)]
    unsafe {
        let title_raw = params.title.trim().to_string();
        let message_raw = params.message.trim().to_string();
        let title = if title_raw.is_empty() { "确认".to_string() } else { title_raw };
        let message = if message_raw.is_empty() { "确定要继续吗？".to_string() } else { message_raw };
        let title_wide = to_wide_null(&title);
        let message_wide = to_wide_null(&message);
        let style = MB_OKCANCEL | MB_ICONWARNING | MB_DEFBUTTON2;
        let ret = MessageBoxW(None, PCWSTR(message_wide.as_ptr()), PCWSTR(title_wide.as_ptr()), style);
        return Ok(ret == IDOK);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
fn get_basic_file_info(params: GetBasicFileInfoParams) -> Result<BasicFileInfo, String> {
    #[cfg(windows)]
    unsafe {
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        let pb = std::path::PathBuf::from(&p);
        let is_dir = pb.is_dir();
        let size_bytes = if is_dir {
            None
        } else {
            std::fs::metadata(&pb).ok().map(|m| m.len())
        };
        let mut info = SHFILEINFOW::default();
        let wide = to_wide_null(&p);
        let flags = SHGFI_TYPENAME;
        let ret = SHGetFileInfoW(
            PCWSTR(wide.as_ptr()),
            FILE_FLAGS_AND_ATTRIBUTES(0),
            Some(&mut info),
            std::mem::size_of::<SHFILEINFOW>() as u32,
            flags,
        );
        let type_name = if ret == 0 {
            None
        } else {
            let raw = String::from_utf16_lossy(&info.szTypeName);
            let s = raw.trim_end_matches('\0').trim().to_string();
            if s.is_empty() { None } else { Some(s) }
        };
        Ok(BasicFileInfo {
            is_dir,
            type_name,
            size_bytes,
        })
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(BasicFileInfo {
            is_dir: false,
            type_name: None,
            size_bytes: None,
        })
    }
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ReadTextFileParams {
    path: String,
}

#[tauri::command]
fn read_text_file(params: ReadTextFileParams) -> Result<String, String> {
    let path = params.path.trim().to_string();
    if path.is_empty() {
        return Err("路径不能为空".to_string());
    }
    let pb = std::path::PathBuf::from(&path);
    if !pb.is_file() {
        return Err("不是有效的文件".to_string());
    }
    let content = std::fs::read_to_string(&pb).map_err(|e| e.to_string())?;
    Ok(content)
}

#[tauri::command]
fn confirm_task_dialog(params: ConfirmTaskDialogParams) -> Result<bool, String> {
    #[cfg(windows)]
    unsafe {
        let title = params.title.trim().to_string();
        let instruction = params.instruction.trim().to_string();
        let content = params.content.trim().to_string();
        let ok_text = params.ok_text.trim().to_string();
        let cancel_text = params.cancel_text.trim().to_string();

        let title_wide = to_wide_null(if title.is_empty() { "确认" } else { &title });
        let instruction_wide = to_wide_null(if instruction.is_empty() { "确认操作" } else { &instruction });
        let content_wide = to_wide_null(&content);
        let ok_wide = to_wide_null(if ok_text.is_empty() { "确定" } else { &ok_text });
        let cancel_wide = to_wide_null(if cancel_text.is_empty() { "取消" } else { &cancel_text });

        let mut hicon: Option<HICON> = None;
        if let Some(icon_path) = params.icon_path.as_ref() {
            let ip = icon_path.trim();
            if !ip.is_empty() {
                let mut info = SHFILEINFOW::default();
                let wide = to_wide_null(ip);
                let flags = SHGFI_ICON | SHGFI_LARGEICON;
                let ret = SHGetFileInfoW(
                    PCWSTR(wide.as_ptr()),
                    FILE_FLAGS_AND_ATTRIBUTES(0),
                    Some(&mut info),
                    std::mem::size_of::<SHFILEINFOW>() as u32,
                    flags,
                );
                if ret != 0 && !info.hIcon.0.is_null() {
                    hicon = Some(info.hIcon);
                }
            }
        }
        if hicon.is_none() {
            if let Some(icon_name) = params.icon_name.as_ref() {
                let n = icon_name.trim();
                if !n.is_empty() {
                    let is_dir = params.icon_is_dir.unwrap_or(false);
                    let mut safe = n.to_string();
                    safe = safe
                        .chars()
                        .map(|c| match c {
                            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' => '_',
                            _ => c,
                        })
                        .collect();
                    while safe.ends_with('.') || safe.ends_with(' ') {
                        safe.pop();
                    }
                    if safe.is_empty() {
                        safe = if is_dir { "folder".to_string() } else { "file.txt".to_string() };
                    }
                    let dummy = if is_dir {
                        "folder".to_string()
                    } else if safe.contains('.') {
                        safe
                    } else {
                        format!("{safe}.txt")
                    };
                    let mut info = SHFILEINFOW::default();
                    let wide = to_wide_null(&dummy);
                    let flags = SHGFI_ICON | SHGFI_USEFILEATTRIBUTES | SHGFI_LARGEICON;
                    let attrs = if is_dir {
                        windows::Win32::Storage::FileSystem::FILE_ATTRIBUTE_DIRECTORY
                    } else {
                        FILE_ATTRIBUTE_NORMAL
                    };
                    let ret = SHGetFileInfoW(
                        PCWSTR(wide.as_ptr()),
                        attrs,
                        Some(&mut info),
                        std::mem::size_of::<SHFILEINFOW>() as u32,
                        flags,
                    );
                    if ret != 0 && !info.hIcon.0.is_null() {
                        hicon = Some(info.hIcon);
                    }
                }
            }
        }

        let buttons = [
            TASKDIALOG_BUTTON {
                nButtonID: 1,
                pszButtonText: PCWSTR(ok_wide.as_ptr()),
            },
            TASKDIALOG_BUTTON {
                nButtonID: 2,
                pszButtonText: PCWSTR(cancel_wide.as_ptr()),
            },
        ];

        let mut cfg: TASKDIALOGCONFIG = std::mem::zeroed();
        cfg.cbSize = std::mem::size_of::<TASKDIALOGCONFIG>() as u32;
        cfg.pszWindowTitle = PCWSTR(title_wide.as_ptr());
        cfg.pszMainInstruction = PCWSTR(instruction_wide.as_ptr());
        cfg.pszContent = if content.is_empty() {
            PCWSTR::null()
        } else {
            PCWSTR(content_wide.as_ptr())
        };
        cfg.cButtons = buttons.len() as u32;
        cfg.pButtons = buttons.as_ptr();
        cfg.nDefaultButton = 2;
        cfg.dwFlags = if hicon.is_some() {
            TDF_ALLOW_DIALOG_CANCELLATION | TDF_USE_HICON_MAIN
        } else {
            TDF_ALLOW_DIALOG_CANCELLATION
        };
        cfg.cxWidth = params.width.unwrap_or(420);
        if let Some(icon) = hicon {
            cfg.Anonymous1.hMainIcon = icon;
        }

        let mut pressed: i32 = 0;
        TaskDialogIndirect(&cfg, Some(&mut pressed), None, None).map_err(|e| e.to_string())?;
        if let Some(icon) = hicon {
            let _ = DestroyIcon(icon);
        }
        Ok(pressed == 1)
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
fn get_icons_png_base64_batch(params: GetIconsBatchParams) -> Result<Vec<IconBatchItem>, String> {
    #[cfg(windows)]
    {
        let size = params.size.unwrap_or(64);
        let mut out: Vec<IconBatchItem> = Vec::new();
        for p in params.paths.into_iter().take(512) {
            let path = p.trim().to_string();
            if path.is_empty() {
                continue;
            }
            if is_dynamic_icon_path(&path) {
                let effective_size = size.max(128);
                let icon_png_base64 = match icon_png_base64_for_any_jumbo(&path, effective_size) {
                    Ok(Some(v)) => Some(v),
                    _ => icon_png_base64_for_any(&path, Some(effective_size)).ok().flatten(),
                };
                out.push(IconBatchItem { path, icon_png_base64 });
                continue;
            }
            let key = IconCacheKey {
                path: path.clone(),
                size,
            };

            if let Ok(mut cache) = icon_cache().lock() {
                if let Some(v) = cache.get(&key) {
                    out.push(IconBatchItem {
                        path,
                        icon_png_base64: v,
                    });
                    continue;
                }
            }

            let icon_png_base64 = if size >= 48 {
                match icon_png_base64_for_any_jumbo(&path, size) {
                    Ok(Some(v)) => Some(v),
                    _ => icon_png_base64_for_any(&path, Some(size)).ok().flatten(),
                }
            } else {
                icon_png_base64_for_any(&path, Some(size)).ok().flatten()
            };

            if let Ok(mut cache) = icon_cache().lock() {
                cache.insert(key, icon_png_base64.clone());
            }
            out.push(IconBatchItem {
                path,
                icon_png_base64,
            });
        }
        return Ok(out);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(vec![])
    }
}

#[derive(Debug, Deserialize)]
struct ListShellFolderParams {
    shell_path: String,
}

#[tauri::command]
async fn list_shell_folder(params: ListShellFolderParams) -> Result<Vec<ShellEntryItem>, String> {
    #[cfg(windows)]
    {
        let shell_path = params.shell_path;
        return tauri::async_runtime::spawn_blocking(move || {
            list_shell_folder_impl(&shell_path).map_err(|e| e.to_string())
        })
        .await
        .map_err(|e| e.to_string())?;
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(vec![])
    }
}

#[derive(Debug, Deserialize)]
struct GetThumbsBatchParams {
    paths: Vec<String>,
    #[serde(default)]
    size: Option<u32>,
}

#[derive(Debug, Serialize)]
struct ThumbBatchItem {
    path: String,
    thumb_png_base64: Option<String>,
}

#[derive(Debug, Deserialize)]
struct GetMediaMetadataParams {
    path: String,
}

#[derive(Debug, Serialize, Default, Clone)]
struct MediaMetadata {
    kind: String,
    width: Option<u32>,
    height: Option<u32>,
    duration_ms: Option<u64>,
    frame_rate: Option<f64>,
    video_bitrate: Option<u32>,
    audio_bitrate: Option<u32>,
    video_codec: Option<String>,
    audio_codec: Option<String>,
}

#[cfg(windows)]
fn thumb_png_base64_for_image(path: &str, size: u32) -> io::Result<Option<String>> {
    use image::ImageEncoder;
    let p = path.trim();
    if p.is_empty() {
        return Ok(None);
    }
    if is_shell_like_path(p) {
        return Ok(None);
    }
    let pb = PathBuf::from(p);
    if !pb.is_file() {
        return Ok(None);
    }
    let size = size.clamp(24, 1024);
    let img = match image::open(&pb) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };
    let resized = img.resize(size, size, image::imageops::FilterType::Triangle);
    let rgba = resized.to_rgba8();
    let w = rgba.width();
    let h = rgba.height();
    let mut png: Vec<u8> = Vec::new();
    {
        let mut cursor = std::io::Cursor::new(&mut png);
        let enc = image::codecs::png::PngEncoder::new(&mut cursor);
        enc.write_image(
            &rgba,
            w,
            h,
            image::ExtendedColorType::Rgba8,
        )
        .map_err(|e: image::ImageError| io::Error::new(io::ErrorKind::Other, e.to_string()))?;
    }
    let b64 = base64::engine::general_purpose::STANDARD.encode(png);
    Ok(Some(b64))
}

#[cfg(windows)]
fn png_base64_from_hbitmap(hbitmap: HBITMAP) -> io::Result<Option<String>> {
    unsafe {
        if hbitmap.0.is_null() {
            return Ok(None);
        }
        let mut bmp: BITMAP = std::mem::zeroed();
        if GetObjectW(
            hbitmap.into(),
            std::mem::size_of::<BITMAP>() as i32,
            Some(&mut bmp as *mut _ as *mut _),
        ) == 0
        {
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            return Ok(None);
        }

        let width = bmp.bmWidth;
        let height = bmp.bmHeight;
        if width <= 0 || height <= 0 {
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            return Ok(None);
        }

        let mut bi = BITMAPINFO::default();
        bi.bmiHeader = BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: width,
            biHeight: -height,
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0 as u32,
            ..Default::default()
        };

        let dc = CreateCompatibleDC(None);
        if dc.0.is_null() {
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            return Ok(None);
        }

        let old = SelectObject(dc, HGDIOBJ(hbitmap.0));
        if old.0.is_null() {
            let _ = DeleteDC(dc);
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
            return Ok(None);
        }

        let mut buf = vec![0u8; (width as usize) * (height as usize) * 4];
        let got = GetDIBits(
            dc,
            hbitmap,
            0,
            height as u32,
            Some(buf.as_mut_ptr() as *mut _),
            &mut bi,
            DIB_RGB_COLORS,
        );

        let _ = SelectObject(dc, old);
        let _ = DeleteDC(dc);
        let _ = DeleteObject(HGDIOBJ(hbitmap.0));

        if got == 0 {
            return Ok(None);
        }

        for px in buf.chunks_exact_mut(4) {
            let b = px[0];
            let g = px[1];
            let r = px[2];
            let a = px[3];
            px[0] = r;
            px[1] = g;
            px[2] = b;
            px[3] = a;
        }

        let img = image::RgbaImage::from_raw(width as u32, height as u32, buf)
            .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "hbitmap buffer"))?;
        let mut png = Vec::new();
        img.write_to(&mut std::io::Cursor::new(&mut png), image::ImageFormat::Png)
            .map_err(|e| io::Error::new(io::ErrorKind::Other, e.to_string()))?;
        Ok(Some(base64::engine::general_purpose::STANDARD.encode(png)))
    }
}

#[cfg(windows)]
fn thumb_png_base64_for_shell(path: &str, size: u32) -> io::Result<Option<String>> {
    let p = path.trim();
    if p.is_empty() {
        return Ok(None);
    }
    if is_shell_like_path(p) {
        return Ok(None);
    }
    let pb = PathBuf::from(p);
    if !pb.exists() {
        return Ok(None);
    }
    
    // 先检查缓存，避免重复生成
    let key = ThumbCacheKey {
        path: p.to_string(),
        size,
    };
    if let Ok(mut cache) = thumb_cache().lock() {
        if let Some(v) = cache.get(&key) {
            return Ok(v.clone());
        }
    }
    
    let size = size.clamp(24, 1024);
    let _com = com_init();
    unsafe {
        let item = match shell_item2_from_parsing_name(p)? {
            Some(v) => v,
            None => return Ok(None),
        };
        let factory: windows::Win32::UI::Shell::IShellItemImageFactory = match item.cast() {
            Ok(v) => v,
            Err(_) => return Ok(None),
        };
        let sz = windows::Win32::Foundation::SIZE {
            cx: size as i32,
            cy: size as i32,
        };
        // 使用 SIIGBF_THUMBNAILONLY 优先使用缩略图缓存，速度更快
        // 如果失败，回退到 SIIGBF_BIGGERSIZEOK
        let flags = windows::Win32::UI::Shell::SIIGBF_THUMBNAILONLY;
        let hbitmap = factory.GetImage(sz, flags).ok();
        let hbitmap = match hbitmap {
            Some(bmp) => Some(bmp),
            None => {
                // 回退到默认方式
                let flags = windows::Win32::UI::Shell::SIIGBF_BIGGERSIZEOK;
                factory.GetImage(sz, flags).ok()
            }
        };
        match hbitmap {
            Some(bmp) => {
                let result = png_base64_from_hbitmap(bmp)?;
                // 存入缓存
                if let Ok(mut cache) = thumb_cache().lock() {
                    cache.insert(key, result.clone());
                }
                Ok(result)
            }
            None => Ok(None),
        }
    }
}

#[tauri::command]
fn get_image_thumbs_png_base64_batch(params: GetThumbsBatchParams) -> Result<Vec<ThumbBatchItem>, String> {
    #[cfg(windows)]
    {
        let size = params.size.unwrap_or(256);
        let mut out: Vec<ThumbBatchItem> = Vec::new();
        for p in params.paths.into_iter().take(256) {
            let path = p.trim().to_string();
            if path.is_empty() {
                continue;
            }
            let key = ThumbCacheKey {
                path: path.clone(),
                size,
            };
            if let Ok(mut cache) = thumb_cache().lock() {
                if let Some(v) = cache.get(&key) {
                    out.push(ThumbBatchItem {
                        path,
                        thumb_png_base64: v,
                    });
                    continue;
                }
            }
            let v = thumb_png_base64_for_image(&path, size).ok().flatten();
            if let Ok(mut cache) = thumb_cache().lock() {
                cache.insert(key, v.clone());
            }
            out.push(ThumbBatchItem {
                path,
                thumb_png_base64: v,
            });
        }
        return Ok(out);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(vec![])
    }
}

#[tauri::command]
async fn get_shell_thumbs_png_base64_batch(params: GetThumbsBatchParams) -> Result<Vec<ThumbBatchItem>, String> {
    #[cfg(windows)]
    {
        use rayon::prelude::*;
        
        let size = params.size.unwrap_or(256);
        let paths: Vec<String> = params.paths.into_iter().take(64).collect();
        
        // 使用 spawn_blocking 在后台线程执行缩略图生成，避免阻塞主线程
        let result = tauri::async_runtime::spawn_blocking(move || {
            // 使用并行迭代器加速处理
            let items: Vec<ThumbBatchItem> = paths
                .par_iter()
                .filter_map(|path| {
                    let path = path.trim().to_string();
                    if path.is_empty() {
                        return None;
                    }
                    
                    // 先检查缓存
                    let key = ThumbCacheKey {
                        path: path.clone(),
                        size,
                    };
                    if let Ok(mut cache) = thumb_cache().lock() {
                        if let Some(v) = cache.get(&key) {
                            return Some(ThumbBatchItem {
                                path,
                                thumb_png_base64: v.clone(),
                            });
                        }
                    }
                    
                    // 生成缩略图（带超时）
                    let v = thumb_png_base64_for_shell(&path, size).ok().flatten();
                    if let Ok(mut cache) = thumb_cache().lock() {
                        cache.insert(key, v.clone());
                    }
                    
                    Some(ThumbBatchItem {
                        path,
                        thumb_png_base64: v,
                    })
                })
                .collect();
            
            items
        }).await.map_err(|e| e.to_string())?;
        
        return Ok(result);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(vec![])
    }
}

#[tauri::command]
fn get_media_metadata(params: GetMediaMetadataParams) -> Result<MediaMetadata, String> {
    #[cfg(windows)]
    {
        let path = params.path.trim().to_string();
        if path.is_empty() {
            return Ok(MediaMetadata::default());
        }
        if is_shell_like_path(&path) {
            return Ok(MediaMetadata::default());
        }
        let pb = PathBuf::from(&path);
        if !pb.is_file() {
            return Ok(MediaMetadata::default());
        }

        let _com = com_init();
        let mut out = MediaMetadata::default();

        unsafe {
            let item = match shell_item2_from_parsing_name(&path).map_err(|e| e.to_string())? {
                Some(v) => v,
                None => return Ok(out),
            };

            let img_w = property_key_from_name("System.Image.HorizontalSize")
                .and_then(|k| shell_item2_get_u32(&item, &k))
                .filter(|v| *v > 0);
            let img_h = property_key_from_name("System.Image.VerticalSize")
                .and_then(|k| shell_item2_get_u32(&item, &k))
                .filter(|v| *v > 0);

            let vid_w = property_key_from_name("System.Video.FrameWidth")
                .and_then(|k| shell_item2_get_u32(&item, &k))
                .filter(|v| *v > 0);
            let vid_h = property_key_from_name("System.Video.FrameHeight")
                .and_then(|k| shell_item2_get_u32(&item, &k))
                .filter(|v| *v > 0);

            let duration_100ns = property_key_from_name("System.Media.Duration")
                .and_then(|k| shell_item2_get_u64(&item, &k))
                .filter(|v| *v > 0);
            if let Some(v) = duration_100ns {
                out.duration_ms = Some(v / 10_000);
            }

            let frame_rate_raw = property_key_from_name("System.Video.FrameRate")
                .and_then(|k| shell_item2_get_u32(&item, &k))
                .filter(|v| *v > 0);
            if let Some(v) = frame_rate_raw {
                let fps = if v >= 1000 { (v as f64) / 1000.0 } else { v as f64 };
                out.frame_rate = Some(fps);
            }

            out.video_bitrate = property_key_from_name("System.Video.EncodingBitrate")
                .and_then(|k| shell_item2_get_u32(&item, &k))
                .filter(|v| *v > 0);
            out.audio_bitrate = property_key_from_name("System.Audio.EncodingBitrate")
                .and_then(|k| shell_item2_get_u32(&item, &k))
                .filter(|v| *v > 0);

            out.video_codec = property_key_from_name("System.Video.Compression")
                .and_then(|k| shell_item2_get_string(&item, &k))
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty());
            out.audio_codec = property_key_from_name("System.Audio.Format")
                .and_then(|k| shell_item2_get_string(&item, &k))
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty());

            let has_video = out.duration_ms.is_some()
                || vid_w.is_some()
                || vid_h.is_some()
                || out.frame_rate.is_some()
                || out.video_bitrate.is_some()
                || out.audio_bitrate.is_some()
                || out.video_codec.is_some()
                || out.audio_codec.is_some();

            if has_video {
                out.kind = "video".to_string();
                out.width = vid_w.or(img_w);
                out.height = vid_h.or(img_h);
            } else if img_w.is_some() || img_h.is_some() {
                out.kind = "image".to_string();
                out.width = img_w;
                out.height = img_h;
            }
        }

        if out.kind != "video" && (out.width.is_none() || out.height.is_none()) {
            if let Ok(reader) = image::ImageReader::open(&pb).and_then(|r| r.with_guessed_format()) {
                if let Ok((w, h)) = reader.into_dimensions() {
                    if w > 0 && h > 0 {
                        out.kind = "image".to_string();
                        out.width = Some(w);
                        out.height = Some(h);
                    }
                }
            }
        }

        Ok(out)
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(MediaMetadata::default())
    }
}

#[derive(Debug, Deserialize)]
struct ScanGalleryImagesParams {
    request_id: u64,
    #[serde(default)]
    max_items: Option<u64>,
}

#[cfg(windows)]
fn known_folder_path(folder_id: &windows::core::GUID) -> Option<String> {
    unsafe {
        use windows::Win32::UI::Shell::{SHGetKnownFolderPath, KF_FLAG_DEFAULT};
        let out: PWSTR = SHGetKnownFolderPath(folder_id, KF_FLAG_DEFAULT, None).ok()?;
        if out.is_null() {
            return None;
        }
        let s = out.to_string().ok();
        CoTaskMemFree(Some(out.0 as _));
        s
    }
}

#[cfg(windows)]
fn default_gallery_sources() -> Vec<String> {
    use windows::Win32::UI::Shell::{
        FOLDERID_CameraRoll, FOLDERID_Pictures, FOLDERID_SavedPictures, FOLDERID_Screenshots,
    };
    let mut list = Vec::new();
    for id in [&FOLDERID_Pictures, &FOLDERID_CameraRoll, &FOLDERID_Screenshots, &FOLDERID_SavedPictures] {
        if let Some(p) = known_folder_path(id) {
            let t = p.trim().to_string();
            if !t.is_empty() && PathBuf::from(&t).is_dir() {
                list.push(t);
            }
        }
    }
    list.sort_by(|a, b| normalize_compare_path_win(a).cmp(&normalize_compare_path_win(b)));
    list.dedup_by(|a, b| normalize_compare_path_win(a) == normalize_compare_path_win(b));
    list
}

#[cfg(windows)]
fn is_dynamic_icon_path(path: &str) -> bool {
    let p = path.trim();
    if p.is_empty() {
        return false;
    }
    let low = p.to_ascii_lowercase();
    low == "shell:recyclebinfolder"
        || low.contains("recyclebinfolder")
        || low.contains("645ff040-5081-101b-9f08-00aa002f954e")
}

#[cfg(windows)]
fn is_image_ext(ext: &str) -> bool {
    matches!(
        ext,
        "jpg" | "jpeg" | "png" | "gif" | "bmp" | "webp" | "tif" | "tiff"
    )
}

#[tauri::command]
fn scan_gallery_images(window: tauri::Window, params: ScanGalleryImagesParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        tauri::async_runtime::spawn_blocking(move || -> Result<bool, String> {
            let window_label = window.label().to_string();
            let cancel_key = format!("{window_label}::gallery");
            if let Ok(mut map) = gallery_latest_map().lock() {
                map.insert(cancel_key.clone(), params.request_id);
            }

            #[derive(Serialize, Clone)]
            struct GalleryProgressPayload {
                request_id: u64,
                scanned: u64,
                found: u64,
                done: bool,
            }

            #[derive(Serialize, Clone)]
            struct GalleryBatchPayload {
                request_id: u64,
                items: Vec<DirEntryItem>,
            }

            let max_items = params.max_items.unwrap_or(8000).clamp(1, 20000) as u64;
            let sources = default_gallery_sources();
            let mut seen: HashSet<String> = HashSet::new();
            let mut batch: Vec<DirEntryItem> = Vec::new();
            let mut scanned: u64 = 0;
            let mut found: u64 = 0;
            let mut last_emit = std::time::Instant::now();

            for src in sources {
                let walker = WalkDir::new(src)
                    .follow_links(false)
                    .into_iter()
                    .filter_map(|e| e.ok());
                for entry in walker {
                    scanned = scanned.saturating_add(1);
                    if scanned % 256 == 0 {
                        let latest = gallery_latest_map()
                            .lock()
                            .ok()
                            .and_then(|m| m.get(&cancel_key).copied())
                            .unwrap_or(params.request_id);
                        if latest != params.request_id {
                            let _ = window.emit(
                                "gallery_scan_progress",
                                GalleryProgressPayload {
                                    request_id: params.request_id,
                                    scanned,
                                    found,
                                    done: true,
                                },
                            );
                            if let Ok(mut map) = gallery_latest_map().lock() {
                                if map.get(&cancel_key).copied() == Some(params.request_id) {
                                    map.remove(&cancel_key);
                                }
                            }
                            return Ok(true);
                        }
                    }

                    if last_emit.elapsed() >= std::time::Duration::from_millis(120) {
                        let _ = window.emit(
                            "gallery_scan_progress",
                            GalleryProgressPayload {
                                request_id: params.request_id,
                                scanned,
                                found,
                                done: false,
                            },
                        );
                        last_emit = std::time::Instant::now();
                    }

                    if !entry.file_type().is_file() {
                        continue;
                    }
                    let ext = entry
                        .path()
                        .extension()
                        .and_then(|s| s.to_str())
                        .map(|s| s.to_ascii_lowercase())
                        .unwrap_or_default();
                    if !is_image_ext(&ext) {
                        continue;
                    }

                    let path = entry.path().to_string_lossy().to_string();
                    if path.trim().is_empty() {
                        continue;
                    }
                    let key = normalize_compare_path_win(&path);
                    if !seen.insert(key) {
                        continue;
                    }

                    let name = entry.file_name().to_string_lossy().to_string();
                    let metadata = entry.metadata().ok();
                    let size = metadata.as_ref().map(|m| m.len());
                    let modified_ms = metadata
                        .as_ref()
                        .and_then(|m| m.modified().ok())
                        .and_then(system_time_to_ms);

                    found = found.saturating_add(1);
                    batch.push(DirEntryItem {
                        name,
                        path,
                        is_dir: false,
                        size: Some(size.unwrap_or(0)),
                        modified_ms,
                    });

                    if batch.len() >= 96 {
                        batch.sort_by(|a, b| b.modified_ms.unwrap_or(0).cmp(&a.modified_ms.unwrap_or(0)));
                        let payload = GalleryBatchPayload {
                            request_id: params.request_id,
                            items: batch.drain(..).collect(),
                        };
                        let _ = window.emit("gallery_scan_batch", payload);
                    }

                    if found >= max_items {
                        break;
                    }
                }
                if found >= max_items {
                    break;
                }
            }

            if !batch.is_empty() {
                batch.sort_by(|a, b| b.modified_ms.unwrap_or(0).cmp(&a.modified_ms.unwrap_or(0)));
                let _ = window.emit(
                    "gallery_scan_batch",
                    GalleryBatchPayload {
                        request_id: params.request_id,
                        items: batch,
                    },
                );
            }
            let _ = window.emit(
                "gallery_scan_progress",
                GalleryProgressPayload {
                    request_id: params.request_id,
                    scanned,
                    found,
                    done: true,
                },
            );
            if let Ok(mut map) = gallery_latest_map().lock() {
                if map.get(&cancel_key).copied() == Some(params.request_id) {
                    map.remove(&cancel_key);
                }
            }
            Ok(true)
        });
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
fn open_path(path: String) -> Result<bool, String> {
    #[cfg(windows)]
    {
        use std::process::Command;
        Command::new("explorer")
            .arg(&path)
            .spawn()
            .map_err(|e| e.to_string())?;
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = path;
        Ok(false)
    }
}

#[tauri::command]
fn get_drag_icon_path() -> Result<String, String> {
    let mut p = std::env::temp_dir();
    p.push("filemgr_drag_icon_256.png");
    if !p.exists() {
        fs::write(&p, include_bytes!("../../F.png")).map_err(|e| e.to_string())?;
    }
    Ok(p.to_string_lossy().to_string())
}

#[tauri::command]
fn get_drag_icon_path_for_target(params: GetIconParams) -> Result<String, String> {
    #[cfg(windows)]
    {
        let target = params.path.trim().to_string();
        if target.is_empty() {
            return get_drag_icon_path();
        }

        let size = params.size.unwrap_or(48).clamp(16, 512);
        let mut preferred_png_base64: Option<String> = None;
        {
            let pb = std::path::PathBuf::from(&target);
            if pb.is_file() {
                let ext = pb
                    .extension()
                    .and_then(|s| s.to_str())
                    .unwrap_or("")
                    .to_ascii_lowercase();
                if !ext.is_empty() && is_image_ext(&ext) {
                    preferred_png_base64 = thumb_png_base64_for_image(&target, size).ok().flatten();
                }
            }
        }

        let icon_png_base64 = preferred_png_base64.or_else(|| icon_png_base64_for_drag_target(&target, size).ok().flatten());
        if let Some(b64) = icon_png_base64 {
            let mut png = base64::engine::general_purpose::STANDARD
                .decode(b64)
                .map_err(|e| e.to_string())?;
            if size >= 16 {
                if let Ok(img) = image::load_from_memory(&png) {
                    if img.width() > size || img.height() > size {
                        let resized = img.resize_exact(size, size, image::imageops::FilterType::Lanczos3);
                        let mut out_png = Vec::new();
                        resized
                            .write_to(&mut std::io::Cursor::new(&mut out_png), image::ImageFormat::Png)
                            .map_err(|e| e.to_string())?;
                        png = out_png;
                    }
                }
            }

            let mut hasher = std::collections::hash_map::DefaultHasher::new();
            std::hash::Hasher::write(&mut hasher, target.as_bytes());
            std::hash::Hasher::write_u32(&mut hasher, size);
            if let Ok(meta) = std::fs::metadata(&target) {
                if let Ok(modified) = meta.modified() {
                    if let Ok(d) = modified.duration_since(std::time::UNIX_EPOCH) {
                        std::hash::Hasher::write_u64(&mut hasher, d.as_secs());
                        std::hash::Hasher::write_u32(&mut hasher, d.subsec_nanos());
                    }
                }
            }
            let h = std::hash::Hasher::finish(&hasher);

            let mut out = std::env::temp_dir();
            out.push(format!("filemgr_drag_icon_{h}_{size}.png"));
            if !out.exists() {
                fs::write(&out, png).map_err(|e| e.to_string())?;
            }
            return Ok(out.to_string_lossy().to_string());
        }

        get_drag_icon_path()
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        get_drag_icon_path()
    }
}

#[derive(Debug, Deserialize)]
struct OpenInTerminalParams {
    path: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct NativeContextMenuItem {
    kind: String,
    #[serde(default)]
    id: Option<String>,
    #[serde(default)]
    label: Option<String>,
    #[serde(default)]
    enabled: bool,
    #[serde(default)]
    glyph: Option<String>,
    #[serde(default)]
    children: Option<Vec<NativeContextMenuItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct ShowNativeContextMenuParams {
    x: i32,
    y: i32,
    #[serde(default)]
    owner_hwnd: Option<u64>,
    #[serde(default)]
    theme_color: Option<String>, // New field for frontend theme color
    #[serde(default)]
    items: Vec<NativeContextMenuItem>,
}

#[derive(Debug, Deserialize)]
struct NativeContextMenuResponse {
    #[serde(rename = "selectedId", default)]
    selected_id: Option<String>,
    #[serde(default)]
    error: Option<String>,
}

#[cfg(windows)]
fn resolve_native_menu_sidecar_exe() -> Option<PathBuf> {
    let exe_name = "filemgr-native-menu.exe";
    if let Ok(cwd) = std::env::current_dir() {
        let p = cwd.join("bin").join(exe_name);
        if p.exists() {
            return Some(p);
        }
        let p2 = cwd.join("src-tauri").join("bin").join(exe_name);
        if p2.exists() {
            return Some(p2);
        }
    }
    if let Ok(cur) = std::env::current_exe() {
        if let Some(dir) = cur.parent() {
            let p = dir.join(exe_name);
            if p.exists() {
                return Some(p);
            }
            let p2 = dir.join("..").join("Resources").join(exe_name);
            if p2.exists() {
                return Some(p2);
            }
            let p3 = dir.join("resources").join(exe_name);
            if p3.exists() {
                return Some(p3);
            }
        }
    }
    None
}

#[cfg(windows)]
unsafe fn create_menu_glyph_bitmap(glyph: &str, override_color: Option<u32>) -> Option<windows::Win32::Graphics::Gdi::HBITMAP> {
    use windows::Win32::Foundation::RECT;
    use windows::Win32::Graphics::Gdi::{
        CreateCompatibleDC, CreateDIBSection, CreateFontW, DeleteDC, DeleteObject, DrawTextW, SelectObject, SetBkMode,
        SetTextColor, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET, DEFAULT_PITCH, DEFAULT_QUALITY,
        DIB_RGB_COLORS, DRAW_TEXT_FORMAT, DT_CENTER, DT_SINGLELINE, DT_VCENTER, FF_DONTCARE, FW_NORMAL, HBITMAP, HGDIOBJ,
        OUT_DEFAULT_PRECIS, HFONT, TRANSPARENT,
    };
    use windows::core::PCWSTR;

    let g = glyph.trim();
    if g.is_empty() {
        return None;
    }

    // Increase size to 24x24 for better visibility
    let size = 24;
    let font_size = -18; // Larger font

    let mut bmi = BITMAPINFO::default();
    bmi.bmiHeader = BITMAPINFOHEADER {
        biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
        biWidth: size,
        biHeight: -size,
        biPlanes: 1,
        biBitCount: 32,
        biCompression: BI_RGB.0,
        ..Default::default()
    };
    let mut bits: *mut core::ffi::c_void = std::ptr::null_mut();
    let bmp: HBITMAP = match CreateDIBSection(None, &bmi, DIB_RGB_COLORS, &mut bits, None, 0) {
        Ok(v) => v,
        Err(_) => return None,
    };
    if bmp.0.is_null() || bits.is_null() {
        return None;
    }
    std::ptr::write_bytes(bits, 0, (size * size * 4) as usize);

    let dc = CreateCompatibleDC(None);
    if dc.0.is_null() {
        let _ = DeleteObject(HGDIOBJ(bmp.0));
        return None;
    }
    let old_bmp = SelectObject(dc, HGDIOBJ(bmp.0));
    let _ = SetBkMode(dc, TRANSPARENT);

    // Draw white text, then use the brightness as alpha channel
    let _ = SetTextColor(dc, windows::Win32::Foundation::COLORREF(0x00FFFFFF));

    let face = to_wide_null("Segoe MDL2 Assets");
    let font: HFONT = CreateFontW(
        font_size,
        0,
        0,
        0,
        FW_NORMAL.0 as i32,
        0,
        0,
        0,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        DEFAULT_QUALITY,
        u32::from(DEFAULT_PITCH.0 | FF_DONTCARE.0),
        PCWSTR(face.as_ptr()),
    );
    let mut old_font: HGDIOBJ = HGDIOBJ(std::ptr::null_mut());
    if !font.0.is_null() {
        old_font = SelectObject(dc, HGDIOBJ(font.0));
    }

    let mut wide = to_wide_null(g);
    if !wide.is_empty() {
        wide.pop();
    }
    let mut rc = RECT {
        left: 0,
        top: 0,
        right: size,
        bottom: size,
    };
    let _ = DrawTextW(dc, &mut wide, &mut rc, DRAW_TEXT_FORMAT((DT_CENTER | DT_VCENTER | DT_SINGLELINE).0));

    let pixels = bits as *mut u32;
    
    // Determine target color (0x00RRGGBB)
    let target_rgb = if let Some(c) = override_color {
        c
    } else {
        // Use manual link for GetSysColor since windows crate import is tricky
        #[link(name = "user32")]
        extern "system" {
            fn GetSysColor(nIndex: i32) -> u32;
        }
        let sys_color = GetSysColor(7); // COLOR_MENUTEXT = 7
        // Convert to 0x00RRGGBB for our manual pixel manipulation
        let r = (sys_color & 0xFF) as u32;
        let g = ((sys_color >> 8) & 0xFF) as u32;
        let b = ((sys_color >> 16) & 0xFF) as u32;
        (r << 16) | (g << 8) | b
    };

    for i in 0..((size * size) as usize) {
        let px = *pixels.add(i);
        // Extract alpha from R channel (since we drew white)
        let alpha = (px & 0xFF) as u32;
        if alpha > 0 {
             // Apply alpha to target color
             *pixels.add(i) = (alpha << 24) | target_rgb;
        } else {
             *pixels.add(i) = 0;
        }
    }

    if !old_font.0.is_null() {
        let _ = SelectObject(dc, old_font);
    }
    if !font.0.is_null() {
        let _ = DeleteObject(HGDIOBJ(font.0));
    }
    if !old_bmp.0.is_null() {
        let _ = SelectObject(dc, old_bmp);
    }
    let _ = DeleteDC(dc);

    Some(bmp)
}

#[cfg(windows)]
fn get_theme_accent_color() -> u32 {
    // Try DWM colorization color first (Accent Color)
    // DwmGetColorizationColor(pcrColorization, pfOpaqueBlend)
    use windows::Win32::Graphics::Dwm::DwmGetColorizationColor;
    use windows::core::BOOL;
    
    let mut color = 0u32;
    let mut opaque = BOOL(0);
    unsafe {
        if DwmGetColorizationColor(&mut color, &mut opaque).is_ok() {
            // DwmGetColorizationColor returns 0xAARRGGBB
            // We want 0x00RRGGBB
            return color & 0x00FFFFFF;
        }
    }
    
    // Fallback to COLOR_HIGHLIGHT (System Highlight Color)
    #[link(name = "user32")]
    extern "system" {
        fn GetSysColor(nIndex: i32) -> u32;
    }
    // COLOR_HIGHLIGHT = 13
    unsafe {
        let c = GetSysColor(13);
        // GetSysColor returns 0x00BBGGRR, we need 0x00RRGGBB
        let r = (c & 0xFF) as u32;
        let g = ((c >> 8) & 0xFF) as u32;
        let b = ((c >> 16) & 0xFF) as u32;
        (r << 16) | (g << 8) | b
    }
}

#[cfg(windows)]
fn parse_hex_color(hex: &str) -> Option<u32> {
    let s = hex.trim().trim_start_matches('#');
    if s.len() == 6 {
        if let Ok(val) = u32::from_str_radix(s, 16) {
            // hex is RRGGBB, we want 0x00RRGGBB
            return Some(val & 0x00FFFFFF);
        }
    }
    None
}

#[cfg(windows)]
unsafe fn append_custom_native_menu_items(
    menu: windows::Win32::UI::WindowsAndMessaging::HMENU,
    items: &[NativeContextMenuItem],
    next_cmd_id: &mut u32,
    id_map: &mut HashMap<u32, String>,
    bitmaps: &mut Vec<windows::Win32::Graphics::Gdi::HBITMAP>,
    theme_color_override: Option<u32>, // New parameter
) {
    use windows::core::PCWSTR;
    use windows::Win32::UI::WindowsAndMessaging::{
        AppendMenuW, CreatePopupMenu, SetMenuInfo, SetMenuItemInfoW, MENUINFO, MENUITEMINFOW, MIM_STYLE, MNS_CHECKORBMP, MF_GRAYED,
        MF_POPUP, MF_SEPARATOR, MF_STRING, MIIM_BITMAP,
    };

    for it in items {
        let kind = it.kind.trim().to_ascii_lowercase();
        if kind == "sep" || kind == "separator" {
            let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());
            continue;
        }
        if kind == "submenu" {
            if let Ok(child) = CreatePopupMenu() {
                let mut mi: MENUINFO = std::mem::zeroed();
                mi.cbSize = std::mem::size_of::<MENUINFO>() as u32;
                mi.fMask = MIM_STYLE;
                mi.dwStyle = MNS_CHECKORBMP;
                let _ = SetMenuInfo(child, &mut mi);

                let children = it.children.clone().unwrap_or_default();
                append_custom_native_menu_items(child, &children, next_cmd_id, id_map, bitmaps, theme_color_override);
                let mut flags = MF_POPUP;
                if !it.enabled {
                    flags |= MF_GRAYED;
                }
                let label = it.label.clone().unwrap_or_default();
                let wide = to_wide_null(label.trim());
                let _ = AppendMenuW(menu, flags, child.0 as usize, PCWSTR(wide.as_ptr()));
            }
            continue;
        }
        if kind != "item" {
            continue;
        }

        let id = it.id.clone().unwrap_or_default().trim().to_string();
        let label = it.label.clone().unwrap_or_default().trim().to_string();
        if id.is_empty() || label.is_empty() {
            continue;
        }
        let cmd_id = *next_cmd_id;
        *next_cmd_id = (*next_cmd_id).saturating_add(1);
        id_map.insert(cmd_id, id.clone());

        let mut flags = MF_STRING;
        if !it.enabled {
            flags |= MF_GRAYED;
        }
        let wide = to_wide_null(&label);
        let _ = AppendMenuW(menu, flags, cmd_id as usize, PCWSTR(wide.as_ptr()));

        let glyph = it.glyph.clone().unwrap_or_default();
        let color = if theme_color_override.is_some() {
            theme_color_override
        } else {
            Some(get_theme_accent_color())
        };
        if let Some(bmp) = create_menu_glyph_bitmap(&glyph, color) {
            bitmaps.push(bmp);
            let mut info: MENUITEMINFOW = std::mem::zeroed();
            info.cbSize = std::mem::size_of::<MENUITEMINFOW>() as u32;
            info.fMask = MIIM_BITMAP;
            info.hbmpItem = bmp;
            let _ = SetMenuItemInfoW(menu, cmd_id, false, &mut info);
        }
    }
}

#[cfg(windows)]
struct MenuWin11StyleHookGuard(windows::Win32::UI::Accessibility::HWINEVENTHOOK);

#[cfg(windows)]
impl Drop for MenuWin11StyleHookGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = windows::Win32::UI::Accessibility::UnhookWinEvent(self.0);
        }
    }
}

#[cfg(windows)]
extern "system" fn menu_win_event_proc(
    _hook: windows::Win32::UI::Accessibility::HWINEVENTHOOK,
    _event: u32,
    hwnd: windows::Win32::Foundation::HWND,
    _id_object: i32,
    _id_child: i32,
    _id_event_thread: u32,
    _dwms_event_time: u32,
) {
    unsafe {
        if hwnd.0.is_null() {
            return;
        }

        let mut pid: u32 = 0;
        let _ = windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId(hwnd, Some(&mut pid));
        if pid == 0 || pid != windows::Win32::System::Threading::GetCurrentProcessId() {
            return;
        }

        let mut cls: [u16; 64] = [0; 64];
        let n = windows::Win32::UI::WindowsAndMessaging::GetClassNameW(hwnd, &mut cls);
        if n == 0 {
            return;
        }
        let name = String::from_utf16_lossy(&cls[..(n as usize)]).trim().to_string();
        if name != "#32768" {
            return;
        }

        let corner: i32 = 2;
        let _ = windows::Win32::Graphics::Dwm::DwmSetWindowAttribute(
            hwnd,
            windows::Win32::Graphics::Dwm::DWMWINDOWATTRIBUTE(33),
            &corner as *const _ as _,
            std::mem::size_of::<i32>() as u32,
        );

        let backdrop: i32 = 3;
        let _ = windows::Win32::Graphics::Dwm::DwmSetWindowAttribute(
            hwnd,
            windows::Win32::Graphics::Dwm::DWMWINDOWATTRIBUTE(38),
            &backdrop as *const _ as _,
            std::mem::size_of::<i32>() as u32,
        );

        #[link(name = "user32")]
        extern "system" {
            fn RedrawWindow(
                hwnd: windows::Win32::Foundation::HWND,
                lprcupdate: *const windows::Win32::Foundation::RECT,
                hrgnupdate: windows::Win32::Graphics::Gdi::HRGN,
                flags: u32,
            ) -> i32;
        }
        let flags = 0x0001u32 | 0x0004u32 | 0x0100u32 | 0x0400u32;
        let _ = RedrawWindow(
            hwnd,
            std::ptr::null(),
            windows::Win32::Graphics::Gdi::HRGN(std::ptr::null_mut()),
            flags,
        );
    }
}

#[cfg(windows)]
fn install_menu_win11_style_hook() -> Option<MenuWin11StyleHookGuard> {
    unsafe {
        use windows::Win32::UI::Accessibility::SetWinEventHook;
        let event_object_show: u32 = 0x8002;
        let hook = SetWinEventHook(
            event_object_show,
            event_object_show,
            None,
            Some(menu_win_event_proc),
            0,
            0,
            0,
        );
        if hook.0.is_null() {
            return None;
        }
        Some(MenuWin11StyleHookGuard(hook))
    }
}

#[tauri::command]
fn show_native_context_menu(_window: tauri::Window, params: ShowNativeContextMenuParams) -> Result<Option<String>, String> {
    #[cfg(windows)]
    {
        use std::time::Instant;
        use windows::Win32::Foundation::POINT;
        use windows::Win32::Graphics::Gdi::{DeleteObject, HGDIOBJ};
        use windows::Win32::UI::WindowsAndMessaging::{
            CreatePopupMenu, DestroyMenu, GetCursorPos, PostMessageW, SetForegroundWindow, SetMenuInfo, TrackPopupMenuEx, MENUINFO,
            MIM_STYLE, MNS_CHECKORBMP, TPM_RETURNCMD, TPM_RIGHTBUTTON, WM_NULL,
        };

        let hwnd = _window.hwnd().map_err(|e| e.to_string())?;
        let (menu_x, menu_y) = unsafe {
            let mut pt = POINT::default();
            if GetCursorPos(&mut pt).is_ok() {
                (pt.x, pt.y)
            } else {
                (params.x, params.y)
            }
        };

        if let Some(exe) = resolve_native_menu_sidecar_exe() {
            let start = Instant::now();
            let mut sidecar_params = params.clone();
            sidecar_params.x = menu_x;
            sidecar_params.y = menu_y;
            sidecar_params.owner_hwnd = Some(hwnd.0 as usize as u64);
            let input = serde_json::to_string(&sidecar_params).map_err(|e| e.to_string())?;
            let tmp = std::env::temp_dir();
            let pid = std::process::id();
            let ts = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map(|d| d.as_nanos())
                .unwrap_or(0);
            let seq = NATIVE_MENU_IO_SEQ.fetch_add(1, Ordering::Relaxed);
            let in_path = tmp.join(format!("filemgr-native-menu-{pid}-{ts}-{seq}.in.json"));
            let out_path = tmp.join(format!("filemgr-native-menu-{pid}-{ts}-{seq}.out.json"));

            let sidecar_resp = (|| -> Result<Option<String>, String> {
                std::fs::write(&in_path, input.as_bytes()).map_err(|e| e.to_string())?;
                let status = std::process::Command::new(&exe)
                    .arg("--in")
                    .arg(&in_path)
                    .arg("--out")
                    .arg(&out_path)
                    .status()
                    .map_err(|e| e.to_string())?;
                if !status.success() {
                    return Err("native menu sidecar exited".to_string());
                }
                let s = std::fs::read_to_string(&out_path).map_err(|e| e.to_string())?;
                let resp: NativeContextMenuResponse = serde_json::from_str(s.trim()).map_err(|e| e.to_string())?;
                if let Some(err) = resp.error {
                    let msg = err.trim().to_string();
                    if !msg.is_empty() {
                        return Err(msg);
                    }
                }
                Ok(resp
                    .selected_id
                    .map(|x| x.trim().to_string())
                    .filter(|x| !x.is_empty()))
            })();

            let _ = std::fs::remove_file(&in_path);
            let _ = std::fs::remove_file(&out_path);

            match sidecar_resp {
                Ok(selected) => {
                    if selected.is_none() && start.elapsed().as_millis() < 80 {
                    } else {
                        return Ok(selected);
                    }
                }
                Err(_) => {}
            }
        }

        unsafe {
            let hmenu = CreatePopupMenu().map_err(|e| e.to_string())?;
            let mut mi: MENUINFO = std::mem::zeroed();
            mi.cbSize = std::mem::size_of::<MENUINFO>() as u32;
            mi.fMask = MIM_STYLE;
            mi.dwStyle = MNS_CHECKORBMP;
            let _ = SetMenuInfo(hmenu, &mut mi);

            let mut next_cmd_id: u32 = 1;
            let mut id_map: HashMap<u32, String> = HashMap::new();
            let mut bitmaps: Vec<windows::Win32::Graphics::Gdi::HBITMAP> = vec![];

            // Parse frontend theme color if provided
            let theme_color = if let Some(hex) = &params.theme_color {
                parse_hex_color(hex)
            } else {
                None
            };

            append_custom_native_menu_items(hmenu, &params.items, &mut next_cmd_id, &mut id_map, &mut bitmaps, theme_color);

            let _ = SetForegroundWindow(hwnd);
            let _style_hook = install_menu_win11_style_hook();
            let cmd = TrackPopupMenuEx(hmenu, (TPM_RETURNCMD | TPM_RIGHTBUTTON).0, menu_x, menu_y, hwnd, None);
            let _ = PostMessageW(Some(hwnd), WM_NULL, windows::Win32::Foundation::WPARAM(0), windows::Win32::Foundation::LPARAM(0));
            let _ = DestroyMenu(hmenu);
            for bmp in bitmaps {
                let _ = DeleteObject(HGDIOBJ(bmp.0));
            }

            let cmd_id = cmd.0 as u32;
            if cmd_id == 0 {
                return Ok(None);
            }
            return Ok(id_map.get(&cmd_id).cloned());
        }
    }
    #[cfg(not(windows))]
    {
        let _ = _window;
        let _ = params;
        Ok(None)
    }
}

#[tauri::command]
fn open_in_terminal(params: OpenInTerminalParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        use std::path::Path;
        use std::process::Command;
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }
        if is_shell_like_path(&p) {
            return Err("不支持对 shell 路径打开终端".to_string());
        }
        let src = Path::new(&p);
        let dir = if src.is_dir() {
            src.to_path_buf()
        } else {
            src.parent().ok_or_else(|| "无效路径".to_string())?.to_path_buf()
        };
        if !dir.exists() {
            return Err("路径不存在".to_string());
        }
        let dir_str = dir.to_string_lossy().to_string();
        if Command::new("wt").arg("-d").arg(&dir_str).spawn().is_ok() {
            return Ok(true);
        }
        Command::new("cmd")
            .arg("/K")
            .arg(format!("pushd \"{}\"", dir_str))
            .spawn()
            .map_err(|e| e.to_string())?;
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(false)
    }
}

#[derive(Debug, Deserialize)]
struct CompressToZipParams {
    paths: Vec<String>,
}

#[cfg(windows)]
fn unique_archive_dest_path(parent: &std::path::Path, base_name: &str, ext: &str) -> std::path::PathBuf {
    let mut safe = base_name.trim().to_string();
    if safe.is_empty() {
        safe = "压缩文件".to_string();
    }
    safe = safe
        .chars()
        .map(|c| match c {
            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' => '_',
            _ => c,
        })
        .collect();

    let mut i = 0u32;
    loop {
        let name = if i == 0 {
            format!("{safe}.{ext}")
        } else {
            format!("{safe} ({i}).{ext}")
        };
        let candidate = parent.join(name);
        if !candidate.exists() {
            return candidate;
        }
        i = i.saturating_add(1);
        if i > 9999 {
            return parent.join(format!("{safe} ({i}).{ext}"));
        }
    }
}

#[cfg(windows)]
fn escape_ps_single_quoted(s: &str) -> String {
    s.replace('\'', "''")
}

#[tauri::command]
fn compress_to_zip(params: CompressToZipParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        use std::path::{Path, PathBuf};
        use std::process::Command;

        let list = normalize_path_list(params.paths, 64);
        if list.is_empty() {
            return Err("未选择要压缩的项".to_string());
        }
        if list.iter().any(|p| is_shell_like_path(p)) {
            return Err("不支持对 shell 路径压缩".to_string());
        }

        let mut parent0: Option<PathBuf> = None;
        let mut parent0_cmp: Option<String> = None;
        let mut name0: Option<String> = None;
        for p in &list {
            let src = Path::new(p);
            if !src.exists() {
                return Err(format!("路径不存在：{p}"));
            }
            let parent = src.parent().ok_or_else(|| "无效路径".to_string())?;
            let parent_s = parent.to_string_lossy().to_string();
            if parent0.is_none() {
                parent0 = Some(parent.to_path_buf());
                parent0_cmp = Some(parent_s.to_ascii_lowercase());
                name0 = Some(
                    src.file_name()
                        .map(|x| x.to_string_lossy().to_string())
                        .unwrap_or_else(|| "压缩文件".to_string()),
                );
            } else if parent_s.to_ascii_lowercase() != parent0_cmp.clone().unwrap_or_default() {
                return Err("仅支持同一文件夹内的文件/文件夹压缩".to_string());
            }
        }

        let parent_dir = parent0.ok_or_else(|| "无效路径".to_string())?;
        let base_name = if list.len() == 1 {
            name0.unwrap_or_else(|| "压缩文件".to_string())
        } else {
            "压缩文件".to_string()
        };
        let dest = unique_archive_dest_path(&parent_dir, &base_name, "zip");
        let dest_s = dest.to_string_lossy().to_string();

        let items_ps = list
            .iter()
            .map(|p| format!("'{}'", escape_ps_single_quoted(p)))
            .collect::<Vec<_>>()
            .join(", ");
        let script = format!(
            "$ErrorActionPreference='Stop'; Compress-Archive -LiteralPath @({items}) -DestinationPath '{dest}'",
            items = items_ps,
            dest = escape_ps_single_quoted(&dest_s)
        );

        let status = Command::new("powershell")
            .args(["-NoProfile", "-Command", &script])
            .status()
            .map_err(|e| e.to_string())?;
        if !status.success() {
            return Err("压缩失败".to_string());
        }
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
fn compress_to_7z(params: CompressToZipParams) -> Result<bool, String> {
    #[cfg(windows)]
    {
        use std::path::{Path, PathBuf};
        use std::process::Command;

        let list = normalize_path_list(params.paths, 64);
        if list.is_empty() {
            return Err("未选择要压缩的项".to_string());
        }
        if list.iter().any(|p| is_shell_like_path(p)) {
            return Err("不支持对 shell 路径压缩".to_string());
        }

        let mut parent0: Option<PathBuf> = None;
        let mut parent0_cmp: Option<String> = None;
        let mut name0: Option<String> = None;
        for p in &list {
            let src = Path::new(p);
            if !src.exists() {
                return Err(format!("路径不存在：{p}"));
            }
            let parent = src.parent().ok_or_else(|| "无效路径".to_string())?;
            let parent_s = parent.to_string_lossy().to_string();
            if parent0.is_none() {
                parent0 = Some(parent.to_path_buf());
                parent0_cmp = Some(parent_s.to_ascii_lowercase());
                name0 = Some(
                    src.file_name()
                        .map(|x| x.to_string_lossy().to_string())
                        .unwrap_or_else(|| "压缩文件".to_string()),
                );
            } else if parent_s.to_ascii_lowercase() != parent0_cmp.clone().unwrap_or_default() {
                return Err("仅支持同一文件夹内的文件/文件夹压缩".to_string());
            }
        }

        let parent_dir = parent0.ok_or_else(|| "无效路径".to_string())?;
        let base_name = if list.len() == 1 {
            name0.unwrap_or_else(|| "压缩文件".to_string())
        } else {
            "压缩文件".to_string()
        };
        let dest = unique_archive_dest_path(&parent_dir, &base_name, "7z");
        let dest_name = dest
            .file_name()
            .map(|x| x.to_string_lossy().to_string())
            .ok_or_else(|| "无效路径".to_string())?;

        let items = list
            .iter()
            .map(|p| {
                Path::new(p)
                    .file_name()
                    .map(|x| x.to_string_lossy().to_string())
                    .ok_or_else(|| format!("无效路径：{p}"))
            })
            .collect::<Result<Vec<_>, _>>()?;

        let status = Command::new("7z")
            .current_dir(&parent_dir)
            .arg("a")
            .arg("-t7z")
            .arg(&dest_name)
            .args(items)
            .status()
            .map_err(|_| "未找到 7z（可安装 7-Zip 并确保 7z 在 PATH 中）".to_string())?;
        if !status.success() {
            return Err("压缩失败".to_string());
        }
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = params;
        Ok(false)
    }
}

#[derive(Debug, Deserialize)]
struct RenamePathParams {
    path: String,
    new_name: String,
}

#[derive(Debug, Deserialize)]
struct DeletePathsParams {
    paths: Vec<String>,
    #[serde(default)]
    recycle: Option<bool>,
}

#[derive(Debug, Deserialize)]
struct PastePathsToDirParams {
    src_paths: Vec<String>,
    dest_dir: String,
    mode: String,
}

#[derive(Debug, Deserialize)]
struct CreateDirParams {
    parent_dir: String,
    #[serde(default)]
    name: Option<String>,
}

#[derive(Debug, Deserialize)]
struct CreateFileParams {
    parent_dir: String,
    #[serde(default)]
    name: Option<String>,
    #[serde(default)]
    initial_content: Option<String>,
}

#[derive(Debug, Deserialize)]
struct ShowPropertiesParams {
    path: String,
}

fn is_shell_like_path(p: &str) -> bool {
    let s = p.trim();
    if s.is_empty() {
        return true;
    }
    s.to_ascii_lowercase().starts_with("shell:")
}

fn normalize_path_list(paths: Vec<String>, limit: usize) -> Vec<String> {
    let mut out: Vec<String> = paths
        .into_iter()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect();
    out.sort();
    out.dedup();
    out.truncate(limit.max(1));
    out
}

fn validate_new_name(name: &str) -> Result<String, String> {
    let trimmed = name.trim();
    if trimmed.is_empty() {
        return Err("名称不能为空".to_string());
    }
    if trimmed.contains('/') || trimmed.contains('\\') {
        return Err("名称不能包含路径分隔符".to_string());
    }
    #[cfg(windows)]
    {
        let invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*'];
        if trimmed.chars().any(|c| invalid.contains(&c)) {
            return Err("名称包含非法字符".to_string());
        }
        if trimmed.ends_with('.') || trimmed.ends_with(' ') {
            return Err("名称不能以空格或句点结尾".to_string());
        }
    }
    Ok(trimmed.to_string())
}

fn sanitize_default_name(name: &str) -> String {
    let mut safe = name.trim().to_string();
    if safe.is_empty() {
        safe = "新建".to_string();
    }
    #[cfg(windows)]
    {
        safe = safe
            .chars()
            .map(|c| match c {
                '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' => '_',
                _ => c,
            })
            .collect();
        while safe.ends_with('.') || safe.ends_with(' ') {
            safe.pop();
        }
        if safe.is_empty() {
            safe = "新建".to_string();
        }
    }
    safe
}

fn split_name_ext(name: &str) -> (String, Option<String>) {
    let s = name.trim();
    if s.is_empty() {
        return ("".to_string(), None);
    }
    let mut last_dot: Option<usize> = None;
    for (i, ch) in s.char_indices() {
        if ch == '.' {
            last_dot = Some(i);
        }
    }
    if let Some(i) = last_dot {
        if i > 0 && i + 1 < s.len() {
            let base = s[..i].to_string();
            let ext = s[i + 1..].to_string();
            return (base, Some(ext));
        }
    }
    (s.to_string(), None)
}

fn unique_child_path(parent: &std::path::Path, name: &str) -> std::path::PathBuf {
    let safe = name.trim();
    let (base_raw, ext_raw) = split_name_ext(safe);
    let mut base = base_raw.trim().to_string();
    if base.is_empty() {
        base = "新建".to_string();
    }
    let ext = ext_raw.map(|x| x.trim().to_string()).filter(|x| !x.is_empty());

    let mut i = 0u32;
    loop {
        let file_name = if i == 0 {
            match &ext {
                Some(e) => format!("{base}.{e}"),
                None => base.clone(),
            }
        } else {
            match &ext {
                Some(e) => format!("{base} ({i}).{e}"),
                None => format!("{base} ({i})"),
            }
        };
        let candidate = parent.join(&file_name);
        if !candidate.exists() {
            return candidate;
        }
        i = i.saturating_add(1);
        if i > 9999 {
            return parent.join(file_name);
        }
    }
}

fn copy_path_recursive(src: &std::path::Path, dest: &std::path::Path) -> io::Result<()> {
    if src.is_file() {
        if let Some(parent) = dest.parent() {
            fs::create_dir_all(parent)?;
        }
        fs::copy(src, dest)?;
        return Ok(());
    }
    if src.is_dir() {
        fs::create_dir_all(dest)?;
        for entry in WalkDir::new(src).min_depth(1) {
            let entry = entry?;
            let rel = entry.path().strip_prefix(src).unwrap_or(entry.path());
            let target = dest.join(rel);
            if entry.file_type().is_dir() {
                fs::create_dir_all(&target)?;
            } else {
                if let Some(parent) = target.parent() {
                    fs::create_dir_all(parent)?;
                }
                fs::copy(entry.path(), &target)?;
            }
        }
        return Ok(());
    }
    Err(io::Error::new(io::ErrorKind::NotFound, "source not found"))
}

fn remove_path_recursive(path: &std::path::Path) -> io::Result<()> {
    if path.is_dir() {
        fs::remove_dir_all(path)?;
        return Ok(());
    }
    fs::remove_file(path)?;
    Ok(())
}

fn move_path_to(src: &std::path::Path, dest: &std::path::Path) -> io::Result<()> {
    if let Some(parent) = dest.parent() {
        fs::create_dir_all(parent)?;
    }
    match fs::rename(src, dest) {
        Ok(()) => Ok(()),
        Err(_) => {
            copy_path_recursive(src, dest)?;
            remove_path_recursive(src)?;
            Ok(())
        }
    }
}

#[tauri::command]
async fn rename_path(params: RenamePathParams) -> Result<bool, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let p = params.path.trim().to_string();
        if is_shell_like_path(&p) {
            return Err("不支持对 shell 路径重命名".to_string());
        }
        let new_name = validate_new_name(&params.new_name)?;
        let src = std::path::PathBuf::from(&p);
        if !src.exists() {
            return Err("路径不存在".to_string());
        }
        let parent = src.parent().ok_or_else(|| "无效路径".to_string())?;
        let dest = parent.join(new_name);
        if dest == src {
            return Ok(true);
        }
        if dest.exists() {
            return Err("目标名称已存在".to_string());
        }
        fs::rename(&src, &dest).map_err(|e| e.to_string())?;
        Ok(true)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn create_dir(params: CreateDirParams) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let parent_dir = params.parent_dir.trim().to_string();
        if parent_dir.is_empty() {
            return Err("路径不能为空".to_string());
        }
        if is_shell_like_path(&parent_dir) {
            return Err("不支持对 shell 路径新建".to_string());
        }
        let parent = std::path::PathBuf::from(&parent_dir);
        if !parent.exists() {
            return Err("路径不存在".to_string());
        }
        if !parent.is_dir() {
            return Err("目标不是文件夹".to_string());
        }

        let name_opt = params.name;
        let has_name = name_opt.is_some();
        let raw_name = name_opt.unwrap_or_else(|| "新建文件夹".to_string());
        let name = if raw_name.trim().is_empty() {
            sanitize_default_name("新建文件夹")
        } else if has_name {
            validate_new_name(&raw_name)?
        } else {
            sanitize_default_name(&raw_name)
        };
        let dest = if has_name {
            let d = parent.join(&name);
            if d.exists() {
                return Err("目标名称已存在".to_string());
            }
            d
        } else {
            unique_child_path(&parent, &name)
        };
        fs::create_dir(&dest).map_err(|e| e.to_string())?;
        Ok(dest.to_string_lossy().to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn create_file(params: CreateFileParams) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let parent_dir = params.parent_dir.trim().to_string();
        if parent_dir.is_empty() {
            return Err("路径不能为空".to_string());
        }
        if is_shell_like_path(&parent_dir) {
            return Err("不支持对 shell 路径新建".to_string());
        }
        let parent = std::path::PathBuf::from(&parent_dir);
        if !parent.exists() {
            return Err("路径不存在".to_string());
        }
        if !parent.is_dir() {
            return Err("目标不是文件夹".to_string());
        }

        let name_opt = params.name;
        let has_name = name_opt.is_some();
        let raw_name = name_opt.unwrap_or_else(|| "新建文件.txt".to_string());
        let name = if raw_name.trim().is_empty() {
            sanitize_default_name("新建文件.txt")
        } else if has_name {
            validate_new_name(&raw_name)?
        } else {
            sanitize_default_name(&raw_name)
        };
        let dest = if has_name {
            let d = parent.join(&name);
            if d.exists() {
                return Err("目标名称已存在".to_string());
            }
            d
        } else {
            unique_child_path(&parent, &name)
        };
        let mut f = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&dest)
            .map_err(|e| e.to_string())?;
        if let Some(content) = params.initial_content {
            use std::io::Write;
            f.write_all(content.as_bytes()).map_err(|e| e.to_string())?;
        }
        Ok(dest.to_string_lossy().to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
fn empty_recycle_bin() -> Result<bool, String> {
    #[cfg(windows)]
    unsafe {
        let flags = SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND;
        SHEmptyRecycleBinW(None, PCWSTR::null(), flags).map_err(|e| e.to_string())?;
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        Ok(false)
    }
}

#[cfg(windows)]
unsafe fn delete_paths_recycle_bin(hwnd: HWND, paths: &[String]) -> Result<bool, String> {
    if paths.is_empty() {
        return Ok(true);
    }
    let mut wide_multi: Vec<u16> = Vec::new();
    for p in paths {
        let w = to_wide_null(p);
        if w.len() <= 1 {
            continue;
        }
        wide_multi.extend_from_slice(&w[..w.len() - 1]);
        wide_multi.push(0);
    }
    wide_multi.push(0);
    if wide_multi.len() < 2 {
        return Ok(true);
    }

    let mut op = SHFILEOPSTRUCTW::default();
    op.hwnd = hwnd;
    op.wFunc = FO_DELETE;
    op.pFrom = PCWSTR(wide_multi.as_ptr());
    op.fFlags = ((FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT).0) as u16;

    let res = SHFileOperationW(&mut op);
    if res != 0 {
        return Err(format!("删除失败（{res}）"));
    }
    Ok(!op.fAnyOperationsAborted.as_bool())
}

#[cfg(windows)]
unsafe fn delete_paths_recycle_bin_chunked(hwnd: HWND, paths: &[String]) -> Result<bool, String> {
    if paths.is_empty() {
        return Ok(true);
    }
    let mut aborted = false;
    let mut chunk: Vec<String> = Vec::new();
    let mut wide_len: usize = 1;
    let max_wide_len: usize = 30_000;

    for p in paths {
        let w = to_wide_null(p);
        if w.len() <= 1 {
            continue;
        }
        if !chunk.is_empty() && wide_len.saturating_add(w.len()) > max_wide_len {
            let ok = delete_paths_recycle_bin(hwnd, &chunk)?;
            if !ok {
                aborted = true;
            }
            chunk.clear();
            wide_len = 1;
        }
        chunk.push(p.to_string());
        wide_len = wide_len.saturating_add(w.len());
    }

    if !chunk.is_empty() {
        let ok = delete_paths_recycle_bin(hwnd, &chunk)?;
        if !ok {
            aborted = true;
        }
    }

    Ok(!aborted)
}

#[tauri::command]
async fn delete_paths(window: tauri::Window, params: DeletePathsParams) -> Result<bool, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let recycle = params.recycle.unwrap_or(true);
        let paths = normalize_path_list(params.paths, 100_000);
        let paths: Vec<String> = paths.into_iter().filter(|p| !is_shell_like_path(p)).collect();
        if paths.is_empty() {
            return Ok(true);
        }

        #[cfg(windows)]
        unsafe {
            if recycle {
                let _com = com_init();
                let hwnd = window.hwnd().map_err(|e| e.to_string())?;
                return delete_paths_recycle_bin_chunked(hwnd, &paths);
            }
        }

        for p in &paths {
            let path = std::path::PathBuf::from(p);
            remove_path_recursive(&path).map_err(|e| e.to_string())?;
        }
        Ok(true)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn paste_paths_to_dir(params: PastePathsToDirParams) -> Result<bool, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let src_paths = normalize_path_list(params.src_paths, 256);
        let dest_dir = params.dest_dir.trim().to_string();
        if is_shell_like_path(&dest_dir) {
            return Err("目标目录无效".to_string());
        }
        let mode = params.mode.trim().to_ascii_lowercase();
        if mode != "copy" && mode != "move" {
            return Err("mode 只能为 copy 或 move".to_string());
        }

        let dest_dir_path = std::path::PathBuf::from(&dest_dir);
        if !dest_dir_path.is_dir() {
            return Err("目标不是目录".to_string());
        }

        for src in src_paths {
            if is_shell_like_path(&src) {
                continue;
            }
            let src_path = std::path::PathBuf::from(&src);
            let name = src_path
                .file_name()
                .and_then(|s| s.to_str())
                .map(|s| s.to_string())
                .ok_or_else(|| "无法获取源文件名".to_string())?;
            let dest_path = dest_dir_path.join(name);
            if dest_path.exists() {
                return Err("目标已存在".to_string());
            }
            if mode == "move" {
                move_path_to(&src_path, &dest_path).map_err(|e| e.to_string())?;
            } else {
                copy_path_recursive(&src_path, &dest_path).map_err(|e| e.to_string())?;
            }
        }
        Ok(true)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
fn show_properties(window: tauri::Window, params: ShowPropertiesParams) -> Result<bool, String> {
    let p = params.path.trim().to_string();
    if p.is_empty() {
        return Err("路径不能为空".to_string());
    }
    #[cfg(windows)]
    unsafe {
        let hwnd = window.hwnd().map_err(|e| e.to_string())?;
        let path = to_wide_null(&p);
        let ok = SHObjectProperties(Some(hwnd), SHOP_FILEPATH, PCWSTR(path.as_ptr()), PCWSTR(std::ptr::null())).as_bool();
        if ok {
            return Ok(true);
        }

        let verb = to_wide_null("properties");
        let res = ShellExecuteW(
            Some(hwnd),
            PCWSTR(verb.as_ptr()),
            PCWSTR(path.as_ptr()),
            PCWSTR(std::ptr::null()),
            PCWSTR(std::ptr::null()),
            SW_SHOW,
        );
        if res.0 as isize <= 32 {
            return Err(format!("打开属性失败（{code}）", code = res.0 as isize));
        }
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        Ok(false)
    }
}

#[derive(Debug, Deserialize)]
struct ShowShellContextMenuParams {
    paths: Vec<String>,
}

#[derive(Debug, Deserialize)]
struct ShowFolderBackgroundContextMenuParams {
    path: String,
}

#[cfg(windows)]
thread_local! {
    static SHELL_MENU_HOOK: RefCell<Option<ShellMenuHook>> = const { RefCell::new(None) };
}

#[cfg(windows)]
struct ShellMenuHook {
    hwnd: HWND,
    old_wndproc: isize,
    menu2: Option<IContextMenu2>,
    menu3: Option<IContextMenu3>,
}

#[cfg(windows)]
extern "system" fn shell_menu_wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    let mut handled: Option<LRESULT> = None;
    SHELL_MENU_HOOK.with(|cell| {
        let borrowed = cell.borrow();
        let hook = match borrowed.as_ref() {
            Some(v) => v,
            None => return,
        };
        if hook.hwnd != hwnd {
            return;
        }
        let is_menu_msg = matches!(msg, WM_INITMENUPOPUP | WM_DRAWITEM | WM_MEASUREITEM | WM_MENUCHAR);
        if !is_menu_msg {
            return;
        }
        unsafe {
            if let Some(m3) = hook.menu3.as_ref() {
                let mut out = LRESULT(0);
                if m3.HandleMenuMsg2(msg, wparam, lparam, Some(&mut out)).is_ok() {
                    handled = Some(out);
                }
                return;
            }
            if let Some(m2) = hook.menu2.as_ref() {
                let _ = m2.HandleMenuMsg(msg, wparam, lparam);
                handled = Some(LRESULT(0));
            }
        }
    });
    if let Some(v) = handled {
        return v;
    }
    let old = SHELL_MENU_HOOK.with(|cell| cell.borrow().as_ref().map(|h| h.old_wndproc).unwrap_or(0));
    unsafe { CallWindowProcW(Some(std::mem::transmute(old)), hwnd, msg, wparam, lparam) }
}

#[cfg(windows)]
struct ShellMenuHookGuard;

#[cfg(windows)]
impl ShellMenuHookGuard {
    fn install(hwnd: HWND, menu2: Option<IContextMenu2>, menu3: Option<IContextMenu3>) -> Result<Self, String> {
        unsafe {
            let prev = SetWindowLongPtrW(hwnd, GWLP_WNDPROC, shell_menu_wndproc as _);
            SHELL_MENU_HOOK.with(|cell| {
                *cell.borrow_mut() = Some(ShellMenuHook {
                    hwnd,
                    old_wndproc: prev,
                    menu2,
                    menu3,
                });
            });
        }
        Ok(Self)
    }
}

#[cfg(windows)]
impl Drop for ShellMenuHookGuard {
    fn drop(&mut self) {
        SHELL_MENU_HOOK.with(|cell| {
            let mut borrowed = cell.borrow_mut();
            let hook = borrowed.take();
            if let Some(h) = hook {
                unsafe {
                    let _ = SetWindowLongPtrW(h.hwnd, GWLP_WNDPROC, h.old_wndproc);
                }
            }
        });
    }
}

#[cfg(windows)]
unsafe fn shell_context_menu_for_paths(hwnd: HWND, paths: &[String]) -> Result<(IContextMenu, Option<IContextMenu2>, Option<IContextMenu3>), String> {
    use std::path::Path;

    let mut normalized: Vec<String> = paths
        .iter()
        .map(|s| s.trim().to_string())
        .filter(|s| !s.is_empty())
        .collect();
    normalized.dedup();
    normalized.truncate(64);
    let p0 = normalized.get(0).cloned().unwrap_or_default();
    if p0.is_empty() {
        return Err("empty path".to_string());
    }

    let mut use_paths = normalized;
    if use_paths.len() > 1 {
        if p0.starts_with("shell:") {
            use_paths.truncate(1);
        } else {
            let parent0 = Path::new(&p0)
                .parent()
                .map(|p| p.to_string_lossy().to_string())
                .unwrap_or_default()
                .to_ascii_lowercase();
            if parent0.is_empty() {
                use_paths.truncate(1);
            } else {
                let same_parent = use_paths.iter().all(|p| {
                    if p.starts_with("shell:") {
                        return false;
                    }
                    Path::new(p)
                        .parent()
                        .map(|pp| pp.to_string_lossy().to_string())
                        .unwrap_or_default()
                        .to_ascii_lowercase()
                        == parent0
                });
                if !same_parent {
                    use_paths.truncate(1);
                }
            }
        }
    }

    let mut pidl_guards: Vec<PidlGuard> = Vec::with_capacity(use_paths.len());
    let mut child_ptrs: Vec<*const ITEMIDLIST> = Vec::with_capacity(use_paths.len());
    let mut parent: Option<IShellFolder> = None;

    for (idx, p) in use_paths.iter().enumerate() {
        let wide = to_wide_null(p);
        let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
        SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None).map_err(|e| e.to_string())?;
        if pidl.is_null() {
            continue;
        }
        let pidl_guard = PidlGuard(pidl);

        let mut child: *mut ITEMIDLIST = std::ptr::null_mut();
        let folder = SHBindToParent::<IShellFolder>(pidl_guard.0, Some(&mut child)).map_err(|e| e.to_string())?;
        if child.is_null() {
            continue;
        }

        if idx == 0 {
            parent = Some(folder);
        } else {
            let _ = folder;
        }

        child_ptrs.push(child as *const ITEMIDLIST);
        pidl_guards.push(pidl_guard);
    }

    let parent = parent.ok_or_else(|| "bind to parent failed".to_string())?;
    if child_ptrs.is_empty() {
        return Err("no valid paths".to_string());
    }

    let cm = parent
        .GetUIObjectOf::<IContextMenu>(hwnd, child_ptrs.as_slice(), None)
        .map_err(|e| e.to_string())?;
    let m2 = cm.cast::<IContextMenu2>().ok();
    let m3 = cm.cast::<IContextMenu3>().ok();
    Ok((cm, m2, m3))
}

#[tauri::command]
fn show_shell_context_menu(window: tauri::Window, params: ShowShellContextMenuParams) -> Result<bool, String> {
    #[cfg(windows)]
    unsafe {
        let _com = com_init();
        let hwnd = window.hwnd().map_err(|e| e.to_string())?;
        let (cm, m2, m3) = shell_context_menu_for_paths(hwnd, &params.paths)?;
        let _guard = ShellMenuHookGuard::install(hwnd, m2, m3)?;

        let hmenu: HMENU = CreatePopupMenu().map_err(|e| e.to_string())?;

        let id_first: u32 = 1;
        let id_last: u32 = 0x7fff;
        cm.QueryContextMenu(hmenu, 0, id_first, id_last, CMF_NORMAL)
            .ok()
            .map_err(|e| e.to_string())?;

        let mut pt = POINT::default();
        GetCursorPos(&mut pt).map_err(|e| e.to_string())?;
        let _ = SetForegroundWindow(hwnd);
        let cmd = TrackPopupMenuEx(
            hmenu,
            (TPM_RETURNCMD | TPM_RIGHTBUTTON).0,
            pt.x,
            pt.y,
            hwnd,
            None,
        );
        let _ = PostMessageW(Some(hwnd), WM_NULL, WPARAM(0), LPARAM(0));

        let _ = DestroyMenu(hmenu);

        let cmd_id = cmd.0 as u32;
        if cmd_id == 0 {
            return Ok(true);
        }

        let verb_offset = cmd_id.saturating_sub(id_first) as usize;
        let mut invoke: CMINVOKECOMMANDINFOEX = std::mem::zeroed();
        invoke.cbSize = std::mem::size_of::<CMINVOKECOMMANDINFOEX>() as u32;
        invoke.fMask = CMIC_MASK_PTINVOKE | 0x00004000;
        invoke.hwnd = hwnd;
        invoke.lpVerb = PCSTR(verb_offset as usize as *const u8);
        invoke.lpVerbW = PCWSTR(verb_offset as usize as *const u16);
        invoke.nShow = 1;
        invoke.ptInvoke = pt;

        cm.InvokeCommand(&mut invoke as *mut _ as *mut _)
            .map_err(|e| e.to_string())?;
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
fn show_folder_background_context_menu(window: tauri::Window, params: ShowFolderBackgroundContextMenuParams) -> Result<bool, String> {
    #[cfg(windows)]
    unsafe {
        let _com = com_init();
        let hwnd = window.hwnd().map_err(|e| e.to_string())?;
        let p = params.path.trim().to_string();
        if p.is_empty() {
            return Err("路径不能为空".to_string());
        }

        let desktop = SHGetDesktopFolder().map_err(|e| e.to_string())?;
        let wide = to_wide_null(&p);
        let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
        SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None).map_err(|e| e.to_string())?;
        if pidl.is_null() {
            return Err("无法解析路径".to_string());
        }
        let pidl_guard = PidlGuard(pidl);

        let folder: IShellFolder = desktop
            .BindToObject::<_, IShellFolder>(pidl_guard.0, None)
            .map_err(|e| e.to_string())?;

        let cm: IContextMenu = folder.CreateViewObject(hwnd).map_err(|e| e.to_string())?;
        let m2 = cm.cast::<IContextMenu2>().ok();
        let m3 = cm.cast::<IContextMenu3>().ok();
        let _guard = ShellMenuHookGuard::install(hwnd, m2, m3)?;

        let hmenu: HMENU = CreatePopupMenu().map_err(|e| e.to_string())?;
        let id_first: u32 = 1;
        let id_last: u32 = 0x7fff;
        cm.QueryContextMenu(hmenu, 0, id_first, id_last, CMF_NORMAL)
            .ok()
            .map_err(|e| e.to_string())?;

        let mut pt = POINT::default();
        GetCursorPos(&mut pt).map_err(|e| e.to_string())?;
        let _ = SetForegroundWindow(hwnd);
        let cmd = TrackPopupMenuEx(
            hmenu,
            (TPM_RETURNCMD | TPM_RIGHTBUTTON).0,
            pt.x,
            pt.y,
            hwnd,
            None,
        );
        let _ = PostMessageW(Some(hwnd), WM_NULL, WPARAM(0), LPARAM(0));
        let _ = DestroyMenu(hmenu);

        let cmd_id = cmd.0 as u32;
        if cmd_id == 0 {
            return Ok(true);
        }

        let verb_offset = cmd_id.saturating_sub(id_first) as usize;
        let mut invoke: CMINVOKECOMMANDINFOEX = std::mem::zeroed();
        invoke.cbSize = std::mem::size_of::<CMINVOKECOMMANDINFOEX>() as u32;
        invoke.fMask = CMIC_MASK_PTINVOKE | 0x00004000;
        invoke.hwnd = hwnd;
        invoke.lpVerb = PCSTR(verb_offset as usize as *const u8);
        invoke.lpVerbW = PCWSTR(verb_offset as usize as *const u16);
        invoke.nShow = 1;
        invoke.ptInvoke = pt;
        cm.InvokeCommand(&mut invoke as *mut _ as *mut _)
            .map_err(|e| e.to_string())?;
        return Ok(true);
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
fn get_system_accent_color() -> Option<String> {
    #[cfg(windows)]
    unsafe {
        use windows::Win32::{
            Graphics::Dwm::DwmGetColorizationColor,
        };
        use windows::core::BOOL;

        let mut dwm_color: u32 = 0;
        let mut opaque = BOOL(0);
        if DwmGetColorizationColor(&mut dwm_color, &mut opaque).is_ok() {
            let r = ((dwm_color >> 16) & 0xff) as u8;
            let g = ((dwm_color >> 8) & 0xff) as u8;
            let b = (dwm_color & 0xff) as u8;
            return Some(format!("#{:02x}{:02x}{:02x}", r, g, b));
        }
        None
    }

    #[cfg(not(windows))]
    {
        None
    }
}

#[tauri::command]
fn show_quick_access_context_menu(window: tauri::Window, params: ShowQuickAccessContextMenuParams) -> Result<bool, String> {
    #[cfg(windows)]
    unsafe {
        let _com = com_init();
        let hwnd = window.hwnd().map_err(|e| e.to_string())?;
        let target = normalize_compare_path_win(&params.path);
        if target.is_empty() {
            return Err("路径不能为空".to_string());
        }

        let desktop = SHGetDesktopFolder().map_err(|e| e.to_string())?;
        let wide = to_wide_null(quick_access_namespace());
        let mut pidl: *mut ITEMIDLIST = std::ptr::null_mut();
        SHParseDisplayName(PCWSTR(wide.as_ptr()), None, &mut pidl, 0, None).map_err(|e| e.to_string())?;
        if pidl.is_null() {
            return Err("无法打开快速访问".to_string());
        }
        let pidl_guard = PidlGuard(pidl);
        let folder: IShellFolder = desktop
            .BindToObject::<_, IShellFolder>(pidl_guard.0, None)
            .map_err(|e| e.to_string())?;

        let mut enum_list: Option<IEnumIDList> = None;
        let flags = SHCONTF_FOLDERS.0 as u32;
        let hr = folder.EnumObjects(HWND(std::ptr::null_mut()), flags, &mut enum_list);
        if hr.0 < 0 {
            return Err("无法枚举快速访问".to_string());
        }
        let enum_list = enum_list.ok_or_else(|| "无法枚举快速访问".to_string())?;

        loop {
            let mut fetched: u32 = 0;
            let mut rgelt: [*mut ITEMIDLIST; 1] = [std::ptr::null_mut()];
            let hr = enum_list.Next(&mut rgelt, Some(&mut fetched as *mut u32));
            let child_pidl = rgelt[0];
            if hr.0 < 0 || fetched == 0 || child_pidl.is_null() {
                break;
            }
            let child_guard = PidlGuard(child_pidl);

            let mut disp: STRRET = std::mem::zeroed();
            if folder.GetDisplayNameOf(child_guard.0, SHGDN_FORPARSING, &mut disp).is_err() {
                continue;
            }
            let path = strret_to_string(&mut disp, child_guard.0).unwrap_or_default();
            if normalize_compare_path_win(&path) != target {
                continue;
            }

            let child_ptrs = [child_guard.0 as *const ITEMIDLIST];
            let cm = folder
                .GetUIObjectOf::<IContextMenu>(hwnd, child_ptrs.as_slice(), None)
                .map_err(|e| e.to_string())?;
            let m2 = cm.cast::<IContextMenu2>().ok();
            let m3 = cm.cast::<IContextMenu3>().ok();
            let _guard = ShellMenuHookGuard::install(hwnd, m2, m3)?;

            let hmenu: HMENU = CreatePopupMenu().map_err(|e| e.to_string())?;
            let id_first: u32 = 1;
            let id_last: u32 = 0x7fff;
            cm.QueryContextMenu(hmenu, 0, id_first, id_last, CMF_NORMAL)
                .ok()
                .map_err(|e| e.to_string())?;

            let mut pt = POINT::default();
            GetCursorPos(&mut pt).map_err(|e| e.to_string())?;
            let _ = SetForegroundWindow(hwnd);
            let cmd = TrackPopupMenuEx(
                hmenu,
                (TPM_RETURNCMD | TPM_RIGHTBUTTON).0,
                pt.x,
                pt.y,
                hwnd,
                None,
            );
            let _ = PostMessageW(Some(hwnd), WM_NULL, WPARAM(0), LPARAM(0));
            let _ = DestroyMenu(hmenu);

            let cmd_id = cmd.0 as u32;
            if cmd_id == 0 {
                return Ok(true);
            }

            let verb_offset = cmd_id.saturating_sub(id_first) as usize;
            let mut invoke: CMINVOKECOMMANDINFOEX = std::mem::zeroed();
            invoke.cbSize = std::mem::size_of::<CMINVOKECOMMANDINFOEX>() as u32;
            invoke.fMask = CMIC_MASK_PTINVOKE | 0x00004000;
            invoke.hwnd = hwnd;
            invoke.lpVerb = PCSTR(verb_offset as usize as *const u8);
            invoke.lpVerbW = PCWSTR(verb_offset as usize as *const u16);
            invoke.nShow = 1;
            invoke.ptInvoke = pt;
            cm.InvokeCommand(&mut invoke as *mut _ as *mut _)
                .map_err(|e| e.to_string())?;
            return Ok(true);
        }
        Err("未找到对应的快速访问项".to_string())
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        let _ = params;
        Ok(false)
    }
}

#[tauri::command]
fn set_window_effect(window: tauri::Window, effect: String) {
    #[cfg(target_os = "windows")]
    {
        let _ = clear_mica(&window);
        let _ = clear_acrylic(&window);
        let _ = clear_blur(&window);
        match effect.as_str() {
            "mica" => {
                let _ = apply_mica(&window, None);
            }
            "acrylic" => {
                let _ = apply_acrylic(&window, Some((0, 0, 0, 0)));
            }
            "acrylic_strong" => {
                let _ = apply_acrylic(&window, Some((0, 0, 0, 0)));
            }
            "blur" => {
                let _ = apply_blur(&window, Some((0, 0, 0, 0)));
            }
            "transparent" => {}
            "none" => {}
            _ => {}
        }
    }
}

#[tauri::command]
fn set_window_corner_preference(window: tauri::Window, preference: String) -> Result<bool, String> {
    #[cfg(windows)]
    unsafe {
        use windows::Win32::Graphics::Dwm::{DwmSetWindowAttribute, DWMWINDOWATTRIBUTE};
        let hwnd = window.hwnd().map_err(|e| e.to_string())?;
        let v: i32 = match preference.as_str() {
            "dont_round" => 1,
            "round" => 2,
            "round_small" => 3,
            _ => 0,
        };
        let attr = DWMWINDOWATTRIBUTE(33);
        Ok(DwmSetWindowAttribute(hwnd, attr, &v as *const _ as _, std::mem::size_of::<i32>() as u32).is_ok())
    }
    #[cfg(not(windows))]
    {
        let _ = window;
        let _ = preference;
        Ok(false)
    }
}

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_drag::init())
        .setup(|app| {
            #[cfg(windows)]
            {
                let id = app.config().identifier.clone();
                win_app_id::set_app_user_model_id(id.clone());
                let exe = std::env::current_exe().ok().map(|p| p.to_string_lossy().to_string()).unwrap_or_default();
                let wide = to_wide_null(&id);
                unsafe {
                    let _ = SetCurrentProcessExplicitAppUserModelID(PCWSTR(wide.as_ptr()));
                }
                let _ = tauri::async_runtime::spawn_blocking({
                    let id2 = id.clone();
                    let exe2 = exe.clone();
                    move || {
                        if !exe2.trim().is_empty() {
                            let _ = win_app_id::ensure_start_menu_shortcut_windows("FileMgr", &exe2, &id2);
                        }
                    }
                });
                let _ = tauri::async_runtime::spawn_blocking(move || {
                    let _ = set_jump_list_windows(vec![]);
                });
                let app_handle = app.handle().clone();
                std::thread::spawn(move || {
                    let mut last_sig: u64 = 0;
                    loop {
                        let sig = match list_quick_access_windows() {
                            Ok(list) => {
                                let mut h: u64 = 14695981039346656037;
                                for it in list {
                                    let p = normalize_compare_path_win(&it.path);
                                    for b in p.as_bytes() {
                                        h ^= *b as u64;
                                        h = h.wrapping_mul(1099511628211);
                                    }
                                    h ^= if it.pinned { 1 } else { 0 };
                                    h = h.wrapping_mul(1099511628211);
                                }
                                h
                            }
                            Err(_) => 0,
                        };
                        if sig != last_sig {
                            last_sig = sig;
                            if let Some(w) = app_handle.get_webview_window("main") {
                                let _ = w.emit("sidebar_should_refresh", ());
                            }
                        }
                        std::thread::sleep(Duration::from_millis(1500));
                    }
                });
            }
            let start_path = start_path_from_args();
            if let Some(p) = start_path {
                if let Some(w) = app.get_webview_window("main") {
                    let js = format!(
                        "window.__FILEMGR_START_PATH = {};",
                        serde_json::to_string(&p).unwrap_or_else(|_| "\"\"".to_string())
                    );
                    let _ = w.eval(&js);
                }
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            list_dir,
            dir_stats,
            dir_stats_stream,
            dir_stats_cancel,
            folder_size,
            folder_size_stream,
            search_dir,
            list_roots,
            set_jump_list,
            list_quick_access,
            is_in_quick_access,
            pin_to_quick_access,
            remove_from_quick_access,
            show_quick_access_context_menu,
            list_roots_detailed,
            get_icon_png_base64,
            get_stock_icon_png_base64,
            get_new_item_icon_png_base64,
            confirm_message_box,
            confirm_task_dialog,
            get_basic_file_info,
            get_icons_png_base64_batch,
            get_image_thumbs_png_base64_batch,
            get_shell_thumbs_png_base64_batch,
            get_media_metadata,
            list_shell_folder,
            scan_gallery_images,
            open_path,
            get_drag_icon_path,
            get_drag_icon_path_for_target,
            show_native_context_menu,
            open_in_terminal,
            compress_to_zip,
            compress_to_7z,
            rename_path,
            create_dir,
            create_file,
            empty_recycle_bin,
            delete_paths,
            paste_paths_to_dir,
            show_properties,
            show_shell_context_menu,
            show_folder_background_context_menu,
            get_system_accent_color,
            set_window_effect,
            set_window_corner_preference,
            read_text_file
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

fn start_path_from_args() -> Option<String> {
    let mut it = std::env::args().skip(1);
    while let Some(a) = it.next() {
        if let Some(v) = a.strip_prefix("--path=") {
            let p = v.trim().trim_matches('"').to_string();
            if !p.is_empty() {
                return Some(p);
            }
            continue;
        }
        if a == "--path" || a == "--tabPath" {
            if let Some(v) = it.next() {
                let p = v.trim().trim_matches('"').to_string();
                if !p.is_empty() {
                    return Some(p);
                }
            }
            continue;
        }
        if let Some(v) = a.strip_prefix("--tabPath=") {
            let p = v.trim().trim_matches('"').to_string();
            if !p.is_empty() {
                return Some(p);
            }
            continue;
        }
    }
    None
}

#[cfg(all(test, windows))]
mod tests {
    use super::*;

    #[test]
    fn drag_icon_path_for_target_is_png() {
        let exe = std::env::current_exe().unwrap();
        let params = GetIconParams {
            path: exe.to_string_lossy().to_string(),
            size: Some(48),
        };
        let p = get_drag_icon_path_for_target(params).unwrap();
        let bytes = std::fs::read(&p).unwrap();
        assert!(bytes.starts_with(&[137, 80, 78, 71, 13, 10, 26, 10]));
    }

    #[test]
    fn jumbo_icon_is_not_tiny() {
        let p = std::env::var("SystemRoot")
            .map(|r| format!(r"{}\System32\notepad.exe", r.trim_end_matches(['\\', '/'])))
            .unwrap_or_else(|_| r"C:\Windows\System32\notepad.exe".to_string());
        let b64 = icon_png_base64_for_any_jumbo(&p, 256).ok().flatten().unwrap_or_default();
        assert!(!b64.is_empty());
        let png = base64::engine::general_purpose::STANDARD.decode(b64).unwrap();
        let img = image::load_from_memory(&png).unwrap();
        assert!(img.width() >= 48);
        assert!(img.height() >= 48);
    }
}

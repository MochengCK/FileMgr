use serde::{Deserialize, Serialize};
use std::{
    fs,
    io::{self, BufRead, Write},
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

#[cfg(windows)]
use base64::Engine as _;
#[cfg(windows)]
use windows::core::PCWSTR;
#[cfg(windows)]
use windows::Win32::{
    Graphics::Gdi::{
        CreateCompatibleDC, DeleteDC, DeleteObject, GetDIBits, GetObjectW, SelectObject, BITMAP,
        BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HBITMAP, HGDIOBJ,
    },
    Storage::FileSystem::{GetDriveTypeW, GetVolumeInformationW, FILE_FLAGS_AND_ATTRIBUTES},
    UI::{
        Shell::{SHGetFileInfoW, SHFILEINFOW, SHGFI_ICON, SHGFI_LARGEICON, SHGFI_SMALLICON},
        WindowsAndMessaging::{DestroyIcon, GetIconInfo, ICONINFO},
    },
};

#[derive(Debug, Deserialize)]
struct Request {
    id: u64,
    method: String,
    #[serde(default)]
    params: serde_json::Value,
}

#[derive(Debug, Serialize)]
struct ResponseOk<T: Serialize> {
    id: u64,
    ok: bool,
    result: T,
}

#[derive(Debug, Serialize)]
struct ResponseErr {
    id: u64,
    ok: bool,
    error: ErrorBody,
}

#[derive(Debug, Serialize)]
struct ErrorBody {
    message: String,
}

#[derive(Debug, Deserialize)]
struct ListDirParams {
    path: String,
}

#[derive(Debug, Deserialize)]
struct GetIconParams {
    path: String,
    #[serde(default)]
    size: Option<u32>,
}

#[derive(Debug, Serialize)]
struct RootItemDetailed {
    path: String,
    label: String,
    icon_png_base64: Option<String>,
}

#[derive(Debug, Serialize)]
struct DirEntryItem {
    name: String,
    path: String,
    is_dir: bool,
    size: Option<u64>,
    modified_ms: Option<u128>,
}

fn system_time_to_ms(t: SystemTime) -> Option<u128> {
    t.duration_since(UNIX_EPOCH).ok().map(|d| d.as_millis())
}

fn list_dir(path: &str) -> io::Result<Vec<DirEntryItem>> {
    let mut items = Vec::new();
    let dir_path = PathBuf::from(path);
    for entry in fs::read_dir(&dir_path)? {
        let entry = entry?;
        let file_name = entry.file_name().to_string_lossy().to_string();
        let full_path = entry.path();
        let metadata = entry.metadata().ok();
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

#[cfg(windows)]
fn to_wide_null(s: &str) -> Vec<u16> {
    let mut v: Vec<u16> = s.encode_utf16().collect();
    v.push(0);
    v
}

#[cfg(windows)]
fn hbitmap_to_rgba(hbmp: HBITMAP) -> io::Result<(u32, u32, Vec<u8>)> {
    unsafe {
        let mut bmp = BITMAP::default();
        let got = GetObjectW(
            HGDIOBJ(hbmp.0),
            std::mem::size_of::<BITMAP>() as i32,
            Some((&mut bmp) as *mut _ as _),
        );
        if got == 0 {
            return Err(io::Error::new(io::ErrorKind::Other, "GetObjectW failed"));
        }

        let width = bmp.bmWidth as i32;
        let height = bmp.bmHeight as i32;
        if width <= 0 || height == 0 {
            return Err(io::Error::new(io::ErrorKind::Other, "invalid bitmap size"));
        }

        let abs_height = height.unsigned_abs();
        let mut bi = BITMAPINFO::default();
        bi.bmiHeader = BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: width,
            biHeight: abs_height as i32,
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0 as u32,
            biSizeImage: 0,
            biXPelsPerMeter: 0,
            biYPelsPerMeter: 0,
            biClrUsed: 0,
            biClrImportant: 0,
        };

        let bytes_per_row = (width as u32) * 4;
        let mut buf = vec![0u8; (bytes_per_row * abs_height) as usize];

        let hdc = CreateCompatibleDC(None);
        if hdc.0.is_null() {
            return Err(io::Error::new(
                io::ErrorKind::Other,
                "CreateCompatibleDC failed",
            ));
        }
        let prev = SelectObject(hdc, HGDIOBJ(hbmp.0));
        let res = GetDIBits(
            hdc,
            hbmp,
            0,
            abs_height as u32,
            Some(buf.as_mut_ptr() as _),
            &mut bi,
            DIB_RGB_COLORS,
        );
        SelectObject(hdc, prev);
        let _ = DeleteDC(hdc);

        if res == 0 {
            return Err(io::Error::new(io::ErrorKind::Other, "GetDIBits failed"));
        }

        let mut rgba = vec![0u8; buf.len()];
        let row_len = bytes_per_row as usize;
        for y in 0..(abs_height as usize) {
            let src_row = if height > 0 {
                abs_height as usize - 1 - y
            } else {
                y
            };
            let src_off = src_row * row_len;
            let dst_off = y * row_len;
            for x in 0..width as usize {
                let i = src_off + x * 4;
                let o = dst_off + x * 4;
                let b = buf[i];
                let g = buf[i + 1];
                let r = buf[i + 2];
                let a = buf[i + 3];
                rgba[o] = r;
                rgba[o + 1] = g;
                rgba[o + 2] = b;
                rgba[o + 3] = a;
            }
        }

        Ok((width as u32, abs_height, rgba))
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
        if ret == 0 || info.hIcon.0.is_null() {
            return Ok(None);
        }

        let hicon = info.hIcon;
        let mut icon_info = ICONINFO::default();
        let ok = GetIconInfo(hicon, &mut icon_info).is_ok();
        if !ok {
            let _ = DestroyIcon(hicon);
            return Ok(None);
        }

        let hbmp_color = icon_info.hbmColor;
        let hbmp_mask = icon_info.hbmMask;
        let result = if !hbmp_color.0.is_null() {
            let (w, h, rgba) = hbitmap_to_rgba(hbmp_color)?;
            let img = image::RgbaImage::from_raw(w, h, rgba)
                .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "invalid rgba buffer"))?;
            let mut out = Vec::new();
            img.write_to(&mut std::io::Cursor::new(&mut out), image::ImageFormat::Png)
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e.to_string()))?;
            Some(base64::engine::general_purpose::STANDARD.encode(out))
        } else if !hbmp_mask.0.is_null() {
            let (w, h, rgba) = hbitmap_to_rgba(hbmp_mask)?;
            let img = image::RgbaImage::from_raw(w, h, rgba)
                .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "invalid rgba buffer"))?;
            let mut out = Vec::new();
            img.write_to(&mut std::io::Cursor::new(&mut out), image::ImageFormat::Png)
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e.to_string()))?;
            Some(base64::engine::general_purpose::STANDARD.encode(out))
        } else {
            None
        };

        if !hbmp_color.0.is_null() {
            let _ = DeleteObject(HGDIOBJ(hbmp_color.0));
        }
        if !hbmp_mask.0.is_null() {
            let _ = DeleteObject(HGDIOBJ(hbmp_mask.0));
        }
        let _ = DestroyIcon(hicon);

        Ok(result)
    }
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
    let vol_name = if ok {
        utf16_z_to_string(&vol)
    } else {
        String::new()
    };
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
fn list_roots_detailed() -> io::Result<Vec<RootItemDetailed>> {
    let mut roots = Vec::new();
    for i in 0..26u8 {
        let letter = (b'A' + i) as char;
        let path = format!("{letter}:\\");
        if !std::path::Path::new(&path).exists() {
            continue;
        }
        let label = drive_label_for_root(&path);
        let icon_png_base64 = icon_png_base64_for_path(&path, Some(16)).ok().flatten();
        roots.push(RootItemDetailed {
            path,
            label,
            icon_png_base64,
        });
    }
    Ok(roots)
}

fn write_json_line<T: Serialize>(value: &T) {
    let mut out = io::stdout().lock();
    if let Ok(s) = serde_json::to_string(value) {
        let _ = out.write_all(s.as_bytes());
        let _ = out.write_all(b"\n");
        let _ = out.flush();
    }
}

fn main() {
    let stdin = io::stdin();
    for line in stdin.lock().lines() {
        let line = match line {
            Ok(l) => l,
            Err(_) => break,
        };
        if line.trim().is_empty() {
            continue;
        }

        let req: Request = match serde_json::from_str(&line) {
            Ok(v) => v,
            Err(_) => continue,
        };

        match req.method.as_str() {
            "list_dir" => {
                let params: Result<ListDirParams, _> = serde_json::from_value(req.params);
                match params {
                    Ok(p) => match list_dir(&p.path) {
                        Ok(items) => write_json_line(&ResponseOk {
                            id: req.id,
                            ok: true,
                            result: items,
                        }),
                        Err(e) => write_json_line(&ResponseErr {
                            id: req.id,
                            ok: false,
                            error: ErrorBody {
                                message: e.to_string(),
                            },
                        }),
                    },
                    Err(e) => write_json_line(&ResponseErr {
                        id: req.id,
                        ok: false,
                        error: ErrorBody {
                            message: e.to_string(),
                        },
                    }),
                }
            }
            "list_roots_detailed" => {
                #[cfg(windows)]
                match list_roots_detailed() {
                    Ok(items) => write_json_line(&ResponseOk {
                        id: req.id,
                        ok: true,
                        result: items,
                    }),
                    Err(e) => write_json_line(&ResponseErr {
                        id: req.id,
                        ok: false,
                        error: ErrorBody {
                            message: e.to_string(),
                        },
                    }),
                }
                #[cfg(not(windows))]
                write_json_line(&ResponseOk {
                    id: req.id,
                    ok: true,
                    result: vec![RootItemDetailed {
                        path: "/".to_string(),
                        label: "/".to_string(),
                        icon_png_base64: None,
                    }],
                });
            }
            "get_icon_png_base64" => {
                let params: Result<GetIconParams, _> = serde_json::from_value(req.params);
                match params {
                    Ok(p) => {
                        #[cfg(windows)]
                        match icon_png_base64_for_path(&p.path, p.size) {
                            Ok(value) => write_json_line(&ResponseOk {
                                id: req.id,
                                ok: true,
                                result: value,
                            }),
                            Err(e) => write_json_line(&ResponseErr {
                                id: req.id,
                                ok: false,
                                error: ErrorBody {
                                    message: e.to_string(),
                                },
                            }),
                        }
                        #[cfg(not(windows))]
                        write_json_line(&ResponseOk {
                            id: req.id,
                            ok: true,
                            result: Option::<String>::None,
                        });
                    }
                    Err(e) => write_json_line(&ResponseErr {
                        id: req.id,
                        ok: false,
                        error: ErrorBody {
                            message: e.to_string(),
                        },
                    }),
                }
            }
            _ => write_json_line(&ResponseErr {
                id: req.id,
                ok: false,
                error: ErrorBody {
                    message: "unknown method".to_string(),
                },
            }),
        }
    }
}

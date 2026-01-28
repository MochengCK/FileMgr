use std::sync::OnceLock;

use windows::Win32::{
    Foundation::PROPERTYKEY,
    System::Com::{CoCreateInstance, CoTaskMemAlloc, CoTaskMemFree, IPersistFile, CLSCTX_INPROC_SERVER},
    System::Com::StructuredStorage::{PropVariantClear, PROPVARIANT},
    UI::Shell::{
        PropertiesSystem::{IPropertyStore, PSGetPropertyKeyFromName},
        ShellLink, IShellLinkW, SHGetKnownFolderPath, FOLDERID_Programs, KNOWN_FOLDER_FLAG,
    },
};

use windows::core::{Interface, PCWSTR, PWSTR};

static APP_USER_MODEL_ID: OnceLock<String> = OnceLock::new();

pub(crate) fn set_app_user_model_id(id: String) {
    let _ = APP_USER_MODEL_ID.set(id);
}

pub(crate) fn app_user_model_id() -> &'static str {
    APP_USER_MODEL_ID.get().map(|s| s.as_str()).unwrap_or("com.filemgr.app")
}

fn pwstr_z_to_string(ptr: PWSTR) -> String {
    if ptr.0.is_null() {
        return String::new();
    }
    unsafe {
        let mut len = 0usize;
        while *ptr.0.add(len) != 0 {
            len += 1;
        }
        let slice = std::slice::from_raw_parts(ptr.0, len);
        String::from_utf16_lossy(slice)
    }
}

pub(crate) fn ensure_start_menu_shortcut_windows(app_name: &str, exe_path: &str, app_id: &str) -> Result<(), String> {
    let _com = crate::com_init();
    unsafe {
        let programs = SHGetKnownFolderPath(&FOLDERID_Programs, KNOWN_FOLDER_FLAG(0), None).map_err(|e| e.to_string())?;
        let base = pwstr_z_to_string(programs);
        if !programs.0.is_null() {
            CoTaskMemFree(Some(programs.0 as _));
        }
        if base.trim().is_empty() {
            return Err("Programs 目录为空".to_string());
        }
        let lnk_path = format!("{}\\{}.lnk", base.trim_end_matches('\\'), app_name);
        let exe_wide = crate::to_wide_null(exe_path);
        let lnk_wide = crate::to_wide_null(&lnk_path);

        let link: IShellLinkW =
            CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER).map_err(|e: windows::core::Error| e.to_string())?;
        link.SetPath(PCWSTR(exe_wide.as_ptr())).map_err(|e: windows::core::Error| e.to_string())?;

        let store: IPropertyStore = link.cast().map_err(|e: windows::core::Error| e.to_string())?;
        let mut key = PROPERTYKEY::default();
        let key_name = crate::to_wide_null("System.AppUserModel.ID");
        if PSGetPropertyKeyFromName(PCWSTR(key_name.as_ptr()), &mut key).is_ok() {
            let id_wide = crate::to_wide_null(app_id);
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

        let pf: IPersistFile = link.cast().map_err(|e: windows::core::Error| e.to_string())?;
        pf.Save(PCWSTR(lnk_wide.as_ptr()), true)
            .map_err(|e: windows::core::Error| e.to_string())?;
        Ok(())
    }
}

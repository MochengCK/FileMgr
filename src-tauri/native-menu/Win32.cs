using System.Runtime.InteropServices;

namespace FileMgr.NativeMenu;

internal static class Win32
{
  internal const uint MF_STRING = 0x00000000;
  internal const uint MF_SEPARATOR = 0x00000800;
  internal const uint MF_POPUP = 0x00000010;
  internal const uint MF_GRAYED = 0x00000001;

  internal const uint TPM_RIGHTBUTTON = 0x0002;
  internal const uint TPM_RETURNCMD = 0x0100;

  internal const uint MIIM_BITMAP = 0x00000080;

  internal const int COLOR_MENUTEXT = 7;
  internal const uint WM_NULL = 0x0000;

  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
  internal struct MENUITEMINFO
  {
    public uint cbSize;
    public uint fMask;
    public uint fType;
    public uint fState;
    public uint wID;
    public nint hSubMenu;
    public nint hbmpChecked;
    public nint hbmpUnchecked;
    public nint dwItemData;
    public nint dwTypeData;
    public uint cch;
    public nint hbmpItem;
  }

  [DllImport("user32.dll", SetLastError = true)]
  internal static extern nint CreatePopupMenu();

  [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  internal static extern bool AppendMenuW(nint hMenu, uint uFlags, nuint uIDNewItem, string? lpNewItem);

  [DllImport("user32.dll", SetLastError = true)]
  internal static extern bool DestroyMenu(nint hMenu);

  [DllImport("user32.dll", SetLastError = true)]
  internal static extern nint TrackPopupMenuEx(nint hMenu, uint uFlags, int x, int y, nint hwnd, nint lptpm);

  [DllImport("user32.dll")]
  internal static extern nint GetForegroundWindow();

  [DllImport("user32.dll")]
  internal static extern bool SetForegroundWindow(nint hWnd);

  [DllImport("user32.dll", SetLastError = true)]
  internal static extern bool PostMessageW(nint hWnd, uint msg, nint wParam, nint lParam);

  [DllImport("user32.dll")]
  internal static extern uint GetSysColor(int nIndex);

  [DllImport("user32.dll", SetLastError = true)]
  internal static extern bool SetMenuItemInfoW(nint hMenu, uint item, bool fByPosition, ref MENUITEMINFO info);

  [DllImport("gdi32.dll", SetLastError = true)]
  internal static extern bool DeleteObject(nint hObject);
}


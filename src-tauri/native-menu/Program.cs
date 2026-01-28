using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text;
using System.Text.Json;

namespace FileMgr.NativeMenu;

internal static class Program
{
  private static readonly JsonSerializerOptions JsonOptions = new()
  {
    PropertyNameCaseInsensitive = true
  };

  private sealed class MenuBuildResult
  {
    public nint Menu { get; init; }
    public Dictionary<uint, string> CommandIdToItemId { get; init; } = new();
    public List<nint> Bitmaps { get; init; } = new();
  }

  private static async Task<int> Main()
  {
    try
    {
      Console.InputEncoding = Encoding.UTF8;
      Console.OutputEncoding = Encoding.UTF8;

      var inPath = "";
      var outPath = "";
      var args = Environment.GetCommandLineArgs().Skip(1).ToArray();
      for (var i = 0; i < args.Length; i++)
      {
        var a = args[i] ?? "";
        if (a == "--in" && i + 1 < args.Length)
        {
          inPath = args[i + 1] ?? "";
          i++;
          continue;
        }
        if (a.StartsWith("--in=", StringComparison.OrdinalIgnoreCase))
        {
          inPath = a.Substring("--in=".Length);
          continue;
        }
        if (a == "--out" && i + 1 < args.Length)
        {
          outPath = args[i + 1] ?? "";
          i++;
          continue;
        }
        if (a.StartsWith("--out=", StringComparison.OrdinalIgnoreCase))
        {
          outPath = a.Substring("--out=".Length);
          continue;
        }
      }

      if (inPath.Trim().Length > 0 && outPath.Trim().Length > 0)
      {
        var input = await File.ReadAllTextAsync(inPath, Encoding.UTF8);
        var req = JsonSerializer.Deserialize<NativeMenuRequest>(input, JsonOptions) ?? new NativeMenuRequest();
        var resp = await RunMenu(req);
        var json = JsonSerializer.Serialize(resp, JsonOptions);
        await File.WriteAllTextAsync(outPath, json, Encoding.UTF8);
        return 0;
      }

      var stdinInput = await Console.In.ReadToEndAsync();
      var stdinReq = JsonSerializer.Deserialize<NativeMenuRequest>(stdinInput, JsonOptions) ?? new NativeMenuRequest();

      var stdinResp = await RunMenu(stdinReq);
      await WriteResponse(stdinResp);

      return 0;
    }
    catch (Exception ex)
    {
      try
      {
        await WriteResponse(new NativeMenuResponse { SelectedId = null, Error = ex.Message });
      }
      catch
      {
      }
      return 1;
    }
  }

  private static Task<NativeMenuResponse> RunMenu(NativeMenuRequest req)
  {
    var build = BuildMenu(req.Items ?? new List<NativeMenuItem>());
    try
    {
      var owner = req.OwnerHwnd != 0 ? unchecked((nint)req.OwnerHwnd) : Win32.GetForegroundWindow();
      if (owner == 0) owner = Win32.GetForegroundWindow();
      Win32.SetForegroundWindow(owner);
      var cmd = Win32.TrackPopupMenuEx(build.Menu, Win32.TPM_RETURNCMD | Win32.TPM_RIGHTBUTTON, req.X, req.Y, owner, 0);
      Win32.PostMessageW(owner, Win32.WM_NULL, 0, 0);

      var cmdId = unchecked((uint)cmd);
      build.CommandIdToItemId.TryGetValue(cmdId, out var selectedId);
      return Task.FromResult(new NativeMenuResponse { SelectedId = selectedId });
    }
    finally
    {
      Win32.DestroyMenu(build.Menu);
      foreach (var h in build.Bitmaps)
      {
        if (h != 0) Win32.DeleteObject(h);
      }
    }
  }

  private static Task WriteResponse(NativeMenuResponse resp)
  {
    var json = JsonSerializer.Serialize(resp, JsonOptions);
    return Console.Out.WriteAsync(json);
  }

  private static MenuBuildResult BuildMenu(List<NativeMenuItem> items)
  {
    var menu = Win32.CreatePopupMenu();
    if (menu == 0) throw new InvalidOperationException("CreatePopupMenu 失败");

    var nextCmdId = 1u;
    var map = new Dictionary<uint, string>();
    var bitmaps = new List<nint>();
    AppendItems(menu, items, ref nextCmdId, map, bitmaps);
    return new MenuBuildResult { Menu = menu, CommandIdToItemId = map, Bitmaps = bitmaps };
  }

  private static void AppendItems(nint menu, List<NativeMenuItem> items, ref uint nextCmdId, Dictionary<uint, string> map, List<nint> bitmaps)
  {
    foreach (var it in items)
    {
      var kind = (it.Kind ?? "").Trim().ToLowerInvariant();
      if (kind == "sep" || kind == "separator")
      {
        Win32.AppendMenuW(menu, Win32.MF_SEPARATOR, 0, null);
        continue;
      }

      if (kind == "submenu")
      {
        var childMenu = Win32.CreatePopupMenu();
        if (childMenu == 0) continue;
        var children = it.Children ?? new List<NativeMenuItem>();
        AppendItems(childMenu, children, ref nextCmdId, map, bitmaps);
        var flags = Win32.MF_POPUP | (it.Enabled ? 0u : Win32.MF_GRAYED);
        var label = (it.Label ?? "").Trim();
        if (label.Length == 0) label = " ";
        Win32.AppendMenuW(menu, flags, unchecked((nuint)childMenu), label);
        continue;
      }

      if (kind == "item")
      {
        var id = (it.Id ?? "").Trim();
        var label = (it.Label ?? "").Trim();
        if (label.Length == 0) label = " ";
        if (id.Length == 0)
        {
          Win32.AppendMenuW(menu, Win32.MF_STRING | (it.Enabled ? 0u : Win32.MF_GRAYED), 0, label);
          continue;
        }

        var cmdId = nextCmdId;
        nextCmdId = nextCmdId == uint.MaxValue ? nextCmdId : nextCmdId + 1;
        map[cmdId] = id;
        Win32.AppendMenuW(menu, Win32.MF_STRING | (it.Enabled ? 0u : Win32.MF_GRAYED), cmdId, label);

        var bmp = TryRenderGlyphBitmap(it.Glyph);
        if (bmp != 0)
        {
          bitmaps.Add(bmp);
          var info = new Win32.MENUITEMINFO
          {
            cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf<Win32.MENUITEMINFO>(),
            fMask = Win32.MIIM_BITMAP,
            hbmpItem = bmp
          };
          Win32.SetMenuItemInfoW(menu, cmdId, false, ref info);
        }
      }
    }
  }

  private static nint TryRenderGlyphBitmap(string? glyph)
  {
    var g = (glyph ?? "").Trim();
    if (g.Length == 0) return 0;

    var colorRef = Win32.GetSysColor(Win32.COLOR_MENUTEXT);
    var color = Color.FromArgb(
      255,
      (int)(colorRef & 0xff),
      (int)((colorRef >> 8) & 0xff),
      (int)((colorRef >> 16) & 0xff)
    );

    using var bmp = new Bitmap(16, 16, PixelFormat.Format32bppArgb);
    using var gfx = Graphics.FromImage(bmp);
    gfx.Clear(Color.Transparent);
    gfx.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

    using var font = new Font("Segoe MDL2 Assets", 12.0f, FontStyle.Regular, GraphicsUnit.Pixel);
    using var brush = new SolidBrush(color);
    using var fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
    gfx.DrawString(g, font, brush, new RectangleF(0, 0, 16, 16), fmt);

    var hBmp = bmp.GetHbitmap(Color.Transparent);
    return hBmp;
  }
}

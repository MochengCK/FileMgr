using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace FileMgr.NativeMenu;

public static class NativeMenuBuilder
{
  public static MenuFlyout Build(List<NativeMenuItem> items, Action<string> onPick)
  {
    var flyout = new MenuFlyout();
    foreach (var it in items ?? new List<NativeMenuItem>())
    {
      var el = BuildItem(it, onPick);
      if (el != null) flyout.Items.Add(el);
    }
    return flyout;
  }

  private static MenuFlyoutItemBase? BuildItem(NativeMenuItem it, Action<string> onPick)
  {
    var kind = (it.Kind ?? "").Trim();
    if (kind == "sep" || kind == "separator")
    {
      return new MenuFlyoutSeparator();
    }

    if (kind == "submenu")
    {
      var sub = new MenuFlyoutSubItem
      {
        Text = (it.Label ?? "").Trim(),
        IsEnabled = it.Enabled
      };
      var icon = CreateIcon(it.Glyph);
      if (icon != null) sub.Icon = icon;

      foreach (var child in it.Children ?? new List<NativeMenuItem>())
      {
        var el = BuildItem(child, onPick);
        if (el != null) sub.Items.Add(el);
      }
      return sub;
    }

    if (kind == "item")
    {
      var id = (it.Id ?? "").Trim();
      if (id.Length == 0) return null;
      var mi = new MenuFlyoutItem
      {
        Text = (it.Label ?? "").Trim(),
        IsEnabled = it.Enabled
      };
      var icon = CreateIcon(it.Glyph);
      if (icon != null) mi.Icon = icon;
      mi.Click += (_, __) => onPick(id);
      return mi;
    }

    return null;
  }

  private static IconElement? CreateIcon(string? glyph)
  {
    var g = (glyph ?? "").Trim();
    if (g.Length == 0) return null;
    return new FontIcon
    {
      Glyph = g,
      FontFamily = new FontFamily("Segoe MDL2 Assets"),
      FontSize = 14
    };
  }
}


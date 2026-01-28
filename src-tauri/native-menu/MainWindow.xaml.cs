using Microsoft.UI.Text;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using System.Text.Json;
using Windows.Foundation;
using Windows.Graphics;

namespace FileMgr.NativeMenu;

public sealed partial class MainWindow : Window
{
  private static readonly JsonSerializerOptions JsonOptions = new()
  {
    PropertyNameCaseInsensitive = true
  };

  private int _written;

  public MainWindow()
  {
    InitializeComponent();
    Loaded += OnLoaded;
  }

  private async void OnLoaded(object sender, RoutedEventArgs e)
  {
    try
    {
      var inPath = (App.InPath ?? "").Trim();
      var outPath = (App.OutPath ?? "").Trim();
      if (inPath.Length == 0 || outPath.Length == 0)
      {
        await WriteResponse(outPath, new NativeMenuResponse { SelectedId = null, Error = "参数错误" });
        Application.Current.Exit();
        return;
      }

      NativeMenuRequest? req = null;
      try
      {
        var json = await File.ReadAllTextAsync(inPath);
        req = JsonSerializer.Deserialize<NativeMenuRequest>(json, JsonOptions);
      }
      catch (Exception ex)
      {
        await WriteResponse(outPath, new NativeMenuResponse { SelectedId = null, Error = ex.Message });
        Application.Current.Exit();
        return;
      }

      if (req == null)
      {
        await WriteResponse(outPath, new NativeMenuResponse { SelectedId = null, Error = "请求为空" });
        Application.Current.Exit();
        return;
      }

      ConfigureWindow(req.X, req.Y);

      var flyout = NativeMenuBuilder.Build(req.Items ?? new List<NativeMenuItem>(), OnPick);
      flyout.Closed += async (_, __) =>
      {
        if (System.Threading.Interlocked.Exchange(ref _written, 1) != 0) return;
        await WriteResponse(outPath, new NativeMenuResponse { SelectedId = null, Error = null });
        Application.Current.Exit();
      };

      var opts = new FlyoutShowOptions
      {
        Position = new Point(0, 0),
        Placement = FlyoutPlacementMode.BottomEdgeAlignedLeft
      };
      flyout.ShowAt(Root, opts);

      async void OnPick(string id)
      {
        if (System.Threading.Interlocked.Exchange(ref _written, 1) != 0) return;
        await WriteResponse(outPath, new NativeMenuResponse { SelectedId = id, Error = null });
        Application.Current.Exit();
      }
    }
    catch (Exception ex)
    {
      var outPath = (App.OutPath ?? "").Trim();
      if (System.Threading.Interlocked.Exchange(ref _written, 1) == 0)
      {
        await WriteResponse(outPath, new NativeMenuResponse { SelectedId = null, Error = ex.Message });
      }
      Application.Current.Exit();
    }
  }

  private void ConfigureWindow(int x, int y)
  {
    var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
    var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
    var appWindow = AppWindow.GetFromWindowId(windowId);
    if (appWindow == null) return;

    if (appWindow.Presenter is OverlappedPresenter p)
    {
      p.IsResizable = false;
      p.IsMaximizable = false;
      p.IsMinimizable = false;
      p.IsAlwaysOnTop = true;
    }

    appWindow.Move(new PointInt32(x, y));
    appWindow.Resize(new SizeInt32(1, 1));
  }

  private static async Task WriteResponse(string outPath, NativeMenuResponse resp)
  {
    if (outPath.Trim().Length == 0) return;
    var json = JsonSerializer.Serialize(resp, JsonOptions);
    var dir = Path.GetDirectoryName(outPath);
    if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
    await File.WriteAllTextAsync(outPath, json);
  }
}


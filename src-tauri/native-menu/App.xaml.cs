using Microsoft.UI.Xaml;

namespace FileMgr.NativeMenu;

public sealed partial class App : Application
{
  internal static string? InPath { get; set; }
  internal static string? OutPath { get; set; }

  private Window? _window;

  public App()
  {
    InitializeComponent();
  }

  protected override void OnLaunched(LaunchActivatedEventArgs args)
  {
    _window = new MainWindow();
    _window.Activate();
  }
}


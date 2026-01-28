using System.Text.Json.Serialization;

namespace FileMgr.NativeMenu;

public sealed class NativeMenuRequest
{
  [JsonPropertyName("x")]
  public int X { get; set; }

  [JsonPropertyName("y")]
  public int Y { get; set; }

  [JsonPropertyName("owner_hwnd")]
  public ulong OwnerHwnd { get; set; }

  [JsonPropertyName("items")]
  public List<NativeMenuItem> Items { get; set; } = new();
}

public sealed class NativeMenuItem
{
  [JsonPropertyName("kind")]
  public string Kind { get; set; } = "";

  [JsonPropertyName("id")]
  public string? Id { get; set; }

  [JsonPropertyName("label")]
  public string? Label { get; set; }

  [JsonPropertyName("enabled")]
  public bool Enabled { get; set; } = true;

  [JsonPropertyName("glyph")]
  public string? Glyph { get; set; }

  [JsonPropertyName("children")]
  public List<NativeMenuItem>? Children { get; set; }
}

public sealed class NativeMenuResponse
{
  [JsonPropertyName("selectedId")]
  public string? SelectedId { get; set; }

  [JsonPropertyName("error")]
  public string? Error { get; set; }
}

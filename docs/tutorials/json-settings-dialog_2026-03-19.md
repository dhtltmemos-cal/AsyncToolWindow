# Hướng dẫn: JSON Settings Dialog (§JsonSettings)

**Tính năng:** JSON editor dialog mở từ Tools menu  
**Ngày:** 2026-03-19  
**Phiên bản:** `feat: json-settings-dialog`

---

## Mở dialog

**Tools → Async Tool Window Sample Settings (JSON)...**

---

## Tính năng

| Nút | Chức năng |
|-----|-----------|
| ✨ Format JSON | Reformat với indentation chuẩn, giữ caret |
| ✔ Validate | Kiểm tra cú pháp JSON, hiện kết quả inline |
| ↩ Reset Default | Load lại JSON mặc định vào editor |
| 🔄 Reload from File | Tải lại từ file, bỏ thay đổi chưa lưu |
| 📋 Copy | Copy toàn bộ JSON vào clipboard |
| 📌 Paste | Paste từ clipboard + auto-format nếu valid |
| 💾 Save | Validate → lưu file → log key thay đổi → đóng |

---

## Live feedback khi gõ

- **Status bar màu xanh lá** = JSON hợp lệ
- **Status bar màu cam** = cú pháp lỗi / braces mismatch
- **Char/line counter** cập nhật realtime góc phải

---

## Khi Save — Log ra Output Window

```
[JsonConfig] ── Save ──────────────────────────
[JsonConfig] 3 key(s) changed:
[JsonConfig]   ServerUrl: "https://old.com" → "https://new.com"
[JsonConfig]   TimeoutSeconds: "30" → "60"
[JsonConfig]   Advanced.RetryCount: "3" → "5"
[JsonConfig] Saved → C:\Users\...\AsyncToolWindowSample.json
```

---

## Sử dụng JsonConfigService trong code

```csharp
// Typed properties
string url  = JsonConfig.ServerUrl;
int timeout = JsonConfig.TimeoutSeconds;
bool debug  = JsonConfig.DebugMode;

// Dot-path cho nested
int retry = JsonConfig.Get<int>("Advanced.RetryCount", 3);
string ua = JsonConfig.Get<string>("Advanced.UserAgent", "");

// Full JObject snapshot
JObject snap = JsonConfig.GetSnapshot();

// Subscribe khi save
JsonConfig.ConfigSaved += changedKeys =>
{
    foreach (string key in changedKeys)
        Console.WriteLine($"Changed: {key}");
};
```

---

## Cấu trúc JSON mặc định

```json
{
  "ServerUrl": "https://api.example.com/v1",
  "ApiKey": "",
  "TimeoutSeconds": 30,
  "MaxResults": 50,
  "RefreshInterval": 5,
  "OutputFormat": "JSON",
  "EnableLogging": true,
  "DebugMode": false,
  "ShowStatusBar": true,
  "Tags": ["vsix", "sample"],
  "Advanced": {
    "RetryCount": 3,
    "RetryDelayMs": 500,
    "UserAgent": "AsyncToolWindowSample/1.0"
  }
}
```

File lưu tại: `%AppData%\AsyncToolWindowSample\AsyncToolWindowSample.json`

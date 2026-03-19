using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Task = System.Threading.Tasks.Task;

namespace AsyncToolWindowSample.Services
{
    // ======================================================================//
    //  Result DTOs — thay thế ValueTuple để tương thích C# 7.0 / .NET 4.6  //
    // ======================================================================//

    /// <summary>Kết quả Parse JSON.</summary>
    public sealed class JsonParseResult
    {
        public bool    Valid  { get; set; }
        public JObject Obj    { get; set; }
        public string  Error  { get; set; }

        public static JsonParseResult Ok(JObject obj)
            => new JsonParseResult { Valid = true, Obj = obj };

        public static JsonParseResult Fail(string error)
            => new JsonParseResult { Valid = false, Error = error };
    }

    /// <summary>Kết quả Format JSON.</summary>
    public sealed class JsonFormatResult
    {
        public string Result { get; set; }
        public string Error  { get; set; }

        public bool IsOk => Error == null;

        public static JsonFormatResult Ok(string result)
            => new JsonFormatResult { Result = result };

        public static JsonFormatResult Fail(string error)
            => new JsonFormatResult { Error = error };
    }

    /// <summary>Kết quả Save JSON.</summary>
    public sealed class JsonSaveResult
    {
        public bool                  Success     { get; set; }
        public IReadOnlyList<string> ChangedKeys { get; set; }
        public string                Error       { get; set; }

        public static JsonSaveResult Ok(IReadOnlyList<string> changedKeys)
            => new JsonSaveResult { Success = true, ChangedKeys = changedKeys };

        public static JsonSaveResult Fail(string error)
            => new JsonSaveResult
            {
                Success     = false,
                ChangedKeys = new List<string>(),
                Error       = error
            };
    }

    /// <summary>Kết quả Reset.</summary>
    public sealed class JsonResetResult
    {
        public bool   Success { get; set; }
        public string Error   { get; set; }

        public static JsonResetResult Ok()            => new JsonResetResult { Success = true };
        public static JsonResetResult Fail(string e)  => new JsonResetResult { Error = e };
    }

    // ======================================================================//
    //  JsonConfigService                                                      //
    // ======================================================================//

    /// <summary>
    /// Quản lý cấu hình extension dưới dạng JSON object.
    ///   - Lưu/đọc file JSON trong %AppData%\AsyncToolWindowSample\
    ///   - Lấy từng key bằng typed getter hoặc dot-path
    ///   - Phát hiện và log các key thay đổi khi Save
    ///   - Parse, validate và format JSON
    ///   - Tương thích C# 7.0 / .NET 4.6 (không dùng ValueTuple)
    /// All public methods must be called on the UI thread.
    /// </summary>
    public sealed class JsonConfigService
    {
        // ------------------------------------------------------------------ //
        //  Constants                                                           //
        // ------------------------------------------------------------------ //

        private const string ConfigFileName = "AsyncToolWindowSample.json";

        /// <summary>Schema mặc định — dùng khi file chưa tồn tại.</summary>
        public static readonly string DefaultJson = JsonConvert.SerializeObject(
            new
            {
                ServerUrl       = "https://api.example.com/v1",
                ApiKey          = "",
                TimeoutSeconds  = 30,
                MaxResults      = 50,
                RefreshInterval = 5,
                OutputFormat    = "JSON",
                EnableLogging   = true,
                DebugMode       = false,
                ShowStatusBar   = true,
                Tags            = new[] { "vsix", "sample" },
                Advanced        = new
                {
                    RetryCount   = 3,
                    RetryDelayMs = 500,
                    UserAgent    = "AsyncToolWindowSample/1.0"
                }
            },
            Formatting.Indented);

        // ------------------------------------------------------------------ //
        //  Fields                                                              //
        // ------------------------------------------------------------------ //

        private readonly AsyncPackage        _package;
        private readonly OutputWindowService _outputWindow;

        private string  _configFilePath;
        private JObject _current;

        // ------------------------------------------------------------------ //
        //  Events                                                              //
        // ------------------------------------------------------------------ //

        /// <summary>Fired sau khi Save thành công. Payload: danh sách key thay đổi.</summary>
        public event Action<IReadOnlyList<string>> ConfigSaved;

        // ------------------------------------------------------------------ //
        //  Constructor                                                         //
        // ------------------------------------------------------------------ //

        public JsonConfigService(AsyncPackage package, OutputWindowService outputWindow)
        {
            _package      = package      ?? throw new ArgumentNullException(nameof(package));
            _outputWindow = outputWindow ?? throw new ArgumentNullException(nameof(outputWindow));
        }

        // ================================================================== //
        //  Initialization                                                      //
        // ================================================================== //

        public async Task InitializeAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder  = Path.Combine(appData, "AsyncToolWindowSample");
            Directory.CreateDirectory(folder);
            _configFilePath = Path.Combine(folder, ConfigFileName);

            _current = LoadFromFile();
            _outputWindow.Log($"[JsonConfig] Initialized. File: {_configFilePath}");
        }

        // ================================================================== //
        //  Public read API                                                     //
        // ================================================================== //

        /// <summary>Trả về bản sao JObject hiện tại (immutable snapshot).</summary>
        public JObject GetSnapshot() => (JObject)_current.DeepClone();

        /// <summary>
        /// Lấy giá trị theo dot-path, ví dụ "Advanced.RetryCount".
        /// Dùng overload với explicit fallback để tránh 'default literal' (CS8107).
        /// </summary>
        public T Get<T>(string dotPath, T fallback)
        {
            try
            {
                JToken token = ResolvePath(_current, dotPath);
                if (token == null) return fallback;
                return token.ToObject<T>();
            }
            catch
            {
                return fallback;
            }
        }

        // ── Typed convenience properties ──────────────────────────────────
        public string ServerUrl       => Get<string>("ServerUrl",       "https://api.example.com/v1");
        public string ApiKey          => Get<string>("ApiKey",          "");
        public int    TimeoutSeconds  => Get<int>   ("TimeoutSeconds",  30);
        public int    MaxResults      => Get<int>   ("MaxResults",      50);
        public int    RefreshInterval => Get<int>   ("RefreshInterval", 5);
        public string OutputFormat    => Get<string>("OutputFormat",    "JSON");
        public bool   EnableLogging   => Get<bool>  ("EnableLogging",   true);
        public bool   DebugMode       => Get<bool>  ("DebugMode",       false);
        public bool   ShowStatusBar   => Get<bool>  ("ShowStatusBar",   true);

        /// <summary>Trả về chuỗi JSON hiện tại (formatted).</summary>
        public string GetCurrentJson(Formatting formatting = Formatting.Indented)
            => _current.ToString(formatting);

        /// <summary>Đường dẫn file config.</summary>
        public string ConfigFilePath => _configFilePath ?? "(not initialized)";

        // ================================================================== //
        //  Validate                                                            //
        // ================================================================== //

        /// <summary>
        /// Kiểm tra chuỗi JSON có hợp lệ không.
        /// Trả về JsonParseResult (không dùng ValueTuple).
        /// </summary>
        public JsonParseResult TryParse(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return JsonParseResult.Fail("JSON string is empty.");
            try
            {
                JObject obj = JObject.Parse(json);
                return JsonParseResult.Ok(obj);
            }
            catch (JsonException ex)
            {
                return JsonParseResult.Fail("JSON parse error: " + ex.Message);
            }
        }

        // ================================================================== //
        //  Format                                                              //
        // ================================================================== //

        /// <summary>
        /// Format lại chuỗi JSON với indentation chuẩn.
        /// Trả về JsonFormatResult (không dùng ValueTuple).
        /// </summary>
        public JsonFormatResult FormatJson(string json)
        {
            JsonParseResult parsed = TryParse(json);
            if (!parsed.Valid)
                return JsonFormatResult.Fail(parsed.Error);
            return JsonFormatResult.Ok(parsed.Obj.ToString(Formatting.Indented));
        }

        // ================================================================== //
        //  Save                                                               //
        // ================================================================== //

        /// <summary>
        /// Lưu JSON mới vào file. Detect và log các key thay đổi.
        /// Trả về JsonSaveResult (không dùng ValueTuple).
        /// </summary>
        public JsonSaveResult Save(string newJson)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // 1. Parse
            JsonParseResult parsed = TryParse(newJson);
            if (!parsed.Valid)
                return JsonSaveResult.Fail(parsed.Error);

            JObject newObj = parsed.Obj;

            // 2. Diff
            List<string> changedKeys = DiffKeys(_current, newObj);

            // 3. Log diff
            _outputWindow.Log("[JsonConfig] ── Save ──────────────────────────");
            if (changedKeys.Count == 0)
            {
                _outputWindow.Log("[JsonConfig] No changes detected.");
            }
            else
            {
                _outputWindow.Log("[JsonConfig] " + changedKeys.Count + " key(s) changed:");
                foreach (string key in changedKeys)
                {
                    string oldVal = GetFlatValue(_current, key);
                    string newVal = GetFlatValue(newObj,   key);

                    // Mask ApiKey
                    if (key.IndexOf("ApiKey", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        oldVal = Mask(oldVal);
                        newVal = Mask(newVal);
                    }
                    _outputWindow.Log("[JsonConfig]   " + key + ": \"" + oldVal + "\" → \"" + newVal + "\"");
                }
            }

            // 4. Write file
            try
            {
                File.WriteAllText(
                    _configFilePath,
                    newObj.ToString(Formatting.Indented),
                    Encoding.UTF8);
                _outputWindow.Log("[JsonConfig] Saved → " + _configFilePath);
            }
            catch (Exception ex)
            {
                string msg = "[JsonConfig] ERROR writing file: " + ex.Message;
                _outputWindow.Log(msg);
                return JsonSaveResult.Fail(msg);
            }

            // 5. Update in-memory
            _current = newObj;

            // 6. Fire event
            if (ConfigSaved != null)
                ConfigSaved(changedKeys);

            return JsonSaveResult.Ok(changedKeys);
        }

        // ================================================================== //
        //  Reset                                                               //
        // ================================================================== //

        /// <summary>Reset về DefaultJson và lưu file.</summary>
        public JsonResetResult ResetToDefault()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _outputWindow.Log("[JsonConfig] Resetting to default...");
            JsonSaveResult result = Save(DefaultJson);
            return result.Success
                ? JsonResetResult.Ok()
                : JsonResetResult.Fail(result.Error);
        }

        // ================================================================== //
        //  Private helpers                                                     //
        // ================================================================== //

        private JObject LoadFromFile()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!File.Exists(_configFilePath))
            {
                _outputWindow.Log("[JsonConfig] File not found – using defaults.");
                try { File.WriteAllText(_configFilePath, DefaultJson, Encoding.UTF8); }
                catch { /* ignore */ }
                return JObject.Parse(DefaultJson);
            }

            try
            {
                string raw = File.ReadAllText(_configFilePath, Encoding.UTF8);
                JObject obj = JObject.Parse(raw);
                _outputWindow.Log("[JsonConfig] Loaded from file.");
                return obj;
            }
            catch (Exception ex)
            {
                _outputWindow.Log("[JsonConfig] ERROR loading: " + ex.Message + " – using defaults.");
                return JObject.Parse(DefaultJson);
            }
        }

        /// <summary>
        /// Flatten cả hai JObject thành key→value (dot-path),
        /// so sánh và trả về list key khác nhau.
        /// </summary>
        private static List<string> DiffKeys(JObject oldObj, JObject newObj)
        {
            Dictionary<string, string> oldFlat = Flatten(oldObj, "");
            Dictionary<string, string> newFlat = Flatten(newObj, "");
            var changed = new List<string>();

            // Keys có trong new (thêm hoặc thay đổi)
            foreach (KeyValuePair<string, string> kv in newFlat)
            {
                string oldVal;
                bool exists = oldFlat.TryGetValue(kv.Key, out oldVal);
                if (!exists || oldVal != kv.Value)
                    changed.Add(kv.Key);
            }

            // Keys bị xóa khỏi new
            foreach (KeyValuePair<string, string> kv in oldFlat)
            {
                if (!newFlat.ContainsKey(kv.Key))
                    changed.Add(kv.Key + " (removed)");
            }

            return changed;
        }

        /// <summary>Flatten JObject thành dictionary "Parent.Child" → "value".</summary>
        private static Dictionary<string, string> Flatten(JObject obj, string prefix)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (JProperty prop in obj.Properties())
            {
                string key = string.IsNullOrEmpty(prefix)
                    ? prop.Name
                    : prefix + "." + prop.Name;

                JObject childObj = prop.Value as JObject;
                JArray  childArr = prop.Value as JArray;

                if (childObj != null)
                {
                    // Recurse into nested object
                    foreach (KeyValuePair<string, string> kv in Flatten(childObj, key))
                        result[kv.Key] = kv.Value;
                }
                else if (childArr != null)
                {
                    result[key] = childArr.ToString(Formatting.None);
                }
                else
                {
                    result[key] = prop.Value != null ? prop.Value.ToString() : "null";
                }
            }

            return result;
        }

        private static string GetFlatValue(JObject obj, string flatKey)
        {
            // flatKey có thể kết thúc bằng " (removed)"
            string clean = flatKey.Replace(" (removed)", "");
            Dictionary<string, string> flat = Flatten(obj, "");
            string val;
            return flat.TryGetValue(clean, out val) ? val : "(deleted)";
        }

        private static JToken ResolvePath(JObject obj, string dotPath)
        {
            string[] parts = dotPath.Split('.');
            JToken current = obj;
            foreach (string part in parts)
            {
                JObject jo = current as JObject;
                if (jo == null) return null;
                JToken next;
                if (!jo.TryGetValue(part, StringComparison.OrdinalIgnoreCase, out next))
                    return null;
                current = next;
            }
            return current;
        }

        private static string Mask(string s)
        {
            if (string.IsNullOrEmpty(s)) return "(empty)";
            if (s.Length <= 4)           return "****";
            return "****" + s.Substring(s.Length - 4);
        }
    }
}

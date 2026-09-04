using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace ML.ImportExcel.Core.Localization
{
    /// <summary>
    /// 多语言消息目录：每个语系一个 JSON 文件（{"key":"template"}）。
    /// 默认会把内置 en / zh 词条**自动生成到“运行目录\\ExcelLang”**，方便发布后直接
    /// 修改更新或新增语系；也可通过 <see cref="ExcelImportOptions.LanguageDirectory"/>
    /// 指定其它目录（同样会生成内置词条）。目录下形如 {culture}.json 的文件会覆盖内置词条。
    /// 语言解析：支持 "zh-CN"/"zh-Hans"/"en-US" 等带区域写法；找不到指定语系回退默认英语(en)。
    /// </summary>
    public sealed class MessageCatalog
    {
        public const string DefaultLanguage = "en";
        public const string DefaultDirectoryName = "ExcelLang";
        private const string ResourcePrefix = "ML.ImportExcel.Lang.";
        private const string ResourceSuffix = ".json";

        private static readonly object SyncRoot = new object();
        private readonly Dictionary<string, Dictionary<string, string>> _langs =
            new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>实际生效的语言包目录（生成/读取目录）。</summary>
        public string EffectiveLanguageDirectory { get; }

        public MessageCatalog(string externalLanguageDirectory = null)
        {
            LoadEmbeddedLanguages();

            // 未指定目录时：默认“运行目录\ExcelLang”（程序集所在目录下，开发时即 bin 输出目录）
            EffectiveLanguageDirectory = string.IsNullOrWhiteSpace(externalLanguageDirectory)
                ? Path.Combine(AppContext.BaseDirectory, DefaultDirectoryName)
                : externalLanguageDirectory;

            // 把内置 en/zh 词条生成到该目录（仅当文件缺失时写出，用户后续可直接修改）
            EnsureDefaultLanguageFiles(EffectiveLanguageDirectory);

            LoadExternalLanguages(EffectiveLanguageDirectory);

            // 保证至少存在默认英语
            if (!_langs.ContainsKey(DefaultLanguage))
                _langs[DefaultLanguage] = new Dictionary<string, string>(StringComparer.Ordinal);
        }

        /// <summary>把内置语言包(en/zh…)写出到指定目录（文件已存在则跳过，不影响用户修改）。</summary>
        private static void EnsureDefaultLanguageFiles(string directory)
        {
            lock (SyncRoot)
            {
                try { Directory.CreateDirectory(directory); }
                catch { return; }

                var asm = typeof(MessageCatalog).Assembly;
                foreach (var name in asm.GetManifestResourceNames())
                {
                    if (!name.StartsWith(ResourcePrefix, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!name.EndsWith(ResourceSuffix, StringComparison.OrdinalIgnoreCase)) continue;

                    // 资源名形如 ML.ImportExcel.Lang.en.json → 文件 en.json
                    var fileName = name.Substring(ResourcePrefix.Length);
                    var langFile = Path.Combine(directory, fileName);
                    if (File.Exists(langFile)) continue;

                    try
                    {
                        using (var s = asm.GetManifestResourceStream(name))
                        using (var r = new StreamReader(s, System.Text.Encoding.UTF8))
                        {
                            var content = r.ReadToEnd();
                            if (!string.IsNullOrWhiteSpace(content))
                                File.WriteAllText(langFile, content, System.Text.Encoding.UTF8);
                        }
                    }
                    catch
                    {
                        // 写入失败不阻塞（仍可使用内置词条）
                    }
                }
            }
        }

        private void LoadEmbeddedLanguages()
        {
            var asm = typeof(MessageCatalog).Assembly;
            foreach (var name in asm.GetManifestResourceNames())
            {
                if (!name.StartsWith(ResourcePrefix, StringComparison.OrdinalIgnoreCase)) continue;
                if (!name.EndsWith(ResourceSuffix, StringComparison.OrdinalIgnoreCase)) continue;

                var lang = name.Substring(ResourcePrefix.Length,
                    name.Length - ResourcePrefix.Length - ResourceSuffix.Length);
                if (string.IsNullOrEmpty(lang)) continue;

                using (var s = asm.GetManifestResourceStream(name))
                {
                    if (s == null) continue;
                    MergeFromJson(lang, s);
                }
            }
        }

        private void LoadExternalLanguages(string directory)
        {
            if (!Directory.Exists(directory)) return;
            foreach (var file in Directory.GetFiles(directory, "*.json", SearchOption.TopDirectoryOnly))
            {
                var lang = Path.GetFileNameWithoutExtension(file);
                if (string.IsNullOrEmpty(lang)) continue;
                using (var s = File.OpenRead(file))
                    MergeFromJson(lang, s);
            }
        }

        private void MergeFromJson(string lang, Stream stream)
        {
            try
            {
                // 说明：JsonSerializer.Deserialize<T>(Stream) 同步重载是 net6 新增，net5 需先读成字符串
                string text;
                using (var reader = new StreamReader(stream))
                    text = reader.ReadToEnd();

                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(text);
                if (dict == null || dict.Count == 0) return;

                if (!_langs.TryGetValue(lang, out var target))
                {
                    target = new Dictionary<string, string>(StringComparer.Ordinal);
                    _langs[lang] = target;
                }
                foreach (var kv in dict)
                    if (!string.IsNullOrEmpty(kv.Key))
                        target[kv.Key] = kv.Value ?? "";
            }
            catch
            {
                // 单个语言包解析失败不影响其它语言
            }
        }

        /// <summary>把调用方语系归一化为已加载语言；未知语系回退默认英语。</summary>
        public string ResolveLanguage(string requested)
        {
            if (string.IsNullOrWhiteSpace(requested)) return DefaultLanguage;
            var r = requested.Trim();
            if (_langs.ContainsKey(r)) return r;

            // 带区域写法：zh-CN → zh；en-GB → en
            var dash = r.IndexOf('-');
            if (dash > 0)
            {
                var baseLang = r.Substring(0, dash);
                if (_langs.ContainsKey(baseLang)) return baseLang;
            }

            // 常见别名
            if (r.StartsWith("zh", StringComparison.OrdinalIgnoreCase) && _langs.ContainsKey("zh")) return "zh";
            if (r.StartsWith("en", StringComparison.OrdinalIgnoreCase) && _langs.ContainsKey("en")) return "en";

            return DefaultLanguage;
        }

        /// <summary>是否已加载指定语系。</summary>
        public bool HasLanguage(string language) => _langs.ContainsKey(ResolveLanguage(language));

        /// <summary>按语系取文案并用参数格式化；词条缺失时回退默认英语；仍无则返回 key。</summary>
        public string Get(string language, string key, params object[] args)
        {
            var resolved = ResolveLanguage(language);
            if (_langs.TryGetValue(resolved, out var dict) && dict.TryGetValue(key, out var template))
                return Format(template, args);

            // 词条缺失 → 回退默认英语
            if (!string.Equals(resolved, DefaultLanguage, StringComparison.OrdinalIgnoreCase)
                && _langs.TryGetValue(DefaultLanguage, out var en) && en.TryGetValue(key, out var enTemplate))
                return Format(enTemplate, args);

            return key;
        }

        private static string Format(string template, object[] args)
        {
            try
            {
                return args == null || args.Length == 0
                    ? template
                    : string.Format(CultureInfo.InvariantCulture, template, args);
            }
            catch
            {
                return template;
            }
        }

        /// <summary>已加载的语言代码列表（如 en、zh、…）。</summary>
        public IReadOnlyList<string> AvailableLanguages => _langs.Keys.ToList();
    }
}

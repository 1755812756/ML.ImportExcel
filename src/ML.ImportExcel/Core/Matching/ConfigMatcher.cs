using System;
using System.Collections.Generic;
using System.Linq;
using ML.ImportExcel.Core.Models;

namespace ML.ImportExcel.Core.Matching
{
    /// <summary>自动匹配候选结果。</summary>
    public sealed class MatchCandidate
    {
        public ImportConfig Config { get; set; }

        /// <summary>命中率 = 匹配到的模板列数 / 模板列总数。</summary>
        public double Score { get; set; }

        /// <summary>模板（配置记录的列头）总列数。</summary>
        public int Expected { get; set; }

        /// <summary>命中的模板列数。</summary>
        public int Matched { get; set; }

        /// <summary>命中的“已绑定实体字段”列数（数据能否被充分利用）。</summary>
        public int BoundMatched { get; set; }

        /// <summary>命中的列在文件中是否保持与模板一致的相对顺序。</summary>
        public bool OrderMatch { get; set; }
    }

    /// <summary>
    /// 依据“Excel 列头”自动匹配导入配置（实体）。
    /// 匹配签名 = 配置中记录的“全部 Excel 列头”（含未绑定实体字段的列），因为模板特征列
    /// 最能代表“这是哪张报表”。评分优先级：
    ///   1) 命中率高的优先；
    ///   2) 命中率相同 → 命中列数多的（模板更具体/超集）优先——避免 A B C D 模板
    ///      被 A B C D E / A B C D E F 文件误匹配；
    ///   3) 仍相同 → 命中“已绑定字段”列更多的优先（数据利用率更高）；
    ///   4) 仍相同 → 命中列在文件中与模板顺序一致的优先（区分同列集不同顺序的报表）；
    ///   5) 仍相同 → Id 更小者。
    /// 阈值判定（命中率 &lt; MatchThreshold 视为不匹配）由调用方在服务中执行。
    /// </summary>
    public static class ConfigMatcher
    {
        /// <summary>
        /// 列头归一化：半角转换 + 转小写 + 去除空白、标点与括号注释（保留字母/数字/中日韩文字）。
        /// 兼容全角半角、大小写、空格、零宽字符以及标点/括号写法差异
        /// （如 “入学日期（必填）” 与 “入学日期”、“ＡＢＣ” 与 “ABC” 视为一致）。
        /// </summary>
        public static string Normalize(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var t = RemoveBracketed(s.Trim());
            if (t.Length == 0) return "";
            var sb = new System.Text.StringBuilder(t.Length);
            foreach (var raw in t)
            {
                var ch = raw;
                // 全角 → 半角（仅对可映射的 ASCII 区段）
                if (ch >= '\uFF01' && ch <= '\uFF5E')
                    ch = (char)(ch - 0xFEE0);

                if (ch == ' ' || ch == '\u3000' || ch == '\t' || ch == '\r' || ch == '\n' || ch == '\u200B' || ch == '\uFEFF')
                    continue;

                // 仅保留字母/数字/中日韩等文字，其它标点一律忽略
                if (!char.IsLetterOrDigit(ch)) continue;

                sb.Append(char.ToLowerInvariant(ch));
            }
            return sb.ToString();
        }

        /// <summary>去除括号注释内容（中文/英文圆括号、方括号、书名号式括注），如 “入学日期（必填）”→ “入学日期”。</summary>
        private static string RemoveBracketed(string s)
        {
            var pairs = new[]
            {
                (open: '(', close: ')'),
                (open: '（', close: '）'),
                (open: '[', close: ']'),
                (open: '【', close: '】')
            };
            while (true)
            {
                var bestOpen = -1;
                char bestClose = '\0';
                foreach (var p in pairs)
                {
                    var idx = s.LastIndexOf(p.open);
                    if (idx > bestOpen)
                    {
                        bestOpen = idx;
                        bestClose = p.close;
                    }
                }
                if (bestOpen < 0) return s;
                var closeAt = s.IndexOf(bestClose, bestOpen + 1);
                if (closeAt < 0) return s;
                s = s.Remove(bestOpen, closeAt - bestOpen + 1);
            }
        }

        /// <summary>
        /// 在全部配置中选出与给定文件列头最匹配的一个。
        /// fileHeaders 为文件中实际的列头文本（按列顺序）列表。
        /// 返回 null 表示没有可参与匹配的配置。
        /// </summary>
        public static MatchCandidate BestMatch(IEnumerable<ImportConfig> configs, IEnumerable<string> fileHeaders)
        {
            if (configs == null || fileHeaders == null) return null;

            var filePos = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var idx = 0;
            foreach (var h in fileHeaders)
            {
                var n = Normalize(h);
                if (n.Length > 0 && !filePos.ContainsKey(n))
                    filePos[n] = idx;
                idx++;
            }

            MatchCandidate best = null;
            foreach (var cfg in configs)
            {
                if (cfg.Fields == null || cfg.Fields.Count == 0) continue;

                // 模板特征 = 配置记录的所有非空列头（含未绑定字段的列），按配置列序排列
                var sig = new List<(string Header, string Field, string Norm)>();
                var hasBound = false;
                foreach (var f in cfg.Fields.OrderBy(x => x.ColumnIndex))
                {
                    if (string.IsNullOrWhiteSpace(f.Header)) continue;
                    var n = Normalize(f.Header);
                    if (n.Length == 0) continue;
                    sig.Add((f.Header, f.Field ?? "", n));
                    if (!string.IsNullOrWhiteSpace(f.Field)) hasBound = true;
                }

                // 没有任何绑定字段的配置不能解析数据，不参与匹配
                if (!hasBound || sig.Count == 0) continue;

                var matched = 0;
                var boundMatched = 0;
                var prevPos = -1;
                var orderOk = true;
                foreach (var s in sig)
                {
                    if (filePos.TryGetValue(s.Norm, out var pos))
                    {
                        matched++;
                        if (!string.IsNullOrWhiteSpace(s.Field)) boundMatched++;
                        if (pos <= prevPos) orderOk = false; // 未保持相对顺序
                        prevPos = pos;
                    }
                }

                var score = (double)matched / sig.Count;
                var isBetter = best == null;
                if (!isBetter)
                {
                    if (score - best.Score > 1e-9)
                        isBetter = true;                                    // ① 命中率
                    else if (Math.Abs(score - best.Score) < 1e-9)
                    {
                        if (matched > best.Matched)
                            isBetter = true;                                // ② 命中列数（更具体）
                        else if (matched == best.Matched && boundMatched > best.BoundMatched)
                            isBetter = true;                                // ③ 命中绑定字段列数
                        else if (matched == best.Matched
                                 && boundMatched == best.BoundMatched
                                 && orderOk && !best.OrderMatch)
                            isBetter = true;                                // ④ 列顺序一致
                        else if (matched == best.Matched
                                 && boundMatched == best.BoundMatched
                                 && orderOk == best.OrderMatch
                                 && cfg.Id < best.Config.Id)
                            isBetter = true;                                // ⑤ Id
                    }
                }

                if (isBetter)
                {
                    best = new MatchCandidate
                    {
                        Config = cfg,
                        Score = score,
                        Expected = sig.Count,
                        Matched = matched,
                        BoundMatched = boundMatched,
                        OrderMatch = orderOk
                    };
                }
            }
            return best;
        }
    }
}

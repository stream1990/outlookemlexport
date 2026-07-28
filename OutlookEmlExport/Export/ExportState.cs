using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace OutlookEmlExport.Export
{
    internal sealed class ExportState
    {
        private readonly string _path;
        private readonly HashSet<string> _done = new HashSet<string>(StringComparer.Ordinal);

        public Dictionary<string, int> Stats { get; private set; }

        public ExportState(string path)
        {
            _path = path;
            Stats = new Dictionary<string, int>();
            Stats["exported"] = 0;
            Stats["skipped"] = 0;
            Stats["failed"] = 0;
            Stats["folders"] = 0;

            if (!File.Exists(path))
                return;

            try
            {
                var json = File.ReadAllText(path);
                var ser = new JavaScriptSerializer();
                ser.MaxJsonLength = int.MaxValue;
                var data = ser.Deserialize<Dictionary<string, object>>(json);
                if (data == null)
                    return;

                object doneObj;
                if (data.TryGetValue("done", out doneObj) && doneObj is IEnumerable)
                {
                    foreach (var item in (IEnumerable)doneObj)
                    {
                        string s = item == null ? null : item.ToString();
                        if (!string.IsNullOrEmpty(s))
                            _done.Add(s);
                    }
                }

                object statsObj;
                if (data.TryGetValue("stats", out statsObj) && statsObj is IDictionary)
                {
                    foreach (DictionaryEntry kv in (IDictionary)statsObj)
                    {
                        try
                        {
                            string key = kv.Key == null ? null : kv.Key.ToString();
                            if (!string.IsNullOrEmpty(key))
                                Stats[key] = Convert.ToInt32(kv.Value);
                        }
                        catch
                        {
                            // ignore bad values
                        }
                    }
                }
            }
            catch
            {
                // 断点文件损坏时从头开始
            }
        }

        public bool IsDone(string entryId)
        {
            return _done.Contains(entryId);
        }

        public void MarkDone(string entryId)
        {
            if (!string.IsNullOrEmpty(entryId))
                _done.Add(entryId);
        }

        public void Inc(string key, int delta)
        {
            if (!Stats.ContainsKey(key))
                Stats[key] = 0;
            Stats[key] = Stats[key] + delta;
        }

        public void Inc(string key)
        {
            Inc(key, 1);
        }

        public void Save()
        {
            var payload = new Dictionary<string, object>();
            payload["done"] = _done.OrderBy(x => x, StringComparer.Ordinal).ToArray();
            payload["stats"] = Stats;
            payload["updated_at"] = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");

            var ser = new JavaScriptSerializer();
            ser.MaxJsonLength = int.MaxValue;
            File.WriteAllText(_path, PrettyJson(ser.Serialize(payload)));
        }

        private static string PrettyJson(string compact)
        {
            try
            {
                int indent = 0;
                var sb = new StringBuilder(compact.Length * 2);
                bool inString = false;
                for (int i = 0; i < compact.Length; i++)
                {
                    char c = compact[i];
                    if (c == '"' && (i == 0 || compact[i - 1] != '\\'))
                        inString = !inString;

                    if (inString)
                    {
                        sb.Append(c);
                        continue;
                    }

                    switch (c)
                    {
                        case '{':
                        case '[':
                            sb.Append(c);
                            sb.AppendLine();
                            indent++;
                            sb.Append(new string(' ', indent * 2));
                            break;
                        case '}':
                        case ']':
                            sb.AppendLine();
                            indent = Math.Max(0, indent - 1);
                            sb.Append(new string(' ', indent * 2));
                            sb.Append(c);
                            break;
                        case ',':
                            sb.Append(c);
                            sb.AppendLine();
                            sb.Append(new string(' ', indent * 2));
                            break;
                        case ':':
                            sb.Append(": ");
                            break;
                        default:
                            sb.Append(c);
                            break;
                    }
                }
                return sb.ToString();
            }
            catch
            {
                return compact;
            }
        }
    }
}

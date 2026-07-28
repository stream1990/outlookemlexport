using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookEmlExport.Export
{
    internal static class Utils
    {
        private static readonly Regex InvalidFsChars = new Regex(
            @"[<>:""/\\|?*\x00-\x1f]",
            RegexOptions.Compiled);

        private static readonly Regex WhitespaceRun = new Regex(
            @"\s+",
            RegexOptions.Compiled);

        public static string SafeName(string name, int maxLen)
        {
            name = (name ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(name))
                name = "untitled";

            name = InvalidFsChars.Replace(name, "_");
            name = WhitespaceRun.Replace(name, " ").Trim(' ', '.');
            if (string.IsNullOrEmpty(name))
                name = "untitled";

            if (name.Length > maxLen)
                name = name.Substring(0, maxLen).TrimEnd(' ', '.');

            return name;
        }

        public static string SafeName(string name)
        {
            return SafeName(name, 80);
        }

        public static string ShortId(string text, int n)
        {
            using (var sha1 = SHA1.Create())
            {
                var bytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(text ?? string.Empty));
                var hex = BitConverter.ToString(bytes).Replace("-", string.Empty).ToLowerInvariant();
                return hex.Substring(0, Math.Min(n, hex.Length));
            }
        }

        public static string ShortId(string text)
        {
            return ShortId(text, 10);
        }

        public static void EnsureDir(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        public static string ComStr(object value)
        {
            if (value == null)
                return string.Empty;
            try
            {
                string s = Convert.ToString(value);
                return s == null ? string.Empty : s.Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string UniqueEmlPath(string folderDir, string baseName)
        {
            var candidate = Path.Combine(folderDir, baseName + ".eml");
            if (!File.Exists(candidate))
                return candidate;

            int n = 2;
            while (true)
            {
                candidate = Path.Combine(folderDir, baseName + "_" + n + ".eml");
                if (!File.Exists(candidate))
                    return candidate;
                n++;
            }
        }

        /// <summary>
        /// 将显示名格式化为 RFC 地址，例如 "Name" &lt;addr@x.com&gt;
        /// </summary>
        public static string FormatAddr(string name, string address)
        {
            name = (name ?? string.Empty).Trim();
            address = (address ?? string.Empty).Trim();

            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(address))
                return string.Empty;
            if (string.IsNullOrEmpty(name))
                return address;
            if (string.IsNullOrEmpty(address))
                return name;

            if (name.IndexOfAny(new[] { '"', ',', '<', '>', '@', '\\' }) >= 0
                || name.IndexOf(' ') >= 0)
            {
                name = "\"" + name.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
            }

            return name + " <" + address + ">";
        }

        public static string MakeMessageId(string idString)
        {
            if (string.IsNullOrEmpty(idString))
                idString = ShortId(Guid.NewGuid().ToString("N"), 12);
            return "<" + idString + "@outlook.local>";
        }

        public static bool IsExchangeDn(string address)
        {
            return !string.IsNullOrEmpty(address)
                   && address.IndexOf('/') >= 0
                   && address.IndexOf('=') >= 0;
        }
    }
}

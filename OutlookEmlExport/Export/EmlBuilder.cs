using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OutlookEmlExport.Export
{
    /// <summary>
    /// 从 Outlook MailItem 组装标准 MIME/EML。
    /// 优先复用 PR_TRANSPORT_MESSAGE_HEADERS，再补正文与附件。
    /// 使用 late-binding COM，不依赖 Interop 程序集。
    /// </summary>
    internal static class EmlBuilder
    {
        private const string PrTransportMessageHeadersW =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001F";
        private const string PrTransportMessageHeadersA =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        private static readonly HashSet<string> SkipTransportKeys = CreateSkipKeys();

        private static HashSet<string> CreateSkipKeys()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string[] keys = new string[]
            {
                "subject", "from", "to", "cc", "bcc", "date",
                "content-type", "content-transfer-encoding", "mime-version",
            };
            foreach (var k in keys)
                set.Add(k);
            return set;
        }

        public static byte[] BuildEmlBytes(object mail)
        {
            var headers = new List<KeyValuePair<string, string>>();
            var headerKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Action<string, string, bool> addHeader = delegate(string key, string value, bool overwrite)
            {
                if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
                    return;
                if (headerKeys.Contains(key))
                {
                    if (!overwrite)
                        return;
                    for (int i = 0; i < headers.Count; i++)
                    {
                        if (string.Equals(headers[i].Key, key, StringComparison.OrdinalIgnoreCase))
                        {
                            headers[i] = new KeyValuePair<string, string>(key, value);
                            return;
                        }
                    }
                }
                headers.Add(new KeyValuePair<string, string>(key, value));
                headerKeys.Add(key);
            };

            string subject = Utils.ComStr(GetProp(mail, "Subject"));
            string senderName = Utils.ComStr(GetProp(mail, "SenderName"));
            string senderEmail = Utils.ComStr(GetProp(mail, "SenderEmailAddress"));
            if (Utils.IsExchangeDn(senderEmail))
                senderEmail = string.Empty;

            if (string.IsNullOrEmpty(senderEmail))
            {
                try
                {
                    object sender = GetProp(mail, "Sender");
                    if (sender != null)
                    {
                        try
                        {
                            dynamic dSender = sender;
                            object eu = dSender.GetExchangeUser();
                            if (eu != null)
                                senderEmail = Utils.ComStr(((dynamic)eu).PrimarySmtpAddress);
                        }
                        catch { /* ignore */ }

                        if (string.IsNullOrEmpty(senderEmail))
                        {
                            try
                            {
                                dynamic dSender = sender;
                                object ea = dSender.GetExchangeDistributionList();
                                if (ea != null)
                                    senderEmail = Utils.ComStr(((dynamic)ea).PrimarySmtpAddress);
                            }
                            catch { /* ignore */ }
                        }
                    }
                }
                catch { /* ignore */ }
            }

            object recipients = GetProp(mail, "Recipients");
            string toH = RecipientsToHeader(recipients, 1);
            if (string.IsNullOrEmpty(toH))
                toH = Utils.ComStr(GetProp(mail, "To"));
            string ccH = RecipientsToHeader(recipients, 2);
            if (string.IsNullOrEmpty(ccH))
                ccH = Utils.ComStr(GetProp(mail, "CC"));
            string bccH = RecipientsToHeader(recipients, 3);
            if (string.IsNullOrEmpty(bccH))
                bccH = Utils.ComStr(GetProp(mail, "BCC"));

            if (!string.IsNullOrEmpty(subject))
                addHeader("Subject", EncodeHeaderIfNeeded(subject), false);

            if (!string.IsNullOrEmpty(senderName) || !string.IsNullOrEmpty(senderEmail))
            {
                if (!string.IsNullOrEmpty(senderName) && !string.IsNullOrEmpty(senderEmail))
                    addHeader("From", EncodeHeaderIfNeeded(Utils.FormatAddr(senderName, senderEmail)), false);
                else if (!string.IsNullOrEmpty(senderEmail))
                    addHeader("From", senderEmail, false);
                else
                    addHeader("From", EncodeHeaderIfNeeded(senderName), false);
            }

            if (!string.IsNullOrEmpty(toH))
                addHeader("To", EncodeHeaderIfNeeded(toH), false);
            if (!string.IsNullOrEmpty(ccH))
                addHeader("Cc", EncodeHeaderIfNeeded(ccH), false);
            if (!string.IsNullOrEmpty(bccH))
                addHeader("Bcc", EncodeHeaderIfNeeded(bccH), false);

            string sent = FormatOutlookTime(GetProp(mail, "SentOn"));
            string received = FormatOutlookTime(GetProp(mail, "ReceivedTime"));
            if (!string.IsNullOrEmpty(sent))
                addHeader("Date", sent, false);
            else if (!string.IsNullOrEmpty(received))
                addHeader("Date", received, false);

            string headersRaw = GetMapiProperty(mail, PrTransportMessageHeadersW);
            if (headersRaw == null)
                headersRaw = GetMapiProperty(mail, PrTransportMessageHeadersA);
            if (!string.IsNullOrEmpty(headersRaw))
            {
                string[] lines = headersRaw
                    .Replace("\r\n", "\n")
                    .Replace('\r', '\n')
                    .Split(new[] { '\n' }, StringSplitOptions.None);
                foreach (string rawLine in lines)
                {
                    if (string.IsNullOrEmpty(rawLine) || rawLine[0] == ' ' || rawLine[0] == '\t')
                        continue;
                    int colon = rawLine.IndexOf(':');
                    if (colon <= 0)
                        continue;
                    string key = rawLine.Substring(0, colon).Trim();
                    string val = rawLine.Substring(colon + 1).Trim();
                    if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(val))
                        continue;
                    if (SkipTransportKeys.Contains(key))
                        continue;
                    if (headerKeys.Contains(key))
                        continue;
                    addHeader(key, val, false);
                }
            }

            if (!headerKeys.Contains("Message-ID") && !headerKeys.Contains("Message-Id"))
            {
                string entry = Utils.ComStr(GetProp(mail, "EntryID"));
                if (string.IsNullOrEmpty(entry))
                    entry = Utils.ShortId(subject + (sent ?? string.Empty), 12);
                addHeader("Message-ID", Utils.MakeMessageId(Utils.ShortId(entry, 12)), false);
            }

            addHeader("MIME-Version", "1.0", false);

            string html = Utils.ComStr(GetProp(mail, "HTMLBody"));
            string body = Utils.ComStr(GetProp(mail, "Body"));
            List<AttData> attachments = CollectAttachments(mail);

            string boundaryMixed = "----=_Mixed_" + Utils.ShortId(Guid.NewGuid().ToString("N"), 16);
            string boundaryAlt = "----=_Alt_" + Utils.ShortId(Guid.NewGuid().ToString("N"), 16);

            var sb = new StringBuilder();
            foreach (var h in headers)
            {
                if (string.Equals(h.Key, "Content-Type", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(h.Key, "Content-Transfer-Encoding", StringComparison.OrdinalIgnoreCase))
                    continue;
                sb.Append(h.Key).Append(": ").Append(h.Value ?? string.Empty).Append("\r\n");
            }

            bool hasHtml = !string.IsNullOrEmpty(html)
                           && html.Trim().ToLowerInvariant() != "<html><body></body></html>";
            bool hasAtt = attachments.Count > 0;

            if (hasAtt)
            {
                sb.Append("Content-Type: multipart/mixed; boundary=\"").Append(boundaryMixed).Append("\"\r\n");
                sb.Append("\r\n");
                sb.Append("This is a multi-part message in MIME format.\r\n\r\n");
                sb.Append("--").Append(boundaryMixed).Append("\r\n");
                AppendBodyPart(sb, body, html, hasHtml, boundaryAlt);
                foreach (var att in attachments)
                {
                    sb.Append("--").Append(boundaryMixed).Append("\r\n");
                    AppendAttachmentPart(sb, att.FileName, att.Data);
                }
                sb.Append("--").Append(boundaryMixed).Append("--\r\n");
            }
            else
            {
                AppendBodyPart(sb, body, html, hasHtml, boundaryAlt);
            }

            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        public static string MailFilename(object mail)
        {
            string entry = Utils.ComStr(GetProp(mail, "EntryID"));
            if (string.IsNullOrEmpty(entry))
            {
                var rnd = new byte[8];
                using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
                    rng.GetBytes(rnd);
                entry = BitConverter.ToString(rnd).Replace("-", string.Empty).ToLowerInvariant();
            }

            string sid = Utils.ShortId(entry, 10);
            object dt = GetProp(mail, "ReceivedTime");
            if (dt == null)
                dt = GetProp(mail, "SentOn");

            string datePart;
            if (dt is DateTime)
                datePart = ((DateTime)dt).ToString("yyyy-MM-dd_HHmmss");
            else
                datePart = "unknown-date";

            string subject = Utils.SafeName(Utils.ComStr(GetProp(mail, "Subject")), 60);
            if (string.IsNullOrEmpty(subject) || subject == "untitled")
                subject = "(no subject)";

            return datePart + "_" + subject + "_" + sid;
        }

        private static void AppendBodyPart(
            StringBuilder sb,
            string body,
            string html,
            bool hasHtml,
            string boundaryAlt)
        {
            if (hasHtml && !string.IsNullOrEmpty(body))
            {
                sb.Append("Content-Type: multipart/alternative; boundary=\"").Append(boundaryAlt).Append("\"\r\n\r\n");
                sb.Append("--").Append(boundaryAlt).Append("\r\n");
                AppendTextPart(sb, "text/plain; charset=utf-8", body ?? string.Empty);
                sb.Append("--").Append(boundaryAlt).Append("\r\n");
                AppendTextPart(sb, "text/html; charset=utf-8", html);
                sb.Append("--").Append(boundaryAlt).Append("--\r\n");
            }
            else if (hasHtml)
            {
                AppendTextPart(sb, "text/html; charset=utf-8", html);
            }
            else
            {
                AppendTextPart(sb, "text/plain; charset=utf-8", body ?? string.Empty);
            }
        }

        private static void AppendTextPart(StringBuilder sb, string contentType, string text)
        {
            sb.Append("Content-Type: ").Append(contentType).Append("\r\n");
            sb.Append("Content-Transfer-Encoding: base64\r\n\r\n");
            sb.Append(ToBase64Lines(Encoding.UTF8.GetBytes(text ?? string.Empty))).Append("\r\n");
        }

        private static void AppendAttachmentPart(StringBuilder sb, string fileName, byte[] data)
        {
            string safe = Utils.SafeName(fileName ?? "attachment");
            sb.Append("Content-Type: application/octet-stream; name=\"").Append(EscapeQuoted(safe)).Append("\"\r\n");
            sb.Append("Content-Transfer-Encoding: base64\r\n");
            sb.Append("Content-Disposition: attachment; filename=\"").Append(EscapeQuoted(safe)).Append("\"\r\n");
            sb.Append("\r\n");
            sb.Append(ToBase64Lines(data ?? new byte[0])).Append("\r\n");
        }

        private static string EscapeQuoted(string s)
        {
            return (s ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        private static string ToBase64Lines(byte[] data)
        {
            if (data == null)
                data = new byte[0];
            string b64 = Convert.ToBase64String(data);
            if (b64.Length <= 76)
                return b64;
            var sb = new StringBuilder(b64.Length + b64.Length / 76 * 2);
            for (int i = 0; i < b64.Length; i += 76)
            {
                if (i > 0)
                    sb.Append("\r\n");
                sb.Append(b64, i, Math.Min(76, b64.Length - i));
            }
            return sb.ToString();
        }

        private static string EncodeHeaderIfNeeded(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            bool needEncode = false;
            foreach (char c in value)
            {
                if (c > 127)
                {
                    needEncode = true;
                    break;
                }
            }
            if (!needEncode)
                return value;

            byte[] bytes = Encoding.UTF8.GetBytes(value);
            return "=?utf-8?B?" + Convert.ToBase64String(bytes) + "?=";
        }

        private static string RecipientsToHeader(object recipients, int wantType)
        {
            if (recipients == null)
                return string.Empty;

            var parts = new List<string>();
            try
            {
                dynamic recips = recipients;
                int count = Convert.ToInt32(recips.Count);
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic r = recips.Item(i);
                        if (Convert.ToInt32(r.Type) != wantType)
                            continue;

                        string name = Utils.ComStr(r.Name);
                        string addr = Utils.ComStr(r.Address);
                        if (Utils.IsExchangeDn(addr))
                            addr = string.Empty;

                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(addr)
                            && !string.Equals(name, addr, StringComparison.OrdinalIgnoreCase))
                            parts.Add(Utils.FormatAddr(name, addr));
                        else if (!string.IsNullOrEmpty(addr))
                            parts.Add(addr);
                        else if (!string.IsNullOrEmpty(name))
                            parts.Add(name);
                    }
                    catch
                    {
                        // skip bad recipient
                    }
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Join(", ", parts.ToArray());
        }

        private static string FormatOutlookTime(object dt)
        {
            if (dt == null)
                return null;
            try
            {
                if (dt is DateTime)
                {
                    DateTime d = (DateTime)dt;
                    if (d.Kind == DateTimeKind.Local || d.Kind == DateTimeKind.Unspecified)
                    {
                        TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(d);
                        string sign = offset >= TimeSpan.Zero ? "+" : "-";
                        TimeSpan abs = offset.Duration();
                        return d.ToString("ddd, dd MMM yyyy HH:mm:ss ", CultureInfo.InvariantCulture)
                               + sign
                               + abs.Hours.ToString("00")
                               + abs.Minutes.ToString("00");
                    }
                    return d.ToString("ddd, dd MMM yyyy HH:mm:ss +0000", CultureInfo.InvariantCulture);
                }
                return Convert.ToString(dt, CultureInfo.InvariantCulture);
            }
            catch
            {
                try { return Convert.ToString(dt); }
                catch { return null; }
            }
        }

        private static string GetMapiProperty(object mail, string schema)
        {
            try
            {
                dynamic dMail = mail;
                object pa = dMail.PropertyAccessor;
                object val = ((dynamic)pa).GetProperty(schema);
                if (val == null)
                    return null;
                return Convert.ToString(val);
            }
            catch
            {
                return null;
            }
        }

        private static object GetProp(object comObj, string name)
        {
            if (comObj == null)
                return null;
            try
            {
                dynamic d = comObj;
                switch (name)
                {
                    case "Subject": return d.Subject;
                    case "SenderName": return d.SenderName;
                    case "SenderEmailAddress": return d.SenderEmailAddress;
                    case "Sender": return d.Sender;
                    case "Recipients": return d.Recipients;
                    case "To": return d.To;
                    case "CC": return d.CC;
                    case "BCC": return d.BCC;
                    case "SentOn": return d.SentOn;
                    case "ReceivedTime": return d.ReceivedTime;
                    case "EntryID": return d.EntryID;
                    case "HTMLBody": return d.HTMLBody;
                    case "Body": return d.Body;
                    case "Attachments": return d.Attachments;
                    case "Class": return d.Class;
                    case "MessageClass": return d.MessageClass;
                    case "Name": return d.Name;
                    default: return null;
                }
            }
            catch
            {
                return null;
            }
        }

        private sealed class AttData
        {
            public string FileName;
            public byte[] Data;
        }

        private static List<AttData> CollectAttachments(object mail)
        {
            var list = new List<AttData>();
            try
            {
                dynamic dMail = mail;
                dynamic atts = dMail.Attachments;
                int count = Convert.ToInt32(atts.Count);
                if (count <= 0)
                    return list;

                string tmpRoot = Path.Combine(Path.GetTempPath(), "ol_eml_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tmpRoot);
                try
                {
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            dynamic att = atts.Item(i);
                            string fname = Utils.SafeName(Utils.ComStr(att.FileName));
                            if (string.IsNullOrEmpty(fname) || fname == "untitled")
                                fname = "attachment_" + i;
                            string saveAs = Path.Combine(tmpRoot, i + "_" + fname);
                            att.SaveAsFile(saveAs);
                            list.Add(new AttData
                            {
                                FileName = fname,
                                Data = File.ReadAllBytes(saveAs),
                            });
                        }
                        catch
                        {
                            // 嵌入图片/受保护附件可能失败，跳过
                        }
                    }
                }
                finally
                {
                    try { Directory.Delete(tmpRoot, true); }
                    catch { /* ignore */ }
                }
            }
            catch
            {
                // no attachments
            }
            return list;
        }
    }
}

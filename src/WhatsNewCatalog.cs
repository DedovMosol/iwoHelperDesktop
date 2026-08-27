using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace ExcelMerger
{
    internal static class WhatsNewCatalog
    {
        private const string ResourceName = "whatsnew.json";

        internal static string CurrentVersion
        {
            get { return Assembly.GetExecutingAssembly().GetName().Version.ToString(3); }
        }

        internal static bool IsPacked()
        {
            using (Stream stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(ResourceName))
                return stream != null && stream.Length > 32;
        }

        internal static bool ShouldShow(UserSettings settings, string version)
        {
            return settings != null && settings.ShowWhatsNewOnStart &&
                !string.Equals(settings.LastWhatsNewVersion, version,
                    StringComparison.OrdinalIgnoreCase);
        }

        internal static List<string> Items(string version, string language)
        {
            string json = "";
            try
            {
                using (Stream stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream(ResourceName))
                using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                    json = reader.ReadToEnd();
            }
            catch { }
            string notes = UpdateChecker.ExtractNotes(json, version, language);
            var result = new List<string>();
            foreach (string line in (notes ?? "").Replace("\r\n", "\n").Split('\n'))
            {
                string item = line.Trim();
                if (item.StartsWith("•", StringComparison.Ordinal))
                    item = item.Substring(1).Trim();
                if (item.Length > 0)
                    result.Add(item);
            }
            if (result.Count == 0)
                result.Add(Loc.T("whatsnew.fallback"));
            return result;
        }
    }
}

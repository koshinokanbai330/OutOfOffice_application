using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OutOfOfficeAddin.Services
{
    /// <summary>
    /// Persists and restores the To / Cc recipient lists to/from
    /// %USERPROFILE%\Documents\mailingList.txt.
    ///
    /// File format (UTF-8, OS line endings):
    ///   To: addr1; addr2
    ///   Cc: addr3
    /// </summary>
    public static class MailingListService
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "mailingList.txt");

        /// <summary>
        /// Loads saved To and Cc recipients.
        /// Returns empty lists if the file does not exist or cannot be parsed.
        /// </summary>
        public static (List<string> to, List<string> cc) Load()
        {
            var to = new List<string>();
            var cc = new List<string>();

            if (!File.Exists(FilePath))
                return (to, cc);

            try
            {
                foreach (var line in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    if (line.StartsWith("To:", StringComparison.OrdinalIgnoreCase))
                    {
                        to.AddRange(ParseAddresses(line.Substring(3)));
                    }
                    else if (line.StartsWith("Cc:", StringComparison.OrdinalIgnoreCase))
                    {
                        cc.AddRange(ParseAddresses(line.Substring(3)));
                    }
                }
            }
            catch
            {
                // Return whatever we have so far; a corrupt file won't break the add-in.
            }

            return (to, cc);
        }

        /// <summary>
        /// Saves the given To and Cc lists, overwriting any previous file.
        /// </summary>
        public static void Save(IEnumerable<string> to, IEnumerable<string> cc)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));

            var sb = new StringBuilder();
            sb.AppendLine("To: " + string.Join("; ", to));
            sb.AppendLine("Cc: " + string.Join("; ", cc));

            File.WriteAllText(FilePath, sb.ToString(), Encoding.UTF8);
        }

        private static IEnumerable<string> ParseAddresses(string raw)
        {
            foreach (var part in raw.Split(';'))
            {
                var trimmed = part.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                    yield return trimmed;
            }
        }
    }
}

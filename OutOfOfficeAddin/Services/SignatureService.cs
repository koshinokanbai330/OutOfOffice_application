using System;
using System.IO;
using System.Text;

namespace OutOfOfficeAddin.Services
{
    /// <summary>
    /// Reads the user's default Outlook HTML signature from the well-known
    /// signature folder so it can be appended to OOF messages.
    ///
    /// Outlook stores signatures in:
    ///   %APPDATA%\Microsoft\Signatures\
    /// Each signature appears as three files:
    ///   {name}.htm  – HTML version (used here)
    ///   {name}.rtf
    ///   {name}.txt
    ///
    /// The active signature for new messages is recorded in the registry key
    ///   HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles\{profile}\9375CFF0413111d3B88A00104B2A6676\{account}
    /// value "New Signature".  Reading the registry is complex and fragile,
    /// so this service simply returns the first .htm file it finds.
    /// If no file is found an empty string is returned.
    /// </summary>
    public static class SignatureService
    {
        private static readonly string SignatureFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Microsoft", "Signatures");

        /// <summary>
        /// Returns the HTML content of the user's first available Outlook signature,
        /// or an empty string if none is found.
        /// When multiple signatures exist, the first one alphabetically is used.
        /// Check <see cref="AvailableSignatureNames"/> to see which signatures are present.
        /// </summary>
        public static string GetDefaultSignatureHtml()
        {
            if (!Directory.Exists(SignatureFolder))
                return string.Empty;

            var files = Directory.GetFiles(SignatureFolder, "*.htm");
            if (files.Length == 0)
                return string.Empty;

            // Sort for deterministic selection when multiple signatures exist
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);

            try
            {
                return File.ReadAllText(files[0], Encoding.UTF8);
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the names (without extension) of all available Outlook signatures.
        /// Useful for informing the user which signature was selected.
        /// </summary>
        public static string[] AvailableSignatureNames()
        {
            if (!Directory.Exists(SignatureFolder))
                return Array.Empty<string>();

            var files = Directory.GetFiles(SignatureFolder, "*.htm");
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);
            var names = new string[files.Length];
            for (int i = 0; i < files.Length; i++)
                names[i] = Path.GetFileNameWithoutExtension(files[i]);
            return names;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace OutOfOfficeAddin.Services
{
    /// <summary>
    /// Downloads the travel-allowance Excel template and fills it with
    /// the business trip data.
    ///
    /// Template sheets:
    ///   "日帰り One-Day"  – used for single-day and multi-day trips
    ///   "宿泊 Overnight"  – additionally used for multi-day (2+ nights) trips
    ///
    /// Column layout is discovered dynamically by searching for a header row
    /// and then writing to the first blank row below it.
    /// </summary>
    public class ExcelService
    {
        private const string TemplateUrl =
            "https://inside-docupedia.bosch.com/confluence/download/attachments/1560215212/" +
            "BEGJ_%E5%87%BA%E5%BC%B5%E6%97%A5%E5%BD%93%E7%B2%BE%E7%AE%97%E6%9B%B8.xlsx" +
            "?version=1&modificationDate=1610517538000&api=v2";

        // Header text patterns (partial match, case-insensitive)
        private const string ColHeaderDate = "日にち";
        private const string ColHeaderDestination = "出張先";
        private const string ColHeaderDepart = "出発";
        private const string ColHeaderStart = "始業";
        private const string ColHeaderFinish = "終業";
        private const string ColHeaderReturn = "帰着";

        private static readonly TimeSpan TimeDepart = new TimeSpan(7, 0, 0);
        private static readonly TimeSpan TimeStart = new TimeSpan(9, 0, 0);
        private static readonly TimeSpan TimeFinish = new TimeSpan(18, 0, 0);
        private static readonly TimeSpan TimeReturn = new TimeSpan(21, 0, 0);

        /// <summary>
        /// Downloads the template, fills the data and saves the workbook.
        /// </summary>
        /// <param name="startDate">First trip date.</param>
        /// <param name="endDate">Last trip date.</param>
        /// <param name="destination">Location / destination city.</param>
        /// <param name="familyName">User's family name (used in filename).</param>
        /// <param name="saveFolder">Local folder where the file will be saved.</param>
        /// <returns>Full path of the saved file.</returns>
        public async Task<string> CreateAsync(
            DateTime startDate,
            DateTime endDate,
            string destination,
            string familyName,
            string saveFolder)
        {
            // Download template to a temp file
            var tempPath = Path.Combine(Path.GetTempPath(), $"OOF_template_{Guid.NewGuid()}.xlsx");
            await DownloadTemplateAsync(tempPath);

            var fileName = $"BT-Allowance-{familyName}-{startDate:yyyyMMdd}.xlsx";
            var savePath = Path.Combine(saveFolder, fileName);

            var dates = BuildDateList(startDate, endDate);
            var multiDay = dates.Count > 1;

            using (var wb = new XLWorkbook(tempPath))
            {
                FillSheet(wb, "日帰り One-Day", dates, destination);

                if (multiDay)
                    FillSheet(wb, "宿泊 Overnight", dates, destination);

                Directory.CreateDirectory(saveFolder);
                wb.SaveAs(savePath);
            }

            // Clean up temp file
            try { File.Delete(tempPath); } catch { }

            return savePath;
        }

        // -------------------------------------------------------------------

        private static async Task DownloadTemplateAsync(string tempPath)
        {
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(30);
                var bytes = await client.GetByteArrayAsync(TemplateUrl);
                File.WriteAllBytes(tempPath, bytes);
            }
        }

        private static List<DateTime> BuildDateList(
            DateTime start, DateTime end)
        {
            var list = new List<DateTime>();
            for (var d = start.Date; d <= end.Date; d = d.AddDays(1))
                list.Add(d);
            return list;
        }

        private static void FillSheet(
            XLWorkbook wb,
            string sheetName,
            List<DateTime> dates,
            string destination)
        {
            if (!wb.TryGetWorksheet(sheetName, out var ws))
                return; // Sheet not found – skip silently

            // Find header row and column indices
            int headerRow = FindHeaderRow(ws);
            if (headerRow < 1) return;

            int colDate = FindColumn(ws, headerRow, ColHeaderDate);
            int colDest = FindColumn(ws, headerRow, ColHeaderDestination);
            int colDepart = FindColumn(ws, headerRow, ColHeaderDepart);
            int colStart = FindColumn(ws, headerRow, ColHeaderStart);
            int colFinish = FindColumn(ws, headerRow, ColHeaderFinish);
            int colReturn = FindColumn(ws, headerRow, ColHeaderReturn);

            // Find the first blank row below the header
            int dataRow = FindFirstBlankRow(ws, headerRow + 1, colDate > 0 ? colDate : 1);

            foreach (var date in dates)
            {
                if (colDate > 0)
                    ws.Cell(dataRow, colDate).Value = date.Date;

                if (colDest > 0)
                    ws.Cell(dataRow, colDest).Value = destination;

                if (colDepart > 0)
                    ws.Cell(dataRow, colDepart).Value = date.Date + TimeDepart;

                if (colStart > 0)
                    ws.Cell(dataRow, colStart).Value = date.Date + TimeStart;

                if (colFinish > 0)
                    ws.Cell(dataRow, colFinish).Value = date.Date + TimeFinish;

                if (colReturn > 0)
                    ws.Cell(dataRow, colReturn).Value = date.Date + TimeReturn;

                dataRow++;
            }
        }

        /// <summary>Searches rows 1–50 for the first row that contains the date header.</summary>
        private static int FindHeaderRow(IXLWorksheet ws)
        {
            for (int r = 1; r <= 50; r++)
            {
                foreach (var cell in ws.Row(r).CellsUsed())
                {
                    if (ContainsText(cell.GetString(), ColHeaderDate))
                        return r;
                }
            }
            return -1;
        }

        /// <summary>Returns the column index of the header containing <paramref name="headerText"/>.</summary>
        private static int FindColumn(IXLWorksheet ws, int headerRow, string headerText)
        {
            var row = ws.Row(headerRow);
            foreach (var cell in row.CellsUsed())
            {
                if (ContainsText(cell.GetString(), headerText))
                    return cell.Address.ColumnNumber;
            }
            return -1;
        }

        /// <summary>Returns the first row at or below <paramref name="startRow"/> where column <paramref name="col"/> is empty.</summary>
        private static int FindFirstBlankRow(IXLWorksheet ws, int startRow, int col)
        {
            const int MaxRows = 10000;
            int r = startRow;
            int limit = startRow + MaxRows;
            while (r < limit && !ws.Cell(r, col).IsEmpty())
                r++;
            if (r >= limit)
                throw new InvalidOperationException(
                    $"Could not find a blank row in column {col} within {MaxRows} rows of row {startRow}.");
            return r;
        }

        private static bool ContainsText(string cellValue, string pattern)
            => cellValue != null &&
               cellValue.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0;
    }
}

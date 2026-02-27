using System;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OutOfOfficeAddin.Services
{
    /// <summary>
    /// Sets the user's Outlook automatic-reply (OOF) settings via Microsoft Graph.
    /// </summary>
    public class OofService
    {
        private readonly GraphAuthService _auth;

        public OofService(GraphAuthService auth)
        {
            _auth = auth;
        }

        /// <summary>
        /// Enables automatic replies for the given date range.
        /// </summary>
        /// <param name="startDate">First day of absence.</param>
        /// <param name="endDate">Last day of absence (back-date text will show endDate + 1).</param>
        /// <param name="signatureHtml">User's HTML signature to append to both messages.</param>
        public async Task SetAsync(DateTime startDate, DateTime endDate, string signatureHtml)
        {
            var token = await _auth.AcquireTokenAsync();
            var backDate = endDate.AddDays(1);

            var internalHtml = BuildInternalHtml(backDate, signatureHtml);
            var externalHtml = BuildExternalHtml(backDate, signatureHtml);

            var payload = new
            {
                automaticRepliesSetting = new
                {
                    status = "scheduled",
                    scheduledStartDateTime = new
                    {
                        dateTime = startDate.Date.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
                        timeZone = "Tokyo Standard Time"
                    },
                    scheduledEndDateTime = new
                    {
                        // End at 23:59 on the last day so the OOF covers the full day
                        dateTime = endDate.Date.AddDays(1).AddSeconds(-1)
                            .ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
                        timeZone = "Tokyo Standard Time"
                    },
                    internalReplyMessage = internalHtml,
                    externalReplyMessage = externalHtml
                }
            };

            var json = JsonConvert.SerializeObject(payload);
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", token);

                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // PatchAsync is .NET 5+; use SendAsync with HttpMethod.Patch for .NET Framework 4.8
                var request = new HttpRequestMessage(new HttpMethod("PATCH"),
                    "https://graph.microsoft.com/v1.0/me/mailboxSettings")
                {
                    Content = content
                };

                var response = await client.SendAsync(request);

                if (!response.IsSuccessStatusCode)
                {
                    var body = await response.Content.ReadAsStringAsync();
                    throw new InvalidOperationException(
                        $"Graph API returned {(int)response.StatusCode}: {body}");
                }
            }
        }

        // ------------------------------------------------------------------ helpers

        private static string BuildInternalHtml(DateTime backDate, string signatureHtml)
        {
            // English date: MMM dd, yyyy (invariant English month abbreviation)
            var engDate = backDate.ToString("MMM dd, yyyy", CultureInfo.GetCultureInfo("en-US"));

            var body =
                $"<p>Dear Sender,</p>" +
                $"<p>Thank you for your email.<br/>" +
                $"I will be back {engDate}. Email will be read with delay.</p>";

            return WrapHtml(body, signatureHtml);
        }

        private static string BuildExternalHtml(DateTime backDate, string signatureHtml)
        {
            var engDate = backDate.ToString("MMM dd, yyyy", CultureInfo.GetCultureInfo("en-US"));
            var jpDate = backDate.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);

            var body =
                $"<p>Dear Sender,</p>" +
                $"<p>Thank you for your email.<br/>" +
                $"I will be back {engDate}. Email will be read with delay.</p>" +
                $"<p>ご連絡ありがとうございます。<br/>" +
                $"申し訳ありませんが、{jpDate} まで不在のため対応できません。<br/>" +
                $"ご理解いただけますと幸いです。</p>";

            return WrapHtml(body, signatureHtml);
        }

        private static string WrapHtml(string body, string signatureHtml)
        {
            var sig = string.IsNullOrWhiteSpace(signatureHtml)
                ? string.Empty
                : $"<hr/>{signatureHtml}";

            return $"<html><body>{body}{sig}</body></html>";
        }
    }
}

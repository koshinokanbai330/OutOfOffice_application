using System;
using System.Collections.Generic;

namespace OutOfOfficeAddin.Models
{
    /// <summary>
    /// Holds all user-entered data for an out-of-office request.
    /// </summary>
    public class OutOfOfficeRequest
    {
        public LeaveType LeaveType { get; set; } = LeaveType.FullDayOff;
        public DateTime StartDate { get; set; } = DateTime.Today;
        public DateTime EndDate { get; set; } = DateTime.Today;

        /// <summary>Subject shown in the meeting (auto-generated from family name + leave type).</summary>
        public string Subject { get; set; } = string.Empty;

        public string Location { get; set; } = string.Empty;

        /// <summary>Resolved display names / SMTP addresses for the To field.</summary>
        public List<string> ToRecipients { get; set; } = new List<string>();

        /// <summary>Resolved display names / SMTP addresses for the Cc field.</summary>
        public List<string> CcRecipients { get; set; } = new List<string>();

        public bool SetAutoReplies { get; set; } = true;

        // --- Business Trip only ---
        public bool CreateExcel { get; set; } = true;
        public string ExcelSaveFolder { get; set; } = string.Empty;
    }
}

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OutOfOfficeAddin.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutOfOfficeAddin.Services
{
    /// <summary>
    /// Creates and optionally sends an all-day meeting request via Outlook interop.
    /// </summary>
    public class MeetingService
    {
        private readonly Outlook.Application _outlookApp;

        public MeetingService(Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp;
        }

        /// <summary>
        /// Creates a draft meeting item in the default calendar.
        /// If <paramref name="send"/> is true the item is sent; otherwise it is saved as a draft.
        /// </summary>
        /// <returns>
        /// The <see cref="Outlook.AppointmentItem"/> that was created.
        /// The caller is responsible for releasing the COM object.
        /// </returns>
        public Outlook.AppointmentItem CreateOrSend(OutOfOfficeRequest request, bool send)
        {
            Outlook.AppointmentItem appt = null;

            try
            {
                appt = (Outlook.AppointmentItem)_outlookApp.CreateItem(
                    Outlook.OlItemType.olAppointmentItem);

                appt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                appt.AllDayEvent = true;
                appt.Start = request.StartDate.Date;
                appt.End = request.EndDate.Date.AddDays(1); // Outlook all-day: End = Start + n days
                appt.Subject = request.Subject;
                appt.Location = request.Location;

                // Show as Free
                appt.BusyStatus = Outlook.OlBusyStatus.olFree;

                // Reminder off
                appt.ReminderSet = false;

                // Recipients
                foreach (var addr in request.ToRecipients)
                {
                    var recip = appt.Recipients.Add(addr);
                    recip.Type = (int)Outlook.OlMeetingRecipientType.olRequired;
                    recip.Resolve();
                }

                foreach (var addr in request.CcRecipients)
                {
                    var recip = appt.Recipients.Add(addr);
                    recip.Type = (int)Outlook.OlMeetingRecipientType.olOptional;
                    recip.Resolve();
                }

                if (send)
                {
                    appt.Send();
                }
                else
                {
                    appt.Save();
                }

                return appt;
            }
            catch
            {
                // Release COM object on failure to avoid leaks
                if (appt != null)
                {
                    try { Marshal.ReleaseComObject(appt); } catch { }
                }
                throw;
            }
        }

        /// <summary>
        /// Opens the Outlook "Select Names" dialog and returns the chosen addresses
        /// as a semicolon-separated string (display-name format that Outlook can resolve).
        /// Returns null if the user cancels.
        /// </summary>
        public string ShowAddressBook(string title = "Select Recipients")
        {
            Outlook.SelectNamesDialog dlg = null;
            try
            {
                dlg = _outlookApp.Session.GetSelectNamesDialog();
                dlg.Caption = title;
                dlg.NumberOfRecipientSelectors = Outlook.OlRecipientSelectors.olShowTo;
                dlg.ForceResolution = false;
                dlg.AllowMultipleSelection = true;
                dlg.Display();

                var result = dlg.Recipients;
                if (result.Count == 0)
                    return null;

                var addresses = new List<string>();
                foreach (Outlook.Recipient r in result)
                    addresses.Add(r.Address ?? r.Name);

                return string.Join("; ", addresses);
            }
            finally
            {
                if (dlg != null)
                {
                    try { Marshal.ReleaseComObject(dlg); } catch { }
                }
            }
        }
    }
}

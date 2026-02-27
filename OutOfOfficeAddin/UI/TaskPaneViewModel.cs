using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutOfOfficeAddin.Models;
using OutOfOfficeAddin.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutOfOfficeAddin.UI
{
    /// <summary>
    /// ViewModel for the Out-of-Office task pane.
    /// Exposes all bindings and orchestrates service calls.
    /// </summary>
    public class TaskPaneViewModel : INotifyPropertyChanged
    {
        // ---- injected services ----
        private readonly MeetingService _meetingService;
        private readonly OofService _oofService;
        private readonly ExcelService _excelService;
        private readonly GraphAuthService _authService;

        // ---- leave type items ----
        public IReadOnlyList<LeaveTypeItem> LeaveTypeItems { get; } = new[]
        {
            new LeaveTypeItem(LeaveType.BusinessTrip, "Business Trip"),
            new LeaveTypeItem(LeaveType.FullDayOff, "Full Day Off"),
            new LeaveTypeItem(LeaveType.AmHalfDayOff, "AM Half Day Off"),
            new LeaveTypeItem(LeaveType.PmHalfDayOff, "PM Half Day Off"),
        };

        public TaskPaneViewModel(Outlook.Application outlookApp)
        {
            _authService = new GraphAuthService();
            _meetingService = new MeetingService(outlookApp);
            _oofService = new OofService(_authService);
            _excelService = new ExcelService();

            // Determine family name from the current user's account
            try
            {
                var currentUser = outlookApp.Session.CurrentUser;
                // Try Exchange-specific LastName first; fall back to first word of display name
                string lastName = null;
                try
                {
                    var exchUser = currentUser.AddressEntry.GetExchangeUser();
                    if (exchUser != null && !string.IsNullOrWhiteSpace(exchUser.LastName))
                        lastName = exchUser.LastName;
                }
                catch { /* Not an Exchange account or unsupported */ }

                if (string.IsNullOrWhiteSpace(lastName))
                {
                    // Fall back to the first token of the display name
                    var parts = currentUser.Name?.Split(new[] { ' ' }, 2,
                        StringSplitOptions.RemoveEmptyEntries);
                    lastName = (parts != null && parts.Length > 0) ? parts[0] : "User";
                }
                _familyName = lastName;
            }
            catch { _familyName = "User"; }

            SelectedLeaveType = LeaveTypeItems[1]; // default: Full Day Off
            StartDate = DateTime.Today;
            EndDate = DateTime.Today;
            SetAutoReplies = true;
            CreateExcel = true;

            // Load saved mailing list
            var (to, cc) = MailingListService.Load();
            ToText = string.Join("; ", to);
            CcText = string.Join("; ", cc);

            CreateDraftCommand = new RelayCommand(async () => await ExecuteAsync(send: false));
            SendCommand = new RelayCommand(async () => await ExecuteAsync(send: true));
            CancelCommand = new RelayCommand(OnCancel);
            AddToFromAbCommand = new RelayCommand(() => AddFromAddressBook(isTo: true));
            AddCcFromAbCommand = new RelayCommand(() => AddFromAddressBook(isTo: false));
            BrowseExcelFolderCommand = new RelayCommand(OnBrowseExcelFolder);
        }

        // ------------------------------------------------------------------ Properties

        private string _familyName;

        private LeaveTypeItem _selectedLeaveType;
        public LeaveTypeItem SelectedLeaveType
        {
            get => _selectedLeaveType;
            set
            {
                if (SetField(ref _selectedLeaveType, value))
                {
                    UpdateSubject();
                    UpdateLocation();
                    OnPropertyChanged(nameof(IsBusinessTrip));
                }
            }
        }

        private DateTime _startDate;
        public DateTime StartDate
        {
            get => _startDate;
            set { if (SetField(ref _startDate, value)) UpdatePreviewMessages(); }
        }

        private DateTime _endDate;
        public DateTime EndDate
        {
            get => _endDate;
            set { if (SetField(ref _endDate, value)) UpdatePreviewMessages(); }
        }

        private string _subject = string.Empty;
        public string Subject
        {
            get => _subject;
            private set => SetField(ref _subject, value);
        }

        private string _location = string.Empty;
        public string Location
        {
            get => _location;
            set => SetField(ref _location, value);
        }

        private string _toText = string.Empty;
        public string ToText
        {
            get => _toText;
            set => SetField(ref _toText, value);
        }

        private string _ccText = string.Empty;
        public string CcText
        {
            get => _ccText;
            set => SetField(ref _ccText, value);
        }

        private bool _setAutoReplies;
        public bool SetAutoReplies
        {
            get => _setAutoReplies;
            set
            {
                if (SetField(ref _setAutoReplies, value))
                    UpdatePreviewMessages();
            }
        }

        private string _internalMessagePreview = string.Empty;
        public string InternalMessagePreview
        {
            get => _internalMessagePreview;
            private set => SetField(ref _internalMessagePreview, value);
        }

        private string _externalMessagePreview = string.Empty;
        public string ExternalMessagePreview
        {
            get => _externalMessagePreview;
            private set => SetField(ref _externalMessagePreview, value);
        }

        public bool IsBusinessTrip =>
            _selectedLeaveType?.LeaveType == LeaveType.BusinessTrip;

        private bool _createExcel;
        public bool CreateExcel
        {
            get => _createExcel;
            set => SetField(ref _createExcel, value);
        }

        private string _excelSaveFolder = string.Empty;
        public string ExcelSaveFolder
        {
            get => _excelSaveFolder;
            set => SetField(ref _excelSaveFolder, value);
        }

        private string _statusLog = string.Empty;
        public string StatusLog
        {
            get => _statusLog;
            private set => SetField(ref _statusLog, value);
        }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            private set => SetField(ref _isBusy, value);
        }

        // ------------------------------------------------------------------ Commands

        public RelayCommand CreateDraftCommand { get; }
        public RelayCommand SendCommand { get; }
        public RelayCommand CancelCommand { get; }
        public RelayCommand AddToFromAbCommand { get; }
        public RelayCommand AddCcFromAbCommand { get; }
        public RelayCommand BrowseExcelFolderCommand { get; }

        // ------------------------------------------------------------------ private helpers

        private void UpdateSubject()
        {
            if (_selectedLeaveType != null)
                Subject = SubjectHelper.Build(_familyName, _selectedLeaveType.LeaveType);
        }

        private void UpdateLocation()
        {
            if (_selectedLeaveType != null)
                Location = SubjectHelper.DefaultLocation(_selectedLeaveType.LeaveType);
        }

        private void UpdatePreviewMessages()
        {
            if (!SetAutoReplies)
            {
                InternalMessagePreview = "(Auto-reply disabled)";
                ExternalMessagePreview = "(Auto-reply disabled)";
                return;
            }

            var backDate = EndDate.AddDays(1);
            var engDate = backDate.ToString("MMM dd, yyyy",
                System.Globalization.CultureInfo.GetCultureInfo("en-US"));
            var jpDate = backDate.ToString("yyyy/MM/dd",
                System.Globalization.CultureInfo.InvariantCulture);

            InternalMessagePreview =
                $"Dear Sender,\n\n" +
                $"Thank you for your email.\n" +
                $"I will be back {engDate}. Email will be read with delay.\n\n" +
                $"[Your signature]";

            ExternalMessagePreview =
                $"Dear Sender,\n\n" +
                $"Thank you for your email.\n" +
                $"I will be back {engDate}. Email will be read with delay.\n\n" +
                $"ご連絡ありがとうございます。\n" +
                $"申し訳ありませんが、{jpDate} まで不在のため対応できません。\n" +
                $"ご理解いただけますと幸いです。\n\n" +
                $"[Your signature]";
        }

        private void AddFromAddressBook(bool isTo)
        {
            var result = _meetingService.ShowAddressBook(
                isTo ? "Select To Recipients" : "Select Cc Recipients");

            if (result == null) return;

            if (isTo)
                ToText = string.IsNullOrWhiteSpace(ToText) ? result : ToText + "; " + result;
            else
                CcText = string.IsNullOrWhiteSpace(CcText) ? result : CcText + "; " + result;
        }

        private void OnBrowseExcelFolder()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select folder to save Excel allowance sheet";
                if (!string.IsNullOrWhiteSpace(ExcelSaveFolder))
                    dlg.SelectedPath = ExcelSaveFolder;

                if (dlg.ShowDialog() == DialogResult.OK)
                    ExcelSaveFolder = dlg.SelectedPath;
            }
        }

        private void OnCancel()
        {
            // Reset form to defaults
            SelectedLeaveType = LeaveTypeItems[1];
            StartDate = DateTime.Today;
            EndDate = DateTime.Today;
            ToText = string.Empty;
            CcText = string.Empty;
            StatusLog = "Cancelled.";
        }

        private static List<string> ParseAddresses(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<string>();

            var list = new List<string>();
            foreach (var part in text.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = part.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                    list.Add(trimmed);
            }
            return list;
        }

        private async Task ExecuteAsync(bool send)
        {
            if (IsBusy) return;

            // Validate
            if (send && string.IsNullOrWhiteSpace(ToText))
            {
                AppendLog("ERROR: To field is required when sending.");
                return;
            }
            if (send && IsBusinessTrip && CreateExcel && string.IsNullOrWhiteSpace(ExcelSaveFolder))
            {
                AppendLog("ERROR: Excel save folder is required when 'Create and fill allowance Excel' is enabled.");
                return;
            }
            if (EndDate < StartDate)
            {
                AppendLog("ERROR: End date must be on or after Start date.");
                return;
            }

            IsBusy = true;
            StatusLog = string.Empty;

            var request = new OutOfOfficeRequest
            {
                LeaveType = _selectedLeaveType?.LeaveType ?? LeaveType.FullDayOff,
                StartDate = StartDate,
                EndDate = EndDate,
                Subject = Subject,
                Location = Location,
                ToRecipients = ParseAddresses(ToText),
                CcRecipients = ParseAddresses(CcText),
                SetAutoReplies = SetAutoReplies,
                CreateExcel = CreateExcel,
                ExcelSaveFolder = ExcelSaveFolder,
            };

            bool meetingDone = false;
            bool oofDone = false;
            bool excelDone = false;

            try
            {
                // 1. Create / send meeting
                AppendLog(send ? "Creating and sending meeting…" : "Creating draft meeting…");
                _meetingService.CreateOrSend(request, send);
                meetingDone = true;
                AppendLog(send ? "✔ Meeting sent." : "✔ Draft saved.");

                // Save mailing list after successful send
                if (send)
                    MailingListService.Save(request.ToRecipients, request.CcRecipients);

                // 2. OOF
                if (request.SetAutoReplies)
                {
                    AppendLog("Setting auto-reply via Microsoft Graph…");
                    var sig = SignatureService.GetDefaultSignatureHtml();
                    await _oofService.SetAsync(request.StartDate, request.EndDate, sig);
                    oofDone = true;
                    AppendLog("✔ Auto-reply configured.");
                }

                // 3. Excel (business trip only)
                if (IsBusinessTrip && request.CreateExcel)
                {
                    AppendLog("Downloading Excel template and filling data…");
                    var savedPath = await _excelService.CreateAsync(
                        request.StartDate,
                        request.EndDate,
                        request.Location,
                        _familyName,
                        request.ExcelSaveFolder);
                    excelDone = true;
                    AppendLog($"✔ Excel saved: {savedPath}");
                }

                AppendLog("All tasks completed successfully.");
            }
            catch (Exception ex)
            {
                AppendLog($"ERROR: {ex.Message}");

                // Summary of what completed vs what failed
                var summary = new System.Text.StringBuilder();
                summary.AppendLine("Status summary:");
                summary.AppendLine($"  Meeting: {(meetingDone ? "✔ Done" : "✘ Not done")}");
                if (request.SetAutoReplies)
                    summary.AppendLine($"  Auto-reply: {(oofDone ? "✔ Done" : "✘ Failed")}");
                if (IsBusinessTrip && request.CreateExcel)
                    summary.AppendLine($"  Excel: {(excelDone ? "✔ Done" : "✘ Failed")}");

                AppendLog(summary.ToString());
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void AppendLog(string message)
        {
            StatusLog = string.IsNullOrEmpty(StatusLog)
                ? message
                : StatusLog + Environment.NewLine + message;
        }

        // ------------------------------------------------------------------ INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(name);
            return true;
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>Wraps a <see cref="LeaveType"/> with a display label for the ComboBox.</summary>
    public class LeaveTypeItem
    {
        public LeaveType LeaveType { get; }
        public string DisplayName { get; }
        public LeaveTypeItem(LeaveType type, string name) { LeaveType = type; DisplayName = name; }
        public override string ToString() => DisplayName;
    }
}

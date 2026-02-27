using OutOfOfficeAddin.Models;

namespace OutOfOfficeAddin.Services
{
    /// <summary>
    /// Generates the meeting subject based on leave type and the user's family name.
    /// </summary>
    public static class SubjectHelper
    {
        /// <summary>
        /// Returns the subject string, e.g. "Yamada BT", "Yamada OFF".
        /// </summary>
        public static string Build(string familyName, LeaveType leaveType)
        {
            switch (leaveType)
            {
                case LeaveType.BusinessTrip:
                    return $"{familyName} BT";
                case LeaveType.FullDayOff:
                    return $"{familyName} OFF";
                case LeaveType.AmHalfDayOff:
                    return $"{familyName} AM OFF";
                case LeaveType.PmHalfDayOff:
                    return $"{familyName} PM OFF";
                default:
                    return $"{familyName} OFF";
            }
        }

        /// <summary>
        /// Returns the default location for the given leave type.
        /// Business Trip has no default; all Off types default to "Home".
        /// </summary>
        public static string DefaultLocation(LeaveType leaveType)
        {
            return leaveType == LeaveType.BusinessTrip ? string.Empty : "Home";
        }
    }
}

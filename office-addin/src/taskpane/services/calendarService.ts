import { getAccessToken } from "./authService";
import { LeaveType, getSubjectSuffix, getDefaultLocation } from "./subjectHelper";
import { toISODateStr, addDays } from "./subjectHelper";

export interface CalendarEventRequest {
  familyName: string;
  leaveType: LeaveType;
  startDate: Date;
  endDate: Date;
  location: string;
  toRecipients: string[];
  ccRecipients: string[];
}

export async function createCalendarEvent(req: CalendarEventRequest): Promise<{ id: string }> {
  const token = await getAccessToken();

  const subject = `${req.familyName} ${getSubjectSuffix(req.leaveType)}`;
  // All-day events in Graph: endDate is exclusive (next day)
  const endDateExclusive = addDays(req.endDate, 1);

  const body = {
    subject,
    start: {
      dateTime: `${toISODateStr(req.startDate)}T00:00:00`,
      timeZone: "UTC",
    },
    end: {
      dateTime: `${toISODateStr(endDateExclusive)}T00:00:00`,
      timeZone: "UTC",
    },
    isAllDay: true,
    location: { displayName: req.location || getDefaultLocation(req.leaveType) },
    showAs: "free",
    isReminderOn: false,
    attendees: [
      ...req.toRecipients.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
      ...req.ccRecipients.map((email) => ({
        emailAddress: { address: email },
        type: "optional",
      })),
    ],
  };

  const resp = await fetch("https://graph.microsoft.com/v1.0/me/events", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Failed to create calendar event: ${resp.status} ${err}`);
  }

  return resp.json();
}



export type LeaveType = "BT" | "OFF" | "AM_OFF" | "PM_OFF";

export const LEAVE_TYPE_LABELS: Record<LeaveType, string> = {
  BT: "Business Trip",
  OFF: "Full Day Off",
  AM_OFF: "AM Half Day Off",
  PM_OFF: "PM Half Day Off",
};

export function getSubjectSuffix(type: LeaveType): string {
  switch (type) {
    case "BT": return "BT";
    case "OFF": return "OFF";
    case "AM_OFF": return "AM OFF";
    case "PM_OFF": return "PM OFF";
  }
}

export function getDefaultLocation(type: LeaveType): string {
  return type === "BT" ? "" : "Home";
}

export function buildSubject(familyName: string, type: LeaveType): string {
  return `${familyName} ${getSubjectSuffix(type)}`;
}

export function formatDate(d: Date): string {
  return d.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
}

export function addDays(d: Date, n: number): Date {
  const result = new Date(d);
  result.setDate(result.getDate() + n);
  return result;
}

export function toISODateStr(d: Date): string {
  return d.toISOString().split("T")[0];
}

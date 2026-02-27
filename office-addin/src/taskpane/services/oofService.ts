import { getAccessToken } from "./authService";
import { addDays, toISODateStr } from "./subjectHelper";

export interface OofSettings {
  startDate: Date;
  endDate: Date;
  internalMessage: string;
  externalMessage: string;
}

function buildInternalMessage(endDate: Date, signature: string): string {
  const returnDate = addDays(endDate, 1);
  const returnStr = returnDate.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const msg = `<p>Thank you for your message. I am currently out of office and will return on ${returnStr}.</p>
<p>I will respond to your email as soon as possible upon my return.</p>
<p>If you need immediate assistance, please contact my colleague.</p>`;
  return signature ? `${msg}${signature}` : msg;
}

function buildExternalMessage(endDate: Date, signature: string): string {
  const returnDate = addDays(endDate, 1);
  const returnStr = returnDate.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const returnStrJa = returnDate.toLocaleDateString("ja-JP", { year: "numeric", month: "long", day: "numeric" });
  const msg = `<p>Thank you for your message. I am currently out of office and will return on ${returnStr}.</p>
<p>I will respond to your email as soon as possible upon my return.</p>
<p>If you need immediate assistance, please contact my colleague.</p>
<hr/>
<p>メールありがとうございます。現在不在にしており、${returnStrJa}に戻る予定です。</p>
<p>返信までしばらくお待ちください。</p>`;
  return signature ? `${msg}${signature}` : msg;
}

export function buildOofMessages(endDate: Date, signature: string) {
  return {
    internal: buildInternalMessage(endDate, signature),
    external: buildExternalMessage(endDate, signature),
  };
}

function toISODateTimeStr(date: Date): string {
  return `${toISODateStr(date)}T00:00:00`;
}

export async function setOof(settings: OofSettings): Promise<void> {
  const token = await getAccessToken();
  const scheduledEndDate = addDays(settings.endDate, 1);

  const body = {
    automaticRepliesSetting: {
      status: "scheduled",
      scheduledStartDateTime: {
        dateTime: toISODateTimeStr(settings.startDate),
        timeZone: "UTC",
      },
      scheduledEndDateTime: {
        dateTime: toISODateTimeStr(scheduledEndDate),
        timeZone: "UTC",
      },
      internalReplyMessage: settings.internalMessage,
      externalReplyMessage: settings.externalMessage,
      externalAudience: "all",
    },
  };

  const resp = await fetch("https://graph.microsoft.com/v1.0/me/mailboxSettings", {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Failed to set OOF: ${resp.status} ${err}`);
  }
}

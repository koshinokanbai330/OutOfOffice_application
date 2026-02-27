import { getAccessToken } from "./authService";

const APP_FOLDER_PATH = "/drive/special/approot";
const FILE_NAME = "mailingList.json";

export interface MailingList {
  to: string[];
  cc: string[];
  updatedAt: string;
}

export async function loadMailingList(): Promise<MailingList | null> {
  try {
    const token = await getAccessToken();
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0${APP_FOLDER_PATH}:/${FILE_NAME}:/content`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (resp.status === 404) return null;
    if (!resp.ok) throw new Error(`Graph error: ${resp.status}`);
    return resp.json();
  } catch {
    return null;
  }
}

export async function saveMailingList(list: MailingList): Promise<void> {
  const token = await getAccessToken();
  const body = JSON.stringify({ ...list, updatedAt: new Date().toISOString() });
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0${APP_FOLDER_PATH}:/${FILE_NAME}:/content`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body,
    }
  );
  if (!resp.ok) throw new Error(`Failed to save mailing list: ${resp.status}`);
}

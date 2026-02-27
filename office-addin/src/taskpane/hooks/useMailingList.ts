import { useState, useEffect } from "react";
import { loadMailingList, saveMailingList, MailingList } from "../services/mailingListService";

export function useMailingList() {
  const [mailingList, setMailingList] = useState<MailingList | null>(null);

  useEffect(() => {
    loadMailingList().then((list) => {
      if (list) setMailingList(list);
    });
  }, []);

  const save = async (to: string[], cc: string[]) => {
    const updated: MailingList = { to, cc, updatedAt: new Date().toISOString() };
    await saveMailingList(updated);
    setMailingList(updated);
  };

  return { mailingList, save };
}

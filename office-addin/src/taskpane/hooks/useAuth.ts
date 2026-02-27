import { useState, useEffect } from "react";
import { getAccessToken } from "../services/authService";

/* global Office */

export interface UserProfile {
  displayName: string;
  surname: string;
  givenName: string;
  mail: string;
}

export function useAuth() {
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    (async () => {
      try {
        const token = await getAccessToken();
        const resp = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,surname,givenName,mail", {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (!resp.ok) throw new Error(`Profile fetch failed: ${resp.status}`);
        const data = await resp.json();
        setProfile(data);
      } catch (e) {
        setError(String(e));
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  return { profile, error, loading };
}

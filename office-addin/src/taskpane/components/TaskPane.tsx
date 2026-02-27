import React, { useState, useEffect, useCallback } from "react";
import { LeaveType, LEAVE_TYPE_LABELS, buildSubject, getDefaultLocation, addDays } from "../services/subjectHelper";
import { createCalendarEvent } from "../services/calendarService";
import { setOof, buildOofMessages } from "../services/oofService";
import { saveMailingList, loadMailingList } from "../services/mailingListService";
import { createExcel } from "../services/excelService";
import { useAuth } from "../hooks/useAuth";
import StatusLog from "./StatusLog";

/* global fetch */

const FIELD_STYLE: React.CSSProperties = {
  width: "100%",
  boxSizing: "border-box",
  padding: "4px 6px",
  fontSize: 13,
  border: "1px solid #ccc",
  borderRadius: 3,
  marginBottom: 6,
};

const LABEL_STYLE: React.CSSProperties = {
  fontSize: 12,
  fontWeight: 600,
  color: "#444",
  display: "block",
  marginBottom: 2,
};

const SECTION_STYLE: React.CSSProperties = {
  marginBottom: 10,
};

function parseEmails(raw: string): string[] {
  return raw.split(";").map((s) => s.trim()).filter(Boolean);
}

const today = new Date();
const todayStr = today.toISOString().split("T")[0];

const TaskPane: React.FC = () => {
  const { profile, error: authError, loading: authLoading } = useAuth();

  const [leaveType, setLeaveType] = useState<LeaveType>("BT");
  const [startDate, setStartDate] = useState<string>(todayStr);
  const [endDate, setEndDate] = useState<string>(todayStr);
  const [location, setLocation] = useState<string>("");
  const [toField, setToField] = useState<string>("");
  const [ccField, setCcField] = useState<string>("");
  const [setAutoReply, setSetAutoReply] = useState<boolean>(true);
  const [signature, setSignature] = useState<string>("");
  const [createExcelCheck, setCreateExcelCheck] = useState<boolean>(true);
  const [destination, setDestination] = useState<string>("");
  const [logs, setLogs] = useState<string[]>([]);
  const [busy, setBusy] = useState<boolean>(false);
  const [templateBytes, setTemplateBytes] = useState<ArrayBuffer | null>(null);

  // Load mailing list on mount
  useEffect(() => {
    loadMailingList().then((list) => {
      if (list) {
        setToField(list.to.join("; "));
        setCcField(list.cc.join("; "));
      }
    });
  }, []);

  // Update location default when type changes
  useEffect(() => {
    setLocation(getDefaultLocation(leaveType));
  }, [leaveType]);

  // Fetch template bytes on mount
  useEffect(() => {
    fetch("/assets/template.xlsx")
      .then((r) => r.arrayBuffer())
      .then(setTemplateBytes)
      .catch((err) => console.warn("Failed to load Excel template:", err));
  }, []);

  const familyName = profile?.surname || profile?.displayName?.split(" ")[0] || "User";
  const subject = buildSubject(familyName, leaveType);
  const endDateObj = endDate ? new Date(endDate) : new Date();
  const oofMessages = buildOofMessages(endDateObj, signature);

  const log = useCallback((msg: string) => {
    setLogs((prev) => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);
  }, []);

  const handleCancel = () => {
    setLeaveType("BT");
    setStartDate(todayStr);
    setEndDate(todayStr);
    setLocation(getDefaultLocation("BT"));
    setToField("");
    setCcField("");
    setSetAutoReply(true);
    setSignature("");
    setCreateExcelCheck(true);
    setDestination("");
    setLogs([]);
  };

  const runSend = async (asDraft: boolean) => {
    if (!asDraft && !toField.trim()) {
      setLogs(["ERROR: To field is required for sending."]);
      return;
    }
    setBusy(true);
    setLogs([]);
    const toList = parseEmails(toField);
    const ccList = parseEmails(ccField);
    const start = new Date(startDate);
    const end = new Date(endDate);

    try {
      log("Creating calendar event...");
      await createCalendarEvent({
        familyName,
        leaveType,
        startDate: start,
        endDate: end,
        location,
        toRecipients: asDraft ? [] : toList,
        ccRecipients: asDraft ? [] : ccList,
      });
      log(`✓ Calendar event ${asDraft ? "draft " : ""}created: ${subject}`);

      if (!asDraft) {
        // Save mailing list
        try {
          await saveMailingList({ to: toList, cc: ccList, updatedAt: new Date().toISOString() });
          log("✓ Mailing list saved to OneDrive.");
        } catch (e) {
          log(`WARNING: Could not save mailing list: ${e}`);
        }

        // OOF
        if (setAutoReply) {
          try {
            log("Setting automatic replies (OOF)...");
            await setOof({
              startDate: start,
              endDate: end,
              internalMessage: oofMessages.internal,
              externalMessage: oofMessages.external,
            });
            log("✓ Automatic replies configured.");
          } catch (e) {
            log(`ERROR: OOF failed: ${e}`);
          }
        }

        // Excel
        if (leaveType === "BT" && createExcelCheck) {
          if (!templateBytes) {
            log("WARNING: Excel template not loaded, skipping Excel creation.");
          } else {
            try {
              log("Creating travel allowance Excel...");
              const url = await createExcel(
                { familyName, startDate: start, endDate: end, destination },
                templateBytes
              );
              log(`✓ Excel saved to OneDrive: ${url}`);
            } catch (e) {
              log(`ERROR: Excel creation failed: ${e}`);
            }
          }
        }
      }
    } catch (e) {
      log(`ERROR: ${e}`);
    } finally {
      setBusy(false);
    }
  };

  if (authLoading) {
    return (
      <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
        <p>Loading...</p>
      </div>
    );
  }

  if (authError) {
    return (
      <div style={{ padding: 16, fontFamily: "Segoe UI, sans-serif" }}>
        <p style={{ color: "red" }}>Authentication error: {authError}</p>
      </div>
    );
  }

  return (
    <div style={{ padding: "12px 14px", fontFamily: "Segoe UI, sans-serif", fontSize: 13 }}>
      <h3 style={{ margin: "0 0 12px", fontSize: 15, color: "#0078d4" }}>Out of Office</h3>

      {/* Type */}
      <div style={SECTION_STYLE}>
        <label style={LABEL_STYLE}>Type</label>
        <select
          style={FIELD_STYLE}
          value={leaveType}
          onChange={(e) => setLeaveType(e.target.value as LeaveType)}
          disabled={busy}
        >
          {(Object.keys(LEAVE_TYPE_LABELS) as LeaveType[]).map((k) => (
            <option key={k} value={k}>{LEAVE_TYPE_LABELS[k]}</option>
          ))}
        </select>
      </div>

      {/* Dates */}
      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
        <div style={{ flex: 1 }}>
          <label style={LABEL_STYLE}>Start date</label>
          <input type="date" style={FIELD_STYLE} value={startDate}
            onChange={(e) => { setStartDate(e.target.value); if (e.target.value > endDate) setEndDate(e.target.value); }}
            disabled={busy} />
        </div>
        <div style={{ flex: 1 }}>
          <label style={LABEL_STYLE}>End date</label>
          <input type="date" style={FIELD_STYLE} value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
            min={startDate}
            disabled={busy} />
        </div>
      </div>

      {/* Subject (read-only) */}
      <div style={SECTION_STYLE}>
        <label style={LABEL_STYLE}>Subject (auto)</label>
        <input type="text" style={{ ...FIELD_STYLE, background: "#f5f5f5", color: "#555" }}
          value={subject} readOnly />
      </div>

      {/* Location */}
      <div style={SECTION_STYLE}>
        <label style={LABEL_STYLE}>Location</label>
        <input type="text" style={FIELD_STYLE} value={location}
          onChange={(e) => setLocation(e.target.value)} disabled={busy}
          placeholder={leaveType === "BT" ? "" : "Home"} />
      </div>

      {/* To / Cc */}
      <div style={SECTION_STYLE}>
        <label style={LABEL_STYLE}>To (semicolon-separated)</label>
        <input type="text" style={FIELD_STYLE} value={toField}
          onChange={(e) => setToField(e.target.value)} disabled={busy}
          placeholder="email1@example.com; email2@example.com" />
      </div>
      <div style={SECTION_STYLE}>
        <label style={LABEL_STYLE}>Cc (semicolon-separated)</label>
        <input type="text" style={FIELD_STYLE} value={ccField}
          onChange={(e) => setCcField(e.target.value)} disabled={busy}
          placeholder="email3@example.com" />
      </div>

      {/* Auto replies checkbox */}
      <div style={{ ...SECTION_STYLE, display: "flex", alignItems: "center", gap: 6 }}>
        <input type="checkbox" id="setAutoReply" checked={setAutoReply}
          onChange={(e) => setSetAutoReply(e.target.checked)} disabled={busy} />
        <label htmlFor="setAutoReply" style={{ fontSize: 12 }}>Set automatic replies for this period</label>
      </div>

      {/* OOF previews */}
      {setAutoReply && (
        <>
          <div style={SECTION_STYLE}>
            <label style={LABEL_STYLE}>Internal message (preview)</label>
            <div style={{
              ...FIELD_STYLE,
              background: "#f9f9f9",
              minHeight: 60,
              overflow: "auto",
              fontSize: 11,
            }}
              dangerouslySetInnerHTML={{ __html: oofMessages.internal }} />
          </div>
          <div style={SECTION_STYLE}>
            <label style={LABEL_STYLE}>External message (preview)</label>
            <div style={{
              ...FIELD_STYLE,
              background: "#f9f9f9",
              minHeight: 60,
              overflow: "auto",
              fontSize: 11,
            }}
              dangerouslySetInnerHTML={{ __html: oofMessages.external }} />
          </div>
        </>
      )}

      {/* Signature */}
      <div style={SECTION_STYLE}>
        <label style={LABEL_STYLE}>Signature (optional, HTML allowed)</label>
        <textarea
          style={{ ...FIELD_STYLE, minHeight: 60, resize: "vertical", fontFamily: "monospace", fontSize: 11 }}
          value={signature}
          onChange={(e) => setSignature(e.target.value)}
          disabled={busy}
          placeholder="<p>Best regards,<br/>Your Name</p>"
        />
      </div>

      {/* Business Trip section */}
      {leaveType === "BT" && (
        <div style={{
          border: "1px solid #0078d4",
          borderRadius: 4,
          padding: "8px 10px",
          marginBottom: 10,
          background: "#f0f6ff",
        }}>
          <div style={{ fontWeight: 600, color: "#0078d4", marginBottom: 8, fontSize: 13 }}>
            Business Trip
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
            <input type="checkbox" id="createExcel" checked={createExcelCheck}
              onChange={(e) => setCreateExcelCheck(e.target.checked)} disabled={busy} />
            <label htmlFor="createExcel" style={{ fontSize: 12 }}>
              Create and fill allowance Excel (saved to OneDrive)
            </label>
          </div>
          {createExcelCheck && (
            <div>
              <label style={LABEL_STYLE}>Destination</label>
              <input type="text" style={FIELD_STYLE} value={destination}
                onChange={(e) => setDestination(e.target.value)} disabled={busy}
                placeholder="e.g. Tokyo, Japan" />
            </div>
          )}
        </div>
      )}

      {/* Buttons */}
      <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
        <button
          style={{
            flex: 1,
            padding: "7px 0",
            background: "#fff",
            border: "1px solid #0078d4",
            color: "#0078d4",
            borderRadius: 3,
            cursor: busy ? "not-allowed" : "pointer",
            fontSize: 13,
          }}
          onClick={() => runSend(true)}
          disabled={busy}
        >
          Create draft
        </button>
        <button
          style={{
            flex: 1,
            padding: "7px 0",
            background: busy ? "#aaa" : "#0078d4",
            border: "none",
            color: "#fff",
            borderRadius: 3,
            cursor: busy ? "not-allowed" : "pointer",
            fontSize: 13,
          }}
          onClick={() => runSend(false)}
          disabled={busy}
        >
          Send
        </button>
        <button
          style={{
            flex: 1,
            padding: "7px 0",
            background: "#fff",
            border: "1px solid #ccc",
            color: "#555",
            borderRadius: 3,
            cursor: busy ? "not-allowed" : "pointer",
            fontSize: 13,
          }}
          onClick={handleCancel}
          disabled={busy}
        >
          Cancel
        </button>
      </div>

      <StatusLog messages={logs} />
    </div>
  );
};

export default TaskPane;

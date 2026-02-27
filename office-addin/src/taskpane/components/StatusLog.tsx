import React from "react";

interface Props {
  messages: string[];
}

const StatusLog: React.FC<Props> = ({ messages }) => {
  if (messages.length === 0) return null;
  return (
    <div style={{
      marginTop: 12,
      padding: "8px 10px",
      background: "#f3f3f3",
      border: "1px solid #ddd",
      borderRadius: 4,
      fontSize: 12,
      maxHeight: 120,
      overflowY: "auto",
      fontFamily: "monospace",
    }}>
      {messages.map((m, i) => (
        <div key={i} style={{ color: m.startsWith("ERROR") ? "#c00" : "#333" }}>{m}</div>
      ))}
    </div>
  );
};

export default StatusLog;

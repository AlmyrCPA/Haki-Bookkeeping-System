import React from "react";
import { Plus } from "lucide-react";

// Dropdown in the sidebar, above the nav. Reads its list from /clientIndex (via
// App.jsx). Switching just re-points the Firestore onSnapshot listener in
// BookkeepingApp at the newly selected clientId.
export default function ClientSwitcher({ clients, currentClientId, onSwitch, onAddClient }) {
  return (
    <>
      <div className="client-switch-label">Client</div>
      <div className="client-switch-row">
        <select
          className="io-select"
          value={currentClientId || ""}
          onChange={(e) => onSwitch(e.target.value)}
        >
          {clients.length === 0 && <option value="">No clients yet</option>}
          {clients.map((c) => (
            <option key={c.id} value={c.id}>{c.name || "(unnamed)"}</option>
          ))}
        </select>
        <button
          type="button"
          className="client-add-btn"
          title="Add new client"
          onClick={onAddClient}
        >
          <Plus size={15} />
        </button>
      </div>
    </>
  );
}

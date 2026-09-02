import React, { useState } from "react";
import { collection, doc, setDoc } from "firebase/firestore";
import { db } from "./firebase";
import {
  EntryModalShell,
  LabeledField,
  Field,
  makeInitialData,
  emptyCompany,
  PARTY_TYPES,
} from "./BookkeepingApp";

// Phase 1 "Add New Client" modal. Only Name / TIN / Taxpayer Type up front —
// enough to identify the client in the switcher. Everything else (RDO, address,
// VAT status, full Chart of Accounts) is filled in later from Company Details.
export default function AddClientModal({ onCancel, onCreate }) {
  const [name, setName] = useState("");
  const [tin, setTin] = useState("");
  const [taxpayerType, setTaxpayerType] = useState("Non-Individual");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const handleCreate = async () => {
    if (!name.trim()) { setError("Company name is required."); return; }
    setError("");
    setBusy(true);
    try {
      const clientId = doc(collection(db, "clients")).id;
      const seed = {
        ...makeInitialData(),
        company: { ...emptyCompany, name: name.trim(), tin: tin.trim(), taxpayerType },
      };
      await setDoc(doc(db, "clients", clientId), seed);
      await setDoc(doc(db, "clientIndex", clientId), { name: name.trim(), tin: tin.trim() });
      onCreate(clientId);
    } catch (err) {
      console.error("Failed to create client:", err);
      setError("Couldn't create the client. Please try again.");
      setBusy(false);
    }
  };

  return (
    <EntryModalShell
      title="Add New Client"
      submitLabel={busy ? "Creating…" : "Create client"}
      onCancel={onCancel}
      onSubmit={busy ? () => {} : handleCreate}
    >
      <LabeledField label="Company Name">
        <Field value={name} onChange={setName} placeholder="Your Company Name Inc." />
      </LabeledField>
      <LabeledField label="TIN">
        <Field value={tin} onChange={setTin} placeholder="000-000-000-000" />
      </LabeledField>
      <LabeledField label="Taxpayer Type">
        <Field type="select" options={PARTY_TYPES} value={taxpayerType} onChange={setTaxpayerType} />
      </LabeledField>
      {error && <div className="qf-hint" style={{ color: "#A13D3D" }}>{error}</div>}
      <div className="qf-hint">
        Everything else — RDO, address, VAT status, the full Chart of Accounts, etc. — can be
        filled in from Company Details once the client is created.
      </div>

      {/* Self-contained styling so the modal renders correctly even before the main
          app tree (and its <Style/>) is mounted — e.g. on the first-client screen. */}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600&family=IBM+Plex+Sans:wght@400;500;600&display=swap');
        .modal-overlay { position: fixed; inset: 0; background: rgba(15,42,77,0.35); display: flex; align-items: center; justify-content: center; z-index: 60; padding: 20px; font-family: 'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif; }
        .modal-box { background: #fff; border-radius: 12px; max-width: 620px; width: 100%; max-height: 86vh; overflow-y: auto; box-shadow: 0 24px 64px rgba(15,42,77,0.28); color: #16233A; }
        .modal-head { display: flex; align-items: center; justify-content: space-between; padding: 18px 22px; border-bottom: 1px solid #C7DAF0; position: sticky; top: 0; background: #fff; }
        .modal-head h2 { font-family: 'Fraunces', serif; font-size: 17px; font-weight: 600; color: #0F2A4D; margin: 0; }
        .modal-close { background: none; border: none; cursor: pointer; color: #55697F; padding: 4px; border-radius: 6px; }
        .modal-body { padding: 20px 22px; display: grid; grid-template-columns: 1fr 1fr; gap: 14px 18px; }
        .modal-foot { display: flex; justify-content: flex-end; gap: 10px; padding: 14px 22px; border-top: 1px solid #C7DAF0; position: sticky; bottom: 0; background: #fff; }
        .btn-outline { background: #fff; border: 1px solid #C7DAF0; color: #16233A; padding: 9px 18px; border-radius: 7px; font-size: 13px; font-weight: 500; cursor: pointer; }
        .primary-btn { margin-left: auto; background: #1D5FA8; color: #fff; border: none; padding: 9px 16px; border-radius: 7px; font-size: 13px; font-weight: 500; cursor: pointer; white-space: nowrap; }
        .form-field { display: flex; flex-direction: column; gap: 5px; }
        .form-field.wide { grid-column: 1 / -1; }
        .form-field label { font-size: 11.5px; color: #55697F; font-weight: 500; }
        .ledger-field { width: 100%; background: #fff; border: 1px solid #C7DAF0; border-radius: 6px; padding: 8px 10px; font-size: 12.5px; font-family: inherit; color: #16233A; outline: none; }
        .ledger-field:focus { border-color: #1D5FA8; }
        .qf-hint { font-size: 11.5px; color: #55697F; line-height: 1.45; grid-column: 1 / -1; }
        @media (max-width: 900px) { .modal-body { grid-template-columns: 1fr; } }
      `}</style>
    </EntryModalShell>
  );
}

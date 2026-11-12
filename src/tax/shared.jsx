// Shared building blocks for the BIR return worksheets (1601-C, 2550Q/2551Q, 0619-E/1601-EQ,
// 1701Q/1702Q, 1701/1702-RT, 1604-C/1604-E).
//
// Each return is described by a schema (sections of numbered lines) plus a compute function built on
// makeCalc(). The same schema + computed values drive the on-screen form, the A4 PDF encoding worksheet
// and the Excel export, so the three never disagree.
//
// Storage: data[storeKey][periodKey] = {
//   m: { [key]: typed value },          manual items and header answers
//   overrides: { [key]: value },        typed replacements for auto-filled items ("Reset to computed" deletes)
//   payment: { cashDebit, check, taxDebit, others: { bank, number, date, amount, specify } }, ror: { refNo, date },
//   attach: { [key]: true }, filed, dateFiled, filedRef, dueDate, preparedSnapshot, ...form-specific fields
// }
import React, { useState } from "react";
import { Pencil, Undo2, Check, Info, FileDown, CheckCircle2, Circle } from "lucide-react";
import * as XLSX from "xlsx";
import {
  fmtPlain, fmtParen, num, round2, MONTHS, pad2, todayMDY, todayFileStamp, parseAppDate, toMDY,
  Field, LabeledField, NumberField, EntryModalShell, BANK_ACCOUNTS, loadJsPDF, withTimeout, useReportExport,
  ExportStatus, reliefFormatTIN,
} from "../BookkeepingApp";

/* ============================== TAX CONSTANTS ============================== */

// Graduated income tax table for individuals (annual taxable income), keyed by the first year it
// applies so a later change never rewrites an earlier return. Each bracket: [upper limit, base tax, rate].
export const GRADUATED_TAX_TABLES = [
  { from: 2023, label: "TRAIN, 2023 onward", brackets: [[250000, 0, 0], [400000, 0, 0.15], [800000, 22500, 0.20], [2000000, 102500, 0.25], [8000000, 402500, 0.30], [Infinity, 2202500, 0.35]] },
  { from: 2018, label: "TRAIN, 2018 to 2022", brackets: [[250000, 0, 0], [400000, 0, 0.20], [800000, 30000, 0.25], [2000000, 130000, 0.30], [8000000, 490000, 0.32], [Infinity, 2410000, 0.35]] },
];
export function taxTableFor(year) {
  return GRADUATED_TAX_TABLES.find((t) => Number(year) >= t.from) || GRADUATED_TAX_TABLES[GRADUATED_TAX_TABLES.length - 1];
}
export function graduatedTax(income, year) {
  const x = num(income);
  if (x <= 0) return 0;
  let lower = 0;
  for (const [upper, base, rate] of taxTableFor(year).brackets) {
    if (x <= upper) return round2(base + (x - lower) * rate);
    lower = upper;
  }
  return 0;
}
export const BENEFITS_CEILING = 90000;      // 13th month pay and other benefits exempt up to this amount
export const EXEMPT_COMP_THRESHOLD = 250000; // annual compensation exempt for those with no other income
export const EIGHT_PCT_LIMIT = 3000000;      // gross sales / receipts cap for the 8% option
export const TIN_BRANCH_DIGITS = 5;          // eBIRForms 7.9.6.0 (RMC 36-2026) widened the branch code; confirm

export const WORKSHEET_NOTE = "Worksheet for eBIRForms encoding. Not the official BIR form.";

/* ============================== PERIOD HELPERS ============================== */

export const monthKey = (year, monthIdx) => `${year}-${pad2(monthIdx + 1)}`;
export const quarterKey = (year, q) => `${year}-Q${q}`;
export const mdy = (d) => (d ? `${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}/${d.getFullYear()}` : "");
export const quarterRange = (year, q) => ({ start: new Date(year, (q - 1) * 3, 1), end: new Date(year, q * 3, 0) });
export const quarterOfMonth = (monthIdx) => Math.floor(monthIdx / 3) + 1;

// The taxable year ending in `fyEndMonth` (1-12) of `year`. Calendar year when fyEndMonth is 12.
export function fiscalYearRange(year, fyEndMonth = 12) {
  const endM = Number(fyEndMonth) || 12;
  return { start: new Date(year - (endM === 12 ? 0 : 1), endM === 12 ? 0 : endM, 1), end: new Date(year, endM, 0) };
}
// Quarter q (1-3) of the fiscal year ending in fyEndMonth of `year`: cumulative start, quarter start and end.
export function fiscalQuarterRange(year, fyEndMonth, q) {
  const fy = fiscalYearRange(year, fyEndMonth);
  const qStart = new Date(fy.start.getFullYear(), fy.start.getMonth() + (q - 1) * 3, 1);
  const qEnd = new Date(fy.start.getFullYear(), fy.start.getMonth() + q * 3, 0);
  return { ytdStart: fy.start, start: qStart, end: qEnd };
}
export const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
export const inDates = (dateStr, start, end) => { const d = parseAppDate(dateStr); return !!d && d >= start && d <= end; };

/* ============================== COMPANY / PART I ============================== */

export function companyInfo(data) {
  const c = data.company || {};
  const indiv = c.taxpayerType === "Individual";
  const tin9 = String(c.tin || "").replace(/\D/g, "").slice(0, 9);
  const branch = String(c.branchCode || "").replace(/\D/g, "") || "0000";
  const name = indiv
    ? [c.surname, [c.firstName, c.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", ") || c.name || ""
    : (c.registeredName || c.name || "");
  return {
    c, indiv, tin9, branch,
    tinText: tin9 ? `${reliefFormatTIN(tin9)}-${branch}` : "",
    rdo: c.rdo || "", name,
    address: c.address || [c.addr1, c.addr2].filter(Boolean).join(", "),
    zip: c.zipCode || "", contact: c.contactNumber || "", email: c.email || "",
    category: c.withholdingCategory || "Private", topWA: c.topWithholdingAgent || "No",
    signatory: c.authorizedSignatory || "", signatoryTitle: c.signatoryPosition || "", signatoryTin: c.signatoryTin || "",
  };
}

// Lists blank Company Details fields a return needs, with a link to the page.
export function CompanyPrereqWarning({ data, fields, onNavigate }) {
  const c = data.company || {};
  const missing = fields.filter(([k]) => !String(c[k] ?? "").trim()).map(([, label]) => label);
  if (!missing.length) return null;
  return (
    <div className="callout">
      <Info size={16} />
      <div>
        <div className="callout-title">Company details incomplete</div>
        <div className="callout-body">Missing in Company Details: {missing.join(", ")}.{" "}
          {onNavigate && <button className="tx-link" onClick={() => onNavigate("company")}>Edit in Company Details</button>}
        </div>
      </div>
    </div>
  );
}

export function BackgroundBlock({ title = "Part I – Background Information", items, onNavigate, children }) {
  return (
    <div className="tx-card">
      <div className="tx-card-head">
        <span>{title}</span>
        {onNavigate && <button className="tx-link" onClick={() => onNavigate("company")}>Edit in Company Details</button>}
      </div>
      <div className="tx-bg-grid">
        {items.filter(Boolean).map(([no, label, value]) => (
          <div key={no + label} className="tx-bg-item">
            <span className="tx-no">{no}</span>
            <span className="tx-bg-label">{label}</span>
            <span className={"tx-bg-value" + (String(value ?? "").trim() ? "" : " empty")}>{String(value ?? "").trim() || "— not set —"}</span>
          </div>
        ))}
      </div>
      {children}
    </div>
  );
}

/* ============================== RETURN STORAGE ============================== */

export function useReturn(data, setData, storeKey, periodKey, defaults) {
  const stored = ((data[storeKey] || {})[periodKey]) || null;
  const form = stored || { m: {}, overrides: {}, ...(defaults || {}) };
  const write = (fn) => setData((d) => {
    const all = d[storeKey] || {};
    const cur = all[periodKey] || { m: {}, overrides: {}, ...(defaults || {}) };
    return { ...d, [storeKey]: { ...all, [periodKey]: fn(cur) } };
  });
  return {
    form, exists: !!stored,
    patch: (obj) => write((cur) => ({ ...cur, ...obj })),
    update: (fn) => write(fn),
    setM: (k, v) => write((cur) => ({ ...cur, m: { ...(cur.m || {}), [k]: v } })),
    setOverride: (k, v) => write((cur) => ({ ...cur, overrides: { ...(cur.overrides || {}), [k]: v } })),
    resetOverride: (k) => write((cur) => { const o = { ...(cur.overrides || {}) }; delete o[k]; return { ...cur, overrides: o }; }),
  };
}

// Value engine. c.a(key, computed) = auto item (override wins); c.man(key) = typed item; c.set(key, v) = formula.
export function makeCalc(form) {
  const ov = (form && form.overrides) || {};
  const m = (form && form.m) || {};
  const auto = {}, v = {};
  const has = (k) => ov[k] !== undefined && ov[k] !== null && ov[k] !== "";
  return {
    auto, v,
    overridden: (k) => has(k),
    m: (k) => num(m[k]),
    t: (k, dflt = "") => (m[k] === undefined || m[k] === null || m[k] === "" ? dflt : m[k]),
    a: (k, computed) => { const cv = round2(num(computed)); auto[k] = cv; v[k] = has(k) ? round2(num(ov[k])) : cv; return v[k]; },
    man: (k) => { v[k] = round2(num(m[k])); return v[k]; },
    set: (k, val) => { v[k] = round2(num(val)); return v[k]; },
  };
}

/* ============================== FORM RENDERING ============================== */

function AutoCell({ value, computed, overridden, onOverride, onReset, src, disabled }) {
  const [editing, setEditing] = useState(false);
  if (!disabled && (editing || overridden)) {
    return (
      <div className={"tx-auto" + (overridden ? " ovr" : "")} title={src}>
        {overridden && <span className="tx-dot" title="Overridden" />}
        <NumberField value={value} onChange={(x) => onOverride(x)} align="right" />
        <button className="tx-icon-btn" title={`Reset to computed (${fmtParen(computed)})`} onClick={() => { onReset(); setEditing(false); }}><Undo2 size={12} /></button>
      </div>
    );
  }
  return (
    <div className="tx-auto" title={src ? `Source: ${src}` : undefined}>
      <span className="tx-amt">{fmtParen(value)}</span>
      {!disabled && <button className="tx-icon-btn" title="Override this auto-filled amount" onClick={() => setEditing(true)}><Pencil size={11} /></button>}
    </div>
  );
}

// A single auto-filled amount outside a FormSection (e.g. inside a schedule table).
export function AutoAmount({ calc, rt, k, src, disabled }) {
  return <AutoCell value={calc.v[k]} computed={calc.auto[k]} overridden={calc.overridden(k)} src={src} disabled={disabled}
    onOverride={(x) => rt.setOverride(k, x)} onReset={() => rt.resetOverride(k)} />;
}

const asArr = (x) => (Array.isArray(x) ? x : [x]);

function renderCell(line, k, t, i, calc, rt, disabledAll) {
  if (!k || !t) return <td key={i} className="tx-blank"></td>;
  const disabled = disabledAll || (typeof line.dis === "function" ? line.dis() : !!line.dis);
  const val = calc.v[k];
  const amtClass = line.due ? (val < 0 ? " tx-over" : val > 0 ? " tx-due" : "") : "";
  if (t === "calc") return <td key={i} className={"num tx-calc" + amtClass}>{fmtParen(val)}</td>;
  if (t === "auto") {
    return (
      <td key={i} className={"num" + amtClass}>
        <AutoCell value={val} computed={calc.auto[k]} overridden={calc.overridden(k)} src={line.src} disabled={disabled}
          onOverride={(x) => rt.setOverride(k, x)} onReset={() => rt.resetOverride(k)} />
      </td>
    );
  }
  if (t === "man") return <td key={i} className="num"><NumberField value={(rt.form.m || {})[k]} onChange={(x) => rt.setM(k, x)} align="right" disabled={disabled} /></td>;
  const raw = (rt.form.m || {})[k];
  const shown = raw === undefined || raw === null || raw === "" ? (line.dflt ?? "") : raw;
  if (t === "yesno") return <td key={i}><Field type="select" options={["No", "Yes"]} value={shown} onChange={(x) => rt.setM(k, x)} disabled={disabled} /></td>;
  if (t === "select") return <td key={i}><Field type="select" options={line.opts || []} value={shown} onChange={(x) => rt.setM(k, x)} disabled={disabled} /></td>;
  if (t === "date") return <td key={i}><Field type="date" value={shown} onChange={(x) => rt.setM(k, x)} disabled={disabled} /></td>;
  if (t === "int") return <td key={i}><Field value={shown} onChange={(x) => rt.setM(k, String(x).replace(/[^\d]/g, ""))} align="right" disabled={disabled} /></td>;
  if (t === "ro") return <td key={i}><span className="tx-ro">{line.text ?? shown}</span></td>;
  return <td key={i}><Field value={shown} onChange={(x) => rt.setM(k, x)} disabled={disabled} placeholder={line.ph} /></td>;
}

// section = { title, cols?: ["Amount"], note?, lines: [{ no, label, k, t, src, strong, due, dis, hide, f, note }] }
export function FormSection({ section, calc, rt, disabled }) {
  if (section.hide) return null;
  const cols = section.cols || ["Amount"];
  const lines = section.lines.filter((l) => l && !l.hide);
  return (
    <div className="tx-card">
      {section.title && <div className="tx-card-head"><span>{section.title}</span>{section.headRight}</div>}
      {section.note && <div className="tx-note">{section.note}</div>}
      <div className="table-scroll">
        <table className="tx-form">
          {cols.length > 1 || section.showColHead ? (
            <thead><tr><th style={{ width: 52 }}>Item</th><th>Description</th>{cols.map((c) => <th key={c} className="num-head tx-col">{c}</th>)}</tr></thead>
          ) : null}
          <tbody>
            {lines.map((l, idx) => {
              if (l.heading) return <tr key={"h" + idx} className="tx-subhead"><td colSpan={2 + cols.length}>{l.heading}</td></tr>;
              const ks = asArr(l.k), ts = asArr(l.t);
              return (
                <tr key={(l.no || "") + (ks[0] || idx)} className={l.strong ? "tx-strong" : ""}>
                  <td className="tx-no">{l.no}</td>
                  <td className="tx-label">{l.label}{l.note && <div className="tx-line-note">{l.note}</div>}</td>
                  {cols.map((_, i) => renderCell(l, ks[i], ts.length > 1 ? ts[i] : (i < ks.length ? ts[0] : null), i, calc, rt, disabled))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== HEADER / STATUS ============================== */

export function SubTabs({ tabs, value, onChange }) {
  return (
    <div className="tx-tabs">
      {tabs.map(([k, label]) => <button key={k} className={"tx-tab" + (value === k ? " active" : "")} onClick={() => onChange(k)}>{label}</button>)}
    </div>
  );
}

export function ReturnStatusBar({ rt, dueDefault, dueNote = "verify weekends / holidays / eFPS", onFiled, onUnfiled, children, extra }) {
  const [modal, setModal] = useState(null);
  const f = rt.form;
  const due = f.dueDate || (dueDefault ? mdy(dueDefault) : "");
  return (
    <div className="tx-statusbar">
      <div className="tx-selectors">{children}</div>
      <div className="tx-status-right">
        {due && (
          <div className="tx-due" title={dueNote}>
            <span>Due</span>
            <div style={{ width: 118 }}><Field type="date" value={due} onChange={(v) => rt.patch({ dueDate: v })} /></div>
            <span className="tx-due-note">{dueNote}</span>
          </div>
        )}
        {extra}
        <span className={"tx-chip" + (f.filed ? " filed" : "")}>{f.filed ? `Filed ${f.dateFiled || ""}${f.filedRef ? " · " + f.filedRef : ""}` : "Draft"}</span>
        {f.filed
          ? <button className="io-btn" onClick={() => { rt.patch({ filed: false }); onUnfiled && onUnfiled(); }}>Mark as draft</button>
          : <button className="io-btn" onClick={() => setModal({ date: todayMDY(), ref: "" })}><Check size={13} /> Mark as filed</button>}
      </div>
      {modal && (
        <EntryModalShell title="Mark as filed" submitLabel="Mark as filed" onCancel={() => setModal(null)}
          onSubmit={() => { rt.patch({ filed: true, dateFiled: toMDY(modal.date), filedRef: modal.ref }); onFiled && onFiled(modal); setModal(null); }}>
          <LabeledField label="Date filed"><Field type="date" value={modal.date} onChange={(v) => setModal((m) => ({ ...m, date: v }))} /></LabeledField>
          <LabeledField label="Reference (filing reference / TRN, optional)" wide><Field value={modal.ref} onChange={(v) => setModal((m) => ({ ...m, ref: v }))} /></LabeledField>
        </EntryModalShell>
      )}
    </div>
  );
}

export function YearSelect({ years, value, onChange }) {
  return <select className="io-select" value={value} onChange={(e) => onChange(Number(e.target.value))}>{years.map((y) => <option key={y} value={y}>{y}</option>)}</select>;
}

/* ============================== PAYMENT / DECLARATION ============================== */

export const PAYMENT_ROWS = [["cashDebit", "Cash/Bank Debit Memo"], ["check", "Check"], ["taxDebit", "Tax Debit Memo"], ["others", "Others (specify)"]];
export const paymentTotal = (form) => round2(PAYMENT_ROWS.reduce((s, [k]) => s + num(((form.payment || {})[k] || {}).amount), 0));

export function PaymentDetails({ rt, title = "Part III – Details of Payment", firstItem, amountDue }) {
  const p = rt.form.payment || {};
  const ror = rt.form.ror || {};
  const setRow = (k, field, v) => rt.update((cur) => ({ ...cur, payment: { ...(cur.payment || {}), [k]: { ...((cur.payment || {})[k] || {}), [field]: v } } }));
  const paid = paymentTotal(rt.form);
  const diff = round2(paid - num(amountDue));
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>{title}</span></div>
      <div className="table-scroll">
        <table className="tx-form">
          <thead><tr><th style={{ width: 52 }}>Item</th><th>Particulars</th><th>Drawee Bank / Agency</th><th>Number</th><th>Date (MM/DD/YYYY)</th><th className="num-head">Amount</th></tr></thead>
          <tbody>
            {PAYMENT_ROWS.map(([k, label], i) => {
              const row = p[k] || {};
              return (
                <tr key={k}>
                  <td className="tx-no">{firstItem ? Number(firstItem) + i : ""}</td>
                  <td className="tx-label">{k === "others" ? <Field value={row.specify} onChange={(v) => setRow(k, "specify", v)} placeholder="Others (specify)" /> : label}</td>
                  <td style={{ minWidth: 150 }}><Field type="combo" options={BANK_ACCOUNTS} freeText value={row.bank} onChange={(v) => setRow(k, "bank", v)} /></td>
                  <td><Field value={row.number} onChange={(v) => setRow(k, "number", v)} /></td>
                  <td style={{ width: 120 }}><Field type="date" value={row.date} onChange={(v) => setRow(k, "date", v)} /></td>
                  <td className="num" style={{ width: 140 }}><NumberField value={row.amount} onChange={(v) => setRow(k, "amount", v)} align="right" /></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div className="tx-inline-fields">
        <LabeledField label="Machine Validation / ROR reference No."><Field value={ror.refNo} onChange={(v) => rt.update((cur) => ({ ...cur, ror: { ...(cur.ror || {}), refNo: v } }))} /></LabeledField>
        <LabeledField label="Date"><Field type="date" value={ror.date} onChange={(v) => rt.update((cur) => ({ ...cur, ror: { ...(cur.ror || {}), date: v } }))} /></LabeledField>
      </div>
      {amountDue !== undefined && (
        <div className="tx-note">Paid {fmtPlain(paid)} vs. amount due {fmtPlain(Math.max(0, num(amountDue)))}: difference {fmtParen(round2(paid - Math.max(0, num(amountDue))))}{Math.abs(diff) < 0.005 ? " (matches)" : ""}. Information only.</div>
      )}
    </div>
  );
}

export function Declaration({ data, rt, showTreasurer, signatoryLabel = "Signature over printed name of taxpayer / authorized representative" }) {
  const c = data.company || {};
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>Declaration</span></div>
      <div className="tx-note">I/We declare under the penalties of perjury that this return, and all its attachments, have been made in good faith, verified by me/us, and to the best of my/our knowledge and belief, is true and correct, pursuant to the provisions of the National Internal Revenue Code, as amended, and the regulations issued under authority thereof.</div>
      <div className="tx-bg-grid">
        <div className="tx-bg-item"><span className="tx-no"></span><span className="tx-bg-label">{signatoryLabel}</span><span className="tx-bg-value">{c.authorizedSignatory || "— set in Company Details —"}</span></div>
        <div className="tx-bg-item"><span className="tx-no"></span><span className="tx-bg-label">Title / Position</span><span className="tx-bg-value">{c.signatoryPosition || "—"}</span></div>
        <div className="tx-bg-item"><span className="tx-no"></span><span className="tx-bg-label">TIN of signatory</span><span className="tx-bg-value">{c.signatoryTin || "—"}</span></div>
        {showTreasurer && <>
          <div className="tx-bg-item"><span className="tx-no"></span><span className="tx-bg-label">Treasurer / Assistant Treasurer</span><span className="tx-bg-value">{c.treasurerName || "— set in Company Details —"}</span></div>
          <div className="tx-bg-item"><span className="tx-no"></span><span className="tx-bg-label">Title / TIN</span><span className="tx-bg-value">{[c.treasurerTitle, c.treasurerTin].filter(Boolean).join(" · ") || "—"}</span></div>
        </>}
      </div>
      <div className="tx-inline-fields">
        <LabeledField label="Tax Agent Accreditation No. / Attorney's Roll No."><Field value={(rt.form.m || {}).agentNo ?? (c.taxAgentAccreditationNo || c.taxAgentRollNo || "")} onChange={(v) => rt.setM("agentNo", v)} /></LabeledField>
        <LabeledField label="Date of Issue"><Field type="date" value={(rt.form.m || {}).agentIssue ?? (c.taxAgentIssueDate || "")} onChange={(v) => rt.setM("agentIssue", v)} /></LabeledField>
        <LabeledField label="Date of Expiry"><Field type="date" value={(rt.form.m || {}).agentExpiry ?? (c.taxAgentExpiryDate || "")} onChange={(v) => rt.setM("agentExpiry", v)} /></LabeledField>
      </div>
    </div>
  );
}

/* ============================== RECONCILIATION / CHECKLIST ============================== */

// items: [{ label, a, b, aLabel?, bLabel?, info? }]. A difference under ₱1 counts as reconciled; `info`
// rows explain a difference (e.g. year-end adjustment) instead of flagging it.
export function ReconStrip({ title = "Reconciliation", items, causes, colA = "Form", colB = "Source" }) {
  const rows = items.filter(Boolean);
  const bad = rows.filter((r) => !r.info && r.b !== null && r.b !== undefined && Math.abs(round2(num(r.a) - num(r.b))) >= 1);
  const ok = bad.length === 0;
  return (
    <div className={"vs-check " + (ok ? "ok" : "warn")}>
      <div className="vs-check-head">{ok ? <><Check size={14} /> {title}: Reconciled</> : <><Info size={14} /> {title}: {bad.length} difference{bad.length === 1 ? "" : "s"}</>}</div>
      <div className="vs-check-grid">
        <span></span><span className="vs-check-col">{colA}</span><span className="vs-check-col">{colB}</span><span className="vs-check-col">Difference</span>
        {rows.map((r, i) => {
          const showB = r.b !== null && r.b !== undefined;
          const d = showB ? round2(num(r.a) - num(r.b)) : null;
          return (
            <React.Fragment key={i}>
              <span>{r.label}{r.aLabel || r.bLabel ? <span className="tx-recon-sub"> {r.aLabel ? `(${r.aLabel}` : "("}{r.bLabel ? ` vs ${r.bLabel})` : ")"}</span> : null}{r.info && Math.abs(d || 0) >= 0.005 ? <span className="tx-recon-sub"> {r.info}</span> : null}</span>
              <span className="num">{r.aText ?? fmtPlain(r.a)}</span>
              <span className="num">{showB ? (r.bText ?? fmtPlain(r.b)) : "—"}</span>
              <span className={"num" + (d && Math.abs(d) >= 1 && !r.info ? " tx-due" : "")}>{showB ? (r.aText !== undefined ? String(d) : fmtParen(d)) : ""}</span>
            </React.Fragment>
          );
        })}
      </div>
      {!ok && causes && <div className="vs-check-causes">Likely causes: {causes}</div>}
    </div>
  );
}

// items: [[key, label]]. Checked count can feed "Number of sheets / attachments".
export function AttachmentChecklist({ rt, items, title = "Attachment checklist", onOpen }) {
  const a = rt.form.attach || {};
  const toggle = (k) => rt.update((cur) => ({ ...cur, attach: { ...(cur.attach || {}), [k]: !(cur.attach || {})[k] } }));
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>{title}</span><span className="tx-muted">{items.filter(([k]) => a[k]).length} of {items.length} ticked</span></div>
      <div className="tx-checklist">
        {items.map(([k, label, openKey]) => (
          <div key={k} className="tx-check-item">
            <button className="tx-check-btn" onClick={() => toggle(k)}>{a[k] ? <CheckCircle2 size={16} className="ok" /> : <Circle size={16} />}<span>{label}</span></button>
            {openKey && onOpen && <button className="tx-link" onClick={() => onOpen(openKey)}>Open</button>}
          </div>
        ))}
      </div>
    </div>
  );
}
export const attachCount = (form, items) => items.filter(([k]) => (form.attach || {})[k]).length;

export function Notice({ tone = "warn", children }) {
  return <div className={"tx-notice " + tone}><Info size={14} /><div>{children}</div></div>;
}

/* ============================== EXPORT: PDF + EXCEL ============================== */

// Converts a schema section + computed values into an export block.
export function sectionBlock(section, calc, form) {
  if (section.hide) return null;
  const cols = section.cols || ["Amount"];
  const m = (form && form.m) || {};
  const rows = section.lines.filter((l) => l && !l.hide).map((l) => {
    if (l.heading) return { heading: l.heading };
    const ks = asArr(l.k), ts = asArr(l.t);
    const cells = cols.map((_, i) => {
      const k = ks[i]; const t = ts.length > 1 ? ts[i] : (i < ks.length ? ts[0] : null);
      if (!k || !t) return { v: "" };
      if (["calc", "auto", "man"].includes(t)) return { v: calc.v[k] ?? 0, k, f: Array.isArray(l.f) ? l.f[i] : (i === 0 ? l.f : undefined), num: true };
      if (t === "ro") return { v: l.text ?? "" };
      const raw = m[k];
      return { v: raw === undefined || raw === null || raw === "" ? (l.dflt ?? "") : String(raw) };
    });
    return { no: l.no || "", label: l.label, cells, strong: !!l.strong };
  });
  return { type: "form", title: section.title, cols, rows };
}
export const infoBlock = (title, items) => ({ type: "info", title, rows: items.filter(Boolean).map(([no, label, value]) => [no, label, String(value ?? "")]) });
export function paymentBlock(form, title = "Details of Payment") {
  const p = form.payment || {};
  const body = PAYMENT_ROWS.map(([k, label]) => { const r = p[k] || {}; return [k === "others" && r.specify ? `Others: ${r.specify}` : label, r.bank || "", r.number || "", r.date || "", num(r.amount)]; });
  const ror = form.ror || {};
  body.push([`Machine Validation / ROR: ${ror.refNo || ""}`, "", "", ror.date || "", ""]);
  return { type: "table", title, head: ["Particulars", "Drawee Bank / Agency", "Number", "Date", "Amount"], body, numCols: [4] };
}
export function signatureBlock(data) {
  const c = data.company || {};
  return { name: c.authorizedSignatory || "", title: c.signatoryPosition || "", tin: c.signatoryTin || "", treasurer: c.treasurerName ? `${c.treasurerName}${c.treasurerTitle ? ", " + c.treasurerTitle : ""}${c.treasurerTin ? " · TIN " + c.treasurerTin : ""}` : "" };
}

const fmtCell = (v) => (typeof v === "number" ? fmtParen(v) : String(v ?? ""));

// spec = { title, formNo, version, period, company: companyInfo(), blocks, signature, filename }
export async function exportWorksheetPdf(spec) {
  const jsPDF = await withTimeout(loadJsPDF(), 8000);
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const marginX = 12;
  const pageW = () => doc.internal.pageSize.getWidth();
  const pageH = () => doc.internal.pageSize.getHeight();
  let y = 14;
  doc.setFont("helvetica", "bold"); doc.setFontSize(12.5);
  doc.text(spec.title, marginX, y); y += 5.5;
  doc.setFont("helvetica", "normal"); doc.setFontSize(9);
  if (spec.version) { doc.text(spec.version, marginX, y); y += 4.5; }
  doc.text(`${spec.company.name || ""}   TIN ${spec.company.tinText || "—"}   RDO ${spec.company.rdo || "—"}`, marginX, y); y += 4.5;
  if (spec.period) { doc.text(spec.period, marginX, y); y += 4.5; }
  doc.setTextColor(161, 61, 61); doc.setFont("helvetica", "bold");
  doc.text(WORKSHEET_NOTE, marginX, y); y += 3;
  doc.setTextColor(0); doc.setFont("helvetica", "normal");

  const ensure = (h) => { if (y + h > pageH() - 16) { doc.addPage("a4", "portrait"); y = 14; } };
  const titleOf = (t) => { if (!t) return; ensure(12); y += 5; doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.text(t, marginX, y); doc.setFont("helvetica", "normal"); y += 1.5; };
  const head = [31, 95, 168];
  spec.blocks.filter(Boolean).forEach((b) => {
    if (b.landscape) { doc.addPage("a4", "landscape"); y = 12; }
    titleOf(b.title);
    if (b.type === "info") {
      doc.autoTable({ startY: y + 1, margin: { left: marginX, right: marginX }, body: b.rows, theme: "grid",
        styles: { fontSize: 8, cellPadding: 1.3 }, columnStyles: { 0: { cellWidth: 12, halign: "center" }, 1: { cellWidth: 70, textColor: [85, 105, 127] } } });
    } else if (b.type === "form") {
      const strong = new Set();
      const body = b.rows.map((r, i) => {
        if (r.heading) { strong.add(i); return [{ content: r.heading, colSpan: 2 + b.cols.length, styles: { fillColor: [232, 241, 251], fontStyle: "bold" } }]; }
        if (r.strong) strong.add(i);
        return [r.no, r.label, ...r.cells.map((c) => (c.num ? fmtParen(c.v) : c.v))];
      });
      const colStyles = { 0: { cellWidth: 12, halign: "center" } };
      b.cols.forEach((_, i) => { colStyles[2 + i] = { halign: "right", cellWidth: b.cols.length > 1 ? 30 : 38 }; });
      doc.autoTable({ startY: y + 1, margin: { left: marginX, right: marginX }, head: [["Item", "Description", ...b.cols]], body, theme: "grid",
        styles: { fontSize: 8, cellPadding: 1.3, overflow: "linebreak" }, headStyles: { fillColor: head, textColor: 255 }, columnStyles: colStyles,
        didParseCell: (d) => { if (d.section === "body" && strong.has(d.row.index)) d.cell.styles.fontStyle = "bold"; } });
    } else if (b.type === "table") {
      const numCols = new Set(b.numCols || []);
      const fmtRow = (r) => r.map((v, i) => (numCols.has(i) && typeof v === "number" ? fmtParen(v) : fmtCell(v)));
      const colStyles = {}; numCols.forEach((i) => { colStyles[i] = { halign: "right" }; });
      doc.autoTable({ startY: y + 1, margin: { left: marginX, right: marginX }, head: [b.head], body: b.body.map(fmtRow), foot: b.foot ? [fmtRow(b.foot)] : undefined, theme: "grid",
        styles: { fontSize: b.small ? 6.5 : 7.5, cellPadding: 1.1, overflow: "linebreak" }, headStyles: { fillColor: head, textColor: 255, fontSize: b.small ? 6.5 : 7.5 },
        footStyles: { fillColor: [232, 241, 251], textColor: [15, 42, 77], fontStyle: "bold" }, columnStyles: colStyles, showFoot: "lastPage" });
    } else if (b.type === "text") {
      ensure(10);
      const lines = doc.splitTextToSize(b.text, pageW() - marginX * 2);
      doc.setFontSize(8); doc.text(lines, marginX, y + 4); y += 4 + lines.length * 3.6;
      return;
    }
    y = doc.lastAutoTable.finalY + 2;
  });
  if (spec.signature) {
    if (doc.internal.pageSize.getWidth() > doc.internal.pageSize.getHeight()) { doc.addPage("a4", "portrait"); y = 14; }
    ensure(30); y += 10;
    doc.setFontSize(8.5);
    doc.line(marginX, y, marginX + 85, y);
    doc.text(spec.signature.name || " ", marginX, y + 4);
    doc.text(`${spec.signature.title || "Title"}${spec.signature.tin ? "   TIN " + spec.signature.tin : ""}`, marginX, y + 8);
    doc.setTextColor(85, 105, 127); doc.text("Signature over printed name of taxpayer / authorized representative", marginX, y + 12); doc.setTextColor(0);
    if (spec.signature.treasurer) { doc.line(115, y, 198, y); doc.text(spec.signature.treasurer, 115, y + 4); doc.setTextColor(85, 105, 127); doc.text("Treasurer / Assistant Treasurer", 115, y + 8); doc.setTextColor(0); }
  }
  const pages = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pages; i++) {
    doc.setPage(i);
    doc.setFontSize(7); doc.setTextColor(120);
    doc.text(`${WORKSHEET_NOTE}  ·  Printed ${todayMDY()}  ·  Page ${i} of ${pages}`, marginX, doc.internal.pageSize.getHeight() - 7);
    doc.setTextColor(0);
  }
  doc.save(spec.filename || `${spec.formNo || "BIR"}-worksheet-${todayFileStamp()}.pdf`);
}

// One workbook: the return (header, form blocks with live formulas for totals) on the first sheet;
// table blocks with a `sheet` name go to their own sheets (alphalists, annexes).
export function exportWorksheetExcel(spec) {
  const wb = XLSX.utils.book_new();
  const aoa = [[spec.title], [spec.version || ""], [`${spec.company.name || ""}  TIN ${spec.company.tinText || ""}  RDO ${spec.company.rdo || ""}`], [spec.period || ""], [WORKSHEET_NOTE], []];
  const addr = {};
  const formulaCells = []; // [rowIdx, colIdx, value, fn]
  const numCells = [];
  const extraSheets = [];
  spec.blocks.filter(Boolean).forEach((b) => {
    if (b.type === "table" && b.sheet) { extraSheets.push(b); return; }
    if (b.title) aoa.push([b.title]);
    if (b.type === "info") { b.rows.forEach((r) => aoa.push(r)); }
    else if (b.type === "form") {
      aoa.push(["Item", "Description", ...b.cols]);
      b.rows.forEach((r) => {
        if (r.heading) { aoa.push(["", r.heading]); return; }
        const rowIdx = aoa.length;
        aoa.push([r.no, r.label, ...r.cells.map((c) => c.v)]);
        r.cells.forEach((c, i) => {
          if (!c.num) return;
          const ref = XLSX.utils.encode_cell({ r: rowIdx, c: 2 + i });
          if (c.k) addr[c.k] = ref;
          numCells.push(ref);
          if (c.f) formulaCells.push([ref, c.v, c.f]);
        });
      });
    } else if (b.type === "table") {
      aoa.push(b.head);
      [...b.body, ...(b.foot ? [b.foot] : [])].forEach((r) => {
        const rowIdx = aoa.length; aoa.push(r);
        (b.numCols || []).forEach((i) => { if (typeof r[i] === "number") numCells.push(XLSX.utils.encode_cell({ r: rowIdx, c: i })); });
      });
    } else if (b.type === "text") aoa.push([b.text]);
    aoa.push([]);
  });
  if (spec.signature) { aoa.push(["Signatory", spec.signature.name, spec.signature.title, spec.signature.tin]); if (spec.signature.treasurer) aoa.push(["Treasurer", spec.signature.treasurer]); }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  numCells.forEach((ref) => { if (ws[ref] && typeof ws[ref].v === "number") ws[ref].z = "#,##0.00;(#,##0.00)"; });
  formulaCells.forEach(([ref, v, fn]) => {
    const r = (k) => addr[k] || "0";
    try { const f = fn(r); if (f) ws[ref] = { t: "n", v: num(v), f, z: "#,##0.00;(#,##0.00)" }; } catch (e) { /* keep the value */ }
  });
  ws["!cols"] = [{ wch: 8 }, { wch: 62 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, ws, (spec.sheetName || spec.formNo || "Return").slice(0, 31));
  extraSheets.forEach((b) => {
    const sAoa = [[b.title || b.sheet], [], b.head, ...b.body, ...(b.foot ? [b.foot] : [])];
    const s = XLSX.utils.aoa_to_sheet(sAoa);
    (b.numCols || []).forEach((ci) => { for (let ri = 3; ri < sAoa.length; ri++) { const ref = XLSX.utils.encode_cell({ r: ri, c: ci }); if (s[ref] && typeof s[ref].v === "number") s[ref].z = "#,##0.00;(#,##0.00)"; } });
    s["!cols"] = b.head.map((h) => ({ wch: Math.max(10, Math.min(30, String(h).length + 4)) }));
    let name = String(b.sheet).replace(/[\\/?*[\]:]/g, " ").slice(0, 31);
    let n = 2; while (wb.SheetNames.includes(name)) name = `${String(b.sheet).slice(0, 27)} (${n++})`;
    XLSX.utils.book_append_sheet(wb, s, name);
  });
  XLSX.writeFile(wb, (spec.filename || `${spec.formNo || "BIR"}-worksheet-${todayFileStamp()}.pdf`).replace(/\.pdf$/, ".xlsx"));
}

export function useWorksheetExport(buildSpec) {
  return useReportExport(() => exportWorksheetPdf(buildSpec()), () => exportWorksheetExcel(buildSpec()));
}
export function WorksheetButtons({ hook, label = "Encoding worksheet" }) {
  return (
    <div className="io-toolbar">
      <button className="io-btn" onClick={hook.exportPDF}><FileDown size={13} /> {label} PDF (A4)</button>
      <button className="io-btn" onClick={hook.exportExcel}><FileDown size={13} /> Excel</button>
    </div>
  );
}
export { ExportStatus };

/* ============================== STYLES ============================== */

export function TaxStyle() {
  return (
    <style>{`
      .tx-tabs { display: flex; gap: 6px; margin: 0 0 14px; flex-wrap: wrap; }
      .tx-tab { border: 1px solid var(--line); background: var(--white); color: var(--ink-soft); padding: 7px 14px; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; }
      .tx-tab.active { background: var(--green); border-color: var(--green); color: #fff; }
      .tx-statusbar { display: flex; justify-content: space-between; align-items: center; gap: 12px; flex-wrap: wrap; background: var(--white); border: 1px solid var(--line); border-radius: 10px; padding: 10px 12px; margin-bottom: 14px; }
      .tx-selectors { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
      .tx-status-right { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
      .tx-due { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-soft); }
      .tx-due-note { font-size: 10.5px; color: #8a6d1f; }
      .tx-chip { font-size: 11.5px; font-weight: 600; padding: 4px 10px; border-radius: 999px; background: #FFF4D6; color: #7A5A00; border: 1px solid #F0D98A; }
      .tx-chip.filed { background: #E3F4E8; color: #1E6B3A; border-color: #A9D8B7; }
      .tx-card { background: var(--white); border: 1px solid var(--line); border-radius: 10px; margin-bottom: 14px; overflow: hidden; }
      .tx-card-head { display: flex; justify-content: space-between; align-items: center; gap: 10px; padding: 9px 12px; background: var(--paper-deep); color: var(--green-deep); font-weight: 600; font-size: 13px; border-bottom: 1px solid var(--line); }
      .tx-note { padding: 8px 12px; font-size: 12px; color: var(--ink-soft); }
      .tx-muted { font-size: 11.5px; color: var(--ink-soft); font-weight: 500; }
      .tx-form { width: 100%; border-collapse: collapse; font-size: 12.5px; }
      .tx-form th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.02em; color: var(--green-deep); padding: 7px 8px; border-bottom: 1px solid var(--line); background: #F4F8FD; white-space: nowrap; }
      .tx-form th.tx-col { min-width: 150px; }
      .tx-form td { padding: 4px 8px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
      .tx-form td.num { text-align: right; width: 170px; font-family: 'IBM Plex Mono', monospace; }
      .tx-form td.tx-blank { background: #F8FAFD; }
      .tx-form tr.tx-strong td { font-weight: 700; background: #F8FBFF; }
      .tx-form tr.tx-subhead td { font-weight: 600; font-size: 11.5px; color: var(--green-deep); background: #F4F8FD; text-transform: uppercase; letter-spacing: 0.02em; }
      .tx-no { color: var(--ink-soft); font-size: 11.5px; text-align: center; width: 52px; white-space: nowrap; }
      .tx-label { color: var(--ink); }
      .tx-line-note { font-size: 11px; color: var(--ink-soft); margin-top: 2px; }
      .tx-calc { color: var(--green-deep); }
      .tx-due { }
      td.tx-due, .num.tx-due, span.tx-due { color: var(--red) !important; }
      td.tx-over, .num.tx-over { color: #1E6B3A !important; }
      .tx-auto { display: flex; align-items: center; justify-content: flex-end; gap: 4px; }
      .tx-auto .ledger-field { max-width: 140px; }
      .tx-auto.ovr .ledger-field { border-color: #E3B341; background: #FFFBEF; }
      .tx-dot { width: 7px; height: 7px; border-radius: 50%; background: #E3B341; flex: none; }
      .tx-icon-btn { border: none; background: transparent; color: var(--ink-soft); cursor: pointer; padding: 2px; border-radius: 4px; display: inline-flex; }
      .tx-icon-btn:hover { background: var(--paper-deep); color: var(--green); }
      .tx-amt { padding: 4px 2px; }
      .tx-ro { color: var(--ink-soft); }
      .tx-link { border: none; background: transparent; color: var(--green); font-weight: 600; cursor: pointer; padding: 0; font-size: 12px; text-decoration: underline; }
      .tx-bg-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 0; }
      .tx-bg-item { display: grid; grid-template-columns: 40px 1fr; gap: 2px 6px; padding: 7px 10px; border-bottom: 1px solid var(--line-soft); border-right: 1px solid var(--line-soft); }
      .tx-bg-item .tx-no { grid-row: span 2; }
      .tx-bg-label { font-size: 11px; color: var(--ink-soft); }
      .tx-bg-value { font-size: 13px; font-weight: 500; word-break: break-word; }
      .tx-bg-value.empty { color: #B07D10; font-weight: 400; font-style: italic; }
      .tx-inline-fields { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 10px; padding: 10px 12px; }
      .tx-notice { display: flex; gap: 8px; align-items: flex-start; padding: 9px 12px; border-radius: 8px; font-size: 12.5px; margin-bottom: 12px; }
      .tx-notice.warn { background: #FFF7E0; border: 1px solid #F0D98A; color: #6B4E00; }
      .tx-notice.danger { background: #FDECEC; border: 1px solid #E8B4B4; color: #8A2626; }
      .tx-notice.info { background: #EEF5FD; border: 1px solid var(--line); color: var(--green-deep); }
      .tx-notice.ok { background: #E9F6EE; border: 1px solid #A9D8B7; color: #1E6B3A; }
      .tx-checklist { padding: 6px 12px 10px; display: grid; gap: 4px; }
      .tx-check-item { display: flex; justify-content: space-between; align-items: center; gap: 10px; }
      .tx-check-btn { display: flex; align-items: center; gap: 8px; border: none; background: transparent; cursor: pointer; font-size: 12.5px; color: var(--ink); text-align: left; padding: 4px 0; }
      .tx-check-btn .ok { color: #1E6B3A; }
      .tx-recon-sub { color: var(--ink-soft); font-size: 11px; }
      .tx-grid2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap: 14px; }
      .tx-amber { color: #9A6B00; font-size: 11.5px; }
      .tx-pill { display: inline-block; font-size: 10.5px; padding: 1px 6px; border-radius: 999px; background: var(--paper-deep); color: var(--green-deep); margin-left: 4px; }
      .tx-pill.warn { background: #FFF4D6; color: #7A5A00; }
      .tx-table td, .tx-table th { white-space: nowrap; }
      .tx-table td.wrap { white-space: normal; }
      @media (max-width: 720px) { .tx-grid2 { grid-template-columns: 1fr; } .tx-form td.num { width: 130px; } }
    `}</style>
  );
}

// Re-exported for convenience in the form modules.
export { fmtPlain, fmtParen, num, round2, MONTHS };

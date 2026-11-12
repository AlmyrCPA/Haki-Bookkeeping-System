// Employees Record (Setup) — one master list of employees that the 1601-C, the Alphalist of Employees
// (2316) and the 1604-C alphalist all read from, so a name / TIN / address correction is made once.
import React, { useMemo, useState } from "react";
import { UserRound, Trash2, Pencil, Link2 } from "lucide-react";
import {
  Field, LabeledField, EntryModalShell, SectionHeader, ReadCell, ConfirmModal, AddRowBtn, DelBtn, SearchBar, matchesSearch,
  useRowSelection, SelectAllCheckbox, RowCheckbox, SelectionBar, usePagination, PageSizeSelect, PaginationBar,
  ImportExportBar, ImportStatus, useJournalImport, ExportBar, ExportStatus, useJournalExport, EMP_STATUS_OPTIONS,
  SEP_REASON_OPTIONS, normalizeTin, parseAppDate, toMDY, uid, num, round2, pad2,
} from "../BookkeepingApp";
import { TaxStyle } from "./shared";

export const EMPLOYEE_RECORD_DEFAULTS = {
  empNo: "", tin: "", lastName: "", firstName: "", middleName: "", suffix: "", dob: "", nationality: "Filipino",
  phone: "", email: "", addr1: "", addr2: "", zipCode: "", foreignAddress: "", sssNo: "", philhealthNo: "", pagibigNo: "",
  dateHired: "", dateSeparated: "", empStatus: "R", sepReason: "NA", position: "", department: "", employerType: "Main",
  active: "Yes", region: "", isMWE: "N", smwPerDay: 0, smwPerMonth: 0, smwPerYear: 0, factorDaysPerYear: 261,
  defaultMonthlyBasic: 0, substitutedFiling: "Y", prevEmployerName: "", prevEmployerTin: "", prevEmployerAddress: "", prevEmployerZip: "",
};

export const employeeDisplayName = (e) => {
  if (!e) return "";
  const first = [e.firstName, e.middleName].filter(Boolean).join(" ");
  return [e.lastName, [first, e.suffix].filter(Boolean).join(" ")].filter(Boolean).join(", ") || "(no name)";
};
export const tinDigits = (t) => String(t ?? "").replace(/\D/g, "");
export const formatEmployeeTin = (t) => {
  const d = tinDigits(t);
  if (!d) return "";
  const core = d.slice(0, 9).padStart(9, "0");
  const branch = d.slice(9);
  return `${core.slice(0, 3)}-${core.slice(3, 6)}-${core.slice(6, 9)}${branch ? "-" + branch : ""}`;
};
export const isMweRecord = (e) => String(e && e.isMWE).toUpperCase() === "Y";

// Was the employee on the payroll at any time in [start, end]?
export function employedDuring(e, start, end) {
  const hired = parseAppDate(e.dateHired);
  const sep = parseAppDate(e.dateSeparated);
  if (hired && hired > end) return false;
  if (sep && sep < start) return false;
  return true;
}

// The alphalist row's identity fields, read from the record. Period From/To are clipped to the row's year.
export function alphalistFieldsFromRecord(rec, row) {
  const year = Number(row.year) || new Date().getFullYear();
  const hired = parseAppDate(rec.dateHired);
  const sep = parseAppDate(rec.dateSeparated);
  const jan1 = new Date(year, 0, 1), dec31 = new Date(year, 11, 31);
  const from = hired && hired > jan1 ? hired : jan1;
  const to = sep && sep < dec31 ? sep : dec31;
  const mdy = (d) => `${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}/${d.getFullYear()}`;
  return {
    ...row,
    employeeId: rec.id,
    tin: rec.tin, lastName: rec.lastName, firstName: rec.firstName, middleName: rec.middleName,
    nationality: rec.nationality && rec.nationality.toLowerCase() !== "filipino" ? rec.nationality : (row.nationality && !rec.nationality ? row.nationality : ""),
    address: [rec.addr1, rec.addr2].filter(Boolean).join(", ") || row.address || "",
    zipCode: rec.zipCode || row.zipCode || "",
    isMWE: isMweRecord(rec) ? "Y" : "N",
    empStatus: rec.empStatus || row.empStatus || "R",
    empFrom: mdy(from), empTo: mdy(to),
    sepReason: sep && sep <= dec31 ? (rec.sepReason && rec.sepReason !== "NA" ? rec.sepReason : "R") : "NA",
    substitutedFiling: rec.substitutedFiling === "N" ? "N" : "Y",
    smwPerDay: rec.smwPerDay ?? row.smwPerDay, smwPerMonth: rec.smwPerMonth ?? row.smwPerMonth,
    smwPerYear: rec.smwPerYear ?? row.smwPerYear, factorDaysPerYear: rec.factorDaysPerYear || row.factorDaysPerYear || 261,
    prevEmployerName: rec.prevEmployerName || "", prevEmployerTin: rec.prevEmployerTin || "",
  };
}

const yn = (v, dflt) => {
  const s = String(v ?? "").trim().toUpperCase();
  if (["Y", "YES", "TRUE", "1"].includes(s)) return "Y";
  if (["N", "NO", "FALSE", "0"].includes(s)) return "N";
  return dflt;
};

export function validateEmployeeRecord(v, all) {
  const errs = [];
  const d = tinDigits(v.tin);
  if (!d) errs.push("TIN is required.");
  else if (d.length !== 9 && d.length !== 12) errs.push("TIN must have 9 digits, or 12 digits with the branch code.");
  else if ((all || []).some((e) => e.id !== v.id && normalizeTin(e.tin) === normalizeTin(v.tin))) errs.push("Another employee already has this TIN.");
  if (!String(v.lastName || "").trim() || !String(v.firstName || "").trim()) errs.push("Last name and first name are required.");
  if (isMweRecord(v) && !(num(v.smwPerDay) > 0 || num(v.smwPerMonth) > 0 || num(v.smwPerYear) > 0)) errs.push("Minimum wage earners need the statutory minimum wage rate (per day, month or year).");
  const h = parseAppDate(v.dateHired), s = parseAppDate(v.dateSeparated);
  if (h && s && s < h) errs.push("Date Separated can't be before Date Hired.");
  return errs;
}

export function normalizeEmployeeRecord(v) {
  const out = { ...EMPLOYEE_RECORD_DEFAULTS, ...v };
  ["dob", "dateHired", "dateSeparated"].forEach((k) => { out[k] = out[k] ? toMDY(String(out[k]).trim()) : ""; });
  out.isMWE = yn(out.isMWE, "N");
  out.substitutedFiling = yn(out.substitutedFiling, "Y");
  if (out.dateSeparated) out.active = "No";
  else if (!["Yes", "No"].includes(out.active)) out.active = "Yes";
  ["smwPerDay", "smwPerMonth", "smwPerYear", "factorDaysPerYear", "defaultMonthlyBasic"].forEach((k) => { out[k] = round2(num(out[k])); });
  if (!out.factorDaysPerYear) out.factorDaysPerYear = 261;
  if (!out.nationality) out.nationality = "Filipino";
  if (!["Main", "Secondary"].includes(out.employerType)) out.employerType = "Main";
  return out;
}

function nextEmpNo(records) {
  const nums = (records || []).map((e) => parseInt(String(e.empNo || "").replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
  const n = nums.length ? Math.max(...nums) + 1 : 1;
  return String(n).padStart(4, "0");
}

// Excel import: normalize values, assign employee numbers, and skip TINs already on file (or repeated in the file).
export function normalizeImportedEmployeeRecords(rawRows, data) {
  const existing = [...(data.employeeRecords || [])];
  const out = [];
  rawRows.forEach((r) => {
    const rec = normalizeEmployeeRecord({ ...r, id: r.id || uid() });
    const core = normalizeTin(rec.tin);
    if (!core || [...existing, ...out].some((e) => normalizeTin(e.tin) === core)) return;
    if (!rec.empNo) rec.empNo = nextEmpNo([...existing, ...out]);
    out.push(rec);
  });
  return out;
}

// Where an Employees Record entry is used: saved 1601-C lines and Alphalist rows.
export function employeeUsage(data, id) {
  let returns = 0, alpha = 0;
  Object.values(data.form1601c || {}).forEach((f) => { if ((f.lines || []).some((l) => l.employeeId === id)) returns++; });
  (data.employees || []).forEach((r) => { if (r.employeeId === id) alpha++; });
  return { returns, alpha, total: returns + alpha };
}

function EmployeeRecordModal({ data, initial, onCancel, onSubmit }) {
  const all = data.employeeRecords || [];
  const [v, setV] = useState(() => initial ? { ...EMPLOYEE_RECORD_DEFAULTS, ...initial } : { ...EMPLOYEE_RECORD_DEFAULTS, empNo: nextEmpNo(all) });
  const [errs, setErrs] = useState([]);
  const set = (k, x) => setV((p) => {
    const n = { ...p, [k]: x };
    if (k === "dateSeparated" && x) n.active = "No";
    return n;
  });
  const foreign = v.nationality && v.nationality.trim().toLowerCase() !== "filipino";
  const submit = () => {
    const rec = normalizeEmployeeRecord({ ...v, id: v.id || uid() });
    const e = validateEmployeeRecord(rec, all);
    if (e.length) { setErrs(e); return; }
    onSubmit(rec);
  };
  return (
    <EntryModalShell title={initial ? "Edit Employee Record" : "Add Employee Record"} submitLabel={initial ? "Save changes" : "Add employee"} onCancel={onCancel} onSubmit={submit}>
      {errs.length > 0 && <div className="tx-notice danger"><div>{errs.map((e) => <div key={e}>{e}</div>)}</div></div>}
      <div className="sub-block-title" style={{ marginTop: 0 }}>Identity</div>
      <div className="form-grid">
        <LabeledField label="Employee No."><Field value={v.empNo} onChange={(x) => set("empNo", x)} /></LabeledField>
        <LabeledField label="TIN (000-000-000-0000)"><Field value={v.tin} onChange={(x) => set("tin", x)} placeholder="000-000-000-0000" /></LabeledField>
        <LabeledField label="Last Name"><Field value={v.lastName} onChange={(x) => set("lastName", x)} /></LabeledField>
        <LabeledField label="First Name"><Field value={v.firstName} onChange={(x) => set("firstName", x)} /></LabeledField>
        <LabeledField label="Middle Name"><Field value={v.middleName} onChange={(x) => set("middleName", x)} /></LabeledField>
        <LabeledField label="Suffix"><Field value={v.suffix} onChange={(x) => set("suffix", x)} placeholder="Jr., III" /></LabeledField>
        <LabeledField label="Date of Birth"><Field type="date" value={v.dob} onChange={(x) => set("dob", x)} /></LabeledField>
        <LabeledField label="Nationality"><Field value={v.nationality} onChange={(x) => set("nationality", x)} /></LabeledField>
      </div>
      <div className="sub-block-title">Contact and Address</div>
      <div className="form-grid">
        <LabeledField label="Telephone / Mobile"><Field value={v.phone} onChange={(x) => set("phone", x)} /></LabeledField>
        <LabeledField label="Email"><Field value={v.email} onChange={(x) => set("email", x)} /></LabeledField>
        <LabeledField label="Local Home Address, line 1" wide><Field value={v.addr1} onChange={(x) => set("addr1", x)} /></LabeledField>
        <LabeledField label="Local Home Address, line 2" wide><Field value={v.addr2} onChange={(x) => set("addr2", x)} /></LabeledField>
        <LabeledField label="ZIP Code"><Field value={v.zipCode} onChange={(x) => set("zipCode", x)} /></LabeledField>
        {foreign && <LabeledField label="Foreign Address" wide><Field value={v.foreignAddress} onChange={(x) => set("foreignAddress", x)} /></LabeledField>}
      </div>
      <div className="sub-block-title">Government IDs (optional)</div>
      <div className="form-grid">
        <LabeledField label="SSS No."><Field value={v.sssNo} onChange={(x) => set("sssNo", x)} /></LabeledField>
        <LabeledField label="PhilHealth No."><Field value={v.philhealthNo} onChange={(x) => set("philhealthNo", x)} /></LabeledField>
        <LabeledField label="Pag-IBIG No."><Field value={v.pagibigNo} onChange={(x) => set("pagibigNo", x)} /></LabeledField>
      </div>
      <div className="sub-block-title">Employment</div>
      <div className="form-grid">
        <LabeledField label="Date Hired"><Field type="date" value={v.dateHired} onChange={(x) => set("dateHired", x)} /></LabeledField>
        <LabeledField label="Date Separated"><Field type="date" value={v.dateSeparated} onChange={(x) => set("dateSeparated", x)} /></LabeledField>
        <LabeledField label="Employment Status"><Field type="select" options={EMP_STATUS_OPTIONS} value={v.empStatus} onChange={(x) => set("empStatus", x)} /></LabeledField>
        <LabeledField label="Reason for Separation"><Field type="select" options={SEP_REASON_OPTIONS} value={v.sepReason} onChange={(x) => set("sepReason", x)} /></LabeledField>
        <LabeledField label="Position"><Field value={v.position} onChange={(x) => set("position", x)} /></LabeledField>
        <LabeledField label="Department"><Field value={v.department} onChange={(x) => set("department", x)} /></LabeledField>
        <LabeledField label="Employer Type"><Field type="select" options={["Main", "Secondary"]} value={v.employerType} onChange={(x) => set("employerType", x)} /></LabeledField>
        <LabeledField label="Active"><Field type="select" options={["Yes", "No"]} value={v.active} onChange={(x) => set("active", x)} /></LabeledField>
      </div>
      <div className="sub-block-title">Tax Setup</div>
      <div className="form-grid">
        <LabeledField label="Statutory Minimum Wage Earner"><Field type="select" options={[{ value: "N", label: "No" }, { value: "Y", label: "Yes" }]} value={v.isMWE} onChange={(x) => set("isMWE", x)} /></LabeledField>
        <LabeledField label="Substituted Filing"><Field type="select" options={[{ value: "Y", label: "Yes" }, { value: "N", label: "No (other income / multiple employers)" }]} value={v.substitutedFiling} onChange={(x) => set("substitutedFiling", x)} /></LabeledField>
        <LabeledField label="Default Monthly Basic Pay"><Field type="number" align="right" value={v.defaultMonthlyBasic} onChange={(x) => set("defaultMonthlyBasic", x)} /></LabeledField>
        {isMweRecord(v) && (<>
          <LabeledField label="Region Assigned"><Field value={v.region} onChange={(x) => set("region", x)} placeholder="e.g. NCR, Region VII" /></LabeledField>
          <LabeledField label="SMW per Day"><Field type="number" align="right" value={v.smwPerDay} onChange={(x) => set("smwPerDay", x)} /></LabeledField>
          <LabeledField label="SMW per Month"><Field type="number" align="right" value={v.smwPerMonth} onChange={(x) => set("smwPerMonth", x)} /></LabeledField>
          <LabeledField label="SMW per Year"><Field type="number" align="right" value={v.smwPerYear} onChange={(x) => set("smwPerYear", x)} /></LabeledField>
          <LabeledField label="Factor (days per year)"><Field type="number" align="right" value={v.factorDaysPerYear} onChange={(x) => set("factorDaysPerYear", x)} /></LabeledField>
        </>)}
      </div>
      <div className="sub-block-title">Previous Employer (for 2316, if hired mid-year)</div>
      <div className="form-grid">
        <LabeledField label="Name"><Field value={v.prevEmployerName} onChange={(x) => set("prevEmployerName", x)} /></LabeledField>
        <LabeledField label="TIN"><Field value={v.prevEmployerTin} onChange={(x) => set("prevEmployerTin", x)} /></LabeledField>
        <LabeledField label="Registered Address" wide><Field value={v.prevEmployerAddress} onChange={(x) => set("prevEmployerAddress", x)} /></LabeledField>
        <LabeledField label="ZIP"><Field value={v.prevEmployerZip} onChange={(x) => set("prevEmployerZip", x)} /></LabeledField>
      </div>
    </EntryModalShell>
  );
}

// One-time migration: create Employees Record entries from existing alphalist rows (matched by TIN) and link the rows.
function planAlphalistLink(data) {
  const records = data.employeeRecords || [];
  const byTin = new Map(records.map((e) => [normalizeTin(e.tin), e]));
  const creates = new Map();
  const links = [];
  (data.employees || []).forEach((r) => {
    if (r.employeeId) return;
    const core = normalizeTin(r.tin);
    if (!core) { links.push({ row: r, action: "skip", reason: "No TIN" }); return; }
    const hit = byTin.get(core);
    if (hit) { links.push({ row: r, action: "link", rec: hit }); return; }
    if (!creates.has(core)) {
      creates.set(core, normalizeEmployeeRecord({
        id: uid(), tin: r.tin, lastName: r.lastName, firstName: r.firstName, middleName: r.middleName,
        nationality: r.nationality || "Filipino", addr1: r.address, zipCode: r.zipCode, isMWE: r.isMWE,
        empStatus: r.empStatus || "R", sepReason: r.sepReason || "NA", dateHired: r.empFrom, dateSeparated: r.sepReason && r.sepReason !== "NA" ? r.empTo : "",
        substitutedFiling: r.substitutedFiling || "Y", smwPerDay: r.smwPerDay, smwPerMonth: r.smwPerMonth, smwPerYear: r.smwPerYear, factorDaysPerYear: r.factorDaysPerYear,
      }));
    }
    links.push({ row: r, action: "create", rec: creates.get(core) });
  });
  let n = parseInt(nextEmpNo(records), 10);
  [...creates.values()].forEach((rec) => { rec.empNo = String(n++).padStart(4, "0"); });
  return { creates: [...creates.values()], links };
}

export function EmployeesRecordPage({ data, setData }) {
  const records = data.employeeRecords || [];
  const [search, setSearch] = useState("");
  const [showInactive, setShowInactive] = useState(true);
  const [modal, setModal] = useState(null);
  const [confirm, setConfirm] = useState(null);
  const [blocked, setBlocked] = useState(null);
  const [linkPlan, setLinkPlan] = useState(null);
  const sel = useRowSelection();
  const importHook = useJournalImport("employeerecords", data, ({ newRows }) => setData((d) => ({ ...d, employeeRecords: [...(d.employeeRecords || []), ...newRows] })));

  const filtered = useMemo(() => records
    .filter((e) => showInactive || e.active !== "No")
    .filter((e) => matchesSearch(e, ["empNo", "tin", "lastName", "firstName", "middleName", "position", "department"], search))
    .sort((a, b) => employeeDisplayName(a).localeCompare(employeeDisplayName(b))), [records, search, showInactive]);
  const pg = usePagination(filtered, 50);
  const exportHook = useJournalExport("employeerecords", data, filtered, {});

  const save = (rec) => {
    setData((d) => {
      const list = d.employeeRecords || [];
      return { ...d, employeeRecords: list.some((e) => e.id === rec.id) ? list.map((e) => e.id === rec.id ? rec : e) : [...list, rec] };
    });
    setModal(null);
  };
  const tryDelete = (ids) => {
    const used = ids.map((id) => ({ id, u: employeeUsage(data, id) })).filter((x) => x.u.total > 0);
    if (used.length) {
      setBlocked(used.map((x) => { const e = records.find((r) => r.id === x.id); return `${employeeDisplayName(e)}: used in ${x.u.returns} 1601-C return${x.u.returns === 1 ? "" : "s"} / ${x.u.alpha} alphalist row${x.u.alpha === 1 ? "" : "s"}`; }));
      return;
    }
    setConfirm({ ids });
  };
  const doDelete = () => {
    const ids = new Set(confirm.ids);
    setData((d) => ({ ...d, employeeRecords: (d.employeeRecords || []).filter((e) => !ids.has(e.id)) }));
    sel.clear(); setConfirm(null);
  };
  const markInactive = () => {
    const ids = sel.selected;
    setData((d) => ({ ...d, employeeRecords: (d.employeeRecords || []).map((e) => ids.has(e.id) ? { ...e, active: "No" } : e) }));
    sel.clear();
  };
  const unlinkedAlpha = (data.employees || []).filter((r) => !r.employeeId).length;
  const applyLinkPlan = () => {
    const { creates, links } = linkPlan;
    const linkMap = new Map(links.filter((l) => l.action !== "skip").map((l) => [l.row.id, l.rec.id]));
    setData((d) => ({
      ...d,
      employeeRecords: [...(d.employeeRecords || []), ...creates],
      employees: (d.employees || []).map((r) => linkMap.has(r.id) ? { ...r, employeeId: linkMap.get(r.id) } : r),
    }));
    setLinkPlan(null);
  };

  return (
    <div>
      <TaxStyle />
      <SectionHeader icon={UserRound} title="Employees Record" subtitle="Master list of employees. The 1601-C, the Alphalist of Employees (2316) and the 1604-C alphalist read names, TINs, MWE status, employment dates and previous-employer details from here."
        right={<div className="header-actions"><ImportExportBar journalKey="employeerecords" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} /></div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      {unlinkedAlpha > 0 && (
        <div className="tx-notice info">
          <div>{unlinkedAlpha} Alphalist of Employees row{unlinkedAlpha === 1 ? " is" : "s are"} not linked to an Employees Record entry.{" "}
            <button className="tx-link" onClick={() => setLinkPlan(planAlphalistLink(data))}>Create Employees Record entries from existing alphalist rows</button></div>
        </div>
      )}
      {blocked && (
        <div className="tx-notice danger">
          <div><strong>Can't delete — these employees are in use.</strong> Mark them inactive instead.{blocked.map((b) => <div key={b}>{b}</div>)}
            <button className="tx-link" onClick={() => setBlocked(null)}>Dismiss</button></div>
        </div>
      )}
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn" onClick={markInactive}>Mark inactive</button>
          <button className="io-btn danger" onClick={() => tryDelete([...sel.selected])}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <SearchBar value={search} onChange={setSearch} placeholder="Search name, TIN, employee no., position…" />
            <label className="tx-muted" style={{ display: "flex", gap: 6, alignItems: "center" }}><input type="checkbox" checked={showInactive} onChange={(e) => setShowInactive(e.target.checked)} /> Show inactive</label>
            <div style={{ width: 120 }}><PageSizeSelect pg={pg} /></div>
          </div>
          <AddRowBtn onClick={() => setModal({ mode: "add" })}>Add employee</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr>
              <th style={{ width: 32 }}><SelectAllCheckbox ids={pg.pageRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
              <th>Employee No.</th><th style={{ minWidth: 220 }}>Name</th><th>TIN</th><th>Status</th><th>MWE</th><th>Date Hired</th><th>Date Separated</th><th>Active</th><th style={{ width: 70 }}></th><th style={{ width: 36 }}></th>
            </tr></thead>
            <tbody>
              {pg.pageRows.length === 0 && <tr><td colSpan={11} className="empty-row">{records.length ? "No employees match this search." : "No employees yet — add your first one above, or import an Excel file."}</td></tr>}
              {pg.pageRows.map((e) => (
                <tr key={e.id} className={sel.selected.has(e.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={e.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><ReadCell>{e.empNo}</ReadCell></td>
                  <td><ReadCell>{employeeDisplayName(e)}{e.nationality && e.nationality.toLowerCase() !== "filipino" ? <span className="tx-pill">{e.nationality}</span> : null}</ReadCell></td>
                  <td><ReadCell>{formatEmployeeTin(e.tin)}</ReadCell></td>
                  <td><ReadCell align="center">{e.empStatus}</ReadCell></td>
                  <td><ReadCell align="center">{isMweRecord(e) ? "Yes" : "No"}</ReadCell></td>
                  <td><ReadCell>{e.dateHired}</ReadCell></td>
                  <td><ReadCell>{e.dateSeparated}</ReadCell></td>
                  <td><ReadCell align="center">{e.active || "Yes"}</ReadCell></td>
                  <td className="text-center"><button className="io-btn" style={{ padding: "4px 8px" }} onClick={() => setModal({ mode: "edit", row: e })}><Pencil size={12} /> Edit</button></td>
                  <td className="text-center"><DelBtn onClick={() => tryDelete([e.id])} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <PaginationBar pg={pg} />
      </div>
      <div className="rc-warn">TINs, birth dates and addresses are personal data. Access follows the client allowlist in firestore.rules. Deleting an employee used in a 1601-C return or an alphalist row is blocked; mark the employee inactive instead.</div>
      {modal && <EmployeeRecordModal data={data} initial={modal.mode === "edit" ? modal.row : null} onCancel={() => setModal(null)} onSubmit={save} />}
      {confirm && (
        <ConfirmModal title="Delete employees" danger confirmLabel="Delete"
          message={`Delete ${confirm.ids.length} employee record${confirm.ids.length === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirm(null)} onConfirm={doDelete} />
      )}
      {linkPlan && (
        <EntryModalShell title="Link alphalist rows to Employees Record" submitLabel={`Create ${linkPlan.creates.length} and link`} onCancel={() => setLinkPlan(null)} onSubmit={applyLinkPlan}>
          <p className="section-head-note">Rows are matched by TIN. {linkPlan.creates.length} new Employees Record entr{linkPlan.creates.length === 1 ? "y" : "ies"} will be created; {linkPlan.links.filter((l) => l.action === "link").length} row(s) link to an existing entry; {linkPlan.links.filter((l) => l.action === "skip").length} row(s) without a TIN are left as they are.</p>
          <div className="table-scroll" style={{ maxHeight: 340 }}>
            <table className="ledger-table">
              <thead><tr><th>Year</th><th>Alphalist row</th><th>TIN</th><th>Action</th></tr></thead>
              <tbody>
                {linkPlan.links.map((l) => (
                  <tr key={l.row.id}>
                    <td>{l.row.year}</td><td>{[l.row.lastName, l.row.firstName].filter(Boolean).join(", ")}</td><td>{l.row.tin}</td>
                    <td>{l.action === "create" ? <><Link2 size={12} /> Create #{l.rec.empNo}</> : l.action === "link" ? <><Link2 size={12} /> Link to {employeeDisplayName(l.rec)}</> : <span className="tx-amber">{l.reason}</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </EntryModalShell>
      )}
    </div>
  );
}

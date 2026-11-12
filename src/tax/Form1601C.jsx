// BIR Form 1601-C (January 2018 ENCS) — Monthly Remittance Return of Income Taxes Withheld on
// Compensation. Employees come from the Employees Record; their monthly figures roll up into Part II.
// Item numbers follow published descriptions of the form: keep them in ITEMS_1601C so a different
// numbering on the user's copy is a one-line change.
import React, { useMemo, useState } from "react";
import { FileText, Plus, Trash2, Copy, Users } from "lucide-react";
import {
  Field, NumberField, SectionHeader, getAvailableYears, MONTHS, uid, num, round2, fmtPlain, fmtParen,
  useRowSelection, SelectAllCheckbox, RowCheckbox, SelectionBar, EntryModalShell, LabeledField,
} from "../BookkeepingApp";
import {
  TaxStyle, makeCalc, useReturn, FormSection, ReturnStatusBar, BackgroundBlock, CompanyPrereqWarning, PaymentDetails,
  Declaration, companyInfo, monthKey, sectionBlock, infoBlock, paymentBlock, signatureBlock, useWorksheetExport,
  WorksheetButtons, ExportStatus, YearSelect, Notice,
} from "./shared";
import { employeeDisplayName, formatEmployeeTin, isMweRecord, employedDuring } from "./EmployeesRecord";

export const LINE_FIELDS_1601C = [
  ["totalComp", "Total Compensation", "14"],
  ["mweBasic", "Statutory Minimum Wage (MWE)", "15"],
  ["mweOtherPay", "Holiday / OT / Night Diff / Hazard (MWE)", "16"],
  ["thirteenthOther", "13th Month Pay & Other Benefits", "17"],
  ["deMinimis", "De Minimis Benefits", "18"],
  ["mandatoryContrib", "SSS / GSIS / PHIC / HDMF & Union Dues", "19"],
  ["otherNonTax", "Other Non-Taxable Compensation", "20"],
  ["notSubjectToWht", "Taxable Comp. Not Subject to WHT", "23"],
  ["taxWithheld", "Tax Withheld", "25"],
];

export function computeLine1601c(l) {
  const totalNonTax = round2(num(l.mweBasic) + num(l.mweOtherPay) + num(l.thirteenthOther) + num(l.deMinimis) + num(l.mandatoryContrib) + num(l.otherNonTax));
  const taxable = round2(num(l.totalComp) - totalNonTax);
  const netTaxable = round2(taxable - num(l.notSubjectToWht));
  return { totalNonTax, taxable, netTaxable };
}

export function lineWarnings1601c(l, rec) {
  const c = computeLine1601c(l);
  const w = [];
  if (c.totalNonTax > num(l.totalComp) + 0.004) w.push("Non-taxable amounts exceed Total Compensation");
  const mwe = rec ? isMweRecord(rec) : false;
  if (!mwe && (num(l.mweBasic) || num(l.mweOtherPay))) w.push("MWE columns used, but the record says SMW = No");
  if (mwe && num(l.taxWithheld) > 0) w.push("Minimum wage earner with tax withheld");
  if (num(l.taxWithheld) > 0 && c.netTaxable <= 0) w.push("Tax withheld with zero net taxable compensation");
  return w;
}

// The whole return's computation. Also used by the 1604-C (Part II monthly remittances, reconciliation).
export function compute1601c(form) {
  const c = makeCalc(form);
  const none = c.t("anyTaxWithheld", "") === "No";
  const lines = form.lines || [];
  const sum = (k) => (none ? 0 : round2(lines.reduce((s, l) => s + num(l[k]), 0)));
  c.set("i14", sum("totalComp"));
  c.set("i15", sum("mweBasic"));
  c.set("i16", sum("mweOtherPay"));
  c.set("i17", sum("thirteenthOther"));
  c.set("i18", sum("deMinimis"));
  c.set("i19", sum("mandatoryContrib"));
  c.set("i20", sum("otherNonTax"));
  c.set("i21", c.v.i15 + c.v.i16 + c.v.i17 + c.v.i18 + c.v.i19 + c.v.i20);
  c.set("i22", c.v.i14 - c.v.i21);
  c.set("i23", sum("notSubjectToWht"));
  c.set("i24", c.v.i22 - c.v.i23);
  c.set("i25", sum("taxWithheld"));
  c.set("i26", none ? 0 : c.m("i26"));
  c.set("i27", c.v.i25 + c.v.i26);
  c.set("i28", c.t("amended") === "Yes" ? c.m("i28") : 0);
  c.set("i29", c.m("i29"));
  c.set("i30", c.v.i28 + c.v.i29);
  c.set("i31", c.v.i27 - c.v.i30);
  c.man("i32"); c.man("i33"); c.man("i34");
  c.set("i35", c.v.i32 + c.v.i33 + c.v.i34);
  c.set("i36", c.v.i31 + c.v.i35);
  return c;
}

export const due1601c = (year, monthIdx) => (monthIdx === 11 ? new Date(year + 1, 0, 25) : new Date(year, monthIdx + 1, 10));

const r = (k) => (ref) => ref(k);
function sections1601c(c, form, info) {
  const amended = c.t("amended") === "Yes";
  const none = c.t("anyTaxWithheld", "") === "No";
  return {
    header: {
      title: "Return Header (items 1 to 5)", lines: [
        { no: "1", label: "For the Month (MM/YYYY)", k: "_month", t: "ro", text: info.monthLabel },
        { no: "2", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
        { no: "3", label: "Any Taxes Withheld?", k: "anyTaxWithheld", t: "select", opts: ["Yes", "No"], dflt: info.defaultAnyTax, note: "Defaults to Yes when total tax withheld is more than zero. \"No\" zeroes the form." },
        { no: "4", label: "Number of Sheet/s Attached", k: "sheetsAttached", t: "int" },
        { no: "5", label: "ATC", k: "_atc", t: "ro", text: info.atc },
      ],
    },
    relief: {
      title: "Part I (items 13 / 13A)", lines: [
        { no: "13", label: "Payees availing of tax relief under a special law or tax treaty?", k: "taxRelief", t: "yesno", dflt: "No" },
        { no: "13A", label: "If yes, specify", k: "taxReliefSpecify", t: "text", hide: c.t("taxRelief") !== "Yes" },
      ],
    },
    part2: {
      title: "Part II – Computation of Tax", lines: [
        { no: "14", label: "Total Amount of Compensation", k: "i14", t: "calc" },
        { heading: "Less: Non-Taxable / Exempt Compensation" },
        { no: "15", label: "Statutory Minimum Wage (for MWEs)", k: "i15", t: "calc" },
        { no: "16", label: "Holiday Pay, Overtime Pay, Night Shift Differential Pay, Hazard Pay (for MWEs only)", k: "i16", t: "calc" },
        { no: "17", label: "13th Month Pay and Other Benefits", k: "i17", t: "calc" },
        { no: "18", label: "De Minimis Benefits", k: "i18", t: "calc" },
        { no: "19", label: "SSS, GSIS, PHIC, HDMF Mandatory Contributions & Union Dues (employee's share only)", k: "i19", t: "calc" },
        { no: "20", label: "Other Non-Taxable Compensation (specify)", k: "i20", t: "calc" },
        { no: "21", label: "Total Non-Taxable Compensation (sum of items 15 to 20)", k: "i21", t: "calc", strong: true, f: (ref) => `${ref("i15")}+${ref("i16")}+${ref("i17")}+${ref("i18")}+${ref("i19")}+${ref("i20")}` },
        { no: "22", label: "Total Taxable Compensation (item 14 less item 21)", k: "i22", t: "calc", f: (ref) => `${ref("i14")}-${ref("i21")}` },
        { no: "23", label: "Less: Taxable compensation not subject to withholding tax (for employees, other than MWEs, receiving ₱250,000 & below for the year)", k: "i23", t: "calc" },
        { no: "24", label: "Net Taxable Compensation (item 22 less item 23)", k: "i24", t: "calc", strong: true, f: (ref) => `${ref("i22")}-${ref("i23")}` },
        { no: "25", label: "Total Taxes Withheld", k: "i25", t: "calc", strong: true },
        { no: "26", label: "Add/(Less): Adjustment of Taxes Withheld from Previous Month/s (from Part IV-A item 4; negative allowed)", k: "i26", t: "man", dis: none },
        { no: "27", label: "Taxes Withheld for Remittance (sum of items 25 and 26)", k: "i27", t: "calc", strong: true, f: (ref) => `${ref("i25")}+${ref("i26")}` },
        { no: "28", label: "Less: Tax Remitted in Return Previously Filed, if this is an amended return", k: "i28", t: "man", dis: !amended },
        { no: "29", label: "Other Remittances Made (specify)", k: "i29", t: "man" },
        { no: "", label: "Other remittances: specify", k: "i29Specify", t: "text", hide: !c.v.i29 },
        { no: "30", label: "Total Tax Remittances Made (sum of items 28 and 29)", k: "i30", t: "calc", f: (ref) => `${ref("i28")}+${ref("i29")}` },
        { no: "31", label: "Tax Still Due / (Over-remittance) (item 27 less item 30)", k: "i31", t: "calc", due: true, f: (ref) => `${ref("i27")}-${ref("i30")}` },
        { heading: "Add: Penalties (manual — the EOPT Act changed the surcharge and interest rules)" },
        { no: "32", label: "Surcharge", k: "i32", t: "man" },
        { no: "33", label: "Interest", k: "i33", t: "man" },
        { no: "34", label: "Compromise", k: "i34", t: "man" },
        { no: "35", label: "Total Penalties (sum of items 32 to 34)", k: "i35", t: "calc", f: (ref) => `${ref("i32")}+${ref("i33")}+${ref("i34")}` },
        { no: "36", label: "TOTAL AMOUNT STILL DUE / (OVER-REMITTANCE) (sum of items 31 and 35)", k: "i36", t: "calc", strong: true, due: true, f: (ref) => `${ref("i31")}+${ref("i35")}` },
      ],
    },
  };
}

function EmployeePickerModal({ candidates, onCancel, onPick }) {
  const [picked, setPicked] = useState(new Set());
  const [q, setQ] = useState("");
  const list = candidates.filter((e) => `${employeeDisplayName(e)} ${e.tin} ${e.empNo}`.toLowerCase().includes(q.toLowerCase()));
  const toggle = (id) => setPicked((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  return (
    <EntryModalShell title="Add employees" submitLabel={`Add ${picked.size || ""}`.trim()} onCancel={onCancel} onSubmit={() => picked.size && onPick([...picked])}>
      <LabeledField label="Search" wide><Field value={q} onChange={setQ} placeholder="Name, TIN or employee no." /></LabeledField>
      <div className="table-scroll" style={{ maxHeight: 360 }}>
        <table className="ledger-table">
          <thead><tr><th></th><th>Employee</th><th>TIN</th><th>MWE</th></tr></thead>
          <tbody>
            {list.length === 0 && <tr><td colSpan={4} className="empty-row">No active employees left to add for this month.</td></tr>}
            {list.map((e) => (
              <tr key={e.id} onClick={() => toggle(e.id)} style={{ cursor: "pointer" }}>
                <td className="text-center"><input type="checkbox" className="row-check" checked={picked.has(e.id)} readOnly /></td>
                <td>{employeeDisplayName(e)}</td><td>{formatEmployeeTin(e.tin)}</td><td>{isMweRecord(e) ? "Yes" : "No"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </EntryModalShell>
  );
}

const newLine = (rec, copyFrom) => {
  const l = { id: uid(), employeeId: rec.id, totalComp: 0, mweBasic: 0, mweOtherPay: 0, thirteenthOther: 0, deMinimis: 0, mandatoryContrib: 0, otherNonTax: 0, otherNonTaxSpecify: "", notSubjectToWht: 0, taxWithheld: 0 };
  // Copy keeps the regular pay; one-off 13th month / other benefits and the tax withheld are left for the user.
  if (copyFrom) return { ...l, ...copyFrom, id: uid(), totalComp: round2(num(copyFrom.totalComp) - num(copyFrom.thirteenthOther)), thirteenthOther: 0, taxWithheld: 0 };
  const basic = num(rec.defaultMonthlyBasic);
  if (basic) { l.totalComp = basic; if (isMweRecord(rec)) l.mweBasic = basic; }
  return l;
};

export function Form1601CPage({ data, setData, onNavigate }) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [monthIdx, setMonthIdx] = useState(now.getMonth());
  const key = monthKey(year, monthIdx);
  const rt = useReturn(data, setData, "form1601c", key, { lines: [] });
  const form = rt.form;
  const info0 = companyInfo(data);
  const records = data.employeeRecords || [];
  const recById = useMemo(() => new Map(records.map((e) => [e.id, e])), [records]);
  const lines = form.lines || [];
  const totalTax = round2(lines.reduce((s, l) => s + num(l.taxWithheld), 0));
  const info = {
    monthLabel: `${String(monthIdx + 1).padStart(2, "0")}/${year}`,
    atc: info0.category === "Government" ? "WW020" : "WW010",
    defaultAnyTax: totalTax > 0 || lines.length === 0 ? "Yes" : "No",
  };
  // Only an explicit "No" zeroes the form; the default merely suggests the answer for the worksheet.
  const calc = useMemo(() => compute1601c(form), [form]);
  const secs = sections1601c(calc, form, info);
  const none = calc.t("anyTaxWithheld", "") === "No";
  const sel = useRowSelection();
  const [picker, setPicker] = useState(false);

  const monthStart = new Date(year, monthIdx, 1), monthEndD = new Date(year, monthIdx + 1, 0);
  const inMonth = new Set(lines.map((l) => l.employeeId));
  // Inactive records stay pickable for the months before their separation date.
  const candidates = records.filter((e) => (e.active !== "No" || e.dateSeparated) && employedDuring(e, monthStart, monthEndD) && !inMonth.has(e.id));
  const setLines = (fn) => rt.update((cur) => ({ ...cur, lines: fn(cur.lines || []) }));
  const addEmployees = (ids) => { setLines((ls) => [...ls, ...ids.map((id) => newLine(recById.get(id)))]); setPicker(false); };
  const prevKey = monthIdx === 0 ? monthKey(year - 1, 11) : monthKey(year, monthIdx - 1);
  const prevLines = ((data.form1601c || {})[prevKey] || {}).lines || [];
  const copyPrev = () => setLines((ls) => {
    const have = new Set(ls.map((l) => l.employeeId));
    return [...ls, ...prevLines.filter((l) => !have.has(l.employeeId) && recById.has(l.employeeId) && employedDuring(recById.get(l.employeeId), monthStart, monthEndD)).map((l) => newLine(recById.get(l.employeeId), l))];
  });
  const setCell = (id, k, v) => setLines((ls) => ls.map((l) => l.id === id ? { ...l, [k]: v } : l));
  const removeSelected = () => { setLines((ls) => ls.filter((l) => !sel.selected.has(l.id))); sel.clear(); };

  const tot = (k) => round2(lines.reduce((s, l) => s + num(l[k]), 0));
  const totC = lines.reduce((a, l) => { const c = computeLine1601c(l); return { nt: a.nt + c.totalNonTax, tx: a.tx + c.taxable, net: a.net + c.netTaxable }; }, { nt: 0, tx: 0, net: 0 });

  const due = due1601c(year, monthIdx);
  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1601-C — Monthly Remittance Return of Income Taxes Withheld on Compensation",
    version: "January 2018 (ENCS)", formNo: "1601-C", period: `For the month of ${MONTHS[monthIdx]} ${year}`, company: info0,
    blocks: [
      sectionBlock(secs.header, calc, { m: { anyTaxWithheld: info.defaultAnyTax, ...(form.m || {}) } }),
      infoBlock("Part I – Background Information", bgItems(info0)),
      sectionBlock(secs.relief, calc, form),
      sectionBlock(secs.part2, calc, form),
      paymentBlock(form, "Part III – Details of Payment (items 37 to 40)"),
      { type: "table", title: "Annex: Employees' compensation for the month", sheet: "Employees", small: true,
        head: ["Employee", "TIN", ...LINE_FIELDS_1601C.slice(0, 7).map(([, h, n]) => `${h} (${n})`), "Total Non-Taxable (21)", "Taxable (22)", "Not Subj. to WHT (23)", "Net Taxable (24)", "Tax Withheld (25)"],
        body: lines.map((l) => { const rec = recById.get(l.employeeId); const c = computeLine1601c(l); return [employeeDisplayName(rec), formatEmployeeTin(rec && rec.tin), ...LINE_FIELDS_1601C.slice(0, 7).map(([k]) => num(l[k])), c.totalNonTax, c.taxable, num(l.notSubjectToWht), c.netTaxable, num(l.taxWithheld)]; }),
        foot: ["TOTAL", "", ...LINE_FIELDS_1601C.slice(0, 7).map(([k]) => tot(k)), round2(totC.nt), round2(totC.tx), tot("notSubjectToWht"), round2(totC.net), tot("taxWithheld")],
        numCols: Array.from({ length: 12 }, (_, i) => i + 2), landscape: true },
    ],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1601C_${String(monthIdx + 1).padStart(2, "0")}${year}_worksheet.pdf`,
  }));

  return (
    <div>
      <TaxStyle />
      <SectionHeader icon={FileText} title="BIR Form 1601-C" subtitle="Monthly Remittance Return of Income Taxes Withheld on Compensation (January 2018 ENCS). An encoding worksheet for eBIRForms / eFPS — no .dat file."
        right={<WorksheetButtons hook={exp} />} />
      <ExportStatus status={exp.status} />
      <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={[["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"], ["withholdingCategory", "Category of Withholding Agent"]]} />
      <ReturnStatusBar rt={rt} dueDefault={due} dueNote={monthIdx === 11 ? "December: January 25 · verify" : "verify weekends / holidays / eFPS"}>
        <select className="io-select" value={monthIdx} onChange={(e) => setMonthIdx(Number(e.target.value))}>{MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}</select>
        <YearSelect years={getAvailableYears(data)} value={year} onChange={setYear} />
      </ReturnStatusBar>
      {calc.t("amended") === "Yes" && <Notice>Amended return: item 28 is unlocked. Key in the figures of the original filing (tax remitted in the return previously filed).</Notice>}
      <FormSection section={secs.header} calc={calc} rt={{ ...rt, form: { ...form, m: { anyTaxWithheld: info.defaultAnyTax, ...(form.m || {}) } } }} />
      <BackgroundBlock items={bgItems(info0)} onNavigate={onNavigate} />
      <FormSection section={secs.relief} calc={calc} rt={rt} />

      <div className="tx-card">
        <div className="tx-card-head">
          <span>Employees' compensation for the month</span>
          <div className="io-toolbar">
            <button className="io-btn" onClick={() => setPicker(true)} disabled={none}><Plus size={13} /> Add employee</button>
            <button className="io-btn" onClick={() => addEmployees(candidates.map((e) => e.id))} disabled={none || !candidates.length}><Users size={13} /> Add all active employees ({candidates.length})</button>
            <button className="io-btn" onClick={copyPrev} disabled={none || !prevLines.length} title="Copies the employee list and pay of the previous month. Tax withheld is left blank."><Copy size={13} /> Copy lines from previous month</button>
          </div>
        </div>
        {records.length === 0 && <div className="tx-note">No employees yet. Add them in Setup → Employees Record. {onNavigate && <button className="tx-link" onClick={() => onNavigate("employeerecords")}>Open Employees Record</button>}</div>}
        {sel.selected.size > 0 && <SelectionBar count={sel.selected.size} onClear={sel.clear}><button className="io-btn danger" onClick={removeSelected}><Trash2 size={13} /> Remove selected</button></SelectionBar>}
        <div className="table-scroll">
          <table className="ledger-table tx-table">
            <thead><tr>
              <th style={{ width: 32 }}><SelectAllCheckbox ids={lines.map((l) => l.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
              <th style={{ minWidth: 200 }}>Employee</th><th>TIN</th>
              {LINE_FIELDS_1601C.slice(0, 7).map(([k, h, n]) => <th key={k} className="num-head" style={{ minWidth: 120 }} title={`Feeds item ${n}`}>{h} <span className="tx-pill">{n}</span></th>)}
              <th className="num-head">Total Non-Taxable</th><th className="num-head">Taxable Comp.</th>
              <th className="num-head" style={{ minWidth: 120 }}>Not Subj. to WHT <span className="tx-pill">23</span></th>
              <th className="num-head">Net Taxable</th>
              <th className="num-head" style={{ minWidth: 120 }}>Tax Withheld <span className="tx-pill">25</span></th>
            </tr></thead>
            <tbody>
              {lines.length === 0 && <tr><td colSpan={15} className="empty-row">No employees on this month's return yet.</td></tr>}
              {lines.map((l) => {
                const rec = recById.get(l.employeeId);
                const c = computeLine1601c(l);
                const warns = lineWarnings1601c(l, rec);
                return (
                  <tr key={l.id} className={sel.selected.has(l.id) ? "row-selected" : ""}>
                    <td className="text-center"><RowCheckbox id={l.id} selected={sel.selected} toggle={sel.toggle} /></td>
                    <td className="wrap">{rec ? employeeDisplayName(rec) : <span className="tx-amber">Employee record deleted</span>}{rec && isMweRecord(rec) && <span className="tx-pill">MWE</span>}
                      {warns.map((w) => <div key={w} className="tx-amber">⚠ {w}</div>)}
                      {num(l.otherNonTax) !== 0 && <div style={{ maxWidth: 220 }}><Field value={l.otherNonTaxSpecify} onChange={(v) => setCell(l.id, "otherNonTaxSpecify", v)} placeholder="Other non-taxable: specify" /></div>}
                    </td>
                    <td>{rec ? formatEmployeeTin(rec.tin) : ""}</td>
                    {LINE_FIELDS_1601C.slice(0, 7).map(([k]) => <td key={k}><NumberField value={l[k]} onChange={(v) => setCell(l.id, k, v)} align="right" disabled={none} /></td>)}
                    <td className="num">{fmtPlain(c.totalNonTax)}</td>
                    <td className="num">{fmtParen(c.taxable)}</td>
                    <td><NumberField value={l.notSubjectToWht} onChange={(v) => setCell(l.id, "notSubjectToWht", v)} align="right" disabled={none} /></td>
                    <td className="num">{fmtParen(c.netTaxable)}</td>
                    <td><NumberField value={l.taxWithheld} onChange={(v) => setCell(l.id, "taxWithheld", v)} align="right" disabled={none} /></td>
                  </tr>
                );
              })}
            </tbody>
            {lines.length > 0 && (
              <tfoot><tr>
                <td colSpan={3} className="totals-label">Totals ({lines.length})</td>
                {LINE_FIELDS_1601C.slice(0, 7).map(([k]) => <td key={k} className="num">{fmtPlain(tot(k))}</td>)}
                <td className="num">{fmtPlain(totC.nt)}</td><td className="num">{fmtParen(totC.tx)}</td>
                <td className="num">{fmtPlain(tot("notSubjectToWht"))}</td><td className="num">{fmtParen(totC.net)}</td><td className="num">{fmtPlain(tot("taxWithheld"))}</td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>

      {none && <Notice tone="info">Any Taxes Withheld = No: the employee lines and Part II are zeroed.</Notice>}
      <FormSection section={secs.part2} calc={calc} rt={rt} />
      <PaymentDetails rt={rt} title="Part III – Details of Payment (items 37 to 40)" firstItem={37} amountDue={calc.v.i36} />
      <Declaration data={data} rt={rt} />
      {picker && <EmployeePickerModal candidates={candidates} onCancel={() => setPicker(false)} onPick={addEmployees} />}
    </div>
  );
}

export const bgItems = (i) => [
  ["6", "TIN", i.tinText], ["7", "RDO Code", i.rdo], ["8", "Withholding Agent's Name", i.name],
  ["9", "Registered Address", i.address], ["9A", "ZIP Code", i.zip], ["10", "Contact Number", i.contact],
  ["11", "Category of Withholding Agent", i.category], ["12", "Email Address", i.email],
];

// BIR Form 0619-E (January 2018) Monthly Remittance Form of Creditable Income Taxes Withheld (Expanded)
// and BIR Form 1601-EQ (January 2019 ENCS) Quarterly Remittance Return. Figures come from the EWT
// Schedule (computeEwtSchedule / ewtSourceRows), the same rows the QAP lists. Marking a return as filed
// writes its payment into the EWT Schedule's PAID columns so it isn't typed twice.
import React, { useMemo, useState } from "react";
import { FileText } from "lucide-react";
import {
  SectionHeader, getAvailableYears, MONTHS, num, round2, fmtPlain, computeEwtSchedule, ewtSourceRows, buildQapRows,
  atcKey, inRange,
} from "../BookkeepingApp";
import {
  TaxStyle, makeCalc, useReturn, FormSection, ReturnStatusBar, BackgroundBlock, CompanyPrereqWarning, PaymentDetails,
  Declaration, companyInfo, monthKey, quarterKey, quarterRange, sectionBlock, infoBlock, paymentBlock, signatureBlock,
  useWorksheetExport, WorksheetButtons, ExportStatus, YearSelect, Notice, ReconStrip, SubTabs, AttachmentChecklist,
  PAYMENT_ROWS, mdy,
} from "./shared";

export const due0619e = (year, monthIdx) => new Date(year, monthIdx + 1, 10);
export const due1601eq = (year, q) => new Date(year, q * 3 + 1, 0); // last day of the month after the quarter

const PREREQ = [["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"], ["withholdingCategory", "Category of Withholding Agent"]];

/* ------------------------------ 0619-E ------------------------------ */

export function compute0619e(form, monthEwt) {
  const c = makeCalc(form);
  c.a("i14", monthEwt);
  c.set("i15", c.t("amended") === "Yes" ? c.m("i15") : 0);
  c.set("i16", c.v.i14 - c.v.i15);
  c.man("i17a"); c.man("i17b"); c.man("i17c");
  c.set("i17d", c.v.i17a + c.v.i17b + c.v.i17c);
  c.set("i18", c.v.i16 + c.v.i17d);
  return c;
}

function sections0619e(c, info) {
  const amended = c.t("amended") === "Yes";
  return {
    header: { title: "Header (items 1 to 6)", lines: [
      { no: "1", label: "For the Month of (MM/YYYY)", k: "_m", t: "ro", text: info.monthText },
      { no: "2", label: "Due Date (MM/DD/YYYY)", k: "_due", t: "ro", text: info.dueText },
      { no: "3", label: "Amended Form?", k: "amended", t: "yesno", dflt: "No" },
      { no: "4", label: "Any Taxes Withheld?", k: "anyTaxWithheld", t: "select", opts: ["Yes", "No"], dflt: info.defaultAny },
      { no: "5", label: "ATC", k: "_atc", t: "ro", text: "WME10 (confirm against your copy)" },
      { no: "6", label: "Tax Type Code", k: "_ttc", t: "ro", text: "WE" },
    ]},
    part2: { title: "Part II – Tax Remittance", lines: [
      { no: "14", label: "Amount of Remittance", k: "i14", t: "auto", src: "EWT Schedule: the month's total EWT (Purchase Journal and Cash Disbursements rows with an ATC)" },
      { no: "15", label: "Less: Amount Remitted from Previously Filed Form, if this is an amended form", k: "i15", t: "man", dis: !amended },
      { no: "16", label: "Net Amount of Remittance (item 14 less item 15)", k: "i16", t: "calc", strong: true, f: (r) => `${r("i14")}-${r("i15")}` },
      { heading: "Add: Penalties (manual)" },
      { no: "17A", label: "Surcharge", k: "i17a", t: "man" },
      { no: "17B", label: "Interest", k: "i17b", t: "man" },
      { no: "17C", label: "Compromise", k: "i17c", t: "man" },
      { no: "17D", label: "Total Penalties (sum of items 17A to 17C)", k: "i17d", t: "calc", f: (r) => `${r("i17a")}+${r("i17b")}+${r("i17c")}` },
      { no: "18", label: "TOTAL AMOUNT OF REMITTANCE (sum of items 16 and 17D)", k: "i18", t: "calc", strong: true, due: true, f: (r) => `${r("i16")}+${r("i17d")}` },
    ]},
  };
}
const bg0619 = (i) => [["7", "TIN", i.tinText], ["8", "RDO Code", i.rdo], ["9", "Withholding Agent's Name", i.name], ["10", "Registered Address", i.address], ["10A", "ZIP Code", i.zip], ["11", "Contact Number", i.contact], ["12", "Category of Withholding Agent", i.category], ["13", "Email Address", i.email]];
const bg1601eq = (i) => [["6", "TIN", i.tinText], ["7", "RDO Code", i.rdo], ["8", "Withholding Agent's Name", i.name], ["9", "Registered Address", i.address], ["9A", "ZIP Code", i.zip], ["10", "Contact Number", i.contact], ["11", "Category of Withholding Agent", i.category], ["12", "Email Address", i.email]];

// Writes a filed return's payment into the EWT Schedule settlement block (data.ewtSchedule[year][monthIdx]).
function ewtWriteThrough(setData, year, monthIdx, form, paid) {
  const p = form.payment || {};
  const firstPaid = PAYMENT_ROWS.map(([k]) => p[k] || {}).find((r) => num(r.amount)) || {};
  setData((d) => {
    const all = d.ewtSchedule || {};
    const yr = all[year] || {};
    return { ...d, ewtSchedule: { ...all, [year]: { ...yr, [monthIdx]: { ...(yr[monthIdx] || {}), paid: round2(paid), dateFiled: form.dateFiled || "", datePaid: firstPaid.date || form.dateFiled || "", bank: firstPaid.bank || "", refNo: (form.ror || {}).refNo || firstPaid.number || form.filedRef || "" } } } };
  });
}

function Form0619E({ data, setData, onNavigate, year, setYear, onGoQuarter }) {
  const now = new Date();
  const [monthIdx, setMonthIdx] = useState(() => { const m = now.getMonth(); return m % 3 === 2 ? m - 1 : m; });
  const info0 = companyInfo(data);
  const rt = useReturn(data, setData, "form0619e", monthKey(year, monthIdx));
  const sched = useMemo(() => computeEwtSchedule(data, year), [data.purchases, data.disbursements, data.atc, data.ewtSchedule, year]);
  const monthEwt = round2(sched.months[monthIdx].totalEwt);
  const form = rt.form;
  const calc = useMemo(() => compute0619e(form, monthEwt), [form, monthEwt]);
  const third = monthIdx % 3 === 2;
  const due = due0619e(year, monthIdx);
  const info = { monthText: `${String(monthIdx + 1).padStart(2, "0")}/${year}`, dueText: form.dueDate || mdy(due), defaultAny: calc.v.i14 > 0 ? "Yes" : "No" };
  const secs = sections0619e(calc, info);
  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 0619-E — Monthly Remittance Form of Creditable Income Taxes Withheld (Expanded)", version: "January 2018", formNo: "0619-E",
    period: `For the month of ${MONTHS[monthIdx]} ${year}`, company: info0,
    blocks: [sectionBlock(secs.header, calc, form), infoBlock("Part I – Background Information", bg0619(info0)), sectionBlock(secs.part2, calc, form), paymentBlock(form, "Part III – Details of Payment (items 19 to 22)")],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_0619E_${String(monthIdx + 1).padStart(2, "0")}${year}_worksheet.pdf`,
  }));
  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}><WorksheetButtons hook={exp} /></div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={PREREQ} />
    <ReturnStatusBar rt={rt} dueDefault={third ? null : due} dueNote="10th of the following month · eFPS filers follow their own deadlines · verify"
      onFiled={(m) => ewtWriteThrough(setData, year, monthIdx, { ...form, dateFiled: m.date, filedRef: m.ref }, calc.v.i16)}>
      <select className="io-select" value={monthIdx} onChange={(e) => setMonthIdx(Number(e.target.value))}>{MONTHS.map((m, i) => <option key={m} value={i}>{m}{i % 3 === 2 ? " (paid with 1601-EQ)" : ""}</option>)}</select>
      <YearSelect years={getAvailableYears(data)} value={year} onChange={setYear} />
    </ReturnStatusBar>
    {third ? (
      <Notice tone="info">{MONTHS[monthIdx]} is the third month of the quarter: it is paid with the 1601-EQ, so there is no 0619-E for it. <button className="tx-link" onClick={() => onGoQuarter(Math.floor(monthIdx / 3) + 1)}>Open the 1601-EQ</button></Notice>
    ) : (<>
      {calc.t("amended") === "Yes" && <Notice>Amended form: item 15 is the amount remitted on the original. Key in the original's figures.</Notice>}
      <FormSection section={secs.header} calc={calc} rt={{ ...rt, form: { ...form, m: { anyTaxWithheld: info.defaultAny, ...(form.m || {}) } } }} />
      <BackgroundBlock items={bg0619(info0)} onNavigate={onNavigate} />
      <FormSection section={secs.part2} calc={calc} rt={rt} />
      <PaymentDetails rt={rt} title="Part III – Details of Payment (items 19 to 22)" firstItem={19} amountDue={calc.v.i18} />
      <Declaration data={data} rt={rt} />
      <div className="rc-warn">Marking the 0619-E as filed writes item 16, the filing date and the payment details into the EWT Schedule's PAID columns for {MONTHS[monthIdx]}.</div>
    </>)}
  </>);
}

/* ------------------------------ 1601-EQ ------------------------------ */

// Per-ATC tax base and tax withheld for the quarter, from the EWT Schedule's source rows.
export function ewtByAtcForQuarter(data, year, q) {
  const { start, end } = quarterRange(year, q);
  const atcRef = new Map((data.atc || []).map((a) => [atcKey(a.code), a]));
  const map = new Map();
  ewtSourceRows(data).forEach((e) => {
    if (!inRange(e.date, start, end)) return;
    const k = atcKey(e.atc);
    const g = map.get(k) || { atc: k, base: 0, tax: 0, rate: num((atcRef.get(k) || {}).rate) || e.rate };
    g.base += e.taxBase; g.tax += e.ewt;
    map.set(k, g);
  });
  return [...map.values()].map((g) => ({ ...g, base: round2(g.base), tax: round2(g.tax) })).sort((a, b) => a.atc.localeCompare(b.atc));
}

export function compute1601eq(form, atcRows, m1, m2, prevOver) {
  const c = makeCalc(form);
  const none = c.t("anyTaxWithheld", "") === "No";
  let total = 0;
  atcRows.forEach((r) => {
    c.a(`base_${r.atc}`, none ? 0 : r.base);
    c.a(`tax_${r.atc}`, none ? 0 : r.tax);
    total += c.v[`tax_${r.atc}`];
  });
  c.set("i19", total);
  c.a("i20", m1);
  c.a("i21", m2);
  c.set("i22", c.t("amended") === "Yes" ? c.m("i22") : 0);
  c.a("i23", prevOver);
  c.man("i24");
  c.set("i25", c.v.i20 + c.v.i21 + c.v.i22 + c.v.i23 + c.v.i24);
  c.set("i26", c.v.i19 - c.v.i25);
  c.man("i27"); c.man("i28"); c.man("i29");
  c.set("i30", c.v.i27 + c.v.i28 + c.v.i29);
  c.set("i31", c.v.i26 + c.v.i30);
  return c;
}

const OVERPAY_OPTS = ["To be refunded", "To be issued a Tax Credit Certificate", "To be carried over to the next quarter"];

function sections1601eq(c, info, atcRows) {
  const amended = c.t("amended") === "Yes";
  const atcLines = atcRows.map((r, i) => ({ no: i < 6 ? String(13 + i) : `${18}+`, label: `${r.atc} @ ${round2(r.rate * 100)}%`, k: [`base_${r.atc}`, `tax_${r.atc}`], t: "auto", src: "EWT Schedule / QAP rows for the quarter" }));
  return {
    header: { title: "Header", lines: [
      { no: "1", label: "For the Year", k: "_y", t: "ro", text: String(info.year) },
      { no: "2", label: "Quarter", k: "_q", t: "ro", text: `${["1st", "2nd", "3rd", "4th"][info.q - 1]} Quarter` },
      { no: "3", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "4", label: "Any Taxes Withheld?", k: "anyTaxWithheld", t: "select", opts: ["Yes", "No"], dflt: info.defaultAny },
      { no: "5", label: "Number of Sheet/s Attached (confirm item number on your copy)", k: "sheetsAttached", t: "int", dflt: String(info.attachCount) },
      { no: "", label: "Tax relief under special law / treaty? (confirm whether in the header or Part I)", k: "taxRelief", t: "yesno", dflt: info.c.taxRelief || "No" },
    ]},
    atc: { title: "Part II – Computation of Tax (items 13 to 19)", cols: ["Tax Base (consolidated for the quarter)", "Tax Withheld (consolidated for the quarter)"], lines: [
      ...atcLines,
      ...(atcRows.length === 0 ? [{ no: "13", label: "No ATC activity this quarter", k: [null, null], t: [null, null] }] : []),
      { no: "19", label: "Total Taxes Withheld for the Quarter (sum of items 13 to 18)", k: [null, "i19"], t: [null, "calc"], strong: true },
    ]},
    part2: { title: "Part II (continued) – Remittances", lines: [
      { heading: "Less: Remittances Made" },
      { no: "20", label: "1st Month of the Quarter (net remittance, item 16 of the 0619-E)", k: "i20", t: "auto", src: info.m1src },
      { no: "21", label: "2nd Month of the Quarter (net remittance, item 16 of the 0619-E)", k: "i21", t: "auto", src: info.m2src },
      { no: "22", label: "Tax Remitted in Return Previously Filed, if this is an amended return", k: "i22", t: "man", dis: !amended },
      { no: "23", label: "Over-remittance from Previous Quarter of the same taxable year", k: "i23", t: "auto", src: "Previous quarter's 1601-EQ over-remittance marked “carried over”" },
      { no: "24", label: "Other Payments Made (please attach proof of payment, BIR Form No. 0605)", k: "i24", t: "man" },
      { no: "25", label: "Total Remittances Made (sum of items 20 to 24)", k: "i25", t: "calc", f: (r) => `${r("i20")}+${r("i21")}+${r("i22")}+${r("i23")}+${r("i24")}` },
      { no: "26", label: "Tax Still Due / (Over-remittance) (item 19 less item 25)", k: "i26", t: "calc", due: true, f: (r) => `${r("i19")}-${r("i25")}` },
      { heading: "Add: Penalties (manual)" },
      { no: "27", label: "Surcharge", k: "i27", t: "man" },
      { no: "28", label: "Interest", k: "i28", t: "man" },
      { no: "29", label: "Compromise", k: "i29", t: "man" },
      { no: "30", label: "Total Penalties (sum of items 27 to 29)", k: "i30", t: "calc", f: (r) => `${r("i27")}+${r("i28")}+${r("i29")}` },
      { no: "31", label: "TOTAL AMOUNT STILL DUE / (Over-remittance) (sum of items 26 and 30)", k: "i31", t: "calc", strong: true, due: true, f: (r) => `${r("i26")}+${r("i30")}` },
      { no: "", label: "If over-remittance, mark one (confirm wording on your copy)", k: "overpayChoice", t: "select", opts: OVERPAY_OPTS, hide: !(c.v.i31 < 0) },
    ]},
  };
}

function net0619(data, year, monthIdx) {
  const f = (data.form0619e || {})[monthKey(year, monthIdx)];
  if (!f) return null;
  const sched = computeEwtSchedule(data, year);
  return compute0619e(f, round2(sched.months[monthIdx].totalEwt)).v.i16;
}

const ATTACH_1601EQ = [["qap", "QAP acknowledgment (\"Validation Successful\") or filed QAP attached", "qap"], ["form0605", "BIR Form 0605 (proof of other payments, if item 24 is used)"]];

function Form1601EQ({ data, setData, postings, onNavigate, year, setYear, q, setQ }) {
  const info0 = companyInfo(data);
  const m1i = (q - 1) * 3, m2i = m1i + 1, m3i = m1i + 2;
  const m1 = net0619(data, year, m1i), m2 = net0619(data, year, m2i);
  const prevForm = q > 1 ? (data.form1601eq || {})[quarterKey(year, q - 1)] : null;
  const atcRows = useMemo(() => ewtByAtcForQuarter(data, year, q), [data.purchases, data.disbursements, data.atc, year, q]);
  let prevOver = 0;
  if (prevForm && ((prevForm.m || {}).overpayChoice || "") === OVERPAY_OPTS[2]) {
    const prevAtc = ewtByAtcForQuarter(data, year, q - 1);
    const pm1 = net0619(data, year, m1i - 3), pm2 = net0619(data, year, m2i - 3);
    const pc = compute1601eq(prevForm, prevAtc, pm1 || 0, pm2 || 0, 0);
    if (pc.v.i31 < 0) prevOver = round2(-pc.v.i31);
  }
  const rt = useReturn(data, setData, "form1601eq", quarterKey(year, q), { snapshot0619: { m1, m2 } });
  const form = rt.form;
  const calc = useMemo(() => compute1601eq(form, atcRows, m1 || 0, m2 || 0, prevOver), [form, atcRows, m1, m2, prevOver]);
  const info = {
    year, q, c: data.company || {}, defaultAny: atcRows.length ? "Yes" : "No",
    attachCount: ATTACH_1601EQ.filter(([k]) => (form.attach || {})[k]).length,
    m1src: m1 === null ? `No 0619-E saved for ${MONTHS[m1i]}: type the amount remitted` : `0619-E for ${MONTHS[m1i]}, item 16`,
    m2src: m2 === null ? `No 0619-E saved for ${MONTHS[m2i]}: type the amount remitted` : `0619-E for ${MONTHS[m2i]}, item 16`,
  };
  const secs = sections1601eq(calc, info, atcRows);
  const snap = form.snapshot0619 || {};
  const changed = rt.exists && ((snap.m1 ?? null) !== (m1 ?? null) || (snap.m2 ?? null) !== (m2 ?? null));

  const { start, end } = quarterRange(year, q);
  const sched = useMemo(() => computeEwtSchedule(data, year), [data.purchases, data.disbursements, data.atc, data.ewtSchedule, year]);
  const qap = useMemo(() => round2(buildQapRows(data, q, year).reduce((s, r) => s + num(r.ewt), 0)), [data, q, year]);
  const gl = useMemo(() => round2((postings || []).filter((p) => p.code === "2300" && inRange(p.date, start, end)).reduce((s, p) => s + p.credit, 0)), [postings, year, q]);

  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1601-EQ — Quarterly Remittance Return of Creditable Income Taxes Withheld (Expanded)", version: "January 2019 (ENCS)", formNo: "1601-EQ",
    period: `${["1st", "2nd", "3rd", "4th"][q - 1]} Quarter ${year}`, company: info0,
    blocks: [
      sectionBlock(secs.header, calc, { ...form, m: { anyTaxWithheld: info.defaultAny, ...(form.m || {}) } }),
      infoBlock("Part I – Background Information", bg1601eq(info0)),
      { type: "table", title: "Part II – ATC rows (items 13 to 18)", head: ["Item", "ATC", "Tax Base", "Tax Rate", "Tax Withheld"],
        body: atcRows.map((r, i) => [i < 6 ? String(13 + i) : "18+", r.atc, calc.v[`base_${r.atc}`], `${round2(r.rate * 100)}%`, calc.v[`tax_${r.atc}`]]),
        foot: ["19", "Total Taxes Withheld for the Quarter", "", "", calc.v.i19], numCols: [2, 4] },
      sectionBlock(secs.part2, calc, form),
      { type: "table", title: "Monthly remittances", head: ["Month", "Source", "Amount"], body: [
        [MONTHS[m1i], m1 === null ? "manual" : "0619-E item 16", calc.v.i20], [MONTHS[m2i], m2 === null ? "manual" : "0619-E item 16", calc.v.i21], [MONTHS[m3i], "paid with this 1601-EQ (item 31)", Math.max(0, calc.v.i31)]], numCols: [2] },
      paymentBlock(form, "Part III – Details of Payment (items 32 to 35)"),
    ], signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1601EQ_Q${q}${year}_worksheet.pdf`,
  }));

  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}><WorksheetButtons hook={exp} /></div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={PREREQ} />
    <ReturnStatusBar rt={rt} dueDefault={due1601eq(year, q)} dueNote="last day of the month after the quarter · verify"
      onFiled={(m) => { rt.patch({ snapshot0619: { m1, m2 } }); ewtWriteThrough(setData, year, m3i, { ...form, dateFiled: m.date, filedRef: m.ref }, Math.max(0, calc.v.i31)); }}>
      <select className="io-select" value={q} onChange={(e) => setQ(Number(e.target.value))}>{[1, 2, 3, 4].map((x) => <option key={x} value={x}>{["1st", "2nd", "3rd", "4th"][x - 1]} Quarter</option>)}</select>
      <YearSelect years={getAvailableYears(data)} value={year} onChange={setYear} />
    </ReturnStatusBar>
    {changed && <Notice>A 0619-E for this quarter changed since this 1601-EQ was prepared (items 20 / 21 now read {fmtPlain(m1 || 0)} and {fmtPlain(m2 || 0)}; were {fmtPlain(snap.m1 || 0)} and {fmtPlain(snap.m2 || 0)}). Typed overrides are kept. <button className="tx-link" onClick={() => rt.patch({ snapshot0619: { m1, m2 } })}>Acknowledge</button></Notice>}
    {calc.t("amended") === "Yes" && <Notice>Amended return: item 22 is unlocked. Key in the figures of the original filing.</Notice>}
    {atcRows.length > 6 && <Notice>{atcRows.length} ATCs have activity this quarter. The paper form has six lines (items 13 to 18); all of them are shown and totalled here. Encode the extra lines as the eBIRForms package allows.</Notice>}
    <FormSection section={secs.header} calc={calc} rt={{ ...rt, form: { ...form, m: { anyTaxWithheld: info.defaultAny, sheetsAttached: String(info.attachCount), ...(form.m || {}) } } }} />
    <BackgroundBlock items={bg1601eq(info0)} onNavigate={onNavigate} />
    <FormSection section={secs.atc} calc={calc} rt={rt} />
    <FormSection section={secs.part2} calc={calc} rt={rt} />
    <PaymentDetails rt={rt} title="Part III – Details of Payment (items 32 to 35)" firstItem={32} amountDue={calc.v.i31} />
    <AttachmentChecklist rt={rt} items={ATTACH_1601EQ} onOpen={(k) => onNavigate && onNavigate(k)} />
    <ReconStrip title="EWT check" colA="1601-EQ item 19" colB="Source" items={[
      { label: "vs EWT Schedule, quarter total", a: calc.v.i19, b: round2(sched.quarters[q - 1].totalEwt) },
      { label: "vs QAP total for the quarter", a: calc.v.i19, b: qap },
      { label: "vs GL 2300 Withholding Taxes Expanded (credits in the quarter)", a: calc.v.i19, b: gl },
    ]} causes="ATCs missing an EWT Group; ATC-tagged rows with a bad date; manual General Journal credits to 2300; or an override of an ATC row." />
    <Declaration data={data} rt={rt} />
    <div className="rc-warn">Marking the 1601-EQ as filed writes the amount paid with it (item 31, if positive) into the EWT Schedule's PAID column for {MONTHS[m3i]}.</div>
  </>);
}

export function EwtReturnsPage({ data, setData, postings, onNavigate }) {
  const now = new Date();
  const [tab, setTab] = useState("0619e");
  const [year, setYear] = useState(now.getFullYear());
  const [q, setQ] = useState(Math.floor(now.getMonth() / 3) + 1);
  return (
    <div>
      <TaxStyle />
      <SectionHeader icon={FileText} title="0619-E / 1601-EQ" subtitle="Expanded withholding tax remittances. Auto-filled from the EWT Schedule; encoding worksheets for eBIRForms / eFPS (no .dat). The QAP is the 1601-EQ attachment." />
      <SubTabs tabs={[["0619e", "0619-E (Monthly)"], ["1601eq", "1601-EQ (Quarterly)"]]} value={tab} onChange={setTab} />
      {tab === "0619e"
        ? <Form0619E data={data} setData={setData} onNavigate={onNavigate} year={year} setYear={setYear} onGoQuarter={(qq) => { setQ(qq); setTab("1601eq"); }} />
        : <Form1601EQ data={data} setData={setData} postings={postings} onNavigate={onNavigate} year={year} setYear={setYear} q={q} setQ={setQ} />}
    </div>
  );
}

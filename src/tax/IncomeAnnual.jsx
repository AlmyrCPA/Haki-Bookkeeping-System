// BIR Form 1701 (January 2018 ENCS, individuals) / 1702-RT (January 2018 ENCS, v2018C, corporations at
// the regular rate). Annual counterparts of the 1701Q / 1702Q tab, reusing its figures (incomeFigures),
// the ITR deduction-line mapping on the Chart of Accounts and the saved quarterly returns.
import React, { useMemo, useState } from "react";
import { FileText, Plus } from "lucide-react";
import { SectionHeader, getAvailableYears, num, round2, fmtPlain, uid, DelBtn, Field, NumberField, LabeledField, isVatCompany, computeVatSchedule, aggregateVatByYear, ITR_LINES, normalizeTin } from "../BookkeepingApp";
import {
  TaxStyle, makeCalc, useReturn, FormSection, ReturnStatusBar, BackgroundBlock, CompanyPrereqWarning, PaymentDetails,
  Declaration, companyInfo, quarterKey, fiscalYearRange, fiscalQuarterRange, sectionBlock, infoBlock, paymentBlock, signatureBlock,
  useWorksheetExport, WorksheetButtons, ExportStatus, YearSelect, Notice, ReconStrip, graduatedTax, paymentTotal,
  AttachmentChecklist, attachCount, EIGHT_PCT_LIMIT, taxTableFor,
} from "./shared";
import { incomeFigures, cwtInRange, mcitApplies, corpRate, faDepreciationForYear } from "./incomeData";
import { build1701q, build1702q, electionFromCompany, RateSelect } from "./IncomeQuarterly";

export const dueAnnual = (fy) => new Date(fy.end.getFullYear(), fy.end.getMonth() + 4, 15);
const lineNo = (l) => l.split(" ")[0];
const lineKey = (l) => `d_${lineNo(l)}`;

/* ------------------------------ NOLCO / MCIT carry tables ------------------------------ */

const nolcoE = (r) => round2(num(r.nol) - (num(r.appliedPrev) + num(r.expired) + num(r.appliedCurrent)));
// Pre-fill from last year's saved return: unapplied balances carry over; last year's own net operating loss becomes a new row.
function prefillNolco(prevForm, prevYearNol, prevYear) {
  const rows = [];
  ((prevForm && prevForm.nolco) || []).forEach((r) => {
    if (nolcoE(r) > 0.004) rows.push({ id: uid(), year: r.year, nol: num(r.nol), appliedPrev: round2(num(r.appliedPrev) + num(r.appliedCurrent)), expired: 0, appliedCurrent: 0 });
  });
  if (prevYearNol > 0) rows.push({ id: uid(), year: String(prevYear), nol: round2(prevYearNol), appliedPrev: 0, expired: 0, appliedCurrent: 0 });
  return rows;
}
const mcitG = (r) => round2(Math.max(0, num(r.mcit) - num(r.normal)) - (num(r.appliedPrev) + num(r.expired) + num(r.appliedCurrent)));
function prefillMcit(prevForm, prevNormal, prevMcit, prevYear) {
  const rows = [];
  ((prevForm && prevForm.mcitCarry) || []).forEach((r) => { if (mcitG(r) > 0.004) rows.push({ id: uid(), year: r.year, normal: num(r.normal), mcit: num(r.mcit), appliedPrev: round2(num(r.appliedPrev) + num(r.appliedCurrent)), expired: 0, appliedCurrent: 0 }); });
  if (prevMcit > prevNormal) rows.push({ id: uid(), year: String(prevYear), normal: round2(prevNormal), mcit: round2(prevMcit), appliedPrev: 0, expired: 0, appliedCurrent: 0 });
  return rows.slice(-3);
}

function CarryTable({ title, note, rows, onChange, cols, computedCol, totalCol, totalLabel }) {
  const set = (id, k, v) => onChange(rows.map((r) => r.id === id ? { ...r, [k]: v } : r));
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>{title}</span><button className="io-btn" onClick={() => onChange([...rows, { id: uid() }])}><Plus size={13} /> Add year</button></div>
      {note && <div className="tx-note">{note}</div>}
      <div className="table-scroll">
        <table className="ledger-table tx-table">
          <thead><tr>{cols.map(([k, h]) => <th key={k} className={k === "year" ? "" : "num-head"}>{h}</th>)}<th className="num-head">{computedCol[0]}</th><th style={{ width: 36 }}></th></tr></thead>
          <tbody>
            {rows.length === 0 && <tr><td colSpan={cols.length + 2} className="empty-row">No carry-over rows.</td></tr>}
            {rows.map((r) => (
              <tr key={r.id}>
                {cols.map(([k]) => <td key={k}>{k === "year" ? <Field value={r.year} onChange={(v) => set(r.id, "year", v.replace(/\D/g, "").slice(0, 4))} placeholder="YYYY" /> : <NumberField value={r[k]} onChange={(v) => set(r.id, k, v)} align="right" />}</td>)}
                <td className="num">{fmtPlain(computedCol[1](r))}</td>
                <td className="text-center"><DelBtn onClick={() => onChange(rows.filter((x) => x.id !== r.id))} /></td>
              </tr>
            ))}
          </tbody>
          {rows.length > 0 && <tfoot><tr><td colSpan={cols.length - 1} className="totals-label">{totalLabel}</td><td className="num">{fmtPlain(round2(rows.reduce((s, r) => s + num(r[totalCol]), 0)))}</td><td></td><td></td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

function ListTable({ title, rows, onChange, cols, note }) {
  const set = (id, k, v) => onChange(rows.map((r) => r.id === id ? { ...r, [k]: v } : r));
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>{title}</span><button className="io-btn" onClick={() => onChange([...rows, { id: uid() }])}><Plus size={13} /> Add row</button></div>
      {note && <div className="tx-note">{note}</div>}
      <div className="table-scroll">
        <table className="ledger-table tx-table">
          <thead><tr>{cols.map(([k, h, t]) => <th key={k} className={t === "num" ? "num-head" : ""}>{h}</th>)}<th style={{ width: 36 }}></th></tr></thead>
          <tbody>
            {rows.length === 0 && <tr><td colSpan={cols.length + 1} className="empty-row">No rows.</td></tr>}
            {rows.map((r) => (
              <tr key={r.id}>
                {cols.map(([k, , t, opts]) => <td key={k}>{t === "num" ? <NumberField value={r[k]} onChange={(v) => set(r.id, k, v)} align="right" /> : t === "select" ? <Field type="select" options={opts} value={r[k]} onChange={(v) => set(r.id, k, v)} /> : <Field value={r[k]} onChange={(v) => set(r.id, k, v)} />}</td>)}
                <td className="text-center"><DelBtn onClick={() => onChange(rows.filter((x) => x.id !== r.id))} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
const listSum = (rows, k, filter) => round2((rows || []).filter((r) => !filter || filter(r)).reduce((s, r) => s + num(r[k]), 0));
const tableBlock = (title, cols, rows, extra) => ({ type: "table", title, head: cols.map(([, h]) => h), body: (rows || []).map((r) => cols.map(([k, , t]) => (t === "num" || (t === undefined && k !== "year") ? num(r[k]) : r[k] || ""))), numCols: cols.map(([k, , t], i) => ((t === "num" || (t === undefined && k !== "year")) ? i : -1)).filter((i) => i >= 0), ...(extra || {}) });

// Itemized deduction lines (1701 Schedule 4 / 1702-RT Schedule I), from the Chart of Accounts ITR line mapping.
function deductionLines(c, fig, itemized) {
  let total = 0;
  ITR_LINES.forEach((l) => { const v = c.a(lineKey(l), itemized ? fig.lines[l] : 0); total += v; });
  c.set("d_18", total);
}
function deductionSection(title, itemized, faDep, c) {
  return { title, hide: !itemized, note: "From the Chart of Accounts ITR Deduction Line mapping. Accounts without a line count under 17d Others; Non-deductible accounts are excluded.", lines: [
    ...ITR_LINES.map((l) => ({ no: lineNo(l), label: l.slice(lineNo(l).length + 1) + (lineNo(l) === "5" && Math.abs(c.v.d_5 - faDep) >= 1 ? ` (Fixed Asset Ledger: ${fmtPlain(faDep)})` : ""), k: lineKey(l), t: "auto", src: `Expense accounts mapped to “${l}”` })),
    { no: "18", label: "Total Ordinary Allowable Itemized Deductions (sum of items 1 to 17)", k: "d_18", t: "calc", strong: true, f: (r) => ITR_LINES.map((l) => r(lineKey(l))).join("+") },
  ]};
}

function prevYearForm(data, store, year) { return (data[store] || {})[String(year - 1)] || null; }

const ATTACH_ANNUAL = [
  ["proofFiling", "Proof of filing (FRN or TRRC)"], ["proofPayment", "Proof of payment"],
  ["afs", "Audited Financial Statements and notes (Certificate of Independent CPA when gross sales exceed ₱3M — verify)", "income"],
  ["smr", "Statement of Management Responsibility"], ["f2307", "BIR Forms 2307 (Certificates of Creditable Tax Withheld at Source)", "form2307"],
  ["f2316", "BIR Form 2316 / 1606, where applicable"], ["sawt", "SAWT acknowledgement", "sawt"],
];

/* ------------------------------ 1701 ------------------------------ */

export function build1701(data, postings, year) {
  const c0 = data.company || {};
  const store = "form1701";
  const form = (data[store] || {})[String(year)] || { m: {}, overrides: {} };
  const fy = fiscalYearRange(year, c0.fiscalYearEndMonth || 12);
  const fig = incomeFigures(data, postings, fy.start, fy.end);
  // Election from the Q1 1701Q of the year, else Company Details; editable here with a warning.
  const q1 = (data.form1701q || {})[quarterKey(year, 1)];
  const dflt = q1 ? { rate: (q1.m || {}).taxRate || electionFromCompany(c0).rate, method: (q1.m || {}).method || electionFromCompany(c0).method } : electionFromCompany(c0);
  const c = makeCalc(form);
  const rate = c.t("taxRate", dflt.rate);
  const eight = rate === "8%";
  const osd = !eight && c.t("method", dflt.method) === "OSD";
  const mixed = c0.filerKind === "Mixed income";
  const comp = form.compRows || [];
  // Schedules 1 and 2: compensation (mixed-income earners)
  c.a("i4", mixed ? listSum(comp, "comp", (r) => r.owner !== "Spouse") : 0);
  c.man("i5");
  c.set("i6", c.v.i4 - c.v.i5);
  c.set("i7", graduatedTax(Math.max(0, c.v.i6), year));
  // Schedule 4 deductions, 5 special, 6 NOLCO
  deductionLines(c, fig, !eight && !osd);
  c.set("sp_total", listSum(form.special5, "amount", (r) => r.owner !== "Spouse"));
  c.set("nolco_total", eight || osd ? 0 : listSum(form.nolco, "appliedCurrent"));
  // Schedule 3.A graduated
  c.a("i8", fig.sales); c.a("i9", fig.returns);
  c.set("i10", c.v.i8 - c.v.i9);
  c.a("i11", osd ? 0 : fig.costOfSales);
  c.set("i12", c.v.i10 - c.v.i11);
  c.set("i13", osd ? 0 : c.v.d_18);
  c.set("i14", osd ? 0 : c.v.sp_total);
  c.set("i15", c.v.nolco_total);
  c.set("i16", c.v.i13 + c.v.i14 + c.v.i15);
  c.set("i17", osd ? c.v.i10 * 0.4 : 0);
  c.set("i18", osd ? c.v.i10 - c.v.i17 : c.v.i12 - c.v.i16);
  c.a("i19", fig.otherIncome); c.man("i20"); c.man("i21");
  c.set("i22", c.v.i19 + c.v.i20 + c.v.i21);
  c.set("i23", c.v.i18 + c.v.i22);
  c.set("i24", c.v.i6 + c.v.i23);
  c.set("i25", graduatedTax(Math.max(0, c.v.i24), year));
  // Schedule 3.B 8%
  c.a("i26", fig.netSales); c.a("i27", fig.otherIncome);
  c.set("i28", c.v.i26 + c.v.i27);
  c.set("i29", mixed ? 0 : 250000);
  c.set("i30", Math.max(0, c.v.i28 - c.v.i29));
  c.set("i31", c.v.i30 * 0.08);
  c.set("i32", c.v.i7 + c.v.i31);
  // Part VI
  c.set("p6_1", eight ? c.v.i32 : c.v.i25);
  c.man("p6_2"); c.man("p6_3");
  c.set("p6_4", c.v.p6_2 - c.v.p6_3);
  c.set("p6_5", c.v.p6_1 + c.v.p6_4);
  // Part VII credits
  const prev = prevYearForm(data, store, year);
  let prevExcess = 0;
  if (prev && /carried over/i.test((prev.m || {}).overpayChoice || "")) { const pc = build1701(data, postings, year - 1).calc; if (pc.v.i31a < 0) prevExcess = -pc.v.i31a; }
  c.a("p7_1", prevExcess);
  let qPaid = 0;
  for (let k = 1; k <= 3; k++) { const f = (data.form1701q || {})[quarterKey(year, k)]; if (f && f.filed) qPaid += paymentTotal(f) || Math.max(0, build1701q(data, postings, year, k).calc.v.i30); }
  c.a("p7_2", qPaid);
  const q3 = fiscalQuarterRange(year, c0.fiscalYearEndMonth || 12, 3);
  c.a("p7_3", cwtInRange(data, fy.start, q3.end));
  c.a("p7_4", cwtInRange(data, new Date(q3.end.getFullYear(), q3.end.getMonth(), q3.end.getDate() + 1), fy.end));
  c.a("p7_5", listSum(comp, "taxWithheld", (r) => r.owner !== "Spouse"));
  c.set("p7_6", c.t("amended") === "Yes" ? c.m("p7_6") : 0);
  c.man("p7_7"); c.man("p7_8"); c.man("p7_9");
  c.set("p7_10", c.v.p7_1 + c.v.p7_2 + c.v.p7_3 + c.v.p7_4 + c.v.p7_5 + c.v.p7_6 + c.v.p7_7 + c.v.p7_8 + c.v.p7_9);
  // Part II
  c.set("i22a", c.v.p6_5); c.set("i23a", c.v.p7_10);
  c.set("i24a", c.v.i22a - c.v.i23a);
  const inst = c.t("installment") === "Yes";
  c.set("i25a", inst ? Math.min(c.m("i25a"), Math.max(0, c.v.i22a) * 0.5) : 0);
  c.set("i26a", c.v.i24a - c.v.i25a);
  c.man("i27a"); c.man("i28a"); c.man("i29a");
  c.set("i30a", c.v.i27a + c.v.i28a + c.v.i29a);
  c.set("i31a", c.v.i26a + c.v.i30a);
  ["22", "23", "25", "27", "28", "29"].forEach((n) => c.man(`i${n}b`));
  c.set("i24b", c.v.i22b - c.v.i23b); c.set("i26b", c.v.i24b - c.v.i25b);
  c.set("i30b", c.v.i27b + c.v.i28b + c.v.i29b); c.set("i31b", c.v.i26b + c.v.i30b);
  c.set("i32a", c.v.i31a + c.v.i31b);
  // Part IX reconciliation
  c.a("r1", fig.netIncomeBooks); c.a("r2", fig.nonDeductibleTotal); c.man("r3");
  c.set("r4", c.v.r1 + c.v.r2 + c.v.r3);
  c.man("r6"); c.a("r7", fig.finalTaxIncome); c.set("r8", c.v.i14);
  c.a("r9", (osd ? c.v.i17 - (fig.costOfSales + fig.deductions) : 0) + c.v.i15 - c.v.i20 - c.v.i21);
  c.set("r10", c.v.r6 + c.v.r7 + c.v.r8 + c.v.r9);
  c.set("r11", c.v.r4 - c.v.r10);
  const nol = round2(Math.max(0, -(c.v.i12 - c.v.i13)));
  return { calc: c, form, fig, fy, eight, osd, mixed, dflt, nol, q1 };
}

function sections1701(b, info) {
  const { calc: c, eight, osd, mixed } = b;
  const amended = c.t("amended") === "Yes";
  const spouse = info.spouse;
  const cols = spouse ? ["A Taxpayer / Filer", "B Spouse"] : ["A Taxpayer / Filer"];
  const two = (n, ta = "calc", tb = "man") => (spouse ? { k: [`i${n}a`, `i${n}b`], t: [ta, tb] } : { k: `i${n}a`, t: ta });
  return {
    header: { title: "Header (items 1 to 3)", lines: [
      { no: "1", label: "For the Year", k: "_y", t: "ro", text: String(info.year) },
      { no: "2", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "3", label: "Short Period Return?", k: "shortPeriod", t: "yesno", dflt: "No" },
      { no: "", label: "Short period: from / to", k: "shortFromTo", t: "text", ph: "MM/DD/YYYY – MM/DD/YYYY", hide: c.t("shortPeriod") !== "Yes" },
    ]},
    election: { title: "Items 7, 21 and 21A", note: b.q1 ? "Inherited from the 1st quarter 1701Q. Change only with care: the election was made on the Q1 return." : "From Company Details (no 1st quarter 1701Q saved for this year). The election is made on the Q1 return.", lines: [
      { no: "7", label: "ATC (II011 to II017; confirm against your eBIRForms list)", k: "atc", t: "text", dflt: info.atc },
      { no: "21", label: "Tax Rate", k: "taxRate", t: "select", opts: ["Graduated", "8%"], dflt: b.dflt.rate },
      { no: "21A", label: "Method of Deduction", k: "method", t: "select", opts: ["Itemized", "OSD"], dflt: b.dflt.method === "N/A under 8%" ? "Itemized" : b.dflt.method, hide: eight },
      { no: "", label: "Pay in two installments? (tax due over ₱2,000; second installment due October 15 — confirm threshold)", k: "installment", t: "yesno", dflt: "No" },
    ]},
    part2: { title: "Part II – Total Tax Payable", cols, lines: [
      { no: "22", label: "Tax Due (from Part VI item 5)", ...two("22") },
      { no: "23", label: "Less: Total Tax Credits / Payments (from Part VII item 10)", ...two("23") },
      { no: "24", label: "Tax Payable / (Overpayment) (item 22 less item 23)", ...two("24", "calc", "calc"), due: true, f: spouse ? [(r) => `${r("i22a")}-${r("i23a")}`, (r) => `${r("i22b")}-${r("i23b")}`] : (r) => `${r("i22a")}-${r("i23a")}` },
      { no: "25", label: "Less: Portion of tax payable allowed for 2nd installment, due on or before October 15 (50% or less of item 22)", ...two("25", "man", "man"), dis: c.t("installment") !== "Yes" },
      { no: "26", label: "Amount of Tax Payable / (Overpayment) (item 24 less item 25)", ...two("26", "calc", "calc"), due: true },
      { heading: "Add: Penalties (manual)" },
      { no: "27", label: "Interest", ...two("27", "man", "man") },
      { no: "28", label: "Surcharge", ...two("28", "man", "man") },
      { no: "29", label: "Compromise", ...two("29", "man", "man") },
      { no: "30", label: "Total Penalties (sum of items 27 to 29)", ...two("30", "calc", "calc") },
      { no: "31", label: "Total Amount Payable / (Overpayment) (sum of items 26 and 30)", ...two("31", "calc", "calc"), strong: true, due: true },
      { no: "32", label: "Aggregate Amount Payable / (Overpayment) (taxpayer + spouse)", k: "i32a", t: "calc", strong: true, due: true, hide: !spouse },
      { no: "", label: "If overpayment, mark one (confirm wording)", k: "overpayChoice", t: "select", opts: ["To be refunded", "To be issued a Tax Credit Certificate", "To be carried over as tax credit for next year"], hide: !(c.v.i31a < 0) },
      { no: "33", label: "Number of Attachments", k: "attachments", t: "int", dflt: String(info.attachCount) },
    ]},
    sched2: { title: "Schedule 2 – Taxable Compensation Income (items 4 to 7)", hide: !mixed, lines: [
      { no: "4", label: "Gross Compensation Income (Schedule 1 total, taxpayer)", k: "i4", t: "auto", src: "Schedule 1 compensation rows marked Taxpayer" },
      { no: "5", label: "Less: Non-Taxable / Exempt Compensation", k: "i5", t: "man" },
      { no: "6", label: "Taxable Compensation Income (item 4 less item 5)", k: "i6", t: "calc", f: (r) => `${r("i4")}-${r("i5")}` },
      { no: "7", label: "Tax Due on Compensation Income (graduated table on item 6; used with the 8% option)", k: "i7", t: "calc" },
    ]},
    sched3a: { title: `Schedule 3.A – Taxable Business Income, Graduated Rates (items 8 to 25)`, hide: eight, lines: [
      { no: "8", label: "Sales / Revenues / Receipts / Fees (net of output VAT)", k: "i8", t: "auto", src: "Income Statement revenue accounts for the year" },
      { no: "9", label: "Less: Sales Returns, Allowances and Discounts", k: "i9", t: "auto", src: "Contra-revenue accounts (returns, discounts)" },
      { no: "10", label: "Net Sales / Revenues / Receipts / Fees (item 8 less item 9)", k: "i10", t: "calc", f: (r) => `${r("i8")}-${r("i9")}` },
      { no: "11", label: "Less: Cost of Sales / Services (itemized only)", k: "i11", t: "auto", src: "Income Statement cost of sales (with the inventory log)", dis: osd },
      { no: "12", label: "Gross Income / (Loss) from Operation (item 10 less item 11)", k: "i12", t: "calc", f: (r) => `${r("i10")}-${r("i11")}` },
      { no: "13", label: "Total Ordinary Allowable Itemized Deductions (Schedule 4 item 18)", k: "i13", t: "calc" },
      { no: "14", label: "Total Special Allowable Itemized Deductions (Schedule 5)", k: "i14", t: "calc" },
      { no: "15", label: "Allowance for NOLCO (Schedule 6 item 8)", k: "i15", t: "calc" },
      { no: "16", label: "Total Allowable Itemized Deductions (sum of items 13 to 15)", k: "i16", t: "calc", f: (r) => `${r("i13")}+${r("i14")}+${r("i15")}` },
      { no: "17", label: "Optional Standard Deduction (40% of item 10, net sales)", k: "i17", t: "calc" },
      { no: "18", label: "Net Income / (Loss) (itemized: item 12 less item 16; OSD: item 10 less item 17)", k: "i18", t: "calc", strong: true },
      { no: "19", label: "Other non-operating income (Other Income accounts not subject to final tax)", k: "i19", t: "auto", src: "Other Income accounts not flagged Subject to Final Tax" },
      { no: "20", label: "Other non-operating income (specify)", k: "i20", t: "man" },
      { no: "21", label: "Amount received / share in income by a partner from GPP", k: "i21", t: "man" },
      { no: "22", label: "Total Other Non-Operating Income (sum of items 19 to 21)", k: "i22", t: "calc", f: (r) => `${r("i19")}+${r("i20")}+${r("i21")}` },
      { no: "23", label: "Total Taxable Income – Business (sum of items 18 and 22)", k: "i23", t: "calc", strong: true },
      { no: "24", label: "Total Taxable Income – Compensation and Business (items 6 and 23)", k: "i24", t: "calc", strong: true, f: (r) => `${r("i6")}+${r("i23")}` },
      { no: "25", label: `Total Tax Due – Graduated (${taxTableFor(info.year).label}) → Part VI item 1`, k: "i25", t: "calc", strong: true },
    ]},
    sched3b: { title: "Schedule 3.B – 8% Income Tax Rate (items 26 to 32)", hide: !eight, lines: [
      { no: "26", label: "Sales / Revenues / Receipts / Fees (net of returns and discounts)", k: "i26", t: "auto", src: "Income Statement revenue accounts, net of contra-revenue" },
      { no: "27", label: "Add: Other Non-Operating Income", k: "i27", t: "auto", src: "Other Income accounts not flagged Subject to Final Tax" },
      { no: "28", label: "Total Income (sum of items 26 and 27)", k: "i28", t: "calc", f: (r) => `${r("i26")}+${r("i27")}` },
      { no: "29", label: `Less: Amount allowed as deduction under Sec. 24(A)(2)(b) – ₱250,000${mixed ? " (not applicable: taxpayer has compensation income)" : ""}`, k: "i29", t: "calc" },
      { no: "30", label: "Taxable Income / (Loss) (item 28 less item 29)", k: "i30", t: "calc", f: (r) => `MAX(${r("i28")}-${r("i29")},0)` },
      { no: "31", label: "Tax Due – Business Income (item 30 × 8%)", k: "i31", t: "calc", f: (r) => `${r("i30")}*0.08` },
      { no: "32", label: "Total Tax Due – Compensation and Business, Flat Rate (items 7 and 31) → Part VI item 1", k: "i32", t: "calc", strong: true, f: (r) => `${r("i7")}+${r("i31")}` },
    ]},
    sched4: deductionSection("Schedule 4 – Ordinary Allowable Itemized Deductions (items 1 to 18)", !eight && !osd, info.faDep, c),
    part6: { title: "Part VI – Summary of Income Tax Due", lines: [
      { no: "1", label: "Regular Income Tax Due (Schedule 3.A item 25 or Schedule 3.B item 32)", k: "p6_1", t: "calc" },
      { no: "2", label: "Special Rate Income Tax Due (Part X; manual in v1)", k: "p6_2", t: "man" },
      { no: "3", label: "Less: Share of Other Government Agency, if remitted directly", k: "p6_3", t: "man" },
      { no: "4", label: "Net Special Income Tax Due (item 2 less item 3)", k: "p6_4", t: "calc", f: (r) => `${r("p6_2")}-${r("p6_3")}` },
      { no: "5", label: "Total Income Tax Due (items 1 and 4) → Part II item 22", k: "p6_5", t: "calc", strong: true, f: (r) => `${r("p6_1")}+${r("p6_4")}` },
    ]},
    part7: { title: "Part VII – Tax Credits / Payments", lines: [
      { no: "1", label: "Prior Year's Excess Credits", k: "p7_1", t: "auto", src: "Last year's 1701 overpayment marked “carried over”" },
      { no: "2", label: "Tax Payments for the First Three Quarters", k: "p7_2", t: "auto", src: "Saved 1701Q returns marked Filed: amount paid" },
      { no: "3", label: "Creditable Tax Withheld for the First Three Quarters", k: "p7_3", t: "auto", src: "Creditable Withholding Taxes tab, Q1 to Q3" },
      { no: "4", label: "Creditable Tax Withheld per BIR Form No. 2307 for the 4th Quarter", k: "p7_4", t: "auto", src: "Creditable Withholding Taxes tab, Q4" },
      { no: "5", label: "Creditable Tax Withheld per BIR Form No. 2316 (Schedule 1)", k: "p7_5", t: "auto", src: "Schedule 1 tax withheld, taxpayer" },
      { no: "6", label: "Tax Paid in Return Previously Filed, if this is an amended return", k: "p7_6", t: "man", dis: !amended },
      { no: "7", label: "Foreign Tax Credits, if applicable", k: "p7_7", t: "man" },
      { no: "8", label: "Special Tax Credits", k: "p7_8", t: "man" },
      { no: "9", label: "Other Tax Credits / Payments (specify)", k: "p7_9", t: "man" },
      { no: "10", label: "Total Tax Credits / Payments (sum of items 1 to 9) → Part II item 23", k: "p7_10", t: "calc", strong: true, f: (r) => ["p7_1", "p7_2", "p7_3", "p7_4", "p7_5", "p7_6", "p7_7", "p7_8", "p7_9"].map(r).join("+") },
    ]},
    part9: { title: "Part IX – Reconciliation of Net Income per Books Against Taxable Income", lines: [
      { no: "1", label: "Net Income / (Loss) per Books", k: "r1", t: "auto", src: "Income Statement accounts for the year" },
      { no: "2", label: "Add: Non-deductible expenses / taxable other income (Non-deductible accounts)", k: "r2", t: "auto", src: "Expense accounts mapped Non-deductible" },
      { no: "3", label: "Add: Other taxable income not in the books (specify)", k: "r3", t: "man" },
      { no: "4", label: "Total (sum of items 1 to 3)", k: "r4", t: "calc" },
      { no: "6", label: "Less: Non-taxable income", k: "r6", t: "man" },
      { no: "7", label: "Less: Income subjected to final tax", k: "r7", t: "auto", src: "Income accounts flagged Subject to Final Tax" },
      { no: "8", label: "Less: Special deductions (Schedule 5)", k: "r8", t: "calc" },
      { no: "9", label: "Less: Other adjustments (OSD in lieu of actual cost and expenses, NOLCO, GPP / other income lines)", k: "r9", t: "auto", src: "OSD less actual cost of sales and deductible expenses, plus NOLCO, less items 20 and 21" },
      { no: "10", label: "Total (sum of items 6 to 9)", k: "r10", t: "calc" },
      { no: "11", label: "Net Taxable Income / (Loss) (item 4 less item 10)", k: "r11", t: "calc", strong: true },
    ]},
  };
}

const bg1701 = (i, c) => [["4", "TIN", i.tinText], ["5", "RDO Code", i.rdo], ["6", "Taxpayer Type", c.filerKind === "Professional" ? "Professional" : c.filerKind === "Mixed income" ? "Single Proprietor / Compensation Earner" : "Single Proprietor"],
  ["8", "Taxpayer's Name", i.name], ["9", "Registered Address", i.address], ["9A", "ZIP Code", i.zip], ["10", "Date of Birth", c.dateOfBirth], ["11", "Email Address", i.email],
  ["12", "Citizenship", c.citizenship], ["13", "Claiming Foreign Tax Credits?", c.claimsForeignTaxCredit || "No"], ["14", "Foreign Tax Number", c.foreignTaxNumber || "N/A"], ["15", "Contact Number", i.contact],
  ["16", "Civil Status", c.civilStatus], ["17", "If married, does the spouse have income?", c.civilStatus === "Married" ? (c.spouseHasIncome || "No") : "N/A"], ["18", "Filing Status", c.civilStatus === "Married" && c.spouseHasIncome === "Yes" ? "Joint" : "N/A"],
  ["19", "Income exempt from income tax?", c.incomeExempt || "No"], ["20", "Income subject to special / preferential rate?", c.incomeSpecialRate || "No"]];

const COMP_COLS = [["employer", "Employer Name", "text"], ["employerTin", "Employer TIN", "text"], ["comp", "Compensation Income", "num"], ["taxWithheld", "Tax Withheld", "num"], ["owner", "Taxpayer / Spouse", "select", ["Taxpayer", "Spouse"]]];
const SPECIAL_COLS = [["desc", "Description", "text"], ["basis", "Legal Basis", "text"], ["amount", "Amount", "num"], ["owner", "Taxpayer / Spouse", "select", ["Taxpayer", "Spouse"]]];
const NOLCO_COLS = [["year", "Year Incurred"], ["nol", "A Net Operating Loss"], ["appliedPrev", "B NOLCO Applied Previous Years"], ["expired", "C NOLCO Expired"], ["appliedCurrent", "D NOLCO Applied Current Year"]];

function Screen1701({ data, setData, postings, onNavigate, year, selectors }) {
  const info0 = companyInfo(data);
  const c0 = data.company || {};
  const b = useMemo(() => build1701(data, postings, year), [data, postings, year]);
  const prev = prevYearForm(data, "form1701", year);
  const prevB = prev ? build1701(data, postings, year - 1) : null;
  const rt = useReturn(data, setData, "form1701", String(year), { nolco: prevB ? prefillNolco(prev, prevB.nol, year - 1) : [], compRows: [], special5: [], qSnapshot: b.calc.auto.p7_2 });
  const calc = b.calc;
  const spouse = c0.civilStatus === "Married" && c0.spouseHasIncome === "Yes";
  const faDep = faDepreciationForYear(data, year);
  const info = { year, spouse, atc: c0.filerKind === "Professional" ? (b.eight ? "II017" : "II012") : (b.eight ? "II017" : (b.mixed ? "II013" : "II011")), attachCount: attachCount(rt.form, ATTACH_ANNUAL), faDep };
  const secs = sections1701(b, info);
  const form = rt.form;
  const isVat = isVatCompany(data);
  const vs = useMemo(() => (isVat ? computeVatSchedule(data, year, aggregateVatByYear(data)) : null), [data, year, isVat]);
  const qChanged = rt.exists && form.qSnapshot !== undefined && Math.abs(num(form.qSnapshot) - calc.auto.p7_2) >= 0.005;
  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1701 — Annual Income Tax Return for Individuals (including MIXED Income Earner), Estates and Trusts", version: "January 2018 (ENCS)", formNo: "1701",
    period: `For the year ${year}`, company: info0,
    blocks: [sectionBlock(secs.header, calc, form), infoBlock("Part I – Background Information", bg1701(info0, c0)), sectionBlock(secs.election, calc, form), sectionBlock(secs.part2, calc, form),
      paymentBlock(form, "Part III – Details of Payment (items 34 to 37)"),
      b.mixed ? tableBlock("Schedule 1 – Compensation Income and Tax Withheld", COMP_COLS, form.compRows) : null,
      sectionBlock(secs.sched2, calc, form), sectionBlock(secs.sched3a, calc, form), sectionBlock(secs.sched3b, calc, form), sectionBlock(secs.sched4, calc, form),
      (form.special5 || []).length ? tableBlock("Schedule 5 – Special Allowable Itemized Deductions", SPECIAL_COLS, form.special5) : null,
      (form.nolco || []).length ? tableBlock("Schedule 6 – NOLCO", [...NOLCO_COLS], form.nolco) : null,
      sectionBlock(secs.part6, calc, form), sectionBlock(secs.part7, calc, form), sectionBlock(secs.part9, calc, form), annualQuarterBlock(data, postings, "1701", year)],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1701_${year}_worksheet.pdf`,
  }));
  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}>
      <button className="io-btn" onClick={() => onNavigate("form2307")}>Open Creditable Withholding Taxes</button>
      <button className="io-btn" onClick={() => onNavigate("sawt")}>Open SAWT</button>
      <button className="io-btn" onClick={() => onNavigate("income")}>Open Income Statement</button>
      <WorksheetButtons hook={exp} />
    </div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={[["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"], ["dateOfBirth", "Date of Birth"]]} />
    <ReturnStatusBar rt={rt} dueDefault={dueAnnual(b.fy)} dueNote="April 15 (calendar year; RMC 20-2026 for TY 2025) · verify">{selectors}</ReturnStatusBar>
    <Notice tone="info">Also exist (not built): 1701A (business / professional income only, OSD or 8%) and 1701-MS (micro and small taxpayers, manual filing). Under RMC 20-2026 micro and small individual filers may use 1701, 1701A or 1701-MS; this worksheet is labelled for the 1701.</Notice>
    {b.eight && calc.v.i28 > EIGHT_PCT_LIMIT && <Notice tone="danger">Gross receipts plus non-operating income exceed ₱3,000,000. The 8% option no longer applies. Review with the accountant; the election was not changed.</Notice>}
    {qChanged && <Notice>A quarterly 1701Q changed since this return was prepared (quarterly payments now {fmtPlain(calc.auto.p7_2)}, were {fmtPlain(form.qSnapshot)}). Typed values are kept. <button className="tx-link" onClick={() => rt.patch({ qSnapshot: calc.auto.p7_2 })}>Acknowledge</button></Notice>}
    {calc.t("amended") === "Yes" && <Notice>Amended return: Part VII item 6 is unlocked. Key in the corrected figures.</Notice>}
    <div className="tx-note" style={{ padding: "0 0 8px" }}>{isVat ? "Sales exclude 12% output VAT." : "Non-VAT: sales are the gross receipts as recorded."}</div>
    <FormSection section={secs.header} calc={calc} rt={rt} />
    <BackgroundBlock items={bg1701(info0, c0)} onNavigate={onNavigate} />
    <FormSection section={secs.election} calc={calc} rt={rt} />
    {spouse && <SpouseBg rt={rt} c={c0} />}
    {b.mixed && <ListTable title="Schedule 1 – Compensation Income and Tax Withheld (items 1 to 3B)" rows={form.compRows || []} onChange={(rows) => rt.patch({ compRows: rows })} cols={COMP_COLS} note="One row per employer (BIR Form 2316). Taxpayer rows feed Schedule 2 item 4 and Part VII item 5." />}
    <FormSection section={secs.sched2} calc={calc} rt={rt} />
    <FormSection section={secs.sched3a} calc={calc} rt={rt} />
    <FormSection section={secs.sched3b} calc={calc} rt={rt} />
    <FormSection section={secs.sched4} calc={calc} rt={rt} />
    {!b.eight && !b.osd && <UnmappedLink fig={b.fig} onNavigate={onNavigate} />}
    {!b.eight && !b.osd && <ListTable title="Schedule 5 – Special Allowable Itemized Deductions (items 1 to 6)" rows={form.special5 || []} onChange={(rows) => rt.patch({ special5: rows })} cols={SPECIAL_COLS} />}
    {!b.eight && !b.osd && <CarryTable title="Schedule 6 – Computation of NOLCO (items 4 to 8)" rows={form.nolco || []} onChange={(rows) => rt.patch({ nolco: rows })} cols={NOLCO_COLS}
      computedCol={["E Unapplied (A − (B + C + D))", nolcoE]} totalCol="appliedCurrent" totalLabel="8 Total NOLCO applied (column D) → Schedule 3.A item 15"
      note={`6.A: gross income ${fmtPlain(calc.v.i12)} less ordinary deductions ${fmtPlain(calc.v.i13)} = net operating loss ${fmtPlain(b.nol)}. Pre-filled from last year's saved return. Losses for 2020 and 2021 carry over 5 years (Bayanihan Act), others 3 years; type the Expired column (nothing expires automatically).`} />}
    <FormSection section={secs.part6} calc={calc} rt={rt} />
    <FormSection section={secs.part7} calc={calc} rt={rt} />
    <FormSection section={secs.part2} calc={calc} rt={rt} />
    <PaymentDetails rt={rt} title="Part III – Details of Payment (items 34 to 37)" firstItem={34} amountDue={calc.v.i31a} />
    {(c0.incomeExempt === "Yes" || c0.incomeSpecialRate === "Yes") && <ManualPart rt={rt} title="Part VIII – Tax Relief Availment (manual)" />}
    <FormSection section={secs.part9} calc={calc} rt={rt} />
    <AnnualRecon b={b} data={data} vs={vs} salesKey={b.eight ? "i26" : "i10"} dedKey={!b.eight && !b.osd ? "i13" : null} cwtKeys={["p7_3", "p7_4"]} qKey="p7_2" qDue={annualQuarterDue(data, postings, "1701", year)} recon={[calc.v.r11, b.eight ? calc.v.i30 : calc.v.i23]} faDep={faDep} />
    <AttachmentChecklist rt={rt} items={ATTACH_ANNUAL} title="Attachment checklist (feeds item 33; filed through eAFS within 15 days per RMC 20-2026)" onOpen={onNavigate} />
    <Declaration data={data} rt={rt} />
  </>);
}

function SpouseBg({ rt, c }) {
  const [open, setOpen] = useState(false);
  const m = rt.form.m || {};
  const F = (k, label, dflt = "") => <LabeledField label={label}><Field value={m[k] ?? dflt} onChange={(v) => rt.setM(k, v)} /></LabeledField>;
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>Part IV – Background Information of Spouse (items 1 to 12A)</span><button className="tx-link" onClick={() => setOpen(!open)}>{open ? "Collapse" : "Expand"}</button></div>
      {open && <div className="tx-inline-fields">
        {F("spTin", "1 Spouse TIN", c.spouseTin)}{F("spRdo", "2 RDO Code")}{F("spType", "3 Taxpayer Type")}{F("spAtc", "4 ATC")}{F("spName", "5 Spouse's Name", c.spouseName)}
        {F("spContact", "6 Contact Number")}{F("spCitizenship", "7 Citizenship")}{F("spForeignCredit", "8 Claiming Foreign Tax Credits?")}{F("spForeignTax", "9 Foreign Tax Number")}
        {F("spExempt", "10 Income exempt?")}{F("spSpecial", "11 Income at special rate?")}{F("spRate", "12 Tax Rate")}{F("spMethod", "12A Method of Deduction")}
      </div>}
    </div>
  );
}
function ManualPart({ rt, title }) {
  const m = rt.form.m || {};
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>{title}</span></div>
      <div className="tx-inline-fields">
        <LabeledField label="Description / legal basis" wide><Field value={m.reliefDesc} onChange={(v) => rt.setM("reliefDesc", v)} /></LabeledField>
        <LabeledField label="Amount"><NumberField value={m.reliefAmt} onChange={(v) => rt.setM("reliefAmt", v)} align="right" /></LabeledField>
      </div>
    </div>
  );
}
function UnmappedLink({ fig, onNavigate }) {
  if (!fig.unmapped.length && !fig.nonDeductible.length) return null;
  return (
    <div className="tx-note" style={{ padding: "0 0 10px" }}>
      {fig.nonDeductible.length > 0 && <>Non-deductible: {fmtPlain(fig.nonDeductibleTotal)} excluded. </>}
      {fig.unmapped.length > 0 && <span className="tx-amber">{fig.unmapped.length} unmapped account{fig.unmapped.length === 1 ? "" : "s"} ({fig.unmapped.map((x) => `${x.code} ${x.name}`).join(", ")}) rolled into 17 Others. </span>}
      <button className="tx-link" onClick={() => onNavigate("coa")}>Map in Chart of Accounts</button>
    </div>
  );
}

function annualQuarterDue(data, postings, kind, year) {
  let t = 0;
  for (let k = 1; k <= 3; k++) {
    const store = kind === "1701" ? "form1701q" : "form1702q";
    if (!(data[store] || {})[quarterKey(year, k)]) continue;
    const bq = kind === "1701" ? build1701q(data, postings, year, k) : build1702q(data, postings, year, k);
    t += Math.max(0, bq.calc.v[kind === "1701" ? "i30" : "i24"]);
  }
  return round2(t);
}
function annualQuarterBlock(data, postings, kind, year) {
  const store = kind === "1701" ? "form1701q" : "form1702q";
  const body = [1, 2, 3].map((k) => {
    const f = (data[store] || {})[quarterKey(year, k)];
    if (!f) return [`Q${k}`, "Not saved", 0, 0];
    const bq = kind === "1701" ? build1701q(data, postings, year, k) : build1702q(data, postings, year, k);
    const ti = kind === "1701" ? (bq.eight ? bq.calc.v.i51 : bq.calc.v.i46) : bq.calc.v.s2_10;
    return [`Q${k}`, f.filed ? `Filed ${f.dateFiled || ""}` : "Draft", ti, f.filed ? (paymentTotal(f) || Math.max(0, bq.calc.v[kind === "1701" ? "i30" : "i24"])) : 0];
  });
  return { type: "table", title: "Quarterly summary (Q1 to Q3 taxable income to date and payments)", head: ["Quarter", "Status", "Taxable income to date", "Payment"], body, numCols: [2, 3] };
}

function AnnualRecon({ b, vs, salesKey, dedKey, cwtKeys, qKey, qDue, recon, faDep }) {
  const { calc, fig } = b;
  return (
    <ReconStrip title="Books check" colA="Return" colB="Books / source" items={[
      { label: "Net sales vs Income Statement net sales", a: calc.v[salesKey], b: fig.netSales },
      vs ? { label: "Net sales vs VAT Schedule sales tax base for the year", a: calc.v[salesKey], b: round2(vs.total.totalSales) } : null,
      dedKey ? { label: `Itemized deductions vs Income Statement expenses less non-deductible${fig.unmapped.length ? ` (${fig.unmapped.length} unmapped)` : ""}`, a: calc.v[dedKey], b: round2(fig.totalExpenses - fig.nonDeductibleTotal) } : null,
      { label: "Creditable tax (Q1–Q3 + Q4) vs Creditable Withholding Taxes tab for the year", a: round2(calc.v[cwtKeys[0]] + calc.v[cwtKeys[1]]), b: round2(calc.auto[cwtKeys[0]] + calc.auto[cwtKeys[1]]) },
      { label: "Quarterly payments vs saved quarterly returns' amounts payable", a: calc.v[qKey], b: qDue, info: "returns not marked Filed are not counted as paid" },
      { label: "Net income reconciliation vs computed taxable income", a: recon[0], b: recon[1] },
      dedKey ? { label: "Depreciation (line 5) vs Fixed Asset Ledger", a: calc.v.d_5, b: faDep } : null,
    ]} causes="unmapped expense accounts, 2307s dated outside the year, manual General Journal entries, quarterly returns not marked Filed, or overrides." />
  );
}

/* ------------------------------ 1702-RT ------------------------------ */

export function build1702(data, postings, year) {
  const c0 = data.company || {};
  const store = "form1702";
  const form = (data[store] || {})[String(year)] || { m: {}, overrides: {} };
  const fy = fiscalYearRange(year, c0.fiscalYearEndMonth || 12);
  const fig = incomeFigures(data, postings, fy.start, fy.end);
  const q1 = (data.form1702q || {})[quarterKey(year, 1)];
  const c = makeCalc(form);
  const method = c.t("method", (q1 && (q1.m || {}).method) || c0.deductionMethod || "Itemized");
  const osd = method === "OSD";
  deductionLines(c, fig, !osd);
  c.set("sp_total", listSum(form.special2, "amount"));
  c.set("nolco_total", osd ? 0 : listSum(form.nolco, "appliedCurrent"));
  c.set("mcit_applied", listSum(form.mcitCarry, "appliedCurrent"));
  c.a("i27", fig.sales); c.a("i28", fig.returns);
  c.set("i29", c.v.i27 - c.v.i28);
  c.a("i30", fig.costOfSales);
  c.set("i31", c.v.i29 - c.v.i30);
  c.a("i32", fig.otherIncome);
  c.set("i33", c.v.i31 + c.v.i32);
  c.set("i34", osd ? 0 : c.v.d_18);
  c.set("i35", osd ? 0 : c.v.sp_total);
  c.set("i36", c.v.nolco_total);
  c.set("i37", c.v.i34 + c.v.i35 + c.v.i36);
  c.set("i38", osd ? c.v.i33 * 0.4 : 0);
  c.set("i39", osd ? c.v.i33 - c.v.i38 : c.v.i33 - c.v.i37);
  const rate = form.rate === undefined || form.rate === "" ? corpRate(c0) : num(form.rate);
  c.rate = rate;
  c.set("i41", Math.max(0, c.v.i39) * rate);
  const mc = mcitApplies(c0, year);
  c.set("i42", mc.applies ? Math.max(0, c.v.i33) * 0.02 : 0);
  c.set("i43", Math.max(c.v.i41, c.v.i42));
  const prev = prevYearForm(data, store, year);
  let prevExcess = 0;
  if (prev && /carried over/i.test((prev.m || {}).overpayChoice || "")) { const pc = build1702(data, postings, year - 1).calc; if (pc.v.i21 < 0) prevExcess = -pc.v.i21; }
  c.a("i44", prevExcess);
  let qMcit = 0, qReg = 0;
  for (let k = 1; k <= 3; k++) {
    const f = (data.form1702q || {})[quarterKey(year, k)];
    if (!f || !f.filed) continue;
    const bq = build1702q(data, postings, year, k).calc;
    const paid = paymentTotal(f) || Math.max(0, bq.v.i24);
    if (bq.v.s3_5 > bq.v.s2_12) qMcit += paid; else qReg += paid;
  }
  c.a("i45", qMcit); c.a("i46", qReg);
  c.set("i47", c.v.mcit_applied);
  const q3 = fiscalQuarterRange(year, c0.fiscalYearEndMonth || 12, 3);
  c.a("i48", cwtInRange(data, fy.start, q3.end));
  c.a("i49", cwtInRange(data, new Date(q3.end.getFullYear(), q3.end.getMonth(), q3.end.getDate() + 1), fy.end));
  c.man("i50");
  c.set("i51", c.t("amended") === "Yes" ? c.m("i51") : 0);
  c.man("i52"); c.man("i53"); c.man("i54");
  c.set("i55", ["i44", "i45", "i46", "i47", "i48", "i49", "i50", "i51", "i52", "i53", "i54"].reduce((s, k) => s + c.v[k], 0));
  c.set("i56", c.v.i43 - c.v.i55);
  // Part II
  c.set("i14", c.v.i43); c.set("i15", c.v.i55); c.set("i16", c.v.i14 - c.v.i15);
  c.man("i17"); c.man("i18"); c.man("i19");
  c.set("i20", c.v.i17 + c.v.i18 + c.v.i19);
  c.set("i21", c.v.i16 + c.v.i20);
  // Part V
  c.set("i57", c.v.i35 * rate); c.set("i58", c.v.i52); c.set("i59", c.v.i57 + c.v.i58);
  // Schedule III
  c.set("s3_1", c.v.i33); c.set("s3_2", osd ? 0 : c.v.d_18); c.set("s3_3", c.v.s3_1 - c.v.s3_2);
  // Schedule V
  c.a("r1", fig.netIncomeBooks); c.a("r2", fig.nonDeductibleTotal); c.man("r3");
  c.set("r4", c.v.r1 + c.v.r2 + c.v.r3);
  c.man("r5"); c.a("r6", fig.finalTaxIncome); c.set("r7", c.v.i35);
  c.a("r8", (osd ? c.v.i38 - (fig.costOfSales + fig.deductions) : 0) + c.v.i36);
  c.set("r9", c.v.r5 + c.v.r6 + c.v.r7 + c.v.r8);
  c.set("r10", c.v.r4 - c.v.r9);
  return { calc: c, form, fig, fy, osd, method, mc, q1, nol: round2(Math.max(0, -c.v.s3_3)) };
}

function sections1702(b, info) {
  const { calc: c, osd, mc } = b;
  const amended = c.t("amended") === "Yes";
  return {
    header: { title: "Header (items 1 to 5)", lines: [
      { no: "1", label: "For the", k: "_cal", t: "ro", text: info.fiscal ? "Fiscal" : "Calendar" },
      { no: "2", label: "Year Ended (MM/YYYY)", k: "_ye", t: "ro", text: info.yearEnded },
      { no: "3", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "4", label: "Short Period Return?", k: "shortPeriod", t: "yesno", dflt: "No" },
      { no: "", label: "Short period: from / to", k: "shortFromTo", t: "text", ph: "MM/DD/YYYY – MM/DD/YYYY", hide: c.t("shortPeriod") !== "Yes" },
      { no: "5", label: "ATC (regular rate IC010; MCIT IC055 — confirm in your eBIRForms)", k: "atc", t: "text", dflt: c.v.i42 > c.v.i41 ? "IC055" : "IC010" },
      { no: "13", label: "Method of Deductions", k: "method", t: "select", opts: ["Itemized", "OSD"], dflt: b.method, note: b.q1 ? "Inherited from the 1st quarter 1702Q." : "From Company Details." },
    ]},
    part2: { title: "Part II – Total Tax Payable", lines: [
      { no: "14", label: "Tax Due (from Part IV item 43)", k: "i14", t: "calc" },
      { no: "15", label: "Less: Total Tax Credits / Payments (from Part IV item 55)", k: "i15", t: "calc" },
      { no: "16", label: "Net Tax Payable / (Overpayment) (item 14 less item 15)", k: "i16", t: "calc", due: true, f: (r) => `${r("i14")}-${r("i15")}` },
      { heading: "Add: Penalties (manual)" },
      { no: "17", label: "Surcharge", k: "i17", t: "man" },
      { no: "18", label: "Interest", k: "i18", t: "man" },
      { no: "19", label: "Compromise", k: "i19", t: "man" },
      { no: "20", label: "Total Penalties (sum of items 17 to 19)", k: "i20", t: "calc", f: (r) => `${r("i17")}+${r("i18")}+${r("i19")}` },
      { no: "21", label: "TOTAL AMOUNT PAYABLE / (Overpayment) (sum of items 16 and 20)", k: "i21", t: "calc", strong: true, due: true, f: (r) => `${r("i16")}+${r("i20")}` },
      { no: "", label: "If overpayment, mark one", k: "overpayChoice", t: "select", opts: ["To be refunded", "To be issued a Tax Credit Certificate", "To be carried over as tax credit for next year"], hide: !(c.v.i21 < 0) },
      { no: "22", label: "Number of Attachments", k: "attachments", t: "int", dflt: String(info.attachCount) },
    ]},
    part4: { title: "Part IV – Computation of Tax (items 27 to 56)", lines: [
      { no: "27", label: "Sales / Receipts / Revenues / Fees (net of output VAT)", k: "i27", t: "auto", src: "Income Statement revenue accounts for the year" },
      { no: "28", label: "Less: Sales Returns, Allowances and Discounts", k: "i28", t: "auto", src: "Contra-revenue accounts" },
      { no: "29", label: "Net Sales / Receipts / Revenues / Fees (item 27 less item 28)", k: "i29", t: "calc", f: (r) => `${r("i27")}-${r("i28")}` },
      { no: "30", label: "Less: Cost of Sales / Services", k: "i30", t: "auto", src: "Income Statement cost of sales" },
      { no: "31", label: "Gross Income from Operation (item 29 less item 30)", k: "i31", t: "calc", f: (r) => `${r("i29")}-${r("i30")}` },
      { no: "32", label: "Add: Other Taxable Income Not Subjected to Final Tax", k: "i32", t: "auto", src: "Other Income accounts not flagged Subject to Final Tax" },
      { no: "33", label: "Total Taxable Income (sum of items 31 and 32)", k: "i33", t: "calc", strong: true, f: (r) => `${r("i31")}+${r("i32")}` },
      { heading: "Less: Deductions" },
      { no: "34", label: "Ordinary Allowable Itemized Deductions (Schedule I item 18)", k: "i34", t: "calc" },
      { no: "35", label: "Special Allowable Itemized Deductions (Schedule II item 5)", k: "i35", t: "calc" },
      { no: "36", label: "NOLCO (only for those taxable under Sec. 27(A to C), Sec. 28(A)(1) and (A)(6)(b)) (Schedule IIIA item 8)", k: "i36", t: "calc" },
      { no: "37", label: "Total Deductions (sum of items 34 to 36)", k: "i37", t: "calc", f: (r) => `${r("i34")}+${r("i35")}+${r("i36")}` },
      { no: "38", label: "Optional Standard Deduction (40% of item 33)", k: "i38", t: "calc" },
      { no: "39", label: `Net Taxable Income / (Loss) (${osd ? "item 33 less item 38" : "item 33 less item 37"})`, k: "i39", t: "calc", strong: true },
      { no: "40", label: "Applicable Income Tax Rate", k: "_rate", t: "ro", text: `${round2(c.rate * 100)}%` },
      { no: "41", label: "Income Tax Due other than MCIT (item 39 × item 40)", k: "i41", t: "calc" },
      { no: "42", label: `Minimum Corporate Income Tax (MCIT) (2% of item 33)${mc.applies ? "" : ` — not applicable until ${mc.firstYear}`}`, k: "i42", t: "calc" },
      { no: "43", label: "Tax Due (the higher of item 41 and item 42) → Part II item 14", k: "i43", t: "calc", strong: true, f: (r) => `MAX(${r("i41")},${r("i42")})` },
      { heading: "Less: Tax Credits / Payments" },
      { no: "44", label: "Prior Year's Excess Credits other than MCIT", k: "i44", t: "auto", src: "Last year's 1702-RT overpayment marked “carried over”" },
      { no: "45", label: "Income Tax Payment under MCIT from Previous Quarter/s", k: "i45", t: "auto", src: "1702Q returns marked Filed where MCIT was higher" },
      { no: "46", label: "Income Tax Payment under Regular / Normal Rate from Previous Quarter/s", k: "i46", t: "auto", src: "1702Q returns marked Filed where the regular rate applied" },
      { no: "47", label: "Excess MCIT Applied this Current Taxable Year (Schedule IV item 4)", k: "i47", t: "calc" },
      { no: "48", label: "Creditable Tax Withheld from Previous Quarter/s per BIR Form No. 2307", k: "i48", t: "auto", src: "Creditable Withholding Taxes tab, Q1 to Q3" },
      { no: "49", label: "Creditable Tax Withheld per BIR Form No. 2307 for the 4th Quarter", k: "i49", t: "auto", src: "Creditable Withholding Taxes tab, Q4" },
      { no: "50", label: "Foreign Tax Credits, if applicable", k: "i50", t: "man" },
      { no: "51", label: "Tax Paid in Return Previously Filed, if this is an amended return", k: "i51", t: "man", dis: !amended },
      { no: "52", label: "Special Tax Credits (Part V item 58)", k: "i52", t: "man" },
      { no: "53", label: "Other Tax Credits / Payments (specify)", k: "i53", t: "man" },
      { no: "54", label: "Other Tax Credits / Payments (specify)", k: "i54", t: "man" },
      { no: "55", label: "Total Tax Credits / Payments (sum of items 44 to 54) → Part II item 15", k: "i55", t: "calc", strong: true, f: (r) => ["i44", "i45", "i46", "i47", "i48", "i49", "i50", "i51", "i52", "i53", "i54"].map(r).join("+") },
      { no: "56", label: "Net Tax Payable / (Overpayment) (item 43 less item 55)", k: "i56", t: "calc", strong: true, due: true, f: (r) => `${r("i43")}-${r("i55")}` },
    ]},
    part5: { title: "Part V – Tax Relief Availment", hide: !(c.v.i35 || c.v.i52), lines: [
      { no: "57", label: "Special Allowable Itemized Deductions (item 35 × applicable rate)", k: "i57", t: "calc" },
      { no: "58", label: "Add: Special Tax Credits (item 52)", k: "i58", t: "calc" },
      { no: "59", label: "Total Tax Relief Availment (sum of items 57 and 58)", k: "i59", t: "calc", strong: true },
    ]},
    schedI: deductionSection("Schedule I – Ordinary Allowable Itemized Deductions (items 1 to 18)", !osd, info.faDep, c),
    schedIII: { title: "Schedule III – Computation of Net Operating Loss (items 1 to 3)", hide: osd, lines: [
      { no: "1", label: "Gross Income (Part IV item 33)", k: "s3_1", t: "calc" },
      { no: "2", label: "Less: Ordinary Allowable Itemized Deductions (Schedule I item 18)", k: "s3_2", t: "calc" },
      { no: "3", label: "Net Operating Loss (item 1 less item 2)", k: "s3_3", t: "calc", strong: true },
    ]},
    schedV: { title: "Schedule V – Reconciliation of Net Income per Books Against Taxable Income (items 1 to 10)", lines: [
      { no: "1", label: "Net Income / (Loss) per Books", k: "r1", t: "auto", src: "Income Statement accounts for the year" },
      { no: "2", label: "Add: Non-deductible expenses / taxable other income", k: "r2", t: "auto", src: "Expense accounts mapped Non-deductible" },
      { no: "3", label: "Add: Other taxable income not in the books (specify)", k: "r3", t: "man" },
      { no: "4", label: "Total (sum of items 1 to 3)", k: "r4", t: "calc" },
      { no: "5", label: "Less: Non-taxable income", k: "r5", t: "man" },
      { no: "6", label: "Less: Income subjected to final tax", k: "r6", t: "auto", src: "Income accounts flagged Subject to Final Tax" },
      { no: "7", label: "Less: Special deductions (Schedule II)", k: "r7", t: "calc" },
      { no: "8", label: "Less: Other deductions / adjustments (OSD in lieu of actual cost and expenses, NOLCO)", k: "r8", t: "auto", src: "OSD less actual cost of sales and deductible expenses, plus NOLCO applied" },
      { no: "9", label: "Total (sum of items 5 to 8)", k: "r9", t: "calc" },
      { no: "10", label: "Net Taxable Income / (Loss) (item 4 less item 9) — should equal Part IV item 39", k: "r10", t: "calc", strong: true },
    ]},
  };
}
const bg1702 = (i, c, method) => [["6", "TIN", i.tinText], ["7", "RDO Code", i.rdo], ["8", "Registered Name", i.name], ["9", "Registered Address", i.address], ["9A", "ZIP Code", i.zip],
  ["10", "Date of Incorporation / Organization", c.dateIncorporation], ["11", "Contact Number", i.contact], ["12", "Email Address", i.email],
  ["13", "Method of Deductions", method === "OSD" ? "Optional Standard Deduction (OSD), Sec. 34(L)" : "Itemized Deductions, Sec. 34(A-J)"]];
const MCIT_COLS = [["year", "Year"], ["normal", "A Normal Income Tax as Adjusted"], ["mcit", "B MCIT"], ["appliedPrev", "D Excess MCIT Applied Previous Years"], ["expired", "E Expired"], ["appliedCurrent", "F Applied this Year"]];
const SPECIAL2_COLS = [["desc", "Description", "text"], ["basis", "Legal Basis", "text"], ["amount", "Amount", "num"]];

function Screen1702({ data, setData, postings, onNavigate, year, selectors }) {
  const info0 = companyInfo(data);
  const c0 = data.company || {};
  const b = useMemo(() => build1702(data, postings, year), [data, postings, year]);
  const prev = prevYearForm(data, "form1702", year);
  const prevB = prev ? build1702(data, postings, year - 1) : null;
  const rt = useReturn(data, setData, "form1702", String(year), {
    nolco: prevB ? prefillNolco(prev, prevB.nol, year - 1) : [], special2: [],
    mcitCarry: prevB ? prefillMcit(prev, prevB.calc.v.i41, prevB.calc.v.i42, year - 1) : [], qSnapshot: round2(b.calc.auto.i45 + b.calc.auto.i46),
  });
  const calc = b.calc;
  const fyEnd = Number(c0.fiscalYearEndMonth) || 12;
  const faDep = faDepreciationForYear(data, year);
  const info = { year, fiscal: fyEnd !== 12, yearEnded: `${String(fyEnd).padStart(2, "0")}/${year}`, attachCount: attachCount(rt.form, ATTACH_ANNUAL), faDep };
  const secs = sections1702(b, info);
  const form = rt.form;
  const isVat = isVatCompany(data);
  const vs = useMemo(() => (isVat && fyEnd === 12 ? computeVatSchedule(data, year, aggregateVatByYear(data)) : null), [data, year, isVat, fyEnd]);
  const needsMx = c0.incomeSpecialRate === "Yes" || c0.incomeExempt === "Yes" || c0.incomeTaxRegime === "Special rate";
  const qNow = round2(calc.auto.i45 + calc.auto.i46);
  const qChanged = rt.exists && form.qSnapshot !== undefined && Math.abs(num(form.qSnapshot) - qNow) >= 0.005;
  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1702-RT — Annual Income Tax Return for Corporations, Partnerships and Other Non-Individual Taxpayers Subject Only to REGULAR Income Tax Rate", version: "January 2018 (ENCS), eBIRForms v2018C", formNo: "1702-RT",
    period: `Year ended ${info.yearEnded}`, company: info0,
    blocks: [sectionBlock(secs.header, calc, form), infoBlock("Part I – Background Information", bg1702(info0, c0, b.method)), sectionBlock(secs.part2, calc, form),
      paymentBlock(form, "Part III – Details of Payment (items 23 to 26)"), sectionBlock(secs.part4, calc, form), sectionBlock(secs.part5, calc, form), sectionBlock(secs.schedI, calc, form),
      (form.special2 || []).length ? tableBlock("Schedule II – Special Allowable Itemized Deductions", SPECIAL2_COLS, form.special2) : null,
      sectionBlock(secs.schedIII, calc, form),
      (form.nolco || []).length ? tableBlock("Schedule IIIA – Computation of Available NOLCO", NOLCO_COLS, form.nolco) : null,
      (form.mcitCarry || []).length ? tableBlock("Schedule IV – Computation of MCIT", MCIT_COLS, form.mcitCarry) : null,
      sectionBlock(secs.schedV, calc, form), annualQuarterBlock(data, postings, "1702", year)],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1702RT_${year}_worksheet.pdf`,
  }));
  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}>
      <button className="io-btn" onClick={() => onNavigate("form2307")}>Open Creditable Withholding Taxes</button>
      <button className="io-btn" onClick={() => onNavigate("sawt")}>Open SAWT</button>
      <button className="io-btn" onClick={() => onNavigate("income")}>Open Income Statement</button>
      <WorksheetButtons hook={exp} />
    </div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={[["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"], ["dateIncorporation", "Date of Incorporation"]]} />
    <ReturnStatusBar rt={rt} dueDefault={dueAnnual(b.fy)} dueNote={fyEnd === 12 ? "April 15 · verify" : "15th day of the 4th month after year-end · verify"}>{selectors}</ReturnStatusBar>
    {needsMx && <Notice tone="danger">This taxpayer needs 1702-MX / 1702-EX (income at special / preferential rates, or exempt income, per Company Details). This screen covers 1702-RT only.</Notice>}
    {qChanged && <Notice>A quarterly 1702Q changed since this return was prepared (quarterly payments now {fmtPlain(qNow)}, were {fmtPlain(form.qSnapshot)}). Typed values are kept. <button className="tx-link" onClick={() => rt.patch({ qSnapshot: qNow })}>Acknowledge</button></Notice>}
    {calc.t("amended") === "Yes" && <Notice>Amended return: item 51 is unlocked. Key in the corrected figures.</Notice>}
    <div className="tx-note" style={{ padding: "0 0 8px" }}>{isVat ? "Sales exclude 12% output VAT." : "Non-VAT: sales are the gross receipts as recorded."}</div>
    <FormSection section={secs.header} calc={calc} rt={rt} />
    <BackgroundBlock items={bg1702(info0, c0, b.method)} onNavigate={onNavigate} />
    <FormSection section={{ ...secs.part4, headRight: <span className="tx-muted">Rate (item 40): <RateSelect value={form.rate} dflt={corpRate(c0)} onChange={(v) => rt.patch({ rate: v })} /></span> }} calc={calc} rt={rt} />
    <FormSection section={secs.part5} calc={calc} rt={rt} />
    <FormSection section={secs.schedI} calc={calc} rt={rt} />
    {!b.osd && <UnmappedLink fig={b.fig} onNavigate={onNavigate} />}
    {!b.osd && <ListTable title="Schedule II – Special Allowable Itemized Deductions (items 1 to 5)" rows={form.special2 || []} onChange={(rows) => rt.patch({ special2: rows })} cols={SPECIAL2_COLS} />}
    <FormSection section={secs.schedIII} calc={calc} rt={rt} />
    {!b.osd && <CarryTable title="Schedule IIIA – Computation of Available NOLCO (items 4 to 8)" rows={form.nolco || []} onChange={(rows) => rt.patch({ nolco: rows })} cols={NOLCO_COLS}
      computedCol={["E Unapplied (A − (B + C + D))", nolcoE]} totalCol="appliedCurrent" totalLabel="8 Total NOLCO applied (column D) → item 36"
      note="Pre-filled from last year's saved return. Losses for 2020 and 2021 carry over 5 years (Bayanihan Act), others 3 years; type the Expired column." />}
    <CarryTable title="Schedule IV – Computation of Minimum Corporate Income Tax (MCIT)" rows={form.mcitCarry || []} onChange={(rows) => rt.patch({ mcitCarry: rows })} cols={MCIT_COLS}
      computedCol={["G Balance (C − (D + E + F)), C = B − A", mcitG]} totalCol="appliedCurrent" totalLabel="4 Total excess MCIT applied (column F) → item 47"
      note={`Excess MCIT over normal income tax can be credited within the next 3 years. ${b.mc.applies ? (b.mc.known ? `MCIT applies from ${b.mc.firstYear}.` : "Set Year Business Started in Company Details to grey out MCIT in the first 4 years.") : `MCIT is not applicable for ${year} (first 4 taxable years).`}`} />
    <FormSection section={secs.schedV} calc={calc} rt={rt} />
    <FormSection section={secs.part2} calc={calc} rt={rt} />
    <PaymentDetails rt={rt} title="Part III – Details of Payment (items 23 to 26)" firstItem={23} amountDue={calc.v.i21} />
    <AnnualRecon b={b} data={data} vs={vs} salesKey="i29" dedKey={b.osd ? null : "i34"} cwtKeys={["i48", "i49"]} qKey="i46" qDue={round2(annualQuarterDue(data, postings, "1702", year) - calc.v.i45)} recon={[calc.v.r10, calc.v.i39]} faDep={faDep} />
    <AttachmentChecklist rt={rt} items={ATTACH_ANNUAL} title="Attachment checklist (feeds item 22; filed through eAFS within 15 days per RMC 20-2026)" onOpen={onNavigate} />
    <Declaration data={data} rt={rt} showTreasurer signatoryLabel="President / Principal Officer / Authorized Representative" />
    <div className="rc-warn">Rate: 25%, or 20% when net taxable income is ₱5,000,000 or less and total assets (excluding land) are ₱100,000,000 or less (CREATE). Foreign corporations, proprietary educational institutions and hospitals (10%), and PEZA / BOI entities fall outside 1702-RT.</div>
  </>);
}

/* ------------------------------ Page ------------------------------ */

export function IncomeAnnualPage({ data, setData, postings, onNavigate }) {
  const now = new Date();
  const [year, setYear] = useState(now.getMonth() < 4 ? now.getFullYear() - 1 : now.getFullYear());
  const indiv = (data.company || {}).taxpayerType === "Individual";
  const selectors = <YearSelect years={getAvailableYears(data)} value={year} onChange={setYear} />;
  const props = { data, setData, postings, onNavigate, year, selectors };
  return (
    <div>
      <TaxStyle />
      <SectionHeader icon={FileText} title={indiv ? "BIR Form 1701 – Annual Income Tax Return (Individuals)" : "BIR Form 1702-RT – Annual Income Tax Return (Regular Rate)"}
        subtitle={`Follows Company Details → Taxpayer Type (${indiv ? "Individual" : "Non-Individual"}). Full-year figures from the books and the saved quarterly returns. Encoding worksheet for eBIRForms / eFPS — no .dat file.`} />
      {indiv ? <Screen1701 {...props} /> : <Screen1702 {...props} />}
    </div>
  );
}

export { normalizeTin };

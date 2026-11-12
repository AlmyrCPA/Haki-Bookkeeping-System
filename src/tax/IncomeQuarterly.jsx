// BIR Form 1701Q (January 2018 ENCS, individuals) / 1702Q (January 2018 ENCS, v2018C, corporations).
// The tab follows Company Details → Taxpayer Type. Q1 to Q3 only; Q4 is the annual return.
//
// The form's "this quarter" lines (sales, cost of sales, deductions…) are the quarter's own figures and
// the previous quarters' taxable income is added on its own line, so taxable income to date is the
// cumulative year-to-date figure the tax table (or rate) is applied to. Previous quarters come from the
// saved returns; when a quarter wasn't saved in Haki, from the books.
import React, { useMemo, useState } from "react";
import { FileText } from "lucide-react";
import { SectionHeader, getAvailableYears, num, round2, fmtPlain, isVatCompany, computeVatSchedule, aggregateVatByYear, Field, LabeledField } from "../BookkeepingApp";
import {
  TaxStyle, makeCalc, useReturn, FormSection, ReturnStatusBar, BackgroundBlock, CompanyPrereqWarning, PaymentDetails,
  Declaration, companyInfo, quarterKey, fiscalQuarterRange, sectionBlock, infoBlock, paymentBlock, signatureBlock,
  useWorksheetExport, WorksheetButtons, ExportStatus, YearSelect, Notice, ReconStrip, graduatedTax, paymentTotal,
  addDays, mdy, EIGHT_PCT_LIMIT, taxTableFor,
} from "./shared";
import { incomeFigures, cwtInRange, mcitApplies, corpRate } from "./incomeData";

const QN = ["1st", "2nd", "3rd", "4th"];
export const due1701q = (year, q) => new Date(year, [4, 7, 10][q - 1], 15);
export const due1702q = (rangeEnd) => addDays(rangeEnd, 60);

export function electionFromCompany(c) {
  const e = c.taxRateElection || "Graduated + Itemized";
  if (e === "8%") return { rate: "8%", method: "N/A under 8%" };
  return { rate: "Graduated", method: e.includes("OSD") ? "OSD" : "Itemized" };
}
// Q2 / Q3 inherit the election made on Q1's return.
function electionFor(data, store, year, q, form) {
  const c = data.company || {};
  const q1 = q > 1 ? (data[store] || {})[quarterKey(year, 1)] : null;
  const src = q1 ? (q1.m || {}) : (form.m || {});
  if (store === "form1701q") {
    const d = electionFromCompany(c);
    return { rate: src.taxRate || d.rate, method: src.taxRate === "8%" ? "N/A under 8%" : (src.method || d.method), locked: !!q1 };
  }
  return { method: src.method || c.deductionMethod || "Itemized", locked: !!q1 };
}

function prevPayments(data, store, year, q, build) {
  let paid = 0;
  for (let k = 1; k < q; k++) {
    const f = (data[store] || {})[quarterKey(year, k)];
    if (!f || !f.filed) continue;
    const pt = paymentTotal(f);
    paid += pt || Math.max(0, build(k).calc.v[store === "form1701q" ? "i30" : "i24"]);
  }
  return round2(paid);
}

/* ------------------------------ 1701Q ------------------------------ */

export function build1701q(data, postings, year, q) {
  const c0 = data.company || {};
  const store = "form1701q";
  const form = (data[store] || {})[quarterKey(year, q)] || { m: {}, overrides: {} };
  const rng = fiscalQuarterRange(year, c0.fiscalYearEndMonth || 12, q);
  const el = electionFor(data, store, year, q, form);
  const eight = el.rate === "8%";
  const osd = !eight && el.method === "OSD";
  const fig = incomeFigures(data, postings, rng.start, rng.end);
  const mixed = c0.filerKind === "Mixed income";
  const c = makeCalc(form);
  // Previous quarters' cumulative taxable income.
  let prev = 0, prevSrc = "No previous quarter in this taxable year";
  if (q > 1) {
    const saved = (data[store] || {})[quarterKey(year, q - 1)];
    if (saved) { const pc = build1701q(data, postings, year, q - 1).calc; prev = eight ? pc.v.i51 : pc.v.i46; prevSrc = `Saved ${QN[q - 2]} quarter return`; }
    else {
      const pf = incomeFigures(data, postings, rng.ytdStart, addDays(rng.start, -1));
      prev = eight ? pf.netSales + pf.otherIncome : osd ? round2((pf.netSales + pf.otherIncome) * 0.6) : pf.grossIncome + pf.otherIncome - pf.deductions;
      prevSrc = `Books, ${mdy(rng.ytdStart)} to ${mdy(addDays(rng.start, -1))} (no saved ${QN[q - 2]} quarter return)`;
    }
  }
  if (!eight) {
    c.a("i36", fig.netSales);
    c.a("i37", osd ? 0 : fig.costOfSales);
    c.set("i38", c.v.i36 - c.v.i37);
    c.a("i39", fig.otherIncome);
    c.man("i40");
    c.set("i41", c.v.i38 + c.v.i39 + c.v.i40);
    c.a("i42", osd ? 0 : fig.deductions);
    c.set("i43", osd ? c.v.i41 * 0.4 : 0);
    c.set("i44", c.v.i41 - (osd ? c.v.i43 : c.v.i42));
    c.a("i45", prev);
    c.set("i46", c.v.i44 + c.v.i45);
    c.set("taxDueG", graduatedTax(Math.max(0, c.v.i46), year));
  } else {
    c.a("i47", fig.netSales);
    c.a("i48", fig.otherIncome);
    c.set("i49", c.v.i47 + c.v.i48);
    c.a("i50", prev);
    c.set("i51", c.v.i49 + c.v.i50);
    c.set("i52", mixed ? 0 : 250000);
    c.set("i53", Math.max(0, c.v.i51 - c.v.i52));
    c.set("i54", c.v.i53 * 0.08);
  }
  const build = (k) => build1701q(data, postings, year, k);
  c.man("i55");
  c.a("i56", prevPayments(data, store, year, q, build));
  c.a("i57", cwtInRange(data, rng.ytdStart, addDays(rng.start, -1)));
  c.a("i58", cwtInRange(data, rng.start, rng.end));
  c.set("i59", c.t("amended") === "Yes" ? c.m("i59") : 0);
  c.man("i60"); c.man("i61");
  c.set("i62", c.v.i55 + c.v.i56 + c.v.i57 + c.v.i58 + c.v.i59 + c.v.i60 + c.v.i61);
  c.set("i26", eight ? c.v.i54 : c.v.taxDueG);
  c.set("i63", c.v.i26 - c.v.i62);
  c.man("i64"); c.man("i65"); c.man("i66");
  c.set("i68", c.v.i64 + c.v.i65 + c.v.i66);
  c.set("i27", c.v.i62);
  c.set("i28", c.v.i26 - c.v.i27);
  c.set("i29", c.v.i68);
  c.set("i30", c.v.i28 + c.v.i29);
  // Spouse column (typed).
  c.man("s26"); c.man("s27"); c.set("s28", c.v.s26 - c.v.s27); c.man("s29"); c.set("s30", c.v.s28 + c.v.s29);
  c.set("i31", c.v.i30 + c.v.s30);
  const grossTest = eight ? c.v.i51 : round2((c.v.i36 || 0) + (c.v.i39 || 0) + prev);
  return { calc: c, form, fig, rng, el, eight, osd, mixed, prevSrc, grossTest };
}

function atc1701q(c, el) {
  if (c.filerKind === "Mixed income") return el.rate === "8%" ? "II017" : "II013";
  if (el.rate === "8%") return "II017";
  return c.filerKind === "Professional" ? "II012" : "II011";
}

function sections1701q(b, info) {
  const { calc: c, el, eight, osd } = b;
  const amended = c.t("amended") === "Yes";
  const spouse = info.spouse;
  const cols = spouse ? ["A Taxpayer / Filer", "B Spouse"] : ["Taxpayer / Filer"];
  const pair = (k, t, sk, st) => (spouse ? { k: [k, sk], t: [t, st] } : { k, t });
  return {
    header: { title: "Header (items 1 to 4)", lines: [
      { no: "1", label: "For the Year", k: "_y", t: "ro", text: String(info.year) },
      { no: "2", label: "Quarter", k: "_q", t: "ro", text: `${QN[info.q - 1]} Quarter` },
      { no: "3", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "4", label: "Number of Sheet/s Attached", k: "sheetsAttached", t: "int" },
    ]},
    election: { title: "Tax rate and deduction method (items 15 / 16)", note: el.locked ? "Elected on the 1st quarter return: later quarters follow it (locked)." : "Elected on the 1st quarter return and kept for the year.", lines: [
      { no: "7A", label: "ATC (confirm against your eBIRForms list, II011 to II017)", k: "atc", t: "text", dflt: info.atc },
      { no: "15", label: `Tax Rate${el.locked ? " (elected on Q1)" : ""}`, k: "taxRate", t: "select", opts: ["Graduated", "8%"], dflt: el.rate, dis: el.locked },
      { no: "16", label: `Method of Deduction${el.locked ? " (elected on Q1)" : ""}`, k: "method", t: "select", opts: ["Itemized", "OSD"], dflt: el.method, dis: el.locked || eight, hide: eight },
    ]},
    part3: { title: "Part III – Total Tax Payable", cols, lines: [
      { no: "26", label: `Tax Due (from Part V, ${eight ? "Schedule II item 54" : "Schedule I item 46"})`, ...pair("i26", "calc", "s26", "man") },
      { no: "27", label: "Less: Tax Credits / Payments (from Part V, Schedule III item 62)", ...pair("i27", "calc", "s27", "man") },
      { no: "28", label: "Tax Payable / (Overpayment) (item 26 less item 27)", ...pair("i28", "calc", "s28", "calc"), due: true, f: spouse ? [(r) => `${r("i26")}-${r("i27")}`, (r) => `${r("s26")}-${r("s27")}`] : (r) => `${r("i26")}-${r("i27")}` },
      { no: "29", label: "Add: Total Penalties (from Part V, Schedule IV item 68)", ...pair("i29", "calc", "s29", "man") },
      { no: "30", label: "Total Amount Payable / (Overpayment) (sum of items 28 and 29)", ...pair("i30", "calc", "s30", "calc"), strong: true, due: true, f: spouse ? [(r) => `${r("i28")}+${r("i29")}`, (r) => `${r("s28")}+${r("s29")}`] : (r) => `${r("i28")}+${r("i29")}` },
      { no: "31", label: "Aggregate Amount Payable / (Overpayment) (taxpayer + spouse)", k: "i31", t: "calc", strong: true, due: true, hide: !spouse },
      { no: "", label: "If overpayment, mark one (confirm wording)", k: "overpayChoice", t: "select", opts: ["To be refunded", "To be issued a Tax Credit Certificate", "To be carried over to the next quarter"], hide: !(c.v.i30 < 0) },
    ]},
    schedI: { title: "Schedule I – Graduated Rates (items 36 to 46)", hide: eight, note: "Items 36 to 44 are this quarter's figures; item 45 adds the previous quarters, so item 46 is the taxable income to date.", lines: [
      { no: "36", label: "Sales / Revenues / Receipts / Fees (net of returns and discounts; excludes 12% output VAT)", k: "i36", t: "auto", src: "Income Statement revenue accounts for the quarter, less sales returns / discounts" },
      { no: "37", label: "Less: Cost of Sales / Services (itemized only)", k: "i37", t: "auto", src: "Income Statement cost of sales (inventory from the Income Statement inventory log)", dis: osd },
      { no: "38", label: "Gross Income / (Loss) from Operation (item 36 less item 37)", k: "i38", t: "calc", f: (r) => `${r("i36")}-${r("i37")}` },
      { no: "39", label: "Add: Other Non-Operating Income", k: "i39", t: "auto", src: "Other Income accounts not flagged Subject to Final Tax" },
      { no: "40", label: "Amount Received / Share in Income by a Partner from General Professional Partnership (GPP)", k: "i40", t: "man" },
      { no: "41", label: "Total Gross Income (sum of items 38 to 40)", k: "i41", t: "calc", strong: true, f: (r) => `${r("i38")}+${r("i39")}+${r("i40")}` },
      { no: "42", label: "Less: Allowable Itemized Deductions", k: "i42", t: "auto", src: "Expense accounts for the quarter, less those mapped Non-deductible", dis: osd },
      { no: "43", label: "Less: Optional Standard Deduction (40% of item 41)", k: "i43", t: "calc", f: (r) => `${r("i41")}*0.4*${osd ? 1 : 0}` },
      { no: "44", label: "Net Income / (Loss) for the Quarter (item 41 less item 42 or 43)", k: "i44", t: "calc", strong: true },
      { no: "45", label: "Add: Taxable Income / (Loss) Previous Quarter/s", k: "i45", t: "auto", src: b.prevSrc },
      { no: "46", label: "Total Taxable Income / (Loss) To Date (sum of items 44 and 45)", k: "i46", t: "calc", strong: true, f: (r) => `${r("i44")}+${r("i45")}` },
      { no: "", label: `Tax Due (graduated table, ${taxTableFor(info.year).label}, applied to item 46)`, k: "taxDueG", t: "calc", strong: true },
    ]},
    schedII: { title: "Schedule II – 8% Gross Sales / Receipts (items 47 to 54)", hide: !eight, lines: [
      { no: "47", label: "Sales / Revenues / Receipts / Fees (net of returns and discounts)", k: "i47", t: "auto", src: "Income Statement revenue accounts for the quarter" },
      { no: "48", label: "Add: Non-Operating Income", k: "i48", t: "auto", src: "Other Income accounts not flagged Subject to Final Tax" },
      { no: "49", label: "Total Income for the Quarter (sum of items 47 and 48)", k: "i49", t: "calc", f: (r) => `${r("i47")}+${r("i48")}` },
      { no: "50", label: "Add: Total Taxable Income / (Loss) Previous Quarter/s", k: "i50", t: "auto", src: b.prevSrc },
      { no: "51", label: "Cumulative Taxable Income / (Loss) as of This Quarter (sum of items 49 and 50)", k: "i51", t: "calc", strong: true, f: (r) => `${r("i49")}+${r("i50")}` },
      { no: "52", label: `Less: Allowable reduction from gross sales / receipts and other non-operating income of ₱250,000${b.mixed ? " (not allowed for a mixed-income earner)" : ""}`, k: "i52", t: "calc" },
      { no: "53", label: "Taxable Income / (Loss) To Date (item 51 less item 52)", k: "i53", t: "calc", f: (r) => `MAX(${r("i51")}-${r("i52")},0)` },
      { no: "54", label: "Tax Due (item 53 × 8%)", k: "i54", t: "calc", strong: true, f: (r) => `${r("i53")}*0.08` },
    ]},
    schedIII: { title: "Schedule III – Tax Credits / Payments (items 55 to 63)", lines: [
      { no: "55", label: "Prior Year's Excess Credits", k: "i55", t: "man" },
      { no: "56", label: "Tax Payment/s for the Previous Quarter/s", k: "i56", t: "auto", src: "Previous quarters' returns marked Filed: amount paid (Part IV)" },
      { no: "57", label: "Creditable Tax Withheld for the Previous Quarter/s", k: "i57", t: "auto", src: "Creditable Withholding Taxes tab, earlier quarters of the year" },
      { no: "58", label: "Creditable Tax Withheld per BIR Form No. 2307 for this Quarter", k: "i58", t: "auto", src: "Creditable Withholding Taxes tab, rows dated in this quarter" },
      { no: "59", label: "Tax Paid in Return Previously Filed, if this is an amended return", k: "i59", t: "man", dis: !amended },
      { no: "60", label: "Foreign Tax Credits, if applicable", k: "i60", t: "man" },
      { no: "61", label: "Other Tax Credits / Payments (specify)", k: "i61", t: "man" },
      { no: "", label: "Other credits: specify", k: "i61Specify", t: "text", hide: !c.v.i61 },
      { no: "62", label: "Total Tax Credits / Payments (sum of items 55 to 61)", k: "i62", t: "calc", strong: true, f: (r) => `${r("i55")}+${r("i56")}+${r("i57")}+${r("i58")}+${r("i59")}+${r("i60")}+${r("i61")}` },
      { no: "63", label: "Tax Payable / (Overpayment)", k: "i63", t: "calc", due: true },
    ]},
    schedIV: { title: "Schedule IV – Penalties (manual)", lines: [
      { no: "64", label: "Surcharge", k: "i64", t: "man" },
      { no: "65", label: "Interest", k: "i65", t: "man" },
      { no: "66", label: "Compromise", k: "i66", t: "man" },
      { no: "68", label: "Total Penalties (sum of items 64 to 66) → item 29", k: "i68", t: "calc", f: (r) => `${r("i64")}+${r("i65")}+${r("i66")}` },
    ]},
  };
}

const bg1701q = (i, c, el) => [["5", "TIN", i.tinText], ["6", "RDO Code", i.rdo], ["7", "Taxpayer Type", c.filerKind === "Professional" ? "Professional" : "Single Proprietor"], ["8", "Taxpayer's Name", i.name],
  ["9", "Date of Birth", c.dateOfBirth], ["10", "Registered Address", i.address], ["10A", "ZIP Code", i.zip], ["11", "Email Address", i.email], ["12", "Citizenship", c.citizenship],
  ["13", "Foreign Tax Number", c.foreignTaxNumber || "N/A"], ["14", "Claiming Foreign Tax Credits?", c.claimsForeignTaxCredit || "No"], ["16A", "Contact Number", i.contact]];

function SpouseInfo({ rt, c }) {
  const [open, setOpen] = useState(false);
  const m = rt.form.m || {};
  const F = (k, label, dflt = "") => <LabeledField label={label}><Field value={m[k] ?? dflt} onChange={(v) => rt.setM(k, v)} /></LabeledField>;
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>Part II – Information on Spouse (items 17 to 25A)</span><button className="tx-link" onClick={() => setOpen(!open)}>{open ? "Collapse" : "Expand"}</button></div>
      {open && <div className="tx-inline-fields">
        {F("spTin", "17 Spouse TIN", c.spouseTin)}{F("spRdo", "18 RDO Code")}{F("spType", "19 Filer Type")}{F("spAtc", "19A ATC")}
        {F("spName", "20 Spouse's Name", c.spouseName)}{F("spCitizenship", "21 Citizenship")}{F("spForeignTax", "22 Foreign Tax Number")}
        {F("spForeignCredit", "23 Claiming Foreign Tax Credits?")}{F("spRate", "24 Tax Rate")}{F("spMethod", "25 Method of Deduction")}
      </div>}
    </div>
  );
}

function Screen1701Q({ data, setData, postings, onNavigate, year, q, selectors }) {
  const info0 = companyInfo(data);
  const c0 = data.company || {};
  const rt = useReturn(data, setData, "form1701q", quarterKey(year, q));
  const b = useMemo(() => build1701q(data, postings, year, q), [data, postings, year, q]);
  const spouse = c0.civilStatus === "Married" && c0.spouseHasIncome === "Yes";
  const info = { year, q, spouse, atc: atc1701q(c0, b.el) };
  const secs = sections1701q(b, info);
  const calc = b.calc;
  const isVat = isVatCompany(data);
  const vs = useMemo(() => (isVat ? computeVatSchedule(data, year, aggregateVatByYear(data)) : null), [data, year, isVat]);
  const vatBase = vs ? round2(vs.quarters[q - 1].totalSales) : null;
  const rtView = { ...rt, form: { ...rt.form, m: { taxRate: b.el.rate, method: b.el.method, ...(rt.form.m || {}) } } };
  if (b.el.locked) { rtView.form.m.taxRate = b.el.rate; rtView.form.m.method = b.el.method; }
  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1701Q — Quarterly Income Tax Return for Individuals, Estates and Trusts", version: "January 2018 (ENCS)", formNo: "1701Q",
    period: `${QN[q - 1]} Quarter ${year}`, company: info0,
    blocks: [sectionBlock(secs.header, calc, rt.form), infoBlock("Part I – Background Information", bg1701q(info0, c0, b.el)), sectionBlock(secs.election, calc, rtView.form),
      sectionBlock(secs.part3, calc, rt.form), paymentBlock(rt.form, "Part IV – Details of Payment (items 32 to 35)"),
      sectionBlock(secs.schedI, calc, rt.form), sectionBlock(secs.schedII, calc, rt.form), sectionBlock(secs.schedIII, calc, rt.form), sectionBlock(secs.schedIV, calc, rt.form),
      quarterSummaryBlock(data, "form1701q", year, q, (k) => build1701q(data, postings, year, k), b.eight ? "i51" : "i46", "i30")],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1701Q_Q${q}${year}_worksheet.pdf`,
  }));
  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}>
      <button className="io-btn" onClick={() => onNavigate("form2307")}>Open Creditable Withholding Taxes</button>
      <button className="io-btn" onClick={() => onNavigate("income")}>Open Income Statement</button>
      <WorksheetButtons hook={exp} />
    </div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={[["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"], ["dateOfBirth", "Date of Birth"], ["citizenship", "Citizenship"]]} />
    <ReturnStatusBar rt={rt} dueDefault={due1701q(year, q)} extra={b.eight ? <span className="tx-pill warn">No 2551Q required</span> : null}>{selectors}</ReturnStatusBar>
    {b.eight && <Notice tone="info">8% covers income tax and percentage tax. Do not file 2551Q for this taxpayer for the year.</Notice>}
    {b.grossTest > EIGHT_PCT_LIMIT && b.eight && <Notice tone="danger">Gross receipts exceed ₱3,000,000. The 8% option no longer applies; the taxpayer becomes subject to graduated rates and VAT / percentage tax for the year. Review with the accountant.</Notice>}
    {calc.t("amended") === "Yes" && <Notice>Amended return: item 59 is unlocked. Key in the figures of the original filing.</Notice>}
    <div className="tx-note" style={{ padding: "0 0 8px" }}>{isVat ? "Sales exclude 12% output VAT." : "Non-VAT: sales are the gross receipts as recorded."} NOLCO is deducted on the annual return only.</div>
    <FormSection section={secs.header} calc={calc} rt={rt} />
    <BackgroundBlock items={bg1701q(info0, c0, b.el)} onNavigate={onNavigate} />
    <FormSection section={secs.election} calc={calc} rt={rtView} />
    {spouse && <SpouseInfo rt={rt} c={c0} />}
    <FormSection section={secs.schedI} calc={calc} rt={rt} />
    <FormSection section={secs.schedII} calc={calc} rt={rt} />
    {!b.eight && <UnmappedNote fig={b.fig} onNavigate={onNavigate} />}
    <FormSection section={secs.schedIII} calc={calc} rt={rt} />
    <FormSection section={secs.schedIV} calc={calc} rt={rt} />
    <FormSection section={secs.part3} calc={calc} rt={rt} />
    <PaymentDetails rt={rt} title="Part IV – Details of Payment (items 32 to 35)" firstItem={32} amountDue={calc.v.i30} />
    <IncomeRecon b={b} q={q} vatBase={vatBase} salesKey={b.eight ? "i47" : "i36"} dedKey={b.eight ? null : "i42"} cwtKey="i58" taxableKey={b.eight ? "i49" : "i44"} />
    <Declaration data={data} rt={rt} />
    <div className="rc-warn">2307s are claimed in the quarter the income was received (the Creditable Withholding Taxes row date). A 2307 received after the quarter closed for a prior quarter's income can be claimed in a later quarter or on the annual return.</div>
  </>);
}

function UnmappedNote({ fig, onNavigate }) {
  if (!fig.unmapped.length && !fig.nonDeductible.length) return null;
  return (
    <div className="tx-note" style={{ padding: "0 0 10px" }}>
      {fig.nonDeductible.length > 0 && <>Non-deductible: {fmtPlain(fig.nonDeductibleTotal)} excluded ({fig.nonDeductible.map((x) => x.name).join(", ")}). </>}
      {fig.unmapped.length > 0 && <span className="tx-amber">{fig.unmapped.length} expense account{fig.unmapped.length === 1 ? "" : "s"} without an ITR deduction line ({fig.unmapped.map((x) => x.name).join(", ")}) counted under Others. </span>}
      <button className="tx-link" onClick={() => onNavigate("coa")}>Review in Chart of Accounts</button>
    </div>
  );
}

function IncomeRecon({ b, vatBase, salesKey, dedKey, cwtKey, taxableKey }) {
  const { calc, fig } = b;
  return (
    <ReconStrip title="Books check" colA="Return" colB="Books" items={[
      { label: "Sales / receipts vs Income Statement net sales for the period", a: calc.v[salesKey], b: fig.netSales },
      vatBase !== null && vatBase !== undefined ? { label: "Sales vs VAT Schedule sales base for the quarter (net of VAT)", a: calc.v[salesKey], b: vatBase } : null,
      dedKey && !b.osd ? { label: `Itemized deductions vs Income Statement expenses less non-deductible (${fmtPlain(fig.nonDeductibleTotal)})`, a: calc.v[dedKey], b: round2(fig.totalExpenses - fig.nonDeductibleTotal) } : null,
      { label: "Creditable tax this quarter vs Creditable Withholding Taxes tab", a: calc.v[cwtKey], b: calc.auto[cwtKey] },
      { label: "Net income per books vs taxable income this quarter (non-deductible, final-tax income, OSD vs actual)", a: calc.v[taxableKey] ?? 0, b: fig.netIncomeBooks, info: "information only" },
    ]} causes="Cash Receipts CWT rows dated outside the quarter, income accounts not set up as Revenue / Other Income, manual General Journal entries, 2307s not yet recorded, or overrides." />
  );
}

export function quarterSummaryBlock(data, store, year, q, build, taxableKey, payKey) {
  const body = [];
  for (let k = 1; k <= q; k++) {
    const f = (data[store] || {})[quarterKey(year, k)];
    const cc = build(k).calc;
    body.push([`${QN[k - 1]} Quarter`, f ? (f.filed ? `Filed ${f.dateFiled || ""}` : "Draft") : "Not saved", cc.v[taxableKey] || 0, f && f.filed ? (paymentTotal(f) || Math.max(0, cc.v[payKey])) : 0]);
  }
  return { type: "table", title: "Quarterly summary (taxable income to date and payments)", head: ["Quarter", "Status", "Taxable income to date", "Payment"], body, numCols: [2, 3] };
}

/* ------------------------------ 1702Q ------------------------------ */

export function build1702q(data, postings, year, q) {
  const c0 = data.company || {};
  const store = "form1702q";
  const form = (data[store] || {})[quarterKey(year, q)] || { m: {}, overrides: {} };
  const rng = fiscalQuarterRange(year, c0.fiscalYearEndMonth || 12, q);
  const el = electionFor(data, store, year, q, form);
  const osd = el.method === "OSD";
  const fig = incomeFigures(data, postings, rng.start, rng.end);
  const c = makeCalc(form);
  let prevTI = 0, prevGI = 0, prevSrc = "No previous quarter in this taxable year";
  if (q > 1) {
    const saved = (data[store] || {})[quarterKey(year, q - 1)];
    if (saved) { const pc = build1702q(data, postings, year, q - 1).calc; prevTI = pc.v.s2_10; prevGI = pc.v.s3_3; prevSrc = `Saved ${QN[q - 2]} quarter return`; }
    else {
      const pf = incomeFigures(data, postings, rng.ytdStart, addDays(rng.start, -1));
      const tgi = pf.grossIncome + pf.otherIncome;
      prevTI = osd ? round2(tgi * 0.6) : round2(tgi - pf.deductions);
      prevGI = pf.grossIncome;
      prevSrc = `Books, ${mdy(rng.ytdStart)} to ${mdy(addDays(rng.start, -1))} (no saved ${QN[q - 2]} quarter return)`;
    }
  }
  // Schedule 2, regular rate
  c.a("s2_1", fig.netSales);
  c.a("s2_2", fig.costOfSales);
  c.set("s2_3", c.v.s2_1 - c.v.s2_2);
  c.a("s2_4", fig.otherIncome);
  c.set("s2_5", c.v.s2_3 + c.v.s2_4);
  c.a("s2_6", osd ? 0 : fig.deductions);
  c.set("s2_7", osd ? c.v.s2_5 * 0.4 : 0);
  c.set("s2_8", c.v.s2_5 - (osd ? c.v.s2_7 : c.v.s2_6));
  c.a("s2_9", prevTI);
  c.set("s2_10", c.v.s2_8 + c.v.s2_9);
  const rate = form.rate === undefined || form.rate === "" ? corpRate(c0) : num(form.rate);
  c.rate = rate;
  c.set("s2_12", Math.max(0, c.v.s2_10) * rate);
  // Schedule 3, MCIT
  const mc = mcitApplies(c0, year);
  c.set("s3_1", c.v.s2_3);
  c.a("s3_2", prevGI);
  c.set("s3_3", c.v.s3_1 + c.v.s3_2);
  c.set("s3_5", mc.applies ? Math.max(0, c.v.s3_3) * 0.02 : 0);
  c.man("s3_6");
  c.set("s2_13", Math.max(c.v.s2_12, c.v.s3_5));
  // Schedule 1, special rate (manual)
  ["1a", "1b", "2a", "2b", "4a", "4b", "6a", "6b", "8a", "8b", "10b"].forEach((k) => c.man(`s1_${k}`));
  c.set("s1_3a", c.v.s1_1a - c.v.s1_2a); c.set("s1_3b", c.v.s1_1b - c.v.s1_2b);
  c.set("s1_5a", c.v.s1_3a + c.v.s1_4a); c.set("s1_5b", c.v.s1_3b + c.v.s1_4b);
  c.set("s1_7a", c.v.s1_5a - c.v.s1_6a); c.set("s1_7b", c.v.s1_5b - c.v.s1_6b);
  c.set("s1_9a", c.v.s1_7a + c.v.s1_8a); c.set("s1_9b", c.v.s1_7b + c.v.s1_8b);
  c.set("s1_11b", Math.max(0, c.v.s1_9b) * (c.v.s1_10b / 100));
  // Schedule 4, credits
  const build = (k) => build1702q(data, postings, year, k);
  c.man("s4_1");
  c.a("s4_2", prevPayments(data, store, year, q, build));
  c.a("s4_3", cwtInRange(data, rng.ytdStart, addDays(rng.start, -1)));
  c.a("s4_4", cwtInRange(data, rng.start, rng.end));
  c.set("s4_5", c.t("amended") === "Yes" ? c.m("s4_5") : 0);
  c.man("s4_6"); c.man("s4_6a");
  c.set("s4_7", c.v.s4_1 + c.v.s4_2 + c.v.s4_3 + c.v.s4_4 + c.v.s4_5 + c.v.s4_6 + c.v.s4_6a);
  // Part II
  c.set("i14", c.v.s1_11b);
  c.set("i15", c.v.s2_13);
  c.set("i16", c.v.s1_11b);
  c.set("i17", c.v.i15 + c.v.i16);
  c.set("i18", c.v.s4_7);
  c.set("i19", c.v.i17 - c.v.i18);
  c.man("i20"); c.man("i21"); c.man("i22");
  c.set("i23", c.v.i20 + c.v.i21 + c.v.i22);
  c.set("i24", c.v.i19 + c.v.i23);
  return { calc: c, form, fig, rng, el, osd, mc, prevSrc };
}

function sections1702q(b, info) {
  const { calc: c, el, osd, mc } = b;
  const amended = c.t("amended") === "Yes";
  const special = info.special;
  const AB = ["A Exempt", "B Special"];
  return {
    header: { title: "Header (items 1 to 5A)", lines: [
      { no: "1", label: "For the", k: "_cal", t: "ro", text: info.fiscal ? "Fiscal" : "Calendar" },
      { no: "2", label: "Year Ended (MM/YYYY)", k: "_ye", t: "ro", text: info.yearEnded },
      { no: "3", label: "Quarter", k: "_q", t: "ro", text: `${QN[info.q - 1]} Quarter` },
      { no: "4", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "5", label: "Short Period Return?", k: "shortPeriod", t: "yesno", dflt: "No" },
      { no: "", label: "Short period: from / to", k: "shortFromTo", t: "text", ph: "MM/DD/YYYY – MM/DD/YYYY", hide: c.t("shortPeriod") !== "Yes" },
      { no: "5A", label: "ATC (confirm the ATC for your regime after RMC 66-2021)", k: "atc", t: "text", dflt: "IC010" },
    ]},
    part1x: { title: "Part I (items 12 / 13)", note: el.locked ? "Method of deduction was elected on the 1st quarter return (locked)." : "The method of deduction is elected on the 1st quarter return and kept for the year.", lines: [
      { no: "12", label: `Method of Deduction${el.locked ? " (elected on Q1)" : ""}`, k: "method", t: "select", opts: ["Itemized", "OSD"], dflt: el.method, dis: el.locked },
      { no: "13", label: "Tax Relief Availed?", k: "taxRelief", t: "yesno", dflt: info.c.taxRelief || "No" },
      { no: "13A", label: "If yes, specify", k: "taxReliefSpecify", t: "text", dflt: info.c.taxReliefSpecify || "", hide: c.t("taxRelief", info.c.taxRelief || "No") !== "Yes" },
    ]},
    part2: { title: "Part II – Total Tax Payable (confirm item numbers against your copy)", lines: [
      { no: "14", label: "Income Tax Due – Special Rate (from Schedule 1 item 11)", k: "i14", t: "calc", hide: !special },
      { no: "15", label: "Income Tax Due – Regular Rate (from Schedule 2 item 13)", k: "i15", t: "calc" },
      { no: "16", label: "Add: Income Tax Due – Special Rate (Schedule 1 item 11)", k: "i16", t: "calc" },
      { no: "17", label: "Aggregate Income Tax Due (sum of items 15 and 16)", k: "i17", t: "calc", strong: true, f: (r) => `${r("i15")}+${r("i16")}` },
      { no: "18", label: "Less: Total Tax Credits / Payments (from Schedule 4 item 7)", k: "i18", t: "calc" },
      { no: "19", label: "Net Tax Payable / (Overpayment) (item 17 less item 18)", k: "i19", t: "calc", due: true, f: (r) => `${r("i17")}-${r("i18")}` },
      { heading: "Add: Penalties (manual)" },
      { no: "20", label: "Surcharge", k: "i20", t: "man" },
      { no: "21", label: "Interest", k: "i21", t: "man" },
      { no: "22", label: "Compromise", k: "i22", t: "man" },
      { no: "23", label: "Total Penalties (sum of items 20 to 22)", k: "i23", t: "calc", f: (r) => `${r("i20")}+${r("i21")}+${r("i22")}` },
      { no: "24", label: "TOTAL AMOUNT PAYABLE / (Overpayment) (sum of items 19 and 23)", k: "i24", t: "calc", strong: true, due: true, f: (r) => `${r("i19")}+${r("i23")}` },
      { no: "25", label: "If overpayment, mark one", k: "overpayChoice", t: "select", opts: ["To be refunded", "To be issued a Tax Credit Certificate", "To be carried over as tax credit for next period"], hide: !(c.v.i24 < 0) },
    ]},
    sched1: { title: "Schedule 1 – Special Rate (manual entry for v1)", cols: AB, hide: !special, note: "Special-rate (PEZA / BOI) companies: encode figures manually.", lines: [
      { no: "1", label: "Sales / Revenues / Receipts / Fees", k: ["s1_1a", "s1_1b"], t: "man" },
      { no: "2", label: "Less: Cost of Sales / Services", k: ["s1_2a", "s1_2b"], t: "man" },
      { no: "3", label: "Gross Income from Operation", k: ["s1_3a", "s1_3b"], t: "calc" },
      { no: "4", label: "Add: Non-Operating and Other Income", k: ["s1_4a", "s1_4b"], t: "man" },
      { no: "5", label: "Total Gross Income", k: ["s1_5a", "s1_5b"], t: "calc" },
      { no: "6", label: "Less: Deductions (or 40% OSD)", k: ["s1_6a", "s1_6b"], t: "man" },
      { no: "7", label: "Taxable Income this Quarter", k: ["s1_7a", "s1_7b"], t: "calc" },
      { no: "8", label: "Add: Taxable Income Previous Quarter/s", k: ["s1_8a", "s1_8b"], t: "man" },
      { no: "9", label: "Total Taxable Income to Date", k: ["s1_9a", "s1_9b"], t: "calc" },
      { no: "10", label: "Tax Rate (%)", k: [null, "s1_10b"], t: [null, "man"] },
      { no: "11", label: "Income Tax Due – Special Rate", k: [null, "s1_11b"], t: [null, "calc"], strong: true },
    ]},
    sched2: { title: "Schedule 2 – Regular Rate (items 1 to 13)", note: "Items 1 to 8 are this quarter's figures; item 9 adds the previous quarters, so item 10 is the taxable income to date.", lines: [
      { no: "1", label: "Sales / Revenues / Receipts / Fees (net of returns; excludes output VAT)", k: "s2_1", t: "auto", src: "Income Statement revenue accounts for the quarter" },
      { no: "2", label: "Less: Cost of Sales / Services", k: "s2_2", t: "auto", src: "Income Statement cost of sales" },
      { no: "3", label: "Gross Income from Operation (item 1 less item 2)", k: "s2_3", t: "calc", f: (r) => `${r("s2_1")}-${r("s2_2")}` },
      { no: "4", label: "Add: Non-Operating and Other Income", k: "s2_4", t: "auto", src: "Other Income accounts not flagged Subject to Final Tax" },
      { no: "5", label: "Total Gross Income (sum of items 3 and 4)", k: "s2_5", t: "calc", strong: true, f: (r) => `${r("s2_3")}+${r("s2_4")}` },
      { no: "6", label: "Less: Deductions — Itemized", k: "s2_6", t: "auto", src: "Expense accounts for the quarter, less those mapped Non-deductible", dis: osd },
      { no: "7", label: "Less: Optional Standard Deduction (40% of item 5)", k: "s2_7", t: "calc" },
      { no: "8", label: "Net Taxable Income / (Loss) This Quarter", k: "s2_8", t: "calc", strong: true },
      { no: "9", label: "Add: Taxable Income / (Loss) Previous Quarter/s", k: "s2_9", t: "auto", src: b.prevSrc },
      { no: "10", label: "Total Taxable Income / (Loss) To Date (sum of items 8 and 9)", k: "s2_10", t: "calc", strong: true, f: (r) => `${r("s2_8")}+${r("s2_9")}` },
      { no: "11", label: `Tax Rate (${round2(c.rate * 100)}%; regime from Company Details, editable below)`, k: "_rate", t: "ro", text: `${round2(c.rate * 100)}%` },
      { no: "12", label: "Income Tax Due – Regular Rate (item 10 × item 11)", k: "s2_12", t: "calc" },
      { no: "", label: `MCIT Due (Schedule 3 item 5)${mc.applies ? "" : " — not applicable yet"}`, k: "s3_5", t: "calc" },
      { no: "13", label: "Income Tax Due (the higher of item 12 and MCIT) → Part II", k: "s2_13", t: "calc", strong: true, f: (r) => `MAX(${r("s2_12")},${r("s3_5")})` },
    ]},
    sched3: { title: "Schedule 3 – Minimum Corporate Income Tax (MCIT)", note: mc.applies ? (mc.known ? `MCIT applies from ${mc.firstYear} (4th taxable year after commencement). Confirm eligibility.` : "Year business started is blank in Company Details, so MCIT is shown as applicable. Confirm eligibility.") : `MCIT applies from the 4th taxable year after commencement (${mc.firstYear}); not applicable for ${info.year}.`, lines: [
      { no: "1", label: "Gross Income – this quarter (sales less cost of sales)", k: "s3_1", t: "calc" },
      { no: "2", label: "Add: Gross Income – Previous Quarter/s", k: "s3_2", t: "auto", src: b.prevSrc },
      { no: "3", label: "Total Gross Income to Date", k: "s3_3", t: "calc", f: (r) => `${r("s3_1")}+${r("s3_2")}` },
      { no: "4", label: "MCIT Rate", k: "_mr", t: "ro", text: "2%" },
      { no: "5", label: "Minimum Corporate Income Tax (item 3 × 2%)", k: "s3_5", t: "calc", strong: true },
      { no: "6", label: "Excess MCIT carry-forward from prior years (information)", k: "s3_6", t: "man" },
    ]},
    sched4: { title: "Schedule 4 – Tax Credits / Payments", lines: [
      { no: "1", label: "Prior Year's Excess Credits", k: "s4_1", t: "man" },
      { no: "2", label: "Tax Payment/s for the Previous Quarter/s", k: "s4_2", t: "auto", src: "Previous quarters' returns marked Filed: amount paid" },
      { no: "3", label: "Creditable Tax Withheld for the Previous Quarter/s", k: "s4_3", t: "auto", src: "Creditable Withholding Taxes tab, earlier quarters" },
      { no: "4", label: "Creditable Tax Withheld per BIR Form No. 2307 for this Quarter", k: "s4_4", t: "auto", src: "Creditable Withholding Taxes tab, this quarter" },
      { no: "5", label: "Tax Paid in Return Previously Filed, if this is an amended return", k: "s4_5", t: "man", dis: !amended },
      { no: "6", label: "Foreign Tax Credits, if applicable", k: "s4_6", t: "man" },
      { no: "6A", label: "Other Credits / Payments (specify)", k: "s4_6a", t: "man" },
      { no: "7", label: "Total Tax Credits / Payments → Part II item 18", k: "s4_7", t: "calc", strong: true, f: (r) => `${r("s4_1")}+${r("s4_2")}+${r("s4_3")}+${r("s4_4")}+${r("s4_5")}+${r("s4_6")}+${r("s4_6a")}` },
    ]},
  };
}

const bg1702q = (i, c) => [["6", "TIN", i.tinText], ["7", "RDO Code", i.rdo], ["8", "Registered Name", i.name], ["9", "Registered Address", i.address], ["9A", "ZIP Code", i.zip], ["10", "Contact Number", i.contact], ["11", "Email Address", i.email]];

function Screen1702Q({ data, setData, postings, onNavigate, year, q, selectors }) {
  const info0 = companyInfo(data);
  const c0 = data.company || {};
  const rt = useReturn(data, setData, "form1702q", quarterKey(year, q));
  const b = useMemo(() => build1702q(data, postings, year, q), [data, postings, year, q]);
  const fyEnd = Number(c0.fiscalYearEndMonth) || 12;
  const info = { year, q, c: c0, fiscal: fyEnd !== 12, yearEnded: `${String(fyEnd).padStart(2, "0")}/${year}`, special: c0.incomeSpecialRate === "Yes" || c0.incomeTaxRegime === "Special rate" };
  const secs = sections1702q(b, info);
  const calc = b.calc;
  const rtView = { ...rt, form: { ...rt.form, m: { method: b.el.method, ...(rt.form.m || {}), ...(b.el.locked ? { method: b.el.method } : {}) } } };
  const isVat = isVatCompany(data);
  const vs = useMemo(() => (isVat && fyEnd === 12 ? computeVatSchedule(data, year, aggregateVatByYear(data)) : null), [data, year, isVat, fyEnd]);
  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1702Q — Quarterly Income Tax Return for Corporations, Partnerships and Other Non-Individual Taxpayers", version: "January 2018 (ENCS), eBIRForms v2018C", formNo: "1702Q",
    period: `${QN[q - 1]} Quarter, year ended ${info.yearEnded}`, company: info0,
    blocks: [sectionBlock(secs.header, calc, rt.form), infoBlock("Part I – Background Information", bg1702q(info0, c0)), sectionBlock(secs.part1x, calc, rtView.form),
      sectionBlock(secs.part2, calc, rt.form), paymentBlock(rt.form, "Part III – Details of Payment (items 27 to 30)"),
      sectionBlock(secs.sched1, calc, rt.form), sectionBlock(secs.sched2, calc, rt.form), sectionBlock(secs.sched3, calc, rt.form), sectionBlock(secs.sched4, calc, rt.form),
      quarterSummaryBlock(data, "form1702q", year, q, (k) => build1702q(data, postings, year, k), "s2_10", "i24")],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1702Q_Q${q}${year}_worksheet.pdf`,
  }));
  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}>
      <button className="io-btn" onClick={() => onNavigate("form2307")}>Open Creditable Withholding Taxes</button>
      <button className="io-btn" onClick={() => onNavigate("income")}>Open Income Statement</button>
      <WorksheetButtons hook={exp} />
    </div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={[["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"]]} />
    <ReturnStatusBar rt={rt} dueDefault={due1702q(b.rng.end)} dueNote="60 days after the quarter closes · verify weekends / holidays / eFPS / extensions">{selectors}</ReturnStatusBar>
    {calc.t("amended") === "Yes" && <Notice>Amended return: Schedule 4 item 5 is unlocked. Key in the figures of the original filing.</Notice>}
    <div className="tx-note" style={{ padding: "0 0 8px" }}>{isVat ? "Sales exclude 12% output VAT." : "Non-VAT: sales are the gross receipts as recorded."} NOLCO is deducted on the annual return only.</div>
    <FormSection section={secs.header} calc={calc} rt={rt} />
    <BackgroundBlock items={bg1702q(info0, c0)} onNavigate={onNavigate} />
    <FormSection section={secs.part1x} calc={calc} rt={rtView} />
    <FormSection section={secs.sched1} calc={calc} rt={rt} />
    <FormSection section={{ ...secs.sched2, headRight: <span className="tx-muted">Rate: <RateSelect value={rt.form.rate} dflt={corpRate(c0)} onChange={(v) => rt.patch({ rate: v })} /></span> }} calc={calc} rt={rt} />
    <UnmappedNote fig={b.fig} onNavigate={onNavigate} />
    <div className={b.mc.applies ? "" : "tx-dim"} style={b.mc.applies ? undefined : { opacity: 0.6 }}><FormSection section={secs.sched3} calc={calc} rt={rt} /></div>
    <FormSection section={secs.sched4} calc={calc} rt={rt} />
    <FormSection section={secs.part2} calc={calc} rt={rt} />
    <PaymentDetails rt={rt} title="Part III – Details of Payment (items 27 to 30)" firstItem={27} amountDue={calc.v.i24} />
    <IncomeRecon b={b} q={q} vatBase={vs ? round2(vs.quarters[q - 1].totalSales) : null} salesKey="s2_1" dedKey="s2_6" cwtKey="s4_4" taxableKey="s2_8" />
    <Declaration data={data} rt={rt} showTreasurer />
    <div className="rc-warn">Rate: 25%, or 20% when net taxable income is ₱5,000,000 or less and total assets (excluding land) are ₱100,000,000 or less (CREATE). Proprietary educational institutions / hospitals (10%), resident foreign corporations and special-rate entities use Schedule 1 / another rate. MCIT is 2% (the 1% CREATE relief ended June 30, 2023).</div>
  </>);
}

export function RateSelect({ value, dflt, onChange }) {
  const v = value === undefined || value === "" ? dflt : num(value);
  return (
    <select className="io-select" value={String(v)} onChange={(e) => onChange(Number(e.target.value) === dflt ? "" : Number(e.target.value))}>
      <option value="0.25">25%</option><option value="0.2">20%</option>
    </select>
  );
}

/* ------------------------------ Page ------------------------------ */

export function IncomeQuarterlyPage({ data, setData, postings, onNavigate }) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [q, setQ] = useState(Math.min(3, Math.floor(now.getMonth() / 3) + 1));
  const c0 = data.company || {};
  const indiv = c0.taxpayerType === "Individual";
  const otherSaved = Object.keys(data[indiv ? "form1702q" : "form1701q"] || {}).length;
  const selectors = (<>
    <select className="io-select" value={q} onChange={(e) => setQ(Number(e.target.value))}>{[1, 2, 3, 4].map((x) => <option key={x} value={x}>{QN[x - 1]} Quarter</option>)}</select>
    <YearSelect years={getAvailableYears(data)} value={year} onChange={setYear} />
  </>);
  const props = { data, setData, postings, onNavigate, year, q, selectors };
  return (
    <div>
      <TaxStyle />
      <SectionHeader icon={FileText} title={indiv ? "BIR Form 1701Q – Quarterly Income Tax Return (Individuals)" : "BIR Form 1702Q – Quarterly Income Tax Return (Corporations)"}
        subtitle={`Follows Company Details → Taxpayer Type (${indiv ? "Individual" : "Non-Individual"}). Q1 to Q3 only. Encoding worksheet for eBIRForms / eFPS — no .dat file.`} />
      {otherSaved > 0 && <div className="tx-note" style={{ padding: "0 0 8px" }}>{otherSaved} saved {indiv ? "1702Q" : "1701Q"} return(s) from an earlier taxpayer type are kept.</div>}
      {q === 4 ? (<>
        <div className="tx-statusbar"><div className="tx-selectors">{selectors}</div></div>
        <Notice tone="info">Q4 is reported in the Annual Income Tax Return ({indiv ? "1701" : "1702-RT"}). <button className="tx-link" onClick={() => onNavigate("incomeannual")}>Open the annual return</button></Notice>
      </>) : indiv ? <Screen1701Q {...props} /> : <Screen1702Q {...props} />}
    </div>
  );
}

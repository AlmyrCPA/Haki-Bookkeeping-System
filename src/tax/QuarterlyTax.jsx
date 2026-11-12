// BIR Form 2550Q (January 2023 ENCS) Quarterly VAT Return / BIR Form 2551Q (January 2018 ENCS)
// Quarterly Percentage Tax Return. The tab shows the form that matches Company Details → VAT Status;
// both are saved per quarter so nothing is lost when the status changes.
import React, { useMemo, useState } from "react";
import { FileText, Plus } from "lucide-react";
import {
  Field, NumberField, SectionHeader, getAvailableYears, num, round2, fmtPlain, fmtParen, uid, DelBtn,
  computeVatSchedule, aggregateVatByYear, purchaseSlspiField, isVatCompany, inRange, EntryModalShell,
} from "../BookkeepingApp";
import {
  TaxStyle, makeCalc, useReturn, FormSection, ReturnStatusBar, BackgroundBlock, CompanyPrereqWarning, PaymentDetails,
  Declaration, companyInfo, quarterKey, quarterRange, sectionBlock, infoBlock, paymentBlock, signatureBlock, useWorksheetExport,
  WorksheetButtons, ExportStatus, YearSelect, Notice, ReconStrip, AutoAmount,
} from "./shared";

const QUARTER_NAMES = ["1st", "2nd", "3rd", "4th"];
export const due2550q = (year, q) => (q === 4 ? new Date(year + 1, 0, 25) : new Date(year, q * 3, 25));

/* ------------------------------ 2550Q data ------------------------------ */

const isNonResidentType = (vatType) => /non-?resident/i.test(String(vatType || ""));

export function vatFiguresForQuarter(data, year, q) {
  const { start, end } = quarterRange(year, q);
  const govNames = new Set((data.customers || []).filter((c) => c.government === "Yes").map((c) => String(c.name || "").toLowerCase()));
  const s = { pvtA: 0, pvtB: 0, govA: 0, govB: 0, zero: 0, exempt: 0, domA: 0, domB: 0, nrA: 0, nrB: 0, impA: 0, impB: 0, noInput: 0 };
  (data.sales || []).forEach((r) => {
    if (!inRange(r.date, start, end)) return;
    const gov = govNames.has(String(r.customer || "").toLowerCase());
    if (gov) { s.govA += num(r.vatable); s.govB += num(r.outputVat); } else { s.pvtA += num(r.vatable); s.pvtB += num(r.outputVat); }
    s.zero += num(r.zeroRated); s.exempt += num(r.exempt);
  });
  (data.purchases || []).forEach((r) => {
    if (!inRange(r.date, start, end)) return;
    const field = purchaseSlspiField(data, r);
    const vatBearing = ["Other than capital goods", "Services", "Capital goods"].includes(field);
    if (vatBearing && isNonResidentType(r.vatType)) { s.nrA += num(r.vatable); s.nrB += num(r.inputVat); s.noInput += num(r.nonvat); }
    else if (vatBearing) { s.domA += num(r.vatable); s.domB += num(r.inputVat); s.noInput += num(r.nonvat); }
    else s.noInput += num(r.vatable) + num(r.nonvat);
  });
  (data.importation || []).forEach((r) => {
    if (!inRange(r.date, start, end)) return;
    s.impA += num(r.dutiableValue); s.impB += num(r.vatPaid);
  });
  Object.keys(s).forEach((k) => { s[k] = round2(s[k]); });
  return s;
}

const SCHED1_COLS = [["date", "A Date Purchased / Imported", "date"], ["source", "B Source (D/I)", "text"], ["desc", "C Description", "text"],
  ["amount", "D Amount Exceeding ₱1M", "num"], ["balPrev", "E Balance of Input Tax from Previous Period", "num"], ["estLife", "F Estimated Life (months)", "num"],
  ["recLife", "G Recognized Life (months)", "num"], ["remLife", "H Remaining Life", "num"], ["allowable", "I Allowable Input Tax for the Period", "num"], ["balNext", "J Balance to Carry to Next Period", "num"]];
const SCHED3_COLS = [["period", "Period Covered", "text"], ["agent", "Name of Withholding Agent", "agent"], ["incomePayment", "Income Payment", "num"], ["taxWithheld", "Total Tax Withheld", "num"]];
const SCHED4_COLS = [["period", "Period Covered", "text"], ["miller", "Name of Miller", "text"], ["taxpayer", "Name of Taxpayer", "text"], ["orNo", "Official Receipt No.", "text"], ["amount", "Amount Paid", "num"]];
const schedSum = (rows, k) => round2((rows || []).reduce((s, r) => s + num(r[k]), 0));

export function compute2550q(form, figs, sched) {
  const c = makeCalc(form);
  const s1 = form.sched1 || [], s3 = form.sched3 || [], s4 = form.sched4 || [];
  c.a("i31a", figs.pvtA); c.a("i31b", figs.pvtB);
  c.a("i32a", figs.govA); c.a("i32b", figs.govB);
  c.a("i33a", figs.zero); c.a("i34a", figs.exempt);
  c.set("i35a", c.v.i31a + c.v.i32a + c.v.i33a + c.v.i34a);
  c.set("i35b", c.v.i31b + c.v.i32b);
  c.a("i36b", Math.max(0, num(sched && sched.S)));
  c.set("i37b", schedSum(s1, "balPrev"));
  c.man("i38b"); c.man("i39b"); c.man("i40b");
  c.set("i41b", c.v.i36b + c.v.i37b + c.v.i38b + c.v.i39b + c.v.i40b);
  c.a("i42a", figs.domA); c.a("i42b", figs.domB);
  c.a("i43a", figs.nrA); c.a("i43b", figs.nrB);
  c.a("i44a", figs.impA); c.a("i44b", figs.impB);
  c.man("i45a"); c.man("i45b");
  c.a("i46a", figs.noInput);
  c.man("i47a");
  c.set("i48a", c.v.i42a + c.v.i43a + c.v.i44a + c.v.i45a + c.v.i46a + c.v.i47a);
  c.set("i48b", c.v.i42b + c.v.i43b + c.v.i44b + c.v.i45b);
  c.set("i49b", c.v.i41b + c.v.i48b);
  c.set("i50b", schedSum(s1, "balNext"));
  // Schedule 2: input tax attributable to VAT-exempt sales.
  const direct = c.m("s2direct");
  const notDirect = round2(c.v.i48b - direct);
  const ratable = c.v.i35a ? round2((c.v.i34a / c.v.i35a) * notDirect) : 0;
  c.set("s2notDirect", notDirect); c.set("s2ratable", ratable);
  c.a("i51b", direct + ratable);
  c.man("i52b"); c.man("i53b");
  c.set("i54b", c.v.i50b + c.v.i51b + c.v.i52b + c.v.i53b);
  c.set("i55b", c.v.i49b - c.v.i54b);
  c.set("i56b", c.v.i35b - c.v.i55b);
  // Part II
  c.set("i15", c.v.i56b);
  c.set("i16", schedSum(s3, "taxWithheld"));
  c.set("i17", schedSum(s4, "amount"));
  c.set("i18", c.t("amended") === "Yes" ? c.m("i18") : 0);
  c.man("i19");
  c.set("i20", c.v.i16 + c.v.i17 + c.v.i18 + c.v.i19);
  c.set("i21", c.v.i15 - c.v.i20);
  c.man("i22"); c.man("i23"); c.man("i24");
  c.set("i25", c.v.i22 + c.v.i23 + c.v.i24);
  c.set("i26", c.v.i21 + c.v.i25);
  return c;
}

function sections2550q(c, info) {
  const amended = c.t("amended") === "Yes";
  const AB = ["A Sales / Purchases", "B Output / Input Tax"];
  return {
    header: { title: "Return Header (items 1 to 7)", lines: [
      { no: "1", label: "For the", k: "_cal", t: "ro", text: "Calendar" },
      { no: "2", label: "Year Ended (MM/YYYY)", k: "_ye", t: "ro", text: `12/${info.year}` },
      { no: "3", label: "Quarter", k: "_q", t: "ro", text: `${QUARTER_NAMES[info.q - 1]} Quarter` },
      { no: "4", label: "Return Period (From – To)", k: "_rp", t: "ro", text: info.periodText },
      { no: "5", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "6", label: "Short Period Return?", k: "shortPeriod", t: "yesno", dflt: "No" },
      { no: "7", label: "Number of Sheet/s Attached", k: "sheetsAttached", t: "int" },
    ]},
    relief: { title: "Part I (items 14 / 14A)", lines: [
      { no: "14", label: "Are you availing of tax relief under Special Law or International Tax Treaty?", k: "taxRelief", t: "yesno", dflt: info.c.taxRelief || "No" },
      { no: "14A", label: "If yes, specify", k: "taxReliefSpecify", t: "text", dflt: info.c.taxReliefSpecify || "", hide: c.t("taxRelief", info.c.taxRelief || "No") !== "Yes" },
    ]},
    part2: { title: "Part II – Total Tax Payable", lines: [
      { no: "15", label: "Net VAT Payable / (Excess Input Tax) (from Part IV, item 56)", k: "i15", t: "calc", due: true },
      { heading: "Less: Tax Credits / Payments" },
      { no: "16", label: "Creditable VAT Withheld (from Part V – Schedule 3, column D)", k: "i16", t: "calc" },
      { no: "17", label: "Advance VAT Payments (from Part V – Schedule 4)", k: "i17", t: "calc" },
      { no: "18", label: "VAT paid in return previously filed, if this is an amended return", k: "i18", t: "man", dis: !amended },
      { no: "19", label: "Other Credits / Payment (specify)", k: "i19", t: "man" },
      { no: "", label: "Other credits: specify", k: "i19Specify", t: "text", hide: !c.v.i19 },
      { no: "20", label: "Total Tax Credits / Payment (sum of items 16 to 19)", k: "i20", t: "calc", f: (r) => `${r("i16")}+${r("i17")}+${r("i18")}+${r("i19")}` },
      { no: "21", label: "Tax Still Payable / (Excess Credits) (item 15 less item 20)", k: "i21", t: "calc", due: true, f: (r) => `${r("i15")}-${r("i20")}` },
      { heading: "Add: Penalties (manual)" },
      { no: "22", label: "Surcharge", k: "i22", t: "man" },
      { no: "23", label: "Interest", k: "i23", t: "man" },
      { no: "24", label: "Compromise", k: "i24", t: "man" },
      { no: "25", label: "Total Penalties (sum of items 22 to 24)", k: "i25", t: "calc", f: (r) => `${r("i22")}+${r("i23")}+${r("i24")}` },
      { no: "26", label: "TOTAL AMOUNT PAYABLE / (Excess Credits) (sum of items 21 and 25)", k: "i26", t: "calc", strong: true, due: true, f: (r) => `${r("i21")}+${r("i25")}` },
    ]},
    part4: { title: "Part IV – Details of VAT Computation", cols: AB, lines: [
      { heading: "Sales and output tax" },
      { no: "31", label: "VATable Sales — Private (Sales Journal, non-Government customers)", k: ["i31a", "i31b"], t: "auto", src: "Sales Journal: VATable sales and output VAT to customers not flagged Government" },
      { no: "32", label: "VATable Sales — Government (customers flagged Government)", k: ["i32a", "i32b"], t: "auto", src: "Sales Journal: customers flagged Government in Customers Master" },
      { no: "33", label: "Zero-Rated Sales", k: ["i33a", null], t: ["auto", null], src: "Sales Journal: zero-rated" },
      { no: "34", label: "Exempt Sales", k: ["i34a", null], t: ["auto", null], src: "Sales Journal: VAT-exempt" },
      { no: "35", label: "Total Sales & Output Tax Due (sum of items 31 to 34)", k: ["i35a", "i35b"], t: "calc", strong: true, f: [(r) => `${r("i31a")}+${r("i32a")}+${r("i33a")}+${r("i34a")}`, (r) => `${r("i31b")}+${r("i32b")}`] },
      { heading: "Less: Allowable Input Tax" },
      { no: "36", label: "Input Tax Carried Over from Previous Quarter", k: [null, "i36b"], t: [null, "auto"], src: "VAT Schedule: carry-over into this quarter (previous quarter's excess input tax)" },
      { no: "37", label: "Input Tax Deferred on Capital Goods Exceeding ₱1 Million from Previous Quarter (Schedule 1, column E)", k: [null, "i37b"], t: [null, "calc"] },
      { no: "38", label: "Transitional Input Tax", k: [null, "i38b"], t: [null, "man"] },
      { no: "39", label: "Presumptive Input Tax", k: [null, "i39b"], t: [null, "man"] },
      { no: "40", label: "Others (specify)", k: [null, "i40b"], t: [null, "man"] },
      { no: "41", label: "Total (sum of items 36 to 40)", k: [null, "i41b"], t: [null, "calc"], strong: true, f: [null, (r) => `${r("i36b")}+${r("i37b")}+${r("i38b")}+${r("i39b")}+${r("i40b")}`] },
      { heading: "Current Transactions" },
      { no: "42", label: "Domestic Purchases", k: ["i42a", "i42b"], t: "auto", src: "Purchase Journal: VATable purchases whose VAT Type SLSPI Field is Services, Capital goods or Other than capital goods" },
      { no: "43", label: "Services Rendered by Non-residents", k: ["i43a", "i43b"], t: "auto", src: "Purchase Journal: VAT Type naming non-resident services" },
      { no: "44", label: "Importations", k: ["i44a", "i44b"], t: "auto", src: "Importation Ledger: dutiable value and VAT paid" },
      { no: "45", label: "Others (specify)", k: ["i45a", "i45b"], t: "man" },
      { no: "46", label: "Domestic Purchases with No Input Tax", k: ["i46a", null], t: ["auto", null], src: "Purchase Journal: Non-VAT amounts and purchases on Exempt / Zero-Rated VAT Types" },
      { no: "47", label: "VAT-Exempt Importations", k: ["i47a", null], t: ["man", null] },
      { no: "48", label: "Total Current Purchases / Input Tax (sum of items 42 to 47)", k: ["i48a", "i48b"], t: "calc", strong: true, f: [(r) => `${r("i42a")}+${r("i43a")}+${r("i44a")}+${r("i45a")}+${r("i46a")}+${r("i47a")}`, (r) => `${r("i42b")}+${r("i43b")}+${r("i44b")}+${r("i45b")}`] },
      { no: "49", label: "Total Available Input Tax (sum of items 41B and 48B)", k: [null, "i49b"], t: [null, "calc"], strong: true, f: [null, (r) => `${r("i41b")}+${r("i48b")}`] },
      { heading: "Less: Deductions from Input Tax" },
      { no: "50", label: "Input Tax on Purchases of Capital Goods Exceeding ₱1 Million Deferred for the Succeeding Period (Schedule 1, column J)", k: [null, "i50b"], t: [null, "calc"] },
      { no: "51", label: "Input Tax Attributable to VAT-Exempt Sales (Schedule 2)", k: [null, "i51b"], t: [null, "auto"], src: "Schedule 2: direct + ratable portion" },
      { no: "52", label: "VAT Refund / TCC Claimed", k: [null, "i52b"], t: [null, "man"] },
      { no: "53", label: "Others (specify)", k: [null, "i53b"], t: [null, "man"] },
      { no: "54", label: "Total (sum of items 50 to 53)", k: [null, "i54b"], t: [null, "calc"], f: [null, (r) => `${r("i50b")}+${r("i51b")}+${r("i52b")}+${r("i53b")}`] },
      { no: "55", label: "Total Allowable Input Tax (item 49 less item 54)", k: [null, "i55b"], t: [null, "calc"], strong: true, f: [null, (r) => `${r("i49b")}-${r("i54b")}`] },
      { no: "56", label: "Net VAT Payable / (Excess Input Tax) (item 35B less item 55)", k: [null, "i56b"], t: [null, "calc"], strong: true, due: true, f: [null, (r) => `${r("i35b")}-${r("i55b")}`] },
    ]},
    sched2: { title: "Schedule 2 – Input Tax Attributable to VAT-Exempt Sales", lines: [
      { no: "", label: "Input tax directly attributable to VAT-exempt sale", k: "s2direct", t: "man" },
      { no: "", label: "Input tax not directly attributable (total current input tax, item 48B, less the direct amount)", k: "s2notDirect", t: "calc" },
      { no: "", label: "Ratable portion (Exempt sales ÷ Total sales × not directly attributable)", k: "s2ratable", t: "calc" },
      { no: "", label: "Total → item 51", k: "i51b", t: "calc", strong: true },
    ]},
  };
}

/* ------------------------------ 2551Q data ------------------------------ */

export const ptDefaultRate = (year, q) => {
  const idx = year * 4 + (q - 1);
  return idx >= 2020 * 4 + 2 && idx <= 2023 * 4 + 1 ? 0.01 : 0.03; // July 2020 – June 2023: 1% (CREATE)
};
export function grossReceiptsForQuarter(data, year, q) {
  const { start, end } = quarterRange(year, q);
  return round2((data.sales || []).filter((r) => inRange(r.date, start, end)).reduce((s, r) => s + num(r.vatable) + num(r.exempt) + num(r.zeroRated), 0));
}

export function compute2551q(form, gross, rateDflt) {
  const c = makeCalc(form);
  const rows = form.ptRows || [];
  c.a("pt1amt", gross);
  const rate1 = form.ptRate1 === undefined || form.ptRate1 === "" ? rateDflt : num(form.ptRate1);
  c.set("pt1due", c.v.pt1amt * rate1);
  let total = c.v.pt1due;
  rows.forEach((r) => { total += round2(num(r.amount) * num(r.rate)); });
  c.set("i7", total);
  c.set("i14", c.v.i7);
  c.man("i15");
  c.set("i16", c.t("amended") === "Yes" ? c.m("i16") : 0);
  c.man("i17");
  c.set("i18", c.v.i15 + c.v.i16 + c.v.i17);
  c.set("i19", c.v.i14 - c.v.i18);
  c.man("i20"); c.man("i21"); c.man("i22");
  c.set("i23", c.v.i20 + c.v.i21 + c.v.i22);
  c.set("i24", c.v.i19 + c.v.i23);
  c.rate1 = rate1;
  return c;
}

function sections2551q(c, info) {
  const amended = c.t("amended") === "Yes";
  return {
    header: { title: "Return Header (items 1 to 5)", lines: [
      { no: "1", label: "For the", k: "_cal", t: "ro", text: "Calendar" },
      { no: "2", label: "Year Ended (MM/YYYY)", k: "_ye", t: "ro", text: `12/${info.year}` },
      { no: "3", label: "Quarter", k: "_q", t: "ro", text: `${QUARTER_NAMES[info.q - 1]} Quarter` },
      { no: "4", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "5", label: "Number of Sheet/s Attached", k: "sheetsAttached", t: "int" },
    ]},
    relief: { title: "Part I (items 12 / 13)", lines: [
      { no: "12", label: "Are you availing of tax relief under Special Law or International Tax Treaty?", k: "taxRelief", t: "yesno", dflt: info.c.taxRelief || "No" },
      { no: "12A", label: "If yes, specify", k: "taxReliefSpecify", t: "text", dflt: info.c.taxReliefSpecify || "", hide: c.t("taxRelief", info.c.taxRelief || "No") !== "Yes" },
      { no: "13", label: "Income tax rate (individuals only)", k: "rateElection", t: "select", opts: ["Graduated income tax rate on net taxable income", "8% income tax rate on gross sales/receipts/others"], dflt: info.c.taxRateElection === "8%" ? "8% income tax rate on gross sales/receipts/others" : "Graduated income tax rate on net taxable income", hide: !info.indiv },
    ]},
    part2: { title: "Part II – Total Tax Payable", lines: [
      { no: "14", label: "Total Tax Due (from Schedule 1, item 7)", k: "i14", t: "calc" },
      { heading: "Less: Tax Credits / Payments" },
      { no: "15", label: "Creditable Percentage Tax Withheld per BIR Form No. 2307", k: "i15", t: "man", note: "Use “Pick from Creditable Withholding Taxes” to total the quarter's 2307s." },
      { no: "16", label: "Tax Paid in Return Previously Filed, if this is an amended return", k: "i16", t: "man", dis: !amended },
      { no: "17", label: "Other Tax Credit / Payment (specify)", k: "i17", t: "man" },
      { no: "", label: "Other tax credit: specify", k: "i17Specify", t: "text", hide: !c.v.i17 },
      { no: "18", label: "Total Tax Credits / Payments (sum of items 15 to 17)", k: "i18", t: "calc", f: (r) => `${r("i15")}+${r("i16")}+${r("i17")}` },
      { no: "19", label: "Tax Still Payable / (Overpayment) (item 14 less item 18)", k: "i19", t: "calc", due: true, f: (r) => `${r("i14")}-${r("i18")}` },
      { heading: "Add: Penalties (manual)" },
      { no: "20", label: "Surcharge", k: "i20", t: "man" },
      { no: "21", label: "Interest", k: "i21", t: "man" },
      { no: "22", label: "Compromise", k: "i22", t: "man" },
      { no: "23", label: "Total Penalties (sum of items 20 to 22)", k: "i23", t: "calc", f: (r) => `${r("i20")}+${r("i21")}+${r("i22")}` },
      { no: "24", label: "TOTAL AMOUNT PAYABLE / (Overpayment) (sum of items 19 and 23)", k: "i24", t: "calc", strong: true, due: true, f: (r) => `${r("i19")}+${r("i23")}` },
      { no: "", label: "If overpayment, mark one", k: "overpayChoice", t: "select", opts: ["To be refunded", "To be issued a Tax Credit Certificate"], hide: !(c.v.i24 < 0) },
    ]},
  };
}

/* ------------------------------ Schedules UI ------------------------------ */

function ScheduleTable({ title, rows, cols, onChange, totalKey, totalLabel, agentOptions, note }) {
  const set = (id, k, v) => onChange(rows.map((r) => r.id === id ? { ...r, [k]: v } : r));
  return (
    <div className="tx-card">
      <div className="tx-card-head"><span>{title}</span><button className="io-btn" onClick={() => onChange([...rows, { id: uid() }])}><Plus size={13} /> Add row</button></div>
      {note && <div className="tx-note">{note}</div>}
      <div className="table-scroll">
        <table className="ledger-table tx-table">
          <thead><tr>{cols.map(([k, h, t]) => <th key={k} className={t === "num" ? "num-head" : ""} style={{ minWidth: t === "num" ? 120 : 140 }}>{h}</th>)}<th style={{ width: 36 }}></th></tr></thead>
          <tbody>
            {rows.length === 0 && <tr><td colSpan={cols.length + 1} className="empty-row">No rows.</td></tr>}
            {rows.map((r) => (
              <tr key={r.id}>
                {cols.map(([k, , t]) => (
                  <td key={k}>
                    {t === "num" ? <NumberField value={r[k]} onChange={(v) => set(r.id, k, v)} align="right" />
                      : t === "date" ? <Field type="date" value={r[k]} onChange={(v) => set(r.id, k, v)} />
                      : t === "agent" ? <Field type="combo" options={agentOptions || []} freeText value={r[k]} onChange={(v) => set(r.id, k, v)} />
                      : <Field value={r[k]} onChange={(v) => set(r.id, k, v)} />}
                  </td>
                ))}
                <td className="text-center"><DelBtn onClick={() => onChange(rows.filter((x) => x.id !== r.id))} /></td>
              </tr>
            ))}
          </tbody>
          {totalKey && rows.length > 0 && <tfoot><tr><td colSpan={cols.length - 1} className="totals-label">{totalLabel}</td><td className="num">{fmtPlain(schedSum(rows, totalKey))}</td><td></td></tr></tfoot>}
        </table>
      </div>
    </div>
  );
}

const schedBlock = (title, cols, rows, sheet) => ({
  type: "table", title, sheet, head: cols.map(([, h]) => h),
  body: (rows || []).map((r) => cols.map(([k, , t]) => (t === "num" ? num(r[k]) : r[k] || ""))),
  numCols: cols.map(([, , t], i) => (t === "num" ? i : -1)).filter((i) => i >= 0),
});

/* ------------------------------ Page ------------------------------ */

export function QuarterlyTaxPage({ data, setData, postings, onNavigate }) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [q, setQ] = useState(Math.floor(now.getMonth() / 3) + 1);
  const vatByStatus = isVatCompany(data);
  const [formOverride, setFormOverride] = useState(null);
  const which = formOverride || (vatByStatus ? "2550Q" : "2551Q");
  const other = which === "2550Q" ? "2551Q" : "2550Q";
  const info0 = companyInfo(data);
  const props = { data, setData, postings, onNavigate, year, q, info0 };
  return (
    <div>
      <TaxStyle />
      <SectionHeader icon={FileText} title={which === "2550Q" ? "BIR Form 2550Q – Quarterly VAT Return" : "BIR Form 2551Q – Quarterly Percentage Tax Return"}
        subtitle={which === "2550Q" ? "January 2023 (ENCS). The monthly 2550M was discontinued from January 1, 2023. Encoding worksheet for eBIRForms / eFPS — no .dat file." : "January 2018 (ENCS). The monthly 2551M is discontinued. Encoding worksheet for eBIRForms / eFPS — no .dat file."} />
      <div className="tx-note" style={{ padding: "0 0 10px" }}>
        Company Details → VAT Status: <strong>{vatByStatus ? "VAT Registered" : "Non-VAT"}</strong>.{" "}
        <button className="tx-link" onClick={() => setFormOverride(formOverride ? null : other)}>{formOverride ? `Back to the ${vatByStatus ? "2550Q" : "2551Q"}` : `Open the ${other} instead`}</button>
        {formOverride && " (for a company that changed registration mid-year; each form's returns are saved separately)."}
      </div>
      {which === "2550Q" ? <Vat2550Q {...props} setYear={setYear} setQ={setQ} /> : <Pt2551Q {...props} setYear={setYear} setQ={setQ} />}
    </div>
  );
}

function PeriodSelectors({ data, year, setYear, q, setQ }) {
  return (<>
    <select className="io-select" value={q} onChange={(e) => setQ(Number(e.target.value))}>{[1, 2, 3, 4].map((x) => <option key={x} value={x}>{QUARTER_NAMES[x - 1]} Quarter</option>)}</select>
    <YearSelect years={getAvailableYears(data)} value={year} onChange={setYear} />
  </>);
}

const bg2550 = (i) => [["8", "TIN", i.tinText], ["9", "RDO Code", i.rdo], ["10", "Taxpayer's Name", i.name], ["11", "Registered Address", i.address], ["11A", "ZIP Code", i.zip], ["12", "Contact Number", i.contact], ["13", "Email Address", i.email]];
const bg2551 = (i) => [["6", "TIN", i.tinText], ["7", "RDO Code", i.rdo], ["8", "Taxpayer's Name", i.name], ["9", "Registered Address", i.address], ["9A", "ZIP Code", i.zip], ["10", "Contact Number", i.contact], ["11", "Email Address", i.email]];
const PREREQ = [["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"]];

function Vat2550Q({ data, setData, postings, onNavigate, year, q, setYear, setQ, info0 }) {
  const rt = useReturn(data, setData, "form2550q", quarterKey(year, q), { sched1: [], sched3: [], sched4: [] });
  const form = rt.form;
  const figs = useMemo(() => vatFiguresForQuarter(data, year, q), [data.sales, data.purchases, data.importation, data.customers, data.vatTypes, year, q]);
  const vs = useMemo(() => computeVatSchedule(data, year, aggregateVatByYear(data)), [data, year]);
  const Q = vs.quarters[q - 1];
  const calc = useMemo(() => compute2550q(form, figs, Q), [form, figs, Q]);
  const { start, end } = quarterRange(year, q);
  const info = { year, q, c: data.company || {}, periodText: `${String(start.getMonth() + 1).padStart(2, "0")}/01/${year} – ${String(end.getMonth() + 1).padStart(2, "0")}/${end.getDate()}/${year}` };
  const secs = sections2550q(calc, info);
  const setSched = (k) => (rows) => rt.patch({ [k]: rows });
  const agentOptions = (data.customers || []).map((c) => c.name).filter(Boolean);

  const gl = useMemo(() => {
    let output = 0, input = 0;
    (postings || []).forEach((p) => {
      if (!inRange(p.date, start, end)) return;
      if (p.code === "2200") output += p.credit - p.debit;
      if (p.code === "1400") input += p.debit - p.credit;
    });
    return { output: round2(output), input: round2(input) };
  }, [postings, year, q]);
  const carryIn = Math.max(0, Q.S); // a negative carry-over is an unpaid balance, not input tax
  const schedAllowable = round2(carryIn + Q.totalInputVat - Q.T);
  const schedNet = round2(Q.totalOutputVat - schedAllowable);

  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 2550Q — Quarterly Value-Added Tax (VAT) Return", version: "January 2023 (ENCS)", formNo: "2550Q",
    period: `${QUARTER_NAMES[q - 1]} Quarter ${year} (${info.periodText})`, company: info0,
    blocks: [
      sectionBlock(secs.header, calc, form), infoBlock("Part I – Background Information", bg2550(info0)), sectionBlock(secs.relief, calc, form),
      sectionBlock(secs.part2, calc, form), paymentBlock(form, "Part III – Details of Payment (items 27 to 30)"), sectionBlock(secs.part4, calc, form),
      schedBlock("Schedule 1 – Amortized Input Tax from Capital Goods", SCHED1_COLS, form.sched1, "Schedule 1"),
      sectionBlock(secs.sched2, calc, form),
      schedBlock("Schedule 3 – Creditable VAT Withheld", SCHED3_COLS, form.sched3, "Schedule 3"),
      schedBlock("Schedule 4 – Advance VAT Payment", SCHED4_COLS, form.sched4, "Schedule 4"),
    ], signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_2550Q_Q${q}${year}_worksheet.pdf`,
  }));

  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}><WorksheetButtons hook={exp} /></div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={PREREQ} />
    <ReturnStatusBar rt={rt} dueDefault={due2550q(year, q)} dueNote="25 days after quarter-end · verify weekends / holidays"><PeriodSelectors data={data} year={year} setYear={setYear} q={q} setQ={setQ} /></ReturnStatusBar>
    {calc.t("amended") === "Yes" && <Notice>Amended return: item 18 is unlocked. Key in the figures of the original filing.</Notice>}
    {!(data.customers || []).some((c) => c.government === "Yes") && <Notice tone="info">No customer is flagged Government in Customers Master, so all VATable sales go to item 31. Set the Government column to split item 32.</Notice>}
    <FormSection section={secs.header} calc={calc} rt={rt} />
    <BackgroundBlock items={bg2550(info0)} onNavigate={onNavigate} />
    <FormSection section={secs.relief} calc={calc} rt={rt} />
    <FormSection section={secs.part4} calc={calc} rt={rt} />
    <FormSection section={secs.part2} calc={calc} rt={rt} />
    <PaymentDetails rt={rt} title="Part III – Details of Payment (items 27 to 30)" firstItem={27} amountDue={calc.v.i26} />
    <div className="sub-block-title">Part V – Schedules</div>
    <ScheduleTable title="Schedule 1 – Amortized Input Tax from Capital Goods (column E total → item 37; column J total → item 50)" rows={form.sched1 || []} cols={SCHED1_COLS} onChange={setSched("sched1")}
      note="Manual only: kept on the form for legacy balances. Haki's VAT Schedule claims capital goods input tax outright." />
    <FormSection section={secs.sched2} calc={calc} rt={rt} />
    <ScheduleTable title="Schedule 3 – Creditable VAT Withheld (total → item 16)" rows={form.sched3 || []} cols={SCHED3_COLS} onChange={setSched("sched3")} totalKey="taxWithheld" totalLabel="Total" agentOptions={agentOptions} />
    <ScheduleTable title="Schedule 4 – Advance VAT Payment (total → item 17)" rows={form.sched4 || []} cols={SCHED4_COLS} onChange={setSched("sched4")} totalKey="amount" totalLabel="Total" />
    <ReconStrip title="VAT Schedule check" colA="2550Q" colB="VAT Schedule / GL" items={[
      { label: "Item 35B output tax vs VAT Schedule output VAT for the quarter", a: calc.v.i35b, b: round2(Q.totalOutputVat) },
      { label: "Item 55 allowable input tax vs VAT Schedule (carry-over + input VAT − input allocable to exempt)", a: calc.v.i55b, b: schedAllowable },
      { label: "Item 56 net VAT payable vs VAT Schedule output VAT less allowable input tax", a: calc.v.i56b, b: schedNet },
      { label: "Item 35B vs GL 2200 Output VAT (credit − debit)", a: calc.v.i35b, b: gl.output },
      { label: "Items 42B + 43B vs GL 1400 Input VAT (debit − credit)", a: round2(calc.v.i42b + calc.v.i43b), b: gl.input },
    ]} causes="manual entries in Part V (Schedules 1 and 2, items 38 to 40, 45, 47, 52, 53); an override of item 36 or any auto line; manual General Journal entries to 2200 / 1400; input VAT on Exempt / Zero-Rated / Non-VAT purchases; or entries dated outside the quarter." />
    {Q.S < -0.004 && <div className="tx-note">The VAT Schedule carries an unpaid balance of {fmtPlain(-Q.S)} from earlier quarters into this one. It is not input tax, so item 36 leaves it out; record payments in the VAT Schedule (PAID) to clear it.</div>}
    <Declaration data={data} rt={rt} />
  </>);
}

function Pick2307Modal({ data, year, q, onCancel, onApply }) {
  const { start, end } = quarterRange(year, q);
  const rows = (data.form2307 || []).filter((r) => inRange(r.date, start, end));
  const [picked, setPicked] = useState(new Set());
  const total = round2(rows.filter((r) => picked.has(r.id)).reduce((s, r) => s + num(r.cwt), 0));
  return (
    <EntryModalShell title="Creditable percentage tax withheld (2307s this quarter)" submitLabel={`Use ${fmtPlain(total)}`} onCancel={onCancel} onSubmit={() => onApply(total)}>
      <p className="section-head-note">Tick the 2307s for percentage tax withheld. Their total goes to item 15.</p>
      <div className="table-scroll" style={{ maxHeight: 340 }}>
        <table className="ledger-table">
          <thead><tr><th></th><th>Date</th><th>Withholding agent</th><th>ATC</th><th className="num-head">Income payment</th><th className="num-head">Tax withheld</th></tr></thead>
          <tbody>
            {rows.length === 0 && <tr><td colSpan={6} className="empty-row">No Creditable Withholding Taxes rows dated in this quarter.</td></tr>}
            {rows.map((r) => (
              <tr key={r.id} onClick={() => setPicked((s) => { const n = new Set(s); n.has(r.id) ? n.delete(r.id) : n.add(r.id); return n; })} style={{ cursor: "pointer" }}>
                <td className="text-center"><input type="checkbox" className="row-check" readOnly checked={picked.has(r.id)} /></td>
                <td>{r.date}</td><td>{r.customerName || [r.surname, r.firstName].filter(Boolean).join(", ")}</td><td>{r.atc}</td>
                <td className="num">{fmtPlain(r.incomePayment)}</td><td className="num">{fmtPlain(r.cwt)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </EntryModalShell>
  );
}

function Pt2551Q({ data, setData, onNavigate, year, q, setYear, setQ, info0 }) {
  const rt = useReturn(data, setData, "form2551q", quarterKey(year, q), { ptRows: [] });
  const form = rt.form;
  const gross = useMemo(() => grossReceiptsForQuarter(data, year, q), [data.sales, year, q]);
  const rateDflt = ptDefaultRate(year, q);
  const calc = useMemo(() => compute2551q(form, gross, rateDflt), [form, gross, rateDflt]);
  const info = { year, q, c: data.company || {}, indiv: info0.indiv };
  const secs = sections2551q(calc, info);
  const [pick, setPick] = useState(false);
  const eightPct = info0.indiv && (calc.t("rateElection", (data.company || {}).taxRateElection === "8%" ? "8%" : "") || "").startsWith("8%");
  const rows = form.ptRows || [];
  const setRows = (rs) => rt.patch({ ptRows: rs });
  const atcSugg = ["PT010", "PT040", "PT041", "PT060", "PT070", "PT090", "PT100", "PT105", "PT110", "PT120", "PT130", "PT140", "PT150", "PT160", "PT170", "PT180", "PT190", "PT200", "PT201"];

  const schedBody = [["PT010", calc.v.pt1amt, `${round2(calc.rate1 * 100)}%`, calc.v.pt1due], ...rows.map((r) => [r.atc || "", num(r.amount), `${round2(num(r.rate) * 100)}%`, round2(num(r.amount) * num(r.rate))])];
  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 2551Q — Quarterly Percentage Tax Return", version: "January 2018 (ENCS)", formNo: "2551Q",
    period: `${QUARTER_NAMES[q - 1]} Quarter ${year}`, company: info0,
    blocks: [
      sectionBlock(secs.header, calc, form), infoBlock("Part I – Background Information", bg2551(info0)), sectionBlock(secs.relief, calc, form),
      sectionBlock(secs.part2, calc, form), paymentBlock(form, "Part III – Details of Payment (items 25 to 28)"),
      { type: "table", title: "Schedule 1 – Computation of Tax", head: ["ATC", "Taxable Amount", "Tax Rate", "Tax Due"], body: schedBody, foot: ["7 Total Tax Due", "", "", calc.v.i7], numCols: [1, 3] },
    ], signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_2551Q_Q${q}${year}_worksheet.pdf`,
  }));

  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}><WorksheetButtons hook={exp} /></div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={PREREQ} />
    <ReturnStatusBar rt={rt} dueDefault={due2550q(year, q)} dueNote="25 days after quarter-end · verify weekends / holidays"><PeriodSelectors data={data} year={year} setYear={setYear} q={q} setQ={setQ} /></ReturnStatusBar>
    {eightPct && <Notice tone="danger">This taxpayer elected the 8% income tax rate. The 8% covers income tax and percentage tax: do not file 2551Q for this taxpayer for the year.</Notice>}
    {calc.t("amended") === "Yes" && <Notice>Amended return: item 16 is unlocked. Key in the figures of the original filing.</Notice>}
    <FormSection section={secs.header} calc={calc} rt={rt} />
    <BackgroundBlock items={bg2551(info0)} onNavigate={onNavigate} />
    <FormSection section={secs.relief} calc={calc} rt={rt} />
    <div className="tx-card">
      <div className="tx-card-head"><span>Schedule 1 – Computation of Tax (page 2)</span><button className="io-btn" onClick={() => setRows([...rows, { id: uid(), atc: "", amount: 0, rate: 0 }])} disabled={rows.length >= 5}><Plus size={13} /> Add ATC row</button></div>
      <div className="table-scroll">
        <table className="tx-form">
          <thead><tr><th style={{ width: 52 }}>Row</th><th>ATC</th><th className="num-head">Taxable Amount</th><th className="num-head">Tax Rate</th><th className="num-head">Tax Due</th><th style={{ width: 36 }}></th></tr></thead>
          <tbody>
            <tr>
              <td className="tx-no">1</td><td>PT010 <span className="tx-muted">Percentage tax, Sec. 116 (non-VAT)</span></td>
              <td className="num">
                <AutoAmount calc={calc} rt={rt} k="pt1amt" src="Sales Journal: the quarter's gross receipts (VATable + exempt + zero-rated, without output VAT)" />
              </td>
              <td className="num" style={{ width: 120 }}><RateInput value={form.ptRate1 === undefined || form.ptRate1 === "" ? rateDflt : form.ptRate1} onChange={(v) => rt.patch({ ptRate1: v })} />
                {form.ptRate1 !== undefined && form.ptRate1 !== "" && num(form.ptRate1) !== rateDflt && <button className="tx-link" onClick={() => rt.patch({ ptRate1: "" })}>reset</button>}</td>
              <td className="num">{fmtPlain(calc.v.pt1due)}</td><td></td>
            </tr>
            {rows.map((r, i) => (
              <tr key={r.id}>
                <td className="tx-no">{i + 2}</td>
                <td style={{ maxWidth: 180 }}><Field type="combo" options={atcSugg} freeText value={r.atc} onChange={(v) => setRows(rows.map((x) => x.id === r.id ? { ...x, atc: v } : x))} /></td>
                <td className="num"><NumberField value={r.amount} onChange={(v) => setRows(rows.map((x) => x.id === r.id ? { ...x, amount: v } : x))} align="right" /></td>
                <td className="num"><RateInput value={r.rate} onChange={(v) => setRows(rows.map((x) => x.id === r.id ? { ...x, rate: v } : x))} /></td>
                <td className="num">{fmtPlain(round2(num(r.amount) * num(r.rate)))}</td>
                <td className="text-center"><DelBtn onClick={() => setRows(rows.filter((x) => x.id !== r.id))} /></td>
              </tr>
            ))}
          </tbody>
          <tfoot><tr className="tx-strong"><td className="tx-no">7</td><td colSpan={3}>Total Tax Due (sum of rows 1 to 6) → item 14</td><td className="num">{fmtPlain(calc.v.i7)}</td><td></td></tr></tfoot>
        </table>
      </div>
      <div className="tx-note">Default rate: {round2(rateDflt * 100)}% for this quarter (1% from July 2020 to June 2023 under CREATE; 3% otherwise). Confirm PT010 and the rate with your references.</div>
    </div>
    <FormSection section={{ ...secs.part2, headRight: <button className="io-btn" onClick={() => setPick(true)}>Pick from Creditable Withholding Taxes</button> }} calc={calc} rt={rt} />
    <PaymentDetails rt={rt} title="Part III – Details of Payment (items 25 to 28)" firstItem={25} amountDue={calc.v.i24} />
    <ReconStrip title="Gross receipts check" colA="Schedule 1" colB="Sales Journal" items={[{ label: "PT010 taxable amount vs Sales Journal gross receipts for the quarter", a: calc.v.pt1amt, b: gross }]} causes="an override of the PT010 taxable amount." />
    <Declaration data={data} rt={rt} />
    {pick && <Pick2307Modal data={data} year={year} q={q} onCancel={() => setPick(false)} onApply={(t) => { rt.setM("i15", t); setPick(false); }} />}
  </>);
}

function RateInput({ value, onChange }) {
  const [raw, setRaw] = useState(null);
  const shown = raw !== null ? raw : `${round2(num(value) * 100)}`;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 3, justifyContent: "flex-end" }}>
      <input className="ledger-field num" style={{ maxWidth: 70 }} value={shown} onChange={(e) => setRaw(e.target.value)} onBlur={() => { if (raw !== null) { onChange(round2(num(raw)) / 100); setRaw(null); } }} />%
    </div>
  );
}

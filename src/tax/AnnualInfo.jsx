// BIR Forms 1604-C and 1604-E (January 2018 ENCS) — annual information returns of taxes withheld on
// compensation and of creditable (expanded) withholding taxes, with their alphalists.
//   1604-C reads the Employees Record, the saved 1601-C returns and (for previous-employer figures)
//          the Alphalist of Employees tab.
//   1604-E reads the saved 1601-EQ returns and the EWT source rows the QAP / EWT Schedule use.
// No DAT generator yet: the official files come from the BIR Alphalist Data Entry and Validation Module.
import React, { useMemo, useState } from "react";
import { FileText, ChevronDown, Plus, FileDown } from "lucide-react";
import {
  SectionHeader, getAvailableYears, MONTHS, num, round2, fmtPlain, fmtParen, uid, DelBtn, Field, NumberField, SearchBar,
  computeEwtSchedule, ewtSourceRows, buildQapRows, atcKey, inRange, normalizeTin, renderAlphalistSchedulePdf, loadJsPDF,
  withTimeout, reliefFormatTIN, reliefPadBranch, computeEmployeeRow, useReportExport, todayFileStamp, ITR_LINES, parseAppDate,
} from "../BookkeepingApp";
import * as XLSX from "xlsx";
import {
  TaxStyle, makeCalc, useReturn, FormSection, ReturnStatusBar, BackgroundBlock, CompanyPrereqWarning, Declaration, companyInfo,
  monthKey, quarterKey, sectionBlock, infoBlock, signatureBlock, useWorksheetExport, WorksheetButtons, ExportStatus,
  YearSelect, Notice, ReconStrip, SubTabs, AttachmentChecklist, attachCount, graduatedTax, BENEFITS_CEILING, PAYMENT_ROWS,
  paymentTotal, mdy, WORKSHEET_NOTE,
} from "./shared";
import { compute1601c } from "./Form1601C";
import { compute1601eq, ewtByAtcForQuarter } from "./EwtReturns";
import { employeeDisplayName, formatEmployeeTin, isMweRecord, employedDuring } from "./EmployeesRecord";
import { compute0619e } from "./EwtReturns";

const PREREQ = [["tin", "TIN"], ["rdo", "RDO Code"], ["address", "Registered Address"], ["zipCode", "ZIP"], ["contactNumber", "Contact Number"], ["email", "Email Address"], ["withholdingCategory", "Category of Withholding Agent"], ["topWithholdingAgent", "Top Withholding Agent"]];
const bgInfo = (i) => [["4", "TIN", i.tinText], ["5", "RDO Code", i.rdo], ["6", "Withholding Agent's Name", i.name], ["7", "Registered Address", i.address], ["7A", "ZIP Code", i.zip],
  ["8", "Category of Withholding Agent", i.category], ["8A", "Top Withholding Agent?", i.topWA], ["9", "Contact Number", i.contact], ["10", "Email Address", i.email]];
const EMP_STATUS_TEXT = { R: "Regular", CP: "Casual/Probationary", C: "Casual", P: "Probationary", S: "Seasonal" };
const SEP_TEXT = { NA: "", T: "Terminated", R: "Resigned", D: "Died", RT: "Retired" };

function firstPayment(form) {
  const p = form.payment || {};
  const row = PAYMENT_ROWS.map(([k]) => p[k] || {}).find((r) => num(r.amount)) || {};
  return { date: row.date || form.dateFiled || "", bank: row.bank || "", ref: (form.ror || {}).refNo || row.number || form.filedRef || "" };
}

/* ============================== 1604-C ============================== */

// One alphalist row per employee for the year: present-employer figures summed from the 1601-C lines,
// previous-employer figures from the linked Alphalist of Employees row (the 2316 they submitted).
export function buildCompAlphalist(data, year) {
  const records = data.employeeRecords || [];
  const sums = new Map();
  for (let m = 0; m < 12; m++) {
    const f = (data.form1601c || {})[monthKey(year, m)];
    if (!f || ((f.m || {}).anyTaxWithheld === "No")) continue;
    (f.lines || []).forEach((l) => {
      const s = sums.get(l.employeeId) || { totalComp: 0, mweBasic: 0, mweOtherPay: 0, thirteenthOther: 0, deMinimis: 0, mandatoryContrib: 0, otherNonTax: 0, notSubjectToWht: 0, taxWithheld: 0, months: Array(12).fill(0), taxMonths: Array(12).fill(0) };
      ["totalComp", "mweBasic", "mweOtherPay", "thirteenthOther", "deMinimis", "mandatoryContrib", "otherNonTax", "notSubjectToWht", "taxWithheld"].forEach((k) => { s[k] += num(l[k]); });
      s.months[m] += num(l.totalComp); s.taxMonths[m] += num(l.taxWithheld);
      sums.set(l.employeeId, s);
    });
  }
  const alphaRows = new Map((data.employees || []).filter((r) => Number(r.year) === year && r.employeeId).map((r) => [r.employeeId, computeEmployeeRow(r)]));
  const jan1 = new Date(year, 0, 1), dec31 = new Date(year, 11, 31);
  const rows = records.filter((e) => sums.has(e.id) || employedDuring(e, jan1, dec31) && e.dateHired).map((e) => {
    const s = sums.get(e.id) || { totalComp: 0, mweBasic: 0, mweOtherPay: 0, thirteenthOther: 0, deMinimis: 0, mandatoryContrib: 0, otherNonTax: 0, notSubjectToWht: 0, taxWithheld: 0, months: Array(12).fill(0), taxMonths: Array(12).fill(0) };
    const mwe = isMweRecord(e);
    const benefitsExempt = Math.min(s.thirteenthOther, BENEFITS_CEILING);
    const benefitsTaxable = round2(Math.max(0, s.thirteenthOther - BENEFITS_CEILING));
    const nonTax = round2(s.mweBasic + s.mweOtherPay + benefitsExempt + s.deMinimis + s.mandatoryContrib + s.otherNonTax + s.notSubjectToWht);
    const gross = round2(s.totalComp);
    const taxable = round2(Math.max(0, gross - nonTax));
    const a = alphaRows.get(e.id);
    const prev = a ? { gross: a.prevGross, nonTax: a.prevTotalNonTax, taxable: a.prevTotalTax, withheld: num(a.taxWithheldPrev), status: a.prevEmpStatus, from: a.prevFrom, to: a.prevTo, sep: a.prevSepReason } : { gross: 0, nonTax: 0, taxable: 0, withheld: 0 };
    const totalTaxable = round2(taxable + num(prev.taxable));
    const taxDue = graduatedTax(totalTaxable, year);
    const withheldPresent = round2(s.taxWithheld);
    const adjustment = round2(taxDue - withheldPresent - num(prev.withheld)); // + collected in December / next payroll, − refunded
    const hired = parseAppDate(e.dateHired);
    const sep = parseAppDate(e.dateSeparated);
    return {
      id: e.id, rec: e, mwe, name: employeeDisplayName(e), lastName: e.lastName, firstName: e.firstName, middleName: e.middleName,
      tin: e.tin, nationality: e.nationality || "Filipino", status: e.empStatus || "R",
      from: mdy(hired && hired > jan1 ? hired : jan1), to: mdy(sep && sep < dec31 ? sep : dec31),
      sepReason: sep && sep <= dec31 ? (SEP_TEXT[e.sepReason] || "Resigned") : "",
      gross, nonTax, taxable, benefitsTaxable, mweBasic: round2(s.mweBasic), mweOtherPay: round2(s.mweOtherPay), deMinimis: round2(s.deMinimis),
      contributions: round2(s.mandatoryContrib), benefitsExempt: round2(benefitsExempt), otherNonTax: round2(s.otherNonTax + s.notSubjectToWht),
      prev, prevEmployerTin: e.prevEmployerTin || "", prevEmployerName: e.prevEmployerName || "",
      totalTaxable, taxDue, withheldPresent, withheldPrev: num(prev.withheld), adjustment,
      withheldAsAdjusted: round2(withheldPresent + adjustment),
      substituted: e.substitutedFiling !== "N" && !(num(prev.gross) > 0) ? "Yes" : "No",
      smw: { day: num(e.smwPerDay), month: num(e.smwPerMonth), year: num(e.smwPerYear), factor: num(e.factorDaysPerYear) || 261, region: e.region || "" },
      months: s.months, taxMonths: s.taxMonths, linkedAlpha: !!a, noLines: !sums.has(e.id),
    };
  }).sort((x, y) => x.name.localeCompare(y.name));
  return { sched1: rows.filter((r) => !r.mwe), sched2: rows.filter((r) => r.mwe), all: rows };
}

function compRemittances(data, year) {
  return MONTHS.map((mName, m) => {
    const f = (data.form1601c || {})[monthKey(year, m)];
    if (!f) return { month: mName, missing: true };
    const c = compute1601c(f);
    const p = firstPayment(f);
    const paid = paymentTotal(f);
    return { month: mName, filed: !!f.filed, date: p.date, bank: p.bank, ref: p.ref, taxes: c.v.i25, adj: c.v.i26, penalties: c.v.i35, total: paid || round2(c.v.i27 - c.v.i28 + c.v.i35) };
  });
}

function compute1604c(form, rem, alpha) {
  const c = makeCalc(form);
  const sum = (k) => round2(rem.reduce((s, r) => s + num(r[k]), 0));
  c.set("p2_taxes", sum("taxes")); c.set("p2_adj", sum("adj")); c.set("p2_pen", sum("penalties")); c.set("p2_total", sum("total"));
  const all = alpha.all;
  c.set("p3_count", alpha.sched1.length); c.set("p3_mwe", alpha.sched2.length);
  c.set("p3_taxable", round2(all.reduce((s, r) => s + r.taxable, 0)));
  c.set("p3_adjusted", round2(all.reduce((s, r) => s + r.withheldAsAdjusted, 0)));
  c.a("i12", Math.max(0, round2(c.v.p2_total - c.v.p2_pen - c.v.p3_adjusted)));
  return c;
}

function sections1604c(c, info) {
  return {
    header: { title: "Header (items 1 to 3)", lines: [
      { no: "1", label: "For the Year", k: "_y", t: "ro", text: String(info.year) },
      { no: "2", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
      { no: "3", label: "Number of Sheet/s Attached", k: "sheetsAttached", t: "int", dflt: String(info.attachCount) },
      { no: "", label: "Nil return (registered withholding agent with no withholding this year — still file)", k: "nil", t: "yesno", dflt: info.nilDefault },
    ]},
    part1x: { title: "Part I (items 11 to 13)", lines: [
      { no: "11", label: "Were there any refunds of over-withheld tax released to employees after the year-end adjustment?", k: "refunds", t: "yesno", dflt: "No" },
      { no: "11A", label: "If yes, date of refund", k: "refundDate", t: "date", hide: c.t("refunds") !== "Yes" },
      { no: "12", label: "Total amount of over-remittance on tax withheld from compensation", k: "i12", t: "auto", src: "Total remitted per the 1601-Cs (less penalties) over the alphalist's tax withheld as adjusted" },
      { no: "13", label: "Month of first crediting of the over-remittance", k: "i13Month", t: "select", opts: MONTHS.map((m) => `${m} ${info.year + 1}`), hide: !(c.v.i12 > 0) },
    ]},
    part3: { title: "Part III – Alphalist summary (confirm item numbers on your copy)", lines: [
      { no: "", label: "Number of employees (Schedule 1)", k: "_cnt", t: "ro", text: String(c.v.p3_count) },
      { no: "", label: "Number of minimum wage earners (Schedule 2)", k: "_mwe", t: "ro", text: String(c.v.p3_mwe) },
      { no: "", label: "Total taxable compensation (present employer)", k: "p3_taxable", t: "calc" },
      { no: "", label: "Total tax withheld as adjusted (year-end)", k: "p3_adjusted", t: "calc", strong: true },
      { no: "", label: "Alphalist submitted?", k: "alphaSubmitted", t: "yesno", dflt: "No" },
      { no: "", label: "Submission reference (\"Validation Successful\" message / acknowledgement)", k: "alphaRef", t: "text", hide: c.t("alphaSubmitted") !== "Yes" },
    ]},
  };
}

const ATTACH_1604C = [["ackEmp", "Proof of submission: Alphalist of Employees (\"Validation Successful\" / acknowledgement receipt)"], ["ackMwe", "Proof of submission: Alphalist of Minimum Wage Earners"], ["auth", "Authorization letter (if filed by a representative)"], ["f2316", "BIR Forms 2316 (support for the alphalist)", "alphalist"]];

const S1_HEAD = ["Seq", "Name (Last, First, Middle)", "TIN", "Nationality", "Status", "From", "To", "Reason for Separation", "Gross Compensation", "13th Month & Other Benefits (exempt)", "De Minimis", "SSS/PHIC/HDMF & Union Dues", "Salaries ≤₱250K & Other Non-Taxable", "Total Non-Taxable", "Taxable Compensation", "Prev. Employer TIN", "Prev. Gross", "Prev. Non-Taxable", "Prev. Taxable", "Total Taxable", "Tax Due", "Withheld – Present", "Withheld – Previous", "Year-End Adjustment", "Substituted Filing", "Tax Withheld as Adjusted"];
const s1Row = (r, i) => [i + 1, r.name, formatEmployeeTin(r.tin), r.nationality, EMP_STATUS_TEXT[r.status] || r.status, r.from, r.to, r.sepReason, r.gross, r.benefitsExempt, r.deMinimis, r.contributions, r.otherNonTax, r.nonTax, r.taxable, r.prevEmployerTin, num(r.prev.gross), num(r.prev.nonTax), num(r.prev.taxable), r.totalTaxable, r.taxDue, r.withheldPresent, r.withheldPrev, r.adjustment, r.substituted, r.withheldAsAdjusted];
const S2_HEAD = ["Seq", "Name (Last, First, Middle)", "TIN", "Nationality", "Status", "Region", "From", "To", "SMW/Day", "SMW/Month", "SMW/Year", "Factor", "Gross Compensation", "Basic / SMW", "Holiday, OT, Night Diff., Hazard", "13th Month & Other Benefits (exempt)", "De Minimis", "SSS/PHIC/HDMF", "Other Non-Taxable", "Taxable (benefits over ₱90K, other)", "Prev. Employer TIN", "Prev. Taxable", "Tax Due", "Withheld – Present", "Withheld – Previous", "Year-End Adjustment", "Tax Withheld as Adjusted"];
const s2Row = (r, i) => [i + 1, r.name, formatEmployeeTin(r.tin), r.nationality, EMP_STATUS_TEXT[r.status] || r.status, r.smw.region, r.from, r.to, r.smw.day, r.smw.month, r.smw.year, r.smw.factor, r.gross, r.mweBasic, r.mweOtherPay, r.benefitsExempt, r.deMinimis, r.contributions, r.otherNonTax, r.taxable, r.prevEmployerTin, num(r.prev.taxable), r.taxDue, r.withheldPresent, r.withheldPrev, r.adjustment, r.withheldAsAdjusted];
const S1_NUM = [8, 9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 22, 23, 25];
const S2_NUM = [8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 21, 22, 23, 24, 25, 26];

async function downloadCompAlphalistPdf(data, year, alpha) {
  const jsPDF = await withTimeout(loadJsPDF(), 8000);
  const info = companyInfo(data);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const sched = [[1, "SCHEDULE 1 – ALPHALIST OF EMPLOYEES", S1_HEAD, alpha.sched1.map(s1Row), S1_NUM], [2, "SCHEDULE 2 – ALPHALIST OF MINIMUM WAGE EARNERS", S2_HEAD, alpha.sched2.map(s2Row), S2_NUM]];
  sched.forEach(([n, title, head, body, numCols], idx) => {
    if (idx > 0) doc.addPage("a4", "landscape");
    doc.setFont("helvetica", "bold"); doc.setFontSize(10);
    doc.text(`BIR FORM 1604-C – ${title}`, 148.5, 10, { align: "center" });
    doc.setFont("helvetica", "normal"); doc.setFontSize(8);
    doc.text(`FOR THE YEAR ${year}   ·   WITHHOLDING AGENT: ${info.name}   ·   TIN ${info.tinText}`, 148.5, 15, { align: "center" });
    const colStyles = {}; numCols.forEach((i) => { colStyles[i] = { halign: "right" }; });
    const fmtRow = (r) => r.map((v, i) => (numCols.includes(i) && typeof v === "number" ? fmtParen(v) : String(v ?? "")));
    const foot = head.map((_, i) => (i === 1 ? "TOTAL" : numCols.includes(i) && i > 11 ? fmtParen(round2(body.reduce((s, r) => s + num(r[i]), 0))) : ""));
    doc.autoTable({ startY: 19, head: [head], body: body.map(fmtRow), foot: body.length ? [foot] : undefined, theme: "grid", margin: { left: 6, right: 6 },
      styles: { fontSize: 5.4, cellPadding: 0.8, overflow: "linebreak" }, headStyles: { fillColor: [31, 95, 168], textColor: 255, fontSize: 5.2 },
      footStyles: { fillColor: [232, 241, 251], textColor: [15, 42, 77], fontStyle: "bold" }, columnStyles: colStyles, showFoot: "lastPage" });
    if (!body.length) doc.text("No employees in this schedule.", 8, 30);
  });
  const pages = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pages; i++) { doc.setPage(i); doc.setFontSize(6.5); doc.setTextColor(120); doc.text(`${WORKSHEET_NOTE} Page ${i} of ${pages}`, 6, 205); doc.setTextColor(0); }
  doc.save(`${info.tin9 || "TIN"}_1604C_Alphalist_${year}.pdf`);
}
function downloadCompAlphalistExcel(data, year, alpha) {
  const info = companyInfo(data);
  const wb = XLSX.utils.book_new();
  [["Schedule 1 Employees", S1_HEAD, alpha.sched1.map(s1Row), S1_NUM], ["Schedule 2 MWE", S2_HEAD, alpha.sched2.map(s2Row), S2_NUM]].forEach(([name, head, body, numCols]) => {
    const foot = head.map((_, i) => (i === 1 ? "TOTAL" : numCols.includes(i) && i > 11 ? round2(body.reduce((s, r) => s + num(r[i]), 0)) : ""));
    const aoa = [[`BIR FORM 1604-C – ${name}`], [`FOR THE YEAR ${year} · ${info.name} · TIN ${info.tinText}`], [], head, ...body, foot];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let r = 4; r < aoa.length; r++) numCols.forEach((ci) => { const ref = XLSX.utils.encode_cell({ r, c: ci }); if (ws[ref] && typeof ws[ref].v === "number") ws[ref].z = "#,##0.00;(#,##0.00)"; });
    ws["!cols"] = head.map((h) => ({ wch: Math.max(10, Math.min(28, String(h).length)) }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  XLSX.writeFile(wb, `${info.tin9 || "TIN"}_1604C_Alphalist_${year}.xlsx`);
}

// The Alphalist of Employees tab's Schedule 1 / 2 as an A4 landscape PDF (rows = computeEmployeeRow rows).
export async function downloadEmployeeAlphalistPdf(data, rows, year, scheduleNo) {
  const jsPDF = await withTimeout(loadJsPDF(), 8000);
  const info = companyInfo(data);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setFont("helvetica", "bold"); doc.setFontSize(10);
  doc.text(`BIR FORM 1604-C – SCHEDULE ${scheduleNo}: ${scheduleNo === 2 ? "ALPHALIST OF MINIMUM WAGE EARNERS" : "ALPHALIST OF EMPLOYEES"}`, 148.5, 10, { align: "center" });
  doc.setFont("helvetica", "normal"); doc.setFontSize(8);
  doc.text(`FOR THE YEAR ${year}   ·   ${info.name}   ·   TIN ${info.tinText}`, 148.5, 15, { align: "center" });
  const head = ["Seq", "TIN", "Name", "Status", "From", "To", "Gross (Present)", "Non-Taxable", "Taxable (Present)", "Taxable (Previous)", "Total Taxable", "Tax Due", "Withheld Prev.", "Withheld Present", "Withheld Dec.", "Refunded", "Subst. Filing"];
  const body = rows.map((r, i) => [i + 1, reliefFormatTIN(normalizeTin(r.tin)), [r.lastName, r.firstName, r.middleName].filter(Boolean).join(", "), r.empStatus, r.empFrom, r.empTo, r.pGross, r.pTotalNonTax, r.pTotalTax, r.prevTotalTax, r.totalTaxableComp, num(r.taxDue), num(r.taxWithheldPrev), num(r.taxWithheldPresent), r.amtWithheldDec, r.amtRefunded, r.substitutedFiling]);
  const numCols = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
  const foot = head.map((_, i) => (i === 2 ? "TOTAL" : numCols.includes(i) ? fmtPlain(round2(body.reduce((s, r) => s + num(r[i]), 0))) : ""));
  const colStyles = {}; numCols.forEach((i) => { colStyles[i] = { halign: "right" }; });
  doc.autoTable({ startY: 19, head: [head], body: body.map((r) => r.map((v, i) => (numCols.includes(i) ? fmtPlain(v) : String(v ?? "")))), foot: [foot], theme: "grid", margin: { left: 6, right: 6 },
    styles: { fontSize: 6.5, cellPadding: 1 }, headStyles: { fillColor: [31, 95, 168], textColor: 255 }, footStyles: { fillColor: [232, 241, 251], textColor: [15, 42, 77], fontStyle: "bold" }, columnStyles: colStyles, showFoot: "lastPage" });
  doc.save(`${info.tin9 || "TIN"}_Alphalist_Sched${scheduleNo}_${year}.pdf`);
}

function AlphaTable({ rows, mwe, search }) {
  const [open, setOpen] = useState(null);
  const shown = rows.filter((r) => !search || `${r.name} ${r.tin}`.toLowerCase().includes(search.toLowerCase()));
  const t = (k) => round2(shown.reduce((s, r) => s + num(typeof k === "function" ? k(r) : r[k]), 0));
  return (
    <div className="table-scroll">
      <table className="ledger-table tx-table">
        <thead><tr>
          <th></th><th style={{ minWidth: 200 }}>Employee</th><th>TIN (optional)</th><th>Status</th><th>Period</th>
          <th className="num-head">Gross</th><th className="num-head">Non-Taxable</th><th className="num-head">Taxable</th>
          <th className="num-head">Prev. Taxable</th><th className="num-head">Tax Due</th><th className="num-head">Withheld (Present + Prev.)</th>
          <th className="num-head">Year-End Adj.</th><th>Subst.</th><th className="num-head">Withheld as Adjusted</th>
        </tr></thead>
        <tbody>
          {shown.length === 0 && <tr><td colSpan={14} className="empty-row">No {mwe ? "minimum wage earners" : "employees"} for this year.</td></tr>}
          {shown.map((r) => (
            <React.Fragment key={r.id}>
              <tr onClick={() => setOpen(open === r.id ? null : r.id)} style={{ cursor: "pointer" }}>
                <td><ChevronDown size={13} style={{ transform: open === r.id ? "none" : "rotate(-90deg)" }} /></td>
                <td className="wrap">{r.name}{r.nationality.toLowerCase() !== "filipino" && <span className="tx-pill">{r.nationality}</span>}{r.noLines && <div className="tx-amber">No 1601-C lines this year</div>}{r.sepReason && <span className="tx-pill warn">{r.sepReason}</span>}</td>
                <td>{formatEmployeeTin(r.tin)}</td><td>{EMP_STATUS_TEXT[r.status] || r.status}</td><td>{r.from} – {r.to}</td>
                <td className="num">{fmtPlain(r.gross)}</td><td className="num">{fmtPlain(r.nonTax)}</td><td className="num">{fmtPlain(r.taxable)}</td>
                <td className="num">{fmtPlain(r.prev.taxable)}</td><td className="num">{fmtPlain(r.taxDue)}</td><td className="num">{fmtPlain(r.withheldPresent + r.withheldPrev)}</td>
                <td className={"num" + (r.adjustment < 0 ? " tx-over" : r.adjustment > 0 ? " tx-due" : "")}>{fmtParen(r.adjustment)}</td>
                <td>{r.substituted}</td><td className="num">{fmtPlain(r.withheldAsAdjusted)}</td>
              </tr>
              {open === r.id && (
                <tr><td></td><td colSpan={13} className="wrap">
                  <div className="tx-bg-grid" style={{ background: "#F8FBFF" }}>
                    {[["2316 Gross compensation (present)", r.gross], ["Basic / SMW (MWE)", r.mweBasic], ["Holiday, OT, night diff., hazard (MWE)", r.mweOtherPay], ["13th month & other benefits, exempt (up to ₱90,000)", r.benefitsExempt],
                      ["13th month & other benefits over ₱90,000 (taxable)", r.benefitsTaxable], ["De minimis", r.deMinimis], ["SSS / PHIC / HDMF & union dues", r.contributions], ["Salaries ≤ ₱250,000 & other non-taxable", r.otherNonTax],
                      ["Previous employer", `${r.prevEmployerName || "—"} ${r.prevEmployerTin ? "· TIN " + r.prevEmployerTin : ""}`], ["Previous employer gross / taxable", `${fmtPlain(r.prev.gross)} / ${fmtPlain(r.prev.taxable)}`],
                      ["Tax withheld by previous employer", r.withheldPrev], ["Monthly compensation (Jan–Dec)", r.months.map((x) => fmtPlain(x)).join(" · ")]].map(([l, v]) => (
                      <div key={l} className="tx-bg-item"><span className="tx-no"></span><span className="tx-bg-label">{l}</span><span className="tx-bg-value">{typeof v === "number" ? fmtPlain(v) : v}</span></div>
                    ))}
                  </div>
                  {!r.linkedAlpha && <div className="tx-note">Previous-employer amounts come from the employee's row in the Alphalist of Employees tab (linked to this Employees Record entry); none is linked for {r.from.slice(-4)}.</div>}
                </td></tr>
              )}
            </React.Fragment>
          ))}
        </tbody>
        {shown.length > 0 && <tfoot><tr>
          <td colSpan={5} className="totals-label">Totals ({shown.length})</td>
          <td className="num">{fmtPlain(t("gross"))}</td><td className="num">{fmtPlain(t("nonTax"))}</td><td className="num">{fmtPlain(t("taxable"))}</td>
          <td className="num">{fmtPlain(t((r) => r.prev.taxable))}</td><td className="num">{fmtPlain(t("taxDue"))}</td><td className="num">{fmtPlain(t((r) => r.withheldPresent + r.withheldPrev))}</td>
          <td className="num">{fmtParen(t("adjustment"))}</td><td></td><td className="num">{fmtPlain(t("withheldAsAdjusted"))}</td>
        </tr></tfoot>}
      </table>
    </div>
  );
}

function findCompWhtAccount(coa) {
  return (coa || []).find((a) => a.type === "Liability" && /withholding.*compensation|compensation.*withholding|wtc|withholding tax.*salar/i.test(a.name || ""));
}

function Form1604C({ data, setData, postings, onNavigate, year, selectors }) {
  const info0 = companyInfo(data);
  const rt = useReturn(data, setData, "form1604c", String(year));
  const rem = useMemo(() => compRemittances(data, year), [data.form1601c, year]);
  const alpha = useMemo(() => buildCompAlphalist(data, year), [data.employeeRecords, data.form1601c, data.employees, year]);
  const calc = useMemo(() => compute1604c(rt.form, rem, alpha), [rt.form, rem, alpha]);
  const [aTab, setATab] = useState("s1");
  const [search, setSearch] = useState("");
  const nilDefault = rem.every((r) => r.missing || !num(r.taxes)) ? "Yes" : "No";
  const info = { year, attachCount: attachCount(rt.form, ATTACH_1604C), nilDefault };
  const secs = sections1604c(calc, info);
  const remBody = rem.map((r) => (r.missing ? [r.month, "", "", "", 0, 0, 0, 0] : [r.month + (r.filed ? "" : " (draft)"), r.date, r.bank, r.ref, r.taxes, r.adj, r.penalties, r.total]));
  const remFoot = ["TOTAL", "", "", "", calc.v.p2_taxes, calc.v.p2_adj, calc.v.p2_pen, calc.v.p2_total];
  const salaryAccts = (data.coa || []).filter((a) => a.type === "Expense" && a.itrLine === ITR_LINES[12]);
  const glSalaries = round2((postings || []).filter((p) => salaryAccts.some((a) => a.code === p.code) && inRange(p.date, new Date(year, 0, 1), new Date(year, 11, 31))).reduce((s, p) => s + p.debit - p.credit, 0));
  const whtAcct = findCompWhtAccount(data.coa);
  const glWht = whtAcct ? round2((postings || []).filter((p) => p.code === whtAcct.code && inRange(p.date, new Date(year, 0, 1), new Date(year, 11, 31))).reduce((s, p) => s + p.credit, 0)) : null;
  const allGross = round2(alpha.all.reduce((s, r) => s + r.gross, 0));
  const totalAdjust = round2(alpha.all.reduce((s, r) => s + r.adjustment, 0));
  const alphaExp = useReportExport(() => downloadCompAlphalistPdf(data, year, alpha), () => downloadCompAlphalistExcel(data, year, alpha));

  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1604-C — Annual Information Return of Income Taxes Withheld on Compensation", version: "January 2018 (ENCS)", formNo: "1604-C",
    period: `For the year ${year}`, company: info0,
    blocks: [sectionBlock(secs.header, calc, rt.form), infoBlock("Part I – Background Information", bgInfo(info0)), sectionBlock(secs.part1x, calc, rt.form),
      { type: "table", title: "Part II – Summary of Remittances per BIR Form No. 1601-C", head: ["Month", "Date of Remittance", "Drawee Bank / Agency", "ROR / Check No.", "Taxes Withheld", "Adjustment", "Penalties", "Total Amount Remitted"], body: remBody, foot: remFoot, numCols: [4, 5, 6, 7] },
      sectionBlock(secs.part3, calc, rt.form),
      { type: "table", title: "Schedule 1 – Alphalist of Employees", sheet: "Schedule 1", head: S1_HEAD, body: alpha.sched1.map(s1Row), numCols: S1_NUM, landscape: true, small: true },
      { type: "table", title: "Schedule 2 – Alphalist of Minimum Wage Earners", sheet: "Schedule 2 MWE", head: S2_HEAD, body: alpha.sched2.map(s2Row), numCols: S2_NUM, landscape: true, small: true }],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1604C_${year}_worksheet.pdf`,
  }));

  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}><WorksheetButtons hook={exp} /></div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={PREREQ} />
    <ReturnStatusBar rt={rt} dueDefault={new Date(year + 1, 0, 31)} dueNote="January 31 of the following year · verify weekends / holidays / eFPS">{selectors}</ReturnStatusBar>
    {calc.t("nil", nilDefault) === "Yes" && <Notice tone="info">Nil return: no taxes withheld on compensation this year. A registered withholding agent still files the 1604-C.</Notice>}
    {calc.t("amended") === "Yes" && <Notice>Amended return: re-key the corrected figures; the previously filed lines apply.</Notice>}
    <FormSection section={secs.header} calc={calc} rt={rt} />
    <BackgroundBlock items={bgInfo(info0)} onNavigate={onNavigate} />
    <FormSection section={secs.part1x} calc={calc} rt={rt} />
    <div className="tx-card">
      <div className="tx-card-head"><span>Part II – Summary of Remittances per BIR Form No. 1601-C (confirm item numbers)</span><button className="tx-link" onClick={() => onNavigate("form1601c")}>Open 1601-C</button></div>
      <div className="table-scroll">
        <table className="ledger-table tx-table">
          <thead><tr><th>Month</th><th>Date of Remittance</th><th>Drawee Bank / Agency</th><th>ROR / Check No.</th><th className="num-head">Taxes Withheld</th><th className="num-head">Adjustment</th><th className="num-head">Penalties</th><th className="num-head">Total Amount Remitted</th></tr></thead>
          <tbody>
            {rem.map((r) => r.missing ? (
              <tr key={r.month}><td>{r.month}</td><td colSpan={7} className="tx-amber">No 1601-C saved for this month</td></tr>
            ) : (
              <tr key={r.month}><td>{r.month}{!r.filed && <span className="tx-pill warn">Draft</span>}</td><td>{r.date}</td><td>{r.bank}</td><td>{r.ref}</td>
                <td className="num">{fmtPlain(r.taxes)}</td><td className="num">{fmtParen(r.adj)}</td><td className="num">{fmtPlain(r.penalties)}</td><td className="num">{fmtPlain(r.total)}</td></tr>
            ))}
          </tbody>
          <tfoot><tr><td colSpan={4} className="totals-label">Total</td><td className="num">{fmtPlain(calc.v.p2_taxes)}</td><td className="num">{fmtParen(calc.v.p2_adj)}</td><td className="num">{fmtPlain(calc.v.p2_pen)}</td><td className="num">{fmtPlain(calc.v.p2_total)}</td></tr></tfoot>
        </table>
      </div>
    </div>
    <div className="tx-card">
      <div className="tx-card-head">
        <span>Alphalist – Schedule 1 (employees) / Schedule 2 (minimum wage earners)</span>
        <div className="io-toolbar">
          <button className="io-btn" onClick={alphaExp.exportPDF}><FileDown size={13} /> Alphalist PDF (A4 landscape)</button>
          <button className="io-btn" onClick={alphaExp.exportExcel}><FileDown size={13} /> Alphalist Excel</button>
        </div>
      </div>
      <ExportStatus status={alphaExp.status} />
      <div className="tx-note">Built from the Employees Record and the year's 1601-C lines. Exempt: 13th month and other benefits up to ₱90,000, de minimis, mandatory contributions, MWE basic / holiday / OT / night differential / hazard pay, and salaries of ₱250,000 and below (the 1601-C "not subject to withholding" column). Tax due uses the annual graduated table for {year}. Click a row for its 2316 figures.</div>
      <div className="table-toolbar" style={{ padding: "0 12px" }}>
        <SubTabs tabs={[["s1", `Schedule 1 (${alpha.sched1.length})`], ["s2", `Schedule 2 – MWE (${alpha.sched2.length})`]]} value={aTab} onChange={setATab} />
        <SearchBar value={search} onChange={setSearch} placeholder="Search name or TIN…" />
      </div>
      <AlphaTable rows={aTab === "s1" ? alpha.sched1 : alpha.sched2} mwe={aTab === "s2"} search={search} />
      <div className="tx-note">DAT file: not generated here. The official file comes from the BIR Alphalist Data Entry and Validation Module (or the eBIRForms import). Send a sample 1604-C DAT file and the module's validation log to add a DAT export.</div>
    </div>
    <FormSection section={secs.part3} calc={calc} rt={rt} />
    <ReconStrip title="1604-C check" colA="1604-C" colB="Source" items={[
      { label: "Part II total taxes withheld vs alphalist tax withheld as adjusted", a: calc.v.p2_taxes, b: calc.v.p3_adjusted, info: `the difference is the year-end adjustment (${fmtParen(totalAdjust)}: over-withholding refunded, or under-withholding collected in the first payroll of the next year)` },
      { label: `Alphalist gross compensation vs GL salaries (accounts mapped to "13 Salaries, wages and allowances")`, a: allGross, b: glSalaries, info: "salaries in the GL may include amounts outside the 1601-C (accruals, non-employees)" },
      { label: whtAcct ? `1601-C twelve months vs GL ${whtAcct.code} ${whtAcct.name} (credits)` : "1601-C twelve months vs GL withholding tax on compensation payable (no such account in the Chart of Accounts)", a: calc.v.p2_taxes, b: glWht },
    ]} causes="1601-C returns not saved or not marked Filed, manual General Journal entries, or employees missing from the 1601-C lines." />
    <AttachmentChecklist rt={rt} items={ATTACH_1604C} title="Attachment checklist (feeds item 3)" onOpen={onNavigate} />
    <Declaration data={data} rt={rt} />
    <div className="rc-warn">Penalty for not filing is ₱1,000 per failure, up to ₱25,000 a year. Surcharge, interest and compromise are manual (not computed).</div>
  </>);
}

/* ============================== 1604-E ============================== */

// Payee × ATC for the year from the EWT source rows (purchases + disbursements — never Cash Receipts CWT).
export function buildExpandedAlphalist(data, year) {
  const supp = (name) => (data.suppliers || []).find((s) => (s.name || "").toLowerCase() === String(name || "").toLowerCase());
  const atcRef = new Map((data.atc || []).map((a) => [atcKey(a.code), a]));
  const map = new Map();
  ewtSourceRows(data).forEach((e) => {
    if (!inRange(e.date, new Date(year, 0, 1), new Date(year, 11, 31))) return;
    const s = supp(e.party);
    const tin = String((s && s.tin) || (e.row && e.row.tin) || "").replace(/\D/g, "");
    const isIndiv = s && s.type === "Individual";
    const corp = isIndiv ? "" : ((s && (s.registeredName || s.name)) || e.party || "");
    const ind = isIndiv ? [s.surname, [s.firstName, s.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", ") : "";
    const atc = atcKey(e.atc);
    const key = `${tin || "NOTIN:" + String(e.party || "").toLowerCase()}|${atc}`;
    const g = map.get(key) || { key, tin, branch: tin.length > 9 ? tin.slice(9) : "0000", corp, ind, party: e.party, atc, rate: num((atcRef.get(atc) || {}).rate) || e.rate, base: 0, tax: 0, months: Array(12).fill(0), taxMonths: Array(12).fill(0) };
    const dd = parseAppDate(e.date); const m = dd ? dd.getMonth() : NaN;
    g.base += e.taxBase; g.tax += e.ewt;
    if (!isNaN(m)) { g.months[m] += e.taxBase; g.taxMonths[m] += e.ewt; }
    map.set(key, g);
  });
  const rows = [...map.values()].map((g) => ({ ...g, base: round2(g.base), tax: round2(g.tax), ratePct: round2(g.rate * 100), tinOk: g.tin.length === 9 || g.tin.length === 12 || g.tin.length === 13 || g.tin.length === 14 }))
    .sort((a, b) => (a.corp || a.ind).localeCompare(b.corp || b.ind) || a.atc.localeCompare(b.atc));
  const byName = new Map();
  rows.forEach((r) => { const n = String(r.party || "").toLowerCase(); const s = byName.get(n) || new Set(); if (r.tin) s.add(r.tin.slice(0, 9)); byName.set(n, s); });
  const dupNames = [...byName.entries()].filter(([, s]) => s.size > 1).map(([n]) => n);
  return { rows, badTin: rows.filter((r) => !r.tinOk), dupNames };
}

function exemptSuggestions(data, year, existing) {
  const flagged = (data.suppliers || []).filter((s) => s.exemptFromWht === "Yes");
  const have = new Set((existing || []).map((r) => normalizeTin(r.tin) || String(r.name || "").toLowerCase()));
  return flagged.map((s) => {
    const amt = round2((data.purchases || []).filter((r) => (r.supplier || "").toLowerCase() === (s.name || "").toLowerCase() && !r.atc && !num(r.ewt) && inRange(r.date, new Date(year, 0, 1), new Date(year, 11, 31))).reduce((t, r) => t + num(r.vatable) + num(r.nonvat), 0)
      + (data.disbursements || []).filter((r) => (r.vendor || "").toLowerCase() === (s.name || "").toLowerCase() && !r.atc && !num(r.ewt) && inRange(r.date, new Date(year, 0, 1), new Date(year, 11, 31))).reduce((t, r) => t + num(r.amount), 0));
    return { tin: s.tin || "", name: s.type === "Individual" ? [s.surname, s.firstName, s.middleName].filter(Boolean).join(", ") || s.name : (s.registeredName || s.name), amount: amt };
  }).filter((x) => x.amount > 0 && !have.has(normalizeTin(x.tin) || String(x.name || "").toLowerCase()));
}

function eqRemittances(data, year) {
  return [1, 2, 3, 4].map((q) => {
    const f = (data.form1601eq || {})[quarterKey(year, q)];
    if (!f) return { q, missing: true };
    const atcRows = ewtByAtcForQuarter(data, year, q);
    const m1i = (q - 1) * 3;
    const net = (mi) => { const g = (data.form0619e || {})[monthKey(year, mi)]; return g ? compute0619e(g, round2(computeEwtSchedule(data, year).months[mi].totalEwt)).v.i16 : 0; };
    const c = compute1601eq(f, atcRows, net(m1i), net(m1i + 1), 0);
    const p = firstPayment(f);
    return { q, filed: !!f.filed, date: p.date, bank: p.bank, ref: p.ref, taxes: c.v.i19, penalties: c.v.i30 };
  });
}

const ATTACH_1604E = [["ackPayees", "Proof of submission: Alphalist of Payees (\"Validation Successful\" / acknowledgement receipt)"], ["auth", "Authorization letter (if filed by a representative)"]];
const QN = ["1st Quarter", "2nd Quarter", "3rd Quarter", "4th Quarter"];

function Form1604E({ data, setData, postings, onNavigate, year, selectors }) {
  const info0 = companyInfo(data);
  const rt = useReturn(data, setData, "form1604e", String(year), { exemptRows: [], adj: {} });
  const form = rt.form;
  const rem = useMemo(() => eqRemittances(data, year), [data.form1601eq, data.form0619e, data.purchases, data.disbursements, data.atc, year]);
  const alpha = useMemo(() => buildExpandedAlphalist(data, year), [data.purchases, data.disbursements, data.suppliers, data.atc, year]);
  const [monthly, setMonthly] = useState(false);
  const [search, setSearch] = useState("");
  const adj = form.adj || {};
  const remRows = rem.map((r) => (r.missing ? r : { ...r, adj: num(adj[r.q]), total: round2(r.taxes + num(adj[r.q]) + r.penalties) }));
  const tot = (k) => round2(remRows.reduce((s, r) => s + num(r[k]), 0));
  const alphaTax = round2(alpha.rows.reduce((s, r) => s + r.tax, 0));
  const alphaBase = round2(alpha.rows.reduce((s, r) => s + r.base, 0));
  const sched = useMemo(() => computeEwtSchedule(data, year), [data.purchases, data.disbursements, data.atc, data.ewtSchedule, year]);
  const qapBase = useMemo(() => round2([1, 2, 3, 4].reduce((s, q) => s + buildQapRows(data, q, year).reduce((t, r) => t + num(r.taxBase), 0), 0)), [data, year]);
  const gl2300 = round2((postings || []).filter((p) => p.code === "2300" && inRange(p.date, new Date(year, 0, 1), new Date(year, 11, 31))).reduce((s, p) => s + p.credit, 0));
  const exemptRows = form.exemptRows || [];
  const sugg = exemptSuggestions(data, year, exemptRows);
  const setExempt = (rows) => rt.patch({ exemptRows: rows });
  const calc = useMemo(() => makeCalc(form), [form]);
  const header = { title: "Header (items 1 to 3)", lines: [
    { no: "1", label: "For the Year", k: "_y", t: "ro", text: String(year) },
    { no: "2", label: "Amended Return?", k: "amended", t: "yesno", dflt: "No" },
    { no: "3", label: "Number of Sheet/s Attached", k: "sheetsAttached", t: "int", dflt: String(attachCount(form, ATTACH_1604E)) },
    { no: "", label: "Nil return (registered withholding agent with no withholding this year — still file)", k: "nil", t: "yesno", dflt: alpha.rows.length ? "No" : "Yes" },
  ]};
  const shown = alpha.rows.filter((r) => !search || `${r.corp} ${r.ind} ${r.tin} ${r.atc}`.toLowerCase().includes(search.toLowerCase()));
  const pdfRows = alpha.rows.map((r, i) => ({ seq: i + 1, tin: r.tin ? `${reliefFormatTIN(r.tin.slice(0, 9))}-${reliefPadBranch(r.branch)}` : "— no TIN —", corp: r.corp, ind: r.ind, atc: r.atc, base: r.base, rate: r.ratePct, tax: r.tax }));
  const alphaExp = useReportExport(async () => {
    const jsPDF = await withTimeout(loadJsPDF(), 8000);
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    renderAlphalistSchedulePdf(doc, { titleLines: ["BIR FORM 1604-E - ALPHALIST OF PAYEES SUBJECT TO EXPANDED WITHHOLDING TAX", "ANNUAL INFORMATION RETURN OF CREDITABLE INCOME TAXES WITHHELD (EXPANDED)", `FOR THE YEAR ENDED DECEMBER 31,   ${year}`],
      agentLabel: "WITHHOLDING AGENT'S NAME :", agentName: info0.name, tinText: info0.tinText, rows: pdfRows });
    if (exemptRows.length) {
      doc.addPage("a4", "landscape");
      doc.setFont("helvetica", "bold"); doc.setFontSize(10);
      doc.text("BIR FORM 1604-E - ALPHALIST OF PAYEES OF INCOME PAYMENTS EXEMPT FROM WITHHOLDING TAX", 421, 40, { align: "center" });
      doc.autoTable({ startY: 60, head: [["Seq", "TIN", "Name of Payee", "Amount of Income Payment"]], body: exemptRows.map((r, i) => [i + 1, r.tin || "", r.name || "", fmtPlain(num(r.amount))]),
        foot: [["", "", "TOTAL", fmtPlain(round2(exemptRows.reduce((s, r) => s + num(r.amount), 0)))]], theme: "grid", styles: { fontSize: 8 }, columnStyles: { 3: { halign: "right" } } });
    }
    doc.save(`${info0.tin9 || "TIN"}_1604E_Alphalist_${year}.pdf`);
  }, () => {
    const wb = XLSX.utils.book_new();
    const head = ["SEQ NO", "TIN", "BRANCH", "CORPORATION (Registered Name)", "INDIVIDUAL (Last, First, Middle)", "ATC", "AMOUNT OF INCOME PAYMENT", "TAX RATE (%)", "AMOUNT OF TAX WITHHELD"];
    const body = alpha.rows.map((r, i) => [i + 1, r.tin.slice(0, 9), r.branch, r.corp, r.ind, r.atc, r.base, r.ratePct, r.tax]);
    const ws = XLSX.utils.aoa_to_sheet([[`BIR FORM 1604-E — Alphalist of Payees Subject to Expanded Withholding Tax — ${year}`], [`${info0.name} · TIN ${info0.tinText}`], [], head, ...body, ["", "", "", "", "TOTAL", "", alphaBase, "", alphaTax]]);
    ws["!cols"] = head.map((h) => ({ wch: Math.max(10, Math.min(34, h.length + 2)) }));
    XLSX.utils.book_append_sheet(wb, ws, "Payees EWT");
    const ws2 = XLSX.utils.aoa_to_sheet([["Alphalist of Payees of Income Payments Exempt from Withholding Tax"], [], ["SEQ NO", "TIN", "NAME OF PAYEE", "AMOUNT OF INCOME PAYMENT"], ...exemptRows.map((r, i) => [i + 1, r.tin || "", r.name || "", num(r.amount)])]);
    XLSX.utils.book_append_sheet(wb, ws2, "Exempt Payees");
    const ws3 = XLSX.utils.aoa_to_sheet([["Per-payee monthly breakdown (income payment / tax withheld)"], [], ["Payee", "TIN", "ATC", ...MONTHS.flatMap((m) => [`${m.slice(0, 3)} base`, `${m.slice(0, 3)} tax`])], ...alpha.rows.map((r) => [r.corp || r.ind || r.party, r.tin, r.atc, ...r.months.flatMap((b, i) => [round2(b), round2(r.taxMonths[i])])])]);
    XLSX.utils.book_append_sheet(wb, ws3, "Monthly breakdown");
    XLSX.writeFile(wb, `${info0.tin9 || "TIN"}_1604E_Alphalist_${year}.xlsx`);
  });

  const exp = useWorksheetExport(() => ({
    title: "BIR Form No. 1604-E — Annual Information Return of Creditable Income Taxes Withheld (Expanded) / Income Payments Exempt from Withholding Tax", version: "January 2018 (ENCS)", formNo: "1604-E",
    period: `For the year ${year}`, company: info0,
    blocks: [sectionBlock(header, calc, form), infoBlock("Part I – Background Information", bgInfo(info0)),
      { type: "table", title: "Part II – Schedule 1: Summary of Remittances per BIR Form No. 1601-EQ", head: ["Quarter", "Date of Remittance", "Drawee Bank / Agency", "ROR No.", "Taxes Withheld", "Adjustment", "Penalties", "Total Amount Remitted"],
        body: remRows.map((r) => (r.missing ? [QN[r.q - 1], "no 1601-EQ saved", "", "", 0, 0, 0, 0] : [QN[r.q - 1] + (r.filed ? "" : " (draft)"), r.date, r.bank, r.ref, r.taxes, r.adj, r.penalties, r.total])),
        foot: ["TOTAL", "", "", "", tot("taxes"), tot("adj"), tot("penalties"), tot("total")], numCols: [4, 5, 6, 7] },
      { type: "table", title: "Alphalist of Payees Subject to Creditable (Expanded) Withholding Tax", sheet: "Payees EWT", landscape: true, small: true, head: ["Seq", "TIN", "Corporation", "Individual", "ATC", "Income Payment", "Tax Rate (%)", "Tax Withheld"],
        body: alpha.rows.map((r, i) => [i + 1, r.tin ? `${reliefFormatTIN(r.tin.slice(0, 9))}-${reliefPadBranch(r.branch)}` : "— no TIN —", r.corp, r.ind, r.atc, r.base, r.ratePct, r.tax]), foot: ["", "", "TOTAL", "", "", alphaBase, "", alphaTax], numCols: [5, 7] },
      { type: "table", title: "Alphalist of Payees of Income Payments Exempt from Withholding Tax", sheet: "Exempt Payees", head: ["Seq", "TIN", "Name of Payee", "Amount of Income Payment"],
        body: exemptRows.map((r, i) => [i + 1, r.tin || "", r.name || "", num(r.amount)]), numCols: [3] }],
    signature: signatureBlock(data), filename: `${info0.tin9 || "TIN"}_1604E_${year}_worksheet.pdf`,
  }));

  return (<>
    <div className="io-toolbar" style={{ justifyContent: "flex-end", marginBottom: 8 }}><WorksheetButtons hook={exp} /></div>
    <ExportStatus status={exp.status} />
    <CompanyPrereqWarning data={data} onNavigate={onNavigate} fields={PREREQ} />
    <ReturnStatusBar rt={rt} dueDefault={new Date(year + 1, 2, 1)} dueNote="March 1 of the following year · verify weekends / holidays / eFPS">{selectors}</ReturnStatusBar>
    {calc.t("nil", alpha.rows.length ? "No" : "Yes") === "Yes" && <Notice tone="info">Nil return: no expanded withholding this year. A registered withholding agent still files the 1604-E.</Notice>}
    <FormSection section={header} calc={calc} rt={rt} />
    <BackgroundBlock items={bgInfo(info0)} onNavigate={onNavigate}>
      <div className="tx-note">Confirm whether your copy of the 1604-E Part I also asks the over-remittance questions (items 11 to 13 on the 1604-C).</div>
    </BackgroundBlock>
    <div className="tx-card">
      <div className="tx-card-head"><span>Part II – Schedule 1: Summary of Remittances per BIR Form No. 1601-EQ</span><button className="tx-link" onClick={() => onNavigate("ewtreturns")}>Open 0619-E / 1601-EQ</button></div>
      <div className="tx-note">Quarterly remittances only (the revised form no longer lists the monthly 0619-E payments).</div>
      <div className="table-scroll">
        <table className="ledger-table tx-table">
          <thead><tr><th>Quarter</th><th>Date of Remittance</th><th>Drawee Bank / Agency</th><th>ROR No.</th><th className="num-head">Taxes Withheld (item 19)</th><th className="num-head">Adjustment</th><th className="num-head">Penalties</th><th className="num-head">Total Amount Remitted</th></tr></thead>
          <tbody>
            {remRows.map((r) => r.missing ? (
              <tr key={r.q}><td>{QN[r.q - 1]}</td><td colSpan={7} className="tx-amber">No 1601-EQ saved for this quarter (the alphalist is still built)</td></tr>
            ) : (
              <tr key={r.q}><td>{QN[r.q - 1]}{!r.filed && <span className="tx-pill warn">Draft</span>}</td><td>{r.date}</td><td>{r.bank}</td><td>{r.ref}</td>
                <td className="num">{fmtPlain(r.taxes)}</td><td style={{ width: 140 }}><NumberField value={adj[r.q]} onChange={(v) => rt.patch({ adj: { ...adj, [r.q]: v } })} align="right" /></td>
                <td className="num">{fmtPlain(r.penalties)}</td><td className="num">{fmtPlain(r.total)}</td></tr>
            ))}
          </tbody>
          <tfoot><tr><td colSpan={4} className="totals-label">Total</td><td className="num">{fmtPlain(tot("taxes"))}</td><td className="num">{fmtParen(tot("adj"))}</td><td className="num">{fmtPlain(tot("penalties"))}</td><td className="num">{fmtPlain(tot("total"))}</td></tr></tfoot>
        </table>
      </div>
    </div>
    {(alpha.badTin.length > 0 || alpha.dupNames.length > 0) && (
      <Notice tone="danger">
        {alpha.badTin.length > 0 && <div><strong>{alpha.badTin.length} payee line{alpha.badTin.length === 1 ? "" : "s"} with a missing or malformed TIN — fix before filing</strong> (Suppliers Master): {alpha.badTin.map((r) => `${r.corp || r.ind || r.party} (${r.atc})`).join(", ")}.</div>}
        {alpha.dupNames.length > 0 && <div>Same payee name with two TINs: {alpha.dupNames.join(", ")}. Each variant is listed as its own row.</div>}
        <button className="tx-link" onClick={() => onNavigate("suppliers")}>Open Suppliers Master</button>
      </Notice>
    )}
    <div className="tx-card">
      <div className="tx-card-head">
        <span>Alphalist of Payees Subject to Creditable (Expanded) Withholding Tax — consolidated per payee per ATC</span>
        <div className="io-toolbar">
          <button className="io-btn" onClick={() => setMonthly(!monthly)}>{monthly ? "Hide" : "Show"} monthly breakdown</button>
          <button className="io-btn" onClick={alphaExp.exportPDF}><FileDown size={13} /> Alphalist PDF (A4 landscape)</button>
          <button className="io-btn" onClick={alphaExp.exportExcel}><FileDown size={13} /> Alphalist Excel</button>
        </div>
      </div>
      <ExportStatus status={alphaExp.status} />
      <div className="tx-note">Purchase Journal and Cash Disbursements rows with an ATC (the EWT Schedule / QAP rows). Cash Receipts CWT is tax withheld from the company and is not included. Rate is shown as a percentage; the "nature of income payment" column was removed from the revised form.</div>
      <div className="table-toolbar" style={{ padding: "0 12px" }}><SearchBar value={search} onChange={setSearch} placeholder="Search payee, TIN or ATC…" /></div>
      <div className="table-scroll">
        <table className="ledger-table tx-table">
          <thead><tr><th>Seq</th><th>TIN</th><th style={{ minWidth: 200 }}>Corporation</th><th style={{ minWidth: 180 }}>Individual</th><th>ATC</th>
            {monthly ? MONTHS.map((m) => <th key={m} className="num-head">{m.slice(0, 3)}</th>) : null}
            <th className="num-head">Income Payment</th><th className="num-head">Rate (%)</th><th className="num-head">Tax Withheld</th></tr></thead>
          <tbody>
            {shown.length === 0 && <tr><td colSpan={monthly ? 20 : 8} className="empty-row">No ATC-tagged payments for {year}.</td></tr>}
            {shown.map((r, i) => (
              <tr key={r.key}>
                <td>{i + 1}</td><td>{r.tin ? `${reliefFormatTIN(r.tin.slice(0, 9))}-${reliefPadBranch(r.branch)}` : <span className="tx-amber">no TIN</span>}{r.tin && !r.tinOk && <span className="tx-pill warn">check</span>}</td>
                <td className="wrap">{r.corp}</td><td className="wrap">{r.ind}</td><td>{r.atc}</td>
                {monthly ? r.months.map((b, mi) => <td key={mi} className="num" title={`Tax withheld ${fmtPlain(r.taxMonths[mi])}`}>{b ? fmtPlain(b) : "—"}</td>) : null}
                <td className="num">{fmtPlain(r.base)}</td><td className="num">{r.ratePct.toFixed(2)}</td><td className="num">{fmtPlain(r.tax)}</td>
              </tr>
            ))}
          </tbody>
          {shown.length > 0 && <tfoot><tr><td colSpan={monthly ? 17 : 5} className="totals-label">Total ({shown.length} line{shown.length === 1 ? "" : "s"})</td><td className="num">{fmtPlain(round2(shown.reduce((s, r) => s + r.base, 0)))}</td><td></td><td className="num">{fmtPlain(round2(shown.reduce((s, r) => s + r.tax, 0)))}</td></tr></tfoot>}
        </table>
      </div>
      <div className="tx-note">DAT file: not generated here (the 1604-E DAT has a different record layout from the QAP). Send a sample 1604-E DAT file and the validation log to add it.</div>
    </div>
    <div className="tx-card">
      <div className="tx-card-head"><span>Alphalist of Payees of Income Payments Exempt from Withholding Tax</span><button className="io-btn" onClick={() => setExempt([...exemptRows, { id: uid(), tin: "", name: "", amount: 0 }])}><Plus size={13} /> Add row</button></div>
      <div className="tx-note">Payments exempt from withholding but subject to income tax (e.g. below the withholding threshold with a sworn declaration). Suggestions come from suppliers ticked "Exempt from Withholding" in Suppliers Master, using their purchases with no ATC and no EWT.</div>
      {sugg.length > 0 && <div className="tx-note">{sugg.map((s) => <span key={s.name} style={{ marginRight: 10 }}>{s.name} ({fmtPlain(s.amount)}) <button className="tx-link" onClick={() => setExempt([...exemptRows, { id: uid(), ...s }])}>Add</button></span>)}</div>}
      <div className="table-scroll">
        <table className="ledger-table tx-table">
          <thead><tr><th>TIN</th><th>Name of Payee</th><th className="num-head">Amount of Income Payment</th><th style={{ width: 36 }}></th></tr></thead>
          <tbody>
            {exemptRows.length === 0 && <tr><td colSpan={4} className="empty-row">No exempt payees.</td></tr>}
            {exemptRows.map((r) => (
              <tr key={r.id}>
                <td><Field value={r.tin} onChange={(v) => setExempt(exemptRows.map((x) => x.id === r.id ? { ...x, tin: v } : x))} /></td>
                <td><Field value={r.name} onChange={(v) => setExempt(exemptRows.map((x) => x.id === r.id ? { ...x, name: v } : x))} /></td>
                <td><NumberField value={r.amount} onChange={(v) => setExempt(exemptRows.map((x) => x.id === r.id ? { ...x, amount: v } : x))} align="right" /></td>
                <td className="text-center"><DelBtn onClick={() => setExempt(exemptRows.filter((x) => x.id !== r.id))} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
    <ReconStrip title="1604-E check" colA="Alphalist" colB="Source" items={[
      { label: "Alphalist tax withheld vs the four 1601-EQ item 19 totals", a: alphaTax, b: rem.some((r) => r.missing) ? null : tot("taxes"), info: rem.some((r) => r.missing) ? "" : undefined },
      { label: "Alphalist tax withheld vs EWT Schedule year total", a: alphaTax, b: round2(sched.total.totalEwt) },
      { label: "Alphalist tax withheld vs GL 2300 Withholding Taxes Expanded (credits)", a: alphaTax, b: gl2300 },
      { label: "Alphalist income payments vs QAP quarters (tax base)", a: alphaBase, b: qapBase },
      { label: "Payees with missing or invalid TIN", a: alpha.badTin.length, b: 0, aText: String(alpha.badTin.length), bText: "0" },
    ]} causes="ATCs missing an EWT Group, rows dated outside the year, manual General Journal entries, 1601-EQ returns not saved or not marked Filed, or suppliers missing a TIN." />
    {rem.some((r) => r.missing) && <div className="tx-note">The comparison with the 1601-EQs is shown once all four quarters are saved.</div>}
    <AttachmentChecklist rt={rt} items={ATTACH_1604E} title="Attachment checklist (feeds item 3)" />
    <Declaration data={data} rt={rt} />
  </>);
}

/* ============================== Page ============================== */

export function AnnualInfoPage({ data, setData, postings, onNavigate }) {
  const now = new Date();
  const [tab, setTab] = useState("c");
  const [year, setYear] = useState(now.getMonth() < 3 ? now.getFullYear() - 1 : now.getFullYear());
  const selectors = <YearSelect years={getAvailableYears(data)} value={year} onChange={setYear} />;
  const props = { data, setData, postings, onNavigate, year, selectors };
  return (
    <div>
      <TaxStyle />
      <SectionHeader icon={FileText} title="1604-C / 1604-E – Annual Information Returns" subtitle="January 2018 (ENCS). Annual returns of taxes withheld on compensation (1604-C, due January 31) and of expanded withholding taxes (1604-E, due March 1), with their alphalists. One consolidated return per company." />
      <SubTabs tabs={[["c", "1604-C (Compensation)"], ["e", "1604-E (Expanded)"]]} value={tab} onChange={setTab} />
      {tab === "c" ? <Form1604C {...props} /> : <Form1604E {...props} />}
    </div>
  );
}

export { todayFileStamp, uid };

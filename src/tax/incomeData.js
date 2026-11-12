// Income tax return figures (1701Q / 1702Q / 1701 / 1702-RT), read from the same GL postings as the
// Income Statement (computeIncomeStatement handles inventory-based cost of goods sold). Nothing here
// loops over the journals directly.
import { num, round2, computeIncomeStatement, inRange, ITR_LINES, ITR_NON_DEDUCTIBLE, parseAppDate, computeAssetRow } from "../BookkeepingApp";

const isContraRevenue = (a) => /discount|return|allowance/i.test(a.name || "") || /contra/i.test(a.notes || "");
const isOtherIncome = (a) => /other income/i.test(`${a.category || ""} ${a.accountType || ""}`) || /other income|interest income|gain/i.test(a.name || "");
export const isFinalTaxAccount = (a) => a.finalTax === "Yes";

function inventoryFor(data, year, monthIdx) {
  const log = data.inventoryLog || [];
  const beg = log.find((x) => x.year === year && x.month === 0);
  const end = log.find((x) => x.year === year && x.month === monthIdx);
  return { beg: num(beg && beg.beginning), end: num(end && end.ending) };
}

// All figures for [start, end]. Returns per-ITR-line deductions, the unmapped / non-deductible lists, etc.
export function incomeFigures(data, postings, start, end) {
  const coa = data.coa || [];
  const byCode = new Map();
  (postings || []).forEach((p) => {
    if (!inRange(p.date, start, end)) return;
    const s = byCode.get(p.code) || { d: 0, c: 0 };
    s.d += p.debit; s.c += p.credit;
    byCode.set(p.code, s);
  });
  const credit = (code) => { const s = byCode.get(code) || { d: 0, c: 0 }; return s.c - s.d; };
  const debit = (code) => -credit(code);

  let sales = 0, returns = 0, otherIncome = 0, finalTaxIncome = 0;
  const otherIncomeAccounts = [];
  coa.filter((a) => a.type === "Revenue").forEach((a) => {
    const amt = credit(a.code);
    if (!amt) return;
    if (isContraRevenue(a)) returns += -amt;
    else if (isOtherIncome(a)) {
      if (isFinalTaxAccount(a)) finalTaxIncome += amt;
      else { otherIncome += amt; otherIncomeAccounts.push({ code: a.code, name: a.name, amount: round2(amt) }); }
    } else sales += amt;
  });

  // Cost of sales: the Income Statement's COGS (with inventory) + cost of services, plus any other
  // Cost of Sales account it doesn't cover.
  const endMonth = end.getMonth();
  const inv = inventoryFor(data, end.getFullYear(), endMonth);
  const isFullFromJan = start.getMonth() === 0 && start.getDate() === 1;
  const stmt = computeIncomeStatement(postings || [], coa, start, end, isFullFromJan ? inv.beg : 0, num(inv.end));
  const covered = new Set(["5001", "5002", "5003"]);
  let otherCos = 0;
  coa.filter((a) => a.type === "Cost of Sales" && !covered.has(a.code) && a.category !== "Cost of services").forEach((a) => { otherCos += debit(a.code); });
  const costOfSales = round2(stmt.cogs + stmt.totalCostOfServices + otherCos);

  const lines = Object.fromEntries(ITR_LINES.map((l) => [l, 0]));
  const unmapped = [], nonDeductible = [];
  let totalExpenses = 0;
  coa.filter((a) => a.type === "Expense").forEach((a) => {
    const amt = round2(debit(a.code));
    if (!amt) return;
    totalExpenses += amt;
    const line = a.itrLine || "";
    if (line === ITR_NON_DEDUCTIBLE) { nonDeductible.push({ code: a.code, name: a.name, amount: amt }); return; }
    if (!ITR_LINES.includes(line)) { unmapped.push({ code: a.code, name: a.name, amount: amt }); lines["17d Others"] += amt; return; }
    lines[line] += amt;
  });
  Object.keys(lines).forEach((k) => { lines[k] = round2(lines[k]); });
  const deductions = round2(Object.values(lines).reduce((s, v) => s + v, 0));
  const nonDeductibleTotal = round2(nonDeductible.reduce((s, x) => s + x.amount, 0));
  const r = (x) => round2(x);
  return {
    sales: r(sales), returns: r(returns), netSales: r(sales - returns), costOfSales, grossIncome: r(sales - returns - costOfSales),
    otherIncome: r(otherIncome), otherIncomeAccounts, finalTaxIncome: r(finalTaxIncome),
    lines, deductions, unmapped, nonDeductible, nonDeductibleTotal, totalExpenses: r(totalExpenses),
    netIncomeBooks: r(sales - returns - costOfSales + otherIncome + finalTaxIncome - totalExpenses),
    isNetSales: stmt.netSales, isNetIncome: stmt.netIncome,
  };
}

// Creditable tax withheld per 2307 (the Creditable Withholding Taxes tab) dated in [start, end].
export function cwtInRange(data, start, end) {
  return round2((data.form2307 || []).filter((r) => inRange(r.date, start, end)).reduce((s, r) => s + num(r.cwt), 0));
}

// Fixed Asset Ledger depreciation for the periods posted in a given year.
export function faDepreciationForYear(data, year) {
  let t = 0;
  (data.fixedAssets || []).forEach((a) => {
    const c = computeAssetRow(a);
    (a.postedPeriods || []).forEach((k) => { if (parseInt(String(k).slice(0, 4), 10) === year) t += c.monthlyDep; });
  });
  return round2(t);
}

// MCIT applies from the 4th taxable year following the year operations commenced.
export function mcitApplies(company, year) {
  const y0 = parseInt(company.yearBusinessStarted, 10);
  if (!y0) return { applies: true, known: false };
  return { applies: year >= y0 + 4, known: true, firstYear: y0 + 4 };
}

export const corpRate = (company) => (company.incomeTaxRegime === "Regular 20%" ? 0.20 : 0.25);
export { parseAppDate };

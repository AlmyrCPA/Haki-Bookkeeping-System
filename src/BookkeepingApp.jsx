import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import {
  LayoutDashboard, Building2, BookOpenText, Users, Truck, ListChecks,
  Receipt, Wallet, ArrowDownToLine, ArrowUpFromLine, ScrollText, Scale,
  ClipboardList, TrendingUp, Landmark, Plus, Trash2, Menu, X, Info,
  ChevronDown, Save, RotateCcw, FileDown, FileUp, BarChart3, RefreshCw, Search, FileText, LogOut,
  Settings, Check,
} from "lucide-react";
import * as XLSX from "xlsx";
import { doc, onSnapshot, setDoc } from "firebase/firestore";
import { db } from "./firebase";

/* ============================== CONSTANTS ============================== */

const PESO = new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", minimumFractionDigits: 2 });
const fmt = (n) => PESO.format(Number(n) || 0);
const fmtPlain = (n) => (Number(n) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtParen = (n) => { const v = Number(n) || 0; return v < 0 ? `(${fmtPlain(Math.abs(v))})` : fmtPlain(v); };
const formatDMY = (d) => `${pad2(d.getDate())} ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.getMonth()]} ${d.getFullYear()}`;
const num = (v) => { const n = parseFloat(String(v ?? "").replace(/,/g, "")); return isNaN(n) ? 0 : n; };
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;

const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
// The year range is no longer a fixed cap — it always includes a sensible window around today,
// but widens automatically to cover whatever years actually appear in the entered data (journals,
// importation, etc.), so nothing is ever hidden just because it falls outside a hardcoded range.
function getAvailableYears(data) {
  const currentYear = new Date().getFullYear();
  const years = new Set([currentYear - 1, currentYear, currentYear + 1, currentYear + 2]);
  const collect = (rows, dateField) => {
    (rows || []).forEach((r) => {
      const d = parseAppDate(r[dateField]);
      if (d) years.add(d.getFullYear());
    });
  };
  if (data) {
    collect(data.sales, "date");
    collect(data.purchases, "date");
    collect(data.disbursements, "date");
    collect(data.receipts, "date");
    collect(data.generalJournal, "date");
    collect(data.importation, "date");
  }
  const min = Math.min(...years), max = Math.max(...years);
  const result = [];
  for (let y = min; y <= max; y++) result.push(y);
  return result;
}
const QUARTERS = [
  { label: "Q1 (Jan–Mar)", months: [0,1,2] },
  { label: "Q2 (Apr–Jun)", months: [3,4,5] },
  { label: "Q3 (Jul–Sep)", months: [6,7,8] },
  { label: "Q4 (Oct–Dec)", months: [9,10,11] },
];

const BANK_ACCOUNTS = ["Cash on Hand", "Cash in Bank"];
const TERMS = ["Cash", "Credit"];
export const PARTY_TYPES = ["Individual", "Non-Individual"];
const ITEM_TYPES = ["Services", "Other than Capital Goods", "Capital Goods"];
const COA_TYPES = ["Asset", "Liability", "Equity", "Revenue", "Cost of Sales", "Expense"];
const COA_SUBCLASSES = ["Current asset","Non Current Asset","Current Liabilities","Non-current Liabilities","Equity","Revenue","Cost of Goods Sold","Cost of services","Operating expense","Other Income","Other Expense","Income Tax expense"];
const COA_ACCOUNT_TYPES = ["Bank","Account Receivable","Current asset","Fixed Assets","Non Current Asset","Accounts Payable","Current Liabilities","Non-current Liabilities","Equity","Revenue","Cost of Goods Sold","Cost of services","Operating expense","Other Income","Other Expense","Income Tax expense"];

const DEFAULT_COA = [
  ["1000","Cash on Hand","Asset","Current asset","Bank",""],
  ["1010","Cash in Bank","Asset","Current asset","Bank",""],
  ["1200","Accounts Receivable","Asset","Current asset","Account Receivable",""],
  ["1300","Inventories","Asset","Current asset","Current asset",""],
  ["1400","Input VAT","Asset","Current asset","Current asset","Auto-summarized on VAT Summary"],
  ["1500","Creditable Withholding Tax","Asset","Current asset","Current asset","From Cash Receipts"],
  ["1600","Security Deposit","Asset","Current asset","Current asset",""],
  ["1700","Transportation Vehicle","Asset","Non Current Asset","Fixed Assets",""],
  ["1701","Furniture and Fixtures","Asset","Non Current Asset","Fixed Assets",""],
  ["1710","Accumulated Depn - Trans. Vehicle","Asset","Non Current Asset","Fixed Assets","Contra-asset"],
  ["1711","Accumulated Depn - Furniture and Fixture","Asset","Non Current Asset","Fixed Assets","Contra-asset"],
  ["2000","Accounts Payable","Liability","Current Liabilities","Accounts Payable",""],
  ["2020","VAT Payable","Liability","Current Liabilities","Current Liabilities",""],
  ["2100","Accrued Expenses","Liability","Current Liabilities","Current Liabilities",""],
  ["2200","Output VAT","Liability","Current Liabilities","Current Liabilities","Auto-summarized on VAT Summary"],
  ["2300","Withholding Taxes Expanded","Liability","Current Liabilities","Current Liabilities","From Cash Disbursements"],
  ["2400","SSS, PhilHealth and HDMF Payable","Liability","Current Liabilities","Current Liabilities",""],
  ["2500","Advances from Stockholders","Liability","Non-current Liabilities","Non-current Liabilities",""],
  ["2600","Loans Payable","Liability","Non-current Liabilities","Non-current Liabilities",""],
  ["3000","Capital Stock","Equity","Equity","Equity",""],
  ["3001","Retained Earnings","Equity","Equity","Equity","Auto-computed from Net Income"],
  ["3002","Owner's Drawing","Equity","Equity","Equity","Contra-equity"],
  ["4000","Sales","Revenue","Revenue","Revenue",""],
  ["4001","Sales Discount","Revenue","Revenue","Revenue","Contra-revenue"],
  ["4002","Other Income","Revenue","Other Income","Other Income",""],
  ["5001","Purchases","Cost of Sales","Cost of Goods Sold","Cost of Goods Sold",""],
  ["5002","Handling Costs","Cost of Sales","Cost of Goods Sold","Cost of Goods Sold",""],
  ["5003","Purchase Returns and Allowances","Cost of Sales","Cost of Goods Sold","Cost of Goods Sold","Contra-cost-of-sales"],
  ["5004","Logistics Expense","Cost of Sales","Cost of services","Cost of services",""],
  ["5005","Freight","Cost of Sales","Cost of services","Cost of services",""],
  ["5006","Other Services","Cost of Sales","Cost of services","Cost of services",""],
  ["6010","Salaries and Wages","Expense","Operating expense","Operating expense",""],
  ["6020","Depreciation","Expense","Operating expense","Operating expense",""],
  ["6030","Insurance","Expense","Operating expense","Operating expense",""],
  ["6040","Taxes and Licenses","Expense","Operating expense","Operating expense",""],
  ["6050","Transportation and Travel","Expense","Operating expense","Operating expense",""],
  ["6070","Rent Expense","Expense","Operating expense","Operating expense",""],
  ["6080","Association Dues","Expense","Operating expense","Operating expense",""],
  ["6090","Communication, Light and Water","Expense","Operating expense","Operating expense",""],
  ["6120","Gas and Oil","Expense","Operating expense","Operating expense",""],
  ["6130","Office Supplies","Expense","Operating expense","Operating expense",""],
  ["6140","Depreciation - Other","Expense","Operating expense","Operating expense",""],
  ["6150","Representation","Expense","Operating expense","Operating expense",""],
  ["6160","Repairs and Maintenance","Expense","Operating expense","Operating expense",""],
  ["6170","Miscellaneous","Expense","Operating expense","Operating expense",""],
  ["6180","SSS, PhilHealth and HDMF Contributions","Expense","Operating expense","Operating expense",""],
  ["6190","Brokerage Fee","Expense","Operating expense","Operating expense",""],
  ["7100","Income Tax Expense","Expense","Income Tax expense","Income Tax expense",""],
].map(([code,name,type,category,accountType,notes]) => ({ id: code, code, name, type, category, accountType, notes }));

const DEFAULT_ITEMS = [
  ["GASOLINE","6050","Services"],
  ["MATERIALS","5001","Other than Capital Goods"],
  ["PARKING FEE","6050","Services"],
  ["RENT","6070","Services"],
  ["PARKING RENT","6070","Services"],
  ["REPAIRS - GOODS","6160","Other than Capital Goods"],
  ["OFFICE SUPPLIES","6130","Other than Capital Goods"],
  ["TOLL FEE","6050","Services"],
  ["LPG","6120","Services"],
  ["WATER","6090","Services"],
  ["TELEPHONE","6090","Services"],
  ["ELECTRICITY","6090","Services"],
  ["CUSA","6080","Services"],
].map(([item,account,type]) => ({ id: item, item, account, type }));

const DEFAULT_ATC = [
  ["WI010",0.10,"Professional/talent fees – individuals (no sworn declaration, or gross income over ₱3M)"],
  ["WI011",0.05,"Professional/talent fees – individuals (sworn declaration on file, ₱3M or below)"],
  ["WC010",0.15,"Professional fees – corporations (gross income over ₱720,000)"],
  ["WC011",0.10,"Professional fees – corporations (gross income ₱720,000 or below)"],
  ["WI100",0.05,"Rental of real property – individual lessor"],
  ["WC100",0.05,"Rental of real property – corporate lessor"],
  ["WI120",0.02,"Contractors (engineering/building/specialty) – individual"],
  ["WC120",0.02,"Contractors (engineering/building/specialty) – corporate"],
  ["WI158",0.01,"Purchases of goods – Top Withholding Agent, individual supplier"],
  ["WC158",0.01,"Purchases of goods – Top Withholding Agent, corporate supplier"],
  ["WI160",0.02,"Purchases of services – Top Withholding Agent, individual supplier"],
  ["WC160",0.02,"Purchases of services – Top Withholding Agent, corporate supplier"],
].map(([code,rate,desc], i) => ({ id: "atc"+i, code, rate, desc }));

const VAT_TYPE_BOOKS = ["Purchase journal", "Sales journal"];
const VAT_TYPE_SLSPI_FIELDS = ["Services", "Capital goods", "Other than capital goods", "Exempt", "Zero-Rated", "Exempt Sales", "Zero-Rated Sales", "Taxable Sales"];
const DEFAULT_VAT_TYPES = [
  ["12% VAT on Purchases (Services)", "Purchase journal", "Other than capital goods"],
  ["12% VAT on Purchases (Goods)", "Purchase journal", "Services"],
  ["12% VAT on Purchases (Capital Goods)", "Purchase journal", "Capital goods"],
  ["0% VAT-exempt Purchases", "Purchase journal", "Exempt"],
  ["0% Zero-rated VAT purchases", "Purchase journal", "Zero-Rated"],
  ["12% Services by Non-residents", "Purchase journal", "Services"],
  ["0% Non-VAT purchases", "Purchase journal", ""],
  ["0% Non VAT sales", "Sales journal", ""],
  ["0% VAT Exempt Sales", "Sales journal", "Exempt Sales"],
  ["0% Zero rated Sales", "Sales journal", "Zero-Rated Sales"],
  ["12% VATable Sales (Goods)", "Sales journal", "Taxable Sales"],
  ["12% VATable Sales (Services)", "Sales journal", "Taxable Sales"],
].map(([vatType, books, slspiField], i) => ({ id: "vt" + i, vatType, books, slspiField }));

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const pad2 = (n) => String(n).padStart(2, "0");
const todayMDY = () => { const d = new Date(); return `${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}/${d.getFullYear()}`; };
const todayFileStamp = () => { const d = new Date(); return `${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}-${d.getFullYear()}`; };
// Parses either "MM/DD/YYYY" (current storage format) or legacy "YYYY-MM-DD" (older backups) into a Date at local midnight.
function parseAppDate(str) {
  if (!str) return null;
  const s = String(str).trim();
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) return new Date(Number(mdy[3]), Number(mdy[1]) - 1, Number(mdy[2]));
  const ymd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (ymd) return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
// Normalizes any recognized date string to canonical "MM/DD/YYYY" for storage/display.
function toMDY(str) {
  const d = parseAppDate(str);
  return d ? `${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}/${d.getFullYear()}` : (str || "");
}

export const emptyCompany = {
  name: "", registeredName: "", address: "", zipCode: "", tin: "", rdo: "", lineOfBusiness: "",
  atp: "", dateRegistered: "", preparedBy: "", taxpayerType: "Non-Individual",
  surname: "", firstName: "", middleName: "",
  branchCode: "0000", addr1: "", addr2: "", tradeName: "",
  vatStatus: "VAT Registered",
  // Payor's Authorized Representative — feeds the 2307 certificate's signature block.
  authorizedSignatory: "", signatoryPosition: "", signatoryTin: "", signatureImage: "",
};

const NAV = [
  { group: "Overview", items: [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  ]},
  { group: "Setup", items: [
    { key: "company", label: "Company Details", icon: Building2 },
    { key: "coa", label: "Chart of Accounts", icon: BookOpenText },
    { key: "customers", label: "Customers Master", icon: Users },
    { key: "suppliers", label: "Suppliers Master", icon: Truck },
    { key: "items", label: "Item Record", icon: ListChecks },
    { key: "atc", label: "ATC Reference", icon: ListChecks },
    { key: "vattypes", label: "VAT Type", icon: Scale },
  ]},
  { group: "Journals", items: [
    { key: "sales", label: "Sales Journal", icon: Receipt },
    { key: "purchases", label: "Purchase Journal", icon: Wallet },
    { key: "disbursements", label: "Cash Disbursements", icon: ArrowUpFromLine },
    { key: "receipts", label: "Cash Receipts", icon: ArrowDownToLine },
    { key: "generaljournal", label: "General Journal", icon: ScrollText },
    { key: "fixedassets", label: "Fixed Asset Ledger", icon: Landmark },
    { key: "generalledger", label: "General Ledger", icon: BookOpenText },
  ]},
  { group: "Reports", items: [
    { key: "vat", label: "VAT Summary", icon: Scale },
    { key: "ledger", label: "Monthly Ledger", icon: ClipboardList },
    { key: "transactions", label: "Transaction Details", icon: Search },
    { key: "trial", label: "Trial Balance", icon: ClipboardList },
    { key: "income", label: "Income Statement", icon: TrendingUp },
    { key: "balance", label: "Balance Sheet", icon: Landmark },
  ]},
  { group: "Tax Compliance", items: [
    { key: "importation", label: "Importation Ledger", icon: ArrowDownToLine },
    { key: "slspi", label: "SLSPI", icon: FileText },
    { key: "qap", label: "QAP", icon: FileText },
    { key: "sawt", label: "SAWT", icon: FileText },
    { key: "form2307", label: "Creditable Withholding Taxes", icon: FileText },
    { key: "alphalist", label: "Alphalist of Employees", icon: FileText },
  ]},
];

/* ============================== INITIAL STATE ============================== */

export function makeInitialData() {
  return {
    company: { ...emptyCompany },
    coa: DEFAULT_COA,
    items: DEFAULT_ITEMS,
    atc: DEFAULT_ATC,
    vatTypes: DEFAULT_VAT_TYPES,
    customers: [],
    suppliers: [],
    sales: [],
    purchases: [],
    disbursements: [],
    receipts: [],
    generalJournal: [],
    inventoryLog: [], // {id, year, month, beginning, ending}
    importation: [],
    fixedAssets: [],
    form2307: [],
    employees: [],
  };
}

function makeSampleAddOns() {
  return {
    customers: [
      { id: uid(), tin: "008-064-433", name: "ARQUEE CORP", registeredName: "ARQUEE CORP", firstName: "", middleName: "", surname: "", address: "MANDAUE CITY, CEBU", type: "Non-Individual", account: "SALES", taxCode: 0.12 },
      { id: uid(), tin: "230-805-640", name: "DELA CERNA LIBRADO PRACULLUS", registeredName: "", firstName: "LIBRADO", middleName: "PRACULLUS", surname: "DELA CERNA", address: "STA ROSA, LAGUNA", type: "Individual", account: "SALES", taxCode: 0.12 },
    ],
    suppliers: [
      { id: uid(), tin: "259-225-658", name: "BAJARO JOHN PAULO D.", registeredName: "", firstName: "JOHN PAULO", middleName: "DECO", surname: "BAJARO", address: "METRO MANILA", itemCode: "MATERIALS", accountTitle: "Purchases", type: "Individual", taxCode: "Non-VAT", businessStyle: "MR DIY" },
      { id: uid(), tin: "200-035-311", name: "ACE HARDWARE PHILIPPINES INC.", registeredName: "ACE HARDWARE PHILIPPINES INC.", firstName: "", middleName: "", surname: "", address: "SM SEASIDE CITY CEBU", itemCode: "REPAIRS - GOODS", accountTitle: "Repairs and Maintenance", type: "Non-Individual", taxCode: "VAT", businessStyle: "" },
    ],
    sales: [
      computeSalesRow({ id: uid(), date: "01/05/2026", siNo: "SI-0001", customer: "ARQUEE CORP", tin: "008-064-433", address: "MANDAUE CITY, CEBU", desc: "Assorted merchandise sold", vatable: 5000, exempt: 0, zeroRated: 0, terms: "Cash", coaCode: "4000", bankAccount: "Cash on Hand" }),
    ],
    purchases: [
      computePurchaseRow({ id: uid(), date: "01/06/2026", invNo: "INV-8891", supplier: "ACE HARDWARE PHILIPPINES INC.", tin: "200-035-311", address: "SM SEASIDE CITY CEBU", itemCode: "REPAIRS - GOODS", desc: "REPAIRS - GOODS", vatType: "12% VAT on Purchases (Goods)", atc: "WC158", atcRate: 0.01, vatable: 10000, nonvat: 0, terms: "Credit", coaCode: "6160", bankAccount: "Cash on Hand" }),
    ],
    disbursements: [
      computeDisbRow({ id: uid(), date: "01/07/2026", tin: "", vendor: "Meralco", cvNo: "CV-001", desc: "January electricity bill", atc: "", atcRate: 0, amount: 3500, bankAccount: "Cash in Bank", coaCode: "6090" }),
      computeDisbRow({ id: uid(), date: "01/20/2026", tin: "259-225-658", vendor: "BAJARO JOHN PAULO D.", cvNo: "CV-002", desc: "Bookkeeping services, January", atc: "WI011", atcRate: 0.05, amount: 8000, bankAccount: "Cash in Bank", coaCode: "6160" }),
    ],
    receipts: [
      computeReceiptRow({ id: uid(), date: "01/08/2026", from: "Juan Dela Cruz Store", orNo: "OR-2201", desc: "Collection of December receivable", atc: "", atcRate: 0, amount: 5000, cwt: 0, bankAccount: "Cash on Hand", coaCode: "1200" }),
      computeReceiptRow({ id: uid(), date: "01/22/2026", from: "ARQUEE CORP", orNo: "OR-2202", desc: "Rental income, January", atc: "WI100", atcRate: 0.05, amount: 20000, cwt: 1000, bankAccount: "Cash in Bank", coaCode: "1200" }),
    ],
    generalJournal: [
      { id: uid(), jvNo: "JV-001", date: "01/31/2026", particulars: "To record January rent adjustment", lines: [
        { id: uid(), account: "6070", debit: 1000, credit: 0 },
        { id: uid(), account: "1010", debit: 0, credit: 1000 },
      ]},
    ],
  };
}

// Normalizes any legacy (e.g. YYYY-MM-DD) date strings in loaded/imported data to MM/DD/YYYY.
function normalizeDatesInData(d) {
  const fixRows = (arr) => (arr || []).map((r) => ({ ...r, date: toMDY(r.date) }));
  return {
    ...d,
    sales: fixRows(d.sales),
    purchases: fixRows(d.purchases),
    disbursements: fixRows(d.disbursements),
    receipts: fixRows(d.receipts),
    generalJournal: (d.generalJournal || []).map((jv) => ({ ...jv, date: toMDY(jv.date) })),
  };
}

/* ============================== ROW COMPUTE HELPERS ============================== */

function computeSalesRow(r) {
  const vatable = num(r.vatable), exempt = num(r.exempt), zeroRated = num(r.zeroRated);
  const outputVat = round2(vatable * 0.12);
  const total = round2(vatable + exempt + zeroRated + outputVat);
  return { ...r, outputVat, total };
}

function computePurchaseRow(r) {
  const vatable = num(r.vatable), nonvat = num(r.nonvat), atcRate = num(r.atcRate);
  const inputVat = round2(vatable * 0.12);
  const total = round2(vatable + nonvat + inputVat);
  const ewt = round2((vatable + nonvat) * atcRate);
  const net = round2(total - ewt);
  return { ...r, inputVat, total, ewt, net };
}

function computeDisbRow(r) {
  const amount = num(r.amount), atcRate = num(r.atcRate);
  const ewt = round2(amount * atcRate);
  return { ...r, ewt, net: round2(amount - ewt) };
}

function computeReceiptRow(r) {
  const amount = num(r.amount), cwt = num(r.cwt);
  return { ...r, net: round2(amount - cwt) };
}

// Philippine TINs are a 9-digit core identifier, optionally followed by a 3-4 digit branch code
// (e.g. "123-456-789" vs "123-456-789-000"). Matching on the full string would silently fail to
// find a customer whenever one side includes the branch code and the other doesn't — so this
// compares on the core 9 digits only, which is what actually identifies the taxpayer.
function normalizeTin(tin) {
  return String(tin ?? "").replace(/\D/g, "").slice(0, 9);
}

// BIR Form 2307 log — a manually-maintained record of certificates received. CWT is a plain
// product (Income Payment × Rate), unlike Cash Receipts' VAT-exclusive CWT formula, since a 2307
// certificate states the withheld amount directly rather than being derived from a VAT-inclusive
// collection.
function compute2307Row(r) {
  const cwt = round2(num(r.incomePayment) * num(r.atcRate));
  return { ...r, cwt };
}

// BIR Form 2316 / Alphalist of Employees. Taxable compensation is simply gross minus the
// non-taxable/exempt portion (statutory minimum wage, de minimis within limits, mandatory
// contributions, 13th month & other benefits within the ₱90,000 threshold, etc. — the user
// determines what's non-taxable and enters the net figure directly, matching how the BIR's own
// Alphalist Data Entry Module expects these two summary figures rather than a full itemized break).
// BIR Form 1604-C, Schedule 1 (regular employees) / Schedule 2 (minimum wage earners) — field
// numbers in comments match the official Alphalist layout so the mapping stays traceable.
// Non-taxable and taxable amounts are entered as their individual components (13th month, de
// minimis, mandatory contributions, basic salary, etc.) and every subtotal is derived exactly the
// way the BIR form derives it, rather than asking for pre-summed figures.
function computeEmployeeRow(r) {
  const mwe = r.isMWE === "Y";
  // Present employer — non-taxable (7b..7e), plus MWE's own non-taxable pay components when applicable
  const pNonTaxParts = [num(r.p13thMonthNonTax), num(r.pDeMinimis), num(r.pContributions), num(r.pSalariesNonTax)];
  if (mwe) pNonTaxParts.push(num(r.mweBasicPay), num(r.mweHolidayPay), num(r.mweOvertimePay), num(r.mweNightDiff), num(r.mweHazardPay));
  const pTotalNonTax = round2(pNonTaxParts.reduce((a, v) => a + v, 0)); // 7f
  // Present employer — taxable (7g..7i)
  const pTotalTax = round2(num(r.pBasicSalary) + num(r.p13thMonthExcess) + num(r.pSalariesTax)); // 7j
  const pGross = round2(pTotalNonTax + pTotalTax); // 7a

  // Previous employer, same shape (12b..12j)
  const prevTotalNonTax = round2(num(r.prev13thMonthNonTax) + num(r.prevDeMinimis) + num(r.prevContributions) + num(r.prevSalariesNonTax)); // 12f
  const prevTotalTax = round2(num(r.prevBasicSalary) + num(r.prev13thMonthExcess) + num(r.prevSalariesTax)); // 12j
  const prevGross = round2(prevTotalNonTax + prevTotalTax); // 12a

  const totalTaxableComp = round2(pTotalTax + prevTotalTax); // 13
  const totalWithheld = round2(num(r.taxWithheldPrev) + num(r.taxWithheldPresent)); // 15a+15b
  const netAfterCredit = round2(num(r.taxDue) - totalWithheld - num(r.taxCredit5pct));
  const amtWithheldDec = Math.max(0, netAfterCredit); // 17a
  const amtRefunded = Math.max(0, -netAfterCredit); // 17b

  return {
    ...r,
    pTotalNonTax, pTotalTax, pGross,
    prevTotalNonTax, prevTotalTax, prevGross,
    totalTaxableComp, amtWithheldDec, amtRefunded,
  };
}

// Landed Cost = Dutiable Value + Customs Duty + Other BOC Charges.
// VAT Paid on Importation = Dutiable Value × VAT Rate (matches the reference Importation Ledger).
function computeImportationRow(r) {
  const dutiableValue = num(r.dutiableValue), customsDuty = num(r.customsDuty), otherCharges = num(r.otherCharges);
  const vatRate = r.vatRate === "" || r.vatRate == null ? 0.12 : num(r.vatRate);
  const landedCost = round2(dutiableValue + customsDuty + otherCharges);
  const vatPaid = round2(dutiableValue * vatRate);
  return { ...r, landedCost, vatPaid };
}

// Straight-line depreciation, the standard method for simple bookkeeping. Accumulated depreciation
// reflects what's ACTUALLY been posted to the General Journal (via postedPeriods), not a theoretical
// time-elapsed calculation — so the ledger always ties out to what's really in the books.
function computeAssetRow(r) {
  const cost = num(r.cost), salvage = num(r.salvageValue), lifeYears = num(r.usefulLifeYears);
  const depreciableBase = Math.max(0, round2(cost - salvage));
  const totalMonths = Math.round(lifeYears * 12);
  const monthlyDep = totalMonths > 0 ? round2(depreciableBase / totalMonths) : 0;
  const postedCount = (r.postedPeriods || []).length;
  const accumulatedDep = Math.min(depreciableBase, round2(monthlyDep * postedCount));
  const bookValue = round2(cost - accumulatedDep);
  const fullyDepreciated = depreciableBase > 0 && accumulatedDep >= depreciableBase;
  return { ...r, monthlyDep, accumulatedDep, bookValue, fullyDepreciated };
}

/* ============================== EXCEL IMPORT / EXPORT ============================== */


const SALES_COLS = [
  { header: "Date (MM/DD/YYYY)", field: "date" },
  { header: "SI/OR No.", field: "siNo" },
  { header: "Customer Name", field: "customer" },
  { header: "TIN", field: "tin" },
  { header: "Address", field: "address" },
  { header: "Description", field: "desc" },
  { header: "VATable Sales", field: "vatable", type: "number" },
  { header: "VAT-Exempt", field: "exempt", type: "number" },
  { header: "Zero-Rated", field: "zeroRated", type: "number" },
  { header: "Terms (Cash/Credit)", field: "terms" },
  { header: "Account Code", field: "coaCode" },
  { header: "Bank Account", field: "bankAccount" },
];
const PURCHASE_COLS = [
  { header: "Date (MM/DD/YYYY)", field: "date" },
  { header: "Supplier Inv./OR", field: "invNo" },
  { header: "Supplier Name", field: "supplier" },
  { header: "TIN", field: "tin" },
  { header: "Address", field: "address" },
  { header: "Item Code", field: "itemCode" },
  { header: "Description", field: "desc" },
  { header: "VAT Type", field: "vatType" },
  { header: "ATC Code", field: "atc" },
  { header: "VATable", field: "vatable", type: "number" },
  { header: "Non-VAT", field: "nonvat", type: "number" },
  { header: "Terms (Cash/Credit)", field: "terms" },
  { header: "Account Code", field: "coaCode" },
  { header: "Bank Account", field: "bankAccount" },
];
const DISB_COLS = [
  { header: "Date (MM/DD/YYYY)", field: "date" },
  { header: "TIN", field: "tin" },
  { header: "Suppliers Name", field: "vendor" },
  { header: "Payment Ref.", field: "cvNo" },
  { header: "Description", field: "desc" },
  { header: "ATC Code", field: "atc" },
  { header: "Amount", field: "amount", type: "number" },
  { header: "Bank Account", field: "bankAccount" },
  { header: "Account Code", field: "coaCode" },
];
const RECEIPT_COLS = [
  { header: "Date (MM/DD/YYYY)", field: "date" },
  { header: "Received From", field: "from" },
  { header: "OR/Ref No.", field: "orNo" },
  { header: "Description", field: "desc" },
  { header: "ATC Code", field: "atc" },
  { header: "Amount", field: "amount", type: "number" },
  { header: "CWT", field: "cwt", type: "number" },
  { header: "Bank Account", field: "bankAccount" },
  { header: "Account Code", field: "coaCode" },
];
const FORM2307_COLS = [
  { header: "Month (MM/DD/YYYY)", field: "date" },
  { header: "TIN", field: "tin" },
  { header: "Customer's Name", field: "customerName" },
  { header: "Surname", field: "surname" },
  { header: "First Name", field: "firstName" },
  { header: "Middle Name", field: "middleName" },
  { header: "ATC Code", field: "atc" },
  { header: "Income Payment", field: "incomePayment", type: "number" },
];
const EMPLOYEE_COLS = [
  { header: "Year", field: "year", type: "number" },
  { header: "TIN", field: "tin" },
  { header: "Last Name", field: "lastName" },
  { header: "First Name", field: "firstName" },
  { header: "Middle Name", field: "middleName" },
  { header: "Nationality (foreigners only)", field: "nationality" },
  { header: "Registered Address", field: "address" },
  { header: "ZIP Code", field: "zipCode" },
  { header: "Statutory Minimum Wage Earner (Y/N)", field: "isMWE" },
  { header: "Employment Status (R/CP/C/P/S)", field: "empStatus" },
  { header: "Employment Period From", field: "empFrom" },
  { header: "Employment Period To", field: "empTo" },
  { header: "Reason of Separation", field: "sepReason" },
  { header: "13th Month Pay & Other Benefits (Non-Taxable)", field: "p13thMonthNonTax", type: "number" },
  { header: "De Minimis Benefits", field: "pDeMinimis", type: "number" },
  { header: "SSS/GSIS/PHIC/HDMF & Union Dues", field: "pContributions", type: "number" },
  { header: "Salaries (₱250K & below) & Other Non-Taxable", field: "pSalariesNonTax", type: "number" },
  { header: "Basic Salary (Taxable, net of contributions)", field: "pBasicSalary", type: "number" },
  { header: "13th Month Pay & Other Benefits (Taxable Excess)", field: "p13thMonthExcess", type: "number" },
  { header: "Salaries & Other Forms of Compensation (Taxable)", field: "pSalariesTax", type: "number" },
  { header: "Basic SMW Per Day", field: "smwPerDay", type: "number" },
  { header: "Basic SMW Per Month", field: "smwPerMonth", type: "number" },
  { header: "Basic SMW Per Year", field: "smwPerYear", type: "number" },
  { header: "Factor Used (No. of Days/Year)", field: "factorDaysPerYear", type: "number" },
  { header: "MWE Basic/SMW Pay", field: "mweBasicPay", type: "number" },
  { header: "MWE Holiday Pay", field: "mweHolidayPay", type: "number" },
  { header: "MWE Overtime Pay", field: "mweOvertimePay", type: "number" },
  { header: "MWE Night Shift Differential", field: "mweNightDiff", type: "number" },
  { header: "MWE Hazard Pay", field: "mweHazardPay", type: "number" },
  { header: "Previous Employer Status", field: "prevEmpStatus" },
  { header: "Previous Employment Period From", field: "prevFrom" },
  { header: "Previous Employment Period To", field: "prevTo" },
  { header: "Previous Reason of Separation", field: "prevSepReason" },
  { header: "Previous: 13th Month & Other Benefits (Non-Taxable)", field: "prev13thMonthNonTax", type: "number" },
  { header: "Previous: De Minimis Benefits", field: "prevDeMinimis", type: "number" },
  { header: "Previous: SSS/GSIS/PHIC/HDMF & Union Dues", field: "prevContributions", type: "number" },
  { header: "Previous: Salaries (₱250K & below) & Other Non-Taxable", field: "prevSalariesNonTax", type: "number" },
  { header: "Previous: Basic Salary (Taxable)", field: "prevBasicSalary", type: "number" },
  { header: "Previous: 13th Month & Other Benefits (Taxable Excess)", field: "prev13thMonthExcess", type: "number" },
  { header: "Previous: Salaries & Other Forms (Taxable)", field: "prevSalariesTax", type: "number" },
  { header: "Tax Due", field: "taxDue", type: "number" },
  { header: "Tax Withheld — Previous Employer", field: "taxWithheldPrev", type: "number" },
  { header: "Tax Withheld — Present Employer", field: "taxWithheldPresent", type: "number" },
  { header: "5% Tax Credit (PERA Act of 2008)", field: "taxCredit5pct", type: "number" },
  { header: "Substituted Filing (Y/N)", field: "substitutedFiling" },
];
const GJ_COLS = [
  { header: "JV No.", field: "jvNo" },
  { header: "Date (MM/DD/YYYY)", field: "date" },
  { header: "Particulars", field: "particulars" },
  { header: "Account Code", field: "account" },
  { header: "Debit", field: "debit", type: "number" },
  { header: "Credit", field: "credit", type: "number" },
];

const JOURNAL_SHEETS = {
  sales: { title: "Sales Journal", cols: SALES_COLS, dataKey: "sales" },
  purchases: { title: "Purchase Journal", cols: PURCHASE_COLS, dataKey: "purchases" },
  disbursements: { title: "Cash Disbursements", cols: DISB_COLS, dataKey: "disbursements" },
  receipts: { title: "Cash Receipts", cols: RECEIPT_COLS, dataKey: "receipts" },
  generaljournal: { title: "General Journal", cols: GJ_COLS, dataKey: "generalJournal" },
  form2307: { title: "Creditable Withholding Taxes", cols: FORM2307_COLS, dataKey: "form2307" },
  employees: { title: "Alphalist of Employees", cols: EMPLOYEE_COLS, dataKey: "employees" },
};
const REFERENCE_SHEET_NAMES = ["Instructions", "Chart of Accounts", "Item Record", "Customers Master", "Suppliers Master", "ATC Reference"];

function buildReferenceSheets(wb, data) {
  const coaRows = [["Code", "Account Name", "Type", "Category"], ...data.coa.map((a) => [a.code, a.name, a.type, a.category])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(coaRows), "Chart of Accounts");

  const itemRows = [["Item", "Default Account Code", "Type"], ...data.items.map((i) => [i.item, i.account, i.type])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(itemRows), "Item Record");

  const custRows = [["TIN", "Customer Name", "Address", "Type", "Account"], ...data.customers.map((c) => [c.tin, c.name, c.address, c.type, c.account])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(custRows), "Customers Master");

  const suppRows = [["TIN", "Supplier Name", "Address", "Item Code", "Account Title", "Type"], ...data.suppliers.map((s) => [s.tin, s.name, s.address, s.itemCode, s.accountTitle, s.type])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(suppRows), "Suppliers Master");

  const atcRows = [["ATC Code", "Rate", "Description"], ...data.atc.map((a) => [a.code, a.rate, a.desc])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(atcRows), "ATC Reference");
}

function buildInstructionsSheet(journalKey, data) {
  const meta = JOURNAL_SHEETS[journalKey];
  const lines = [
    [`${meta.title} — Import Template`],
    [""],
    ["How to use this file"],
    [`1. Fill out the "${meta.title}" sheet starting on row 2. Keep the header row and column order as-is.`],
    ["2. Enter dates as MM/DD/YYYY, e.g. 01/31/2026."],
    ["3. \"Account Code\" must match a code on the \"Chart of Accounts\" sheet."],
  ];
  if (journalKey === "sales") {
    lines.push(["4. \"Customer Name\" should match a name on the \"Customers Master\" sheet so TIN and Address auto-fill on import."]);
    lines.push(["5. \"Terms\" must be Cash or Credit. \"Bank Account\" must be Cash on Hand or Cash in Bank (ignored when Terms = Credit)."]);
  } else if (journalKey === "purchases") {
    lines.push(["4. \"Supplier Name\" should match a name on the \"Suppliers Master\" sheet, and \"Item Code / Description\" should match the \"Item Record\" sheet, so details and account auto-fill on import."]);
    lines.push(["5. \"Terms\" must be Cash or Credit. \"Bank Account\" must be Cash on Hand or Cash in Bank (ignored when Terms = Credit)."]);
  } else if (journalKey === "disbursements" || journalKey === "receipts") {
    lines.push(["4. \"Bank Account\" must be Cash on Hand or Cash in Bank."]);
  } else if (journalKey === "generaljournal") {
    lines.push(["4. Give every line of the same voucher the same \"JV No.\" — matching JV numbers are grouped back into one voucher on import."]);
    lines.push(["5. Make sure each voucher's total Debit equals its total Credit."]);
  }
  lines.push([""]);
  lines.push(["VAT, EWT/CWT, and Net/Total amounts are calculated automatically by the app — leave those out."]);
  lines.push([""]);
  lines.push(["Adding your own Excel dropdowns (optional)"]);
  lines.push(["Select the column, then Data > Data Validation > List, and set the source to a range on the reference sheets below, for example:"]);
  lines.push([`  Account Code column   ->   ='Chart of Accounts'!$A$2:$A$${data.coa.length + 1}`]);
  if (journalKey === "sales") lines.push([`  Customer Name column   ->   ='Customers Master'!$B$2:$B$${data.customers.length + 1}`]);
  if (journalKey === "purchases") {
    lines.push([`  Supplier Name column   ->   ='Suppliers Master'!$B$2:$B$${data.suppliers.length + 1}`]);
    lines.push([`  Item Code column   ->   ='Item Record'!$A$2:$A$${data.items.length + 1}`]);
  }
  lines.push([""]);
  lines.push(["Reference sheets included in this workbook: Chart of Accounts, Item Record, Customers Master, Suppliers Master, ATC Reference."]);
  lines.push(["Importing this file ADDS new rows/vouchers — it never overwrites or removes what's already in the app."]);
  const ws = XLSX.utils.aoa_to_sheet(lines);
  ws["!cols"] = [{ wch: 100 }];
  return ws;
}

function buildEntryAOA(journalKey, data) {
  const meta = JOURNAL_SHEETS[journalKey];
  const header = meta.cols.map((c) => c.header);
  const rows = [];
  if (journalKey === "generaljournal") {
    (data.generalJournal || []).forEach((jv) => {
      (jv.lines || []).forEach((l) => {
        rows.push([jv.jvNo || "", jv.date || "", jv.particulars || "", l.account || "", l.debit || "", l.credit || ""]);
      });
    });
  } else {
    (data[meta.dataKey] || []).forEach((r) => {
      rows.push(meta.cols.map((c) => r[c.field] ?? ""));
    });
  }
  return [header, ...rows];
}

function downloadJournalTemplate(journalKey, data) {
  const meta = JOURNAL_SHEETS[journalKey];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildInstructionsSheet(journalKey, data), "Instructions");
  const entryWs = XLSX.utils.aoa_to_sheet(buildEntryAOA(journalKey, data));
  entryWs["!cols"] = meta.cols.map((c) => ({ wch: Math.max(14, c.header.length + 2) }));
  XLSX.utils.book_append_sheet(wb, entryWs, meta.title.slice(0, 31));
  buildReferenceSheets(wb, data);
  XLSX.writeFile(wb, `${meta.title.replace(/\s+/g, "-")}-Template.xlsx`);
}

async function importJournalExcel(journalKey, file, data) {
  const meta = JOURNAL_SHEETS[journalKey];
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.includes(meta.title)
    ? meta.title
    : (wb.SheetNames.find((n) => !REFERENCE_SHEET_NAMES.includes(n)) || wb.SheetNames[0]);
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error("No entry sheet found in this workbook.");
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, dateNF: "mm/dd/yyyy", defval: "" });
  const rows = aoa.slice(1).filter((r) => r.some((cell) => String(cell ?? "").trim() !== ""));

  if (journalKey === "generaljournal") {
    const groups = {};
    const order = [];
    rows.forEach((r) => {
      const [jvNo, date, particulars, account, debit, credit] = r;
      const key = String(jvNo ?? "").trim() || `JV-IMPORT-${uid()}`;
      if (!groups[key]) { groups[key] = { id: uid(), jvNo: key, date: toMDY(String(date ?? "").trim()), particulars: String(particulars ?? "").trim(), lines: [] }; order.push(key); }
      if (!String(account ?? "").trim() && !num(debit) && !num(credit)) return;
      groups[key].lines.push({ id: uid(), account: String(account ?? "").trim(), debit: round2(num(debit)), credit: round2(num(credit)) });
    });
    const newVouchers = order.map((k) => groups[k]).filter((v) => v.lines.length > 0);
    return { newVouchers };
  }

  const rawRows = rows.map((r) => {
    const obj = { id: uid() };
    meta.cols.forEach((c, i) => {
      if (c.field === "date") { obj.date = toMDY(String(r[i] ?? "").trim()); return; }
      obj[c.field] = c.type === "number" ? round2(num(r[i])) : String(r[i] ?? "").trim();
    });
    return obj;
  });

  let newRows;
  if (journalKey === "sales") {
    newRows = rawRows.map((r) => {
      const cust = data.customers.find((c) => (c.name || "").toLowerCase() === (r.customer || "").toLowerCase());
      if (cust) { if (!r.tin) r.tin = cust.tin; if (!r.address) r.address = cust.address; }
      if (!r.coaCode) r.coaCode = "4000";
      if (!r.terms) r.terms = "Cash";
      if (!r.bankAccount) r.bankAccount = "Cash on Hand";
      return computeSalesRow(r);
    });
  } else if (journalKey === "purchases") {
    newRows = rawRows.map((r) => {
      const supp = data.suppliers.find((s) => (s.name || "").toLowerCase() === (r.supplier || "").toLowerCase());
      if (supp) { if (!r.tin) r.tin = supp.tin; if (!r.address) r.address = supp.address; if (!r.itemCode) r.itemCode = supp.itemCode || ""; }
      const item = data.items.find((i) => (i.item || "").toLowerCase() === (r.itemCode || r.desc || "").toLowerCase());
      if (item) { if (!r.coaCode) r.coaCode = item.account; if (!r.desc) r.desc = item.item; }
      const atc = data.atc.find((a) => (a.code || "").toLowerCase() === (r.atc || "").toLowerCase());
      r.atcRate = atc ? num(atc.rate) : 0;
      if (atc) r.atc = atc.code;
      if (!r.coaCode) r.coaCode = "5001";
      if (!r.terms) r.terms = "Cash";
      if (!r.bankAccount) r.bankAccount = "Cash on Hand";
      return computePurchaseRow(r);
    });
  } else if (journalKey === "disbursements") {
    newRows = rawRows.map((r) => {
      const supp = data.suppliers.find((s) => (s.name || "").toLowerCase() === (r.vendor || "").toLowerCase());
      if (supp && !r.tin) r.tin = supp.tin;
      const atc = data.atc.find((a) => (a.code || "").toLowerCase() === (r.atc || "").toLowerCase());
      r.atcRate = atc ? num(atc.rate) : 0;
      if (atc) r.atc = atc.code;
      if (!r.bankAccount) r.bankAccount = "Cash on Hand";
      if (!r.coaCode) r.coaCode = "6170";
      return computeDisbRow(r);
    });
  } else if (journalKey === "receipts") {
    newRows = rawRows.map((r) => {
      const atc = data.atc.find((a) => (a.code || "").toLowerCase() === (r.atc || "").toLowerCase());
      r.atcRate = atc ? num(atc.rate) : 0;
      if (atc) r.atc = atc.code;
      if (atc && !r.cwt) r.cwt = round2(cwtTaxBase(data, r.amount) * r.atcRate);
      if (!r.bankAccount) r.bankAccount = "Cash on Hand";
      if (!r.coaCode) r.coaCode = "1200";
      return computeReceiptRow(r);
    });
  } else if (journalKey === "form2307") {
    newRows = rawRows.map((r) => {
      const cust = data.customers.find((c) => normalizeTin(c.tin) && normalizeTin(c.tin) === normalizeTin(r.tin));
      if (cust) {
        if (!r.customerName) r.customerName = cust.registeredName || cust.name || "";
        if (!r.surname) r.surname = cust.surname || "";
        if (!r.firstName) r.firstName = cust.firstName || "";
        if (!r.middleName) r.middleName = cust.middleName || "";
        if (!r.custType) r.custType = cust.type || "";
      }
      const atc = data.atc.find((a) => (a.code || "").toLowerCase() === (r.atc || "").toLowerCase());
      r.atcRate = atc ? num(atc.rate) : 0;
      if (atc) r.atc = atc.code;
      return compute2307Row(r);
    });
  } else if (journalKey === "employees") {
    newRows = rawRows.map((r) => {
      if (!r.year) r.year = new Date().getFullYear();
      const mwe = String(r.isMWE ?? "").trim().toUpperCase();
      r.isMWE = mwe === "Y" || mwe === "YES" || mwe === "TRUE" ? "Y" : "N";
      const subFiling = String(r.substitutedFiling ?? "").trim().toUpperCase();
      r.substitutedFiling = subFiling === "Y" || subFiling === "YES" ? "Y" : (subFiling === "N" || subFiling === "NO" ? "N" : "");
      return computeEmployeeRow(r);
    });
  }
  return { newRows };
}

function useJournalImport(journalKey, data, onImported) {
  const fileInputRef = useRef(null);
  const [status, setStatus] = useState(null);
  const triggerImport = () => fileInputRef.current?.click();
  const handleFileChange = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const result = await importJournalExcel(journalKey, file, data);
      onImported(result);
      const count = result.newVouchers ? result.newVouchers.length : result.newRows.length;
      const noun = journalKey === "generaljournal" ? (count === 1 ? "voucher" : "vouchers") : (count === 1 ? "row" : "rows");
      setStatus({ type: "success", text: `Imported ${count} ${noun} from "${file.name}".` });
    } catch (err) {
      setStatus({ type: "error", text: `Couldn't read "${file.name}". Make sure it's an .xlsx file using this journal's template with the column order unchanged.` });
    }
    setTimeout(() => setStatus(null), 5000);
  };
  return { fileInputRef, status, triggerImport, handleFileChange };
}

function ImportExportBar({ journalKey, data, importHook }) {
  const title = JOURNAL_SHEETS[journalKey].title;
  return (
    <div className="io-toolbar">
      <button className="io-btn" onClick={() => downloadJournalTemplate(journalKey, data)}><FileDown size={13} /> Download {title} template</button>
      <button className="io-btn" onClick={importHook.triggerImport}><FileUp size={13} /> Import {title} Excel</button>
      <input ref={importHook.fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={importHook.handleFileChange} />
    </div>
  );
}
function ImportStatus({ status }) {
  if (!status) return null;
  return <div className={"import-msg" + (status.type === "error" ? " error" : "")}>{status.text}</div>;
}

/* ============================== PDF / A4 EXCEL EXPORT ============================== */

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
    const s = document.createElement("script");
    s.src = src;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(s);
  });
}

let jsPDFLoadPromise = null;
function loadJsPDF() {
  if (window.jspdf && window.jspdf.jsPDF) return Promise.resolve(window.jspdf.jsPDF);
  if (!jsPDFLoadPromise) {
    jsPDFLoadPromise = loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js")
      .then(() => loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"))
      .then(() => {
        if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("PDF engine did not initialize");
        return window.jspdf.jsPDF;
      })
      .catch((err) => { jsPDFLoadPromise = null; throw err; });
  }
  return jsPDFLoadPromise;
}

// Builds a flat {headers, body, totalsRow} table for a journal, matching what's shown on screen
// (including computed columns like VAT/Net/Total), for PDF and print-Excel export.
function buildExportTable(journalKey, rows, coaByCode) {
  const acctLabel = (code) => { const a = coaByCode[code]; return a ? `${code} · ${a.name}` : (code || ""); };

  if (journalKey === "sales") {
    const headers = ["Date","SI/OR No.","Customer","TIN","Address","Description","VATable Sales","VAT-Exempt","Zero-Rated","Output VAT","Total","Terms","Account","Bank Account"];
    const body = rows.map((r) => [r.date, r.siNo, r.customer, r.tin, r.address, r.desc, fmtPlain(r.vatable), fmtPlain(r.exempt), fmtPlain(r.zeroRated), fmtPlain(r.outputVat), fmtPlain(r.total), r.terms, acctLabel(r.coaCode), r.bankAccount]);
    const t = rows.reduce((a, r) => ({ vatable: a.vatable + num(r.vatable), exempt: a.exempt + num(r.exempt), zeroRated: a.zeroRated + num(r.zeroRated), outputVat: a.outputVat + num(r.outputVat), total: a.total + num(r.total) }), { vatable: 0, exempt: 0, zeroRated: 0, outputVat: 0, total: 0 });
    const totalsRow = ["", "", "", "", "", "TOTALS", fmtPlain(t.vatable), fmtPlain(t.exempt), fmtPlain(t.zeroRated), fmtPlain(t.outputVat), fmtPlain(t.total), "", "", ""];
    return { headers, body, totalsRow, numericCols: [6,7,8,9,10] };
  }
  if (journalKey === "purchases") {
    const headers = ["Date","Supplier Inv./OR","Supplier","TIN","Address","Item Code","Description","VAT Type","ATC","Rate","VATable","Non-VAT","Input VAT","Total","EWT","Net Amount","Terms","Account","Bank Account"];
    const body = rows.map((r) => [r.date, r.invNo, r.supplier, r.tin, r.address, r.itemCode, r.desc, r.vatType, r.atc, r.atc ? `${round2(num(r.atcRate) * 100)}%` : "", fmtPlain(r.vatable), fmtPlain(r.nonvat), fmtPlain(r.inputVat), fmtPlain(r.total), fmtPlain(r.ewt), fmtPlain(r.net), r.terms, acctLabel(r.coaCode), r.bankAccount]);
    const t = rows.reduce((a, r) => ({ vatable: a.vatable + num(r.vatable), nonvat: a.nonvat + num(r.nonvat), inputVat: a.inputVat + num(r.inputVat), total: a.total + num(r.total), ewt: a.ewt + num(r.ewt), net: a.net + num(r.net) }), { vatable: 0, nonvat: 0, inputVat: 0, total: 0, ewt: 0, net: 0 });
    const totalsRow = ["", "", "", "", "", "", "", "", "TOTALS", fmtPlain(t.vatable), fmtPlain(t.nonvat), fmtPlain(t.inputVat), fmtPlain(t.total), fmtPlain(t.ewt), fmtPlain(t.net), "", "", ""];
    return { headers, body, totalsRow, numericCols: [10,11,12,13,14,15] };
  }
  if (journalKey === "disbursements") {
    const headers = ["Date","TIN","Suppliers Name","Payment Ref.","Description","ATC","Rate","Amount","EWT","Net Amount","Bank Account","Account"];
    const body = rows.map((r) => [r.date, r.tin, r.vendor, r.cvNo, r.desc, r.atc, r.atc ? `${round2(num(r.atcRate) * 100)}%` : "", fmtPlain(r.amount), fmtPlain(r.ewt), fmtPlain(r.net), r.bankAccount, acctLabel(r.coaCode)]);
    const t = rows.reduce((a, r) => ({ amount: a.amount + num(r.amount), ewt: a.ewt + num(r.ewt), net: a.net + num(r.net) }), { amount: 0, ewt: 0, net: 0 });
    const totalsRow = ["", "", "", "", "", "", "TOTALS", fmtPlain(t.amount), fmtPlain(t.ewt), fmtPlain(t.net), "", ""];
    return { headers, body, totalsRow, numericCols: [7,8,9] };
  }
  if (journalKey === "receipts") {
    const headers = ["Date","Received From","OR/Ref No.","Description","ATC","Rate","Amount","CWT","Net Amount","Bank Account","Account"];
    const body = rows.map((r) => [r.date, r.from, r.orNo, r.desc, r.atc, r.atc ? `${round2(num(r.atcRate) * 100)}%` : "", fmtPlain(r.amount), fmtPlain(r.cwt), fmtPlain(r.net), r.bankAccount, acctLabel(r.coaCode)]);
    const t = rows.reduce((a, r) => ({ amount: a.amount + num(r.amount), cwt: a.cwt + num(r.cwt), net: a.net + num(r.net) }), { amount: 0, cwt: 0, net: 0 });
    const totalsRow = ["", "", "", "", "", "TOTALS", fmtPlain(t.amount), fmtPlain(t.cwt), fmtPlain(t.net), "", ""];
    return { headers, body, totalsRow, numericCols: [6,7,8] };
  }
  if (journalKey === "form2307") {
    const headers = ["Month","TIN","Customer's Name","Surname","First Name","Middle Name","ATC","Rate","Income Payment","CWT"];
    const body = rows.map((r) => [r.date, r.tin, r.customerName, r.surname, r.firstName, r.middleName, r.atc, r.atc ? `${round2(num(r.atcRate) * 100)}%` : "", fmtPlain(r.incomePayment), fmtPlain(r.cwt)]);
    const t = rows.reduce((a, r) => ({ incomePayment: a.incomePayment + num(r.incomePayment), cwt: a.cwt + num(r.cwt) }), { incomePayment: 0, cwt: 0 });
    const totalsRow = ["", "", "", "", "", "", "", "TOTALS", fmtPlain(t.incomePayment), fmtPlain(t.cwt)];
    return { headers, body, totalsRow, numericCols: [8,9] };
  }
  if (journalKey === "employees") {
    const headers = ["Year","TIN","Last Name","First Name","Middle Name","MWE","Gross (Present)","Non-Taxable","Taxable","Total Taxable Comp.","Tax Due","Tax Withheld"];
    const body = rows.map((r) => [r.year, r.tin, r.lastName, r.firstName, r.middleName, r.isMWE, fmtPlain(r.pGross), fmtPlain(r.pTotalNonTax), fmtPlain(r.pTotalTax), fmtPlain(r.totalTaxableComp), fmtPlain(r.taxDue), fmtPlain(round2(num(r.taxWithheldPrev) + num(r.taxWithheldPresent)))]);
    const t = rows.reduce((a, r) => ({ gross: a.gross + num(r.pGross), nonTax: a.nonTax + num(r.pTotalNonTax), tax: a.tax + num(r.pTotalTax), taxableComp: a.taxableComp + num(r.totalTaxableComp), due: a.due + num(r.taxDue), wh: a.wh + num(r.taxWithheldPrev) + num(r.taxWithheldPresent) }), { gross: 0, nonTax: 0, tax: 0, taxableComp: 0, due: 0, wh: 0 });
    const totalsRow = ["", "", "", "", "", "TOTALS", fmtPlain(t.gross), fmtPlain(t.nonTax), fmtPlain(t.tax), fmtPlain(t.taxableComp), fmtPlain(t.due), fmtPlain(t.wh)];
    return { headers, body, totalsRow, numericCols: [6,7,8,9,10,11] };
  }
  if (journalKey === "generaljournal") {
    const headers = ["JV No.","Date","Particulars","Account","Debit","Credit"];
    const body = [];
    let debitSum = 0, creditSum = 0;
    rows.forEach((jv) => {
      (jv.lines || []).forEach((l) => {
        debitSum += num(l.debit); creditSum += num(l.credit);
        body.push([jv.jvNo, jv.date, jv.particulars, acctLabel(l.account), fmtPlain(l.debit), fmtPlain(l.credit)]);
      });
      body.push(["", "", "", "", "", ""]); // blank row after each journal entry
    });
    const totalsRow = ["", "", "TOTALS", "", fmtPlain(round2(debitSum)), fmtPlain(round2(creditSum))];
    return { headers, body, totalsRow, numericCols: [4,5] };
  }
  return { headers: [], body: [], totalsRow: [], numericCols: [] };
}

async function exportJournalPDF(journalKey, data, rows, coaByCode) {
  const jsPDFCtor = await withTimeout(loadJsPDF(), 8000);
  const meta = JOURNAL_SHEETS[journalKey];
  const { headers, body, totalsRow } = buildExportTable(journalKey, rows, coaByCode);
  const doc = new jsPDFCtor({ orientation: "landscape", unit: "mm", format: "a4" });
  const marginLeft = 8;
  doc.setFontSize(13);
  doc.setFont(undefined, "bold");
  doc.text(data.company.name || "Your Company Name Inc.", marginLeft, 12);
  doc.setFontSize(9.5);
  doc.setFont(undefined, "normal");
  doc.text(`${meta.title} — Printed ${todayMDY()}`, marginLeft, 18);
  const fontSize = headers.length > 12 ? 6.5 : headers.length > 8 ? 8 : 9.5;
  doc.autoTable({
    head: [headers], body, foot: [totalsRow],
    startY: 23, margin: { left: marginLeft, right: marginLeft },
    styles: { fontSize, cellPadding: 1.6, overflow: "linebreak" },
    headStyles: { fillColor: [31, 95, 168], textColor: 255, fontStyle: "bold" },
    footStyles: { fillColor: [232, 241, 251], textColor: [15, 42, 77], fontStyle: "bold" },
    theme: "grid",
  });
  doc.save(`${meta.title.replace(/\s+/g, "-")}-A4-${todayFileStamp()}.pdf`);
}

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error("timed out")), ms)),
  ]);
}

function patchXlsxPageSetup(arrayBuf, JSZipCtor) {
  return JSZipCtor.loadAsync(arrayBuf).then((zip) => {
    const sheetPath = "xl/worksheets/sheet1.xml";
    const file = zip.file(sheetPath);
    if (!file) return zip.generateAsync({ type: "blob" });
    return file.async("string").then((xml) => {
      const sheetPrXml = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>';
      const pageSetupXml = '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>';
      if (!xml.includes("<sheetPr")) xml = xml.replace(/(<worksheet[^>]*>)/, "$1" + sheetPrXml);
      if (xml.includes("<pageMargins")) xml = xml.replace(/(<pageMargins[^/]*\/>)/, "$1" + pageSetupXml);
      else xml = xml.replace("</worksheet>", pageSetupXml + "</worksheet>");
      zip.file(sheetPath, xml);
      return zip.generateAsync({ type: "blob" });
    });
  });
}

let jsZipLoadPromise = null;
function loadJSZip() {
  if (window.JSZip) return Promise.resolve(window.JSZip);
  if (!jsZipLoadPromise) {
    jsZipLoadPromise = loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js")
      .then(() => { if (!window.JSZip) throw new Error("JSZip did not initialize"); return window.JSZip; })
      .catch((err) => { jsZipLoadPromise = null; throw err; });
  }
  return jsZipLoadPromise;
}

// Builds the workbook via SheetJS (headers/data/totals, column widths, merges), then — since the
// free SheetJS build parses "!pageSetup" on read but silently drops it on write — patches the raw
// worksheet XML afterward via JSZip to actually embed real landscape/A4 page setup. If JSZip can't
// be loaded (e.g. no internet on first use), still delivers a correct, working file without that patch.
async function exportJournalExcelPrint(journalKey, data, rows, coaByCode) {
  const meta = JOURNAL_SHEETS[journalKey];
  const { headers, body, totalsRow } = buildExportTable(journalKey, rows, coaByCode);
  const wb = XLSX.utils.book_new();
  const aoa = [
    [data.company.name || "Your Company Name Inc."],
    [meta.title],
    [`Printed ${todayMDY()}`],
    [],
    headers,
    ...body,
    totalsRow,
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = headers.map((h) => ({ wch: Math.max(10, Math.min(24, h.length + 4)) }));
  ws["!margins"] = { left: 0.3, right: 0.3, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
  if (headers.length > 1) {
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: headers.length - 1 } },
    ];
  }
  XLSX.utils.book_append_sheet(wb, ws, meta.title.slice(0, 31));
  const filename = `${meta.title.replace(/\s+/g, "-")}-A4-${todayFileStamp()}.xlsx`;

  try {
    const JSZipCtor = await withTimeout(loadJSZip(), 8000);
    const arrayBuf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = await withTimeout(patchXlsxPageSetup(arrayBuf, JSZipCtor), 8000);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    return { patched: true };
  } catch (err) {
    XLSX.writeFile(wb, filename);
    return { patched: false };
  }
}

function useJournalExport(journalKey, data, rows, coaByCode) {
  const [status, setStatus] = useState(null);
  const exportPDF = async () => {
    setStatus({ type: "pending", text: "Generating PDF…" });
    try {
      await exportJournalPDF(journalKey, data, rows, coaByCode);
      setStatus({ type: "success", text: "PDF downloaded." });
    } catch (err) {
      setStatus({ type: "error", text: "Couldn't generate the PDF — this needs an internet connection the first time. Please check your connection and try again." });
    }
    setTimeout(() => setStatus(null), 5000);
  };
  const exportExcel = async () => {
    setStatus({ type: "pending", text: "Generating Excel file…" });
    try {
      const result = await exportJournalExcelPrint(journalKey, data, rows, coaByCode);
      setStatus({ type: "success", text: result.patched ? "Excel file downloaded (A4 landscape)." : "Excel file downloaded — set Landscape/A4 manually when printing (needs internet the first time for automatic page setup)." });
    } catch (err) {
      setStatus({ type: "error", text: "Couldn't generate the Excel file." });
    }
    setTimeout(() => setStatus(null), 6000);
  };
  return { status, exportPDF, exportExcel };
}

function ExportBar({ exportHook }) {
  return (
    <div className="io-toolbar">
      <button className="io-btn" onClick={exportHook.exportPDF}><FileDown size={13} /> Export PDF (A4)</button>
      <button className="io-btn" onClick={exportHook.exportExcel}><FileDown size={13} /> Export Excel (A4)</button>
    </div>
  );
}
function ExportStatus({ status }) {
  if (!status) return null;
  return <div className={"import-msg" + (status.type === "error" ? " error" : "")}>{status.text}</div>;
}

// Flattens the General Ledger's grouped (header/opening/transactions/closing per account) view
// into row arrays for PDF/Excel export, tagging each row's kind so PDF export can style it.
function buildGeneralLedgerExportRows(groups) {
  const rows = [];
  const rowKinds = [];
  groups.forEach((g) => {
    rows.push([`${g.code} ${g.name}`, "", "", "", "", "", "", "", ""]);
    rowKinds.push("header");
    rows.push(["", "", `${g.code} ${g.name} Opening Balances`, "", "", "", "", "", fmtPlain(g.opening)]);
    rowKinds.push("opening");
    g.rows.forEach((r) => {
      rows.push([r.date, TRANSACTION_TYPE_LABELS[r.source] || r.source, r.desc || "", r.ref || "", r.acctName, r.code, r.debit ? fmtPlain(r.debit) : "", r.credit ? fmtPlain(r.credit) : "", fmtPlain(r.runningBalance)]);
      rowKinds.push("txn");
    });
    rows.push(["", "", `${g.code} ${g.name} Closing Balances`, "", "", "", "", "", fmtPlain(g.closing)]);
    rowKinds.push("closing");
  });
  return { rows, rowKinds };
}

const GL_EXPORT_HEADERS = ["Date","Transaction Type","Description","Reference","Account Name","Account Code","Debit (PHP)","Credit (PHP)","Running Balance (PHP)"];

async function exportGeneralLedgerPDF(data, groups, year) {
  const jsPDFCtor = await withTimeout(loadJsPDF(), 8000);
  const { rows, rowKinds } = buildGeneralLedgerExportRows(groups);
  const doc = new jsPDFCtor({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setFontSize(13); doc.setFont(undefined, "bold");
  doc.text(data.company.name || "Your Company Name Inc.", 8, 12);
  doc.setFontSize(9.5); doc.setFont(undefined, "normal");
  doc.text(`General Ledger — ${year} — Printed ${todayMDY()}`, 8, 18);
  doc.autoTable({
    head: [GL_EXPORT_HEADERS], body: rows,
    startY: 23, margin: { left: 8, right: 8 },
    styles: { fontSize: 7, cellPadding: 1.4, overflow: "linebreak" },
    headStyles: { fillColor: [31, 95, 168], textColor: 255, fontStyle: "bold" },
    theme: "grid",
    didParseCell: (hookData) => {
      if (hookData.section !== "body") return;
      const kind = rowKinds[hookData.row.index];
      if (kind === "header") { hookData.cell.styles.fontStyle = "bold"; hookData.cell.styles.fillColor = [232, 241, 251]; hookData.cell.styles.textColor = [15, 42, 77]; }
      else if (kind === "opening" || kind === "closing") { hookData.cell.styles.fontStyle = "bold"; hookData.cell.styles.fillColor = [245, 248, 252]; }
    },
  });
  doc.save(`General-Ledger-${year}-A4-${todayFileStamp()}.pdf`);
}

async function exportGeneralLedgerExcelPrint(data, groups, year) {
  const { rows } = buildGeneralLedgerExportRows(groups);
  const wb = XLSX.utils.book_new();
  const aoa = [
    [data.company.name || "Your Company Name Inc."],
    [`General Ledger — ${year}`],
    [`Printed ${todayMDY()}`],
    [],
    GL_EXPORT_HEADERS,
    ...rows,
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = GL_EXPORT_HEADERS.map((h) => ({ wch: Math.max(10, Math.min(26, h.length + 4)) }));
  ws["!margins"] = { left: 0.3, right: 0.3, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: GL_EXPORT_HEADERS.length - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: GL_EXPORT_HEADERS.length - 1 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: GL_EXPORT_HEADERS.length - 1 } },
  ];
  XLSX.utils.book_append_sheet(wb, ws, "General Ledger");
  const filename = `General-Ledger-${year}-A4-${todayFileStamp()}.xlsx`;
  try {
    const JSZipCtor = await withTimeout(loadJSZip(), 8000);
    const arrayBuf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = await withTimeout(patchXlsxPageSetup(arrayBuf, JSZipCtor), 8000);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    return { patched: true };
  } catch (err) {
    XLSX.writeFile(wb, filename);
    return { patched: false };
  }
}

function useGeneralLedgerExport(data, groups, year) {
  const [status, setStatus] = useState(null);
  const exportPDF = async () => {
    setStatus({ type: "pending", text: "Generating PDF…" });
    try {
      await exportGeneralLedgerPDF(data, groups, year);
      setStatus({ type: "success", text: "PDF downloaded." });
    } catch (err) {
      setStatus({ type: "error", text: "Couldn't generate the PDF — this needs an internet connection the first time. Please check your connection and try again." });
    }
    setTimeout(() => setStatus(null), 5000);
  };
  const exportExcel = async () => {
    setStatus({ type: "pending", text: "Generating Excel file…" });
    try {
      const result = await exportGeneralLedgerExcelPrint(data, groups, year);
      setStatus({ type: "success", text: result.patched ? "Excel file downloaded (A4 landscape)." : "Excel file downloaded — set Landscape/A4 manually when printing (needs internet the first time for automatic page setup)." });
    } catch (err) {
      setStatus({ type: "error", text: "Couldn't generate the Excel file." });
    }
    setTimeout(() => setStatus(null), 6000);
  };
  return { status, exportPDF, exportExcel };
}

/* ============================== POSTINGS ENGINE ============================== */

function buildPostings(data) {
  const postings = [];
  const push = (date, code, debit, credit, source, ref, journalKey, rowId, lineId, reclassable, party, desc) => {
    if (!date || !code) return;
    if (!debit && !credit) return;
    const id = `${journalKey}:${rowId}:${lineId || ""}:${code}:${postings.length}`;
    postings.push({ id, date, code, debit: debit || 0, credit: credit || 0, source, ref, journalKey, rowId, lineId: lineId || null, reclassable: !!reclassable, party: party || "", desc: desc || "" });
  };

  data.sales.forEach((r) => {
    if (!r.date) return;
    const bankCode = r.terms === "Credit" ? "1200" : (r.bankAccount === "Cash in Bank" ? "1010" : "1000");
    const salesAmt = round2(num(r.vatable) + num(r.exempt) + num(r.zeroRated));
    if (!num(r.total) && !salesAmt) return;
    push(r.date, bankCode, num(r.total), 0, "Sales Journal", r.siNo, "sales", r.id, null, false, r.customer, r.desc);
    push(r.date, r.coaCode || "4000", 0, salesAmt, "Sales Journal", r.siNo, "sales", r.id, null, true, r.customer, r.desc);
    push(r.date, "2200", 0, num(r.outputVat), "Sales Journal", r.siNo, "sales", r.id, null, false, r.customer, r.desc);
  });

  data.purchases.forEach((r) => {
    if (!r.date) return;
    const expenseAmt = round2(num(r.vatable) + num(r.nonvat));
    if (!expenseAmt && !num(r.inputVat)) return;
    push(r.date, r.coaCode || "5001", expenseAmt, 0, "Purchase Journal", r.invNo, "purchases", r.id, null, true, r.supplier, r.desc);
    push(r.date, "1400", num(r.inputVat), 0, "Purchase Journal", r.invNo, "purchases", r.id, null, false, r.supplier, r.desc);
    push(r.date, "2300", 0, num(r.ewt), "Purchase Journal", r.invNo, "purchases", r.id, null, false, r.supplier, r.desc);
    const creditCode = r.terms === "Credit" ? "2000" : (r.bankAccount === "Cash in Bank" ? "1010" : "1000");
    push(r.date, creditCode, 0, num(r.net), "Purchase Journal", r.invNo, "purchases", r.id, null, false, r.supplier, r.desc);
  });

  data.disbursements.forEach((r) => {
    if (!r.date || !num(r.amount)) return;
    push(r.date, r.coaCode || "6170", num(r.amount), 0, "Cash Disbursements", r.cvNo, "disbursements", r.id, null, true, r.vendor, r.desc);
    push(r.date, "2300", 0, num(r.ewt), "Cash Disbursements", r.cvNo, "disbursements", r.id, null, false, r.vendor, r.desc);
    const bankCode = r.bankAccount === "Cash in Bank" ? "1010" : "1000";
    push(r.date, bankCode, 0, num(r.net), "Cash Disbursements", r.cvNo, "disbursements", r.id, null, false, r.vendor, r.desc);
  });

  data.receipts.forEach((r) => {
    if (!r.date || !num(r.amount)) return;
    const bankCode = r.bankAccount === "Cash in Bank" ? "1010" : "1000";
    push(r.date, bankCode, num(r.net), 0, "Cash Receipts", r.orNo, "receipts", r.id, null, false, r.from, r.desc);
    push(r.date, "1500", num(r.cwt), 0, "Cash Receipts", r.orNo, "receipts", r.id, null, false, r.from, r.desc);
    push(r.date, r.coaCode || "1200", 0, num(r.amount), "Cash Receipts", r.orNo, "receipts", r.id, null, true, r.from, r.desc);
  });

  data.generalJournal.forEach((jv) => {
    (jv.lines || []).forEach((l) => {
      if (!jv.date || !l.account || (!num(l.debit) && !num(l.credit))) return;
      push(jv.date, l.account, num(l.debit), num(l.credit), "General Journal", jv.jvNo, "generaljournal", jv.id, l.id, true, "", jv.particulars);
    });
  });

  return postings;
}

// Applies a reclass (new account code) to the journal row/line a posting originated from.
function reclassPosting(setData, posting, newCode) {
  const { journalKey, rowId, lineId } = posting;
  setData((d) => {
    if (journalKey === "sales") return { ...d, sales: d.sales.map((r) => r.id === rowId ? computeSalesRow({ ...r, coaCode: newCode }) : r) };
    if (journalKey === "purchases") return { ...d, purchases: d.purchases.map((r) => r.id === rowId ? computePurchaseRow({ ...r, coaCode: newCode }) : r) };
    if (journalKey === "disbursements") return { ...d, disbursements: d.disbursements.map((r) => r.id === rowId ? computeDisbRow({ ...r, coaCode: newCode }) : r) };
    if (journalKey === "receipts") return { ...d, receipts: d.receipts.map((r) => r.id === rowId ? computeReceiptRow({ ...r, coaCode: newCode }) : r) };
    if (journalKey === "generaljournal") {
      return { ...d, generalJournal: d.generalJournal.map((jv) => jv.id === rowId ? { ...jv, lines: jv.lines.map((l) => l.id === lineId ? { ...l, account: newCode } : l) } : jv) };
    }
    return d;
  });
}


function monthEnd(year, monthIdx) { return new Date(year, monthIdx + 1, 0); }
function inRange(dateStr, start, end) {
  const d = parseAppDate(dateStr);
  return d && d >= start && d <= end;
}
function upTo(dateStr, end) {
  const d = parseAppDate(dateStr);
  return d && d <= end;
}

/* ============================== SMALL UI PRIMITIVES ============================== */

function SearchableSelect({ value, onChange, options, placeholder, disabled, align, onAddNew, addNewLabel, freeText: freeTextProp }) {
  const freeText = freeTextProp !== undefined ? freeTextProp : !!onAddNew; // free text allowed when explicitly requested, or implied by having an "add new" affordance
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [highlight, setHighlight] = useState(0);
  const [menuPos, setMenuPos] = useState(null);
  const wrapRef = useRef(null);
  const inputRef = useRef(null);

  const norm = (s) => String(s ?? "").toLowerCase();
  const optVal = (o) => ((typeof o === "string" || typeof o === "number") ? o : o.value);
  const optLabel = (o) => ((typeof o === "string" || typeof o === "number") ? o : o.label);

  const selected = options.find((o) => optVal(o) === value);
  const selectedLabel = freeText ? (value || "") : (selected ? optLabel(selected) : "");

  useEffect(() => { if (!open) setQuery(selectedLabel || ""); }, [selectedLabel, open]);

  const commitFreeText = () => {
    const q = query.trim();
    if (freeText && q !== (value || "")) onChange(q);
  };

  useEffect(() => {
    function onDocMouseDown(e) {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) {
        setOpen(false);
        if (freeText) commitFreeText(); else setQuery(selectedLabel || "");
      }
    }
    document.addEventListener("mousedown", onDocMouseDown);
    return () => document.removeEventListener("mousedown", onDocMouseDown);
  }, [selectedLabel, freeText, query, value]);

  const computePos = () => {
    if (!inputRef.current) return;
    const rect = inputRef.current.getBoundingClientRect();
    setMenuPos({ top: rect.bottom + 4, left: rect.left, width: Math.max(rect.width, 240) });
  };

  const openMenu = (clearQuery) => {
    if (disabled) return;
    computePos();
    setOpen(true);
    setHighlight(0);
    if (clearQuery && !freeText) setQuery("");
  };

  const filtered = query.trim() === "" ? options : options.filter((o) => norm(optLabel(o)).includes(norm(query)) || norm(optVal(o)).includes(norm(query)));
  const hasExactMatch = options.some((o) => norm(optLabel(o)) === norm(query.trim()) || norm(optVal(o)) === norm(query.trim()));
  const showAddNew = !!onAddNew && query.trim() !== "" && !hasExactMatch;

  const pick = (o) => { onChange(optVal(o)); setQuery(optLabel(o)); setOpen(false); };
  const pickAddNew = () => { const q = query.trim(); setOpen(false); onAddNew(q); };

  const handleKeyDown = (e) => {
    if (!open && (e.key === "ArrowDown" || e.key === "Enter")) { openMenu(true); return; }
    if (!open) return;
    const total = filtered.length + (showAddNew ? 1 : 0);
    if (e.key === "ArrowDown") { e.preventDefault(); setHighlight((h) => Math.min(h + 1, total - 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setHighlight((h) => Math.max(h - 1, 0)); }
    else if (e.key === "Enter") {
      e.preventDefault();
      if (showAddNew && highlight === filtered.length) pickAddNew();
      else if (filtered[highlight]) pick(filtered[highlight]);
      else if (freeText) { commitFreeText(); setOpen(false); }
    }
    else if (e.key === "Escape") { setOpen(false); setQuery(selectedLabel || ""); }
  };

  return (
    <div className="combo-wrap" ref={wrapRef}>
      <input
        ref={inputRef}
        type="text"
        className={"ledger-field combo-input" + (align === "right" ? " num" : "")}
        value={query}
        placeholder={placeholder || "Search…"}
        disabled={disabled}
        onFocus={() => openMenu(true)}
        onChange={(e) => { setQuery(e.target.value); if (!open) { computePos(); setOpen(true); } else computePos(); setHighlight(0); }}
        onKeyDown={handleKeyDown}
      />
      {open && menuPos && (
        <div className="combo-menu" style={{ position: "fixed", top: menuPos.top, left: menuPos.left, width: menuPos.width }}>
          {filtered.length === 0 && !showAddNew && <div className="combo-empty">No matches</div>}
          {filtered.map((o, i) => (
            <div key={optVal(o)} className={"combo-option" + (i === highlight ? " active" : "") + (optVal(o) === value ? " selected" : "")}
              onMouseDown={(e) => { e.preventDefault(); pick(o); }} onMouseEnter={() => setHighlight(i)}>
              {optLabel(o)}
            </div>
          ))}
          {showAddNew && (
            <div className={"combo-option combo-add-new" + (highlight === filtered.length ? " active" : "")}
              onMouseDown={(e) => { e.preventDefault(); pickAddNew(); }} onMouseEnter={() => setHighlight(filtered.length)}>
              + Add "{query.trim()}" as new {addNewLabel || "entry"}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// Safely evaluates a simple arithmetic expression like "112000/1.12" or "=(500*3)+250".
// Only digits, + - * / ( ) and decimal points are allowed — no eval(), no identifiers, nothing else.
function evalFormula(expr) {
  const cleaned = String(expr ?? "").trim().replace(/^=/, "");
  if (!cleaned) return null;
  if (!/^[0-9+\-*/().\s]+$/.test(cleaned)) return null;
  const s = cleaned.replace(/\s+/g, "");
  let i = 0;
  const parseExpr = () => {
    let v = parseTerm();
    while (s[i] === "+" || s[i] === "-") {
      const op = s[i++]; const rhs = parseTerm();
      v = op === "+" ? v + rhs : v - rhs;
    }
    return v;
  };
  const parseTerm = () => {
    let v = parseFactor();
    while (s[i] === "*" || s[i] === "/") {
      const op = s[i++]; const rhs = parseFactor();
      v = op === "*" ? v * rhs : v / rhs;
    }
    return v;
  };
  const parseFactor = () => {
    if (s[i] === "-") { i++; return -parseFactor(); }
    if (s[i] === "+") { i++; return parseFactor(); }
    if (s[i] === "(") {
      i++;
      const v = parseExpr();
      if (s[i] === ")") i++; else throw new Error("unbalanced");
      return v;
    }
    const start = i;
    while (i < s.length && /[0-9.]/.test(s[i])) i++;
    if (start === i) throw new Error("bad token");
    return parseFloat(s.slice(start, i));
  };
  try {
    const result = parseExpr();
    if (i !== s.length) return null;
    return isFinite(result) ? result : null;
  } catch (e) {
    return null;
  }
}

function NumberField({ value, onChange, align, placeholder, disabled }) {
  const [focused, setFocused] = useState(false);
  const [raw, setRaw] = useState(value === 0 || value == null || value === "" ? "" : String(value));

  useEffect(() => {
    if (!focused) setRaw(value === 0 || value == null || value === "" ? "" : String(value));
  }, [value, focused]);

  const displayValue = focused ? raw : (raw === "" ? "" : fmtPlain(num(raw)));
  const alignClass = align === "right" ? " num" : (align === "center" ? " center" : "");

  return (
    <input
      type="text" inputMode="decimal"
      className={"ledger-field" + alignClass}
      value={displayValue}
      placeholder={placeholder || "0.00"}
      disabled={disabled}
      onFocus={() => setFocused(true)}
      onChange={(e) => { setRaw(e.target.value); onChange(e.target.value); }}
      onBlur={(e) => {
        setFocused(false);
        const typed = e.target.value;
        if (typed !== "") {
          const isFormula = typed.trim().startsWith("=") || /[+\-*/()]/.test(typed.trim().slice(1));
          const evaluated = isFormula ? evalFormula(typed) : null;
          const rounded = round2(evaluated != null ? evaluated : num(typed));
          setRaw(String(rounded));
          onChange(rounded);
        }
      }}
    />
  );
}

// Stores the fraction (0.12) but displays/edits it as a percent (12.00%), matching how
// people naturally think about withholding/VAT rates.
function PercentField({ value, onChange, align, disabled }) {
  const [focused, setFocused] = useState(false);
  const toPct = (v) => (v === "" || v == null ? "" : String(round2(num(v) * 100)));
  const [raw, setRaw] = useState(toPct(value));

  useEffect(() => { if (!focused) setRaw(toPct(value)); }, [value, focused]);

  const displayValue = focused ? raw : (raw === "" ? "" : `${Number(raw).toFixed(2)}%`);
  const alignClass = align === "right" ? " num" : (align === "center" ? " center" : "");

  return (
    <input
      type="text" inputMode="decimal"
      className={"ledger-field" + alignClass}
      value={displayValue}
      placeholder="0.00%"
      disabled={disabled}
      onFocus={() => setFocused(true)}
      onChange={(e) => setRaw(e.target.value)}
      onBlur={(e) => {
        setFocused(false);
        const typed = e.target.value;
        if (typed !== "") {
          const pct = round2(num(typed));
          setRaw(String(pct));
          onChange(round2(pct / 100));
        }
      }}
    />
  );
}

export function Field({ value, onChange, type = "text", options = [], align, placeholder, list, disabled, onAddNew, addNewLabel, freeText }) {
  const base = "ledger-field" + (align === "right" ? " num" : (align === "center" ? " center" : ""));
  if (type === "combo") {
    return <SearchableSelect value={value} onChange={onChange} options={options} placeholder={placeholder} disabled={disabled} align={align} onAddNew={onAddNew} addNewLabel={addNewLabel} freeText={freeText} />;
  }
  if (type === "select") {
    return (
      <select className={base} value={value ?? ""} onChange={(e) => onChange(e.target.value)} disabled={disabled}>
        <option value="">—</option>
        {options.map((o) => {
          const val = (typeof o === "string" || typeof o === "number") ? o : o.value;
          const lab = (typeof o === "string" || typeof o === "number") ? o : o.label;
          return <option key={val} value={val}>{lab}</option>;
        })}
      </select>
    );
  }
  if (type === "date") {
    const formatTyped = (raw) => {
      const digits = raw.replace(/\D/g, "").slice(0, 8);
      if (digits.length > 4) return `${digits.slice(0,2)}/${digits.slice(2,4)}/${digits.slice(4)}`;
      if (digits.length > 2) return `${digits.slice(0,2)}/${digits.slice(2)}`;
      return digits;
    };
    return <input type="text" inputMode="numeric" className={base} value={value || ""} placeholder="MM/DD/YYYY" maxLength={10}
      onChange={(e) => onChange(formatTyped(e.target.value))}
      onBlur={(e) => { if (e.target.value) onChange(toMDY(e.target.value)); }}
      disabled={disabled} />;
  }
  if (type === "number") {
    return <NumberField value={value} onChange={onChange} align={align} placeholder={placeholder} disabled={disabled} />;
  }
  if (type === "percent") {
    return <PercentField value={value} onChange={onChange} align={align} disabled={disabled} />;
  }
  return <input type="text" className={base} value={value ?? ""} placeholder={placeholder} list={list} onChange={(e) => onChange(e.target.value)} disabled={disabled} />;
}

function ReadCell({ children, align }) {
  return <div className={"ledger-read" + (align === "right" ? " num" : (align === "center" ? " center" : ""))}>{children}</div>;
}

export function HakiLogo({ size = 34 }) {
  return (
    <svg viewBox="0 0 512 512" width={size} height={size} xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="boltGradIcon" x1="204" y1="309" x2="320" y2="164" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor="#8A6FFF" />
          <stop offset="100%" stopColor="#FF7A6E" />
        </linearGradient>
      </defs>
      <rect x="0" y="0" width="512" height="512" rx="112" fill="#0B0F19" />
      <g transform="translate(85,76) scale(1.5)">
        <rect x="60" y="50" width="36" height="140" rx="8" fill="#F5F4F8" />
        <polygon points="128,50 168,50 160,190 136,190" fill="#F5F4F8" />
        <polygon points="96,145 96,117 150,77 150,105" fill="url(#boltGradIcon)" />
      </g>
    </svg>
  );
}

function SectionHeader({ icon: Icon, title, subtitle, right }) {
  return (
    <div className="section-head">
      <div className="flex items-start gap-3">
        <div className="section-icon"><Icon size={18} strokeWidth={2} /></div>
        <div>
          <h1>{title}</h1>
          {subtitle && <p>{subtitle}</p>}
        </div>
      </div>
      {right}
    </div>
  );
}

function AddRowBtn({ onClick, children }) {
  return <button onClick={onClick} className="add-row-btn"><Plus size={14} /> {children || "Add row"}</button>;
}
function DelBtn({ onClick }) {
  return <button onClick={onClick} className="del-btn" title="Delete row"><Trash2 size={13} /></button>;
}

/* ============================== MASTER TABLE (generic) ============================== */

function downloadMasterTemplate(title, columns, rows) {
  const wb = XLSX.utils.book_new();
  const instrLines = [
    [`${title} — Import Template`],
    [""],
    ["How to use this file"],
    [`1. Fill out the "${title}" sheet starting on row 2. Keep the header row and column order as-is.`],
    ["2. For dropdown-style columns (Type, Account, etc.), type the exact value used in the app."],
    ["3. For percentage fields, enter the decimal value, e.g. 0.12 for 12%."],
    [""],
    ["Importing this file ADDS new rows — it never overwrites or removes what's already in the app."],
  ];
  const instrWs = XLSX.utils.aoa_to_sheet(instrLines);
  instrWs["!cols"] = [{ wch: 100 }];
  XLSX.utils.book_append_sheet(wb, instrWs, "Instructions");
  const header = columns.map((c) => c.label);
  const dataRows = rows.map((r) => columns.map((c) => r[c.key] ?? ""));
  const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
  ws["!cols"] = columns.map((c) => ({ wch: Math.max(14, c.label.length + 2) }));
  XLSX.utils.book_append_sheet(wb, ws, title.slice(0, 31));
  XLSX.writeFile(wb, `${title.replace(/\s+/g, "-")}-Template.xlsx`);
}

async function importMasterExcel(title, columns, file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.includes(title.slice(0, 31)) ? title.slice(0, 31) : (wb.SheetNames.find((n) => n !== "Instructions") || wb.SheetNames[0]);
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error("No entry sheet found in this workbook.");
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, dateNF: "mm/dd/yyyy", defval: "" });
  const rows = aoa.slice(1).filter((r) => r.some((cell) => String(cell ?? "").trim() !== ""));
  return rows.map((r) => {
    const obj = { id: uid() };
    columns.forEach((c, i) => {
      const raw = r[i];
      obj[c.key] = c.type === "number" ? round2(num(raw)) : (c.type === "percent" ? num(raw) : String(raw ?? "").trim());
    });
    return obj;
  });
}

function useMasterImport(title, columns, onImported) {
  const fileInputRef = useRef(null);
  const [status, setStatus] = useState(null);
  const triggerImport = () => fileInputRef.current?.click();
  const handleFileChange = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const newRows = await importMasterExcel(title, columns, file);
      onImported(newRows);
      setStatus({ type: "success", text: `Imported ${newRows.length} row${newRows.length === 1 ? "" : "s"} from "${file.name}".` });
    } catch (err) {
      setStatus({ type: "error", text: `Couldn't read "${file.name}". Make sure it's an .xlsx file using this table's template with the column order unchanged.` });
    }
    setTimeout(() => setStatus(null), 5000);
  };
  return { fileInputRef, status, triggerImport, handleFileChange };
}

function MasterImportExportBar({ title, columns, rows, importHook }) {
  return (
    <div className="io-toolbar">
      <button className="io-btn" onClick={() => downloadMasterTemplate(title, columns, rows)}><FileDown size={13} /> Download {title} template</button>
      <button className="io-btn" onClick={importHook.triggerImport}><FileUp size={13} /> Import {title} Excel</button>
      <input ref={importHook.fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={importHook.handleFileChange} />
    </div>
  );
}

// Prevents the classic "select text near a modal's edge, mouse drifts outside while
// releasing" from being misread as an outside click and closing the modal. Only closes
// when BOTH the mousedown and the click landed directly on the backdrop itself.
function useOverlayClose(onClose) {
  const downOnSelf = useRef(false);
  return {
    onMouseDown: (e) => { downOnSelf.current = e.target === e.currentTarget; },
    onClick: (e) => { if (downOnSelf.current && e.target === e.currentTarget) onClose(); },
  };
}

function AddRecordModal({ title, columns, initial, onCancel, onSubmit }) {
  const [values, setValues] = useState(initial);
  const set = (k, v) => setValues((prev) => ({ ...prev, [k]: v }));
  const overlay = useOverlayClose(onCancel);
  return (
    <div className="modal-overlay" {...overlay}>
      <div className="modal-box" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h2>{title}</h2>
          <button className="modal-close" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="modal-body">
          {columns.map((c) => (
            <LabeledField key={c.key} label={c.label}>
              <Field value={values[c.key]} type={c.type} options={c.options} onChange={(v) => set(c.key, v)} />
            </LabeledField>
          ))}
        </div>
        <div className="modal-foot">
          <button className="btn-outline" onClick={onCancel}>Cancel</button>
          <button className="primary-btn" onClick={() => onSubmit(values)}>Add</button>
        </div>
      </div>
    </div>
  );
}

function MasterTable({ title, columns, rows, onChange, onDelete, addLabel, onCreate, onBulkAdd, initialValues }) {
  const [modalOpen, setModalOpen] = useState(false);
  const [search, setSearch] = useState("");
  const importHook = useMasterImport(title, columns, (newRows) => onBulkAdd(newRows));
  const emptyValues = () => {
    const base = Object.fromEntries(columns.map((c) => [c.key, (c.type === "number" || c.type === "percent") ? 0 : ""]));
    return { ...base, ...(initialValues || {}) };
  };
  const handleSubmit = (values) => { onCreate({ id: uid(), ...values }); setModalOpen(false); };
  const filteredRows = useMemo(() => {
    if (!search.trim()) return rows;
    const q = search.trim().toLowerCase();
    return rows.filter((row) => columns.some((c) => String(row[c.key] ?? "").toLowerCase().includes(q)));
  }, [rows, search, columns]);
  const emptyMsg = rows.length === 0 ? "No entries yet — add your first row above." : "No entries match this search.";
  return (
    <div className="ledger-wrap">
      <div className="table-toolbar">
        <div className="toolbar-left">
          <MasterImportExportBar title={title} columns={columns} rows={rows} importHook={importHook} />
          <SearchBar value={search} onChange={setSearch} placeholder={`Search ${title.toLowerCase()}…`} />
        </div>
        <AddRowBtn onClick={() => setModalOpen(true)}>{addLabel || "Add row"}</AddRowBtn>
      </div>
      <ImportStatus status={importHook.status} />
      <div className="table-scroll">
        <table className="ledger-table">
          <thead>
            <tr>
              {columns.map((c) => <th key={c.key} style={{ minWidth: c.width || 130 }}>{c.label}</th>)}
              <th style={{ width: 36 }}></th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.length === 0 && (
              <tr><td colSpan={columns.length + 1} className="empty-row">{emptyMsg}</td></tr>
            )}
            {filteredRows.map((row) => (
              <tr key={row.id}>
                {columns.map((c) => (
                  <td key={c.key}>
                    <Field value={row[c.key]} type={c.type} options={c.options} align={c.align}
                      onChange={(v) => onChange(row.id, c.key, v)} />
                  </td>
                ))}
                <td className="text-center"><DelBtn onClick={() => onDelete(row.id)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {modalOpen && (
        <AddRecordModal title={addLabel || "Add record"} columns={columns} initial={emptyValues()}
          onCancel={() => setModalOpen(false)} onSubmit={handleSubmit} />
      )}
    </div>
  );
}

/* ============================== APP ============================== */

export default function BookkeepingApp({ clientId, clientSwitcher, onSignOut }) {
  const [data, setData] = useState(makeInitialData());
  const [loaded, setLoaded] = useState(false);
  const [page, setPage] = useState("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved

  // Tracks the last document we know Firestore holds, so the debounced save can
  // skip no-op writes (and the echo of our own snapshot) instead of looping.
  const lastSyncedRef = useRef(null);
  const dataRef = useRef(data);
  dataRef.current = data;

  // Load + live sync — replaces window.storage.get / localStorage.getItem.
  // Switching clients just re-points this listener at the new clientId.
  useEffect(() => {
    if (!clientId) { setLoaded(false); return; }
    setLoaded(false);
    lastSyncedRef.current = null;
    const ref = doc(db, "clients", clientId);
    const unsub = onSnapshot(
      ref,
      (snap) => {
        if (snap.metadata.hasPendingWrites) return; // our own not-yet-acked write
        const merged = snap.exists()
          ? { ...makeInitialData(), ...normalizeDatesInData(snap.data()) }
          : makeInitialData();
        const mergedJson = JSON.stringify(merged);
        lastSyncedRef.current = mergedJson;
        // Only replace local state when the remote copy actually differs — avoids
        // clobbering the user's in-progress edits with the echo of our own write.
        if (mergedJson !== JSON.stringify(dataRef.current)) setData(merged);
        setLoaded(true);
      },
      (err) => { console.error("Firestore sync failed:", err); setLoaded(true); }
    );
    return unsub;
  }, [clientId]);

  // Save — replaces window.storage.set / localStorage.setItem. Same 500ms debounce
  // Haki always used; guarded so an unchanged `data` never triggers a write.
  useEffect(() => {
    if (!loaded || !clientId) return;
    const json = JSON.stringify(data);
    if (json === lastSyncedRef.current) return;
    setSaveState("saving");
    const t = setTimeout(async () => {
      try {
        lastSyncedRef.current = json;
        await setDoc(doc(db, "clients", clientId), JSON.parse(json), { merge: true });
        setSaveState("saved");
      } catch (e) { console.error("Firestore save failed:", e); setSaveState("idle"); }
    }, 500);
    return () => clearTimeout(t);
  }, [data, loaded, clientId]);

  const loadSample = () => setData((d) => ({ ...d, ...makeSampleAddOns() }));
  const [confirmDialog, setConfirmDialog] = useState(null);
  const resetAll = () => {
    setConfirmDialog({
      title: "Reset workbook", danger: true, confirmLabel: "Reset workbook",
      message: "Clear all data and start fresh? This cannot be undone.",
      onConfirm: () => { setData(makeInitialData()); setConfirmDialog(null); },
    });
  };

  const backupInputRef = useRef(null);
  const [backupMsg, setBackupMsg] = useState(null);
  const handleExportBackup = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `haki-backup-${todayFileStamp()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  const handleImportBackupFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      // A restored backup should uphold the same "kept in order" invariant as live editing —
      // otherwise accounts/customers/suppliers/items added before that behavior existed (or added
      // via any path that bypassed it) stay stuck wherever they happened to land in the saved file.
      const byCode = (a, b) => String(a.code || "").localeCompare(String(b.code || ""), undefined, { numeric: true, sensitivity: "base" });
      const byName = (a, b) => String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" });
      const restored = { ...makeInitialData(), ...normalizeDatesInData(parsed) };
      if (Array.isArray(restored.coa)) restored.coa = [...restored.coa].sort(byCode);
      if (Array.isArray(restored.customers)) restored.customers = [...restored.customers].sort(byName);
      if (Array.isArray(restored.suppliers)) restored.suppliers = [...restored.suppliers].sort(byName);
      if (Array.isArray(restored.items)) restored.items = [...restored.items].sort(byName);
      setConfirmDialog({
        title: "Restore backup", danger: true, confirmLabel: "Restore backup",
        message: "Restore this backup? This replaces everything currently in Haki with the backup's contents.",
        onConfirm: () => { setData(restored); setConfirmDialog(null); setBackupMsg({ type: "success", text: "Backup restored." }); setTimeout(() => setBackupMsg(null), 4000); },
      });
    } catch (err) {
      setBackupMsg({ type: "error", text: "Couldn't read that file. Make sure it's a .json backup exported from Haki." });
      setTimeout(() => setBackupMsg(null), 5000);
    }
  };

  const postings = useMemo(() => buildPostings(data), [data]);
  const coaMap = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);

  if (!loaded) {
    return <div className="ph-books boot"><div className="boot-card">Opening the books…</div><Style /></div>;
  }

  const pageMeta = NAV.flatMap((g) => g.items).find((i) => i.key === page);

  return (
    <div className="ph-books">
      <Style />
      <div className={"sidebar" + (navOpen ? " open" : "")}>
        <div className="brand">
          <div className="brand-mark"><HakiLogo size={34} /></div>
          <div>
            <div className="brand-title">Haki</div>
            <div className="brand-sub">Books of Accounts</div>
          </div>
          <button className="nav-close" onClick={() => setNavOpen(false)}><X size={18} /></button>
        </div>
        {clientSwitcher && <div className="client-switch-wrap">{clientSwitcher}</div>}
        <nav>
          {NAV.map((g) => (
            <div key={g.group} className="nav-group">
              <div className="nav-group-label">{g.group}</div>
              {g.items.map((it) => (
                <button key={it.key} className={"nav-item" + (page === it.key ? " active" : "")}
                  onClick={() => { setPage(it.key); setNavOpen(false); }}>
                  <it.icon size={16} strokeWidth={2} /><span>{it.label}</span>
                </button>
              ))}
            </div>
          ))}
        </nav>
        <div className="sidebar-foot">
          <div className="save-indicator">{saveState === "saving" ? "Saving…" : "Saved to the cloud"}</div>
          <button className={"nav-item settings-tab" + (page === "settings" ? " active" : "")}
            onClick={() => { setPage("settings"); setNavOpen(false); }}>
            <Settings size={16} strokeWidth={2} /><span>Settings</span>
          </button>
          <input ref={backupInputRef} type="file" accept=".json" style={{ display: "none" }} onChange={handleImportBackupFile} />
        </div>
      </div>

      {navOpen && <div className="scrim" onClick={() => setNavOpen(false)} />}


      <div className="main">
        <div className="topbar">
          <button className="nav-open" onClick={() => setNavOpen(true)}><Menu size={20} /></button>
          <div className="topbar-title">{data.company.name || "Your Company Name Inc."}</div>
          <div className="topbar-tin">{data.company.tin ? `TIN: ${data.company.tin}` : ""}</div>
        </div>
        <div className="content">
          {page === "dashboard" && <Dashboard data={data} postings={postings} coaMap={coaMap} onLoadSample={loadSample} />}
          {page === "company" && <CompanyPage data={data} setData={setData} saveState={saveState} />}
          {page === "coa" && <CoaPage data={data} setData={setData} />}
          {page === "customers" && <CustomersPage data={data} setData={setData} />}
          {page === "suppliers" && <SuppliersPage data={data} setData={setData} />}
          {page === "items" && <ItemRecordPage data={data} setData={setData} />}
          {page === "atc" && <AtcReferencePage data={data} setData={setData} />}
          {page === "vattypes" && <VatTypesPage data={data} setData={setData} />}
          {page === "sales" && <SalesPage data={data} setData={setData} />}
          {page === "purchases" && <PurchasesPage data={data} setData={setData} />}
          {page === "disbursements" && <DisbursementsPage data={data} setData={setData} />}
          {page === "receipts" && <ReceiptsPage data={data} setData={setData} />}
          {page === "generaljournal" && <GeneralJournalPage data={data} setData={setData} />}
          {page === "fixedassets" && <FixedAssetLedgerPage data={data} setData={setData} />}
          {page === "generalledger" && <GeneralLedgerDetailPage data={data} postings={postings} coaMap={coaMap} />}
          {page === "vat" && <VatSummaryPage data={data} postings={postings} />}
          {page === "ledger" && <GeneralLedgerPage data={data} postings={postings} coaMap={coaMap} />}
          {page === "transactions" && <TransactionDetailsPage data={data} setData={setData} postings={postings} coaMap={coaMap} />}
          {page === "trial" && <TrialBalancePage data={data} postings={postings} coaMap={coaMap} />}
          {page === "income" && <IncomeStatementPage data={data} setData={setData} postings={postings} coaMap={coaMap} />}
          {page === "balance" && <BalanceSheetPage data={data} postings={postings} coaMap={coaMap} />}
          {page === "importation" && <ImportationLedgerPage data={data} setData={setData} />}
          {page === "slspi" && <SLSPIPage data={data} />}
          {page === "qap" && <QAPPage data={data} />}
          {page === "sawt" && <SAWTPage data={data} />}
          {page === "form2307" && <Form2307Page data={data} setData={setData} />}
          {page === "alphalist" && <AlphalistEmployeesPage data={data} setData={setData} />}
          {page === "settings" && (
            <SettingsPage
              backupMsg={backupMsg}
              onExportBackup={handleExportBackup}
              onImportBackup={() => backupInputRef.current?.click()}
              onResetWorkbook={resetAll}
              onSignOut={onSignOut}
            />
          )}
        </div>
      </div>
      {confirmDialog && (
        <ConfirmModal title={confirmDialog.title} message={confirmDialog.message} confirmLabel={confirmDialog.confirmLabel} danger={confirmDialog.danger}
          onCancel={() => setConfirmDialog(null)} onConfirm={confirmDialog.onConfirm} />
      )}
    </div>
  );
}

/* ============================== SETTINGS ============================== */

function SettingsPage({ backupMsg, onExportBackup, onImportBackup, onResetWorkbook, onSignOut }) {
  return (
    <div>
      <SectionHeader icon={Settings} title="Settings" subtitle="Back up or restore this client's books, start over, or sign out." />

      <div className="settings-section">
        <div className="settings-section-head">
          <h3>Full backup</h3>
          <p>Everything for this client — company, masters &amp; all journals — as a single JSON file.</p>
        </div>
        <div className="settings-actions">
          <button className="settings-btn" onClick={onExportBackup}><FileDown size={15} /> Export full backup</button>
          <button className="settings-btn" onClick={onImportBackup}><FileUp size={15} /> Import full backup</button>
        </div>
        {backupMsg && <div className={"backup-msg" + (backupMsg.type === "error" ? " error" : "")}>{backupMsg.text}</div>}
      </div>

      <div className="settings-section">
        <div className="settings-section-head">
          <h3>Danger zone</h3>
          <p>Reset clears all data for this client and starts fresh. This cannot be undone.</p>
        </div>
        <div className="settings-actions">
          <button className="settings-btn danger" onClick={onResetWorkbook}><RotateCcw size={15} /> Reset workbook</button>
        </div>
      </div>

      {onSignOut && (
        <div className="settings-section">
          <div className="settings-section-head">
            <h3>Account</h3>
            <p>Sign out of Haki on this device.</p>
          </div>
          <div className="settings-actions">
            <button className="settings-btn" onClick={onSignOut}><LogOut size={15} /> Sign out</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================== DASHBOARD ============================== */

function Dashboard({ data, postings, coaMap, onLoadSample }) {
  const [year, setYear] = useState(2026);
  const isEmpty = data.sales.length === 0 && data.purchases.length === 0 && data.disbursements.length === 0 && data.receipts.length === 0;

  const monthly = useMemo(() => {
    return MONTHS.map((m, idx) => {
      const start = new Date(year, idx, 1), end = monthEnd(year, idx);
      let sales = 0, outputVat = 0, inputVat = 0, expenses = 0;
      postings.forEach((p) => {
        if (!inRange(p.date, start, end)) return;
        if (p.code === "4000") sales += p.credit - p.debit;
        if (p.code === "2200") outputVat += p.credit - p.debit;
        if (p.code === "1400") inputVat += p.debit - p.credit;
        const acct = coaMap[p.code];
        if (acct && acct.type === "Expense") expenses += p.debit - p.credit;
      });
      return { month: m.slice(0,3), sales: round2(sales), outputVat: round2(outputVat), inputVat: round2(inputVat), expenses: round2(expenses) };
    });
  }, [postings, year, coaMap]);

  const expenseBreakdown = useMemo(() => {
    const byCat = {};
    const start = new Date(year, 0, 1), end = monthEnd(year, 11);
    postings.forEach((p) => {
      if (!inRange(p.date, start, end)) return;
      const acct = coaMap[p.code];
      if (!acct || acct.type !== "Expense") return;
      const net = p.debit - p.credit;
      byCat[acct.name] = (byCat[acct.name] || 0) + net;
    });
    return Object.entries(byCat).filter(([,v]) => v > 0.004).map(([name, value]) => ({ name, value: round2(value) })).sort((a,b) => b.value - a.value).slice(0, 8);
  }, [postings, year, coaMap]);

  const ytd = useMemo(() => {
    let sales = 0, expenses = 0, output = 0, input = 0;
    const start = new Date(year, 0, 1), end = monthEnd(year, 11);
    postings.forEach((p) => {
      if (!inRange(p.date, start, end)) return;
      if (p.code === "4000") sales += p.credit - p.debit;
      if (p.code === "2200") output += p.credit - p.debit;
      if (p.code === "1400") input += p.debit - p.credit;
      const acct = coaMap[p.code];
      if (acct && acct.type === "Expense") expenses += p.debit - p.credit;
    });
    return { sales: round2(sales), expenses: round2(expenses), vatPayable: round2(output - input) };
  }, [postings, year, coaMap]);

  const PIE_COLORS = ["#1D5FA8","#0F2A4D","#4A90D9","#2E86AB","#7EB2E3","#123A66","#9FC6EA","#5B8AC4"];

  return (
    <div>
      <SectionHeader icon={LayoutDashboard} title="Dashboard" subtitle="At a glance — sales, expenses, and VAT position for the year"
        right={
          <select className="year-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>
            {getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        } />

      {isEmpty && (
        <div className="callout">
          <Info size={16} />
          <div>
            <div className="callout-title">This workbook is empty.</div>
            <div className="callout-body">Start logging transactions in the Journals, or load a few sample entries to see how the reports come together.</div>
          </div>
          <button className="primary-btn" onClick={onLoadSample}>Load sample data</button>
        </div>
      )}

      <div className="kpi-row">
        <div className="kpi-card">
          <div className="kpi-label">Sales — {year}</div>
          <div className="kpi-value">{fmt(ytd.sales)}</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Expenses — {year}</div>
          <div className="kpi-value">{fmt(ytd.expenses)}</div>
        </div>
        <div className="kpi-card">
          <div className="kpi-label">Net VAT Payable — {year}</div>
          <div className={"kpi-value" + (ytd.vatPayable < 0 ? " neg" : "")}>{fmt(ytd.vatPayable)}</div>
        </div>
      </div>

      <div className="chart-grid">
        <div className="chart-card">
          <div className="chart-title">Monthly Sales Trend</div>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={monthly} margin={{ top: 8, right: 8, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#DCE8F5" />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#57697F" }} axisLine={{ stroke: "#C7D9EC" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#57697F" }} axisLine={false} tickLine={false} tickFormatter={(v) => `₱${(v/1000).toFixed(0)}k`} />
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #C7D9EC" }} />
              <Bar dataKey="sales" fill="#1D5FA8" radius={[3,3,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="chart-card">
          <div className="chart-title">Expense Breakdown</div>
          {expenseBreakdown.length === 0 ? <div className="chart-empty">No expenses recorded for {year} yet.</div> : (
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={expenseBreakdown} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} label={(e) => e.name}>
                {expenseBreakdown.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
              </Pie>
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #C7D9EC" }} />
            </PieChart>
          </ResponsiveContainer>
          )}
        </div>

        <div className="chart-card wide">
          <div className="chart-title">VAT Position — Output vs Input</div>
          <ResponsiveContainer width="100%" height={240}>
            <LineChart data={monthly} margin={{ top: 8, right: 8, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#DCE8F5" />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#57697F" }} axisLine={{ stroke: "#C7D9EC" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#57697F" }} axisLine={false} tickLine={false} tickFormatter={(v) => `₱${(v/1000).toFixed(0)}k`} />
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 6, border: "1px solid #C7D9EC" }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Line type="monotone" dataKey="outputVat" name="Output VAT" stroke="#0F2A4D" strokeWidth={2} dot={{ r: 3 }} />
              <Line type="monotone" dataKey="inputVat" name="Input VAT" stroke="#4A90D9" strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ============================== COMPANY DETAILS ============================== */

function CompanyPage({ data, setData, saveState }) {
  // Company Details is the one page that does NOT autosave on every keystroke. Edits update a
  // local draft; nothing is written to Firestore until "Save changes" is clicked, which commits
  // the draft into `data` (the app's normal debounced Firestore save then picks it up).
  const [draft, setDraft] = useState(data.company);
  const savedJson = JSON.stringify(data.company);
  // Follow the underlying company object if it changes out from under us — a client switch or a
  // live update from another device — rather than stranding a stale draft.
  useEffect(() => { setDraft(data.company); }, [savedJson]);
  const c = draft;
  const set = (k, v) => setDraft((p) => ({ ...p, [k]: v }));
  const isIndividual = c.taxpayerType === "Individual";
  const dirty = JSON.stringify(draft) !== savedJson;

  // Post-write confirmation: after committing, wait for the app's global save state to actually
  // reach "saved" before showing "Saved" — not just optimistically on click.
  const [savePhase, setSavePhase] = useState("idle"); // idle | pending | writing | done
  useEffect(() => {
    setSavePhase((p) => {
      if (p === "pending" && saveState === "saving") return "writing";
      if (p === "writing" && saveState === "saved") return "done";
      return p;
    });
  }, [saveState]);
  useEffect(() => {
    if (savePhase !== "done") return;
    const t = setTimeout(() => setSavePhase("idle"), 2500);
    return () => clearTimeout(t);
  }, [savePhase]);
  // Fallback so the button never stays stuck on "Saving…" if the global save state
  // transition is missed (e.g. an unchanged commit).
  useEffect(() => {
    if (savePhase !== "pending" && savePhase !== "writing") return;
    const t = setTimeout(() => setSavePhase((p) => (p === "pending" || p === "writing") ? "done" : p), 4000);
    return () => clearTimeout(t);
  }, [savePhase]);
  const handleSave = () => { setData((d) => ({ ...d, company: draft })); setSavePhase("pending"); };
  const saving = savePhase === "pending" || savePhase === "writing";

  return (
    <div>
      <SectionHeader icon={Building2} title="Company Details" subtitle="Single source of company info — every report and journal header pulls from here. Changes here are saved only when you click Save changes." />
      <div className="form-card">
        <div className="form-grid">
          <LabeledField label="Business / Trade Name"><Field value={c.name} onChange={(v) => set("name", v)} placeholder="Your Company Name Inc." /></LabeledField>
          <LabeledField label="Registered Name"><Field value={c.registeredName} onChange={(v) => set("registeredName", v)} /></LabeledField>
          <LabeledField label="Registered Address" wide><Field value={c.address} onChange={(v) => set("address", v)} /></LabeledField>
          <LabeledField label="Zip Code"><Field value={c.zipCode} onChange={(v) => set("zipCode", v)} placeholder="0000" /></LabeledField>
          <LabeledField label="TIN"><Field value={c.tin} onChange={(v) => set("tin", v)} placeholder="000-000-000-000" /></LabeledField>
          <LabeledField label="RDO Code"><Field value={c.rdo} onChange={(v) => set("rdo", v)} /></LabeledField>
          <LabeledField label="Line of Business"><Field value={c.lineOfBusiness} onChange={(v) => set("lineOfBusiness", v)} /></LabeledField>
          <LabeledField label="Permit to Use No. (ATP)"><Field value={c.atp} onChange={(v) => set("atp", v)} /></LabeledField>
          <LabeledField label="Date Registered / ATP Effective"><Field type="date" value={c.dateRegistered} onChange={(v) => set("dateRegistered", v)} /></LabeledField>
          <LabeledField label="Prepared By"><Field value={c.preparedBy} onChange={(v) => set("preparedBy", v)} /></LabeledField>
          <LabeledField label="Taxpayer Type">
            <Field type="select" value={c.taxpayerType} options={PARTY_TYPES} onChange={(v) => set("taxpayerType", v)} />
          </LabeledField>
          <LabeledField label="VAT Status">
            <Field type="select" value={c.vatStatus} options={["VAT Registered","Non-VAT"]} onChange={(v) => set("vatStatus", v)} />
          </LabeledField>
          {isIndividual && (<>
            <LabeledField label="Surname"><Field value={c.surname} onChange={(v) => set("surname", v)} /></LabeledField>
            <LabeledField label="First Name"><Field value={c.firstName} onChange={(v) => set("firstName", v)} /></LabeledField>
            <LabeledField label="Middle Name"><Field value={c.middleName} onChange={(v) => set("middleName", v)} /></LabeledField>
          </>)}
        </div>
        <p className="section-head-note" style={{ marginTop: 12, marginBottom: 0 }}>VAT Status determines how Cash Receipts computes CWT and how SAWT computes its tax base: if VAT Registered, amounts are treated as VAT-inclusive (divided by 1.12 before applying the withholding rate); if Non-VAT, the full amount is used as-is.</p>
      </div>

      <div className="sub-block-title" style={{ marginTop: 28 }}>Tax Compliance Filing Details</div>
      <p className="section-head-note">Used by the Tax Compliance tab (SLSPI, QAP, SAWT) when generating BIR .dat files — address is split into two lines since that's how the RELIEF/Alphalist format expects it. The Authorized Signatory block fills in the Payor's signature on generated BIR Form 2307 certificates.</p>
      <div className="form-card">
        <div className="form-grid">
          <LabeledField label="Branch Code"><Field value={c.branchCode} onChange={(v) => set("branchCode", v)} placeholder="0000" /></LabeledField>
          <LabeledField label="Trade Name (if different)"><Field value={c.tradeName} onChange={(v) => set("tradeName", v)} placeholder="Defaults to Business Name" /></LabeledField>
          <LabeledField label="Address Line 1" wide><Field value={c.addr1} onChange={(v) => set("addr1", v)} placeholder="Street / Barangay" /></LabeledField>
          <LabeledField label="Address Line 2" wide><Field value={c.addr2} onChange={(v) => set("addr2", v)} placeholder="City / Province" /></LabeledField>
          <LabeledField label="Authorized Signatory"><Field value={c.authorizedSignatory} onChange={(v) => set("authorizedSignatory", v)} placeholder="Name of the person signing 2307s" /></LabeledField>
          <LabeledField label="Position"><Field value={c.signatoryPosition} onChange={(v) => set("signatoryPosition", v)} placeholder="e.g. Finance Manager, President" /></LabeledField>
          <LabeledField label="Signatory's TIN"><Field value={c.signatoryTin} onChange={(v) => set("signatoryTin", v)} placeholder="000-000-000 (the signatory's own TIN)" /></LabeledField>
          <LabeledField label="E-Signature" wide>
            <SignatureUploadField value={c.signatureImage} onChange={(v) => set("signatureImage", v)} />
          </LabeledField>
        </div>
      </div>

      <div className="company-save-bar">
        {savePhase === "done" && <span className="save-confirm"><Check size={14} /> Saved</span>}
        {dirty && savePhase === "idle" && <span className="save-pending-note">Unsaved changes</span>}
        <button className="primary-btn" disabled={!dirty || saving} onClick={handleSave}>
          {saving ? "Saving…" : "Save changes"}
        </button>
      </div>
    </div>
  );
}

function SignatureUploadField({ value, onChange }) {
  const fileInputRef = useRef(null);
  const handleFileChange = (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onChange(reader.result); // already a data: URL
    reader.readAsDataURL(file);
  };
  return (
    <div className="signature-upload">
      {value ? (
        <>
          <img src={value} alt="Signature preview" style={{ maxHeight: 60, maxWidth: 200, border: "1px solid var(--line)", borderRadius: 6, background: "#fff", padding: 4 }} />
          <button type="button" className="io-btn" onClick={() => onChange("")}>Remove</button>
        </>
      ) : (
        <button type="button" className="io-btn" onClick={() => fileInputRef.current?.click()}><FileUp size={13} /> Upload signature image</button>
      )}
      <input ref={fileInputRef} type="file" accept="image/png,image/jpeg" style={{ display: "none" }} onChange={handleFileChange} />
    </div>
  );
}
export function LabeledField({ label, children, wide }) {
  return <div className={"form-field" + (wide ? " wide" : "")}><label>{label}</label>{children}</div>;
}

/* ============================== CHART OF ACCOUNTS ============================== */

function CoaPage({ data, setData }) {
  const cols = [
    { key: "code", label: "Code", width: 80 },
    { key: "name", label: "Account Name", width: 260 },
    { key: "type", label: "Class", type: "select", options: COA_TYPES, width: 130 },
    { key: "category", label: "Subclass", type: "combo", options: COA_SUBCLASSES, width: 170 },
    { key: "accountType", label: "Account Type", type: "combo", options: COA_ACCOUNT_TYPES, width: 170 },
    { key: "notes", label: "Notes", width: 220 },
  ];
  const onChange = (id, key, v) => setData((d) => ({ ...d, coa: d.coa.map((r) => r.id === id ? { ...r, [key]: v } : r) }));
  const sortByCode = (arr) => [...arr].sort((a, b) => String(a.code || "").localeCompare(String(b.code || ""), undefined, { numeric: true, sensitivity: "base" }));
  const onCreate = (row) => setData((d) => ({ ...d, coa: sortByCode([...d.coa, row]) }));
  const onBulkAdd = (rows) => setData((d) => ({ ...d, coa: sortByCode([...d.coa, ...rows]) }));
  const onDelete = (id) => setData((d) => ({ ...d, coa: d.coa.filter((r) => r.id !== id) }));
  return (
    <div>
      <SectionHeader icon={BookOpenText} title="Chart of Accounts" subtitle="Account codes and names referenced across every journal. Add rows as needed." />
      <MasterTable title="Chart of Accounts" columns={cols} rows={data.coa} onChange={onChange} onCreate={onCreate} onBulkAdd={onBulkAdd} onDelete={onDelete}
        addLabel="Add account" initialValues={{ type: "Expense" }} />
    </div>
  );
}

/* ============================== CUSTOMERS / SUPPLIERS ============================== */

function CustomersPage({ data, setData }) {
  const cols = [
    { key: "tin", label: "TIN", width: 130 },
    { key: "name", label: "Customers Name", width: 200 },
    { key: "registeredName", label: "Registered Name", width: 180 },
    { key: "firstName", label: "First Name", width: 120 },
    { key: "middleName", label: "Middle Name", width: 120 },
    { key: "surname", label: "Surname", width: 120 },
    { key: "address", label: "Address", width: 200 },
    { key: "type", label: "Type", type: "select", options: PARTY_TYPES, width: 130 },
    { key: "account", label: "Account", width: 110 },
    { key: "taxCode", label: "Tax Code", type: "percent", align: "right", width: 90 },
  ];
  const onChange = (id, key, v) => setData((d) => ({ ...d, customers: d.customers.map((r) => r.id === id ? { ...r, [key]: v } : r) }));
  const sortByName = (arr) => [...arr].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" }));
  const onCreate = (row) => setData((d) => ({ ...d, customers: sortByName([...d.customers, row]) }));
  const onBulkAdd = (rows) => setData((d) => ({ ...d, customers: sortByName([...d.customers, ...rows]) }));
  const onDelete = (id) => setData((d) => ({ ...d, customers: d.customers.filter((r) => r.id !== id) }));
  return (
    <div>
      <SectionHeader icon={Users} title="Customers Master" subtitle="Add every customer once. The Sales Journal looks up TIN and Address automatically by name." />
      <MasterTable title="Customers Master" columns={cols} rows={data.customers} onChange={onChange} onCreate={onCreate} onBulkAdd={onBulkAdd} onDelete={onDelete}
        addLabel="Add customer" initialValues={{ type: "Non-Individual", account: "SALES", taxCode: 0.12 }} />
    </div>
  );
}

function SuppliersPage({ data, setData }) {
  const itemOptions = data.items.map((i) => ({ value: i.item, label: i.item }));
  const acctOptions = data.coa.map((a) => ({ value: a.name, label: `${a.code} · ${a.name}` }));
  const cols = [
    { key: "tin", label: "TIN", width: 130 },
    { key: "name", label: "Suppliers Name", width: 200 },
    { key: "registeredName", label: "Registered Name", width: 180 },
    { key: "firstName", label: "First Name", width: 120 },
    { key: "middleName", label: "Middle Name", width: 120 },
    { key: "surname", label: "Surname", width: 120 },
    { key: "address", label: "Address", width: 200 },
    { key: "zipCode", label: "Zip Code", width: 90 },
    { key: "itemCode", label: "Item Code", type: "combo", options: itemOptions, width: 160 },
    { key: "accountTitle", label: "Account Title", type: "combo", options: acctOptions, width: 220 },
    { key: "type", label: "Type", type: "select", options: PARTY_TYPES, width: 130 },
    { key: "taxCode", label: "Tax Code", width: 90 },
    { key: "businessStyle", label: "Business Style", width: 140 },
  ];
  const onChange = (id, key, v) => setData((d) => ({ ...d, suppliers: d.suppliers.map((r) => r.id === id ? { ...r, [key]: v } : r) }));
  const sortByName = (arr) => [...arr].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" }));
  const onCreate = (row) => setData((d) => ({ ...d, suppliers: sortByName([...d.suppliers, row]) }));
  const onBulkAdd = (rows) => setData((d) => ({ ...d, suppliers: sortByName([...d.suppliers, ...rows]) }));
  const onDelete = (id) => setData((d) => ({ ...d, suppliers: d.suppliers.filter((r) => r.id !== id) }));
  return (
    <div>
      <SectionHeader icon={Truck} title="Suppliers Master" subtitle="Add every supplier once. The Purchase Journal looks up TIN, Address, and Description automatically by name." />
      <MasterTable title="Suppliers Master" columns={cols} rows={data.suppliers} onChange={onChange} onCreate={onCreate} onBulkAdd={onBulkAdd} onDelete={onDelete}
        addLabel="Add supplier" initialValues={{ type: "Non-Individual", taxCode: "VAT" }} />
    </div>
  );
}

/* ============================== ITEM RECORD + ATC REFERENCE ============================== */

function ItemRecordPage({ data, setData }) {
  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const itemCols = [
    { key: "item", label: "Item", width: 180 },
    { key: "account", label: "Default Account", type: "combo", options: acctOptions, width: 220 },
    { key: "type", label: "Type", type: "select", options: ITEM_TYPES, width: 180 },
  ];
  const sortItems = (arr) => [...arr].sort((a, b) => (a.item || "").localeCompare(b.item || "", undefined, { sensitivity: "base" }));
  const onItemChange = (id, key, v) => setData((d) => ({ ...d, items: d.items.map((r) => r.id === id ? { ...r, [key]: v } : r) }));
  const onItemCreate = (row) => setData((d) => ({ ...d, items: sortItems([...d.items, row]) }));
  const onItemBulkAdd = (rows) => setData((d) => ({ ...d, items: sortItems([...d.items, ...rows]) }));
  const onItemDelete = (id) => setData((d) => ({ ...d, items: d.items.filter((r) => r.id !== id) }));

  return (
    <div>
      <SectionHeader icon={ListChecks} title="Item Record" subtitle="Maps purchase items to a default account. Item Code fields in the Purchase Journal pull directly from here." />
      <MasterTable title="Item Record" columns={itemCols} rows={data.items} onChange={onItemChange} onCreate={onItemCreate} onBulkAdd={onItemBulkAdd} onDelete={onItemDelete}
        addLabel="Add item" initialValues={{ type: "Services" }} />
    </div>
  );
}

function AtcReferencePage({ data, setData }) {
  const atcCols = [
    { key: "code", label: "ATC Code", width: 100 },
    { key: "rate", label: "Rate", type: "percent", align: "center", width: 90 },
    { key: "desc", label: "Description", width: 420 },
  ];
  const onAtcChange = (id, key, v) => setData((d) => ({ ...d, atc: d.atc.map((r) => r.id === id ? { ...r, [key]: v } : r) }));
  const onAtcCreate = (row) => setData((d) => ({ ...d, atc: [...d.atc, row] }));
  const onAtcBulkAdd = (rows) => setData((d) => ({ ...d, atc: [...d.atc, ...rows] }));
  const onAtcDelete = (id) => setData((d) => ({ ...d, atc: d.atc.filter((r) => r.id !== id) }));

  return (
    <div>
      <SectionHeader icon={ListChecks} title="ATC Reference" subtitle="Alphanumeric Tax Code table — the BIR withholding-tax codes used across QAP and SAWT." />
      <MasterTable title="ATC Reference" columns={atcCols} rows={data.atc} onChange={onAtcChange} onCreate={onAtcCreate} onBulkAdd={onAtcBulkAdd} onDelete={onAtcDelete}
        addLabel="Add ATC code" />
    </div>
  );
}

/* ============================== VAT TYPE ============================== */

function VatTypesPage({ data, setData }) {
  const cols = [
    { key: "vatType", label: "VAT Type", width: 260 },
    { key: "books", label: "Books", type: "select", options: VAT_TYPE_BOOKS, width: 160 },
    { key: "slspiField", label: "SLSPI Field", type: "select", options: VAT_TYPE_SLSPI_FIELDS, width: 190 },
  ];
  const onChange = (id, key, v) => setData((d) => ({ ...d, vatTypes: d.vatTypes.map((r) => r.id === id ? { ...r, [key]: v } : r) }));
  const onCreate = (row) => setData((d) => ({ ...d, vatTypes: [...d.vatTypes, row] }));
  const onBulkAdd = (rows) => setData((d) => ({ ...d, vatTypes: [...d.vatTypes, ...rows] }));
  const onDelete = (id) => setData((d) => ({ ...d, vatTypes: d.vatTypes.filter((r) => r.id !== id) }));
  return (
    <div>
      <SectionHeader icon={Scale} title="VAT Type" subtitle="Edit the names or add rows as needed. Codes are referenced by the journals." />
      <MasterTable title="VAT Type" columns={cols} rows={data.vatTypes} onChange={onChange} onCreate={onCreate} onBulkAdd={onBulkAdd} onDelete={onDelete}
        addLabel="Add VAT type" initialValues={{ books: "Purchase journal" }} />
    </div>
  );
}

/* ============================== SALES JOURNAL ============================== */

function QuickAddPartyModal({ kind, initialName, onCancel, onSubmit }) {
  const isSupplier = kind === "supplier";
  const label = isSupplier ? "Supplier" : "Customer";
  const [v, setV] = useState({ name: initialName || "", tin: "", address: "", type: "Non-Individual" });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const handleSubmit = () => {
    const base = { id: uid(), name: v.name, registeredName: "", firstName: "", middleName: "", surname: "", address: v.address, tin: v.tin, type: v.type, taxCode: isSupplier ? "VAT" : 0.12 };
    onSubmit(isSupplier ? { ...base, itemCode: "", accountTitle: "", businessStyle: "" } : { ...base, account: "" });
  };
  return (
    <EntryModalShell title={`Add New ${label}`} submitLabel={`Add ${label.toLowerCase()}`} onCancel={onCancel} onSubmit={handleSubmit}>
      <LabeledField label={`${label} Name`} wide><Field value={v.name} onChange={(x) => set("name", x)} /></LabeledField>
      <LabeledField label="TIN"><Field value={v.tin} onChange={(x) => set("tin", x)} /></LabeledField>
      <LabeledField label="Type"><Field type="select" options={PARTY_TYPES} value={v.type} onChange={(x) => set("type", x)} /></LabeledField>
      <LabeledField label="Address" wide><Field value={v.address} onChange={(x) => set("address", x)} /></LabeledField>
      <div className="qf-hint">This creates a new record in the {label}s Master with just the essentials — add Surname/First/Middle Name, Business Style, and other details there later if needed.</div>
    </EntryModalShell>
  );
}

function ConfirmModal({ title, message, confirmLabel, danger, onCancel, onConfirm }) {
  const overlay = useOverlayClose(onCancel);
  return (
    <div className="modal-overlay" {...overlay}>
      <div className="modal-box confirm-box" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h2>{title || "Are you sure?"}</h2>
          <button className="modal-close" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="modal-body confirm-body"><p>{message}</p></div>
        <div className="modal-foot">
          <button className="btn-outline" onClick={onCancel}>Cancel</button>
          <button className={"primary-btn" + (danger ? " danger" : "")} onClick={onConfirm}>{confirmLabel || "Confirm"}</button>
        </div>
      </div>
    </div>
  );
}

export function EntryModalShell({ title, onCancel, onSubmit, submitLabel, children }) {
  const overlay = useOverlayClose(onCancel);
  return (
    <div className="modal-overlay" {...overlay}>
      <div className="modal-box" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h2>{title}</h2>
          <button className="modal-close" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="modal-body">{children}</div>
        <div className="modal-foot">
          <button className="btn-outline" onClick={onCancel}>Cancel</button>
          <button className="primary-btn" onClick={onSubmit}>{submitLabel || "Add"}</button>
        </div>
      </div>
    </div>
  );
}

function ComputedPreview({ items }) {
  return (
    <div className="computed-preview">
      {items.map(([label, value]) => (
        <div key={label} className="computed-preview-item"><span>{label}</span><strong>{value}</strong></div>
      ))}
    </div>
  );
}

function SalesEntryModal({ data, setData, onCancel, onSubmit }) {
  const [v, setV] = useState({ date: todayMDY(), siNo: "", customer: "", tin: "", address: "", desc: "", vatable: 0, exempt: 0, zeroRated: 0, terms: "Cash", coaCode: "4000", bankAccount: "Cash on Hand" });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const acctOptions = data.coa.filter((a) => a.type === "Revenue").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const custOptions = data.customers.map((c) => c.name).filter(Boolean);
  const [quickAdd, setQuickAdd] = useState(null); // typed name string, or null
  const onCustomer = (name) => {
    const c = data.customers.find((x) => x.name.toLowerCase() === (name || "").toLowerCase());
    setV((p) => ({ ...p, customer: name, ...(c ? { tin: c.tin, address: c.address } : {}) }));
  };
  const onQuickAddSubmit = (customer) => {
    setData((d) => ({ ...d, customers: [...d.customers, customer].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    setV((p) => ({ ...p, customer: customer.name, tin: customer.tin, address: customer.address }));
    setQuickAdd(null);
  };
  const computed = computeSalesRow(v);
  return (
    <EntryModalShell title="Add Sale" submitLabel="Add sale" onCancel={onCancel}
      onSubmit={() => onSubmit(computeSalesRow({ id: uid(), ...v }))}>
      <LabeledField label="Date"><Field type="date" value={v.date} onChange={(x) => set("date", x)} /></LabeledField>
      <LabeledField label="SI/OR No."><Field value={v.siNo} onChange={(x) => set("siNo", x)} placeholder="SI-0001" /></LabeledField>
      <LabeledField label="Customer Name"><Field type="combo" options={custOptions} value={v.customer} onChange={onCustomer}
        onAddNew={(typed) => setQuickAdd(typed)} addNewLabel="customer" /></LabeledField>
      <LabeledField label="TIN"><Field value={v.tin} onChange={(x) => set("tin", x)} /></LabeledField>
      <LabeledField label="Address" wide><Field value={v.address} onChange={(x) => set("address", x)} /></LabeledField>
      <LabeledField label="Description" wide><Field value={v.desc} onChange={(x) => set("desc", x)} /></LabeledField>
      <LabeledField label="VATable Sales"><Field type="number" align="right" value={v.vatable} onChange={(x) => set("vatable", x)} /></LabeledField>
      <LabeledField label="VAT-Exempt"><Field type="number" align="right" value={v.exempt} onChange={(x) => set("exempt", x)} /></LabeledField>
      <LabeledField label="Zero-Rated"><Field type="number" align="right" value={v.zeroRated} onChange={(x) => set("zeroRated", x)} /></LabeledField>
      <LabeledField label="Terms"><Field type="select" options={TERMS} value={v.terms} onChange={(x) => set("terms", x)} /></LabeledField>
      <LabeledField label="Account"><Field type="combo" options={acctOptions} value={v.coaCode} onChange={(x) => set("coaCode", x)} /></LabeledField>
      <LabeledField label="Bank Account"><Field type="select" options={BANK_ACCOUNTS} value={v.bankAccount} onChange={(x) => set("bankAccount", x)} disabled={v.terms === "Credit"} /></LabeledField>
      <ComputedPreview items={[["Output VAT", fmt(computed.outputVat)], ["Total", fmt(computed.total)]]} />
      {quickAdd !== null && (
        <QuickAddPartyModal kind="customer" initialName={quickAdd} onCancel={() => setQuickAdd(null)} onSubmit={onQuickAddSubmit} />
      )}
    </EntryModalShell>
  );
}

function PurchaseEntryModal({ data, setData, onCancel, onSubmit }) {
  const [v, setV] = useState({ date: todayMDY(), invNo: "", supplier: "", tin: "", address: "", itemCode: "", desc: "", vatType: "", atc: "", atcRate: 0, vatable: 0, nonvat: 0, terms: "Cash", coaCode: "5001", bankAccount: "Cash on Hand" });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const acctOptions = data.coa.filter((a) => ["Cost of Sales","Expense","Asset"].includes(a.type)).map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const itemOptions = data.items.map((i) => ({ value: i.item, label: i.item }));
  const suppOptions = data.suppliers.map((s) => s.name).filter(Boolean);
  const vatTypeOptions = data.vatTypes.filter((v) => v.books === "Purchase journal").map((vt) => ({ value: vt.vatType, label: vt.vatType }));
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const [quickAdd, setQuickAdd] = useState(null);
  const onSupplier = (name) => {
    const s = data.suppliers.find((x) => x.name.toLowerCase() === (name || "").toLowerCase());
    if (!s) return setV((p) => ({ ...p, supplier: name }));
    const item = data.items.find((i) => i.item === s.itemCode);
    setV((p) => ({ ...p, supplier: name, tin: s.tin, address: s.address, itemCode: s.itemCode || p.itemCode, desc: item ? item.item : p.desc, coaCode: item ? item.account : p.coaCode }));
  };
  const onQuickAddSubmit = (supplier) => {
    setData((d) => ({ ...d, suppliers: [...d.suppliers, supplier].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    setV((p) => ({ ...p, supplier: supplier.name, tin: supplier.tin, address: supplier.address }));
    setQuickAdd(null);
  };
  const onItemCode = (code) => {
    const item = data.items.find((i) => i.item === code);
    setV((p) => (item ? { ...p, itemCode: code, desc: item.item, coaCode: item.account } : { ...p, itemCode: code }));
  };
  const onAtc = (code) => {
    const atc = data.atc.find((a) => a.code === code);
    setV((p) => ({ ...p, atc: code, atcRate: atc ? num(atc.rate) : 0 }));
  };
  const computed = computePurchaseRow(v);
  return (
    <EntryModalShell title="Add Purchase" submitLabel="Add purchase" onCancel={onCancel}
      onSubmit={() => onSubmit(computePurchaseRow({ id: uid(), ...v }))}>
      <LabeledField label="Date"><Field type="date" value={v.date} onChange={(x) => set("date", x)} /></LabeledField>
      <LabeledField label="Supplier Inv./OR"><Field value={v.invNo} onChange={(x) => set("invNo", x)} /></LabeledField>
      <LabeledField label="Supplier Name"><Field type="combo" options={suppOptions} value={v.supplier} onChange={onSupplier}
        onAddNew={(typed) => setQuickAdd(typed)} addNewLabel="supplier" /></LabeledField>
      <LabeledField label="TIN"><Field value={v.tin} onChange={(x) => set("tin", x)} /></LabeledField>
      <LabeledField label="Address" wide><Field value={v.address} onChange={(x) => set("address", x)} /></LabeledField>
      <LabeledField label="Item Code"><Field type="combo" options={itemOptions} value={v.itemCode} onChange={onItemCode} /></LabeledField>
      <LabeledField label="Description"><Field value={v.desc} onChange={(x) => set("desc", x)} /></LabeledField>
      <LabeledField label="VAT Type" wide><Field type="combo" options={vatTypeOptions} value={v.vatType} onChange={(x) => set("vatType", x)} /></LabeledField>
      <LabeledField label="ATC"><Field type="combo" options={atcOptions} value={v.atc} onChange={onAtc} /></LabeledField>
      <LabeledField label="Rate"><ReadCell align="center">{v.atc ? `${round2(v.atcRate * 100)}%` : "—"}</ReadCell></LabeledField>
      <LabeledField label="VATable"><Field type="number" align="right" value={v.vatable} onChange={(x) => set("vatable", x)} /></LabeledField>
      <LabeledField label="Non-VAT"><Field type="number" align="right" value={v.nonvat} onChange={(x) => set("nonvat", x)} /></LabeledField>
      <LabeledField label="Terms"><Field type="select" options={TERMS} value={v.terms} onChange={(x) => set("terms", x)} /></LabeledField>
      <LabeledField label="Account"><Field type="combo" options={acctOptions} value={v.coaCode} onChange={(x) => set("coaCode", x)} /></LabeledField>
      <LabeledField label="Bank Account"><Field type="select" options={BANK_ACCOUNTS} value={v.bankAccount} onChange={(x) => set("bankAccount", x)} disabled={v.terms === "Credit"} /></LabeledField>
      <ComputedPreview items={[["Input VAT", fmt(computed.inputVat)], ["Total", fmt(computed.total)], ["EWT", fmt(computed.ewt)], ["Net Amount", fmt(computed.net)]]} />
      {quickAdd !== null && (
        <QuickAddPartyModal kind="supplier" initialName={quickAdd} onCancel={() => setQuickAdd(null)} onSubmit={onQuickAddSubmit} />
      )}
    </EntryModalShell>
  );
}

function DisbEntryModal({ data, setData, onCancel, onSubmit }) {
  const [v, setV] = useState({ date: todayMDY(), tin: "", vendor: "", cvNo: "", desc: "", atc: "", atcRate: 0, amount: 0, bankAccount: "Cash on Hand", coaCode: "6170" });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const acctOptions = data.coa.filter((a) => a.type === "Expense" || a.type === "Asset" || a.type === "Liability").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const suppOptions = data.suppliers.map((s) => s.name).filter(Boolean);
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const [quickAdd, setQuickAdd] = useState(null);
  const onSupplier = (name) => {
    const s = data.suppliers.find((x) => x.name.toLowerCase() === (name || "").toLowerCase());
    setV((p) => ({ ...p, vendor: name, ...(s ? { tin: s.tin } : {}) }));
  };
  const onQuickAddSubmit = (supplier) => {
    setData((d) => ({ ...d, suppliers: [...d.suppliers, supplier].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    setV((p) => ({ ...p, vendor: supplier.name, tin: supplier.tin }));
    setQuickAdd(null);
  };
  const onAtc = (code) => {
    const atc = data.atc.find((a) => a.code === code);
    setV((p) => ({ ...p, atc: code, atcRate: atc ? num(atc.rate) : 0 }));
  };
  const computed = computeDisbRow(v);
  return (
    <EntryModalShell title="Add Disbursement" submitLabel="Add disbursement" onCancel={onCancel}
      onSubmit={() => onSubmit(computeDisbRow({ id: uid(), ...v }))}>
      <LabeledField label="Date"><Field type="date" value={v.date} onChange={(x) => set("date", x)} /></LabeledField>
      <LabeledField label="TIN"><Field value={v.tin} onChange={(x) => set("tin", x)} /></LabeledField>
      <LabeledField label="Suppliers Name"><Field type="combo" options={suppOptions} value={v.vendor} onChange={onSupplier}
        onAddNew={(typed) => setQuickAdd(typed)} addNewLabel="supplier" /></LabeledField>
      <LabeledField label="Payment Ref."><Field value={v.cvNo} onChange={(x) => set("cvNo", x)} placeholder="CV-001" /></LabeledField>
      <LabeledField label="Description" wide><Field value={v.desc} onChange={(x) => set("desc", x)} /></LabeledField>
      <LabeledField label="ATC"><Field type="combo" options={atcOptions} value={v.atc} onChange={onAtc} /></LabeledField>
      <LabeledField label="Rate"><ReadCell align="center">{v.atc ? `${round2(v.atcRate * 100)}%` : "—"}</ReadCell></LabeledField>
      <LabeledField label="Amount"><Field type="number" align="right" value={v.amount} onChange={(x) => set("amount", x)} /></LabeledField>
      <LabeledField label="Bank Account"><Field type="select" options={BANK_ACCOUNTS} value={v.bankAccount} onChange={(x) => set("bankAccount", x)} /></LabeledField>
      <LabeledField label="Account"><Field type="combo" options={acctOptions} value={v.coaCode} onChange={(x) => set("coaCode", x)} /></LabeledField>
      <ComputedPreview items={[["EWT", fmt(computed.ewt)], ["Net Amount", fmt(computed.net)]]} />
      {quickAdd !== null && (
        <QuickAddPartyModal kind="supplier" initialName={quickAdd} onCancel={() => setQuickAdd(null)} onSubmit={onQuickAddSubmit} />
      )}
    </EntryModalShell>
  );
}

function ReceiptEntryModal({ data, setData, onCancel, onSubmit }) {
  const [v, setV] = useState({ date: todayMDY(), from: "", orNo: "", desc: "", atc: "", atcRate: 0, amount: 0, cwt: 0, bankAccount: "Cash on Hand", coaCode: "1200" });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const custOptions = data.customers.map((c) => c.name).filter(Boolean);
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const [quickAdd, setQuickAdd] = useState(null);
  const onQuickAddSubmit = (customer) => {
    setData((d) => ({ ...d, customers: [...d.customers, customer].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    setV((p) => ({ ...p, from: customer.name }));
    setQuickAdd(null);
  };
  const onAtc = (code) => {
    const atc = data.atc.find((a) => a.code === code);
    const rate = atc ? num(atc.rate) : 0;
    setV((p) => ({ ...p, atc: code, atcRate: rate, cwt: round2(cwtTaxBase(data, p.amount) * rate) }));
  };
  const onAmount = (val) => {
    setV((p) => ({ ...p, amount: val, cwt: p.atc ? round2(cwtTaxBase(data, val) * num(p.atcRate)) : p.cwt }));
  };
  const computed = computeReceiptRow(v);
  return (
    <EntryModalShell title="Add Receipt" submitLabel="Add receipt" onCancel={onCancel}
      onSubmit={() => onSubmit(computeReceiptRow({ id: uid(), ...v }))}>
      <LabeledField label="Date"><Field type="date" value={v.date} onChange={(x) => set("date", x)} /></LabeledField>
      <LabeledField label="Received From"><Field type="combo" options={custOptions} value={v.from} onChange={(x) => set("from", x)}
        onAddNew={(typed) => setQuickAdd(typed)} addNewLabel="customer" /></LabeledField>
      <LabeledField label="OR/Ref No."><Field value={v.orNo} onChange={(x) => set("orNo", x)} placeholder="OR-001" /></LabeledField>
      <LabeledField label="Description" wide><Field value={v.desc} onChange={(x) => set("desc", x)} /></LabeledField>
      <LabeledField label="ATC"><Field type="combo" options={atcOptions} value={v.atc} onChange={onAtc} /></LabeledField>
      <LabeledField label="Rate"><ReadCell align="center">{v.atc ? `${round2(v.atcRate * 100)}%` : "—"}</ReadCell></LabeledField>
      <LabeledField label="Amount"><Field type="number" align="right" value={v.amount} onChange={onAmount} /></LabeledField>
      <LabeledField label="CWT"><Field type="number" align="right" value={v.cwt} onChange={(x) => set("cwt", x)} /></LabeledField>
      <LabeledField label="Bank Account"><Field type="select" options={BANK_ACCOUNTS} value={v.bankAccount} onChange={(x) => set("bankAccount", x)} /></LabeledField>
      <LabeledField label="Account"><Field type="combo" options={acctOptions} value={v.coaCode} onChange={(x) => set("coaCode", x)} /></LabeledField>
      <ComputedPreview items={[["Net Amount", fmt(computed.net)]]} />
      <div className="qf-hint">CWT auto-fills from ATC × Amount, but stays fully editable — type over it (formulas work too) for non-VAT books or special cases.</div>
      {quickAdd !== null && (
        <QuickAddPartyModal kind="customer" initialName={quickAdd} onCancel={() => setQuickAdd(null)} onSubmit={onQuickAddSubmit} />
      )}
    </EntryModalShell>
  );
}

function JournalVoucherModal({ data, onCancel, onSubmit, nextJvNo, initial, title, submitLabel }) {
  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const [jvNo, setJvNo] = useState(initial?.jvNo || nextJvNo);
  const [date, setDate] = useState(initial?.date || todayMDY());
  const [particulars, setParticulars] = useState(initial?.particulars || "");
  const [lines, setLines] = useState(
    initial?.lines && initial.lines.length
      ? initial.lines.map((l) => ({ id: uid(), account: l.account, debit: l.debit, credit: l.credit }))
      : [{ id: uid(), account: "", debit: 0, credit: 0 }, { id: uid(), account: "", debit: 0, credit: 0 }]
  );
  const updateLine = (id, patch) => setLines((ls) => ls.map((l) => l.id === id ? { ...l, ...patch } : l));
  const addLine = () => setLines((ls) => [...ls, { id: uid(), account: "", debit: 0, credit: 0 }]);
  const deleteLine = (id) => setLines((ls) => ls.filter((l) => l.id !== id));
  const debitSum = round2(lines.reduce((a, l) => a + num(l.debit), 0));
  const creditSum = round2(lines.reduce((a, l) => a + num(l.credit), 0));
  const balanced = debitSum === creditSum && debitSum > 0;
  return (
    <EntryModalShell title={title || "Add Journal Voucher"} submitLabel={submitLabel || "Add voucher"} onCancel={onCancel}
      onSubmit={() => onSubmit({ id: uid(), jvNo, date, particulars, lines: lines.map((l) => ({ id: uid(), account: l.account, debit: num(l.debit), credit: num(l.credit) })) })}>
      <LabeledField label="JV No."><Field value={jvNo} onChange={setJvNo} placeholder="JV-001" /></LabeledField>
      <LabeledField label="Date"><Field type="date" value={date} onChange={setDate} /></LabeledField>
      <LabeledField label="Particulars" wide><Field value={particulars} onChange={setParticulars} placeholder="Particulars / description of entry" /></LabeledField>
      <div className="modal-lines">
        <table className="ledger-table">
          <thead><tr><th style={{minWidth:200}}>Account</th><th className="num-head" style={{minWidth:110}}>Debit</th><th className="num-head" style={{minWidth:110}}>Credit</th><th style={{width:32}}></th></tr></thead>
          <tbody>
            {lines.map((l) => (
              <tr key={l.id}>
                <td><Field type="combo" options={acctOptions} value={l.account} onChange={(v) => updateLine(l.id, { account: v })} /></td>
                <td><Field type="number" align="right" value={l.debit} onChange={(v) => updateLine(l.id, { debit: v, credit: v ? 0 : l.credit })} /></td>
                <td><Field type="number" align="right" value={l.credit} onChange={(v) => updateLine(l.id, { credit: v, debit: v ? 0 : l.debit })} /></td>
                <td className="text-center"><DelBtn onClick={() => deleteLine(l.id)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
        <AddRowBtn onClick={addLine}>Add line</AddRowBtn>
      </div>
      <ComputedPreview items={[["Debit total", fmt(debitSum)], ["Credit total", fmt(creditSum)], ["Status", debitSum === 0 && creditSum === 0 ? "Empty" : balanced ? "Balanced ✓" : "Out of balance ⚠"]]} />
    </EntryModalShell>
  );
}

function JournalPreviewModal({ voucher, coaByCode, onClose, onCopy }) {
  const debitSum = round2(voucher.lines.reduce((a, l) => a + num(l.debit), 0));
  const creditSum = round2(voucher.lines.reduce((a, l) => a + num(l.credit), 0));
  const overlay = useOverlayClose(onClose);
  return (
    <div className="modal-overlay" {...overlay}>
      <div className="modal-box" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h2>{voucher.jvNo || "Journal Entry"}</h2>
          <button className="modal-close" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <LabeledField label="Date"><ReadCell>{voucher.date || "—"}</ReadCell></LabeledField>
          <LabeledField label="JV No."><ReadCell>{voucher.jvNo || "—"}</ReadCell></LabeledField>
          <LabeledField label="Particulars" wide><ReadCell>{voucher.particulars || "—"}</ReadCell></LabeledField>
          <div className="modal-lines">
            <table className="ledger-table">
              <thead><tr><th style={{minWidth:220}}>Account</th><th className="num-head" style={{minWidth:110}}>Debit</th><th className="num-head" style={{minWidth:110}}>Credit</th></tr></thead>
              <tbody>
                {voucher.lines.map((l) => {
                  const acct = coaByCode[l.account];
                  return (
                    <tr key={l.id}>
                      <td>{l.account}{acct ? ` · ${acct.name}` : ""}</td>
                      <td className="num">{l.debit ? fmt(l.debit) : "—"}</td>
                      <td className="num">{l.credit ? fmt(l.credit) : "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot><tr><td className="totals-label">Total</td><td className="num">{fmt(debitSum)}</td><td className="num">{fmt(creditSum)}</td></tr></tfoot>
            </table>
          </div>
        </div>
        <div className="modal-foot">
          <button className="btn-outline" onClick={onClose}>Close</button>
          <button className="primary-btn" onClick={onCopy}>Make a copy</button>
        </div>
      </div>
    </div>
  );
}

/* ============================== PERIOD FILTER + ROW SELECTION ============================== */

function matchesPeriod(dateStr, filter) {
  if (filter.mode === "All") return true;
  const d = parseAppDate(dateStr);
  if (!d) return false;
  if (filter.mode === "Year") return d.getFullYear() === filter.year;
  if (filter.mode === "Month") return d.getFullYear() === filter.year && d.getMonth() === filter.monthIdx;
  if (filter.mode === "Quarter") {
    const q = QUARTERS[filter.quarterIdx];
    return d.getFullYear() === filter.year && q.months.includes(d.getMonth());
  }
  return true;
}

function usePeriodFilter() {
  const [filter, setFilter] = useState({ mode: "All", year: 2026, monthIdx: 0, quarterIdx: 0 });
  return [filter, setFilter];
}

function PeriodFilterBar({ filter, setFilter, years }) {
  return (
    <div className="period-filter-bar">
      <select className="io-select" value={filter.mode} onChange={(e) => setFilter((f) => ({ ...f, mode: e.target.value }))}>
        <option value="All">All dates</option>
        <option value="Month">Monthly</option>
        <option value="Quarter">Quarterly</option>
        <option value="Year">Yearly</option>
      </select>
      {filter.mode !== "All" && (
        <select className="io-select" value={filter.year} onChange={(e) => setFilter((f) => ({ ...f, year: Number(e.target.value) }))}>
          {(years || getAvailableYears(null)).map((y) => <option key={y} value={y}>{y}</option>)}
        </select>
      )}
      {filter.mode === "Month" && (
        <select className="io-select" value={filter.monthIdx} onChange={(e) => setFilter((f) => ({ ...f, monthIdx: Number(e.target.value) }))}>
          {MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}
        </select>
      )}
      {filter.mode === "Quarter" && (
        <select className="io-select" value={filter.quarterIdx} onChange={(e) => setFilter((f) => ({ ...f, quarterIdx: Number(e.target.value) }))}>
          {QUARTERS.map((q, i) => <option key={q.label} value={i}>{q.label}</option>)}
        </select>
      )}
    </div>
  );
}

function matchesSearch(row, fields, query) {
  if (!query || !query.trim()) return true;
  const q = query.trim().toLowerCase();
  return fields.some((f) => String(row[f] ?? "").toLowerCase().includes(q));
}

function SearchBar({ value, onChange, placeholder }) {
  return (
    <div className="search-bar">
      <Search size={14} />
      <input type="text" value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder || "Search…"} />
      {value && <button className="search-clear" onClick={() => onChange("")}><X size={13} /></button>}
    </div>
  );
}

function useRowSelection() {
  const [selected, setSelected] = useState(() => new Set());
  const toggle = (id) => setSelected((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  const toggleAll = (ids) => setSelected((s) => {
    const allSelected = ids.length > 0 && ids.every((id) => s.has(id));
    return allSelected ? new Set() : new Set(ids);
  });
  const clear = () => setSelected(new Set());
  return { selected, toggle, toggleAll, clear };
}

function SelectAllCheckbox({ ids, selected, toggleAll }) {
  const allSelected = ids.length > 0 && ids.every((id) => selected.has(id));
  return <input type="checkbox" className="row-check" checked={allSelected} onChange={() => toggleAll(ids)} disabled={ids.length === 0} title="Select all" />;
}

function RowCheckbox({ id, selected, toggle }) {
  return <input type="checkbox" className="row-check" checked={selected.has(id)} onChange={() => toggle(id)} />;
}

function SelectionBar({ count, onClear, children }) {
  return (
    <div className="selection-bar">
      <span>{count} selected</span>
      <div className="selection-actions">{children}</div>
      <button className="io-btn" onClick={onClear}>Clear</button>
    </div>
  );
}

function QuickFixModal({ data, count, onCancel, onApply }) {
  const acctOptions = data.coa.filter((a) => ["Cost of Sales","Expense","Asset"].includes(a.type)).map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const vatTypeOptions = data.vatTypes.filter((v) => v.books === "Purchase journal").map((vt) => vt.vatType);
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const [applyAccount, setApplyAccount] = useState(false);
  const [coaCode, setCoaCode] = useState("");
  const [applyTerms, setApplyTerms] = useState(false);
  const [terms, setTerms] = useState("Cash");
  const [applyBank, setApplyBank] = useState(false);
  const [bankAccount, setBankAccount] = useState("Cash on Hand");
  const [applySupplier, setApplySupplier] = useState(false);
  const [supplier, setSupplier] = useState("");
  const [applyVatType, setApplyVatType] = useState(false);
  const [vatType, setVatType] = useState("");
  const [applyAtc, setApplyAtc] = useState(false);
  const [atc, setAtc] = useState("");
  const [error, setError] = useState("");
  const suppNames = data.suppliers.map((s) => s.name).filter(Boolean);

  const handleApply = () => {
    const patch = {};
    if (applyAccount && coaCode) patch.coaCode = coaCode;
    if (applyTerms) patch.terms = terms;
    if (applyBank) patch.bankAccount = bankAccount;
    if (applyVatType) patch.vatType = vatType;
    if (applyAtc) {
      const a = data.atc.find((x) => x.code === atc);
      patch.atc = atc;
      patch.atcRate = a ? num(a.rate) : 0;
    }
    if (applySupplier && supplier) {
      const s = data.suppliers.find((x) => x.name.toLowerCase() === supplier.toLowerCase());
      patch.supplier = supplier;
      if (s) { patch.tin = s.tin; patch.address = s.address; }
    }
    if (Object.keys(patch).length === 0) { setError("Choose at least one field to update."); return; }
    onApply(patch);
  };

  return (
    <EntryModalShell title={`Quick Fix — ${count} transaction${count === 1 ? "" : "s"}`} submitLabel="Apply changes" onCancel={onCancel} onSubmit={handleApply}>
      <datalist id="qf-supp-list">{suppNames.map((n) => <option key={n} value={n} />)}</datalist>
      {error && <div className="qf-error">{error}</div>}
      <div className="qf-row">
        <label className="qf-check"><input type="checkbox" checked={applyAccount} onChange={(e) => setApplyAccount(e.target.checked)} /> Update Account</label>
        <Field type="combo" options={acctOptions} value={coaCode} onChange={setCoaCode} disabled={!applyAccount} />
      </div>
      <div className="qf-row">
        <label className="qf-check"><input type="checkbox" checked={applySupplier} onChange={(e) => setApplySupplier(e.target.checked)} /> Update Supplier</label>
        <Field value={supplier} list="qf-supp-list" onChange={setSupplier} disabled={!applySupplier} />
      </div>
      <div className="qf-row">
        <label className="qf-check"><input type="checkbox" checked={applyVatType} onChange={(e) => setApplyVatType(e.target.checked)} /> Update VAT Type</label>
        <Field type="combo" options={vatTypeOptions} value={vatType} onChange={setVatType} disabled={!applyVatType} />
      </div>
      <div className="qf-row">
        <label className="qf-check"><input type="checkbox" checked={applyAtc} onChange={(e) => setApplyAtc(e.target.checked)} /> Update ATC</label>
        <Field type="combo" options={atcOptions} value={atc} onChange={setAtc} disabled={!applyAtc} />
      </div>
      <div className="qf-row">
        <label className="qf-check"><input type="checkbox" checked={applyTerms} onChange={(e) => setApplyTerms(e.target.checked)} /> Update Terms</label>
        <Field type="select" options={TERMS} value={terms} onChange={setTerms} disabled={!applyTerms} />
      </div>
      <div className="qf-row">
        <label className="qf-check"><input type="checkbox" checked={applyBank} onChange={(e) => setApplyBank(e.target.checked)} /> Update Bank Account</label>
        <Field type="select" options={BANK_ACCOUNTS} value={bankAccount} onChange={setBankAccount} disabled={!applyBank} />
      </div>
      <div className="qf-hint">Only checked fields are applied to the {count} selected transaction{count === 1 ? "" : "s"} — everything else on those rows stays as-is.</div>
    </EntryModalShell>
  );
}

function SalesPage({ data, setData }) {
  const acctOptions = data.coa.filter((a) => a.type === "Revenue").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const custOptions = data.customers.map((c) => c.name).filter(Boolean);
  const importHook = useJournalImport("sales", data, ({ newRows }) => setData((d) => ({ ...d, sales: [...d.sales, ...newRows] })));
  const [modalOpen, setModalOpen] = useState(false);
  const [quickAddCustomer, setQuickAddCustomer] = useState(null); // { rowId } | null
  const [filter, setFilter] = usePeriodFilter();
  const [search, setSearch] = useState("");
  const sel = useRowSelection();

  const update = (id, patch) => setData((d) => ({
    ...d, sales: d.sales.map((r) => r.id === id ? computeSalesRow({ ...r, ...patch }) : r),
  }));
  const onCustomer = (id, name) => {
    const c = data.customers.find((x) => x.name.toLowerCase() === (name || "").toLowerCase());
    update(id, c ? { customer: name, tin: c.tin, address: c.address } : { customer: name });
  };
  const onQuickAddCustomer = (rowId, typedName) => setQuickAddCustomer({ rowId, typedName });
  const onQuickAddCustomerSubmit = (customer) => {
    setData((d) => ({ ...d, customers: [...d.customers, customer].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    onCustomer(quickAddCustomer.rowId, customer.name);
    setQuickAddCustomer(null);
  };
  const onAdd = () => setModalOpen(true);
  const onDelete = (id) => setData((d) => ({ ...d, sales: d.sales.filter((r) => r.id !== id) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, sales: d.sales.filter((r) => !sel.selected.has(r.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };

  const filteredRows = useMemo(() => data.sales.filter((r) => matchesPeriod(r.date, filter) && matchesSearch(r, ["siNo","customer","tin","address","desc"], search)), [data.sales, filter, search]);
  const totals = filteredRows.reduce((a, r) => ({
    vatable: a.vatable + num(r.vatable), exempt: a.exempt + num(r.exempt), zeroRated: a.zeroRated + num(r.zeroRated),
    outputVat: a.outputVat + num(r.outputVat), total: a.total + num(r.total),
  }), { vatable: 0, exempt: 0, zeroRated: 0, outputVat: 0, total: 0 });
  const emptyMsg = data.sales.length === 0 ? "No sales logged yet — add your first row above." : "No sales match this filter.";
  const exportHook = useJournalExport("sales", data, filteredRows, coaByCode);

  return (
    <div>
      <SectionHeader icon={Receipt} title="Sales Journal" subtitle="Log every sale / Official Receipt or Sales Invoice issued. Output VAT computes automatically."
        right={<div className="header-actions"><ImportExportBar journalKey="sales" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} /></div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
            <SearchBar value={search} onChange={setSearch} placeholder="Search SI No., customer, TIN, description…" />
          </div>
          <AddRowBtn onClick={onAdd}>Add sale</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead>
              <tr>
                <th style={{width:32}}><SelectAllCheckbox ids={filteredRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
                <th style={{minWidth:120}}>Date</th><th style={{minWidth:100}}>SI/OR No.</th><th style={{minWidth:180}}>Customer Name</th>
                <th style={{minWidth:120}}>TIN</th><th style={{minWidth:170}}>Address</th><th style={{minWidth:180}}>Description</th>
                <th style={{minWidth:110}} className="num-head">VATable Sales</th><th style={{minWidth:110}} className="num-head">VAT-Exempt</th>
                <th style={{minWidth:110}} className="num-head">Zero-Rated</th><th style={{minWidth:110}} className="num-head">Output VAT</th>
                <th style={{minWidth:110}} className="num-head">Total</th><th style={{minWidth:90}}>Terms</th>
                <th style={{minWidth:170}}>Account</th><th style={{minWidth:130}}>Bank Account</th><th style={{width:36}}></th>
              </tr>
            </thead>
            <tbody>
              {filteredRows.length === 0 && <tr><td colSpan={16} className="empty-row">{emptyMsg}</td></tr>}
              {filteredRows.map((r) => (
                <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><Field type="date" value={r.date} onChange={(v) => update(r.id, { date: v })} /></td>
                  <td><Field value={r.siNo} onChange={(v) => update(r.id, { siNo: v })} placeholder="SI-0001" /></td>
                  <td><Field type="combo" options={custOptions} value={r.customer} onChange={(v) => onCustomer(r.id, v)}
                    onAddNew={(typed) => onQuickAddCustomer(r.id, typed)} addNewLabel="customer" /></td>
                  <td><Field value={r.tin} onChange={(v) => update(r.id, { tin: v })} /></td>
                  <td><Field value={r.address} onChange={(v) => update(r.id, { address: v })} /></td>
                  <td><Field value={r.desc} onChange={(v) => update(r.id, { desc: v })} /></td>
                  <td><Field type="number" align="right" value={r.vatable} onChange={(v) => update(r.id, { vatable: v })} /></td>
                  <td><Field type="number" align="right" value={r.exempt} onChange={(v) => update(r.id, { exempt: v })} /></td>
                  <td><Field type="number" align="right" value={r.zeroRated} onChange={(v) => update(r.id, { zeroRated: v })} /></td>
                  <td><ReadCell align="right">{fmtPlain(r.outputVat)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.total)}</ReadCell></td>
                  <td><Field type="select" options={TERMS} value={r.terms} onChange={(v) => update(r.id, { terms: v })} /></td>
                  <td><Field type="combo" options={acctOptions} value={r.coaCode} onChange={(v) => update(r.id, { coaCode: v })} /></td>
                  <td><Field type="select" options={BANK_ACCOUNTS} value={r.bankAccount} onChange={(v) => update(r.id, { bankAccount: v })} disabled={r.terms === "Credit"} /></td>
                  <td className="text-center"><DelBtn onClick={() => onDelete(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {filteredRows.length > 0 && (
              <tfoot><tr>
                <td colSpan={7} className="totals-label">Totals</td>
                <td className="num">{fmtPlain(totals.vatable)}</td><td className="num">{fmtPlain(totals.exempt)}</td>
                <td className="num">{fmtPlain(totals.zeroRated)}</td><td className="num">{fmtPlain(totals.outputVat)}</td>
                <td className="num">{fmt(totals.total)}</td><td colSpan={4}></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
      {modalOpen && (
        <SalesEntryModal data={data} setData={setData} onCancel={() => setModalOpen(false)}
          onSubmit={(row) => { setData((d) => ({ ...d, sales: [...d.sales, row] })); setModalOpen(false); }} />
      )}
      {quickAddCustomer && (
        <QuickAddPartyModal kind="customer" initialName={quickAddCustomer.typedName}
          onCancel={() => setQuickAddCustomer(null)} onSubmit={onQuickAddCustomerSubmit} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected sale${sel.selected.size === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
    </div>
  );
}

/* ============================== PURCHASE JOURNAL ============================== */

function PurchasesPage({ data, setData }) {
  const acctOptions = data.coa.filter((a) => ["Cost of Sales","Expense","Asset"].includes(a.type)).map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const suppNames = data.suppliers.map((s) => s.name).filter(Boolean);
  const itemOptions = data.items.map((i) => ({ value: i.item, label: i.item }));
  const vatTypeOptions = data.vatTypes.filter((vt) => vt.books === "Purchase journal").map((vt) => ({ value: vt.vatType, label: vt.vatType }));
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const importHook = useJournalImport("purchases", data, ({ newRows }) => setData((d) => ({ ...d, purchases: [...d.purchases, ...newRows] })));
  const [modalOpen, setModalOpen] = useState(false);
  const [quickFixOpen, setQuickFixOpen] = useState(false);
  const [quickAddSupplier, setQuickAddSupplier] = useState(null); // { rowId, typedName } | null
  const [filter, setFilter] = usePeriodFilter();
  const [search, setSearch] = useState("");
  const sel = useRowSelection();
  const update = (id, patch) => setData((d) => ({ ...d, purchases: d.purchases.map((r) => r.id === id ? computePurchaseRow({ ...r, ...patch }) : r) }));
  const onSupplier = (id, name) => {
    const s = data.suppliers.find((x) => x.name.toLowerCase() === (name || "").toLowerCase());
    if (!s) return update(id, { supplier: name });
    const item = data.items.find((i) => i.item === s.itemCode);
    update(id, { supplier: name, tin: s.tin, address: s.address, itemCode: s.itemCode || "", desc: item ? item.item : "", coaCode: item ? item.account : "" });
  };
  const onQuickAddSupplierSubmit = (supplier) => {
    setData((d) => ({ ...d, suppliers: [...d.suppliers, supplier].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    onSupplier(quickAddSupplier.rowId, supplier.name);
    setQuickAddSupplier(null);
  };
  const onItemCode = (id, code) => {
    const item = data.items.find((i) => i.item === code);
    update(id, item ? { itemCode: code, desc: item.item, coaCode: item.account } : { itemCode: code });
  };
  const onAtc = (id, code) => {
    const atc = data.atc.find((a) => a.code === code);
    update(id, { atc: code, atcRate: atc ? num(atc.rate) : 0 });
  };
  const onAdd = () => setModalOpen(true);
  const onDelete = (id) => setData((d) => ({ ...d, purchases: d.purchases.filter((r) => r.id !== id) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, purchases: d.purchases.filter((r) => !sel.selected.has(r.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };
  const onQuickFixApply = (patch) => {
    setData((d) => ({ ...d, purchases: d.purchases.map((r) => sel.selected.has(r.id) ? computePurchaseRow({ ...r, ...patch }) : r) }));
    setQuickFixOpen(false);
    sel.clear();
  };

  const filteredRows = useMemo(() => data.purchases.filter((r) => matchesPeriod(r.date, filter) && matchesSearch(r, ["invNo","supplier","tin","address","desc","itemCode","vatType"], search)), [data.purchases, filter, search]);
  const totals = filteredRows.reduce((a, r) => ({
    vatable: a.vatable + num(r.vatable), nonvat: a.nonvat + num(r.nonvat), inputVat: a.inputVat + num(r.inputVat),
    total: a.total + num(r.total), ewt: a.ewt + num(r.ewt), net: a.net + num(r.net),
  }), { vatable: 0, nonvat: 0, inputVat: 0, total: 0, ewt: 0, net: 0 });
  const emptyMsg = data.purchases.length === 0 ? "No purchases logged yet — add your first row above." : "No purchases match this filter.";
  const exportHook = useJournalExport("purchases", data, filteredRows, coaByCode);
  const gen2307 = useGenerate2307(data, "purchases");
  const generate2307Selected = () => gen2307.generate(data.purchases.filter((r) => sel.selected.has(r.id)));

  return (
    <div>
      <SectionHeader icon={Wallet} title="Purchase Journal" subtitle="Log every supplier purchase. Item Code, Account, and VAT Type feed the SLP generator. Input VAT computes automatically."
        right={<div className="header-actions"><ImportExportBar journalKey="purchases" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} /></div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      <ExportStatus status={gen2307.status} />
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn accent" onClick={() => setQuickFixOpen(true)}>Quick Fix</button>
          <button className="io-btn" onClick={generate2307Selected} disabled={gen2307.status?.type === "pending"}><FileDown size={13} /> Generate 2307</button>
          <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
            <SearchBar value={search} onChange={setSearch} placeholder="Search invoice no., supplier, TIN, description…" />
          </div>
          <AddRowBtn onClick={onAdd}>Add purchase</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead>
              <tr>
                <th style={{width:32}}><SelectAllCheckbox ids={filteredRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
                <th style={{minWidth:120}}>Date</th><th style={{minWidth:110}}>Supplier Inv./OR</th><th style={{minWidth:190}}>Supplier Name</th>
                <th style={{minWidth:120}}>TIN</th><th style={{minWidth:160}}>Address</th>
                <th style={{minWidth:150}}>Item Code</th><th style={{minWidth:160}}>Description</th><th style={{minWidth:200}}>VAT Type</th>
                <th style={{minWidth:180}}>ATC</th><th style={{minWidth:80}} className="num-head">Rate</th>
                <th style={{minWidth:110}} className="num-head">VATable</th><th style={{minWidth:110}} className="num-head">Non-VAT</th>
                <th style={{minWidth:100}} className="num-head">Input VAT</th><th style={{minWidth:110}} className="num-head">Total</th>
                <th style={{minWidth:100}} className="num-head">EWT</th><th style={{minWidth:110}} className="num-head">Net Amount</th>
                <th style={{minWidth:90}}>Terms</th><th style={{minWidth:190}}>Account</th><th style={{minWidth:130}}>Bank Account</th><th style={{width:36}}></th>
              </tr>
            </thead>
            <tbody>
              {filteredRows.length === 0 && <tr><td colSpan={21} className="empty-row">{emptyMsg}</td></tr>}
              {filteredRows.map((r) => (
                <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><Field type="date" value={r.date} onChange={(v) => update(r.id, { date: v })} /></td>
                  <td><Field value={r.invNo} onChange={(v) => update(r.id, { invNo: v })} /></td>
                  <td><Field type="combo" options={suppNames} value={r.supplier} onChange={(v) => onSupplier(r.id, v)}
                    onAddNew={(typed) => setQuickAddSupplier({ rowId: r.id, typedName: typed })} addNewLabel="supplier" /></td>
                  <td><Field value={r.tin} onChange={(v) => update(r.id, { tin: v })} /></td>
                  <td><Field value={r.address} onChange={(v) => update(r.id, { address: v })} /></td>
                  <td><Field type="combo" options={itemOptions} value={r.itemCode} onChange={(v) => onItemCode(r.id, v)} /></td>
                  <td><Field value={r.desc} onChange={(v) => update(r.id, { desc: v })} /></td>
                  <td><Field type="combo" options={vatTypeOptions} value={r.vatType} onChange={(v) => update(r.id, { vatType: v })} /></td>
                  <td><Field type="combo" options={atcOptions} value={r.atc} onChange={(v) => onAtc(r.id, v)} /></td>
                  <td><ReadCell align="center">{r.atc ? `${round2(num(r.atcRate) * 100)}%` : "—"}</ReadCell></td>
                  <td><Field type="number" align="right" value={r.vatable} onChange={(v) => update(r.id, { vatable: v })} /></td>
                  <td><Field type="number" align="right" value={r.nonvat} onChange={(v) => update(r.id, { nonvat: v })} /></td>
                  <td><ReadCell align="right">{fmt(r.inputVat)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.total)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.ewt)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.net)}</ReadCell></td>
                  <td><Field type="select" options={TERMS} value={r.terms} onChange={(v) => update(r.id, { terms: v })} /></td>
                  <td><Field type="combo" options={acctOptions} value={r.coaCode} onChange={(v) => update(r.id, { coaCode: v })} /></td>
                  <td><Field type="select" options={BANK_ACCOUNTS} value={r.bankAccount} onChange={(v) => update(r.id, { bankAccount: v })} disabled={r.terms === "Credit"} /></td>
                  <td className="text-center"><DelBtn onClick={() => onDelete(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {filteredRows.length > 0 && (
              <tfoot><tr>
                <td colSpan={11} className="totals-label">Totals</td>
                <td className="num">{fmt(totals.vatable)}</td><td className="num">{fmt(totals.nonvat)}</td>
                <td className="num">{fmt(totals.inputVat)}</td><td className="num">{fmt(totals.total)}</td>
                <td className="num">{fmt(totals.ewt)}</td><td className="num">{fmt(totals.net)}</td><td colSpan={4}></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
      {modalOpen && (
        <PurchaseEntryModal data={data} setData={setData} onCancel={() => setModalOpen(false)}
          onSubmit={(row) => { setData((d) => ({ ...d, purchases: [...d.purchases, row] })); setModalOpen(false); }} />
      )}
      {quickAddSupplier && (
        <QuickAddPartyModal kind="supplier" initialName={quickAddSupplier.typedName}
          onCancel={() => setQuickAddSupplier(null)} onSubmit={onQuickAddSupplierSubmit} />
      )}
      {quickFixOpen && (
        <QuickFixModal data={data} count={sel.selected.size} onCancel={() => setQuickFixOpen(false)} onApply={onQuickFixApply} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected purchase${sel.selected.size === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
    </div>
  );
}

/* ============================== CASH DISBURSEMENTS ============================== */

function DisbursementsPage({ data, setData }) {
  const acctOptions = data.coa.filter((a) => a.type === "Expense" || a.type === "Asset" || a.type === "Liability").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const suppOptions = data.suppliers.map((s) => s.name).filter(Boolean);
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const importHook = useJournalImport("disbursements", data, ({ newRows }) => setData((d) => ({ ...d, disbursements: [...d.disbursements, ...newRows] })));
  const [modalOpen, setModalOpen] = useState(false);
  const [quickAddSupplier, setQuickAddSupplier] = useState(null); // { rowId, typedName } | null
  const [filter, setFilter] = usePeriodFilter();
  const [search, setSearch] = useState("");
  const sel = useRowSelection();
  const update = (id, patch) => setData((d) => ({ ...d, disbursements: d.disbursements.map((r) => r.id === id ? computeDisbRow({ ...r, ...patch }) : r) }));
  const onSupplier = (id, name) => {
    const s = data.suppliers.find((x) => x.name.toLowerCase() === (name || "").toLowerCase());
    update(id, s ? { vendor: name, tin: s.tin } : { vendor: name });
  };
  const onQuickAddSupplierSubmit = (supplier) => {
    setData((d) => ({ ...d, suppliers: [...d.suppliers, supplier].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    onSupplier(quickAddSupplier.rowId, supplier.name);
    setQuickAddSupplier(null);
  };
  const onAtc = (id, code) => {
    const atc = data.atc.find((a) => a.code === code);
    update(id, { atc: code, atcRate: atc ? num(atc.rate) : 0 });
  };
  const onAdd = () => setModalOpen(true);
  const onDelete = (id) => setData((d) => ({ ...d, disbursements: d.disbursements.filter((r) => r.id !== id) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, disbursements: d.disbursements.filter((r) => !sel.selected.has(r.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };
  const filteredRows = useMemo(() => data.disbursements.filter((r) => matchesPeriod(r.date, filter) && matchesSearch(r, ["cvNo","vendor","desc","tin"], search)), [data.disbursements, filter, search]);
  const totals = filteredRows.reduce((a, r) => ({ amount: a.amount + num(r.amount), ewt: a.ewt + num(r.ewt), net: a.net + num(r.net) }), { amount: 0, ewt: 0, net: 0 });
  const emptyMsg = data.disbursements.length === 0 ? "No disbursements logged yet — add your first row above." : "No disbursements match this filter.";
  const exportHook = useJournalExport("disbursements", data, filteredRows, coaByCode);
  const gen2307 = useGenerate2307(data, "disbursements");
  const generate2307Selected = () => gen2307.generate(data.disbursements.filter((r) => sel.selected.has(r.id)));

  return (
    <div>
      <SectionHeader icon={ArrowUpFromLine} title="Cash Disbursements" subtitle="Log all cash/bank/check payments — operating expenses, not just purchases. ATC feeds the QAP generator whenever EWT applies."
        right={<div className="header-actions"><ImportExportBar journalKey="disbursements" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} /></div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      <ExportStatus status={gen2307.status} />
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn" onClick={generate2307Selected} disabled={gen2307.status?.type === "pending"}><FileDown size={13} /> Generate 2307</button>
          <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
            <SearchBar value={search} onChange={setSearch} placeholder="Search ref. no., supplier, TIN, description…" />
          </div>
          <AddRowBtn onClick={onAdd}>Add disbursement</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr>
              <th style={{width:32}}><SelectAllCheckbox ids={filteredRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
              <th style={{minWidth:120}}>Date</th><th style={{minWidth:120}}>TIN</th><th style={{minWidth:190}}>Suppliers Name</th><th style={{minWidth:110}}>Payment Ref.</th>
              <th style={{minWidth:200}}>Description</th><th style={{minWidth:180}}>ATC</th><th style={{minWidth:80}} className="num-head">Rate</th>
              <th style={{minWidth:110}} className="num-head">Amount</th>
              <th style={{minWidth:100}} className="num-head">EWT</th><th style={{minWidth:110}} className="num-head">Net Amount</th>
              <th style={{minWidth:130}}>Bank Account</th><th style={{minWidth:190}}>Account</th><th style={{width:36}}></th>
            </tr></thead>
            <tbody>
              {filteredRows.length === 0 && <tr><td colSpan={14} className="empty-row">{emptyMsg}</td></tr>}
              {filteredRows.map((r) => (
                <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><Field type="date" value={r.date} onChange={(v) => update(r.id, { date: v })} /></td>
                  <td><Field value={r.tin} onChange={(v) => update(r.id, { tin: v })} /></td>
                  <td><Field type="combo" options={suppOptions} value={r.vendor} onChange={(v) => onSupplier(r.id, v)}
                    onAddNew={(typed) => setQuickAddSupplier({ rowId: r.id, typedName: typed })} addNewLabel="supplier" /></td>
                  <td><Field value={r.cvNo} onChange={(v) => update(r.id, { cvNo: v })} placeholder="CV-001" /></td>
                  <td><Field value={r.desc} onChange={(v) => update(r.id, { desc: v })} /></td>
                  <td><Field type="combo" options={atcOptions} value={r.atc} onChange={(v) => onAtc(r.id, v)} /></td>
                  <td><ReadCell align="center">{r.atc ? `${round2(num(r.atcRate) * 100)}%` : "—"}</ReadCell></td>
                  <td><Field type="number" align="right" value={r.amount} onChange={(v) => update(r.id, { amount: v })} /></td>
                  <td><ReadCell align="right">{fmt(r.ewt)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.net)}</ReadCell></td>
                  <td><Field type="select" options={BANK_ACCOUNTS} value={r.bankAccount} onChange={(v) => update(r.id, { bankAccount: v })} /></td>
                  <td><Field type="combo" options={acctOptions} value={r.coaCode} onChange={(v) => update(r.id, { coaCode: v })} /></td>
                  <td className="text-center"><DelBtn onClick={() => onDelete(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {filteredRows.length > 0 && (
              <tfoot><tr><td colSpan={8} className="totals-label">Totals</td>
                <td className="num">{fmt(totals.amount)}</td><td className="num">{fmt(totals.ewt)}</td><td className="num">{fmt(totals.net)}</td><td colSpan={3}></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
      {modalOpen && (
        <DisbEntryModal data={data} setData={setData} onCancel={() => setModalOpen(false)}
          onSubmit={(row) => { setData((d) => ({ ...d, disbursements: [...d.disbursements, row] })); setModalOpen(false); }} />
      )}
      {quickAddSupplier && (
        <QuickAddPartyModal kind="supplier" initialName={quickAddSupplier.typedName}
          onCancel={() => setQuickAddSupplier(null)} onSubmit={onQuickAddSupplierSubmit} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected disbursement${sel.selected.size === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
    </div>
  );
}

/* ============================== CASH RECEIPTS ============================== */

function ReceiptsPage({ data, setData }) {
  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const custOptions = data.customers.map((c) => c.name).filter(Boolean);
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const importHook = useJournalImport("receipts", data, ({ newRows }) => setData((d) => ({ ...d, receipts: [...d.receipts, ...newRows] })));
  const [modalOpen, setModalOpen] = useState(false);
  const [quickAddCustomer, setQuickAddCustomer] = useState(null); // { rowId, typedName } | null
  const [filter, setFilter] = usePeriodFilter();
  const [search, setSearch] = useState("");
  const sel = useRowSelection();
  const update = (id, patch) => setData((d) => ({ ...d, receipts: d.receipts.map((r) => r.id === id ? computeReceiptRow({ ...r, ...patch }) : r) }));
  const onAmount = (id, val) => {
    const row = data.receipts.find((r) => r.id === id);
    const patch = { amount: val };
    if (row && row.atc) patch.cwt = round2(cwtTaxBase(data, val) * num(row.atcRate));
    update(id, patch);
  };
  const onAtc = (id, code) => {
    const row = data.receipts.find((r) => r.id === id);
    const atc = data.atc.find((a) => a.code === code);
    const rate = atc ? num(atc.rate) : 0;
    update(id, { atc: code, atcRate: rate, cwt: round2(cwtTaxBase(data, row ? row.amount : 0) * rate) });
  };
  const onQuickAddCustomerSubmit = (customer) => {
    setData((d) => ({ ...d, customers: [...d.customers, customer].sort((a, b) => (a.name || "").localeCompare(b.name || "", undefined, { sensitivity: "base" })) }));
    update(quickAddCustomer.rowId, { from: customer.name });
    setQuickAddCustomer(null);
  };
  const onAdd = () => setModalOpen(true);
  const onDelete = (id) => setData((d) => ({ ...d, receipts: d.receipts.filter((r) => r.id !== id) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, receipts: d.receipts.filter((r) => !sel.selected.has(r.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };
  const filteredRows = useMemo(() => data.receipts.filter((r) => matchesPeriod(r.date, filter) && matchesSearch(r, ["orNo","from","desc"], search)), [data.receipts, filter, search]);
  const totals = filteredRows.reduce((a, r) => ({ amount: a.amount + num(r.amount), cwt: a.cwt + num(r.cwt), net: a.net + num(r.net) }), { amount: 0, cwt: 0, net: 0 });
  const emptyMsg = data.receipts.length === 0 ? "No receipts logged yet — add your first row above." : "No receipts match this filter.";
  const exportHook = useJournalExport("receipts", data, filteredRows, coaByCode);

  return (
    <div>
      <SectionHeader icon={ArrowDownToLine} title="Cash Receipts" subtitle="Log cash/bank collections that are NOT a sale. CWT auto-fills from ATC × Amount but stays editable — feeds QAP whenever CWT applies."
        right={<div className="header-actions"><ImportExportBar journalKey="receipts" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} /></div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
            <SearchBar value={search} onChange={setSearch} placeholder="Search OR no., received from, description…" />
          </div>
          <AddRowBtn onClick={onAdd}>Add receipt</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr>
              <th style={{width:32}}><SelectAllCheckbox ids={filteredRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
              <th style={{minWidth:120}}>Date</th><th style={{minWidth:170}}>Received From</th><th style={{minWidth:110}}>OR/Ref No.</th>
              <th style={{minWidth:200}}>Description</th><th style={{minWidth:180}}>ATC</th><th style={{minWidth:80}} className="num-head">Rate</th>
              <th style={{minWidth:110}} className="num-head">Amount</th>
              <th style={{minWidth:100}} className="num-head">CWT</th><th style={{minWidth:110}} className="num-head">Net Amount</th>
              <th style={{minWidth:130}}>Bank Account</th><th style={{minWidth:190}}>Account</th><th style={{width:36}}></th>
            </tr></thead>
            <tbody>
              {filteredRows.length === 0 && <tr><td colSpan={13} className="empty-row">{emptyMsg}</td></tr>}
              {filteredRows.map((r) => (
                <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><Field type="date" value={r.date} onChange={(v) => update(r.id, { date: v })} /></td>
                  <td><Field type="combo" options={custOptions} value={r.from} onChange={(v) => update(r.id, { from: v })}
                    onAddNew={(typed) => setQuickAddCustomer({ rowId: r.id, typedName: typed })} addNewLabel="customer" /></td>
                  <td><Field value={r.orNo} onChange={(v) => update(r.id, { orNo: v })} placeholder="OR-001" /></td>
                  <td><Field value={r.desc} onChange={(v) => update(r.id, { desc: v })} /></td>
                  <td><Field type="combo" options={atcOptions} value={r.atc} onChange={(v) => onAtc(r.id, v)} /></td>
                  <td><ReadCell align="center">{r.atc ? `${round2(num(r.atcRate) * 100)}%` : "—"}</ReadCell></td>
                  <td><Field type="number" align="right" value={r.amount} onChange={(v) => onAmount(r.id, v)} /></td>
                  <td><Field type="number" align="right" value={r.cwt} onChange={(v) => update(r.id, { cwt: v })} /></td>
                  <td><ReadCell align="right">{fmt(r.net)}</ReadCell></td>
                  <td><Field type="select" options={BANK_ACCOUNTS} value={r.bankAccount} onChange={(v) => update(r.id, { bankAccount: v })} /></td>
                  <td><Field type="combo" options={acctOptions} value={r.coaCode} onChange={(v) => update(r.id, { coaCode: v })} /></td>
                  <td className="text-center"><DelBtn onClick={() => onDelete(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {filteredRows.length > 0 && (
              <tfoot><tr><td colSpan={7} className="totals-label">Totals</td>
                <td className="num">{fmt(totals.amount)}</td><td className="num">{fmt(totals.cwt)}</td><td className="num">{fmt(totals.net)}</td><td colSpan={3}></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
      {modalOpen && (
        <ReceiptEntryModal data={data} setData={setData} onCancel={() => setModalOpen(false)}
          onSubmit={(row) => { setData((d) => ({ ...d, receipts: [...d.receipts, row] })); setModalOpen(false); }} />
      )}
      {quickAddCustomer && (
        <QuickAddPartyModal kind="customer" initialName={quickAddCustomer.typedName}
          onCancel={() => setQuickAddCustomer(null)} onSubmit={onQuickAddCustomerSubmit} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected receipt${sel.selected.size === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
    </div>
  );
}

/* ============================== GENERAL JOURNAL ============================== */

function GeneralJournalPage({ data, setData }) {
  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const importHook = useJournalImport("generaljournal", data, ({ newVouchers }) => setData((d) => ({ ...d, generalJournal: [...d.generalJournal, ...newVouchers] })));
  const [modalOpen, setModalOpen] = useState(false);
  const [filter, setFilter] = usePeriodFilter();
  const [search, setSearch] = useState("");
  const [previewVoucher, setPreviewVoucher] = useState(null);
  const [copySource, setCopySource] = useState(null);
  const sel = useRowSelection();

  const updateJv = (id, patch) => setData((d) => ({ ...d, generalJournal: d.generalJournal.map((jv) => jv.id === id ? { ...jv, ...patch } : jv) }));
  const updateLine = (jvId, lineId, patch) => setData((d) => ({
    ...d, generalJournal: d.generalJournal.map((jv) => jv.id === jvId ? { ...jv, lines: jv.lines.map((l) => l.id === lineId ? { ...l, ...patch } : l) } : jv),
  }));
  const addLine = (jvId) => setData((d) => ({ ...d, generalJournal: d.generalJournal.map((jv) => jv.id === jvId ? { ...jv, lines: [...jv.lines, { id: uid(), account: "", debit: 0, credit: 0 }] } : jv) }));
  const deleteLine = (jvId, lineId) => setData((d) => ({ ...d, generalJournal: d.generalJournal.map((jv) => jv.id === jvId ? { ...jv, lines: jv.lines.filter((l) => l.id !== lineId) } : jv) }));
  const addJv = () => setModalOpen(true);
  const deleteJv = (jvId) => setData((d) => ({ ...d, generalJournal: d.generalJournal.filter((jv) => jv.id !== jvId) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, generalJournal: d.generalJournal.filter((jv) => !sel.selected.has(jv.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };

  const filteredVouchers = useMemo(() => data.generalJournal.filter((jv) => matchesPeriod(jv.date, filter)), [data.generalJournal, filter]);
  const emptyMsg = data.generalJournal.length === 0 ? "No journal vouchers yet — add one above." : "No vouchers match this filter.";
  const exportHook = useJournalExport("generaljournal", data, filteredVouchers, coaByCode);

  const searchResults = useMemo(() => {
    if (!search.trim()) return [];
    const q = search.trim().toLowerCase();
    return data.generalJournal.filter((jv) => {
      if ((jv.jvNo || "").toLowerCase().includes(q)) return true;
      if ((jv.particulars || "").toLowerCase().includes(q)) return true;
      if ((jv.date || "").toLowerCase().includes(q)) return true;
      return jv.lines.some((l) => {
        const acct = coaByCode[l.account];
        return `${l.account} ${acct ? acct.name : ""}`.toLowerCase().includes(q);
      });
    });
  }, [data.generalJournal, search, coaByCode]);

  const onMakeCopy = () => {
    setCopySource(previewVoucher);
    setPreviewVoucher(null);
  };
  const onPostCopy = (jv) => {
    setData((d) => ({ ...d, generalJournal: [...d.generalJournal, jv] }));
    setCopySource(null);
    setSearch("");
  };

  return (
    <div>
      <SectionHeader icon={ScrollText} title="General Journal" subtitle="Adjusting/correcting entries that don't belong in the other journals — depreciation, accruals, opening balances."
        right={<div className="header-actions"><ImportExportBar journalKey="generaljournal" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} /><AddRowBtn onClick={addJv}>Add journal voucher</AddRowBtn></div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      <div className="gj-search-bar-wrap"><SearchBar value={search} onChange={setSearch} placeholder="Search JV No., date, particulars, or account…" /></div>

      {search.trim() ? (
        <div className="gj-search-results">
          {searchResults.length === 0 && <div className="empty-panel">No journal entries match "{search}".</div>}
          {searchResults.map((jv) => {
            const debitSum = round2(jv.lines.reduce((a, l) => a + num(l.debit), 0));
            return (
              <div key={jv.id} className="gj-search-row" onClick={() => setPreviewVoucher(jv)}>
                <span className="gj-search-jvno">{jv.jvNo || "—"}</span>
                <span className="gj-search-date">{jv.date}</span>
                <span className="gj-search-particulars">{jv.particulars || "—"}</span>
                <span className="gj-search-amount">{fmt(debitSum)}</span>
              </div>
            );
          })}
        </div>
      ) : (
        <>
          <div className="table-toolbar gj-filter-bar">
            <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
            {filteredVouchers.length > 0 && (
              <label className="qf-check gj-select-all">
                <SelectAllCheckbox ids={filteredVouchers.map((jv) => jv.id)} selected={sel.selected} toggleAll={sel.toggleAll} /> Select all shown
              </label>
            )}
          </div>
          {sel.selected.size > 0 && (
            <SelectionBar count={sel.selected.size} onClear={sel.clear}>
              <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
            </SelectionBar>
          )}
          {filteredVouchers.length === 0 && <div className="empty-panel">{emptyMsg}</div>}
          {filteredVouchers.map((jv) => {
            const debitSum = round2(jv.lines.reduce((a, l) => a + num(l.debit), 0));
            const creditSum = round2(jv.lines.reduce((a, l) => a + num(l.credit), 0));
            const balanced = debitSum === creditSum && debitSum > 0;
            return (
              <div key={jv.id} className={"jv-card" + (sel.selected.has(jv.id) ? " row-selected" : "")}>
                <div className="jv-head">
                  <RowCheckbox id={jv.id} selected={sel.selected} toggle={sel.toggle} />
                  <Field value={jv.jvNo} onChange={(v) => updateJv(jv.id, { jvNo: v })} placeholder="JV-001" />
                  <Field type="date" value={jv.date} onChange={(v) => updateJv(jv.id, { date: v })} />
                  <Field value={jv.particulars} onChange={(v) => updateJv(jv.id, { particulars: v })} placeholder="Particulars / description of entry" />
                  <span className={"jv-status" + (balanced ? " ok" : "")}>{debitSum === 0 && creditSum === 0 ? "Empty" : balanced ? "Balanced" : "Out of balance"}</span>
                  <button className="del-btn" onClick={() => deleteJv(jv.id)}><Trash2 size={14} /></button>
                </div>
                <div className="table-toolbar">
                  <AddRowBtn onClick={() => addLine(jv.id)}>Add line</AddRowBtn>
                </div>
                <table className="ledger-table">
                  <thead><tr><th style={{minWidth:220}}>Account</th><th className="num-head" style={{minWidth:120}}>Debit</th><th className="num-head" style={{minWidth:120}}>Credit</th><th style={{width:36}}></th></tr></thead>
                  <tbody>
                    {jv.lines.map((l) => (
                      <tr key={l.id}>
                        <td><Field type="combo" options={acctOptions} value={l.account} onChange={(v) => updateLine(jv.id, l.id, { account: v })} /></td>
                        <td><Field type="number" align="right" value={l.debit} onChange={(v) => updateLine(jv.id, l.id, { debit: v, credit: v ? 0 : l.credit })} /></td>
                        <td><Field type="number" align="right" value={l.credit} onChange={(v) => updateLine(jv.id, l.id, { credit: v, debit: v ? 0 : l.debit })} /></td>
                        <td className="text-center"><DelBtn onClick={() => deleteLine(jv.id, l.id)} /></td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot><tr><td className="totals-label">Line totals</td><td className="num">{fmt(debitSum)}</td><td className="num">{fmt(creditSum)}</td><td></td></tr></tfoot>
                </table>
              </div>
            );
          })}
        </>
      )}

      {modalOpen && (
        <JournalVoucherModal data={data} nextJvNo={`JV-${String(data.generalJournal.length + 1).padStart(3, "0")}`}
          onCancel={() => setModalOpen(false)}
          onSubmit={(jv) => { setData((d) => ({ ...d, generalJournal: [...d.generalJournal, jv] })); setModalOpen(false); }} />
      )}
      {previewVoucher && (
        <JournalPreviewModal voucher={previewVoucher} coaByCode={coaByCode} onClose={() => setPreviewVoucher(null)} onCopy={onMakeCopy} />
      )}
      {copySource && (
        <JournalVoucherModal data={data} nextJvNo={`JV-${String(data.generalJournal.length + 1).padStart(3, "0")}`}
          initial={{ date: todayMDY(), particulars: copySource.particulars, lines: copySource.lines }}
          title={`Copy of ${copySource.jvNo || "Journal Entry"}`} submitLabel="Post"
          onCancel={() => setCopySource(null)} onSubmit={onPostCopy} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected voucher${sel.selected.size === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
    </div>
  );
}

/* ============================== GENERAL LEDGER (DETAILED, PER-ACCOUNT) ============================== */

const TRANSACTION_TYPE_LABELS = {
  "Sales Journal": "Sales",
  "Purchase Journal": "Purchase",
  "Cash Disbursements": "Disbursement",
  "Cash Receipts": "Receipt",
  "General Journal": "Journal Entry",
};

function GeneralLedgerDetailPage({ data, postings, coaMap }) {
  const [year, setYear] = useState(2026);
  const start = new Date(year, 0, 1), end = monthEnd(year, 11);
  const priorCutoff = new Date(year - 1, 11, 31);

  const groups = useMemo(() => {
    const openingByCode = {};
    postings.forEach((p) => {
      if (upTo(p.date, priorCutoff)) openingByCode[p.code] = round2((openingByCode[p.code] || 0) + p.debit - p.credit);
    });

    const byCode = {};
    postings.forEach((p) => {
      if (!inRange(p.date, start, end)) return;
      if (!byCode[p.code]) byCode[p.code] = [];
      byCode[p.code].push(p);
    });

    const codes = new Set([
      ...Object.keys(openingByCode).filter((c) => openingByCode[c] !== 0),
      ...Object.keys(byCode),
    ]);

    return Array.from(codes).sort().map((code) => {
      const acct = coaMap[code];
      const opening = openingByCode[code] || 0;
      const txns = (byCode[code] || []).slice().sort((a, b) => {
        const da = parseAppDate(a.date), db = parseAppDate(b.date);
        return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
      });
      let running = opening;
      const rows = txns.map((p) => {
        running = round2(running + p.debit - p.credit);
        const acctRow = coaMap[p.code];
        return { ...p, acctName: acctRow ? acctRow.name : p.code, runningBalance: running };
      });
      return { code, name: acct ? acct.name : code, opening, rows, closing: running };
    });
  }, [postings, coaMap, year, priorCutoff, start, end]);

  const totals = groups.reduce((a, g) => g.rows.reduce((b, r) => ({ debit: b.debit + r.debit, credit: b.credit + r.credit }), a), { debit: 0, credit: 0 });
  const exportHook = useGeneralLedgerExport(data, groups, year);

  return (
    <div>
      <SectionHeader icon={BookOpenText} title="General Ledger" subtitle="Every transaction posted to every account for the selected year, with opening/closing balances and a running balance per account."
        right={<div className="header-actions">
          <ExportBar exportHook={exportHook} />
          <select className="year-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>{getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}</select>
        </div>} />
      <ExportStatus status={exportHook.status} />
      <div className="form-card gl-info-card">
        <div className="form-grid">
          <LabeledField label="Business Name"><ReadCell>{data.company.name || "Your Company Name Inc."}</ReadCell></LabeledField>
          <LabeledField label="Business Address"><ReadCell>{data.company.address || "—"}</ReadCell></LabeledField>
          <LabeledField label="Prepared By"><ReadCell>{data.company.preparedBy || "—"}</ReadCell></LabeledField>
          <LabeledField label="TIN"><ReadCell>{data.company.tin || "—"}</ReadCell></LabeledField>
          <LabeledField label="Period" wide><ReadCell>{formatDMY(start)} – {formatDMY(end)}</ReadCell></LabeledField>
        </div>
      </div>

      <div className="ledger-wrap">
        <div className="table-scroll">
          <table className="ledger-table report-table">
            <thead>
              <tr>
                <th style={{minWidth:110}}>Date</th><th style={{minWidth:120}}>Transaction Type</th><th style={{minWidth:200}}>Description</th>
                <th style={{minWidth:120}}>Reference</th><th style={{minWidth:220}}>Account Name</th><th style={{minWidth:90}}>Account Code</th>
                <th style={{minWidth:110}} className="num-head">Debit (PHP)</th><th style={{minWidth:110}} className="num-head">Credit (PHP)</th>
                <th style={{minWidth:130}} className="num-head">Running Balance (PHP)</th>
              </tr>
            </thead>
            <tbody>
              {groups.length === 0 && <tr><td colSpan={9} className="empty-row">No transactions or balances for {year} yet.</td></tr>}
              {groups.map((g) => (
                <React.Fragment key={g.code}>
                  <tr className="gl-group-header"><td colSpan={9}>{g.code} {g.name}</td></tr>
                  <tr className="gl-open-close">
                    <td colSpan={8}>{g.code} {g.name} Opening Balances</td>
                    <td className="num">{fmt(g.opening)}</td>
                  </tr>
                  {g.rows.map((r) => (
                    <tr key={r.id}>
                      <td>{r.date}</td>
                      <td>{TRANSACTION_TYPE_LABELS[r.source] || r.source}</td>
                      <td>{r.desc || "—"}</td>
                      <td>{r.ref || "—"}</td>
                      <td>{r.acctName}</td>
                      <td>{r.code}</td>
                      <td className="num">{r.debit ? fmt(r.debit) : "—"}</td>
                      <td className="num">{r.credit ? fmt(r.credit) : "—"}</td>
                      <td className="num">{fmt(r.runningBalance)}</td>
                    </tr>
                  ))}
                  <tr className="gl-open-close closing">
                    <td colSpan={8}>{g.code} {g.name} Closing Balances</td>
                    <td className="num">{fmt(g.closing)}</td>
                  </tr>
                </React.Fragment>
              ))}
            </tbody>
            {groups.length > 0 && (
              <tfoot><tr>
                <td colSpan={6} className="totals-label">Total movement — {year}</td>
                <td className="num">{fmt(totals.debit)}</td><td className="num">{fmt(totals.credit)}</td><td></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
    </div>
  );
}

/* ============================== VAT SUMMARY ============================== */

function VatSummaryPage({ data, postings }) {
  const [year, setYear] = useState(2026);
  const rows = MONTHS.map((m, idx) => {
    const start = new Date(year, idx, 1), end = monthEnd(year, idx);
    let output = 0, input = 0;
    postings.forEach((p) => {
      if (!inRange(p.date, start, end)) return;
      if (p.code === "2200") output += p.credit - p.debit;
      if (p.code === "1400") input += p.debit - p.credit;
    });
    return { month: m, output: round2(output), input: round2(input), net: round2(output - input) };
  });
  const totals = rows.reduce((a, r) => ({ output: a.output + r.output, input: a.input + r.input, net: a.net + r.net }), { output: 0, input: 0, net: 0 });

  return (
    <div>
      <SectionHeader icon={Scale} title="VAT Summary" subtitle="Monthly Output VAT vs Input VAT — mirrors BIR Form 2550M/2550Q logic. Verify against actual invoices before filing."
        right={<select className="year-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>{getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}</select>} />
      <div className="ledger-wrap">
        <div className="table-scroll">
          <table className="ledger-table report-table">
            <thead><tr><th>Month</th><th className="num-head">Output VAT</th><th className="num-head">Input VAT</th><th className="num-head">Net VAT Payable</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.month}><td>{r.month}</td><td className="num">{fmt(r.output)}</td><td className="num">{fmt(r.input)}</td>
                  <td className={"num" + (r.net < 0 ? " neg" : "")}>{fmt(r.net)}</td></tr>
              ))}
            </tbody>
            <tfoot><tr><td className="totals-label">Total — {year}</td><td className="num">{fmt(totals.output)}</td><td className="num">{fmt(totals.input)}</td>
              <td className={"num" + (totals.net < 0 ? " neg" : "")}>{fmt(totals.net)}</td></tr></tfoot>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ============================== GENERAL LEDGER ============================== */

function GeneralLedgerPage({ data, postings, coaMap }) {
  const [year, setYear] = useState(2026);
  const accountsWithActivity = useMemo(() => {
    const start = new Date(year, 0, 1), end = monthEnd(year, 11);
    const codes = new Set();
    postings.forEach((p) => { if (inRange(p.date, start, end)) codes.add(p.code); });
    return data.coa.filter((a) => codes.has(a.code));
  }, [postings, year, data.coa]);

  const rows = accountsWithActivity.map((acct) => {
    const monthsNet = MONTHS.map((_, idx) => {
      const start = new Date(year, idx, 1), end = monthEnd(year, idx);
      let debit = 0, credit = 0;
      postings.forEach((p) => { if (p.code === acct.code && inRange(p.date, start, end)) { debit += p.debit; credit += p.credit; } });
      return round2(debit - credit);
    });
    const yearTotal = round2(monthsNet.reduce((a, b) => a + b, 0));
    return { acct, monthsNet, yearTotal };
  });

  return (
    <div>
      <SectionHeader icon={ClipboardList} title="Monthly Ledger" subtitle="Monthly net movement per account, rolled up automatically from all journals."
        right={<select className="year-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>{getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}</select>} />
      {rows.length === 0 ? <div className="empty-panel">No posted activity for {year} yet.</div> : (
      <div className="ledger-wrap">
        <div className="table-scroll">
          <table className="ledger-table report-table ml-compact">
            <thead><tr><th style={{minWidth:190, position:"sticky", left:0}}>Account</th>
              {MONTHS.map((m) => <th key={m} className="num-head" style={{minWidth:76}}>{m.slice(0,3)}</th>)}
              <th className="num-head" style={{minWidth:100}}>Year Total</th></tr></thead>
            <tbody>
              {rows.map(({ acct, monthsNet, yearTotal }) => (
                <tr key={acct.code}>
                  <td style={{position:"sticky", left:0, background:"var(--paper)"}}>{acct.code} · {acct.name}</td>
                  {monthsNet.map((v, i) => <td key={i} className="num">{v === 0 ? "—" : fmtParen(v)}</td>)}
                  <td className="num" style={{ fontWeight: 600 }}>{fmtParen(yearTotal)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      )}
    </div>
  );
}

/* ============================== TRANSACTION DETAILS ============================== */

function ReclassModal({ data, count, currentCode, onCancel, onApply }) {
  const acctOptions = data.coa.filter((a) => a.code !== currentCode).map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const [newCode, setNewCode] = useState("");
  const [error, setError] = useState("");
  return (
    <EntryModalShell title={`Reclass — ${count} transaction${count === 1 ? "" : "s"}`} submitLabel="Apply reclass" onCancel={onCancel}
      onSubmit={() => { if (!newCode) { setError("Choose the account to reclass to."); return; } onApply(newCode); }}>
      {error && <div className="qf-error">{error}</div>}
      <LabeledField label="Reclass to Account" wide>
        <Field type="combo" options={acctOptions} value={newCode} onChange={setNewCode} />
      </LabeledField>
      <div className="qf-hint">This moves the {count} selected transaction{count === 1 ? "" : "s"} from their current account to the one you pick — the rest of each transaction (amounts, dates, VAT) stays the same.</div>
    </EntryModalShell>
  );
}

function TransactionDetailsPage({ data, setData, postings, coaMap }) {
  const [coaCode, setCoaCode] = useState("");
  const [filter, setFilter] = usePeriodFilter();
  const sel = useRowSelection();
  const [reclassOpen, setReclassOpen] = useState(false);

  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));

  const rows = useMemo(() => {
    if (!coaCode) return [];
    return postings
      .filter((p) => p.code === coaCode && matchesPeriod(p.date, filter))
      .sort((a, b) => (parseAppDate(a.date) || 0) - (parseAppDate(b.date) || 0));
  }, [postings, coaCode, filter]);

  const totals = rows.reduce((a, r) => ({ debit: a.debit + r.debit, credit: a.credit + r.credit }), { debit: 0, credit: 0 });
  const netMovement = round2(totals.debit - totals.credit);
  const selectableIds = rows.filter((r) => r.reclassable).map((r) => r.id);

  const onReclassApply = (newCode) => {
    rows.forEach((r) => { if (sel.selected.has(r.id)) reclassPosting(setData, r, newCode); });
    setReclassOpen(false);
    sel.clear();
  };

  return (
    <div>
      <SectionHeader icon={Search} title="Transaction Details" subtitle="Drill into every posting for one account, filtered by period — reclass any that were miscoded." />

      <div className="td-controls">
        <div className="td-field">
          <label>Chart of Accounts</label>
          <Field type="combo" options={acctOptions} value={coaCode} onChange={(v) => { setCoaCode(v); sel.clear(); }} />
        </div>
        <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
      </div>

      {!coaCode && <div className="empty-panel">Pick an account above to see every transaction posted to it.</div>}

      {coaCode && (
        <>
          <div className="kpi-row td-kpi-row">
            <div className="kpi-card"><div className="kpi-label">Total Debit</div><div className="kpi-value">{fmt(totals.debit)}</div></div>
            <div className="kpi-card"><div className="kpi-label">Total Credit</div><div className="kpi-value">{fmt(totals.credit)}</div></div>
            <div className="kpi-card"><div className="kpi-label">Net Movement</div><div className={"kpi-value" + (netMovement < 0 ? " neg" : "")}>{fmt(netMovement)}</div></div>
          </div>

          {sel.selected.size > 0 && (
            <SelectionBar count={sel.selected.size} onClear={sel.clear}>
              <button className="io-btn accent" onClick={() => setReclassOpen(true)}>Reclass (Quick Fix)</button>
            </SelectionBar>
          )}

          <div className="ledger-wrap">
            <div className="table-scroll">
              <table className="ledger-table report-table">
                <thead>
                  <tr>
                    <th style={{width:32}}><SelectAllCheckbox ids={selectableIds} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
                    <th style={{minWidth:110}}>Date</th><th style={{minWidth:150}}>Source</th><th style={{minWidth:120}}>Reference</th>
                    <th style={{minWidth:170}}>Supplier / Customer</th><th style={{minWidth:200}}>Description</th>
                    <th style={{minWidth:110}} className="num-head">Debit</th><th style={{minWidth:110}} className="num-head">Credit</th>
                    <th style={{minWidth:90}}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.length === 0 && <tr><td colSpan={9} className="empty-row">No transactions posted to this account for the selected period.</td></tr>}
                  {rows.map((r) => (
                    <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                      <td className="text-center">
                        {r.reclassable ? <RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /> : <span title="Generated automatically — not directly reclassable" className="td-locked">🔒</span>}
                      </td>
                      <td>{r.date}</td>
                      <td>{r.source}</td>
                      <td>{r.ref || "—"}</td>
                      <td>{r.party || "—"}</td>
                      <td>{r.desc || "—"}</td>
                      <td className="num">{r.debit ? fmt(r.debit) : "—"}</td>
                      <td className="num">{r.credit ? fmt(r.credit) : "—"}</td>
                      <td>{r.reclassable ? <span className="td-tag reclass">Reclassable</span> : <span className="td-tag system">System</span>}</td>
                    </tr>
                  ))}
                </tbody>
                {rows.length > 0 && (
                  <tfoot><tr>
                    <td colSpan={6} className="totals-label">Total</td>
                    <td className="num">{fmt(totals.debit)}</td><td className="num">{fmt(totals.credit)}</td><td></td>
                  </tr></tfoot>
                )}
              </table>
            </div>
          </div>
        </>
      )}

      {reclassOpen && (
        <ReclassModal data={data} count={sel.selected.size} currentCode={coaCode} onCancel={() => setReclassOpen(false)} onApply={onReclassApply} />
      )}
    </div>
  );
}

/* ============================== TRIAL BALANCE ============================== */

function TrialBalancePage({ data, postings, coaMap }) {
  const [year, setYear] = useState(2026);
  const [month, setMonth] = useState(11);
  const end = monthEnd(year, month);
  const [, forceRefresh] = useState(0);

  const rows = useMemo(() => {
    const sums = {};
    postings.forEach((p) => {
      if (!upTo(p.date, end)) return;
      if (!sums[p.code]) sums[p.code] = { debit: 0, credit: 0 };
      sums[p.code].debit += p.debit; sums[p.code].credit += p.credit;
    });
    return data.coa.map((acct) => {
      const s = sums[acct.code] || { debit: 0, credit: 0 };
      const net = round2(s.debit - s.credit);
      return { acct, debit: net > 0 ? net : 0, credit: net < 0 ? -net : 0 };
    }).filter((r) => r.debit !== 0 || r.credit !== 0);
  }, [postings, end, data.coa]);

  const totals = rows.reduce((a, r) => ({ debit: a.debit + r.debit, credit: a.credit + r.credit }), { debit: 0, credit: 0 });
  const diff = round2(totals.debit - totals.credit);

  return (
    <div className="pnl-shell">
      <div className="pnl-toolbar">
        <div className="pnl-toolbar-top">
          <h1>Trial Balance</h1>
          <div className="pnl-badge">Every account, Debit vs. Credit <ChevronDown size={13} /></div>
        </div>
        <div className="pnl-daterange-label">As of: <strong>{MONTHS[month]} {year}</strong></div>
        <div className="pnl-controls-row">
          <div className="pnl-date-input">{formatDMY(end)}</div>
          <select className="pnl-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>
            {getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
          <select className="pnl-select" value={month} onChange={(e) => setMonth(Number(e.target.value))}>
            {MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}
          </select>
          <div className="pnl-toolbar-actions">
            <button className="pnl-icon-btn" title="Download"><FileDown size={15} /></button>
            <button className="pnl-refresh-btn" onClick={() => forceRefresh((n) => n + 1)}><RefreshCw size={13} /> Refresh</button>
          </div>
        </div>
      </div>

      <div className="pnl-card">
        <div className="pnl-card-head">
          <div className="pnl-card-title">Trial Balance</div>
          <div className="pnl-company">{data.company.name || "Your Company Name Inc."}</div>
          <div className="pnl-currency">PHP (Philippine Peso)</div>
          <div className="pnl-period">As of {formatDMY(end)}.</div>
        </div>

        {rows.length === 0 ? (
          <div className="pnl-table"><div className="dim" style={{ padding: "14px 4px" }}>No posted activity as of {MONTHS[month]} {year} yet.</div></div>
        ) : (
          <table className="pnl-table">
            <tbody>
              <tr className="pnl-section"><td>Account</td><td className="num">Debit</td><td className="num">Credit</td></tr>
              {rows.map((r) => (
                <tr className="pnl-line" key={r.acct.code}>
                  <td>{r.acct.code} · {r.acct.name}</td>
                  <td className="num">{r.debit ? fmtPlain(r.debit) : "—"}</td>
                  <td className="num">{r.credit ? fmtPlain(r.credit) : "—"}</td>
                </tr>
              ))}
              <tr className="pnl-grand pnl-final">
                <td>Total</td><td className="num">{fmtPlain(totals.debit)}</td><td className="num">{fmtPlain(totals.credit)}</td>
              </tr>
            </tbody>
          </table>
        )}
      </div>

      <div className={"check-strip" + (Math.abs(diff) < 0.01 ? " ok" : " bad")}>
        {Math.abs(diff) < 0.01 ? "✓ Trial balance is in balance." : `⚠ Out of balance by ${fmt(diff)} — check General Journal entries.`}
      </div>
    </div>
  );
}

/* ============================== INCOME STATEMENT ============================== */

function computeIncomeStatement(postings, coa, start, end, begInv, endInv) {
  const sumCode = (code) => { let d = 0, c = 0; postings.forEach((p) => { if (p.code === code && inRange(p.date, start, end)) { d += p.debit; c += p.credit; } }); return { d, c }; };
  const sales = sumCode("4000").c - sumCode("4000").d;
  const salesDiscount = sumCode("4001").d - sumCode("4001").c;
  const netSales = round2(sales - salesDiscount);
  const purchases = sumCode("5001").d - sumCode("5001").c;
  const handling = sumCode("5002").d - sumCode("5002").c;
  const purchReturns = sumCode("5003").c - sumCode("5003").d;
  const cogs = round2(begInv + purchases + handling - purchReturns - endInv);
  const costOfServicesAccounts = coa.filter((a) => a.type === "Cost of Sales" && a.category === "Cost of services");
  const costOfServices = costOfServicesAccounts.map((a) => {
    const s = sumCode(a.code);
    return { name: a.name, amount: round2(s.d - s.c) };
  }).filter((e) => e.amount !== 0);
  const totalCostOfServices = round2(costOfServices.reduce((a, e) => a + e.amount, 0));
  const grossProfit = round2(netSales - cogs - totalCostOfServices);
  const otherIncome = sumCode("4002").c - sumCode("4002").d;
  const expenseAccounts = coa.filter((a) => a.type === "Expense");
  const expenses = expenseAccounts.map((a) => {
    const s = sumCode(a.code);
    return { name: a.name, amount: round2(s.d - s.c) };
  }).filter((e) => e.amount !== 0);
  const totalExpenses = round2(expenses.reduce((a, e) => a + e.amount, 0));
  const netIncome = round2(grossProfit + otherIncome - totalExpenses);
  return { sales: round2(sales), salesDiscount: round2(salesDiscount), netSales, begInv, purchases: round2(purchases), handling: round2(handling), purchReturns: round2(purchReturns), endInv, cogs, costOfServices, totalCostOfServices, grossProfit, otherIncome: round2(otherIncome), expenses, totalExpenses, netIncome };
}

function IncomeStatementPage({ data, setData, postings, coaMap }) {
  const [year, setYear] = useState(2026);
  const [periodType, setPeriodType] = useState("Full Year");
  const [monthIdx, setMonthIdx] = useState(0);
  const [quarterIdx, setQuarterIdx] = useState(0);

  let start, end, label;
  if (periodType === "Month") { start = new Date(year, monthIdx, 1); end = monthEnd(year, monthIdx); label = `${MONTHS[monthIdx]} ${year}`; }
  else if (periodType === "Quarter") { const q = QUARTERS[quarterIdx]; start = new Date(year, q.months[0], 1); end = monthEnd(year, q.months[2]); label = `${q.label} ${year}`; }
  else { start = new Date(year, 0, 1); end = monthEnd(year, 11); label = `Calendar Year ${year}`; }

  // Beginning Inventory is anchored to the YEAR (the Jan 1 balance) regardless of which period
  // you're viewing — stored under month 0 so Month/Quarter/Full-Year views all agree on it.
  const yearBeginRecord = data.inventoryLog.find((x) => x.year === year && x.month === 0) || { beginning: 0 };
  const inv = data.inventoryLog.find((x) => x.year === year && x.month === (periodType === "Month" ? monthIdx : 11)) || { beginning: 0, ending: 0 };
  const setBeginning = (v) => {
    setData((d) => {
      const existing = d.inventoryLog.find((x) => x.year === year && x.month === 0);
      if (existing) return { ...d, inventoryLog: d.inventoryLog.map((x) => x === existing ? { ...x, beginning: num(v) } : x) };
      return { ...d, inventoryLog: [...d.inventoryLog, { id: uid(), year, month: 0, beginning: num(v), ending: 0 }] };
    });
  };
  const setEnding = (v) => {
    const m = periodType === "Month" ? monthIdx : 11;
    setData((d) => {
      const existing = d.inventoryLog.find((x) => x.year === year && x.month === m);
      if (existing) return { ...d, inventoryLog: d.inventoryLog.map((x) => x === existing ? { ...x, ending: num(v) } : x) };
      return { ...d, inventoryLog: [...d.inventoryLog, { id: uid(), year, month: m, beginning: 0, ending: num(v) }] };
    });
  };

  const stmt = useMemo(() => computeIncomeStatement(postings, data.coa, start, end, num(yearBeginRecord.beginning), num(inv.ending)), [postings, data.coa, start, end, yearBeginRecord, inv]);
  const operatingProfit = round2(stmt.grossProfit - stmt.totalExpenses);
  const [, forceRefresh] = useState(0);

  return (
    <div className="pnl-shell">
      <div className="pnl-toolbar">
        <div className="pnl-toolbar-top">
          <h1>Profit &amp; Loss Statement</h1>
          <div className="pnl-badge">Standard Profit &amp; Loss <ChevronDown size={13} /></div>
        </div>
        <div className="pnl-daterange-label">Date Range: <strong>{label}</strong></div>
        <div className="pnl-controls-row">
          <div className="pnl-date-input">{formatDMY(start)}</div>
          <div className="pnl-date-input">{formatDMY(end)}</div>
          <select className="pnl-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>
            {getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
          <select className="pnl-select" value={periodType} onChange={(e) => setPeriodType(e.target.value)}>
            <option>Full Year</option><option>Quarter</option><option>Month</option>
          </select>
          {periodType === "Quarter" && <select className="pnl-select" value={quarterIdx} onChange={(e) => setQuarterIdx(Number(e.target.value))}>{QUARTERS.map((q, i) => <option key={q.label} value={i}>{q.label}</option>)}</select>}
          {periodType === "Month" && <select className="pnl-select" value={monthIdx} onChange={(e) => setMonthIdx(Number(e.target.value))}>{MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}</select>}
          <div className="pnl-toolbar-actions">
            <button className="pnl-icon-btn" title="Chart view"><BarChart3 size={15} /></button>
            <button className="pnl-icon-btn" title="Download"><FileDown size={15} /></button>
            <button className="pnl-refresh-btn" onClick={() => forceRefresh((n) => n + 1)}><RefreshCw size={13} /> Refresh</button>
          </div>
        </div>
      </div>

      <div className="pnl-card">
        <div className="pnl-card-head">
          <div className="pnl-card-title">Profit &amp; Loss</div>
          <div className="pnl-company">{data.company.name || "Your Company Name Inc."}</div>
          <div className="pnl-currency">PHP (Philippine Peso)</div>
          <div className="pnl-period">For the period {formatDMY(start)} to {formatDMY(end)}.</div>
        </div>

        <table className="pnl-table">
          <tbody>
            <tr className="pnl-section"><td colSpan={2}>Revenue</td></tr>
            <tr className="pnl-line"><td>Sales</td><td className="num">{fmtPlain(stmt.sales)}</td></tr>
            <tr className="pnl-line"><td>Sales Returns</td><td className="num">({fmtPlain(stmt.salesDiscount)})</td></tr>
            <tr className="pnl-total"><td>Total Revenue</td><td className="num">{fmtPlain(stmt.netSales)}</td></tr>

            <tr className="pnl-section"><td colSpan={2}>Cost of Goods Sold</td></tr>
            <tr className="pnl-line">
              <td>Beginning Inventory</td>
              <td className="num pnl-input-cell"><Field type="number" align="right" value={yearBeginRecord.beginning} onChange={setBeginning} /></td>
            </tr>
            <tr className="pnl-line"><td>Purchases</td><td className="num">{fmtPlain(stmt.purchases)}</td></tr>
            <tr className="pnl-line"><td>Purchase Returns &amp; Allowances</td><td className="num">({fmtPlain(stmt.purchReturns)})</td></tr>
            <tr className="pnl-line"><td>Handling Cost</td><td className="num">{fmtPlain(stmt.handling)}</td></tr>
            <tr className="pnl-line">
              <td>Ending Inventory</td>
              <td className="num pnl-input-cell">(<Field type="number" align="right" value={inv.ending} onChange={setEnding} />)</td>
            </tr>
            <tr className="pnl-total"><td>Total Cost of Goods Sold</td><td className="num">{fmtPlain(stmt.cogs)}</td></tr>

            <tr className="pnl-section"><td colSpan={2}>Cost of Services</td></tr>
            {stmt.costOfServices.length === 0 && <tr className="pnl-line"><td className="dim">No cost of services posted this period</td><td className="num">—</td></tr>}
            {stmt.costOfServices.map((e) => <tr className="pnl-line" key={e.name}><td>{e.name}</td><td className="num">{fmtPlain(e.amount)}</td></tr>)}
            <tr className="pnl-total"><td>Total Cost of Services</td><td className="num">{fmtPlain(stmt.totalCostOfServices)}</td></tr>

            <tr className="pnl-grand"><td>Gross Profit</td><td className="num">{fmtPlain(stmt.grossProfit)}</td></tr>

            <tr className="pnl-section"><td colSpan={2}>Operating Expense</td></tr>
            {stmt.expenses.length === 0 && <tr className="pnl-line"><td className="dim">No expenses posted this period</td><td className="num">—</td></tr>}
            {stmt.expenses.map((e) => <tr className="pnl-line" key={e.name}><td>{e.name}</td><td className="num">{fmtPlain(e.amount)}</td></tr>)}
            <tr className="pnl-total"><td>Total Operating Expense</td><td className="num">{fmtPlain(stmt.totalExpenses)}</td></tr>
            <tr className="pnl-grand"><td>Operating Profit</td><td className="num">{fmtPlain(operatingProfit)}</td></tr>

            {stmt.otherIncome !== 0 && <tr className="pnl-line"><td>Other Income</td><td className="num">{fmtPlain(stmt.otherIncome)}</td></tr>}
            <tr className="pnl-grand pnl-final"><td>Net Profit</td><td className="num">{fmtPlain(stmt.netIncome)}</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== BALANCE SHEET ============================== */

function BalanceSheetPage({ data, postings, coaMap }) {
  const [year, setYear] = useState(2026);
  const [month, setMonth] = useState(11);
  const end = monthEnd(year, month);
  const start0 = new Date(2020, 0, 1);
  const [, forceRefresh] = useState(0);

  const result = useMemo(() => {
    const sums = {};
    postings.forEach((p) => { if (!upTo(p.date, end)) return; if (!sums[p.code]) sums[p.code] = { d: 0, c: 0 }; sums[p.code].d += p.debit; sums[p.code].c += p.credit; });

    const assets = data.coa.filter((a) => a.type === "Asset").map((a) => { const s = sums[a.code] || { d: 0, c: 0 }; return { name: a.name, amount: round2(s.d - s.c) }; }).filter((r) => r.amount !== 0);
    const liabilities = data.coa.filter((a) => a.type === "Liability").map((a) => { const s = sums[a.code] || { d: 0, c: 0 }; return { name: a.name, amount: round2(s.c - s.d) }; }).filter((r) => r.amount !== 0);

    const capitalStock = (() => { const s = sums["3000"] || { d: 0, c: 0 }; return round2(s.c - s.d); })();
    const drawing = (() => { const s = sums["3002"] || { d: 0, c: 0 }; return round2(s.d - s.c); })();
    const reDirectPost = (() => { const s = sums["3001"] || { d: 0, c: 0 }; return round2(s.c - s.d); })();

    let netIncome = 0;
    // Dynamically pulls every Revenue/Cost of Sales/Expense account from the COA — not a hardcoded
    // list — so any custom account (like a Cost of Services line) is never silently excluded from
    // Net Income, which would otherwise overstate Retained Earnings and throw the sheet out of balance.
    const incomeStatementCodes = data.coa.filter((a) => ["Revenue", "Cost of Sales", "Expense"].includes(a.type)).map((a) => a.code);
    incomeStatementCodes.forEach((code) => {
      const s = sums[code] || { d: 0, c: 0 };
      const acct = coaMap[code];
      if (!acct) return;
      if (acct.type === "Revenue") netIncome += (s.c - s.d) * (code === "4001" ? -1 : 1);
      else netIncome -= (s.d - s.c);
    });
    netIncome = round2(netIncome);

    const totalAssets = round2(assets.reduce((a, r) => a + r.amount, 0));
    const totalLiabilities = round2(liabilities.reduce((a, r) => a + r.amount, 0));
    const totalEquity = round2(capitalStock + reDirectPost + netIncome - drawing);
    const check = round2(totalAssets - totalLiabilities - totalEquity);

    return { assets, liabilities, capitalStock, drawing, retainedEarnings: round2(reDirectPost + netIncome), totalAssets, totalLiabilities, totalEquity, check };
  }, [postings, end, data.coa, coaMap]);

  return (
    <div className="pnl-shell">
      <div className="pnl-toolbar">
        <div className="pnl-toolbar-top">
          <h1>Balance Sheet</h1>
          <div className="pnl-badge">Assets, Liabilities &amp; Equity <ChevronDown size={13} /></div>
        </div>
        <div className="pnl-daterange-label">As of: <strong>{MONTHS[month]} {year}</strong></div>
        <div className="pnl-controls-row">
          <div className="pnl-date-input">{formatDMY(end)}</div>
          <select className="pnl-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>
            {getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
          <select className="pnl-select" value={month} onChange={(e) => setMonth(Number(e.target.value))}>
            {MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}
          </select>
          <div className="pnl-toolbar-actions">
            <button className="pnl-icon-btn" title="Download"><FileDown size={15} /></button>
            <button className="pnl-refresh-btn" onClick={() => forceRefresh((n) => n + 1)}><RefreshCw size={13} /> Refresh</button>
          </div>
        </div>
      </div>

      <div className="pnl-bs-grid">
        <div className="pnl-card">
          <div className="pnl-card-head">
            <div className="pnl-card-title">Assets</div>
            <div className="pnl-company">{data.company.name || "Your Company Name Inc."}</div>
            <div className="pnl-currency">PHP (Philippine Peso)</div>
            <div className="pnl-period">As of {formatDMY(end)}.</div>
          </div>
          <table className="pnl-table">
            <tbody>
              {result.assets.length === 0 && <tr className="pnl-line"><td className="dim">No asset balances yet</td><td className="num">—</td></tr>}
              {result.assets.map((a) => <tr className="pnl-line" key={a.name}><td>{a.name}</td><td className="num">{fmtPlain(a.amount)}</td></tr>)}
              <tr className="pnl-grand pnl-final"><td>Total Assets</td><td className="num">{fmtPlain(result.totalAssets)}</td></tr>
            </tbody>
          </table>
        </div>

        <div className="pnl-card">
          <div className="pnl-card-head">
            <div className="pnl-card-title">Liabilities &amp; Equity</div>
            <div className="pnl-company">{data.company.name || "Your Company Name Inc."}</div>
            <div className="pnl-currency">PHP (Philippine Peso)</div>
            <div className="pnl-period">As of {formatDMY(end)}.</div>
          </div>
          <table className="pnl-table">
            <tbody>
              <tr className="pnl-section"><td colSpan={2}>Liabilities</td></tr>
              {result.liabilities.length === 0 && <tr className="pnl-line"><td className="dim">No liability balances yet</td><td className="num">—</td></tr>}
              {result.liabilities.map((a) => <tr className="pnl-line" key={a.name}><td>{a.name}</td><td className="num">{fmtPlain(a.amount)}</td></tr>)}
              <tr className="pnl-total"><td>Total Liabilities</td><td className="num">{fmtPlain(result.totalLiabilities)}</td></tr>
              <tr className="pnl-section"><td colSpan={2}>Equity</td></tr>
              <tr className="pnl-line"><td>Capital Stock</td><td className="num">{fmtPlain(result.capitalStock)}</td></tr>
              <tr className="pnl-line"><td>Retained Earnings (net income to date)</td><td className="num">{fmtPlain(result.retainedEarnings)}</td></tr>
              <tr className="pnl-line"><td>Owner's Drawing</td><td className="num">({fmtPlain(result.drawing)})</td></tr>
              <tr className="pnl-total"><td>Total Equity</td><td className="num">{fmtPlain(result.totalEquity)}</td></tr>
              <tr className="pnl-grand pnl-final"><td>Total Liabilities &amp; Equity</td><td className="num">{fmtPlain(round2(result.totalLiabilities + result.totalEquity))}</td></tr>
            </tbody>
          </table>
        </div>
      </div>

      <div className={"check-strip" + (Math.abs(result.check) < 0.01 ? " ok" : " bad")}>
        {Math.abs(result.check) < 0.01 ? "✓ Balance sheet balances." : `⚠ Out of balance by ${fmt(result.check)} — check journal entries.`}
      </div>
    </div>
  );
}

/* ============================== TAX COMPLIANCE — SLSPI / QAP / SAWT DAT GENERATION ============================== */
// Ported from a standalone RELIEF/Alphalist DAT-file generator whose record layouts were confirmed
// against real, validated BIR reference files. Kept as close to that source logic as possible for
// correctness — this is a compliance artifact, not a place to improvise formatting.

function reliefClean(s) {
  return String(s ?? "")
    .replace(/&/g, " and ")
    .replace(/[^a-zA-Z0-9.\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
function reliefPadTin(t) { return t ? String(t).padStart(9, "0") : ""; }
function reliefFormatTIN(tin) {
  const t = String(tin ?? "").replace(/\D/g, "").padStart(9, "0");
  return `${t.slice(0, 3)}-${t.slice(3, 6)}-${t.slice(6, 9)}`;
}
function reliefLastDayOfMonth(year, month) { return new Date(year, month, 0).getDate(); }
function reliefPadBranch(b) { return String(b || "0000").replace(/\D/g, "").padStart(4, "0"); }
function reliefEntityParts(d) {
  if (d.isIndividual) return ["", d.indivSurname, d.indivFirstName, d.indivMiddleName];
  return [d.companyName, "", "", ""];
}
function reliefToMMDDYYYY(v) {
  if (v instanceof Date && !isNaN(v)) return `${pad2(v.getMonth() + 1)}/${pad2(v.getDate())}/${v.getFullYear()}`;
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return `${pad2(d.getUTCMonth() + 1)}/${pad2(d.getUTCDate())}/${d.getUTCFullYear()}`;
  }
  if (typeof v === "string" && v.trim() !== "") {
    const t = v.trim();
    const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return `${pad2(m[1])}/${pad2(m[2])}/${m[3]}`;
    const parsed = new Date(t);
    if (!isNaN(parsed)) return `${pad2(parsed.getMonth() + 1)}/${pad2(parsed.getDate())}/${parsed.getFullYear()}`;
    return t;
  }
  return "";
}
function reliefCellToMonthYear(v) {
  if (v instanceof Date && !isNaN(v)) return { month: v.getMonth() + 1, year: v.getFullYear(), relative: false };
  if (typeof v === "number") {
    if (v >= 1 && v <= 3 && Number.isInteger(v)) return { month: v, year: null, relative: true };
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d)) return { month: d.getUTCMonth() + 1, year: d.getUTCFullYear(), relative: false };
  }
  if (typeof v === "string" && v.trim() !== "") {
    const t = v.trim();
    if (/^[123]$/.test(t)) return { month: parseInt(t, 10), year: null, relative: true };
    const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return { month: parseInt(m[1], 10), year: parseInt(m[3], 10), relative: false };
    const parsed = new Date(t);
    if (!isNaN(parsed)) return { month: parsed.getMonth() + 1, year: parsed.getFullYear(), relative: false };
  }
  return null;
}

const SLSPI_CONFIG = {
  SLP: {
    marker: "P", title: "Summary List of Purchases",
    expectedHeaders: ["TIN","SUPPLIERS NAME","SURNAME","FIRST NAME","MIDDLE NAME","ADDRESS 1","ADDRESS 2","VAT EXEMPT","ZERO RATED","SERVICES","CAPITAL GOODS","OTHER THAN CAPITAL GOODS","INPUT VAT"],
    note: "Amounts kept at full floating-point precision (no rounding), matching a validated SLP DAT file. Always run the file through the BIR Validation Module before emailing esubmission@bir.gov.ph.",
  },
  SLS: {
    marker: "S", title: "Summary List of Sales",
    expectedHeaders: ["TIN","CUSTOMERS NAME","SURNAME","FIRST NAME","MIDDLE NAME","ADDRESS 1","ADDRESS 2","EXEMPT SALES","ZERO RATED SALES","TAXABLE SALES","OUTPUT VAT"],
    note: "Detail amounts fixed to 2 decimals; header totals truncated to whole numbers, matching a validated SLS DAT file. Always run the file through the BIR Validation Module before emailing esubmission@bir.gov.ph.",
  },
  SLI: {
    marker: "I", title: "Summary List of Importation",
    expectedHeaders: ["IMPORT ENTRY NO","ASSESSMENT DATE","SUPPLIERS NAME","IMPORTATION DATE","COUNTRY ORIGIN","TOTAL LANDED COST","OTHER CHARGES","EXEMPT","TAXABLE","VAT","OR NUMBER","VAT PAYMENT DATE"],
    note: "Text fields quoted; detail amounts unrounded; header totals forced to 2 decimals, matching a validated Importation DAT file. Always run the file through the BIR Validation Module before emailing esubmission@bir.gov.ph.",
  },
};

// A VAT-registered business's collected amounts are presumed VAT-inclusive, so the withholding
// base is the amount net of VAT (÷1.12). A Non-VAT business has no VAT baked in, so the full
// amount is the base. Used for both Cash Receipts' CWT and SAWT's tax base.
function cwtTaxBase(data, amount) {
  const isVat = (data.company.vatStatus || "VAT Registered") !== "Non-VAT";
  return isVat ? round2(num(amount) / 1.12) : round2(num(amount));
}

function buildFilingDetails(data, period) {
  const c = data.company;
  const isIndividual = c.taxpayerType === "Individual";
  const displayName = isIndividual
    ? [c.surname, [c.firstName, c.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", ")
    : (c.name || c.registeredName || "");
  return {
    companyTin: (c.tin || "").replace(/\D/g, ""),
    rdoCode: c.rdo || "",
    branchCode: c.branchCode || "0000",
    taxpayerType: c.taxpayerType, isIndividual,
    companyName: c.name || c.registeredName || "",
    indivSurname: c.surname || "", indivFirstName: c.firstName || "", indivMiddleName: c.middleName || "",
    displayName,
    tradeName: c.tradeName || displayName,
    addr1: c.addr1 || c.address || "",
    addr2: c.addr2 || "",
    zipCode: c.zipCode || "",
    preparedBy: c.preparedBy || "",
    // Payor's Authorized Representative for the 2307 signature block.
    authorizedSignatory: c.authorizedSignatory || "",
    signatoryPosition: c.signatoryPosition || "",
    signatoryTin: c.signatoryTin ? reliefFormatTIN(c.signatoryTin) : "",
    signatureImage: c.signatureImage || "",
    month: period.month, year: period.year, quarter: period.quarter,
    sawtFormType: period.sawtFormType,
  };
}

function downloadDatFile(record) {
  const lines = [record.hRow.join(","), ...record.dRows.map((r) => r.join(",")), ...(record.cRow ? [record.cRow.join(",")] : [])];
  const blob = new Blob([lines.join("\r\n")], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = record.filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function downloadQapZip(groups, d) {
  const JSZipCtor = await withTimeout(loadJSZip(), 8000);
  const zip = new JSZipCtor();
  groups.forEach((g) => {
    const lines = [g.hRow.join(","), ...g.dRows.map((r) => r.join(",")), g.cRow.join(",")];
    zip.file(g.filename, lines.join("\r\n"));
  });
  const blob = await withTimeout(zip.generateAsync({ type: "blob" }), 8000);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `${d.companyTin}_QAP_Q${d.quarter}_${d.year}_DAT_files.zip`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function buildSlspiRecord(subMode, d, sourceRows) {
  const day = reliefLastDayOfMonth(d.year, d.month);
  const returnPeriod = `${pad2(d.month)}/${pad2(day)}/${d.year}`;
  const marker = SLSPI_CONFIG[subMode].marker;

  if (subMode === "SLP") {
    const numf = (n) => String(n);
    const totals = { vatExempt: 0, zeroRated: 0, services: 0, capitalGoods: 0, otherCapitalGoods: 0, inputVat: 0 };
    sourceRows.forEach((r) => {
      totals.vatExempt += r.vatExempt; totals.zeroRated += r.zeroRated; totals.services += r.services;
      totals.capitalGoods += r.capitalGoods; totals.otherCapitalGoods += r.otherCapitalGoods; totals.inputVat += r.inputVat;
    });
    const hRow = ["H", marker, d.companyTin, ...reliefEntityParts(d).map(reliefClean), reliefClean(d.tradeName), reliefClean(d.addr1), reliefClean(d.addr2),
      numf(totals.vatExempt), numf(totals.zeroRated), numf(totals.services), numf(totals.capitalGoods),
      numf(totals.otherCapitalGoods), numf(totals.inputVat), numf(totals.inputVat), "0", d.rdoCode, returnPeriod, "12", ""];
    const dRows = sourceRows.map((r) => ["D", marker, r.tin, reliefClean(r.partyName), reliefClean(r.surname), reliefClean(r.firstName), reliefClean(r.middleName),
      reliefClean(r.address1), reliefClean(r.address2), numf(r.vatExempt), numf(r.zeroRated), numf(r.services), numf(r.capitalGoods),
      numf(r.otherCapitalGoods), numf(r.inputVat), d.companyTin, returnPeriod, "", "", "", "", ""]);
    return { hRow, dRows, cRow: null, totals, totalLabel: "Total Input VAT", totalVal: totals.inputVat, returnPeriod, filename: `${d.companyTin}${marker}${pad2(d.month)}${d.year}_SLP.dat` };
  }
  if (subMode === "SLS") {
    const fmt2 = (n) => Number(n).toFixed(2);
    const truncInt = (n) => String(Math.trunc(n));
    const totals = { exemptSales: 0, zeroRatedSales: 0, taxableSales: 0, outputVat: 0 };
    sourceRows.forEach((r) => {
      totals.exemptSales += r.exemptSales; totals.zeroRatedSales += r.zeroRatedSales;
      totals.taxableSales += r.taxableSales; totals.outputVat += r.outputVat;
    });
    const hRow = ["H", marker, d.companyTin, ...reliefEntityParts(d).map(reliefClean), reliefClean(d.tradeName), reliefClean(d.addr1), reliefClean(d.addr2),
      truncInt(totals.exemptSales), truncInt(totals.zeroRatedSales), truncInt(totals.taxableSales), truncInt(totals.outputVat), d.rdoCode, returnPeriod, "12"];
    const dRows = sourceRows.map((r) => ["D", marker, reliefPadTin(r.tin), reliefClean(r.partyName), reliefClean(r.surname), reliefClean(r.firstName), reliefClean(r.middleName),
      reliefClean(r.address1), reliefClean(r.address2), fmt2(r.exemptSales), fmt2(r.zeroRatedSales), fmt2(r.taxableSales), fmt2(r.outputVat), d.companyTin, returnPeriod, "", ""]);
    return { hRow, dRows, cRow: null, totals, totalLabel: "Total Output VAT", totalVal: totals.outputVat, returnPeriod, filename: `${d.companyTin}${marker}${pad2(d.month)}${d.year}_SLS.dat` };
  }
  const q = (s) => `"${reliefClean(s).replace(/"/g, "")}"`;
  const numf = (n) => String(n);
  const fmt2 = (n) => Number(n).toFixed(2);
  const totals = { totalLandedCost: 0, otherCharges: 0, exempt: 0, taxable: 0, vat: 0 };
  sourceRows.forEach((r) => {
    totals.totalLandedCost += r.totalLandedCost; totals.otherCharges += r.otherCharges;
    totals.exempt += r.exempt; totals.taxable += r.taxable; totals.vat += r.vat;
  });
  const hRow = ["H", marker, q(d.companyTin), ...reliefEntityParts(d).map(q), q(d.tradeName), q(d.addr1), q(d.addr2),
    fmt2(totals.totalLandedCost), fmt2(totals.otherCharges), fmt2(totals.exempt), fmt2(totals.taxable), fmt2(totals.vat), d.rdoCode, returnPeriod, "12"];
  const dRows = sourceRows.map((r) => ["D", marker, q(r.importEntryNo), r.assessmentDate, q(r.suppliersName), r.importationDate,
    q(r.countryOrigin), numf(r.totalLandedCost), numf(r.otherCharges), numf(r.exempt), numf(r.taxable), numf(r.vat), q(r.orNumber), r.vatPaymentDate, d.companyTin, returnPeriod]);
  return { hRow, dRows, cRow: null, totals, totalLabel: "Total VAT", totalVal: totals.vat, returnPeriod, filename: `${d.companyTin}${marker}${pad2(d.month)}${d.year}_SLI.dat` };
}

function reliefResolveRowMonthYear(r, d) {
  const quarterStartMonth = (d.quarter - 1) * 3 + 1;
  if (!r.monthInfo) return { month: ((quarterStartMonth + 2 - 1) % 12) + 1, year: d.year };
  if (r.monthInfo.relative) {
    const calMonth = quarterStartMonth + (r.monthInfo.month - 1);
    const calYear = calMonth > 12 ? d.year + 1 : d.year;
    return { month: ((calMonth - 1) % 12) + 1, year: calYear };
  }
  return { month: r.monthInfo.month, year: r.monthInfo.year };
}

function buildQapGroups(d, sourceRows) {
  const qOpt = (s) => (s ? `"${reliefClean(s).replace(/"/g, "")}"` : "");
  const fmt2 = (n) => Number(n).toFixed(2);
  const byKey = new Map();
  sourceRows.forEach((r) => {
    const { month, year } = reliefResolveRowMonthYear(r, d);
    const key = `${year}-${pad2(month)}`;
    if (!byKey.has(key)) byKey.set(key, { month, year, rows: [] });
    byKey.get(key).rows.push(r);
  });
  const sortedKeys = [...byKey.keys()].sort();
  return sortedKeys.map((key) => {
    const { month, year, rows: rowsForMonth } = byKey.get(key);
    const monthReturnPeriod = `${pad2(month)}/${year}`;
    const hRow = ["HQAP", "H1601EQ", d.companyTin, reliefPadBranch(d.branchCode), qOpt(d.displayName), monthReturnPeriod, d.rdoCode];
    const dRows = rowsForMonth.map((r, i) => ["D1", "1601EQ", String(i + 1), reliefPadTin(r.tin), reliefPadBranch(r.branchCode),
      qOpt(r.partyName), qOpt(r.surname), qOpt(r.firstName), qOpt(r.middleName), monthReturnPeriod, r.atc, fmt2(r.rate), fmt2(r.taxBase), fmt2(r.ewt)]);
    const totals = { taxBase: 0, ewt: 0 };
    rowsForMonth.forEach((r) => { totals.taxBase += r.taxBase; totals.ewt += r.ewt; });
    const cRow = ["C1", "1601EQ", d.companyTin, reliefPadBranch(d.branchCode), monthReturnPeriod, fmt2(totals.taxBase), fmt2(totals.ewt)];
    return { calMonth: month, calYear: year, returnPeriod: monthReturnPeriod, hRow, dRows, cRow, totals, filename: `${d.companyTin}${reliefPadBranch(d.branchCode)}${pad2(month)}${year}1601EQ.dat` };
  });
}

function buildSawtRecord(d, sourceRows) {
  const round2str = (n) => String(round2(n));
  const fmt2 = (n) => Number(n).toFixed(2);
  const sawtPeriod = `${pad2(d.month)}/${d.year}`;
  const ft = d.sawtFormType || (d.isIndividual ? "1701Q" : "1702Q");
  const isLegacyNonIndividual = ft === "1702Q";

  let hNameParts, dNameFn, amtFn, hTrailing, cTrailing;
  if (isLegacyNonIndividual) {
    hNameParts = reliefEntityParts(d).map(reliefClean);
    dNameFn = reliefClean; amtFn = round2str;
    hTrailing = ["", "", "", "", ""]; cTrailing = ["", "", "", "", "", "", "", ""];
  } else {
    const alwaysQuote = (s) => `"${reliefClean(s)}"`;
    const qOpt = (s) => (s ? `"${reliefClean(s).replace(/"/g, "")}"` : "");
    hNameParts = reliefEntityParts(d).map(alwaysQuote);
    dNameFn = qOpt; amtFn = fmt2;
    hTrailing = []; cTrailing = [];
  }

  const padBranch = reliefPadBranch(d.branchCode);
  const hRow = ["HSAWT", `H${ft}`, d.companyTin, padBranch, ...hNameParts, sawtPeriod, d.rdoCode, ...hTrailing];
  const dRows = sourceRows.map((r, i) => ["DSAWT", `D${ft}`, String(i + 1), reliefPadTin(r.tin), reliefPadBranch(r.branchCode),
    dNameFn(r.partyName), dNameFn(r.surname), dNameFn(r.firstName), dNameFn(r.middleName), sawtPeriod, "", r.atc, amtFn(r.rate), amtFn(r.taxBase), amtFn(r.cwt)]);
  const totals = { taxBase: 0, cwt: 0 };
  sourceRows.forEach((r) => { totals.taxBase += r.taxBase; totals.cwt += r.cwt; });
  const cRow = ["CSAWT", `C${ft}`, d.companyTin, padBranch, sawtPeriod, amtFn(totals.taxBase), amtFn(totals.cwt), ...cTrailing];
  return { hRow, dRows, cRow, totals, returnPeriod: sawtPeriod, filename: `${d.companyTin}${padBranch}${pad2(d.month)}${d.year}${ft}.dat` };
}

/* ============================== SLSPI / QAP / SAWT — HUMAN-READABLE EXCEL & PDF REPORTS ============================== */
// These are the reconciliation-style reports (not the .dat submission files) — for internal review
// and paper trail, matching the reference tool's report layouts as closely as possible.

function reliefFormatLongDate(d) {
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}

function qapSlotIndex(r, d) {
  const { month, year } = reliefResolveRowMonthYear(r, d);
  const quarterStartMonth = (d.quarter - 1) * 3 + 1;
  for (let s = 0; s < 3; s++) {
    const cm = quarterStartMonth + s;
    const cy = cm > 12 ? d.year + 1 : d.year;
    const nm = ((cm - 1) % 12) + 1;
    if (nm === month && cy === year) return s;
  }
  return 2;
}

function useReportExport(pdfFn, excelFn) {
  const [status, setStatus] = useState(null);
  const exportPDF = async () => {
    setStatus({ type: "pending", text: "Generating PDF report…" });
    try { await pdfFn(); setStatus({ type: "success", text: "PDF report downloaded." }); }
    catch (e) { setStatus({ type: "error", text: "Couldn't generate the PDF report — this needs an internet connection the first time. Please check your connection and try again." }); }
    setTimeout(() => setStatus(null), 5000);
  };
  const exportExcel = () => {
    try { excelFn(); setStatus({ type: "success", text: "Excel report downloaded." }); }
    catch (e) { setStatus({ type: "error", text: "Couldn't generate the Excel report." }); }
    setTimeout(() => setStatus(null), 4000);
  };
  return { status, exportPDF, exportExcel };
}

function downloadSlspiExcelReport(subMode, data, sourceRows, month, year) {
  const d = buildFilingDetails(data, { month, year, quarter: null, sawtFormType: null });
  const reportDate = new Date(year, month, 0);
  const partyName = (r) => r.partyName || [r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", ");

  let aoa, colCount, moneyCols;
  if (subMode === "SLP") {
    const rows = sourceRows.map((r) => {
      const taxable = r.services + r.capitalGoods + r.otherCapitalGoods;
      const gross = r.vatExempt + r.zeroRated + taxable;
      const grossTaxable = taxable + r.inputVat;
      return [reportDate, r.tin ? reliefFormatTIN(r.tin) : "", r.partyName || "", [r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", "),
        [r.address1, r.address2].filter(Boolean).join(" "), gross, r.vatExempt, r.zeroRated, taxable, r.services, r.capitalGoods, r.otherCapitalGoods, r.inputVat, grossTaxable];
    });
    const totalCols = [5,6,7,8,9,10,11,12,13];
    const totalsRow = new Array(14).fill(""); totalsRow[0] = "Grand Total :";
    totalCols.forEach((c) => { totalsRow[c] = rows.reduce((s, r) => s + (typeof r[c] === "number" ? r[c] : 0), 0); });
    aoa = [
      ["PURCHASE TRANSACTION"], ["RECONCILIATION OF LISTING FOR ENFORCEMENT"], [],
      [`TIN : ${reliefFormatTIN(d.companyTin)}`], [`OWNER'S NAME: ${d.displayName}`],
      [`OWNER'S TRADE NAME : ${d.tradeName}`], [`OWNER'S ADDRESS: ${[d.addr1, d.addr2].filter(Boolean).join(" ")}`], [],
      ["TAXABLE","TAXPAYER","REGISTERED NAME","NAME OF SUPPLIER","SUPPLIER'S ADDRESS","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF"],
      ["MONTH","IDENTIFICATION","","(Last Name, First Name, Middle Name)","","GROSS PURCHASE","EXEMPT PURCHASE","ZERO-RATED PURCHASE","TAXABLE PURCHASE","PURCHASE OF SERVICES","PURCHASE OF CAPITAL GOODS","PURCHASE OF GOODS OTHER THAN CAPITAL GOODS","INPUT TAX","GROSS TAXABLE PURCHASE"],
      ["","NUMBER","","","","","","","","","","","",""],
      ["(1)","(2)","(3)","(4)","(5)","(6)","(7)","(8)","(9)","(10)","(11)","(12)","(13)","(14)"],
      ...rows, totalsRow, ["END OF REPORT"],
    ];
    colCount = 14; moneyCols = [5,6,7,8,9,10,11,12,13];
  } else if (subMode === "SLS") {
    const rows = sourceRows.map((r) => {
      const gross = r.exemptSales + r.zeroRatedSales + r.taxableSales;
      const grossTaxable = r.taxableSales + r.outputVat;
      return [reportDate, r.tin ? reliefFormatTIN(r.tin) : "", r.partyName || "", [r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", "),
        [r.address1, r.address2].filter(Boolean).join(" "), gross, r.exemptSales, r.zeroRatedSales, r.taxableSales, r.outputVat, grossTaxable];
    });
    const totalCols = [5,6,7,8,9,10];
    const totalsRow = new Array(11).fill(""); totalsRow[0] = "Grand Total :";
    totalCols.forEach((c) => { totalsRow[c] = rows.reduce((s, r) => s + (typeof r[c] === "number" ? r[c] : 0), 0); });
    aoa = [
      ["SALES TRANSACTION"], ["RECONCILIATION OF LISTING FOR ENFORCEMENT"], [],
      [`TIN : ${reliefFormatTIN(d.companyTin)}`], [`OWNER'S NAME: ${d.displayName}`],
      [`OWNER'S TRADE NAME : ${d.tradeName}`], [`OWNER'S ADDRESS: ${[d.addr1, d.addr2].filter(Boolean).join(" ")}`], [],
      ["TAXABLE","TAXPAYER","REGISTERED NAME","NAME OF CUSTOMER","CUSTOMER'S ADDRESS","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF"],
      ["MONTH","IDENTIFICATION","","(Last Name, First Name, Middle Name)","","GROSS SALES","EXEMPT SALES","ZERO-RATED SALES","TAXABLE SALES","OUTPUT TAX","GROSS TAXABLE SALES"],
      ["","NUMBER","","","","","","","","",""],
      ["(1)","(2)","(3)","(4)","(5)","(6)","(7)","(8)","(9)","(10)","(11)"],
      ...rows, totalsRow, ["END OF REPORT"],
    ];
    colCount = 11; moneyCols = [5,6,7,8,9,10];
  } else {
    const rows = sourceRows.map((r) => {
      const gross = r.totalLandedCost + r.otherCharges;
      const grossTaxable = r.taxable + r.vat;
      return [reportDate, r.importEntryNo, r.suppliersName, r.countryOrigin, r.assessmentDate, gross, r.exempt, r.taxable, r.vat, grossTaxable, r.orNumber];
    });
    const totalCols = [5,6,7,8,9];
    const totalsRow = new Array(11).fill(""); totalsRow[0] = "Grand Total :";
    totalCols.forEach((c) => { totalsRow[c] = rows.reduce((s, r) => s + (typeof r[c] === "number" ? r[c] : 0), 0); });
    aoa = [
      ["IMPORTATION TRANSACTION"], ["RECONCILIATION OF LISTING FOR ENFORCEMENT"], [],
      [`TIN : ${reliefFormatTIN(d.companyTin)}`], [`OWNER'S NAME: ${d.displayName}`],
      [`OWNER'S TRADE NAME : ${d.tradeName}`], [`OWNER'S ADDRESS: ${[d.addr1, d.addr2].filter(Boolean).join(" ")}`], [],
      ["TAXABLE","IMPORT ENTRY","SUPPLIER","COUNTRY OF","ASSESSMENT","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","AMOUNT OF","OR"],
      ["MONTH","NUMBER","NAME","ORIGIN","DATE","TOTAL LANDED COST","EXEMPT","TAXABLE","VAT","GROSS TAXABLE","NUMBER"],
      ["","","","","","","","","","",""],
      ["(1)","(2)","(3)","(4)","(5)","(6)","(7)","(8)","(9)","(10)","(11)"],
      ...rows, totalsRow, ["END OF REPORT"],
    ];
    colCount = 11; moneyCols = [5,6,7,8,9];
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = new Array(colCount).fill(0).map((_, i) => i === 1 ? { wch: 14 } : (i === 2 || i === 3) ? { wch: 26 } : i === 4 ? { wch: 24 } : { wch: 13 });
  const dataStartRow = 12;
  const dataRowCount = aoa.length - dataStartRow - 2;
  for (let i = 0; i < dataRowCount + 1; i++) {
    const rowIdx = dataStartRow + i;
    moneyCols.forEach((c) => {
      const ref = XLSX.utils.encode_cell({ r: rowIdx, c });
      if (ws[ref] && typeof ws[ref].v === "number") ws[ref].z = "#,##0.00";
    });
    const monthRef = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
    if (ws[monthRef] && ws[monthRef].v instanceof Date) ws[monthRef].z = "mm/dd/yyyy";
  }

  const wb = XLSX.utils.book_new();
  const sheetName = `${d.companyTin}${SLSPI_CONFIG[subMode].marker}${pad2(month)}${year}`.slice(0, 31);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `${d.companyTin}_${subMode}_Report_${pad2(month)}${year}.xlsx`);
}

async function downloadSlspiPdfReport(subMode, data, sourceRows, month, year) {
  const jsPDFCtor = await withTimeout(loadJsPDF(), 8000);
  const d = buildFilingDetails(data, { month, year, quarter: null, sawtFormType: null });

  let rows, totals, head, colStyles;
  if (subMode === "SLP") {
    rows = sourceRows.map((r) => {
      const taxableNet = r.services + r.capitalGoods + r.otherCapitalGoods;
      const totalPurchase = r.vatExempt + r.zeroRated + taxableNet;
      const grossTaxable = taxableNet + r.inputVat;
      const name = r.partyName || [r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", ");
      return { tin: r.tin ? reliefFormatTIN(r.tin) : "", nameAddr: `${name}\n${[r.address1, r.address2].filter(Boolean).join(" ")}`,
        col1: totalPurchase, col2: r.zeroRated, col3: r.vatExempt, col4: grossTaxable,
        stack: `${fmtPlain(r.services)}\n${fmtPlain(r.otherCapitalGoods)}\n${fmtPlain(r.capitalGoods)}`, col5: taxableNet, col6: r.inputVat };
    });
    totals = rows.reduce((a, r) => { a.col1+=r.col1; a.col2+=r.col2; a.col3+=r.col3; a.col4+=r.col4; a.col5+=r.col5; a.col6+=r.col6; return a; }, { col1:0,col2:0,col3:0,col4:0,col5:0,col6:0 });
    head = [["T.I.N.", "Registered Name", "Total Purchase", "Zero Rated", "Exempt", "Gross Taxable", "Services /Other\nThan Goods /\nCapital Goods", "Taxable Net", "Input Tax"]];
    colStyles = { 0:{cellWidth:65}, 1:{cellWidth:150}, 2:{cellWidth:70,halign:"right"}, 3:{cellWidth:60,halign:"right"}, 4:{cellWidth:60,halign:"right"}, 5:{cellWidth:70,halign:"right"}, 6:{cellWidth:90,halign:"right"}, 7:{cellWidth:70,halign:"right"}, 8:{cellWidth:70,halign:"right"} };
  } else if (subMode === "SLS") {
    rows = sourceRows.map((r) => {
      const totalSales = r.exemptSales + r.zeroRatedSales + r.taxableSales;
      const grossTaxable = r.taxableSales + r.outputVat;
      const name = r.partyName || [r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", ");
      return { tin: r.tin ? reliefFormatTIN(r.tin) : "", nameAddr: `${name}\n${[r.address1, r.address2].filter(Boolean).join(" ")}`,
        col1: totalSales, col2: r.zeroRatedSales, col3: r.exemptSales, col4: grossTaxable, col5: r.taxableSales, col6: r.outputVat };
    });
    totals = rows.reduce((a, r) => { a.col1+=r.col1; a.col2+=r.col2; a.col3+=r.col3; a.col4+=r.col4; a.col5+=r.col5; a.col6+=r.col6; return a; }, { col1:0,col2:0,col3:0,col4:0,col5:0,col6:0 });
    head = [["T.I.N.", "Registered Name", "Total Sales", "Zero Rated", "Exempt", "Gross Taxable", "Taxable Net", "Output Tax"]];
    colStyles = { 0:{cellWidth:70}, 1:{cellWidth:170}, 2:{cellWidth:80,halign:"right"}, 3:{cellWidth:70,halign:"right"}, 4:{cellWidth:70,halign:"right"}, 5:{cellWidth:85,halign:"right"}, 6:{cellWidth:80,halign:"right"}, 7:{cellWidth:80,halign:"right"} };
  } else {
    rows = sourceRows.map((r) => ({ tin: r.importEntryNo, nameAddr: `${r.suppliersName} - ${r.countryOrigin}`,
      col1: r.totalLandedCost + r.otherCharges, col2: r.totalLandedCost, col3: r.otherCharges, col4: r.exempt, col5: r.taxable, col6: r.vat }));
    totals = rows.reduce((a, r) => { a.col1+=r.col1; a.col2+=r.col2; a.col3+=r.col3; a.col4+=r.col4; a.col5+=r.col5; a.col6+=r.col6; return a; }, { col1:0,col2:0,col3:0,col4:0,col5:0,col6:0 });
    head = [["Entry No.", "Name of Seller /\nCountry of Origin", "Total Landed Cost", "Dutiable Value", "All Charges Before\nRelease from\nCustom's Custody", "Exempt", "Taxable Goods", "VAT Paid"]];
    colStyles = { 0:{cellWidth:60}, 1:{cellWidth:170}, 2:{cellWidth:85,halign:"right"}, 3:{cellWidth:80,halign:"right"}, 4:{cellWidth:100,halign:"right"}, 5:{cellWidth:60,halign:"right"}, 6:{cellWidth:75,halign:"right"}, 7:{cellWidth:75,halign:"right"} };
  }

  const doc = new jsPDFCtor({ orientation: "landscape", unit: "pt", format: "letter" });
  const left = 30; let y = 32;
  doc.setFont("helvetica", "normal"); doc.setFontSize(9.5);
  doc.text("Bureau of Internal Revenue", left, y); y += 12;
  doc.text("Reconciliation of Listings for Enforcement System", left, y); y += 12;
  doc.text(subMode === "SLP" ? "Summary List of Purchases" : subMode === "SLS" ? "Summary List of Sales" : "Summary List of Importation", left, y); y += 12;
  doc.text(`Report Date : ${reliefFormatLongDate(new Date())}`, left, y); y += 16;

  const boxTop = y;
  doc.setFont("helvetica", "bold");
  doc.text("T.I.N. of Owner :", left, y);
  doc.setFont("helvetica", "bolditalic");
  doc.text(reliefFormatTIN(d.companyTin), left + 95, y);
  y += 14;
  doc.setFont("helvetica", "normal");
  doc.text("Registered Name :", left, y);
  doc.text(d.displayName, left + 95, y);
  doc.text("Owners Address:", left + 430, y);
  doc.text(d.addr1, left + 520, y);
  y += 12;
  doc.text("Trade Name :", left, y);
  doc.text(d.tradeName, left + 95, y);
  doc.text(d.addr2, left + 520, y);
  y += 14;
  doc.line(left, y - 4, 782, y - 4);
  doc.text("Taxable Month", left, y + 8);
  doc.setFont("helvetica", "italic");
  doc.text(`${MONTHS[month - 1]}, ${year}`, left + 95, y + 8);
  y += 20;
  doc.line(left, y, 782, y);
  doc.line(left, boxTop - 14, left, y);
  doc.line(782, boxTop - 14, 782, y);
  doc.line(left, boxTop - 14, 782, boxTop - 14);

  const body = rows.map((r) => subMode === "SLP"
    ? [r.tin, r.nameAddr, fmtPlain(r.col1), fmtPlain(r.col2), fmtPlain(r.col3), fmtPlain(r.col4), r.stack, fmtPlain(r.col5), fmtPlain(r.col6)]
    : [r.tin, r.nameAddr, fmtPlain(r.col1), fmtPlain(r.col2), fmtPlain(r.col3), fmtPlain(r.col4), fmtPlain(r.col5), fmtPlain(r.col6)]);

  const totalsLine = (label) => {
    const base = [{ content: label, colSpan: 2, styles: { halign: "right", fontStyle: "bold" } }, fmtPlain(totals.col1), fmtPlain(totals.col2), fmtPlain(totals.col3), fmtPlain(totals.col4)];
    return subMode === "SLP" ? [...base, "", fmtPlain(totals.col5), fmtPlain(totals.col6)] : [...base, fmtPlain(totals.col5), fmtPlain(totals.col6)];
  };
  body.push(totalsLine(subMode === "SLI" ? "Total for the Taxable Month" : "Taxable Month Totals"));
  body.push(totalsLine("Report Totals"));
  if (subMode !== "SLI") body.push(totalsLine("Page 1   PageTotals"));

  doc.autoTable({
    head, body, startY: y + 6, margin: { left, right: 30 },
    styles: { font: "helvetica", fontSize: 8, cellPadding: 3, valign: "top", lineColor: [0,0,0], lineWidth: 0.3 },
    headStyles: { fillColor: false, textColor: [0,0,0], fontStyle: "bold", halign: "center", lineWidth: 0.5 },
    columnStyles: colStyles,
    didParseCell: (hookData) => {
      if (hookData.section === "body" && hookData.row.index >= rows.length) {
        hookData.cell.styles.fontStyle = "bold";
        hookData.cell.styles.lineWidth = { top: 0.7, bottom: 0.7, left: 0.3, right: 0.3 };
      }
    },
  });
  doc.save(`${d.companyTin}_${subMode}_Report_${pad2(month)}${year}.pdf`);
}

function downloadQapExcelReport(data, allRows, quarter, year) {
  const d = buildFilingDetails(data, { month: null, year, quarter, sawtFormType: null });
  const quarterEndMonth = quarter * 3;
  const groups = []; const groupIndex = {};
  allRows.forEach((r) => {
    const key = `${r.tin}|${r.branchCode}|${r.atc}`;
    if (!(key in groupIndex)) {
      groupIndex[key] = groups.length;
      groups.push({ tin: r.tin, branchCode: r.branchCode, partyName: r.partyName, surname: r.surname, firstName: r.firstName, middleName: r.middleName, atc: r.atc, rate: r.rate, m: [0,0,0], ewtM: [0,0,0] });
    }
    const g = groups[groupIndex[key]];
    const slot = qapSlotIndex(r, d);
    g.m[slot] += r.taxBase; g.ewtM[slot] += r.ewt;
    if (!g.rate) g.rate = r.rate;
  });

  const rows = groups.map((g, i) => {
    const totalIncome = g.m[0] + g.m[1] + g.m[2];
    const totalTax = g.ewtM[0] + g.ewtM[1] + g.ewtM[2];
    return [i + 1, `${reliefFormatTIN(g.tin)}-${reliefPadBranch(g.branchCode)}`, g.partyName || "", [g.surname, [g.firstName, g.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", "), g.atc, "",
      g.m[0], g.m[0] ? g.rate : "", g.ewtM[0], g.m[1], g.m[1] ? g.rate : "", g.ewtM[1], g.m[2], g.m[2] ? g.rate : "", g.ewtM[2], totalIncome, totalTax];
  });
  const totalCols = [6,8,9,11,12,14,15,16];
  const totalsRow = new Array(17).fill(""); totalsRow[0] = "Grand Total :";
  totalCols.forEach((c) => { totalsRow[c] = rows.reduce((s, r) => s + (typeof r[c] === "number" ? r[c] : 0), 0); });

  const aoa = [
    ["Attachment to BIR Form 1601-EQ"],
    ["QUARTERLY ALPHABETICAL LIST OF PAYEES SUBJECTED TO EXPANDED WITHHOLDING TAX & PAYEES WHOSE INCOME PAYMENTS ARE EXEMPT"],
    [`FOR THE QUARTER ENDING ${MONTHS[quarterEndMonth - 1].toUpperCase()}, ${year}`], [],
    [`TIN : ${d.companyTin}-${reliefPadBranch(d.branchCode)}`], [`WITHHOLDING AGENT'S NAME: ${d.displayName}`], [],
    ["","","","","","","1ST MONTH OF THE QUARTER","","","2ND MONTH OF THE QUARTER","","","3RD MONTH OF THE QUARTER","","","TOTAL FOR THE QUARTER",""],
    ["SEQ","TAXPAYER","CORPORATION","INDIVIDUAL","ATC CODE","NATURE OF PAYMENT","AMOUNT OF","TAX RATE","AMOUNT OF","AMOUNT OF","TAX RATE","AMOUNT OF","AMOUNT OF","TAX RATE","AMOUNT OF","TOTAL","TOTAL"],
    ["NO","IDENTIFICATION","(Registered Name)","(Last Name, First Name, Middle Name)","","","INCOME PAYMENT","","TAX WITHHELD","INCOME PAYMENT","","TAX WITHHELD","INCOME PAYMENT","","TAX WITHHELD","INCOME PAYMENT","TAX WITHHELD"],
    ["(1)","(2)","(3)","(4)","(5)","","(6)","(7)","(8)","(9)","(10)","(11)","(12)","(13)","(14)","(15)","(16)"],
    ...rows, totalsRow, ["END OF REPORT"],
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{wch:6},{wch:18},{wch:26},{wch:30},{wch:8},{wch:20},{wch:13},{wch:9},{wch:13},{wch:13},{wch:9},{wch:13},{wch:13},{wch:9},{wch:13},{wch:14},{wch:13}];
  const dataStartRow = 11;
  for (let i = 0; i < rows.length + 1; i++) {
    const rowIdx = dataStartRow + i;
    [6,8,9,11,12,14,15,16].forEach((c) => {
      const ref = XLSX.utils.encode_cell({ r: rowIdx, c });
      if (ws[ref] && typeof ws[ref].v === "number") ws[ref].z = "#,##0.00";
    });
  }

  const wb = XLSX.utils.book_new();
  const sheetName = `${d.companyTin}QAP${pad2(quarterEndMonth)}${year}`.slice(0, 31);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `${d.companyTin}_QAP_Report_Q${quarter}${year}.xlsx`);
}

async function downloadQapPdfReport(data, allRows, quarter, year) {
  const jsPDFCtor = await withTimeout(loadJsPDF(), 8000);
  const d = buildFilingDetails(data, { month: null, year, quarter, sawtFormType: null });
  const quarterEndMonth = quarter * 3;
  const branch3 = reliefPadBranch(d.branchCode).slice(-3);
  const byMonth = { 1: [], 2: [], 3: [] };
  allRows.forEach((r) => { byMonth[qapSlotIndex(r, d) + 1].push(r); });

  const doc = new jsPDFCtor({ orientation: "portrait", unit: "pt", format: "letter" });
  const left = 40; let y = 40;
  doc.setFont("helvetica", "bold"); doc.setFontSize(10);
  doc.text("BIR FORM 1601EQ - SCHEDULE 1", 306, y, { align: "center" });
  doc.text("PAGE   1", 560, y, { align: "right" }); y += 14;
  doc.text("ALPHABETICAL LIST OF PAYEES FROM WHOM TAXES WERE WITHHELD", 306, y, { align: "center" }); y += 14;
  doc.setFont("helvetica", "normal");
  doc.text(`FOR THE QUARTER ENDING ${MONTHS[quarterEndMonth - 1].toUpperCase()},   ${year}`, 306, y, { align: "center" }); y += 30;

  doc.setFontSize(9);
  doc.text("WITHHOLDING AGENT'S NAME :", left, y);
  doc.setFont("helvetica", "bold");
  doc.text(d.displayName, left + 170, y);
  doc.line(left + 168, y + 3, 470, y + 3);
  doc.setFont("helvetica", "normal");
  doc.text("TIN:", 490, y);
  doc.setFont("helvetica", "bold");
  doc.text(`${reliefFormatTIN(d.companyTin)}-${branch3}`, 515, y);
  doc.line(513, y + 3, 570, y + 3);
  y += 30;

  const colX = { seq: left, tin: left + 35, corp: left + 130, ind: left + 300, atc: left + 435, amt: left + 470, rate: left + 530, ewt: left + 560 };
  doc.setFont("helvetica", "bold"); doc.setFontSize(8);
  [["SEQ.","TAXPAYER","CORPORATION","INDIVIDUAL","ATC","AMOUNT OF","TAX","AMOUNT OF TAX"],
   ["NO.","IDENTIFICATION","(Registered Name)","(LastName, FirstName  MiddleName)","","INCOME","RATE","WITHHELD"],
   ["","NUMBER (TIN)","","","","PAYMENT","",""]].forEach((line, li) => {
    doc.text(line[0], colX.seq, y + li * 10); doc.text(line[1], colX.tin, y + li * 10);
    doc.text(line[2], colX.corp, y + li * 10); doc.text(line[3], colX.ind, y + li * 10);
    doc.text(line[4], colX.atc, y + li * 10);
    doc.text(line[5], colX.amt, y + li * 10, { align: "right" });
    doc.text(line[6], colX.rate, y + li * 10, { align: "right" });
    doc.text(line[7], colX.ewt, y + li * 10, { align: "right" });
  });
  y += 34;
  doc.setFont("helvetica", "normal");
  doc.text("(1)", colX.seq, y); doc.text("(2)", colX.tin, y); doc.text("(3)", colX.corp, y);
  doc.text("(4)", colX.ind, y); doc.text("(5)", colX.atc, y);
  doc.text("(6)", colX.amt, y, { align: "right" }); doc.text("(7)", colX.rate, y, { align: "right" }); doc.text("(8)", colX.ewt, y, { align: "right" });
  y += 6;
  doc.text("----", colX.seq, y); doc.text("------------", colX.tin, y);
  doc.text("------------------------------", colX.corp, y); doc.text("------------------------------", colX.ind, y);
  doc.text("----", colX.atc, y); doc.text("-----------", colX.amt, y, { align: "right" });
  doc.text("------", colX.rate, y, { align: "right" }); doc.text("-----------", colX.ewt, y, { align: "right" });
  y += 16;

  let grandTotal = 0;
  [1,2,3].forEach((m) => {
    byMonth[m].forEach((r, i) => {
      grandTotal += r.ewt;
      doc.text(String(i + 1), colX.seq, y);
      doc.text(`${reliefFormatTIN(r.tin)}-${reliefPadBranch(r.branchCode)}`, colX.tin, y);
      doc.text(r.partyName || "", colX.corp, y);
      doc.text([r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", "), colX.ind, y);
      doc.text(r.atc, colX.atc, y);
      doc.text(fmtPlain(r.taxBase), colX.amt, y, { align: "right" });
      doc.text(Number(r.rate).toFixed(2), colX.rate, y, { align: "right" });
      doc.text(fmtPlain(r.ewt), colX.ewt, y, { align: "right" });
      y += 22;
      if (y > 740) { doc.addPage(); y = 40; }
    });
  });
  y += 14;
  doc.text("---------------", colX.ewt, y, { align: "right" }); y += 12;
  doc.setFont("helvetica", "bold");
  doc.text("PAGE TOTAL", left, y);
  doc.text(fmtPlain(grandTotal), colX.ewt, y, { align: "right" });
  doc.save(`${d.companyTin}_QAP_Report_Q${quarter}${year}.pdf`);
}

function downloadSawtExcelReport(data, sourceRows, month, year, formType) {
  const d = buildFilingDetails(data, { month, year, quarter: null, sawtFormType: formType });
  const branch4 = reliefPadBranch(d.branchCode);
  const rows = sourceRows.map((r, i) => [i + 1, `${reliefFormatTIN(r.tin)}-${reliefPadBranch(r.branchCode || branch4)}`,
    r.partyName || "", [r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", "), r.atc, "", r.taxBase, r.rate, r.cwt]);
  const totalsRow = new Array(9).fill(""); totalsRow[0] = "Grand Total :";
  totalsRow[6] = rows.reduce((s, r) => s + r[6], 0);
  totalsRow[8] = rows.reduce((s, r) => s + r[8], 0);

  const aoa = [
    [`BIR FORM ${formType}`], ["SUMMARY ALPHALIST OF WITHHOLDING TAXES (SAWT)"],
    [`FOR THE MONTH OF ${MONTHS[month - 1].toUpperCase()}, ${year}`], [],
    [`TIN : ${d.companyTin}-${branch4}`], [`PAYEE'S NAME: ${d.displayName}`], [],
    ["SEQ","TAXPAYER","CORPORATION","INDIVIDUAL","ATC CODE","NATURE OF PAYMENT","AMOUNT OF","TAX RATE","AMOUNT OF"],
    ["NO","IDENTIFICATION","(Registered Name)","(Last Name, First Name, Middle Name)","","","INCOME PAYMENT","","TAX WITHHELD"],
    ["","NUMBER","","","","","","",""],
    ["(1)","(2)","(3)","(4)","(5)","","(6)","(7)","(8)"],
    ...rows, totalsRow, ["END OF REPORT"],
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{wch:6},{wch:18},{wch:28},{wch:28},{wch:9},{wch:20},{wch:14},{wch:9},{wch:14}];
  const dataStartRow = 11;
  for (let i = 0; i < rows.length + 1; i++) {
    const rowIdx = dataStartRow + i;
    [6,7,8].forEach((c) => {
      const ref = XLSX.utils.encode_cell({ r: rowIdx, c });
      if (ws[ref] && typeof ws[ref].v === "number") ws[ref].z = "#,##0.00";
    });
  }

  const wb = XLSX.utils.book_new();
  const sheetName = `${d.companyTin}SAWT${pad2(month)}${year}`.slice(0, 31);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `${d.companyTin}_SAWT_Report_${pad2(month)}${year}.xlsx`);
}

async function downloadSawtPdfReport(data, sourceRows, month, year, formType) {
  const jsPDFCtor = await withTimeout(loadJsPDF(), 8000);
  const d = buildFilingDetails(data, { month, year, quarter: null, sawtFormType: formType });
  const branch3 = reliefPadBranch(d.branchCode).slice(-3);

  const doc = new jsPDFCtor({ orientation: "portrait", unit: "pt", format: "letter" });
  const left = 40; let y = 40;
  doc.setFont("helvetica", "bold"); doc.setFontSize(10);
  doc.text(`BIR FORM ${formType}`, 306, y, { align: "center" });
  doc.text("PAGE   1", 560, y, { align: "right" }); y += 14;
  doc.text("SUMMARY ALPHALIST OF WITHHOLDING TAXES (SAWT)", 306, y, { align: "center" }); y += 14;
  doc.setFont("helvetica", "normal");
  doc.text(`FOR THE MONTH OF ${MONTHS[month - 1].toUpperCase()},   ${year}`, 306, y, { align: "center" }); y += 30;

  doc.setFontSize(9);
  doc.text("PAYEE'S NAME", left, y);
  doc.setFont("helvetica", "bold");
  doc.text(d.displayName, left + 90, y);
  doc.line(left + 88, y + 3, 430, y + 3);
  doc.setFont("helvetica", "normal");
  doc.text("TIN :", 460, y);
  doc.setFont("helvetica", "bold");
  doc.text(`${reliefFormatTIN(d.companyTin)}-${branch3}`, 490, y);
  doc.line(488, y + 3, 560, y + 3);
  y += 30;

  const colX = { seq: left, tin: left + 35, corp: left + 130, ind: left + 300, atc: left + 435, amt: left + 490, rate: left + 535, ewt: left + 570 };
  doc.setFont("helvetica", "bold"); doc.setFontSize(8);
  [["SEQ.","TAXPAYER","CORPORATION","INDIVIDUAL","ATC","AMOUNT OF","TAX","AMOUNT OF TAX"],
   ["NO.","IDENTIFICATION","(Registered Name)","(LastName, FirstName  MiddleName)","","INCOME","RATE","WITHHELD"],
   ["","NUMBER (TIN)","","","","PAYMENT","",""]].forEach((line, li) => {
    doc.text(line[0], colX.seq, y + li * 10); doc.text(line[1], colX.tin, y + li * 10);
    doc.text(line[2], colX.corp, y + li * 10); doc.text(line[3], colX.ind, y + li * 10);
    doc.text(line[4], colX.atc, y + li * 10);
    doc.text(line[5], colX.amt, y + li * 10, { align: "right" });
    doc.text(line[6], colX.rate, y + li * 10, { align: "right" });
    doc.text(line[7], colX.ewt, y + li * 10, { align: "right" });
  });
  y += 34;
  doc.setFont("helvetica", "normal");
  doc.text("(1)", colX.seq, y); doc.text("(2)", colX.tin, y); doc.text("(3)", colX.corp, y);
  doc.text("(4)", colX.ind, y); doc.text("(5)", colX.atc, y);
  doc.text("(6)", colX.amt, y, { align: "right" }); doc.text("(7)", colX.rate, y, { align: "right" }); doc.text("(8)", colX.ewt, y, { align: "right" });
  y += 6;
  doc.text("----", colX.seq, y); doc.text("------------", colX.tin, y);
  doc.text("------------------------------", colX.corp, y); doc.text("------------------------------", colX.ind, y);
  doc.text("----", colX.atc, y); doc.text("-----------", colX.amt, y, { align: "right" });
  doc.text("------", colX.rate, y, { align: "right" }); doc.text("-----------", colX.ewt, y, { align: "right" });
  y += 16;

  let grandTotal = 0;
  sourceRows.forEach((r, i) => {
    grandTotal += r.cwt;
    doc.text(String(i + 1), colX.seq, y);
    doc.text(`${reliefFormatTIN(r.tin)}-${reliefPadBranch(r.branchCode)}`, colX.tin, y);
    if (r.partyName) doc.text(r.partyName, colX.corp, y);
    doc.text(r.atc, colX.atc, y);
    doc.text(Number(r.rate).toFixed(2), colX.rate, y, { align: "right" });
    y += 12;
    const indName = [r.surname, [r.firstName, r.middleName].filter(Boolean).join(" ")].filter(Boolean).join(", ");
    if (indName) doc.text(indName, colX.ind, y);
    doc.text(fmtPlain(r.taxBase), colX.amt, y, { align: "right" });
    doc.text(fmtPlain(r.cwt), colX.ewt, y, { align: "right" });
    y += 16;
    if (y > 740) { doc.addPage(); y = 40; }
  });
  y += 10;
  doc.text("---------------", colX.ewt, y, { align: "right" }); y += 12;
  doc.setFont("helvetica", "bold");
  doc.text("PAGE TOTAL", left, y);
  doc.text(fmtPlain(grandTotal), colX.ewt, y, { align: "right" });
  doc.save(`${d.companyTin}_SAWT_Report_${pad2(month)}${year}.pdf`);
}

function ReliefCompanyWarning({ filingDetails, needAddr }) {
  const missing = !filingDetails.companyTin || !filingDetails.rdoCode || (needAddr && !filingDetails.addr1);
  if (!missing) return null;
  return (
    <div className="callout">
      <Info size={16} />
      <div>
        <div className="callout-title">Company details incomplete</div>
        <div className="callout-body">Fill in TIN, RDO Code{needAddr ? ", and Address Line 1" : ""} on the Company Details page before generating a .dat file.</div>
      </div>
    </div>
  );
}

/* ============================== IMPORTATION LEDGER ============================== */

function ImportationEntryModal({ data, onCancel, onSubmit }) {
  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const suppOptions = data.suppliers.map((s) => s.name).filter(Boolean);
  const [v, setV] = useState({
    date: todayMDY(), importEntryNo: "", refNo: "", supplier: "", countryOrigin: "", desc: "",
    dutiableValue: 0, customsDuty: 0, otherCharges: 0, vatRate: 0.12, orNo: "", coaCode: "1300",
    assessmentDate: todayMDY(), vatPaymentDate: "",
  });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const computed = computeImportationRow(v);
  return (
    <EntryModalShell title="Add Importation" submitLabel="Add importation" onCancel={onCancel}
      onSubmit={() => onSubmit(computeImportationRow({ id: uid(), ...v }))}>
      <LabeledField label="Date of Importation"><Field type="date" value={v.date} onChange={(x) => set("date", x)} /></LabeledField>
      <LabeledField label="Import Entry No. (BOC)"><Field value={v.importEntryNo} onChange={(x) => set("importEntryNo", x)} placeholder="IE-2026-000123" /></LabeledField>
      <LabeledField label="Reference No. (BL/AWB)"><Field value={v.refNo} onChange={(x) => set("refNo", x)} /></LabeledField>
      <LabeledField label="Foreign Supplier / Seller"><Field type="combo" options={suppOptions} freeText value={v.supplier} onChange={(x) => set("supplier", x)} /></LabeledField>
      <LabeledField label="Country of Origin"><Field value={v.countryOrigin} onChange={(x) => set("countryOrigin", x)} /></LabeledField>
      <LabeledField label="Description of Goods" wide><Field value={v.desc} onChange={(x) => set("desc", x)} /></LabeledField>
      <LabeledField label="Dutiable / Assessed Value"><Field type="number" align="right" value={v.dutiableValue} onChange={(x) => set("dutiableValue", x)} /></LabeledField>
      <LabeledField label="Customs Duty"><Field type="number" align="right" value={v.customsDuty} onChange={(x) => set("customsDuty", x)} /></LabeledField>
      <LabeledField label="Other BOC Charges"><Field type="number" align="right" value={v.otherCharges} onChange={(x) => set("otherCharges", x)} /></LabeledField>
      <LabeledField label="VAT Rate"><Field type="percent" value={v.vatRate} onChange={(x) => set("vatRate", x)} /></LabeledField>
      <LabeledField label="BOC OR No."><Field value={v.orNo} onChange={(x) => set("orNo", x)} /></LabeledField>
      <LabeledField label="Account"><Field type="combo" options={acctOptions} value={v.coaCode} onChange={(x) => set("coaCode", x)} /></LabeledField>
      <LabeledField label="Assessment Date"><Field type="date" value={v.assessmentDate} onChange={(x) => set("assessmentDate", x)} /></LabeledField>
      <LabeledField label="VAT Payment Date"><Field type="date" value={v.vatPaymentDate} onChange={(x) => set("vatPaymentDate", x)} /></LabeledField>
      <ComputedPreview items={[["Landed Cost", fmt(computed.landedCost)], ["VAT Paid on Importation", fmt(computed.vatPaid)]]} />
    </EntryModalShell>
  );
}

function ImportationLedgerPage({ data, setData }) {
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const acctOptions = data.coa.map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const suppOptions = data.suppliers.map((s) => s.name).filter(Boolean);
  const [modalOpen, setModalOpen] = useState(false);
  const [filter, setFilter] = usePeriodFilter();
  const [search, setSearch] = useState("");
  const sel = useRowSelection();

  const update = (id, patch) => setData((d) => ({ ...d, importation: d.importation.map((r) => r.id === id ? computeImportationRow({ ...r, ...patch }) : r) }));
  const onDelete = (id) => setData((d) => ({ ...d, importation: d.importation.filter((r) => r.id !== id) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, importation: d.importation.filter((r) => !sel.selected.has(r.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };

  const filteredRows = useMemo(() => (data.importation || []).filter((r) => matchesPeriod(r.date, filter) && matchesSearch(r, ["importEntryNo","refNo","supplier","countryOrigin","desc","orNo"], search)), [data.importation, filter, search]);
  const totals = filteredRows.reduce((a, r) => ({ dutiableValue: a.dutiableValue + num(r.dutiableValue), customsDuty: a.customsDuty + num(r.customsDuty), otherCharges: a.otherCharges + num(r.otherCharges), landedCost: a.landedCost + num(r.landedCost), vatPaid: a.vatPaid + num(r.vatPaid) }), { dutiableValue: 0, customsDuty: 0, otherCharges: 0, landedCost: 0, vatPaid: 0 });
  const emptyMsg = (data.importation || []).length === 0 ? "No importations logged yet — add your first row above." : "No importations match this filter.";

  return (
    <div>
      <SectionHeader icon={ArrowDownToLine} title="Importation Ledger" subtitle="Every import transaction — landed cost and VAT paid on importation compute automatically. Feeds the SLI tab." />
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
            <SearchBar value={search} onChange={setSearch} placeholder="Search entry no., supplier, country, description…" />
          </div>
          <AddRowBtn onClick={() => setModalOpen(true)}>Add importation</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead>
              <tr>
                <th style={{width:32}}><SelectAllCheckbox ids={filteredRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
                <th style={{minWidth:110}}>Date</th><th style={{minWidth:130}}>Entry No.</th><th style={{minWidth:120}}>Ref No.</th>
                <th style={{minWidth:190}}>Foreign Supplier / Seller</th><th style={{minWidth:110}}>Country</th><th style={{minWidth:160}}>Description</th>
                <th style={{minWidth:110}} className="num-head">Dutiable Value</th><th style={{minWidth:100}} className="num-head">Customs Duty</th>
                <th style={{minWidth:100}} className="num-head">Other Charges</th><th style={{minWidth:110}} className="num-head">Landed Cost</th>
                <th style={{minWidth:100}} className="num-head">VAT Paid</th><th style={{minWidth:100}}>OR No.</th><th style={{minWidth:190}}>Account</th>
                <th style={{width:36}}></th>
              </tr>
            </thead>
            <tbody>
              {filteredRows.length === 0 && <tr><td colSpan={15} className="empty-row">{emptyMsg}</td></tr>}
              {filteredRows.map((r) => (
                <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><Field type="date" value={r.date} onChange={(v) => update(r.id, { date: v })} /></td>
                  <td><Field value={r.importEntryNo} onChange={(v) => update(r.id, { importEntryNo: v })} /></td>
                  <td><Field value={r.refNo} onChange={(v) => update(r.id, { refNo: v })} /></td>
                  <td><Field type="combo" options={suppOptions} freeText value={r.supplier} onChange={(v) => update(r.id, { supplier: v })} /></td>
                  <td><Field value={r.countryOrigin} onChange={(v) => update(r.id, { countryOrigin: v })} /></td>
                  <td><Field value={r.desc} onChange={(v) => update(r.id, { desc: v })} /></td>
                  <td><Field type="number" align="right" value={r.dutiableValue} onChange={(v) => update(r.id, { dutiableValue: v })} /></td>
                  <td><Field type="number" align="right" value={r.customsDuty} onChange={(v) => update(r.id, { customsDuty: v })} /></td>
                  <td><Field type="number" align="right" value={r.otherCharges} onChange={(v) => update(r.id, { otherCharges: v })} /></td>
                  <td><ReadCell align="right">{fmt(r.landedCost)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.vatPaid)}</ReadCell></td>
                  <td><Field value={r.orNo} onChange={(v) => update(r.id, { orNo: v })} /></td>
                  <td><Field type="combo" options={acctOptions} value={r.coaCode} onChange={(v) => update(r.id, { coaCode: v })} /></td>
                  <td className="text-center"><DelBtn onClick={() => onDelete(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {filteredRows.length > 0 && (
              <tfoot><tr>
                <td colSpan={7} className="totals-label">Totals</td>
                <td className="num">{fmt(totals.dutiableValue)}</td><td className="num">{fmt(totals.customsDuty)}</td>
                <td className="num">{fmt(totals.otherCharges)}</td><td className="num">{fmt(totals.landedCost)}</td>
                <td className="num">{fmt(totals.vatPaid)}</td><td colSpan={3}></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
      {modalOpen && (
        <ImportationEntryModal data={data} onCancel={() => setModalOpen(false)}
          onSubmit={(row) => { setData((d) => ({ ...d, importation: [...(d.importation || []), row] })); setModalOpen(false); }} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected importation${sel.selected.size === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
    </div>
  );
}

/* ============================== 2307 LOG ============================== */
// A manually-maintained log of BIR Form 2307 certificates. Customer identity fields snapshot from
// Customers Master at the moment a TIN is entered/changed (same pattern as Purchase Journal's
// supplier auto-fill) — not a live lookup — so the log stays self-contained for Excel export/import
// and doesn't silently change if Customers Master is edited later.

function Form2307EntryModal({ data, onCancel, onSubmit }) {
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const [v, setV] = useState({ date: todayMDY(), tin: "", customerName: "", surname: "", firstName: "", middleName: "", custType: "", atc: "", atcRate: 0, incomePayment: 0 });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const onTin = (tin) => {
    const cust = data.customers.find((c) => normalizeTin(c.tin) && normalizeTin(c.tin) === normalizeTin(tin));
    setV((p) => ({
      ...p, tin,
      customerName: cust ? (cust.registeredName || cust.name || "") : "",
      surname: cust ? (cust.surname || "") : "",
      firstName: cust ? (cust.firstName || "") : "",
      middleName: cust ? (cust.middleName || "") : "",
      custType: cust ? (cust.type || "") : "",
    }));
  };
  const onAtc = (code) => {
    const atc = data.atc.find((a) => a.code === code);
    setV((p) => ({ ...p, atc: code, atcRate: atc ? num(atc.rate) : 0 }));
  };
  const computed = compute2307Row(v);
  const custFound = !!data.customers.find((c) => normalizeTin(c.tin) && normalizeTin(c.tin) === normalizeTin(v.tin));
  return (
    <EntryModalShell title="Add CWT Entry" submitLabel="Add entry" onCancel={onCancel}
      onSubmit={() => onSubmit(compute2307Row({ id: uid(), ...v }))}>
      <LabeledField label="Month"><Field type="date" value={v.date} onChange={(x) => set("date", x)} /></LabeledField>
      <LabeledField label="TIN"><Field value={v.tin} onChange={onTin} placeholder="000-000-000-000" /></LabeledField>
      <LabeledField label="Customer's Name" wide><ReadCell>{v.tin ? (custFound ? (v.customerName || "—") : "No matching customer in Customers Master") : "Enter a TIN above"}</ReadCell></LabeledField>
      <LabeledField label="Surname"><ReadCell>{v.surname || "—"}</ReadCell></LabeledField>
      <LabeledField label="First Name"><ReadCell>{v.firstName || "—"}</ReadCell></LabeledField>
      <LabeledField label="Middle Name"><ReadCell>{v.middleName || "—"}</ReadCell></LabeledField>
      <LabeledField label="ATC"><Field type="combo" options={atcOptions} value={v.atc} onChange={onAtc} /></LabeledField>
      <LabeledField label="Rate"><ReadCell align="center">{v.atc ? `${round2(v.atcRate * 100)}%` : "—"}</ReadCell></LabeledField>
      <LabeledField label="Income Payment"><Field type="number" align="right" value={v.incomePayment} onChange={(x) => set("incomePayment", x)} /></LabeledField>
      <ComputedPreview items={[["CWT", fmt(computed.cwt)]]} />
      <div className="qf-hint">Customer's Name, Surname, First Name and Middle Name auto-fill from Customers Master the moment a matching TIN is entered — add the customer there first if they're not found.</div>
    </EntryModalShell>
  );
}

function Form2307Page({ data, setData }) {
  const atcOptions = data.atc.map((a) => ({ value: a.code, label: `${a.code} — ${a.desc} (${round2(num(a.rate) * 100)}%)` }));
  const importHook = useJournalImport("form2307", data, ({ newRows }) => setData((d) => ({ ...d, form2307: [...(d.form2307 || []), ...newRows] })));
  const [modalOpen, setModalOpen] = useState(false);
  const [filter, setFilter] = usePeriodFilter();
  const [search, setSearch] = useState("");
  const sel = useRowSelection();

  const update = (id, patch) => setData((d) => ({ ...d, form2307: d.form2307.map((r) => r.id === id ? compute2307Row({ ...r, ...patch }) : r) }));
  const onTin = (id, tin) => {
    const cust = data.customers.find((c) => normalizeTin(c.tin) && normalizeTin(c.tin) === normalizeTin(tin));
    update(id, {
      tin,
      customerName: cust ? (cust.registeredName || cust.name || "") : "",
      surname: cust ? (cust.surname || "") : "",
      firstName: cust ? (cust.firstName || "") : "",
      middleName: cust ? (cust.middleName || "") : "",
      custType: cust ? (cust.type || "") : "",
    });
  };
  const [refreshStatus, setRefreshStatus] = useState(null);
  const onRefreshFromCustomers = () => {
    let matched = 0, unmatched = 0;
    setData((d) => ({
      ...d,
      form2307: (d.form2307 || []).map((r) => {
        const cust = data.customers.find((c) => normalizeTin(c.tin) && normalizeTin(c.tin) === normalizeTin(r.tin));
        if (cust) matched++; else if (r.tin) unmatched++;
        return cust ? compute2307Row({
          ...r,
          customerName: cust.registeredName || cust.name || "",
          surname: cust.surname || "",
          firstName: cust.firstName || "",
          middleName: cust.middleName || "",
          custType: cust.type || "",
        }) : r;
      }),
    }));
    setRefreshStatus({ type: "success", text: `Re-matched ${matched} entr${matched === 1 ? "y" : "ies"} against Customers Master.${unmatched ? ` ${unmatched} still have no matching TIN.` : ""}` });
    setTimeout(() => setRefreshStatus(null), 6000);
  };
  const onAtc = (id, code) => {
    const atc = data.atc.find((a) => a.code === code);
    update(id, { atc: code, atcRate: atc ? num(atc.rate) : 0 });
  };
  const onAdd = () => setModalOpen(true);
  const onDelete = (id) => setData((d) => ({ ...d, form2307: d.form2307.filter((r) => r.id !== id) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, form2307: d.form2307.filter((r) => !sel.selected.has(r.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };

  const filteredRows = useMemo(() => (data.form2307 || []).filter((r) => matchesPeriod(r.date, filter) && matchesSearch(r, ["tin","customerName","surname","firstName","middleName","atc"], search)), [data.form2307, filter, search]);
  const totals = filteredRows.reduce((a, r) => ({ incomePayment: a.incomePayment + num(r.incomePayment), cwt: a.cwt + num(r.cwt) }), { incomePayment: 0, cwt: 0 });
  const emptyMsg = (data.form2307 || []).length === 0 ? "No entries logged yet — add your first row above." : "No entries match this filter.";
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const exportHook = useJournalExport("form2307", data, filteredRows, coaByCode);

  return (
    <div>
      <SectionHeader icon={FileText} title="Creditable Withholding Taxes" subtitle="BIR Form 2307 log — Customer's Name and individual name parts auto-fill from Customers Master by TIN; Rate auto-fills from the selected ATC."
        right={<div className="header-actions">
          <button className="io-btn" onClick={onRefreshFromCustomers} title="Re-match every row's TIN against the current Customers Master — useful if a customer's name was added or fixed after these rows were entered">
            <RefreshCw size={13} /> Refresh from Customers Master
          </button>
          <ImportExportBar journalKey="form2307" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} />
        </div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      <ExportStatus status={refreshStatus} />
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <PeriodFilterBar filter={filter} setFilter={setFilter} years={getAvailableYears(data)} />
            <SearchBar value={search} onChange={setSearch} placeholder="Search TIN, customer, ATC…" />
          </div>
          <AddRowBtn onClick={onAdd}>Add CWT entry</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr>
              <th style={{width:32}}><SelectAllCheckbox ids={filteredRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
              <th style={{minWidth:110}}>Month</th><th style={{minWidth:130}}>TIN</th><th style={{minWidth:190}}>Customer's Name</th>
              <th style={{minWidth:140}}>Surname</th><th style={{minWidth:130}}>First Name</th><th style={{minWidth:130}}>Middle Name</th>
              <th style={{minWidth:180}}>ATC</th><th style={{minWidth:80}} className="num-head">Rate</th>
              <th style={{minWidth:130}} className="num-head">Income Payment</th><th style={{minWidth:110}} className="num-head">CWT</th>
              <th style={{width:36}}></th>
            </tr></thead>
            <tbody>
              {filteredRows.length === 0 && <tr><td colSpan={12} className="empty-row">{emptyMsg}</td></tr>}
              {filteredRows.map((r) => (
                <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><Field type="date" value={r.date} onChange={(v) => update(r.id, { date: v })} /></td>
                  <td><Field value={r.tin} onChange={(v) => onTin(r.id, v)} placeholder="000-000-000-000" /></td>
                  <td><ReadCell>{r.tin ? (r.custType ? (r.customerName || "—") : <span style={{ color: "var(--red)" }}>No match in Customers Master</span>) : "—"}</ReadCell></td>
                  <td><ReadCell>{r.surname || "—"}</ReadCell></td>
                  <td><ReadCell>{r.firstName || "—"}</ReadCell></td>
                  <td><ReadCell>{r.middleName || "—"}</ReadCell></td>
                  <td><Field type="combo" options={atcOptions} value={r.atc} onChange={(v) => onAtc(r.id, v)} /></td>
                  <td><ReadCell align="center">{r.atc ? `${round2(num(r.atcRate) * 100)}%` : "—"}</ReadCell></td>
                  <td><Field type="number" align="right" value={r.incomePayment} onChange={(v) => update(r.id, { incomePayment: v })} /></td>
                  <td><ReadCell align="right">{fmt(r.cwt)}</ReadCell></td>
                  <td className="text-center"><DelBtn onClick={() => onDelete(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {filteredRows.length > 0 && (
              <tfoot><tr>
                <td colSpan={9} className="totals-label">Totals</td>
                <td className="num">{fmt(totals.incomePayment)}</td><td className="num">{fmt(totals.cwt)}</td><td></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
      {modalOpen && (
        <Form2307EntryModal data={data} onCancel={() => setModalOpen(false)}
          onSubmit={(row) => { setData((d) => ({ ...d, form2307: [...(d.form2307 || []), row] })); setModalOpen(false); }} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected entr${sel.selected.size === 1 ? "y" : "ies"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
    </div>
  );
}

/* ============================== ALPHALIST OF EMPLOYEES ============================== */
// BIR Form 1604-C's two schedules — Schedule 1 (regular employees) and Schedule 2 (statutory
// minimum wage earners) — plus BIR Form 2316 certificates. Field names below carry the official
// line-number references in comments so the mapping stays traceable back to the actual form.
// Unlike SLSPI/QAP/SAWT — ported from a tool whose record layouts were confirmed against real,
// validated BIR files — this DAT layout follows the general published Alphalist structure and a
// real sample workbook, but hasn't been independently confirmed the same way. Always run it
// through the BIR Alphalist Data Entry & Validation Module before actual submission.

const EMP_STATUS_OPTIONS = [
  { value: "R", label: "R — Regular" },
  { value: "CP", label: "CP — Casual/Probationary" },
  { value: "C", label: "C — Casual" },
  { value: "P", label: "P — Probationary" },
  { value: "S", label: "S — Seasonal" },
];
const SEP_REASON_OPTIONS = [
  { value: "NA", label: "NA — Not applicable" },
  { value: "T", label: "T — Terminated" },
  { value: "R", label: "R — Resigned" },
  { value: "D", label: "D — Died" },
  { value: "RT", label: "RT — Retired" },
];

function buildAlphalistDatRecord(filingDetails, rows, year, scheduleNo) {
  const d = filingDetails;
  const companyTin = reliefPadTin(d.companyTin);
  const branch = reliefPadBranch(d.branchCode);
  const q = (s) => `"${String(s ?? "").replace(/"/g, "'")}"`;
  const amt = (n) => round2(num(n)).toFixed(2);
  const scheduleLabel = scheduleNo === 2 ? "ALPHALIST OF MINIMUM WAGE EARNERS" : "ALPHALIST OF EMPLOYEES";

  const hRow = [q("H"), q(companyTin), q(branch), q(d.displayName), q(d.addr1), q(String(year)), q(scheduleLabel)];
  const dRows = rows.map((r, i) => {
    const base = [
      q("D"), q(String(i + 1)), q(reliefPadTin(normalizeTin(r.tin))),
      q(r.lastName || ""), q(r.firstName || ""), q(r.middleName || ""),
      q(r.nationality || ""), q(r.empStatus || ""), q(toMDY(r.empFrom) || ""), q(toMDY(r.empTo) || ""), q(r.sepReason || "NA"),
    ];
    const totals = [amt(r.pGross), amt(r.pTotalNonTax), amt(r.pTotalTax), amt(r.prevGross), amt(r.prevTotalNonTax), amt(r.prevTotalTax),
      amt(r.totalTaxableComp), amt(r.taxDue), amt(r.taxWithheldPrev), amt(r.taxWithheldPresent), amt(r.taxCredit5pct),
      amt(r.amtWithheldDec), amt(r.amtRefunded), q(r.substitutedFiling || "")];
    if (scheduleNo === 2) {
      return [...base, amt(r.smwPerDay), amt(r.smwPerMonth), amt(r.smwPerYear), amt(r.factorDaysPerYear),
        amt(r.mweBasicPay), amt(r.mweHolidayPay), amt(r.mweOvertimePay), amt(r.mweNightDiff), amt(r.mweHazardPay), ...totals];
    }
    return [...base, ...totals];
  });
  const totals = rows.reduce((a, r) => ({
    gross: a.gross + num(r.pGross), nonTax: a.nonTax + num(r.pTotalNonTax), tax: a.tax + num(r.pTotalTax),
    taxableComp: a.taxableComp + num(r.totalTaxableComp), due: a.due + num(r.taxDue),
    wh: a.wh + num(r.taxWithheldPrev) + num(r.taxWithheldPresent),
  }), { gross: 0, nonTax: 0, tax: 0, taxableComp: 0, due: 0, wh: 0 });
  const cRow = [q("C"), q(String(rows.length)), amt(totals.gross), amt(totals.nonTax), amt(totals.tax), amt(totals.taxableComp), amt(totals.due), amt(totals.wh)];

  return { hRow, dRows, cRow, totals, filename: `${companyTin}${branch}ALPHALIST_SCHED${scheduleNo}_${year}.dat` };
}

async function downloadAlphalistDat(record) {
  const lines = [record.hRow.join(","), ...record.dRows.map((r) => r.join(",")), record.cRow.join(",")];
  const blob = new Blob([lines.join("\r\n")], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = record.filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

async function build2316Pdf(jsPDFCtor, data, filingDetails, employee, year) {
  const doc = new jsPDFCtor({ orientation: "portrait", unit: "pt", format: "letter" });
  const left = 42; const right = 570; let y = 36;
  const fullName = [employee.lastName, employee.firstName, employee.middleName].filter(Boolean).join(", ");
  const mwe = employee.isMWE === "Y";

  doc.setFont("helvetica", "bold"); doc.setFontSize(11);
  doc.text("BIR FORM 2316", 306, y, { align: "center" }); y += 13;
  doc.setFontSize(8.5); doc.setFont("helvetica", "normal");
  doc.text("Certificate of Compensation Payment/Tax Withheld For Compensation Payment With or Without Tax Withheld", 306, y, { align: "center", maxWidth: 480 }); y += 20;
  doc.setFont("helvetica", "bold"); doc.text(`For the Year ${year}`, 306, y, { align: "center" }); y += 18;

  doc.setFont("helvetica", "bold"); doc.setFontSize(9); doc.text("Part I — Employee Information", left, y); y += 13;
  doc.setFont("helvetica", "normal"); doc.setFontSize(8.5);
  doc.text(`TIN: ${reliefFormatTIN(normalizeTin(employee.tin))}`, left, y);
  doc.text(`Statutory Minimum Wage Earner: ${mwe ? "Yes" : "No"}`, 320, y); y += 12;
  doc.text(`Name: ${fullName}`, left, y); y += 12;
  doc.text(`Registered Address: ${employee.address || ""}${employee.zipCode ? "  " + employee.zipCode : ""}`, left, y); y += 12;
  if (employee.nationality) { doc.text(`Nationality: ${employee.nationality}`, left, y); y += 12; }
  doc.text(`Employment Status: ${employee.empStatus || ""}   Period: ${employee.empFrom || ""} to ${employee.empTo || ""}`, left, y); y += 16;

  doc.setFont("helvetica", "bold"); doc.setFontSize(9); doc.text("Part II — Employer Information", left, y); y += 13;
  doc.setFont("helvetica", "normal"); doc.setFontSize(8.5);
  doc.text(`TIN: ${reliefFormatTIN(filingDetails.companyTin)}`, left, y);
  doc.text(`RDO Code: ${filingDetails.rdoCode}`, 320, y); y += 12;
  doc.text(`Employer's Name: ${filingDetails.displayName}`, left, y); y += 12;
  doc.text(`Registered Address: ${filingDetails.addr1}`, left, y); y += 16;

  const nonTaxBody = mwe
    ? [["Basic/SMW Pay", fmtPlain(employee.mweBasicPay)], ["Holiday Pay", fmtPlain(employee.mweHolidayPay)],
       ["Overtime Pay", fmtPlain(employee.mweOvertimePay)], ["Night Shift Differential", fmtPlain(employee.mweNightDiff)],
       ["Hazard Pay", fmtPlain(employee.mweHazardPay)], ["13th Month Pay & Other Benefits", fmtPlain(employee.p13thMonthNonTax)],
       ["De Minimis Benefits", fmtPlain(employee.pDeMinimis)], ["SSS/GSIS/PHIC/HDMF & Union Dues", fmtPlain(employee.pContributions)],
       ["Salaries (₱250K & below) & Other Forms", fmtPlain(employee.pSalariesNonTax)]]
    : [["13th Month Pay & Other Benefits", fmtPlain(employee.p13thMonthNonTax)], ["De Minimis Benefits", fmtPlain(employee.pDeMinimis)],
       ["SSS/GSIS/PHIC/HDMF & Union Dues", fmtPlain(employee.pContributions)], ["Salaries (₱250K & below) & Other Forms", fmtPlain(employee.pSalariesNonTax)]];

  doc.autoTable({
    startY: y, margin: { left, right: 612 - right },
    head: [["Present Employer — Non-Taxable/Exempt Compensation", "Amount"]],
    body: [...nonTaxBody, [{ content: "Total Non-Taxable", styles: { fontStyle: "bold" } }, { content: fmtPlain(employee.pTotalNonTax), styles: { fontStyle: "bold" } }]],
    styles: { font: "helvetica", fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [230, 236, 245], textColor: [20, 30, 50], fontStyle: "bold", fontSize: 8 },
    columnStyles: { 0: { cellWidth: 350 }, 1: { cellWidth: 128, halign: "right" } },
  });
  y = doc.lastAutoTable.finalY + 10;

  doc.autoTable({
    startY: y, margin: { left, right: 612 - right },
    head: [["Present Employer — Taxable Compensation", "Amount"]],
    body: [
      ["Basic Salary (net of contributions)", fmtPlain(employee.pBasicSalary)],
      ["13th Month Pay & Other Benefits (excess)", fmtPlain(employee.p13thMonthExcess)],
      ["Salaries & Other Forms of Compensation", fmtPlain(employee.pSalariesTax)],
      [{ content: "Total Taxable", styles: { fontStyle: "bold" } }, { content: fmtPlain(employee.pTotalTax), styles: { fontStyle: "bold" } }],
      [{ content: "Gross Compensation Income (Present Employer)", styles: { fontStyle: "bold" } }, { content: fmtPlain(employee.pGross), styles: { fontStyle: "bold" } }],
    ],
    styles: { font: "helvetica", fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [230, 236, 245], textColor: [20, 30, 50], fontStyle: "bold", fontSize: 8 },
    columnStyles: { 0: { cellWidth: 350 }, 1: { cellWidth: 128, halign: "right" } },
  });
  y = doc.lastAutoTable.finalY + 10;

  if (y > 620) { doc.addPage(); y = 40; }
  doc.autoTable({
    startY: y, margin: { left, right: 612 - right },
    head: [["Part III — Tax Summary (Present + Previous Employer)", "Amount"]],
    body: [
      ["Total Taxable Compensation Income", fmtPlain(employee.totalTaxableComp)],
      ["Tax Due", fmtPlain(employee.taxDue)],
      ["Tax Withheld — Previous Employer", fmtPlain(employee.taxWithheldPrev)],
      ["Tax Withheld — Present Employer", fmtPlain(employee.taxWithheldPresent)],
      ["5% Tax Credit (PERA Act of 2008)", fmtPlain(employee.taxCredit5pct)],
      [{ content: "Amount Withheld & Paid For in December / Last Salary", styles: { fontStyle: "bold" } }, { content: fmtPlain(employee.amtWithheldDec), styles: { fontStyle: "bold" } }],
      ["Amount Refunded to Employee (if over-withheld)", fmtPlain(employee.amtRefunded)],
      ["Substituted Filing", employee.substitutedFiling === "Y" ? "Yes" : (employee.substitutedFiling === "N" ? "No" : "")],
    ],
    styles: { font: "helvetica", fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [230, 236, 245], textColor: [20, 30, 50], fontStyle: "bold", fontSize: 8 },
    columnStyles: { 0: { cellWidth: 350 }, 1: { cellWidth: 128, halign: "right" } },
  });
  y = doc.lastAutoTable.finalY + 34;

  if (y > 700) { doc.addPage(); y = 40; }
  doc.setFont("helvetica", "normal"); doc.setFontSize(7.5);
  doc.text("This certifies that the above-named employee received the compensation stated, and that the tax indicated has been withheld and remitted to the Bureau of Internal Revenue, in accordance with the Tax Code, as amended.", left, y, { maxWidth: right - left });
  y += 46;
  doc.line(left, y, left + 220, y);
  doc.text("Employer's Representative / Date", left, y + 11);
  return doc;
}

async function downloadOne2316(data, filingDetails, employee, year) {
  const jsPDFCtor = await withTimeout(loadJsPDF(), 8000);
  const doc = await build2316Pdf(jsPDFCtor, data, filingDetails, employee, year);
  const name = [employee.lastName, employee.firstName].filter(Boolean).join("_").replace(/\s+/g, "");
  doc.save(`2316_${name || normalizeTin(employee.tin)}_${year}.pdf`);
}

async function downloadAll2316Zip(data, filingDetails, rows, year) {
  const [jsPDFCtor, JSZipCtor] = await Promise.all([withTimeout(loadJsPDF(), 8000), withTimeout(loadJSZip(), 8000)]);
  const zip = new JSZipCtor();
  for (const employee of rows) {
    const doc = await build2316Pdf(jsPDFCtor, data, filingDetails, employee, year);
    const name = [employee.lastName, employee.firstName].filter(Boolean).join("_").replace(/\s+/g, "") || normalizeTin(employee.tin);
    zip.file(`2316_${name}_${year}.pdf`, doc.output("blob"));
  }
  const blob = await withTimeout(zip.generateAsync({ type: "blob" }), 8000);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `2316_Certificates_${year}.zip`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

/* ============================== BIR FORM 2307 GENERATION ============================== */
// Certificate-drawing logic adapted from the verified 2307-generation.jsx reference. The layout
// math and jsPDF calls are kept verbatim (verified against real BIR Form 2307 samples); only the
// data plumbing is wired to this app's real Purchase Journal / Cash Disbursements rows, ATC
// Reference, Suppliers Master and buildFilingDetails().

const BIR_SEAL_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAHoAAABsCAMAAAHlD2bGAAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOjZjMjYwY2E2LTE3ZjktNDQzMS1hMzExLTdjNGY1NTdiNGM2YwAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNok2NcCf3gmrKkje+a8N23egAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6OWZhZjYwZDAtN2VkNy00NGVlLTljMTctMjg3Mzk4M2QxZTAzbHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaO+yYOxOOZjhcSs+dFXwho4AAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIOKHDecwq45K+C1Au2pt5CKG0rvoURiWrdncp9BvwKW7pGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2gnkQspUl3SGqPx41NhHFN5AAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFgg4HAOMMylyn3L3J3P9tNQxo9DV/+iFjLea/CH6FDNHdxkbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYTo2YzI2MGNhNi0xN2Y5LTQ0MzEtYTMxMS03YzRmNTU3YjRjNmMvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDpkMTk0N2Q1YS01MTNkLTQ3ZGYtYTM2ZC0zZjAzZTNlOWI4ODdyY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFgg4ocN5zCrjkr4LUC7am3kIobSu+hRGJat2dyn0G/ApbuiY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIMUnKqreyfZgnA7yR6tBxFi3H9W42qqduxxAxtGAy0DKomN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYIJiratROHbjCNcZ4wmjOIMesnGAml6GKRa8SvxkZXW29dGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD2WEB21xAK0IFkaJiTVQevCn0jTgIrjuVLb6UYc+a5bfJDoBeF7Ww07Z6gnLPYb6tzfqPSGGzM05e3STROhDABGWAGY/GOlwAAAAFzUkdCAK7OHOkAAAAEZ0FNQQAAsY8L/GEFAAACr1BMVEX////7+/vZ2dmysrJdXV08PDwwMDAzMzMxMTE1NTV0dHS/v7/t7e35+fmnp6c2NjYqKiomJiYkJCQlJSUnJycoKChLS0vo6Oj6+vqjo6MvLy8jIyMiIiIrKyuSkpL39/dpaWn8/Pzu7u4hISE5OTnm5uY4ODhKSko3NzcpKSna2trW1tYsLCzU1NT9/f3GxsZ8fHzc3NxAQEAtLS1ubm7r6+vY2NilpaWioqLV1dU7Oztqamrv7+/BwcFaWlpXV1f4+Pjx8fFxcXE9PT1ERERZWVmbm5s/Pz9BQUFTU1NkZGTd3d3n5+dcXFxQUFC4uLi5ublNTU2urq7+/v6IiIjw8PCwsLAuLi7j4+PJycljY2NgYGAgICC+vr56enpbW1v19fW3t7dra2ucnJxhYWE+Pj7R0dFFRUVtbW2ZmZlsbGw6Ojp7e3vs7Ozk5OROTk5+fn6pqanMzMxzc3Py8vLCwsJVVVWdnZ2rq6uOjo6vr6+Wlpbh4eFYWFj29vaDg4OHh4eFhYVwcHC0tLRWVlYyMjLFxcW7u7uYmJh1dXV4eHjAwMDS0tJRUVFiYmKQkJB9fX3Ozs5SUlKTk5NeXl6CgoKLi4vLy8uamprHx8eJiYnDw8PX19fT09Pz8/OEhIRCQkKtra1HR0eUlJTPz8+zs7O2trZmZmb09PQfHx92dnafn580NDSAgIBISEhvb29MTEx/f39ycnLf39/ExMRnZ2doaGh3d3dDQ0Oenp7l5eW8vLx5eXlfX1/p6emmpqZGRkZPT08dHR2srKzNzc2oqKiVlZWGhoZlZWWBgYGXl5dJSUnq6upUVFSkpKS1tbWKiore3t6goKDb29uPj4+MjIzi4uIeHh66urrg4ODIyMi9vb3KysqRkZHQ0NCNjY2hoaGqqqocHByxsbEAAAD7FYZkAAAA5XRSTlP///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8A/6eJEAAAAAlwSFlzAAAh1QAAIdUBBJy0nQAADc5JREFUWEftmcuVIyuzhZnhBCNsYcYUJ3LKABfSFEzAHtKS++0gUyVVlVSqPo+1/rvO7i4pk1cEQTyRexfBuXE+CiG4GnuidSG4npxL3R1ng4u95ND5Xq/BlZDDEVKLa1JoIZQjjJ2H1cCiR1CzvQohHSmE5Nr17tzkPelhgcHR3b0D+u9f/wDd+fMJaK3usv9Yk22z0fO9BA8TIYy53qP3g32kurYZXYzHUcK+w6WQ2EA4Nvi+Fkydln5bviSkwo6vhlC6qw4xnO/gbvRfgdg+Hz+Dc3CDTh0gzJ2tF4x5mjdOfXe5fpZ345R8LN75HflsdN+0QgjZRdq2cLCI5u5BtC7osaXgD5iQRsxexOrqBJCurAhX/Ui+SGOMpRv0mvpwO1tA4XwrNy1cCOjTOIJvCDo1aM+72SCFCI3e8iwpowuMPHsuQNX43Gqu96Qf4F8I+J+GtOw5betCiN9gmZ/++qPSGQLajm3UiUhxB5/7ZQsocPUOnWbAY79MKHSfOBgGycQeuqVSHlGn0JvneFzMZ4/BM6GFhlB1YD209mClMebUQkmBo5k+ixPNWJAtNg9PSUqFvmG184O6senZOTiOY8Jpyx/Mh1D5qAeLY7czhonC5rvuVCIKsksp1TCOwdatT+DJ4+eWNlc0UqM+ugGcImqZ2uES3u2hV4OxkoDVqLuxqftu2WbnKEOJiCbhN0yyFzJjc9hM9GOwDxeQ/gcCqs45l7zvbAhHer82ON/zcrKfe9UPtWLOERZX2yMiBorkH87yAZLLg3n+jyK34MefbUQiOHG2vIvlfYZ3nLqdOxbk+3NhX7ijVZuFLjvdQ8r30fU9sFIGpI2w5NrAJECaveFkNVG9D/7+ARpsDy2H0XHQ3sV6hDgwXubTgbJsNuIbMHdnmwk/gLW3aX4ZHcdlECTowbaeM28d3e+NwJXEd1OoR3Jt+oJTQBZyYN9D7biI5InXbgwscmg0zKLePMCOBoUzD3iADwOP3OX8WIKwsnyN4EYaE48WkCdgVZtxBw1CuJPMgsyGhMg8CY6K/+3IIe+Kh+sMcU825wasj0wtu+w2Egu92HSlHCwww+5JSDK+tI2dg/t8bpoAqjlECJY2Ssxjz6nibnGrsEaIxXkyhO9HWJYT6DlwEErUGoR1ZCUPJE9bJQT36LaL0CNWzlXjmDjvpWfbljZto8yZSiswDp7ke9EMIuekGC/H3zviiPLAteNI0yZ/aAf7HUJBTVzevNwXomlJ+Yrii84DeVi4ezZb7hSQPiQFHmKSpAU4MEKcjQnEwidgIGLZq/PdDwvmtms2slXbs5ps5LdYvd2UKcwhkxzKuZRJB7PQF5OBBqALZg36kF4VeWkTxeu5FyI5oL6JW/rCYH52S48wWgvfWdXPUNhMdsj/4X8Ce/LUDH+GpSgYyfn+G8hNWlizfOaXIBW8VHtlz7+ARQ0ViZR8smsKp7Sc2g8wXiUqMj5lm1jIBuvi5UceKFA1SKkjZViR+2V2VvJigelp+BVsBN+nnPZtDDsxqjAkr87P5eIdjD35dVshwojChhiPu6ha91OcvRsFMXnPyKMTcmo/iD7LVb2cfjtaxVH59LrLExIXcYzW9WJ2DX75kXpwQFZ/yx1Sg5BMnLOezx6QjrAaE/p97h4+dtKIvpG8OTbwvKZXfUtvzr5lYgktKnl984W1DoUDZW3PPOSKMdtM4/CjK5SwWYrigziqHUNdQdHGfoWKpd0T+4hno3Xocf4scgzSJrnXkTj9J7MhAPbqlXlMahyyM5ItZrWUoqpuiNfyqfo+QU7jPTUNEoPBWaS0R7T40woVjyJpY11VRV+hs/B5j5SS5ThcpES5UHs3RoYJhgabcI+zjVyDVFrCyRRwlkGg3DCyBMefIvupVDcwgVIyF11DIalE2sTDytnkJ8gDKgNE3JKzR8Am5lPREwQ82fDK8zSdv7pvpL+RBevA7r/MpqVQqpGipuI3NsnfgYnpHkTpD3mm35CbG99lXNAeMixS46SzdaYtyhUP5UKz5p3Sndm6gPuWNl/SCVUIR2xlVke6uDWYDV7ZlLQpopBfZ0NW+iTdQDk36FNDK1Ur3k2qdZQEqc/QMLqvs9W0sQCyN9uceJQ1iA0ze5sIkoSOYWzJZjxgtW0mpYCmG+zFFUlQD5bLLJf7CdW8Pyky5uw7xQwym1sp0lW3z546CZgG6h7hK5bhK8nT1eH0suSyEi2PnfjeFtFv+BasWWbSOGZ4PXSO7PQIuW+69Fiu7UnuNc1Zx8r0c9/KOLXlHglEfbN5T0irgwkMwxObkuIUZZnMH7qOwWnp3vcc/AVomh0nrNtc5dc7Smte2TZN42f7uoEZgkej7QqVfDPL0vaG4awQpu08wdkVAyfVpHY+ZR/2rZVylmDmKp/Ajla8UcUkfCNVDC2EQrZyiyVPk3t1nkv7is7nWnqkkC6dmlodH/3fwaxSg3huvsrifJ2oDT7VwxB4lS6bZulh3Rjkuku1CYbbtoL7lwr0AVaH2fob3kCbZxnzwep4wfaJcxQVrS4YdXOFEKSzr7m+A4UMoqXotBdH3fpOwnQHI5ao4ezhbHwfJsATZ9NvcPmX33J9QebhX5/Rf/gP/+ELzHFewFn+W7h3GDJe/r2okf4+WFV7ggBPtXDitxcxvwSB9P4qPM+PglJd/e+WfFyOXEmE+dRdWWqyBEIxwisv1S+SfOma8sRfPoJ1I2m4raXfW5Az7lklFZUDbRQCcGDRiuzuwl+4llLEX2Bd0oazuVwPHxiUGxbu7O0WeX4I8U+hBNrAZvRjUj+uo6yDoqPMNkrddDeepkk7KvkjI9SQ29b/5Pivycqp0laRnk53IvhCbebbmLqodbkkRvqdFFbyJSNmQlGFfM7/KqKfcM604wP2oyKlFGfeek8+TanUhqmlGMmF8kzQXz89ik1h3bf8et/rLmcd3UKgSkgpNd8oCs50Ka0LzhzqbC0VNMwr+10VB7CU7uL+XRjHKMmZR6t48h3KaQRES0mm1q1fF7z6zZLU13dqrSvtrcw1C1mvb0PyxmqgTi6bPUV+KsM7KiEoZ/lSGRel5fLncnIbuZxMH/3mqDE/u1/nWH4pcWasBzJiCLNoonJrbvTRdEsjQXQVgn3ZD4l730xU1DsoofJnmnEGv3bw8gcIdhQX/dZyb6r+YUfFyak+26o6Tq9Dj8pzODIv32EPyWWj/attr5I2Yhm19T01/hWdAfvo83SsUm9apkpbvocfKH3k9BF7sl9/dehRW7BS9S3gFsZSJJeTr90L64JwbkhcP3bxzzazPBerd8r51GIr2u0u+vRSRvDJgI848wqiYXcFPjWUbJZEid6S7THmPKbu1Qzr4mWhjt6GfjnkH5St7VxQW7l7ew7zJcsnwC9uRCW7FIfPY1KLui0lu/BRw6LNW2y6tNoZbD+f0Yj/sVAKUX38rOlrQ1IeeayRCJmoWLIbtINFw+HlOUV7UT43Tq01K9I/JnLXMSS4xjQx+ls8WRSewfwP0ybrJ9dzgG9cJHqzi4ZdBhlNo3h+qIGHHCPi2ihxa+S8MqU+3hw3ZEIAr/2aLXazRJ+TazpatMw497rcETHRtp3b13okcrEzHaxZXIF9c4utnqRfbtt0SfqtC5zhQoV0g5PZpm6ZDu/2jDWvbYrBJUpeUlYlz/6kJroemVc0EU69fHXaZ9TAQKC94aPMplKkyNa9p/YzWsn71vFd+4Y+6LonDRc33DebaxtimdqlOyak87FnXSpq1de7XvrNkYIYHRrb2Ks09uzBaRWci+TdZq55l3+Rr5x2NYm8kANZBCMRme06WywTfrgdsDFBPpK6nlilXQ9n16T0+tIywYlkIN2SX/28O5KrU7ve5N0ilnY4uTRaDnEugf2gZUCDFJtsaW9uUn/yz0gCPeNl04FrF9XuogvJAamSSBfZyI6l8XIddVsSu/H6Ahp3Prqwe1f3MvbK8abeIwcntX1QGLNIxJpzrqRNMZe593KMy1AkuctJ/QC2dM0q3e+tFbKT5rHQ0zcIIl6JmFc0OdhWgT3GkbEQ5z5+fVBu8TZWjieULaEu0mQ+ZLPDrCDziAX4vSolszM5UsXziUHcLvnEjfJ5T/ou0k2ktY1tYx/QH6woB2KE1qmHeQniOPBkgzRZvz41zPKcL5U/n94DvoVPr/pVtr3rFxJCiN07Q/k4EpIxF7Le7R9qB2Hv0T1UBK4tVDJAX+8DpYxotfHuI0lRP4MX/zzuBeEn0oOVQK5PoOytUwR5caXhRvlX8ra1PmZgU8GyFMkVbZWQh7jT/nmOCD9DZMjZ2aXyDZrzS9J2kB8lEwKk1AAzm8zlajhtmiClFvLVMTrpKh7QZhjM7O6WeQ826cyq5CEingJv3dM2CYlGXz7UFk8E9jL0Y0skd77BNJBB5+vbiJoEtIKCl2K+lyH5Pe3w4slJyjb3iOPOkwPGk+nXRktS1gIn1utvcNE+p8qX8kU4wtSh7MmTOjmieZS0U3OROMtb7kb8RvgPKAMoLGBi+lWKpjP6qBwoFeSqStMUWlQXoSrCdiRWiPwZbsSvrWOp84vC7miEXJi8tUid0gLvJ+Df4HJWiHj5t+WmOoePq8m7IiO2x1+238M/xr+Ze7+GznZB5LPEq/143rI5XLIvYfkV39A+e/97IQ9FkVEeg6/uFAT7teQfxHIT3+DNkPyXcDrsB9hB/xtYP+5fwMbP9n8Ndd//6dP9fwfn/g9aAEyD+ciRGQAAAABJRU5ErkJggg==";
const BIR_BARCODE_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAPAAAAA2CAIAAAB2hoNsAAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOjkzNjA0MTFhLTA0MjMtNGQ3Ni1iOGNiLTUyYjJhNTIzYzQyYwAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNoSSwi4korTdeBZKvPZXuFzgAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6ZTA5MTE0NzUtZTQwOC00MGQwLWIxYzEtODJjNWUwNWNjZDI3bHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaOin5qV+tBKU9u3ceCof/48AAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIBsXv9bqPVeTxi0rMoi/GWp6HocrJqj8AXG92jtMd5S9pGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2hrzFSrQ3X/lSsaRTb9lFBLAAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFggDySalt9DYCpok6YHclj+8015CTD/cCSU5Q9wZbMr4pNkbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYTo5MzYwNDExYS0wNDIzLTRkNzYtYjhjYi01MmIyYTUyM2M0MmMvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDo4YWUzNzkzNi04YzU4LTRkODYtODVmNy00ZWZhNTVjYjEyNDJyY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggGxe/1uo9V5PGLSsyiL8ZanoehysmqPwBcb3aO0x3lL2iY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIG7IAVipamG5O8bO/SJ+4J5rObFd9GQjI7najzlSVE8BomN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYILerdPRK4rO22wMXZFfigfwPfXridSkTHXTGWO9TUB1wdGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD2WECCGEDNCCDT4qqYWt3637HoFYrvmlG0zgKj+hBhlWjITe65An20EeApczhiWmKCZ7jzYb6BLcOEy3+qUj/J0AY5+JIChQAAAAFzUkdCAK7OHOkAAAAEZ0FNQQAAsY8L/GEFAAAACXBIWXMAAA7DAAAOwwHHb6hkAAACDUlEQVR4Xu3YwaojQQxD0WT+/58zMwgagbBxeG8l7lmETrXLroXoRb1fX/p8Pv9+3+//G/dnp/V06Zayv3j93nmSHdLeZzqDy73Trnx2ucvtb8Vr0n6Sy17ZK11OcdNJRCt/9AfoQKBRhUCjCoFGFQKNKgQaVQg0qhBoVLneeD/yWnu/7hatp3sHubxNXp+z3L4u0xTnfaZn5/1lmjJVXtanlZ3qfW8+O+8vlymS/d10BtEKX2hUIdCoQqBRhUCjCoFGFQKNKgQaVQg0qlxvvB95rb1fd8u9Xispe6bsLLk3397nTp1zer6d7FPkJ51zl/j69Jymid/Wp336fkKt8IVGFQKNKgQaVQg0qhBoVCHQqEKgUYVAo8r1xvuR19r7dbfs9ck7XGpkmi7eZ9orOX1a8V2yd5ap5jIlaya5102dp7nut86Q8gwuz5Yn4QuNKgQaVQg0qhBoVCHQqEKgUYVAowqBRpXrjfcjr7X3627Ruky7XHZwlyniNdPEqc90htw71Xul88rpJNO63Ceq5if1U4342+wv2WGS012eKk/CFxpVCDSqEGhUIdCoQqBRhUCjCoFGFQKNKtcb70dea+/X3aL1ydQn5VufldMvfbx+4pX7rpw47fLK7HafKNnNV9Jek3O/7aYV2fe6nOvyPDmFLzSqEGhUIdCoQqBRhUCjCoFGFQKNKgQaRV6vvyaB+HZYZ+5KAAAAAElFTkSuQmCC";

const fmtCert = (n) => Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const QUARTER_2307_BOUNDS = {
  1: { fromMonth: "01", fromDay: "01", toMonth: "03", toDay: "31" },
  2: { fromMonth: "04", fromDay: "01", toMonth: "06", toDay: "30" },
  3: { fromMonth: "07", fromDay: "01", toMonth: "09", toDay: "30" },
  4: { fromMonth: "10", fromDay: "01", toMonth: "12", toDay: "31" },
};

// Flattens a Purchase Journal or Cash Disbursements row into the common shape the 2307 grouping
// needs. Income payment = the EWT tax base: (VATable + Non-VAT) for purchases, the full amount
// for disbursements — matching how computePurchaseRow / computeDisbRow derive `ewt`.
function normalize2307Row(row, journalKey) {
  const isPurch = journalKey === "purchases";
  return {
    name: (isPurch ? row.supplier : row.vendor) || "",
    tin: row.tin || "",
    address: row.address || "",
    date: row.date,
    atc: row.atc || "",
    atcRate: num(row.atcRate),
    ewt: num(row.ewt),
    incomePayment: isPurch ? num(row.vatable) + num(row.nonvat) : num(row.amount),
  };
}

// Groups the selected journal rows into one certificate payload per supplier: filter to rows
// with an ATC and EWT > 0, require a single quarter, group by supplier (normalized TIN + name),
// sub-group by ATC, bucket amounts into month-of-quarter columns, sum EWT. Returns
// { groups, quarterInfo } or { error }.
function build2307SupplierGroups(selectedRows, journalKey, data) {
  const norm = selectedRows.map((r) => normalize2307Row(r, journalKey));
  const qualifying = norm.filter((r) => r.atc && r.ewt > 0);
  if (qualifying.length === 0) {
    return { error: "None of the selected rows have both an ATC and withholding tax (EWT > 0) — nothing to certify." };
  }

  const quarters = new Set();
  for (const r of qualifying) {
    const d = parseAppDate(r.date);
    if (!d) return { error: "One or more selected rows has an unrecognized date — fix it before generating." };
    quarters.add(`${d.getFullYear()}-Q${Math.floor(d.getMonth() / 3) + 1}`);
  }
  if (quarters.size > 1) {
    return { error: `The selected rows span ${quarters.size} quarters (${[...quarters].sort().join(", ")}). Narrow the selection to a single quarter and try again.` };
  }

  const qKey = [...quarters][0];
  const year = Number(qKey.slice(0, 4));
  const quarter = Number(qKey.slice(-1));
  const qStartMonth = (quarter - 1) * 3; // 0-indexed first month of the quarter

  const atcByCode = Object.fromEntries((data.atc || []).map((a) => [a.code, a]));
  const suppliers = data.suppliers || [];

  const bySupplier = new Map();
  for (const r of qualifying) {
    const key = `${normalizeTin(r.tin)}|${r.name.trim().toLowerCase()}`;
    if (!bySupplier.has(key)) bySupplier.set(key, []);
    bySupplier.get(key).push(r);
  }

  const groups = [];
  for (const rows of bySupplier.values()) {
    const first = rows[0];
    const firstTin = normalizeTin(first.tin);
    const sup =
      (firstTin && suppliers.find((s) => normalizeTin(s.tin) === firstTin)) ||
      suppliers.find((s) => (s.name || "").trim().toLowerCase() === first.name.trim().toLowerCase());
    const payeeName = sup
      ? (sup.type === "Individual"
          ? [sup.surname, sup.firstName, sup.middleName].filter(Boolean).join(", ")
          : (sup.registeredName || sup.name || first.name))
      : first.name;

    const byAtc = new Map();
    for (const r of rows) {
      if (!byAtc.has(r.atc)) byAtc.set(r.atc, []);
      byAtc.get(r.atc).push(r);
    }
    const ewtLines = [];
    for (const [atc, arows] of byAtc) {
      const months = [0, 0, 0];
      let taxWithheld = 0;
      for (const r of arows) {
        const d = parseAppDate(r.date);
        const idx = d.getMonth() - qStartMonth;
        if (idx >= 0 && idx <= 2) months[idx] += r.incomePayment;
        taxWithheld += r.ewt;
      }
      ewtLines.push({
        natureLabel: (atcByCode[atc] && atcByCode[atc].desc) || atc,
        atc,
        m1: round2(months[0]), m2: round2(months[1]), m3: round2(months[2]),
        taxWithheld: round2(taxWithheld),
      });
    }
    ewtLines.sort((a, b) => a.atc.localeCompare(b.atc));

    groups.push({
      payeeName: (payeeName || "").trim() || "(unnamed supplier)",
      payeeTin: firstTin,
      payeeAddress: (sup && sup.address) || first.address || "",
      payeeZip: (sup && sup.zipCode) || "",
      payeePosition: null, // Haki doesn't collect an individual payee signatory's name/position
      ewtLines,
    });
  }
  groups.sort((a, b) => a.payeeName.localeCompare(b.payeeName));

  const b = QUARTER_2307_BOUNDS[quarter];
  const quarterInfo = {
    fromMonth: b.fromMonth, fromDay: b.fromDay, fromYear: String(year),
    toMonth: b.toMonth, toDay: b.toDay, toYear: String(year),
    quarter, year,
  };
  return { groups, quarterInfo };
}

async function build2307Pdf(jsPDFCtor, filingDetails, supplierGroup, quarterInfo) {
  const doc = new jsPDFCtor({ orientation: "portrait", unit: "pt", format: "a4" });
  const left = 34, right = 561;
  let y = 30;
  // One consistent gray value used for every shaded section in the form (Part I/II headers,
  // both Part III table headers, CONFORME, and both signature blocks) — a prior draft had three
  // different values (235, 224, and white/255) scattered across these, which is what "equal
  // shading across the whole form" is fixing.
  const GRAY = [222, 222, 222];

  // ---- Header / masthead ----
  // Matches the real BIR Form 2307 top band: one full-width bordered rectangle split into
  // three columns \u2014 [ For BIR Use Only / BCS Item  over  BIR Form No. 2307 January 2018 (ENCS) ]
  // | [ seal + Republic of the Philippines \u2026 over the title ] | [ barcode over "2307 01/18ENCS" ].
  doc.setDrawColor(0); doc.setLineWidth(0.75);
  const headTop = y;
  const headH = 72;
  const fbRowH = 26;                     // height of the "For BIR Use Only" sub-cell
  const colForm = left + 132;            // left column / centre column divider
  const colBarcode = right - 106;        // centre column / barcode column divider

  // Outer rectangle, column dividers (full height), and the horizontal split in the left column.
  doc.rect(left, headTop, right - left, headH);
  doc.line(colForm, headTop, colForm, headTop + headH);
  doc.line(colBarcode, headTop, colBarcode, headTop + headH);
  doc.line(left, headTop + fbRowH, colForm, headTop + fbRowH);

  // Left column, top sub-cell.
  doc.setFontSize(6.5); doc.setFont("helvetica", "normal");
  doc.text("For BIR", left + 5, headTop + 10);
  doc.text("Use Only", left + 5, headTop + 20);
  doc.text("BCS/", left + 62, headTop + 10);
  doc.text("Item:", left + 62, headTop + 20);

  // Left column, bottom sub-cell.
  const formY = headTop + fbRowH;
  doc.setFontSize(6.5); doc.setFont("helvetica", "normal");
  doc.text("BIR Form No.", left + 5, formY + 11);
  doc.setFontSize(16); doc.setFont("helvetica", "bold");
  doc.text("2307", left + 6, formY + 32);
  doc.setFontSize(6.5); doc.setFont("helvetica", "normal");
  doc.text("January 2018 (ENCS)", left + 5, formY + 43);

  // Centre column: seal + Republic of the Philippines block (upper), form title (lower).
  const titleCx = (colForm + colBarcode) / 2;
  doc.addImage(`data:image/png;base64,${BIR_SEAL_BASE64}`, "PNG", titleCx - 94, headTop + 4, 30, 28);
  doc.setFontSize(8); doc.setFont("helvetica", "bold");
  doc.text("Republic of the Philippines", titleCx + 10, headTop + 13, { align: "center" });
  doc.text("Department of Finance", titleCx + 10, headTop + 22, { align: "center" });
  doc.text("Bureau of Internal Revenue", titleCx + 10, headTop + 31, { align: "center" });
  doc.setFontSize(14); doc.setFont("helvetica", "bold");
  doc.text("Certificate of Creditable", titleCx, headTop + 51, { align: "center" });
  doc.text("Tax Withheld At Source", titleCx, headTop + 66, { align: "center" });

  // Barcode column.
  doc.addImage(`data:image/png;base64,${BIR_BARCODE_BASE64}`, "PNG", colBarcode + 4, headTop + 8, right - colBarcode - 8, 20);
  doc.setFontSize(6.5); doc.setFont("helvetica", "normal");
  doc.text("2307 01/18ENCS", right - 4, headTop + headH - 6, { align: "right" });

  y = headTop + headH;
  doc.setFontSize(7.5); doc.setFont("helvetica", "normal");
  doc.text("Fill in all applicable spaces. Mark all appropriate boxes with an \u201cX\u201d.", left, y + 9);
  y += 12;

  // ---- Field 1: Period ----
  doc.rect(left, y, right - left, 16);
  doc.setFontSize(7.5); doc.setFont("helvetica", "normal");
  doc.text("1  For the Period", left + 3, y + 10);
  doc.text("From", left + 90, y + 10);
  const periodBoxes = (x0, vals) => {
    let x = x0;
    [vals.m, vals.d, vals.yr].forEach((v, i) => {
      const w = i === 2 ? 30 : 18;
      doc.rect(x, y + 2, w, 12);
      doc.text(v, x + w / 2, y + 11, { align: "center" });
      x += w + 2;
    });
  };
  periodBoxes(left + 115, { m: quarterInfo.fromMonth, d: quarterInfo.fromDay, yr: quarterInfo.fromYear });
  doc.text("(MM/DD/YYYY)", left + 200, y + 10);
  doc.text("To", left + 260, y + 10);
  periodBoxes(left + 278, { m: quarterInfo.toMonth, d: quarterInfo.toDay, yr: quarterInfo.toYear });
  doc.text("(MM/DD/YYYY)", left + 363, y + 10);
  y += 16;

  const sectionBar = (label) => {
    doc.setFillColor(...GRAY);
    doc.rect(left, y, right - left, 12, "F");
    doc.rect(left, y, right - left, 12);
    doc.setFont("helvetica", "bold"); doc.setFontSize(7.5);
    doc.text(label, 297.5, y + 9, { align: "center" });
    y += 12;
  };
  const tinRow = (num, label, tin) => {
    doc.rect(left, y, right - left, 16);
    doc.setFont("helvetica", "normal"); doc.setFontSize(7.5);
    doc.text(`${num}  ${label}`, left + 3, y + 10);
    let x = left + 200;
    [tin.slice(0, 3), tin.slice(3, 6), tin.slice(6, 9), "0000"].forEach((seg) => {
      doc.rect(x, y + 2, 40, 12);
      doc.text(seg, x + 20, y + 11, { align: "center" });
      x += 42;
    });
    y += 16;
  };

  // ---- Part I - Payee Information ----
  sectionBar("Part I - Payee   Information");
  tinRow(2, "Taxpayer Identification Number (TIN)", supplierGroup.payeeTin.padEnd(9, "0"));

  doc.rect(left, y, right - left, 20);
  doc.setFont("helvetica", "normal"); doc.setFontSize(7);
  doc.text("3  Payee's Name (Last Name, First Name, Middle Name for Individual OR Registered Name for Non-Individual)", left + 3, y + 8);
  doc.setFont("helvetica", "bold"); doc.setFontSize(8);
  doc.text(supplierGroup.payeeName, left + 3, y + 17);
  y += 20;

  doc.rect(left, y, right - left, 16);
  doc.setFont("helvetica", "normal"); doc.setFontSize(7);
  doc.text("4  Registered Address", left + 3, y + 6);
  doc.text("4A  Zip Code", right - 85, y + 6);
  doc.setFont("helvetica", "bold"); doc.setFontSize(7.5);
  doc.text(supplierGroup.payeeAddress || "", left + 3, y + 14);
  doc.rect(right - 30, y + 2, 26, 11);
  if (supplierGroup.payeeZip) doc.text(supplierGroup.payeeZip, right - 17, y + 11, { align: "center" });
  y += 16;

  doc.rect(left, y, right - left, 12);
  doc.setFont("helvetica", "normal"); doc.setFontSize(7);
  doc.text("5  Foreign Address, if", left + 3, y + 9);
  y += 12;

  // ---- Part II - Payor Information ----
  sectionBar("Part II - Payor   Information");
  tinRow(6, "Taxpayer Identification Number (TIN)", filingDetails.companyTin.padEnd(9, "0"));

  doc.rect(left, y, right - left, 20);
  doc.setFont("helvetica", "normal"); doc.setFontSize(7);
  doc.text("7  Payor's Name (Last Name, First Name, Middle Name for Individual OR Registered Name for Non-Individual)", left + 3, y + 8);
  doc.setFont("helvetica", "bold"); doc.setFontSize(8);
  doc.text(filingDetails.displayName, left + 3, y + 17);
  y += 20;

  doc.rect(left, y, right - left, 16);
  doc.setFont("helvetica", "normal"); doc.setFontSize(7);
  doc.text("8  Registered Address", left + 3, y + 6);
  doc.text("8A  Zip Code", right - 85, y + 6);
  doc.setFont("helvetica", "bold"); doc.setFontSize(7.5);
  doc.text(filingDetails.addr1 || "", left + 3, y + 14, { maxWidth: right - left - 90 });
  doc.rect(right - 30, y + 2, 26, 11);
  if (filingDetails.zipCode) doc.text(filingDetails.zipCode, right - 17, y + 11, { align: "center" });
  y += 16;

  // ---- Part III ----
  sectionBar("Part III - Details of Monthly Income Payments and Tax Withheld for the Quarter");

  const computedEwt = supplierGroup.ewtLines.map((r) => {
    const total = Math.round((r.m1 + r.m2 + r.m3) * 100) / 100;
    // Tax Withheld For the Quarter = the actual EWT already recorded on the journal rows for
    // this ATC (summed upstream), not a re-derivation of total * rate.
    return { ...r, total, taxWithheld: Math.round((r.taxWithheld || 0) * 100) / 100 };
  });
  const ewtTotal = computedEwt.reduce((s, r) => s + r.total, 0);
  const ewtTaxTotal = computedEwt.reduce((s, r) => s + r.taxWithheld, 0);
  const blankRowCount = Math.max(0, 8 - computedEwt.length); // pad to ~8 rows total, matching the verified template capacity

  doc.autoTable({
    startY: y,
    margin: { left, right: 595.28 - right },
    head: [
      [{ content: "Income Payments Subject to Expanded\nWithholding Tax", rowSpan: 2 }, { content: "ATC", rowSpan: 2 },
       { content: "AMOUNT OF INCOME PAYMENTS", colSpan: 4 }, { content: "Tax Withheld For the Quarter", rowSpan: 2 }],
      ["1st Month of\nthe Quarter", "2nd Month of\nthe Quarter", "3rd Month of\nthe Quarter", "Total"],
    ],
    body: [
      ...computedEwt.map((r) => [r.natureLabel, r.atc, fmtCert(r.m1), fmtCert(r.m2), fmtCert(r.m3), fmtCert(r.total), fmtCert(r.taxWithheld)]),
      ...Array(blankRowCount).fill(["", "", "", "", "", "", ""]),
      [{ content: "Total", styles: { fontStyle: "bold" } }, "", "", "", "", { content: fmtCert(ewtTotal), styles: { fontStyle: "bold" } }, { content: fmtCert(ewtTaxTotal), styles: { fontStyle: "bold" } }],
    ],
    styles: { font: "helvetica", fontSize: 6.5, cellPadding: 3, lineColor: [0, 0, 0], lineWidth: 0.5, valign: "middle" },
    headStyles: { fillColor: GRAY, textColor: [0, 0, 0], fontStyle: "bold", halign: "center", fontSize: 6 },
    columnStyles: { 0: { cellWidth: 180 }, 1: { cellWidth: 50, halign: "center" }, 2: { cellWidth: 55, halign: "right" }, 3: { cellWidth: 55, halign: "right" }, 4: { cellWidth: 55, halign: "right" }, 5: { cellWidth: 55, halign: "right" }, 6: { cellWidth: 77, halign: "right" } },
  });
  y = doc.lastAutoTable.finalY;

  // ---- Table 2: Money Payments Subject to Withholding of Business Tax (re-added per request —
  // always empty for now, since no current ATC/data distinguishes this category from standard EWT) ----
  doc.autoTable({
    startY: y,
    margin: { left, right: 595.28 - right },
    body: [
      [{ content: "Money Payments Subject to Withholding\nof Business Tax (Government & Private)", styles: { fontStyle: "bold" } }, "", "", "", "", "", ""],
      ...Array(6).fill(["", "", "", "", "", "", ""]),
      [{ content: "Total", styles: { fontStyle: "bold" } }, "", "", "", "", { content: "-", styles: { fontStyle: "bold" } }, { content: "-", styles: { fontStyle: "bold" } }],
    ],
    styles: { font: "helvetica", fontSize: 6.5, cellPadding: 3, lineColor: [0, 0, 0], lineWidth: 0.5, valign: "middle" },
    columnStyles: { 0: { cellWidth: 180 }, 1: { cellWidth: 50, halign: "center" }, 2: { cellWidth: 55, halign: "right" }, 3: { cellWidth: 55, halign: "right" }, 4: { cellWidth: 55, halign: "right" }, 5: { cellWidth: 55, halign: "right" }, 6: { cellWidth: 77, halign: "right" } },
  });
  y = doc.lastAutoTable.finalY + 8;

  // ---- Declaration + signatures ----
  // Bordered box, matching the real reference: a full-width rectangle with the declaration text
  // inside, rather than free-floating text.
  const declBoxTop = y;
  doc.setFont("helvetica", "normal"); doc.setFontSize(6.5);
  const declaration = "We declare, under the penalties of perjury, that this certificate has been made in good faith, verified by me, and to the best of my knowledge and belief, is true and correct, pursuant to the provisions of the National Internal Revenue Code, as amended, and the regulations issued under authority thereof. Further, we give our consent to the processing of our information as contemplated under the *Data Privacy Act of 2012 (R.A No. 10173) for legitimate and lawful purposes.";
  doc.text(declaration, left + 4, y + 10, { maxWidth: right - left - 8 });
  y += 26;
  doc.rect(left, declBoxTop, right - left, y - declBoxTop);

  // Signature block: ONE shaded, bordered box containing the signature (or blank space), the
  // typed Name(Position/TIN) line, and the two static instructional lines together — matching
  // the real reference's merged gray box, rather than free-floating text with underlines.
  // Below that, a separate bordered row with vertical dividers for Tax Agent Accreditation /
  // Date of Issuance / Date of Expiry — three real columns, not just label text with lines.
  // nameLines: array of 1-2 strings to print above the two static instructional lines.
  const signatureBlock = (nameLines, roleLine, signatureImageDataUrl) => {
    const boxTop = y;
    const sigAreaHeight = signatureImageDataUrl ? 30 : 16;
    const nameLinesCount = nameLines.filter(Boolean).length;
    const nameAreaHeight = nameLinesCount * 10;
    const boxHeight = sigAreaHeight + nameAreaHeight + 2 + 8 + 6;

    // Draw the shaded box first, then a white inset box specifically around the printed-name
    // area (where a physical signature would go) — matching the real reference, which uses white
    // there to visually set the signature line apart from the informational gray around it. Text
    // is drawn once, last, on top of both shapes in the correct order.
    doc.setFillColor(...GRAY);
    doc.rect(left, boxTop, right - left, boxHeight, "FD");

    const whiteBoxTop = boxTop + sigAreaHeight - 8;
    const whiteBoxHeight = nameAreaHeight + 6;
    doc.setFillColor(255, 255, 255);
    doc.rect(297.5 - 110, whiteBoxTop, 220, whiteBoxHeight, "FD");

    let ty = boxTop + sigAreaHeight;
    if (signatureImageDataUrl) {
      try { doc.addImage(signatureImageDataUrl, "PNG", 297.5 - 60, boxTop + 2, 120, 26); } catch (e) { /* text below still prints */ }
    }
    doc.setFont("helvetica", "bold"); doc.setFontSize(8);
    nameLines.filter(Boolean).forEach((line) => { doc.text(line, 297.5, ty, { align: "center" }); ty += 10; });
    ty += 2;
    doc.setFont("helvetica", "normal"); doc.setFontSize(6.5);
    doc.text(roleLine, 297.5, ty, { align: "center" }); ty += 8;
    doc.text("(Indicate Title/Designation and TIN)", 297.5, ty, { align: "center" }); // always static
    y = boxTop + boxHeight;

    // Tax Agent Accreditation row — bordered, three columns with vertical dividers.
    const taHeight = 22;
    const col1W = (right - left) * 0.46, col2W = (right - left) * 0.27;
    doc.rect(left, y, right - left, taHeight);
    doc.line(left + col1W, y, left + col1W, y + taHeight);
    doc.line(left + col1W + col2W, y, left + col1W + col2W, y + taHeight);
    doc.setFontSize(6.5); doc.setFont("helvetica", "normal");
    doc.text("Tax Agent Accreditation No./Attorney's Roll No. (if applicable)", left + col1W / 2, y + 9, { align: "center", maxWidth: col1W - 6 });
    doc.text("Date of Issuance", left + col1W + col2W / 2, y + 9, { align: "center" });
    doc.text("Date of Expiry", left + col1W + col2W + col2W / 2, y + 9, { align: "center" });
    y += taHeight;
  };

  // Payor: one combined line "NAME (POSITION/TIN)" when signatory data is on file — matches the
  // real reference exactly (e.g. "MARIA LINLEY S. SANTOS (PRESIDENT/450-103-423)"). Position and
  // TIN share one parenthetical, separated by a slash — not separate elements as an earlier draft
  // had it. Falls back to just the name if position/TIN haven't been entered yet.
  const payorName = filingDetails.authorizedSignatory || filingDetails.preparedBy;
  const payorLine = payorName && filingDetails.signatoryPosition && filingDetails.signatoryTin
    ? `${payorName} (${filingDetails.signatoryPosition}/${filingDetails.signatoryTin})`
    : payorName;
  signatureBlock([payorLine], "Signature over Printed Name of Payor/Payor's  Authorized Representative/Tax Agent", filingDetails.signatureImage);

  doc.setFillColor(...GRAY);
  doc.rect(left, y, right - left, 12, "F");
  doc.rect(left, y, right - left, 12);
  doc.setFont("helvetica", "bold"); doc.setFontSize(7.5);
  doc.text("CONFORME:", 297.5, y + 9, { align: "center" });
  y += 12;

  // Payee: name on its own line, then "POSITION/TIN" on a second line below it when available —
  // same parenthetical-free, slash-separated convention as the payor's line, just on its own row
  // since the payee's name already sits on the line above. Haki doesn't currently collect an
  // individual payee signatory's name/position (only the supplier entity's own name/TIN), so this
  // second line only appears if that data is ever supplied — otherwise just the name prints.
  // Always text, never an uploaded image, since Haki only stores this business's own signature.
  const formatTinDashed = (t) => {
    const digits = String(t ?? "").replace(/\D/g, "");
    return digits.length >= 9 ? `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6, 9)}` : digits;
  };
  const payeePositionTinLine = supplierGroup.payeePosition && supplierGroup.payeeTin
    ? `${supplierGroup.payeePosition} / ${formatTinDashed(supplierGroup.payeeTin)}`
    : null;
  signatureBlock([supplierGroup.payeeName, payeePositionTinLine], "Signature over Printed Name of Payee/Payee's Authorized Representative/Tax Agent", null);

  doc.setFontSize(6);
  doc.text("*NOTE: The BIR Data Privacy is in the BIR website (www.bir.gov.ph)", left, y + 8);

  return doc;
}

// Single supplier — download the PDF directly (mirrors downloadOne2316).
async function downloadOne2307(filingDetails, supplierGroup, quarterInfo) {
  const jsPDFCtor = await withTimeout(loadJsPDF(), 8000);
  const doc = await build2307Pdf(jsPDFCtor, filingDetails, supplierGroup, quarterInfo);
  const name = supplierGroup.payeeName.replace(/[^a-zA-Z0-9]+/g, "_");
  doc.save(`2307_${name}_${quarterInfo.year}Q${quarterInfo.quarter}.pdf`);
}

// Multiple suppliers — one PDF each, zipped (mirrors downloadAll2316Zip).
async function downloadAll2307Zip(filingDetails, supplierGroups, quarterInfo) {
  const [jsPDFCtor, JSZipCtor] = await Promise.all([withTimeout(loadJsPDF(), 8000), withTimeout(loadJSZip(), 8000)]);
  const zip = new JSZipCtor();
  for (const group of supplierGroups) {
    const doc = await build2307Pdf(jsPDFCtor, filingDetails, group, quarterInfo);
    const name = group.payeeName.replace(/[^a-zA-Z0-9]+/g, "_");
    zip.file(`2307_${name}_${quarterInfo.year}Q${quarterInfo.quarter}.pdf`, doc.output("blob"));
  }
  const blob = await withTimeout(zip.generateAsync({ type: "blob" }), 8000);
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `2307_Certificates_${quarterInfo.year}Q${quarterInfo.quarter}.zip`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

// Shared hook for the "Generate 2307" button on the Purchase Journal and Cash Disbursements pages.
function useGenerate2307(data, journalKey) {
  const [status, setStatus] = useState(null);
  const filingDetails = useMemo(
    () => buildFilingDetails(data, { month: null, year: null, quarter: null, sawtFormType: null }),
    [data]
  );
  const generate = async (selectedRows) => {
    const { groups, quarterInfo, error } = build2307SupplierGroups(selectedRows, journalKey, data);
    if (error) {
      setStatus({ type: "error", text: error });
      setTimeout(() => setStatus(null), 7000);
      return;
    }
    setStatus({ type: "pending", text: `Generating ${groups.length} certificate${groups.length === 1 ? "" : "s"}…` });
    try {
      if (groups.length === 1) {
        await downloadOne2307(filingDetails, groups[0], quarterInfo);
        setStatus({ type: "success", text: "2307 certificate downloaded." });
      } else {
        await downloadAll2307Zip(filingDetails, groups, quarterInfo);
        setStatus({ type: "success", text: `Downloaded ${groups.length} certificates (Q${quarterInfo.quarter} ${quarterInfo.year}) in a ZIP.` });
      }
    } catch (e) {
      setStatus({ type: "error", text: "Couldn't generate the certificate(s) — check your connection and try again." });
    }
    setTimeout(() => setStatus(null), 6000);
  };
  return { status, generate };
}

function EmployeeEntryModal({ data, initial, onCancel, onSubmit }) {
  const [v, setV] = useState(initial || {
    year: new Date().getFullYear(), tin: "", lastName: "", firstName: "", middleName: "", nationality: "", address: "", zipCode: "",
    isMWE: "N", empStatus: "R", empFrom: "", empTo: "", sepReason: "NA",
    p13thMonthNonTax: 0, pDeMinimis: 0, pContributions: 0, pSalariesNonTax: 0,
    pBasicSalary: 0, p13thMonthExcess: 0, pSalariesTax: 0,
    smwPerDay: 0, smwPerMonth: 0, smwPerYear: 0, factorDaysPerYear: 261,
    mweBasicPay: 0, mweHolidayPay: 0, mweOvertimePay: 0, mweNightDiff: 0, mweHazardPay: 0,
    prevEmpStatus: "", prevFrom: "", prevTo: "", prevSepReason: "NA",
    prev13thMonthNonTax: 0, prevDeMinimis: 0, prevContributions: 0, prevSalariesNonTax: 0,
    prevBasicSalary: 0, prev13thMonthExcess: 0, prevSalariesTax: 0,
    taxDue: 0, taxWithheldPrev: 0, taxWithheldPresent: 0, taxCredit5pct: 0, substitutedFiling: "Y",
  });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const computed = computeEmployeeRow(v);
  const isEdit = !!initial;
  return (
    <EntryModalShell title={isEdit ? "Edit Employee" : "Add Employee"} submitLabel={isEdit ? "Save changes" : "Add employee"} onCancel={onCancel}
      onSubmit={() => onSubmit(computeEmployeeRow({ id: v.id || uid(), ...v }))}>
      <div className="sub-block-title" style={{ marginTop: 0 }}>Identity</div>
      <LabeledField label="Year"><Field type="select" options={getAvailableYears(data)} value={v.year} onChange={(x) => set("year", Number(x))} /></LabeledField>
      <LabeledField label="TIN"><Field value={v.tin} onChange={(x) => set("tin", x)} placeholder="000-000-000-000" /></LabeledField>
      <LabeledField label="Last Name"><Field value={v.lastName} onChange={(x) => set("lastName", x)} /></LabeledField>
      <LabeledField label="First Name"><Field value={v.firstName} onChange={(x) => set("firstName", x)} /></LabeledField>
      <LabeledField label="Middle Name"><Field value={v.middleName} onChange={(x) => set("middleName", x)} /></LabeledField>
      <LabeledField label="Nationality (foreigners only)"><Field value={v.nationality} onChange={(x) => set("nationality", x)} /></LabeledField>
      <LabeledField label="Registered Address" wide><Field value={v.address} onChange={(x) => set("address", x)} /></LabeledField>
      <LabeledField label="ZIP Code"><Field value={v.zipCode} onChange={(x) => set("zipCode", x)} /></LabeledField>
      <LabeledField label="Statutory Minimum Wage Earner"><Field type="select" options={[{ value: "N", label: "No — Schedule 1" }, { value: "Y", label: "Yes — Schedule 2" }]} value={v.isMWE} onChange={(x) => set("isMWE", x)} /></LabeledField>

      <div className="sub-block-title">Present Employment</div>
      <LabeledField label="Employment Status"><Field type="select" options={EMP_STATUS_OPTIONS} value={v.empStatus} onChange={(x) => set("empStatus", x)} /></LabeledField>
      <LabeledField label="Period From"><Field type="date" value={v.empFrom} onChange={(x) => set("empFrom", x)} /></LabeledField>
      <LabeledField label="Period To"><Field type="date" value={v.empTo} onChange={(x) => set("empTo", x)} /></LabeledField>
      <LabeledField label="Reason of Separation"><Field type="select" options={SEP_REASON_OPTIONS} value={v.sepReason} onChange={(x) => set("sepReason", x)} /></LabeledField>

      {v.isMWE === "Y" && (<>
        <div className="sub-block-title">Statutory Minimum Wage Detail (Schedule 2)</div>
        <LabeledField label="Basic SMW Per Day"><Field type="number" align="right" value={v.smwPerDay} onChange={(x) => set("smwPerDay", x)} /></LabeledField>
        <LabeledField label="Basic SMW Per Month"><Field type="number" align="right" value={v.smwPerMonth} onChange={(x) => set("smwPerMonth", x)} /></LabeledField>
        <LabeledField label="Basic SMW Per Year"><Field type="number" align="right" value={v.smwPerYear} onChange={(x) => set("smwPerYear", x)} /></LabeledField>
        <LabeledField label="Factor Used (No. of Days/Year)"><Field type="number" align="right" value={v.factorDaysPerYear} onChange={(x) => set("factorDaysPerYear", x)} /></LabeledField>
        <LabeledField label="Basic/SMW Pay (Non-Taxable)"><Field type="number" align="right" value={v.mweBasicPay} onChange={(x) => set("mweBasicPay", x)} /></LabeledField>
        <LabeledField label="Holiday Pay"><Field type="number" align="right" value={v.mweHolidayPay} onChange={(x) => set("mweHolidayPay", x)} /></LabeledField>
        <LabeledField label="Overtime Pay"><Field type="number" align="right" value={v.mweOvertimePay} onChange={(x) => set("mweOvertimePay", x)} /></LabeledField>
        <LabeledField label="Night Shift Differential"><Field type="number" align="right" value={v.mweNightDiff} onChange={(x) => set("mweNightDiff", x)} /></LabeledField>
        <LabeledField label="Hazard Pay"><Field type="number" align="right" value={v.mweHazardPay} onChange={(x) => set("mweHazardPay", x)} /></LabeledField>
      </>)}

      <div className="sub-block-title">Present Employer — Non-Taxable/Exempt Compensation</div>
      <LabeledField label="13th Month Pay & Other Benefits"><Field type="number" align="right" value={v.p13thMonthNonTax} onChange={(x) => set("p13thMonthNonTax", x)} /></LabeledField>
      <LabeledField label="De Minimis Benefits"><Field type="number" align="right" value={v.pDeMinimis} onChange={(x) => set("pDeMinimis", x)} /></LabeledField>
      <LabeledField label="SSS/GSIS/PHIC/HDMF & Union Dues"><Field type="number" align="right" value={v.pContributions} onChange={(x) => set("pContributions", x)} /></LabeledField>
      <LabeledField label="Salaries (₱250K & below) & Other Forms"><Field type="number" align="right" value={v.pSalariesNonTax} onChange={(x) => set("pSalariesNonTax", x)} /></LabeledField>

      <div className="sub-block-title">Present Employer — Taxable Compensation</div>
      <LabeledField label="Basic Salary (net of contributions)"><Field type="number" align="right" value={v.pBasicSalary} onChange={(x) => set("pBasicSalary", x)} /></LabeledField>
      <LabeledField label="13th Month Pay & Other Benefits (excess)"><Field type="number" align="right" value={v.p13thMonthExcess} onChange={(x) => set("p13thMonthExcess", x)} /></LabeledField>
      <LabeledField label="Salaries & Other Forms of Compensation"><Field type="number" align="right" value={v.pSalariesTax} onChange={(x) => set("pSalariesTax", x)} /></LabeledField>
      <ComputedPreview items={[["Total Non-Taxable", fmt(computed.pTotalNonTax)], ["Total Taxable", fmt(computed.pTotalTax)], ["Gross Compensation", fmt(computed.pGross)]]} />

      <div className="sub-block-title">Previous Employer (if applicable)</div>
      <LabeledField label="Employment Status"><Field type="select" options={EMP_STATUS_OPTIONS} value={v.prevEmpStatus} onChange={(x) => set("prevEmpStatus", x)} /></LabeledField>
      <LabeledField label="Period From"><Field type="date" value={v.prevFrom} onChange={(x) => set("prevFrom", x)} /></LabeledField>
      <LabeledField label="Period To"><Field type="date" value={v.prevTo} onChange={(x) => set("prevTo", x)} /></LabeledField>
      <LabeledField label="Reason of Separation"><Field type="select" options={SEP_REASON_OPTIONS} value={v.prevSepReason} onChange={(x) => set("prevSepReason", x)} /></LabeledField>
      <LabeledField label="Non-Taxable Compensation (Previous)"><Field type="number" align="right" value={v.prev13thMonthNonTax} onChange={(x) => set("prev13thMonthNonTax", x)} placeholder="13th month & other benefits" /></LabeledField>
      <LabeledField label="De Minimis (Previous)"><Field type="number" align="right" value={v.prevDeMinimis} onChange={(x) => set("prevDeMinimis", x)} /></LabeledField>
      <LabeledField label="Contributions & Union Dues (Previous)"><Field type="number" align="right" value={v.prevContributions} onChange={(x) => set("prevContributions", x)} /></LabeledField>
      <LabeledField label="Salaries Non-Taxable (Previous)"><Field type="number" align="right" value={v.prevSalariesNonTax} onChange={(x) => set("prevSalariesNonTax", x)} /></LabeledField>
      <LabeledField label="Basic Salary Taxable (Previous)"><Field type="number" align="right" value={v.prevBasicSalary} onChange={(x) => set("prevBasicSalary", x)} /></LabeledField>
      <LabeledField label="13th Month Excess (Previous)"><Field type="number" align="right" value={v.prev13thMonthExcess} onChange={(x) => set("prev13thMonthExcess", x)} /></LabeledField>
      <LabeledField label="Salaries Taxable (Previous)"><Field type="number" align="right" value={v.prevSalariesTax} onChange={(x) => set("prevSalariesTax", x)} /></LabeledField>
      <ComputedPreview items={[["Total Taxable (Previous)", fmt(computed.prevTotalTax)], ["Gross (Previous)", fmt(computed.prevGross)]]} />

      <div className="sub-block-title">Tax Summary</div>
      <LabeledField label="Tax Due (Jan–Dec)"><Field type="number" align="right" value={v.taxDue} onChange={(x) => set("taxDue", x)} /></LabeledField>
      <LabeledField label="Tax Withheld — Previous Employer"><Field type="number" align="right" value={v.taxWithheldPrev} onChange={(x) => set("taxWithheldPrev", x)} /></LabeledField>
      <LabeledField label="Tax Withheld — Present Employer"><Field type="number" align="right" value={v.taxWithheldPresent} onChange={(x) => set("taxWithheldPresent", x)} /></LabeledField>
      <LabeledField label="5% Tax Credit (PERA Act of 2008)"><Field type="number" align="right" value={v.taxCredit5pct} onChange={(x) => set("taxCredit5pct", x)} /></LabeledField>
      <LabeledField label="Substituted Filing"><Field type="select" options={[{ value: "Y", label: "Yes" }, { value: "N", label: "No" }]} value={v.substitutedFiling} onChange={(x) => set("substitutedFiling", x)} /></LabeledField>
      <ComputedPreview items={[["Total Taxable Compensation", fmt(computed.totalTaxableComp)], ["Withheld/Paid in December", fmt(computed.amtWithheldDec)], ["Refunded to Employee", fmt(computed.amtRefunded)]]} />
    </EntryModalShell>
  );
}

function AlphalistEmployeesPage({ data, setData }) {
  const [year, setYear] = useState(new Date().getFullYear());
  const importHook = useJournalImport("employees", data, ({ newRows }) => setData((d) => ({ ...d, employees: [...(d.employees || []), ...newRows] })));
  const [modalState, setModalState] = useState(null); // null | { mode: "add" } | { mode: "edit", row }
  const [search, setSearch] = useState("");
  const sel = useRowSelection();
  const [genStatus, setGenStatus] = useState(null);

  const update = (id, patch) => setData((d) => ({ ...d, employees: d.employees.map((r) => r.id === id ? computeEmployeeRow({ ...r, ...patch }) : r) }));
  const onDelete = (id) => setData((d) => ({ ...d, employees: d.employees.filter((r) => r.id !== id) }));
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const onDeleteSelected = () => setConfirmDeleteOpen(true);
  const confirmDeleteSelected = () => {
    setData((d) => ({ ...d, employees: d.employees.filter((r) => !sel.selected.has(r.id)) }));
    sel.clear();
    setConfirmDeleteOpen(false);
  };

  const allRows = useMemo(() => (data.employees || []).map(computeEmployeeRow), [data.employees]);
  const filteredRows = useMemo(() => allRows.filter((r) => Number(r.year) === year && matchesSearch(r, ["tin","lastName","firstName","middleName"], search)), [allRows, year, search]);
  const sched1Rows = useMemo(() => filteredRows.filter((r) => r.isMWE !== "Y"), [filteredRows]);
  const sched2Rows = useMemo(() => filteredRows.filter((r) => r.isMWE === "Y"), [filteredRows]);
  const totals = filteredRows.reduce((a, r) => ({ gross: a.gross + num(r.pGross), nonTax: a.nonTax + num(r.pTotalNonTax), taxable: a.taxable + num(r.totalTaxableComp), due: a.due + num(r.taxDue), wh: a.wh + num(r.taxWithheldPrev) + num(r.taxWithheldPresent) }), { gross: 0, nonTax: 0, taxable: 0, due: 0, wh: 0 });
  const emptyMsg = (data.employees || []).length === 0 ? "No employees logged yet — add your first one above." : `No employees on file for ${year}.`;
  const filingDetails = useMemo(() => buildFilingDetails(data, { month: null, year, quarter: null, sawtFormType: null }), [data, year]);
  const exportHook = useJournalExport("employees", data, filteredRows, {});

  const onDownloadDat = (scheduleNo) => {
    const rows = scheduleNo === 2 ? sched2Rows : sched1Rows;
    if (rows.length === 0) return;
    downloadAlphalistDat(buildAlphalistDatRecord(filingDetails, rows, year, scheduleNo));
  };
  const onDownloadAll2316 = async () => {
    setGenStatus({ type: "pending", text: "Generating certificates…" });
    try { await downloadAll2316Zip(data, filingDetails, filteredRows, year); setGenStatus({ type: "success", text: `Downloaded ${filteredRows.length} certificate(s) in a ZIP.` }); }
    catch (e) { setGenStatus({ type: "error", text: "Couldn't generate the certificates — check your connection and try again." }); }
    setTimeout(() => setGenStatus(null), 5000);
  };
  const onDownloadOne2316 = async (employee) => {
    setGenStatus({ type: "pending", text: "Generating certificate…" });
    try { await downloadOne2316(data, filingDetails, employee, year); setGenStatus({ type: "success", text: "Certificate downloaded." }); }
    catch (e) { setGenStatus({ type: "error", text: "Couldn't generate the certificate — check your connection and try again." }); }
    setTimeout(() => setGenStatus(null), 4000);
  };

  return (
    <div>
      <SectionHeader icon={FileText} title="Alphalist of Employees" subtitle="BIR Form 2316 certificates and BIR Form 1604-C's two schedules — Schedule 1 (regular employees) and Schedule 2 (minimum wage earners)."
        right={<div className="header-actions"><ImportExportBar journalKey="employees" data={data} importHook={importHook} /><ExportBar exportHook={exportHook} /></div>} />
      <ImportStatus status={importHook.status} />
      <ExportStatus status={exportHook.status} />
      <ReliefCompanyWarning filingDetails={filingDetails} needAddr />
      {sel.selected.size > 0 && (
        <SelectionBar count={sel.selected.size} onClear={sel.clear}>
          <button className="io-btn danger" onClick={onDeleteSelected}><Trash2 size={13} /> Delete selected</button>
        </SelectionBar>
      )}
      <div className="ledger-wrap">
        <div className="table-toolbar">
          <div className="toolbar-left">
            <select className="io-select" value={year} onChange={(e) => setYear(Number(e.target.value))}>
              {getAvailableYears(data).map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
            <SearchBar value={search} onChange={setSearch} placeholder="Search TIN, last name, first name…" />
          </div>
          <AddRowBtn onClick={() => setModalState({ mode: "add" })}>Add employee</AddRowBtn>
        </div>
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr>
              <th style={{width:32}}><SelectAllCheckbox ids={filteredRows.map((r) => r.id)} selected={sel.selected} toggleAll={sel.toggleAll} /></th>
              <th style={{minWidth:130}}>TIN</th><th style={{minWidth:140}}>Last Name</th><th style={{minWidth:120}}>First Name</th><th style={{minWidth:120}}>Middle Name</th>
              <th style={{minWidth:70}}>MWE</th><th style={{minWidth:80}}>Status</th>
              <th style={{minWidth:120}} className="num-head">Gross (Present)</th><th style={{minWidth:120}} className="num-head">Non-Taxable</th>
              <th style={{minWidth:130}} className="num-head">Total Taxable Comp.</th><th style={{minWidth:100}} className="num-head">Tax Due</th>
              <th style={{minWidth:110}} className="num-head">Tax Withheld</th>
              <th style={{minWidth:130}}></th><th style={{width:36}}></th>
            </tr></thead>
            <tbody>
              {filteredRows.length === 0 && <tr><td colSpan={14} className="empty-row">{emptyMsg}</td></tr>}
              {filteredRows.map((r) => (
                <tr key={r.id} className={sel.selected.has(r.id) ? "row-selected" : ""}>
                  <td className="text-center"><RowCheckbox id={r.id} selected={sel.selected} toggle={sel.toggle} /></td>
                  <td><ReadCell>{reliefFormatTIN(normalizeTin(r.tin))}</ReadCell></td>
                  <td><ReadCell>{r.lastName}</ReadCell></td>
                  <td><ReadCell>{r.firstName}</ReadCell></td>
                  <td><ReadCell>{r.middleName}</ReadCell></td>
                  <td><ReadCell align="center">{r.isMWE === "Y" ? "Yes" : "No"}</ReadCell></td>
                  <td><ReadCell align="center">{r.empStatus}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.pGross)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.pTotalNonTax)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.totalTaxableComp)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.taxDue)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(round2(num(r.taxWithheldPrev) + num(r.taxWithheldPresent)))}</ReadCell></td>
                  <td className="text-center">
                    <button className="io-btn" style={{ padding: "4px 8px", marginRight: 4 }} onClick={() => setModalState({ mode: "edit", row: r })}>Edit</button>
                    <button className="io-btn" style={{ padding: "4px 8px" }} onClick={() => onDownloadOne2316(r)} title="Download this employee's 2316"><FileDown size={12} /></button>
                  </td>
                  <td className="text-center"><DelBtn onClick={() => onDelete(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {filteredRows.length > 0 && (
              <tfoot><tr>
                <td colSpan={7} className="totals-label">Totals</td>
                <td className="num">{fmt(totals.gross)}</td><td className="num">{fmt(totals.nonTax)}</td>
                <td className="num">{fmt(totals.taxable)}</td><td className="num">{fmt(totals.due)}</td><td className="num">{fmt(totals.wh)}</td><td colSpan={2}></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>
      {filteredRows.length > 0 && (
        <>
          <ExportStatus status={genStatus} />
          <div className="ledger-wrap"><div className="table-toolbar">
            <div className="header-actions">
              <button className="io-btn" onClick={() => onDownloadDat(1)} disabled={sched1Rows.length === 0}><FileDown size={13} /> Schedule 1 .dat ({sched1Rows.length} regular)</button>
              <button className="io-btn" onClick={() => onDownloadDat(2)} disabled={sched2Rows.length === 0}><FileDown size={13} /> Schedule 2 .dat ({sched2Rows.length} MWE)</button>
              <button className="io-btn" onClick={onDownloadAll2316}><FileDown size={13} /> Download all 2316 certificates (.zip)</button>
            </div>
          </div></div>
        </>
      )}
      {modalState && (
        <EmployeeEntryModal data={data} initial={modalState.mode === "edit" ? modalState.row : null}
          onCancel={() => setModalState(null)}
          onSubmit={(row) => {
            setData((d) => ({
              ...d,
              employees: modalState.mode === "edit"
                ? d.employees.map((x) => x.id === row.id ? row : x)
                : [...(d.employees || []), row],
            }));
            setModalState(null);
          }} />
      )}
      {confirmDeleteOpen && (
        <ConfirmModal title="Delete selected" danger confirmLabel="Delete"
          message={`Delete ${sel.selected.size} selected employee record${sel.selected.size === 1 ? "" : "s"}? This cannot be undone.`}
          onCancel={() => setConfirmDeleteOpen(false)} onConfirm={confirmDeleteSelected} />
      )}
      <div className="rc-warn">Schedule 1 covers regular employees; Schedule 2 covers statutory minimum wage earners — each employee's ".dat" file is generated separately based on the Statutory Minimum Wage Earner flag. This layout follows the BIR's general published Alphalist structure and a real sample workbook, but hasn't been independently confirmed the same way SLSPI/QAP/SAWT were. Always run it through the BIR Alphalist Data Entry & Validation Module before actual submission.</div>
    </div>
  );
}

/* ============================== FIXED ASSET LEDGER ============================== */

const ASSET_CATEGORIES = ["Transportation Vehicle","Furniture and Fixtures","Office Equipment","Machinery and Equipment","Building and Improvements","Other"];

function FixedAssetEntryModal({ data, onCancel, onSubmit }) {
  const assetAcctOptions = data.coa.filter((a) => a.type === "Asset").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const expAcctOptions = data.coa.filter((a) => a.type === "Expense").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const [v, setV] = useState({
    name: "", code: "", category: "Office Equipment", acquisitionDate: todayMDY(),
    cost: 0, salvageValue: 0, usefulLifeYears: 5,
    assetAccount: "", accumDepAccount: "", depExpenseAccount: "6020",
    status: "Active", postedPeriods: [],
  });
  const set = (k, val) => setV((p) => ({ ...p, [k]: val }));
  const computed = computeAssetRow(v);
  return (
    <EntryModalShell title="Add Fixed Asset" submitLabel="Add asset" onCancel={onCancel}
      onSubmit={() => onSubmit(computeAssetRow({ id: uid(), ...v }))}>
      <LabeledField label="Asset Name" wide><Field value={v.name} onChange={(x) => set("name", x)} placeholder="Delivery Van - Toyota Hiace" /></LabeledField>
      <LabeledField label="Asset Code"><Field value={v.code} onChange={(x) => set("code", x)} placeholder="FA-001" /></LabeledField>
      <LabeledField label="Category"><Field type="select" options={ASSET_CATEGORIES} value={v.category} onChange={(x) => set("category", x)} /></LabeledField>
      <LabeledField label="Acquisition Date"><Field type="date" value={v.acquisitionDate} onChange={(x) => set("acquisitionDate", x)} /></LabeledField>
      <LabeledField label="Cost"><Field type="number" align="right" value={v.cost} onChange={(x) => set("cost", x)} /></LabeledField>
      <LabeledField label="Salvage Value"><Field type="number" align="right" value={v.salvageValue} onChange={(x) => set("salvageValue", x)} /></LabeledField>
      <LabeledField label="Useful Life (Years)"><Field type="number" align="right" value={v.usefulLifeYears} onChange={(x) => set("usefulLifeYears", x)} /></LabeledField>
      <LabeledField label="Asset Account"><Field type="combo" options={assetAcctOptions} value={v.assetAccount} onChange={(x) => set("assetAccount", x)} /></LabeledField>
      <LabeledField label="Accumulated Depreciation Account"><Field type="combo" options={assetAcctOptions} value={v.accumDepAccount} onChange={(x) => set("accumDepAccount", x)} /></LabeledField>
      <LabeledField label="Depreciation Expense Account"><Field type="combo" options={expAcctOptions} value={v.depExpenseAccount} onChange={(x) => set("depExpenseAccount", x)} /></LabeledField>
      <LabeledField label="Status"><Field type="select" options={["Active","Disposed"]} value={v.status} onChange={(x) => set("status", x)} /></LabeledField>
      <ComputedPreview items={[["Monthly Depreciation", fmt(computed.monthlyDep)], ["Book Value (at start)", fmt(v.cost)]]} />
    </EntryModalShell>
  );
}

function FixedAssetLedgerPage({ data, setData }) {
  const coaByCode = useMemo(() => Object.fromEntries(data.coa.map((a) => [a.code, a])), [data.coa]);
  const assetAcctOptions = data.coa.filter((a) => a.type === "Asset").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const expAcctOptions = data.coa.filter((a) => a.type === "Expense").map((a) => ({ value: a.code, label: `${a.code} · ${a.name}` }));
  const [modalOpen, setModalOpen] = useState(false);
  const assets = useMemo(() => (data.fixedAssets || []).map(computeAssetRow), [data.fixedAssets]);

  const update = (id, patch) => setData((d) => ({ ...d, fixedAssets: d.fixedAssets.map((r) => r.id === id ? computeAssetRow({ ...r, ...patch }) : r) }));
  const onDelete = (id) => setData((d) => ({ ...d, fixedAssets: d.fixedAssets.filter((r) => r.id !== id) }));
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  const [postMonth, setPostMonth] = useState(new Date().getMonth() + 1);
  const [postYear, setPostYear] = useState(new Date().getFullYear());
  const periodKey = `${postYear}-${pad2(postMonth)}`;
  const periodLabel = `${MONTHS[postMonth - 1]} ${postYear}`;

  const dueThisPeriod = useMemo(() => {
    return assets.filter((a) => {
      if (a.status === "Disposed") return false;
      if (a.fullyDepreciated) return false;
      if (a.monthlyDep <= 0) return false;
      const acqDate = parseAppDate(a.acquisitionDate);
      if (!acqDate) return false;
      const acqKey = `${acqDate.getFullYear()}-${pad2(acqDate.getMonth() + 1)}`;
      if (periodKey < acqKey) return false; // can't depreciate before acquisition
      if ((a.postedPeriods || []).includes(periodKey)) return false; // already posted this period
      return true;
    });
  }, [assets, periodKey]);
  const eligible = useMemo(() => dueThisPeriod.filter((a) => a.depExpenseAccount && a.accumDepAccount), [dueThisPeriod]);
  const missingAccounts = useMemo(() => dueThisPeriod.filter((a) => !a.depExpenseAccount || !a.accumDepAccount), [dueThisPeriod]);

  const [postStatus, setPostStatus] = useState(null);

  const onPost = () => {
    if (eligible.length === 0) return;
    const expByAccount = {}, accumByAccount = {};
    eligible.forEach((a) => {
      expByAccount[a.depExpenseAccount] = round2((expByAccount[a.depExpenseAccount] || 0) + a.monthlyDep);
      accumByAccount[a.accumDepAccount] = round2((accumByAccount[a.accumDepAccount] || 0) + a.monthlyDep);
    });
    const lines = [
      ...Object.entries(expByAccount).map(([account, debit]) => ({ id: uid(), account, debit, credit: 0 })),
      ...Object.entries(accumByAccount).map(([account, credit]) => ({ id: uid(), account, debit: 0, credit })),
    ];
    const jvNo = `JV-${String(data.generalJournal.length + 1).padStart(3, "0")}`;
    const monthEndDate = monthEnd(postYear, postMonth - 1);
    const jv = { id: uid(), jvNo, date: `${pad2(postMonth)}/${pad2(monthEndDate.getDate())}/${postYear}`, particulars: `Monthly depreciation — ${periodLabel}`, lines };
    setData((d) => ({
      ...d,
      generalJournal: [...d.generalJournal, jv],
      fixedAssets: d.fixedAssets.map((a) => eligible.some((e) => e.id === a.id) ? { ...a, postedPeriods: [...(a.postedPeriods || []), periodKey] } : a),
    }));
    setPostStatus({ type: "success", text: `Posted ${jvNo} — ${periodLabel} depreciation for ${eligible.length} asset${eligible.length === 1 ? "" : "s"}.` });
    setTimeout(() => setPostStatus(null), 5000);
  };

  const totals = assets.reduce((a, r) => ({ cost: a.cost + num(r.cost), accumDep: a.accumDep + num(r.accumulatedDep), bookValue: a.bookValue + num(r.bookValue) }), { cost: 0, accumDep: 0, bookValue: 0 });

  return (
    <div>
      <SectionHeader icon={Landmark} title="Fixed Asset Ledger" subtitle="Straight-line depreciation, computed automatically — post it to the General Journal whenever you're ready."
        right={<AddRowBtn onClick={() => setModalOpen(true)}>Add asset</AddRowBtn>} />
      <div className="ledger-wrap">
        <div className="table-scroll">
          <table className="ledger-table">
            <thead><tr>
              <th style={{minWidth:190}}>Asset Name</th><th style={{minWidth:90}}>Code</th><th style={{minWidth:160}}>Category</th>
              <th style={{minWidth:120}}>Acquisition Date</th><th style={{minWidth:110}} className="num-head">Cost</th>
              <th style={{minWidth:110}} className="num-head">Salvage Value</th><th style={{minWidth:90}} className="num-head">Life (Yrs)</th>
              <th style={{minWidth:110}} className="num-head">Monthly Dep.</th><th style={{minWidth:120}} className="num-head">Accum. Dep.</th>
              <th style={{minWidth:110}} className="num-head">Book Value</th>
              <th style={{minWidth:170}}>Asset Account</th><th style={{minWidth:200}}>Accum. Dep. Account</th><th style={{minWidth:190}}>Dep. Expense Account</th>
              <th style={{minWidth:100}}>Status</th><th style={{width:36}}></th>
            </tr></thead>
            <tbody>
              {assets.length === 0 && <tr><td colSpan={15} className="empty-row">No fixed assets yet — add your first one above.</td></tr>}
              {assets.map((r) => (
                <tr key={r.id}>
                  <td><Field value={r.name} onChange={(v) => update(r.id, { name: v })} /></td>
                  <td><Field value={r.code} onChange={(v) => update(r.id, { code: v })} /></td>
                  <td><Field type="select" options={ASSET_CATEGORIES} value={r.category} onChange={(v) => update(r.id, { category: v })} /></td>
                  <td><Field type="date" value={r.acquisitionDate} onChange={(v) => update(r.id, { acquisitionDate: v })} /></td>
                  <td><Field type="number" align="right" value={r.cost} onChange={(v) => update(r.id, { cost: v })} /></td>
                  <td><Field type="number" align="right" value={r.salvageValue} onChange={(v) => update(r.id, { salvageValue: v })} /></td>
                  <td><Field type="number" align="right" value={r.usefulLifeYears} onChange={(v) => update(r.id, { usefulLifeYears: v })} /></td>
                  <td><ReadCell align="right">{fmt(r.monthlyDep)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.accumulatedDep)}</ReadCell></td>
                  <td><ReadCell align="right">{fmt(r.bookValue)}</ReadCell></td>
                  <td><Field type="combo" options={assetAcctOptions} value={r.assetAccount} onChange={(v) => update(r.id, { assetAccount: v })} /></td>
                  <td><Field type="combo" options={assetAcctOptions} value={r.accumDepAccount} onChange={(v) => update(r.id, { accumDepAccount: v })} /></td>
                  <td><Field type="combo" options={expAcctOptions} value={r.depExpenseAccount} onChange={(v) => update(r.id, { depExpenseAccount: v })} /></td>
                  <td><Field type="select" options={["Active","Disposed"]} value={r.status} onChange={(v) => update(r.id, { status: v })} /></td>
                  <td className="text-center"><DelBtn onClick={() => setConfirmDeleteId(r.id)} /></td>
                </tr>
              ))}
            </tbody>
            {assets.length > 0 && (
              <tfoot><tr>
                <td colSpan={4} className="totals-label">Totals</td>
                <td className="num">{fmt(totals.cost)}</td><td colSpan={2}></td>
                <td></td><td className="num">{fmt(totals.accumDep)}</td><td className="num">{fmt(totals.bookValue)}</td>
                <td colSpan={5}></td>
              </tr></tfoot>
            )}
          </table>
        </div>
      </div>

      <div className="sub-block-title" style={{ marginTop: 28 }}>Post Monthly Depreciation</div>
      <div className="form-card">
        <div className="form-grid">
          <LabeledField label="Month"><Field type="select" options={MONTHS.map((m, i) => ({ value: i + 1, label: m }))} value={postMonth} onChange={(v) => setPostMonth(Number(v))} /></LabeledField>
          <LabeledField label="Year"><Field type="select" options={getAvailableYears(data)} value={postYear} onChange={(v) => setPostYear(Number(v))} /></LabeledField>
        </div>
      </div>
      {missingAccounts.length > 0 && (
        <div className="callout">
          <Info size={16} />
          <div>
            <div className="callout-title">{missingAccounts.length} asset{missingAccounts.length === 1 ? "" : "s"} skipped — missing account setup</div>
            <div className="callout-body">{missingAccounts.map((a) => a.name).join(", ")} {missingAccounts.length === 1 ? "is" : "are"} due for {periodLabel} depreciation but {missingAccounts.length === 1 ? "hasn't" : "haven't"} got both an Asset Account and Accumulated Depreciation Account set above. Fill those in to include {missingAccounts.length === 1 ? "it" : "them"} in this posting.</div>
          </div>
        </div>
      )}
      {eligible.length === 0 ? (
        <div className="empty-panel">No assets are due for depreciation in {periodLabel} — either already posted, not yet acquired, fully depreciated, disposed, or missing required accounts.</div>
      ) : (
        <>
          <div className="ledger-wrap">
            <div className="table-scroll">
              <table className="ledger-table report-table">
                <thead><tr><th>Asset</th><th>Expense Account</th><th>Accum. Dep. Account</th><th className="num-head">Amount</th></tr></thead>
                <tbody>
                  {eligible.map((a) => (
                    <tr key={a.id}>
                      <td>{a.name}</td>
                      <td>{coaByCode[a.depExpenseAccount] ? `${a.depExpenseAccount} · ${coaByCode[a.depExpenseAccount].name}` : a.depExpenseAccount}</td>
                      <td>{coaByCode[a.accumDepAccount] ? `${a.accumDepAccount} · ${coaByCode[a.accumDepAccount].name}` : a.accumDepAccount}</td>
                      <td className="num">{fmt(a.monthlyDep)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot><tr>
                  <td colSpan={3} className="totals-label">Total to post — {periodLabel}</td>
                  <td className="num">{fmt(round2(eligible.reduce((s, a) => s + a.monthlyDep, 0)))}</td>
                </tr></tfoot>
              </table>
            </div>
          </div>
          <ExportStatus status={postStatus} />
          <div className="ledger-wrap"><div className="table-toolbar">
            <AddRowBtn onClick={onPost}>Post {periodLabel} depreciation to General Journal</AddRowBtn>
          </div></div>
        </>
      )}

      {modalOpen && (
        <FixedAssetEntryModal data={data} onCancel={() => setModalOpen(false)}
          onSubmit={(row) => { setData((d) => ({ ...d, fixedAssets: [...(d.fixedAssets || []), row] })); setModalOpen(false); }} />
      )}
      {confirmDeleteId && (
        <ConfirmModal title="Delete asset" danger confirmLabel="Delete"
          message="Delete this fixed asset? Its depreciation history and posted-period tracking will be lost. Any depreciation already posted to the General Journal stays there — you'd need to remove those entries separately if needed."
          onCancel={() => setConfirmDeleteId(null)} onConfirm={() => { onDelete(confirmDeleteId); setConfirmDeleteId(null); }} />
      )}
    </div>
  );
}

/* ============================== SLS / SLP / SLI — AUTO-SOURCED FROM HAKI'S OWN JOURNALS ============================== */
// Instead of asking for an Excel upload, SLS/SLP/SLI pull directly from the Sales Journal, Purchase
// Journal, and Importation Ledger, aggregate by payee for the period, and sort alphabetically.

function reliefPartySortKey(masterMatch, fallbackName) {
  if (masterMatch && masterMatch.type === "Individual") return masterMatch.surname || fallbackName || "";
  if (masterMatch) return masterMatch.name || masterMatch.registeredName || fallbackName || "";
  return fallbackName || "";
}

function buildSlsRowsFromSales(data, month, year) {
  const start = new Date(year, month - 1, 1), end = monthEnd(year, month - 1);
  const relevant = data.sales.filter((r) => inRange(r.date, start, end));
  const groups = new Map();
  relevant.forEach((r) => {
    const cust = data.customers.find((c) => (c.name || "").toLowerCase() === (r.customer || "").toLowerCase());
    const tin = ((cust && cust.tin) || r.tin || "").replace(/\D/g, "");
    const key = tin ? tin.replace(/\D/g, "") : `NOTIN:${(r.customer || "").toLowerCase()}`;
    if (!groups.has(key)) {
      const isIndiv = cust && cust.type === "Individual";
      groups.set(key, {
        tin, partyName: isIndiv ? "" : ((cust && (cust.name || cust.registeredName)) || r.customer || ""),
        surname: isIndiv ? cust.surname : "", firstName: isIndiv ? cust.firstName : "", middleName: isIndiv ? cust.middleName : "",
        address1: (cust && cust.address) || r.address || "", address2: "",
        exemptSales: 0, zeroRatedSales: 0, taxableSales: 0, outputVat: 0,
        sortKey: reliefPartySortKey(cust, r.customer),
      });
    }
    const g = groups.get(key);
    g.exemptSales += num(r.exempt); g.zeroRatedSales += num(r.zeroRated); g.taxableSales += num(r.vatable); g.outputVat += num(r.outputVat);
  });
  const rows = Array.from(groups.values()).map((g) => ({ ...g, exemptSales: round2(g.exemptSales), zeroRatedSales: round2(g.zeroRatedSales), taxableSales: round2(g.taxableSales), outputVat: round2(g.outputVat) }));
  rows.sort((a, b) => (a.sortKey || "").localeCompare(b.sortKey || ""));
  return rows;
}

function buildSlpRowsFromPurchases(data, month, year) {
  const start = new Date(year, month - 1, 1), end = monthEnd(year, month - 1);
  const relevant = data.purchases.filter((r) => inRange(r.date, start, end));
  const groups = new Map();
  relevant.forEach((r) => {
    const supp = data.suppliers.find((s) => (s.name || "").toLowerCase() === (r.supplier || "").toLowerCase());
    const tin = ((supp && supp.tin) || r.tin || "").replace(/\D/g, "");
    const cleanTin = tin;
    if (!cleanTin) return; // suppliers with no TIN are excluded from SLP
    const key = cleanTin;
    if (!groups.has(key)) {
      const isIndiv = supp && supp.type === "Individual";
      groups.set(key, {
        tin, partyName: isIndiv ? "" : ((supp && (supp.name || supp.registeredName)) || r.supplier || ""),
        surname: isIndiv ? supp.surname : "", firstName: isIndiv ? supp.firstName : "", middleName: isIndiv ? supp.middleName : "",
        address1: (supp && supp.address) || r.address || "", address2: "",
        vatExempt: 0, zeroRated: 0, services: 0, capitalGoods: 0, otherCapitalGoods: 0, inputVat: 0,
        sortKey: reliefPartySortKey(supp, r.supplier),
      });
    }
    const g = groups.get(key);
    g.inputVat += num(r.inputVat);
    const amt = round2(num(r.vatable) + num(r.nonvat));
    const vt = data.vatTypes.find((v) => v.vatType === r.vatType);
    const slspiField = vt ? vt.slspiField : "Services"; // no VAT Type selected yet — default matches prior behavior
    if (slspiField === "Capital goods") g.capitalGoods += amt;
    else if (slspiField === "Other than capital goods") g.otherCapitalGoods += amt;
    else if (slspiField === "Exempt") g.vatExempt += amt;
    else if (slspiField === "Zero-Rated") g.zeroRated += amt;
    else if (slspiField === "Services") g.services += amt;
    // a blank SLSPI field (e.g. "0% Non-VAT purchases") is intentionally excluded from all 5 SLP buckets
  });
  const rows = Array.from(groups.values()).map((g) => ({ ...g, vatExempt: round2(g.vatExempt), zeroRated: round2(g.zeroRated), services: round2(g.services), capitalGoods: round2(g.capitalGoods), otherCapitalGoods: round2(g.otherCapitalGoods), inputVat: round2(g.inputVat) }));
  rows.sort((a, b) => (a.sortKey || "").localeCompare(b.sortKey || ""));
  return rows;
}

function buildSliRowsFromImportation(data, month, year) {
  const start = new Date(year, month - 1, 1), end = monthEnd(year, month - 1);
  const rows = (data.importation || []).filter((r) => inRange(r.date, start, end)).map((r) => ({
    importEntryNo: r.importEntryNo || "", assessmentDate: r.assessmentDate || "", suppliersName: r.supplier || "",
    importationDate: r.date || "", countryOrigin: r.countryOrigin || "", totalLandedCost: num(r.landedCost), otherCharges: num(r.otherCharges),
    exempt: 0, taxable: num(r.dutiableValue), vat: num(r.vatPaid), orNumber: r.orNo || "", vatPaymentDate: r.vatPaymentDate || "",
  }));
  rows.sort((a, b) => (a.suppliersName || "").localeCompare(b.suppliersName || ""));
  return rows;
}

// ATC-tagged Purchase Journal entries are the basis for QAP — one row per transaction, carrying
// its own actual date so no "relative month" guessing is needed like the old upload path required.
function buildQapRowsFromPurchases(data, quarter, year) {
  const startMonth = (quarter - 1) * 3 + 1;
  const endMonth = startMonth + 2;
  const rows = [];
  data.purchases.forEach((r) => {
    if (!r.atc) return;
    const d = parseAppDate(r.date);
    if (!d) return;
    const m = d.getMonth() + 1, y = d.getFullYear();
    if (y !== year || m < startMonth || m > endMonth) return;
    const supp = data.suppliers.find((s) => (s.name || "").toLowerCase() === (r.supplier || "").toLowerCase());
    const isIndiv = supp && supp.type === "Individual";
    const tin = ((supp && supp.tin) || r.tin || "").replace(/\D/g, "");
    rows.push({
      tin, branchCode: "0000",
      partyName: isIndiv ? "" : ((supp && (supp.name || supp.registeredName)) || r.supplier || ""),
      surname: isIndiv ? supp.surname : "", firstName: isIndiv ? supp.firstName : "", middleName: isIndiv ? supp.middleName : "",
      atc: r.atc, rate: num(r.atcRate), taxBase: round2(num(r.vatable) + num(r.nonvat)), ewt: num(r.ewt),
      monthInfo: { month: m, year: y, relative: false },
      sortKey: reliefPartySortKey(supp, r.supplier),
    });
  });
  rows.sort((a, b) => (a.sortKey || "").localeCompare(b.sortKey || "") || a.monthInfo.month - b.monthInfo.month);
  return rows;
}

// Same idea as buildQapRowsFromPurchases, but for Cash Disbursements — only included when the row
// actually has an ATC code AND a resulting EWT (a disbursement can carry an ATC with $0 EWT if the
// rate or amount is blank, which shouldn't feed QAP).
function buildQapRowsFromDisbursements(data, quarter, year) {
  const startMonth = (quarter - 1) * 3 + 1;
  const endMonth = startMonth + 2;
  const rows = [];
  data.disbursements.forEach((r) => {
    if (!r.atc || !(num(r.ewt) > 0)) return;
    const d = parseAppDate(r.date);
    if (!d) return;
    const m = d.getMonth() + 1, y = d.getFullYear();
    if (y !== year || m < startMonth || m > endMonth) return;
    const supp = data.suppliers.find((s) => (s.name || "").toLowerCase() === (r.vendor || "").toLowerCase());
    const isIndiv = supp && supp.type === "Individual";
    const tin = ((supp && supp.tin) || r.tin || "").replace(/\D/g, "");
    rows.push({
      tin, branchCode: "0000",
      partyName: isIndiv ? "" : ((supp && (supp.name || supp.registeredName)) || r.vendor || ""),
      surname: isIndiv ? supp.surname : "", firstName: isIndiv ? supp.firstName : "", middleName: isIndiv ? supp.middleName : "",
      atc: r.atc, rate: num(r.atcRate), taxBase: round2(num(r.amount)), ewt: num(r.ewt),
      monthInfo: { month: m, year: y, relative: false },
      sortKey: reliefPartySortKey(supp, r.vendor),
    });
  });
  rows.sort((a, b) => (a.sortKey || "").localeCompare(b.sortKey || "") || a.monthInfo.month - b.monthInfo.month);
  return rows;
}

// Cash Receipts entries with an ATC and a resulting CWT feed QAP too — the party here is looked
// up in Customers Master (not Suppliers), since "Received From" ties to a customer/payor.
function buildQapRowsFromReceipts(data, quarter, year) {
  const startMonth = (quarter - 1) * 3 + 1;
  const endMonth = startMonth + 2;
  const rows = [];
  data.receipts.forEach((r) => {
    if (!r.atc || !(num(r.cwt) > 0)) return;
    const d = parseAppDate(r.date);
    if (!d) return;
    const m = d.getMonth() + 1, y = d.getFullYear();
    if (y !== year || m < startMonth || m > endMonth) return;
    const cust = data.customers.find((c) => (c.name || "").toLowerCase() === (r.from || "").toLowerCase());
    const isIndiv = cust && cust.type === "Individual";
    const tin = ((cust && cust.tin) || r.tin || "").replace(/\D/g, "");
    rows.push({
      tin, branchCode: "0000",
      partyName: isIndiv ? "" : ((cust && (cust.name || cust.registeredName)) || r.from || ""),
      surname: isIndiv ? cust.surname : "", firstName: isIndiv ? cust.firstName : "", middleName: isIndiv ? cust.middleName : "",
      atc: r.atc, rate: num(r.atcRate), taxBase: round2(num(r.amount)), ewt: num(r.cwt),
      monthInfo: { month: m, year: y, relative: false },
      sortKey: reliefPartySortKey(cust, r.from),
    });
  });
  rows.sort((a, b) => (a.sortKey || "").localeCompare(b.sortKey || "") || a.monthInfo.month - b.monthInfo.month);
  return rows;
}

// SAWT lists taxes withheld FROM this business — exactly the CWT recorded against Cash Receipts
// entries. Same source data as buildQapRowsFromReceipts, but keeps the natural "cwt" field name
// since that's what buildSawtRecord / the SAWT reports expect, and is scoped by month, not quarter.
// SAWT now sources from the 2307 Log — each entry there already carries its own snapshot of the
// customer's name parts and type, so no fresh Customers Master lookup is needed here.
function buildSawtRowsFromForm2307(data, month, year) {
  const start = new Date(year, month - 1, 1), end = monthEnd(year, month - 1);
  const rows = [];
  (data.form2307 || []).forEach((r) => {
    if (!(num(r.cwt) > 0)) return;
    if (!inRange(r.date, start, end)) return;
    const tin = normalizeTin(r.tin);
    if (!tin) return; // only transactions with BOTH a CWT and a TIN belong in SAWT
    const isIndiv = r.custType === "Individual";
    rows.push({
      tin, branchCode: "0000",
      partyName: isIndiv ? "" : (r.customerName || ""),
      surname: isIndiv ? (r.surname || "") : "", firstName: isIndiv ? (r.firstName || "") : "", middleName: isIndiv ? (r.middleName || "") : "",
      atc: r.atc, rate: num(r.atcRate), taxBase: round2(num(r.incomePayment)), cwt: num(r.cwt),
      sortKey: (r.customerName || r.surname || "").toLowerCase(),
    });
  });
  rows.sort((a, b) => (a.sortKey || "").localeCompare(b.sortKey || ""));
  return rows;
}

function SLSPIPage({ data }) {
  const [subMode, setSubMode] = useState("SLP");
  const [month, setMonth] = useState(1);
  const [year, setYear] = useState(2026);
  const cfg = SLSPI_CONFIG[subMode];

  const sourceRows = useMemo(() => {
    if (subMode === "SLP") return buildSlpRowsFromPurchases(data, month, year);
    if (subMode === "SLS") return buildSlsRowsFromSales(data, month, year);
    return buildSliRowsFromImportation(data, month, year);
  }, [data, subMode, month, year]);

  const filingDetails = useMemo(() => buildFilingDetails(data, { month, year, quarter: null, sawtFormType: null }), [data, month, year]);
  const record = sourceRows.length ? buildSlspiRecord(subMode, filingDetails, sourceRows) : null;
  const reportExport = useReportExport(
    () => downloadSlspiPdfReport(subMode, data, sourceRows, month, year),
    () => downloadSlspiExcelReport(subMode, data, sourceRows, month, year)
  );

  const sourceLabel = subMode === "SLP" ? "Purchase Journal" : subMode === "SLS" ? "Sales Journal" : "Importation Ledger";
  const previewName = (r) => subMode === "SLI" ? r.suppliersName : (r.partyName || [r.surname, r.firstName, r.middleName].filter(Boolean).join(" "));
  const previewTin = (r) => subMode === "SLI" ? "" : (r.tin ? reliefFormatTIN(r.tin) : "— no TIN —");
  const previewAmt = (r) => subMode === "SLP" ? r.inputVat : subMode === "SLS" ? r.outputVat : r.vat;

  return (
    <div>
      <SectionHeader icon={FileText} title="SLSPI" subtitle="Summary List of Sales, Purchases & Importation — auto-sourced from your journals, generates the .dat file for BIR eSubmission." />
      <div className="rc-mode-picker">
        {["SLP","SLS","SLI"].map((m) => (
          <button key={m} className={"rc-mode-btn" + (subMode === m ? " active" : "")} onClick={() => setSubMode(m)}>
            <div className="rc-mode-title">{m}</div>
            <div className="rc-mode-sub">{SLSPI_CONFIG[m].title}</div>
          </button>
        ))}
      </div>
      <ReliefCompanyWarning filingDetails={filingDetails} needAddr />
      <div className="form-card">
        <div className="form-grid">
          <LabeledField label="Return Month"><Field type="select" options={MONTHS.map((m, i) => ({ value: i + 1, label: m }))} value={month} onChange={(v) => setMonth(Number(v))} /></LabeledField>
          <LabeledField label="Return Year"><Field type="select" options={getAvailableYears(data)} value={year} onChange={(v) => setYear(Number(v))} /></LabeledField>
        </div>
      </div>

      {sourceRows.length === 0 ? (
        <div className="empty-panel">No {sourceLabel.toLowerCase()} entries found for {MONTHS[month - 1]} {year}.</div>
      ) : (
        <>
          <div className="kpi-row">
            <div className="kpi-card"><div className="kpi-label">Payees (from {sourceLabel})</div><div className="kpi-value">{sourceRows.length}</div></div>
            <div className="kpi-card"><div className="kpi-label">Return Period</div><div className="kpi-value" style={{ fontSize: 16 }}>{record ? record.returnPeriod : "—"}</div></div>
            <div className="kpi-card"><div className="kpi-label">{record ? record.totalLabel : ""}</div><div className="kpi-value">{record ? fmt(record.totalVal) : "—"}</div></div>
          </div>
          <div className="ledger-wrap">
            <div className="table-scroll">
              <table className="ledger-table report-table">
                <thead><tr><th>Payee (alphabetical)</th><th>TIN</th><th className="num-head">{subMode === "SLP" ? "Input VAT" : subMode === "SLS" ? "Output VAT" : "VAT"}</th></tr></thead>
                <tbody>
                  {sourceRows.map((r, i) => (
                    <tr key={i}><td>{previewName(r) || "—"}</td><td>{previewTin(r)}</td><td className="num">{fmt(previewAmt(r))}</td></tr>
                  ))}
                </tbody>
                <tfoot><tr>
                  <td colSpan={2} className="totals-label">Total — {sourceRows.length} payee{sourceRows.length === 1 ? "" : "s"}</td>
                  <td className="num">{fmt(round2(sourceRows.reduce((s, r) => s + num(previewAmt(r)), 0)))}</td>
                </tr></tfoot>
              </table>
            </div>
          </div>
          {record && (
            <div className="ledger-wrap"><div className="table-toolbar">
              <AddRowBtn onClick={() => downloadDatFile(record)}>Download .dat file</AddRowBtn>
              <div className="header-actions">
                <button className="io-btn" onClick={reportExport.exportExcel}><FileDown size={13} /> Excel Report</button>
                <button className="io-btn" onClick={reportExport.exportPDF}><FileDown size={13} /> PDF Report</button>
              </div>
            </div></div>
          )}
          <ExportStatus status={reportExport.status} />
        </>
      )}
      <div className="rc-warn">
        {subMode === "SLP" && "Rows are pulled from the Purchase Journal for the selected month, grouped and summed by supplier (matched against the Suppliers Master by TIN when possible). Suppliers with no TIN on file are excluded from SLP, per BIR requirements. "}
        {subMode === "SLS" && "Rows are pulled from the Sales Journal for the selected month, grouped and summed by customer (matched against the Customers Master by TIN when possible). Customers with no TIN are still included. "}
        {subMode === "SLI" && "Rows are pulled from the Importation Ledger for the selected month, one row per import entry (not aggregated, since each entry has its own Import Entry No. and OR No.). "}
        Listed alphabetically. {cfg.note}
      </div>
    </div>
  );
}

function QAPPage({ data }) {
  const [quarter, setQuarter] = useState(1);
  const [year, setYear] = useState(2026);

  const allRows = useMemo(() => {
    const rows = [...buildQapRowsFromPurchases(data, quarter, year), ...buildQapRowsFromDisbursements(data, quarter, year), ...buildQapRowsFromReceipts(data, quarter, year)];
    rows.sort((a, b) => (a.sortKey || "").localeCompare(b.sortKey || "") || a.monthInfo.month - b.monthInfo.month);
    return rows;
  }, [data, quarter, year]);
  const filingDetails = useMemo(() => buildFilingDetails(data, { month: null, year, quarter, sawtFormType: null }), [data, year, quarter]);
  const groups = allRows.length ? buildQapGroups(filingDetails, allRows) : [];
  const [dlStatus, setDlStatus] = useState(null);

  const onDownload = async () => {
    setDlStatus({ type: "pending", text: "Building ZIP…" });
    try { await downloadQapZip(groups, filingDetails); setDlStatus({ type: "success", text: `Downloaded ${groups.length} file(s) in a ZIP.` }); }
    catch (e) { setDlStatus({ type: "error", text: "Couldn't build the ZIP — check your connection and try again." }); }
    setTimeout(() => setDlStatus(null), 5000);
  };
  const reportExport = useReportExport(
    () => downloadQapPdfReport(data, allRows, quarter, year),
    () => downloadQapExcelReport(data, allRows, quarter, year)
  );

  return (
    <div>
      <SectionHeader icon={FileText} title="QAP" subtitle="Quarterly Alphalist of Payees (BIR Form 1601-EQ attachment) — auto-sourced from ATC-tagged Purchase Journal, Cash Disbursements, and Cash Receipts entries." />
      <ReliefCompanyWarning filingDetails={filingDetails} needAddr={false} />
      <div className="form-card">
        <div className="form-grid">
          <LabeledField label="Quarter"><Field type="select" options={QUARTERS.map((q, i) => ({ value: i + 1, label: q.label }))} value={quarter} onChange={(v) => setQuarter(Number(v))} /></LabeledField>
          <LabeledField label="Year"><Field type="select" options={getAvailableYears(data)} value={year} onChange={(v) => setYear(Number(v))} /></LabeledField>
        </div>
      </div>
      {allRows.length === 0 ? (
        <div className="empty-panel">No ATC-tagged transactions found for {QUARTERS[quarter - 1].label}, {year}. Set an ATC code on the relevant rows in the Purchase Journal, Cash Disbursements, or Cash Receipts (with a resulting EWT/CWT) to have them appear here.</div>
      ) : (
        <>
          <div className="ledger-wrap">
            <div className="table-scroll">
              <table className="ledger-table report-table">
                <thead><tr><th>Payee (alphabetical)</th><th>ATC</th><th>Month</th><th className="num-head">Tax Base</th><th className="num-head">EWT</th></tr></thead>
                <tbody>
                  {allRows.map((r, i) => (
                    <tr key={i}>
                      <td>{r.partyName || [r.surname, r.firstName, r.middleName].filter(Boolean).join(" ") || "—"}</td>
                      <td>{r.atc}</td>
                      <td>{MONTHS[r.monthInfo.month - 1]}</td>
                      <td className="num">{fmt(r.taxBase)}</td>
                      <td className="num">{fmt(r.ewt)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot><tr>
                  <td colSpan={3} className="totals-label">Total — {allRows.length} transaction{allRows.length === 1 ? "" : "s"}</td>
                  <td className="num">{fmt(round2(allRows.reduce((s, r) => s + r.taxBase, 0)))}</td>
                  <td className="num">{fmt(round2(allRows.reduce((s, r) => s + r.ewt, 0)))}</td>
                </tr></tfoot>
              </table>
            </div>
          </div>
          <div className="kpi-row">
            <div className="kpi-card"><div className="kpi-label">Files</div><div className="kpi-value">{groups.length}</div></div>
            <div className="kpi-card"><div className="kpi-label">Total Records</div><div className="kpi-value">{groups.reduce((s, g) => s + g.dRows.length, 0)}</div></div>
            <div className="kpi-card"><div className="kpi-label">Total EWT</div><div className="kpi-value">{fmt(groups.reduce((s, g) => s + g.totals.ewt, 0))}</div></div>
          </div>
          <ExportStatus status={dlStatus} />
          <div className="ledger-wrap"><div className="table-toolbar">
            <AddRowBtn onClick={onDownload}>Download {groups.length} .dat file{groups.length > 1 ? "s" : ""} (.zip)</AddRowBtn>
            <div className="header-actions">
              <button className="io-btn" onClick={reportExport.exportExcel}><FileDown size={13} /> Excel Report</button>
              <button className="io-btn" onClick={reportExport.exportPDF}><FileDown size={13} /> PDF Report</button>
            </div>
          </div></div>
          <ExportStatus status={reportExport.status} />
        </>
      )}
      <div className="rc-warn">Rows are pulled from the Purchase Journal, Cash Disbursements, and Cash Receipts for the selected quarter — any transaction with an ATC code (and a resulting EWT/CWT greater than zero) feeds QAP automatically, grouped by the transaction's actual month within the quarter. Text fields are quoted only when non-blank; per-row branch codes default to 0000 unless set on the Supplier/Customer. Always run each file through the BIR Alphalist Validation Module before submitting.</div>
    </div>
  );
}

function SAWTPage({ data }) {
  const [month, setMonth] = useState(1);
  const [year, setYear] = useState(2026);
  const [formType, setFormType] = useState("1702Q");

  const sourceRows = useMemo(() => buildSawtRowsFromForm2307(data, month, year), [data, month, year]);
  const filingDetails = useMemo(() => buildFilingDetails(data, { month, year, quarter: null, sawtFormType: formType }), [data, month, year, formType]);
  const record = sourceRows.length ? buildSawtRecord(filingDetails, sourceRows) : null;
  const reportExport = useReportExport(
    () => downloadSawtPdfReport(data, sourceRows, month, year, formType),
    () => downloadSawtExcelReport(data, sourceRows, month, year, formType)
  );

  return (
    <div>
      <SectionHeader icon={FileText} title="SAWT" subtitle="Summary Alphalist of Withholding Tax — auto-sourced from Creditable Withholding Taxes entries with CWT withheld from you." />
      <ReliefCompanyWarning filingDetails={filingDetails} needAddr={false} />
      <div className="form-card">
        <div className="form-grid">
          <LabeledField label="SAWT Form Type"><Field type="select" options={[{ value: "1702Q", label: "1702Q — Non-Individual, Quarterly" }, { value: "1701", label: "1701 — Individual, Annual" }, { value: "1701Q", label: "1701Q — Individual, Quarterly" }]} value={formType} onChange={setFormType} /></LabeledField>
          <LabeledField label="Return Month"><Field type="select" options={MONTHS.map((m, i) => ({ value: i + 1, label: m }))} value={month} onChange={(v) => setMonth(Number(v))} /></LabeledField>
          <LabeledField label="Return Year"><Field type="select" options={getAvailableYears(data)} value={year} onChange={(v) => setYear(Number(v))} /></LabeledField>
        </div>
      </div>
      {sourceRows.length === 0 ? (
        <div className="empty-panel">No Creditable Withholding Taxes entries with both a CWT and a TIN found for {MONTHS[month - 1]}, {year}. Add an entry in the Creditable Withholding Taxes tab — it needs a TIN and a resulting CWT greater than zero to appear here.</div>
      ) : (
        <>
          <div className="ledger-wrap">
            <div className="table-scroll">
              <table className="ledger-table report-table">
                <thead><tr><th>Payor (alphabetical)</th><th>ATC</th><th className="num-head">Tax Base</th><th className="num-head">CWT</th></tr></thead>
                <tbody>
                  {sourceRows.map((r, i) => (
                    <tr key={i}>
                      <td>{r.partyName || [r.surname, r.firstName, r.middleName].filter(Boolean).join(" ") || "—"}</td>
                      <td>{r.atc}</td>
                      <td className="num">{fmt(r.taxBase)}</td>
                      <td className="num">{fmt(r.cwt)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot><tr>
                  <td colSpan={2} className="totals-label">Total — {sourceRows.length} payee{sourceRows.length === 1 ? "" : "s"}</td>
                  <td className="num">{fmt(round2(sourceRows.reduce((s, r) => s + r.taxBase, 0)))}</td>
                  <td className="num">{fmt(round2(sourceRows.reduce((s, r) => s + r.cwt, 0)))}</td>
                </tr></tfoot>
              </table>
            </div>
          </div>
          <div className="kpi-row">
            <div className="kpi-card"><div className="kpi-label">Records</div><div className="kpi-value">{record.dRows.length}</div></div>
            <div className="kpi-card"><div className="kpi-label">Return Period</div><div className="kpi-value" style={{ fontSize: 16 }}>{record.returnPeriod}</div></div>
            <div className="kpi-card"><div className="kpi-label">Total CWT</div><div className="kpi-value">{fmt(record.totals.cwt)}</div></div>
          </div>
          <div className="ledger-wrap"><div className="table-toolbar">
            <AddRowBtn onClick={() => downloadDatFile(record)}>Download .dat file</AddRowBtn>
            <div className="header-actions">
              <button className="io-btn" onClick={reportExport.exportExcel}><FileDown size={13} /> Excel Report</button>
              <button className="io-btn" onClick={reportExport.exportPDF}><FileDown size={13} /> PDF Report</button>
            </div>
          </div></div>
          <ExportStatus status={reportExport.status} />
        </>
      )}
      <div className="rc-warn">Rows are pulled from Creditable Withholding Taxes for the selected month — an entry feeds SAWT only when it has both a CWT greater than zero and a TIN, sorted alphabetically. Tax base is the Income Payment amount as entered there, used as-is (no VAT adjustment). 1702Q (non-individual, quarterly) follows a confirmed pattern. 1701 (individual, annual) uses selective quoting. 1701Q (individual, quarterly) is inferred from 1701's pattern and not independently confirmed — always run it through the BIR Alphalist Validation Module before submitting.</div>
    </div>
  );
}

/* ============================== STYLES ============================== */

function Style() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

      .ph-books {
        --paper: #F6F9FD;
        --paper-deep: #E8F1FB;
        --ink: #16233A;
        --ink-soft: #55697F;
        --line: #C7DAF0;
        --line-soft: #DCE8F7;
        --green: #1D5FA8;
        --green-deep: #0F2A4D;
        --gold: #2E86AB;
        --red: #A13D3D;
        --white: #FFFFFF;
        font-family: 'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif;
        color: var(--ink);
        background: var(--paper);
        display: flex;
        min-height: 100%;
        width: 100%;
        position: relative;
      }
      .ph-books * { box-sizing: border-box; }
      .ph-books.boot { align-items: center; justify-content: center; min-height: 480px; }
      .boot-card { font-family: 'Fraunces', serif; font-size: 18px; color: var(--green-deep); }

      /* Sidebar */
      .sidebar {
        width: 248px; flex-shrink: 0; background: linear-gradient(180deg, var(--green-deep), var(--green));
        color: #E8F1FC; display: flex; flex-direction: column; min-height: 100%;
        position: sticky; top: 0; align-self: flex-start; height: 100vh; max-height: 900px;
      }
      .brand { display: flex; align-items: center; gap: 10px; padding: 20px 18px 16px; border-bottom: 1px solid rgba(255,255,255,0.14); }
      .brand-mark { width: 34px; height: 34px; border-radius: 8px; display: flex; align-items: center; justify-content: center; overflow: hidden; flex-shrink: 0; }
      .brand-mark svg { display: block; border-radius: 8px; }
      .brand-title { font-family: 'Fraunces', serif; font-weight: 600; font-size: 16px; letter-spacing: 0.2px; }
      .brand-sub { font-size: 11px; color: #A9C8EC; margin-top: 1px; }
      .nav-close { display: none; background: none; border: none; color: #E8F1FC; margin-left: auto; cursor: pointer; }
      nav { flex: 1; overflow-y: auto; padding: 14px 10px; }
      .nav-group { margin-bottom: 16px; }
      .nav-group-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.09em; color: #8FB4DC; padding: 4px 10px 6px; font-weight: 600; }
      .nav-item { width: 100%; display: flex; align-items: center; gap: 9px; padding: 8px 10px; border-radius: 7px; background: none; border: none; color: #D6E6F8; font-size: 13px; text-align: left; cursor: pointer; margin-bottom: 2px; transition: background 0.12s ease; }
      .nav-item:hover { background: rgba(255,255,255,0.08); }
      .nav-item.active { background: rgba(255,255,255,0.16); color: #FFFFFF; font-weight: 500; }
      .sidebar-foot { padding: 12px 16px 16px; border-top: 1px solid rgba(255,255,255,0.14); }
      .backup-row { display: flex; gap: 8px; margin-bottom: 8px; flex-wrap: wrap; }
      .backup-msg { font-size: 11px; color: #9EC2E8; background: rgba(255,255,255,0.08); border-radius: 5px; padding: 6px 8px; margin-bottom: 8px; }
      .backup-msg.error { color: #FFC9C0; background: rgba(220,90,70,0.18); }
      .backup-row .ghost-btn { flex: 1; justify-content: center; }
      .save-indicator { font-size: 11px; color: #9EC2E8; margin-bottom: 8px; }
      .backup-hint { font-size: 10.5px; color: #7FA3C7; margin-bottom: 6px; line-height: 1.4; }
      .ghost-btn { display: inline-flex; align-items: center; gap: 6px; background: none; border: 1px solid rgba(255,255,255,0.25); color: #DCEBFA; font-size: 11.5px; padding: 6px 10px; border-radius: 6px; cursor: pointer; }
      .ghost-btn:hover { background: rgba(255,255,255,0.08); }
      .settings-tab { margin-bottom: 0; }

      /* Settings page */
      .settings-section { border: 1px solid var(--line); border-radius: 10px; background: var(--white); padding: 18px 20px; margin-bottom: 16px; max-width: 640px; }
      .settings-section-head h3 { font-family: 'Fraunces', serif; font-size: 15px; font-weight: 600; color: var(--green-deep); margin: 0 0 4px; }
      .settings-section-head p { font-size: 12.5px; color: var(--ink-soft); margin: 0 0 14px; line-height: 1.5; }
      .settings-actions { display: flex; flex-wrap: wrap; gap: 10px; }
      .settings-btn { display: inline-flex; align-items: center; gap: 7px; background: var(--white); border: 1px solid var(--line); color: var(--ink); font-size: 13px; font-weight: 500; padding: 9px 14px; border-radius: 7px; cursor: pointer; }
      .settings-btn:hover { background: var(--paper-deep); border-color: var(--green); }
      .settings-btn.danger { color: var(--red); border-color: #E7C9C9; }
      .settings-btn.danger:hover { background: #FBECEC; border-color: var(--red); }
      .content .backup-msg { margin-top: 12px; font-size: 12.5px; color: var(--green-deep); background: var(--paper-deep); border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; }
      .content .backup-msg.error { color: var(--red); background: #FBECEC; border-color: #E7C9C9; }

      /* Client switcher (sidebar, above nav) */
      .client-switch-wrap { padding: 12px 14px; border-bottom: 1px solid rgba(255,255,255,0.14); display: flex; flex-direction: column; gap: 7px; }
      .client-switch-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.09em; color: #8FB4DC; font-weight: 600; }
      .client-switch-row { display: flex; gap: 6px; align-items: center; }
      .io-select { flex: 1; min-width: 0; background: rgba(255,255,255,0.10); color: #FFFFFF; border: 1px solid rgba(255,255,255,0.25); border-radius: 6px; padding: 7px 8px; font-size: 12.5px; font-family: inherit; cursor: pointer; }
      .io-select option { color: #16233A; }
      .client-add-btn { flex-shrink: 0; width: 30px; height: 30px; display: inline-flex; align-items: center; justify-content: center; background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.25); border-radius: 6px; color: #E8F1FC; cursor: pointer; }
      .client-add-btn:hover { background: rgba(255,255,255,0.20); }
      .qf-hint { font-size: 11.5px; color: var(--ink-soft); line-height: 1.45; grid-column: 1 / -1; margin-top: 2px; }

      .scrim { display: none; }

      /* Main */
      .main { flex: 1; min-width: 0; display: flex; flex-direction: column; }
      .topbar { display: none; align-items: center; gap: 10px; padding: 12px 16px; background: var(--white); border-bottom: 1px solid var(--line-soft); position: sticky; top: 0; z-index: 5; }
      .nav-open { background: none; border: none; cursor: pointer; color: var(--green-deep); }
      .topbar-title { font-family: 'Fraunces', serif; font-weight: 600; font-size: 14px; }
      .topbar-tin { font-size: 11px; color: var(--ink-soft); margin-left: auto; font-family: 'IBM Plex Mono', monospace; }
      .content { padding: 32px 40px 60px; max-width: 1240px; width: 100%; margin: 0 auto; }

      .section-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; margin-bottom: 22px; flex-wrap: wrap; }
      .section-icon { width: 38px; height: 38px; border-radius: 9px; background: var(--paper-deep); border: 1px solid var(--line); color: var(--green); display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .section-head h1 { font-family: 'Fraunces', serif; font-size: 22px; font-weight: 600; margin: 0 0 4px; color: var(--green-deep); }
      .section-head p { font-size: 13px; color: var(--ink-soft); margin: 0; max-width: 560px; line-height: 1.5; }

      .callout { display: flex; align-items: center; gap: 14px; background: var(--white); border: 1px solid var(--line); border-left: 3px solid var(--gold); border-radius: 8px; padding: 14px 16px; margin-bottom: 24px; }
      .callout svg { color: var(--gold); flex-shrink: 0; }
      .callout-title { font-weight: 600; font-size: 13.5px; }
      .callout-body { font-size: 12.5px; color: var(--ink-soft); margin-top: 2px; }
      .primary-btn { margin-left: auto; background: var(--green); color: #fff; border: none; padding: 9px 16px; border-radius: 7px; font-size: 13px; font-weight: 500; cursor: pointer; white-space: nowrap; }
      .primary-btn:hover { background: var(--green-deep); }
      .primary-btn.danger { background: var(--red); }
      .primary-btn.danger:hover { background: #7f322f; }
      .primary-btn:disabled { background: var(--line); color: var(--ink-soft); cursor: default; }
      .primary-btn:disabled:hover { background: var(--line); }

      /* Company Details — explicit Save */
      .company-save-bar { display: flex; align-items: center; justify-content: flex-end; gap: 14px; margin-top: 22px; }
      .company-save-bar .primary-btn { margin-left: 0; }
      .save-confirm { display: inline-flex; align-items: center; gap: 5px; color: var(--green); font-size: 13px; font-weight: 600; }
      .save-pending-note { font-size: 12px; color: var(--ink-soft); }
      .signature-upload { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }

      .confirm-box { max-width: 420px; }
      .confirm-body p { margin: 0; font-size: 13px; color: var(--ink); line-height: 1.55; }

      .header-actions { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
      .io-toolbar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
      .io-btn { display: inline-flex; align-items: center; gap: 6px; background: var(--white); border: 1px solid var(--line); color: var(--green-deep); font-size: 12px; font-weight: 500; padding: 7px 12px; border-radius: 7px; cursor: pointer; white-space: nowrap; }
      .io-btn:hover { background: var(--paper-deep); border-color: var(--green); }
      .io-btn.danger { color: var(--red); }
      .io-btn.danger:hover { background: #FBEAE7; border-color: var(--red); }
      .io-btn.accent { background: var(--green); color: #fff; border-color: var(--green); }
      .io-btn.accent:hover { background: var(--green-deep); }
      .io-select { background: var(--white); border: 1px solid var(--line); color: var(--ink); font-size: 12px; padding: 7px 10px; border-radius: 7px; }
      .period-filter-bar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
      .row-check { width: 15px; height: 15px; accent-color: var(--green); cursor: pointer; }
      tr.row-selected, .jv-card.row-selected { background: var(--paper-deep); }
      .selection-bar { display: flex; align-items: center; gap: 12px; background: var(--paper-deep); border: 1px solid var(--line); border-left: 3px solid var(--green); border-radius: 8px; padding: 9px 14px; margin-bottom: 10px; font-size: 12.5px; color: var(--green-deep); font-weight: 500; flex-wrap: wrap; }
      .selection-actions { display: flex; gap: 8px; margin-left: 4px; flex-wrap: wrap; }
      .selection-bar > .io-btn { margin-left: auto; }
      .gj-filter-bar { justify-content: space-between; }
      .gj-select-all { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-soft); cursor: pointer; }
      .qf-row { grid-column: 1 / -1; display: flex; align-items: center; gap: 12px; }
      .qf-check { display: flex; align-items: center; gap: 7px; font-size: 12.5px; min-width: 170px; color: var(--ink); }
      .qf-check input { width: 15px; height: 15px; accent-color: var(--green); cursor: pointer; }
      .qf-row .ledger-field, .qf-row select { flex: 1; border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; }
      .qf-hint { grid-column: 1 / -1; font-size: 11.5px; color: var(--ink-soft); background: var(--paper-deep); border-radius: 6px; padding: 8px 10px; }
      .qf-error { grid-column: 1 / -1; font-size: 12px; color: var(--red); background: #FBEAE7; border: 1px solid #F2CFC9; border-radius: 6px; padding: 8px 10px; }

      .td-controls { display: flex; align-items: flex-end; gap: 16px; flex-wrap: wrap; background: var(--white); border: 1px solid var(--line); border-radius: 10px; padding: 14px 16px; margin-bottom: 18px; }
      .td-field { display: flex; flex-direction: column; gap: 5px; min-width: 260px; }
      .td-field label { font-size: 11.5px; color: var(--ink-soft); font-weight: 500; }
      .td-field .ledger-field { border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; }
      .td-kpi-row { margin-bottom: 16px; }
      .td-locked { opacity: 0.5; font-size: 12px; }
      .td-tag { font-size: 10.5px; padding: 3px 8px; border-radius: 20px; font-weight: 600; }
      .td-tag.reclass { background: #E3F0FC; color: var(--green-deep); }
      .td-tag.system { background: var(--paper-deep); color: var(--ink-soft); }

      .combo-wrap { position: relative; width: 100%; }
      .combo-input { width: 100%; cursor: text; }
      .combo-menu { background: var(--white); border: 1px solid var(--line); border-radius: 8px; box-shadow: 0 12px 32px rgba(15,42,77,0.18); max-height: 240px; overflow-y: auto; z-index: 200; padding: 4px; }
      .combo-option { padding: 7px 10px; font-size: 12.5px; cursor: pointer; border-radius: 5px; color: var(--ink); }
      .combo-add-new { color: var(--green-deep); font-weight: 600; border-top: 1px solid var(--line-soft); margin-top: 2px; padding-top: 8px; }
      .combo-option:hover, .combo-option.active { background: var(--paper-deep); }
      .combo-option.selected { font-weight: 600; color: var(--green-deep); }
      .combo-empty { padding: 8px 10px; font-size: 12px; color: var(--ink-soft); }
      .import-msg { margin: -10px 0 18px; padding: 9px 14px; border-radius: 8px; background: var(--paper-deep); color: var(--green-deep); font-size: 12.5px; border: 1px solid var(--line); }
      .import-msg.error { background: #FBEAE7; color: var(--red); border-color: #F2CFC9; }

      .modal-overlay { position: fixed; inset: 0; background: rgba(15,42,77,0.35); display: flex; align-items: center; justify-content: center; z-index: 60; padding: 20px; }
      .modal-box { background: var(--white); border-radius: 12px; max-width: 620px; width: 100%; max-height: 86vh; overflow-y: auto; box-shadow: 0 24px 64px rgba(15,42,77,0.28); }
      .modal-head { display: flex; align-items: center; justify-content: space-between; padding: 18px 22px; border-bottom: 1px solid var(--line); position: sticky; top: 0; background: var(--white); }
      .modal-head h2 { font-family: 'Fraunces', serif; font-size: 17px; font-weight: 600; color: var(--green-deep); margin: 0; }
      .modal-close { background: none; border: none; cursor: pointer; color: var(--ink-soft); padding: 4px; border-radius: 6px; }
      .modal-close:hover { background: var(--paper-deep); }
      .modal-body { padding: 20px 22px; display: grid; grid-template-columns: 1fr 1fr; gap: 14px 18px; }
      .modal-foot { display: flex; justify-content: flex-end; gap: 10px; padding: 14px 22px; border-top: 1px solid var(--line); position: sticky; bottom: 0; background: var(--white); }
      .btn-outline { background: var(--white); border: 1px solid var(--line); color: var(--ink); padding: 9px 18px; border-radius: 7px; font-size: 13px; font-weight: 500; cursor: pointer; }
      .btn-outline:hover { background: var(--paper-deep); }
      .computed-preview { grid-column: 1 / -1; display: flex; gap: 22px; flex-wrap: wrap; background: var(--paper-deep); border: 1px solid var(--line); border-radius: 8px; padding: 10px 14px; }
      .computed-preview-item { font-size: 12.5px; }
      .computed-preview-item span { color: var(--ink-soft); margin-right: 6px; }
      .computed-preview-item strong { font-family: 'IBM Plex Mono', monospace; color: var(--green-deep); }
      .modal-lines { grid-column: 1 / -1; }
      .modal-lines .ledger-table { margin-bottom: 8px; }

      .kpi-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 22px; }
      .kpi-card { background: var(--white); border: 1px solid var(--line); border-radius: 10px; padding: 16px 18px; }
      .kpi-label { font-size: 11.5px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 6px; }
      .kpi-value { font-family: 'IBM Plex Mono', monospace; font-size: 22px; font-weight: 500; color: var(--green-deep); }
      .kpi-value.neg { color: var(--red); }

      .chart-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
      .chart-card { background: var(--white); border: 1px solid var(--line); border-radius: 10px; padding: 16px 18px; }
      .chart-card.wide { grid-column: 1 / -1; }
      .chart-title { font-size: 12.5px; font-weight: 600; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 10px; }
      .chart-empty { display: flex; align-items: center; justify-content: center; height: 220px; color: var(--ink-soft); font-size: 12.5px; }

      /* Forms */
      .form-card { background: var(--white); border: 1px solid var(--line); border-radius: 10px; padding: 22px 24px; }
      .gl-info-card { margin-bottom: 18px; padding: 16px 20px; }
      .gl-info-card .form-grid { gap: 12px 20px; }
      .gl-group-header td { background: var(--paper-deep); font-weight: 700; font-family: 'Fraunces', serif; font-size: 13.5px; color: var(--green-deep); padding: 10px; border-top: 2px solid var(--line); border-bottom: 1px solid var(--line); }
      .gl-open-close td { font-weight: 600; background: rgba(31,95,168,0.045); border-bottom: 1px solid var(--line-soft); color: var(--ink); }
      .gl-open-close.closing td { border-bottom: 2px solid var(--line); }
      .ml-compact th, .ml-compact td { font-size: 11px; padding: 6px 8px; white-space: nowrap; }
      .ml-compact td.num, .ml-compact th.num-head { font-family: 'IBM Plex Mono', monospace; letter-spacing: -0.2px; }
      .ml-compact thead th { font-size: 10px; padding: 8px 8px; }
      .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px 20px; }
      .form-field { display: flex; flex-direction: column; gap: 5px; }
      .form-field.wide { grid-column: 1 / -1; }
      .form-field label { font-size: 11.5px; color: var(--ink-soft); font-weight: 500; }
      .form-field .ledger-field { border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; }

      /* Tables */
      .sub-block-title { font-family: 'Fraunces', serif; font-weight: 600; font-size: 15px; color: var(--green-deep); margin-bottom: 10px; }
      .section-head-note { font-size: 12.5px; color: var(--ink-soft); margin: -18px 0 16px; max-width: 620px; line-height: 1.5; }
      .rc-mode-picker { display: flex; gap: 12px; margin-bottom: 18px; flex-wrap: wrap; }
      .rc-mode-btn { flex: 1; min-width: 180px; background: var(--white); border: 1.5px solid var(--line); border-radius: 10px; padding: 14px 16px; cursor: pointer; text-align: left; box-shadow: none; transition: border-color .15s ease, background .15s ease, transform .12s ease; }
      .rc-mode-btn:hover { border-color: var(--green); transform: translateY(-1px); }
      .rc-mode-btn.active { border-color: var(--green); background: var(--paper-deep); }
      .rc-mode-title { font-family: 'Fraunces', serif; font-size: 16px; font-weight: 600; color: var(--green-deep); margin-bottom: 2px; }
      .rc-mode-sub { font-size: 11.5px; color: var(--ink-soft); }
      .rc-upload-card { margin-top: 14px; padding: 4px 10px 12px; }
      .rc-note { font-size: 11.5px; color: var(--ink-soft); padding: 0 6px; margin: 6px 0 2px; line-height: 1.5; }
      .rc-warn { font-size: 12px; color: var(--ink-soft); background: var(--white); border: 1px solid var(--line); border-left: 3px solid var(--gold); border-radius: 8px; padding: 12px 14px; margin-top: 18px; line-height: 1.6; }
      .ledger-wrap { background: var(--white); border: 1px solid var(--line); border-radius: 10px; padding: 4px; margin-bottom: 8px; }
      .table-toolbar { display: flex; justify-content: space-between; align-items: center; padding: 6px 6px 2px; flex-wrap: wrap; gap: 10px; }
      .table-toolbar .add-row-btn { margin: 0; }
      .toolbar-left { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; flex: 1; }
      .search-bar { display: flex; align-items: center; gap: 6px; border: 1px solid var(--line); border-radius: 7px; padding: 6px 10px; background: var(--white); color: var(--ink-soft); min-width: 220px; }
      .search-bar input { border: none; outline: none; background: transparent; font-size: 12.5px; color: var(--ink); flex: 1; font-family: inherit; }
      .search-bar input::placeholder { color: var(--ink-soft); }
      .search-clear { background: none; border: none; cursor: pointer; color: var(--ink-soft); display: flex; padding: 2px; border-radius: 4px; }
      .search-clear:hover { background: var(--paper-deep); color: var(--ink); }
      .gj-search-bar-wrap { margin-bottom: 14px; max-width: 420px; }
      .gj-search-bar-wrap .search-bar { width: 100%; }
      .gj-search-results { background: var(--white); border: 1px solid var(--line); border-radius: 10px; overflow: hidden; margin-bottom: 12px; }
      .gj-search-row { display: grid; grid-template-columns: 110px 110px 1fr 130px; gap: 10px; align-items: center; padding: 10px 14px; border-bottom: 1px solid var(--line-soft); cursor: pointer; font-size: 12.5px; }
      .gj-search-row:last-child { border-bottom: none; }
      .gj-search-row:hover { background: var(--paper-deep); }
      .gj-search-jvno { font-weight: 600; color: var(--green-deep); }
      .gj-search-date { color: var(--ink-soft); font-family: 'IBM Plex Mono', monospace; }
      .gj-search-particulars { color: var(--ink); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .gj-search-amount { text-align: right; font-family: 'IBM Plex Mono', monospace; color: var(--green-deep); font-weight: 600; }
      .table-scroll { overflow-x: auto; border-radius: 7px; }
      table.ledger-table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
      .ledger-table thead th { background: var(--paper-deep); color: var(--green-deep); font-weight: 600; text-align: left; padding: 9px 10px; border-bottom: 2px solid var(--line); position: sticky; top: 0; white-space: nowrap; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.02em; }
      .ledger-table thead th.num-head { text-align: right; }
      .ledger-table tbody td { border-bottom: 1px solid var(--line-soft); padding: 2px 4px; vertical-align: middle; }
      .ledger-table tbody tr:hover { background: #FAFCFF; }
      .ledger-table tbody tr:nth-child(even) { background: rgba(31,92,76,0.025); }
      .ledger-table tfoot td { padding: 9px 10px; border-top: 2px solid var(--line); font-weight: 600; background: var(--paper-deep); }
      .ledger-table td.num, .ledger-table th.num-head, td.num { font-family: 'IBM Plex Mono', monospace; text-align: right; }
      .totals-label { text-align: right; color: var(--ink-soft); }
      .empty-row, .empty-panel { text-align: center; color: var(--ink-soft); font-size: 12.5px; padding: 26px 10px; }
      .empty-panel { background: var(--white); border: 1px dashed var(--line); border-radius: 10px; }

      .ledger-field { width: 100%; background: transparent; border: 1px solid transparent; border-radius: 5px; padding: 7px 8px; font-size: 12.5px; font-family: inherit; color: var(--ink); outline: none; }
      .ledger-field:hover { border-color: var(--line); }
      .ledger-field:focus { border-color: var(--green); background: var(--paper-deep); }
      .ledger-field.num { font-family: 'IBM Plex Mono', monospace; text-align: right; }
      .ledger-field.center { text-align: center; }
      .ledger-field:disabled { opacity: 0.45; }
      .ledger-read { padding: 7px 8px; font-family: 'IBM Plex Mono', monospace; color: var(--ink-soft); font-size: 12.5px; }
      .ledger-read.num { text-align: right; }
      .ledger-read.center { text-align: center; }

      .add-row-btn { display: inline-flex; align-items: center; gap: 6px; background: none; border: 1px dashed var(--line); color: var(--green); font-size: 12.5px; font-weight: 500; padding: 8px 14px; border-radius: 7px; cursor: pointer; margin: 6px 2px 2px; }
      .add-row-btn:hover { background: var(--paper-deep); border-color: var(--green); }
      .del-btn { background: none; border: none; color: #B9877E; cursor: pointer; padding: 4px; border-radius: 5px; display: inline-flex; }
      .del-btn:hover { background: #FBEAE7; color: var(--red); }

      .year-select, .period-picker select { border: 1px solid var(--line); background: var(--white); border-radius: 6px; padding: 7px 10px; font-size: 12.5px; color: var(--ink); font-family: inherit; }
      .period-picker { display: flex; gap: 8px; }

      /* JV cards */
      .jv-card { background: var(--white); border: 1px solid var(--line); border-radius: 10px; padding: 14px; margin-bottom: 16px; }
      .jv-head { display: grid; grid-template-columns: 20px 110px 130px 1fr auto auto; gap: 8px; align-items: center; margin-bottom: 8px; }
      .jv-status { font-size: 11px; padding: 4px 9px; border-radius: 20px; background: #FBEAE7; color: var(--red); font-weight: 600; white-space: nowrap; }
      .jv-status.ok { background: #E3F0FC; color: var(--green); }

      .check-strip { margin-top: 10px; padding: 10px 14px; border-radius: 8px; font-size: 12.5px; font-weight: 500; }
      .check-strip.ok { background: #E3F0FC; color: var(--green-deep); }
      .check-strip.bad { background: #FBEAE7; color: var(--red); }

      .inv-log { display: flex; align-items: center; gap: 10px; background: var(--white); border: 1px solid var(--line); border-radius: 8px; padding: 10px 14px; margin-bottom: 16px; font-size: 12.5px; flex-wrap: wrap; }
      .inv-log span { color: var(--ink-soft); }
      .inv-log .ledger-field { width: 120px; border: 1px solid var(--line); border-radius: 5px; }

      .stmt-table td { padding: 7px 10px; }
      .stmt-h td { font-weight: 700; color: var(--green-deep); padding-top: 14px; border-bottom: 1px solid var(--line); font-family: 'Fraunces', serif; font-size: 13.5px; }
      .stmt-sub td { font-weight: 600; border-top: 1px solid var(--line-soft); }
      .stmt-final td { font-weight: 700; font-size: 14px; border-top: 2px solid var(--green); color: var(--green-deep); padding-top: 10px; }
      .dim { color: var(--ink-soft); font-style: italic; }

      .bs-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }

      /* Profit & Loss — dark report skin */
      .pnl-shell { --pnl-bg-1: #0B1220; --pnl-bg-2: #121a2e; --pnl-bg-3: #1a2338; --pnl-border: #2a3550;
        --pnl-text: #E8EDF7; --pnl-text-dim: #8b9ab3; --pnl-link: #6AA3FF; --pnl-accent: #3B82F6;
        border-radius: 12px; overflow: hidden; font-family: 'IBM Plex Sans', sans-serif; box-shadow: 0 8px 28px rgba(11,18,32,0.18); }
      .pnl-toolbar { background: var(--pnl-bg-1); color: var(--pnl-text); padding: 14px 18px 16px; }
      .pnl-toolbar-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
      .pnl-toolbar-top h1 { font-family: 'Fraunces', serif; font-size: 16px; font-weight: 600; margin: 0; color: #fff; }
      .pnl-badge { display: inline-flex; align-items: center; gap: 6px; background: rgba(255,255,255,0.06); border: 1px solid var(--pnl-border); color: var(--pnl-text); font-size: 11.5px; padding: 6px 10px; border-radius: 6px; }
      .pnl-daterange-label { font-size: 11.5px; color: var(--pnl-text-dim); margin-bottom: 10px; }
      .pnl-daterange-label strong { color: var(--pnl-link); font-weight: 500; }
      .pnl-controls-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
      .pnl-date-input { background: var(--pnl-bg-2); border: 1px solid var(--pnl-border); color: var(--pnl-text); font-size: 12px; padding: 7px 10px; border-radius: 6px; font-family: 'IBM Plex Mono', monospace; }
      .pnl-select { background: var(--pnl-bg-2); border: 1px solid var(--pnl-border); color: var(--pnl-text); font-size: 12px; padding: 7px 10px; border-radius: 6px; }
      .pnl-toolbar-actions { display: flex; align-items: center; gap: 8px; margin-left: auto; }
      .pnl-icon-btn { background: var(--pnl-bg-2); border: 1px solid var(--pnl-border); color: var(--pnl-text-dim); width: 30px; height: 30px; border-radius: 6px; display: inline-flex; align-items: center; justify-content: center; cursor: pointer; }
      .pnl-icon-btn:hover { color: var(--pnl-text); border-color: var(--pnl-link); }
      .pnl-refresh-btn { display: inline-flex; align-items: center; gap: 6px; background: var(--pnl-accent); border: none; color: #fff; font-size: 12.5px; font-weight: 500; padding: 8px 14px; border-radius: 6px; cursor: pointer; }
      .pnl-refresh-btn:hover { background: #2f6fe0; }
      .pnl-card { background: var(--pnl-bg-2); padding: 18px 22px 10px; }
      .pnl-card-head { padding-bottom: 14px; margin-bottom: 6px; border-bottom: 1px solid var(--pnl-border); }
      .pnl-card-title { font-family: 'Fraunces', serif; font-size: 15px; font-weight: 600; color: #fff; }
      .pnl-company { font-size: 12.5px; color: var(--pnl-link); margin-top: 4px; }
      .pnl-currency { font-size: 11.5px; color: var(--pnl-text-dim); margin-top: 2px; }
      .pnl-period { font-size: 11.5px; color: var(--pnl-text-dim); margin-top: 2px; }
      table.pnl-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
      .pnl-table td { padding: 7px 4px; color: var(--pnl-text); border-bottom: 1px solid rgba(255,255,255,0.04); }
      .pnl-table td.num { text-align: right; font-family: 'IBM Plex Mono', monospace; }
      .pnl-table tr.pnl-section td { background: var(--pnl-bg-3); font-weight: 600; color: #fff; padding: 8px 10px; border-radius: 4px; border-bottom: none; }
      .pnl-table tr.pnl-line td:first-child { color: var(--pnl-link); padding-left: 10px; }
      .pnl-table tr.pnl-total td { font-weight: 600; border-top: 1px solid var(--pnl-border); border-bottom: none; padding-top: 9px; padding-bottom: 9px; }
      .pnl-table tr.pnl-grand td { font-weight: 700; font-size: 13.5px; border-top: 2px solid var(--pnl-accent); border-bottom: none; padding-top: 10px; padding-bottom: 10px; color: #fff; }
      .pnl-table tr.pnl-final td { border-top: 2px solid #22c55e; }
      .pnl-input-cell { display: flex; align-items: center; justify-content: flex-end; gap: 2px; }
      .pnl-input-cell .ledger-field { width: 120px; background: var(--pnl-bg-3); border: 1px solid var(--pnl-border); border-radius: 5px; color: var(--pnl-text); font-family: 'IBM Plex Mono', monospace; }
      .pnl-input-cell .ledger-field:focus { background: #202b45; }
      .pnl-table .dim { color: var(--pnl-text-dim); font-style: italic; }
      .pnl-bs-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }
      @media (max-width: 900px) { .pnl-bs-grid { grid-template-columns: 1fr; } }
      @media (max-width: 900px) {
        .pnl-toolbar-top { flex-wrap: wrap; gap: 8px; }
        .pnl-controls-row { flex-direction: column; align-items: stretch; }
        .pnl-toolbar-actions { margin-left: 0; justify-content: flex-end; }
        .pnl-card { padding: 16px 14px 10px; }
      }

      @media (max-width: 900px) {
        .sidebar { position: fixed; left: -260px; top: 0; height: 100%; z-index: 30; transition: left 0.2s ease; width: 250px; box-shadow: 4px 0 24px rgba(0,0,0,0.2); }
        .sidebar.open { left: 0; }
        .nav-close { display: block; }
        .scrim { display: block; position: fixed; inset: 0; background: rgba(0,0,0,0.35); z-index: 25; }
        .topbar { display: flex; }
        .content { padding: 20px 14px 50px; }
        .form-grid { grid-template-columns: 1fr; }
        .modal-body { grid-template-columns: 1fr; }
        .kpi-row { grid-template-columns: 1fr; }
        .chart-grid { grid-template-columns: 1fr; }
        .bs-grid { grid-template-columns: 1fr; }
        .jv-head { grid-template-columns: 20px 1fr 1fr; grid-template-areas: "cb a b" "cb c c" "cb d e"; }
        .gj-search-row { grid-template-columns: 1fr 1fr; grid-template-rows: auto auto; }
        .gj-search-amount { text-align: left; }
      }
    `}</style>
  );
}

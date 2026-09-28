// src/utils/analytics.js

function toNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const cleaned = String(value)
    .replace(/₦/g, "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim();

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : 0;
}

function getValue(row, keys) {
  for (const key of keys) {
    if (
      row &&
      row[key] !== undefined &&
      row[key] !== null &&
      String(row[key]).trim() !== ""
    ) {
      return row[key];
    }
  }

  return "";
}

export function getLoanAmount(row) {
  return toNumber(
    getValue(row, [
      "loanAmount",
      "Loan Amount",
      "LoanAmount",
      "loan_amount",
      "amount",
      "Amount",
    ])
  );
}

export function getPrincipalBalance(row) {
  return toNumber(
    getValue(row, [
      "principalBalance",
      "Principal Bal.",
      "Principal Balance",
      "PrincipalBal",
      "principal_bal",
      "outstandingBalance",
      "Outstanding Balance",
    ])
  );
}

export function getAccountOfficer(row) {
  return (
    getValue(row, [
      "accountOfficer",
      "Account Officer",
      "AccountOfficer",
      "account_officer",
      "officer",
      "Officer",
      "loanOfficer",
      "Loan Officer",
    ]) || "Unassigned"
  );
}

export function getProduct(row) {
  return (
    getValue(row, [
      "product",
      "Product",
      "loanProduct",
      "Loan Product",
      "loan_product",
    ]) || "Unknown"
  );
}

export function getCustomerName(row) {
  return (
    getValue(row, [
      "customerName",
      "Customer Name",
      "CustomerName",
      "customer_name",
      "name",
      "Name",
    ]) || "Unknown Customer"
  );
}

export function parseDate(value) {
  if (!value) {
    return null;
  }

  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  // Excel serial date
  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));

    const date = new Date(
      excelEpoch.getTime() +
        value * 24 * 60 * 60 * 1000
    );

    return Number.isNaN(date.getTime()) ? null : date;
  }

  const stringValue = String(value).trim();

  if (!stringValue) {
    return null;
  }

  // Standard JavaScript date
  const normalDate = new Date(stringValue);

  if (!Number.isNaN(normalDate.getTime())) {
    return normalDate;
  }

  // DD/MM/YYYY
  const slashMatch = stringValue.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
  );

  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]) - 1;
    const year = Number(slashMatch[3]);

    const date = new Date(year, month, day);

    if (!Number.isNaN(date.getTime())) {
      return date;
    }
  }

  // DD-MM-YYYY
  const dashMatch = stringValue.match(
    /^(\d{1,2})-(\d{1,2})-(\d{4})$/
  );

  if (dashMatch) {
    const day = Number(dashMatch[1]);
    const month = Number(dashMatch[2]) - 1;
    const year = Number(dashMatch[3]);

    const date = new Date(year, month, day);

    if (!Number.isNaN(date.getTime())) {
      return date;
    }
  }

  return null;
}

export function calculateKPIs(data = []) {
  if (!Array.isArray(data)) {
    return {
      totalLoans: 0,
      totalDisbursed: 0,
      totalPrincipalBalance: 0,
      averageLoanSize: 0,
    };
  }

  const totalLoans = data.length;

  const totalDisbursed = data.reduce(
    (total, row) => total + getLoanAmount(row),
    0
  );

  const totalPrincipalBalance = data.reduce(
    (total, row) => total + getPrincipalBalance(row),
    0
  );

  const averageLoanSize =
    totalLoans > 0
      ? totalDisbursed / totalLoans
      : 0;

  return {
    totalLoans,
    totalDisbursed,
    totalPrincipalBalance,
    averageLoanSize,
  };
}

export function portfolioByOfficer(data = []) {
  if (!Array.isArray(data)) {
    return [];
  }

  const grouped = {};

  data.forEach((row) => {
    const officer = getAccountOfficer(row);

    const loanAmount = getLoanAmount(row);

    const principalBalance =
      getPrincipalBalance(row);

    if (!grouped[officer]) {
      grouped[officer] = {
        name: officer,
        officer,
        count: 0,
        loanCount: 0,
        amount: 0,
        loanAmount: 0,
        principalBalance: 0,
      };
    }

    grouped[officer].count += 1;
    grouped[officer].loanCount += 1;

    grouped[officer].amount += loanAmount;
    grouped[officer].loanAmount += loanAmount;

    grouped[officer].principalBalance +=
      principalBalance;
  });

  return Object.values(grouped).sort(
    (a, b) => b.loanAmount - a.loanAmount
  );
}

export function productDistribution(data = []) {
  if (!Array.isArray(data)) {
    return [];
  }

  const grouped = {};

  data.forEach((row) => {
    const product = getProduct(row);

    const loanAmount = getLoanAmount(row);

    if (!grouped[product]) {
      grouped[product] = {
        name: product,
        product,
        count: 0,
        loanCount: 0,
        amount: 0,
        loanAmount: 0,
      };
    }

    grouped[product].count += 1;
    grouped[product].loanCount += 1;

    grouped[product].amount += loanAmount;
    grouped[product].loanAmount += loanAmount;
  });

  return Object.values(grouped).sort(
    (a, b) => b.loanAmount - a.loanAmount
  );
}

export { toNumber };
export function loanTrendByMonth(data = []) {
  if (!Array.isArray(data)) {
    return [];
  }

  const grouped = {};

  data.forEach((row) => {
    const rawDate =
      row.disbursementDate ??
      row["Disbursement Date"] ??
      row.DisbursementDate ??
      row.disbursement_date;

    const date = parseDate(rawDate);

    if (!date) return;

    const year = date.getFullYear();
    const month = date.getMonth();

    const key = `${year}-${String(month + 1).padStart(2, "0")}`;

    if (!grouped[key]) {
      grouped[key] = {
        key,
        month: date.toLocaleString("en-US", {
          month: "short",
          year: "numeric",
        }),
        count: 0,
        loanCount: 0,
        amount: 0,
        loanAmount: 0,
      };
    }

    const amount = getLoanAmount(row);

    grouped[key].count += 1;
    grouped[key].loanCount += 1;
    grouped[key].amount += amount;
    grouped[key].loanAmount += amount;
  });

  return Object.values(grouped).sort((a, b) =>
    a.key.localeCompare(b.key)
  );
}
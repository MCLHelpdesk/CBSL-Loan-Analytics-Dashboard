import React, { useMemo, useState } from "react";
import { useMsal } from "@azure/msal-react";
import FileUpload from "./component/FileUpload";

/* =========================================================
   DATA HELPERS
========================================================= */

const getValue = (row, keys) => {
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
};

const toNumber = (value) => {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const cleaned = String(value)
    .replace(/₦/g, "")
    .replace(/,/g, "")
    .replace(/\s/g, "")
    .replace(/%/g, "")
    .trim();

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : 0;
};

const parseDate = (value) => {
  if (!value) return null;

  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  // Excel serial date
  if (typeof value === "number") {
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(
      excelEpoch.getTime() + value * 24 * 60 * 60 * 1000
    );

    return Number.isNaN(date.getTime()) ? null : date;
  }

  const text = String(value).trim();

  if (!text) return null;

  // YYYY-MM-DD
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) {
    const [year, month, day] = text.split("-").map(Number);
    const date = new Date(year, month - 1, day);

    return Number.isNaN(date.getTime()) ? null : date;
  }

  // DD/MM/YYYY
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(text)) {
    const [day, month, year] = text.split("/").map(Number);
    const date = new Date(year, month - 1, day);

    return Number.isNaN(date.getTime()) ? null : date;
  }

  // DD-MM-YYYY
  if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(text)) {
    const [day, month, year] = text.split("-").map(Number);
    const date = new Date(year, month - 1, day);

    return Number.isNaN(date.getTime()) ? null : date;
  }

  const parsed = new Date(text);

  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const formatDate = (value) => {
  const date = parseDate(value);

  if (!date) return "—";

  return date.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
};

const formatCurrency = (value) => {
  return new Intl.NumberFormat("en-NG", {
    style: "currency",
    currency: "NGN",
    maximumFractionDigits: 0,
  }).format(toNumber(value));
};

const formatCompactCurrency = (value) => {
  const number = toNumber(value);

  if (number >= 1_000_000_000) {
    return `₦${(number / 1_000_000_000).toFixed(1)}B`;
  }

  if (number >= 1_000_000) {
    return `₦${(number / 1_000_000).toFixed(1)}M`;
  }

  if (number >= 1_000) {
    return `₦${(number / 1_000).toFixed(1)}K`;
  }

  return formatCurrency(number);
};

const normalizeText = (value) => {
  return String(value || "").trim();
};

/* =========================================================
   EXCEL FIELD DEFINITIONS
========================================================= */

const LOAN_AMOUNT_KEYS = [
  "loanAmount",
  "Loan Amount",
  "LoanAmount",
  "loan_amount",
  "amountDisbursed",
  "Amount Disbursed",
  "Disbursed Amount",
  "disbursedAmount",
  "Amount",
];

const PRINCIPAL_KEYS = [
  "principalBalance",
  "Principal Bal.",
  "Principal Bal",
  "Principal Balance",
  "PrincipalBal",
  "principal_bal",
  "outstandingBalance",
  "Outstanding Balance",
  "Outstanding Principal",
  "outstandingPrincipal",
];

const OFFICER_KEYS = [
  "accountOfficer",
  "Account Officer",
  "AccountOfficer",
  "account_officer",
  "Loan Officer",
  "loanOfficer",
  "Officer",
  "officer",
];

const PRODUCT_KEYS = [
  "product",
  "Product",
  "loanProduct",
  "Loan Product",
  "loan_product",
  "Product Name",
];

const CUSTOMER_KEYS = [
  "customerName",
  "Customer Name",
  "CustomerName",
  "customer_name",
  "Customer",
  "Name",
];

const ACCOUNT_KEYS = [
  "accountNo",
  "Account No.",
  "Account No",
  "Account Number",
  "AccountNumber",
  "accountNumber",
];

const PHONE_KEYS = [
  "phoneNo",
  "Phone No.",
  "Phone No",
  "Phone",
  "phone",
  "phoneNumber",
  "Phone Number",
];

const BVN_KEYS = [
  "BVN",
  "bvn",
  "Bvn",
];

const GENDER_KEYS = [
  "Gender",
  "gender",
];

const ADDRESS_KEYS = [
  "Address",
  "address",
];

const DISBURSEMENT_KEYS = [
  "disbursementDate",
  "Disbursement Date",
  "DisbursementDate",
  "disbursement_date",
  "Date Disbursed",
  "Disbursed Date",
];

const MATURITY_KEYS = [
  "maturationDate",
  "Maturation Date",
  "Maturity Date",
  "MaturityDate",
  "maturityDate",
  "maturity_date",
];

const EFFECTIVE_KEYS = [
  "effectiveDate",
  "Effective Date",
  "EffectiveDate",
  "effective_date",
];

const REPAID_KEYS = [
  "amountRepaid",
  "Amount Repaid",
  "Repaid Amount",
  "Amount Repaid.",
  "repaidAmount",
  "amount_repaid",
  "totalRepaid",
  "Total Repaid",
];

/* =========================================================
   MAIN COMPONENT
========================================================= */

export default function Dashboard() {
  const { instance, accounts } = useMsal();

  const [loanData, setLoanData] = useState([]);
  const [activeSection, setActiveSection] = useState("dashboard");

  const [portfolioSearch, setPortfolioSearch] = useState("");
  const [portfolioOfficerFilter, setPortfolioOfficerFilter] = useState("all");
  const [portfolioProductFilter, setPortfolioProductFilter] = useState("all");

  const [selectedOfficer, setSelectedOfficer] = useState(null);

  /* =========================================================
     FIELD ACCESSORS
  ========================================================= */

  const getLoanAmount = (row) =>
    toNumber(getValue(row, LOAN_AMOUNT_KEYS));

  const getPrincipalBalance = (row) =>
    toNumber(getValue(row, PRINCIPAL_KEYS));

  const getOfficer = (row) =>
    normalizeText(getValue(row, OFFICER_KEYS)) || "Unassigned";

  const getProduct = (row) =>
    normalizeText(getValue(row, PRODUCT_KEYS)) || "Unclassified";

  const getCustomer = (row) =>
    normalizeText(getValue(row, CUSTOMER_KEYS)) || "Unknown Customer";

  const getAccount = (row) =>
    normalizeText(getValue(row, ACCOUNT_KEYS)) || "—";

  const getPhone = (row) =>
    normalizeText(getValue(row, PHONE_KEYS)) || "—";

  const getBVN = (row) =>
    normalizeText(getValue(row, BVN_KEYS)) || "—";

  const getGender = (row) =>
    normalizeText(getValue(row, GENDER_KEYS)) || "—";

  const getAddress = (row) =>
    normalizeText(getValue(row, ADDRESS_KEYS)) || "—";

  const getDisbursementDate = (row) =>
    parseDate(getValue(row, DISBURSEMENT_KEYS));

  const getMaturityDate = (row) =>
    parseDate(getValue(row, MATURITY_KEYS));

  const getEffectiveDate = (row) =>
    parseDate(getValue(row, EFFECTIVE_KEYS));

  const getRepaidAmount = (row) => {
    const explicitRepaid = getValue(row, REPAID_KEYS);

    if (
      explicitRepaid !== undefined &&
      explicitRepaid !== null &&
      String(explicitRepaid).trim() !== ""
    ) {
      return Math.max(toNumber(explicitRepaid), 0);
    }

    return Math.max(
      getLoanAmount(row) - getPrincipalBalance(row),
      0
    );
  };

  /* =========================================================
     FILE LOADING
  ========================================================= */

  const handleDataLoaded = (data) => {
    if (!Array.isArray(data)) {
      setLoanData([]);
      return;
    }

    setLoanData(data);
    setActiveSection("dashboard");
  };

  /* =========================================================
     GLOBAL ANALYTICS
  ========================================================= */

  const kpis = useMemo(() => {
    const totalLoans = loanData.length;

    const totalDisbursed = loanData.reduce(
      (sum, row) => sum + getLoanAmount(row),
      0
    );

    const outstandingPrincipal = loanData.reduce(
      (sum, row) => sum + getPrincipalBalance(row),
      0
    );

    const amountRepaid = loanData.reduce(
      (sum, row) => sum + getRepaidAmount(row),
      0
    );

    const repaymentRate =
      totalDisbursed > 0
        ? (amountRepaid / totalDisbursed) * 100
        : 0;

    const officers = new Set(
      loanData
        .map(getOfficer)
        .filter((officer) => officer && officer !== "Unassigned")
    );

    const products = new Set(
      loanData
        .map(getProduct)
        .filter((product) => product && product !== "Unclassified")
    );

    return {
      totalLoans,
      totalDisbursed,
      outstandingPrincipal,
      amountRepaid,
      repaymentRate,
      accountOfficers: officers.size,
      products: products.size,
    };
  }, [loanData]);

  /* =========================================================
     MONTHLY TREND
  ========================================================= */

  const monthlyTrend = useMemo(() => {
    const groups = {};

    loanData.forEach((row) => {
      const date = getDisbursementDate(row);

      if (!date) return;

      const year = date.getFullYear();
      const month = date.getMonth();

      const key = `${year}-${String(month + 1).padStart(2, "0")}`;

      if (!groups[key]) {
        groups[key] = {
          key,
          year,
          month,
          label: date.toLocaleDateString("en-US", {
            month: "short",
            year: "numeric",
          }),
          count: 0,
          amount: 0,
        };
      }

      groups[key].count += 1;
      groups[key].amount += getLoanAmount(row);
    });

    return Object.values(groups).sort((a, b) =>
      a.key.localeCompare(b.key)
    );
  }, [loanData]);

  /* =========================================================
     OFFICER SUMMARY
  ========================================================= */

  const officerData = useMemo(() => {
    const groups = {};

    loanData.forEach((row) => {
      const officer = getOfficer(row);

      if (!groups[officer]) {
        groups[officer] = {
          name: officer,
          count: 0,
          amount: 0,
          outstanding: 0,
          repaid: 0,
        };
      }

      groups[officer].count += 1;
      groups[officer].amount += getLoanAmount(row);
      groups[officer].outstanding += getPrincipalBalance(row);
      groups[officer].repaid += getRepaidAmount(row);
    });

    return Object.values(groups)
      .map((item) => ({
        ...item,
        average:
          item.count > 0 ? item.amount / item.count : 0,
        repaymentRate:
          item.amount > 0
            ? (item.repaid / item.amount) * 100
            : 0,
      }))
      .sort((a, b) => b.amount - a.amount);
  }, [loanData]);

  /* =========================================================
     PRODUCT SUMMARY
  ========================================================= */

  const productData = useMemo(() => {
    const groups = {};

    loanData.forEach((row) => {
      const product = getProduct(row);

      if (!groups[product]) {
        groups[product] = {
          name: product,
          count: 0,
          amount: 0,
          outstanding: 0,
        };
      }

      groups[product].count += 1;
      groups[product].amount += getLoanAmount(row);
      groups[product].outstanding += getPrincipalBalance(row);
    });

    return Object.values(groups).sort(
      (a, b) => b.amount - a.amount
    );
  }, [loanData]);

  /* =========================================================
     MONTHLY WEEKLY ANALYSIS
  ========================================================= */

  const monthlyWeeklyData = useMemo(() => {
    const groups = {};

    loanData.forEach((row) => {
      const date = getDisbursementDate(row);

      if (!date) return;

      const year = date.getFullYear();
      const month = date.getMonth();
      const day = date.getDate();

      const key = `${year}-${String(month + 1).padStart(2, "0")}`;

      if (!groups[key]) {
        groups[key] = {
          key,
          year,
          month,
          monthName: date.toLocaleDateString("en-US", {
            month: "long",
          }),
          monthLabel: date.toLocaleDateString("en-US", {
            month: "short",
            year: "numeric",
          }),
          weeks: [
            {
              week: 1,
              startDay: 1,
              endDay: 7,
              count: 0,
              amount: 0,
            },
            {
              week: 2,
              startDay: 8,
              endDay: 14,
              count: 0,
              amount: 0,
            },
            {
              week: 3,
              startDay: 15,
              endDay: 21,
              count: 0,
              amount: 0,
            },
            {
              week: 4,
              startDay: 22,
              endDay: 28,
              count: 0,
              amount: 0,
            },
            {
              week: 5,
              startDay: 29,
              endDay: new Date(
                year,
                month + 1,
                0
              ).getDate(),
              count: 0,
              amount: 0,
            },
          ],
        };
      }

      let week;

      if (day <= 7) week = 1;
      else if (day <= 14) week = 2;
      else if (day <= 21) week = 3;
      else if (day <= 28) week = 4;
      else week = 5;

      const targetWeek = groups[key].weeks.find(
        (item) => item.week === week
      );

      if (targetWeek) {
        targetWeek.count += 1;
        targetWeek.amount += getLoanAmount(row);
      }
    });

    return Object.values(groups).sort((a, b) =>
      a.key.localeCompare(b.key)
    );
  }, [loanData]);

  /* =========================================================
     OFFICER PERFORMANCE
  ========================================================= */

  const officerPerformance = useMemo(() => {
    const officers = {};

    loanData.forEach((row) => {
      const officer = getOfficer(row);
      const date = getDisbursementDate(row);

      if (!officers[officer]) {
        officers[officer] = {
          name: officer,
          totalLoans: 0,
          totalAmount: 0,
          outstanding: 0,
          repaid: 0,
          months: {},
        };
      }

      officers[officer].totalLoans += 1;
      officers[officer].totalAmount += getLoanAmount(row);
      officers[officer].outstanding += getPrincipalBalance(row);
      officers[officer].repaid += getRepaidAmount(row);

      if (!date) return;

      const year = date.getFullYear();
      const month = date.getMonth();
      const day = date.getDate();

      const key = `${year}-${String(month + 1).padStart(2, "0")}`;

      if (!officers[officer].months[key]) {
        officers[officer].months[key] = {
          key,
          monthLabel: date.toLocaleDateString("en-US", {
            month: "short",
            year: "numeric",
          }),
          weeks: [
            {
              week: 1,
              count: 0,
              amount: 0,
            },
            {
              week: 2,
              count: 0,
              amount: 0,
            },
            {
              week: 3,
              count: 0,
              amount: 0,
            },
            {
              week: 4,
              count: 0,
              amount: 0,
            },
            {
              week: 5,
              count: 0,
              amount: 0,
            },
          ],
        };
      }

      let week;

      if (day <= 7) week = 1;
      else if (day <= 14) week = 2;
      else if (day <= 21) week = 3;
      else if (day <= 28) week = 4;
      else week = 5;

      const targetWeek =
        officers[officer].months[key].weeks.find(
          (item) => item.week === week
        );

      if (targetWeek) {
        targetWeek.count += 1;
        targetWeek.amount += getLoanAmount(row);
      }
    });

    return Object.values(officers)
      .map((officer) => ({
        ...officer,
        averageLoan:
          officer.totalLoans > 0
            ? officer.totalAmount / officer.totalLoans
            : 0,
        repaymentRate:
          officer.totalAmount > 0
            ? (officer.repaid / officer.totalAmount) * 100
            : 0,
        months: Object.values(officer.months).sort((a, b) =>
          a.key.localeCompare(b.key)
        ),
      }))
      .sort((a, b) => b.totalAmount - a.totalAmount);
  }, [loanData]);

  /* =========================================================
     RECENT LOANS
  ========================================================= */

  const recentLoans = useMemo(() => {
    return [...loanData]
      .sort((a, b) => {
        const dateA = getDisbursementDate(a);
        const dateB = getDisbursementDate(b);

        return (
          (dateB?.getTime() || 0) -
          (dateA?.getTime() || 0)
        );
      })
      .slice(0, 8);
  }, [loanData]);

  /* =========================================================
     PORTFOLIO FILTERS
  ========================================================= */

  const officerOptions = useMemo(() => {
    return [...new Set(loanData.map(getOfficer))].sort();
  }, [loanData]);

  const productOptions = useMemo(() => {
    return [...new Set(loanData.map(getProduct))].sort();
  }, [loanData]);

  const filteredPortfolio = useMemo(() => {
    const search = portfolioSearch.trim().toLowerCase();

    return [...loanData]
      .filter((row) => {
        if (
          portfolioOfficerFilter !== "all" &&
          getOfficer(row) !== portfolioOfficerFilter
        ) {
          return false;
        }

        if (
          portfolioProductFilter !== "all" &&
          getProduct(row) !== portfolioProductFilter
        ) {
          return false;
        }

        if (!search) return true;

        const searchableText = [
          getCustomer(row),
          getAccount(row),
          getOfficer(row),
          getProduct(row),
          getPhone(row),
          getBVN(row),
          getAddress(row),
        ]
          .join(" ")
          .toLowerCase();

        return searchableText.includes(search);
      })
      .sort((a, b) => {
        const dateA = getDisbursementDate(a);
        const dateB = getDisbursementDate(b);

        return (
          (dateB?.getTime() || 0) -
          (dateA?.getTime() || 0)
        );
      });
  }, [
    loanData,
    portfolioSearch,
    portfolioOfficerFilter,
    portfolioProductFilter,
  ]);

  const filteredPortfolioStats = useMemo(() => {
    const totalDisbursed = filteredPortfolio.reduce(
      (sum, row) => sum + getLoanAmount(row),
      0
    );

    const outstanding = filteredPortfolio.reduce(
      (sum, row) => sum + getPrincipalBalance(row),
      0
    );

    const repaid = filteredPortfolio.reduce(
      (sum, row) => sum + getRepaidAmount(row),
      0
    );

    return {
      count: filteredPortfolio.length,
      totalDisbursed,
      outstanding,
      repaid,
      repaymentRate:
        totalDisbursed > 0
          ? (repaid / totalDisbursed) * 100
          : 0,
    };
  }, [filteredPortfolio]);

  /* =========================================================
     LOGOUT
  ========================================================= */

  const logout = () => {
    instance.logoutRedirect();
  };

  /* =========================================================
     NAVIGATION
  ========================================================= */

  const menuItems = [
    {
      id: "dashboard",
      icon: "▦",
      label: "Dashboard",
    },
    {
      id: "portfolio",
      icon: "◫",
      label: "Loan Portfolio",
    },
    {
      id: "officers",
      icon: "♙",
      label: "Account Officers",
    },
    {
      id: "reports",
      icon: "▤",
      label: "Reports",
    },
  ];

  const navigateTo = (section) => {
    setActiveSection(section);

    window.scrollTo({
      top: 0,
      behavior: "smooth",
    });
  };

  /* =========================================================
     EMPTY STATE
  ========================================================= */

  const EmptyData = ({
    title = "No data available",
    message = "Upload your Excel loan portfolio to begin.",
  }) => (
    <div className="empty-state">
      <div className="empty-icon">▧</div>
      <h3>{title}</h3>
      <p>{message}</p>
    </div>
  );

  /* =========================================================
     KPI CARD
  ========================================================= */

  const KpiCard = ({
    title,
    value,
    subtitle,
    icon,
    accent,
  }) => (
    <div className={`kpi-card ${accent || ""}`}>
      <div className="kpi-top">
        <span className="kpi-title">{title}</span>
        <span className="kpi-icon">{icon}</span>
      </div>

      <div className="kpi-value">{value}</div>

      {subtitle && (
        <div className="kpi-subtitle">{subtitle}</div>
      )}
    </div>
  );

  /* =========================================================
     BAR COMPONENT
  ========================================================= */

  const HorizontalBar = ({
    label,
    value,
    max,
    amount,
    color = "#2563eb",
  }) => {
    const percentage =
      max > 0 ? Math.max((value / max) * 100, 3) : 0;

    return (
      <div className="bar-row">
        <div className="bar-label">
          <span title={label}>{label}</span>
          <strong>{amount}</strong>
        </div>

        <div className="bar-track">
          <div
            className="bar-fill"
            style={{
              width: `${percentage}%`,
              background: color,
            }}
          />
        </div>
      </div>
    );
  };

  /* =========================================================
     DASHBOARD OVERVIEW
  ========================================================= */

  const DashboardOverview = () => (
    <>
      <section className="upload-panel">
        <div>
          <div className="eyebrow">DATA IMPORT</div>
          <h2>Upload Loan Portfolio</h2>
          <p>
            Upload your Excel portfolio to refresh all
            dashboard analytics.
          </p>
        </div>

        <div className="upload-area">
          <FileUpload onDataLoaded={handleDataLoaded} />
        </div>
      </section>

      <section className="kpi-grid">
        <KpiCard
          title="Total Loans"
          value={kpis.totalLoans.toLocaleString()}
          subtitle="Loan accounts in portfolio"
          icon="▣"
          accent="blue"
        />

        <KpiCard
          title="Total Disbursed"
          value={formatCompactCurrency(
            kpis.totalDisbursed
          )}
          subtitle="Gross loan disbursement"
          icon="₦"
          accent="orange"
        />

        <KpiCard
          title="Amount Repaid"
          value={formatCompactCurrency(
            kpis.amountRepaid
          )}
          subtitle={`${kpis.repaymentRate.toFixed(
            1
          )}% portfolio repayment`}
          icon="✓"
          accent="green"
        />

        <KpiCard
          title="Outstanding Principal"
          value={formatCompactCurrency(
            kpis.outstandingPrincipal
          )}
          subtitle="Current principal balance"
          icon="◈"
          accent="red"
        />

        <KpiCard
          title="Account Officers"
          value={kpis.accountOfficers.toLocaleString()}
          subtitle={`${kpis.products} loan products`}
          icon="♙"
          accent="purple"
        />
      </section>

      <section className="analytics-grid">
        <div className="panel large-panel">
          <div className="panel-header">
            <div>
              <span className="eyebrow">DISBURSEMENT</span>
              <h3>Loan Disbursement Trend</h3>
            </div>

            <div className="panel-badge">
              {monthlyTrend.length} periods
            </div>
          </div>

          {monthlyTrend.length === 0 ? (
            <EmptyData />
          ) : (
            <div className="trend-chart">
              {monthlyTrend.map((item) => {
                const maxAmount = Math.max(
                  ...monthlyTrend.map((x) => x.amount),
                  1
                );

                const height =
                  (item.amount / maxAmount) * 100;

                return (
                  <div
                    className="trend-column"
                    key={item.key}
                  >
                    <div className="trend-value">
                      {formatCompactCurrency(item.amount)}
                    </div>

                    <div className="trend-bar-area">
                      <div
                        className="trend-bar"
                        style={{
                          height: `${Math.max(
                            height,
                            4
                          )}%`,
                        }}
                      />
                    </div>

                    <div className="trend-label">
                      {item.label}
                    </div>

                    <div className="trend-count">
                      {item.count} loans
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="panel">
          <div className="panel-header">
            <div>
              <span className="eyebrow">OFFICER ANALYSIS</span>
              <h3>Portfolio by Officer</h3>
            </div>

            <button
              className="text-button"
              onClick={() => navigateTo("officers")}
            >
              View all →
            </button>
          </div>

          {officerData.length === 0 ? (
            <EmptyData />
          ) : (
            <div className="bar-list">
              {officerData.slice(0, 7).map((item) => {
                const max = Math.max(
                  ...officerData.map((x) => x.amount),
                  1
                );

                return (
                  <HorizontalBar
                    key={item.name}
                    label={item.name}
                    value={item.amount}
                    max={max}
                    amount={formatCompactCurrency(
                      item.amount
                    )}
                  />
                );
              })}
            </div>
          )}
        </div>
      </section>

      <section className="panel weekly-panel">
        <div className="panel-header">
          <div>
            <span className="eyebrow">WEEKLY ANALYTICS</span>
            <h3>Loan Disbursement by Week</h3>
            <p>
              Week 1 = days 1–7 · Week 2 = 8–14 · Week 3 =
              15–21 · Week 4 = 22–28 · Week 5 = 29–end
            </p>
          </div>
        </div>

        {monthlyWeeklyData.length === 0 ? (
          <EmptyData />
        ) : (
          <div className="monthly-weekly-list">
            {monthlyWeeklyData.map((month) => (
              <div
                className="monthly-week-block"
                key={month.key}
              >
                <div className="month-title">
                  <strong>{month.monthLabel}</strong>

                  <span>
                    {month.weeks.reduce(
                      (sum, week) => sum + week.count,
                      0
                    )}{" "}
                    loans
                  </span>
                </div>

                <div className="week-grid">
                  {month.weeks.map((week) => {
                    const maxMonthAmount = Math.max(
                      ...month.weeks.map(
                        (x) => x.amount
                      ),
                      1
                    );

                    const height =
                      (week.amount /
                        maxMonthAmount) *
                      100;

                    return (
                      <div
                        className="week-card"
                        key={`${month.key}-${week.week}`}
                      >
                        <div className="week-number">
                          W{week.week}
                        </div>

                        <div className="week-chart">
                          <div
                            className="week-bar"
                            style={{
                              height: `${Math.max(
                                height,
                                week.amount > 0
                                  ? 5
                                  : 1
                              )}%`,
                            }}
                          />
                        </div>

                        <strong>
                          {formatCompactCurrency(
                            week.amount
                          )}
                        </strong>

                        <span>
                          {week.count} loan
                          {week.count !== 1
                            ? "s"
                            : ""}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        )}
      </section>

      <section className="analytics-grid">
        <div className="panel">
          <div className="panel-header">
            <div>
              <span className="eyebrow">PRODUCT MIX</span>
              <h3>Product Distribution</h3>
            </div>
          </div>

          {productData.length === 0 ? (
            <EmptyData />
          ) : (
            <div className="bar-list">
              {productData.slice(0, 8).map((item) => {
                const max = Math.max(
                  ...productData.map((x) => x.amount),
                  1
                );

                return (
                  <HorizontalBar
                    key={item.name}
                    label={item.name}
                    value={item.amount}
                    max={max}
                    amount={formatCompactCurrency(
                      item.amount
                    )}
                    color="#7c3aed"
                  />
                );
              })}
            </div>
          )}
        </div>

        <div className="panel">
          <div className="panel-header">
            <div>
              <span className="eyebrow">PORTFOLIO SNAPSHOT</span>
              <h3>Portfolio Composition</h3>
            </div>
          </div>

          <div className="composition">
            <div className="composition-row">
              <span>Gross Disbursed</span>
              <strong>
                {formatCurrency(kpis.totalDisbursed)}
              </strong>
            </div>

            <div className="composition-row">
              <span>Amount Repaid</span>
              <strong className="green-text">
                {formatCurrency(kpis.amountRepaid)}
              </strong>
            </div>

            <div className="composition-row">
              <span>Outstanding Principal</span>
              <strong className="red-text">
                {formatCurrency(
                  kpis.outstandingPrincipal
                )}
              </strong>
            </div>

            <div className="composition-row">
              <span>Repayment Rate</span>
              <strong>
                {kpis.repaymentRate.toFixed(1)}%
              </strong>
            </div>
          </div>
        </div>
      </section>

      <section className="panel">
        <div className="panel-header">
          <div>
            <span className="eyebrow">RECENT ACTIVITY</span>
            <h3>Recent Loan Disbursements</h3>
          </div>

          <button
            className="text-button"
            onClick={() => navigateTo("portfolio")}
          >
            View portfolio →
          </button>
        </div>

        {recentLoans.length === 0 ? (
          <EmptyData />
        ) : (
          <div className="table-wrapper">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Customer</th>
                  <th>Account No.</th>
                  <th>Officer</th>
                  <th>Product</th>
                  <th>Loan Amount</th>
                  <th>Principal Balance</th>
                  <th>Disbursement Date</th>
                </tr>
              </thead>

              <tbody>
                {recentLoans.map((row, index) => (
                  <tr key={index}>
                    <td>
                      <strong>{getCustomer(row)}</strong>
                    </td>

                    <td>{getAccount(row)}</td>

                    <td>{getOfficer(row)}</td>

                    <td>
                      <span className="table-tag">
                        {getProduct(row)}
                      </span>
                    </td>

                    <td className="money-cell">
                      {formatCurrency(
                        getLoanAmount(row)
                      )}
                    </td>

                    <td className="money-cell">
                      {formatCurrency(
                        getPrincipalBalance(row)
                      )}
                    </td>

                    <td>
                      {formatDate(
                        getValue(
                          row,
                          DISBURSEMENT_KEYS
                        )
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </>
  );

  /* =========================================================
     LOAN PORTFOLIO PAGE
  ========================================================= */

  const LoanPortfolioPage = () => (
    <>
      <section className="page-intro">
        <div>
          <span className="eyebrow">PORTFOLIO MANAGEMENT</span>
          <h1>Loan Portfolio</h1>
          <p>
            Search, filter and review every loan account in
            the uploaded portfolio.
          </p>
        </div>

        <div className="record-count">
          {filteredPortfolio.length.toLocaleString()} records
        </div>
      </section>

      <section className="kpi-grid portfolio-kpis">
        <KpiCard
          title="Filtered Loans"
          value={filteredPortfolioStats.count.toLocaleString()}
          subtitle="Current filtered records"
          icon="▣"
          accent="blue"
        />

        <KpiCard
          title="Filtered Disbursed"
          value={formatCompactCurrency(
            filteredPortfolioStats.totalDisbursed
          )}
          subtitle="Gross amount"
          icon="₦"
          accent="orange"
        />

        <KpiCard
          title="Outstanding"
          value={formatCompactCurrency(
            filteredPortfolioStats.outstanding
          )}
          subtitle="Principal balance"
          icon="◈"
          accent="red"
        />

        <KpiCard
          title="Repaid"
          value={formatCompactCurrency(
            filteredPortfolioStats.repaid
          )}
          subtitle={`${filteredPortfolioStats.repaymentRate.toFixed(
            1
          )}% repayment rate`}
          icon="✓"
          accent="green"
        />
      </section>

      <section className="panel filters-panel">
        <div className="filters-header">
          <div>
            <span className="eyebrow">SEARCH & FILTER</span>
            <h3>Portfolio Filters</h3>
          </div>

          <button
            className="clear-button"
            onClick={() => {
              setPortfolioSearch("");
              setPortfolioOfficerFilter("all");
              setPortfolioProductFilter("all");
            }}
          >
            Clear filters
          </button>
        </div>

        <div className="filter-grid">
          <div className="filter-field search-field">
            <label>Search portfolio</label>

            <div className="search-input-wrapper">
              <span>⌕</span>

              <input
                type="text"
                placeholder="Customer, account, officer, product, phone or BVN..."
                value={portfolioSearch}
                onChange={(e) =>
                  setPortfolioSearch(e.target.value)
                }
              />
            </div>
          </div>

          <div className="filter-field">
            <label>Account Officer</label>

            <select
              value={portfolioOfficerFilter}
              onChange={(e) =>
                setPortfolioOfficerFilter(
                  e.target.value
                )
              }
            >
              <option value="all">All Officers</option>

              {officerOptions.map((officer) => (
                <option key={officer} value={officer}>
                  {officer}
                </option>
              ))}
            </select>
          </div>

          <div className="filter-field">
            <label>Loan Product</label>

            <select
              value={portfolioProductFilter}
              onChange={(e) =>
                setPortfolioProductFilter(
                  e.target.value
                )
              }
            >
              <option value="all">All Products</option>

              {productOptions.map((product) => (
                <option key={product} value={product}>
                  {product}
                </option>
              ))}
            </select>
          </div>
        </div>
      </section>

      <section className="panel portfolio-table-panel">
        <div className="panel-header">
          <div>
            <span className="eyebrow">LOAN ACCOUNTS</span>
            <h3>Portfolio Register</h3>
          </div>

          <span className="panel-badge">
            {filteredPortfolio.length.toLocaleString()} results
          </span>
        </div>

        {filteredPortfolio.length === 0 ? (
          <EmptyData
            title="No matching loans"
            message="Try changing your search or filter criteria."
          />
        ) : (
          <div className="table-wrapper portfolio-table-wrapper">
            <table className="data-table portfolio-table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Customer</th>
                  <th>Account No.</th>
                  <th>Officer</th>
                  <th>Product</th>
                  <th>Loan Amount</th>
                  <th>Principal Bal.</th>
                  <th>Repaid</th>
                  <th>Repayment %</th>
                  <th>Disbursement</th>
                  <th>Maturity</th>
                </tr>
              </thead>

              <tbody>
                {filteredPortfolio.map(
                  (row, index) => {
                    const loanAmount =
                      getLoanAmount(row);

                    const principal =
                      getPrincipalBalance(row);

                    const repaid =
                      getRepaidAmount(row);

                    const repaymentRate =
                      loanAmount > 0
                        ? (repaid / loanAmount) * 100
                        : 0;

                    return (
                      <tr key={index}>
                        <td className="row-number">
                          {index + 1}
                        </td>

                        <td>
                          <div className="customer-cell">
                            <strong>
                              {getCustomer(row)}
                            </strong>

                            <small>
                              {getGender(row)}
                            </small>
                          </div>
                        </td>

                        <td>{getAccount(row)}</td>

                        <td>
                          {getOfficer(row)}
                        </td>

                        <td>
                          <span className="table-tag">
                            {getProduct(row)}
                          </span>
                        </td>

                        <td className="money-cell">
                          {formatCurrency(
                            loanAmount
                          )}
                        </td>

                        <td className="money-cell">
                          {formatCurrency(
                            principal
                          )}
                        </td>

                        <td className="money-cell green-text">
                          {formatCurrency(repaid)}
                        </td>

                        <td>
                          <div className="progress-cell">
                            <span>
                              {repaymentRate.toFixed(
                                1
                              )}
                              %
                            </span>

                            <div className="mini-progress">
                              <div
                                style={{
                                  width: `${Math.min(
                                    repaymentRate,
                                    100
                                  )}%`,
                                }}
                              />
                            </div>
                          </div>
                        </td>

                        <td>
                          {formatDate(
                            getValue(
                              row,
                              DISBURSEMENT_KEYS
                            )
                          )}
                        </td>

                        <td>
                          {formatDate(
                            getValue(
                              row,
                              MATURITY_KEYS
                            )
                          )}
                        </td>
                      </tr>
                    );
                  }
                )}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </>
  );

  /* =========================================================
     OFFICER PERFORMANCE CARD
  ========================================================= */

  const OfficerPerformanceCard = ({
    officer,
  }) => {
    const isOpen =
      selectedOfficer === officer.name;

    const maxMonthlyAmount = Math.max(
      ...officer.months.flatMap((month) =>
        month.weeks.map((week) => week.amount)
      ),
      1
    );

    return (
      <div
        className={`officer-card ${
          isOpen ? "officer-card-open" : ""
        }`}
      >
        <button
          className="officer-card-header"
          onClick={() =>
            setSelectedOfficer(
              isOpen ? null : officer.name
            )
          }
        >
          <div className="officer-identity">
            <div className="officer-avatar">
              {officer.name
                .split(" ")
                .filter(Boolean)
                .slice(0, 2)
                .map((part) =>
                  part[0]?.toUpperCase()
                )
                .join("") || "U"}
            </div>

            <div>
              <strong>{officer.name}</strong>
              <span>
                {officer.totalLoans} loan
                {officer.totalLoans !== 1
                  ? "s"
                  : ""}{" "}
                managed
              </span>
            </div>
          </div>

          <div className="officer-header-stats">
            <div>
              <span>Disbursed</span>
              <strong>
                {formatCompactCurrency(
                  officer.totalAmount
                )}
              </strong>
            </div>

            <div>
              <span>Outstanding</span>
              <strong>
                {formatCompactCurrency(
                  officer.outstanding
                )}
              </strong>
            </div>

            <div>
              <span>Repayment</span>
              <strong>
                {officer.repaymentRate.toFixed(1)}%
              </strong>
            </div>

            <span className="expand-icon">
              {isOpen ? "⌃" : "⌄"}
            </span>
          </div>
        </button>

        {isOpen && (
          <div className="officer-card-body">
            <div className="officer-summary-grid">
              <div>
                <span>Total Loans</span>
                <strong>
                  {officer.totalLoans.toLocaleString()}
                </strong>
              </div>

              <div>
                <span>Total Disbursed</span>
                <strong>
                  {formatCurrency(
                    officer.totalAmount
                  )}
                </strong>
              </div>

              <div>
                <span>Average Loan</span>
                <strong>
                  {formatCurrency(
                    officer.averageLoan
                  )}
                </strong>
              </div>

              <div>
                <span>Amount Repaid</span>
                <strong className="green-text">
                  {formatCurrency(
                    officer.repaid
                  )}
                </strong>
              </div>
            </div>

            <div className="officer-weekly-title">
              <div>
                <span className="eyebrow">
                  WEEKLY PERFORMANCE
                </span>
                <h4>
                  {officer.name} — Loan Disbursement
                </h4>
              </div>
            </div>

            {officer.months.length === 0 ? (
              <div className="small-empty">
                No valid disbursement dates found.
              </div>
            ) : (
              <div className="officer-month-list">
                {officer.months.map((month) => (
                  <div
                    className="officer-month"
                    key={month.key}
                  >
                    <div className="officer-month-heading">
                      <strong>
                        {month.monthLabel}
                      </strong>

                      <span>
                        {month.weeks.reduce(
                          (sum, week) =>
                            sum + week.count,
                          0
                        )}{" "}
                        loans
                      </span>
                    </div>

                    <div className="officer-week-chart">
                      {month.weeks.map((week) => {
                        const height =
                          (week.amount /
                            maxMonthlyAmount) *
                          100;

                        return (
                          <div
                            className="officer-week-column"
                            key={`${month.key}-${week.week}`}
                          >
                            <div className="officer-week-value">
                              {week.amount > 0
                                ? formatCompactCurrency(
                                    week.amount
                                  )
                                : "—"}
                            </div>

                            <div className="officer-week-track">
                              <div
                                className="officer-week-fill"
                                style={{
                                  height: `${Math.max(
                                    height,
                                    week.amount > 0
                                      ? 6
                                      : 1
                                  )}%`,
                                }}
                              />
                            </div>

                            <strong>
                              W{week.week}
                            </strong>

                            <span>
                              {week.count} loan
                              {week.count !== 1
                                ? "s"
                                : ""}
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  /* =========================================================
     ACCOUNT OFFICERS PAGE
  ========================================================= */

  const AccountOfficersPage = () => (
    <>
      <section className="page-intro">
        <div>
          <span className="eyebrow">TEAM PERFORMANCE</span>
          <h1>Account Officers</h1>
          <p>
            Analyse loan disbursement, repayment and weekly
            performance by account officer.
          </p>
        </div>

        <div className="record-count">
          {officerPerformance.length} officers
        </div>
      </section>

      <section className="kpi-grid">
        <KpiCard
          title="Total Officers"
          value={officerPerformance.length.toLocaleString()}
          subtitle="Officers represented in portfolio"
          icon="♙"
          accent="purple"
        />

        <KpiCard
          title="Loans Managed"
          value={kpis.totalLoans.toLocaleString()}
          subtitle="Total loan accounts"
          icon="▣"
          accent="blue"
        />

        <KpiCard
          title="Total Disbursed"
          value={formatCompactCurrency(
            kpis.totalDisbursed
          )}
          subtitle="Gross officer portfolio"
          icon="₦"
          accent="orange"
        />

        <KpiCard
          title="Average / Officer"
          value={formatCompactCurrency(
            officerPerformance.length > 0
              ? kpis.totalDisbursed /
                  officerPerformance.length
              : 0
          )}
          subtitle="Average portfolio value"
          icon="◈"
          accent="green"
        />
      </section>

      <section className="panel officer-ranking-panel">
        <div className="panel-header">
          <div>
            <span className="eyebrow">
              OFFICER PORTFOLIO
            </span>
            <h3>Account Officer Performance</h3>
            <p>
              Select an officer to expand weekly
              disbursement performance.
            </p>
          </div>
        </div>

        {officerPerformance.length === 0 ? (
          <EmptyData
            title="No account officer data"
            message="Upload a portfolio containing an Account Officer column."
          />
        ) : (
          <div className="officer-list">
            {officerPerformance.map((officer) => (
              <OfficerPerformanceCard
                key={officer.name}
                officer={officer}
              />
            ))}
          </div>
        )}
      </section>
    </>
  );

  /* =========================================================
     REPORTS PAGE
  ========================================================= */

  const ReportsPage = () => {
    const topOfficer = officerData[0];
    const topProduct = productData[0];

    return (
      <>
        <section className="page-intro">
          <div>
            <span className="eyebrow">
              MANAGEMENT INFORMATION
            </span>
            <h1>Reports</h1>
            <p>
              High-level portfolio information generated
              directly from the uploaded loan data.
            </p>
          </div>
        </section>

        <section className="kpi-grid">
          <KpiCard
            title="Portfolio Value"
            value={formatCompactCurrency(
              kpis.totalDisbursed
            )}
            subtitle="Gross amount disbursed"
            icon="₦"
            accent="orange"
          />

          <KpiCard
            title="Outstanding"
            value={formatCompactCurrency(
              kpis.outstandingPrincipal
            )}
            subtitle="Principal still outstanding"
            icon="◈"
            accent="red"
          />

          <KpiCard
            title="Amount Repaid"
            value={formatCompactCurrency(
              kpis.amountRepaid
            )}
            subtitle={`${kpis.repaymentRate.toFixed(
              1
            )}% repayment rate`}
            icon="✓"
            accent="green"
          />

          <KpiCard
            title="Loan Accounts"
            value={kpis.totalLoans.toLocaleString()}
            subtitle="Total accounts"
            icon="▣"
            accent="blue"
          />
        </section>

        <section className="report-grid">
          <div className="panel">
            <div className="panel-header">
              <div>
                <span className="eyebrow">
                  OFFICER SUMMARY
                </span>
                <h3>Officer Portfolio Summary</h3>
              </div>
            </div>

            {officerData.length === 0 ? (
              <EmptyData />
            ) : (
              <div className="report-list">
                {officerData.map((officer) => (
                  <div
                    className="report-list-row"
                    key={officer.name}
                  >
                    <div className="report-name">
                      <div className="small-avatar">
                        {officer.name
                          .charAt(0)
                          .toUpperCase()}
                      </div>

                      <div>
                        <strong>
                          {officer.name}
                        </strong>
                        <span>
                          {officer.count} loans
                        </span>
                      </div>
                    </div>

                    <div>
                      <span>Disbursed</span>
                      <strong>
                        {formatCompactCurrency(
                          officer.amount
                        )}
                      </strong>
                    </div>

                    <div>
                      <span>Outstanding</span>
                      <strong>
                        {formatCompactCurrency(
                          officer.outstanding
                        )}
                      </strong>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="panel">
            <div className="panel-header">
              <div>
                <span className="eyebrow">
                  PRODUCT SUMMARY
                </span>
                <h3>Product Portfolio</h3>
              </div>
            </div>

            {productData.length === 0 ? (
              <EmptyData />
            ) : (
              <div className="report-list">
                {productData.map((product) => (
                  <div
                    className="report-list-row"
                    key={product.name}
                  >
                    <div className="report-name">
                      <div className="product-icon">
                        ◈
                      </div>

                      <div>
                        <strong>
                          {product.name}
                        </strong>
                        <span>
                          {product.count} loans
                        </span>
                      </div>
                    </div>

                    <div>
                      <span>Disbursed</span>
                      <strong>
                        {formatCompactCurrency(
                          product.amount
                        )}
                      </strong>
                    </div>

                    <div>
                      <span>Outstanding</span>
                      <strong>
                        {formatCompactCurrency(
                          product.outstanding
                        )}
                      </strong>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </section>

        <section className="panel">
          <div className="panel-header">
            <div>
              <span className="eyebrow">
                PORTFOLIO HIGHLIGHTS
              </span>
              <h3>Current Portfolio Snapshot</h3>
            </div>
          </div>

          <div className="highlight-grid">
            <div className="highlight-card">
              <span>Largest Officer Portfolio</span>

              <strong>
                {topOfficer
                  ? topOfficer.name
                  : "—"}
              </strong>

              <small>
                {topOfficer
                  ? formatCurrency(
                      topOfficer.amount
                    )
                  : "No data"}
              </small>
            </div>

            <div className="highlight-card">
              <span>Largest Product Portfolio</span>

              <strong>
                {topProduct
                  ? topProduct.name
                  : "—"}
              </strong>

              <small>
                {topProduct
                  ? formatCurrency(
                      topProduct.amount
                    )
                  : "No data"}
              </small>
            </div>

            <div className="highlight-card">
              <span>Portfolio Repayment Rate</span>

              <strong>
                {kpis.repaymentRate.toFixed(1)}%
              </strong>

              <small>
                {formatCurrency(kpis.amountRepaid)} repaid
              </small>
            </div>

            <div className="highlight-card">
              <span>Outstanding Portfolio</span>

              <strong>
                {formatCompactCurrency(
                  kpis.outstandingPrincipal
                )}
              </strong>

              <small>
                Current principal balance
              </small>
            </div>
          </div>
        </section>
      </>
    );
  };

  /* =========================================================
     PAGE TITLE
  ========================================================= */

  const pageTitles = {
    dashboard: {
      title: "Dashboard",
      subtitle:
        "Loan portfolio overview and performance analytics",
    },
    portfolio: {
      title: "Loan Portfolio",
      subtitle:
        "Detailed loan account analysis and portfolio register",
    },
    officers: {
      title: "Account Officers",
      subtitle:
        "Officer-level loan disbursement and portfolio performance",
    },
    reports: {
      title: "Reports",
      subtitle:
        "Management information and portfolio summaries",
    },
  };

  const currentPage = pageTitles[activeSection];

  /* =========================================================
     MAIN RENDER
  ========================================================= */

  return (
    <div className="dashboard-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-logo">CB</div>

          <div>
            <strong>CBSL</strong>
            <span>Analytics</span>
          </div>
        </div>

        <div className="sidebar-section-title">
          MAIN MENU
        </div>

        <nav className="sidebar-nav">
          {menuItems.map((item) => (
            <button
              key={item.id}
              className={`sidebar-item ${
                activeSection === item.id
                  ? "active"
                  : ""
              }`}
              onClick={() => navigateTo(item.id)}
            >
              <span className="sidebar-icon">
                {item.icon}
              </span>

              <span>{item.label}</span>

              {activeSection === item.id && (
                <span className="active-indicator" />
              )}
            </button>
          ))}
        </nav>

        <div className="sidebar-bottom">
          <div className="sidebar-user">
            <div className="user-avatar">
              {accounts?.[0]?.username
                ?.charAt(0)
                ?.toUpperCase() || "U"}
            </div>

            <div className="user-info">
              <strong>
                {accounts?.[0]?.name || "User"}
              </strong>

              <span>Authenticated User</span>
            </div>
          </div>

          <button
            className="logout-button"
            onClick={logout}
          >
            <span>↪</span>
            Sign Out
          </button>
        </div>
      </aside>

      <main className="main-content">
        <header className="topbar">
          <div>
            <div className="breadcrumb">
              CBSL Analytics <span>/</span>{" "}
              {currentPage.title}
            </div>

            <h1>{currentPage.title}</h1>

            <p>{currentPage.subtitle}</p>
          </div>

          <div className="topbar-right">
            <div className="secure-badge">
              <span className="secure-dot" />
              Secure
            </div>

            <div className="top-user">
              <div className="top-user-avatar">
                {accounts?.[0]?.name
                  ?.charAt(0)
                  ?.toUpperCase() || "U"}
              </div>

              <div>
                <strong>
                  {accounts?.[0]?.name ||
                    "Authenticated User"}
                </strong>

                <span>
                  {accounts?.[0]?.username || ""}
                </span>
              </div>
            </div>
          </div>
        </header>

        <div className="content-area">
          {activeSection === "dashboard" && (
            <DashboardOverview />
          )}

          {activeSection === "portfolio" && (
            <LoanPortfolioPage />
          )}

          {activeSection === "officers" && (
            <AccountOfficersPage />
          )}

          {activeSection === "reports" && (
            <ReportsPage />
          )}
        </div>
      </main>

      <style>{`
        * {
          box-sizing: border-box;
        }

        body {
          margin: 0;
          background: #f4f6f9;
          color: #172033;
          font-family:
            Inter,
            -apple-system,
            BlinkMacSystemFont,
            "Segoe UI",
            sans-serif;
        }

        button,
        input,
        select {
          font-family: inherit;
        }

        button {
          cursor: pointer;
        }

        .dashboard-shell {
          min-height: 100vh;
          display: flex;
          background: #f4f6f9;
        }

        /* SIDEBAR */

        .sidebar {
          width: 245px;
          min-height: 100vh;
          background: #101827;
          color: #fff;
          padding: 24px 14px;
          position: fixed;
          left: 0;
          top: 0;
          bottom: 0;
          display: flex;
          flex-direction: column;
          z-index: 20;
        }

        .brand {
          display: flex;
          align-items: center;
          gap: 11px;
          padding: 0 12px 30px;
        }

        .brand-logo {
          width: 42px;
          height: 42px;
          border-radius: 12px;
          background: #f47b20;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 14px;
          font-weight: 800;
          box-shadow: 0 8px 18px rgba(244, 123, 32, 0.2);
        }

        .brand strong {
          display: block;
          font-size: 17px;
          letter-spacing: 0.4px;
        }

        .brand span {
          color: #8792a6;
          font-size: 11px;
          margin-top: 2px;
          display: block;
        }

        .sidebar-section-title {
          padding: 0 13px;
          margin-bottom: 9px;
          font-size: 10px;
          font-weight: 700;
          color: #667085;
          letter-spacing: 1.1px;
        }

        .sidebar-nav {
          display: flex;
          flex-direction: column;
          gap: 5px;
        }

        .sidebar-item {
          position: relative;
          border: 0;
          background: transparent;
          color: #a9b2c2;
          width: 100%;
          min-height: 47px;
          border-radius: 9px;
          display: flex;
          align-items: center;
          gap: 13px;
          padding: 0 13px;
          font-size: 13px;
          font-weight: 500;
          text-align: left;
          transition: 0.2s ease;
        }

        .sidebar-item:hover {
          background: #182337;
          color: #fff;
        }

        .sidebar-item.active {
          background: #1b2a42;
          color: #fff;
        }

        .sidebar-icon {
          width: 22px;
          text-align: center;
          font-size: 17px;
          color: #8d9ab0;
        }

        .sidebar-item.active .sidebar-icon {
          color: #f47b20;
        }

        .active-indicator {
          position: absolute;
          right: 0;
          width: 3px;
          height: 25px;
          border-radius: 4px 0 0 4px;
          background: #f47b20;
        }

        .sidebar-bottom {
          margin-top: auto;
        }

        .sidebar-user {
          border-top: 1px solid #263246;
          padding: 18px 9px 15px;
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .user-avatar,
        .top-user-avatar {
          flex-shrink: 0;
          width: 35px;
          height: 35px;
          border-radius: 50%;
          background: #263653;
          color: #fff;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 12px;
          font-weight: 700;
        }

        .user-info {
          min-width: 0;
        }

        .user-info strong {
          display: block;
          font-size: 11px;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          max-width: 150px;
        }

        .user-info span {
          color: #718097;
          font-size: 9px;
          display: block;
          margin-top: 3px;
        }

        .logout-button {
          width: 100%;
          border: 1px solid #303c51;
          background: transparent;
          color: #aeb8c8;
          padding: 9px;
          border-radius: 7px;
          font-size: 11px;
          display: flex;
          gap: 8px;
          align-items: center;
          justify-content: center;
        }

        .logout-button:hover {
          border-color: #f47b20;
          color: #fff;
        }

        /* MAIN */

        .main-content {
          flex: 1;
          margin-left: 245px;
          min-width: 0;
        }

        .topbar {
          min-height: 93px;
          background: #fff;
          border-bottom: 1px solid #e7eaf0;
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 19px 34px;
          position: sticky;
          top: 0;
          z-index: 10;
        }

        .breadcrumb {
          color: #8b94a5;
          font-size: 10px;
          margin-bottom: 5px;
        }

        .breadcrumb span {
          padding: 0 5px;
          color: #c2c8d1;
        }

        .topbar h1 {
          margin: 0;
          color: #101827;
          font-size: 21px;
          font-weight: 750;
        }

        .topbar p {
          margin: 4px 0 0;
          color: #7c8799;
          font-size: 11px;
        }

        .topbar-right {
          display: flex;
          align-items: center;
          gap: 22px;
        }

        .secure-badge {
          color: #21864a;
          background: #edf9f1;
          border: 1px solid #d9f1e1;
          border-radius: 30px;
          padding: 6px 10px;
          font-size: 10px;
          font-weight: 650;
          display: flex;
          align-items: center;
          gap: 6px;
        }

        .secure-dot {
          width: 6px;
          height: 6px;
          border-radius: 50%;
          background: #22a15a;
        }

        .top-user {
          display: flex;
          align-items: center;
          gap: 9px;
        }

        .top-user strong {
          display: block;
          font-size: 11px;
          color: #202938;
        }

        .top-user span {
          display: block;
          color: #8992a1;
          font-size: 9px;
          margin-top: 2px;
          max-width: 180px;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .top-user-avatar {
          background: #eef1f6;
          color: #354157;
        }

        .content-area {
          padding: 27px 34px 50px;
        }

        /* UPLOAD */

        .upload-panel {
          background: linear-gradient(
            120deg,
            #162236,
            #1c2d47
          );
          color: #fff;
          border-radius: 12px;
          padding: 23px 25px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 20px;
          margin-bottom: 19px;
          box-shadow: 0 8px 20px rgba(16, 24, 39, 0.08);
        }

        .eyebrow {
          font-size: 9px;
          letter-spacing: 1.1px;
          font-weight: 800;
          color: #f47b20;
          display: block;
          margin-bottom: 6px;
        }

        .upload-panel h2 {
          margin: 0;
          font-size: 18px;
        }

        .upload-panel p {
          margin: 6px 0 0;
          font-size: 11px;
          color: #b5c0d1;
        }

        .upload-area {
          min-width: 230px;
        }

        /* KPI */

        .kpi-grid {
          display: grid;
          grid-template-columns: repeat(5, 1fr);
          gap: 13px;
          margin-bottom: 19px;
        }

        .kpi-card {
          position: relative;
          background: #fff;
          border: 1px solid #e8ebf0;
          border-radius: 10px;
          padding: 17px;
          min-height: 122px;
          overflow: hidden;
        }

        .kpi-card::before {
          content: "";
          position: absolute;
          left: 0;
          top: 0;
          bottom: 0;
          width: 3px;
          background: #2563eb;
        }

        .kpi-card.orange::before {
          background: #f47b20;
        }

        .kpi-card.green::before {
          background: #16a34a;
        }

        .kpi-card.red::before {
          background: #dc2626;
        }

        .kpi-card.purple::before {
          background: #7c3aed;
        }

        .kpi-top {
          display: flex;
          align-items: center;
          justify-content: space-between;
        }

        .kpi-title {
          font-size: 10px;
          color: #737e90;
          font-weight: 600;
        }

        .kpi-icon {
          width: 29px;
          height: 29px;
          border-radius: 7px;
          background: #edf3ff;
          color: #2563eb;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 13px;
        }

        .orange .kpi-icon {
          background: #fff3e8;
          color: #f47b20;
        }

        .green .kpi-icon {
          background: #eaf8ef;
          color: #16a34a;
        }

        .red .kpi-icon {
          background: #fff0f0;
          color: #dc2626;
        }

        .purple .kpi-icon {
          background: #f2edff;
          color: #7c3aed;
        }

        .kpi-value {
          margin-top: 13px;
          color: #111a2a;
          font-size: 21px;
          font-weight: 800;
          letter-spacing: -0.4px;
        }

        .kpi-subtitle {
          margin-top: 5px;
          color: #9aa3b1;
          font-size: 9px;
        }

        /* PANELS */

        .analytics-grid {
          display: grid;
          grid-template-columns: 1.6fr 1fr;
          gap: 17px;
          margin-bottom: 17px;
        }

        .panel {
          background: #fff;
          border: 1px solid #e7eaf0;
          border-radius: 10px;
          padding: 20px;
          min-width: 0;
          box-shadow: 0 2px 8px rgba(18, 31, 53, 0.025);
        }

        .panel-header {
          display: flex;
          justify-content: space-between;
          align-items: flex-start;
          gap: 20px;
          margin-bottom: 18px;
        }

        .panel-header h3 {
          margin: 0;
          font-size: 14px;
          color: #172033;
        }

        .panel-header p {
          margin: 5px 0 0;
          color: #8a94a5;
          font-size: 10px;
        }

        .panel-badge {
          background: #f4f6f9;
          color: #667085;
          padding: 6px 9px;
          border-radius: 5px;
          font-size: 9px;
          white-space: nowrap;
        }

        .text-button {
          border: 0;
          background: transparent;
          color: #2563eb;
          font-size: 10px;
          font-weight: 700;
          padding: 0;
        }

        .text-button:hover {
          color: #1748ae;
        }

        /* TREND */

        .trend-chart {
          height: 265px;
          display: flex;
          align-items: stretch;
          gap: 13px;
          padding-top: 5px;
          overflow-x: auto;
        }

        .trend-column {
          flex: 1;
          min-width: 70px;
          display: flex;
          flex-direction: column;
          align-items: center;
        }

        .trend-value {
          color: #677286;
          font-size: 8px;
          height: 24px;
          white-space: nowrap;
        }

        .trend-bar-area {
          flex: 1;
          width: 100%;
          max-width: 44px;
          display: flex;
          align-items: flex-end;
          background: #f5f7fa;
          border-radius: 6px 6px 3px 3px;
          overflow: hidden;
        }

        .trend-bar {
          width: 100%;
          min-height: 4px;
          background: linear-gradient(
            180deg,
            #f47b20,
            #ef6411
          );
          border-radius: 6px 6px 2px 2px;
        }

        .trend-label {
          margin-top: 8px;
          color: #596477;
          font-size: 8px;
          font-weight: 700;
          white-space: nowrap;
        }

        .trend-count {
          color: #9ba3b1;
          font-size: 8px;
          margin-top: 3px;
        }

        /* BARS */

        .bar-list {
          display: flex;
          flex-direction: column;
          gap: 15px;
        }

        .bar-row {
          min-width: 0;
        }

        .bar-label {
          display: flex;
          justify-content: space-between;
          gap: 10px;
          margin-bottom: 6px;
          font-size: 9px;
        }

        .bar-label span {
          color: #586376;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .bar-label strong {
          color: #273248;
          white-space: nowrap;
          font-size: 9px;
        }

        .bar-track {
          height: 6px;
          background: #eef1f5;
          border-radius: 20px;
          overflow: hidden;
        }

        .bar-fill {
          height: 100%;
          border-radius: 20px;
          min-width: 2px;
        }

        /* WEEKLY */

        .weekly-panel {
          margin-bottom: 17px;
        }

        .monthly-weekly-list {
          display: flex;
          flex-direction: column;
          gap: 22px;
        }

        .monthly-week-block {
          border-bottom: 1px solid #edf0f4;
          padding-bottom: 20px;
        }

        .monthly-week-block:last-child {
          border-bottom: 0;
          padding-bottom: 0;
        }

        .month-title {
          display: flex;
          justify-content: space-between;
          margin-bottom: 13px;
        }

        .month-title strong {
          font-size: 11px;
          color: #263147;
        }

        .month-title span {
          color: #8993a3;
          font-size: 9px;
        }

        .week-grid {
          display: grid;
          grid-template-columns: repeat(5, 1fr);
          gap: 9px;
        }

        .week-card {
          border: 1px solid #edf0f4;
          background: #fafbfc;
          border-radius: 8px;
          padding: 10px;
          text-align: center;
        }

        .week-number {
          color: #7f8999;
          font-size: 9px;
          font-weight: 700;
        }

        .week-chart {
          height: 74px;
          display: flex;
          align-items: flex-end;
          justify-content: center;
          margin: 8px 0;
        }

        .week-bar {
          width: 25px;
          background: linear-gradient(
            180deg,
            #2563eb,
            #1748ae
          );
          border-radius: 5px 5px 2px 2px;
          min-height: 1px;
        }

        .week-card strong {
          display: block;
          color: #263147;
          font-size: 10px;
        }

        .week-card span {
          display: block;
          margin-top: 3px;
          color: #8d96a5;
          font-size: 8px;
        }

        /* COMPOSITION */

        .composition {
          display: flex;
          flex-direction: column;
        }

        .composition-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 15px 0;
          border-bottom: 1px solid #edf0f4;
        }

        .composition-row:last-child {
          border-bottom: 0;
        }

        .composition-row span {
          color: #768195;
          font-size: 10px;
        }

        .composition-row strong {
          color: #263147;
          font-size: 11px;
        }

        .green-text {
          color: #159447 !important;
        }

        .red-text {
          color: #d92d20 !important;
        }

        /* TABLE */

        .table-wrapper {
          overflow-x: auto;
        }

        .data-table {
          width: 100%;
          border-collapse: collapse;
          min-width: 850px;
        }

        .data-table th {
          text-align: left;
          padding: 10px 9px;
          background: #fafbfc;
          color: #8791a2;
          font-size: 8px;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          border-bottom: 1px solid #e8ebf0;
          white-space: nowrap;
        }

        .data-table td {
          padding: 11px 9px;
          border-bottom: 1px solid #f0f2f5;
          color: #596477;
          font-size: 9px;
          white-space: nowrap;
        }

        .data-table tbody tr:hover {
          background: #fafcff;
        }

        .data-table td strong {
          color: #263147;
          font-size: 10px;
        }

        .money-cell {
          font-weight: 700;
          color: #263147 !important;
        }

        .table-tag {
          background: #f1f5ff;
          color: #3158a7;
          border-radius: 4px;
          padding: 4px 7px;
          font-size: 8px;
          font-weight: 700;
        }

        /* PAGE INTRO */

        .page-intro {
          display: flex;
          align-items: flex-end;
          justify-content: space-between;
          margin-bottom: 21px;
        }

        .page-intro h1 {
          margin: 0;
          font-size: 24px;
          color: #172033;
        }

        .page-intro p {
          margin: 6px 0 0;
          color: #808a9b;
          font-size: 11px;
        }

        .record-count {
          background: #fff;
          border: 1px solid #e4e8ee;
          border-radius: 7px;
          padding: 8px 12px;
          color: #657084;
          font-size: 9px;
          font-weight: 700;
        }

        .portfolio-kpis {
          grid-template-columns: repeat(4, 1fr);
        }

        /* FILTERS */

        .filters-panel {
          margin-bottom: 17px;
        }

        .filters-header {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          margin-bottom: 17px;
        }

        .filters-header h3 {
          margin: 0;
          font-size: 14px;
        }

        .clear-button {
          border: 1px solid #dce2ea;
          background: #fff;
          color: #667085;
          padding: 7px 10px;
          border-radius: 6px;
          font-size: 9px;
          font-weight: 700;
        }

        .clear-button:hover {
          border-color: #f47b20;
          color: #f47b20;
        }

        .filter-grid {
          display: grid;
          grid-template-columns: 2fr 1fr 1fr;
          gap: 13px;
        }

        .filter-field label {
          display: block;
          margin-bottom: 6px;
          color: #667085;
          font-size: 9px;
          font-weight: 700;
        }

        .filter-field input,
        .filter-field select {
          width: 100%;
          height: 38px;
          border: 1px solid #dfe4eb;
          background: #fff;
          border-radius: 7px;
          padding: 0 10px;
          color: #344054;
          font-size: 10px;
          outline: none;
        }

        .filter-field input:focus,
        .filter-field select:focus {
          border-color: #8bb0f6;
          box-shadow: 0 0 0 3px #edf3ff;
        }

        .search-input-wrapper {
          position: relative;
        }

        .search-input-wrapper > span {
          position: absolute;
          left: 11px;
          top: 10px;
          color: #8993a4;
          font-size: 14px;
        }

        .search-input-wrapper input {
          padding-left: 31px;
        }

        .portfolio-table-panel {
          padding-bottom: 5px;
        }

        .portfolio-table {
          min-width: 1450px;
        }

        .row-number {
          color: #a0a8b5 !important;
          width: 35px;
        }

        .customer-cell {
          display: flex;
          flex-direction: column;
          gap: 3px;
        }

        .customer-cell small {
          color: #9ba4b2;
          font-size: 8px;
        }

        .progress-cell {
          min-width: 75px;
        }

        .progress-cell > span {
          display: block;
          font-size: 8px;
          color: #667085;
          margin-bottom: 4px;
        }

        .mini-progress {
          height: 4px;
          width: 70px;
          border-radius: 20px;
          background: #edf0f4;
          overflow: hidden;
        }

        .mini-progress > div {
          height: 100%;
          background: #16a34a;
          border-radius: 20px;
        }

        /* OFFICERS */

        .officer-ranking-panel {
          padding: 21px;
        }

        .officer-list {
          display: flex;
          flex-direction: column;
          gap: 10px;
        }

        .officer-card {
          border: 1px solid #e6eaf0;
          border-radius: 9px;
          overflow: hidden;
          background: #fff;
        }

        .officer-card-open {
          border-color: #cbd9f5;
          box-shadow: 0 5px 18px rgba(37, 99, 235, 0.06);
        }

        .officer-card-header {
          width: 100%;
          border: 0;
          background: #fff;
          padding: 13px 15px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 20px;
          text-align: left;
        }

        .officer-card-header:hover {
          background: #fafcff;
        }

        .officer-identity {
          display: flex;
          align-items: center;
          gap: 10px;
          min-width: 220px;
        }

        .officer-avatar {
          width: 35px;
          height: 35px;
          border-radius: 9px;
          background: #eef3ff;
          color: #315db4;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 10px;
          font-weight: 800;
        }

        .officer-identity strong {
          display: block;
          color: #253047;
          font-size: 11px;
        }

        .officer-identity span {
          display: block;
          color: #919aa9;
          font-size: 8px;
          margin-top: 3px;
        }

        .officer-header-stats {
          display: flex;
          align-items: center;
          gap: 28px;
        }

        .officer-header-stats > div {
          min-width: 90px;
        }

        .officer-header-stats span {
          display: block;
          color: #929bab;
          font-size: 8px;
          margin-bottom: 3px;
        }

        .officer-header-stats strong {
          color: #29344a;
          font-size: 10px;
        }

        .expand-icon {
          color: #667085 !important;
          font-size: 16px !important;
          margin: 0 !important;
        }

        .officer-card-body {
          border-top: 1px solid #edf0f4;
          padding: 18px;
          background: #fbfcfe;
        }

        .officer-summary-grid {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 10px;
          margin-bottom: 23px;
        }

        .officer-summary-grid > div {
          background: #fff;
          border: 1px solid #e9edf2;
          border-radius: 7px;
          padding: 12px;
        }

        .officer-summary-grid span {
          display: block;
          color: #8a94a5;
          font-size: 8px;
          margin-bottom: 6px;
        }

        .officer-summary-grid strong {
          color: #263147;
          font-size: 13px;
        }

        .officer-weekly-title {
          margin-bottom: 12px;
        }

        .officer-weekly-title h4 {
          margin: 0;
          font-size: 12px;
          color: #29344a;
        }

        .officer-month-list {
          display: flex;
          flex-direction: column;
          gap: 18px;
        }

        .officer-month {
          background: #fff;
          border: 1px solid #e7ebf0;
          border-radius: 8px;
          padding: 14px;
        }

        .officer-month-heading {
          display: flex;
          justify-content: space-between;
          margin-bottom: 12px;
        }

        .officer-month-heading strong {
          color: #344054;
          font-size: 10px;
        }

        .officer-month-heading span {
          color: #8b95a5;
          font-size: 8px;
        }

        .officer-week-chart {
          display: grid;
          grid-template-columns: repeat(5, 1fr);
          gap: 9px;
          height: 160px;
        }

        .officer-week-column {
          display: flex;
          flex-direction: column;
          align-items: center;
          min-width: 0;
        }

        .officer-week-value {
          color: #667085;
          font-size: 7px;
          height: 19px;
          white-space: nowrap;
        }

        .officer-week-track {
          width: 100%;
          max-width: 38px;
          flex: 1;
          display: flex;
          align-items: flex-end;
          background: #f3f5f8;
          border-radius: 5px 5px 2px 2px;
          overflow: hidden;
        }

        .officer-week-fill {
          width: 100%;
          background: #2563eb;
          border-radius: 5px 5px 2px 2px;
          min-height: 1px;
        }

        .officer-week-column > strong {
          margin-top: 6px;
          color: #667085;
          font-size: 8px;
        }

        .officer-week-column > span {
          color: #9aa3b1;
          font-size: 7px;
          margin-top: 2px;
        }

        /* REPORTS */

        .report-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 17px;
          margin-bottom: 17px;
        }

        .report-list {
          display: flex;
          flex-direction: column;
        }

        .report-list-row {
          display: grid;
          grid-template-columns: 1.5fr 1fr 1fr;
          gap: 15px;
          align-items: center;
          padding: 13px 0;
          border-bottom: 1px solid #edf0f4;
        }

        .report-list-row:last-child {
          border-bottom: 0;
        }

        .report-name {
          display: flex;
          align-items: center;
          gap: 9px;
          min-width: 0;
        }

        .small-avatar,
        .product-icon {
          width: 29px;
          height: 29px;
          border-radius: 7px;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #eef3ff;
          color: #315db4;
          font-size: 9px;
          font-weight: 800;
          flex-shrink: 0;
        }

        .product-icon {
          background: #f2edff;
          color: #7c3aed;
        }

        .report-name strong {
          display: block;
          font-size: 9px;
          color: #344054;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }

        .report-name span {
          display: block;
          color: #99a1ae;
          font-size: 7px;
          margin-top: 3px;
        }

        .report-list-row > div:not(.report-name) span {
          display: block;
          color: #9aa3b1;
          font-size: 7px;
          margin-bottom: 3px;
        }

        .report-list-row > div:not(.report-name) strong {
          color: #344054;
          font-size: 9px;
        }

        .highlight-grid {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 11px;
        }

        .highlight-card {
          border: 1px solid #e8ebf0;
          border-radius: 8px;
          padding: 15px;
          background: #fafbfc;
        }

        .highlight-card span {
          display: block;
          color: #8a94a5;
          font-size: 8px;
          margin-bottom: 8px;
        }

        .highlight-card strong {
          display: block;
          color: #253047;
          font-size: 12px;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .highlight-card small {
          display: block;
          margin-top: 5px;
          color: #9aa3b1;
          font-size: 8px;
        }

        /* EMPTY */

        .empty-state {
          min-height: 170px;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          text-align: center;
          color: #8993a4;
        }

        .empty-icon {
          width: 43px;
          height: 43px;
          border-radius: 12px;
          background: #f1f4f8;
          display: flex;
          align-items: center;
          justify-content: center;
          color: #8d98aa;
          font-size: 19px;
          margin-bottom: 10px;
        }

        .empty-state h3 {
          margin: 0;
          color: #4b5567;
          font-size: 12px;
        }

        .empty-state p {
          margin: 5px 0 0;
          color: #9aa3b1;
          font-size: 9px;
        }

        .small-empty {
          padding: 30px;
          text-align: center;
          color: #929bab;
          font-size: 9px;
        }

        /* RESPONSIVE */

        @media (max-width: 1250px) {
          .kpi-grid {
            grid-template-columns: repeat(3, 1fr);
          }

          .analytics-grid {
            grid-template-columns: 1fr;
          }

          .portfolio-kpis {
            grid-template-columns: repeat(2, 1fr);
          }

          .highlight-grid {
            grid-template-columns: repeat(2, 1fr);
          }
        }

        @media (max-width: 950px) {
          .sidebar {
            width: 72px;
            padding: 20px 9px;
          }

          .brand {
            justify-content: center;
            padding: 0 0 25px;
          }

          .brand > div:last-child,
          .sidebar-section-title,
          .sidebar-item span:not(.sidebar-icon):not(.active-indicator),
          .sidebar-user .user-info,
          .logout-button {
            display: none;
          }

          .sidebar-item {
            justify-content: center;
            padding: 0;
          }

          .sidebar-icon {
            margin: 0;
          }

          .sidebar-user {
            justify-content: center;
            padding-left: 0;
            padding-right: 0;
          }

          .main-content {
            margin-left: 72px;
          }

          .topbar {
            padding: 17px 22px;
          }

          .content-area {
            padding: 22px;
          }

          .officer-header-stats {
            gap: 12px;
          }

          .officer-header-stats > div {
            min-width: 70px;
          }
        }

        @media (max-width: 700px) {
          .topbar {
            align-items: flex-start;
          }

          .topbar-right {
            display: none;
          }

          .content-area {
            padding: 17px 13px 35px;
          }

          .upload-panel {
            flex-direction: column;
            align-items: flex-start;
          }

          .upload-area {
            width: 100%;
          }

          .kpi-grid,
          .portfolio-kpis {
            grid-template-columns: repeat(2, 1fr);
          }

          .filter-grid {
            grid-template-columns: 1fr;
          }

          .week-grid {
            gap: 5px;
          }

          .week-card {
            padding: 7px 3px;
          }

          .officer-card-header {
            align-items: flex-start;
          }

          .officer-header-stats {
            display: none;
          }

          .officer-summary-grid {
            grid-template-columns: repeat(2, 1fr);
          }

          .report-grid {
            grid-template-columns: 1fr;
          }

          .highlight-grid {
            grid-template-columns: 1fr;
          }

          .page-intro {
            align-items: flex-start;
            flex-direction: column;
            gap: 12px;
          }
        }

        @media (max-width: 480px) {
          .kpi-grid,
          .portfolio-kpis {
            grid-template-columns: 1fr;
          }

          .topbar h1 {
            font-size: 18px;
          }

          .upload-panel {
            padding: 18px;
          }

          .panel {
            padding: 15px;
          }

          .officer-summary-grid {
            grid-template-columns: 1fr;
          }
        }
      `}</style>
    </div>
  );
}
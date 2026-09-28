import * as XLSX from "xlsx";

function FileUpload({ onDataLoaded }) {
  const handleFileUpload = (event) => {
    const file = event.target.files[0];

    if (!file) {
      return;
    }

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        // ============================================================
        // 1. READ EXCEL WORKBOOK
        // ============================================================

        const workbook = XLSX.read(e.target.result, {
          type: "binary",
          cellDates: true,
        });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        if (!worksheet) {
          throw new Error("Worksheet not found.");
        }

        // ============================================================
        // 2. READ EXCEL AS RAW ROWS
        // ============================================================

        const rows = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          blankrows: false,
        });

        console.log("=================================");
        console.log("EXCEL FILE:", file.name);
        console.log("SHEET:", sheetName);
        console.log("TOTAL ROWS:", rows.length);
        console.log("RAW DATA:", rows);

        if (rows.length === 0) {
          throw new Error("The Excel file is empty.");
        }

        // ============================================================
        // 3. CLEAN TEXT
        // ============================================================

        const cleanText = (value) => {
          if (value === null || value === undefined) {
            return "";
          }

          return String(value)
            .replace(/\u00A0/g, " ")
            .replace(/\r?\n/g, " ")
            .replace(/\s+/g, " ")
            .trim();
        };

        // ============================================================
        // 4. NORMALIZE COLUMN NAMES
        // ============================================================

        const normalize = (value) => {
          return cleanText(value)
            .toLowerCase()
            .replace(/[.:]/g, "")
            .replace(/\s+/g, " ")
            .trim();
        };

        // ============================================================
        // 5. CONVERT VALUES TO NUMBERS
        // ============================================================

        const toNumber = (value) => {
          if (
            value === null ||
            value === undefined ||
            value === ""
          ) {
            return 0;
          }

          if (typeof value === "number") {
            return Number.isFinite(value) ? value : 0;
          }

          const cleaned = String(value)
            .replace(/₦/g, "")
            .replace(/NGN/gi, "")
            .replace(/,/g, "")
            .replace(/%/g, "")
            .trim();

          if (cleaned === "") {
            return 0;
          }

          const number = Number(cleaned);

          return Number.isFinite(number) ? number : 0;
        };

        // ============================================================
        // 6. FIND THE REAL HEADER ROW
        // ============================================================

        const keywords = [
          "customer",
          "account",
          "loan",
          "amount",
          "balance",
          "product",
          "officer",
          "phone",
          "gender",
        ];

        let headerRowIndex = -1;

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i] || [];

          const rowText = row
            .map((cell) => normalize(cell))
            .join(" ");

          let matches = 0;

          keywords.forEach((keyword) => {
            if (rowText.includes(keyword)) {
              matches++;
            }
          });

          if (matches >= 3) {
            headerRowIndex = i;
            break;
          }
        }

        console.log(
          "HEADER ROW INDEX:",
          headerRowIndex
        );

        if (headerRowIndex === -1) {
          throw new Error(
            "Could not find the actual Excel column headings."
          );
        }

        // ============================================================
        // 7. GET EXCEL HEADERS
        // ============================================================

        const headerRow = rows[headerRowIndex];

        const headers = headerRow.map(
          (header, index) => {
            const value = cleanText(header);

            if (value === "") {
              return `Column_${index + 1}`;
            }

            return value;
          }
        );

        console.log(
          "ACTUAL EXCEL HEADERS:",
          headers
        );

        // ============================================================
        // 8. CREATE NORMALIZED HEADER MAP
        // ============================================================

        const headerMap = {};

        headers.forEach((header) => {
          headerMap[normalize(header)] = header;
        });

        console.log(
          "NORMALIZED HEADER MAP:",
          headerMap
        );

        // ============================================================
        // 9. GET DATA ROWS
        // ============================================================

        const dataRows = rows
          .slice(headerRowIndex + 1)
          .filter((row) => {
            return row.some(
              (cell) => cleanText(cell) !== ""
            );
          });

        console.log(
          "DATA ROWS:",
          dataRows.length
        );

        // ============================================================
        // 10. CONVERT EACH EXCEL ROW INTO A RECORD
        // ============================================================

        const records = dataRows.map(
          (row, index) => {
            const record = {};

            // --------------------------------------------------------
            // Preserve ALL original Excel columns
            // --------------------------------------------------------

            headers.forEach(
              (header, columnIndex) => {
                record[header] =
                  row[columnIndex] !== undefined
                    ? row[columnIndex]
                    : "";
              }
            );

            // --------------------------------------------------------
            // Create normalized record for field lookup
            // --------------------------------------------------------

            const normalizedRecord = {};

            Object.keys(record).forEach(
              (key) => {
                normalizedRecord[
                  normalize(key)
                ] = record[key];
              }
            );

            // ========================================================
            // HELPER: FIND FIELD
            // ========================================================

            const getField = (possibleNames) => {
              for (
                let i = 0;
                i < possibleNames.length;
                i++
              ) {
                const normalizedName =
                  normalize(possibleNames[i]);

                if (
                  Object.prototype.hasOwnProperty.call(
                    normalizedRecord,
                    normalizedName
                  )
                ) {
                  return normalizedRecord[
                    normalizedName
                  ];
                }
              }

              return "";
            };

            // ========================================================
            // CUSTOMER INFORMATION
            // ========================================================

            record.customerId = cleanText(
              getField([
                "Customer ID",
                "Customer No",
                "Customer Number",
                "Customer ID.",
              ])
            );

            record.customerName = cleanText(
              getField([
                "Customer Name",
                "Customer",
                "Name",
                "Account Name",
              ])
            );

            record.accountNo = cleanText(
              getField([
                "Account No.",
                "Account No",
                "Account Number",
                "Account",
              ])
            );

            record.address = cleanText(
              getField([
                "Address",
                "Customer Address",
              ])
            );

            record.gender = cleanText(
              getField([
                "Gender",
                "Sex",
              ])
            );

            record.phoneNo = cleanText(
              getField([
                "Phone No.",
                "Phone No",
                "Phone Number",
                "Phone",
                "Mobile",
                "Mobile Number",
              ])
            );

            // ========================================================
            // ACCOUNT OFFICER
            // ========================================================

            record.accountOfficer = cleanText(
              getField([
                "Account Officer",
                "Account Officer Name",
                "Loan Officer",
                "Officer",
                "Relationship Officer",
              ])
            );

            // ========================================================
            // LOAN AMOUNT / AMOUNT DISBURSED
            // ========================================================

            /*
             * IMPORTANT:
             *
             * Your actual Excel file contains:
             *
             * Loan Amount: 250000
             *
             * Therefore this is the amount that was
             * disbursed by the Account Officer.
             */

            record.loanAmount = toNumber(
              getField([
                "Loan Amount",
                "Amount Disbursed",
                "Disbursed Amount",
                "Disbursement Amount",
                "Amount",
                "Principal Amount",
              ])
            );

            /*
             * Keep a second explicit field for the
             * dashboard/analytics to use if needed.
             */

            record.amountDisbursed =
              record.loanAmount;

            // ========================================================
            // PRINCIPAL BALANCE
            // ========================================================

            /*
             * Your actual Excel heading is:
             *
             * Principal Bal.
             *
             * normalize() converts it to:
             *
             * principal bal
             */

            record.principalBalance = toNumber(
              getField([
                "Principal Bal.",
                "Principal Bal",
                "Principal Balance",
                "Principal Outstanding",
                "Outstanding Balance",
                "Loan Balance",
                "Balance",
              ])
            );

            /*
             * Keep principalBal as an alias so that
             * existing dashboard code using principalBal
             * will also continue to work.
             */

            record.principalBal =
              record.principalBalance;

            // ========================================================
            // LOAN PRODUCT
            // ========================================================

            record.product = cleanText(
              getField([
                "Product",
                "Loan Product",
                "Product Name",
                "Loan Type",
              ])
            );

            // ========================================================
            // DATES
            // ========================================================

            record.disbursementDate =
              getField([
                "Disbursement Date",
                "Disbursed Date",
                "Date Disbursed",
              ]);

            record.effectiveDate =
              getField([
                "Effective Date",
                "Loan Effective Date",
              ]);

            record.maturationDate =
              getField([
                "Maturation Date",
                "Maturity Date",
                "Expiry Date",
              ]);

            // ========================================================
            // INTEREST RATE
            // ========================================================

            record.interestRate = toNumber(
              getField([
                "Interest Rate",
                "Interest",
                "Rate",
              ])
            );

            // ========================================================
            // REPAYMENT METHODS
            // ========================================================

            record.interestRepayment = cleanText(
              getField([
                "Int. Repay.",
                "Interest Repayment",
                "Interest Repayment Method",
              ])
            );

            record.principalRepayment = cleanText(
              getField([
                "Prin. Repay.",
                "Principal Repayment",
                "Principal Repayment Method",
              ])
            );

            // ========================================================
            // LINKED ACCOUNT
            // ========================================================

            record.linkedAccountName = cleanText(
              getField([
                "Linked Account Name",
                "Linked Account",
              ])
            );

            record.linkedAccountNumber = cleanText(
              getField([
                "Linked Account Number",
                "Linked Account No.",
                "Linked Account No",
              ])
            );

            // ========================================================
            // BVN
            // ========================================================

            record.bvn = cleanText(
              getField([
                "BVN",
                "Bank Verification Number",
              ])
            );

            // ========================================================
            // PREVIOUS LOAN
            // ========================================================

            record.hasPreviouslyTakenLoan =
              getField([
                "Has Previously Taken Loan",
                "Previously Taken Loan",
                "Previous Loan",
              ]);

            // ========================================================
            // ORIGINAL EXCEL ROW
            // ========================================================

            record._excelRow =
              headerRowIndex + index + 2;

            // ========================================================
            // RETURN RECORD
            // ========================================================

            return record;
          }
        );

        // ============================================================
        // 11. REMOVE EMPTY RECORDS
        // ============================================================

        const cleanedRecords =
          records.filter((record) => {
            return (
              record.customerId !== "" ||
              record.customerName !== "" ||
              record.accountNo !== "" ||
              record.accountOfficer !== "" ||
              record.loanAmount !== 0 ||
              record.principalBalance !== 0
            );
          });

        // ============================================================
        // 12. CALCULATE TOTAL DISBURSED
        // ============================================================

        const totalDisbursed =
          cleanedRecords.reduce(
            (total, record) => {
              return (
                total +
                Number(record.loanAmount || 0)
              );
            },
            0
          );

        // ============================================================
        // 13. CALCULATE TOTAL OUTSTANDING
        // ============================================================

        const totalOutstanding =
          cleanedRecords.reduce(
            (total, record) => {
              return (
                total +
                Number(
                  record.principalBalance || 0
                )
              );
            },
            0
          );

        // ============================================================
        // 14. CALCULATE OFFICER PORTFOLIO
        // ============================================================

        const officerPortfolio = {};

        cleanedRecords.forEach((record) => {
          const officer =
            cleanText(record.accountOfficer) ||
            "Unassigned";

          const amount =
            Number(record.loanAmount) || 0;

          if (!officerPortfolio[officer]) {
            officerPortfolio[officer] = {
              officer: officer,
              amountDisbursed: 0,
              loanCount: 0,
            };
          }

          officerPortfolio[
            officer
          ].amountDisbursed += amount;

          officerPortfolio[
            officer
          ].loanCount += 1;
        });

        // ============================================================
        // 15. CONVERT OFFICER PORTFOLIO TO ARRAY
        // ============================================================

        const officerPortfolioArray =
          Object.values(
            officerPortfolio
          ).sort(
            (a, b) =>
              b.amountDisbursed -
              a.amountDisbursed
          );

        // ============================================================
        // 16. DEBUG INFORMATION
        // ============================================================

        console.log(
          "================================="
        );

        console.log(
          "NUMBER OF RECORDS:",
          cleanedRecords.length
        );

        console.log(
          "TOTAL DISBURSED:",
          totalDisbursed
        );

        console.log(
          "TOTAL OUTSTANDING:",
          totalOutstanding
        );

        console.log(
          "OFFICER PORTFOLIO:",
          officerPortfolioArray
        );

        console.log(
          "FIRST CLEANED RECORD:",
          cleanedRecords[0]
        );

        console.log(
          "ALL CLEANED RECORDS:",
          cleanedRecords
        );

        console.log(
          "================================="
        );

        // ============================================================
        // 17. SEND DATA TO DASHBOARD
        // ============================================================

        if (
          typeof onDataLoaded === "function"
        ) {
          onDataLoaded(cleanedRecords);
        }

        // ============================================================
        // 18. SUCCESS MESSAGE
        // ============================================================

        if (cleanedRecords.length > 0) {
          console.log(
            `${cleanedRecords.length} loan records loaded successfully.`
          );

          console.log(
            `Total amount disbursed: ₦${totalDisbursed.toLocaleString()}`
          );

          console.log(
            `Total outstanding balance: ₦${totalOutstanding.toLocaleString()}`
          );
        } else {
          alert(
            "The Excel file was read, but no loan records were found."
          );
        }
      } catch (error) {
        console.error(
          "================================="
        );

        console.error(
          "ERROR READING EXCEL FILE:",
          error
        );

        console.error(
          "================================="
        );

        alert(
          "Unable to read the Excel file: " +
            error.message
        );
      }
    };

    reader.onerror = () => {
      console.error(
        "Unable to read the selected Excel file."
      );

      alert(
        "Unable to read the selected Excel file."
      );
    };

    reader.readAsBinaryString(file);
  };

  // ================================================================
  // FILE UPLOAD UI
  // ================================================================

  return (
    <div style={{ marginBottom: "20px" }}>
      <input
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileUpload}
      />
    </div>
  );
}

export default FileUpload;
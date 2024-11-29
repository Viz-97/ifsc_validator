const axios = require("axios");
const Excel = require("exceljs");
const ifsc = require("ifsc");
const readline = require("readline");
const express = require("express");
const fs = require("fs");
const path = require("path");

// Initialize web server
const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000; // Detect the port dynamically

// File setup
const outputFilePath = path.join(__dirname, "output.xlsx");
const excelFilePath = path.join(__dirname, "sample.xlsx");

// Initialize Excel workbook and worksheet for IFSC details
const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("IFSC Details");

// Add header to the worksheet
worksheet.columns = [
  { header: "IFSC", key: "ifsc", width: 20 },
  { header: "BANK", key: "bank", width: 30 },
  { header: "BRANCH", key: "branch", width: 30 },
  { header: "STATUS", key: "status", width: 10 },
];

// Function to write IFSC details to Excel
async function writeToExcel(data) {
  worksheet.addRow(data);
  await workbook.xlsx.writeFile(outputFilePath);
  console.log(
    `Added to Excel: ${data.ifsc}, ${data.bank}, ${data.branch}, ${data.status}`
  );
}

// Function to process an individual IFSC code
async function processIFSC(ifscCode) {
  const isValid = ifsc.validate(ifscCode);
  if (isValid) {
    const details = await ifsc.fetchDetails(ifscCode);
    const result = {
      ifsc: ifscCode,
      bank: details.BANK,
      branch: details.BRANCH,
      status: "VALID",
    };
    await writeToExcel(result); // Write to Excel
    return result;
  } else {
    const result = {
      ifsc: ifscCode,
      bank: null,
      branch: null,
      status: "INVALID",
    };
    await writeToExcel(result); // Write to Excel
    return result;
  }
}

// Function to process the sample.xlsx file and append data to Excel
async function processExcel() {
  const excelWorkbook = new Excel.Workbook();
  await excelWorkbook.xlsx.readFile(excelFilePath);
  const worksheet = excelWorkbook.getWorksheet(1);
  const total = worksheet.rowCount;
  let count = 0;

  for (let i = 1; i <= total; i++) {
    const row = worksheet.getRow(i);
    const code = row.getCell(1).value;

    const isValid = ifsc.validate(code);
    if (isValid) {
      const details = await ifsc.fetchDetails(code);
      const result = {
        ifsc: code,
        bank: details.BANK,
        branch: details.BRANCH,
        status: "VALID",
      };
      await writeToExcel(result); // Write to Excel
    } else {
      const result = {
        ifsc: code,
        bank: null,
        branch: null,
        status: "INVALID",
      };
      await writeToExcel(result); // Write to Excel
    }
    count++;
    if (count % 20 === 0) {
      console.clear();
      console.log(
        `Processed ${count} out of ${total} rows from the Excel file.`
      );
    }
  }

  console.log("Excel file processing complete.");
}

// Initialize console input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
  terminal: true,
});

// Function to get bank details by region
async function getBanksByRegion(region) {
  const excelWorkbook = new Excel.Workbook();
  const worksheet = excelWorkbook.addWorksheet("Region Banks");
  worksheet.columns = [
    { header: "IFSC", key: "ifsc", width: 20 },
    { header: "BANK", key: "bank", width: 30 },
    { header: "BRANCH", key: "branch", width: 30 },
  ];

  try {
    // Make an API call to fetch banks by region (using BankCode API)
    const response = await axios.get(
      `https://www.bankcode.in/api/banks?state=${region}`
    );

    // Log the raw API response to debug
    console.log("Raw API Response:", response.data);

    const banks = response.data;
    if (banks.length === 0) {
      console.log(`No banks found for region: ${region}`);
    } else {
      // Filter banks based on region (case-insensitive)
      banks.forEach((bank) => {
        if (bank.branch.toLowerCase().includes(region.toLowerCase())) {
          worksheet.addRow({
            ifsc: bank.ifsc,
            bank: bank.bank,
            branch: bank.branch,
          });
        }
      });

      // Save the region-specific bank details in a new Excel file
      await excelWorkbook.xlsx.writeFile("region_output.xlsx");
      console.log(
        "Region-specific bank details have been saved to region_output.xlsx"
      );

      // Also log the results to the console
      console.log(`Banks in ${region}:`);
      banks.forEach((bank) => {
        if (bank.branch.toLowerCase().includes(region.toLowerCase())) {
          console.log(
            `IFSC: ${bank.ifsc}, Bank: ${bank.bank}, Branch: ${bank.branch}`
          );
        }
      });
    }
  } catch (error) {
    console.error("Error fetching bank data:", error);
  }
}

// Console input handler
function promptUser() {
  rl.question(
    "Enter an IFSC code or type 'region' to get banks by region (or type 'exit' to quit): ",
    async (input) => {
      if (input.toLowerCase() === "exit") {
        rl.close();
        console.log("Exiting program...");
        process.exit(0);
      } else if (input.toLowerCase() === "region") {
        rl.question("Enter the region: ", async (region) => {
          await getBanksByRegion(region);
          promptUser(); // Keep prompting for input
        });
      } else {
        await processIFSC(input); // Process IFSC code
        promptUser(); // Keep prompting for input
      }
    }
  );
}

// Web API endpoint for IFSC validation
app.post("/validate", async (req, res) => {
  const { ifsc: ifscCode } = req.body;
  if (!ifscCode) {
    return res.status(400).json({ error: "IFSC code is required" });
  }
  try {
    const result = await processIFSC(ifscCode);
    res.status(200).json(result);
  } catch (error) {
    res
      .status(500)
      .json({ error: "Internal server error", details: error.message });
  }
});

// Start web server
app.listen(PORT, () => {
  console.log(`Web server running at http://localhost:${PORT}`); // Logs the dynamically detected port
  console.log("You can use POST /validate to submit IFSC codes.");
});

// Start Excel file processing and console input
console.log("Processing the Excel file...");
processExcel().then(() => {
  console.log(
    "You can enter IFSC codes or a region via the console or use the web API."
  );
  promptUser();
});

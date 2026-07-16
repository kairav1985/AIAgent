console.log("Rows :", data.length);

// Debug
console.log("First Row:", JSON.stringify(data[0], null, 2));
console.log("Action:", data[0]["Action"]);
console.log("Expected Result:", data[0]["Expected Result"]);

// Excel Headers
const headers = [
    "Test Suite",
    "Name",
    "Importance",
    "Summary",
    "Precondition",
    "Action",
    "Expected Result",
    "Actual Result",
    "Test Step Status",
    "Test Case Status"
];

// Convert JSON → Excel Rows
const sheetData = [
    headers,
    ...data.map(tc => [
        tc["Test Suite"] || "",
        tc["Name"] || "",
        tc["Importance"] || "",
        tc["Summary"] || "",
        tc["Precondition"] || "",
        tc["Action"] || "",
        tc["Expected Result"] || "",
        tc["Actual Result"] || "",
        tc["Test Step Status"] || "",
        tc["Test Case Status"] || ""
    ])
];

// Create Worksheet
const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

// Optional Column Width
worksheet["!cols"] = [
    { wch: 25 }, // Test Suite
    { wch: 30 }, // Name
    { wch: 12 }, // Importance
    { wch: 40 }, // Summary
    { wch: 45 }, // Precondition
    { wch: 50 }, // Action
    { wch: 60 }, // Expected Result
    { wch: 20 }, // Actual Result
    { wch: 18 }, // Test Step Status
    { wch: 18 }  // Test Case Status
];

// Create Workbook
const workbook = XLSX.utils.book_new();

// Add Worksheet
XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Test Cases"
);
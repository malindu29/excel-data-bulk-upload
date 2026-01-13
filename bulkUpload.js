import XLSX from "xlsx";
import { faker } from "@faker-js/faker";
import path from "node:path";
import fs from "node:fs";

// ---------- CONFIG ----------
const ROW_COUNT = 500; // number of records to generate
const OUTPUT_DIR = "output";
const ID_TRACKER_FILE = path.join(OUTPUT_DIR, "last_student_id.txt");
const NIC_TRACKER_FILE = path.join(OUTPUT_DIR, "last_nic_index.txt");

// ---------- UTILS ----------
function getLastStudentId() {
  if (fs.existsSync(ID_TRACKER_FILE)) {
    const content = fs.readFileSync(ID_TRACKER_FILE, "utf-8").trim();
    const lastId = parseInt(content, 10);
    return isNaN(lastId) ? 0 : lastId;
  }
  return 0;
}

function saveLastStudentId(id) {
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
  }
  fs.writeFileSync(ID_TRACKER_FILE, String(id), "utf-8");
}

function getLastNicIndex() {
  if (fs.existsSync(NIC_TRACKER_FILE)) {
    const content = fs.readFileSync(NIC_TRACKER_FILE, "utf-8").trim();
    const lastIndex = parseInt(content, 10);
    return isNaN(lastIndex) ? 0 : lastIndex;
  }
  return 0;
}

function saveLastNicIndex(index) {
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
  }
  fs.writeFileSync(NIC_TRACKER_FILE, String(index), "utf-8");
}

function sanitizeName(name) {
  // Remove any characters that are not letters or spaces
  let sanitized = name.replace(/[^a-zA-Z\s]/g, "").trim();
  
  // Ensure name is between 3 and 50 characters
  if (sanitized.length < 3) {
    // Pad with additional text to meet minimum length
    sanitized = sanitized + "son";
  }
  
  if (sanitized.length > 50) {
    sanitized = sanitized.substring(0, 50);
  }
  
  return sanitized;
}

function generateStudentId(index) {
  return `STUDENT-ID${String(index).padStart(5, "0")}`;
}

function generateNIC(index) {
  // Sri Lankan NIC format: Old (9 digits + V/X) or New (12 digits)
  // Regex: ^(?:[0-9]{9}[VXvx]|[0-9]{12})$
  
  // Alternate between old and new formats
  if (index % 2 === 0) {
    // New format: 12 digits
    const year = 2000 + (index % 24);
    const dayOfYear = String(index % 365 + 1).padStart(3, "0");
    const serial = String(index).padStart(5, "0");
    return `${year}${dayOfYear}${serial}`;
  } else {
    // Old format: 9 digits + V or X
    const year = String(91 + (index % 9)).padStart(2, "0");
    const dayOfYear = String(index % 365 + 1).padStart(3, "0");
    const serial = String(index).padStart(4, "0");
    const suffix = index % 3 === 0 ? "X" : "V";
    return `${year}${dayOfYear}${serial}${suffix}`;
  }
}

// ---------- MAIN ----------
function generateBulkExcel(templatePath) {
  if (!fs.existsSync(templatePath)) {
    console.error("❌ Template file not found");
    return;
  }

  const workbook = XLSX.readFile(templatePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Convert sheet to JSON (header-based)
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  // Get the last used Student ID and NIC Index to continue from there
  let currentId = getLastStudentId();
  let currentNicIndex = getLastNicIndex();

  const newRows = [];

  for (let i = 1; i <= ROW_COUNT; i++) {
    currentId++; // Increment to get the next unique ID
    currentNicIndex++; // Increment to get the next unique NIC
    
    const firstName = sanitizeName(faker.person.firstName());
    const lastName = sanitizeName(faker.person.lastName());
    
    // Temporary email service domains that provide real working inboxes
    const tempMailDomains = [
      'guerrillamail.com',
      'maildrop.cc',
      'mailinator.com'
    ];

    newRows.push({
      "Student ID*": generateStudentId(currentId),
      "First Name*": firstName,
      "Last Name*": lastName,
      "Email*": faker.internet.email({
        firstName,
        lastName,
        provider: faker.helpers.arrayElement(tempMailDomains),
      }).toLowerCase(),
      "National Identification Number*": generateNIC(currentNicIndex),
    });
  }

  // Save the last used ID and NIC Index for next run
  saveLastStudentId(currentId);
  saveLastNicIndex(currentNicIndex);

  // Convert back to sheet
  const newSheet = XLSX.utils.json_to_sheet(newRows, {
    header: Object.keys(newRows[0]),
  });

  workbook.Sheets[sheetName] = newSheet;

  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
  }

  const outputPath = path.join(
    OUTPUT_DIR,
    `bulk_upload_${Date.now()}.xlsx`
  );

  XLSX.writeFile(workbook, outputPath);
  console.log(`✅ Bulk Excel created: ${outputPath}`);
}

// ---------- CLI ----------
const templatePath = process.argv[2];

if (!templatePath) {
  console.log("Usage: node bulkUpload.js <template.xlsx>");
  process.exit(1);
}

generateBulkExcel(templatePath);

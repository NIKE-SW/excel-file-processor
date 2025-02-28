import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";

/// Constants & Configuration

const WOOD_TYPES = ["redwood", "whitewood"];
const ID_PREFIXES = {
  redwood: "R1",
  whitewood: "W1",
};

/// Utility Functions (Independent, No State)

// Normalize status value
const normalizeStatus = (status) => {
  if (status.includes("S/F")) return "SF";
  if (status.includes("IV")) return "IV";
  if (status === "V") return "V";
  return status;
};

// Find next non-empty cell in an array
const findNextNonEmptyCell = (array, index) => {
  for (let i = index + 1; i < array.length; i++) {
    if (array[i]) return array[i];
  }
  return null;
};

// Reads an Excel file and returns a Promise that resolves with its ArrayBuffer
const readFileAsArrayBuffer = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    // Resolve the promise when reading is complete
    reader.onload = () => resolve(reader.result);

    // Reject the promise if there's an error
    reader.onerror = () => reject(new Error("Failed to read the file."));

    // Start reading the file as an ArrayBuffer
    reader.readAsArrayBuffer(file);
  });
};

// Generate an item object from a row and context
const getItemObject = (row, context) => {
  const amount = findNextNonEmptyCell(row, context.packIndex);
  const amountIndex = row.findIndex((cell) => cell === amount);
  return {
    id: `${ID_PREFIXES[context.woodType]}${context.width}${
      context.dimensions[amountIndex]
    }0`
      .replaceAll(".", "")
      .replaceAll("*", ""),
    packId: Number(row[context.packIndex]),
    amount,
    status: normalizeStatus(context.status),
  };
};

/// State-Dependent Functions

// Process an Excel workbook and extract relevant data
const processWorkbook = (workbook, setData) => {
  const sheetsList = workbook.SheetNames;
  const output = [];
  let context = {
    woodType: null,
    width: null,
    status: null,
    dimensions: null,
    packIndex: null,
  };

  sheetsList.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    // Converts the sheet content into an array of arrays
    const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    sheetData.forEach((row, index) => {
      const woodType = row.find(
        (cell) => WOOD_TYPES.some((wood) => new RegExp(`(.+)wood`).test(cell))
        /// WOOD_TYPES.some((wood) => wood === cell), more efficient if "redwood" and "whitewood" are the only wood types. According to this excel yes but I am not sure in general.
      );
      if (woodType) {
        const woodTypeIndex = row.findIndex((cell) => cell === woodType);
        context.woodType = woodType;
        context.dimensions = sheetData[index + 1] || [];
        context.width = sheetData[index + 2]?.[0] || "";
        context.status = findNextNonEmptyCell(row, woodTypeIndex);
        context.packIndex = context.dimensions.findIndex(
          (cell) => cell === "PACK"
        );
        return;
      }

      const packId = row[context.packIndex];
      if (packId && new RegExp(`[0-9]{8}\\.[0-9]`).test(packId)) {
        output.push(getItemObject(row, context));
      }

      const width = sheetData[index + 1]?.[0];
      if (width && new RegExp(/[0-9]+\*[0-9]+/).test(width)) {
        context.width = width;
      }
    });
  });

  setData(output);
};

// Export processed data to an Excel file
const exportToExcel = (data) => {
  if (data.length === 0) {
    alert("No data to export!");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  // Append the sheet to the workbook
  XLSX.utils.book_append_sheet(wb, ws, "Processed Data");
  
  XLSX.writeFile(wb, "processed_output.xlsx");
};

/// Event Handlers (Trigger Functions)

// Handle file upload, read file, and process workbook
const handleFileUpload = async (event, setData, setFileName) => {
  const file = event.target.files[0];
  if (!file) {
    alert("Please select a valid Excel file.");
    return;
  }

  try {
    // Wait for the file to be read as an ArrayBuffer
    const bufferArray = await readFileAsArrayBuffer(file);
    // Converts the binary buffer into a JavaScript object with organized sheet data
    const workbook = XLSX.read(bufferArray, { type: "buffer" });

    // Process the workbook
    processWorkbook(workbook, setData);

    // Set the file name only if successful
    setFileName(file.name);
    console.log("File processed successfully!");
  } catch (error) {
    console.error("Error reading or processing the file:", error);
    alert("Failed to process the Excel file.");
  }
};

/// React Component

// ExcelProcessor component handles file input and displays processed data
const ExcelProcessor = () => {
  const [data, setData] = useState([]); // Stores processed data
  const [fileName, setFileName] = useState(""); // Stores uploaded file name
  const fileInputRef = useRef(null);

  // Opens file selection dialog
  const handleButtonClick = () => {
    fileInputRef.current.click();
  };

  return (
    <div className="p-4">
      {/* Hidden File Input */}
      <input
        type="file"
        accept=".xls,.xlsx"
        onChange={(event) => handleFileUpload(event, setData, setFileName)}
        ref={fileInputRef}
        className="hidden"
        style={{ display: "none" }} // Ensure it is fully hidden
      />

      {/* Styled Upload Button */}
      <button
        onClick={handleButtonClick}
        className="px-4 py-2 bg-blue-900 text-white font-semibold rounded-lg shadow-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
      >
        Upload Excel File
      </button>

      {/* File Name Display + Export Button */}
      {fileName && (
        <div className="mt-2">
          <p className="text-sm text-gray-900">Uploaded: {fileName}</p>

          {/* Export Button */}
          {data.length > 0 && (
            <button
              onClick={() => exportToExcel(data)}
              className="mt-2 px-4 py-2 bg-green-900 text-white font-semibold rounded-lg shadow-md hover:bg-green-600"
            >
              Export to Excel File
            </button>
          )}
        </div>
      )}

      {fileName && (
        <>
          {data.length > 0 ? (
            <table className="border-collapse border border-gray-400 mt-4 w-full">
              <thead>
                <tr>
                  {Object.keys(data[0]).map((header, index) => (
                    <th
                      key={index}
                      className="border border-gray-300 p-2 bg-gray-200"
                    >
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.map((row, rowIndex) => (
                  <tr key={rowIndex} className="hover:bg-gray-100">
                    {Object.values(row).map((cell, colIndex) => (
                      <td key={colIndex} className="border border-gray-300 p-2">
                        {cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <p className="text-center text-gray-600 mt-4">
              No data available to display.
            </p>
          )}
        </>
      )}
    </div>
  );
};

export default ExcelProcessor;

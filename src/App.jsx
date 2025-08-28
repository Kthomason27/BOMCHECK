import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";

export default function NpdBomChecker() {
  const [npdFile, setNpdFile] = useState(null);
  const [bomFile, setBomFile] = useState(null);
  const [comparisonTable, setComparisonTable] = useState([]);
  const [npdColumns, setNpdColumns] = useState([]);
  const [bomColumns, setBomColumns] = useState([]);
  const [selectedNpdColumns, setSelectedNpdColumns] = useState([]);
  const [selectedBomColumns, setSelectedBomColumns] = useState([]);
  const [npdSheetNames, setNpdSheetNames] = useState([]);
  const [bomSheetNames, setBomSheetNames] = useState([]);
  const [selectedNpdSheets, setSelectedNpdSheets] = useState([]);
  const [selectedBomSheets, setSelectedBomSheets] = useState([]);
  const [kitFilter, setKitFilter] = useState("all");
  const [positionFilter, setpositionFilter] = useState("all");
  const [sortConfig, setSortConfig] = useState({
    column: null,
    direction: "asc",
    source: "npd",
  });

  // Get unique values for filters
  const uniqueKits = React.useMemo(() => {
    const kits = new Set();
    comparisonTable.forEach((row) => {
      const npdKit = row.npd["Kit #"];
      const bomKit = row.bom["Kit Item"];
      if (npdKit) kits.add(String(npdKit).toLowerCase());
      if (bomKit) kits.add(String(bomKit).toLowerCase());
    });
    return Array.from(kits).sort();
  }, [comparisonTable]);

  const uniquePositions = React.useMemo(() => {
    const positions = new Set();
    comparisonTable.forEach((row) => {
      const npdItem = row.npd["Position #"];
      const bomItem = row.bom["Position"];
      if (npdItem) positions.add(String(npdItem).toLowerCase());
      if (bomItem) positions.add(String(bomItem).toLowerCase());
    });
    return Array.from(positions).sort();
  }, [comparisonTable]);

  const filteredTable = React.useMemo(() => {
    return comparisonTable.filter((row) => {
      // Kit filter
      if (kitFilter !== "all") {
        const npdKit = String(row.npd["Kit #"]).toLowerCase();
        const bomKit = String(row.bom["Kit Item"]).toLowerCase();
        if (npdKit !== kitFilter && bomKit !== kitFilter) {
          return false;
        }
      }

      // Item filter
      if (positionFilter !== "all") {
        const npdItem = String(row.npd["Position #"]).toLowerCase();
        const bomItem = String(row.bom["Position"]).toLowerCase();
        if (npdItem !== positionFilter && bomItem !== positionFilter) {
          return false;
        }
      }

      return true;
    });
  }, [comparisonTable, kitFilter, positionFilter]);

  const sortedTable = React.useMemo(() => {
    if (!sortConfig.column) return filteredTable;

    const sorted = [...filteredTable].sort((a, b) => {
      const aVal =
        sortConfig.source === "npd"
          ? a.npd[sortConfig.column] || ""
          : a.bom[sortConfig.column] || "";
      const bVal =
        sortConfig.source === "npd"
          ? b.npd[sortConfig.column] || ""
          : b.bom[sortConfig.column] || "";

      if (aVal < bVal) return sortConfig.direction === "asc" ? -1 : 1;
      if (aVal > bVal) return sortConfig.direction === "asc" ? 1 : -1;
      return 0;
    });

    return sorted;
  }, [filteredTable, sortConfig]);

  const readExcelFile = async (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const result = workbook.SheetNames.map((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          const headers = json[0];
          const rows = json.slice(1).map((row) => {
            const obj = {};
            headers.forEach((header, index) => {
              obj[header] = row[index];
            });
            return obj;
          });
          return { sheetName, headers, data: rows };
        });
        resolve({ sheets: result, sheetNames: workbook.SheetNames });
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  };

  const handleCompare = async () => {
    try {
      const npdResult = await readExcelFile(npdFile);
      const bomResult = await readExcelFile(bomFile);

      // Flatten selected sheets into one array of rows
      const npdData = selectedNpdSheets.flatMap((sheetName) => {
        const sheet = npdResult.sheets.find((s) => s.sheetName === sheetName);
        return (
          sheet?.data.map((row) => ({ ...row, sourceSheet: sheetName })) || []
        );
      });

      const bomData = selectedBomSheets.flatMap((sheetName) => {
        const sheet = bomResult.sheets.find((s) => s.sheetName === sheetName);
        return (
          sheet?.data.map((row) => ({ ...row, sourceSheet: sheetName })) || []
        );
      });

      // Match key includes Kit #, Position #, and Item #
      const matchKey = (row) =>
        `${String(row["Kit #"] || "").toLowerCase()}|${String(
          row["Position #"] || ""
        ).toLowerCase()}|${String(
          row["Item #"] || row["h2m Item #"] || ""
        ).toLowerCase()}`;

      const bomMatchKey = (row) =>
        `${String(row["Kit Item"] || "").toLowerCase()}|${String(
          row["Position"] || ""
        ).toLowerCase()}|${String(row["Item number"] || "").toLowerCase()}`;

      // Create a map of BOM rows by match key
      const bomMap = new Map();
      for (const row of bomData) {
        bomMap.set(bomMatchKey(row), row);
      }
      const processedBomKeys = new Set();
      // Filter out non-data rows from NPD
      const filteredNpdData = npdData.filter(
        (row) => row["Kit #"] && (row["Item #"] || row["h2m Item #"])
      );

      // Deduplicate NPD rows by match key
      const uniqueNpdData = Array.from(
        new Map(filteredNpdData.map((row) => [matchKey(row), row])).values()
      );

      // Update field comparison to use correct BOM column names
      const fieldComparisons = [
        { npd: "h2m Item #", bom: "Item number" },
        { npd: "Description", bom: "Product name" },
        { npd: "Kit Qty", bom: "Quantity" },
        { npd: "Kit #", bom: "Kit Item" },
        { npd: "BOM/Assembly Notes", bom: "BOM notes" },
        { npd: "Product Notes", bom: "BOM Product notes" },
      ];

      // Build comparison table
      const table = [];

      for (const npdRow of uniqueNpdData) {
        const key = matchKey(npdRow);
        const bomRow = bomMap.get(key);

        let match = true;
        const differences = {};

        if (bomRow) {
          processedBomKeys.add(bomMatchKey(bomRow));

          for (const { npd: npdField, bom: bomField } of fieldComparisons) {
            const npdValue = String(npdRow[npdField] || "");
            const bomValue = String(bomRow[bomField] || "");

            if (npdValue.toLowerCase() !== bomValue.toLowerCase()) {
              match = false;
              differences[npdField] = true;
              differences[bomField] = true;
            }
          }
        } else {
          match = false;
        }
        table.push({
          key: key,
          npd: npdRow,
          bom: bomRow || {},
          match,
          differences,
        });
      }
      for (const bomRow of bomData) {
        const bomKey = bomMatchKey(bomRow);

        if (!processedBomKeys.has(bomKey)) {
          table.push({
            key: bomKey,
            npd: {}, // Empty NPD data
            bom: bomRow,
            match: false, // BOM-only rows are always mismatches
            differences: {}, // No specific field differences since NPD is empty
          });
        }
      }

      setComparisonTable(table);
    } catch (error) {
      console.error("Error comparing files:", error);
    }
  };

  const handleFileChange = async (e, type) => {
    const file = e.target.files[0];
    const result = await readExcelFile(file);

    if (type === "npd") {
      setNpdFile(file);
      setNpdColumns(result.sheets[0].headers);
      setSelectedNpdColumns(result.sheets[0].headers);
      setNpdSheetNames(result.sheetNames);
      setSelectedNpdSheets(result.sheetNames);
    } else {
      setBomFile(file);
      setBomColumns(result.sheets[0].headers);
      setSelectedBomColumns(result.sheets[0].headers);
      setBomSheetNames(result.sheetNames);
      setSelectedBomSheets(result.sheetNames);
    }
  };

  const renderTableHeader = () => {
    return (
      <tr>
        {selectedNpdColumns.map((col) => (
          <th
            key={`npd-${col}`}
            onClick={() => {
              setSortConfig((prev) => ({
                column: col,
                direction:
                  prev.column === col && prev.direction === "asc"
                    ? "desc"
                    : "asc",
                source: "npd",
              }));
            }}
            style={{ cursor: "pointer", userSelect: "none" }}
          >
            NPD: {col}{" "}
            {sortConfig.column === col && sortConfig.source === "npd"
              ? sortConfig.direction === "asc"
                ? "🔼"
                : "🔽"
              : ""}
          </th>
        ))}
        <th>Match</th>
        {selectedBomColumns.map((col) => (
          <th
            key={`bom-${col}`}
            onClick={() => {
              setSortConfig((prev) => ({
                column: col,
                direction:
                  prev.column === col && prev.direction === "asc"
                    ? "desc"
                    : "asc",
                source: "bom",
              }));
            }}
            style={{ cursor: "pointer", userSelect: "none" }}
          >
            BOM: {col}{" "}
            {sortConfig.column === col && sortConfig.source === "bom"
              ? sortConfig.direction === "asc"
                ? "🔼"
                : "🔽"
              : ""}
          </th>
        ))}
      </tr>
    );
  };

  // Helper function to render cell content with missing data indicators
  const renderCellContent = (value, isRowMissing, isDifferent) => {
    if (isRowMissing) {
      return (
        <span
          style={{
            color: "#999",
            fontStyle: "italic",
            fontWeight: isDifferent ? "bold" : "normal",
          }}
        >
          [MISSING]
        </span>
      );
    }

    return (
      <span style={{ fontWeight: isDifferent ? "bold" : "normal" }}>
        {value || ""}
      </span>
    );
  };

  const renderTableRow = ({ npd, bom, match, differences }, index) => {
    // Reverse the logic: mismatched rows are green with ✔️, matched rows are red with ❌
    const isMismatch = !match;
    const isNpdMissing = Object.keys(npd).length === 0;
    const isBomMissing = Object.keys(bom).length === 0;

    return (
      <tr
        key={index}
        style={{ backgroundColor: isMismatch ? "#f8d7da" : "#c8f7c5" }}
      >
        {selectedNpdColumns.map((col) => (
          <td key={`npd-${col}`}>
            {renderCellContent(npd[col], isNpdMissing, differences[col])}
          </td>
        ))}
        <td style={{ textAlign: "center" }}>{isMismatch ? "❌" : "✔️"}</td>
        {selectedBomColumns.map((col) => (
          <td key={`bom-${col}`}>
            {renderCellContent(bom[col], isBomMissing, differences[col])}
          </td>
        ))}
      </tr>
    );
  };

  return (
    <div style={{ padding: "20px" }}>
      <h1>NPD vs BOM Comparison</h1>

      <div>
        <label>NPD File: </label>
        <input type="file" onChange={(e) => handleFileChange(e, "npd")} />
        <div>
          {npdColumns.map((col) => (
            <label key={col} style={{ marginRight: "10px" }}>
              <input
                type="checkbox"
                checked={selectedNpdColumns.includes(col)}
                onChange={(e) => {
                  const newCols = e.target.checked
                    ? [...selectedNpdColumns, col]
                    : selectedNpdColumns.filter((c) => c !== col);
                  setSelectedNpdColumns(newCols);
                }}
              />
              {col}
            </label>
          ))}
        </div>
      </div>

      <div>
        <label>BOM File: </label>
        <input type="file" onChange={(e) => handleFileChange(e, "bom")} />
        <div>
          {bomColumns.map((col) => (
            <label key={col} style={{ marginRight: "10px" }}>
              <input
                type="checkbox"
                checked={selectedBomColumns.includes(col)}
                onChange={(e) => {
                  const newCols = e.target.checked
                    ? [...selectedBomColumns, col]
                    : selectedBomColumns.filter((c) => c !== col);
                  setSelectedBomColumns(newCols);
                }}
              />
              {col}
            </label>
          ))}
        </div>
      </div>

      <button onClick={handleCompare} style={{ margin: "20px 0" }}>
        Compare Files
      </button>

      {comparisonTable.length > 0 && (
        <div>
          {/* Filter controls */}
          <div
            style={{
              marginBottom: "20px",
              display: "flex",
              gap: "20px",
              alignItems: "center",
            }}
          >
            <div>
              <label>Filter by Kit #: </label>
              <select
                value={kitFilter}
                onChange={(e) => setKitFilter(e.target.value)}
                style={{ marginLeft: "5px" }}
              >
                <option value="all">All Kits</option>
                {uniqueKits.map((kit) => (
                  <option key={kit} value={kit}>
                    {kit}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label>Filter by Position #: </label>
              <select
                value={positionFilter}
                onChange={(e) => setpositionFilter(e.target.value)}
                style={{ marginLeft: "5px" }}
              >
                <option value="all">All Positions</option>
                {uniquePositions.map((item) => (
                  <option key={item} value={item}>
                    {item}
                  </option>
                ))}
              </select>
            </div>

            <div style={{ marginLeft: "auto", color: "#666" }}>
              Showing {sortedTable.length} of {comparisonTable.length} rows
            </div>
          </div>

          <table
            border="1"
            cellPadding="5"
            style={{ borderCollapse: "collapse", width: "100%" }}
          >
            <thead>{renderTableHeader()}</thead>
            <tbody>{sortedTable.map(renderTableRow)}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

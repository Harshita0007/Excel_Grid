import * as XLSX from "xlsx";

export const createColumnHeaders = (numCols) => {
  const headers = [];
  for (let i = 0; i < numCols; i++) {
    let header = "",
      num = i;
    while (num >= 0) {
      header = String.fromCharCode(65 + (num % 26)) + header;
      num = Math.floor(num / 26) - 1;
    }
    headers.push(header);
  }
  return headers;
};

export const createInitialGridData = (
  initialData,
  initialRows,
  initialColumns
) => {
  const gridData = [];
  const actualRows = Math.max(initialRows, initialData.length);
  const actualColumns = Math.max(
    initialColumns,
    Math.max(...initialData.map((row) => row?.length || 0), 0)
  );

  for (let i = 0; i < actualRows; i++) {
    gridData[i] = {
      id: `row_${i}_${Date.now()}_${Math.random()}`,
      cells: [],
    };
    for (let j = 0; j < actualColumns; j++) {
      gridData[i].cells[j] = initialData[i]?.[j] || "";
    }
  }
  return gridData;
};

export const filterData = (data, filters) => {
  return data.filter((row) => {
    return Object.entries(filters).every(([colIndex, filterConfig]) => {
      if (!filterConfig || !filterConfig.value) return true;

      const cellValue = row.cells[parseInt(colIndex)] || "";
      const filterValue = filterConfig.value.toLowerCase();
      const filterType = filterConfig.type || "contains";

      switch (filterType) {
        case "contains":
          return cellValue.toString().toLowerCase().includes(filterValue);
        case "equals":
          return cellValue.toString().toLowerCase() === filterValue;
        case "startsWith":
          return cellValue.toString().toLowerCase().startsWith(filterValue);
        case "endsWith":
          return cellValue.toString().toLowerCase().endsWith(filterValue);
        case "greaterThan":
          const numValue = parseFloat(cellValue);
          const filterNum = parseFloat(filterValue);
          return !isNaN(numValue) && !isNaN(filterNum) && numValue > filterNum;
        case "lessThan":
          const numValue2 = parseFloat(cellValue);
          const filterNum2 = parseFloat(filterValue);
          return (
            !isNaN(numValue2) && !isNaN(filterNum2) && numValue2 < filterNum2
          );
        case "notEmpty":
          return cellValue.toString().trim() !== "";
        case "empty":
          return cellValue.toString().trim() === "";
        default:
          return cellValue.toString().toLowerCase().includes(filterValue);
      }
    });
  });
};

export const sortData = (data, sortConfig) => {
  if (!sortConfig.key && sortConfig.key !== 0) return data;

  // Separate rows with empty cells in the sort column
  const nonEmptyRows = [];
  const emptyRows = [];
  data.forEach((row, idx) => {
    const value = row.cells[sortConfig.key];
    if (
      value === undefined ||
      value === null ||
      value.toString().trim() === ""
    ) {
      emptyRows.push({ row, idx });
    } else {
      nonEmptyRows.push({ row, idx });
    }
  });

  // Sort only non-empty rows
  nonEmptyRows.sort((aObj, bObj) => {
    const aValue = aObj.row.cells[sortConfig.key] || "";
    const bValue = bObj.row.cells[sortConfig.key] || "";
    switch (sortConfig.type) {
      case "number":
        const aNum = parseFloat(aValue) || 0;
        const bNum = parseFloat(bValue) || 0;
        return sortConfig.direction === "asc" ? aNum - bNum : bNum - aNum;
      case "date":
        const aDate = new Date(aValue);
        const bDate = new Date(bValue);
        const aTime = isNaN(aDate.getTime()) ? 0 : aDate.getTime();
        const bTime = isNaN(bDate.getTime()) ? 0 : bDate.getTime();
        return sortConfig.direction === "asc" ? aTime - bTime : bTime - aTime;
      case "text":
      default:
        const aNum2 = parseFloat(aValue);
        const bNum2 = parseFloat(bValue);
        if (!isNaN(aNum2) && !isNaN(bNum2)) {
          return sortConfig.direction === "asc" ? aNum2 - bNum2 : bNum2 - aNum2;
        }
        const aStr = aValue.toString().toLowerCase();
        const bStr = bValue.toString().toLowerCase();
        if (aStr < bStr) return sortConfig.direction === "asc" ? -1 : 1;
        if (aStr > bStr) return sortConfig.direction === "asc" ? 1 : -1;
        return 0;
    }
  });

  // Reconstruct the array, preserving empty cell positions
  const result = Array(data.length);
  let nonEmptyIdx = 0;
  let emptyIdx = 0;
  for (let i = 0; i < data.length; i++) {
    if (emptyRows.length > 0 && emptyRows[0].idx === i) {
      result[i] = emptyRows.shift().row;
    } else {
      result[i] = nonEmptyRows[nonEmptyIdx++].row;
    }
  }
  return result;
};

export const getSortIcon = (colIndex, sortConfig) => {
  if (sortConfig.key !== colIndex) return "⇅";
  return sortConfig.direction === "asc" ? "↑" : "↓";
};

export const insertRowAbove = (data, rowIndex, numColumns) => {
  const newRow = {
    id: `row_${Date.now()}_${Math.random()}`,
    cells: Array(numColumns).fill(""),
  };
  const newData = [...data];
  newData.splice(rowIndex, 0, newRow);
  return newData;
};

export const insertRowBelow = (data, rowIndex, numColumns) => {
  const newRow = {
    id: `row_${Date.now()}_${Math.random()}`,
    cells: Array(numColumns).fill(""),
  };
  const newData = [...data];
  newData.splice(rowIndex + 1, 0, newRow);
  return newData;
};

export const insertColumnLeft = (data, colIndex) => {
  return data.map((row) => {
    const newCells = [...row.cells];
    newCells.splice(colIndex, 0, "");
    return { ...row, cells: newCells };
  });
};

export const insertColumnRight = (data, colIndex) => {
  return data.map((row) => {
    const newCells = [...row.cells];
    newCells.splice(colIndex + 1, 0, "");
    return { ...row, cells: newCells };
  });
};

export const insertMultipleRows = (data, startIndex, count, numColumns) => {
  const newRows = [];
  for (let i = 0; i < count; i++) {
    newRows.push({
      id: `row_${Date.now()}_${Math.random()}_${i}`,
      cells: Array(numColumns).fill(""),
    });
  }
  const newData = [...data];
  newData.splice(startIndex, 0, ...newRows);
  return newData;
};

export const insertMultipleColumns = (data, startIndex, count) => {
  return data.map((row) => {
    const newCells = [...row.cells];
    const columnsToInsert = Array(count).fill("");
    newCells.splice(startIndex, 0, ...columnsToInsert);
    return { ...row, cells: newCells };
  });
};

export const addMultipleColumns = (data, count = 5) => {
  return data.map((row) => ({
    ...row,
    cells: [...row.cells, ...Array(count).fill("")],
  }));
};

export const insertMultipleColumnsLeft = (data, colIndex, count = 3) => {
  return insertMultipleColumns(data, colIndex, count);
};

export const insertMultipleColumnsRight = (data, colIndex, count = 3) => {
  return insertMultipleColumns(data, colIndex + 1, count);
};

export const insertMultipleRowsAbove = (
  data,
  rowIndex,
  count = 5,
  numColumns
) => {
  return insertMultipleRows(data, rowIndex, count, numColumns);
};

export const insertMultipleRowsBelow = (
  data,
  rowIndex,
  count = 5,
  numColumns
) => {
  return insertMultipleRows(data, rowIndex + 1, count, numColumns);
};

export const getUniqueColumnValues = (data, columnIndex) => {
  const uniqueValues = new Set();
  data.forEach((row) => {
    const value = row.cells[columnIndex] || "";
    if (value.toString().trim() !== "") {
      uniqueValues.add(value.toString());
    }
  });
  return Array.from(uniqueValues).sort();
};

export const detectColumnDataType = (data, columnIndex) => {
  const values = data
    .map((row) => row.cells[columnIndex] || "")
    .filter((val) => val.toString().trim() !== "");

  if (values.length === 0) return "text";

  const numericCount = values.filter((val) => !isNaN(parseFloat(val))).length;
  const dateCount = values.filter(
    (val) => !isNaN(new Date(val).getTime())
  ).length;

  const numericRatio = numericCount / values.length;
  const dateRatio = dateCount / values.length;

  if (numericRatio > 0.8) return "number";
  if (dateRatio > 0.8) return "date";
  return "text";
};

export const getFilterPresets = () => [
  { label: "Contains", value: "contains", icon: "🔍" },
  { label: "Equals", value: "equals", icon: "=" },
  { label: "Starts With", value: "startsWith", icon: "🔤" },
  { label: "Ends With", value: "endsWith", icon: "🔚" },
  { label: "Greater Than", value: "greaterThan", icon: ">" },
  { label: "Less Than", value: "lessThan", icon: "<" },
  { label: "Not Empty", value: "notEmpty", icon: "📄" },
  { label: "Empty", value: "empty", icon: "📋" },
];

export const updateCellValue = (data, rowId, colIndex, value) => {
  return data.map((row) => {
    if (row.id === rowId) {
      const newCells = [...row.cells];
      newCells[colIndex] = value;
      return { ...row, cells: newCells };
    }
    return row;
  });
};

export const findRowIndexById = (data, rowId) => {
  if (!data || !Array.isArray(data)) return -1;
  return data.findIndex((row) => row && row.id === rowId);
};

export const clearCellData = (data, rowIdentifier, colIndex) => {
  console.log("Clear cell data called:", {
    rowIdentifier,
    colIndex,
    dataLength: data.length,
  });

  return data.map((row, rIndex) => {
    const isMatch =
      typeof rowIdentifier === "number"
        ? rIndex === rowIdentifier
        : row.id === rowIdentifier;

    if (isMatch) {
      const newCells = [...row.cells];

      if (colIndex >= 0 && colIndex < newCells.length) {
        console.log(
          "Clearing cell at row",
          rIndex,
          "col",
          colIndex,
          "old value:",
          newCells[colIndex]
        );
        newCells[colIndex] = "";
      }
      return { ...row, cells: newCells };
    }
    return row;
  });
};

export const clearCellDataById = (data, rowId, colIndex) => {
  console.log("Clear cell by ID:", { rowId, colIndex });

  return data.map((row) => {
    if (row.id === rowId) {
      const newCells = [...row.cells];
      if (colIndex >= 0 && colIndex < newCells.length) {
        console.log("Clearing cell by ID, old value:", newCells[colIndex]);
        newCells[colIndex] = "";
      }
      return { ...row, cells: newCells };
    }
    return row;
  });
};

export const clearCellDataByIndex = (data, rowIndex, colIndex) => {
  console.log("Clear cell by index:", {
    rowIndex,
    colIndex,
    dataLength: data.length,
  });

  if (rowIndex < 0 || rowIndex >= data.length) {
    console.error("Invalid row index:", rowIndex);
    return data;
  }

  return data.map((row, rIndex) => {
    if (rIndex === rowIndex) {
      const newCells = [...row.cells];
      if (colIndex >= 0 && colIndex < newCells.length) {
        console.log("Clearing cell by index, old value:", newCells[colIndex]);
        newCells[colIndex] = "";
      }
      return { ...row, cells: newCells };
    }
    return row;
  });
};

export const clearCellsData = (data, cellPositions) => {
  return data.map((row, rIndex) => {
    const newCells = [...row.cells];
    cellPositions.forEach(({ row: targetRow, col: targetCol }) => {
      if (rIndex === targetRow) {
        newCells[targetCol] = "";
      }
    });
    return { ...row, cells: newCells };
  });
};

export const clearRowData = (data, rowIndex) => {
  return data.map((row, rIndex) => {
    if (rIndex === rowIndex) {
      return {
        ...row,
        cells: row.cells.map(() => ""),
      };
    }
    return row;
  });
};

export const clearColumnData = (data, colIndex) => {
  return data.map((row) => {
    const newCells = [...row.cells];
    newCells[colIndex] = "";
    return { ...row, cells: newCells };
  });
};

export const addNewRow = (data, numColumns) => {
  const newRow = {
    id: `row_${data.length}_${Date.now()}_${Math.random()}`,
    cells: Array(numColumns).fill(""),
  };
  return [...data, newRow];
};

export const addNewColumn = (data) => {
  return data.map((row) => ({
    ...row,
    cells: [...row.cells, ""],
  }));
};

export const deleteRow = (data, rowIndex) => {
  if (data.length <= 1) return data;
  return data.filter((_, index) => index !== rowIndex);
};

export const deleteColumn = (data, colIndex) => {
  const numColumns = data[0]?.cells?.length || 0;
  if (numColumns <= 1) return data;

  return data.map((row) => ({
    ...row,
    cells: row.cells.filter((_, index) => index !== colIndex),
  }));
};

export const clearAllData = (data) => {
  return data.map((row) => ({
    ...row,
    cells: row.cells.map(() => ""),
  }));
};

export const getNextFocusedCell = (
  focusedCell,
  direction,
  maxRows,
  maxCols
) => {
  let newRow = focusedCell.row;
  let newCol = focusedCell.col;

  switch (direction) {
    case "up":
      newRow = newRow > 0 ? newRow - 1 : maxRows - 1;
      break;
    case "down":
      newRow = newRow < maxRows - 1 ? newRow + 1 : 0;
      break;
    case "left":
      newCol = newCol > 0 ? newCol - 1 : maxCols - 1;
      break;
    case "right":
      newCol = newCol < maxCols - 1 ? newCol + 1 : 0;
      break;
    case "tab":
      if (newCol < maxCols - 1) {
        newCol = newCol + 1;
      } else {
        newCol = 0;
        newRow = newRow < maxRows - 1 ? newRow + 1 : 0;
      }
      break;
    case "shift-tab":
      if (newCol > 0) {
        newCol = newCol - 1;
      } else {
        newCol = maxCols - 1;
        newRow = newRow > 0 ? newRow - 1 : maxRows - 1;
      }
      break;
    case "enter":
      newRow = newRow < maxRows - 1 ? newRow + 1 : 0;
      break;
    default:
      break;
  }

  return { row: newRow, col: newCol };
};

export const saveToFile = (data, fileName) => {
  const dataToSave = {
    fileName,
    data: data.map((row) => row.cells),
    timestamp: new Date().toISOString(),
  };

  const blob = new Blob([JSON.stringify(dataToSave, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${fileName}.json`;
  a.click();
  URL.revokeObjectURL(url);
};

export const importFromFile = (file, callback) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const importedData = JSON.parse(e.target.result);
      console.log("Imported data:", importedData);
      console.log("Type of importedData:", typeof importedData);
      if (Array.isArray(importedData)) {
        if (importedData.length > 0) {
          console.log(
            "First element type:",
            typeof importedData[0],
            importedData[0]
          );
        }
        if (
          importedData.length > 0 &&
          typeof importedData[0] === "object" &&
          !Array.isArray(importedData[0])
        ) {
          const headers = Object.keys(importedData[0]);
          const rows = importedData.map((obj) =>
            headers.map((h) => obj[h] ?? "")
          );
          callback([headers, ...rows]);
        } else {
          callback(importedData);
        }
      } else if (importedData && typeof importedData === "object") {
        // Single object: treat as one row
        const headers = Object.keys(importedData);
        const row = headers.map((h) => importedData[h] ?? "");
        callback([headers, row]);
      } else if (importedData.data && Array.isArray(importedData.data)) {
        callback(importedData.data);
      } else {
        throw new Error("Invalid file format");
      }
    } catch (error) {
      alert(
        "Error importing file. Please make sure it's a valid JSON file with the correct format."
      );
      console.error("Import error:", error);
    }
  };
  reader.readAsText(file);
};

export const addToHistory = (history, historyIndex, newData) => {
  const newHistory = history.slice(0, historyIndex + 1);
  newHistory.push(JSON.parse(JSON.stringify(newData)));
  return {
    history: newHistory,
    historyIndex: newHistory.length - 1,
  };
};

export const calculateColumnWidth = (containerWidth, numColumns) => {
  const availableWidth = containerWidth - 60;
  return Math.max(80, Math.floor(availableWidth / numColumns));
};

export const adjustFocusAfterRowDeletion = (focusedCell, dataLength) => {
  if (focusedCell.row >= dataLength) {
    return { ...focusedCell, row: dataLength - 1 };
  }
  return focusedCell;
};

export const adjustFocusAfterColumnDeletion = (focusedCell, numColumns) => {
  if (focusedCell.col >= numColumns) {
    return { ...focusedCell, col: numColumns - 1 };
  }
  return focusedCell;
};

export const autoFitColumnWidth = (data, colIndex) => {
  let maxWidth = 50;

  data.forEach((row) => {
    const cellValue = row.cells[colIndex] || "";
    const textWidth = cellValue.toString().length * 8;
    maxWidth = Math.max(maxWidth, textWidth);
  });

  return Math.min(maxWidth, 300);
};

export const getContextMenuItems = () => [
  { label: "Insert Row Above", action: "insertRowAbove", icon: "⬆️" },
  { label: "Insert Row Below", action: "insertRowBelow", icon: "⬇️" },
  { label: "---", action: null },
  { label: "Insert Column Left", action: "insertColumnLeft", icon: "⬅️" },
  { label: "Insert Column Right", action: "insertColumnRight", icon: "➡️" },
  { label: "---", action: null },
  { label: "Copy", action: "copy", icon: "📋", shortcut: "Ctrl+C" },
  { label: "Cut", action: "cut", icon: "✂️", shortcut: "Ctrl+X" },
  { label: "Paste", action: "paste", icon: "📋", shortcut: "Ctrl+V" },
  { label: "---", action: null },
  { label: "Delete Row", action: "deleteRow", icon: "🗑️" },
  { label: "Delete Column", action: "deleteColumn", icon: "🗑️" },
  { label: "---", action: null },
];

export const handleContextMenuAction = (action, data, focusedCell) => {
  console.log("=== CONTEXT MENU ACTION ===");
  console.log("Action:", action);
  console.log("Focused cell:", focusedCell);
  console.log("Data length:", data.length);
  console.log("===========================");

  if (
    !focusedCell ||
    typeof focusedCell.row === "undefined" ||
    typeof focusedCell.col === "undefined"
  ) {
    console.error("Invalid focused cell:", focusedCell);
    return data;
  }

  const { row, col } = focusedCell;

  switch (action) {
    case "clearCell":
      console.log("Executing clear cell action...");
      // Handle both row index and row ID
      if (typeof row === "number") {
        console.log("Clearing cell by index:", row, col);
        const result = clearCellDataByIndex(data, row, col);
        console.log("Clear cell result:", result);
        return result;
      } else {
        console.log("Clearing cell by ID:", row, col);
        const result = clearCellDataById(data, row, col);
        console.log("Clear cell result:", result);
        return result;
      }

    case "insertRowAbove":
      const rowIdx1 =
        typeof row === "number" ? row : findRowIndexById(data, row);
      if (rowIdx1 === -1) {
        console.error("Row not found for insertRowAbove");
        return data;
      }
      return insertRowAbove(data, rowIdx1, data[0]?.cells?.length || 0);

    case "insertRowBelow":
      const rowIdx2 =
        typeof row === "number" ? row : findRowIndexById(data, row);
      if (rowIdx2 === -1) {
        console.error("Row not found for insertRowBelow");
        return data;
      }
      return insertRowBelow(data, rowIdx2, data[0]?.cells?.length || 0);

    case "insertColumnLeft":
      return insertColumnLeft(data, col);

    case "insertColumnRight":
      return insertColumnRight(data, col);

    case "deleteRow":
      const rowIdx3 =
        typeof row === "number" ? row : findRowIndexById(data, row);
      if (rowIdx3 === -1) {
        console.error("Row not found for deleteRow");
        return data;
      }
      return deleteRow(data, rowIdx3);

    case "deleteColumn":
      return deleteColumn(data, col);

    default:
      console.log("Unknown action:", action);
      return data;
  }
};

export const debugClearCell = (data, focusedCell) => {
  console.log("=== DEBUG CLEAR CELL ===");
  console.log("Data length:", data.length);
  console.log("Focused cell:", focusedCell);
  if (
    focusedCell &&
    typeof focusedCell.row === "number" &&
    focusedCell.row >= 0
  ) {
    console.log(
      "Current cell value:",
      data[focusedCell.row]?.cells[focusedCell.col]
    );
    console.log("Row ID:", data[focusedCell.row]?.id);
  }
  console.log("========================");
};

export const testClearCellFunction = (data, focusedCell) => {
  console.log("=== TESTING CLEAR CELL ===");
  debugClearCell(data, focusedCell);

  if (
    !focusedCell ||
    typeof focusedCell.row === "undefined" ||
    typeof focusedCell.col === "undefined"
  ) {
    console.error("Invalid focused cell for testing");
    return data;
  }

  const result = clearCellData(data, focusedCell.row, focusedCell.col);
  console.log("Test result:", result);
  console.log("==========================");
  return result;
};

export const exportToCSV = (data, fileName) => {
  const rows = data.map((row) => row.cells);
  const csv = rows
    .map((row) =>
      row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(",")
    )
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${fileName}.csv`;
  a.click();
  URL.revokeObjectURL(url);
};

export const exportToXLS = (data, fileName) => {
  const rows = data.map((row) => row.cells);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  const wbout = XLSX.write(wb, { bookType: "xls", type: "array" });
  const blob = new Blob([wbout], { type: "application/vnd.ms-excel" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${fileName}.xls`;
  a.click();
  URL.revokeObjectURL(url);
};

export const exportToXLSX = (data, fileName) => {
  const rows = data.map((row) => row.cells);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${fileName}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
};

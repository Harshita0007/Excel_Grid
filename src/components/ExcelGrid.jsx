import React, { useState, useEffect, useRef, useCallback } from "react";
import {
  createColumnHeaders,
  createInitialGridData,
  filterData,
  sortData,
  getSortIcon,
  updateCellValue,
  insertRowAbove,
  insertRowBelow,
  insertColumnLeft,
  insertColumnRight,
  addNewRow,
  addNewColumn,
  deleteRow,
  deleteColumn,
  clearAllData,
  getNextFocusedCell,
  saveToFile,
  importFromFile,
  addToHistory,
  calculateColumnWidth,
  adjustFocusAfterRowDeletion,
  adjustFocusAfterColumnDeletion,
  getContextMenuItems,
  getFilterPresets,
  exportToCSV,
  exportToXLS,
  exportToXLSX,
} from "./ExcelUtils";

const ContextMenu = ({
  isOpen,
  position,
  onClose,
  onMenuAction,
  focusedCell,
}) => {
  const menuRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (menuRef.current && !menuRef.current.contains(event.target)) {
        onClose();
      }
    };

    if (isOpen) {
      document.addEventListener("mousedown", handleClickOutside);
    }

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  const menuItems = getContextMenuItems();

  return (
    <div
      ref={menuRef}
      style={{
        position: "fixed",
        top: position.y,
        left: position.x,
        backgroundColor: "white",
        border: "1px solid #ccc",
        borderRadius: "4px",
        boxShadow: "0 2px 10px rgba(0,0,0,0.1)",
        zIndex: 1000,
        minWidth: "200px",
        padding: "4px 0",
      }}
    >
      {menuItems.map((item, index) =>
        item.label === "---" ? (
          <div
            key={index}
            style={{ height: "1px", backgroundColor: "#eee", margin: "4px 0" }}
          />
        ) : (
          <div
            key={index}
            style={{
              padding: "8px 12px",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              fontSize: "14px",
              backgroundColor: "white",
              borderRadius: "2px",
            }}
            onMouseEnter={(e) => {
              e.target.style.backgroundColor = "#f0f0f0";
            }}
            onMouseLeave={(e) => {
              e.target.style.backgroundColor = "white";
            }}
            onClick={() => {
              if (item.action) {
                onMenuAction(item.action);
                onClose();
              }
            }}
          >
            <span>
              {item.icon && (
                <span style={{ marginRight: "8px" }}>{item.icon}</span>
              )}
              {item.label}
            </span>
            {item.shortcut && (
              <span style={{ fontSize: "12px", color: "#666" }}>
                {item.shortcut}
              </span>
            )}
          </div>
        )
      )}
    </div>
  );
};

const FileManagerModal = ({
  isOpen,
  onClose,
  onSave,
  onImport,
  onRefresh,
  data,
}) => {
  const [fileName, setFileName] = useState("spreadsheet");
  const fileInputRef = useRef(null);

  const handleDownload = (type) => {
    if (type === "json") {
      onSave(fileName);
    } else if (type === "csv") {
      exportToCSV(data, fileName);
    } else if (type === "xls") {
      exportToXLS(data, fileName);
    } else if (type === "xlsx") {
      exportToXLSX(data, fileName);
    }
    onClose();
  };

  const handleImport = (event) => {
    const file = event.target.files[0];
    if (file) {
      onImport(file);
      onClose();
    }
  };

  const handleRefresh = () => {
    onRefresh();
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: "rgba(0,0,0,0.5)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 1000,
      }}
    >
      <div
        style={{
          backgroundColor: "white",
          padding: "30px",
          borderRadius: "8px",
          minWidth: "350px",
          boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
        }}
      >
        <h3 style={{ marginTop: 0, marginBottom: "20px", color: "#333" }}>
          File Manager
        </h3>

        <div style={{ marginBottom: "20px" }}>
          <label
            style={{
              display: "block",
              marginBottom: "8px",
              fontWeight: "bold",
            }}
          >
            File Name:
          </label>
          <input
            type="text"
            value={fileName}
            onChange={(e) => setFileName(e.target.value)}
            style={{
              width: "100%",
              padding: "8px",
              border: "1px solid #ddd",
              borderRadius: "4px",
              fontSize: "14px",
            }}
          />
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
          <div
            style={{
              display: "flex",
              flexDirection: "row",
              gap: "10px",
              marginBottom: "10px",
              flexWrap: "wrap",
            }}
          >
            <button
              onClick={() => handleDownload("json")}
              style={{
                flex: 1,
                padding: "12px 16px",
                backgroundColor: "#28a745",
                color: "white",
                border: "none",
                borderRadius: "6px",
                cursor: "pointer",
                fontSize: "14px",
                fontWeight: "bold",
                minWidth: "120px",
              }}
            >
              💾 JSON
            </button>
            <button
              onClick={() => handleDownload("csv")}
              style={{
                flex: 1,
                padding: "12px 16px",
                backgroundColor: "#17a2b8",
                color: "white",
                border: "none",
                borderRadius: "6px",
                cursor: "pointer",
                fontSize: "14px",
                fontWeight: "bold",
                minWidth: "120px",
              }}
            >
              📄 CSV
            </button>
            <button
              onClick={() => handleDownload("xls")}
              style={{
                flex: 1,
                padding: "12px 16px",
                backgroundColor: "#ffc107",
                color: "black",
                border: "none",
                borderRadius: "6px",
                cursor: "pointer",
                fontSize: "14px",
                fontWeight: "bold",
                minWidth: "120px",
              }}
            >
              📊 XLS
            </button>
            <button
              onClick={() => handleDownload("xlsx")}
              style={{
                flex: 1,
                padding: "12px 16px",
                backgroundColor: "#6f42c1",
                color: "white",
                border: "none",
                borderRadius: "6px",
                cursor: "pointer",
                fontSize: "14px",
                fontWeight: "bold",
                minWidth: "120px",
              }}
            >
              📈 XLSX
            </button>
          </div>
          <button
            onClick={() => fileInputRef.current?.click()}
            style={{
              padding: "12px 20px",
              backgroundColor: "#007bff",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
              fontSize: "14px",
              fontWeight: "bold",
            }}
          >
            📂 Import JSON
          </button>
          <button
            onClick={handleRefresh}
            style={{
              padding: "12px 20px",
              backgroundColor: "#ffc107",
              color: "black",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
              fontSize: "14px",
              fontWeight: "bold",
            }}
          >
            🔄 Refresh/Reset
          </button>
          <button
            onClick={onClose}
            style={{
              padding: "12px 20px",
              backgroundColor: "#6c757d",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
              fontSize: "14px",
              fontWeight: "bold",
            }}
          >
            Cancel
          </button>
        </div>

        <input
          ref={fileInputRef}
          type="file"
          accept=".json"
          onChange={handleImport}
          style={{ display: "none" }}
        />
      </div>
    </div>
  );
};

// Helper to check if a cell is in the selected range
function cellIsInSelectedRange(row, col, range) {
  if (!range) return false;
  const minRow = Math.min(range.start.row, range.end.row);
  const maxRow = Math.max(range.start.row, range.end.row);
  const minCol = Math.min(range.start.col, range.end.col);
  const maxCol = Math.max(range.start.col, range.end.col);
  return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
}

// Helper to get selected rectangle bounds
function getSelectedRectangle(range) {
  if (!range) return null;
  const minRow = Math.min(range.start.row, range.end.row);
  const maxRow = Math.max(range.start.row, range.end.row);
  const minCol = Math.min(range.start.col, range.end.col);
  const maxCol = Math.max(range.start.col, range.end.col);
  return { minRow, maxRow, minCol, maxCol };
}

const ExcelGrid = ({
  initialData = [],
  rows: initialRows = 20,
  columns: initialColumns = 10,
  onDataChange = () => {},
  currentFileName = "default",
}) => {
  const [data, setData] = useState(() =>
    createInitialGridData(initialData, initialRows, initialColumns)
  );
  const [focusedCell, setFocusedCell] = useState({ row: 0, col: 0 });
  const [editingCell, setEditingCell] = useState(null);
  const [filters, setFilters] = useState({});
  const [sortConfig, setSortConfig] = useState({ key: null, direction: "asc" });
  const [colWidth, setColWidth] = useState(() =>
    Array(initialColumns).fill(120)
  );
  const [showFileManager, setShowFileManager] = useState(false);
  const [contextMenu, setContextMenu] = useState({
    isOpen: false,
    position: { x: 0, y: 0 },
  });
  const [clipboard, setClipboard] = useState({ data: null, type: null });
  const [history, setHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [selectedRange, setSelectedRange] = useState(null);
  const mouseDownRef = useRef(false);
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 500;
  const [showFilterRow, setShowFilterRow] = useState(false);

  const gridRef = useRef(null);
  const gridWrapperRef = useRef(null);
  const currentEditValue = useRef("");

  const currentRows = data.length;
  const currentColumns = data[0]?.cells?.length || 0;
  const columnHeaders = createColumnHeaders(currentColumns);

  const addToHistoryHandler = useCallback(
    (newData) => {
      const historyResult = addToHistory(history, historyIndex, newData);
      setHistory(historyResult.history);
      setHistoryIndex(historyResult.historyIndex);
    },
    [history, historyIndex]
  );

  const undo = useCallback(() => {
    if (historyIndex > 0) {
      const previousData = history[historyIndex - 1];
      setData(previousData);
      setHistoryIndex(historyIndex - 1);
      onDataChange(previousData);
    }
  }, [history, historyIndex, onDataChange]);

  const redo = useCallback(() => {
    if (historyIndex < history.length - 1) {
      const nextData = history[historyIndex + 1];
      setData(nextData);
      setHistoryIndex(historyIndex + 1);
      onDataChange(nextData);
    }
  }, [history, historyIndex, onDataChange]);

  useEffect(() => {
    if (history.length === 0) {
      setHistory([JSON.parse(JSON.stringify(data))]);
      setHistoryIndex(0);
    }
  }, []);

  useEffect(() => {
    const updateColumnWidth = () => {
      if (gridWrapperRef.current) {
        const containerWidth = gridWrapperRef.current.offsetWidth;
        const newColWidth = calculateColumnWidth(
          containerWidth,
          currentColumns
        );
        setColWidth((prev) => {
          if (prev.length === currentColumns) return prev;
          if (prev.length < currentColumns) {
            return [
              ...prev,
              ...Array(currentColumns - prev.length).fill(newColWidth),
            ];
          }
          return prev.slice(0, currentColumns);
        });
      }
    };

    const observer = new ResizeObserver(updateColumnWidth);
    if (gridWrapperRef.current) observer.observe(gridWrapperRef.current);
    updateColumnWidth();

    return () => observer.disconnect();
  }, [currentColumns]);

  const filteredData = filterData(data, filters);
  const sortedData = sortData(filteredData, sortConfig);

  // Pagination logic
  const totalRows = sortedData.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / rowsPerPage));
  const pagedRows = sortedData.slice(
    (currentPage - 1) * rowsPerPage,
    currentPage * rowsPerPage
  );

  const handleCellChange = (rowId, col, value) => {
    const newData = updateCellValue(data, rowId, col, value);
    setData(newData);
    onDataChange(newData);
    addToHistoryHandler(newData);
    setSortConfig((prev) => ({ key: null, direction: prev.direction }));
  };

  const handleContextMenuAction = (action) => {
    let newData;

    switch (action) {
      case "insertRowAbove":
        newData = insertRowAbove(data, focusedCell.row, currentColumns);
        break;
      case "insertRowBelow":
        newData = insertRowBelow(data, focusedCell.row, currentColumns);
        break;
      case "insertColumnLeft":
        newData = insertColumnLeft(data, focusedCell.col);
        break;
      case "insertColumnRight":
        newData = insertColumnRight(data, focusedCell.col);
        break;
      case "copy":
        handleCopy();
        return;
      case "cut":
        handleCut();
        return;
      case "paste":
        handlePaste();
        return;
      case "deleteRow":
        newData = deleteRow(data, focusedCell.row);
        const adjustedFocusRow = adjustFocusAfterRowDeletion(
          focusedCell,
          newData.length
        );
        setFocusedCell(adjustedFocusRow);
        break;
      case "deleteColumn":
        newData = deleteColumn(data, focusedCell.col);
        const adjustedFocusCol = adjustFocusAfterColumnDeletion(
          focusedCell,
          newData[0]?.cells?.length || 0
        );
        setFocusedCell(adjustedFocusCol);
        break;
      default:
        return;
    }

    if (newData) {
      setData(newData);
      onDataChange(newData);
      addToHistoryHandler(newData);
    }
  };

  const handleCopy = () => {
    let text = "";
    if (selectedRange) {
      const { minRow, maxRow, minCol, maxCol } =
        getSelectedRectangle(selectedRange);
      const rows = [];
      for (let r = minRow; r <= maxRow; r++) {
        const row = [];
        for (let c = minCol; c <= maxCol; c++) {
          row.push(sortedData[r]?.cells[c] ?? "");
        }
        rows.push(row.join("\t"));
      }
      text = rows.join("\n");
    } else {
      text = sortedData[focusedCell.row]?.cells[focusedCell.col] || "";
    }
    // Copy to clipboard
    if (navigator.clipboard) {
      navigator.clipboard.writeText(text);
    }
    setClipboard({ data: text, type: "copy" });
  };

  const handleCut = () => {
    let text = "";
    let newData = [...data];
    if (selectedRange) {
      const { minRow, maxRow, minCol, maxCol } =
        getSelectedRectangle(selectedRange);
      const rows = [];
      for (let r = minRow; r <= maxRow; r++) {
        const row = [];
        for (let c = minCol; c <= maxCol; c++) {
          row.push(sortedData[r]?.cells[c] ?? "");
          // Clear cell
          const gridRow = newData.findIndex(
            (rowObj) => rowObj.id === sortedData[r].id
          );
          if (gridRow !== -1) newData[gridRow].cells[c] = "";
        }
        rows.push(row.join("\t"));
      }
      text = rows.join("\n");
    } else {
      text = sortedData[focusedCell.row]?.cells[focusedCell.col] || "";
      const rowId = sortedData[focusedCell.row]?.id;
      if (rowId) {
        newData = updateCellValue(newData, rowId, focusedCell.col, "");
      }
    }
    if (navigator.clipboard) {
      navigator.clipboard.writeText(text);
    }
    setClipboard({ data: text, type: "cut" });
    setData(newData);
    onDataChange(newData);
    addToHistoryHandler(newData);
  };

  const handlePaste = async () => {
    let text = clipboard.data;
    if (!text && navigator.clipboard) {
      text = await navigator.clipboard.readText();
    }
    if (!text) return;
    const rows = text.split("\n").map((row) => row.split("\t"));
    let startRow = selectedRange
      ? Math.min(selectedRange.start.row, selectedRange.end.row)
      : focusedCell.row;
    let startCol = selectedRange
      ? Math.min(selectedRange.start.col, selectedRange.end.col)
      : focusedCell.col;
    let newData = [...data];
    for (let r = 0; r < rows.length; r++) {
      // Use the visible (sorted/filtered/paged) data to get the correct row id
      const gridRow = sortedData[startRow + r];
      if (!gridRow) continue;
      const rowId = gridRow.id;
      const mainDataRowIdx = newData.findIndex((rowObj) => rowObj.id === rowId);
      if (mainDataRowIdx === -1) continue;
      for (let c = 0; c < rows[r].length; c++) {
        if (startCol + c < newData[mainDataRowIdx].cells.length) {
          newData[mainDataRowIdx].cells[startCol + c] = rows[r][c];
        }
      }
    }
    setData(newData);
    onDataChange(newData);
    addToHistoryHandler(newData);
  };

  // Native clipboard event handlers
  const handleNativeCopy = (e) => {
    let text = "";
    if (selectedRange) {
      const { minRow, maxRow, minCol, maxCol } =
        getSelectedRectangle(selectedRange);
      const rows = [];
      for (let r = minRow; r <= maxRow; r++) {
        const row = [];
        for (let c = minCol; c <= maxCol; c++) {
          row.push(sortedData[r]?.cells[c] ?? "");
        }
        rows.push(row.join("\t"));
      }
      text = rows.join("\n");
    } else {
      text = sortedData[focusedCell.row]?.cells[focusedCell.col] || "";
    }
    e.clipboardData.setData("text/plain", text);
    e.preventDefault();
  };

  const handleNativeCut = (e) => {
    let text = "";
    let newData = [...data];
    if (selectedRange) {
      const { minRow, maxRow, minCol, maxCol } =
        getSelectedRectangle(selectedRange);
      const rows = [];
      for (let r = minRow; r <= maxRow; r++) {
        const row = [];
        for (let c = minCol; c <= maxCol; c++) {
          row.push(sortedData[r]?.cells[c] ?? "");
          // Clear cell
          const gridRow = newData.findIndex(
            (rowObj) => rowObj.id === sortedData[r].id
          );
          if (gridRow !== -1) newData[gridRow].cells[c] = "";
        }
        rows.push(row.join("\t"));
      }
      text = rows.join("\n");
    } else {
      text = sortedData[focusedCell.row]?.cells[focusedCell.col] || "";
      const rowId = sortedData[focusedCell.row]?.id;
      if (rowId) {
        newData = updateCellValue(newData, rowId, focusedCell.col, "");
      }
    }
    e.clipboardData.setData("text/plain", text);
    e.preventDefault();
    setData(newData);
    onDataChange(newData);
    addToHistoryHandler(newData);
  };

  const handleNativePaste = (e) => {
    let text = e.clipboardData.getData("text/plain");
    if (!text) return;
    const rows = text.split("\n").map((row) => row.split("\t"));
    let startRow = selectedRange
      ? Math.min(selectedRange.start.row, selectedRange.end.row)
      : focusedCell.row;
    let startCol = selectedRange
      ? Math.min(selectedRange.start.col, selectedRange.end.col)
      : focusedCell.col;
    let newData = [...data];
    for (let r = 0; r < rows.length; r++) {
      // Use the visible (sorted/filtered/paged) data to get the correct row id
      const gridRow = sortedData[startRow + r];
      if (!gridRow) continue;
      const rowId = gridRow.id;
      const mainDataRowIdx = newData.findIndex((rowObj) => rowObj.id === rowId);
      if (mainDataRowIdx === -1) continue;
      for (let c = 0; c < rows[r].length; c++) {
        if (startCol + c < newData[mainDataRowIdx].cells.length) {
          newData[mainDataRowIdx].cells[startCol + c] = rows[r][c];
        }
      }
    }
    setData(newData);
    onDataChange(newData);
    addToHistoryHandler(newData);
    e.preventDefault();
  };

  useEffect(() => {
    const handleGlobalKeyDown = (e) => {
      if (showFileManager || contextMenu.isOpen) return;

      const key = e.key;
      const ctrlKey = e.ctrlKey || e.metaKey;

      if (ctrlKey && key === "z" && !e.shiftKey) {
        e.preventDefault();
        undo();
        return;
      }

      if (ctrlKey && (key === "y" || (key === "z" && e.shiftKey))) {
        e.preventDefault();
        redo();
        return;
      }

      if (ctrlKey && key === "c") {
        e.preventDefault();
        handleCopy();
        return;
      }

      if (ctrlKey && key === "x") {
        e.preventDefault();
        handleCut();
        return;
      }

      if (ctrlKey && key === "v") {
        e.preventDefault();
        handlePaste();
        return;
      }

      if (key === "Delete" || key === "Backspace") {
        const tag = document.activeElement?.tagName;
        if (tag === "INPUT" || tag === "TEXTAREA") return;

        e.preventDefault();
        if (editingCell) return;
        // Check if there is a multi-cell selection
        if (
          selectedRange &&
          (selectedRange.start.row !== selectedRange.end.row ||
            selectedRange.start.col !== selectedRange.end.col)
        ) {
          // Clear all cells in the selected range in a single batch update
          const startRow = Math.min(
            selectedRange.start.row,
            selectedRange.end.row
          );
          const endRow = Math.max(
            selectedRange.start.row,
            selectedRange.end.row
          );
          const startCol = Math.min(
            selectedRange.start.col,
            selectedRange.end.col
          );
          const endCol = Math.max(
            selectedRange.start.col,
            selectedRange.end.col
          );

          // Deep copy of data
          let newData = data.map((row) => ({ ...row, cells: [...row.cells] }));
          for (let row = startRow; row <= endRow; row++) {
            const rowId = sortedData[row]?.id;
            const gridRow = newData.findIndex((rowObj) => rowObj.id === rowId);
            if (gridRow !== -1) {
              for (let col = startCol; col <= endCol; col++) {
                newData[gridRow].cells[col] = "";
              }
            }
          }
          setData(newData);
          onDataChange(newData);
          addToHistoryHandler(newData);
        } else {
          // Single cell deletion (existing behavior)
          const rowId = sortedData[focusedCell.row]?.id;
          if (rowId) {
            handleCellChange(rowId, focusedCell.col, "");
          }
        }
        return;
      }

      let newFocusedCell;
      switch (key) {
        case "Tab":
          e.preventDefault();
          newFocusedCell = getNextFocusedCell(
            focusedCell,
            e.shiftKey ? "shift-tab" : "tab",
            sortedData.length,
            currentColumns
          );
          setFocusedCell(newFocusedCell);
          setEditingCell(null);
          return;
        case "ArrowUp":
          e.preventDefault();
          newFocusedCell = getNextFocusedCell(
            focusedCell,
            "up",
            sortedData.length,
            currentColumns
          );
          setFocusedCell(newFocusedCell);
          setEditingCell(null);
          return;
        case "ArrowDown":
          e.preventDefault();
          newFocusedCell = getNextFocusedCell(
            focusedCell,
            "down",
            sortedData.length,
            currentColumns
          );
          setFocusedCell(newFocusedCell);
          setEditingCell(null);
          return;
        case "ArrowLeft":
          e.preventDefault();
          newFocusedCell = getNextFocusedCell(
            focusedCell,
            "left",
            sortedData.length,
            currentColumns
          );
          setFocusedCell(newFocusedCell);
          setEditingCell(null);
          return;
        case "ArrowRight":
          e.preventDefault();
          newFocusedCell = getNextFocusedCell(
            focusedCell,
            "right",
            sortedData.length,
            currentColumns
          );
          setFocusedCell(newFocusedCell);
          setEditingCell(null);
          return;
        case "Enter":
          e.preventDefault();
          if (editingCell) {
            setEditingCell(null);
            newFocusedCell = getNextFocusedCell(
              focusedCell,
              "enter",
              sortedData.length,
              currentColumns
            );
            setFocusedCell(newFocusedCell);
          } else {
            setEditingCell({ rowIndex: focusedCell.row, col: focusedCell.col });
            const currentValue =
              sortedData[focusedCell.row]?.cells[focusedCell.col] || "";
            currentEditValue.current = currentValue;
          }
          return;
        case "Escape":
          e.preventDefault();
          if (editingCell) {
            const rowId = sortedData[focusedCell.row]?.id;
            if (rowId) {
              handleCellChange(
                rowId,
                focusedCell.col,
                currentEditValue.current
              );
            }
            setEditingCell(null);
          }
          return;
        case "F2":
          e.preventDefault();
          setEditingCell({ rowIndex: focusedCell.row, col: focusedCell.col });
          const currentValue =
            sortedData[focusedCell.row]?.cells[focusedCell.col] || "";
          currentEditValue.current = currentValue;
          return;
        default:
          if (key.length === 1 && !ctrlKey && !editingCell) {
            e.preventDefault();
            setEditingCell({ rowIndex: focusedCell.row, col: focusedCell.col });
            const rowId = sortedData[focusedCell.row]?.id;
            if (rowId) {
              handleCellChange(rowId, focusedCell.col, key);
            }
          }
          return;
      }
    };

    document.addEventListener("keydown", handleGlobalKeyDown);
    return () => document.removeEventListener("keydown", handleGlobalKeyDown);
  }, [
    focusedCell,
    editingCell,
    showFileManager,
    contextMenu.isOpen,
    undo,
    redo,
    sortedData,
    currentColumns,
    handleCopy,
    handleCut,
    handlePaste,
    handleCellChange,
  ]);

  // Mouse selection handlers
  const handleCellMouseDown = (rowIndex, colIndex, e) => {
    if (e.shiftKey && focusedCell) {
      setSelectedRange({
        start: focusedCell,
        end: { row: rowIndex, col: colIndex },
      });
    } else {
      setFocusedCell({ row: rowIndex, col: colIndex });
      setSelectedRange({
        start: { row: rowIndex, col: colIndex },
        end: { row: rowIndex, col: colIndex },
      });
    }
    mouseDownRef.current = true;
  };

  const handleCellMouseEnter = (rowIndex, colIndex) => {
    if (mouseDownRef.current && selectedRange) {
      setSelectedRange((range) => ({
        ...range,
        end: { row: rowIndex, col: colIndex },
      }));
    }
  };

  useEffect(() => {
    const handleMouseUp = () => {
      mouseDownRef.current = false;
    };
    document.addEventListener("mouseup", handleMouseUp);
    return () => document.removeEventListener("mouseup", handleMouseUp);
  }, []);

  // Keyboard selection handlers
  useEffect(() => {
    const handleKeyDown = (e) => {
      if (
        e.shiftKey &&
        ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key)
      ) {
        e.preventDefault();
        let { row, col } = focusedCell;
        let newEnd = selectedRange?.end || { row, col };
        switch (e.key) {
          case "ArrowUp":
            newEnd = { row: Math.max(0, newEnd.row - 1), col: newEnd.col };
            break;
          case "ArrowDown":
            newEnd = {
              row: Math.min(data.length - 1, newEnd.row + 1),
              col: newEnd.col,
            };
            break;
          case "ArrowLeft":
            newEnd = { row: newEnd.row, col: Math.max(0, newEnd.col - 1) };
            break;
          case "ArrowRight":
            newEnd = {
              row: newEnd.row,
              col: Math.min(data[0].cells.length - 1, newEnd.col + 1),
            };
            break;
        }
        setSelectedRange((range) => ({
          start: range?.start || focusedCell,
          end: newEnd,
        }));
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [focusedCell, selectedRange, data]);

  const handleCellClick = (rowIndex, colIndex, e) => {
    setFocusedCell({ row: rowIndex, col: colIndex });
    setEditingCell(null);
    if (!e.shiftKey) setSelectedRange(null);
  };

  const handleCellDoubleClick = (rowIndex, colIndex) => {
    setEditingCell({ rowIndex, col: colIndex });
    const currentValue = sortedData[rowIndex]?.cells[colIndex] || "";
    currentEditValue.current = currentValue;
  };

  const handleSort = (colIndex) => {
    let direction = "asc";
    if (sortConfig.key === colIndex && sortConfig.direction === "asc") {
      direction = "desc";
    }
    setSortConfig({ key: colIndex, direction });
  };
  const handleFilterChange = (colIndex, value) => {
    setFilters((prev) => ({
      ...prev,
      [colIndex]: {
        ...(prev[colIndex] || {}),
        value,
        type: prev[colIndex]?.type || "contains",
      },
    }));
  };

  const handleContextMenu = (e, rowIndex, colIndex) => {
    e.preventDefault();
    setFocusedCell({ row: rowIndex, col: colIndex });
    setContextMenu({
      isOpen: true,
      position: { x: e.clientX, y: e.clientY },
    });
  };

  const handleAddRow = () => {
    const newData = addNewRow(data, currentColumns);
    setData(newData);
    onDataChange(newData);
    addToHistoryHandler(newData);
  };

  const handleAddColumn = () => {
    const newData = addNewColumn(data);
    setData(newData);
    onDataChange(newData);
    addToHistoryHandler(newData);
  };

  const handleDeleteRow = () => {
    if (data.length <= 1) return;
    const newData = deleteRow(data, focusedCell.row);
    const adjustedFocus = adjustFocusAfterRowDeletion(
      focusedCell,
      newData.length
    );
    setData(newData);
    setFocusedCell(adjustedFocus);
    onDataChange(newData);
    addToHistoryHandler(newData);
  };

  const handleDeleteColumn = () => {
    if (currentColumns <= 1) return;
    const newData = deleteColumn(data, focusedCell.col);
    const adjustedFocus = adjustFocusAfterColumnDeletion(
      focusedCell,
      newData[0]?.cells?.length || 0
    );
    setData(newData);
    setFocusedCell(adjustedFocus);
    onDataChange(newData);
    addToHistoryHandler(newData);
  };

  const handleClearAll = () => {
    if (
      window.confirm(
        "Are you sure you want to clear all data? This action cannot be undone."
      )
    ) {
      const newData = clearAllData(data);
      setData(newData);
      onDataChange(newData);
      addToHistoryHandler(newData);
      setFilters({});
      setSortConfig({ key: null, direction: "asc" });
    }
  };

  const handleSave = (fileName) => {
    saveToFile(data, fileName);
  };

  const handleImport = (file) => {
    importFromFile(file, (newData) => {
      console.log("Callback called", newData);
      const gridData = createInitialGridData(
        newData,
        initialRows,
        initialColumns
      );
      setData(gridData);
      onDataChange(gridData);
      addToHistoryHandler(gridData);
      setFilters({});
      setSortConfig({ key: null, direction: "asc" });
    });
  };

  const handleRefresh = () => {
    if (
      window.confirm(
        "Are you sure you want to refresh? This will reset all data to the initial state."
      )
    ) {
      const newData = createInitialGridData(
        initialData,
        initialRows,
        initialColumns
      );
      setData(newData);
      onDataChange(newData);
      addToHistoryHandler(newData);
      setFilters({});
      setSortConfig({ key: null, direction: "asc" });
      setFocusedCell({ row: 0, col: 0 });
    }
  };

  // Modern button style for toolbar
  const toolbarButtonStyle = {
    background: "linear-gradient(90deg, #f3f4f6 0%, #e5e7eb 100%)",
    border: "1px solid #bfc9d9",
    borderRadius: "8px",
    boxShadow: "0 1px 4px rgba(0,0,0,0.04)",
    color: "#222",
    fontWeight: 600,
    fontSize: "1.1rem",
    padding: "10px 22px",
    margin: "0 8px 8px 0",
    cursor: "pointer",
    transition: "background 0.2s, box-shadow 0.2s",
    outline: "none",
    display: "inline-flex",
    alignItems: "center",
    gap: "6px",
  };

  // Column resize state
  const [resizingCol, setResizingCol] = useState(null);
  const [resizeStartX, setResizeStartX] = useState(0);
  const [resizeStartWidth, setResizeStartWidth] = useState(0);

  const handleResizeMouseDown = (colIndex, e) => {
    e.preventDefault();
    setResizingCol(colIndex);
    setResizeStartX(e.clientX);
    setResizeStartWidth(colWidth[colIndex]);
    document.body.style.cursor = "col-resize";
  };
  useEffect(() => {
    if (resizingCol === null) return;
    const handleMouseMove = (e) => {
      const delta = e.clientX - resizeStartX;
      setColWidth((prev) => {
        const next = [...prev];
        next[resizingCol] = Math.max(60, resizeStartWidth + delta);
        return next;
      });
    };
    const handleMouseUp = () => {
      setResizingCol(null);
      document.body.style.cursor = "";
    };
    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };
  }, [resizingCol, resizeStartX, resizeStartWidth]);

  // Count only rows with at least one non-empty cell
  const nonEmptyRowCount = data.filter((row) =>
    row.cells.some((cell) => String(cell).trim() !== "")
  ).length;

  // Count only columns with at least one non-empty cell
  const nonEmptyColCount = (data[0]?.cells || []).filter((_, colIdx) =>
    data.some((row) => String(row.cells[colIdx] ?? "").trim() !== "")
  ).length;

  const [darkMode, setDarkMode] = useState(() => {
    if (typeof window !== "undefined") {
      return localStorage.getItem("excelGridDarkMode") === "true";
    }
    return false;
  });
  useEffect(() => {
    if (typeof window !== "undefined") {
      localStorage.setItem("excelGridDarkMode", darkMode);
    }
  }, [darkMode]);

  return (
    <div
      style={{
        minHeight: "100vh",
        minWidth: "100vw",
        background: darkMode ? "#1e1e2e" : "#f8fafc",
        margin: 0,
        padding: 0,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Segoe UI', Arial, sans-serif",
        overflow: "hidden",
      }}
      tabIndex={0}
      onCopy={handleNativeCopy}
      onCut={handleNativeCut}
      onPaste={handleNativePaste}
    >
      <div
        ref={gridWrapperRef}
        style={{
          position: "relative",
          width: "95vw",
          maxWidth: "1600px",
          minWidth: "900px",
          height: "90vh",
          maxHeight: "900px",
          minHeight: "600px",
          background: darkMode ? "#232336" : "#fff",
          borderRadius: "16px",
          boxShadow: darkMode
            ? "0 2px 24px rgba(30,30,50,0.25)"
            : "0 2px 24px rgba(0,0,0,0.10)",
          display: "flex",
          flexDirection: "column",
          alignItems: "stretch",
          justifyContent: "flex-start",
          overflow: "hidden",
          padding: "24px",
          border: darkMode ? "1px solid #444" : "1px solid #ccc",
        }}
      >
        <button
          onClick={() => setShowFileManager(true)}
          style={{
            position: "absolute",
            top: 24,
            right: 32,
            zIndex: 10,
            padding: "8px 16px",
            background: "linear-gradient(90deg, #6f42c1 0%, #8e54e9 100%)",
            color: "#fff",
            border: "none",
            borderRadius: "18px",
            cursor: "pointer",
            fontSize: "14px",
            fontWeight: "bold",
            boxShadow: "0 4px 16px rgba(111,66,193,0.15)",
            letterSpacing: "0.5px",
            transition: "background 0.2s, box-shadow 0.2s",
          }}
        >
          📁 File Manager
        </button>
        <button
          onClick={() => setDarkMode((d) => !d)}
          style={{
            position: "absolute",
            top: 64,
            right: 32,
            zIndex: 10,
            padding: "8px 16px",
            background: darkMode
              ? "linear-gradient(90deg, #22263a 0%, #3a3a5a 100%)"
              : "linear-gradient(90deg, #6f42c1 0%, #8e54e9 100%)",
            color: "#fff",
            border: "none",
            borderRadius: "18px",
            cursor: "pointer",
            fontSize: "14px",
            fontWeight: "bold",
            boxShadow: "0 4px 16px rgba(111,66,193,0.15)",
            letterSpacing: "0.5px",
            transition: "background 0.2s, box-shadow 0.2s",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: "8px",
          }}
        >
          {darkMode ? "🌙 Dark Mode" : "☀️ Light Mode"}
        </button>
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            margin: 0,
            padding: "40px 40px 0 40px",
          }}
        >
          <h1
            style={{
              fontWeight: 700,
              fontSize: "2.5rem",
              margin: 0,
              display: "flex",
              alignItems: "center",
              gap: 12,
            }}
          >
            <span role="img" aria-label="spreadsheet">
              🧮
            </span>{" "}
            Excel-like Grid Component
          </h1>
          <div
            style={{
              margin: "24px 0 16px 0",
              display: "flex",
              flexWrap: "wrap",
              gap: "0",
            }}
          >
            <button style={toolbarButtonStyle} onClick={handleAddRow}>
              ➕ Add Row
            </button>
            <button style={toolbarButtonStyle} onClick={handleAddColumn}>
              ➕ Add Column
            </button>
            <button style={toolbarButtonStyle} onClick={handleDeleteRow}>
              🗑️ Delete Row
            </button>
            <button style={toolbarButtonStyle} onClick={handleDeleteColumn}>
              🗑️ Delete Column
            </button>
            <button style={toolbarButtonStyle} onClick={handleClearAll}>
              🧹 Clear All
            </button>
            <button
              style={toolbarButtonStyle}
              onClick={() => setShowFilterRow(!showFilterRow)}
            >
              {showFilterRow ? "🔽 Hide Filters" : "🔼 Show Filters"}
            </button>
            <button
              style={{
                ...toolbarButtonStyle,
                opacity: historyIndex <= 0 ? 0.5 : 1,
                cursor: historyIndex <= 0 ? "not-allowed" : "pointer",
              }}
              onClick={undo}
              disabled={historyIndex <= 0}
            >
              ↩️ Undo
            </button>
            <button
              style={{
                ...toolbarButtonStyle,
                opacity: historyIndex >= history.length - 1 ? 0.5 : 1,
                cursor:
                  historyIndex >= history.length - 1
                    ? "not-allowed"
                    : "pointer",
              }}
              onClick={redo}
              disabled={historyIndex >= history.length - 1}
            >
              ↪️ Redo
            </button>
          </div>
        </div>
        <div
          ref={gridWrapperRef}
          style={{
            flex: 1,
            height: "100%",
            overflowY: "auto",
            overflowX: "auto",
            border: "1px solid #ccc",
            background: darkMode ? "#232336" : "#fff",
            borderRadius: "10px",
            boxShadow: "0 2px 12px rgba(0,0,0,0.07)",
            position: "relative",
            width: "100%",
            margin: 0,
            display: "flex",
            flexDirection: "column",
          }}
        >
          <table
            ref={gridRef}
            style={{
              borderCollapse: "collapse",
              width: "100%",
              height: "100%",
              tableLayout: "fixed",
              background: darkMode ? "#232336" : "#fff",
              color: darkMode ? "#e5e7eb" : "#222",
              transition: "background 0.2s, color 0.2s",
            }}
          >
            <thead>
              <tr>
                <th
                  style={{
                    width: "60px",
                    backgroundColor: darkMode ? "#232336" : "#f1f1f1",
                    color: darkMode ? "#e5e7eb" : "#222",
                    position: "sticky",
                    top: 0,
                    left: 0,
                    zIndex: 4,
                    borderBottom: darkMode
                      ? "2px solid #333"
                      : "2px solid #e0e0e0",
                    borderRight: darkMode
                      ? "2px solid #333"
                      : "2px solid #e0e0e0",
                    boxShadow: darkMode
                      ? "2px 0 6px -2px rgba(30,30,50,0.10)"
                      : "2px 0 6px -2px rgba(0,0,0,0.04)",
                  }}
                ></th>
                {columnHeaders.map((header, colIndex) => (
                  <th
                    key={colIndex}
                    onClick={() => handleSort(colIndex)}
                    style={{
                      width: `${colWidth[colIndex]}px`,
                      backgroundColor: darkMode ? "#232336" : "#f1f1f1",
                      color: darkMode ? "#e5e7eb" : "#222",
                      border: darkMode ? "1px solid #333" : "1px solid #ddd",
                      padding: "10px 6px",
                      textAlign: "center",
                      cursor: "pointer",
                      position: "sticky",
                      top: 0,
                      zIndex: 10,
                      fontWeight: 600,
                      fontSize: "1.1rem",
                      borderBottom: darkMode
                        ? "2px solid #333"
                        : "2px solid #e0e0e0",
                      userSelect: resizingCol === colIndex ? "none" : undefined,
                    }}
                  >
                    {header} {getSortIcon(colIndex, sortConfig)}
                    {/* Resize handle */}
                    <span
                      onMouseDown={(e) => handleResizeMouseDown(colIndex, e)}
                      style={{
                        position: "absolute",
                        right: 0,
                        top: 0,
                        width: "8px",
                        height: "100%",
                        cursor: "col-resize",
                        zIndex: 10,
                        userSelect: "none",
                        background:
                          resizingCol === colIndex
                            ? darkMode
                              ? "rgba(33,134,235,0.18)"
                              : "rgba(33,134,235,0.12)"
                            : "transparent",
                        transition: "background 0.2s",
                      }}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </th>
                ))}
              </tr>
              {showFilterRow && (
                <tr>
                  <th
                    style={{
                      width: "60px",
                      backgroundColor: darkMode ? "#232336" : "#f3f6fa",
                      color: darkMode ? "#e5e7eb" : "#222",
                      position: "sticky",
                      top: 48,
                      left: 0,
                      zIndex: 4,
                      boxShadow: darkMode
                        ? "2px 0 6px -2px rgba(30,30,50,0.10)"
                        : "2px 0 6px -2px rgba(0,0,0,0.04)",
                      borderBottom: darkMode
                        ? "2px solid #333"
                        : "2px solid #e0e0e0",
                      borderRight: darkMode
                        ? "2px solid #333"
                        : "2px solid #e0e0e0",
                    }}
                  >
                    <button
                      onClick={() => setShowFilterRow(false)}
                      style={{
                        background: "none",
                        border: "none",
                        cursor: "pointer",
                        fontSize: 18,
                        color: darkMode ? "#aaa" : "#888",
                      }}
                      title="Collapse filter row"
                    >
                      ✖️
                    </button>
                  </th>
                  {columnHeaders.map((_, colIndex) => (
                    <th
                      key={`filter-${colIndex}`}
                      style={{
                        padding: "8px 4px",
                        backgroundColor: darkMode ? "#232336" : "#f3f6fa",
                        color: darkMode ? "#e5e7eb" : "#222",
                        border: darkMode
                          ? "1px solid #333"
                          : "1px solid #e3e8ee",
                        position: "sticky",
                        top: 48,
                        zIndex: 5,
                        height: "60px",
                        boxShadow: darkMode
                          ? "0 2px 6px -2px rgba(30,30,50,0.13)"
                          : "0 2px 6px -2px rgba(0,0,0,0.07)",
                        borderBottom: darkMode
                          ? "2px solid #333"
                          : "2px solid #e0e0e0",
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          flexDirection: "column",
                          gap: "6px",
                        }}
                      >
                        <select
                          value={filters[colIndex]?.type || "contains"}
                          onChange={(e) => {
                            const type = e.target.value;
                            setFilters((prev) => ({
                              ...prev,
                              [colIndex]: {
                                ...(prev[colIndex] || {}),
                                type,
                                value:
                                  type === "notEmpty" || type === "empty"
                                    ? "1"
                                    : "",
                              },
                            }));
                          }}
                          onClick={(e) => e.stopPropagation()}
                          onMouseDown={(e) => e.stopPropagation()}
                          onFocus={(e) => e.stopPropagation()}
                          onKeyDown={(e) => e.stopPropagation()}
                          style={{
                            width: "100%",
                            fontSize: "13px",
                            padding: "4px 8px",
                            border: darkMode
                              ? "1px solid #444"
                              : "1px solid #bfc9d9",
                            borderRadius: "8px",
                            background: darkMode ? "#232336" : "#fff",
                            color: darkMode ? "#e5e7eb" : "#222",
                            boxShadow: darkMode
                              ? "0 1px 2px rgba(30,30,50,0.10)"
                              : "0 1px 2px rgba(0,0,0,0.03)",
                            outline: "none",
                            transition: "border 0.2s",
                          }}
                        >
                          {getFilterPresets().map((preset) => (
                            <option key={preset.value} value={preset.value}>
                              {preset.icon} {preset.label}
                            </option>
                          ))}
                        </select>
                        <input
                          type="text"
                          value={filters[colIndex]?.value || ""}
                          onChange={(e) =>
                            handleFilterChange(colIndex, e.target.value)
                          }
                          onClick={(e) => e.stopPropagation()}
                          onMouseDown={(e) => e.stopPropagation()}
                          onFocus={(e) => e.stopPropagation()}
                          onKeyDown={(e) => e.stopPropagation()}
                          style={{
                            width: "100%",
                            padding: "6px 10px",
                            fontSize: "13px",
                            border: darkMode
                              ? "1px solid #444"
                              : "1px solid #bfc9d9",
                            borderRadius: "8px",
                            boxSizing: "border-box",
                            background: darkMode ? "#232336" : "#fff",
                            color: darkMode ? "#e5e7eb" : "#222",
                            boxShadow: darkMode
                              ? "0 1px 2px rgba(30,30,50,0.10)"
                              : "0 1px 2px rgba(0,0,0,0.03)",
                            outline: "none",
                            transition: "border 0.2s",
                            display:
                              filters[colIndex]?.type === "notEmpty" ||
                              filters[colIndex]?.type === "empty"
                                ? "none"
                                : "block",
                          }}
                        />
                      </div>
                    </th>
                  ))}
                </tr>
              )}
            </thead>
            <tbody>
              {pagedRows.map((row, rowIndex) => {
                const globalRowIndex =
                  (currentPage - 1) * rowsPerPage + rowIndex;
                return (
                  <tr
                    key={row.id}
                    className="excel-grid-row"
                    style={{
                      backgroundColor: darkMode
                        ? globalRowIndex % 2 === 0
                          ? "#232336"
                          : "#282846"
                        : globalRowIndex % 2 === 0
                        ? "#f9f9fb"
                        : "#fff",
                      color: darkMode ? "#e5e7eb" : "#222",
                      transition: "background 0.2s, color 0.2s",
                    }}
                  >
                    <th
                      style={{
                        backgroundColor: darkMode
                          ? "#232336"
                          : focusedCell.row === globalRowIndex
                          ? "#f8f9fa"
                          : "#f1f1f1",
                        color: darkMode ? "#e5e7eb" : "#222",
                        border: darkMode ? "1px solid #333" : "1px solid #ddd",
                        textAlign: "center",
                        position: "sticky",
                        left: 0,
                        zIndex: 2,
                        width: "60px",
                        fontWeight: 700,
                        borderRight: darkMode
                          ? "2px solid #333"
                          : "2px solid #e0e0e0",
                        boxShadow: darkMode
                          ? "2px 0 6px -2px rgba(30,30,50,0.10)"
                          : "2px 0 6px -2px rgba(0,0,0,0.04)",
                      }}
                    >
                      {String(globalRowIndex + 1).padStart(2, "0")}
                    </th>
                    {row.cells.map((cell, colIndex) => {
                      const isFocused =
                        focusedCell.row === globalRowIndex &&
                        focusedCell.col === colIndex;
                      const isEditing =
                        editingCell?.rowIndex === globalRowIndex &&
                        editingCell?.col === colIndex;
                      const isSelected = cellIsInSelectedRange(
                        globalRowIndex,
                        colIndex,
                        selectedRange
                      );
                      return (
                        <td
                          key={colIndex}
                          onClick={(e) =>
                            handleCellClick(globalRowIndex, colIndex, e)
                          }
                          onMouseDown={(e) =>
                            handleCellMouseDown(globalRowIndex, colIndex, e)
                          }
                          onMouseEnter={() =>
                            handleCellMouseEnter(globalRowIndex, colIndex)
                          }
                          onDoubleClick={() =>
                            handleCellDoubleClick(globalRowIndex, colIndex)
                          }
                          onContextMenu={(e) =>
                            handleContextMenu(e, globalRowIndex, colIndex)
                          }
                          style={{
                            width: `${colWidth[colIndex]}px`,
                            padding: "10px 8px",
                            cursor: "pointer",
                            backgroundColor: isFocused
                              ? darkMode
                                ? "#2d3a4a"
                                : "#eaf3fb"
                              : isSelected
                              ? darkMode
                                ? "#1e293b"
                                : "#dbeafe"
                              : "inherit",
                            color: darkMode ? "#e5e7eb" : "#222",
                            borderTop:
                              isSelected &&
                              !cellIsInSelectedRange(
                                globalRowIndex - 1,
                                colIndex,
                                selectedRange
                              )
                                ? darkMode
                                  ? "2px solid #38bdf8"
                                  : "2px solid #2186eb"
                                : darkMode
                                ? "1px solid #333"
                                : "1px solid #ccc",
                            borderBottom:
                              isSelected &&
                              !cellIsInSelectedRange(
                                globalRowIndex + 1,
                                colIndex,
                                selectedRange
                              )
                                ? darkMode
                                  ? "2px solid #38bdf8"
                                  : "2px solid #2186eb"
                                : darkMode
                                ? "1px solid #333"
                                : "1px solid #ccc",
                            borderLeft:
                              isSelected &&
                              !cellIsInSelectedRange(
                                globalRowIndex,
                                colIndex - 1,
                                selectedRange
                              )
                                ? darkMode
                                  ? "2px solid #38bdf8"
                                  : "2px solid #2186eb"
                                : darkMode
                                ? "1px solid #333"
                                : "1px solid #ccc",
                            borderRight:
                              isSelected &&
                              !cellIsInSelectedRange(
                                globalRowIndex,
                                colIndex + 1,
                                selectedRange
                              )
                                ? darkMode
                                  ? "2px solid #38bdf8"
                                  : "2px solid #2186eb"
                                : darkMode
                                ? "1px solid #333"
                                : "1px solid #ccc",
                            boxShadow: isFocused
                              ? darkMode
                                ? "0 0 0 2px #38bdf8 inset"
                                : "0 0 0 2px #2186eb inset"
                              : undefined,
                            fontSize: "1rem",
                            transition:
                              "background 0.2s, outline 0.2s, border 0.2s, color 0.2s",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                            maxWidth: `${colWidth[colIndex]}px`,
                            userSelect: isSelected ? "none" : undefined,
                          }}
                        >
                          {isEditing ? (
                            <input
                              autoFocus
                              value={cell}
                              onChange={(e) =>
                                handleCellChange(
                                  row.id,
                                  colIndex,
                                  e.target.value
                                )
                              }
                              onBlur={() => setEditingCell(null)}
                              onKeyDown={(e) => {
                                if (e.key === "Enter") setEditingCell(null);
                              }}
                              style={{
                                width: "100%",
                                border: "none",
                                background: darkMode ? "#232336" : "#fff",
                                color: darkMode ? "#e5e7eb" : "#222",
                                outline: "none",
                                fontSize: "1rem",
                              }}
                            />
                          ) : (
                            <span
                              style={{
                                display: "block",
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                                width: "100%",
                                maxWidth: "100%",
                                userSelect: isSelected ? "none" : undefined,
                                color: darkMode ? "#e5e7eb" : "#222",
                              }}
                            >
                              {cell}
                            </span>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr>
                <td
                  colSpan={columnHeaders.length + 1}
                  style={{
                    padding: "16px 0",
                    background: "#fff",
                    textAlign: "center",
                    border: "none",
                  }}
                >
                  <div
                    style={{
                      display: "inline-flex",
                      justifyContent: "center",
                      alignItems: "center",
                      gap: 16,
                    }}
                  >
                    <button
                      onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                      disabled={currentPage === 1}
                      style={{
                        padding: "8px 16px",
                        borderRadius: 6,
                        border: "1px solid #ccc",
                        background: currentPage === 1 ? "#eee" : "#fff",
                        cursor: currentPage === 1 ? "not-allowed" : "pointer",
                      }}
                    >
                      Prev
                    </button>
                    <span style={{ fontWeight: 500, fontSize: "1.1em" }}>
                      Page {currentPage} of {totalPages}
                    </span>
                    <button
                      onClick={() =>
                        setCurrentPage((p) => Math.min(totalPages, p + 1))
                      }
                      disabled={currentPage === totalPages}
                      style={{
                        padding: "8px 16px",
                        borderRadius: 6,
                        border: "1px solid #ccc",
                        background:
                          currentPage === totalPages ? "#eee" : "#fff",
                        cursor:
                          currentPage === totalPages
                            ? "not-allowed"
                            : "pointer",
                      }}
                    >
                      Next
                    </button>
                  </div>
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
        <ContextMenu
          isOpen={contextMenu.isOpen}
          position={contextMenu.position}
          onClose={() => setContextMenu({ ...contextMenu, isOpen: false })}
          onMenuAction={handleContextMenuAction}
          focusedCell={focusedCell}
        />
        <FileManagerModal
          isOpen={showFileManager}
          onClose={() => setShowFileManager(false)}
          onSave={handleSave}
          onImport={handleImport}
          onRefresh={handleRefresh}
          data={data}
        />
        {/* Info bar below grid */}
        <div
          style={{
            width: "100%",
            marginTop: 0,
            marginBottom: 0,
            padding: "10px 0 0 0",
            display: "flex",
            justifyContent: "flex-end",
            alignItems: "center",
            gap: "32px",
            fontSize: "1.05rem",
            fontWeight: 500,
            color: darkMode ? "#e5e7eb" : "#334155",
            background: "transparent",
            borderTop: darkMode ? "1px solid #333" : "1px solid #e5e7eb",
            fontFamily: "'JetBrains Mono', 'Fira Mono', 'Menlo', monospace",
            letterSpacing: 0.1,
            userSelect: "none",
          }}
        >
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span
              role="img"
              aria-label="rows"
              style={{ fontSize: "1.1em", opacity: 0.85 }}
            >
              🔢
            </span>
            Rows:{" "}
            <span style={{ fontVariantNumeric: "tabular-nums", marginLeft: 2 }}>
              {nonEmptyRowCount}
            </span>
          </span>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span
              role="img"
              aria-label="columns"
              style={{ fontSize: "1.1em", opacity: 0.85 }}
            >
              📊
            </span>
            Columns:{" "}
            <span style={{ fontVariantNumeric: "tabular-nums", marginLeft: 2 }}>
              {nonEmptyColCount}
            </span>
          </span>
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span
              role="img"
              aria-label="selected cell"
              style={{ fontSize: "1.1em", opacity: 0.85 }}
            >
              🔲
            </span>
            Selected:{" "}
            <span style={{ fontVariantNumeric: "tabular-nums", marginLeft: 2 }}>
              {String.fromCharCode(65 + (focusedCell.col ?? 0))}
              {(focusedCell.row + 1).toString().padStart(2, "0")}
            </span>
          </span>
        </div>
      </div>
    </div>
  );
};

export default ExcelGrid;

import React, { useState, useEffect, useMemo } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx-js-style';
import { 
  Search, 
  X, 
  Settings, 
  Sun, 
  Moon, 
  FileSpreadsheet, 
  ChevronLeft, 
  ChevronRight, 
  History, 
  Trash2, 
  Plus, 
  ExternalLink, 
  Sliders, 
  RefreshCw,
  Info,
  Check,
  Filter,
  PlusCircle,
  Download,
  ArrowUpDown,
  Maximize2,
  ShoppingCart,
  PlusSquare,
  Minus
} from 'lucide-react';

export default function App() {
  // Theme State
  const [theme, setTheme] = useState(() => localStorage.getItem('theme') || 'light');

  // URL input and parse state
  const [sheetUrl, setSheetUrl] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  
  // Sheet Data States
  const [allData, setAllData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [currentSheetInfo, setCurrentSheetInfo] = useState(null);
  
  // App History State (Saved sheets metadata)
  const [history, setHistory] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem('sheet_history')) || [];
    } catch {
      return [];
    }
  });
  
  // Search & Pagination States
  const [searchTerm, setSearchTerm] = useState('');
  const [searchColumn, setSearchColumn] = useState('all'); // 'all' or specific column name
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);
  
  // UI Display Control States
  const [showHistory, setShowHistory] = useState(false);
  const [showConfig, setShowConfig] = useState(false);
  const [advancedFilters, setAdvancedFilters] = useState([]);
  const [showAdvancedFilters, setShowAdvancedFilters] = useState(false);
  const [titleColumn, setTitleColumn] = useState('');
  const [selectedColumns, setSelectedColumns] = useState({}); // column_name -> boolean
  const [nicknameInput, setNicknameInput] = useState('');

  // New States for Features
  const [sortConfig, setSortConfig] = useState({ column: null, direction: 'none' });
  const [selectedRow, setSelectedRow] = useState(null); // Data of the selected row for Detail Modal
  const [showSortMenu, setShowSortMenu] = useState(false);

  // Cart / Pick List States
  const [allPickLists, setAllPickLists] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem('sheet_all_picklists')) || {};
    } catch {
      return {};
    }
  });
  const [showPickList, setShowPickList] = useState(false);
  const [pickingRow, setPickingRow] = useState(null);
  const [pickingQty, setPickingQty] = useState(1);

  // States for Export Config Modal
  const [showExportConfig, setShowExportConfig] = useState(false);
  const [exportType, setExportType] = useState('all'); // 'all' or 'picklist'
  const [exportColsSelection, setExportColsSelection] = useState({}); // { colName: boolean }

  // Derived current pickList
  const currentSheetKey = currentSheetInfo ? `${currentSheetInfo.spreadsheetId}_${currentSheetInfo.gid}` : null;
  const pickList = currentSheetKey && allPickLists[currentSheetKey] ? allPickLists[currentSheetKey] : [];

  // Helper to update the current sheet's pickList
  const setPickList = (updater) => {
    if (!currentSheetKey) return;
    setAllPickLists(prev => {
      const currentList = prev[currentSheetKey] || [];
      const newList = typeof updater === 'function' ? updater(currentList) : updater;
      return { ...prev, [currentSheetKey]: newList };
    });
  };

  // Handle Theme Switching
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
  }, [theme]);

  // Sync History to LocalStorage
  useEffect(() => {
    localStorage.setItem('sheet_history', JSON.stringify(history));
  }, [history]);

  // Sync All Pick Lists to LocalStorage
  useEffect(() => {
    localStorage.setItem('sheet_all_picklists', JSON.stringify(allPickLists));
  }, [allPickLists]);

  // Regex to extract Spreadsheet ID & GID
  const parseSheetUrl = (url) => {
    if (!url) return null;
    
    let spreadsheetId = null;
    let isPublished = false;

    // Check for published format first /d/e/ID
    const pubMatch = url.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (pubMatch) {
      spreadsheetId = pubMatch[1];
      isPublished = true;
    } else {
      // Normal format /d/ID
      const idMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (idMatch) {
        spreadsheetId = idMatch[1];
      }
    }
    
    if (!spreadsheetId) return null;
    
    // GID: sits after gid=
    const gidMatch = url.match(/[#&?]gid=([0-9]+)/);
    
    return {
      spreadsheetId,
      gid: gidMatch ? gidMatch[1] : '0',
      isPublished
    };
  };

  // Main Fetch and Parse Handler
  const loadSheetData = async (url, customName = '', isFromHistory = false) => {
    const parsed = parseSheetUrl(url);
    if (!parsed) {
      setError('Đường dẫn không hợp lệ. Vui lòng nhập link Google Sheet chuẩn.');
      return;
    }
    
    setLoading(true);
    setError(null);
    setSearchTerm('');
    setCurrentPage(1);

    const { spreadsheetId, gid, isPublished } = parsed;
    
    let exportUrl;
    if (isPublished) {
      exportUrl = `https://docs.google.com/spreadsheets/d/e/${spreadsheetId}/pub?gid=${gid}&single=true&output=csv`;
    } else {
      exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`;
    }

    try {
      const response = await fetch(exportUrl);
      if (!response.ok) {
        throw new Error('Không thể tải file CSV. Có thể Sheet chưa được chia sẻ công khai hoặc sai URL.');
      }
      
      const csvText = await response.text();
      
      // Parse CSV Text with PapaParse
      Papa.parse(csvText, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          if (results.errors && results.errors.length > 0 && results.data.length === 0) {
            setError('Lỗi khi phân tích dữ liệu CSV.');
            setLoading(false);
            return;
          }

          const parsedRows = results.data;
          if (parsedRows.length === 0) {
            setError('Sheet trống hoặc không chứa dòng dữ liệu nào hợp lệ.');
            setLoading(false);
            return;
          }

          // Extract headers
          const headers = Object.keys(parsedRows[0]);
          setColumns(headers);
          setAllData(parsedRows);

          // Configure card titles
          // Default to the first column as the title
          const defaultTitleCol = headers[0] || '';
          setTitleColumn(defaultTitleCol);

          // Set all columns selected by default
          const defaultSelected = {};
          headers.forEach(col => {
            defaultSelected[col] = true;
          });
          setSelectedColumns(defaultSelected);

          // Sheet Metadata Info
          const sheetName = customName || `Bảng tính_${spreadsheetId.substring(0, 6)} (Tab: ${gid})`;
          const info = {
            url,
            spreadsheetId,
            gid,
            name: sheetName,
            timestamp: Date.now(),
            rowCount: parsedRows.length,
            headers
          };
          
          setCurrentSheetInfo(info);
          setNicknameInput(sheetName);

          // Add/Update History
          if (!isFromHistory) {
            setHistory(prev => {
              // Filter out duplicate sheet-id & gid
              const filtered = prev.filter(item => 
                !(item.spreadsheetId === spreadsheetId && item.gid === gid)
              );
              return [info, ...filtered].slice(0, 10); // Keep max 10 entries
            });
          } else {
            // Update timestamp in history
            setHistory(prev => 
              prev.map(item => 
                (item.spreadsheetId === spreadsheetId && item.gid === gid) 
                  ? { ...item, timestamp: Date.now() } 
                  : item
              )
            );
          }

          setLoading(false);
          setShowHistory(false);
        },
        error: (err) => {
          setError(`Lỗi PapaParse: ${err.message}`);
          setLoading(false);
        }
      });

    } catch (err) {
      setError(err.message || 'Không kết nối được với máy chủ Google Sheets.');
      setLoading(false);
    }
  };

  // Update sheet nickname/custom name
  const updateSheetNickname = () => {
    if (!nicknameInput.trim() || !currentSheetInfo) return;
    
    const updatedInfo = { ...currentSheetInfo, name: nicknameInput.trim() };
    setCurrentSheetInfo(updatedInfo);
    
    setHistory(prev => 
      prev.map(item => 
        (item.spreadsheetId === currentSheetInfo.spreadsheetId && item.gid === currentSheetInfo.gid)
          ? { ...item, name: nicknameInput.trim() }
          : item
      )
    );
  };

  // Remove single sheet from history
  const deleteHistoryItem = (e, index) => {
    e.stopPropagation();
    setHistory(prev => prev.filter((_, i) => i !== index));
  };

  // Clear all history
  const clearAllHistory = () => {
    if (window.confirm('Bạn có chắc muốn xóa toàn bộ lịch sử đã xem?')) {
      setHistory([]);
    }
  };

  // Advanced filter helper handlers
  const addFilter = () => {
    setAdvancedFilters([...advancedFilters, { id: Date.now(), column: columns[0] || '', operator: 'contains', value: '' }]);
  };

  const updateFilter = (id, field, newValue) => {
    setAdvancedFilters(prev => prev.map(f => f.id === id ? { ...f, [field]: newValue } : f));
  };

  const removeFilter = (id) => {
    setAdvancedFilters(prev => prev.filter(f => f.id !== id));
  };

  // Handle opening Export configuration
  const handleOpenExportAll = () => {
    if (filteredData.length === 0) return;
    const initialSelection = {};
    columns.forEach(col => {
      initialSelection[col] = selectedColumns[col] !== false; // Check only visible ones by default
    });
    setExportColsSelection(initialSelection);
    setExportType('all');
    setShowExportConfig(true);
  };

  const handleOpenExportPickList = () => {
    if (pickList.length === 0) return;
    const initialSelection = {};
    columns.forEach(col => {
      initialSelection[col] = selectedColumns[col] !== false; // Check only visible ones by default
    });
    setExportColsSelection(initialSelection);
    setExportType('picklist');
    setShowExportConfig(true);
  };

  // Perform actual XLSX Export based on selection
  const triggerExport = () => {
    const selectedColsList = columns.filter(col => exportColsSelection[col]);
    if (selectedColsList.length === 0) {
      alert('Vui lòng chọn ít nhất một cột để xuất.');
      return;
    }

    const autoFit = (ws, data) => {
      if (data.length === 0) return;
      const headers = Object.keys(data[0] || {});
      ws['!cols'] = headers.map(header => {
        let max_len = header.length;
        data.forEach(row => {
          const val = row[header] !== undefined && row[header] !== null ? String(row[header]) : "";
          if (val.length > max_len) {
            max_len = val.length;
          }
        });
        return { wch: Math.max(10, max_len + 3) };
      });
    };

    const styleHeaders = (ws) => {
      Object.keys(ws).forEach(key => {
        if (key.match(/^[A-Z]+1$/)) {
          ws[key].s = {
            font: {
              bold: true,
              color: { rgb: "FFFFFF" },
              sz: 11,
              name: "Calibri"
            },
            fill: {
              fgColor: { rgb: "4F46E5" } // Dark blue / Purple theme color
            },
            alignment: {
              horizontal: "center",
              vertical: "center"
            }
          };
        }
      });
    };

    if (exportType === 'all') {
      const exportData = filteredData.map(row => {
        const rowData = {};
        selectedColsList.forEach(col => {
          rowData[col] = row[col];
        });
        return rowData;
      });
      const worksheet = XLSX.utils.json_to_sheet(exportData);
      autoFit(worksheet, exportData);
      styleHeaders(worksheet);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
      XLSX.writeFile(workbook, `${currentSheetInfo.name}_export.xlsx`);
      setShowExportConfig(false);
    } else {
      const exportData = pickList.map(item => {
        const rowData = {};
        selectedColsList.forEach(col => {
          rowData[col] = item.originalRow[col];
        });
        rowData['Số lượng lấy'] = item.quantity;
        return rowData;
      });
      const worksheet = XLSX.utils.json_to_sheet(exportData);
      autoFit(worksheet, exportData);
      styleHeaders(worksheet);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Pick List");
      XLSX.writeFile(workbook, `danh_sach_nhat_hang_${new Date().getTime()}.xlsx`);
      
      setShowExportConfig(false);
      setShowPickList(false);
      
      setTimeout(() => {
        if (window.confirm('Đã tải xong file Excel! Bạn có muốn xóa danh sách hiện tại để nhặt đơn mới không?')) {
          setPickList([]);
        }
      }, 500);
    }
  };

  // Smart Formatting
  const formatCellValue = (value) => {
    if (!value) return <span style={{ color: 'var(--text-tertiary)' }}>—</span>;
    const str = String(value);
    
    // Format Images
    if (str.match(/\.(jpeg|jpg|gif|png|webp)$/i) && str.startsWith('http')) {
      return <img src={str} alt="Data" style={{ maxWidth: '100%', maxHeight: '120px', borderRadius: '4px', marginTop: '4px' }} />;
    }
    
    // Format Links
    if (str.startsWith('http://') || str.startsWith('https://')) {
      return <a href={str} target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent)', textDecoration: 'underline' }}>{str}</a>;
    }
    
    // Format Badges (Simple keyword matching)
    const lower = str.toLowerCase();
    if (['hoàn thành', 'done', 'thành công', 'đã giao'].includes(lower)) {
      return <span className="status-badge success">{str}</span>;
    }
    if (['hủy', 'đã hủy', 'cancel', 'thất bại', 'lỗi'].includes(lower)) {
      return <span className="status-badge danger">{str}</span>;
    }
    if (['đang xử lý', 'pending', 'đang giao', 'chờ'].includes(lower)) {
      return <span className="status-badge warning">{str}</span>;
    }
    
    return str;
  };

  // Filter Data Logic (Memoized for performance)
  const filteredData = useMemo(() => {
    let result = allData;

    if (advancedFilters.length > 0) {
      result = result.filter(row => {
        return advancedFilters.every(filter => {
          const { column, operator, value } = filter;
          if (!column) return true;
          
          const cellValue = String(row[column] || '');
          const cellLower = cellValue.toLowerCase();
          const valLower = String(value || '').toLowerCase();
          
          switch (operator) {
            case 'contains': return cellLower.includes(valLower);
            case 'not_contains': return !cellLower.includes(valLower);
            case 'equals': return cellLower === valLower;
            case 'not_equals': return cellLower !== valLower;
            case 'greater_than': return Number(cellValue) > Number(value);
            case 'less_than': return Number(cellValue) < Number(value);
            case 'starts_with': return cellLower.startsWith(valLower);
            case 'ends_with': return cellLower.endsWith(valLower);
            case 'is_empty': return cellValue.trim() === '';
            case 'is_not_empty': return cellValue.trim() !== '';
            default: return true;
          }
        });
      });
    }

    if (!searchTerm.trim()) return result;
    
    const term = searchTerm.toLowerCase();
    
    return result.filter(row => {
      if (searchColumn === 'all') {
        // Search across all columns
        return Object.values(row).some(val => 
          String(val).toLowerCase().includes(term)
        );
      } else {
        // Search in specific column only
        return String(row[searchColumn] || '').toLowerCase().includes(term);
      }
    });

    // Apply Sorting
    if (sortConfig.column && sortConfig.direction !== 'none') {
      result = [...result].sort((a, b) => {
        const valA = a[sortConfig.column];
        const valB = b[sortConfig.column];
        
        if (valA === valB) return 0;
        if (valA == null) return 1;
        if (valB == null) return -1;
        
        const numA = Number(valA);
        const numB = Number(valB);
        if (!isNaN(numA) && !isNaN(numB)) {
          return sortConfig.direction === 'asc' ? numA - numB : numB - numA;
        }
        
        const strA = String(valA).toLowerCase();
        const strB = String(valB).toLowerCase();
        if (strA < strB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (strA > strB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }

    return result;
  }, [allData, searchTerm, searchColumn, advancedFilters, sortConfig]);

  // Pagination Logic
  const pageCount = Math.ceil(filteredData.length / pageSize) || 1;
  
  // Safe page index alignment
  useEffect(() => {
    if (currentPage > pageCount) {
      setCurrentPage(pageCount);
    }
  }, [pageCount, currentPage]);

  const displayData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredData.slice(start, start + pageSize);
  }, [filteredData, currentPage, pageSize]);

  // Helper to scroll back to top of card list
  const scrollToTop = () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handlePageChange = (page) => {
    if (page >= 1 && page <= pageCount) {
      setCurrentPage(page);
      scrollToTop();
    }
  };

  // Toggle visible status of columns
  const toggleColumnSelection = (col) => {
    setSelectedColumns(prev => ({
      ...prev,
      [col]: !prev[col]
    }));
  };

  return (
    <div className="app-container">
      {/* 1. APP HEADER */}
      <header className="app-header">
        <div className="header-top">
          <div className="brand">
            <span className="brand-icon">S</span>
            <h1 className="brand-name">SheetFinder</h1>
          </div>
          
          <div className="action-buttons">
            <button 
              className="icon-btn" 
              onClick={() => setShowPickList(true)} 
              title="Danh sách nhặt hàng"
              style={{ position: 'relative' }}
            >
              <ShoppingCart size={18} />
              {pickList.length > 0 && (
                <span style={{ position: 'absolute', top: -4, right: -4, background: 'var(--danger)', color: 'white', borderRadius: '50%', width: 16, height: 16, fontSize: 10, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                  {pickList.length}
                </span>
              )}
            </button>
            <button 
              className="icon-btn" 
              onClick={() => setShowHistory(!showHistory)} 
              title="Lịch sử xem"
            >
              <History size={18} />
            </button>
            <button 
              className="icon-btn" 
              onClick={() => setTheme(theme === 'light' ? 'dark' : 'light')} 
              title="Đổi giao diện"
            >
              {theme === 'light' ? <Moon size={18} /> : <Sun size={18} />}
            </button>
          </div>
        </div>

        {/* Dynamic nickname edit or mini status bar */}
        {currentSheetInfo && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.8rem' }}>
            <span style={{ color: 'var(--text-tertiary)' }}>Đang xem:</span>
            <input 
              type="text" 
              value={nicknameInput}
              onChange={(e) => setNicknameInput(e.target.value)}
              onBlur={updateSheetNickname}
              onKeyDown={(e) => e.key === 'Enter' && updateSheetNickname()}
              style={{
                background: 'transparent',
                border: 'none',
                borderBottom: '1px dashed var(--border)',
                fontWeight: '600',
                color: 'var(--accent)',
                padding: '2px 4px',
                flex: 1,
                fontSize: '0.85rem'
              }}
              title="Bấm vào để đổi tên gợi nhớ"
            />
            <button 
              className="icon-btn" 
              style={{ width: '26px', height: '26px', borderRadius: '4px' }}
              onClick={() => loadSheetData(currentSheetInfo.url, currentSheetInfo.name, true)}
              title="Làm mới dữ liệu"
            >
              <RefreshCw size={12} />
            </button>
            <button 
              className="icon-btn" 
              style={{ width: '26px', height: '26px', borderRadius: '4px' }}
              onClick={() => {
                setAllData([]);
                setCurrentSheetInfo(null);
                setSheetUrl('');
              }}
              title="Đóng trang này"
            >
              <X size={12} />
            </button>
          </div>
        )}

        {/* 2. HISTORY DRAWER POPUP */}
        {showHistory && (
          <div className="history-drawer">
            <div className="history-title">
              <span>Lịch sử tra cứu ({history.length})</span>
              {history.length > 0 && (
                <button 
                  onClick={clearAllHistory} 
                  style={{ display: 'flex', alignItems: 'center', gap: '4px', color: 'var(--danger)', fontSize: '0.75rem', fontWeight: '600' }}
                >
                  <Trash2 size={12} /> Xóa hết
                </button>
              )}
            </div>
            
            {history.length === 0 ? (
              <p style={{ textAlign: 'center', fontSize: '0.85rem', color: 'var(--text-tertiary)', padding: '16px 0' }}>
                Chưa có lịch sử bảng tính nào được lưu.
              </p>
            ) : (
              <div className="history-list">
                {history.map((item, idx) => (
                  <div 
                    key={idx} 
                    className="history-item"
                    onClick={() => loadSheetData(item.url, item.name, true)}
                  >
                    <div className="history-item-content">
                      <div className="history-item-name">{item.name}</div>
                      <div className="history-item-meta">
                        {item.rowCount} dòng • Tab: {item.gid} • {new Date(item.timestamp).toLocaleDateString('vi-VN')}
                      </div>
                    </div>
                    <button 
                      className="delete-history-btn" 
                      onClick={(e) => deleteHistoryItem(e, idx)}
                      title="Xóa khỏi lịch sử"
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </header>

      {/* 3. SETUP VIEW (Paste URL) */}
      {!currentSheetInfo && (
        <main className="setup-container">
          <div className="welcome-card">
            <span className="welcome-emoji">📱</span>
            <h2 className="welcome-title">Google Sheet Tra Cứu Nhanh</h2>
            <p className="welcome-desc">
              Hỗ trợ tìm kiếm, lọc và xem bảng dữ liệu mượt mà, không bị lag trên trình duyệt điện thoại.
            </p>
          </div>

          <div className="card input-group">
            <label className="input-label" htmlFor="sheet-url">Nhập đường dẫn Google Sheet công khai:</label>
            <div className="input-wrapper">
              <span className="input-icon">
                <FileSpreadsheet size={18} />
              </span>
              <input
                id="sheet-url"
                type="text"
                placeholder="https://docs.google.com/spreadsheets/d/..."
                value={sheetUrl}
                onChange={(e) => setSheetUrl(e.target.value)}
                className="text-input"
              />
            </div>
            <button
              onClick={() => loadSheetData(sheetUrl)}
              disabled={loading || !sheetUrl.trim()}
              className="primary-btn"
            >
              {loading ? (
                <>
                  <div className="spinner" style={{ width: '16px', height: '16px', borderWidth: '2px' }}></div>
                  Đang tải dữ liệu...
                </>
              ) : (
                <>
                  <Plus size={16} />
                  Tải và hiển thị Sheet
                </>
              )}
            </button>
          </div>

          {error && (
            <div className="error-container">
              <div className="error-title">Không lấy được dữ liệu</div>
              <div className="error-message">{error}</div>
              <div style={{ textAlign: 'left', marginTop: '8px', fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
                <strong style={{ display: 'block', marginBottom: '4px', color: 'var(--text-primary)' }}>Hướng dẫn chia sẻ Sheet:</strong>
                1. Mở Google Sheet trên máy tính.<br />
                2. Click nút <strong>Chia sẻ (Share)</strong> ở góc phải trên.<br />
                3. Đổi Quyền truy cập chung sang <strong>Bất kỳ ai có liên kết (Anyone with the link)</strong>.<br />
                4. Sao chép liên kết đó dán vào ứng dụng này.
              </div>
            </div>
          )}

          {/* Quick instructions setup layout */}
          <section className="instructions-section">
            <h3 className="instruction-title">Lợi ích của SheetFinder</h3>
            <div className="instruction-step">
              <span className="step-num">1</span>
              <div className="step-text">
                <strong>Tìm kiếm siêu tốc:</strong> Tìm kiếm tức thời trên toàn bộ các cột của bảng dữ liệu mà không cần tải lại trang.
              </div>
            </div>
            <div className="instruction-step">
              <span className="step-num">2</span>
              <div className="step-text">
                <strong>Giao diện dạng Thẻ:</strong> Tối ưu cho màn hình dọc điện thoại, không cần cuộn ngang mệt mỏi.
              </div>
            </div>
            <div className="instruction-step">
              <span className="step-num">3</span>
              <div className="step-text">
                <strong>Phân trang thông minh:</strong> Chỉ hiển thị 20 thẻ mỗi trang để tránh gây lag và đơ máy.
              </div>
            </div>
            <div className="instruction-step">
              <span className="step-num">4</span>
              <div className="step-text">
                <strong>Hỗ trợ Offline:</strong> Lưu lịch sử các bảng đã xem để bạn tiện xem lại nhanh chóng bất cứ lúc nào.
              </div>
            </div>
          </section>
        </main>
      )}

      {/* 4. DASHBOARD VIEW (Search, filters & Card render) */}
      {currentSheetInfo && (
        <main className="main-dashboard">
          {/* Sticky search panel */}
          <div className="filter-panel">
            <div className="search-container" style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <div className="search-wrapper" style={{ width: '100%' }}>
                <span className="input-icon" style={{ left: '12px' }}>
                  <Search size={16} />
                </span>
                <input
                  type="text"
                  placeholder={
                    searchColumn === 'all' 
                      ? "Tìm kiếm trên tất cả cột..." 
                      : `Tìm kiếm trong cột "${searchColumn}"...`
                  }
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setCurrentPage(1); // Reset to page 1 on new search
                  }}
                  className="search-input"
                />
                {searchTerm && (
                  <button 
                    className="clear-search-btn" 
                    onClick={() => {
                      setSearchTerm('');
                      setCurrentPage(1);
                    }}
                  >
                    <X size={16} />
                  </button>
                )}
              </div>

              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
                {/* Action Toolbar */}
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                  <button 
                    className="icon-btn" 
                    onClick={handleOpenExportAll}
                    title="Xuất file Excel"
                  >
                    <Download size={18} />
                  </button>
                  <div style={{ position: 'relative' }}>
                    <button 
                      className="icon-btn" 
                      onClick={() => setShowSortMenu(!showSortMenu)}
                      title="Sắp xếp"
                      style={sortConfig.direction !== 'none' ? { backgroundColor: 'var(--accent-light)', color: 'var(--accent)', borderColor: 'var(--accent)' } : {}}
                    >
                      <ArrowUpDown size={18} />
                    </button>
                    {showSortMenu && (
                      <div className="sort-menu-dropdown" style={{ left: 0, right: 'auto', width: 'max-content', minWidth: '220px' }}>
                        <div className="sort-title">Sắp xếp theo:</div>
                        <select 
                          value={sortConfig.column || ''} 
                          onChange={(e) => setSortConfig(prev => ({ ...prev, column: e.target.value, direction: prev.direction === 'none' ? 'asc' : prev.direction }))}
                          className="filter-select"
                          style={{ width: '100%', marginBottom: '8px' }}
                        >
                          <option value="" disabled>Chọn cột</option>
                          {columns.map(col => <option key={col} value={col}>{col}</option>)}
                        </select>
                        <div style={{ display: 'flex', gap: '4px' }}>
                          <button 
                            className={`pill ${sortConfig.direction === 'asc' ? 'active' : ''}`} 
                            onClick={() => setSortConfig(prev => ({ column: prev.column || columns[0], direction: 'asc' }))} 
                            style={{ flex: 1, justifyContent: 'center' }}
                          >A-Z</button>
                          <button 
                            className={`pill ${sortConfig.direction === 'desc' ? 'active' : ''}`} 
                            onClick={() => setSortConfig(prev => ({ column: prev.column || columns[0], direction: 'desc' }))} 
                            style={{ flex: 1, justifyContent: 'center' }}
                          >Z-A</button>
                          <button className="icon-btn" style={{ width: 32, height: 32 }} onClick={() => { setSortConfig({ column: null, direction: 'none' }); setShowSortMenu(false); }}><X size={14}/></button>
                        </div>
                      </div>
                    )}
                  </div>
                  <button 
                    className="icon-btn" 
                    onClick={() => setShowAdvancedFilters(!showAdvancedFilters)}
                    title="Bộ lọc nâng cao"
                    style={{ position: 'relative', ...(showAdvancedFilters || advancedFilters.length > 0 ? { backgroundColor: 'var(--accent-light)', color: 'var(--accent)', borderColor: 'var(--accent)' } : {}) }}
                  >
                    <Filter size={18} />
                    {advancedFilters.length > 0 && (
                      <span style={{ position: 'absolute', top: -4, right: -4, background: 'var(--danger)', color: 'white', borderRadius: '50%', width: 16, height: 16, fontSize: 10, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                        {advancedFilters.length}
                      </span>
                    )}
                  </button>
                  <button 
                    className="icon-btn" 
                    onClick={() => setShowConfig(true)}
                    title="Cấu hình hiển thị"
                  >
                    <Sliders size={18} />
                  </button>
                </div>
              </div>
            </div>

            {/* Advanced Filters Builder */}
            {showAdvancedFilters && (
              <div className="advanced-filters-container">
                <div style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '8px', display: 'flex', justifyContent: 'space-between' }}>
                  <span>Điều kiện lọc (AND):</span>
                  <button onClick={() => setAdvancedFilters([])} style={{ color: 'var(--danger)', fontSize: '0.8rem' }}>Xóa tất cả</button>
                </div>
                {advancedFilters.length === 0 ? (
                  <div style={{ textAlign: 'center', padding: '16px', color: 'var(--text-tertiary)', fontSize: '0.85rem', border: '1px dashed var(--border)', borderRadius: 'var(--radius-sm)' }}>
                    Chưa có điều kiện nào. Bấm thêm điều kiện bên dưới.
                  </div>
                ) : (
                  <div className="filters-list">
                    {advancedFilters.map((f) => (
                      <div key={f.id} className="filter-row">
                        <select className="filter-select" value={f.column} onChange={e => updateFilter(f.id, 'column', e.target.value)}>
                          {columns.map(col => <option key={col} value={col}>{col}</option>)}
                        </select>
                        <select className="filter-select" value={f.operator} onChange={e => updateFilter(f.id, 'operator', e.target.value)}>
                          <option value="contains">Chứa (Contains)</option>
                          <option value="not_contains">Không chứa</option>
                          <option value="equals">Bằng (=)</option>
                          <option value="not_equals">Khác (!=)</option>
                          <option value="greater_than">Lớn hơn (&gt;)</option>
                          <option value="less_than">Nhỏ hơn (&lt;)</option>
                          <option value="starts_with">Bắt đầu bằng</option>
                          <option value="ends_with">Kết thúc bằng</option>
                          <option value="is_empty">Rỗng</option>
                          <option value="is_not_empty">Không rỗng</option>
                        </select>
                        {!['is_empty', 'is_not_empty'].includes(f.operator) && (
                          <input type="text" className="filter-input" placeholder="Giá trị..." value={f.value} onChange={e => updateFilter(f.id, 'value', e.target.value)} />
                        )}
                        <button className="icon-btn" style={{ width: 32, height: 32, flexShrink: 0 }} onClick={() => removeFilter(f.id)}><X size={14} /></button>
                      </div>
                    ))}
                  </div>
                )}
                <button className="primary-btn" style={{ marginTop: '12px', padding: '8px 12px', fontSize: '0.85rem', width: '100%' }} onClick={addFilter}>
                  <PlusCircle size={16} /> Thêm điều kiện lọc
                </button>
              </div>
            )}

            {/* Quick search-column filter pill selector */}
            <div className="filter-pills">
              <button 
                className={`pill ${searchColumn === 'all' ? 'active' : ''}`}
                onClick={() => {
                  setSearchColumn('all');
                  setCurrentPage(1);
                }}
              >
                Tất cả cột
              </button>
              {columns.map((col, idx) => (
                <button
                  key={idx}
                  className={`pill ${searchColumn === col ? 'active' : ''}`}
                  onClick={() => {
                    setSearchColumn(col);
                    setCurrentPage(1);
                  }}
                >
                  Cột: {col}
                </button>
              ))}
            </div>
          </div>

          {/* Statistics Ribbon */}
          <div className="stats-ribbon">
            <div className="stats-summary">
              Tìm thấy <span className="stats-count">{filteredData.length}</span> dòng 
              {searchTerm && ` (trong tổng số ${allData.length})`}
            </div>
            <div className="active-sheet-name" title={currentSheetInfo.name}>
              {currentSheetInfo.name}
            </div>
          </div>

          {/* Main Card render loop */}
          {displayData.length === 0 ? (
            <div className="loading-container" style={{ minHeight: '30vh' }}>
              <Info size={36} style={{ color: 'var(--text-tertiary)' }} />
              <p style={{ color: 'var(--text-secondary)', textAlign: 'center' }}>
                Không tìm thấy dòng nào khớp với từ khóa tìm kiếm.
              </p>
            </div>
          ) : (
            <div className="cards-grid">
              {displayData.map((row, index) => {
                // Calculate original spreadsheet row index (1-based, ignoring header)
                const globalIndex = (currentPage - 1) * pageSize + index + 1;
                
                // Card title is defined by titleColumn state
                const cardTitle = row[titleColumn] || `Dòng thứ #${globalIndex}`;

                return (
                  <article key={index} className="data-card" onClick={() => setSelectedRow(row)} style={{ cursor: 'pointer' }}>
                    <header className="card-header">
                      <h4 className="card-title">{cardTitle}</h4>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span className="card-index">#{globalIndex}</span>
                        <button 
                          className="icon-btn" 
                          onClick={(e) => { e.stopPropagation(); setPickingRow(row); setPickingQty(1); }}
                          title="Thêm vào danh sách nhặt"
                          style={{ color: 'var(--accent)', padding: '4px' }}
                        >
                          <PlusSquare size={16} />
                        </button>
                        <Maximize2 size={14} style={{ color: 'var(--text-tertiary)' }} />
                      </div>
                    </header>
                    
                    <div className="card-body">
                      {columns
                        .filter(col => col !== titleColumn && selectedColumns[col])
                        .map((col, colIdx) => (
                          <div key={colIdx} className="property-row">
                            <span className="property-label">{col}:</span>
                            <span className="property-value">{formatCellValue(row[col])}</span>
                          </div>
                        ))}
                    </div>
                  </article>
                );
              })}
            </div>
          )}

          {/* 5. STICKY BOTTOM PAGINATION BAR */}
          <div className="pagination-sticky">
            <div className="page-controls">
              <button
                className="page-btn"
                disabled={currentPage === 1}
                onClick={() => handlePageChange(currentPage - 1)}
                title="Trang trước"
              >
                <ChevronLeft size={16} />
                Trước
              </button>

              <div className="page-indicator">
                <select 
                  className="jump-select"
                  value={currentPage}
                  onChange={(e) => handlePageChange(Number(e.target.value))}
                >
                  {Array.from({ length: pageCount }, (_, i) => (
                    <option key={i + 1} value={i + 1}>
                      Trang {i + 1} / {pageCount}
                    </option>
                  ))}
                </select>
              </div>

              <button
                className="page-btn"
                disabled={currentPage === pageCount}
                onClick={() => handlePageChange(currentPage + 1)}
                title="Trang sau"
              >
                Sau
                <ChevronRight size={16} />
              </button>
            </div>
          </div>
        </main>
      )}

      {/* 6. CONFIG DRAWER MODAL (Column selection & card title selection) */}
      {showConfig && (
        <div className="drawer-overlay" onClick={() => setShowConfig(false)}>
          <div className="drawer-content" onClick={(e) => e.stopPropagation()}>
            <header className="drawer-header">
              <h3 className="drawer-title">Cấu hình hiển thị Card</h3>
              <button className="close-btn" onClick={() => setShowConfig(false)}>
                <X size={20} />
              </button>
            </header>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
              {/* Select Title Column */}
              <div className="settings-list">
                <span className="settings-section-title">Chọn cột làm tiêu đề Card:</span>
                <div className="radio-group">
                  {columns.map((col, idx) => (
                    <label key={idx} className="radio-option">
                      <input
                        type="radio"
                        name="title-column"
                        value={col}
                        checked={titleColumn === col}
                        onChange={() => setTitleColumn(col)}
                        style={{ accentColor: 'var(--accent)', width: '16px', height: '16px' }}
                      />
                      <span>{col}</span>
                    </label>
                  ))}
                </div>
              </div>

              {/* Show/Hide Columns Toggle */}
              <div className="settings-list">
                <span className="settings-section-title">Hiển thị các cột nội dung:</span>
                <div className="checkbox-group">
                  {columns.map((col, idx) => {
                    const isTitle = col === titleColumn;
                    return (
                      <div 
                        key={idx} 
                        className="checkbox-option"
                        style={isTitle ? { opacity: 0.5, cursor: 'not-allowed' } : {}}
                      >
                        <div className="option-left">
                          {selectedColumns[col] && !isTitle ? <Check size={16} style={{ color: 'var(--success)' }} /> : <span style={{ width: '16px' }}></span>}
                          <span style={isTitle ? { fontWeight: '600' } : {}}>{col} {isTitle && '(Tiêu đề)'}</span>
                        </div>
                        <label className="switch">
                          <input
                            type="checkbox"
                            checked={selectedColumns[col] || false}
                            disabled={isTitle}
                            onChange={() => toggleColumnSelection(col)}
                          />
                          <span className="slider"></span>
                        </label>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Page Size select */}
              <div className="settings-list" style={{ borderTop: '1px solid var(--border)', paddingTop: '16px' }}>
                <span className="settings-section-title">Số dòng mỗi trang:</span>
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(Number(e.target.value));
                    setCurrentPage(1);
                  }}
                  className="jump-select"
                  style={{ width: '100%', textAlign: 'left', padding: '12px' }}
                >
                  <option value={10}>10 dòng / trang</option>
                  <option value={20}>20 dòng / trang</option>
                  <option value={50}>50 dòng / trang</option>
                  <option value={100}>100 dòng / trang</option>
                </select>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 7. DETAIL MODAL */}
      {selectedRow && (
        <div className="drawer-overlay" onClick={() => setSelectedRow(null)}>
          <div className="drawer-content detail-modal" onClick={(e) => e.stopPropagation()}>
            <header className="drawer-header" style={{ position: 'sticky', top: 0, backgroundColor: 'var(--bg-card)', zIndex: 10, paddingBottom: '12px', borderBottom: '1px solid var(--border)' }}>
              <h3 className="drawer-title" style={{ fontSize: '1.2rem' }}>{selectedRow[titleColumn] || 'Chi tiết'}</h3>
              <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                <button 
                  className="primary-btn"
                  onClick={() => { setPickingRow(selectedRow); setPickingQty(1); setSelectedRow(null); }}
                  style={{ padding: '6px 12px', fontSize: '0.85rem' }}
                >
                  <PlusSquare size={14} /> Chọn
                </button>
                <button className="close-btn" onClick={() => setSelectedRow(null)}>
                  <X size={20} />
                </button>
              </div>
            </header>
            <div className="detail-body">
              {columns.map((col, idx) => (
                <div key={idx} className="detail-row">
                  <div className="detail-label">{col}</div>
                  <div className="detail-value">{formatCellValue(selectedRow[col])}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* 8. PICKING QUANTITY MODAL */}
      {pickingRow && (
        <div className="drawer-overlay" onClick={() => setPickingRow(null)} style={{ zIndex: 200, alignItems: 'center', display: 'flex', justifyContent: 'center', padding: '20px' }}>
          <div className="card" onClick={(e) => e.stopPropagation()} style={{ width: '100%', maxWidth: '320px', padding: '20px', animation: 'slideUp 0.2s ease-out' }}>
            <h3 style={{ marginBottom: '16px', fontSize: '1.1rem' }}>Nhập số lượng</h3>
            <p style={{ fontSize: '0.9rem', color: 'var(--text-secondary)', marginBottom: '16px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
              {pickingRow[titleColumn]}
            </p>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '24px', justifyContent: 'center' }}>
              <button className="icon-btn" style={{ width: 40, height: 40, backgroundColor: 'var(--bg-primary)' }} onClick={() => setPickingQty(Math.max(1, pickingQty - 1))}><Minus size={20} /></button>
              <input 
                type="number" 
                value={pickingQty} 
                onChange={(e) => setPickingQty(Math.max(1, parseInt(e.target.value) || 1))}
                className="filter-input"
                style={{ width: '80px', textAlign: 'center', fontSize: '1.2rem', padding: '8px' }}
              />
              <button className="icon-btn" style={{ width: 40, height: 40, backgroundColor: 'var(--bg-primary)' }} onClick={() => setPickingQty(pickingQty + 1)}><Plus size={20} /></button>
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button className="secondary-btn" onClick={() => setPickingRow(null)} style={{ flex: 1, padding: '10px' }}>Hủy</button>
              <button className="primary-btn" style={{ flex: 1, padding: '10px' }} onClick={() => {
                setPickList(prev => {
                  const exists = prev.findIndex(item => JSON.stringify(item.originalRow) === JSON.stringify(pickingRow));
                  if (exists >= 0) {
                    const newList = [...prev];
                    newList[exists].quantity += pickingQty;
                    return newList;
                  }
                  return [...prev, { originalRow: pickingRow, quantity: pickingQty, addedAt: Date.now() }];
                });
                setPickingRow(null);
              }}>
                <Check size={16} /> Xác nhận
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 9. PICK LIST DRAWER */}
      {showPickList && (
        <div className="drawer-overlay" onClick={() => setShowPickList(false)} style={{ zIndex: 150 }}>
          <div className="drawer-content" onClick={(e) => e.stopPropagation()} style={{ height: '85vh', display: 'flex', flexDirection: 'column', padding: 0 }}>
            <header className="drawer-header" style={{ padding: '16px', borderBottom: '1px solid var(--border)' }}>
              <h3 className="drawer-title" style={{ fontSize: '1.1rem' }}>Danh sách đã chọn ({pickList.length})</h3>
              <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                {pickList.length > 0 && (
                  <button 
                    onClick={() => {
                      if (window.confirm('Bạn có chắc muốn xóa toàn bộ danh sách đã chọn?')) {
                        setPickList([]);
                      }
                    }} 
                    style={{ display: 'flex', alignItems: 'center', gap: '4px', color: 'var(--danger)', fontSize: '0.8rem', fontWeight: '600' }}
                  >
                    <Trash2 size={14} /> Xóa hết
                  </button>
                )}
                <button className="close-btn" onClick={() => setShowPickList(false)}>
                  <X size={20} />
                </button>
              </div>
            </header>
            
            {pickList.length === 0 ? (
              <p style={{ textAlign: 'center', fontSize: '0.9rem', color: 'var(--text-tertiary)', padding: '32px 16px' }}>
                Danh sách trống. Bấm [+] trên thẻ để thêm món.
              </p>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
                <div className="history-list" style={{ flex: 1, overflowY: 'auto', padding: '16px' }}>
                  {pickList.map((item, idx) => (
                    <div key={idx} className="history-item" style={{ padding: '12px', cursor: 'default', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', marginBottom: '8px' }}>
                      <div className="history-item-content">
                        <div className="history-item-name">{item.originalRow[titleColumn]}</div>
                        <div className="history-item-meta" style={{ display: 'flex', alignItems: 'center', gap: '12px', marginTop: '6px' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <button 
                              className="icon-btn" 
                              style={{ width: 24, height: 24, padding: 0, backgroundColor: 'var(--bg-secondary)', color: 'var(--text-primary)' }}
                              onClick={() => {
                                setPickList(prev => prev.map((p, i) => i === idx ? { ...p, quantity: Math.max(1, p.quantity - 1) } : p))
                              }}
                            >
                              <Minus size={14} />
                            </button>
                            <span style={{ color: 'var(--accent)', fontWeight: 'bold', fontSize: '0.95rem', minWidth: '24px', textAlign: 'center' }}>{item.quantity}</span>
                            <button 
                              className="icon-btn" 
                              style={{ width: 24, height: 24, padding: 0, backgroundColor: 'var(--bg-secondary)', color: 'var(--text-primary)' }}
                              onClick={() => {
                                setPickList(prev => prev.map((p, i) => i === idx ? { ...p, quantity: p.quantity + 1 } : p))
                              }}
                            >
                              <Plus size={14} />
                            </button>
                          </div>
                          <span style={{ fontSize: '0.75rem', color: 'var(--text-tertiary)', marginLeft: 'auto' }}>{new Date(item.addedAt).toLocaleTimeString('vi-VN')}</span>
                        </div>
                      </div>
                      <button 
                        className="delete-history-btn" 
                        onClick={() => setPickList(prev => prev.filter((_, i) => i !== idx))}
                        title="Xóa món này"
                        style={{ padding: '8px' }}
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                </div>
                <div style={{ padding: '16px', borderTop: '1px solid var(--border)', backgroundColor: 'var(--bg-primary)' }}>
                  <button 
                    className="primary-btn" 
                    style={{ width: '100%', padding: '12px', fontSize: '1rem', justifyContent: 'center' }}
                    onClick={handleOpenExportPickList}
                  >
                    <Download size={18} /> Xuất Excel danh sách này
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* 10. EXPORT CONFIGURATION MODAL */}
      {showExportConfig && (
        <div className="drawer-overlay" onClick={() => setShowExportConfig(false)} style={{ zIndex: 300 }}>
          <div className="drawer-content" onClick={(e) => e.stopPropagation()} style={{ maxHeight: '80vh', display: 'flex', flexDirection: 'column', padding: '20px' }}>
            <header className="drawer-header" style={{ marginBottom: '16px', paddingBottom: '12px' }}>
              <h3 className="drawer-title" style={{ fontSize: '1.2rem' }}>Cấu hình cột xuất Excel</h3>
              <button className="close-btn" onClick={() => setShowExportConfig(false)}>
                <X size={20} />
              </button>
            </header>

            <div style={{ flex: 1, overflowY: 'auto', marginBottom: '20px' }}>
              <div style={{ display: 'flex', gap: '8px', marginBottom: '16px' }}>
                <button 
                  className="pill" 
                  style={{ flex: 1, justifyContent: 'center' }}
                  onClick={() => {
                    const allSelected = {};
                    columns.forEach(col => allSelected[col] = true);
                    setExportColsSelection(allSelected);
                  }}
                >
                  Chọn tất cả
                </button>
                <button 
                  className="pill" 
                  style={{ flex: 1, justifyContent: 'center' }}
                  onClick={() => {
                    const noneSelected = {};
                    columns.forEach(col => noneSelected[col] = false);
                    setExportColsSelection(noneSelected);
                  }}
                >
                  Bỏ chọn tất cả
                </button>
              </div>

              <div className="checkbox-group" style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {columns.map((col, idx) => (
                  <label 
                    key={idx} 
                    className="checkbox-option"
                    style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: '8px 12px', border: '1px solid var(--border)', borderRadius: 'var(--radius-sm)', transition: 'background-color var(--transition-fast)' }}
                  >
                    <span style={{ fontSize: '0.95rem' }}>{col}</span>
                    <input 
                      type="checkbox" 
                      checked={exportColsSelection[col] || false}
                      onChange={() => setExportColsSelection(prev => ({ ...prev, [col]: !prev[col] }))}
                      style={{ width: '18px', height: '18px', accentColor: 'var(--accent)' }}
                    />
                  </label>
                ))}
              </div>
            </div>

            <div style={{ display: 'flex', gap: '8px', paddingTop: '12px', borderTop: '1px solid var(--border)' }}>
              <button className="secondary-btn" onClick={() => setShowExportConfig(false)} style={{ flex: 1, padding: '12px' }}>Hủy</button>
              <button className="primary-btn" onClick={triggerExport} style={{ flex: 1, padding: '12px' }}>
                <Download size={16} /> Tải file Excel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

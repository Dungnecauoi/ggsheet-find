import React, { useState, useEffect, useMemo } from 'react';
import Papa from 'papaparse';
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
  Check
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
  const [titleColumn, setTitleColumn] = useState('');
  const [selectedColumns, setSelectedColumns] = useState({}); // column_name -> boolean
  const [nicknameInput, setNicknameInput] = useState('');

  // Handle Theme Switching
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
  }, [theme]);

  // Sync History to LocalStorage
  useEffect(() => {
    localStorage.setItem('sheet_history', JSON.stringify(history));
  }, [history]);

  // Regex to extract Spreadsheet ID & GID
  const parseSheetUrl = (url) => {
    if (!url) return null;
    
    // Spreadsheets ID: sits between /d/ and /edit (or similar)
    const idMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!idMatch) return null;
    
    // GID: sits after gid=
    const gidMatch = url.match(/[#&?]gid=([0-9]+)/);
    
    return {
      spreadsheetId: idMatch[1],
      gid: gidMatch ? gidMatch[1] : '0'
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

    const { spreadsheetId, gid } = parsed;
    const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${gid}`;

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

  // Filter Data Logic (Memoized for performance)
  const filteredData = useMemo(() => {
    if (!searchTerm.trim()) return allData;
    
    const term = searchTerm.toLowerCase();
    
    return allData.filter(row => {
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
  }, [allData, searchTerm, searchColumn]);

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
            <div className="search-container">
              <div className="search-wrapper">
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

              <button 
                className="icon-btn" 
                onClick={() => setShowConfig(true)}
                title="Cấu hình hiển thị"
              >
                <Sliders size={18} />
              </button>
            </div>

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
                  <article key={index} className="data-card">
                    <header className="card-header">
                      <h4 className="card-title">{cardTitle}</h4>
                      <span className="card-index">#{globalIndex}</span>
                    </header>
                    
                    <div className="card-body">
                      {columns
                        .filter(col => col !== titleColumn && selectedColumns[col])
                        .map((col, colIdx) => (
                          <div key={colIdx} className="property-row">
                            <span className="property-label">{col}:</span>
                            <span className="property-value">{row[col] || <span style={{ color: 'var(--text-tertiary)' }}>—</span>}</span>
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
    </div>
  );
}

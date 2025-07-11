// スプレッドシートビューアー機能実装

// グローバル変数
let currentData = null;
let currentHeaders = null;
let currentPage = 1;
let itemsPerPage = 50;
let sortColumn = -1;
let sortDirection = 'asc';

// 初期化関数の拡張
function initializeSpreadsheetViewer() {
    // 拠点リストの読み込み
    loadLocations();
    
    // データタイプリストの読み込み
    loadDataTypes();
    
    // イベントリスナーの設定
    setupEventListeners();
    
    // 初期データの読み込み
    refreshData();
}

// イベントリスナーの設定
function setupEventListeners() {
    // 検索入力
    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
        searchInput.addEventListener('input', debounce(filterTable, 300));
    }
    
    // セレクトボックスの変更
    document.getElementById('locationSelect')?.addEventListener('change', refreshData);
    document.getElementById('deviceTypeSelect')?.addEventListener('change', filterTable);
    document.getElementById('dataTypeSelect')?.addEventListener('change', handleDataTypeChange);
    
    // ページサイズ変更
    const pageSizeSelect = document.getElementById('pageSizeSelect');
    if (pageSizeSelect) {
        pageSizeSelect.addEventListener('change', () => {
            itemsPerPage = parseInt(pageSizeSelect.value);
            currentPage = 1;
            displayCurrentPage();
        });
    }
}

// データタイプ変更ハンドラ
function handleDataTypeChange() {
    const dataTypeId = document.getElementById('dataTypeSelect').value;
    
    // サマリーデータタイプの場合は専用ビューを表示
    if (dataTypeId === 'SUMMARY') {
        showSummaryView();
    } else {
        refreshData();
    }
}

// サマリービューの表示
function showSummaryView() {
    const dataSection = document.querySelector('.data-section');
    if (!dataSection) return;
    
    dataSection.innerHTML = `
        <div class="card">
            <div class="card-body">
                <h5 class="card-title">
                    <i class="fas fa-chart-bar"></i> サマリーデータ表示
                </h5>
                <div id="summaryContent">
                    <div class="text-center p-5">
                        <div class="spinner-border text-primary" role="status">
                            <span class="visually-hidden">読み込み中...</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    // サマリーデータの読み込み
    loadSummaryData();
}

// サマリーデータの読み込み
function loadSummaryData() {
    const locationId = document.getElementById('locationSelect').value;
    
    google.script.run.withSuccessHandler(displaySummaryData)
        .withFailureHandler(handleError)
        .getSummaryData(locationId);
}

// サマリーデータの表示
function displaySummaryData(data) {
    const summaryContent = document.getElementById('summaryContent');
    if (!summaryContent) return;
    
    if (!data || data.length === 0) {
        summaryContent.innerHTML = `
            <div class="alert alert-info">
                <i class="fas fa-info-circle"></i> サマリーデータがありません。
            </div>
        `;
        return;
    }
    
    // カテゴリ別カード表示
    const cards = createSummaryCards(data);
    summaryContent.innerHTML = `
        <div class="row g-4">
            ${cards}
        </div>
    `;
}

// サマリーカードの作成
function createSummaryCards(data) {
    const categories = parseSummaryData(data);
    let cardsHtml = '';
    
    categories.forEach(category => {
        const iconClass = getSummaryIconClass(category.name);
        const colorClass = getSummaryColorClass(category.name);
        
        cardsHtml += `
            <div class="col-md-6 col-lg-4">
                <div class="card h-100 border-${colorClass}">
                    <div class="card-header bg-${colorClass} text-white">
                        <h6 class="mb-0">
                            <i class="${iconClass}"></i> ${category.name}
                        </h6>
                    </div>
                    <div class="card-body">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>拠点</th>
                                    ${category.deviceTypes.map(type => `<th>${type}</th>`).join('')}
                                    <th>合計</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${category.locations.map(loc => `
                                    <tr>
                                        <td>${loc.name}</td>
                                        ${loc.counts.map(count => `<td>${count || '-'}</td>`).join('')}
                                        <td><strong>${loc.total}</strong></td>
                                    </tr>
                                `).join('')}
                            </tbody>
                            <tfoot>
                                <tr class="table-${colorClass}">
                                    <th>合計</th>
                                    ${category.totals.map(total => `<th>${total}</th>`).join('')}
                                    <th>${category.grandTotal}</th>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
            </div>
        `;
    });
    
    return cardsHtml;
}

// サマリーデータの解析
function parseSummaryData(data) {
    // 実装予定: データ構造に応じた解析ロジック
    return [];
}

// サマリーアイコンクラスの取得
function getSummaryIconClass(categoryName) {
    const iconMap = {
        '返却可能': 'fas fa-check-circle',
        '商談や金額の問題で返却不可': 'fas fa-exclamation-triangle',
        'お客様による返却拒否': 'fas fa-user-times',
        'HW延長保守中': 'fas fa-shield-alt'
    };
    return iconMap[categoryName] || 'fas fa-folder';
}

// サマリーカラークラスの取得
function getSummaryColorClass(categoryName) {
    const colorMap = {
        '返却可能': 'success',
        '商談や金額の問題で返却不可': 'warning',
        'お客様による返却拒否': 'danger',
        'HW延長保守中': 'info'
    };
    return colorMap[categoryName] || 'secondary';
}

// デバウンス関数
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// 拠点リストの読み込み（改善版）
function loadLocations() {
    google.script.run.withSuccessHandler(function(locations) {
        const select = document.getElementById('locationSelect');
        if (!select) return;
        
        select.innerHTML = '<option value="">全ての拠点</option>';
        
        // 管轄でグループ化
        const grouped = groupLocationsByRegion(locations);
        
        Object.entries(grouped).forEach(([region, locs]) => {
            const optgroup = document.createElement('optgroup');
            optgroup.label = region;
            
            locs.forEach(location => {
                const option = document.createElement('option');
                option.value = location.id;
                option.textContent = `${location.code} - ${location.name}`;
                optgroup.appendChild(option);
            });
            
            select.appendChild(optgroup);
        });
        
        // 前回選択した拠点を復元
        const savedLocation = localStorage.getItem('selectedLocation');
        if (savedLocation) {
            select.value = savedLocation;
        }
    }).withFailureHandler(handleError)
    .getLocationList();
}

// 拠点を管轄でグループ化
function groupLocationsByRegion(locations) {
    const grouped = {};
    locations.forEach(location => {
        const region = location.region || 'その他';
        if (!grouped[region]) {
            grouped[region] = [];
        }
        grouped[region].push(location);
    });
    return grouped;
}

// データタイプリストの読み込み（改善版）
function loadDataTypes() {
    google.script.run.withSuccessHandler(function(dataTypes) {
        const select = document.getElementById('dataTypeSelect');
        if (!select) return;
        
        select.innerHTML = '';
        
        dataTypes.forEach(dataType => {
            const option = document.createElement('option');
            option.value = dataType.id;
            option.textContent = dataType.name;
            
            // 説明をツールチップとして追加
            if (dataType.description) {
                option.title = dataType.description;
            }
            
            if (dataType.id === 'NORMAL') {
                option.selected = true;
            }
            
            select.appendChild(option);
        });
    }).withFailureHandler(handleError)
    .getDataTypeList();
}

// データの更新（改善版）
function refreshData() {
    const locationId = document.getElementById('locationSelect').value;
    const dataTypeId = document.getElementById('dataTypeSelect').value;
    
    // 選択を保存
    if (locationId) {
        localStorage.setItem('selectedLocation', locationId);
    }
    
    // ローディング表示
    showLoadingState();
    
    // データの取得
    google.script.run.withSuccessHandler(function(result) {
        currentData = result.data;
        currentHeaders = result.headers;
        currentPage = 1;
        sortColumn = -1;
        sortDirection = 'asc';
        
        displayData(result);
        updateLastUpdateTime();
    }).withFailureHandler(handleError)
    .getSpreadsheetData(locationId, dataTypeId);
}

// ローディング状態の表示
function showLoadingState() {
    const tbody = document.getElementById('tableBody');
    if (tbody) {
        tbody.innerHTML = `
            <tr>
                <td colspan="100" class="text-center p-5">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">読み込み中...</span>
                    </div>
                    <div class="mt-2 text-muted">データを読み込んでいます...</div>
                </td>
            </tr>
        `;
    }
}

// 最終更新時刻の更新
function updateLastUpdateTime() {
    const lastUpdateElement = document.getElementById('lastUpdate');
    if (lastUpdateElement) {
        const now = new Date();
        lastUpdateElement.textContent = now.toLocaleString('ja-JP', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });
    }
}

// データ表示の改善版
function displayData(result) {
    if (!result || !result.headers || !result.data) {
        showNoDataMessage();
        return;
    }
    
    currentData = result.data;
    currentHeaders = result.headers;
    
    // ヘッダーの設定
    setupTableHeaders(result.headers);
    
    // ページネーションの初期化
    initializePagination();
    
    // 最初のページを表示
    displayCurrentPage();
}

// データなしメッセージの表示
function showNoDataMessage() {
    const tbody = document.getElementById('tableBody');
    if (tbody) {
        tbody.innerHTML = `
            <tr>
                <td colspan="100" class="text-center p-5 text-muted">
                    <i class="fas fa-inbox fa-3x mb-3"></i>
                    <div>データがありません</div>
                </td>
            </tr>
        `;
    }
}

// テーブルヘッダーの設定
function setupTableHeaders(headers) {
    const headerRow = document.getElementById('tableHeader');
    if (!headerRow) return;
    
    headerRow.innerHTML = `
        <th width="40">
            <input type="checkbox" id="selectAllCheckbox" onchange="toggleSelectAll()">
        </th>
    `;
    
    headers.forEach((header, index) => {
        const th = document.createElement('th');
        th.innerHTML = `
            <span>${header}</span>
            <i class="fas fa-sort ms-1"></i>
        `;
        th.style.cursor = 'pointer';
        th.onclick = () => sortTable(index);
        headerRow.appendChild(th);
    });
}

// ページネーションの初期化
function initializePagination() {
    const totalPages = Math.ceil(currentData.length / itemsPerPage);
    updatePagination(currentPage, totalPages);
}

// 現在のページを表示
function displayCurrentPage() {
    if (!currentData) return;
    
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const pageData = currentData.slice(startIndex, endIndex);
    
    displayTableData(pageData, startIndex);
    
    const totalPages = Math.ceil(currentData.length / itemsPerPage);
    updatePagination(currentPage, totalPages);
}

// テーブルデータの表示
function displayTableData(data, startIndex) {
    const tbody = document.getElementById('tableBody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    data.forEach((row, index) => {
        const tr = createTableRow(row, startIndex + index);
        tbody.appendChild(tr);
    });
}

// テーブル行の作成
function createTableRow(row, globalIndex) {
    const tr = document.createElement('tr');
    tr.dataset.rowIndex = globalIndex;
    
    // チェックボックス
    const checkTd = document.createElement('td');
    checkTd.innerHTML = `
        <input type="checkbox" class="row-checkbox" 
               data-row-index="${globalIndex}"
               onchange="updateSelectionButtons()">
    `;
    tr.appendChild(checkTd);
    
    // データセル
    row.forEach((cell, cellIndex) => {
        const td = document.createElement('td');
        const header = currentHeaders[cellIndex];
        
        // セルの内容を適切にフォーマット
        if (header === 'デバイス種別') {
            td.innerHTML = getDeviceTypeBadge(cell) + cell;
        } else if (header.includes('ステータス')) {
            td.innerHTML = getStatusBadge(cell);
        } else if (header.includes('日時') || header.includes('日付')) {
            td.textContent = formatDateTime(cell);
        } else {
            td.textContent = cell || '';
        }
        
        // 長いテキストにツールチップを追加
        if (cell && cell.length > 30) {
            td.title = cell;
            td.style.maxWidth = '200px';
            td.style.overflow = 'hidden';
            td.style.textOverflow = 'ellipsis';
            td.style.whiteSpace = 'nowrap';
        }
        
        tr.appendChild(td);
    });
    
    // 行クリックで詳細表示
    tr.style.cursor = 'pointer';
    tr.onclick = (e) => {
        if (e.target.type !== 'checkbox') {
            showRowDetails(globalIndex);
        }
    };
    
    return tr;
}

// 日時のフォーマット
function formatDateTime(dateString) {
    if (!dateString) return '';
    
    try {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        
        return date.toLocaleString('ja-JP', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch (e) {
        return dateString;
    }
}

// 行詳細の表示
function showRowDetails(rowIndex) {
    const row = currentData[rowIndex];
    if (!row) return;
    
    // 詳細モーダルを作成（後で実装）
    console.log('Show details for row:', rowIndex, row);
}

// テーブルのソート（改善版）
function sortTable(columnIndex) {
    if (!currentData) return;
    
    // ソート方向の切り替え
    if (sortColumn === columnIndex) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumn = columnIndex;
        sortDirection = 'asc';
    }
    
    // データのソート
    currentData.sort((a, b) => {
        const aVal = a[columnIndex] || '';
        const bVal = b[columnIndex] || '';
        
        // 数値として比較可能か確認
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return sortDirection === 'asc' ? aNum - bNum : bNum - aNum;
        }
        
        // 文字列として比較
        const result = aVal.toString().localeCompare(bVal.toString(), 'ja');
        return sortDirection === 'asc' ? result : -result;
    });
    
    // ソートアイコンの更新
    updateSortIcons(columnIndex);
    
    // 現在のページを再表示
    displayCurrentPage();
}

// ソートアイコンの更新
function updateSortIcons(activeColumn) {
    const headers = document.querySelectorAll('#tableHeader th');
    headers.forEach((th, index) => {
        if (index === 0) return; // チェックボックス列をスキップ
        
        const icon = th.querySelector('i');
        if (!icon) return;
        
        if (index - 1 === activeColumn) {
            icon.className = sortDirection === 'asc' ? 'fas fa-sort-up' : 'fas fa-sort-down';
            th.classList.add('sorted');
        } else {
            icon.className = 'fas fa-sort';
            th.classList.remove('sorted');
        }
    });
}

// フィルタリング（改善版）
function filterTable() {
    if (!currentData) return;
    
    const searchValue = document.getElementById('searchInput').value.toLowerCase();
    const deviceType = document.getElementById('deviceTypeSelect').value;
    
    // 元のデータから再度フィルタリング
    google.script.run.withSuccessHandler(function(result) {
        let filteredData = result.data;
        
        // 検索フィルタ
        if (searchValue) {
            filteredData = filteredData.filter(row => {
                return row.some(cell => 
                    cell && cell.toString().toLowerCase().includes(searchValue)
                );
            });
        }
        
        // デバイス種別フィルタ
        if (deviceType) {
            const deviceTypeIndex = result.headers.indexOf('デバイス種別');
            if (deviceTypeIndex >= 0) {
                filteredData = filteredData.filter(row => 
                    row[deviceTypeIndex] === deviceType
                );
            }
        }
        
        currentData = filteredData;
        currentPage = 1;
        displayCurrentPage();
        
        // フィルタ結果の表示
        showFilterResults(filteredData.length, result.data.length);
        
    }).withFailureHandler(handleError)
    .getSpreadsheetData(
        document.getElementById('locationSelect').value,
        document.getElementById('dataTypeSelect').value
    );
}

// フィルタ結果の表示
function showFilterResults(filteredCount, totalCount) {
    const filterInfo = document.getElementById('filterInfo');
    if (!filterInfo) {
        const controlsSection = document.querySelector('.controls-section .card-body');
        if (controlsSection) {
            const div = document.createElement('div');
            div.id = 'filterInfo';
            div.className = 'alert alert-info mt-2 mb-0';
            controlsSection.appendChild(div);
        }
    }
    
    const filterInfoElement = document.getElementById('filterInfo');
    if (filterInfoElement) {
        if (filteredCount < totalCount) {
            filterInfoElement.innerHTML = `
                <i class="fas fa-filter"></i> 
                フィルタ結果: ${filteredCount}件 / 全${totalCount}件
                <button class="btn btn-sm btn-link" onclick="clearFilters()">
                    フィルタをクリア
                </button>
            `;
            filterInfoElement.style.display = 'block';
        } else {
            filterInfoElement.style.display = 'none';
        }
    }
}

// フィルタのクリア
function clearFilters() {
    document.getElementById('searchInput').value = '';
    document.getElementById('deviceTypeSelect').value = '';
    refreshData();
}

// ページネーションの更新（改善版）
function updatePagination(currentPage, totalPages) {
    const pagination = document.getElementById('pagination');
    if (!pagination) return;
    
    pagination.innerHTML = '';
    
    if (totalPages <= 1) return;
    
    // ページ情報
    const pageInfo = document.createElement('li');
    pageInfo.className = 'page-item disabled';
    pageInfo.innerHTML = `
        <span class="page-link">
            ${currentPage} / ${totalPages} ページ 
            (${currentData.length}件)
        </span>
    `;
    pagination.appendChild(pageInfo);
    
    // 前へボタン
    const prevLi = createPaginationItem('前へ', currentPage - 1, currentPage === 1);
    pagination.appendChild(prevLi);
    
    // ページ番号
    const pageNumbers = getPageNumbers(currentPage, totalPages);
    pageNumbers.forEach(pageNum => {
        if (pageNum === '...') {
            const li = document.createElement('li');
            li.className = 'page-item disabled';
            li.innerHTML = '<span class="page-link">...</span>';
            pagination.appendChild(li);
        } else {
            const li = createPaginationItem(
                pageNum.toString(), 
                pageNum, 
                false, 
                pageNum === currentPage
            );
            pagination.appendChild(li);
        }
    });
    
    // 次へボタン
    const nextLi = createPaginationItem('次へ', currentPage + 1, currentPage === totalPages);
    pagination.appendChild(nextLi);
}

// ページネーションアイテムの作成
function createPaginationItem(text, page, disabled, active = false) {
    const li = document.createElement('li');
    li.className = `page-item ${disabled ? 'disabled' : ''} ${active ? 'active' : ''}`;
    
    const a = document.createElement('a');
    a.className = 'page-link';
    a.href = '#';
    a.textContent = text;
    
    if (!disabled) {
        a.onclick = (e) => {
            e.preventDefault();
            goToPage(page);
        };
    }
    
    li.appendChild(a);
    return li;
}

// 表示するページ番号の計算
function getPageNumbers(current, total) {
    const delta = 2;
    const range = [];
    const rangeWithDots = [];
    let l;

    for (let i = 1; i <= total; i++) {
        if (i === 1 || i === total || (i >= current - delta && i <= current + delta)) {
            range.push(i);
        }
    }

    range.forEach((i) => {
        if (l) {
            if (i - l === 2) {
                rangeWithDots.push(l + 1);
            } else if (i - l !== 1) {
                rangeWithDots.push('...');
            }
        }
        rangeWithDots.push(i);
        l = i;
    });

    return rangeWithDots;
}

// ページ移動
function goToPage(page) {
    if (!currentData) return;
    
    const totalPages = Math.ceil(currentData.length / itemsPerPage);
    if (page < 1 || page > totalPages) return;
    
    currentPage = page;
    displayCurrentPage();
    
    // ページトップへスクロール
    document.querySelector('.data-section')?.scrollIntoView({ behavior: 'smooth' });
}

// エクスポート機能（改善版）
function exportData() {
    const locationId = document.getElementById('locationSelect').value;
    const dataTypeId = document.getElementById('dataTypeSelect').value;
    const deviceType = document.getElementById('deviceTypeSelect').value;
    const searchValue = document.getElementById('searchInput').value;
    
    // エクスポート中の表示
    const exportBtn = event.target;
    const originalText = exportBtn.innerHTML;
    exportBtn.disabled = true;
    exportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> エクスポート中...';
    
    google.script.run.withSuccessHandler(function(result) {
        exportBtn.disabled = false;
        exportBtn.innerHTML = originalText;
        
        if (result.success) {
            // CSVダウンロード
            downloadCSV(result.data, result.filename);
            showNotification('データをエクスポートしました。', 'success');
        } else {
            showNotification('エクスポートに失敗しました: ' + result.error, 'danger');
        }
    }).withFailureHandler(function(error) {
        exportBtn.disabled = false;
        exportBtn.innerHTML = originalText;
        handleError(error);
    }).exportSpreadsheetData(locationId, dataTypeId, deviceType, searchValue);
}

// CSVダウンロード
function downloadCSV(csvContent, filename) {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    URL.revokeObjectURL(link.href);
}

// ステータス編集（改善版）
function editSelectedStatus() {
    const selectedCheckboxes = document.querySelectorAll('.row-checkbox:checked');
    if (selectedCheckboxes.length === 0) return;
    
    // 選択された行のデータを取得
    const selectedRows = Array.from(selectedCheckboxes).map(checkbox => {
        const rowIndex = parseInt(checkbox.dataset.rowIndex);
        return {
            index: rowIndex,
            data: currentData[rowIndex]
        };
    });
    
    // 編集可能な行を判定
    const editableRows = checkEditableRows(selectedRows);
    
    // モーダルの内容を設定
    setupStatusEditModal(editableRows, selectedRows);
    
    // モーダルを表示
    const modal = new bootstrap.Modal(document.getElementById('statusEditModal'));
    modal.show();
}

// 編集可能な行の判定
function checkEditableRows(selectedRows) {
    if (!currentHeaders) return [];
    
    const statusIndex = currentHeaders.indexOf('0-4.ステータス');
    const depositIndex = currentHeaders.indexOf('1-4.ユーザー機の預り有無');
    
    if (statusIndex === -1 || depositIndex === -1) return [];
    
    return selectedRows.filter(row => {
        const status = row.data[statusIndex];
        const deposit = row.data[depositIndex];
        return status === '1.貸出中' && deposit === '有り';
    });
}

// ステータス編集モーダルのセットアップ
function setupStatusEditModal(editableRows, allSelectedRows) {
    const content = document.getElementById('statusEditContent');
    if (!content) return;
    
    const allEditable = editableRows.length === allSelectedRows.length;
    
    content.innerHTML = `
        <div class="condition-check">
            <h6>設定条件チェック:</h6>
            <div class="condition-item ${allEditable ? 'condition-met' : 'condition-partial'}">
                <i class="fas ${allEditable ? 'fa-check-circle' : 'fa-exclamation-circle'}"></i>
                選択された${allSelectedRows.length}件中、${editableRows.length}件が編集可能です
            </div>
            ${!allEditable ? `
                <div class="alert alert-warning mt-2">
                    <small>編集可能な条件:</small>
                    <ul class="mb-0 small">
                        <li>0-4.ステータス: 1.貸出中</li>
                        <li>1-4.ユーザー機の預り有無: 有り</li>
                    </ul>
                </div>
            ` : ''}
        </div>
        
        ${editableRows.length > 0 ? `
            <div class="mb-3">
                <label for="newStatus" class="form-label">新しい預り機ステータス</label>
                <select class="form-select" id="newStatus">
                    <option value="">選択してください</option>
                    <option value="1.返却可能">1.返却可能</option>
                    <option value="2.商談や金額の問題で返却不可">2.商談や金額の問題で返却不可</option>
                    <option value="3.お客様による返却拒否">3.お客様による返却拒否</option>
                    <option value="4.HW延長保守中">4.HW延長保守中(OS入れ替えやサービス終了を含む)</option>
                    <option value="">(空白 - ステータスなし)</option>
                </select>
            </div>
            
            <div class="mb-3">
                <label for="changeReason" class="form-label">変更理由 <span class="text-danger">*</span></label>
                <textarea class="form-control" id="changeReason" rows="3" 
                          placeholder="変更理由を入力してください（必須）"></textarea>
            </div>
            
            <div id="editableRowsList" class="small text-muted">
                <details>
                    <summary>編集対象の詳細 (${editableRows.length}件)</summary>
                    <ul class="mt-2">
                        ${editableRows.slice(0, 10).map(row => {
                            const idIndex = currentHeaders.indexOf('拠点管理番号');
                            const id = idIndex >= 0 ? row.data[idIndex] : 'N/A';
                            return `<li>${id}</li>`;
                        }).join('')}
                        ${editableRows.length > 10 ? `<li>他 ${editableRows.length - 10}件...</li>` : ''}
                    </ul>
                </details>
            </div>
        ` : `
            <div class="alert alert-info">
                編集可能な行がありません。
            </div>
        `}
    `;
    
    // 保存ボタンの有効/無効を設定
    const saveBtn = document.querySelector('#statusEditModal .btn-primary');
    if (saveBtn) {
        saveBtn.disabled = editableRows.length === 0;
    }
}

// ステータス変更の保存
function saveStatusChanges() {
    const newStatus = document.getElementById('newStatus')?.value;
    const changeReason = document.getElementById('changeReason')?.value;
    
    if (!changeReason || !changeReason.trim()) {
        alert('変更理由を入力してください。');
        document.getElementById('changeReason').focus();
        return;
    }
    
    // 編集対象の行を取得
    const selectedCheckboxes = document.querySelectorAll('.row-checkbox:checked');
    const editableRows = checkEditableRows(
        Array.from(selectedCheckboxes).map(cb => ({
            index: parseInt(cb.dataset.rowIndex),
            data: currentData[parseInt(cb.dataset.rowIndex)]
        }))
    );
    
    if (editableRows.length === 0) return;
    
    // 保存中の表示
    const modal = bootstrap.Modal.getInstance(document.getElementById('statusEditModal'));
    const saveBtn = document.querySelector('#statusEditModal .btn-primary');
    const originalText = saveBtn.innerHTML;
    saveBtn.disabled = true;
    saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 保存中...';
    
    // ステータス更新の実行
    const updates = editableRows.map(row => {
        const idIndex = currentHeaders.indexOf('拠点管理番号');
        return {
            id: row.data[idIndex],
            newStatus: newStatus,
            reason: changeReason,
            timestamp: new Date().toISOString()
        };
    });
    
    google.script.run.withSuccessHandler(function(result) {
        saveBtn.disabled = false;
        saveBtn.innerHTML = originalText;
        
        if (result.success) {
            showNotification(
                `${result.updatedCount}件のステータスを更新しました。`, 
                'success'
            );
            modal.hide();
            refreshData();
        } else {
            showNotification('更新に失敗しました: ' + result.error, 'danger');
        }
    }).withFailureHandler(function(error) {
        saveBtn.disabled = false;
        saveBtn.innerHTML = originalText;
        handleError(error);
    }).updateDepositStatus(updates);
}

// エラーハンドリング（拡張版）
function handleError(error) {
    console.error('エラー:', error);
    
    let message = 'エラーが発生しました';
    if (error.message) {
        message += ': ' + error.message;
    }
    
    showNotification(message, 'danger');
    
    // エラーログを送信（オプション）
    if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run.logError({
            error: error.toString(),
            stack: error.stack,
            userAgent: navigator.userAgent,
            timestamp: new Date().toISOString()
        });
    }
}

// 通知表示（拡張版）
function showNotification(message, type = 'info', duration = 5000) {
    // 既存の通知をチェック
    const existingAlerts = document.querySelectorAll('.alert.notification-alert');
    if (existingAlerts.length >= 3) {
        existingAlerts[0].remove();
    }
    
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type} alert-dismissible fade show notification-alert`;
    alertDiv.style.position = 'fixed';
    alertDiv.style.top = '70px';
    alertDiv.style.right = '20px';
    alertDiv.style.zIndex = '9999';
    alertDiv.style.minWidth = '300px';
    alertDiv.style.maxWidth = '500px';
    
    const icon = getNotificationIcon(type);
    
    alertDiv.innerHTML = `
        <div class="d-flex align-items-center">
            <i class="${icon} me-2"></i>
            <div class="flex-grow-1">${message}</div>
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;
    
    document.body.appendChild(alertDiv);
    
    // アニメーション
    setTimeout(() => alertDiv.classList.add('show'), 10);
    
    // 自動削除
    if (duration > 0) {
        setTimeout(() => {
            alertDiv.classList.remove('show');
            setTimeout(() => alertDiv.remove(), 150);
        }, duration);
    }
}

// 通知アイコンの取得
function getNotificationIcon(type) {
    const icons = {
        'success': 'fas fa-check-circle',
        'danger': 'fas fa-exclamation-circle',
        'warning': 'fas fa-exclamation-triangle',
        'info': 'fas fa-info-circle'
    };
    return icons[type] || icons.info;
}

// デバイス種別バッジの取得（拡張版）
function getDeviceTypeBadge(type) {
    const badges = {
        '端末': {
            class: 'terminal',
            icon: 'fa-desktop',
            label: '端末'
        },
        'プリンタ': {
            class: 'printer',
            icon: 'fa-print',
            label: 'プリンタ'
        },
        'その他': {
            class: 'other',
            icon: 'fa-cube',
            label: 'その他'
        }
    };
    
    const badge = badges[type];
    if (!badge) return '';
    
    return `
        <span class="device-type-badge ${badge.class}">
            <i class="fas ${badge.icon}"></i> ${badge.label}
        </span>
    `;
}

// ステータスバッジの取得（拡張版）
function getStatusBadge(status) {
    if (!status) return '<span class="text-muted">-</span>';
    
    let className = 'status-badge';
    let icon = '';
    
    if (status.includes('貸出中')) {
        className += ' status-active';
        icon = 'fa-check-circle';
    } else if (status.includes('保管中') || status.includes('待機中')) {
        className += ' status-pending';
        icon = 'fa-clock';
    } else if (status.includes('故障') || status.includes('廃棄')) {
        className += ' status-inactive';
        icon = 'fa-times-circle';
    } else {
        className += ' status-default';
        icon = 'fa-circle';
    }
    
    return `
        <span class="${className}">
            <i class="fas ${icon} me-1"></i>${status}
        </span>
    `;
}

// 初期化時に呼び出される関数
if (typeof initializeSpreadsheetViewer === 'function') {
    // 既に定義されている場合は上書きしない
} else {
    window.initializeSpreadsheetViewer = initializeSpreadsheetViewer;
}
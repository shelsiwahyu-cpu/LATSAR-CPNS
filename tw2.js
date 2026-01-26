// ====== BAP MANAGEMENT SYSTEM - VERSION 2 ======

let database;
const DATABASE_NAME = 'BAPSystemDB';
const DATABASE_VERSION = 1;
const OBJECT_STORE = 'bapRecords';

let recordsList = [];
let editingId = null;
let activePage = 1;
let recordsPerPage = 10;
let formItems = [];

// ========== DATABASE OPERATIONS ==========
function initializeDatabase() {
    return new Promise((resolve, reject) => {
        const dbRequest = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
        
        dbRequest.onerror = () => reject(dbRequest.error);
        
        dbRequest.onsuccess = () => {
            database = dbRequest.result;
            resolve(database);
        };
        
        dbRequest.onupgradeneeded = (evt) => {
            database = evt.target.result;
            if (!database.objectStoreNames.contains(OBJECT_STORE)) {
                const store = database.createObjectStore(OBJECT_STORE, { 
                    keyPath: 'id', 
                    autoIncrement: true 
                });
                store.createIndex('letterNumber', 'NOS', { unique: false });
                store.createIndex('receiveDate', 'TGL', { unique: false });
            }
        };
    });
}

function saveRecord(recordData) {
    return new Promise((resolve, reject) => {
        const tx = database.transaction([OBJECT_STORE], 'readwrite');
        const store = tx.objectStore(OBJECT_STORE);
        const saveRequest = recordData.id ? store.put(recordData) : store.add(recordData);
        
        saveRequest.onsuccess = () => resolve(saveRequest.result);
        saveRequest.onerror = () => reject(saveRequest.error);
    });
}

function fetchAllRecords() {
    return new Promise((resolve, reject) => {
        const tx = database.transaction([OBJECT_STORE], 'readonly');
        const getAllRequest = tx.objectStore(OBJECT_STORE).getAll();
        
        getAllRequest.onsuccess = () => resolve(getAllRequest.result);
        getAllRequest.onerror = () => reject(getAllRequest.error);
    });
}

function removeRecord(recordId) {
    return new Promise((resolve, reject) => {
        const tx = database.transaction([OBJECT_STORE], 'readwrite');
        const deleteRequest = tx.objectStore(OBJECT_STORE).delete(recordId);
        
        deleteRequest.onsuccess = () => resolve();
        deleteRequest.onerror = () => reject(deleteRequest.error);
    });
}

// ========== UTILITY FUNCTIONS ==========
function sanitizeHTML(input) {
    const htmlMap = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(input || '').replace(/[&<>"']/g, char => htmlMap[char]);
}

function convertDateFormat(inputDate) {
    if (!inputDate) return '-';
    const dateObj = new Date(inputDate);
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return dateObj.toLocaleDateString('id-ID', options);
}

// ========== TABLE RENDERING ==========
function displayTable(filteredRecords = null) {
    const tableBody = document.querySelector('.userBarang2');
    if (!tableBody) return;

    const displayData = filteredRecords || recordsList;

    if (!Array.isArray(displayData) || displayData.length === 0) {
        const emptyMessage = recordsList.length === 0 
            ? '📦 Belum ada data yang tersimpan' 
            : '🔍 Tidak ada hasil yang ditemukan';
        tableBody.innerHTML = `<tr><td colspan="13" style="text-align:center;padding:30px;color:#888;font-size:16px;">${emptyMessage}</td></tr>`;
        updateEntriesInfo(0, 0, recordsList.length);
        return;
    }

    const indexStart = (activePage - 1) * recordsPerPage;
    const indexEnd = indexStart + recordsPerPage;
    const pageData = displayData.slice(indexStart, indexEnd);

    const tableHTML = pageData.map((record) => {
        const itemsArray = record.items || [];
        
        // Build item names list
        const itemNamesHTML = itemsArray.length > 0 
            ? itemsArray.map((item, i) => `<div>${i + 1}. ${sanitizeHTML(item.NamaBarang || '-')}</div>`).join('')
            : '<span style="color:#999;">-</span>';
        
        // Build quantities and units list
        const quantitiesHTML = itemsArray.length > 0 
            ? itemsArray.map((item, i) => {
                const qty = sanitizeHTML(item.Kuantitas || '-');
                const unit = sanitizeHTML(item.Satuan || '');
                return `<div>${i + 1}. ${qty} ${unit}`.trim() + '</div>';
              }).join('')
            : '<span style="color:#999;">-</span>';
        
        // Build documentation images
        let imagesHTML = '<span style="color:#999;">Tidak ada</span>';
        const hasImages = itemsArray.some(item => item.Image);
        
        if (hasImages) {
            imagesHTML = '<div style="display:flex;flex-wrap:wrap;gap:8px;">';
            itemsArray.forEach((item, idx) => {
                if (item.Image) {
                    imagesHTML += `
                        <div style="position:relative;width:60px;height:60px;">
                            <img src="${item.Image}" 
                                 style="width:100%;height:100%;object-fit:cover;border-radius:6px;border:2px solid #2196F3;cursor:pointer;box-shadow:0 2px 4px rgba(0,0,0,0.1);"
                                 onclick="window.open('${item.Image}', '_blank')"
                                 title="Klik untuk memperbesar">
                            <span style="position:absolute;bottom:2px;right:2px;background:#2196F3;color:white;font-size:11px;padding:2px 5px;border-radius:3px;font-weight:bold;">${idx + 1}</span>
                        </div>`;
                }
            });
            imagesHTML += '</div>';
        }

        return `
            <tr style="border-bottom:1px solid #e5e5e5;transition:background-color 0.2s;" 
                onmouseover="this.style.backgroundColor='#f5f5f5'" 
                onmouseout="this.style.backgroundColor='white'">
                <td style="padding:12px;font-weight:500;">${sanitizeHTML(record.NOS || '-')}</td>
                <td style="padding:12px;">${convertDateFormat(record.TGL)}</td>
                <td style="padding:12px;">${sanitizeHTML(record.PIHAK1 || '-')}</td>
                <td style="padding:12px;">${sanitizeHTML(record.NIP1 || '-')}</td>
                <td style="padding:12px;">${sanitizeHTML(record.JBT1 || '-')}</td>
                <td style="padding:12px;">${sanitizeHTML(record.PIHAK2 || '-')}</td>
                <td style="padding:12px;">${sanitizeHTML(record.NIP2 || '-')}</td>
                <td style="padding:12px;">${sanitizeHTML(record.JBT2 || '-')}</td>
                <td style="padding:12px;max-width:220px;">
                    <div style="max-height:160px;overflow-y:auto;font-size:0.95em;line-height:1.8;">${itemNamesHTML}</div>
                </td>
                <td style="padding:12px;max-width:160px;">
                    <div style="max-height:160px;overflow-y:auto;font-size:0.95em;line-height:1.8;text-align:center;">${quantitiesHTML}</div>
                </td>
                <td style="padding:12px;font-weight:500;color:#2196F3;">${sanitizeHTML(record.NOBAP || '-')}</td>
                <td style="padding:12px;text-align:center;">${imagesHTML}</td>
                <td style="padding:12px;text-align:center;white-space:nowrap;">
                    <button onclick="modifyRecord(${record.id})" 
                            style="padding:8px 14px;margin:3px;background:#4CAF50;color:white;border:none;border-radius:5px;cursor:pointer;transition:all 0.2s;box-shadow:0 2px 4px rgba(0,0,0,0.1);"
                            onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 4px 8px rgba(0,0,0,0.15)'"
                            onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 2px 4px rgba(0,0,0,0.1)'">
                        ✏️ Edit
                    </button>
                    <button onclick="removeRecordConfirm(${record.id})" 
                            style="padding:8px 14px;margin:3px;background:#f44336;color:white;border:none;border-radius:5px;cursor:pointer;transition:all 0.2s;box-shadow:0 2px 4px rgba(0,0,0,0.1);"
                            onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 4px 8px rgba(0,0,0,0.15)'"
                            onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 2px 4px rgba(0,0,0,0.1)'">
                        🗑️ Hapus
                    </button>
                    <button onclick="generateDocument(${record.id})" 
                            style="padding:8px 14px;margin:3px;background:#2196F3;color:white;border:none;border-radius:5px;cursor:pointer;transition:all 0.2s;box-shadow:0 2px 4px rgba(0,0,0,0.1);"
                            onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 4px 8px rgba(0,0,0,0.15)'"
                            onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 2px 4px rgba(0,0,0,0.1)'">
                        📄 Export
                    </button>
                </td>
            </tr>`;
    }).join('');

    tableBody.innerHTML = tableHTML;
    updateEntriesInfo(indexStart + 1, Math.min(indexEnd, displayData.length), displayData.length);
}

function updateEntriesInfo(start, end, total) {
    const infoElement = document.querySelector('.showEntries2');
    if (infoElement) {
        infoElement.textContent = `Menampilkan ${start} sampai ${end} dari ${total} data`;
    }
}

function buildPagination() {
    const totalPageCount = Math.max(1, Math.ceil(recordsList.length / recordsPerPage));
    const paginationContainer = document.querySelector('.pagination2');
    if (!paginationContainer) return;

    let paginationHTML = `
        <button onclick="navigatePage(${activePage - 1})" 
                ${activePage === 1 ? 'disabled' : ''}
                style="padding:8px 16px;margin:0 4px;border:1px solid #ddd;background:${activePage === 1 ? '#f5f5f5' : 'white'};border-radius:4px;cursor:${activePage === 1 ? 'not-allowed' : 'pointer'};">
            ← Prev
        </button>`;

    const maxVisibleButtons = 5;
    const halfVisible = Math.floor(maxVisibleButtons / 2);
    let startPage = Math.max(1, activePage - halfVisible);
    let endPage = Math.min(totalPageCount, startPage + maxVisibleButtons - 1);
    
    if (endPage - startPage < maxVisibleButtons - 1) {
        startPage = Math.max(1, endPage - maxVisibleButtons + 1);
    }

    for (let pageNum = startPage; pageNum <= endPage; pageNum++) {
        const isActive = pageNum === activePage;
        paginationHTML += `
            <button onclick="navigatePage(${pageNum})" 
                    class="${isActive ? 'active' : ''}"
                    style="padding:8px 16px;margin:0 4px;border:1px solid ${isActive ? '#2196F3' : '#ddd'};background:${isActive ? '#2196F3' : 'white'};color:${isActive ? 'white' : '#333'};border-radius:4px;cursor:pointer;font-weight:${isActive ? 'bold' : 'normal'};">
                ${pageNum}
            </button>`;
    }

    paginationHTML += `
        <button onclick="navigatePage(${activePage + 1})" 
                ${activePage === totalPageCount ? 'disabled' : ''}
                style="padding:8px 16px;margin:0 4px;border:1px solid #ddd;background:${activePage === totalPageCount ? '#f5f5f5' : 'white'};border-radius:4px;cursor:${activePage === totalPageCount ? 'not-allowed' : 'pointer'};">
            Next →
        </button>`;
    
    paginationContainer.innerHTML = paginationHTML;
}

function navigatePage(pageNumber) {
    const maxPages = Math.max(1, Math.ceil(recordsList.length / recordsPerPage));
    if (pageNumber < 1 || pageNumber > maxPages) return;
    
    activePage = pageNumber;
    displayTable();
    buildPagination();
}

// ========== SEARCH FUNCTIONALITY ==========
function performSearch() {
    const searchField = document.getElementById('brId2');
    if (!searchField) return;
    
    const searchQuery = searchField.value.toLowerCase().trim();
    
    if (!searchQuery) {
        activePage = 1;
        displayTable();
        buildPagination();
        return;
    }
    
    const matchedRecords = recordsList.filter(record => {
        // Search in main fields
        const mainFieldsMatch = 
            (record.NOS || '').toLowerCase().includes(searchQuery) ||
            (record.PIHAK1 || '').toLowerCase().includes(searchQuery) ||
            (record.PIHAK2 || '').toLowerCase().includes(searchQuery) ||
            (record.NIP1 || '').toLowerCase().includes(searchQuery) ||
            (record.NIP2 || '').toLowerCase().includes(searchQuery) ||
            (record.NOBAP || '').toLowerCase().includes(searchQuery);
        
        // Search in items
        const itemsMatch = (record.items || []).some(item => 
            (item.NamaBarang || '').toLowerCase().includes(searchQuery) ||
            (item.Kuantitas || '').toString().toLowerCase().includes(searchQuery) ||
            (item.Satuan || '').toLowerCase().includes(searchQuery)
        );
        
        return mainFieldsMatch || itemsMatch;
    });
    
    activePage = 1;
    displayTable(matchedRecords);
}

// ========== MODAL OPERATIONS ==========
function openAddModal() {
    editingId = null;
    
    const titleElement = document.getElementById('modalTitle');
    if (titleElement) {
        titleElement.textContent = '➕ Tambah Data BAP Baru';
    }
    
    clearFormFields();

    const modalElement = document.getElementById('modal2');
    if (modalElement) {
        modalElement.classList.add('active');
        modalElement.style.display = 'block';
    }

    setTimeout(() => {
        if (formItems.length === 0) {
            createItemRow();
        }
    }, 100);
}

function closeFormModal() {
    const modalElement = document.getElementById('modal2');
    if (modalElement) {
        modalElement.classList.remove('active');
        modalElement.style.display = 'none';
    }
    clearFormFields();
}

function clearFormFields() {
    const fieldIds = ['NOS2', 'TGL2', 'PIHAK1t2', 'NIP1t2', 'JBT1t2', 'PIHAK2t2', 'NIP2t2', 'JBTt2', 'NOBAP2'];
    
    fieldIds.forEach(fieldId => {
        const field = document.getElementById(fieldId);
        if (field) field.value = '';
    });

    const itemsArea = document.getElementById('itemsContainer');
    if (itemsArea) itemsArea.innerHTML = '';

    const addButton = document.getElementById('addItemBtn');
    if (addButton) addButton.remove();

    formItems = [];
}

// ========== ITEM MANAGEMENT ==========
function createItemRow() {
    const itemsArea = document.getElementById('itemsContainer');
    if (!itemsArea) return;

    const rowIndex = formItems.length;

    const rowElement = document.createElement('div');
    rowElement.className = 'item-row';
    rowElement.style.cssText = 'border:2px solid #ddd;padding:20px;margin-bottom:20px;border-radius:10px;background:#fafafa;box-shadow:0 2px 4px rgba(0,0,0,0.05);';
    
    rowElement.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:18px;padding-bottom:12px;border-bottom:2px solid #e0e0e0;">
            <h4 style="margin:0;color:#2196F3;font-size:18px;font-weight:600;">📦 Barang #${rowIndex + 1}</h4>
            ${rowIndex > 0 ? `
                <button type="button" 
                        onclick="deleteItemRow(${rowIndex})" 
                        style="background:#f44336;color:white;border:none;padding:6px 14px;border-radius:6px;cursor:pointer;font-weight:500;transition:all 0.2s;"
                        onmouseover="this.style.background='#d32f2f'"
                        onmouseout="this.style.background='#f44336'">
                    🗑️ Hapus Item
                </button>` : ''}
        </div>
        
        <div style="margin-bottom:15px;">
            <label style="display:block;margin-bottom:6px;font-weight:600;color:#333;">Nama Barang *</label>
            <input type="text" 
                   id="NamaBarang_${rowIndex}" 
                   placeholder="Contoh: Laptop Dell Latitude" 
                   style="width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;font-size:14px;transition:border-color 0.2s;"
                   onfocus="this.style.borderColor='#2196F3'"
                   onblur="this.style.borderColor='#ddd'">
        </div>
        
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
            <div>
                <label style="display:block;margin-bottom:6px;font-weight:600;color:#333;">Kuantitas *</label>
                <input type="number" 
                       id="Kuantitas_${rowIndex}" 
                       placeholder="10" 
                       min="1"
                       style="width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;font-size:14px;transition:border-color 0.2s;"
                       onfocus="this.style.borderColor='#2196F3'"
                       onblur="this.style.borderColor='#ddd'">
            </div>
            <div>
                <label style="display:block;margin-bottom:6px;font-weight:600;color:#333;">Satuan *</label>
                <input type="text" 
                       id="Satuan_${rowIndex}" 
                       placeholder="Unit/Buah/Pcs" 
                       style="width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;font-size:14px;transition:border-color 0.2s;"
                       onfocus="this.style.borderColor='#2196F3'"
                       onblur="this.style.borderColor='#ddd'">
            </div>
        </div>
        
        <div style="border-top:2px solid #e0e0e0;padding-top:15px;background:#fff;padding:15px;border-radius:6px;">
            <label style="display:block;margin-bottom:10px;font-weight:600;color:#2196F3;font-size:15px;">📷 Upload Dokumentasi Gambar</label>
            <input type="file" 
                   id="Image_${rowIndex}" 
                   accept="image/*" 
                   onchange="processImageUpload(${rowIndex}, event)" 
                   style="width:100%;padding:10px;border:2px dashed #2196F3;border-radius:6px;background:#f0f8ff;cursor:pointer;">
            <div id="Preview_${rowIndex}" style="display:none;position:relative;margin-top:12px;">
                <img id="Img_${rowIndex}" 
                     style="width:100%;max-width:250px;border-radius:8px;border:3px solid #2196F3;box-shadow:0 4px 8px rgba(0,0,0,0.1);">
                <button type="button" 
                        onclick="clearImage(${rowIndex})" 
                        style="position:absolute;top:8px;right:8px;background:#f44336;color:white;border:none;border-radius:50%;width:30px;height:30px;cursor:pointer;font-size:18px;font-weight:bold;box-shadow:0 2px 4px rgba(0,0,0,0.2);"
                        title="Hapus gambar">
                    ×
                </button>
            </div>
        </div>
    `;

    itemsArea.appendChild(rowElement);
    formItems.push({ Image: null });
    refreshAddButton();
}

function deleteItemRow(rowIndex) {
    if (formItems.length <= 1) {
        alert('⚠️ Peringatan: Minimal harus ada 1 item barang!');
        return;
    }

    const itemsArea = document.getElementById('itemsContainer');
    if (itemsArea.children[rowIndex]) {
        itemsArea.children[rowIndex].remove();
    }

    formItems.splice(rowIndex, 1);

    // Rebuild all rows
    const savedItems = [...formItems];
    formItems = [];
    itemsArea.innerHTML = '';
    savedItems.forEach(() => createItemRow());
}

function processImageUpload(rowIndex, uploadEvent) {
    const uploadedFile = uploadEvent.target.files[0];
    if (!uploadedFile) return;
    
    const maxFileSize = 5 * 1024 * 1024; // 5MB
    if (uploadedFile.size > maxFileSize) {
        alert('⚠️ Ukuran file terlalu besar! Maksimal 5MB');
        uploadEvent.target.value = '';
        return;
    }
    
    const fileReader = new FileReader();
    fileReader.onload = (readEvent) => {
        formItems[rowIndex].Image = readEvent.target.result;
        
        const previewArea = document.getElementById(`Preview_${rowIndex}`);
        const imageElement = document.getElementById(`Img_${rowIndex}`);
        
        if (previewArea && imageElement) {
            imageElement.src = readEvent.target.result;
            previewArea.style.display = 'block';
        }
    };
    fileReader.readAsDataURL(uploadedFile);
}

function clearImage(rowIndex) {
    formItems[rowIndex].Image = null;
    
    const fileInput = document.getElementById(`Image_${rowIndex}`);
    const previewArea = document.getElementById(`Preview_${rowIndex}`);
    
    if (fileInput) fileInput.value = '';
    if (previewArea) previewArea.style.display = 'none';
}

function refreshAddButton() {
    let addButton = document.getElementById('addItemBtn');
    
    if (!addButton) {
        addButton = document.createElement('button');
        addButton.id = 'addItemBtn';
        addButton.type = 'button';
        addButton.style.cssText = 'width:100%;padding:14px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;border:none;border-radius:10px;margin-top:20px;cursor:pointer;font-weight:bold;font-size:16px;transition:all 0.3s;box-shadow:0 4px 8px rgba(0,0,0,0.15);';
        addButton.innerHTML = '➕ Tambah Item Barang Lagi';
        addButton.onclick = createItemRow;
        
        addButton.onmouseover = function() {
            this.style.transform = 'translateY(-2px)';
            this.style.boxShadow = '0 6px 12px rgba(0,0,0,0.2)';
        };
        addButton.onmouseout = function() {
            this.style.transform = 'translateY(0)';
            this.style.boxShadow = '0 4px 8px rgba(0,0,0,0.15)';
        };

        const itemsArea = document.getElementById('itemsContainer');
        if (itemsArea && itemsArea.parentNode) {
            itemsArea.parentNode.insertBefore(addButton, itemsArea.nextSibling);
        }
    }
}

// ========== CRUD OPERATIONS ==========
async function refreshData() {
    try {
        recordsList = await fetchAllRecords();
        recordsList.sort((a, b) => (b.id || 0) - (a.id || 0));
        displayTable();
        buildPagination();
    } catch (err) {
        console.error('Error loading data:', err);
        alert('❌ Gagal memuat data dari database');
    }
}

async function submitForm() {
    const collectedItems = [];
    
    for (let i = 0; i < formItems.length; i++) {
        const itemName = document.getElementById(`NamaBarang_${i}`)?.value?.trim() || '';
        const itemQty = document.getElementById(`Kuantitas_${i}`)?.value?.trim() || '';
        const itemUnit = document.getElementById(`Satuan_${i}`)?.value?.trim() || '';

        if (!itemQty || !itemUnit) {
            alert(`⚠️ Item Barang #${i + 1}: Kuantitas dan Satuan wajib diisi!`);
            return;
        }

        collectedItems.push({
            NamaBarang: itemName,
            Kuantitas: itemQty,
            Satuan: itemUnit,
            Image: formItems[i].Image || null
        });
    }

    const recordData = {
        NOS: document.getElementById('NOS2')?.value?.trim() || '',
        TGL: document.getElementById('TGL2')?.value || '',
        PIHAK1: document.getElementById('PIHAK1t2')?.value?.trim() || '',
        NIP1: document.getElementById('NIP1t2')?.value?.trim() || '',
        JBT1: document.getElementById('JBT1t2')?.value?.trim() || '',
        PIHAK2: document.getElementById('PIHAK2t2')?.value?.trim() || '',
        NIP2: document.getElementById('NIP2t2')?.value?.trim() || '',
        JBT2: document.getElementById('JBTt2')?.value?.trim() || '',
        NOBAP: document.getElementById('NOBAP2')?.value?.trim() || '',
        items: collectedItems
    };

    if (!recordData.NOS) {
        alert('⚠️ Nomor Surat wajib diisi!');
        return;
    }
    if (!recordData.PIHAK1) {
        alert('⚠️ Nama Pihak Kesatu wajib diisi!');
        return;
    }

    try {
        if (editingId !== null) {
            recordData.id = editingId;
            await saveRecord(recordData);
            alert('✅ Data berhasil diperbarui!');
        } else {
            await saveRecord(recordData);
            alert('✅ Data baru berhasil disimpan!');
        }

        await refreshData();
        closeFormModal();
    } catch (err) {
        console.error('Save error:', err);
        alert('❌ Gagal menyimpan data ke database');
    }
}

async function modifyRecord(recordId) {
    editingId = recordId;
    const recordData = recordsList.find(rec => rec.id === recordId);
    
    if (!recordData) {
        alert('❌ Data tidak ditemukan!');
        return;
    }

    const titleElement = document.getElementById('modalTitle');
    if (titleElement) {
        titleElement.textContent = '✏️ Edit Data BAP';
    }

    clearFormFields();

    const setFieldValue = (fieldId, value) => {
        const field = document.getElementById(fieldId);
        if (field) field.value = value || '';
    };

    setFieldValue('NOS2', recordData.NOS);
    setFieldValue('TGL2', recordData.TGL);
    setFieldValue('PIHAK1t2', recordData.PIHAK1);
    setFieldValue('NIP1t2', recordData.NIP1);
    setFieldValue('JBT1t2', recordData.JBT1);
    setFieldValue('PIHAK2t2', recordData.PIHAK2);
    setFieldValue('NIP2t2', recordData.NIP2);
    setFieldValue('JBTt2', recordData.JBT2);
    setFieldValue('NOBAP2', recordData.NOBAP);

    const recordItems = recordData.items || [];
    
    setTimeout(() => {
        recordItems.forEach((item, idx) => {
            createItemRow();
            
            setTimeout(() => {
                setFieldValue(`NamaBarang_${idx}`, item.NamaBarang);
                setFieldValue(`Kuantitas_${idx}`, item.Kuantitas);
                setFieldValue(`Satuan_${idx}`, item.Satuan);
                
                if (item.Image) {
                    formItems[idx].Image = item.Image;
                    const imageElement = document.getElementById(`Img_${idx}`);
                    const previewArea = document.getElementById(`Preview_${idx}`);
                    if (imageElement && previewArea) {
                        imageElement.src = item.Image;
                        previewArea.style.display = 'block';
                    }
                }
            }, 50);
        });
    }, 100);

    const modalElement = document.getElementById('modal2');
    if (modalElement) {
        modalElement.classList.add('active');
        modalElement.style.display = 'block';
    }
}

async function removeRecordConfirm(recordId) {
    const recordData = recordsList.find(rec => rec.id === recordId);
    
    const confirmationMessage = recordData 
        ? `⚠️ Konfirmasi Penghapusan\n\nNo Surat: ${recordData.NOS}\nPihak Kesatu: ${recordData.PIHAK1}\n\nApakah Anda yakin ingin menghapus data ini?`
        : '⚠️ Apakah Anda yakin ingin menghapus data ini?';
    
    if (confirm(confirmationMessage)) {
        try {
            await removeRecord(recordId);
            await refreshData();
            alert('✅ Data berhasil dihapus!');
        } catch (err) {
            console.error('Delete error:', err);
            alert('❌ Gagal menghapus data');
        }
    }
}

// ========== DOCUMENT GENERATION ==========
async function generateDocument(recordId) {
    const recordData = recordsList.find(rec => rec.id === recordId);
    
    if (!recordData) {
        alert('❌ Data tidak ditemukan!');
        return;
    }

    try {
        if (!window.docx) {
            alert("⚠️ Library docx belum dimuat. Silakan refresh halaman dan coba lagi.");
            return;
        }

        // Helper: Convert number to Indonesian words
        function numberToWords(num) {
            const units = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
            
            if (num < 12) return units[num];
            if (num < 20) return numberToWords(num - 10) + ' belas';
            if (num < 100) return numberToWords(Math.floor(num / 10)) + ' puluh ' + numberToWords(num % 10);
            if (num < 200) return 'seratus ' + numberToWords(num - 100);
            if (num < 1000) return numberToWords(Math.floor(num / 100)) + ' ratus ' + numberToWords(num % 100);
            if (num < 2000) return 'seribu ' + numberToWords(num - 1000);
            if (num < 1000000) return numberToWords(Math.floor(num / 1000)) + ' ribu ' + numberToWords(num % 1000);
            if (num < 1000000000) return numberToWords(Math.floor(num / 1000000)) + ' juta ' + numberToWords(num % 1000000);
            
            return num.toString();
        }

        function capitalizeFirstLetter(text) {
            if (!text) return text;
            return text.charAt(0).toUpperCase() + text.slice(1);
        }

        function convertBase64ToBuffer(base64String) {
            const parts = base64String.split(',');
            const decodedData = atob(parts[1]);
            let length = decodedData.length;
            const buffer = new Uint8Array(length);
            while (length--) {
                buffer[length] = decodedData.charCodeAt(length);
            }
            return buffer;
        }

        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
                WidthType, AlignmentType, BorderStyle, ImageRun, VerticalAlign } = docx;

        const dateObj = recordData.TGL ? new Date(recordData.TGL) : new Date();
        const dayNames = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];
        const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
                           'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        
        const dayName = dayNames[dateObj.getDay()];
        const dayNum = dateObj.getDate();
        const monthName = monthNames[dateObj.getMonth()];
        const yearNum = dateObj.getFullYear();

        const dayWords = capitalizeFirstLetter(numberToWords(dayNum) || String(dayNum));
        const yearWords = capitalizeFirstLetter(numberToWords(yearNum) || String(yearNum));

        const fullBorders = {
            top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            right: { style: BorderStyle.SINGLE, size: 6, color: "000000" }
        };
        
        const noBorder = {
            top: { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.NONE },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE }
        };

        const documentContent = [];

        // Header Section
        documentContent.push(
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    insideHorizontal: { style: BorderStyle.NONE },
                    insideVertical: { style: BorderStyle.NONE }
                },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                width: { size: 20, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "[LOGO]", size: 20 }) ],
                                        alignment: AlignmentType.CENTER
                                    })
                                ],
                                borders: noBorder,
                                verticalAlign: VerticalAlign.CENTER
                            }),
                            new TableCell({
                                width: { size: 80, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "PEMERINTAH PROVINSI JAWA TIMUR", bold: true, size: 22 }) ],
                                        alignment: AlignmentType.CENTER
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "DINAS SOSIAL", bold: true, size: 22 }) ],
                                        alignment: AlignmentType.CENTER
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "UPT REHABILITASI SOSIAL BINA GRAHITA TUBAN", bold: true, size: 22 }) ],
                                        alignment: AlignmentType.CENTER
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "Jalan Teuku Umar No. 116, Tuban, Tuban Jawa Timur 63215", size: 20 }) ],
                                        alignment: AlignmentType.CENTER
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "Pos-el: uptrsbgtuban@gmail.com", size: 20 }) ],
                                        alignment: AlignmentType.CENTER
                                    })
                                ],
                                borders: noBorder
                            })
                        ]
                    })
                ]
            })
        );

        documentContent.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "_".repeat(79), size: 20 }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 300 }
            })
        );

        // Title
        documentContent.push(
            new Paragraph({
                children: [ new TextRun({ text: "BERITA ACARA SERAH TERIMA", bold: true, size: 24, underline: {} }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({ 
                children: [ new TextRun({ text: `Nomor: ${recordData.NOS || ''}`, size: 22 }) ],
                alignment: AlignmentType.CENTER, 
                spacing: { after: 300 } 
            })
        );

        documentContent.push(
            new Paragraph({
                children: [ new TextRun({ text: "Yang bertanda tangan di bawah ini :", size: 22 }) ],
                spacing: { after: 200 }
            })
        );

        // First Party Details
        documentContent.push(
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 20, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 5, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: recordData.PIHAK1 || '', size: 22 })] })],
                                borders: noBorder,
                                width: { size: 75, type: WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: recordData.NIP1 || '', size: 22 })] })],
                                borders: noBorder
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: recordData.JBT1 || "Pengurus Barang UPT RSBG Tuban", size: 22 })] })],
                                borders: noBorder
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ 
                                    children: [
                                        new TextRun({ text: "Selanjutnya disebut sebagai ", size: 22 }),
                                        new TextRun({ text: "pihak ke I.", bold: true, size: 22 })
                                    ]
                                })],
                                borders: noBorder,
                                columnSpan: 3
                            })
                        ]
                    })
                ],
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    insideHorizontal: { style: BorderStyle.NONE },
                    insideVertical: { style: BorderStyle.NONE }
                }
            })
        );

        documentContent.push(new Paragraph({ text: "", spacing: { after: 200 } }));

        // Second Party Details
        documentContent.push(
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 20, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 5, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: recordData.PIHAK2 || '', size: 22 })] })],
                                borders: noBorder,
                                width: { size: 75, type: WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: recordData.NIP2 || '', size: 22 })] })],
                                borders: noBorder
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: recordData.JBT2 || "", size: 22 })] })],
                                borders: noBorder
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ 
                                    children: [
                                        new TextRun({ text: "Selanjutnya disebut sebagai ", size: 22 }),
                                        new TextRun({ text: "pihak ke II", bold: true, size: 22 })
                                    ]
                                })],
                                borders: noBorder,
                                columnSpan: 3
                            })
                        ]
                    })
                ],
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    insideHorizontal: { style: BorderStyle.NONE },
                    insideVertical: { style: BorderStyle.NONE }
                }
            })
        );

        // Date narrative
        documentContent.push(
            new Paragraph({
                children: [
                    new TextRun({ text: `Pada hari ini `, size: 22 }),
                    new TextRun({ text: `${dayName}`, bold: true, size: 22 }),
                    new TextRun({ text: ` tanggal `, size: 22 }),
                    new TextRun({ text: `${dayWords}`, bold: true, size: 22 }),
                    new TextRun({ text: ` bulan `, size: 22 }),
                    new TextRun({ text: `${monthName}`, bold: true, size: 22 }),
                    new TextRun({ text: ` tahun `, size: 22 }),
                    new TextRun({ text: `${yearWords}`, bold: true, size: 22 }),
                    new TextRun({ text: ` pihak ke I telah menyerahkan barang kepada pihak ke II, dengan rincian sebagai berikut:`, size: 22 })
                ],
                spacing: { before: 300, after: 300 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        // Items Table
        const itemsArray = recordData.items || [];
        const itemTableRows = [];

        itemTableRows.push(
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "No.", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 10, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: fullBorders
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "NAMA BARANG", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: fullBorders
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "KUANTITAS", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: fullBorders
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "SATUAN", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: fullBorders
                    })
                ]
            })
        );

        itemsArray.forEach((item, index) => {
            itemTableRows.push(
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: String(index + 1) + ".", size: 20 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: fullBorders
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: item.NamaBarang || '', size: 20 })],
                                alignment: AlignmentType.LEFT
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: fullBorders
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: String(item.Kuantitas || ''), size: 20 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: fullBorders
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: item.Satuan || '', size: 20 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: fullBorders
                        })
                    ]
                })
            );
        });

        documentContent.push(
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: itemTableRows
            })
        );

        // Closing statement
        documentContent.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "Demikian Berita Acara Serah Terima ini dibuat untuk dapat dilaksanakan sebaik-baiknya.", size: 22 }) ],
                spacing: { before: 300, after: 800 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        // Signature section
        documentContent.push(new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
                insideHorizontal: { style: BorderStyle.NONE },
                insideVertical: { style: BorderStyle.NONE }
            },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            width: { size: 50, type: WidthType.PERCENTAGE },
                            children: [
                                new Paragraph({ 
                                    children: [ new TextRun({ text: "PIHAK KEDUA", bold: true, size: 24 }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ text: "", spacing: { after: 1200 } }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: recordData.PIHAK2 || '', bold: true, size: 24, underline: {} }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: `NIP. ${recordData.NIP2 || ''}`, size: 24 }) ],
                                    alignment: AlignmentType.CENTER 
                                })
                            ],
                            borders: noBorder
                        }),
                        new TableCell({
                            width: { size: 50, type: WidthType.PERCENTAGE },
                            children: [
                                new Paragraph({ 
                                    children: [ new TextRun({ text: "PIHAK KESATU", bold: true, size: 24 }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ text: "", spacing: { after: 1200 } }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: recordData.PIHAK1 || '', bold: true, size: 24, underline: {} }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: `NIP. ${recordData.NIP1 || ''}`, size: 24 }) ],
                                    alignment: AlignmentType.CENTER 
                                })
                            ],
                            borders: noBorder
                        })
                    ]
                })
            ]
        }));

        // Documentation page
        if (itemsArray.some(item => item.Image)) {
            documentContent.push(new Paragraph({ text: "", pageBreakBefore: true }));
            documentContent.push(new Paragraph({ 
                children: [ new TextRun({ text: "Dokumentasi", bold: true, size: 28 }) ], 
                alignment: AlignmentType.CENTER, 
                spacing: { after: 400 } 
            }));

            for (let i = 0; i < itemsArray.length; i++) {
                const item = itemsArray[i];
                
                if (item.Image && typeof item.Image === 'string' && item.Image.startsWith('data:image')) {
                    try {
                        documentContent.push(new Paragraph({ 
                            children: [ 
                                new TextRun({ 
                                    text: `Gambar ${i + 1} - ${item.NamaBarang || 'Item ' + (i + 1)}`,
                                    bold: true, 
                                    size: 22
                                }) 
                            ], 
                            alignment: AlignmentType.CENTER, 
                            spacing: { before: 300, after: 150 } 
                        }));
                        
                        const imageBuffer = convertBase64ToBuffer(item.Image);
                        
                        const imageElement = new ImageRun({
                            data: imageBuffer,
                            transformation: { 
                                width: 450,
                                height: 350 
                            }
                        });
                        
                        documentContent.push(new Paragraph({ 
                            children: [imageElement], 
                            alignment: AlignmentType.CENTER, 
                            spacing: { before: 100, after: 300 } 
                        }));
                    } catch (imageError) {
                        console.error(`❌ Error processing image ${i + 1}:`, imageError);
                    }
                }
            }
        }

        // Create final document
        const finalDocument = new Document({
            styles: {
                default: {
                    document: {
                        run: {
                            font: "Times New Roman",
                            size: 22
                        },
                        paragraph: {
                            spacing: { line: 276 }
                        }
                    }
                }
            },
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 720,
                            right: 1440,
                            bottom: 720,
                            left: 1440
                        }
                    }
                },
                children: documentContent
            }]
        });

        const docBlob = await Packer.toBlob(finalDocument);
        const sanitizedName = (recordData.NOS || `BAP_${recordId}`).replace(/[\\\/:*?"<>|]/g, '_');
        const outputFilename = `Berita Acara Penyerahan Barang.docx`;

        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(docBlob);
        downloadLink.download = outputFilename;
        downloadLink.click();
        URL.revokeObjectURL(downloadLink.href);

        console.log(`✅ Dokumen berhasil dibuat: ${outputFilename}`);

    } catch (error) {
        console.error("❌ Error generating document:", error);
        alert("❌ Gagal membuat dokumen Word. Error: " + error.message);
    }
}

// ========== INITIALIZATION ==========
async function startApplication() {
    try {
        await initializeDatabase();
        await refreshData();

        const addButton = document.getElementById('addBtn2');
        if (addButton) {
            addButton.addEventListener('click', openAddModal);
        }

        const closeButton = document.getElementById('closeModal2');
        if (closeButton) {
            closeButton.addEventListener('click', closeFormModal);
        }

        const sizeSelector = document.getElementById('table_size2');
        if (sizeSelector) {
            sizeSelector.addEventListener('change', (event) => {
                recordsPerPage = parseInt(event.target.value) || 10;
                activePage = 1;
                displayTable();
                buildPagination();
            });
        }

        const searchField = document.getElementById('brId2');
        if (searchField) {
            searchField.addEventListener('input', performSearch);
        }

        window.onclick = (event) => {
            const modalElement = document.getElementById('modal2');
            if (event.target === modalElement) {
                closeFormModal();
            }
        };

        console.log('✅ Aplikasi siap digunakan');
    } catch (error) {
        console.error('❌ Initialization failed:', error);
        alert('❌ Gagal menginisialisasi aplikasi');
    }
}

// Make functions globally accessible
window.modifyRecord = modifyRecord;
window.removeRecordConfirm = removeRecordConfirm;
window.generateDocument = generateDocument;
window.navigatePage = navigatePage;
window.deleteItemRow = deleteItemRow;
window.processImageUpload = processImageUpload;
window.clearImage = clearImage;
window.submitForm = submitForm;
window.closeFormModal = closeFormModal;
window.performSearch = performSearch;

// Start when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startApplication);
} else {
    startApplication();
}
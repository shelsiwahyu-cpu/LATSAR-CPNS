// ====== BAP GENERATOR - FIXED & OPTIMIZED VERSION ======


const DB_NAME = 'BAPGeneratorDB';
const DB_VERSION = 1;
const STORE_NAME = 'barang';

let recordsList = [];
let barangList = [];
let editIndex = null;
let currentPage = 1;
let entriesPerPage = 10;
let itemsInForm = [];

// ========== INDEXEDDB DENGAN ENHANCED ERROR HANDLING ==========
function initIndexedDB() {
    return new Promise((resolve, reject) => {
        console.log('🔄 Inisialisasi IndexedDB...');
        
        // Cek apakah browser support IndexedDB
        if (!window.indexedDB) {
            console.error('❌ Browser tidak support IndexedDB!');
            reject(new Error('IndexedDB tidak tersedia'));
            return;
        }

        const request = indexedDB.open(DB_NAME, DB_VERSION);
        
        request.onerror = (event) => {
            console.error('❌ Error membuka database:', event.target.error);
            reject(request.error);
        };
        
        request.onsuccess = (event) => {
            db = event.target.result;
            console.log('✅ Database berhasil dibuka:', DB_NAME);
            
            // Verifikasi object store
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                console.error('❌ Object store tidak ditemukan!');
                reject(new Error('Object store tidak ada'));
                return;
            }
            
            console.log('✅ Object store tersedia:', STORE_NAME);
            resolve(db);
        };
        
        request.onupgradeneeded = (event) => {
            console.log('🔄 Upgrade database...');
            db = event.target.result;
            
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                console.log('🔄 Membuat object store:', STORE_NAME);
                const objectStore = db.createObjectStore(STORE_NAME, { 
                    keyPath: 'id', 
                    autoIncrement: true 
                });
                objectStore.createIndex('NOS', 'NOS', { unique: false });
                objectStore.createIndex('TGL', 'TGL', { unique: false });
                console.log('✅ Object store dibuat');
            }
        };
        
        request.onblocked = (event) => {
            console.warn('⚠️ Database diblokir, tutup tab lain yang menggunakan database ini');
        };
    });
}

function saveToIndexedDB(data) {
    return new Promise((resolve, reject) => {
        if (!db) {
            console.error('❌ Database belum diinisialisasi!');
            reject(new Error('Database tidak tersedia'));
            return;
        }

        try {
            const transaction = db.transaction([STORE_NAME], 'readwrite');
            
            transaction.onerror = (event) => {
                console.error('❌ Transaction error:', event.target.error);
                reject(event.target.error);
            };
            
            transaction.oncomplete = () => {
                console.log('✅ Transaction complete');
            };
            
            const objectStore = transaction.objectStore(STORE_NAME);
            const request = data.id ? objectStore.put(data) : objectStore.add(data);
            
            request.onsuccess = (event) => {
                const id = event.target.result;
                console.log('✅ Data berhasil disimpan, ID:', id);
                resolve(id);
            };
            
            request.onerror = (event) => {
                console.error('❌ Error menyimpan data:', event.target.error);
                reject(event.target.error);
            };
        } catch (error) {
            console.error('❌ Exception saat menyimpan:', error);
            reject(error);
        }
    });
}

function getAllFromIndexedDB() {
    return new Promise((resolve, reject) => {
        if (!db) {
            console.error('❌ Database belum diinisialisasi!');
            reject(new Error('Database tidak tersedia'));
            return;
        }

        try {
            const transaction = db.transaction([STORE_NAME], 'readonly');
            
            transaction.onerror = (event) => {
                console.error('❌ Transaction error:', event.target.error);
                reject(event.target.error);
            };
            
            const objectStore = transaction.objectStore(STORE_NAME);
            const request = objectStore.getAll();
            
            request.onsuccess = (event) => {
                const data = event.target.result;
                console.log(`✅ Berhasil load ${data.length} data dari database`);
                resolve(data);
            };
            
            request.onerror = (event) => {
                console.error('❌ Error membaca data:', event.target.error);
                reject(event.target.error);
            };
        } catch (error) {
            console.error('❌ Exception saat membaca:', error);
            reject(error);
        }
    });
}

function deleteFromIndexedDB(id) {
    return new Promise((resolve, reject) => {
        if (!db) {
            console.error('❌ Database belum diinisialisasi!');
            reject(new Error('Database tidak tersedia'));
            return;
        }

        try {
            const transaction = db.transaction([STORE_NAME], 'readwrite');
            
            transaction.onerror = (event) => {
                console.error('❌ Transaction error:', event.target.error);
                reject(event.target.error);
            };
            
            const objectStore = transaction.objectStore(STORE_NAME);
            const request = objectStore.delete(id);
            
            request.onsuccess = () => {
                console.log('✅ Data berhasil dihapus, ID:', id);
                resolve();
            };
            
            request.onerror = (event) => {
                console.error('❌ Error menghapus data:', event.target.error);
                reject(event.target.error);
            };
        } catch (error) {
            console.error('❌ Exception saat menghapus:', error);
            reject(error);
        }
    });
}

// ========== HELPERS ==========
function escapeHtml(text) {
    const map = {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'};
    return String(text || '').replace(/[&<>"']/g, m => map[m]);
}

function formatDate(dateString) {
    if (!dateString) return '-';
    const date = new Date(dateString);
    return date.toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' });
}

// ========== RENDER TABLE ==========
function renderTable(filteredData = null) {
    const tbody = document.querySelector('.userBarang');
    if (!tbody) return;

    const data = filteredData || barangList;

    if (!Array.isArray(data) || data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="14" style="text-align: center; padding: 20px; color: gray;">${barangList.length === 0 ? '📦 Belum ada data BAP' : '🔍 Data tidak ditemukan'}</td></tr>`;
        updateShowEntries(0, 0, barangList.length);
        return;
    }

    const startIndex = (currentPage - 1) * entriesPerPage;
    const endIndex = startIndex + entriesPerPage;
    const paginatedData = data.slice(startIndex, endIndex);

    tbody.innerHTML = paginatedData.map((item) => {
        const items = item.items || [];
        
        const namaBarangDisplay = items.length > 0 
            ? items.map((i, idx) => `${idx + 1}. ${escapeHtml((i.NamaBarang || '-').toUpperCase())}`).join('<br>')
            : '-';
        
        const kuantitasSatuanDisplay = items.length > 0 
            ? items.map((i, idx) => `${idx + 1}. ${escapeHtml(i.Kuantitas || '-')} ${escapeHtml((i.Satuan || '').toUpperCase())}`.trim()).join('<br>')
            : '-';

        const merkDisplay = items.length > 0 
            ? items.map((i, idx) => `${idx + 1}. ${escapeHtml((i.Merk || '-').toUpperCase())}`).join('<br>')
            : '-';
        
        let docsDisplay = '<span style="color: #999;">-</span>';
        if (items.some(i => i.Image)) {
            docsDisplay = '<div style="display: flex; flex-wrap: wrap; gap: 5px;">';
            items.forEach((subItem, idx) => {
                if (subItem.Image) {
                    docsDisplay += `<div style="position: relative; width: 50px; height: 50px;">
                    <img src="${subItem.Image}"
                     style="width: 100%; height: 100%; object-fit: cover; border-radius: 4px; border: 2px solid #4CAF50; cursor: pointer;"
                     onclick="window.open('${subItem.Image}', '_blank')"><span style="position: absolute; bottom: 0; right: 0; 
                     background: #4CAF50; color: white; font-size: 10px; padding: 2px 4px; border-radius: 2px;">${idx + 1}</span></div>`;
                }
            });
            docsDisplay += '</div>';
        }

        return `<tr style="border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 10px;">${escapeHtml(item.NOS || '-')}</td>
            <td style="padding: 10px;">${formatDate(item.TGL || '-').toUpperCase()}</td>
            <td style="padding: 10px;">${escapeHtml((item.PIHAK1 || '-').toUpperCase())}</td>
            <td style="padding: 10px;">${escapeHtml(item.NIP1 || '-')}</td>
            <td style="padding: 10px;">${escapeHtml((item.PIHAK2 || '-').toUpperCase())}</td>
            <td style="padding: 10px;">${escapeHtml(item.NIP2 || '-')}</td>
            <td style="padding: 10px;">${escapeHtml((item.JBT2 || '-').toUpperCase())}</td>
            <td style="padding: 10px;">${escapeHtml(item.NOBAP || '-')}</td>
            <td style="padding: 10px; max-width: 200px;"><div style="max-height: 150px; overflow-y: auto; font-size: 0.9em; line-height: 1.6;">${namaBarangDisplay}</div></td>
            <td style="padding: 10px; max-width: 150px;"><div style="max-height: 150px; text-align: center; overflow-y: auto; font-size: 0.9em; line-height: 1.6;">${kuantitasSatuanDisplay}</div></td>
            <td style="padding: 10px; max-width: 150px;"><div style="max-height: 150px; overflow-y: auto; font-size: 0.9em; line-height: 1.6;">${merkDisplay}</div></td>
            <td style="padding: 10px; text-align: center;">${docsDisplay}</td>
            <td style="padding: 10px; text-align: center; white-space: nowrap;">
                <button onclick="editData(${item.id})" 
                            style="padding:8px 14px;margin:3px;background:#4CAF50;color:white;border:none;border-radius:5px;cursor:pointer;transition:all 0.2s;box-shadow:0 2px 4px rgba(0,0,0,0.1);"
                            onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 4px 8px rgba(0,0,0,0.15)'"
                            onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 2px 4px rgba(0,0,0,0.1)'">
                        ✏️ Edit
                    </button>
                    <button onclick="deleteData(${item.id})" 
                            style="padding:8px 14px;margin:3px;background:#f44336;color:white;border:none;border-radius:5px;cursor:pointer;transition:all 0.2s;box-shadow:0 2px 4px rgba(0,0,0,0.1);"
                            onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 4px 8px rgba(0,0,0,0.15)'"
                            onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 2px 4px rgba(0,0,0,0.1)'">
                        🗑️ Hapus
                    </button>
                    <button onclick="exportToWord(${item.id})" 
                            style="padding:8px 14px;margin:3px;background:#2196F3;color:white;border:none;border-radius:5px;cursor:pointer;transition:all 0.2s;box-shadow:0 2px 4px rgba(0,0,0,0.1);"
                            onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 4px 8px rgba(0,0,0,0.15)'"
                            onmouseout="this.style.transform='translateY(0)';this.style.boxShadow='0 2px 4px rgba(0,0,0,0.1)'">
                        📄 Export
                    </button>
                </td>
            </tr>`;
    }).join('');

    updateShowEntries(startIndex + 1, Math.min(endIndex, data.length), data.length);
}

function updateShowEntries(start, end, total) {
    const showEntries = document.querySelector('.showEntries');
    if (showEntries) showEntries.textContent = `Showing ${start} to ${end} of ${total} entries`;
}

function updatePagination() {
    const totalPages = Math.max(1, Math.ceil(barangList.length / entriesPerPage));
    const pagination = document.querySelector('.pagination');
    if (!pagination) return;

    let html = `<button onclick="changePage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>Prev</button>`;

    const maxButtons = 5;
    const half = Math.floor(maxButtons / 2);
    let start = Math.max(1, currentPage - half);
    let end = Math.min(totalPages, start + maxButtons - 1);
    
    if (end - start < maxButtons - 1) start = Math.max(1, end - maxButtons + 1);

    for (let i = start; i <= end; i++) {
        html += `<button onclick="changePage(${i})" class="${i === currentPage ? 'active' : ''}">${i}</button>`;
    }

    html += `<button onclick="changePage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>Next</button>`;
    pagination.innerHTML = html;
}

function changePage(page) {
    const totalPages = Math.max(1, Math.ceil(barangList.length / entriesPerPage));
    if (page < 1 || page > totalPages) return;
    currentPage = page;
    renderTable();
    updatePagination();
}

// ========== SEARCH ==========
function filterTable() {
    const searchInput = document.getElementById('brId');
    if (!searchInput) return;
    
    const searchTerm = searchInput.value.toLowerCase().trim();
    
    if (!searchTerm) {
        currentPage = 1;
        renderTable();
        updatePagination();
        return;
    }
    
    const filtered = barangList.filter(item => {
        const matchMain = (item.NOS || '').toLowerCase().includes(searchTerm) ||
            (item.PIHAK1 || '').toLowerCase().includes(searchTerm) ||
            (item.PIHAK2 || '').toLowerCase().includes(searchTerm) ||
            (item.NOBAP || '').toLowerCase().includes(searchTerm);
        
        const matchItems = (item.items || []).some(i => 
            (i.NamaBarang || '').toLowerCase().includes(searchTerm) ||
            (i.Kuantitas || '').toLowerCase().includes(searchTerm) ||
            (i.Satuan || '').toLowerCase().includes(searchTerm) ||
            (i.Merk || '').toLowerCase().includes(searchTerm)
        );
        
        return matchMain || matchItems;
    });
    
    currentPage = 1;
    renderTable(filtered);
}

// ========== MODAL ==========
function showAddModal() {
    editIndex = null;
    const modalTitle = document.getElementById('modalTitle');
    if (modalTitle) modalTitle.textContent = '➕ Add Tanda Terima';
    
    resetForm();

    const modal = document.getElementById('modal');
    if (modal) {
        modal.classList.add('active');
        modal.style.display = 'block';
    }

    setTimeout(() => {
        if (itemsInForm.length === 0) addNewItemRow();
    }, 100);
}

function closeModal() {
    const modal = document.getElementById('modal');
    if (modal) {
        modal.classList.remove('active');
        modal.style.display = 'none';
    }
    resetForm();
}

function resetForm() {
    // Sesuai dengan ID di HTML
    ['NOS', 'TGL', 'PIHAk1', 'NIP1', 'JBT1', 'PIHAK2', 'NIP2', 'JBT2', 'NOBAP'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });

    const container = document.getElementById('itemsContainer');
    if (container) container.innerHTML = '';

    const addBtn = document.getElementById('addItemBtn');
    if (addBtn) addBtn.remove();

    itemsInForm = [];
}

// ========== ITEM MANAGEMENT ==========
function addNewItemRow() {
    let container = document.getElementById('itemsContainer');
    if (!container) return;

    const itemIndex = itemsInForm.length;

    const itemDiv = document.createElement('div');
    itemDiv.className = 'item-row';
    itemDiv.style.cssText = 'border: 2px solid #e0e0e0; padding: 15px; margin-bottom: 15px; border-radius: 8px; background: #f9f9f9;';
    itemDiv.innerHTML = `
         <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:18px;padding-bottom:12px;border-bottom:2px solid #e0e0e0;">
            <h4 style="margin:0;color:#2196F3;font-size:18px;font-weight:600;">📦 Barang #${itemIndex + 1}</h4>
            ${itemIndex > 0 ? `
                <button type="button" 
                        onclick="removeItemRow(${itemIndex})" 
                        style="background:#f44336;color:white;border:none;padding:6px 14px;border-radius:6px;cursor:pointer;font-weight:500;transition:all 0.2s;"
                        onmouseover="this.style.background='#d32f2f'"
                        onmouseout="this.style.background='#f44336'">
                    🗑️ Hapus Item
                </button>` : ''}
        </div>
        
        <div style="margin-bottom:15px;">
            <label style="display:block;margin-bottom:6px;font-weight:600;color:#333;">Nama Barang *</label>
            <input type="text" 
                   id="NamaBarang_${itemIndex}" 
                   placeholder="Contoh: Laptop Dell Latitude" 
                   style="width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;font-size:14px;transition:border-color 0.2s;"
                   onfocus="this.style.borderColor='#2196F3'"
                   onblur="this.style.borderColor='#ddd'">
        </div>
        
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:15px;margin-bottom:15px;">
            <div>
                <label style="display:block;margin-bottom:6px;font-weight:600;color:#333;">Kuantitas *</label>
                <input type="number" 
                       id="Kuantitas_${itemIndex}" 
                       placeholder="10" 
                       min="1"
                       style="width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;font-size:14px;transition:border-color 0.2s;"
                       onfocus="this.style.borderColor='#2196F3'"
                       onblur="this.style.borderColor='#ddd'">
            </div>
            <div>
                <label style="display:block;margin-bottom:6px;font-weight:600;color:#333;">Satuan *</label>
                <input type="text" 
                       id="Satuan_${itemIndex}" 
                       placeholder="Unit/Buah/Pcs" 
                       style="width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;font-size:14px;transition:border-color 0.2s;"
                       onfocus="this.style.borderColor='#2196F3'"
                       onblur="this.style.borderColor='#ddd'">
            </div>
            <div>
                <label style="display:block;margin-bottom:6px;font-weight:600;color:#333;">Merk</label>
                <input type="text" 
                       id="Merk_${itemIndex}" 
                       placeholder="Merk/Brand" 
                       style="width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;font-size:14px;transition:border-color 0.2s;"
                       onfocus="this.style.borderColor='#2196F3'"
                       onblur="this.style.borderColor='#ddd'">
            </div>
        </div>
        
        <div style="border-top:2px solid #e0e0e0;padding-top:15px;background:#fff;padding:15px;border-radius:6px;">
            <label style="display:block;margin-bottom:10px;font-weight:600;color:#2196F3;font-size:15px;">📷 Upload Dokumentasi Gambar</label>
            <input type="file" 
                   id="Image_${itemIndex}" 
                   accept="image/*" 
                   onchange="processImageUpload(${itemIndex}, event)" 
                   style="width:100%;padding:10px;border:2px dashed #2196F3;border-radius:6px;background:#f0f8ff;cursor:pointer;">
            <div id="Preview_${itemIndex}" style="display:none;position:relative;margin-top:12px;">
                <img id="Img_${itemIndex}" 
                     style="width:100%;max-width:250px;border-radius:8px;border:3px solid #2196F3;box-shadow:0 4px 8px rgba(0,0,0,0.1);">
                <button type="button" 
                        onclick="clearImage(${itemIndex})" 
                        style="position:absolute;top:8px;right:8px;background:#f44336;color:white;border:none;border-radius:50%;width:30px;height:30px;cursor:pointer;font-size:18px;font-weight:bold;box-shadow:0 2px 4px rgba(0,0,0,0.2);"
                        title="Hapus gambar">
                    ×
                </button>
            </div>
        </div>
    `;

    container.appendChild(itemDiv);
    itemsInForm.push({ Image: null });
    updateAddItemButton();
}

function removeItemRow(index) {
    if (itemsInForm.length <= 1) {
        alert('⚠️ Minimal harus ada 1 item!');
        return;
    }

    const container = document.getElementById('itemsContainer');
    if (container.children[index]) container.children[index].remove();

    itemsInForm.splice(index, 1);

    // Rebuild form dengan index yang benar
    const tempItems = [...itemsInForm];
    itemsInForm = [];
    container.innerHTML = '';
    
    tempItems.forEach((item, idx) => {
        addNewItemRow();
        // Restore data
        if (item.Image) {
            itemsInForm[idx].Image = item.Image;
            setTimeout(() => {
                const img = document.getElementById(`Img_${idx}`);
                const preview = document.getElementById(`Preview_${idx}`);
                if (img && preview) {
                    img.src = item.Image;
                    preview.style.display = 'block';
                }
            }, 50);
        }
    });
}

// FIX: Ganti nama fungsi dari handleImageUpload ke processImageUpload
function processImageUpload(itemIndex, event) {
    const file = event.target.files[0];
    if (!file) return;
    
    if (file.size > 5 * 1024 * 1024) {
        alert('⚠️ Ukuran gambar maksimal 5MB');
        event.target.value = '';
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
        itemsInForm[itemIndex].Image = e.target.result;
        const preview = document.getElementById(`Preview_${itemIndex}`);
        const img = document.getElementById(`Img_${itemIndex}`);
        if (preview && img) {
            img.src = e.target.result;
            preview.style.display = 'block';
        }
    };
    reader.readAsDataURL(file);
}

// FIX: Ganti nama fungsi dari removeImage ke clearImage
function clearImage(itemIndex) {
    itemsInForm[itemIndex].Image = null;
    const input = document.getElementById(`Image_${itemIndex}`);
    const preview = document.getElementById(`Preview_${itemIndex}`);
    if (input) input.value = '';
    if (preview) preview.style.display = 'none';
}

function updateAddItemButton() {
    let addBtn = document.getElementById('addItemBtn');
    if (!addBtn) {
        addBtn = document.createElement('button');
        addBtn.id = 'addItemBtn';
        addBtn.type = 'button';
        addBtn.style.cssText = 'width: 100%; padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 8px; margin-top: 15px; cursor: pointer; font-weight: bold; transition: all 0.3s;';
        addBtn.innerHTML = '➕ Tambah Item Lagi';
        addBtn.onclick = addNewItemRow;
        
        addBtn.onmouseover = function() {
            this.style.transform = 'translateY(-2px)';
            this.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
        };
        addBtn.onmouseout = function() {
            this.style.transform = 'translateY(0)';
            this.style.boxShadow = 'none';
        };

        const container = document.getElementById('itemsContainer');
        if (container && container.parentNode) {
            container.parentNode.insertBefore(addBtn, container.nextSibling);
        }
    }
}

// ========== CRUD ==========
// ========== FUNGSI DEBUGGING ==========
async function debugDatabase() {
    console.log('=== DEBUG DATABASE ===');
    console.log('Browser support IndexedDB:', !!window.indexedDB);
    console.log('Database instance:', !!db);
    
    if (db) {
        console.log('Database name:', db.name);
        console.log('Database version:', db.version);
        console.log('Object stores:', Array.from(db.objectStoreNames));
        
        try {
            const allData = await getAllFromIndexedDB();
            console.log('Total data tersimpan:', allData.length);
            console.log('Data:', allData);
        } catch (err) {
            console.error('Error reading data:', err);
        }
    }
    
    console.log('=== END DEBUG ===');
}

// Expose ke window untuk debugging manual
window.debugDatabase = debugDatabase;

// ========== LOAD DATA DENGAN RETRY ==========
async function loadData(retryCount = 0) {
    try {
        console.log('🔄 Loading data...');
        
        if (!db) {
            console.log('⚠️ Database belum ready, mencoba init...');
            await initIndexedDB();
        }
        
        barangList = await getAllFromIndexedDB();
        barangList.sort((a, b) => (b.id || 0) - (a.id || 0));
        
        console.log(`✅ Data berhasil dimuat: ${barangList.length} items`);
        
        renderTable();
        updatePagination();
        
    } catch (error) {
        console.error('❌ Error loading data:', error);
        
        // Retry maksimal 3 kali
        if (retryCount < 3) {
            console.log(`🔄 Retry loading data (attempt ${retryCount + 1}/3)...`);
            setTimeout(() => loadData(retryCount + 1), 1000);
        } else {
            console.error('❌ Gagal memuat data setelah 3 kali percobaan');
            alert('❌ Gagal memuat data. Silakan refresh halaman atau cek console (F12) untuk error detail.');
        }
    }
}


async function saveData() {
    const items = [];
    for (let i = 0; i < itemsInForm.length; i++) {
        const namaBarang = document.getElementById(`NamaBarang_${i}`)?.value?.trim() || '';
        const kuantitas = document.getElementById(`Kuantitas_${i}`)?.value?.trim() || '';
        const satuan = document.getElementById(`Satuan_${i}`)?.value?.trim() || '';
        const merk = document.getElementById(`Merk_${i}`)?.value?.trim() || '';

        if (!kuantitas || !satuan) {
            alert(`⚠️ Item #${i + 1}: Kuantitas dan Satuan harus diisi!`);
            return;
        }

        items.push({
            NamaBarang: namaBarang,
            Kuantitas: kuantitas,
            Satuan: satuan,
            Merk: merk,
            Image: itemsInForm[i].Image || null
        });
    }

    const formData = {
        NOS: document.getElementById('NOS')?.value?.trim() || '',
        TGL: document.getElementById('TGL')?.value || '',
        PIHAK1: document.getElementById('PIHAk1')?.value?.trim() || '',
        NIP1: document.getElementById('NIP1')?.value?.trim() || '',
        JBT1: document.getElementById('JBT1')?.value?.trim() || '',
        PIHAK2: document.getElementById('PIHAK2')?.value?.trim() || '',
        NIP2: document.getElementById('NIP2')?.value?.trim() || '',
        JBT2: document.getElementById('JBT2')?.value?.trim() || '',
        NOBAP: document.getElementById('NOBAP')?.value?.trim() || '',
        items: items
    };

    if (!formData.NOS) {
        alert('⚠️ No Surat harus diisi!');
        return;
    }
    if (!formData.PIHAK1) {
        alert('⚠️ Nama Pihak Kesatu harus diisi!');
        return;
    }

    try {
        if (editIndex !== null) {
            formData.id = editIndex;
            await saveToIndexedDB(formData);
            alert('✅ Data berhasil diperbarui!');
        } else {
            await saveToIndexedDB(formData);
            alert('✅ Data berhasil disimpan!');
        }

        await loadData();
        closeModal();
    } catch (error) {
        console.error('Error saving data:', error);
        alert('❌ Gagal menyimpan data');
    }
}

async function editData(id) {
    editIndex = id;
    const item = barangList.find(b => b.id === id);
    if (!item) {
        alert('❌ Data tidak ditemukan!');
        return;
    }

    const modalTitle = document.getElementById('modalTitle');
    if (modalTitle) modalTitle.textContent = '✏️ Edit Tanda Terima';

    resetForm();

    const setValue = (fieldId, value) => {
        const el = document.getElementById(fieldId);
        if (el) el.value = value || '';
    };

    setValue('NOS', item.NOS);
    setValue('TGL', item.TGL);
    setValue('PIHAk1', item.PIHAK1);
    setValue('NIP1', item.NIP1);
    setValue('JBT1', item.JBT1);
    setValue('PIHAK2', item.PIHAK2);
    setValue('NIP2', item.NIP2);
    setValue('JBT2', item.JBT2);
    setValue('NOBAP', item.NOBAP);

    const items = item.items || [];
    
    setTimeout(() => {
        items.forEach((subItem, index) => {
            addNewItemRow();
            
            setTimeout(() => {
                setValue(`NamaBarang_${index}`, subItem.NamaBarang);
                setValue(`Kuantitas_${index}`, subItem.Kuantitas);
                setValue(`Satuan_${index}`, subItem.Satuan);
                setValue(`Merk_${index}`, subItem.Merk);
                
                if (subItem.Image) {
                    itemsInForm[index].Image = subItem.Image;
                    const img = document.getElementById(`Img_${index}`);
                    const preview = document.getElementById(`Preview_${index}`);
                    if (img && preview) {
                        img.src = subItem.Image;
                        preview.style.display = 'block';
                    }
                }
            }, 50);
        });
    }, 100);

    const modal = document.getElementById('modal');
    if (modal) {
        modal.classList.add('active');
        modal.style.display = 'block';
    }
}

async function deleteData(id) {
    const item = barangList.find(b => b.id === id);
    const confirmMsg = item ? `⚠️ Hapus data:\n\nNo Surat: ${item.NOS}\nPihak Kesatu: ${item.PIHAK1}\n\nLanjutkan?` : '⚠️ Hapus data ini?';
    
    if (confirm(confirmMsg)) {
        try {
            await deleteFromIndexedDB(id);
            await loadData();
            alert('✅ Data berhasil dihapus!');
        } catch (error) {
            console.error('Error deleting data:', error);
            alert('❌ Gagal menghapus data');
        }
    }
}

// ========== EXPORT TO WORD ==========

async function exportToWord(id) {
    const item = barangList.find(b => b.id === id);
    if (!item) {
        alert('❌ Data tidak ditemukan!');
        return;
    }

    try {
        if (!window.docx) {
            alert("⚠️ Library docx belum dimuat. Refresh halaman dan coba lagi.");
            return;
        }

        // Helper functions
        function terbilang(angka) {
            const bilangan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
            
            if (angka < 12) return bilangan[angka];
            if (angka < 20) return terbilang(angka - 10) + ' belas';
            if (angka < 100) return terbilang(Math.floor(angka / 10)) + ' puluh ' + terbilang(angka % 10);
            if (angka < 200) return 'seratus ' + terbilang(angka - 100);
            if (angka < 1000) return terbilang(Math.floor(angka / 100)) + ' ratus ' + terbilang(angka % 100);
            if (angka < 2000) return 'seribu ' + terbilang(angka - 1000);
            if (angka < 1000000) return terbilang(Math.floor(angka / 1000)) + ' ribu ' + terbilang(angka % 1000);
            if (angka < 1000000000) return terbilang(Math.floor(angka / 1000000)) + ' juta ' + terbilang(angka % 1000000);
            
            return angka.toString();
        }

        function capitalizeFirst(str) {
            if (!str) return str;
            return str.charAt(0).toUpperCase() + str.slice(1);
        }

        function dataURLtoArrayBuffer(dataurl) {
            const arr = dataurl.split(',');
            const bstr = atob(arr[1]);
            let n = bstr.length;
            const u8arr = new Uint8Array(n);
            while (n--) {
                u8arr[n] = bstr.charCodeAt(n);
            }
            return u8arr;
        }
        // Helper function untuk Title Case dengan pengecualian akronim
        function toTitleCaseWithAcronym(str) {
            if (!str) return str;
            
            // Daftar akronim yang harus tetap uppercase
            const acronyms = ['UPT', 'RSBG', 'NIP'];
            
            return str
        .toLowerCase()
        .split(' ')
        .map(word => {
            // Cek apakah kata adalah akronim
            if (acronyms.includes(word.toUpperCase())) {
                return word.toUpperCase();
            }
            // Jika bukan, kapitalisasi huruf pertama saja
            return word.charAt(0).toUpperCase() + word.slice(1);
        })
        .join(' ');
        }

        // Helper function untuk Title Case (Huruf Awal Besar)
        function toTitleCase(str) {
         if (!str) return str;
        return str
        .toLowerCase()
        .split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
}
        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
                WidthType, AlignmentType, BorderStyle, ImageRun, VerticalAlign } = docx;

       const tanggal = item.TGL ? new Date(item.TGL) : new Date();
        const hariNama = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];
        const bulanNama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        const hari = hariNama[tanggal.getDay()];
        const tgl = String(tanggal.getDate()).padStart(2, '0'); // Format 2 digit
        const bulan = bulanNama[tanggal.getMonth()];
        const tahun = tanggal.getFullYear();

        const tglTerbilang = capitalizeFirst(terbilang(tgl) || String(tgl));
        const tahunTerbilang = capitalizeFirst(terbilang(tahun) || String(tahun));

        const borders = {
            top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
            right: { style: BorderStyle.SINGLE, size: 6, color: "000000" }
        };
        const noBorders = {
            top: { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.NONE },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE }
        };

        const children = [];

        // ============================================
        // BAGIAN LOGO - LOAD IMAGE
        // ============================================
        let logoImageRun = null;
        
        try {
            console.log('🔄 Mencoba load logo...');
            const logoResponse = await fetch('LOGO PEMPROV.png');
            
            if (logoResponse.ok) {
                console.log('✅ Logo response OK');
                const logoBlob = await logoResponse.blob();
                const logoBuffer = await logoBlob.arrayBuffer();
                
                logoImageRun = new ImageRun({
                    data: logoBuffer,
                    transformation: { 
                        width: 90,
                        height: 100 
                    }
                });
                
                console.log('✅ Logo berhasil dimuat!');
            } else {
                console.log('⚠️ Logo response tidak OK, status:', logoResponse.status);
            }
        } catch (err) {
            console.error('❌ Error saat load logo:', err);
            console.log('⚠️ Logo tidak akan ditampilkan, menggunakan placeholder');
        }

        // ============================================
            // Logo dan Header dalam satu tabel
        children.push(
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
                        height: { value: 1200, rule: 'atLeast' },
                        children: [
                            // CELL LOGO
                            new TableCell({
                                width: { size: 15, type: WidthType.PERCENTAGE },
                                children: logoImageRun ? [
                                    new Paragraph({ 
                                        children: [logoImageRun],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { before: 100, after: 100 }
                                    })
                                ] : [
                                    new Paragraph({ 
                                        children: [ 
                                            new TextRun({ 
                                                text: "LOGO", 
                                                bold: true,
                                                size: 16,
                                                color: "CCCCCC"
                                            }) 
                                        ],
                                        alignment: AlignmentType.CENTER
                                    })
                                ],
                                borders: noBorders,
                                verticalAlign: VerticalAlign.CENTER,
                                margins: {
                                    top: 100,
                                    bottom: 80,
                                    left: 80,
                                    right: 90
                                }
                            }),
                            
                            // CELL HEADER TEXT
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
                                     // EMAIL
                                    new Paragraph({ 
                                        children: [ 
                                            new TextRun({ 
                                                text: "Pos-el: uptrsbgtuban@gmail.com", 
                                                size: 20 
                                            }) 
                                        ],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 0 }  // Tidak ada spacing setelah email
                                    })
                                ],
                                borders: noBorders,
                                verticalAlign: VerticalAlign.CENTER,
                                margins: {
                                    top: 100,
                                    bottom: 20,   // Dikurangi untuk garis lebih dekat
                                    left: 50,
                                    right: 50
                                }
                            })
                        ]
                    })
                ]
            })
        );
         children.push(
            new Paragraph({ 
                text: "",
                border: {
                    bottom: {
                        color: "000000",
                        space: 0,
                        style: BorderStyle.SINGLE,
                        size: 20   // Ketebalan garis
                    }
                },
                spacing: { 
                    before: 0,     // Garis langsung di bawah header (tidak ada gap)
                    after: 300     // Jarak ke judul dokumen
                }
            })
        );
        // Judul
        children.push(
            new Paragraph({
                children: [ new TextRun({ text: "Tanda Terima Pendistribusian Barang", bold: true, size: 24, underline: {} }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({ 
                children: [ new TextRun({ text: `Nomor: ${item.NOS || ''}`, size: 23 }) ],
                alignment: AlignmentType.CENTER, 
                spacing: { after: 300 } 
            })
        );

        children.push(
            new Paragraph({
                children: [ new TextRun({ text: "Dasar Penyaluran Pengeluaran Barang  :", size: 23 }) ],
                spacing: { after: 200 }
            })
        );
         children.push(
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Nomor", size: 23 })] })],
                                borders: noBorders,
                                width: { size: 20, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders,
                                width: { size: 5, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: item.NOS || '', size: 23 })] })],
                                borders: noBorders,
                                width: { size: 75, type: WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Tanggal", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: `${tgl} ${bulan} ${tahun}` || '', size: 23 })] })],
                                
                                borders: noBorders
                            })
                        ]
                    })],
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

        children.push(new Paragraph({ text: "", spacing: { after: 80 } }));


        children.push(
            new Paragraph({
                children: [ new TextRun({ text: "Pihak Yang Menyerahkan :", size: 22 }) ],
                spacing: { after: 200 }
            })
        );

        // Pihak Kesatu
        children.push(
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 23 })] })],
                                borders: noBorders,
                                width: { size: 20, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders,
                                width: { size: 5, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: toTitleCase(item.PIHAK1 || ""), size: 23 })] })],
                                borders: noBorders,
                                width: { size: 75, type: WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: item.NIP1 || '', size: 23 })] })],
                                borders: noBorders
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: toTitleCaseWithAcronym(item.JBT1 || "Pengurus Barang UPT RSBG Tuban"), size: 23 })] })],
                                borders: noBorders
                            })
                        ]
                    }),
                   
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


        children.push(new Paragraph({ text: "", spacing: { after: 80 } }));


        // Pihak Kedua

        children.push(
            new Paragraph({
                children: [ new TextRun({ text: "Pihak Yang Menerima :", size: 23 }) ],
                spacing: { after: 200 }
            })
        );

        

        // Pihak Kedua
        children.push(
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 23 })] })],
                                borders: noBorders,
                                width: { size: 20, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders,
                                width: { size: 5, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: toTitleCase(item.PIHAK2 || ""), size: 23 })] })],
                                borders: noBorders,
                                width: { size: 75, type: WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: item.NIP2 || '', size: 23 })] })],
                                borders: noBorders
                            })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 23 })] })],
                                borders: noBorders
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: toTitleCase(item.JBT2 || ""), size: 23 })] })],
                                borders: noBorders
                            })
                        ]
                    }),
                   
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

    //    Pembatas sebelum tabel barang 

          children.push(new Paragraph({ text: "", spacing: { before: 100, after: 100 } }));

        // Tabel Barang dengan kolom Merk
        const items = item.items || [];
        const tableRows = [];

        tableRows.push(
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "NO.", bold: true, size: 22 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 8, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: borders
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "NAMA BARANG", bold: true, size: 22 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 40, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: borders
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "KUANTITAS", bold: true, size: 22 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 15, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: borders
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "SATUAN", bold: true, size: 22 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 15, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: borders
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "MERK", bold: true, size: 22 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 22, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: borders
                    })
                ]
            })
        );

        items.forEach((subItem, idx) => {
            tableRows.push(
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: String(idx + 1) + ".", size: 22 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: borders
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: toTitleCase(subItem.NamaBarang || ''), size: 22 })],
                                alignment: AlignmentType.LEFT
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: borders
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: String(subItem.Kuantitas || ''), size: 22 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: borders
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({text: toTitleCase(subItem.Satuan || ''), size: 22 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: borders
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: toTitleCase(subItem.Merk || ''), size: 22 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: borders
                        })
                    ]
                })
            );
        });

        children.push(
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: tableRows
            })
        );

        // Penutup
        children.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "Demikian Berita Acara Serah Terima ini dibuat untuk dapat dilaksanakan sebaik-baiknya.", size: 22 }) ],
                spacing: { before: 300, after: 800 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        // Tanda tangan
children.push(new Table({
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
                        
                        new Paragraph({ text: "", spacing: { after: 10 } }), // Spasi untuk sejajar
                        new Paragraph({ 
                            children: [ new TextRun({ text: "Pihak Yang Menerima", bold: true, size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ 
                            children: [ new TextRun({ text: toTitleCase(item.JBT2 || ''), size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ text: "", spacing: { after: 1000 } }),
                        
                        new Paragraph({ 
                            children: [ new TextRun({  text: toTitleCase(item.PIHAK2 || ''), bold: true, size: 24, underline: {} }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ 
                            children: [ new TextRun({ text: `NIP. ${item.NIP2 || ''}`, size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        })
                    ],
                    borders: noBorders
                }),
                new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                        
                        new Paragraph({ 
                        children: [new TextRun({ text: `Tuban, ${tgl} ${bulan} ${tahun}`, size: 22 })],
                        alignment: AlignmentType.CENTER
                        }),
                        new Paragraph({ 
                            children: [ new TextRun({ text: "Pihak Yang Menyerahkan", bold: true, size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ 
                            children: [ new TextRun({ text: toTitleCaseWithAcronym(item.JBT1 || 'Pengurus Barang UPT RSBG Tuban'), size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ text: "", spacing: { after: 1000 } }),
                        new Paragraph({ 
                            children: [ new TextRun({ 
                                text: toTitleCase(item.PIHAK1 || ''), 
                                bold: true, 
                                size: 24, 
                                underline: {} 
                            }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ 
                            children: [ new TextRun({ text: 'NIP. ' + (item.NIP1 || ''), size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        })
                    ],
                    borders: noBorders
                })
            ]
        })
    ]
}));
        // Halaman Dokumentasi
        if (items.some(i => i.Image)) {
            children.push(new Paragraph({ text: "", pageBreakBefore: true }));
            children.push(new Paragraph({ 
                children: [ new TextRun({ text: "Dokumentasi", bold: true, size: 28 }) ], 
                alignment: AlignmentType.CENTER, 
                spacing: { after: 400 } 
            }));

            for (let i = 0; i < items.length; i++) {
                const subItem = items[i];
                
                if (subItem.Image && typeof subItem.Image === 'string' && subItem.Image.startsWith('data:image')) {
                    try {
                        children.push(new Paragraph({ 
                            children: [ 
                                new TextRun({ 
                                    text: `Gambar ${i + 1} - ${subItem.NamaBarang || 'Item ' + (i + 1)}`,
                                    bold: true, 
                                    size: 22
                                }) 
                            ], 
                            alignment: AlignmentType.CENTER, 
                            spacing: { before: 300, after: 150 } 
                        }));
                        
                        const imageBuffer = dataURLtoArrayBuffer(subItem.Image);
                        
                        const imageRun = new ImageRun({
                            data: imageBuffer,
                            transformation: { 
                                width: 450,
                                height: 350 
                            }
                        });
                        
                        children.push(new Paragraph({ 
                            children: [imageRun], 
                            alignment: AlignmentType.CENTER, 
                            spacing: { before: 100, after: 300 } 
                        }));
                    } catch (errImg) {
                        console.error(`❌ Error gambar ${i + 1}:`, errImg);
                    }
                }
            }
        }

        // Create Document
        const doc = new Document({
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
                children
            }]
        });

        const blob = await Packer.toBlob(doc);
        const safeName = (item.NOS || `BAP_${id}`).replace(/[\\\/:*?"<>|]/g, '_');
        const fileName = `Tanda_Terima_Pendistribusian_Barang_${safeName}.docx`;

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        link.click();
        URL.revokeObjectURL(link.href);

        console.log(`✅ File Word berhasil dibuat: ${fileName}`);

    } catch (err) {
        console.error("❌ Gagal export ke Word:", err);
        alert("❌ Gagal membuat file Word. Error: " + err.message);
    }
}

// ========== INIT ==========
async function initApp() {
    try {
        console.log('=== INISIALISASI APLIKASI ===');
        
        // Step 1: Init Database
        console.log('Step 1: Inisialisasi database...');
        await initIndexedDB();
        
        // Step 2: Load Data
        console.log('Step 2: Load data...');
        await loadData();
        
        // Step 3: Setup Event Listeners
        console.log('Step 3: Setup event listeners...');
        
        const addBtn = document.getElementById('addBtn');
        if (addBtn) {
            addBtn.addEventListener('click', showAddModal);
            console.log('✅ Add button listener ready');
        } else {
            console.warn('⚠️ Add button tidak ditemukan');
        }

        const closeBtn = document.getElementById('closeModal');
        if (closeBtn) {
            closeBtn.addEventListener('click', closeModal);
            console.log('✅ Close button listener ready');
        }

        const tableSize = document.getElementById('table_size');
        if (tableSize) {
            tableSize.addEventListener('change', (e) => {
                entriesPerPage = parseInt(e.target.value) || 10;
                currentPage = 1;
                renderTable();
                updatePagination();
            });
            console.log('✅ Table size listener ready');
        }

        const searchInput = document.getElementById('brId');
        if (searchInput) {
            searchInput.addEventListener('input', filterTable);
            console.log('✅ Search listener ready');
        }

        window.onclick = (event) => {
            const modal = document.getElementById('modal');
            if (event.target === modal) closeModal();
        };
        
        console.log('✅ App initialized successfully');
        console.log('=== INISIALISASI SELESAI ===');
        
        // Debug info
        setTimeout(() => {
            console.log(`📊 Status: ${barangList.length} data loaded`);
        }, 500);
        
    } catch (error) {
        console.error('❌ Initialization failed:', error);
        alert('❌ Gagal menginisialisasi aplikasi. Error: ' + error.message + '\n\nSilakan:\n1. Refresh halaman\n2. Buka Console (F12) untuk detail error\n3. Pastikan browser tidak dalam mode incognito');
    }
}

// ========== CEK BROWSER COMPATIBILITY ==========
function checkBrowserCompatibility() {
    console.log('=== CEK BROWSER COMPATIBILITY ===');
    
    const checks = {
        'IndexedDB': !!window.indexedDB,
        'Promise': typeof Promise !== 'undefined',
        'Async/Await': true, // Jika script ini jalan, berarti support
        'FileReader': typeof FileReader !== 'undefined',
        'Blob': typeof Blob !== 'undefined'
    };
    
    console.table(checks);
    
    const unsupported = Object.entries(checks).filter(([key, value]) => !value);
    
    if (unsupported.length > 0) {
        console.error('❌ Browser tidak mendukung:', unsupported.map(([key]) => key).join(', '));
        alert('⚠️ Browser Anda tidak mendukung beberapa fitur yang diperlukan.\n\nFitur tidak didukung: ' + unsupported.map(([key]) => key).join(', ') + '\n\nSilakan gunakan browser modern seperti Chrome, Firefox, atau Edge versi terbaru.');
        return false;
    }
    
    console.log('✅ Browser kompatibel');
    return true;
}

// ========== INIT ==========
console.log('🚀 BAP Generator Starting...');

// Cek compatibility dulu
if (checkBrowserCompatibility()) {
    // Initialize app when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            console.log('📄 DOM Ready');
            initApp();
        });
    } else {
        console.log('📄 DOM Already Ready');
        initApp();
    }
} else {
    console.error('❌ Browser tidak kompatibel, aplikasi tidak akan dijalankan');
}

// Expose functions to window
window.editData = editData;
window.deleteData = deleteData;
window.exportToWord = exportToWord;
window.changePage = changePage;
window.removeItemRow = removeItemRow;
window.processImageUpload = processImageUpload;
window.clearImage = clearImage;
window.saveData = saveData;
window.closeModal = closeModal;
window.filterTable = filterTable;

// Initialize app when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}
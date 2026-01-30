// ====== SISTEM MANAJEMEN NOTA PERMINTAAN - DENGAN GRAND TOTAL SEDERHANA ======

// ----------------- Constants & DB setup -----------------
let db;
const DB_NAME = 'SistemNotaPermintaanDatabase';
const DB_VERSION = 1;
const STORE_NAME = 'dataNotaPermintaanRecords';

// ----------------- App State -----------------
let dataList = [];
let editMode = null;
let currentPage = 1;
let itemsPerPage = 10;
let isInitialized = false;
let isSaving = false;
let itemsInForm = [];

// ----------------- Helpers -----------------
function parseRupiahToFloat(value) {
    if (value == null || value === "") return 0;
    let s = String(value).trim();
    s = s.replace(/\./g, '');
    s = s.replace(/,/g, '.');
    const num = parseFloat(s.replace(/[^\d.]/g, '')) || 0;
    return num;
}

function formatRupiah(value) {
    if (value == null || value === "") return "0";
    const num = parseRupiahToFloat(value);
    return num.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function escapeHTML(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(text || '').replace(/[&<>"']/g, m => map[m]);
}

function formatTanggal(date) {
    if (!date) return '-';
    const d = new Date(date);
    const opts = { year: 'numeric', month: 'long', day: 'numeric' };
    return d.toLocaleDateString('id-ID', opts);
}

// ----------------- IndexedDB functions -----------------
function initDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        
        request.onerror = () => reject(request.error);
        
        request.onsuccess = () => {
            db = request.result;
            resolve(db);
        };
        
        request.onupgradeneeded = (event) => {
            db = event.target.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                const objectStore = db.createObjectStore(STORE_NAME, { 
                    keyPath: 'id', 
                    autoIncrement: true 
                });
                objectStore.createIndex('noSurat', 'NOS', { unique: false });
                objectStore.createIndex('tanggal', 'TGL', { unique: false });
            }
        };
    });
}

function simpanData(data) {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction([STORE_NAME], 'readwrite');
        const store = transaction.objectStore(STORE_NAME);
        const request = data.id ? store.put(data) : store.add(data);
        
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

function ambilSemuaData() {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction([STORE_NAME], 'readonly');
        const request = transaction.objectStore(STORE_NAME).getAll();
        
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

function hapusData(id) {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction([STORE_NAME], 'readwrite');
        const request = transaction.objectStore(STORE_NAME).delete(id);
        
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
    });
}

// ----------------- Multi-Item Functions -----------------
function addItemToForm() {
    const index = itemsInForm.length;
    itemsInForm.push({
        namaBarang: '',
        jumlah: 1,
        satuan: '',
        hargaPerItem: 0,
        subtotal: 0
    });
    
    renderItemsContainer();
}

function removeItemFromForm(index) {
    if (itemsInForm.length <= 1) {
        alert('⚠️ Minimal harus ada 1 item!');
        return;
    }
    
    itemsInForm.splice(index, 1);
    renderItemsContainer();
    calculateGrandTotal();
}

function renderItemsContainer() {
    const container = document.getElementById('itemsContainer');
    if (!container) return;
    
    let html = '';
    
    itemsInForm.forEach((item, index) => {
        html += `
        <div class="item-card" style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 15px; border: 2px solid #e0e0e0;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                <h4 style="margin: 0; color: #3F51B5; font-weight: bold;">📦 Item ${index + 1}</h4>
                ${itemsInForm.length > 1 ? `
                    <button type="button" onclick="removeItemFromForm(${index})" 
                            style="background: #F44336; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px;">
                        🗑️ Hapus
                    </button>
                ` : ''}
            </div>
            
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div style="grid-column: span 2;">
                    <label style="display: block; margin-bottom: 5px; font-weight: 500;">Nama Barang *</label>
                    <input type="text" 
                           id="NamaBarang_${index}" 
                           value="${escapeHTML(item.namaBarang)}"
                           onchange="updateItemData(${index})"
                           style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;">
                </div>
                
                <div>
                    <label style="display: block; margin-bottom: 5px; font-weight: 500;">Jumlah *</label>
                    <input type="number" 
                           id="Jumlah_${index}" 
                           value="${item.jumlah}"
                           min="1"
                           oninput="calculateItemTotal(${index})"
                           style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;">
                </div>

                <div>
                    <label style="display: block; margin-bottom: 5px; font-weight: 500;">Satuan *</label>
                    <input type="text" 
                           id="satuan_${index}" 
                           value="${escapeHTML(item.satuan)}"
                           onchange="updateItemData(${index})"
                           placeholder="Contoh: Pcs, Kg, Unit"
                           style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;">
                </div>
                
                <div style="grid-column: span 2;">
                    <label style="display: block; margin-bottom: 5px; font-weight: 500;">Harga Satuan (Rp) *</label>
                    <input type="number" 
                           id="HrgItem_${index}" 
                           value="${item.hargaPerItem}"
                           min="0"
                           oninput="calculateItemTotal(${index})"
                           style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px;">
                </div>
                
                <div style="grid-column: span 2;">
                    <label style="display: block; margin-bottom: 8px; font-weight: 500; color: #666;">Subtotal (Jumlah × Harga Satuan)</label>
                    <div id="subtotal_${index}" 
                         style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                                color: white; 
                                padding: 14px 20px; 
                                border-radius: 8px; 
                                font-weight: bold; 
                                text-align: center;
                                font-size: 20px;
                                box-shadow: 0 3px 10px rgba(102, 126, 234, 0.3);
                                transition: transform 0.2s;">
                        Rp ${formatRupiah(Math.round(item.subtotal))}
                    </div>
                </div>
            </div>
        </div>`;
    });
    
    // Tombol tambah item
    html += `
    <div style="text-align: center; margin: 20px 0;">
        <button type="button" onclick="addItemToForm()" 
                style="background: #4CAF50; color: white; border: none; padding: 12px 24px; border-radius: 6px; cursor: pointer; font-weight: bold; font-size: 14px; transition: all 0.2s;"
                onmouseover="this.style.transform='scale(1.05)'"
                onmouseout="this.style.transform='scale(1)'">
            ➕ Tambah Item
        </button>
    </div>`;
    
    // Grand Total
    const grandTotal = itemsInForm.reduce((sum, item) => sum + item.subtotal, 0);
    html += `
    <div style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); 
                padding: 24px; 
                border-radius: 10px; 
                margin-top: 20px;
                box-shadow: 0 6px 20px rgba(240, 147, 251, 0.4);">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <h3 style="margin: 0; color: white; font-size: 24px; font-weight: bold; text-shadow: 0 2px 4px rgba(0,0,0,0.2);">
                💰 GRAND TOTAL
            </h3>
            <div id="grandTotal" 
                 style="color: white; font-size: 32px; font-weight: bold; text-shadow: 0 2px 4px rgba(0,0,0,0.3); transition: transform 0.3s;">
                Rp ${formatRupiah(Math.round(grandTotal))}
            </div>
        </div>
    </div>`;
    
    container.innerHTML = html;
}

function updateItemData(index) {
    const namaBarang = document.getElementById(`NamaBarang_${index}`)?.value || '';
    const satuan = document.getElementById(`satuan_${index}`)?.value || '';
    
    itemsInForm[index].namaBarang = namaBarang;
    itemsInForm[index].satuan = satuan;
}

function calculateItemTotal(index) {
    const jumlah = parseInt(document.getElementById(`Jumlah_${index}`)?.value) || 0;
    const hargaPerItem = parseFloat(document.getElementById(`HrgItem_${index}`)?.value) || 0;

    // Perhitungan sederhana: Jumlah × Harga Satuan
    const subtotal = jumlah * hargaPerItem;

    // Update data in array
    itemsInForm[index].jumlah = jumlah;
    itemsInForm[index].hargaPerItem = hargaPerItem;
    itemsInForm[index].subtotal = subtotal;

    const subtotalElement = document.getElementById(`subtotal_${index}`);
    if (subtotalElement) {
        subtotalElement.textContent = 'Rp ' + formatRupiah(Math.round(subtotal));
        
        // Animasi pulse
        subtotalElement.style.transform = 'scale(1.05)';
        setTimeout(() => {
            subtotalElement.style.transform = 'scale(1)';
        }, 200);
    }

    calculateGrandTotal();
}

function calculateGrandTotal() {
    let grandTotal = 0;
    itemsInForm.forEach((item) => {
        grandTotal += item.subtotal;
    });

    const grandTotalElement = document.getElementById('grandTotal');
    if (grandTotalElement) {
        grandTotalElement.textContent = 'Rp ' + formatRupiah(Math.round(grandTotal));
        
        // Animasi pulse
        grandTotalElement.style.transform = 'scale(1.1)';
        setTimeout(() => {
            grandTotalElement.style.transform = 'scale(1)';
        }, 300);
    }
}

// ----------------- Render Table -----------------
function tampilkanTabel(filterData = null) {
    const tbody = document.querySelector('.userBarangNP');
    if (!tbody) return;

    const data = filterData || dataList;

    if (!Array.isArray(data) || data.length === 0) {
        const pesan = dataList.length === 0 
            ? '📦 Belum ada data tersimpan' 
            : '🔍 Pencarian tidak ditemukan';
        tbody.innerHTML = `<tr><td colspan="16" style="text-align:center;padding:30px;color:#888;font-size:16px;">${pesan}</td></tr>`;
        updateInfo(0, 0, dataList.length);
        return;
    }

    const start = (currentPage - 1) * itemsPerPage;
    const end = start + itemsPerPage;
    const pageData = data.slice(start, end);

    const rows = pageData.map((row) => {
        let totalHarga = 0;
        let namaBarangDisplay = '-';
        let jumlahDisplay = '-';
        let hargaSatuanDisplay = '-';
        
        if (row.ITEMS && Array.isArray(row.ITEMS) && row.ITEMS.length > 0) {
            totalHarga = row.ITEMS.reduce((sum, item) => sum + (item.subtotal || 0), 0);
            namaBarangDisplay = row.ITEMS.map(item => item.namaBarang).join(', ');
            
            // PERUBAHAN: Menampilkan jumlah + satuan
            jumlahDisplay = row.ITEMS.map(item => {
                const jumlah = item.jumlah || 0;
                const satuan = item.satuan || '';
                return satuan ? `${jumlah} ${satuan}` : jumlah;
            }).join(', ');
            
            hargaSatuanDisplay = row.ITEMS.map(item => 'Rp ' + formatRupiah(item.hargaPerItem)).join(', ');
        } else {
            const jumlah = parseFloat(row.JUMLAH) || 0;
            const harga = parseFloat(row.HARGA) || 0;
            totalHarga = jumlah * harga;
            namaBarangDisplay = escapeHTML(row.NAMABARANG || '-');
            jumlahDisplay = jumlah || '-';
            hargaSatuanDisplay = harga ? 'Rp ' + formatRupiah(harga) : '-';
        }
        
        return `
            <tr style="border-bottom:1px solid #ddd;transition:background 0.15s;" 
                onmouseover="this.style.backgroundColor='#f9f9f9'" 
                onmouseout="this.style.backgroundColor='white'">
                <td style="padding:10px;font-weight:500;color:#333;">${escapeHTML(row.NOS || '-')}</td>
                <td style="padding:10px;color:#555;">${formatTanggal(row.TGL)}</td>
                <td style="padding:10px;color:#555;">${escapeHTML(row.PIHAK1 || '-')}</td>
                <td style="padding:10px;color:#666;font-size:13px;">${escapeHTML(row.NIP1 || '-')}</td>
                <td style="padding:10px;color:#666;font-size:13px;">${escapeHTML(row.JBT1 || '-')}</td>
                <td style="padding:10px;color:#666;font-size:13px;">${escapeHTML(row.SEKSI || '-')}</td>
                <td style="padding:10px;color:#555;">${escapeHTML(row.PIHAK2 || '-')}</td>
                <td style="padding:10px;color:#666;font-size:13px;">${escapeHTML(row.NIP2 || '-')}</td>
                <td style="padding:10px;color:#666;font-size:13px;">${escapeHTML(row.JBT2 || '-')}</td>
                <td style="padding:10px;font-weight:500;color:#3F51B5;">${escapeHTML(row.NOBAP || '-')}</td>
                <td style="padding:10px;color:#555;">${namaBarangDisplay}</td>
                <td style="padding:10px;color:#555;font-style:italic;">${escapeHTML(row.PERIHAL || '-')}</td>
                <td style="padding:10px;text-align:center;font-weight:500;">${jumlahDisplay}</td>
                <td style="padding:10px;text-align:right;color:#666;">${hargaSatuanDisplay}</td>
                <td style="padding:10px;text-align:right;">
                    <div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:8px 12px;border-radius:6px;font-weight:bold;box-shadow:0 2px 4px rgba(0,0,0,0.1);">
                        Rp ${formatRupiah(totalHarga)}
                    </div>
                </td>
                <td style="padding:10px;text-align:center;white-space:nowrap;">
                    <button onclick="editData(${row.id})" 
                            style="padding:7px 12px;margin:2px;background:#4CAF50;color:white;border:none;border-radius:4px;cursor:pointer;transition:all 0.15s;font-size:13px;"
                            onmouseover="this.style.transform='scale(1.05)'"
                            onmouseout="this.style.transform='scale(1)'">
                        ✏️ Ubah
                    </button>
                    <button onclick="konfirmasiHapus(${row.id})" 
                            style="padding:7px 12px;margin:2px;background:#F44336;color:white;border:none;border-radius:4px;cursor:pointer;transition:all 0.15s;font-size:13px;"
                            onmouseover="this.style.transform='scale(1.05)'"
                            onmouseout="this.style.transform='scale(1)'">
                        🗑️ Hapus
                    </button>
                     <button onclick="konfirmasiExport(${row.id})" 
                            style="padding:7px 12px;margin:2px;background: #2196F3;color:white;border:none;border-radius:4px;cursor:pointer;transition:all 0.15s;font-size:13px;"
                            onmouseover="this.style.transform='scale(1.05)'"
                            onmouseout="this.style.transform='scale(1)'">
                        📥 Export
                    </button>
                </td>
            </tr>`;
    }).join('');

    tbody.innerHTML = rows;
    updateInfo(start + 1, Math.min(end, data.length), data.length);
}

function updateInfo(mulai, akhir, total) {
    const info = document.querySelector('.showEntriesNP');
    if (info) {
        info.textContent = `Menampilkan ${mulai} hingga ${akhir} dari ${total} entri`;
    }
}

function buatPagination() {
    const totalHalaman = Math.max(1, Math.ceil(dataList.length / itemsPerPage));
    const container = document.querySelector('.paginationNP');
    if (!container) return;

    let html = `
        <button onclick="pindahHalaman(${currentPage - 1})" 
                ${currentPage === 1 ? 'disabled' : ''}
                style="padding:7px 14px;margin:0 3px;border:1px solid #ccc;background:${currentPage === 1 ? '#e0e0e0' : 'white'};border-radius:3px;cursor:${currentPage === 1 ? 'not-allowed' : 'pointer'};">
            Prev
        </button>`;

    const maxButtons = 5;
    const half = Math.floor(maxButtons / 2);
    let startPage = Math.max(1, currentPage - half);
    let endPage = Math.min(totalHalaman, startPage + maxButtons - 1);
    
    if (endPage - startPage < maxButtons - 1) {
        startPage = Math.max(1, endPage - maxButtons + 1);
    }

    for (let i = startPage; i <= endPage; i++) {
        const aktif = i === currentPage;
        html += `
            <button onclick="pindahHalaman(${i})" 
                    class="${aktif ? 'active' : ''}"
                    style="padding:7px 14px;margin:0 3px;border:1px solid ${aktif ? '#3F51B5' : '#ccc'};background:${aktif ? '#3F51B5' : 'white'};color:${aktif ? 'white' : '#333'};border-radius:3px;cursor:pointer;font-weight:${aktif ? 'bold' : 'normal'};">
                ${i}
            </button>`;
    }

    html += `
        <button onclick="pindahHalaman(${currentPage + 1})" 
                ${currentPage === totalHalaman ? 'disabled' : ''}
                style="padding:7px 14px;margin:0 3px;border:1px solid #ccc;background:${currentPage === totalHalaman ? '#e0e0e0' : 'white'};border-radius:3px;cursor:${currentPage === totalHalaman ? 'not-allowed' : 'pointer'};">
            Next
        </button>`;
    
    container.innerHTML = html;
}

function pindahHalaman(halaman) {
    const maxHalaman = Math.max(1, Math.ceil(dataList.length / itemsPerPage));
    if (halaman < 1 || halaman > maxHalaman) return;
    
    currentPage = halaman;
    tampilkanTabel();
    buatPagination();
}

// ----------------- Search & Filter -----------------
function filterTable() {
    const input = document.getElementById('brId');
    if (!input) return;
    
    const keyword = input.value.toLowerCase().trim();
    
    if (!keyword) {
        currentPage = 1;
        tampilkanTabel();
        buatPagination();
        return;
    }
    
    const hasil = dataList.filter(row => {
        const fieldUtama = 
            (row.NOS || '').toLowerCase().includes(keyword) ||
            (row.PIHAK1 || '').toLowerCase().includes(keyword) ||
            (row.PIHAK2 || '').toLowerCase().includes(keyword) ||
            (row.NIP1 || '').toLowerCase().includes(keyword) ||
            (row.NIP2 || '').toLowerCase().includes(keyword) ||
            (row.NOBAP || '').toLowerCase().includes(keyword) ||
            (row.NAMABARANG || '').toLowerCase().includes(keyword) ||
            (row.PERIHAL || '').toLowerCase().includes(keyword) ||
            (row.SEKSI || '').toLowerCase().includes(keyword);
        
        let itemMatch = false;
        if (row.ITEMS && Array.isArray(row.ITEMS)) {
            itemMatch = row.ITEMS.some(item => 
                (item.namaBarang || '').toLowerCase().includes(keyword) ||
                (item.satuan || '').toLowerCase().includes(keyword)
            );
        }
        
        return fieldUtama || itemMatch;
    });
    
    currentPage = 1;
    tampilkanTabel(hasil);
}

// ----------------- Modal Functions -----------------
function openModal() {
    editMode = null;
    
    const judul = document.getElementById('modalTitle');
    if (judul) {
        judul.textContent = '➕ Add Nota Permintaan';
    }
    
    bersihkanForm();
    
    itemsInForm = [{
        namaBarang: '',
        jumlah: 1,
        satuan: '',
        hargaPerItem: 0,
        subtotal: 0
    }];
    
    renderItemsContainer();

    const modal = document.getElementById('modalNP');
    if (modal) {
        modal.classList.add('active');
    }
}

function closeModal() {
    const modal = document.getElementById('modalNP');
    if (modal) {
        modal.classList.remove('active');
    }
    bersihkanForm();
    itemsInForm = [];
}

function bersihkanForm() {
    const fields = ['NOS3', 'TGL3', 'PIHAKNP', 'NIP1NP', 'JBT1NP', 'SEKSINP', 'PIHAK2NP', 'NIP2NP', 'JBT2NP', 'perihalNP', 'nobapNP'];
    
    fields.forEach(id => {
        const field = document.getElementById(id);
        if (field) field.value = '';
    });
    
    const container = document.getElementById('itemsContainer');
    if (container) {
        container.innerHTML = '';
    }
}

// ----------------- CRUD Operations -----------------
async function muatUlangData() {
    try {
        dataList = await ambilSemuaData();
        dataList.sort((a, b) => (b.id || 0) - (a.id || 0));
        tampilkanTabel();
        buatPagination();
    } catch (error) {
        console.error('Error loading:', error);
        alert('❌ Gagal memuat data');
    }
}

async function saveData() {
    if (isSaving) {
        console.log('⚠️ Sedang menyimpan, harap tunggu...');
        return;
    }
    
    isSaving = true;
    
    try {
        if (itemsInForm.length === 0) {
            alert('⚠️ Minimal harus ada 1 item!');
            isSaving = false;
            return;
        }
        
        for (let i = 0; i < itemsInForm.length; i++) {
            const item = itemsInForm[i];
            if (!item.namaBarang || item.namaBarang.trim() === '') {
                alert(`⚠️ Nama Barang pada Item ${i + 1} wajib diisi!`);
                isSaving = false;
                return;
            }
            if (!item.jumlah || item.jumlah <= 0) {
                alert(`⚠️ Jumlah pada Item ${i + 1} harus lebih dari 0!`);
                isSaving = false;
                return;
            }
            if (!item.satuan || item.satuan.trim() === '') {
                alert(`⚠️ Satuan pada Item ${i + 1} wajib diisi!`);
                isSaving = false;
                return;
            }
            if (item.hargaPerItem < 0) {
                alert(`⚠️ Harga Satuan pada Item ${i + 1} tidak boleh negatif!`);
                isSaving = false;
                return;
            }
        }
        
        const grandTotal = itemsInForm.reduce((sum, item) => sum + item.subtotal, 0);
        
        const data = {
            NOS: document.getElementById('NOS3')?.value?.trim() || '',
            TGL: document.getElementById('TGL3')?.value || '',
            PIHAK1: document.getElementById('PIHAKNP')?.value?.trim() || '',
            NIP1: document.getElementById('NIP1NP')?.value?.trim() || '',
            JBT1: document.getElementById('JBT1NP')?.value?.trim() || '',
            SEKSI: document.getElementById('SEKSINP')?.value?.trim() || '',
            PIHAK2: document.getElementById('PIHAK2NP')?.value?.trim() || '',
            NIP2: document.getElementById('NIP2NP')?.value?.trim() || '',
            JBT2: document.getElementById('JBT2NP')?.value?.trim() || '',
            PERIHAL: document.getElementById('perihalNP')?.value?.trim() || '',
            NOBAP: document.getElementById('nobapNP')?.value?.trim() || '',
            ITEMS: itemsInForm.map(item => ({...item})),
            TOTALHARGA: grandTotal
        };

        if (!data.NOS) {
            alert('⚠️ No Surat wajib diisi!');
            isSaving = false;
            return;
        }
        if (!data.PIHAK1) {
            alert('⚠️ Pihak Yang Menyetujui wajib diisi!');
            isSaving = false;
            return;
        }

        if (editMode !== null) {
            data.id = editMode;
            await simpanData(data);
            alert('✅ Data berhasil diupdate!');
        } else {
            await simpanData(data);
            alert('✅ Data berhasil disimpan!');
        }

        await muatUlangData();
        closeModal();
        
    } catch (error) {
        console.error('Save error:', error);
        alert('❌ Gagal menyimpan data');
    } finally {
        isSaving = false;
    }
}

async function editData(id) {
    editMode = id;
    const data = dataList.find(r => r.id === id);
    
    if (!data) {
        alert('❌ Data tidak ditemukan!');
        return;
    }

    const judul = document.getElementById('modalTitle');
    if (judul) {
        judul.textContent = '✏️ Edit Data Nota Permintaan';
    }

    bersihkanForm();

    const setVal = (fieldId, val) => {
        const field = document.getElementById(fieldId);
        if (field) field.value = val || '';
    };

    setVal('NOS3', data.NOS);
    setVal('TGL3', data.TGL);
    setVal('PIHAKNP', data.PIHAK1);
    setVal('NIP1NP', data.NIP1);
    setVal('JBT1NP', data.JBT1);
    setVal('SEKSINP', data.SEKSI);
    setVal('PIHAK2NP', data.PIHAK2);
    setVal('NIP2NP', data.NIP2);
    setVal('JBT2NP', data.JBT2);
    setVal('perihalNP', data.PERIHAL);
    setVal('nobapNP', data.NOBAP);
    
    if (data.ITEMS && Array.isArray(data.ITEMS) && data.ITEMS.length > 0) {
        itemsInForm = data.ITEMS.map(item => ({...item}));
    } else {
        itemsInForm = [{
            namaBarang: data.NAMABARANG || '',
            jumlah: parseFloat(data.JUMLAH) || 1,
            satuan: '',
            hargaPerItem: parseFloat(data.HARGA) || 0,
            subtotal: 0
        }];
    }
    
    renderItemsContainer();

    const modal = document.getElementById('modalNP');
    if (modal) {
        modal.classList.add('active');
    }
}

async function konfirmasiHapus(id) {
    const data = dataList.find(r => r.id === id);
    
    const pesan = data 
        ? `⚠️ Yakin ingin menghapus?\n\nNo: ${data.NOS}\nPihak 1: ${data.PIHAK1}`
        : '⚠️ Hapus data ini?';
    
    if (confirm(pesan)) {
        try {
            await hapusData(id);
            await muatUlangData();
            alert('✅ Data berhasil dihapus!');
        } catch (error) {
            console.error('Delete error:', error);
            alert('❌ Gagal menghapus data');
        }
    }
}

// ----------------- Event Listeners & Init -----------------
function initEventListeners() {
    const btnAdd = document.getElementById('addBtnNP');
    if (btnAdd) {
        const newBtn = btnAdd.cloneNode(true);
        btnAdd.parentNode.replaceChild(newBtn, btnAdd);
        newBtn.addEventListener('click', openModal);
    }

    const btnClose = document.getElementById('closeModalNP');
    if (btnClose) {
        const newBtn = btnClose.cloneNode(true);
        btnClose.parentNode.replaceChild(newBtn, btnClose);
        newBtn.addEventListener('click', closeModal);
    }

    const btnSave = document.getElementById('submitBtnNP');
    if (btnSave) {
        const newBtn = btnSave.cloneNode(true);
        btnSave.parentNode.replaceChild(newBtn, btnSave);
        newBtn.addEventListener('click', saveData);
    }

    const selectSize = document.getElementById('table_sizeNP');
    if (selectSize) {
        const newSelect = selectSize.cloneNode(true);
        selectSize.parentNode.replaceChild(newSelect, selectSize);
        newSelect.addEventListener('change', (e) => {
            itemsPerPage = parseInt(e.target.value) || 10;
            currentPage = 1;
            tampilkanTabel();
            buatPagination();
        });
    }

    const btnSearch = document.querySelector('.searchNP');
    if (btnSearch) {
        const newBtn = btnSearch.cloneNode(true);
        btnSearch.parentNode.replaceChild(newBtn, btnSearch);
        newBtn.addEventListener('click', filterTable);
    }

    const inputSearch = document.getElementById('brId');
    if (inputSearch) {
        const newInput = inputSearch.cloneNode(true);
        inputSearch.parentNode.replaceChild(newInput, inputSearch);
        newInput.addEventListener('keyup', filterTable);
    }

    window.onclick = (e) => {
        const modal = document.getElementById('modalNP');
        if (e.target === modal) {
            closeModal();
        }
    };
}

// ----------------- App Init -----------------
async function init() {
    if (isInitialized) {
        console.log('⚠️ Sistem sudah diinisialisasi');
        return;
    }

    try {
        await initDB();
        await muatUlangData();
        initEventListeners();
        isInitialized = true;
        console.log('✅ Sistem siap');
    } catch (error) {
        console.error('❌ Init error:', error);
        alert('❌ Gagal memulai aplikasi');
    }
}

// Global functions
window.editData = editData;
window.konfirmasiHapus = konfirmasiHapus;
window.pindahHalaman = pindahHalaman;
window.saveData = saveData;
window.closeModal = closeModal;
window.filterTable = filterTable;
window.openModal = openModal;
window.addItemToForm = addItemToForm;
window.removeItemFromForm = removeItemFromForm;
window.calculateItemTotal = calculateItemTotal;
window.updateItemData = updateItemData;

// Start
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
} else {
    init();
}

// ========================================
// EXPORT TO WORD - NOTA DINAS 
// FINAL VERSION - SUDAH TERMASUK JABATAN DI TTD
// ========================================

async function exportToWord(id) {
    const item = dataList.find(b => b.id === id);
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
        function toTitleCaseWithAcronym(str) {
            if (!str) return str;
            const acronyms = ['UPT', 'RSBG', 'NIP', 'ATK'];
            return str.toLowerCase().split(' ').map(word => {
                if (acronyms.includes(word.toUpperCase())) {
                    return word.toUpperCase();
                }
                return word.charAt(0).toUpperCase() + word.slice(1);
            }).join(' ');
        }

        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
                WidthType, AlignmentType, BorderStyle, ImageRun, VerticalAlign } = docx;

        const tanggal = item.TGL ? new Date(item.TGL) : new Date();
        const bulanNama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        const tgl = tanggal.getDate();
        const bulan = bulanNama[tanggal.getMonth()];
        const tahun = tanggal.getFullYear();

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
        // HEADER DENGAN LOGO
        // ============================================
        let logoImageRun = null;
        
        try {
            const logoResponse = await fetch('LOGO PEMPROV.png');
            if (logoResponse.ok) {
                const logoBlob = await logoResponse.blob();
                const logoBuffer = await logoBlob.arrayBuffer();
                logoImageRun = new ImageRun({
                    data: logoBuffer,
                    transformation: { width: 90, height: 100 }
                });
            }
        } catch (err) {
            console.log('⚠️ Logo tidak ditemukan, menggunakan placeholder');
        }

        // Header Table
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
                        children: [
                            new TableCell({
                                width: { size: 15, type: WidthType.PERCENTAGE },
                                children: logoImageRun ? [
                                    new Paragraph({ 
                                        children: [logoImageRun],
                                        alignment: AlignmentType.CENTER
                                    })
                                ] : [
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "LOGO", bold: true, size: 16, color: "CCCCCC" }) ],
                                        alignment: AlignmentType.CENTER
                                    })
                                ],
                                borders: noBorders,
                                verticalAlign: VerticalAlign.CENTER
                            }),
                            new TableCell({
                                width: { size: 85, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "DINAS SOSIAL PROVINSI JAWA TIMUR", bold: true, size: 22 }) ],
                                        alignment: AlignmentType.CENTER
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "UPT REHABILITASI SOSIAL BINA GRAHITA TUBAN", bold: true, size: 22 }) ],
                                        alignment: AlignmentType.CENTER
                                    }),
                                    new Paragraph({ 
                                        text: "",
                                        border: {
                                            bottom: { color: "000000", space: 0, style: BorderStyle.SINGLE, size: 15 }
                                        }
                                    })
                                ],
                                borders: noBorders,
                                verticalAlign: VerticalAlign.CENTER
                            })
                        ]
                    })
                ]
            })
        );

        // Spacing setelah header
        children.push(new Paragraph({ text: "", spacing: { after: 100 } }));

        // Judul
        children.push(
            new Paragraph({
                children: [ new TextRun({ text: "NOTA DINAS", bold: true, size: 24, underline: {} }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 150 }
            })
        );

        // Tabel informasi
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
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Kepada Yth", size: 22 })] })], borders: noBorders, width: { size: 20, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders, width: { size: 3, type: WidthType.PERCENTAGE } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: toTitleCaseWithAcronym(item.PIHAK1 || ''), size: 22 })] })], borders: noBorders, width: { size: 77, type: WidthType.PERCENTAGE } })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.NIP1 || '', size: 22 })] })], borders: noBorders })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: toTitleCaseWithAcronym(item.JBT1 || ''), size: 22 })] })], borders: noBorders })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Dari", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: toTitleCaseWithAcronym(item.PIHAK2 || ''), size: 22 })] })], borders: noBorders })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.NIP2 || '', size: 22 })] })], borders: noBorders })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: toTitleCaseWithAcronym(item.JBT2 || ''), size: 22 })] })], borders: noBorders })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Tanggal", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${tgl} ${bulan} ${tahun}`, size: 22 })] })], borders: noBorders })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Nomor", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.NOBAP || '', size: 22 })] })], borders: noBorders })
                        ]
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Perihal", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })], borders: noBorders }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: toTitleCaseWithAcronym(item.PERIHAL || ''), size: 22 })] })], borders: noBorders })
                        ]
                    })
                ]
            })
        );

        // Garis pembatas
        children.push(
            new Paragraph({ 
                text: "",
                border: { bottom: { color: "000000", space: 0, style: BorderStyle.SINGLE, size: 10 } },
                spacing: { before: 150, after: 200 }
            })
        );

        // Paragraf pembuka
        children.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "Dengan hormat,", size: 22 }) ],
                spacing: { after: 200 },
                alignment: AlignmentType.LEFT
            })
        );

        children.push(
            new Paragraph({ 
                children: [ new TextRun({ 
                    text: `Sehubungan dengan pelaksanaan operasional Sub Bagian ${toTitleCaseWithAcronym(item.SEKSI || '')} di UPT Rehabilitasi Sosial Bina Grahita Tuban, bersama ini kami mengajukan permohonan barang sebagaimana terlampir dalam tabel berikut :`, 
                    size: 22 
                }) ],
                spacing: { after: 300 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        // Tabel barang
        const items = item.ITEMS && Array.isArray(item.ITEMS) && item.ITEMS.length > 0 
            ? item.ITEMS 
            : [{
                namaBarang: item.NAMABARANG || '',
                jumlah: item.JUMLAH || 0,
                satuan: item.SATUAN || 'Pcs',
                hargaPerItem: item.HARGA || 0,
                subtotal: (item.JUMLAH || 0) * (item.HARGA || 0)
            }];

        const tableRows = [];
        
        // Header tabel
        tableRows.push(
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "No", bold: true, size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders, width: { size: 8, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Nama Barang", bold: true, size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders, width: { size: 40, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jumlah", bold: true, size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders, width: { size: 15, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Harga Satuan", bold: true, size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders, width: { size: 18, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jumlah Harga", bold: true, size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders, width: { size: 19, type: WidthType.PERCENTAGE } })
                ]
            })
        );

        // Isi tabel
        let totalKeseluruhan = 0;
        items.forEach((subItem, idx) => {
            const jumlah = subItem.jumlah || 0;
            const satuan = subItem.satuan || 'Pcs';
            const hargaSatuan = subItem.hargaPerItem || 0;
            const subtotal = subItem.subtotal || (jumlah * hargaSatuan);
            totalKeseluruhan += subtotal;

            tableRows.push(
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(idx + 1), size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: toTitleCaseWithAcronym(subItem.namaBarang || ''), size: 20 })], alignment: AlignmentType.LEFT })], verticalAlign: VerticalAlign.CENTER, borders: borders }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${jumlah} ${satuan}`, size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formatRupiah(hargaSatuan), size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders }),
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formatRupiah(subtotal), size: 20 })], alignment: AlignmentType.RIGHT })], verticalAlign: VerticalAlign.CENTER, borders: borders })
                    ]
                })
            );
        });

        // Row total
        tableRows.push(
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jumlah", bold: true, size: 20 })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER, borders: borders, columnSpan: 4 }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formatRupiah(totalKeseluruhan), bold: true, size: 20 })], alignment: AlignmentType.RIGHT })], verticalAlign: VerticalAlign.CENTER, borders: borders })
                ]
            })
        );

        children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }));
        children.push(new Paragraph({ text: "", spacing: { after: 400 } }));

        // ============================================
        // TANDA TANGAN - DENGAN JABATAN
        // ============================================
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
                        children: [
                            // KOLOM KIRI - MENGETAHUI
                            new TableCell({
                                width: { size: 50, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: "Mengetahui,", size: 22 }) ],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 100 }
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: toTitleCaseWithAcronym(item.JBT1 || ''), size: 22 }) ],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 1200 }
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: toTitleCaseWithAcronym(item.PIHAK1 || ''), size: 22 }) ],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 50 }
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: `NIP ${item.NIP1 || ''}`, size: 22 }) ],
                                        alignment: AlignmentType.CENTER 
                                    })
                                ],
                                borders: noBorders
                            }),
                            // KOLOM KANAN - TANGGAL DAN PENGAJU
                            new TableCell({
                                width: { size: 50, type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({ 
                                        children: [new TextRun({ text: `Tuban, ${tgl} ${bulan} ${tahun}`, size: 22 })],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 100 }
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: toTitleCaseWithAcronym(item.JBT2 || ''), size: 22 }) ],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 1200 }
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: toTitleCaseWithAcronym(item.PIHAK2 || ''), size: 22 }) ],
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 50 }
                                    }),
                                    new Paragraph({ 
                                        children: [ new TextRun({ text: `NIP ${item.NIP2 || ''}`, size: 22 }) ],
                                        alignment: AlignmentType.CENTER 
                                    })
                                ],
                                borders: noBorders
                            })
                        ]
                    })
                ]
            })
        );

        // Create Document
        const doc = new Document({
            styles: {
                default: {
                    document: {
                        run: { font: "Times New Roman", size: 22 },
                        paragraph: { spacing: { line: 276 } }
                    }
                }
            },
            sections: [{
                properties: {
                    page: {
                        margin: { top: 720, right: 1440, bottom: 720, left: 1440 }
                    }
                },
                children
            }]
        });

        const blob = await Packer.toBlob(doc);
        const safeName = (item.NOS || `NOTA_${id}`).replace(/[\\\/:*?"<>|]/g, '_');
        const fileName = `NOTA_DINAS_${safeName}.docx`;

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

// Fungsi helper untuk konfirmasi export
function konfirmasiExport(id) {
    const data = dataList.find(r => r.id === id);
    const pesan = data 
        ? `📄 Export ke Word?\n\nNo: ${data.NOS}\nPerihal: ${data.PERIHAL}`
        : '📄 Export data ini ke Word?';
    
    if (confirm(pesan)) {
        exportToWord(id);
    }
}

// Register global function
window.exportToWord = exportToWord;
window.konfirmasiExport = konfirmasiExport;

console.log('✅ Export to Word function loaded - VERSION WITH JABATAN IN TTD');

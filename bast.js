// Tambahkan di awal file JavaScript
function checkDocxLibrary() {
    if (typeof window.docx === 'undefined') {
        console.error('❌ Library docx belum dimuat!');
        alert('⚠️ Library docx belum dimuat.\n\nPastikan CDN docx.js sudah di-load di HTML.\nRefresh halaman dan coba lagi.');
        return false;
    }
    console.log('✅ Library docx berhasil dimuat!');
    return true;
}

// Di awal file, sebelum semua fungsi
(function loadDocxLibrary() {
    if (typeof window.docx !== 'undefined') {
        console.log('✅ Docx library already loaded');
        return;
    }
    
    console.log('⏳ Loading docx library...');
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/docx/7.8.2/docx.min.js';
    script.onload = function() {
        console.log('✅ Docx library loaded successfully!');
    };
    script.onerror = function() {
        console.error('❌ Failed to load docx library');
        alert('Gagal memuat library docx. Cek koneksi internet Anda.');
    };
    document.head.appendChild(script);
})();

// ====== BAPGenerator Fixed JS ======

// ----------------- Constants & DB setup -----------------
let db;
const DB_NAME = 'BAPGeneratorDB';
const DB_VERSION = 1;
const STORE_NAME = 'barang';

// ----------------- App State -----------------
let barangList = [];
let editIndex = null;
let previewImageData = null;
let currentPage = 1;
let entriesPerPage = 10;
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

function formatRupiahDecimals(value) {
    if (value == null || value === "") return "0,00";
    const num = parseRupiahToFloat(value);
    return num.toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function terbilang(angka) {
    const bilangan = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas'];
    angka = Number(angka) || 0;
    if (angka < 12) return bilangan[angka];
    if (angka < 20) return terbilang(angka - 10) + ' Belas';
    if (angka < 100) return terbilang(Math.floor(angka / 10)) + ' Puluh ' + terbilang(angka % 10);
    if (angka < 200) return 'Seratus ' + terbilang(angka - 100);
    if (angka < 1000) return terbilang(Math.floor(angka / 100)) + ' Ratus ' + terbilang(angka % 100);
    if (angka < 2000) return 'Seribu ' + terbilang(angka - 1000);
    if (angka < 1000000) return terbilang(Math.floor(angka / 1000)) + ' Ribu ' + terbilang(angka % 1000);
    if (angka < 1000000000) return terbilang(Math.floor(angka / 1000000)) + ' Juta ' + terbilang(angka % 1000000);
    if (angka < 1000000000000) return terbilang(Math.floor(angka / 1000000000)) + ' Miliar ' + terbilang(angka % 1000000000);
    return 'Angka terlalu besar';
}

function capitalizeFirst(s) {
    if (!s) return s;
    return s.charAt(0).toUpperCase() + s.slice(1);
}

function dataURLtoArrayBuffer(dataURL) {
    const base64 = dataURL.split(',')[1];
    if (!base64) return new ArrayBuffer(0);
    const binary = atob(base64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
}

// ----------------- IndexedDB functions -----------------
function initIndexedDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onerror = () => {
            console.error('IndexedDB error:', request.error);
            reject(request.error);
        };

        request.onsuccess = () => {
            db = request.result;
            console.log('IndexedDB initialized successfully');
            resolve(db);
        };

        request.onupgradeneeded = (event) => {
            db = event.target.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                const objectStore = db.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
                objectStore.createIndex('NOS', 'NOS', { unique: false });
                objectStore.createIndex('TGL', 'TGL', { unique: false });
                console.log('Object store created');
            }
        };
    });
}

function saveToIndexedDB(data) {
    return new Promise((resolve, reject) => {
        if (!db) {
            reject(new Error('Database not initialized'));
            return;
        }
        const transaction = db.transaction([STORE_NAME], 'readwrite');
        const objectStore = transaction.objectStore(STORE_NAME);

        if (data.id) {
            const request = objectStore.put(data);
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        } else {
            const request = objectStore.add(data);
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        }
    });
}

function getAllFromIndexedDB() {
    return new Promise((resolve, reject) => {
        if (!db) {
            reject(new Error('Database not initialized'));
            return;
        }
        const transaction = db.transaction([STORE_NAME], 'readonly');
        const objectStore = transaction.objectStore(STORE_NAME);
        const request = objectStore.getAll();

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

function deleteFromIndexedDB(id) {
    return new Promise((resolve, reject) => {
        if (!db) {
            reject(new Error('Database not initialized'));
            return;
        }
        const transaction = db.transaction([STORE_NAME], 'readwrite');
        const objectStore = transaction.objectStore(STORE_NAME);
        const request = objectStore.delete(id);

        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
    });
}

// ----------------- Google Drive Sync -----------------
async function syncToGoogleDrive() {
    try {
        const allData = await getAllFromIndexedDB();

        if (typeof window.gdrive_upload === 'function') {
            const jsonBlob = new Blob([JSON.stringify(allData, null, 2)], { type: 'application/json' });
            const fileName = `BAP_Backup_${new Date().toISOString().split('T')[0]}.json`;

            await window.gdrive_upload({
                file: jsonBlob,
                filename: fileName,
                mimeType: 'application/json'
            });

            alert(`✅ ${allData.length} data berhasil di-backup ke Google Drive!\nFile: ${fileName}`);
        } else {
            const jsonStr = JSON.stringify(allData, null, 2);
            const blob = new Blob([jsonStr], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `BAP_Backup_${new Date().toISOString().split('T')[0]}.json`;
            link.click();
            URL.revokeObjectURL(url);
            alert('⚠️ Google Drive tool belum diaktifkan.\n📥 Backup file telah didownload.');
        }
    } catch (error) {
        console.error('Sync error:', error);
        alert('❌ Gagal sync ke Google Drive: ' + (error.message || error));
    }
}

// ----------------- UI Rendering & Pagination -----------------
function updateShowEntries(start, end, total) {
    const showEntries = document.querySelector('.showEntries');
    if (showEntries) {
        showEntries.textContent = `Showing ${start} to ${end} of ${total} entries`;
    }
}

// ====== FUNGSI RENDER TABLE DENGAN MULTIPLE ITEMS & IMAGES (FIXED) ======

function renderTable(filteredData = null) {
    const tbody = document.querySelector('.userBarang');
    if (!tbody) {
        console.error('Element with class "userBarang" not found');
        return;
    }

    const data = filteredData || barangList;

    // Handle empty data
    if (!Array.isArray(data) || data.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="17" style="text-align: center; padding: 20px; color: gray;">
                    ${barangList.length === 0 ? '📦 Belum ada data' : '🔍 Data tidak ditemukan'}
                </td>
            </tr>
        `;
        if (typeof updateShowEntries === 'function') {
            updateShowEntries(0, 0, barangList.length);
        }
        return;
    }

    // Pagination
    const startIndex = (currentPage - 1) * entriesPerPage;
    const endIndex = startIndex + entriesPerPage;
    const paginatedData = data.slice(startIndex, endIndex);

    // Build table rows
    tbody.innerHTML = paginatedData.map((item) => {
        const items = item.items || [];
      
        let totalHargaSebelumPajak = 0;
        let totalPPN = 0;
        let grandTotal = 0;
        
        
        items.forEach(subItem => {
            const jumlah = parseInt(subItem.Jumlah) || 1;
            
            // HARGA SEBELUM PAJAK per satuan
            let hargaSebelumPajakPerSatuan = parseRupiahToFloat(subItem.Hrgsbl || "0");
            if (hargaSebelumPajakPerSatuan === 0) {
                hargaSebelumPajakPerSatuan = parseRupiahToFloat(subItem.HrgItem || "0");
            }
            
            // PPN per satuan
            const ppnPerSatuan = parseRupiahToFloat(subItem.PPN || "0");
            
            // HARGA SETELAH PAJAK = Harga Sebelum Pajak + (PPN × Quantity)
            const hargaSetelahPajak = hargaSebelumPajakPerSatuan + (ppnPerSatuan * jumlah);
            
            // Akumulasi total
            totalHargaSebelumPajak += hargaSebelumPajakPerSatuan;
            totalPPN += ppnPerSatuan;
            grandTotal += hargaSetelahPajak;
        });

        const itemCount = items.length;
        
        // ===== GABUNGKAN SEMUA URAIAN DENGAN NOMOR =====
        const allUraian = items.length > 0 
            ? items.map((i, idx) => `${idx + 1}. ${i.Uraian || '-'}`).join('<br>')
            : '-';
        
        // ===== DETAIL HARGA PER ITEM =====
        const detailHargaPerItem = items.length > 0
            ? items.map((i, idx) => {
                // Gunakan HargaPerItem jika ada, jika tidak gunakan Hrgsbl, jika tidak ada gunakan HrgItem
                let hargaPerItem = parseRupiahToFloat(i.HargaPerItem || "0");
                if (hargaPerItem === 0) {
                    hargaPerItem = parseRupiahToFloat(i || "0");
                    if (hargaPerItem === 0) {
                        hargaPerItem = parseRupiahToFloat(i.HrgItem || "0");
                    }
                }
                
                return `${idx + 1}. Rp ${formatRupiah(hargaPerItem)}`;
            }).join('<br>')
            : '-';
        
        // Total Harga Per Item
        let totalHargaPerItem = 0;
        items.forEach(i => {
            let hargaPerItem = parseRupiahToFloat(i.HargaPerItem || "0");
            if (hargaPerItem === 0) {
                hargaPerItem = parseRupiahToFloat(i || "0");
                if (hargaPerItem === 0) {
                    hargaPerItem = parseRupiahToFloat(i.HrgItem || "0");
                }
            }
            totalHargaPerItem += hargaPerItem;
        });

        // ===== DETAIL HARGA SEBELUM PAJAK PER ITEM =====
        const detailHargaSebelumPajak = items.length > 0
            ? items.map((i, idx) => {
                // Harga sebelum pajak per satuan
                let hrgsblPerSatuan = parseRupiahToFloat(i.Hrgsbl || "0");
                if (hrgsblPerSatuan === 0) {
                    hrgsblPerSatuan = parseRupiahToFloat(i.HrgItem || "0");
                }
                
                return `${idx + 1}. Rp ${formatRupiah(hrgsblPerSatuan)}`;
            }).join('<br>')
            : '-';

        // ===== DETAIL PPN PER ITEM =====
        const detailPPN = items.length > 0
            ? items.map((i, idx) => {
                const ppnPerSatuan = parseRupiahToFloat(i.PPN || "0");
                return `${idx + 1}. Rp ${formatRupiah(ppnPerSatuan)}`;
            }).join('<br>')
            : '-';
        
        // ===== DETAIL HARGA SETELAH PAJAK PER ITEM =====
        const detailHargaSetelahPajak = items.length > 0
            ? items.map((i, idx) => {
                const jumlah = parseInt(i.Jumlah) || 1;
                
                // Harga sebelum pajak per satuan
                let hrgsblPerSatuan = parseRupiahToFloat(i.Hrgsbl || "0");
                if (hrgsblPerSatuan === 0) {
                    hrgsblPerSatuan = parseRupiahToFloat(i.HrgItem || "0");
                }
                
                // PPN per satuan
                const ppnPerSatuan = parseRupiahToFloat(i.PPN || "0");
                
                // Harga Setelah Pajak = Harga Sebelum Pajak + (PPN × Quantity)
                const hargaSetelahPajak = hrgsblPerSatuan + (ppnPerSatuan * jumlah);
                
                return `${idx + 1}. Rp ${formatRupiah(hargaSetelahPajak)} <span style="color: #666; font-size: 0.85em;">(Qty: ${jumlah})</span>`;
            }).join('<br>')
            : '-';
        
        // ===== GABUNGKAN DOKUMENTASI/GAMBAR SEMUA ITEM =====
        let docsDisplay = '-';
        
        // Jika ada multiple items dengan gambar masing-masing
        if (items.length > 0 && items.some(i => i.Image)) {
            docsDisplay = '<div style="display: flex; flex-direction: column; gap: 8px;">';
            items.forEach((subItem, idx) => {
                if (subItem.Image) {
                    const uraianLabel = subItem.Uraian 
                        ? (subItem.Uraian.length > 20 ? subItem.Uraian.substring(0, 20) + '...' : subItem.Uraian)
                        : `Item ${idx + 1}`;
                    
                    docsDisplay += `
                        <div style="border: 1px solid #e0e0e0; padding: 5px; border-radius: 4px; background: #f9f9f9;">
                            <div style="font-size: 0.75em; color: #666; margin-bottom: 3px;">
                                ${idx + 1}
                            </div>
                            <img src="${subItem.Image}" 
                                 alt="item${idx + 1}" 
                                 style="width: 50px; height: 50px; object-fit: cover; cursor: pointer; border-radius: 4px;" 
                                 onclick="window.open('${subItem.Image}', '_blank')"
                                 onerror="this.src='data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 width=%2250%22 height=%2250%22%3E%3Crect fill=%22%23ddd%22 width=%2250%22 height=%2250%22/%3E%3Ctext x=%2250%25%22 y=%2250%25%22 dominant-baseline=%22middle%22 text-anchor=%22middle%22 font-size=%2212%22 fill=%22%23999%22%3ENo Image%3C/text%3E%3C/svg%3E'">
                        </div>
                    `;
                }
            });
            docsDisplay += '</div>';
        }
        // Jika hanya ada 1 gambar utama (backward compatibility)
        else if (item.Image) {
            docsDisplay = `
                <img src="${item.Image}" 
                     alt="preview" 
                     style="width: 50px; height: 50px; object-fit: cover; cursor: pointer; border-radius: 4px;" 
                     onclick="window.open('${item.Image}', '_blank')"
                     onerror="this.src='data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 width=%2250%22 height=%2250%22%3E%3Crect fill=%22%23ddd%22 width=%2250%22 height=%2250%22/%3E%3Ctext x=%2250%25%22 y=%2250%25%22 dominant-baseline=%22middle%22 text-anchor=%22middle%22 font-size=%2212%22 fill=%22%23999%22%3ENo Image%3C/text%3E%3C/svg%3E'">
            `;
        }

        return `
        <tr style="border-bottom: 1px solid #ddd;">
            <td style="padding: 10px;">${escapeHtml(item.NOS || '-')}</td>
            <td style="padding: 10px;">${escapeHtml(item.TGL || '-')}</td>
            <td style="padding: 10px;">${escapeHtml(item.PIHAK1 || '-')}</td>
            <td style="padding: 10px;">${escapeHtml(item.NIP1 || '-')}</td>
            <td style="padding: 10px;">${escapeHtml(item.PIHAK2 || '-')}</td>
            <td style="padding: 10px;">${escapeHtml(item.NIP2 || '-')}</td>
            <td style="padding: 10px;">${escapeHtml(item.JenBel || '-')}</td>
            <td style="padding: 10px;">${escapeHtml(item.Deskripsi || '-')}</td>
            <td style="padding: 10px; text-align: center;">
                <span style="background: #4CAF50; color: white; padding: 4px 8px; border-radius: 12px; font-size: 0.85em; font-weight: bold;">
                    ${itemCount} item${itemCount > 1 ? 's' : ''}
                </span>
            </td>
            <td style="padding: 10px;">${escapeHtml(item.NOBAP || '-')}</td>
            <td style="padding: 10px; max-width: 250px;">
                <div style="max-height: 150px; overflow-y: auto; font-size: 0.9em; line-height: 1.6;">
                    ${allUraian}
                </div>
            </td>
            <td style="padding: 10px; text-align: right; white-space: nowrap;">
                <div style="max-height: 120px; overflow-y: auto; font-size: 0.85em; line-height: 1.5; text-align: left;">
                    ${items.length > 1 ? detailHargaPerItem + '<br><hr style="margin: 5px 0; border: none; border-top: 1px solid #ddd;">' : (items.length === 1 ? detailHargaPerItem + '<br>' : '')}
                    <strong style="color: #4CAF50;">Total: Rp ${formatRupiah(totalHargaPerItem)}</strong>
                </div>
            </td>
            <td style="padding: 10px; text-align: right; white-space: nowrap;">
                <div style="max-height: 120px; overflow-y: auto; font-size: 0.85em; line-height: 1.5; text-align: left;">
                    ${items.length > 1 ? detailHargaSebelumPajak + '<br><hr style="margin: 5px 0; border: none; border-top: 1px solid #ddd;">' : (items.length === 1 ? detailHargaSebelumPajak + '<br>' : '')}
                    <strong style="color: #1976D2;">Total: Rp ${formatRupiah(totalHargaSebelumPajak)}</strong>
                </div>
            </td>
            <td style="padding: 10px; text-align: right; white-space: nowrap;">
                <div style="max-height: 120px; overflow-y: auto; font-size: 0.85em; line-height: 1.5; text-align: left;">
                    ${items.length > 1 ? detailPPN + '<br><hr style="margin: 5px 0; border: none; border-top: 1px solid #ddd;">' : (items.length === 1 ? detailPPN + '<br>' : '')}
                    <strong style="color: #D32F2F;">Total: Rp ${formatRupiah(totalPPN)}</strong>
                </div>
            </td>
            <td style="padding: 10px; text-align: right; white-space: nowrap;">
                <div style="max-height: 120px; overflow-y: auto; font-size: 0.85em; line-height: 1.5; text-align: left; background: #fff3e0; border-radius: 4px; padding: 8px;">
                    ${items.length > 1 ? detailHargaSetelahPajak + '<br><hr style="margin: 5px 0; border: none; border-top: 1px solid #ddd;">' : (items.length === 1 ? detailHargaSetelahPajak + '<br>' : '')}
                    <strong style="color: #FF6F00; font-size: 1.05em;">Grand Total: Rp ${formatRupiah(grandTotal)}</strong>
                </div>
            </td>
            <td style="padding: 10px; text-align: center; max-width: 150px;">
                <div style="max-height: 250px; overflow-y: auto;">
                    ${docsDisplay}
                </div>
            </td>
            <td style="padding: 10px; text-align: center; white-space: nowrap;">
                <button onclick="editData(${item.id})" 
                        style="padding: 6px 12px; margin: 2px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 0.85em;" 
                        title="Edit">
                    ✏️ Edit
                </button>
                <button onclick="deleteData(${item.id})" 
                        style="padding: 6px 12px; margin: 2px; background: #f44336; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 0.85em;" 
                        title="Hapus">
                    🗑️ Hapus
                </button>
                <button onclick="exportToWord(${item.id})" 
                        style="padding: 6px 12px; margin: 2px; background: #2196F3; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 0.85em;" 
                        title="Download Word">
                    📥 Word
                </button>
            </td>
        </tr>
        `;
    }).join('');

    // Update pagination info
    if (typeof updateShowEntries === 'function') {
        updateShowEntries(startIndex + 1, Math.min(endIndex, data.length), data.length);
    }
}

// ===== HELPER FUNCTION: ESCAPE HTML =====
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, m => map[m]);
}

// ===== HELPER FUNCTION: PARSE RUPIAH TO FLOAT (MENDUKUNG KOMA DESIMAL) =====
// Contoh input: "15.000.000,50" atau "15000000.50" atau "15,000,000.50"
function parseRupiahToFloat(rupiah) {
    if (!rupiah) return 0;
    
    // Konversi ke string
    let str = String(rupiah).trim();
    
    // Hapus prefix "Rp" atau "Rp." jika ada
    str = str.replace(/^Rp\.?\s*/i, '');
    
    // Deteksi format: cek apakah menggunakan koma sebagai desimal atau titik sebagai desimal
    // Format Indonesia: 1.000.000,50 (titik = ribuan, koma = desimal)
    // Format International: 1,000,000.50 (koma = ribuan, titik = desimal)
    
    const lastComma = str.lastIndexOf(',');
    const lastDot = str.lastIndexOf('.');
    
    // Jika koma ada dan muncul setelah titik, maka koma adalah desimal (format Indonesia)
    if (lastComma > lastDot) {
        // Format Indonesia: 1.000.000,50
        str = str.replace(/\./g, ''); // Hapus semua titik (pemisah ribuan)
        str = str.replace(',', '.'); // Ganti koma dengan titik untuk desimal
    } else {
        // Format International atau tanpa desimal: 1,000,000.50 atau 1000000
        str = str.replace(/,/g, ''); // Hapus semua koma (pemisah ribuan)
    }
    
    // Hapus karakter selain angka, titik, dan minus
    str = str.replace(/[^\d.-]/g, '');
    
    // Parse ke float
    const result = parseFloat(str);
    
    return isNaN(result) ? 0 : result;
}

// ===== HELPER FUNCTION: FORMAT FLOAT TO RUPIAH (TANPA PREFIX) =====
// Contoh output: "15.000.000,50" (format Indonesia)
function formatRupiah(angka) {
    if (angka === null || angka === undefined || angka === '') return '0';
    
    // Konversi ke number
    let number = typeof angka === 'string' ? parseFloat(angka) : angka;
    
    if (isNaN(number)) return '0';
    
    // Pisahkan bagian bulat dan desimal
    let parts = number.toFixed(2).split('.');
    let integerPart = parts[0];
    let decimalPart = parts[1];
    
    // Format bagian bulat dengan titik sebagai pemisah ribuan
    integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    
    // Jika ada desimal dan tidak 00, tampilkan dengan koma
    if (decimalPart && decimalPart !== '00') {
        return integerPart + ',' + decimalPart;
    }
    
    // Jika tidak ada desimal atau desimal = 00, tampilkan tanpa koma
    return integerPart;
}

function updatePagination() {
    const totalPages = Math.max(1, Math.ceil(barangList.length / entriesPerPage));
    const pagination = document.querySelector('.pagination');
    if (!pagination) return;

    let paginationHTML = '';

    paginationHTML += `<button onclick="changePage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>Prev</button>`;

    const maxButtons = 5;
    const half = Math.floor(maxButtons / 2);
    let start = Math.max(1, currentPage - half);
    let end = Math.min(totalPages, start + maxButtons - 1);
    if (end - start < maxButtons - 1) {
        start = Math.max(1, end - maxButtons + 1);
    }

    for (let i = start; i <= end; i++) {
        paginationHTML += `<button onclick="changePage(${i})" class="${i === currentPage ? 'active' : ''}">${i}</button>`;
    }

    paginationHTML += `<button onclick="changePage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>Next</button>`;

    pagination.innerHTML = paginationHTML;
}

function changePage(page) {
    const totalPages = Math.max(1, Math.ceil(barangList.length / entriesPerPage));
    if (page < 1 || page > totalPages) return;
    currentPage = page;
    renderTable();
    updatePagination();
}

// ----------------- Search & Filter -----------------
function filterTable() {
    const searchInput = document.getElementById('brId');
    if (!searchInput) return;
    const searchTerm = searchInput.value.toLowerCase();
    const filtered = barangList.filter(item =>
        (item.NOS || '').toLowerCase().includes(searchTerm) ||
        (item.PIHAK1 || '').toLowerCase().includes(searchTerm) ||
        (item.PIHAK2 || '').toLowerCase().includes(searchTerm) ||
        (item.Deskripsi || '').toLowerCase().includes(searchTerm) ||
        (item.JenBel || '').toLowerCase().includes(searchTerm)
    );
    currentPage = 1;
    renderTable(filtered);
    updatePagination();
}

// ----------------- Modal Functions - FIXED -----------------
function closeModal() {
    const modal = document.getElementById('modal');
    if (modal) {
        modal.classList.remove('active');
        modal.style.display = 'none';
    }
    resetForm();
}

function resetForm() {
    const fields = {
        'NOS': '',
        'TGL': '',
        'PIHAK1': '',
        'NIP1': '',
        'PIHAK2': '',
        'NIP2': '',
        'JenBel': '',
        'Deskripsi': '',
        'NOBAP': ''
    };

    for (const [id, value] of Object.entries(fields)) {
        const element = document.getElementById(id);
        if (element) {
            element.value = value;
        }
    }

    const imageInput = document.getElementById('imageInput');
    if (imageInput) imageInput.value = '';

    const previewContainer = document.getElementById('previewContainer');
    if (previewContainer) previewContainer.classList.add('hidden');

    const itemsContainer = document.getElementById('itemsContainer');
    if (itemsContainer) itemsContainer.innerHTML = '';

    const addItemBtn = document.getElementById('addItemBtn');
    if (addItemBtn) addItemBtn.remove();

    const grandTotalDiv = document.getElementById('grandTotalDiv');
    if (grandTotalDiv) grandTotalDiv.remove();

    previewImageData = null;
    itemsInForm = [];
}

function showAddModal() {
    editIndex = null;
    const modalTitle = document.getElementById('modalTitle');
    if (modalTitle) modalTitle.textContent = '➕ Add BAP';
    
    resetForm();

    const modal = document.getElementById('modal');
    if (modal) {
        modal.classList.add('active');
        modal.style.display = 'block';
    }

    setTimeout(() => {
        if (itemsInForm.length === 0) {
            addNewItemRow();
        }
    }, 100);
}

// ----------------- Multiple Items UI -----------------
function addNewItemRow() {
    let itemsContainer = document.getElementById('itemsContainer');
    if (!itemsContainer) {
        const modalContent = document.querySelector('.modalbar .p-6') || document.querySelector('.modalbar') || document.body;
        itemsContainer = document.createElement('div');
        itemsContainer.id = 'itemsContainer';
        itemsContainer.style.marginTop = '20px';
        modalContent.appendChild(itemsContainer);
    }

    const itemIndex = itemsInForm.length;

    const itemDiv = document.createElement('div');
    itemDiv.className = 'item-row';
    itemDiv.style.cssText = 'border: 2px solid #e0e0e0; padding: 15px; margin-bottom: 15px; border-radius: 8px; background: #f9f9f9;';
    itemDiv.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <h4 style="margin: 0; color: #333;">📦 Item #${itemIndex + 1}</h4>
            ${itemIndex > 0 ? `<button type="button" onclick="removeItemRow(${itemIndex})" style="background: #f44336; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer;">🗑️ Hapus Item</button>` : ''}
        </div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 10px;">
            <div>
                <label style="display: block; margin-bottom: 5px; font-weight: 500;">Jumlah/Unit</label>
                <input type="number" id="Jumlah_${itemIndex}" value="1" min="1" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;" onchange="calculateItemTotal(${itemIndex})">
            </div>
            <div>
                <label style="display: block; margin-bottom: 5px; font-weight: 500;">Harga Per Item (Rp)</label>
                <input type="text" id="HrgItem_${itemIndex}" placeholder="0" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;" onchange="calculateItemTotal(${itemIndex})">
            </div>
        </div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 10px;">
            <div>
                <label style="display: block; margin-bottom: 5px; font-weight: 500;">Harga Sebelum Pajak (Rp)</label>
                <input type="text" id="Hrgsbl_${itemIndex}" placeholder="0" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;" readonly>
            </div>
            <div>
                <label style="display: block; margin-bottom: 5px; font-weight: 500;">PPN Per Item (Rp)</label>
                <input type="text" id="PPN_${itemIndex}" placeholder="0" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;" onchange="calculateItemTotal(${itemIndex})">
            </div>
        </div>
        <div style="margin-bottom: 10px;">
            <label style="display: block; margin-bottom: 5px; font-weight: 500;">Uraian Detail</label>
            <textarea id="Uraian_${itemIndex}" rows="3" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;" placeholder="Masukkan uraian barang..."></textarea>
        </div>
        
        <!-- BAGIAN UPLOAD GAMBAR PER ITEM -->
        <div style="border-top: 2px solid #ddd; padding-top: 10px; margin-top: 10px;">
            <label style="display: block; margin-bottom: 8px; font-weight: 500; color: #4CAF50;">📷 Upload Gambar Item</label>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                <!-- Gambar 1 -->
                <div>
                    <label style="font-size: 12px; color: #666;">Gambar 1</label>
                    <input type="file" id="itemImage1_${itemIndex}" accept="image/*" 
                           style="width: 100%; padding: 5px; border: 1px solid #4CAF50; border-radius: 4px; font-size: 12px; margin-bottom: 5px;"
                           onchange="handleItemImageUpload(${itemIndex}, 1, event)">
                    <div id="itemPreview1_${itemIndex}" style="display: none; position: relative;">
                        <img id="itemImg1_${itemIndex}" style="width: 100%; height: 80px; object-fit: cover; border-radius: 4px; border: 2px solid #4CAF50;">
                        <button type="button" onclick="removeItemImage(${itemIndex}, 1)" 
                                style="position: absolute; top: 2px; right: 2px; background: #f44336; color: white; border: none; border-radius: 50%; width: 20px; height: 20px; cursor: pointer; font-size: 12px; padding: 0;">×</button>
                    </div>
                </div>
                <!-- Gambar 2 -->
                <div>
                    <label style="font-size: 12px; color: #666;">Gambar 2 (Opsional)</label>
                    <input type="file" id="itemImage2_${itemIndex}" accept="image/*" 
                           style="width: 100%; padding: 5px; border: 1px solid #2196F3; border-radius: 4px; font-size: 12px; margin-bottom: 5px;"
                           onchange="handleItemImageUpload(${itemIndex}, 2, event)">
                    <div id="itemPreview2_${itemIndex}" style="display: none; position: relative;">
                        <img id="itemImg2_${itemIndex}" style="width: 100%; height: 80px; object-fit: cover; border-radius: 4px; border: 2px solid #2196F3;">
                        <button type="button" onclick="removeItemImage(${itemIndex}, 2)" 
                                style="position: absolute; top: 2px; right: 2px; background: #f44336; color: white; border: none; border-radius: 50%; width: 20px; height: 20px; cursor: pointer; font-size: 12px; padding: 0;">×</button>
                    </div>
                </div>
            </div>
        </div>
        
         <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 15px; border-radius: 8px; text-align: right; margin-top: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
      <strong style="color: white; font-size: 16px;">Subtotal Item: Rp <span id="subtotal_${itemIndex}">0</span></strong>
    </div>
  `;

    itemsContainer.appendChild(itemDiv);
    itemsInForm.push({ Image: null, Image2: null }); // Simpan referensi gambar

    updateAddItemButton();
    updateGrandTotalDisplay();
    setTimeout(() => calculateItemTotal(itemIndex), 20);
}

function handleItemImageUpload(itemIndex, imageNumber, event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // Validasi
    if (file.size > 5 * 1024 * 1024) {
        alert('⚠️ Ukuran file terlalu besar! Maksimal 5MB');
        event.target.value = '';
        return;
    }
    
    if (!file.type.startsWith('image/')) {
        alert('⚠️ File harus berupa gambar!');
        event.target.value = '';
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const base64 = e.target.result;
        
        // Simpan ke itemsInForm
        if (imageNumber === 1) {
            itemsInForm[itemIndex].Image = base64;
        } else {
            itemsInForm[itemIndex].Image2 = base64;
        }
        
        // Tampilkan preview
        const previewDiv = document.getElementById(`itemPreview${imageNumber}_${itemIndex}`);
        const imgElement = document.getElementById(`itemImg${imageNumber}_${itemIndex}`);
        
        if (previewDiv && imgElement) {
            imgElement.src = base64;
            previewDiv.style.display = 'block';
        }
    };
    reader.readAsDataURL(file);
}

// 3. TAMBAHKAN FUNGSI untuk hapus gambar per item
function removeItemImage(itemIndex, imageNumber) {
    if (imageNumber === 1) {
        itemsInForm[itemIndex].Image = null;
        const input = document.getElementById(`itemImage1_${itemIndex}`);
        const preview = document.getElementById(`itemPreview1_${itemIndex}`);
        if (input) input.value = '';
        if (preview) preview.style.display = 'none';
    } else {
        itemsInForm[itemIndex].Image2 = null;
        const input = document.getElementById(`itemImage2_${itemIndex}`);
        const preview = document.getElementById(`itemPreview2_${itemIndex}`);
        if (input) input.value = '';
        if (preview) preview.style.display = 'none';
    }
}


function removeItemRow(index) {
    const itemsContainer = document.getElementById('itemsContainer');
    if (!itemsContainer || itemsContainer.children.length <= index) return;

    if (itemsInForm.length <= 1) {
        alert('Minimal harus ada 1 item!');
        return;
    }

    itemsContainer.children[index].remove();
    itemsInForm.splice(index, 1);

    const tempItems = [...itemsInForm];
    itemsInForm = [];
    itemsContainer.innerHTML = '';
    tempItems.forEach(() => addNewItemRow());
    calculateGrandTotal();
}

function calculateItemTotal(index) {
    const jumlah = parseInt(document.getElementById(`Jumlah_${index}`)?.value) || 1;
    const hargaPerItem = parseRupiahToFloat(document.getElementById(`HrgItem_${index}`)?.value || '0');
    const ppnSatuan = parseRupiahToFloat(document.getElementById(`PPN_${index}`)?.value || '0');

    const hargaSebelumPajak = hargaPerItem * jumlah;

    const hrgsblElement = document.getElementById(`Hrgsbl_${index}`);
    if (hrgsblElement) {
        hrgsblElement.value = formatRupiahDecimals(hargaSebelumPajak);
    }

    const hargaSetelahPajak = hargaPerItem + ppnSatuan;
    const subtotal = hargaSetelahPajak * jumlah;

    const subtotalElement = document.getElementById(`subtotal_${index}`);
    if (subtotalElement) {
        subtotalElement.textContent = formatRupiah(Math.round(subtotal));
    }

    calculateGrandTotal();
}

function calculateGrandTotal() {
    let grandTotal = 0;
    itemsInForm.forEach((_, index) => {
        const jumlah = parseInt(document.getElementById(`Jumlah_${index}`)?.value) || 1;
        const hargaPerItem = parseRupiahToFloat(document.getElementById(`HrgItem_${index}`)?.value || '0');
        const ppnSatuan = parseRupiahToFloat(document.getElementById(`PPN_${index}`)?.value || '0');
        const hargaSetelahPajak = hargaPerItem + ppnSatuan;
        grandTotal += hargaSetelahPajak * jumlah;
    });

    const grandTotalElement = document.getElementById('grandTotal');
    if (grandTotalElement) {
        grandTotalElement.textContent = formatRupiah(Math.round(grandTotal));
    }
}

function updateAddItemButton() {
    let addItemBtn = document.getElementById('addItemBtn');
    if (!addItemBtn) {
        addItemBtn = document.createElement('button');
        addItemBtn.id = 'addItemBtn';
        addItemBtn.type = 'button';
        addItemBtn.style.cssText = 'width: 100%; padding: 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 8px; text-align: center; margin-top: 10px; cursor: pointer; font-weight: bold; font-size: 16px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: transform 0.2s;';
        addItemBtn.textContent = '➕ Tambah Item Lagi';
        addItemBtn.onclick = addNewItemRow;

        const itemsContainer = document.getElementById('itemsContainer');
        if (itemsContainer && itemsContainer.parentNode) {
            itemsContainer.parentNode.insertBefore(addItemBtn, itemsContainer.nextSibling);
        }
    }
}

function updateGrandTotalDisplay() {
    let grandTotalDiv = document.getElementById('grandTotalDiv');
    if (!grandTotalDiv) {
        grandTotalDiv = document.createElement('div');
        grandTotalDiv.id = 'grandTotalDiv';
        grandTotalDiv.style.cssText = 'background: linear-gradient(135deg, #667eea 100%); color: white; padding: 10px; border-radius: 8px; text-align: right; margin-top: 15px; font-size: 18px; box-shadow: 0 4px 8px rgba(0,0,0,0.15);';
        grandTotalDiv.innerHTML = '<strong>GRAND TOTAL: Rp <span id="grandTotal">0</span></strong>';

        const addItemBtn = document.getElementById('addItemBtn');
        if (addItemBtn && addItemBtn.parentNode) {
            addItemBtn.parentNode.insertBefore(grandTotalDiv, addItemBtn.nextSibling);
        }
    }
}

// ----------------- Image Preview -----------------
function initImageListener() {
    const imageInput = document.getElementById('imageInput');
    if (imageInput) {
        imageInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (event) => {
                previewImageData = event.target.result;
                const previewImage = document.getElementById('previewImage');
                const previewContainer = document.getElementById('previewContainer');
                if (previewImage && previewContainer) {
                    previewImage.src = previewImageData;
                    previewContainer.classList.remove('hidden');
                }
            };
            reader.readAsDataURL(file);
        });
    }
}




// ----------------- CRUD - FIXED -----------------
async function loadData() {
    try {
        barangList = await getAllFromIndexedDB();
        barangList.sort((a, b) => (a.id || 0) - (b.id || 0));
        renderTable();
        updatePagination();
    } catch (error) {
        console.error('Error loading data:', error);
        const tbody = document.querySelector('.userBarang');
        if (tbody) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="17" style="text-align: center; padding: 20px; color: red;">
                        ❌ Error loading data: ${error.message || error}
                    </td>
                </tr>
            `;
        }
    }
}

async function saveData() {
    console.log('saveData called, itemsInForm length:', itemsInForm.length);
    
    const items = [];
    for (let i = 0; i < itemsInForm.length; i++) {
        const jumlah = parseInt(document.getElementById(`Jumlah_${i}`)?.value) || 1;
        const hargaPerItem = parseRupiahToFloat(document.getElementById(`HrgItem_${i}`)?.value || '0');
        const hargaSebelumPajak = parseRupiahToFloat(document.getElementById(`Hrgsbl_${i}`)?.value || '0');
        const ppnSatuan = parseRupiahToFloat(document.getElementById(`PPN_${i}`)?.value || '0');
        const uraian = document.getElementById(`Uraian_${i}`)?.value || '';

        const hargaSetelahPajak = hargaPerItem + ppnSatuan;

       items.push({
    Jumlah: String(jumlah),
    HrgItem: String(hargaPerItem),
    Hrgsbl: String(hargaSebelumPajak),
    PPN: String(ppnSatuan),
    Hrgstl: String(hargaSetelahPajak),
    Uraian: uraian,
    Image: itemsInForm[i].Image || null,   // TAMBAHAN
    Image2: itemsInForm[i].Image2 || null  // TAMBAHAN
});
    }

    const formData = {
        NOS: document.getElementById('NOS')?.value || '',
        TGL: document.getElementById('TGL')?.value || '',
        PIHAK1: document.getElementById('PIHAK1')?.value || '',
        NIP1: document.getElementById('NIP1')?.value || '',
        PIHAK2: document.getElementById('PIHAK2')?.value || '',
        NIP2: document.getElementById('NIP2')?.value || '',
        JenBel: document.getElementById('JenBel')?.value || '',
        Deskripsi: document.getElementById('Deskripsi')?.value || '',
        NOBAP: document.getElementById('NOBAP')?.value || '',
        Image: previewImageData,
        items: items
    };

    console.log('FormData:', formData);

    if (!formData.NOS || !formData.PIHAK1 || items.length === 0) {
        alert('⚠️ Isi minimal No Surat, Pihak Kesatu, dan minimal 1 item barang!');
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
        alert('❌ Gagal menyimpan data: ' + (error.message || error));
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
    if (modalTitle) modalTitle.textContent = '✏️ Edit Data BAP';

    resetForm();

    const fields = {
        'NOS': item.NOS,
        'TGL': item.TGL,
        'PIHAK1': item.PIHAK1,
        'NIP1': item.NIP1,
        'PIHAK2': item.PIHAK2,
        'NIP2': item.NIP2,
        'JenBel': item.JenBel,
        'Deskripsi': item.Deskripsi,
        'NOBAP': item.NOBAP
    };

    for (const [idf, value] of Object.entries(fields)) {
        const element = document.getElementById(idf);
        if (element) element.value = value || '';
    }

    const items = item.items || [];
    
    setTimeout(() => {
        items.forEach((subItem, index) => {
            addNewItemRow();
            
            setTimeout(() => {
                const jumlahEl = document.getElementById(`Jumlah_${index}`);
                const hrgItemEl = document.getElementById(`HrgItem_${index}`);
                const hrgsblEl = document.getElementById(`Hrgsbl_${index}`);
                const ppnEl = document.getElementById(`PPN_${index}`);
                const uraianEl = document.getElementById(`Uraian_${index}`);

                if (jumlahEl) jumlahEl.value = subItem.Jumlah || '1';
                if (hrgItemEl) hrgItemEl.value = formatRupiahDecimals(subItem.HrgItem || subItem.Hrgsbl || '0');
                if (hrgsblEl) hrgsblEl.value = formatRupiahDecimals(subItem.Hrgsbl || '0');
                if (ppnEl) ppnEl.value = formatRupiahDecimals(subItem.PPN || '0');
                if (uraianEl) uraianEl.value = subItem.Uraian || '';
                // Load gambar item jika ada
if (subItem.Image) {
    itemsInForm[index].Image = subItem.Image;
    const imgElement = document.getElementById(`itemImg1_${index}`);
    const previewDiv = document.getElementById(`itemPreview1_${index}`);
    if (imgElement && previewDiv) {
        imgElement.src = subItem.Image;
        previewDiv.style.display = 'block';
    }
}

if (subItem.Image2) {
    itemsInForm[index].Image2 = subItem.Image2;
    const imgElement = document.getElementById(`itemImg2_${index}`);
    const previewDiv = document.getElementById(`itemPreview2_${index}`);
    if (imgElement && previewDiv) {
        imgElement.src = subItem.Image2;
        previewDiv.style.display = 'block';
    }
}
                calculateItemTotal(index);
            }, 50);
        });
    }, 100);

    if (item.Image) {
        previewImageData = item.Image;
        const previewImage = document.getElementById('previewImage');
        const previewContainer = document.getElementById('previewContainer');
        if (previewImage && previewContainer) {
            previewImage.src = item.Image;
            previewContainer.classList.remove('hidden');
        }
    }

    const modal = document.getElementById('modal');
    if (modal) {
        modal.classList.add('active');
        modal.style.display = 'block';
    }
}

async function deleteData(id) {
    if (confirm('⚠️ Yakin ingin menghapus data ini?')) {
        try {
            await deleteFromIndexedDB(id);
            await loadData();
            alert('✅ Data berhasil dihapus!');
        } catch (error) {
            console.error('Error deleting data:', error);
            alert('❌ Gagal menghapus data: ' + (error.message || error));
        }
    }
}




// ----------------- Export to Word (FIXED VERSION) -----------------
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

        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
                WidthType, AlignmentType, BorderStyle, ImageRun, VerticalAlign } = docx;

        // Date & names
        const tanggal = item.TGL ? new Date(item.TGL) : new Date();
        const hariNama = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];
        const bulanNama = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
                           'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        const hari = hariNama[tanggal.getDay()];
        const tgl = tanggal.getDate();
        const bulan = bulanNama[tanggal.getMonth()];
        const tahun = tanggal.getFullYear();

        const tglTerbilang = capitalizeFirst(terbilang(tgl) || String(tgl));
        const tahunTerbilang = capitalizeFirst(terbilang(tahun) || String(tahun));

        // Border helper
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

// ---------- HALAMAN 1 ----------
children.push(
    new Paragraph({
        children: [ new TextRun({ text: "BERITA ACARA PENYERAHAN BARANG/JASA", bold: true, size: 24, underline: {} }) ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 }
    }),
    new Paragraph({ 
        children: [ new TextRun({ text: `Nomor : ${item.NOS || ''}`, size: 24 }) ],
        alignment: AlignmentType.CENTER, 
        spacing: { after: 400 } 
    }),
    new Paragraph({
        children: [
            new TextRun({ text: `Pada hari ini `, size: 24 }),
            new TextRun({ text: `${hari}`, italic: true, bold: true, size: 24 }),
            new TextRun({ text: ` tanggal `, size: 24 }),
            new TextRun({ text: `${tglTerbilang}`, bold: true, size: 24 }),
            new TextRun({ text: ` bulan `, size: 24 }),
            new TextRun({ text: `${bulan}`, italic: true, bold: true, size: 24 }),
            new TextRun({ text: ` tahun `, size: 24 }),
            new TextRun({ text: `${tahunTerbilang}`, italic: true, bold: true, size: 24 }),
            new TextRun({ text: `, kami yang bertanda tangan dibawah ini:`, size: 24 })
        ],
        spacing: { after: 300 }
    })
);

// Pihak Kesatu menggunakan Tabel

children.push(
    new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "I. ", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: ":", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: item.PIHAK1 || 'KUSUMA ATMADJA, S.Sos', size: 24 })] })],
                        borders: noBorders,
                        width: { size: 70, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: ":", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: item.NIP1 || '197001021993021003', size: 24 })] })],
                        borders: noBorders,
                        width: { size: 70, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: ":", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Pejabat Pembuat Komitmen", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 70, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "yang selanjutnya disebut ", size: 24 }), new TextRun({ text: "PIHAK KESATU.", bold: true, size: 24 })]
                        })],
                        borders: noBorders,
                        columnSpan: 3,
                        width: { size: 95, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            })
        ],
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.NONE, size: 0 },
            bottom: { style: BorderStyle.NONE, size: 0 },
            left: { style: BorderStyle.NONE, size: 0 },
            right: { style: BorderStyle.NONE, size: 0 },
            insideHorizontal: { style: BorderStyle.NONE, size: 0 },
            insideVertical: { style: BorderStyle.NONE, size: 0 }
        }
    })
);

// Spasi antar pihak
children.push(new Paragraph({ text: "", spacing: { after: 200 } }));

// Pihak Kedua menggunakan Tabel
children.push(
    new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "II. ", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: ":", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: item.PIHAK2 || 'M. NURHASIM', size: 24 })] })],
                        borders: noBorders,
                        width: { size: 70, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "NIP", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: ":", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: item.NIP2 || '19800526 200901 1 003', size: 24 })] })],
                        borders: noBorders,
                        width: { size: 70, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Jabatan", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: ":", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "Pengurus Barang UPT", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 70, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: "", size: 24 })] })],
                        borders: noBorders,
                        width: { size: 5, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "yang selanjutnya disebut ", size: 24 }), new TextRun({ text: "PIHAK KEDUA.", bold: true, size: 24 })]
                        })],
                        borders: noBorders,
                        columnSpan: 3,
                        width: { size: 95, type: WidthType.PERCENTAGE },
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    })
                ]
            })
        ],
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.NONE, size: 0 },
            bottom: { style: BorderStyle.NONE, size: 0 },
            left: { style: BorderStyle.NONE, size: 0 },
            right: { style: BorderStyle.NONE, size: 0 },
            insideHorizontal: { style: BorderStyle.NONE, size: 0 },
            insideVertical: { style: BorderStyle.NONE, size: 0 }
        }
    })
);

children.push(
    new Paragraph({ 
        children: [ new TextRun({ text: "Dengan ini menyatakan bahwa:", size: 24 }) ],
        spacing: { before: 200, after: 120 },
        alignment: AlignmentType.JUSTIFIED
    }),
    new Paragraph({
        children: [
            new TextRun({ text: "1.\t", size: 24 }),
            new TextRun({ text: `PIHAK KESATU telah menyerahkan barang / jasa hasil pekerjaan ${item.JenBel || ''} berupa ${item.Deskripsi || ''} untuk keperluan UPT RSBG Tuban Dinas Sosial Provinsi Jawa Timur sesuai dengan Berita Acara Penerimaan ${item.NOBAP || ''} Tanggal ${String(tanggal.getDate()).padStart(2,'0')} ${bulan} ${tahun}`, size: 24 })
        ],
        spacing: { after: 160 },
        alignment: AlignmentType.JUSTIFIED,
        indent: { left: 720, hanging: 360 }
    }),
    new Paragraph({ 
        children: [ 
            new TextRun({ text: "2.\t", size: 24 }), 
            new TextRun({ text: "PIHAK KEDUA telah menerima dengan baik penyerahan barang/jasa tersebut sebagaimana daftar terlampir;", size: 24 }) 
        ], 
        spacing: { after: 200 },
        alignment: AlignmentType.JUSTIFIED,
        indent: { left: 720, hanging: 360 }
    }),
    new Paragraph({ 
        children: [ new TextRun({ text: "Demikian Berita Acara ini dibuat untuk dipergunakan seperlunya.", size: 24 }) ],
        spacing: { after: 700 },
        alignment: AlignmentType.JUSTIFIED
    })
);
       // Tanda tangan halaman 1 - TANPA GARIS
children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
        top: { style: BorderStyle.NONE, size: 0 },
        bottom: { style: BorderStyle.NONE, size: 0 },
        left: { style: BorderStyle.NONE, size: 0 },
        right: { style: BorderStyle.NONE, size: 0 },
        insideHorizontal: { style: BorderStyle.NONE, size: 0 },
        insideVertical: { style: BorderStyle.NONE, size: 0 }
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
                            children: [ new TextRun({ text: item.PIHAK2 || 'M. NURHASIM', bold: true, size: 24, underline: {} }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ 
                            children: [ new TextRun({ text: `NIP. ${item.NIP2 || '19800526 200901 1 003'}`, size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        })
                    ],
                    borders: {
                        top: { style: BorderStyle.NONE, size: 0 },
                        bottom: { style: BorderStyle.NONE, size: 0 },
                        left: { style: BorderStyle.NONE, size: 0 },
                        right: { style: BorderStyle.NONE, size: 0 }
                    },
                    margins: { top: 0, bottom: 0, left: 0, right: 0 }
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
                            children: [ new TextRun({ text: item.PIHAK1 || 'KUSUMA ATMADJA, S.Sos', bold: true, size: 24, underline: {} }) ],
                            alignment: AlignmentType.CENTER 
                        }),
                        new Paragraph({ 
                            children: [ new TextRun({ text: `NIP. ${item.NIP1 || '197001021993021003'}`, size: 24 }) ],
                            alignment: AlignmentType.CENTER 
                        })
                    ],
                    borders: {
                        top: { style: BorderStyle.NONE, size: 0 },
                        bottom: { style: BorderStyle.NONE, size: 0 },
                        left: { style: BorderStyle.NONE, size: 0 },
                        right: { style: BorderStyle.NONE, size: 0 }
                    },
                    margins: { top: 0, bottom: 0, left: 0, right: 0 }
                })
            ]
        })
    ]
}));


// ---------- HALAMAN 2 - Lampiran ----------
children.push(new Paragraph({ text: "", pageBreakBefore: true }));

// Lampiran heading - Dengan tab stops untuk alignment sempurna (string literal)
children.push(
    new Paragraph({ 
        children: [ new TextRun({ text: "Lampiran Berita Acara Penyerahan Barang/Jasa", size: 22 }) ],
        spacing: { after: 0 }
    }),
    new Paragraph({
        children: [
            new TextRun({ text: "Nomor\t: ", size: 22 }),
            new TextRun({ text: item.NOS || '000.2.1.1/938/1.07.06/2025', size: 22 })
        ],
        spacing: { after: 0 },
        tabStops: [
            { type: "left", position: 1440 }
        ]
    }),
    new Paragraph({
        children: [
            new TextRun({ text: "Tanggal\t: ", size: 22 }),
            new TextRun({ text: `${String(tanggal.getDate()).padStart(2,'0')} ${bulan} ${tahun}`, size: 22 })
        ],
        spacing: { after: 200 },
        tabStops: [
            { type: "left", position: 1440 }
        ]
    })
);

// TABEL DETAIL BARANG
const items = item.items || [];
const tableRows = [];

// Header Row
tableRows.push(
    new TableRow({
        children: [
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: "No.", bold: true, size: 20 })], 
                    alignment: AlignmentType.CENTER
                })],
                width: { size: 5, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                borders: borders,
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
            }),
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: "Banyaknya", bold: true, size: 20 })], 
                    alignment: AlignmentType.CENTER
                })],
                width: { size: 7, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                borders: borders,
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
            }),
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: "Uraian dan Identitas Barang/Jasa", bold: true, size: 20 })], 
                    alignment: AlignmentType.CENTER
                })],
                width: { size: 28, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                borders: borders,
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
            }),
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: "Harga Sebelum Pajak (Rp)", bold: true, size: 20 })], 
                    alignment: AlignmentType.CENTER
                })],
                width: { size: 20, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                borders: borders,
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
            }),
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: "PPN Per Item (Rp)", bold: true, size: 20 })], 
                    alignment: AlignmentType.CENTER
                })],
                width: { size: 18, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                borders: borders,
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
            }),
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: "Harga Setelah Pajak (Rp)", bold: true, size: 20 })], 
                    alignment: AlignmentType.CENTER
                })],
                width: { size: 22, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                borders: borders,
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
            })
        ]
    })
);

// Hitung Grand Total
let grandTotal = 0;

// Data Rows
items.forEach((subItem, index) => {
    const jumlah = parseInt(subItem.Jumlah) || 1;
    const uraian = subItem.Uraian || '';
    
    // Split uraian menjadi baris-baris
    const uraianLines = uraian.split('\n').filter(line => line.trim());
    
    // Buat children untuk uraian cell
    const uraianChildren = [];
    
    if (uraianLines.length > 0) {
        // Baris pertama = nama barang (BOLD)
        uraianChildren.push(
            new Paragraph({
                children: [new TextRun({ text: uraianLines[0], bold: true, size: 21 })],
                spacing: { after: 60 }
            })
        );
        
        // Baris berikutnya = spesifikasi (NORMAL)
        for (let i = 1; i < uraianLines.length; i++) {
            uraianChildren.push(
                new Paragraph({
                    children: [new TextRun({ text: uraianLines[i], size: 21 })],
                    spacing: { after: 40 }
                })
            );
        }
    } else {
        uraianChildren.push(new Paragraph({ children: [new TextRun({ text: "-", size: 21 })] }));
    }
    
    const hargaSebelumPajak = parseRupiahToFloat(subItem.Hrgsbl || 0);
    const ppn = parseRupiahToFloat(subItem.PPN || 0);
    const hargaSetelahPajak = hargaSebelumPajak + ppn * jumlah;
    
    // Tambahkan ke grand total
    grandTotal += hargaSetelahPajak;
    
    tableRows.push(
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ 
                        children: [new TextRun({ text: String(index + 1), size: 18 })], 
                        alignment: AlignmentType.CENTER
                    })],
                    width: { size: 5, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.CENTER,
                    borders: borders,
                    margins: { top: 80, bottom: 80, left: 50, right: 50 }
                }),
                new TableCell({
                    children: [new Paragraph({ 
                        children: [new TextRun({ text: String(jumlah), size: 18 })], 
                        alignment: AlignmentType.CENTER
                    })],
                    width: { size: 7, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.CENTER,
                    borders: borders,
                    margins: { top: 80, bottom: 80, left: 50, right: 50 }
                }),
                new TableCell({
                    children: uraianChildren,
                    width: { size: 28, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.TOP,
                    borders: borders,
                    margins: { top: 80, bottom: 80, left: 100, right: 100 }
                }),
                new TableCell({
                    children: [new Paragraph({ 
                        children: [new TextRun({ text: `Rp ${formatRupiahDecimals(hargaSebelumPajak)}`, size: 18 })], 
                        alignment: AlignmentType.RIGHT
                    })],
                    width: { size: 20, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.CENTER,
                    borders: borders,
                    margins: { top: 80, bottom: 80, left: 50, right: 100 }
                }),
                new TableCell({
                    children: [new Paragraph({ 
                        children: [new TextRun({ text: `Rp ${formatRupiahDecimals(ppn)}`, size: 18 })], 
                        alignment: AlignmentType.RIGHT
                    })],
                    width: { size: 18, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.CENTER,
                    borders: borders,
                    margins: { top: 80, bottom: 80, left: 50, right: 100 }
                }),
                new TableCell({
                    children: [new Paragraph({ 
                        children: [new TextRun({ text: `Rp ${formatRupiahDecimals(hargaSetelahPajak)}`, size: 18 })], 
                        alignment: AlignmentType.RIGHT
                    })],
                    width: { size: 22, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.CENTER,
                    borders: borders,
                    margins: { top: 80, bottom: 80, left: 50, right: 100 }
                })
            ]
        })
    );
});

// Hitung total untuk setiap kolom
let totalHargaSebelumPajak = 0;
let totalPPN = 0;

items.forEach((subItem) => {
    const jumlah = parseInt(subItem.Jumlah) || 1;
    const hargaSebelumPajak = parseRupiahToFloat(subItem.Hrgsbl || 0);
    const ppn = parseRupiahToFloat(subItem.PPN || 0);
    
    totalHargaSebelumPajak += hargaSebelumPajak;
    totalPPN += ppn * jumlah;
});


// ROW KOSONG (SPACE) SEBELUM TERBILANG
tableRows.push(
    new TableRow({
        children: [
            new TableCell({
                children: [new Paragraph({ text: "" })],
                columnSpan: 6,
                borders: borders,
                margins: { top: 50, bottom: 50, left: 50, right: 50 }
            })
        ]
    })
);

// ROW TERBILANG - Format seperti gambar dengan background hitam dan text hitam
const grandTotalTerbilang = capitalizeFirst(terbilang(Math.round(grandTotal))) + " Rupiah";

tableRows.push(
    new TableRow({
        children: [
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: "TERBILANG : " + grandTotalTerbilang, bold: true, size: 20 })], 
                    alignment: AlignmentType.LEFT
                })],
                columnSpan: 5,
                borders: borders,
                verticalAlign: VerticalAlign.CENTER,
                margins: { top: 120, bottom: 120, left: 100, right: 100 }
            }),
            new TableCell({
                children: [new Paragraph({ 
                    children: [new TextRun({ text: `Rp ${formatRupiahDecimals(grandTotal)}`, bold: true, size: 20 })], 
                    alignment: AlignmentType.RIGHT
                })],
                borders: borders,
                verticalAlign: VerticalAlign.CENTER,
                margins: { top: 120, bottom: 120, left: 50, right: 100 }
            })
        ]
    })
);

// Tambahkan tabel ke children
children.push(
    new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE }
    })
);

// Spasi besar antara tabel dan tanda tangan agar tanda tangan di akhir halaman 2
children.push(new Paragraph({ text: "", spacing: { after: 1400 } }));

// Tanda tangan halaman 2
children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                        new Paragraph({ children: [ new TextRun({ text: "PIHAK KEDUA", bold: true, size: 24 }) ], alignment: AlignmentType.CENTER }),
                        new Paragraph({ text: "", spacing: { after: 1200 } }),
                        new Paragraph({ children: [ new TextRun({ text: item.PIHAK2 || 'M. NURHASIM', bold: true, size: 24, underline: {} }) ], alignment: AlignmentType.CENTER }),
                        new Paragraph({ children: [ new TextRun({ text: `NIP. ${item.NIP2 || '19800526 200901 1 003'}`, size: 24 }) ], alignment: AlignmentType.CENTER })
                    ],
                    borders: noBorders
                }),
                new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                        new Paragraph({ children: [ new TextRun({ text: "PIHAK KESATU", bold: true, size: 24 }) ], alignment: AlignmentType.CENTER }),
                        new Paragraph({ text: "", spacing: { after: 1200 } }),
                        new Paragraph({ children: [ new TextRun({ text: item.PIHAK1 || 'KUSUMA ATMADJA, S.Sos', bold: true, size: 24, underline: {} }) ], alignment: AlignmentType.CENTER }),
                        new Paragraph({ children: [ new TextRun({ text: `NIP. ${item.NIP1 || '197001021993021003'}`, size: 24 }) ], alignment: AlignmentType.CENTER })
                    ],
                    borders: noBorders
                })
            ]
        })
    ]
}));


        // ========== HALAMAN 3 - DOKUMENTASI (MULTIPLE IMAGES) ==========
        children.push(new Paragraph({ text: "", pageBreakBefore: true }));
        children.push(new Paragraph({ 
            children: [ new TextRun({ text: "Dokumentasi", bold: true, size: 28 }) ], 
            alignment: AlignmentType.CENTER, 
            spacing: { after: 400 } 
        }));

        // Kumpulkan semua gambar
        const allImages = [];


    items.forEach((subItem, idx) => {
    const itemNumber = idx + 1;

   if (subItem.Image) {
    allImages.push({
        source: subItem.Image,
        label: `Gambar ${itemNumber} `,
    });
}

if (subItem.Image2) {
    allImages.push({
        source: subItem.Image2,
        label: `Gambar ${itemNumber}`,
    });
}
});

        // Tampilkan semua gambar
        if (allImages.length > 0) {
            for (let i = 0; i < allImages.length; i++) {
                const imageObj = allImages[i];
                
                try {
                    console.log(`🔄 Memproses gambar ${i + 1}/${allImages.length}: ${imageObj.label}`);
                    
                    // Validasi gambar adalah base64
                    if (imageObj.source && 
                        typeof imageObj.source === 'string' && 
                        imageObj.source.startsWith('data:image')) {
                        
                        // Label gambar
                        children.push(new Paragraph({ 
                            children: [ 
                                new TextRun({ 
                                    text: imageObj.label,
                                    bold: true, 
                                    size: 22,
                                    color: "333333"
                                }) 
                            ], 
                            alignment: AlignmentType.CENTER, 
                            spacing: { before: 300, after: 150 } 
                        }));
                        
                        // Convert base64 to ArrayBuffer
                        const imageBuffer = dataURLtoArrayBuffer(imageObj.source);
                        
                        // Insert gambar
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
                        
                        console.log(`✅ Gambar ${i + 1} berhasil dimasukkan: ${imageObj.label}`);
                        
                    } else {
                        console.warn(`⚠️ Gambar ${i + 1} bukan format base64 valid`);
                        children.push(new Paragraph({ 
                            children: [ 
                                new TextRun({ 
                                    text: `[Format tidak valid: ${imageObj.label}]`, 
                                    size: 20, 
                                    italics: true, 
                                    color: "FF6600" 
                                }) 
                            ], 
                            alignment: AlignmentType.CENTER, 
                            spacing: { after: 200 } 
                        }));
                    }
                    
                } catch (errImg) {
                    console.error(`❌ Error gambar ${i + 1}:`, errImg);
                    children.push(new Paragraph({ 
                        children: [ 
                            new TextRun({ 
                                text: `[Error loading: ${imageObj.label}]`, 
                                size: 20, 
                                italics: true, 
                                color: "FF0000" 
                            }) 
                        ], 
                        alignment: AlignmentType.CENTER, 
                        spacing: { after: 200 } 
                    }));
                }
            }
        } else {
            console.warn("⚠️ Tidak ada gambar yang ditemukan");
            children.push(new Paragraph({ 
                children: [ 
                    new TextRun({ 
                        text: "(Tidak ada dokumentasi foto)", 
                        size: 22, 
                        italics: true, 
                        color: "999999" 
                    }) 
                ], 
                alignment: AlignmentType.CENTER, 
                spacing: { after: 200 } 
            }));
        }

        // ========== CREATE DOCUMENT ==========
        const doc = new Document({
            styles: {
                default: {
                    document: {
                        run: {
                            font: "Times New Roman",
                            size: 24
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
        const safeName = (item.Deskripsi || `${id}`).replace(/[\\\/:*?"<>|]/g, '');
        const fileName = `BAP_${safeName}.docx`;

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        link.click();
        URL.revokeObjectURL(link.href);

        console.log(`✅ File Word berhasil dibuat: ${fileName}`);

    } catch (err) {
        console.error("❌ Gagal export ke Word:", err);
        alert("❌ Gagal membuat file Word. Cek console untuk detail error.");
    }
}

// ----------------- Init & Event Listeners -----------------
function initTableSizeListener() {
    const tableSize = document.getElementById('table_size');
    if (tableSize) {
        tableSize.addEventListener('change', (e) => {
            entriesPerPage = parseInt(e.target.value) || 10;
            currentPage = 1;
            renderTable();
            updatePagination();
        });
    }
}

function initEventListeners() {
    const addBtn = document.getElementById('addBtn');
    if (addBtn) {
        addBtn.addEventListener('click', showAddModal);
    }

    const downloadBtn = document.getElementById('downloadBtn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', syncToGoogleDrive);
    }

    const searchInput = document.getElementById('brId');
    if (searchInput) {
        searchInput.addEventListener('input', filterTable);
    }

    const modal = document.getElementById('modal');
    if (modal) {
        window.addEventListener('click', (e) => {
            if (e.target === modal) {
                closeModal();
            }
        });
    }

    const saveBtn = document.getElementById('saveBtn');
    if (saveBtn) saveBtn.addEventListener('click', saveData);

    const closeBtn = document.getElementById('closeModalBtn');
    if (closeBtn) closeBtn.addEventListener('click', closeModal);
}


// Expose globals for inline handlers
window.editData = editData;
window.deleteData = deleteData;
window.exportToWord = exportToWord;
window.changePage = changePage;
window.removeItemRow = removeItemRow;
window.calculateItemTotal = calculateItemTotal;
window.saveData = saveData;
window.closeModal = closeModal;
window.filterTable = filterTable;
window.showAddModal = showAddModal;
window.handleItemImageUpload = handleItemImageUpload;
window.removeItemImage = removeItemImage;


// ----------------- App Init -----------------
async function initApp() {
    try {
        console.log('Initializing BAP Generator...');
        await initIndexedDB();
        await loadData();
        initImageListener();
        initTableSizeListener();
        initEventListeners();
        console.log('✅ BAP Generator initialized successfully!');
    } catch (error) {
        console.error('Init error:', error);
        alert('❌ Gagal inisialisasi aplikasi: ' + (error.message || error));
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}
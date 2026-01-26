// ====== SISTEM MANAJEMEN BAP - VERSI 4 ======

let db;
const DB_NAME = 'SistemBAPDatabase_v4';
const DB_VERSION = 1;
const STORE_NAME = 'dataBAPRecords';

let dataList = [];
let editMode = null;
let currentPage = 1;
let itemsPerPage = 10;
let barangItems = [];
let isInitialized = false;
let isSaving = false; // Flag untuk mencegah save ganda

// ========== OPERASI DATABASE ==========
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

// ========== FUNGSI UTILITAS ==========
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

// ========== RENDER TABEL ==========
function tampilkanTabel(filterData = null) {
    const tbody = document.querySelector('.userBarang4');
    if (!tbody) return;

    const data = filterData || dataList;

    if (!Array.isArray(data) || data.length === 0) {
        const pesan = dataList.length === 0 
            ? '📦 Belum ada data tersimpan' 
            : '🔍 Pencarian tidak ditemukan';
        tbody.innerHTML = `<tr><td colspan="13" style="text-align:center;padding:30px;color:#888;font-size:16px;">${pesan}</td></tr>`;
        updateInfo(0, 0, dataList.length);
        return;
    }

    const start = (currentPage - 1) * itemsPerPage;
    const end = start + itemsPerPage;
    const pageData = data.slice(start, end);

    const rows = pageData.map((row) => {
        const items = row.items || [];
        
        const namaBarangList = items.length > 0 
            ? items.map((item, idx) => `<div style="padding:4px 0;">${idx + 1}. ${escapeHTML(item.NamaBarang || '-')}</div>`).join('')
            : '<span style="color:#999;">Kosong</span>';
        
        const jumlahList = items.length > 0 
            ? items.map((item, idx) => {
                const qty = escapeHTML(item.Kuantitas || '-');
                const unit = escapeHTML(item.Satuan || '');
                return `<div style="padding:4px 0;">${idx + 1}. ${qty} ${unit}</div>`.trim();
              }).join('')
            : '<span style="color:#999;">Kosong</span>';
        
        let gambarHTML = '<span style="color:#999;">Tidak ada gambar</span>';
        const adaGambar = items.some(item => item.Image);
        
        if (adaGambar) {
            gambarHTML = '<div style="display:flex;flex-wrap:wrap;gap:6px;">';
            items.forEach((item, i) => {
                if (item.Image) {
                    gambarHTML += `
                        <div style="position:relative;width:55px;height:55px;">
                            <img src="${item.Image}" 
                                 style="width:100%;height:100%;object-fit:cover;border-radius:4px;border:2px solid #3F51B5;cursor:pointer;box-shadow:0 1px 3px rgba(0,0,0,0.1);"
                                 onclick="window.open('${item.Image}', '_blank')"
                                 title="Klik untuk zoom">
                            <span style="position:absolute;bottom:1px;right:1px;background:#3F51B5;color:white;font-size:10px;padding:1px 4px;border-radius:2px;font-weight:bold;">${i + 1}</span>
                        </div>`;
                }
            });
            gambarHTML += '</div>';
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
                <td style="padding:10px;color:#555;">${escapeHTML(row.PIHAK2 || '-')}</td>
                <td style="padding:10px;color:#666;font-size:13px;">${escapeHTML(row.NIP2 || '-')}</td>
                <td style="padding:10px;color:#666;font-size:13px;">${escapeHTML(row.JBT2 || '-')}</td>
                <td style="padding:10px;max-width:200px;">
                    <div style="max-height:140px;overflow-y:auto;font-size:0.9em;line-height:1.6;">${namaBarangList}</div>
                </td>
                <td style="padding:10px;max-width:140px;">
                    <div style="max-height:140px;overflow-y:auto;font-size:0.9em;line-height:1.6;text-align:center;">${jumlahList}</div>
                </td>
                <td style="padding:10px;font-weight:500;color:#3F51B5;">${escapeHTML(row.NOBAP || '-')}</td>
                <td style="padding:10px;text-align:center;">${gambarHTML}</td>
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
                    <button onclick="buatDokumen(${row.id})" 
                            style="padding:7px 12px;margin:2px;background:#2196F3;color:white;border:none;border-radius:4px;cursor:pointer;transition:all 0.15s;font-size:13px;"
                            onmouseover="this.style.transform='scale(1.05)'"
                            onmouseout="this.style.transform='scale(1)'">
                        📄 Cetak
                    </button>
                </td>
            </tr>`;
    }).join('');

    tbody.innerHTML = rows;
    updateInfo(start + 1, Math.min(end, data.length), data.length);
}

function updateInfo(mulai, akhir, total) {
    const info = document.querySelector('.showEntries4');
    if (info) {
        info.textContent = `Menampilkan ${mulai} hingga ${akhir} dari ${total} entri`;
    }
}

function buatPagination() {
    const totalHalaman = Math.max(1, Math.ceil(dataList.length / itemsPerPage));
    const container = document.querySelector('.pagination4');
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

// ========== FITUR PENCARIAN ==========
function filterTable() {
    const input = document.getElementById('brId4');
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
            (row.NOBAP || '').toLowerCase().includes(keyword);
        
        const itemsCocok = (row.items || []).some(item => 
            (item.NamaBarang || '').toLowerCase().includes(keyword) ||
            (item.Kuantitas || '').toString().toLowerCase().includes(keyword) ||
            (item.Satuan || '').toLowerCase().includes(keyword)
        );
        
        return fieldUtama || itemsCocok;
    });
    
    currentPage = 1;
    tampilkanTabel(hasil);
}

// ========== OPERASI MODAL ==========
function openModal() {
    editMode = null;
    
    const judul = document.getElementById('modalTitle');
    if (judul) {
        judul.textContent = '➕ Add BAP';
    }
    
    bersihkanForm();

    const modal = document.getElementById('modal4');
    if (modal) {
        modal.style.display = 'block';
    }

    setTimeout(() => {
        if (barangItems.length === 0) {
            tambahBarisBarang();
        }
    }, 100);
}

function closeModal() {
    const modal = document.getElementById('modal4');
    if (modal) {
        modal.style.display = 'none';
    }
    bersihkanForm();
}

function bersihkanForm() {
    const fields = ['NOS4', 'TGL4', 'PIHAK1t4', 'NIP1t4', 'JBT1t4', 'PIHAK2t4', 'NIP2t4', 'JBTt4', 'NOBAP4'];
    
    fields.forEach(id => {
        const field = document.getElementById(id);
        if (field) field.value = '';
    });

    const container = document.getElementById('itemsContainer');
    if (container) container.innerHTML = '';

    const btn = document.getElementById('addItemBtn');
    if (btn) btn.remove();

    barangItems = [];
}

// ========== MANAJEMEN ITEM BARANG ==========
function tambahBarisBarang() {
    const container = document.getElementById('itemsContainer');
    if (!container) return;

    const index = barangItems.length;

    const div = document.createElement('div');
    div.className = 'item-row';
    div.style.cssText = 'border:2px solid #ccc;padding:18px;margin-bottom:18px;border-radius:8px;background:#f9f9f9;box-shadow:0 1px 3px rgba(0,0,0,0.05);';
    
    div.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;padding-bottom:10px;border-bottom:2px solid #ddd;">
            <h4 style="margin:0;color:#3F51B5;font-size:17px;font-weight:600;">📦 Item Barang #${index + 1}</h4>
            ${index > 0 ? `
                <button type="button" 
                        onclick="hapusBarisBarang(${index})" 
                        style="background:#F44336;color:white;border:none;padding:5px 12px;border-radius:5px;cursor:pointer;font-weight:500;transition:all 0.15s;font-size:13px;"
                        onmouseover="this.style.background='#D32F2F'"
                        onmouseout="this.style.background='#F44336'">
                    ✖ Hapus
                </button>` : ''}
        </div>
        
        <div style="margin-bottom:14px;">
            <label style="display:block;margin-bottom:5px;font-weight:600;color:#222;font-size:14px;">Nama Barang *</label>
            <input type="text" 
                   id="NamaBarang_${index}" 
                   placeholder="Contoh: Komputer Desktop HP" 
                   style="width:100%;padding:9px;border:2px solid #ccc;border-radius:5px;font-size:13px;transition:border-color 0.15s;"
                   onfocus="this.style.borderColor='#3F51B5'"
                   onblur="this.style.borderColor='#ccc'">
        </div>
        
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;">
            <div>
                <label style="display:block;margin-bottom:5px;font-weight:600;color:#222;font-size:14px;">Jumlah *</label>
                <input type="number" 
                       id="Kuantitas_${index}" 
                       placeholder="5" 
                       min="1"
                       style="width:100%;padding:9px;border:2px solid #ccc;border-radius:5px;font-size:13px;transition:border-color 0.15s;"
                       onfocus="this.style.borderColor='#3F51B5'"
                       onblur="this.style.borderColor='#ccc'">
            </div>
            <div>
                <label style="display:block;margin-bottom:5px;font-weight:600;color:#222;font-size:14px;">Satuan *</label>
                <input type="text" 
                       id="Satuan_${index}" 
                       placeholder="Unit/Set/Pcs" 
                       style="width:100%;padding:9px;border:2px solid #ccc;border-radius:5px;font-size:13px;transition:border-color 0.15s;"
                       onfocus="this.style.borderColor='#3F51B5'"
                       onblur="this.style.borderColor='#ccc'">
            </div>
        </div>
        
        <div style="border-top:2px solid #ddd;padding-top:14px;background:white;padding:14px;border-radius:5px;">
            <label style="display:block;margin-bottom:9px;font-weight:600;color:#3F51B5;font-size:14px;">📸 Unggah Foto Barang</label>
            <input type="file" 
                   id="Image_${index}" 
                   accept="image/*" 
                   onchange="prosesGambar(${index}, event)" 
                   style="width:100%;padding:9px;border:2px dashed #3F51B5;border-radius:5px;background:#E8EAF6;cursor:pointer;font-size:13px;">
            <div id="Preview_${index}" style="display:none;position:relative;margin-top:10px;">
                <img id="Img_${index}" 
                     style="width:100%;max-width:220px;border-radius:6px;border:3px solid #3F51B5;box-shadow:0 3px 6px rgba(0,0,0,0.1);">
                <button type="button" 
                        onclick="hapusGambar(${index})" 
                        style="position:absolute;top:6px;right:6px;background:#F44336;color:white;border:none;border-radius:50%;width:28px;height:28px;cursor:pointer;font-size:16px;font-weight:bold;box-shadow:0 2px 4px rgba(0,0,0,0.2);"
                        title="Hapus foto">
                    ×
                </button>
            </div>
        </div>
    `;

    container.appendChild(div);
    barangItems.push({ Image: null });
    refreshTombolTambah();
}

function hapusBarisBarang(index) {
    if (barangItems.length <= 1) {
        alert('⚠️ Minimal harus ada 1 item barang!');
        return;
    }

    const container = document.getElementById('itemsContainer');
    if (container.children[index]) {
        container.children[index].remove();
    }

    barangItems.splice(index, 1);

    const backup = [...barangItems];
    barangItems = [];
    container.innerHTML = '';
    
    const btn = document.getElementById('addItemBtn');
    if (btn) btn.remove();
    
    backup.forEach(() => tambahBarisBarang());
}

function prosesGambar(index, event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) {
        alert('⚠️ File terlalu besar! Maksimal 5MB');
        event.target.value = '';
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
        barangItems[index].Image = e.target.result;
        
        const preview = document.getElementById(`Preview_${index}`);
        const img = document.getElementById(`Img_${index}`);
        
        if (preview && img) {
            img.src = e.target.result;
            preview.style.display = 'block';
        }
    };
    reader.readAsDataURL(file);
}

function hapusGambar(index) {
    barangItems[index].Image = null;
    
    const input = document.getElementById(`Image_${index}`);
    const preview = document.getElementById(`Preview_${index}`);
    
    if (input) input.value = '';
    if (preview) preview.style.display = 'none';
}

function refreshTombolTambah() {
    let btn = document.getElementById('addItemBtn');
    
    if (!btn) {
        btn = document.createElement('button');
        btn.id = 'addItemBtn';
        btn.type = 'button';
        btn.style.cssText = 'width:100%;padding:13px;background:linear-gradient(135deg,#5C6BC0 0%,#7E57C2 100%);color:white;border:none;border-radius:8px;margin-top:18px;cursor:pointer;font-weight:bold;font-size:15px;transition:all 0.25s;box-shadow:0 3px 6px rgba(0,0,0,0.12);';
        btn.innerHTML = '➕ Tambah Barang Lagi';
        btn.onclick = tambahBarisBarang;
        
        btn.onmouseover = function() {
            this.style.transform = 'translateY(-1px)';
            this.style.boxShadow = '0 5px 10px rgba(0,0,0,0.18)';
        };
        btn.onmouseout = function() {
            this.style.transform = 'translateY(0)';
            this.style.boxShadow = '0 3px 6px rgba(0,0,0,0.12)';
        };

        const container = document.getElementById('itemsContainer');
        if (container && container.parentNode) {
            container.parentNode.insertBefore(btn, container.nextSibling);
        }
    }
}

// ========== OPERASI CRUD ==========
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
    // CRITICAL: Cek flag saving untuk mencegah double save
    if (isSaving) {
        console.log('⚠️ Sedang menyimpan, harap tunggu...');
        return;
    }
    
    isSaving = true; // Set flag
    
    try {
        const items = [];
        
        for (let i = 0; i < barangItems.length; i++) {
            const nama = document.getElementById(`NamaBarang_${i}`)?.value?.trim() || '';
            const jumlah = document.getElementById(`Kuantitas_${i}`)?.value?.trim() || '';
            const satuan = document.getElementById(`Satuan_${i}`)?.value?.trim() || '';

            if (!jumlah || !satuan) {
                alert(`⚠️ Item #${i + 1}: Jumlah dan Satuan harus diisi!`);
                isSaving = false; // Reset flag
                return;
            }

            items.push({
                NamaBarang: nama,
                Kuantitas: jumlah,
                Satuan: satuan,
                Image: barangItems[i].Image || null
            });
        }

        const data = {
            NOS: document.getElementById('NOS4')?.value?.trim() || '',
            TGL: document.getElementById('TGL4')?.value || '',
            PIHAK1: document.getElementById('PIHAK1t4')?.value?.trim() || '',
            NIP1: document.getElementById('NIP1t4')?.value?.trim() || '',
            JBT1: document.getElementById('JBT1t4')?.value?.trim() || '',
            PIHAK2: document.getElementById('PIHAK2t4')?.value?.trim() || '',
            NIP2: document.getElementById('NIP2t4')?.value?.trim() || '',
            JBT2: document.getElementById('JBTt4')?.value?.trim() || '',
            NOBAP: document.getElementById('NOBAP4')?.value?.trim() || '',
            items: items
        };

        if (!data.NOS) {
            alert('⚠️ No Surat wajib diisi!');
            isSaving = false; // Reset flag
            return;
        }
        if (!data.PIHAK1) {
            alert('⚠️ Pihak Kesatu wajib diisi!');
            isSaving = false; // Reset flag
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
        isSaving = false; // Reset flag setelah selesai
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
        judul.textContent = '✏️ Edit Data BAP';
    }

    bersihkanForm();

    const setVal = (fieldId, val) => {
        const field = document.getElementById(fieldId);
        if (field) field.value = val || '';
    };

    setVal('NOS4', data.NOS);
    setVal('TGL4', data.TGL);
    setVal('PIHAK1t4', data.PIHAK1);
    setVal('NIP1t4', data.NIP1);
    setVal('JBT1t4', data.JBT1);
    setVal('PIHAK2t4', data.PIHAK2);
    setVal('NIP2t4', data.NIP2);
    setVal('JBTt4', data.JBT2);
    setVal('NOBAP4', data.NOBAP);

    const items = data.items || [];
    
    setTimeout(() => {
        items.forEach((item, i) => {
            tambahBarisBarang();
            
            setTimeout(() => {
                setVal(`NamaBarang_${i}`, item.NamaBarang);
                setVal(`Kuantitas_${i}`, item.Kuantitas);
                setVal(`Satuan_${i}`, item.Satuan);
                
                if (item.Image) {
                    barangItems[i].Image = item.Image;
                    const img = document.getElementById(`Img_${i}`);
                    const preview = document.getElementById(`Preview_${i}`);
                    if (img && preview) {
                        img.src = item.Image;
                        preview.style.display = 'block';
                    }
                }
            }, 50);
        });
    }, 100);

    const modal = document.getElementById('modal4');
    if (modal) {
        modal.style.display = 'block';
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

// ========== GENERATE DOKUMEN WORD ==========
async function buatDokumen(id) {
    const data = dataList.find(r => r.id === id);
    
    if (!data) {
        alert('❌ Data tidak ditemukan!');
        return;
    }

    try {
        if (!window.docx) {
            alert("⚠️ Library docx belum tersedia. Refresh halaman!");
            return;
        }

        function angkaKeKata(n) {
            const satuan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
            
            if (n < 12) return satuan[n];
            if (n < 20) return angkaKeKata(n - 10) + ' belas';
            if (n < 100) return angkaKeKata(Math.floor(n / 10)) + ' puluh ' + angkaKeKata(n % 10);
            if (n < 200) return 'seratus ' + angkaKeKata(n - 100);
            if (n < 1000) return angkaKeKata(Math.floor(n / 100)) + ' ratus ' + angkaKeKata(n % 100);
            if (n < 2000) return 'seribu ' + angkaKeKata(n - 1000);
            if (n < 1000000) return angkaKeKata(Math.floor(n / 1000)) + ' ribu ' + angkaKeKata(n % 1000);
            
            return n.toString();
        }

        function kapitalAwal(str) {
            if (!str) return str;
            return str.charAt(0).toUpperCase() + str.slice(1);
        }

        function base64ToUint8Array(base64) {
            const parts = base64.split(',');
            const decoded = atob(parts[1]);
            let len = decoded.length;
            const arr = new Uint8Array(len);
            while (len--) {
                arr[len] = decoded.charCodeAt(len);
            }
            return arr;
        }

        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
                WidthType, AlignmentType, BorderStyle, ImageRun, VerticalAlign } = docx;

        const tgl = data.TGL ? new Date(data.TGL) : new Date();
        const hari = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];
        const bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
                       'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        
        const namaHari = hari[tgl.getDay()];
        const angkaTanggal = tgl.getDate();
        const namaBulan = bulan[tgl.getMonth()];
        const angkaTahun = tgl.getFullYear();

        const kataTanggal = kapitalAwal(angkaKeKata(angkaTanggal) || String(angkaTanggal));
        const kataTahun = kapitalAwal(angkaKeKata(angkaTahun) || String(angkaTahun));

        const border = {
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

        const content = [];

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "PEMERINTAH PROVINSI JAWA TIMUR", bold: true, size: 24 }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({ 
                children: [ new TextRun({ text: "DINAS SOSIAL", bold: true, size: 24 }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({ 
                children: [ new TextRun({ text: "UPT REHABILITASI SOSIAL BINA GRAHITA TUBAN", bold: true, size: 22 }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({ 
                children: [ new TextRun({ text: "Jl. Teuku Umar No. 116, Tuban 63215", size: 20, italics: true }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 50 }
            }),
            new Paragraph({ 
                children: [ new TextRun({ text: "Email: uptrsbgtuban@gmail.com | Telp: (0356) 321234", size: 20, italics: true }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            })
        );

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "═".repeat(85), size: 20 }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );

        content.push(
            new Paragraph({
                children: [ new TextRun({ text: "BERITA ACARA PENYERAHAN BARANG", bold: true, size: 26, underline: {} }) ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 150 }
            }),
            new Paragraph({ 
                children: [ new TextRun({ text: `Nomor: ${data.NOS || '-'}`, bold: true, size: 22 }) ],
                alignment: AlignmentType.CENTER, 
                spacing: { after: 400 } 
            })
        );

        content.push(
            new Paragraph({
                children: [ 
                    new TextRun({ text: "Pada hari ini, ", size: 22 }),
                    new TextRun({ text: namaHari, bold: true, size: 22 }),
                    new TextRun({ text: ", tanggal ", size: 22 }),
                    new TextRun({ text: kataTanggal, bold: true, size: 22 }),
                    new TextRun({ text: " bulan ", size: 22 }),
                    new TextRun({ text: namaBulan, bold: true, size: 22 }),
                    new TextRun({ text: " tahun ", size: 22 }),
                    new TextRun({ text: kataTahun, bold: true, size: 22 }),
                    new TextRun({ text: ", yang bertandatangan di bawah ini:", size: 22 })
                ],
                spacing: { after: 300 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "PIHAK PERTAMA:", bold: true, size: 22, underline: {} }) ],
                spacing: { before: 200, after: 150 }
            })
        );

        content.push(
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 25, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 5, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: data.PIHAK1 || '', bold: true, size: 22 })] })],
                                borders: noBorder,
                                width: { size: 70, type: WidthType.PERCENTAGE }
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
                                children: [new Paragraph({ children: [new TextRun({ text: data.NIP1 || '', size: 22 })] })],
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
                                children: [new Paragraph({ children: [new TextRun({ text: data.JBT1 || "Pengurus Barang", size: 22 })] })],
                                borders: noBorder
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

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "Selanjutnya disebut sebagai PIHAK PERTAMA.", italics: true, size: 22 }) ],
                spacing: { before: 100, after: 300 }
            })
        );

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "PIHAK KEDUA:", bold: true, size: 22, underline: {} }) ],
                spacing: { before: 200, after: 150 }
            })
        );

        content.push(
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: "Nama", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 25, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: ":", size: 22 })] })],
                                borders: noBorder,
                                width: { size: 5, type: WidthType.PERCENTAGE }
                            }),
                            new TableCell({
                                children: [new Paragraph({ children: [new TextRun({ text: data.PIHAK2 || '', bold: true, size: 22 })] })],
                                borders: noBorder,
                                width: { size: 70, type: WidthType.PERCENTAGE }
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
                                children: [new Paragraph({ children: [new TextRun({ text: data.NIP2 || '', size: 22 })] })],
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
                                children: [new Paragraph({ children: [new TextRun({ text: data.JBT2 || "", size: 22 })] })],
                                borders: noBorder
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

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "Selanjutnya disebut sebagai PIHAK KEDUA.", italics: true, size: 22 }) ],
                spacing: { before: 100, after: 400 }
            })
        );

        content.push(
            new Paragraph({
                children: [
                    new TextRun({ text: "Dengan ini PIHAK PERTAMA telah menyerahkan barang-barang kepada PIHAK KEDUA dengan rincian sebagai berikut:", size: 22 })
                ],
                spacing: { before: 300, after: 300 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        const items = data.items || [];
        const itemRows = [];

        itemRows.push(
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "NO", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 8, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: border,
                        shading: { fill: "E0E0E0" }
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "URAIAN BARANG", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 52, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: border,
                        shading: { fill: "E0E0E0" }
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "JUMLAH", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: border,
                        shading: { fill: "E0E0E0" }
                    }),
                    new TableCell({
                        children: [new Paragraph({ 
                            children: [new TextRun({ text: "SATUAN", bold: true, size: 20 })], 
                            alignment: AlignmentType.CENTER
                        })],
                        width: { size: 20, type: WidthType.PERCENTAGE },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: border,
                        shading: { fill: "E0E0E0" }
                    })
                ]
            })
        );

        items.forEach((item, idx) => {
            itemRows.push(
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: String(idx + 1), size: 20 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: border
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: item.NamaBarang || '', size: 20 })],
                                alignment: AlignmentType.LEFT
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: border
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: String(item.Kuantitas || ''), size: 20 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: border
                        }),
                        new TableCell({
                            children: [new Paragraph({ 
                                children: [new TextRun({ text: item.Satuan || '', size: 20 })], 
                                alignment: AlignmentType.CENTER
                            })],
                            verticalAlign: VerticalAlign.CENTER,
                            borders: border
                        })
                    ]
                })
            );
        });

        content.push(
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: itemRows
            })
        );

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: "Demikian Berita Acara ini dibuat dengan sebenarnya untuk dipergunakan sebagaimana mestinya.", size: 22 }) ],
                spacing: { before: 400, after: 600 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        content.push(
            new Paragraph({ 
                children: [ new TextRun({ text: `Tuban, ${angkaTanggal} ${namaBulan} ${angkaTahun}`, size: 22 }) ],
                alignment: AlignmentType.RIGHT,
                spacing: { after: 400 }
            })
        );

        content.push(new Table({
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
                                    children: [ new TextRun({ text: "Yang Menyerahkan,", size: 22 }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: "PIHAK PERTAMA", bold: true, size: 22 }) ],
                                    alignment: AlignmentType.CENTER,
                                    spacing: { after: 100 }
                                }),
                                new Paragraph({ text: "", spacing: { after: 1000 } }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: data.PIHAK1 || '', bold: true, size: 22, underline: {} }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: `NIP. ${data.NIP1 || ''}`, size: 20 }) ],
                                    alignment: AlignmentType.CENTER 
                                })
                            ],
                            borders: noBorder
                        }),
                        new TableCell({
                            width: { size: 50, type: WidthType.PERCENTAGE },
                            children: [
                                new Paragraph({ 
                                    children: [ new TextRun({ text: "Yang Menerima,", size: 22 }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: "PIHAK KEDUA", bold: true, size: 22 }) ],
                                    alignment: AlignmentType.CENTER,
                                    spacing: { after: 100 }
                                }),
                                new Paragraph({ text: "", spacing: { after: 1000 } }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: data.PIHAK2 || '', bold: true, size: 22, underline: {} }) ],
                                    alignment: AlignmentType.CENTER 
                                }),
                                new Paragraph({ 
                                    children: [ new TextRun({ text: `NIP. ${data.NIP2 || ''}`, size: 20 }) ],
                                    alignment: AlignmentType.CENTER 
                                })
                            ],
                            borders: noBorder
                        })
                    ]
                })
            ]
        }));

        if (items.some(item => item.Image)) {
            content.push(new Paragraph({ text: "", pageBreakBefore: true }));
            content.push(
                new Paragraph({ 
                    children: [ new TextRun({ text: "LAMPIRAN DOKUMENTASI", bold: true, size: 26, underline: {} }) ], 
                    alignment: AlignmentType.CENTER, 
                    spacing: { after: 400 } 
                })
            );

            for (let i = 0; i < items.length; i++) {
                const item = items[i];
                
                if (item.Image && typeof item.Image === 'string' && item.Image.startsWith('data:image')) {
                    try {
                        content.push(new Paragraph({ 
                            children: [ 
                                new TextRun({ 
                                    text: `Foto ${i + 1}: ${item.NamaBarang || 'Barang ' + (i + 1)}`,
                                    bold: true, 
                                    size: 22
                                }) 
                            ], 
                            spacing: { before: 300, after: 150 } 
                        }));
                        
                        const imgBuffer = base64ToUint8Array(item.Image);
                        
                        const img = new ImageRun({
                            data: imgBuffer,
                            transformation: { 
                                width: 400,
                                height: 300 
                            }
                        });
                        
                        content.push(new Paragraph({ 
                            children: [img], 
                            alignment: AlignmentType.CENTER, 
                            spacing: { before: 100, after: 300 } 
                        }));
                    } catch (err) {
                        console.error(`Error gambar ${i + 1}:`, err);
                    }
                }
            }
        }

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
                children: content
            }]
        });

        const blob = await Packer.toBlob(doc);
        const filename = `BAP_${data.NOS || id}.docx`;

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);

        console.log(`✅ Dokumen berhasil: ${filename}`);

    } catch (error) {
        console.error("❌ Error:", error);
        alert("❌ Gagal membuat dokumen: " + error.message);
    }
}

// ========== INISIALISASI ==========
async function init() {
    if (isInitialized) {
        console.log('⚠️ Sistem sudah diinisialisasi');
        return;
    }

    try {
        await initDB();
        await muatUlangData();

        // Hapus semua event listener lama dengan clone
        const btnAdd = document.getElementById('addBtn4');
        if (btnAdd) {
            const newBtn = btnAdd.cloneNode(true);
            btnAdd.parentNode.replaceChild(newBtn, btnAdd);
            newBtn.addEventListener('click', openModal);
        }

        const btnClose = document.getElementById('closeModal4');
        if (btnClose) {
            const newBtn = btnClose.cloneNode(true);
            btnClose.parentNode.replaceChild(newBtn, btnClose);
            newBtn.addEventListener('click', closeModal);
        }

        const btnSave = document.getElementById('submitBtn4');
        if (btnSave) {
            const newBtn = btnSave.cloneNode(true);
            btnSave.parentNode.replaceChild(newBtn, btnSave);
            newBtn.addEventListener('click', saveData);
        }

        const selectSize = document.getElementById('table_size4');
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

        const btnSearch = document.querySelector('.search4');
        if (btnSearch) {
            const newBtn = btnSearch.cloneNode(true);
            btnSearch.parentNode.replaceChild(newBtn, btnSearch);
            newBtn.addEventListener('click', filterTable);
        }

        const inputSearch = document.getElementById('brId4');
        if (inputSearch) {
            const newInput = inputSearch.cloneNode(true);
            inputSearch.parentNode.replaceChild(newInput, inputSearch);
            newInput.addEventListener('keyup', filterTable);
        }

        window.onclick = (e) => {
            const modal = document.getElementById('modal4');
            if (e.target === modal) {
                closeModal();
            }
        };

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
window.buatDokumen = buatDokumen;
window.pindahHalaman = pindahHalaman;
window.hapusBarisBarang = hapusBarisBarang;
window.prosesGambar = prosesGambar;
window.hapusGambar = hapusGambar;
window.saveData = saveData;
window.closeModal = closeModal;
window.filterTable = filterTable;
window.openModal = openModal;

// Start
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
} else {
    init();
} 
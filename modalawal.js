 // Get elements
        const btnMulai = document.getElementById('btnMulai');
        const btnClose = document.getElementById('btnClose');
        const modalDokumen = document.getElementById('modalDokumen');

        // Show modal
        function tampilkanModal() {
            console.log('Menampilkan modal...');
            modalDokumen.classList.add('active');
            document.body.style.overflow = 'hidden';
        }

        // Hide modal
        function tutupModal() {
            console.log('Menutup modal...');
            modalDokumen.classList.remove('active');
            document.body.style.overflow = 'auto';
        }

        // Select document
        function pilihDokumen(halaman) {
            console.log('Memilih dokumen:', halaman);
            window.location.href = halaman;
        }

        // Event listeners
        btnMulai.addEventListener('click', tampilkanModal);
        btnClose.addEventListener('click', tutupModal);

        // Close when clicking outside
        modalDokumen.addEventListener('click', function(e) {
            if (e.target === modalDokumen) {
                tutupModal();
            }
        });

        // Close on ESC key
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' && modalDokumen.classList.contains('active')) {
                tutupModal();
            }
        });

        // Debug
        console.log('Script loaded successfully');
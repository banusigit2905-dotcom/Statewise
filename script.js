if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const tableBody = document.getElementById('tableBody');
    const clearBtn = document.getElementById('clearData');
    const pdfStatus = document.getElementById('pdfStatus');

    // Elemen modal pratinjau PDF
    const pdfPreviewModal = document.getElementById('pdfPreviewModal');
    const pdfPreviewInfo = document.getElementById('pdfPreviewInfo');
    const pdfPreviewBody = document.getElementById('pdfPreviewBody');
    const pdfPreviewTableWrap = document.getElementById('pdfPreviewTableWrap');
    const pdfRawTextWrap = document.getElementById('pdfRawTextWrap');
    const pdfRawText = document.getElementById('pdfRawText');
    const pdfCancelBtn = document.getElementById('pdfCancelBtn');
    const pdfImportBtn = document.getElementById('pdfImportBtn');

    let transactions = JSON.parse(localStorage.getItem('statwise_data')) || [];
    let pendingPdfRows = []; // hasil ekstraksi PDF yang sedang dipratinjau

    // Tampilkan data saat pertama kali load
    renderData();

    // Event listener untuk upload file
    fileInput.addEventListener('change', handleFileUpload);

    // Hapus Semua Data
    clearBtn.addEventListener('click', () => {
        if (confirm('Hapus semua data transaksi?')) {
            transactions = [];
            saveAndRender();
        }
    });

    function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        const nameLower = file.name.toLowerCase();

        if (nameLower.endsWith('.pdf')) {
            handlePdfUpload(file);
            return;
        }

        const reader = new FileReader();
        reader.onload = (event) => {
            const data = event.target.result;
            let workbook;

            if (nameLower.endsWith('.csv')) {
                const text = event.target.result;
                processData(csvToArray(text));
            } else {
                workbook = XLSX.read(data, { type: 'binary' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                processData(json);
            }
        };

        if (nameLower.endsWith('.csv')) {
            reader.readAsText(file);
        } else {
            reader.readAsBinaryString(file);
        }
    }

    // ================== PDF HANDLING ==================

    async function handlePdfUpload(file) {
        if (!window.pdfjsLib) {
            alert('Gagal memuat library PDF. Periksa koneksi internet Anda dan coba lagi.');
            fileInput.value = '';
            return;
        }

        showPdfStatus('Membaca file PDF, mohon tunggu...');

        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

            let allLines = [];
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                showPdfStatus(`Memproses halaman ${pageNum} dari ${pdf.numPages}...`);
                const page = await pdf.getPage(pageNum);
                const content = await page.getTextContent();
                const lines = groupTextItemsIntoLines(content.items);
                allLines = allLines.concat(lines);
            }

            const fullText = allLines.join('\n');
            const parsedRows = parsePdfLines(allLines);

            showPdfStatus('');
            openPdfPreview(parsedRows, fullText);
        } catch (err) {
            console.error(err);
            showPdfStatus('');
            alert('Terjadi kesalahan saat membaca PDF. File mungkin berupa hasil scan (gambar) tanpa teks, atau formatnya tidak didukung.');
        } finally {
            fileInput.value = '';
        }
    }

    function showPdfStatus(msg) {
        if (!msg) {
            pdfStatus.style.display = 'none';
            pdfStatus.innerText = '';
        } else {
            pdfStatus.style.display = 'block';
            pdfStatus.innerText = msg;
        }
    }

    // Menyusun item teks PDF (yang punya koordinat sendiri-sendiri) menjadi baris teks utuh
    function groupTextItemsIntoLines(items) {
        const rows = {};
        items.forEach(item => {
            const y = Math.round(item.transform[5]);
            if (!rows[y]) rows[y] = [];
            rows[y].push(item);
        });

        const sortedY = Object.keys(rows).map(Number).sort((a, b) => b - a);

        return sortedY.map(y => {
            return rows[y]
                .sort((a, b) => a.transform[4] - b.transform[4])
                .map(i => i.str)
                .join(' ')
                .replace(/\s+/g, ' ')
                .trim();
        }).filter(line => line.length > 0);
    }

    // Mencoba mendeteksi baris transaksi (tanggal + deskripsi + nominal) dari teks PDF
    function parsePdfLines(lines) {
        const dateRegex = /^(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|Mei|Jun|Jul|Agu|Sep|Okt|Nov|Des)[a-z]*\s+\d{2,4})/i;
        // Angka bergaya Indonesia, contoh: 1.500.000,00 atau 250000
        const amountRegex = /-?\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|-?\d{4,}(?:[.,]\d{1,2})?/g;

        const results = [];

        lines.forEach(line => {
            const dateMatch = line.match(dateRegex);
            if (!dateMatch) return;

            const date = dateMatch[1];
            let rest = line.slice(dateMatch[0].length).trim();

            const amountMatches = rest.match(amountRegex) || [];
            // Saring token yang terlalu pendek untuk dianggap nominal (misal nomor referensi 2 digit)
            const validAmounts = amountMatches.filter(a => a.replace(/[^\d]/g, '').length >= 3);
            if (validAmounts.length === 0) return;

            let desc = rest;
            validAmounts.forEach(a => {
                desc = desc.replace(a, ' ');
            });
            desc = desc.replace(/\s{2,}/g, ' ').replace(/[|]/g, '').trim();
            if (!desc) desc = '-';

            const lowerLine = line.toLowerCase();
            let debit = 0;
            let credit = 0;

            if (validAmounts.length >= 2) {
                // Asumsi umum e-statement: kolom Debit lalu Kredit
                debit = parseAmountID(validAmounts[0]);
                credit = parseAmountID(validAmounts[1]);
            } else {
                const amt = parseAmountID(validAmounts[0]);
                const isCredit = /\bcr\b|kredit|masuk/.test(lowerLine) || /^\+/.test(validAmounts[0]);
                const isDebit = /\bdb\b|debit|keluar/.test(lowerLine) || /^-/.test(validAmounts[0]);

                if (isCredit && !isDebit) {
                    credit = amt;
                } else {
                    debit = amt;
                }
            }

            results.push({ date, desc, debit, credit });
        });

        return results;
    }

    // Mengonversi angka format Indonesia ("1.500.000,00" atau "250000") menjadi float
    function parseAmountID(str) {
        if (!str) return 0;
        let s = String(str).trim();
        const negative = s.startsWith('-');
        s = s.replace(/^[+-]/, '');

        if (/,\d{1,2}$/.test(s)) {
            // koma sebagai desimal, titik sebagai ribuan
            s = s.replace(/\./g, '').replace(',', '.');
        } else {
            // hanya titik ribuan, tanpa desimal koma
            s = s.replace(/[.,]/g, '');
        }

        const value = parseFloat(s) || 0;
        return negative ? -value : value;
    }

    // ================== PREVIEW MODAL ==================

    function openPdfPreview(rows, rawText) {
        pendingPdfRows = rows.map((r, idx) => ({ ...r, _tempId: idx }));
        renderPdfPreview();
        pdfRawText.value = rawText;

        if (rows.length === 0) {
            pdfPreviewTableWrap.style.display = 'none';
            pdfRawTextWrap.style.display = 'block';
            pdfPreviewInfo.innerText = 'Tidak ada baris transaksi yang terdeteksi secara otomatis.';
            pdfImportBtn.disabled = true;
        } else {
            pdfPreviewTableWrap.style.display = 'block';
            pdfRawTextWrap.style.display = 'none';
            pdfPreviewInfo.innerText = `Ditemukan ${rows.length} transaksi. Periksa dan koreksi jika perlu sebelum diimpor.`;
            pdfImportBtn.disabled = false;
        }

        pdfPreviewModal.style.display = 'flex';
    }

    function renderPdfPreview() {
        pdfPreviewBody.innerHTML = '';
        pendingPdfRows.forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type="text" class="edit-input" value="${escapeHtml(row.date)}" data-field="date" data-id="${row._tempId}"></td>
                <td><input type="text" class="edit-input" value="${escapeHtml(row.desc)}" data-field="desc" data-id="${row._tempId}"></td>
                <td><input type="number" class="edit-input" value="${row.debit}" data-field="debit" data-id="${row._tempId}"></td>
                <td><input type="number" class="edit-input" value="${row.credit}" data-field="credit" data-id="${row._tempId}"></td>
                <td><button class="remove-preview-row" data-id="${row._tempId}">✕</button></td>
            `;
            pdfPreviewBody.appendChild(tr);
        });

        pdfPreviewBody.querySelectorAll('input').forEach(input => {
            input.addEventListener('change', (e) => {
                const id = Number(e.target.dataset.id);
                const field = e.target.dataset.field;
                const row = pendingPdfRows.find(r => r._tempId === id);
                if (!row) return;
                row[field] = (field === 'debit' || field === 'credit') ? (parseFloat(e.target.value) || 0) : e.target.value;
            });
        });

        pdfPreviewBody.querySelectorAll('.remove-preview-row').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = Number(e.target.dataset.id);
                pendingPdfRows = pendingPdfRows.filter(r => r._tempId !== id);
                renderPdfPreview();
                pdfPreviewInfo.innerText = `Ditemukan ${pendingPdfRows.length} transaksi. Periksa dan koreksi jika perlu sebelum diimpor.`;
                pdfImportBtn.disabled = pendingPdfRows.length === 0;
            });
        });
    }

    function escapeHtml(str) {
        const div = document.createElement('div');
        div.innerText = str;
        return div.innerHTML;
    }

    pdfCancelBtn.addEventListener('click', () => {
        pendingPdfRows = [];
        pdfPreviewModal.style.display = 'none';
    });

    pdfImportBtn.addEventListener('click', () => {
        const newData = pendingPdfRows.map(row => ({
            id: Date.now() + Math.random(),
            date: row.date || '-',
            desc: row.desc || '-',
            debit: row.debit || 0,
            credit: row.credit || 0,
            category: 'Umum',
            note: ''
        }));

        transactions = [...transactions, ...newData];
        pendingPdfRows = [];
        pdfPreviewModal.style.display = 'none';
        saveAndRender();
    });

    // ================== CSV/XLSX (existing) ==================

    function processData(rows) {
        const newData = rows.slice(1).filter(row => row.length >= 2).map(row => ({
            id: Date.now() + Math.random(),
            date: row[0] || '-',
            desc: row[1] || '-',
            debit: parseFloat(row[2]) || 0,
            credit: parseFloat(row[3]) || 0,
            category: 'Umum',
            note: ''
        }));

        transactions = [...transactions, ...newData];
        saveAndRender();
    }

    function csvToArray(str, delimiter = ",") {
        const rows = str.slice(str.indexOf("\n") + 1).split("\n");
        return rows.map(row => row.split(delimiter));
    }

    // ================== TABEL UTAMA ==================

    function renderData() {
        tableBody.innerHTML = '';
        let totalInc = 0;
        let totalExp = 0;

        transactions.forEach((item) => {
            totalInc += item.credit;
            totalExp += item.debit;

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.date}</td>
                <td>${item.desc}</td>
                <td style="color:red">-${formatIDR(item.debit)}</td>
                <td style="color:green">+${formatIDR(item.credit)}</td>
                <td><input type="text" value="${item.category}" onchange="updateRow(${item.id}, 'category', this.value)" class="edit-input"></td>
                <td><input type="text" value="${item.note}" onchange="updateRow(${item.id}, 'note', this.value)" class="edit-input"></td>
                <td><button class="delete-row" onclick="deleteRow(${item.id})">✕</button></td>
            `;
            tableBody.appendChild(tr);
        });

        updateSummary(totalInc, totalExp);
    }

    window.updateRow = (id, field, value) => {
        const index = transactions.findIndex(t => t.id === id);
        transactions[index][field] = value;
        localStorage.setItem('statwise_data', JSON.stringify(transactions));
    };

    window.deleteRow = (id) => {
        transactions = transactions.filter(t => t.id !== id);
        saveAndRender();
    };

    function updateSummary(inc, exp) {
        document.getElementById('totalIncome').innerText = formatIDR(inc);
        document.getElementById('totalExpense').innerText = formatIDR(exp);
        document.getElementById('totalBalance').innerText = formatIDR(inc - exp);
    }

    function saveAndRender() {
        localStorage.setItem('statwise_data', JSON.stringify(transactions));
        renderData();
    }

    function formatIDR(amount) {
        return new Intl.NumberFormat('id-ID', {
            style: 'currency',
            currency: 'IDR',
            maximumFractionDigits: 0
        }).format(amount);
    }
});

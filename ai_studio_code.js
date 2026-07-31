document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const tableBody = document.getElementById('tableBody');
    const clearBtn = document.getElementById('clearData');
    
    let transactions = JSON.parse(localStorage.getItem('statwise_data')) || [];

    // Tampilkan data saat pertama kali load
    renderData();

    // Event listener untuk upload file
    fileInput.addEventListener('change', handleFileUpload);

    // Hapus Semua Data
    clearBtn.addEventListener('click', () => {
        if(confirm('Hapus semua data transaksi?')) {
            transactions = [];
            saveAndRender();
        }
    });

    function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            const data = event.target.result;
            let workbook;

            if (file.name.endsWith('.csv')) {
                // Parsing CSV sederhana
                const text = event.target.result;
                processData(csvToArray(text));
            } else {
                // Parsing Excel menggunakan SheetJS
                workbook = XLSX.read(data, { type: 'binary' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                processData(json);
            }
        };

        if (file.name.endsWith('.csv')) {
            reader.readAsText(file);
        } else {
            reader.readAsBinaryString(file);
        }
    }

    function processData(rows) {
        // Asumsi header: Tanggal, Deskripsi, Debit, Kredit
        // Kita lewati baris pertama (header)
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

    function renderData() {
        tableBody.innerHTML = '';
        let totalInc = 0;
        let totalExp = 0;

        transactions.forEach((item, index) => {
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
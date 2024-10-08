const xlsx = require('xlsx');
const fs = require('fs');

// Baca data dari file JSON
const rawData = fs.readFileSync('data.json');
const jsonData = JSON.parse(rawData);

// Fungsi untuk memproses data JSON menjadi format yang diinginkan
const processData = (data) => {
    return data.map(entry => ({
        "Tanggal": entry.tanggal,
        "Nama Nasabah": entry.namaNasabah,
        "Nama Tim Project": entry.namaTimProject.nama,
        "Nama Tim Market": entry.namaMarket.nama,
        "Nama Mitra / Subsidi": entry.namaMitra,
        "Cabang Pengerjaan": entry.cabangPengerjaan.nama,
        "Reports": entry.reports.map(report => `${report.aplikasi} (Rp${report.pencairan})`).join(', '),
        "Jumlah Pencairan": entry.jumlahPencairan,
        "Jumlah Transfer": entry.jumlahTransfer,
        "Keterangan": entry.keterangan,
    }));
};

// Konversi data menjadi format Excel
const convertToExcel = (data) => {
    const processedData = processData(data);

    // Buat worksheet dari data yang telah diproses
    const ws = xlsx.utils.json_to_sheet(processedData);

    // Buat workbook dan tambahkan worksheet
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');

    // Tulis workbook ke file
    xlsx.writeFile(wb, 'DATA PENCAIRAN SEPTEMBER 2024.xlsx');
};

// Jalankan konversi
convertToExcel(jsonData);

console.log('Data telah berhasil dikonversi ke Excel.');

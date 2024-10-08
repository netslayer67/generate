const xlsx = require('xlsx');
const fs = require('fs');

// Baca data dari file JSON
const rawData = fs.readFileSync('gestun.json');
const jsonData = JSON.parse(rawData);

// Fungsi untuk memproses data JSON menjadi format yang diinginkan
const processData = (data) => {
    return data.map(entry => ({
        "Tanggal": entry.tanggal,
        "Nama Nasabah": entry.namaNasabah,
        "Nama Tim Project": entry.namaTimProject ? entry.namaTimProject.nama : null,
        "Nama Tim Market": entry.namaMarket.nama,
        "Nama Mitra / Subsidi": entry.namaMitra,
        "Cabang Pengerjaan": entry.cabangPengerjaan.nama,
        "Nama Aplikasi": entry.aplikasi,
        "Jumlah Gestun": entry.jumlahGestun,
        "Jumlah Transfer": entry.jumlahTransfer,
        "feeToko": entry.feeToko,
        "potonganDp": entry.potonganDp,
        "potonganLainnya": entry.potonganLainnya,
        "Keterangan": entry.keterangan,
        "buktiTF": entry.buktiTransfer.url
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
    xlsx.writeFile(wb, 'DATA GESTUN SEPTEMBER 2024.xlsx');
};

// Jalankan konversi
convertToExcel(jsonData);

console.log('Data telah berhasil dikonversi ke Excel.');

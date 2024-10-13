<?php
require '../../vendor/autoload.php'; // PhpSpreadsheet autoload

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_POST['audience_id']) && isset($_POST['event_id'])) {
    $audience_id = $_POST['audience_id'];
    $event_id = $_POST['event_id'];
    
    // Path ke file Excel yang akan diperbarui
    $filePath = '../../Registration.xlsx';

    // Buka file Excel
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    // Cari baris kosong untuk memasukkan data baru
    $highestRow = $sheet->getHighestRow() + 1; // Mendapatkan baris terakhir

    // Tulis data attendance ke baris kosong
    $sheet->setCellValue('A' . $highestRow, $audience_id);
    $sheet->setCellValue('B' . $highestRow, date('Y-m-d H:i:s')); // Waktu kehadiran

    // Simpan kembali file Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);

    // Kembalikan response berhasil
    echo json_encode([
        'status' => 1,
        'name' => "Participant $audience_id"
    ]);
    exit;
} else {
    // Kembalikan response gagal jika data tidak valid
    echo json_encode([
        'status' => 0,
        'error' => 'Invalid data'
    ]);
    exit;
}

<?php
header('Content-Type: application/json');
require 'koneksi.php'; // Diasumsikan file ini menangani koneksi ke database

// Kriteria untuk query
// Anda bisa mengubah atau menambah kondisi di dalam klausa WHERE di bawah ini
// DATEDIFF(CURDATE(), tgl_kirim) digunakan untuk menghitung selisih hari antara tanggal saat ini dan tgl_kirim
$sql = "SELECT 
            connote, 
            pnrm, 
            produk, 
            tgl_kirim, 
            cod, 
            bsu_cod,
            DATEDIFF(CURDATE(), tgl_kirim) as umur
        FROM 
            tbl_antrn 
        WHERE 
            st = '33' AND
            DATEDIFF(CURDATE(), tgl_kirim) > sla";

$result = $conn->query($sql);

$data = array();
if ($result->num_rows > 0) {
    while($row = $result->fetch_assoc()) {
        $data[] = $row;
    }
}

$conn->close();

echo json_encode($data);
?>

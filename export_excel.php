<?php
session_start();
require 'db.php';

// Include PhpSpreadsheet library
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Redirect if user is not authenticated
if (!isset($_SESSION['user_id'])) {
    header("Location: login.php");
    exit();
}

// Get user parameter
if (!isset($_GET['user'])) {
    echo "User not specified.";
    exit();
}

$username = $_GET['user'];

// Fetch user details
$stmt = $pdo->prepare("SELECT * FROM admin1 WHERE user = ?");
$stmt->execute([$username]);
$user = $stmt->fetch(PDO::FETCH_ASSOC);

if (!$user) {
    echo "User not found.";
    exit();
}

$user_id = $user['id'];

// Fetch user's images and timestamps
$stmt = $pdo->prepare("SELECT image_path, location, uploaded_at FROM uploads WHERE user_id = ? ORDER BY uploaded_at ASC");
$stmt->execute([$user_id]);
$images = $stmt->fetchAll(PDO::FETCH_ASSOC);

if (empty($images)) {
    echo "No records to export.";
    exit();
}

// Create a new Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Set the header row
$headers = ['Name', 'Location', 'Upload_at', 'Time In', 'Time Out', 'Overtime', 'Undertime'];
$sheet->fromArray($headers, NULL, 'A1');

// Prepare data rows
$row = 2;
foreach ($images as $image) {
    // Extract user name, location, and upload timestamp
    $name = $username;
    $location = $image['location'];
    $uploadedAt = $image['uploaded_at'];

    // Calculate Time In and Time Out
    if ($location === 'In') {
        $timeIn = $uploadedAt;
        $timeOut = '';
        $overtime = '';
        $undertime = '';
    } elseif ($location === 'Out') {
        $timeIn = '';
        $timeOut = $uploadedAt;
        $overtime = '';
        $undertime = '';
    } else {
        $timeIn = '';
        $timeOut = '';
        $overtime = $location === 'Overtime' ? $uploadedAt : '';
        $undertime = $location === 'Undertime' ? $uploadedAt : '';
    }

    // Write data to the spreadsheet
    $sheet->setCellValue("A$row", $name);
    $sheet->setCellValue("B$row", $location);
    $sheet->setCellValue("C$row", $uploadedAt);
    $sheet->setCellValue("D$row", $timeIn);
    $sheet->setCellValue("E$row", $timeOut);
    $sheet->setCellValue("F$row", $overtime);
    $sheet->setCellValue("G$row", $undertime);

    $row++;
}

// Set column widths for better visibility
foreach (range('A', 'G') as $column) {
    $sheet->getColumnDimension($column)->setAutoSize(true);
}

// Set the filename
$filename = "Book2_Export_" . $username . ".xlsx";

// Send the file as a download
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $filename . '"');
header('Cache-Control: max-age=0');

// Save the file to output
$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit();
?>

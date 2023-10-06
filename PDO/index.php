<!-- SOURCE : https://code-boxx.com/import-excel-into-mysql-php/#sec-import -->

<?php
// (A) CONNECT TO DATABASE - CHANGE SETTINGS TO YOUR OWN!
$dbhost = "localhost";
$dbname = "dmy_coba_db";
$dbchar = "utf8mb4";
$dbuser = "root";
$dbpass = "";
$pdo = new PDO(
    "mysql:host=$dbhost;charset=$dbchar;dbname=$dbname",
    $dbuser,
    $dbpass,
    [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC
    ]
);

// (B) PHPSPREADSHEET TO LOAD EXCEL FILE
require "../vendor/autoload.php";
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$spreadsheet = $reader->load("file_import.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

// (C) READ DATA + IMPORT
$sql = "INSERT INTO `users` (`name`, `email`) VALUES (?,?)";
foreach ($worksheet->getRowIterator() as $row) {
    // (C1) FETCH DATA FROM WORKSHEET
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);
    $data = [];
    foreach ($cellIterator as $cell) {
        $data[] = $cell->getValue();
    }

    // (C2) INSERT INTO DATABASE
    print_r($data);
    try {
        $stmt = $pdo->prepare($sql);
        $stmt->execute($data);
        echo "OK - USER ID - {$pdo->lastInsertId()}<br>";
    } catch (Exception $ex) {
        echo $ex->getMessage() . "<br>";
    }
    $stmt = null;
}

// (D) CLOSE DATABASE CONNECTION
if ($stmt !== null) {
    $stmt = null;
}
if ($pdo !== null) {
    $pdo = null;
}

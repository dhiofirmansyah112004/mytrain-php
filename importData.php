<?php

// Load the database configuration file 
include_once 'dbConfig.php';

// Include PhpSpreadsheet library autoloader 
require_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

if (isset($_POST['importSubmit'])) {

    // Allowed mime types 
    $excelMimes = array('text/xls', 'text/xlsx', 'application/excel', 'application/vnd.msexcel', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Validate whether selected file is a Excel file 
    if (!empty($_FILES['file']['name']) && in_array($_FILES['file']['type'], $excelMimes)) {

        // If the file is uploaded 
        if (is_uploaded_file($_FILES['file']['tmp_name'])) {
            $reader = new Xlsx();
            $spreadsheet = $reader->load($_FILES['file']['tmp_name']);
            $worksheet = $spreadsheet->getActiveSheet();
            $worksheet_arr = $worksheet->toArray();

            // Remove header row 
            unset($worksheet_arr[0]);

            foreach ($worksheet_arr as $row) {
                $first_name = $row[0];
                $last_name = $row[1];
                $email = $row[2];
                $phone = $row[3];
                $status = $row[4];

                // Check whether member already exists in the database with the same email 
                $prevQuery = "SELECT id FROM members_uploadexcel_7 WHERE email = '" . $email . "'";
                $prevResult = $db->query($prevQuery);

                if ($prevResult->num_rows > 0) {
                    // Update member data in the database 
                    $db->query("UPDATE members_uploadexcel_7 SET first_name = '" . $first_name . "', last_name = '" . $last_name . "', email = '" . $email . "', phone = '" . $phone . "', status = '" . $status . "', modified = NOW() WHERE email = '" . $email . "'");
                } else {
                    // Insert member data in the database 
                    $db->query("INSERT INTO members_uploadexcel_7 (first_name, last_name, email, phone, created, status) VALUES ('" . $first_name . "', '" . $last_name . "', '" . $email . "', '" . $phone . "', NOW(), '" . $status . "' )");
                }
            }

            $qstring = '?status=succ';
        } else {
            $qstring = '?status=err';
        }
    } else {
        $qstring = '?status=invalid_file';
    }
}

// Redirect to the listing page 
header("Location: index.php" . $qstring);

<?php
// Check if form submitted
if (isset($_POST["submit"])) {

    // Check if file uploaded successfully
    if ($_FILES["excel_file"]["error"] == UPLOAD_ERR_OK) {

        // Get file extension
        $file_ext = pathinfo($_FILES["excel_file"]["name"], PATHINFO_EXTENSION);

        // Check if file is Excel file
        if ($file_ext == "xlsx" || $file_ext == "xls") {

            // Read Excel file using PHPExcel library
            require_once "PHPExcel/Classes/PHPExcel.php";
            $input_file = $_FILES["excel_file"]["tmp_name"];
            $excel_reader = PHPExcel_IOFactory::createReaderForFile($input_file);
            $excel_obj = $excel_reader->load($input_file);

            // Get worksheet
            $worksheet = $excel_obj->getActiveSheet();

            // Get highest row and column
            $highest_row = $worksheet->getHighestRow();
            $highest_column = $worksheet->getHighestColumn();

            // Loop through each row
            for ($row = 1; $row <= $highest_row; $row++) {
                // Get cell values
                $column_values = array();
                for ($col = 'A'; $col <= $highest_column; $col++) {
                    $cell_value = $worksheet->getCell($col . $row)->getValue();
                    $column_values[] = $cell_value;
                }

                // Insert data into MySQL
                $db_host = "localhost";
                $db_user = "root";
                $db_password = "";
                $db_name = "xlsx";
                $db_conn = mysqli_connect($db_host, $db_user, $db_password, $db_name);

                $query = "INSERT INTO students (column1, column2, column3) VALUES ('$column_values[0]', '$column_values[1]', '$column_values[2]')";
                mysqli_query($db_conn, $query);
                mysqli_close($db_conn);
            }

            echo "Data uploaded successfully!";
        } else {
            echo "Please upload Excel file!";
        }
    } else {
        echo "File upload failed!";
    }
}

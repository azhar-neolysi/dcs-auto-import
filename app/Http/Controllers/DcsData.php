<?php

require_once "Classes/PHPExcel.php";

mysqli_report(MYSQLI_REPORT_ERROR | MYSQLI_REPORT_STRICT);
//connect to database
$connection = mysqli_connect("localhost", "root", "", "crontest");
if (mysqli_connect_errno()) {
    echo "failed to connect";
}

echo "<h1>Excel Reading</h1><br>";
$query2 = "SELECT valid_to FROM `kt_dcs_data_flow_meter` ORDER BY `kt_dcs_data_flow_meter_id` DESC LIMIT 1";

$result2 = mysqli_query($connection, $query2); //last file updated date selection
if ($result2->num_rows > 0) {

    while ($row = $result2->fetch_assoc()) {
        $vdate = $row['valid_to'];
        // echo $row['valid_to'];

        $prvdate = date('Y-m-d', strtotime("-1 days")); //echo $prvdate; echo '<br>';
        if ($vdate == 0) {
            $Pre_filename = '0';
        }
        //  echo $vdate;echo "<br>";
        $Pre_filename = str_replace('-', '', $vdate);
        $Pre_filename = substr($Pre_filename, 0, 8);  //convert into file name format
    }
} else {
    echo "0 results";
}

$a = 'D';
$ex = '.xlsx';
$fileName = $a . $Pre_filename;

$Pre_filename = $fileName . $ex; //adding ext
$path = [];
$filepath = [];
// $fileName ='D20200504.xls';
$mydir = 'D:\Office\Task Excels\Kothari\dcs dATA';   //Source Directory

$myfiles = array_diff(scandir($mydir), array('.', '..')); //reading files name form source directory
print_r($myfiles);
$length = count($myfiles);
echo $length;
echo '<br>';
$filefound=false;
// foreach ($myfiles as $filename) {
// }
for ($i = 2; $i < $length+2; $i++) {
    echo $myfiles[$i];
    echo '<br>';
    echo $Pre_filename;
    echo '<br>';
    if ($myfiles[$i] > $Pre_filename) { //compare with selected db file with source directory files
          echo '<br>';
        // $path = $filename;
        array_push($path, $myfiles[$i]);
        print_r($path);
        echo '<br>';
        $filefound=true;
        $sflength = count($path);
    } else {
        // if(!$filefound){

        //     $text = 'File Not Found/Already Inserted';
        //     date_default_timezone_set('Asia/Kolkata');
        //     $time = date("Y-m-d H:i:s");
        //     $myfile1 = file_put_contents('messagelogs.txt', $time . $text . PHP_EOL, FILE_APPEND | LOCK_EX);
        //     echo "<br>";
        //     echo 'File Not Found/Already Inserted';
        // }

    }

}
if(!$filefound){

    $text = 'File Not Found/Already Inserted';
    date_default_timezone_set('Asia/Kolkata');
    $time = date("Y-m-d H:i:s");
    $myfile1 = file_put_contents('messagelogs.txt', $time . $text . PHP_EOL, FILE_APPEND | LOCK_EX);
    echo "<br>";
    echo 'File Not Found/Already Inserted';
}


for($i=0;$i<$sflength;$i++){

$filepath = $mydir . DIRECTORY_SEPARATOR . $path[$i];
echo $filepath;

$allowed_ext = ['xls', 'xlsx', 'xlt'];

$reader = PHPExcel_IOFactory::createReaderForFile($filepath); //selected excel reading
$excel_Obj = $reader->load($filepath);
$worksheet = $excel_Obj->getActiveSheet();

$path_ext = pathinfo($filepath, PATHINFO_EXTENSION);



$lastRow = $worksheet->getHighestRow();
$lastCol = $worksheet->getHighestDataColumn();
$colCount = PHPExcel_Cell::columnIndexFromString($lastCol);
echo "<br>";
echo $lastRow . "....<br>";
echo $lastCol;
if (in_array($path_ext, $allowed_ext)) {
    $spreadSheet = PHPExcel_IOFactory::load($filepath);
    $data = $spreadSheet->getActiveSheet()->toArray();
    // echo $data;
} else {
    echo "Invalid File";
    $text = ":Invalid File";
    date_default_timezone_set('Asia/Kolkata');
    $time = date("Y-m-d");

    // $txt = "data-to-add";
    $myfile1 = file_put_contents('messagelogs.txt', $time . $text . PHP_EOL, FILE_APPEND | LOCK_EX);
}

$excel_date = $worksheet->getCell('N3')->getValue();

$unix_date = ($excel_date - 25569) * 86400;
$excel_date = 25569 + ($unix_date / 86400);
$unix_date = ($excel_date - 25569) * 86400;
$vtime = gmdate("Y-m-d H:i:s", $unix_date);
echo $vtime;
$vtime1 = gmdate('d_F_Y', $unix_date);

// return;
$createdBy = '1';


$aliasName = 'DCS_UPLOAD_' . $vtime1;

// return;
date_default_timezone_set('Asia/Kolkata');
$time = date("Y-m-d H:i:s");
$query3 = "SELECT * FROM `kt_dcs_data_instrument` WHERE valid_to ='$vdate'";
// return;

$result3 = mysqli_query($connection, $query3);

if ($result3->num_rows > 0) {
    // output data of each row
    while ($row = $result3->fetch_assoc()) {
        $aliasName1 = $row['alias_name'];
        $licr101opeing = $row['licr-101_closing'];
        $licr101closing = $worksheet->getCell('I124')->getOldCalculatedValue(); //OK
        $lit1104opening = $row['li-t1104_closing'];
        $li1104closing = $worksheet->getCell('M124')->getOldCalculatedValue(); //OK
        $lit1001unit1opening = $row['li-t1001_unit1_closing']; //
        $lit1001unit1closing = 0; //
        $licr201opening = $row['licr-201_closing'];
        $licr201closing = $worksheet->getCell('J124')->getOldCalculatedValue(); //OK
        $lit1105aopening = $row['li-t1105a_closing'];
        $li1105aclosing = $worksheet->getCell('K124')->getOldCalculatedValue(); //OK
        $lit1001unit2opening = $row['li-t1001_unit2_closing']; //
        $lit1001unit2closing = 0; //
        $lit1103aopening = $row['li-t1103a_closing'];
        $li1103aclosing = $worksheet->getCell('K80')->getOldCalculatedValue(); //OK
        $lit1103bopening = $row['li-t1103b_closing'];
        $li1103bclosing = $worksheet->getCell('L80')->getOldCalculatedValue(); //OK
        $lit1103copening = $row['li-t1103c_closing'];
        $li1103cclosing = $worksheet->getCell('M80')->getOldCalculatedValue(); //OK
        $li1001opening = $row['li-1001_closing'];
        $li1001closing = $worksheet->getCell('K36')->getOldCalculatedValue(); //OK
        $li1002opening = $row['li-1002_closing'];
        $li1002closing = $worksheet->getCell('I36')->getOldCalculatedValue(); //OK
        $li1002aopening = $row['li-1002a_closing'];
        $li1002aclosing = $worksheet->getCell('B166')->getOldCalculatedValue(); //OK//
        $li1003aopening = $row['li-1003a_closing'];
        $li1003aclosing = $worksheet->getCell('C166')->getOldCalculatedValue(); //OK//
        $li1003bopening = $row['li-1003b_closing'];
        $li1003bclosing = $worksheet->getCell('D166')->getOldCalculatedValue(); //OK//
        $li1003copening = $row['li-1003c_closing'];
        $li1003cclosing = $worksheet->getCell('E166')->getOldCalculatedValue(); //OK//
        $li2003aopening = $row['li-2003a_closing'];
        $li2003aclosing = $worksheet->getCell('F166')->getOldCalculatedValue(); //OK
        $li2003bopening = $row['li-2003b_closing'];
        $li2003bclosing = $worksheet->getCell('G166')->getOldCalculatedValue(); //OK
        $li1203aopening = $row['li-1203a_closing'];
        $li1203aclosing = $worksheet->getCell('H166')->getOldCalculatedValue(); //OK
        $li1203bopening = $row['li-1203b_closing'];
        $li1203bclosing = $worksheet->getCell('I166')->getOldCalculatedValue(); //OK//
        $li1101aopening = $row['li-1101a_closing'];
        $li1101aclosing = $worksheet->getCell('J166')->getOldCalculatedValue(); //OK
        $li1101bopening = $row['li-1101b_closing'];
        $li1101bclosing = $worksheet->getCell('K166')->getOldCalculatedValue(); //OK
        $li1101copening = $row['li-1101c_closing'];
        $li1101cclosing = $worksheet->getCell('L166')->getOldCalculatedValue(); //OK
        $li1101dopening = $row['li-1101d_closing'];
        $li1101dclosing = $worksheet->getCell('M166')->getOldCalculatedValue(); //OK
        $li1101eopeing = $row['li-1101e_closing']; //
        $li1101eclosing = 0; //
        $li1102aopening = $row['li-1102a_closing'];
        $li1102aclosing = $worksheet->getCell('C207')->getOldCalculatedValue(); //OK
        $li1102bopening = $row['li-1102b_closing'];
        $li1102bclosing = $worksheet->getCell('D207')->getOldCalculatedValue(); //OK
        $li1102copening = $row['li-1102c_closing'];
        $li1102cclosing = $worksheet->getCell('E207')->getOldCalculatedValue(); //OK
        $li1106aopening = $row['li-1106a_closing'];
        $li1106aclosing = $worksheet->getCell('G80')->getOldCalculatedValue(); //OK
        $li1106bopening = $row['li-1106b_closing'];
        $li1106bclosing = $worksheet->getCell('H80')->getOldCalculatedValue(); //OK
        $li1105opening = $row['li-1105_closing'];
        $li1105closing = $worksheet->getCell('I80')->getOldCalculatedValue(); //OK
    }
} else {
    echo "0 results";
}

// return;
$fiq401 = $worksheet->getCell('B36')->getOldCalculatedValue(); //OK
$fcq404 = $worksheet->getCell('C36')->getOldCalculatedValue(); //OK
$fiq112a = $worksheet->getCell('D36')->getOldCalculatedValue(); //OK
$fiq1106a = $worksheet->getCell('F36')->getOldCalculatedValue(); //OK
$fiq1002 = $worksheet->getCell('E36')->getOldCalculatedValue(); //OK
$fcq1102 = $worksheet->getCell('G36')->getOldCalculatedValue(); //OK
$fcq1103 = $worksheet->getCell('H36')->getOldCalculatedValue(); //OK
$fiq1107a = $worksheet->getCell('J36')->getOldCalculatedValue(); //OK
$fcq116 = $worksheet->getCell('M207')->getOldCalculatedValue(); // NOT OK V1006||U1//
$fcq216 = $worksheet->getCell('C249')->getOldCalculatedValue(); // NOT OK U1/U2?
$fiq2004 = $worksheet->getCell('F80')->getOldCalculatedValue(); // NOT OK REGULAR||INTER//
$fcq205 = $worksheet->getCell('I207')->getOldCalculatedValue(); //OK
$fcq114 = $worksheet->getCell('H207')->getOldCalculatedValue(); //OK
$fiqoffgas = $worksheet->getCell('L124')->getOldCalculatedValue(); //OK
$fiqr101trim = $worksheet->getCell('F207')->getOldCalculatedValue(); //OK
$fiq202 = $worksheet->getCell('G124')->getOldCalculatedValue(); //OK
$fiq123 = $worksheet->getCell('B80')->getOldCalculatedValue(); //OK
$fiq134a = $worksheet->getCell('D249')->getOldCalculatedValue(); //OK//134||134a
$fiq126 = $worksheet->getCell('J207')->getOldCalculatedValue(); //OK
$fiq157 = $worksheet->getCell('C124')->getOldCalculatedValue(); //OK
$fiq150 = $worksheet->getCell('B124')->getOldCalculatedValue(); //OK
$fiqr201trim = $worksheet->getCell('F124')->getOldCalculatedValue(); //OK
$fiq302 = $worksheet->getCell('H124')->getOldCalculatedValue(); //OK
$fiq223a = $worksheet->getCell('B249')->getOldCalculatedValue(); //OK
$fiq226 = $worksheet->getCell('K207')->getOldCalculatedValue(); //OK
$fiq234 = $worksheet->getCell('E80')->getOldCalculatedValue(); //234 || 234A
$fiq257 = $worksheet->getCell('D124')->getOldCalculatedValue(); //OK





$fiq134 = $worksheet->getCell('C80')->getOldCalculatedValue();
$fcq223 = $worksheet->getCell('D80')->getOldCalculatedValue();
$fiq2002 = $worksheet->getCell('J80')->getOldCalculatedValue();

$fiq250 = $worksheet->getCell('E124')->getOldCalculatedValue();

$li1201mt = $worksheet->getCell('B207')->getOldCalculatedValue();
$fiq2002 = $worksheet->getCell('G207')->getOldCalculatedValue();
$fiq123a = $worksheet->getCell('L207')->getOldCalculatedValue();


$fiq234a = $worksheet->getCell('E249')->getOldCalculatedValue();
$gcr101trim = $worksheet->getCell('F249')->getOldCalculatedValue();
$gcr201trim = $worksheet->getCell('G249')->getOldCalculatedValue();
$fiqnh3p1 = $worksheet->getCell('H249')->getOldCalculatedValue();
$fiqnh3p2 = $worksheet->getCell('I249')->getOldCalculatedValue();
$tag069 = $worksheet->getCell('J249')->getOldCalculatedValue();
$tag070 = $worksheet->getCell('K249')->getOldCalculatedValue();
$tag071 = $worksheet->getCell('L249')->getOldCalculatedValue();
$tag072 = $worksheet->getCell('M249')->getOldCalculatedValue();




$query1 = "INSERT INTO  kt_dcs_data_flow_meter(`alias_name`,`fiq-401`,`fcq-404`,`fiq-112a`,`fiq-1106a`,`fiq-1002`,`fcq-1102`,`fcq-1103`,`fiq-1107a`,`fcq-116_v1001`,`fcq-216_v1001`,`fiq-2004_regular`,`fiq-2004_intermediate`,`fcq-205`,`fcq-114`,`fiq-offgas`,`fcq-116_unit1`,`fiq-r101_trim`,`fiq-202`,`fiq-123`,`fiq-134`,`fiq-126`,`fiq-157`,`fiq-150`,`fcq-216_unit2`,`fiq-r201_trim`,`fiq-302`,`fiq-223`,`fiq-226`,`fiq-234`,`fiq-257`,`fiq-250`,`valid_from`,`valid_to`,`created_on`,`created_by`) VALUES ('$aliasName','$fiq401','$fcq404 ','$fiq112a','$fiq1106a','$fiq1002','$fcq1102','$fcq1103','$fiq1107a','$fcq116','$fcq216','$fiq2004','$fiq2004','$fcq205','$fcq114','$fiqoffgas','$fcq116','$fiqr101trim','$fiq202','$fiq123','$fiq134a','$fiq126','$fiq157','$fiq150','$fcq216','$fiqr201trim','$fiq302','$fiq223a','$fiq226','$fiq234','$fiq257','$fiq250','$vtime','$vtime','$time','$createdBy')";

$query2 = "INSERT INTO kt_dcs_data_instrument(`alias_name`,`licr-101_opening`,`li-t1104_opening`,`li-t1001_unit1_opening`,`licr-201_opening`,`li-t1105a_opening`,`li-t1001_unit2_opening`,`li-t1103a_opening`,`li-t1103b_opening`,`li-t1103c_opening`,`li-1001_opening`,`li-1002_opening`,`li-1002a_opening`,`li-1003a_opening`,`li-1003b_opening`,`li-1003c_opening`,`li-2003a_opening`,`li-2003b_opening`,`li-1203a_opening`,`li-1203b_opening`,`li-1101a_opening`,`li-1101b_opening`,`li-1101c_opening`,`li-1101d_opening`,`li-1101e_opening`,`li-1102a_opening`,`li-1102b_opening`,`li-1102c_opening`,`li-1106a_opening`,`li-1106b_opening`,`li-1105_opening`,`licr-101_closing`,`li-t1104_closing`,`li-t1001_unit1_closing`,`licr-201_closing`,`li-t1105a_closing`,`li-t1001_unit2_closing`,`li-t1103a_closing`,`li-t1103b_closing`,`li-t1103c_closing`,`li-1001_closing`,`li-1002_closing`,`li-1002a_closing`,`li-1003a_closing`,`li-1003b_closing`,`li-1003c_closing`,`li-2003a_closing`,`li-2003b_closing`,`li-1203a_closing`,`li-1203b_closing`,`li-1101a_closing`,`li-1101b_closing`,`li-1101c_closing`,`li-1101d_closing`,`li-1101e_closing`,`li-1102a_closing`,`li-1102b_closing`,`li-1102c_closing`,`li-1106a_closing`,`li-1106b_closing`,`li-1105_closing`,`valid_from`,`valid_to`,`created_on`,`created_by`)VALUES ('$aliasName','$licr101opeing','$lit1104opening ','$lit1001unit1opening','$licr201opening','$lit1105aopening','$lit1001unit2opening','$lit1103aopening','$lit1103bopening','$lit1103copening','$li1001opening','$li1002opening','$li1002aopening','$li1003aopening','$li1003bopening','$li1003copening','$li2003aopening','$li2003bopening','$li1203aopening','$li1203bopening','$li1101aopening','$li1101bopening','$li1101copening','$li1101dopening','$li1101eopeing','$li1102aopening','$li1102bopening','$li1102copening','$li1106aopening','$li1106bclosing','$li1105opening','$licr101closing','$li1104closing','$lit1001unit1closing','$licr201closing','$li1105aclosing','$lit1001unit2closing','$li1103aclosing','$li1103bclosing','$li1103cclosing','$li1001closing','$li1002closing','$li1002aclosing','$li1003aclosing','$li1003bclosing','$li1003cclosing','$li2003aclosing','$li2003bclosing','$li1203aclosing','$li1203bclosing','$li1101aclosing','$li1101bclosing','$li1101cclosing','$li1101dclosing','$li1101eclosing','$li1102aclosing','$li1102bclosing','$li1102cclosing','$li1106aclosing','$li1106bclosing','$li1105closing','$vtime','$vtime','$time','$createdBy')";

$result_flow_meter = mysqli_query($connection, $query1);
$result_instrument = mysqli_query($connection, $query2);
$msg = true;

}


if (isset($msg)) {
    $text = ':Success';
    date_default_timezone_set('Asia/Kolkata');
    $time = date("Y-m-d H:i:s");

    // $txt = "data-to-add";
    $myfile1 = file_put_contents('messagelogs.txt', $time . $text . PHP_EOL, FILE_APPEND | LOCK_EX);
    echo $_SESSION['message'] = "Success";
} else {
    $text = 'Not Inserted';
    date_default_timezone_set('Asia/Kolkata');
    $time = date("Y-m-d H:i:s");

    // $txt = "data-to-add";
    $myfile1 = file_put_contents('messagelogs.txt', $time . $text . PHP_EOL, FILE_APPEND | LOCK_EX);
    echo $_SESSION['message'] = "Not Inserted";
}

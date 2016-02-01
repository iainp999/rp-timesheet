<?php

require_once('vendor/autoload.php');


$start_date = strtotime(getenv("TS_START_DATE"));
$consultant = getenv("TS_NAME");
$client = getenv("TS_CLIENT");
$project = getenv("TS_PROJECT");
$role = getenv("TS_ROLE");
$activity = getenv("TS_ACTIVITY");

$excel = PHPExcel_IOFactory::load("template.xlsx");
$excel->setActiveSheetIndex(0);

$excel->getActiveSheet()->setCellValue('C4', $consultant);
$excel->getActiveSheet()->setCellValue('C6', $client);
$excel->getActiveSheet()->setCellValue('C8', $project);
$excel->getActiveSheet()->setCellValue('C10', $role);

for ($d = 0; $d < 5; $d++) {

  $the_date = $start_date + ($d * 60 * 60 * 24);

  $day = date('D', $the_date);
  $date = date('d-M-y', $the_date);

  $row = 13 + ($d * 5);

  $excel->getActiveSheet()->setCellValue("B{$row}", $day);
  $excel->getActiveSheet()->setCellValue("C{$row}", $date);
  $excel->getActiveSheet()->setCellValue("D{$row}", $activity);
  $excel->getActiveSheet()->setCellValue("I{$row}", "1");
}

$writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$writer->save('timesheet.xlsx');

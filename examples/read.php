<?php

require_once("../class.directexcel.php");

// single function read
$data = DirectExcel::readArray("inventions.xlsx");

header("content-type: text/plain");
echo "Reading inventions.xlsx\n\n";

// show the data
print_r($data);

?>
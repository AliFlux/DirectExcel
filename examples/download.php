<?php

require_once("../class.directexcel.php");

// data is in array
$data = array(
	"Aircrafts" => array(
		array(
			"Name" => "SR-71 BlackBird",
			"Speed" => "Mach 3+",
			"Built" => "32",
			"Developer" => "Lockheed Martin",
			"Link" => "http://en.wikipedia.org/wiki/Lockheed_SR-71_Blackbird",
		),
		array(
			"Name" => "F-22 Raptor",
			"Speed" => "Mach 1.82",
			"Built" => "195",
			"Developer" => "Lockheed Martin",
			"Link" => "https://en.wikipedia.org/wiki/Lockheed_Martin_F-22_Raptor",
		),
		array(
			"Name" => "C-17 Globemaster III",
			"Speed" => "Mach 0.74",
			"Built" => "250",
			"Developer" => "McDonnell Douglas",
			"Link" => "http://en.wikipedia.org/wiki/Boeing_C-17_Globemaster_III",
		),
	),
	"Aircraft Carrier" => array(
		array(
			"Name" => "USS Nimitz",
			"Classification" => "CVN-68",
			"Length" => "1,092 feet",
			"Motto" => "Teamwork, a Tradition",
			"Link" => "http://en.wikipedia.org/wiki/USS_Nimitz_%28CVN-68%29",
		),
		array(
			"Name" => "USS Gerald R. Ford",
			"Classification" => "CVN-78",
			"Length" => "1,106 feet",
			"Motto" => "",
			"Link" => "http://en.wikipedia.org/wiki/USS_Gerald_R._Ford_%28CVN-78%29",
		),
		array(
			"Name" => "USS Ronald Reagan",
			"Classification" => "CVN-76",
			"Length" => "1,092 feet",
			"Motto" => "Peace Through Strength",
			"Link" => "http://en.wikipedia.org/wiki/USS_Ronald_Reagan_%28CVN-76%29",
		),
	)
);

// One function download
DirectExcel::download($data);

?>
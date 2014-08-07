<?php

// DirectExcel by Ali Ashraf
// facebook.com/Ali.8X
// aliashraf@ieee.org

/*
	The xlsx format is basically a bunch of zipped spreadsheets in XML format
	For experimentation, get any .xlsx file and rename it to .zip
	Then open it and explore its structure
		- /[Content_Types].xml
		- /_rels/.rels
		- /docProps/app.xml
		- /docProps/core.xml
		- /xl/theme/theme1.xml
		- /xl/styles.xml
*/

class DirectExcel {

	// directly make and download excel spreadsheet from data
	static function download($worksheets, $filename = 'spreadsheet.xlsx') 
	{
		// expects that the filename is sanitized
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header("Content-Transfer-Encoding: Binary"); 
		header("Content-disposition: attachment; filename=\"" . ($filename) . "\"");
		echo self::parse($worksheets);
	}

	// write worksheets to a file
	static function write($worksheets, $filename) 
	{
		$parsed = self::parse($worksheets);
		file_put_contents($filename, $parsed);
	}

	// return zip contents from worksheets
	static function parse($worksheets) 
	{

		$dir = self::makeTempDir();

		$ndir = $dir . "/_rels";
		mkdir($ndir);

		$ndir = $dir . "/docProps";
		mkdir($ndir);

		$ndir = $dir . "/xl";
		mkdir($ndir);

		$ndir = $dir . "/xl/_rels";
		mkdir($ndir);

		$ndir = $dir . "/xl/theme";
		mkdir($ndir);

		$ndir = $dir . "/xl/worksheets";
		mkdir($ndir);

		file_put_contents($dir . "/[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');

		file_put_contents($dir . "/_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');

		file_put_contents($dir . "/docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>15.0300</AppVersion></Properties>');

		file_put_contents($dir . "/docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Ali Ashraf</dc:creator><cp:lastModifiedBy>Ali Ashraf</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2013-07-10T20:19:41Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2013-07-10T20:21:15Z</dcterms:modified></cp:coreProperties>');

		file_put_contents($dir . "/xl/theme/theme1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>');

		file_put_contents($dir . "/xl/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><fonts count="1" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>');

		$shared = array();

		$sheetsStr = "";
		$sheetRelationsStr = "";
		$s = 0;
		foreach($worksheets as $key => $fields) {

			$contents = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:{%MAX_LIMIT%}"/><sheetViews><sheetView tabSelected="1" workbookViewId="{%WORKBOOK_VIEW_ID%}"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><cols>{%COLS_DATA%}</cols><sheetData>{%SHEET_DATA%}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';

			$colWidths = array();
			$limit = self::numToAlpha(sizeof($fields[0])) . sizeof($fields);

			$data = '';

			$data .= '<row r="1" spans="1:' . sizeof($fields[0]) . '" x14ac:dyDescent="0.25">';
			$i = 1;
			foreach($fields[0] as $k => $v) {
				$data .= '
				<c r="' . self::numToAlpha($i) . '1" t="s">
					<v>' . sizeof($shared) . '</v>
				</c>';
				array_push($shared, $k);
				$colWidths[$i] = self::getColWidth($k);
				$i++;
			}
			$data .= '</row>';

			$j = 2;
			foreach($fields as $field) {
				$data .= '<row r="' . $j . '" spans="1:' . sizeof($field) . '" x14ac:dyDescent="0.25">';
				$i = 1;
				foreach($field as $k => $value) {

					$data .= '
					<c r="' . self::numToAlpha($i) . $j . '" t="s" s="1">
						<v>' . sizeof($shared) . '</v>
					</c>';
					array_push($shared, $value);

					$colWidths[$i] = max($colWidths[$i], self::getColWidth($value));

					$i++;
				}
				$data .= '</row>';
				$j++;
			}

			$cols = '';
			$i = 1;
			foreach($fields[0] as $k => $v) {
				$cols .= '<col min="' . $i .'" max="' . $i .'" width="' . $colWidths[$i] . '" customWidth="1"/>';
				$i++;
			}

			$contents = strtr($contents, array(
				"{%MAX_LIMIT%}" => $limit,
				"{%SHEET_DATA%}" => $data,
				"{%COLS_DATA%}" => $cols,
				"{%WORKBOOK_VIEW_ID%}" => 0
			));

			file_put_contents($dir . "/xl/worksheets/sheet" . ($s+1) . ".xml", $contents);

			$sheetsStr .= '<sheet name="' . htmlentities($key) . '" sheetId="' . ($s+1) . '" r:id="rId' . ($s+6) . '"/>';
			$sheetRelationsStr .= '<Relationship Id="rId' . ($s+6) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . ($s+1) . '.xml"/>';
			$s++;
		}


		$sharedStr = "";
		foreach($shared as $item) {
			$sharedStr .= '<si>
				<t>' . htmlspecialchars($item) . '</t>
			</si>';
		}

		file_put_contents($dir . "/xl/sharedStrings.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . sizeof($shared) . '" uniqueCount="' . sizeof($shared) . '">' . $sharedStr . '</sst>');
		
		file_put_contents($dir . "/xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/><workbookPr defaultThemeVersion="153222"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="C:\Xelon-Labs" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="14235"/></bookViews><sheets>' . $sheetsStr . '</sheets><calcPr calcId="152511"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>');

		file_put_contents($dir . "/xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' . $sheetRelationsStr . '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>');

		$newDir = self::makeTempDir();
		$zip = $newDir . "/xlsx.xlsx";

		self::compress_dir($dir, $zip);

		self::delTree($dir);

		$return = file_get_contents($zip);
		unlink($zip);
		rmdir($newDir);
		

		return $return;
	}

	// compress a directory
	// an .xlsx file is actually a zip containing xml data
	static function compress_dir($source, $destination, $regEx = "", $blacklist = array())
	{
		if(!extension_loaded('zip') || !file_exists($source)) return false;

		$zip = new ZipArchive();
		if(!$zip->open($destination, ZIPARCHIVE::CREATE)) return false;
		$source = realpath($source);

		if(is_dir($source) === true) {
			$files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($source), RecursiveIteratorIterator::SELF_FIRST);
			foreach($files as $file){

				if(is_dir($file) === true) {
					$zip->addEmptyDir(str_replace($source . '\\','', $file . '\\'));
				}
				else if(is_file($file) === true) {
					$zip->addFromString(str_replace($source . '\\', '', $file), file_get_contents($file));
				}
			}
		} elseif (is_file($source) === true) {
			$zip->addFromString(basename($source), file_get_contents($source));
		}
		return $zip->close();
	}

	static function readZipFile(&$zip, $file) 
	{
		$strings = $zip->getFromName($file);
		if($strings != "") {
			return $strings;
		} else {
			$file = str_replace("/", "\\", $file);
			$strings = $zip->getFromName($file);
			return $strings;
		}
	}
	
	// recursively delete
	static function delTree($dir) 
	{
		if( substr( $dir, -1 ) != "/" ) {
			$dir .= "/";
		}
		$folders = glob( $dir . '*', GLOB_ONLYDIR );
		foreach( $folders as $folder ){
			self::delTree( $folder );
		}
		$others = scandir($dir);
		$others = array_diff($others, array('..', '.'));
		foreach( $others as $other ) {
			$name = $dir . $other;
			if(!in_array($name, $folders)) {
				@unlink( $name );
			}
		}
		return @rmdir( $dir );
	}

	static function makeTempDir() 
	{
		$dir = "xlsx-" . microtime();
		mkdir($dir);
		return $dir;
	}

	static function makeTempFile() 
	{
		$tempfile = tempnam(sys_get_temp_dir(), 'MyFileName');
		return $tempfile;
	}

	static function numToAlpha($num) 
	{
		$str = "";
		$nnum = $num;

		while($nnum > 0) {
			$remainder = ($nnum-1) % 26;
			$str = chr($remainder + ord("A")) . $str;
			$nnum = floor(($nnum-1) / 26);
		}
		return $str;
	}

	// get spreadsheet column width by the input string
	static function getColWidth($str) 
	{
		$wid = 0;
		$first = explode("\n", $str);
		$first = $first[0];
		$size = self::getStringSize($first);
		return ($size["width"] * (50/355)) + 4;
	}

	// get size box of a given string (11 is the font size)
	static function getStringSize($str, $font = "calibri.ttf")
	{
		$mainDir = str_replace("\\", "/", dirname(__FILE__)) . "/";
		$font = $mainDir . "fonts/" . $font;

		$box = imagettfbbox ( 11 , 0 , $font , $str );
		$width = abs($box[4] - $box[0]);
		$height = abs($box[5] - $box[1]);

		return array(
			"width" => $width,
			"height" => $height
		);
	}

	// read the .xlsx file and return its contents
	public static function read($file)
	{
		$za = new ZipArchive(); 
		$za->open($file); 

		$worksheets = array();
		for( $i = 0; $i < $za->numFiles; $i++ ) { 
		    $stat = $za->statIndex( $i ); 
		    $info = pathinfo($stat["name"]);
		    if(($info["dirname"] == "xl/worksheets" || $info["dirname"] == "xl\\worksheets") && $info["extension"] == "xml") {
		    	$name = basename($info["basename"], ".xml");
		    	array_push($worksheets, array("name" => $name, "dir" => $stat["name"]));
		    }
		}

		$strings = self::readZipFile($za, "xl/sharedStrings.xml");
		list($xmlDoc, $xpath) = self::getHtmlXPath($strings);
		$items = $xpath->evaluate("//si");
		$stringsArr = self::getNodeValuesArr($items);

		$strings = self::readZipFile($za, "xl/workbook.xml");
		list($xmlDoc, $xpath) = self::getHtmlXPath($strings);
		$items = $xpath->evaluate("//sheet");
		$workbookArr = self::getSortedNodeAttrs($items, "name", "sheetid");

		$workbook = array();

		$i = 0;
		foreach($worksheets as $worksheet) {
			$arr = array();
			//$data = self::readZipFile($za, $worksheet["dir"]);
			$data = self::readZipFile($za, "xl/worksheets/sheet" . ($i+1) . ".xml");

			list($xmlDoc, $xpath) = self::getHtmlXPath($data);
			$cells = $xpath->evaluate("//row/c");

			foreach($cells as $cell) {
				$v = ($cell->getElementsByTagName('v')->item(0));
				$r = "";
				$r = $cell->getAttribute('r');
				$s = "";
				$s = $cell->getAttribute('s');
				$t = "";
				$t = @$cell->getAttribute('t');
				$value = "";

				if($t == "b") {
					$value = (bool) @$v->nodeValue;
				} else if(@isset($stringsArr[$v->nodeValue])) {
					$value = $stringsArr[$v->nodeValue];
				} else {
					$value = @$v->nodeValue;
				}
				/* else if($s == "3") {
					$value = date("m/d/y", self::excelToUnixTimestamp($v->nodeValue));
				*/

				if($value == "") {
					continue;
				}

				$position = self::alphaToNum($r);

				if(!isset($arr[$position["row"]])) {
					$arr[$position["row"]] = array();
				}
				$arr[$position["row"]][$position["col"]] = $value;
			}

			$workbook[$workbookArr[$i]] = $arr;
			$i++;
		}


		return $workbook;
	}

	// spreadsheets use a grid system, convert them into an associative array
	static function gridToArray($grid, $key = "")
	{
		$arr = array();
		$keys = array_shift($grid);

		foreach($grid as $row) {
			$rarr = array();
			foreach($keys as $k => $v) {
				$rarr[$v] = "";
				$rarr[$v] = @$row[$k];
			}
			if($key != "") {
				$arr[$rarr[$key]] = $rarr;
			} else {
				array_push($arr, $rarr);
			}
		}

		return $arr;
	}

	// read array from spreadsheet file
	public static function readArray($file, $key = "")
	{
		$grids = self::read($file);
		$arr = array();
		foreach($grids as $k => $grid) {
			$arr[$k] = self::gridToArray($grid, $key);
		}
		return $arr;
	}

	static function alphaToNum($alpha)
	{
		preg_match("/[A-Z]+/", $alpha, $matches);
		$a = $matches[0];
		preg_match("/[0-9]+/", $alpha, $matches);
		$n = 1;
		$n = @$matches[0];

		$r = 0;
		$c = 0;
		for($i = strlen($a) -1; $i >= 0; $i--) {
			$alpha = $a[$i];

			$r += (ord($alpha) - 65 + 1) * pow(26, $c);

			$c++;
		}

		return array("row"=>$n, "col"=>$r);
	}

	// Excel uses a different timestamp format in their spreadsheets
	static function excelToUnixTimestamp($excelTime)
	{
		return ($excelTime - 25569) * 86400;
	}

	// XPaths are used for xml traversal and manipulations
	static function getXmlXPath($contents)
	{
		$xmlDoc = new DOMDocument();
		@$xmlDoc->loadXML($contents);
		$xpath = new DOMXPath($xmlDoc);
		return array($xmlDoc, $xpath);
	}

	private static function getHtmlXPath($contents)
	{
		$contents = mb_convert_encoding($contents, 'HTML-ENTITIES', "UTF-8");

		$xmlDoc = new DOMDocument();
		@$xmlDoc->loadHTML($contents);
		$xpath = new DOMXPath($xmlDoc);
		return array($xmlDoc, $xpath);
	}

	private static function getElementXPath($DE)
	{
		$DD= new DOMDocument('1.0', 'utf-8');
		$DD->loadXML( "<html></html>" );
		$DD->documentElement->appendChild($DD->importNode($DE,true));
		$xpath = new DOMXPath($DD);
		return array($DD, $xpath);
	}

	private static function getNodeValues($domnodelist) {
		$o = "";
		for ($i = 0; $i < $domnodelist->length; ++$i) {
			$o .= $domnodelist->item($i)->nodeValue . "\n";
		}
		return $o;
	}

	private static function getNodeValuesArr($domnodelist) {
		$arr = array();
		for ($i = 0; $i < $domnodelist->length; ++$i) {
			array_push($arr, $domnodelist->item($i)->nodeValue);
		}
		return $arr;
	}

	private static function getNodeAttrs($domnodelist, $attr) {
		$arr = array();
		for ($i = 0; $i < $domnodelist->length; ++$i) {
			array_push($arr, $domnodelist->item($i)->getAttribute($attr));
		}
		return $arr;
	}

	private static function getSortedNodeAttrs($domnodelist, $attr, $sortKeyAttribute) {
		$arr = array();
		for ($i = 0; $i < $domnodelist->length; ++$i) {
			$node = $domnodelist->item($i);

			array_push($arr, array(
					"value" => $node->getAttribute($attr),
					"sort" => $node->getAttribute($sortKeyAttribute),
				)
			);
		}
		//usort($arr, 'self::sortNodes');
		$result = array();
		foreach($arr as $v) {
			array_push($result, $v["value"]);
		}
		return $result;
	}

	private static function sortNodes($w1, $w2) {
		return strcmp($w1["sort"], $w2["sort"]);
	}

}

?>
# PHP_EXTENTIONS_WITHOUT_COMPOSER
 php extention without composer or install

# Demo Example

## First, import the needed library and load the Reader of XLSX.

```php
<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

?>
```

## Read the excel file using the load() function. Here test.xlsx is the file name.
```php
<?php

$spreadsheet = $reader->load("test.xlsx");

?>
```

### Get the first sheet in the Excel file and convert it to an array using the toArray() function. And Get the Number of rows in the sheet using the count() function.
```php
<?php

$d=$spreadsheet->getSheet(0)->toArray();

echo count($d);

?>
```

### If you want to iterate all the rows in the excel file, then first convert it to an array and iterate using for or foreach.
```php
<?php
$sheetData = $spreadsheet->getActiveSheet()->toArray();

$i=1;

unset($sheetData[0]);

foreach ($sheetData as $t) {
 // process element here;
// access column by index
	echo $i."---".$t[0].",".$t[1]." <br>";
	$i++;
}
?>
```

## Full Example Code(Reading Excel)
```php
<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();


$spreadsheet = $reader->load("test.xlsx");

$d=$spreadsheet->getSheet(0)->toArray();

echo count($d);

$sheetData = $spreadsheet->getActiveSheet()->toArray();

$i=1;
unset($sheetData[0]);

foreach ($sheetData as $t) {
 // process element here;

	echo $i."---".$t[0].",".$t[1]." <br>";
	$i++;
}
?>
```

### Get the sheet count using the getSheetCount() function.
```php
echo $spreadsheet->getSheetCount();
```

### While the getSheetNames() method will return a list of all worksheets in the workbook, indexed by the order in which their “tabs” would appear when opened in MS Excel (or other appropriate Spreadsheet programs).
```php
echo $spreadsheet->getSheetNames();
```

## Individual worksheets can be accessed by name, or by their index position in the workbook.
```php
// Get the second sheet in the workbook 
// Note that sheets are indexed from 0 
$sheet = $spreadsheet->getSheet(1);
//or
// Retrieve the worksheet called 'Worksheet 1' 
$sheet = $spreadsheet->getSheetByName('Worksheet 1');
```

# Write Excel File

### Like reading, We need to import the needed files for writing also. This time, import the writer class. 
```php
<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet; 
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

?>
```

### Create a new Spreadsheet using the Spreadsheet class.
```php
<?php
// Creates New Spreadsheet 
$spreadsheet = new Spreadsheet(); 
?>
```

### By default one sheet was added to the Spreadsheet. You can get the first sheet using the getSheet() function.
```php
<?php
// Get the first sheet in the workbook 
// Note that sheets are indexed from 0 
$spreadsheet->getSheet(0);
?>
```

### Alternatively you can get the current sheet in the newly created spreadsheet.
```php
<?php

// Retrieve the current active worksheet 
$sheet = $spreadsheet->getActiveSheet();

?>
```

### Now I am creating a sample array of data to insert in the excel file.
```php
<?php
$data_from_db=array();
$data_from_db[0]=array("name"=>"raja","age"=>23);
$data_from_db[1]=array("name"=>"raja1","age"=>43);
?>
```

### We can set the value to the cell by using the setCellValueByColumnAndRow(a,b,c) function.

It takes three parameter

1.a -Column index
2.b -row index
3.c -Cell value
```php
<?php
//set value row
for($i=0;$i<count($data_from_db);$i++)
{

//set value for indi cell
$row=$data_from_db[$i];

//writing cell index start at 1 not 0
$j=1;

	foreach($row as $x => $x_value) {
		$sheet->setCellValueByColumnAndRow($j,$i+1,$x_value);
  		$j=$j+1;
	}

}
?>
```

### Write the created spreadsheet using the save() function.
```php
<?php
// Write an .xlsx file  
$writer = new Xlsx($spreadsheet); 
  
// Save .xlsx file to the files directory 
$writer->save('files/demo.xlsx'); 
?>
```

### Full Code (Writing Excel File)
```php
<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet; 
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 
  
// Creates New Spreadsheet 
$spreadsheet = new Spreadsheet(); 
  
// Retrieve the current active worksheet 
$sheet = $spreadsheet->getActiveSheet(); 

// sample data from db
// call the db get data function here
//delete line from 18 to 20 and call the db function
$data_from_db=array();
$data_from_db[0]=array("name"=>"raja","age"=>23);
$data_from_db[1]=array("name"=>"raja1","age"=>43);

//set column header
//set your own column header
$column_header=["Name","Age"];
$j=1;
foreach($column_header as $x_value) {
		$sheet->setCellValueByColumnAndRow($j,1,$x_value);
  		$j=$j+1;
  		
	}

//set value row
for($i=0;$i<count($data_from_db);$i++)
{

//set value for indi cell
$row=$data_from_db[$i];

$j=1;

	foreach($row as $x => $x_value) {
		$sheet->setCellValueByColumnAndRow($j,$i+2,$x_value);
  		$j=$j+1;
	}

}
   
// Write an .xlsx file  
$writer = new Xlsx($spreadsheet); 
  
// Save .xlsx file to the files directory 
$writer->save('files/demo.xlsx'); 
?>
```

### We can set the cell value using setCellValue() function() by passing cell index name like A1, A2, D3 etc. and value.
```php
<?php
// Set the value of cell A1 
$sheet->setCellValue('A1', 'Raja!'); 
  
// Sets the value of cell B1 
$sheet->setCellValue('B1', '25');
?>
```

## For more reference : [Documentation]([https://pages.github.com/](https://phpspreadsheet.readthedocs.io/en/latest/))

## I hope you find the tutorial will be helpful.

### In this tutorial, we learned how to read the excel file and write the data to an excel on the server-side using the PHP Spreadsheet library. The Spreadsheet library is very simple and easy to use. Documentation is also very clear. We can do many things using the Spreadsheet library. It offers an excel function to apply to the cell also.

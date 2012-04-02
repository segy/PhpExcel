PhpExcel helper for easy generating XLS files

PHPExcel is a great library that can create XLS files. For more information see PHPEXcel project homepage: http://phpexcel.codeplex.com/

I added method for setting font and for easy data adding (see example).

This plugin is for CakePHP 2.x

Short example: 

Controller:

public $helpers = array('PhpExcel.PhpExcel'); 

View:

$this->PhpExcel->createWorksheet();
$this->PhpExcel->setDefaultFont('Calibri', 12);
// add data - starting on first row
$this->PhpExcel->addRow(array('ccc','ddd'));
// skip to 4th row
$this->PhpExcel->setRow(4);
// add data starting from column AC
$this->PhpExcel->addRow(array('fff','ggg'), 'AC');
// add data starting from 5th column E
$this->PhpExcel->addRow(array('iii','jjj'), 4);
// output to browser
$this->PhpExcel->output();
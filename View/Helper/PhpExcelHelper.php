<?php
App::uses('AppHelper', 'Helper');

/**
 * Helper for working with PHPExcel class.
 * PHPExcel has to be in the vendors directory.
 */

class PhpExcelHelper extends AppHelper {
	/**
	 * Instance of PHPExcel class
	 * @var object
	 */
	public $xls;
	/**
	 * Pointer to actual row
	 * @var int
	 */
	protected $row = 1;
	/**
	 * Array of excel columns for simplified access 
	 * @var array
	 */
	protected $columns = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
	
	/**
	 * Constructor
	 */
	public function __construct(View $view, $settings = array()) {
        parent::__construct($view, $settings);
    }
	
	/**
	 * Create new worksheet
	 */
	public function createWorksheet() {
		$this->loadEssentials();
		$this->xls = new PHPExcel();
	}
	
	/**
	 * Create new worksheet from existing file
	 */
	public function loadWorksheet($path) {
		$this->loadEssentials();
		$this->xls = PHPExcel_IOFactory::load($path);
	}
	
	/**
	 * Load vendor classes
	 */
	protected function loadEssentials() {
		// load vendor class
		App::import('Vendor', 'PHPExcel', array('file' => 'phpexcel/Classes/PHPExcel.php')); 
		if (!class_exists('PHPExcel'))
			throw new CakeException('Vendor class PHPExcel not found!');
	}
	
	/**
	 * Set row pointer
	 */
	public function setRow($to) {
		$this->row = (int)$to;
	}
	
	/**
	 * Set default font
	 */
	public function setDefaultFont($name, $size) {
		$this->xls->getDefaultStyle()->getFont()->setName($name);
		$this->xls->getDefaultStyle()->getFont()->setSize($size);
	}
	
	/**
	 * Write array of data to actual row starting from column defined by offset
	 * Offset can be textual or numeric representation
	 */
	public function addRow($data, $offset = 0) {
		// solve textual representation
		if (!is_numeric($offset))
			$offset = $this->columnNumber($offset);
		
		foreach ($data as $d) {
			$this->xls->getActiveSheet()->setCellValueByColumnAndRow($offset++, $this->row, $d);
		}
		$this->row++;
	}
	
	/**
	 * Output file to browser
	 */
	public function output($filename = 'export.xlsx') {
		// set layout
		$this->_View->layout = '';
		// headers
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="'.$filename.'"');
		header('Cache-Control: max-age=0');
		// writer
		$objWriter = PHPExcel_IOFactory::createWriter($this->xls, 'Excel2007');
		$objWriter->save('php://output');
	}
	
	/**
	 * Textual representation of column number
	 */
	protected function columnText($col) {
		$count = count($this->columns);
		$pref = floor($col / $count);
		return ($pref > 0 ? $this->columns[$pref - 1] : '').$this->columns[$col % $count]; 
	}
	
	/**
	 * Numeric representation of column text
	 */
	protected function columnNumber($col) {
		// sanity check
		if (strlen($col) > 2)
			return 0;
		
		$count = count($this->columns);
		if (strlen($col) > 1) {
			$pref = array_search($col{0}, $this->columns) + 1;
			$col = $col{1};
		}
		else
			$pref = 0;
		
		return $pref * $count + array_search($col, $this->columns); 
	}
}
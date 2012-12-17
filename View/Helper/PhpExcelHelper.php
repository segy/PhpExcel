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
	 * Internal table params 
	 * @var array
	 */
	protected $tableParams;
	
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
	 * Start table
	 * inserts table header and sets table params
	 * Possible keys for data:
	 * 	label 	-	table heading
	 * 	width	-	"auto" or units
	 * 	filter	-	true to set excel filter for column
	 * 	wrap	-	true to wrap text in column
	 * Possible keys for params:
	 * 	offset	-	column offset (numeric or text)
	 * 	font	-	font name
	 * 	size	-	font size
	 * 	bold	-	true for bold text
	 * 	italic	-	true for italic text
	 * 	
	 */
	public function addTableHeader($data, $params = array()) {
		// offset
		if (array_key_exists('offset', $params))
			$offset = is_numeric($params['offset']) ? (int)$params['offset'] : PHPExcel_Cell::columnIndexFromString($params['offset']);
		// font name
		if (array_key_exists('font', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setName($params['font_name']);
		// font size
		if (array_key_exists('size', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setSize($params['font_size']);
		// bold
		if (array_key_exists('bold', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setBold($params['bold']);
		// italic
		if (array_key_exists('italic', $params))
			$this->xls->getActiveSheet()->getStyle($this->row)->getFont()->setItalic($params['italic']);
		
		// set internal params that need to be processed after data are inserted
		$this->tableParams = array(
			'header_row' => $this->row,
			'offset' => $offset,
			'row_count' => 0,
			'auto_width' => array(),
			'filter' => array(),
			'wrap' => array()
		);
		
		foreach ($data as $d) {
			// set label
			$this->xls->getActiveSheet()->setCellValueByColumnAndRow($offset, $this->row, $d['label']);
			// set width
			if (array_key_exists('width', $d)) {
				if ($d['width'] == 'auto')
					$this->tableParams['auto_width'][] = $offset;
				else
					$this->xls->getActiveSheet()->getColumnDimensionByColumn($offset)->setWidth((float)$d['width']);
			}
			// filter
			if (array_key_exists('filter', $d) && $d['filter'])
				$this->tableParams['filter'][] = $offset;
			// wrap
			if (array_key_exists('wrap', $d) && $d['wrap'])
				$this->tableParams['wrap'][] = $offset;
			
			$offset++;
		}
		$this->row++;	
	}
	
	/**
	 * Write array of data to actual row
	 */
	public function addTableRow($data) {
		$offset = $this->tableParams['offset'];
		
		foreach ($data as $d) {
			$this->xls->getActiveSheet()->setCellValueByColumnAndRow($offset++, $this->row, $d);
		}
		$this->row++;
		$this->tableParams['row_count']++;
	}
	
	/**
	 * End table
	 * sets params and styles that required data to be inserted
	 */
	public function addTableFooter() {
		// auto width
		foreach ($this->tableParams['auto_width'] as $col)
			$this->xls->getActiveSheet()->getColumnDimensionByColumn($col)->setAutoSize(true);
		// filter (has to be set for whole range)
		if (count($this->tableParams['filter']))
			$this->xls->getActiveSheet()->setAutoFilter(PHPExcel_Cell::stringFromColumnIndex($this->tableParams['filter'][0]).($this->tableParams['header_row']).':'.PHPExcel_Cell::stringFromColumnIndex($this->tableParams['filter'][count($this->tableParams['filter']) - 1]).($this->tableParams['header_row'] + $this->tableParams['row_count']));
		// wrap
		foreach ($this->tableParams['wrap'] as $col)
			$this->xls->getActiveSheet()->getStyle(PHPExcel_Cell::stringFromColumnIndex($col).($this->tableParams['header_row'] + 1).':'.PHPExcel_Cell::stringFromColumnIndex($col).($this->tableParams['header_row'] + $this->tableParams['row_count']))->getAlignment()->setWrapText(true);
	}
	
	/**
	 * Write array of data to actual row starting from column defined by offset
	 * Offset can be textual or numeric representation
	 */
	public function addData($data, $offset = 0) {
		// solve textual representation
		if (!is_numeric($offset))
			$offset = PHPExcel_Cell::columnIndexFromString($offset);
		
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
		// clear memory
		$this->xls->disconnectWorksheets();
	}
	
	/**
	 * Load vendor classes
	 */
	protected function loadEssentials() {
		// load vendor class
		App::import('Vendor', 'PHPExcel/Classes/PHPExcel');
		if (!class_exists('PHPExcel')) {
			throw new CakeException('Vendor class PHPExcel not found!');
		}
	}
}
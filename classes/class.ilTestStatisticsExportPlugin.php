<?php
/* Copyright (c) 1998-2013 ILIAS open source, Extended GPL, see docs/LICENSE */

require_once 'Modules/Test/classes/class.ilTestExportPlugin.php';

/**
 * Abstract parent class for all event hook plugin classes.
 * @author  Christoph Jobst <cjobst@wifa.uni-leipzig.de>
 * @version $Id$
 * @ingroup ModulesTest
 */
class ilTestStatisticsExportPlugin extends ilTestExportPlugin
{
	/**
	 * Get Plugin Name. Must be same as in class name il<Name>Plugin
	 * and must correspond to plugins subdirectory name.
	 * Must be overwritten in plugin class of plugin
	 * (and should be made final)
	 * @return    string    Plugin Name
	 */
	function getPluginName()
	{
		return 'TestStatisticsExport';
	}

	/**
	 * @return string
	 */
	protected function getFormatIdentifier()
	{
		return 'statistics.xlsx';
	}

	/**
	 * @return string
	 */
	public function getFormatLabel()
	{
		return $this->txt('statisticsexport_format');
	}

	/**
	 * @param ilTestExportFilename $filename
	 */
	protected function buildExportFile(ilTestExportFilename $filename)
	{
		//Creating Files with Charts using PHPExcel
		
		require_once './Customizing/global/plugins/Modules/Test/Export/TestStatisticsExport/classes/PHPExcel-1.8/Classes/PHPExcel.php';
		require_once './Customizing/global/plugins/Modules/Test/Export/TestStatisticsExport/classes/PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';
				
		// Create new PHPExcel object
		$objPHPExcel = new PHPExcel();

		//CONTENT #######################################################################################
		
		// Set properties
		$objPHPExcel->getProperties()->setCreator("Chris");
		$objPHPExcel->getProperties()->setLastModifiedBy("Chris");
		$objPHPExcel->getProperties()->setTitle("Office 2007 XLSX  Document");
		$objPHPExcel->getProperties()->setSubject("Office 2007 XLSX  Document");
		$objPHPExcel->getProperties()->setDescription("XSLX Document for Office 2007, generated using PHPExcel.");
		
		// Add data
		$objPHPExcel->setActiveSheetIndex(0);
		$objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Hello');
		$objPHPExcel->getActiveSheet()->SetCellValue('B2', 'world!');
		$objPHPExcel->getActiveSheet()->SetCellValue('C1', 'Hello');
		$objPHPExcel->getActiveSheet()->SetCellValue('D2', 'world!');
		
		// Rename sheet
		$objPHPExcel->getActiveSheet()->setTitle('Simple');

		//END CONTENT #######################################################################################
		
		// Save Excel 2007 file		
		ilUtil::makeDirParents(dirname($filename->getPathname('xls', 'statistics')));
		$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
		$objWriter->save(str_replace(__FILE__, $filename->getPathname('xlsx', 'statistics'),__FILE__));

		//PEAR Spreadsheet - ILIAS internal Excelgenerator
		/*
		require_once './Services/Excel/classes/class.ilExcelWriterAdapter.php';
		
		ilUtil::makeDirParents(dirname($filename->getPathname('xls', 'statistics')));
		
		$excelfile = $filename->getPathname('xls', 'statistics');
		$adapter = new ilExcelWriterAdapter($excelfile, FALSE);

		//$testname = $this->test_obj->getTitle();
		$testname = "helloworld";
		$testname = ilUtil::getASCIIFilename(preg_replace("/\s/", "_", $testname)) . ".xls";

		$workbook = $adapter->getWorkbook();
	
		// sending HTTP headers
		$workbook->send('test.xls');
		
		// Creating a worksheet
		$worksheet =& $workbook->addWorksheet('My first worksheet');
		
		// The actual data
		$worksheet->write(0, 0, 'Name');
		$worksheet->write(0, 1, 'Age');
		$worksheet->write(1, 0, 'John Smith');
		$worksheet->write(1, 1, 30);
		$worksheet->write(2, 0, 'Johann Smitty-Schmidt');
		$worksheet->write(2, 1, 31);
		$worksheet->write(3, 0, 'Juan Herrera');
		$worksheet->write(3, 1, 32);
		
		// Let's send the file
		$workbook->close();
		*/
	}
}
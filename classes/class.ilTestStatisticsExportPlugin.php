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
		return 'teststatistics.csv';
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
	}
}
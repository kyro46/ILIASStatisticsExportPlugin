<?php
/* Copyright (c) 1998-2013 ILIAS open source, Extended GPL, see docs/LICENSE */
require_once 'Modules/Test/classes/class.ilTestExportPlugin.php';

/**
 * Abstract parent class for all event hook plugin classes.
 *
 * @author Christoph Jobst <cjobst@wifa.uni-leipzig.de>
 * @version $Id$
 *          @ingroup ModulesTest
 */
class ilTestStatisticsExportPlugin extends ilTestExportPlugin {
	/**
	 * Get Plugin Name.
	 * Must be same as in class name il<Name>Plugin
	 * and must correspond to plugins subdirectory name.
	 * Must be overwritten in plugin class of plugin
	 * (and should be made final)
	 *
	 * @return string Plugin Name
	 */
	function getPluginName() {
		return 'TestStatisticsExport';
	}
	
	/**
	 *
	 * @return string
	 */
	protected function getFormatIdentifier() {
		return 'statistics.xlsx';
	}
	
	/**
	 *
	 * @return string
	 */
	public function getFormatLabel() {
		return $this->txt ( 'statisticsexport_format' );
	}
	
	/**
	 *
	 * @param ilTestExportFilename $filename        	
	 */
	protected function buildExportFile(ilTestExportFilename $filename) {
		// Creating Files with Charts using PHPExcel
		require_once './Customizing/global/plugins/Modules/Test/Export/TestStatisticsExport/classes/PHPExcel-1.8/Classes/PHPExcel.php';
		// require_once './Customizing/global/plugins/Modules/Test/Export/TestStatisticsExport/classes/PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';
		
		// Create PHPExcel object
		$objPHPExcel = new PHPExcel ();
		
		// Create the first sheet with general data about the test
		$objWorksheet = $objPHPExcel->getActiveSheet ();
		$this->createFirstSheet ( $objWorksheet );
		
		// Create the second sheet, with questionspecific data
		$objPHPExcel->addSheet ( $this->createRawDataDifficulty () );
		
		// Create the third sheet, with questionspecific data for discrimination index
		$objPHPExcel->addSheet ( $this->createDiscriminationIndex () );
		
		// Save XSLX file
		ilUtil::makeDirParents ( dirname ( $filename->getPathname ( 'xlsx', 'statistics' ) ) );
		$objWriter = PHPExcel_IOFactory::createWriter ( $objPHPExcel, 'Excel2007' );
		$objWriter->setIncludeCharts ( TRUE );
		$objWriter->save ( str_replace ( __FILE__, $filename->getPathname ( 'xlsx', 'statistics' ), __FILE__ ) );
	}
	
	/**
	 * Takes a reference to the first active sheet and modifies it
	 *
	 * @param
	 *        	PHPExcel_Worksheet &$objWorksheet
	 */
	public function createFirstSheet(&$objWorksheet) {
		$objWorksheet->fromArray ( array (
				array (
						'',
						2010,
						2011,
						2012 
				),
				array (
						'Q1',
						12,
						15,
						21 
				),
				array (
						'Q2',
						56,
						73,
						86 
				),
				array (
						'Q3',
						52,
						61,
						69 
				),
				array (
						'Q4',
						30,
						32,
						0 
				) 
		) );
		
		// Set the Labels for each data series we want to plot
		// Datatype
		// Cell reference for data
		// Format Code
		// Number of datapoints in series
		// Data values
		// Data Marker
		$dataSeriesLabels1 = array (
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$B$1', NULL, 1 ), // 2010
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$C$1', NULL, 1 ), // 2011
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$D$1', NULL, 1 )  // 2012
				);
		// Set the X-Axis Labels
		// Datatype
		// Cell reference for data
		// Format Code
		// Number of datapoints in series
		// Data values
		// Data Marker
		$xAxisTickValues1 = array (
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$A$2:$A$5', NULL, 4 )  // Q1 to Q4
				);
		// Set the Data values for each data series we want to plot
		// Datatype
		// Cell reference for data
		// Format Code
		// Number of datapoints in series
		// Data values
		// Data Marker
		$dataSeriesValues1 = array (
				new PHPExcel_Chart_DataSeriesValues ( 'Number', 'Worksheet!$B$2:$B$5', NULL, 4 ),
				new PHPExcel_Chart_DataSeriesValues ( 'Number', 'Worksheet!$C$2:$C$5', NULL, 4 ),
				new PHPExcel_Chart_DataSeriesValues ( 'Number', 'Worksheet!$D$2:$D$5', NULL, 4 ) 
		);
		
		// Build the dataseries
		$series1 = new PHPExcel_Chart_DataSeries ( PHPExcel_Chart_DataSeries::TYPE_AREACHART, 		// plotType
		PHPExcel_Chart_DataSeries::GROUPING_PERCENT_STACKED, 		// plotGrouping
		range ( 0, count ( $dataSeriesValues1 ) - 1 ), 		// plotOrder
		$dataSeriesLabels1, 		// plotLabel
		$xAxisTickValues1, 		// plotCategory
		$dataSeriesValues1 ); // plotValues
		                      
		// Set the series in the plot area
		$plotArea1 = new PHPExcel_Chart_PlotArea ( NULL, array (
				$series1 
		) );
		// Set the chart legend
		$legend1 = new PHPExcel_Chart_Legend ( PHPExcel_Chart_Legend::POSITION_TOPRIGHT, NULL, false );
		
		$title1 = new PHPExcel_Chart_Title ( 'Test %age-Stacked Area Chart' );
		$yAxisLabel1 = new PHPExcel_Chart_Title ( 'Value ($k)' );
		
		// Create the chart
		$chart1 = new PHPExcel_Chart ( 'chart1', 		// name
		$title1, 		// title
		$legend1, 		// legend
		$plotArea1, 		// plotArea
		true, 		// plotVisibleOnly
		0, 		// displayBlanksAs
		NULL, 		// xAxisLabel
		$yAxisLabel1 ); // yAxisLabel
		                
		// Set the position where the chart should appear in the worksheet
		$chart1->setTopLeftPosition ( 'A7' );
		$chart1->setBottomRightPosition ( 'H20' );
		
		// Add the chart to the worksheet
		$objWorksheet->addChart ( $chart1 );
		
		// Set the Labels for each data series we want to plot
		// Datatype
		// Cell reference for data
		// Format Code
		// Number of datapoints in series
		// Data values
		// Data Marker
		$dataSeriesLabels2 = array (
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$B$1', NULL, 1 ), // 2010
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$C$1', NULL, 1 ), // 2011
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$D$1', NULL, 1 )  // 2012
				);
		// Set the X-Axis Labels
		// Datatype
		// Cell reference for data
		// Format Code
		// Number of datapoints in series
		// Data values
		// Data Marker
		$xAxisTickValues2 = array (
				new PHPExcel_Chart_DataSeriesValues ( 'String', 'Worksheet!$A$2:$A$5', NULL, 4 )  // Q1 to Q4
				);
		// Set the Data values for each data series we want to plot
		// Datatype
		// Cell reference for data
		// Format Code
		// Number of datapoints in series
		// Data values
		// Data Marker
		$dataSeriesValues2 = array (
				new PHPExcel_Chart_DataSeriesValues ( 'Number', 'Worksheet!$B$2:$B$5', NULL, 4 ),
				new PHPExcel_Chart_DataSeriesValues ( 'Number', 'Worksheet!$C$2:$C$5', NULL, 4 ),
				new PHPExcel_Chart_DataSeriesValues ( 'Number', 'Worksheet!$D$2:$D$5', NULL, 4 ) 
		);
		
		// Build the dataseries
		$series2 = new PHPExcel_Chart_DataSeries ( PHPExcel_Chart_DataSeries::TYPE_BARCHART, 		// plotType
		PHPExcel_Chart_DataSeries::GROUPING_STANDARD, 		// plotGrouping
		range ( 0, count ( $dataSeriesValues2 ) - 1 ), 		// plotOrder
		$dataSeriesLabels2, 		// plotLabel
		$xAxisTickValues2, 		// plotCategory
		$dataSeriesValues2 ); // plotValues
		                      
		// Set additional dataseries parameters
		                      // Make it a vertical column rather than a horizontal bar graph
		$series2->setPlotDirection ( PHPExcel_Chart_DataSeries::DIRECTION_COL );
		
		// Set the series in the plot area
		$plotArea2 = new PHPExcel_Chart_PlotArea ( NULL, array (
				$series2 
		) );
		// Set the chart legend
		$legend2 = new PHPExcel_Chart_Legend ( PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false );
		
		$title2 = new PHPExcel_Chart_Title ( 'Test Column Chart' );
		$yAxisLabel2 = new PHPExcel_Chart_Title ( 'Value ($k)' );
		
		// Create the chart
		$chart2 = new PHPExcel_Chart ( 'chart2', 		// name
		$title2, 		// title
		$legend2, 		// legend
		$plotArea2, 		// plotArea
		true, 		// plotVisibleOnly
		0, 		// displayBlanksAs
		NULL, 		// xAxisLabel
		$yAxisLabel2 ); // yAxisLabel
		                
		// Set the position where the chart should appear in the worksheet
		$chart2->setTopLeftPosition ( 'I7' );
		$chart2->setBottomRightPosition ( 'P20' );
		
		// Add the chart to the worksheet
		$objWorksheet->addChart ( $chart2 );
	}
	
	/**
	 * creates another worksheet
	 *
	 * @return PHPExcel_Worksheet objWorksheet
	 */
	public function createRawDataDifficulty() {
		$objWorksheet = new PHPExcel_Worksheet ( $objPHPExcel );
		$objWorksheet->setTitle ( 'Rohwerte, Schwierigkeitsindex' );
		
		$titles = $this->getTest ()->getQuestionTitlesAndIndexes ();
		//$oids = $this->getTest ()->getQuestions ();
		//asort ( $oids );
		
		$rows = array ();
		$datarow = array ();
		array_push ( $datarow, $this->getTest ()->lng->txt ( "tst_stat_result_resultspoints" ) );
		array_push ( $datarow, $this->getTest ()->lng->txt ( "maximum_points" ) );
		array_push ( $datarow, $this->getTest ()->lng->txt ( "tst_stat_result_qworkedthrough" ) );
		array_push ( $datarow, $this->getTest ()->lng->txt ( "tst_stat_result_qmax" ) );
		array_push ( $datarow, $this->getTest ()->lng->txt ( "tst_stat_result_timeofwork" ) );
		array_push ( $datarow, $this->getTest ()->lng->txt ( "tst_stat_result_atimeofwork" ) );
		// Nötig - aber nicht als eigene Spalte
		// array_push ( $datarow, $this->getTest ()->lng->txt ( "tst_stat_result_total_participants" ) );
		array_push ( $datarow, $this->getTest ()->lng->txt ( "scored_pass" ) );
		// Da nur bewerteter Durchlauf angezeigt - nicht notwendig
		// array_push ( $datarow, $this->getTest ()->lng->txt ( "pass" ) );
		$data = & $this->getTest ()->getCompleteEvaluationData ( TRUE, $filterby, $filtertext );
		
		$counter = 1;
		$maximumPoints = array ();
		foreach ( $data->getParticipants () as $active_id => $userdata ) {
			$remove = FALSE;
			if ($passedonly) {
				if ($data->getParticipant ( $active_id )->getPassed () == FALSE) {
					$remove = TRUE;
				}
			}
			if (! $remove) {
				$datarow2 = array ();
				
				array_push ( $datarow2, $data->getParticipant ( $active_id )->getReached () );
				array_push ( $datarow2, $data->getParticipant ( $active_id )->getMaxpoints () );
				
				array_push ( $datarow2, $data->getParticipant ( $active_id )->getQuestionsWorkedThrough () );
				array_push ( $datarow2, $data->getParticipant ( $active_id )->getNumberOfQuestions () );
				$time = $data->getParticipant ( $active_id )->getTimeOfWork ();
				$time_seconds = $time;
				$time_hours = floor ( $time_seconds / 3600 );
				$time_seconds -= $time_hours * 3600;
				$time_minutes = floor ( $time_seconds / 60 );
				$time_seconds -= $time_minutes * 60;
				array_push ( $datarow2, sprintf ( "%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds ) );
				$time = $data->getParticipant ( $active_id )->getQuestionsWorkedThrough () ? $data->getParticipant ( $active_id )->getTimeOfWork () / $data->getParticipant ( $active_id )->getQuestionsWorkedThrough () : 0;
				$time_seconds = $time;
				$time_hours = floor ( $time_seconds / 3600 );
				$time_seconds -= $time_hours * 3600;
				$time_minutes = floor ( $time_seconds / 60 );
				$time_seconds -= $time_minutes * 60;
				array_push ( $datarow2, sprintf ( "%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds ) );
				
				// Anzahl der Teilnehmer: tst_stat_result_total_participants
				// array_push ( $datarow2, $data->getStatistics ()->getStatistics ()->count () );
				
				// Aenderung: merke dir den bewerteten Durchlauf - nur der soll angezeigt werden
				// $pass tritt an die Stelle des for-Schleifen $pass
				$pass = 0;
				if ($this->getTest ()->getPassScoring () == SCORE_BEST_PASS) {
					array_push ( $datarow2, $data->getParticipant ( $active_id )->getBestPass () + 1 );
					// Entweder der beste Durchlauf...
					$pass = $data->getParticipant ( $active_id )->getBestPass ();
				} else {
					array_push ( $datarow2, $data->getParticipant ( $active_id )->getLastPass () + 1 );
					// Oder der letzte Durchlauf
					$pass = $data->getParticipant ( $active_id )->getLastPass ();
				}
				// Nicht mehr die Durchlaeufe durchlaufen
				// for($pass = 0; $pass <= $data->getParticipant ( $active_id )->getLastPass (); $pass ++) {
				$finishdate = $this->getTest ()->getPassFinishDate ( $active_id, $pass );
				if ($finishdate > 0) {
					/*
					 * if ($pass > 0) { for($i = 1; $i < $col - 1; $i ++) { array_push ( $datarow2, "" ); array_push ( $datarow, "" ); } array_push ( $datarow, "" ); } array_push ( $datarow2, $pass + 1 );
					 */
					if (is_object ( $data->getParticipant ( $active_id ) ) && is_array ( $data->getParticipant ( $active_id )->getQuestions ( $pass ) )) {
						$sortedquestions = array ();
						$sortedquestions = $data->getParticipant ( $active_id )->getQuestions ( $pass );
						asort ( $sortedquestions );
						
						foreach ( $sortedquestions as $question ) {
							$question_data = $data->getParticipant ( $active_id )->getPass ( $pass )->getAnsweredQuestionByQuestionId ( $question ["id"] );
							array_push ( $datarow2, $question_data ["reached"] );
							array_push ( $datarow, preg_replace ( "/<.*?>/", "", $data->getQuestionTitle ( $question ["id"] ) . " (ID=" . $question ["id"] . ")" ) );
							if ($counter == 1) {
								array_push ( $maximumPoints, $question ["points"] );
							}
						}
					}
					// Nur gewerteter Durchlauf - $pass keine Bedingung mehr um Kopfzeile zu schreiben
					if ($this->getTest ()->isRandomTest () || ($counter == 1)) { // && $pass == 0)) {
						array_push ( $rows, $datarow );
					}
					$datarow = array ();
					array_push ( $rows, $datarow2 );
					$datarow2 = array ();
				}
				// } //for
				$counter ++;
			}
		}
		$objWorksheet->fromArray ( $rows, null, 'A1', true );
		
		/*
		 * Ab hier die statistischen Grundwerte 
		 * Anzahl Teilnehmer = $counter -1 
		 * A1-G1 - Kopfzeile 
		 * H1-count($sortedquestions) 
		 * A2-G2 bis A[$counter]-G[$counter] 
		 * Datenzeilen falls Datenfeld leer: nicht bearbeitet 
		 * 
		 * Kurzberechnung: 
		 * $objWorksheet->getHighestColumn(); 
		 * $objWorksheet->getHighestRow();
		 */
		
		$lastRowOfRawData = $objWorksheet->getHighestRow ();
		$lastDataColumn = $objWorksheet->getHighestColumn ();
		//error_log('Last Column: ' . $lastDataColumn);
		
		//Zeile für maximal erreichbare Punktzahl der Aufgabe
		$objWorksheet->setCellValue ( 'G' . ($lastRowOfRawData + 1), 'Maximale Aufgabenpunkte' );
		$objWorksheet->fromArray ( $maximumPoints, null, 'H' . ($lastRowOfRawData + 1) );

		//Zeile für die Anzahl nicht angezeigter Aufgaben
		$objWorksheet->setCellValue ( 'G' . ($lastRowOfRawData + 2), 'Nicht angezeigte Aufgaben' );
		$notShownItems = array();
		$lastDataColumn ++; // Damit die letzte Spalte auch einbezogen wird
		for($column = 'H'; $column != ($lastDataColumn); $column ++) {
			// Anzahl nicht dargestellter Aufgaben ermitteln
			$countNotShown = 0;
			for($row = 2; $row <= $lastRowOfRawData; ++ $row) {
				// Fetch the data of the columns needed
				$value = $objWorksheet->getCell( $column . $row)->getValue();
				if ($value === NULL || $value === '') {
					$countNotShown += 1;
					//error_log('Not Shown Increment to: ' . $countNotShown . 'in Column ' . $column);	
				}
			}
			array_push ( $notShownItems, $countNotShown );
		}
		$objWorksheet->fromArray ( $notShownItems, null, 'H' . ($lastRowOfRawData + 2), true );
		
		
		
		//Schwierigkeitsindex
		$writeRow = $lastRowOfRawData + 3;
		// schon für Anzahl ausgelassener Aufgaben getan: $lastDataColumn ++;
		$objWorksheet->setCellValue ( 'G' . $writeRow, 'Schwierigkeitsindex' );
		for($column = 'H'; $column != ($lastDataColumn); $column ++) {
			/*
			 * Schwierigkeitsindex ermitteln, dabei max. erreichbare Punktzahl und Beantwortugnsstatus nutzen 
			 * Formel nach Lienert 1994, Seite 76, Formel 6.4, modifiziert für Teilpunkte
			 * Schwierigkekit = 100 * (Summe der erreichten Punkte / Summe der erreichbaren Punkte) 
			 * Summe der erreichbaren Punkte = (Anzahl TN - nicht bearbeitete Aufgaben) * maximale Punktzahl
			 */
			//error_log('Current Schwierigkeit Column: ' . $column);
				
			$maxPoints = $objWorksheet->getCell ( $column . ($writeRow - 2) )->getValue (); // Achtung, Positionsabhaengig
			$anzahlTN = $lastRowOfRawData - 1;
			// Anzahl nicht dargestellter Aufgaben ermitteln
			$countNotShown = 0;
			for($row = 2; $row <= $lastRowOfRawData; ++ $row) {
				// Fetch the data of the columns needed
				$value = $objWorksheet->getCell( $column . $row)->getValue();
				if ($value === NULL || $value === '') {
					$countNotShown += 1;
					//error_log('Not Shown Increment to: ' . $countNotShown . 'in Column ' . $column);
						
				}
			}
			$cell = $objWorksheet->getCell ( $column . $writeRow );
			
			if (($anzahlTN - $countNotShown) == 0) {
				$cell->setValue ( 'n.D.' );
						
			} else {
				$cell->setValue ( '=100*(SUM(' . $column . '2:' . $column . $lastRowOfRawData . ')/' . (($anzahlTN - $countNotShown) * $maxPoints) . ')' );
			}
			//error_log('TN:' . $anzahlTN);
			//error_log('Count:' . $countNotShown);
			//error_log('MaxPoints:' . $maxPoints);
				
		}

		/* 
		 * Trennschaerfe
		 * Unter Beachtung eventuell unvollstaendiger Aufgabenbearbeitung
		 * Formel nach Lienert 1994, S.82, Formel 6.12, modifiziert für Teilpunkte
		 * 
		 * -> Kommt aufgrund der umfangreichen Vortabelle in ein separates Worksheet
		 */
		
		//Breite der Spalten automatisch anpassen
		foreach ( range ( 'A', $objWorksheet->getHighestColumn () ) as $columnID ) {
			$objWorksheet->getColumnDimension ( $columnID )->setAutoSize ( true );
		}
		
		// $objWorksheet->setCellValue('A2' , 'mc');
		return $objWorksheet;
	}
	
	/**
	 * creates another worksheet
	 *
	 * @return PHPExcel_Worksheet objWorksheet
	 */
	public function createDiscriminationIndex() {
		$objWorksheet = new PHPExcel_Worksheet ( $objPHPExcel );
		$objWorksheet->setTitle ( 'Trennschaerfeindex' );
		
		
		
		
		//Breite der Spalten automatisch anpassen
		foreach ( range ( 'A', $objWorksheet->getHighestColumn () ) as $columnID ) {
			$objWorksheet->getColumnDimension ( $columnID )->setAutoSize ( true );
		}
		
		return $objWorksheet;
	}
	
	
	
	
}
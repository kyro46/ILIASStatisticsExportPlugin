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
		
		$lastDataColumnCopy = $lastDataColumn;
		
		//Zeile für maximal erreichbare Punktzahl der Aufgabe
		$objWorksheet->setCellValue ( 'G' . ($lastRowOfRawData + 1), 'Maximale Aufgabenpunkte' );
		$objWorksheet->fromArray ( $maximumPoints, null, 'H' . ($lastRowOfRawData + 1) );

		//Zeile für die Anzahl nicht angezeigter Aufgaben
		$objWorksheet->setCellValue ( 'G' . ($lastRowOfRawData + 2), 'Nicht angezeigte Aufgaben' );
		$notShownItems = array();
		$lastDataColumnCopy ++; // Damit die letzte Spalte auch einbezogen wird
		for($column = 'H'; $column != ($lastDataColumnCopy); $column ++) {
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
		// schon für Anzahl ausgelassener Aufgaben getan: $lastDataColumnCopy ++;
		$objWorksheet->setCellValue ( 'G' . $writeRow, 'Schwierigkeitsindex' );
		for($column = 'H'; $column != ($lastDataColumnCopy); $column ++) {
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
		 */
		$questions = $this->getTest ()->getQuestions ();

		//Erste Titelzeile
		$headerrow1 = array ();
		array_push ( $headerrow1, 'Gesamtpunktzahl' );
		array_push ( $headerrow1, 'TN die den jeweilige Punktzahl erreicht haben' );
		array_push ( $headerrow1, ' ' );
		array_push ( $headerrow1, ' ' );
		
		$counter = 7;
		foreach ( $questions as $question ) {
		
			$title = $objWorksheet->getCellByColumnAndRow($counter, 1)->getValue();
				
			array_push ( $headerrow1, 'TN Aufgabe ' . $title . ' richtig' );
			array_push ( $headerrow1, ' ' );
		
			array_push ( $headerrow1, 'TN Aufgabe ' . $title . ' bearbeitet' );
			array_push ( $headerrow1, ' ' );
			array_push ( $headerrow1, ' ' );
		
			array_push ( $headerrow1, 'TN Aufgabe ' . $title . ' falsch oder nicht erreicht' );
			array_push ( $headerrow1, ' ' );
			
			$counter++;
		}
		$objWorksheet->fromArray ( $headerrow1 , null, 'A'.($lastRowOfRawData+6), true );
		
		//Zweite Titelzeile
		$headerrow2 = array ();
		array_push ( $headerrow2, 'X' );
		array_push ( $headerrow2, 'f' );
		array_push ( $headerrow2, 'fX' );
		array_push ( $headerrow2, 'fX^2' );
		
		$counter = 7;
		foreach ( $questions as $question ) {
		
			$title = $objWorksheet->getCellByColumnAndRow($counter, 1)->getValue();
		
			array_push ( $headerrow2, 'f.r' );
			array_push ( $headerrow2, 'f.rX' );
		
			array_push ( $headerrow2, 'f.b' );
			array_push ( $headerrow2, 'f.bX' );
			array_push ( $headerrow2, 'f.bX^2' );
		
			array_push ( $headerrow2, 'f.f, f.a' );
			array_push ( $headerrow2, 'f.u' );
				
			$counter++;
		}
		$objWorksheet->fromArray ( $headerrow2 , null, 'A'.($lastRowOfRawData+7), true );
		
		//error_log($lastRowOfRawData);
		
		//Erste Spalte der Trennschärfenberechnung
		//Verschiedene erreichte Gesamtpunktzahlen	
		// X
		$gesamtpunktzahlen = array();	
		for ($i = 2; $i<=$lastRowOfRawData; $i++) {
			array_push ( $gesamtpunktzahlen , $objWorksheet->getCell( 'A'.$i)->getValue());
		}
		$gesamtpunktzahlenUnique = array_unique($gesamtpunktzahlen);
		asort($gesamtpunktzahlenUnique);
		
		$counter = $lastRowOfRawData+8;
		foreach ($gesamtpunktzahlenUnique as $points) {
			$cell = $objWorksheet->getCell ( 'A' . $counter );
			$cell->setValue ($points);
			$counter++;
		}
		 
		//Zweite Spalte der Trennschärfeberechnung
		//Anzahl der Pbn, die eine bestimmte Punktzahl erreicht haben
		// f
		$anzahlProKriteriumwert = array();
		foreach ($gesamtpunktzahlenUnique as $kriterienwert) {
			$count = 0;
			foreach ($gesamtpunktzahlen as $punktzahlEinesPbn) {
				if($kriterienwert == $punktzahlEinesPbn) {
					$count++;
				}
			}
			array_push( $anzahlProKriteriumwert, $count);
		}
		
		$counter = $lastRowOfRawData+8;
		foreach ($anzahlProKriteriumwert as $anzahl) {
			$cell = $objWorksheet->getCell ( 'B' . $counter );
			$cell->setValue ($anzahl);
			$counter++;
		}
		
		//Dritte Spalte der Trennschärfeberechnung
		//Produkt der Anzahl Probanden mit dem Kriterienwert
		//fX
		$counter = $lastRowOfRawData+8;
		foreach (array_combine($gesamtpunktzahlenUnique, $anzahlProKriteriumwert) as $gesamtzahl => $anzahl) {
			$cell = $objWorksheet->getCell ( 'C' . $counter );
			$cell->setValue ($gesamtzahl * $anzahl);
			$counter++;
		}
		
		//Vierte Spalte der Trennschärfeberechnung
		//Produkt der Anzahl Probanden mit dem Quadrate des Kriterienwertes
		//f(X^2)
		$counter = $lastRowOfRawData+8;
		foreach (array_combine($gesamtpunktzahlenUnique, $anzahlProKriteriumwert) as $gesamtzahl => $anzahl) {
			$cell = $objWorksheet->getCell ( 'D' . $counter );
			$cell->setValue (($gesamtzahl * $gesamtzahl) * $anzahl);
			$counter++;
		}	
		
		//Erste Spalte der Aufgabenspezifischen Trennschärfeberechnung
		//Alle drei Spalten pro Aufgabe wiederholen -> Startspalte Column + 7 = neue Startspalte
		//Anzahl Pbn, die die Aufgabe korrekt gelöst haben (>=50% der Punkte erhalten haben)
		
		/* Ablauf:
		 * 1. Spalte für eine Aufgabe einlesen als Array
		 * Spalten von H bis maxDataColumn durchlaufen, pro Spalte Array mit Punktwerten bilden
		 * 	Array durchlaufen und mit 0.5 * $maximumPoints vergleichen, 
		 * 		wenn größer -> Richtitg-counter++, Gesehen-counter++
		 * 		wenn kleiner -> Falsch-counter++, Gesehen-counter++
		 * 		wenn Arraywert null oder '' -> NichtGesehenCounter++
		 * 
		 * A.x einbeziehen, Erreichte Punktzahl, damit in korrekte Spalte einsortiert
		 */
		$lastDataColumnCopy = $lastDataColumn;
		$lastDataColumnCopy++;
		
		$writeColumn = 'E'; //für das Eintragen in die Trennschärfentabelle zum Iterieren benötigt
		$writeRow = $lastRowOfRawData+8;
		for($column = 'H'; $column != ($lastDataColumnCopy); $column ++) {
			//error_log('Betrete Aufgabenspalte: ' . $column);
					
			foreach ($gesamtpunktzahlenUnique as $gesamtpunkte) {
				//error_log(' Betrachte Gesamtpunktzahl: '. $gesamtpunkte);
				
				$itemCount = 0;
				
				$countCorrect = 0;
				$countWrong = 0;
				$countSeen = 0;
				$countNotSeen = 0;

				for ($row = 2; $row <= $lastRowOfRawData; $row++) {
					$gesamtpunkteOriginal = $objWorksheet->getCell( 'A' . $row)->getValue();
					
					//error_log($gesamtpunkte . ' vergleichen mit '. $gesamtpunkteOriginal);
					
					if ($gesamtpunkteOriginal == $gesamtpunkte) {
						//error_log('Habe hier was gefunden');
						
						$itemCount++;
							
						$value = $objWorksheet->getCell( $column . $row)->getValue();
						//$value = $objWorksheet->getCell( 'I' . $row)->getValue();
							
	
						if ($value === NULL || $value === '') {
							$countNotSeen++;
						} elseif ($value >= (0.5 * $maximumPoints[($itemCount-1)])){
							$countCorrect++;
							$countSeen++;
						} elseif ($value < (0.5 * $maximumPoints[($itemCount-1)])){
							$countWrong++;
							$countSeen++;
						}
						$itemCount++;
					}	
				}
				//error_log(' Für Punktzahl ' . $gesamtpunkte . ' in Spalte ' . $column . ': ' . $countCorrect.' '.$countWrong.' - '.$countSeen.' '.$countNotSeen);

				/*
				 * Nun die counter pro Aufgabe in die Spalten der Trennschärfentabelle eintragen
				* Abstände
				* Erstes f.r für Spalte H ist in E:$lastRowOfRawData+8
				* Dann f.rX   in F
				* Dann f.b    in G
				* Dann f.bX   in H
				* Dann f.bX^2 in I
				* Dann ff,fa  in J
				* Dann f.u    in K
				*
				* Dann von VORN -> Columniterator + 7 und mit nächster Aufgabe fortfahren
				*/
				
				$writeColumnInner = $writeColumn;
				$x = $objWorksheet->getCellByColumnAndRow('A', $writeRow)->getValue();
				
				//f.r
				$cell = $objWorksheet->getCell ( $writeColumnInner . $writeRow );
				$cell->setValue ($countCorrect);
				
				//f.rX
				$writeColumnInner++;
				$cell = $objWorksheet->getCell ( $writeColumnInner . $writeRow );
				$cell->setValue ($countCorrect * $x);
				
				//f.b
				$writeColumnInner++;
				$cell = $objWorksheet->getCell ( $writeColumnInner . $writeRow );
				$cell->setValue ($countSeen);
				
				//f.bX
				$writeColumnInner++;
				$cell = $objWorksheet->getCell ( $writeColumnInner . $writeRow );
				$cell->setValue ($countSeen * $x);
				
				//f.bX^2
				$writeColumnInner++;
				$cell = $objWorksheet->getCell ( $writeColumnInner . $writeRow );
				$cell->setValue ($countSeen * ($x * $x));
				
				//f.f, f.a
				$writeColumnInner++;
				$cell = $objWorksheet->getCell ( $writeColumnInner . $writeRow );
				$cell->setValue ($countWrong);

				//f.u
				$writeColumnInner++;
				$cell = $objWorksheet->getCell ( $writeColumnInner . $writeRow );
				$cell->setValue ($countNotSeen);
				
				//error_log('Bin bei Punktzahl: ' . $gesamtpunkte . ' und Spalte ' . $writeColumnInner);
					

				//error_log('Spalte zurückgesetzt auf: ' . $writeColumn);
				$writeRow++;
				
			}
			
			$writeColumn++;
			$writeColumn++;
			$writeColumn++;
			$writeColumn++;
			$writeColumn++;
			$writeColumn++;
			$writeColumn++;
			$writeRow = $lastRowOfRawData+8;
		}
		
		//error_log('Spaltenende für TS ist vor ' . $writeColumn);
		$stopColumn = $writeColumn;
		
		//Summenberechnungen für die Trennschärfe
		$writeRow = $lastRowOfRawData + 8 + count($gesamtpunktzahlenUnique);
		$objWorksheet->setCellValue ( 'A' . $writeRow, 'Summen -> ' );
		
		for($column = 'B'; $column != $stopColumn; $column ++) {
			$cell = $objWorksheet->getCell ( $column . $writeRow );
			
			$value = 0;
			for ($row = ($lastRowOfRawData + 8); $row < $writeRow ; $row++) {
				$value += $objWorksheet->getCell( $column . $row)->getValue();
			}
			
			$cell->setValue ($value);
			//$cell->setValue ('=SUM(' . $column . ($lastRowOfRawData + 8) . ':' . $column . ($writeRow -1) . ')' );
		}
		
		//Eigentliche Formel für die Trennschärfe implementieren pro Aufgabe
		$writeRow++;
		$objWorksheet->setCellValue ( 'A' . $writeRow, 'Trennschärfe -> ' );
		
		//Allgemeine Daten
		$f = $objWorksheet->getCell( 'B' . ($writeRow-1))->getValue();
		$fX = $objWorksheet->getCell( 'C' . ($writeRow-1))->getValue();
		$fX2 = $objWorksheet->getCell( 'D' . ($writeRow-1))->getValue();

		/*
		error_log($f);
		error_log($fX);
		error_log($fX2);
		*/
		
		$resultcolumn;
		for($column = 'E'; $column != $stopColumn; $column ++) {
			//error_log('Eingang: ' . $column);
			$resultcolumn = $column;
			
			//Aufgabenspezifisch
			$fr = $objWorksheet->getCell( $column . ($writeRow-1))->getValue();
			$column++;
			$frX = $objWorksheet->getCell($column . ($writeRow-1))->getValue();
			$column++;
			$fb = $objWorksheet->getCell($column . ($writeRow-1))->getValue();
			$column++;
			$fbX = $objWorksheet->getCell($column . ($writeRow-1))->getValue();
			$column++;
			$fbX2 = $objWorksheet->getCell($column . ($writeRow-1))->getValue();
			$column++;
			$ffa = $objWorksheet->getCell($column . ($writeRow-1))->getValue();
			$column++;
			$fu = $objWorksheet->getCell($column . ($writeRow-1))->getValue();
			
			//error_log('Ausgang: ' . $column);
				
			/*
			error_log($fr);
			error_log($frX);
			error_log($fb);
			error_log($fbX);
			error_log($fbX2);
			error_log($ffa);
			error_log($fu);
			*/

			$cell = $objWorksheet->getCell ( $resultcolumn . $writeRow );
			//$cell->setValue ('=SUM(' . $column . ($lastRowOfRawData + 8) . ':' . $column . ($writeRow -1) . ')' );
			
			/*
			 * OHNE Berücksichtung nicht bearbeiteter Aufgaben
			 */
			/*
			$produkt1 = ($frX / $fr) - ($fX / $f);				
			$produkt2 = sqrt($fr/($f-$fr));
			$produkt3 = $f/(sqrt($f*$fX2-($fX*$fX)));
			$cell->setValue($produkt1 * $produkt2 * $produkt3);
			*/
			
			/*
			 * MIT Berücksichtigung nicht bearbeiteter Aufgaben
			 */
			$produkt1 = ($frX / $fr) - ($fbX / $fb);
			$produkt2 = sqrt($fr/($fb-$fr));
			$produkt3 = sqrt($fb * ($fb-1)) / (sqrt($f*$fbX2-($fX*$fbX)));
			
			error_log($produkt1 . ' ' . $produkt2 . ' ' . $produkt3);
			$cell->setValue($produkt1 * $produkt2 * $produkt3);
			
		}
		
		
		
		
		
		
		
		
		
		//Breite der Spalten automatisch anpassen
		foreach ( range ( 'A', $objWorksheet->getHighestColumn () ) as $columnID ) {
			$objWorksheet->getColumnDimension ( $columnID )->setAutoSize ( true );
		}
		
		return $objWorksheet;
	}
}
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
		
		$objPHPExcel = new PHPExcel ();
		
		// Create the first sheet with general data about the test
		$objWorksheet = $objPHPExcel->getActiveSheet ();
		$this->createFrameSheet1 ( $objWorksheet );
		
		$this->fillInQuestionDataSheet1 ( $objWorksheet );
		
		$this->calculateSummarySheet1 ( $objWorksheet );
		
		// Create the second sheet, with questionspecific data
		//Durch neue Trennschärfeformel überflüssig
		//$this->createProductArray ( $objPHPExcel);

		//Durch neue Trennschärfeformel überflüssig
		//$this->createTrueFalseArrayAndProductArray($objPHPExcel);
		
		$this->calculateDiscrimationIndex( $objPHPExcel);
		
		// Save XSLX file
		ilUtil::makeDirParents ( dirname ( $filename->getPathname ( 'xlsx', 'statistics' ) ) );
		$objWriter = PHPExcel_IOFactory::createWriter ( $objPHPExcel, 'Excel2007' );
		$objWriter->setIncludeCharts ( TRUE );
		$objWriter->save ( str_replace ( __FILE__, $filename->getPathname ( 'xlsx', 'statistics' ), __FILE__ ) );
	}
	
	/**
	 * Takes a reference to the first active sheet and modifies it
	 *
	 * @param	PHPExcel_Worksheet &$objWorksheet
	 *        	
	 */
	public function createFrameSheet1(&$objWorksheet) {

		$objWorksheet->setTitle ( 'Auswertung' );
		/*
		 * Grunddaten des Tests
		 * Titel, Datum, Fragenanzahl, TN-Anzahl
		 * 
		 * B1-C4
		 */
		$objWorksheet->setCellValue ( 'B1' , 'Titel der Klausur');
		$objWorksheet->setCellValue ( 'C1' , $this->getTest()->getTitle() );
		
		$exportDate = date("Y-m-d H:i:s");
		$objWorksheet->setCellValue ( 'B2' , 'Exportdatum' );
		$objWorksheet->setCellValue ( 'C2' , $exportDate );
		
		$objWorksheet->setCellValue ( 'B3' , 'Anz. Fragen (Pro TN/Insg.)' );
		$objWorksheet->setCellValue ( 'C3' , 'xxx' );
		
		$objWorksheet->setCellValue ( 'B4' , $this->getTest ()->lng->txt ( "tst_stat_result_total_participants" ) );
		$objWorksheet->setCellValue ( 'C4' , 'xxx' );
			
		/*
		 * Grunddaten der Fragen
		 * TN-Nummer, Max. Punkte, Punkte, Mittelwert, Varianz, Standardabweichung
		 * 
		 * A6-F6
		 */
		
		$headerRow = array ();
		array_push ( $headerRow, 'Nr.' );
		array_push ( $headerRow, 'Max. Punkte' );
		array_push ( $headerRow, 'Punkte' );
		array_push ( $headerRow, 'Mittelwert' );
		array_push ( $headerRow, 'Varianz' );
		array_push ( $headerRow, 'Std.Abw.' );
		$objWorksheet->fromArray ( $headerRow, null, 'A6', true );

		$styleArray = array(
				'borders' => array(
						'allborders' => array(
								'style' => PHPExcel_Style_Border::BORDER_THIN
						)
				),
        		'fill' => array(
            		'type' => PHPExcel_Style_Fill::FILL_SOLID,
            			'color' => array('rgb' => '87cefa')
        		)
		);
		$objWorksheet->getStyle('B1:C4')->applyFromArray($styleArray);
		unset($styleArray);
		
		//Breite der Spalten automatisch anpassen
		foreach ( range ( 'A', $objWorksheet->getHighestColumn () ) as $columnID ) {
			$objWorksheet->getColumnDimension ( $columnID )->setAutoSize ( true );
		}
		
		$objWorksheet->setCellValue ( 'G5' , 'Aufgaben' );
		$objWorksheet->freezePane('G7');
		
		/*
		 * ENDE der Grunddaten: A1 bis F6
		 * Aufgabentitel in Zeile (G)6 ff.
		 * Aufgabenwerte ab G7 ff.
		 */
	}
	
	/**
	 * Takes a reference to the first active sheet and fills in questiondata
	 *
	 * @param	PHPExcel_Worksheet &$objWorksheet
	 *        	
	 */
	public function fillInQuestionDataSheet1(&$objWorksheet) {
		$data = & $this->getTest ()->getCompleteEvaluationData ( TRUE, $filterby, $filtertext );
		
		// Anzahl der Teilnehmer
		$numberOfParticipants = $data->getStatistics ()->getStatistics ()->count ();
		$objWorksheet->setCellValue ( 'C4' , $numberOfParticipants );
		
		/*
		 * Aufgaben nach ID geordnet auflisten ab G6
		 * 
		 * Sonderfall: TN haben nicht nur unterschiedliche Aufgabenreihenfolge, sondern auch (t.w.) andere Aufgaben
		 * Sonderfall vom Sonderfall: TN können unterschiedliche maximale Punktzahlen haben
		 */
		$allQuestions = array();
		foreach ( $data->getParticipants () as $active_id => $userdata ) {
			
			// Nur der bewertete Durchlauf soll genutzt werden
			$pass = 0;
			if ($this->getTest ()->getPassScoring () == SCORE_BEST_PASS) {
				$pass = $data->getParticipant ( $active_id )->getBestPass ();
			} else { //der letzte Durchlauf
				$pass = $data->getParticipant ( $active_id )->getLastPass ();
			}
			
			if (is_object ( $data->getParticipant ( $active_id ) ) && is_array ( $data->getParticipant ( $active_id )->getQuestions ( $pass ) )) {
				$participantsQuestions = $data->getParticipant ( $active_id )->getQuestions ( $pass );
				$questionAssoziation = array();
				
				foreach ( $participantsQuestions as $question ) {
					
					$titelAndID = preg_replace ( "/<.*?>/", "", $data->getQuestionTitle ( $question ["id"] ) . " (ID=" . $question ["id"] . ")" ); 
					$id = $question ["id"] ;
						
					$questionAssoziation[$id] = $titelAndID;

				}
				$allQuestions = $allQuestions + $questionAssoziation;
			}
		}
		ksort($allQuestions);
		/*
		 * $allQuestions ist ein assoziatives Array nach dem Muster:
		 * (integer)question_id -> (string)Fragetitel+ID
		 * Enthalten sind duplikatreduziert alle Fragen des Tests in aufsteigender ID-Reihenfolge
		 */
		$objWorksheet->fromArray ( $allQuestions, null, 'G6', true );
		
		//Anzahl der Aufgaben in den Kopfbereich eintragen
		$objWorksheet->setCellValue ( 'C3' , count($questionAssoziation) . '/' . count($allQuestions) );

		
		/*
		 * Aufgabenspalten mit Rohwerten besetzen
		 *
		 */
		$rowCount = 7; //Erste freie Zeile nach den verschiedenen Kopfzeilen
		$maxColumn = $objWorksheet->getHighestColumn(); //Startspalte ist immer G
		$maxColumn++;
		$participantNumber = 1;
		foreach ( $data->getParticipants () as $active_id => $userdata ) {

			//Teilnehmernummer
			$objWorksheet->setCellValue ( 'A'.$rowCount , $participantNumber );
			$participantNumber++;
			
			/*
			 * Max. erreichbare Punkte
			 * Eigene Spalte, da bei Zufallstests nicht zwangsläufig bei jedem identisch
			 */
			$maxPoints = $data->getParticipant ( $active_id )->getMaxpoints ();
			$objWorksheet->setCellValue ( 'B'.$rowCount , $maxPoints );
			
			//Erreichte Punkte
			$reachedPoints = $data->getParticipant ( $active_id )->getReached ();
			$objWorksheet->setCellValue ( 'C'.$rowCount , $reachedPoints );

			//Mittelwert = Erreichte Punkte / Anzahl Aufgaben
			$objWorksheet->setCellValue ( 'D'.$rowCount , '=C'.$rowCount.'/' . count($questionAssoziation));

			//Varianz = Mittelwert - Mittelwert * Mittelwert
			$objWorksheet->setCellValue ( 'E'.$rowCount , '=VARP(G'.$rowCount.':'. $objWorksheet->getHighestColumn() . $rowCount . ')' );
			
			//Standardabweichung = Wurzel(Varianz)
			$objWorksheet->setCellValue ( 'F'.$rowCount , '=sqrt(E'.$rowCount.')');
				
			// Nur der bewertete Durchlauf soll genutzt werden
			$pass = 0;
			if ($this->getTest ()->getPassScoring () == SCORE_BEST_PASS) {
				$pass = $data->getParticipant ( $active_id )->getBestPass ();
			} else { //der letzte Durchlauf
				$pass = $data->getParticipant ( $active_id )->getLastPass ();
			}

			$atLeastOneAnsweredQueston = false;
			if (is_object ( $data->getParticipant ( $active_id ) ) && is_array ( $data->getParticipant ( $active_id )->getQuestions ( $pass ) )) {
				$participantsQuestions = $data->getParticipant ( $active_id )->getQuestions ( $pass );
		
				foreach ( $participantsQuestions as $question ) {
					
					for($column = 'G'; $column != ($maxColumn); $column ++) {
						$question_data = $data->getParticipant ( $active_id )->getPass ( $pass )->getAnsweredQuestionByQuestionId ( $question ["id"] );
											
						$titleFromSheet = $objWorksheet->getCell( $column . '6')->getValue();
						$titelFromObject = preg_replace ( "/<.*?>/", "", $data->getQuestionTitle ( $question ["id"] ) . " (ID=" . $question ["id"] . ")" );
						
						//error_log($titleFromSheet.':'.$titelFromObject);
						//$boolean = $titleFromSheet === $titelFromObject;
						//error_log($boolean);
						
						
						
						if ($titleFromSheet === $titelFromObject) {
							$cell = $objWorksheet->getCell( $column . $rowCount);
							$cell->setValue ( $question_data ["reached"] );
							
							//Sonderfall: Nutzer hat keine einzige Frage beantwortet Teil 1/2
							if ( $question_data ["reached"] != null) {
								$atLeastOneAnsweredQueston = true;
							}
							
						}
					}
				}
			}
			
			//Sonderfall: Nutzer hat keine einzige Frage beantwortet Teil 2/2
			if (!$atLeastOneAnsweredQueston) {
				$objWorksheet->setCellValue ( 'C'.$rowCount , 'Test unbearb. abgebr.' );
				$objWorksheet->setCellValue ( 'D'.$rowCount , null );
				$objWorksheet->setCellValue ( 'E'.$rowCount , null );
				$objWorksheet->setCellValue ( 'F'.$rowCount , null );
			}	
			$rowCount++;	
		}		
	}	

	public function calculateSummarySheet1(&$objWorksheet) {
		$lastRowOfRawData = $objWorksheet->getHighestRow ();
		$lastColumnRawData = $objWorksheet->getHighestColumn();

		$maxColumn = $lastColumnRawData;
		$maxColumn ++;
		
		//Auswertungstabellenbeschriftung
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+3) , 'Spaltensumme');
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+4) , 'Summenprodukt');
		
		
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+6) , 'Erreichbare Punktzahl' );
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+7) , 'Richtige Antworten' );
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+8) , 'Falsche Antworten' );
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+9) , 'Unbeantwortet' );
		
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+11) , 'Mittelwert');
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+12) , 'Varianz' );
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+13) , 'Standardabweichung' );
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+14) , 'Schwierigkeitsindex' );
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+15) , 'Trennschärfe' );
		$objWorksheet->setCellValue ( 'B'.($lastRowOfRawData+16) , 'Trennschärfe (Itemkorrigiert)' );
		
		
		for($column = 'C'; $column != ($maxColumn); $column ++) {
			//Spaltensumme
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+3) , '=SUM(' . $column.'7:'.$column.$lastRowOfRawData.')' );
			
			//Summenprodukt
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+4) , '=SUMPRODUCT(' . $column.'7:'.$column.$lastRowOfRawData.','. $column.'7:'.$column.$lastRowOfRawData . ')' );
		}	
		
		//Anzahl Teilnehmer
		$anzahlTeilnehmer = $objWorksheet->getCell( 'A' . $lastRowOfRawData)->getValue();
		
		//Erreichbare Punktzahl pro Aufgabe
		$data = & $this->getTest ()->getCompleteEvaluationData ( TRUE, $filterby, $filtertext );
		foreach ( $data->getParticipants () as $active_id => $userdata ) {
			
			// Nur der bewertete Durchlauf soll genutzt werden
			$pass = 0;
			if ($this->getTest ()->getPassScoring () == SCORE_BEST_PASS) {
				$pass = $data->getParticipant ( $active_id )->getBestPass ();
			} else { //der letzte Durchlauf
				$pass = $data->getParticipant ( $active_id )->getLastPass ();
			}
		
			if (is_object ( $data->getParticipant ( $active_id ) ) && is_array ( $data->getParticipant ( $active_id )->getQuestions ( $pass ) )) {
				$participantsQuestions = $data->getParticipant ( $active_id )->getQuestions ( $pass );
		
				foreach ( $participantsQuestions as $question ) {
						
					for($column = 'G'; $column != ($maxColumn); $column ++) {
						$question_data = $data->getParticipant ( $active_id )->getPass ( $pass )->getAnsweredQuestionByQuestionId ( $question ["id"] );
							
						$titleFromSheet = $objWorksheet->getCell( $column . '6')->getValue();
						$titelFromObject = preg_replace ( "/<.*?>/", "", $data->getQuestionTitle ( $question ["id"] ) . " (ID=" . $question ["id"] . ")" );
		
						if ($titleFromSheet === $titelFromObject) {
							$cell = $objWorksheet->getCell( $column . ($lastRowOfRawData+6));
							$cell->setValue ( $question_data ["points"] );	
						}
					}
				}
			}
		}// Ende Erreichbare Punktzahl pro Aufgabe

		

		//Mittelwert der erreichten Punkte aller Teilnehmer
		$objWorksheet->setCellValue ( 'C'.($lastRowOfRawData+11) ,  '=AVERAGE(C7:C'. $lastRowOfRawData.')');
		
		//Varianz der erreichten Punkte aller Teilnehmer
		$objWorksheet->setCellValue ( 'C'.($lastRowOfRawData+12) ,  '=VARP(C7:C'. $lastRowOfRawData.')');
		
		//Standardabweichung der erreichten Punkte aller Teilnehmer
		$objWorksheet->setCellValue ( 'C'.($lastRowOfRawData+13) ,  '=sqrt(C'.($lastRowOfRawData+12).')');
	
		for($column = 'G'; $column != ($maxColumn); $column ++) {
			
			//richtige Antworten -> über (>=) 50% der erreichbaren Punktzahl =ZÄHLENWENN(G7:G116;">="&G122*"0,5")
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+7) ,  '=COUNTIF('. $column.'7:' . $column. $lastRowOfRawData  . ',">="&'. $column.($lastRowOfRawData+6) . '*"0,5")');

			//falsche Antworten -> unter (<) 50% der erreichbaren Punktzahl
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+8) ,  '=COUNTIF('. $column.'7:' . $column. $lastRowOfRawData  . ',"<"&'. $column.($lastRowOfRawData+6) . '*"0,5")');

			//unbeantwortete Fragen (wenn Zelle = NIL)
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+9) , '=COUNTIF('. $column.'7:' . $column. $lastRowOfRawData  . ',"")');

			//Mittelwert
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+11) , '=' . $column.($lastRowOfRawData+3) . '/(' . ($anzahlTeilnehmer .'-'. $column.($lastRowOfRawData+9)).')' );

			//Varianz
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+12) , '=VARP(' . $column.'7:'.$column.$lastRowOfRawData.')' );

			//Standardabweichung
			$objWorksheet->setCellValue ( $column.($lastRowOfRawData+13) , '=sqrt('.$column.($lastRowOfRawData+12).')' );

			//Schwierigkeitsindex
			$maxPoints = $objWorksheet->getCell ( $column . ($lastRowOfRawData+6) )->getValue ();
			$countNotShown = $objWorksheet->getCell ( $column.($lastRowOfRawData+9) )->getValue ();
				
			if (($anzahlTeilnehmer - $countNotShown) == 0) {
				$cell->setValue ( 'n.D.' );
			
			} else {
				$objWorksheet->setCellValue ( $column.($lastRowOfRawData+14) , '=100*(SUM(' . $column . '7:' . $column . ($lastRowOfRawData) . ')/' . (($anzahlTeilnehmer - $countNotShown) * $maxPoints) . ')' );
			}
		}
	}
	
	/**
	 * creates another worksheet
	 *
	 * @return PHPExcel_Worksheet objWorksheet
	 */
	public function createProductArray(&$objPHPExcel) {
		
		$produkteAufgaben = new PHPExcel_Worksheet ( $objPHPExcel );
		$produkteAufgaben->setTitle ( 'Aufgaben Produkte' );
		
		$firstSheet = $objPHPExcel->getSheet(0);
		$anzahlTeilnehmer = $firstSheet->getCell( 'C4' )->getValue();
		$endzeile = $anzahlTeilnehmer + 6;
		$lastColumnRawData = $firstSheet->getHighestColumn();
		
		$aufgabenwerte = $firstSheet->rangeToArray('G7:' . $lastColumnRawData . $endzeile,0,true,false);
		//$produkteAufgaben->fromArray($aufgabenwerte, NULL, 'A1', true);
		
		$transponierteAufgabenwerte = PHPExcel_Calculation_LookupRef::TRANSPOSE($aufgabenwerte);
		//$produkteAufgaben->fromArray($transponierteAufgabenwerte, NULL, 'A150', true);
		
		
		$endmatrix = PHPExcel_Calculation_MathTrig::MMULT($transponierteAufgabenwerte, $aufgabenwerte);	
		$produkteAufgaben->fromArray($endmatrix, NULL, 'A1', true);
		
		$lastColumnMMULTData = $produkteAufgaben->getHighestColumn();
		$lastRowMMULTData = $produkteAufgaben->getHighestRow();
		
		$maxColumn = $lastColumnMMULTData;
		$maxColumn++;
		
		$writeRow = $lastRowMMULTData + 2;

		for($column = 'A'; $column != ($maxColumn); $column ++) {
		
			$cell = $produkteAufgaben->getCell ( $column . $writeRow );
			$cell->setValue ( '=SUM(' . $column . '1:' . $column . $lastRowMMULTData . ')' );				
		}
		
		$objPHPExcel->addSheet ( $produkteAufgaben );
	}
	
	public function createTrueFalseArrayAndProductArray($objPHPExcel){

		$firstSheet = $objPHPExcel->getSheet(0);
		
		$richtigFalschWerte = new PHPExcel_Worksheet ( $objPHPExcel );
		$richtigFalschWerte->setTitle ( 'RichtigFalschWerte' );
		$objPHPExcel->addSheet ( $richtigFalschWerte );
		
		$lastColumnRawData = $firstSheet->getHighestColumn();
		$maxColumn = $lastColumnRawData;
		$maxColumn ++;
		
		$anzahlTeilnehmer = $firstSheet->getCell( 'C4' )->getValue();		
	
		//Erreichbare Punktzahl = $anzahlTeilnehmer + 12
		
		$endrow = $anzahlTeilnehmer + 6;
		$richtigFalschColumn = 'A';
		for($column = 'G'; $column != ($maxColumn); $column ++) {

			$erreichbarePunkte = $firstSheet->getCell( $column.($anzahlTeilnehmer + 12) )->getValue();
						
			for ($row = 7; $row <= $endrow; $row++){
		

				if($firstSheet->getCell( $column . $row )->getCalculatedValue() >= (0.5 * (float)$erreichbarePunkte)){
					$richtigFalschWerte->setCellValue( $richtigFalschColumn . ($row-6), 1);

				} else {
					$richtigFalschWerte->setCellValue( $richtigFalschColumn . ($row-6), 0);
				}
			}
			$richtigFalschColumn++;
		}
		
		$lastColumnTrueFalseData = $richtigFalschWerte->getHighestColumn();
		$lastRowTrueFalseData = $richtigFalschWerte->getHighestRow();
		
		$summaryColumn = $lastColumnTrueFalseData;
		$summaryColumn ++;
		for ($row = 1; $row <= $lastRowTrueFalseData; $row++) {
			$richtigFalschWerte->setCellValue ( $summaryColumn.$row ,  '=SUM( A' .$row . ':' . $lastColumnTrueFalseData . $row . ')');
				
		}
		
		
		$richtigFalschWerte->setCellValue( $summaryColumn . ($lastRowTrueFalseData+1), '=AVERAGE(' . $summaryColumn . '1:' . $summaryColumn .$lastRowTrueFalseData . ')');
		$richtigFalschWerte->setCellValue( $summaryColumn . ($lastRowTrueFalseData+2), '=VARP(' . $summaryColumn . '1:' . $summaryColumn .$lastRowTrueFalseData . ')');
		$richtigFalschWerte->setCellValue( $summaryColumn . ($lastRowTrueFalseData+3), '=SQRT(' . $summaryColumn . ($lastRowTrueFalseData+2) . ')');
		
		$richtigFalschProdukte = new PHPExcel_Worksheet ( $objPHPExcel );
		$richtigFalschProdukte->setTitle ( 'RichtigFalsch Produkte' );
		$objPHPExcel->addSheet ( $richtigFalschProdukte );
		
		$aufgabenwerte = $richtigFalschWerte->rangeToArray('A1:' . $lastColumnTrueFalseData . $lastRowTrueFalseData,0,true,false);
		//$produkteAufgaben->fromArray($aufgabenwerte, NULL, 'A1', true);
		
		$transponierteAufgabenwerte = PHPExcel_Calculation_LookupRef::TRANSPOSE($aufgabenwerte);		
		
		$endmatrix = PHPExcel_Calculation_MathTrig::MMULT($transponierteAufgabenwerte, $aufgabenwerte);
		$richtigFalschProdukte->fromArray($endmatrix, NULL, 'A1', true);
		
		$lastColumnMMULTData = $richtigFalschProdukte->getHighestColumn();
		$lastRowMMULTData = $richtigFalschProdukte->getHighestRow();
		
		$maxColumn = $lastColumnMMULTData;
		$maxColumn++;
		
		$writeRow = $lastRowMMULTData + 2;
		
		for($column = 'A'; $column != ($maxColumn); $column ++) {
		
			$cell = $richtigFalschProdukte->getCell ( $column . $writeRow );
			$cell->setValue ( '=SUM(' . $column . '1:' . $column . $lastRowMMULTData . ')' );
		}
		
	}
	
	
	public function calculateDiscrimationIndex(&$objPHPExcel){
		$firstSheet = $objPHPExcel->getSheet(0);
		
		/*
		 * Alte Trennschärfe, mit Excelfunktionen, fehleranfällig
		*/
		
		/*
		$truefalse = $objPHPExcel->getSheet(2);
		$tfmaxc = $truefalse->getHighestColumn();
		$tfmaxr = $truefalse->getHighestRow();
		
		
		$produktSummen = $objPHPExcel->getSheet(3);
		
		$lastColumnRawData = $firstSheet->getHighestColumn();
		$maxColumn = $lastColumnRawData;
		$maxColumn ++;
		
		$anzahlTeilnehmer = $firstSheet->getCell( 'C4' )->getValue();
		$endzeile = $anzahlTeilnehmer + 21;
		
		
		$lastColumnProductData = $produktSummen->getHighestColumn();
		$lastRowProductData = $produktSummen->getHighestRow();

		$maxColumn = $lastColumnRawData;
		$maxColumn ++;
		
		$produkteAufgabenColumn = 'A';
		for($column = 'G'; $column != ($maxColumn); $column ++) {

			$summenwert = $produktSummen->getCell( $produkteAufgabenColumn . $lastRowProductData )->getCalculatedValue();
			$produkteAufgabenColumn++;

			//Durchschnitt Gesamt
			$durchschnittGesamt_TrueFalse = $truefalse->getCell( $tfmaxc.($tfmaxr-2) )->getCalculatedValue();
					
			//Std.Abw. Gesamt
			$stdAbwGes = $truefalse->getCell( $tfmaxc.$tfmaxr )->getCalculatedValue();
				
			
			//Trennschärfeformel
			//error_log('=(((' . $summenwert . '/' . $column . ($anzahlTeilnehmer + 13) . ')-' . $durchschnittGesamt_TrueFalse .')/' . $stdAbwGes . ')');
			//$firstSheet->setCellValue ( $column.($endzeile) , '=(((' . $summenwert . '/' . $column . ($anzahlTeilnehmer + 13) . ')-' . $durchschnittGesamt_TrueFalse .')/' . $stdAbwGes . ')');
			//error_log(	'=(sqrt(' .  $column . ($anzahlTeilnehmer + 13) . '/(' . $anzahlTeilnehmer . '-' .$column . ($anzahlTeilnehmer + 15) . '-' . $column . ($anzahlTeilnehmer + 13) . ')))');
			//$firstSheet->setCellValue ( $column.($endzeile+1) ,		'=(sqrt(' .  $column . ($anzahlTeilnehmer + 13) . '/(' . $anzahlTeilnehmer . '-' .$column . ($anzahlTeilnehmer + 15) . '-' . $column . ($anzahlTeilnehmer + 13) . ')))');
			//$firstSheet->setCellValue ( $column.($endzeile+2) , '=(' . $column.($endzeile) . '*' . $column.($endzeile+1) .' )');
			//$firstSheet->setCellValue ( $column.($endzeile+3) , '=(((' . $summenwert . '/' . $column . ($anzahlTeilnehmer + 13) . ')-' . $durchschnittGesamt_TrueFalse .')/' . $stdAbwGes . ')*' . '(sqrt(' .  $column . ($anzahlTeilnehmer + 13) . '/(' . $anzahlTeilnehmer . '-' .$column . ($anzahlTeilnehmer + 15) . '-' . $column . ($anzahlTeilnehmer + 13) . ')))');

			$firstSheet->setCellValue ( $column.($endzeile) , '=(((' . $summenwert . '/' . $column . ($anzahlTeilnehmer + 13) . ')-' . $durchschnittGesamt_TrueFalse .')/' . $stdAbwGes . ')*' . '(sqrt(' .  $column . ($anzahlTeilnehmer + 13) . '/(' . $anzahlTeilnehmer . '-' .$column . ($anzahlTeilnehmer + 15) . '-' . $column . ($anzahlTeilnehmer + 13) . ')))');
		}
		*/
		
		
		
		
		
		
		/*
		 *
		 * Pearson Correlation
		 *
		 */
		$lastColumnRawData = $firstSheet->getHighestColumn();
		$maxColumn = $lastColumnRawData;
		$maxColumn++;
			
		$anzahlTeilnehmer = $firstSheet->getCell( 'C4' )->getValue();
		$endzeile = $anzahlTeilnehmer + 21;
		
		//First Row Raw Data = 7
		$lastRowRawData = $anzahlTeilnehmer + 6;
		
		for($column = 'G'; $column != ($maxColumn); $column ++) {
			$array_item = array();
			$array_test = array();
			$array_test_unkorrigiert = array();
			for($row = 7; $row <= $lastRowRawData; $row++) {
				$value = $firstSheet->getCell( $column.$row )->getValue();

				if (!($value === null) && !($value === '') && is_numeric($firstSheet->getCell( 'C'.$row )->getValue())) {
					error_log($value . ' inner ' . $column.$row);
					$array_item[] = $value;
					
					$gesamttestwert = $firstSheet->getCell( 'C'.$row )->getValue();
					$itemkorrigierterGesamttestwert = $gesamttestwert - $value;

					$array_test[] = $itemkorrigierterGesamttestwert;
					$array_test_unkorrigiert[] = $gesamttestwert;
						
				}
			}

			//############## Pearson vanilla:
				
			$length= count($array_item);
			$mean1=array_sum($array_item) / $length;
			$mean2=array_sum($array_test_unkorrigiert) / $length;
			
			$a=0;
			$b=0;
			$axb=0;
			$a2=0;
			$b2=0;
			
			for($i=0;$i<$length;$i++)
			{
			$a=$array_item[$i]-$mean1;
			$b=$array_test_unkorrigiert[$i]-$mean2;
			$axb=$axb+($a*$b);
			$a2=$a2+ pow($a,2);
			$b2=$b2+ pow($b,2);
			}
			
			if ($a2 == 0 || $b2 == 0){
			$corr = "NaN";
			} else {
				$corr= $axb / sqrt($a2*$b2);
						
				}
					
				//error_log($axb .' / ' . 'sqrt(' . $a2 . '*' . $b2 . ')');
					
			$firstSheet->setCellValue ( $column.($endzeile) , $corr);
				
			//############## Itemkorrigiert:
			
			
			
			
			$length= count($array_item);
			$mean1=array_sum($array_item) / $length;
			$mean2=array_sum($array_test) / $length;
			
			$a=0;
			$b=0;
			$axb=0;
			$a2=0;
			$b2=0;
			
			for($i=0;$i<$length;$i++)
			{
			$a=$array_item[$i]-$mean1;
			$b=$array_test[$i]-$mean2;
			$axb=$axb+($a*$b);
			$a2=$a2+ pow($a,2);
			$b2=$b2+ pow($b,2);
			}
			
			if ($a2 == 0 || $b2 == 0){
				$corr = "NaN";
			} else {
				$corr= $axb / sqrt($a2*$b2);
				
			}
			
			//error_log($axb .' / ' . 'sqrt(' . $a2 . '*' . $b2 . ')');
			
			$firstSheet->setCellValue ( $column.($endzeile+1) , $corr);
			
		}		
	}
}
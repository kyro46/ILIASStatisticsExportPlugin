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
		
		
		// Create the second sheet, with questionspecific data
		//$objPHPExcel->addSheet ( $this->createRawDataDifficulty () );

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
		 * A1-B4
		 */
		$objWorksheet->setCellValue ( 'A1' , 'Titel der Klausur');
		$objWorksheet->setCellValue ( 'B1' , $this->getTest()->getTitle() );
		
		$exportDate = date("Y-m-d H:i:s");
		$objWorksheet->setCellValue ( 'A2' , 'Exportdatum' );
		$objWorksheet->setCellValue ( 'B2' , $exportDate );
		
		$objWorksheet->setCellValue ( 'A3' , $this->getTest ()->lng->txt ( "tst_stat_result_qmax" ) );
		$objWorksheet->setCellValue ( 'B3' , 'xxx' );
		
		$objWorksheet->setCellValue ( 'A4' , $this->getTest ()->lng->txt ( "tst_stat_result_total_participants" ) );
		$objWorksheet->setCellValue ( 'B4' , 'xxx' );
			
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
		$objWorksheet->getStyle('A1:B4')->applyFromArray($styleArray);
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
		$objWorksheet->setCellValue ( 'B4' , $numberOfParticipants );
		
		/*
		 * Aufgaben nach ID geordnet listen ab G6
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
					
					//$allQuestions = array_merge($allQuestions, array_diff(array($titelAndID), $allQuestions));
						
				}
				$allQuestions = array_merge($allQuestions, array_diff($questionAssoziation, $allQuestions));
				
			}
		}
		
		ksort($allQuestions);
		error_log(implode(",", $allQuestions));
		$objWorksheet->fromArray ( $allQuestions, null, 'G6', true );
		
	}	
	
	/**
	 * creates another worksheet
	 *
	 * @return PHPExcel_Worksheet objWorksheet
	 */
	public function createRawDataDifficulty() {

	}
}
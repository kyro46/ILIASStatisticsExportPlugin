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
		include_once "Services/Excel/classes/class.ilExcelUtils.php";
		$data   = $this->getTest()->getCompleteEvaluationData(TRUE);
		$titles = $this->getTest()->getQuestionTitlesAndIndexes();
		$oids   = $this->getTest()->getQuestions();
		asort($oids);
		
	
		$rows = array();
		$datarow = array();
		$col = 1;
		if ($this->getTest()->getAnonymity())
		{
			array_push($datarow, $this->getTest()->lng->txt("counter"));
			$col++;
		}
		else
		{
			array_push($datarow, $this->getTest()->lng->txt("name"));
			$col++;
			array_push($datarow, $this->getTest()->lng->txt("login"));
			$col++;
			array_push($datarow, $this->getTest()->lng->txt("matriculation"));
			$col++;
		}
		$additionalFields = $this->getTest()->getEvaluationAdditionalFields();
		if (count($additionalFields))
		{
			foreach ($additionalFields as $fieldname)
			{
				array_push($datarow, $this->getTest()->lng->txt($fieldname));
				$col++;
			}
		}
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_resultspoints"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("maximum_points"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_resultsmarks"));
		$col++;
		if ($this->getTest()->ects_output)
		{
			array_push($datarow, $this->getTest()->lng->txt("ects_grade"));
			$col++;
		}
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_qworkedthrough"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_qmax"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_pworkedthrough"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_timeofwork"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_atimeofwork"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_firstvisit"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_lastvisit"));
		$col++;

		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_mark_median"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_rank_participant"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_rank_median"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_total_participants"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("tst_stat_result_median"));
		$col++;
		array_push($datarow, $this->getTest()->lng->txt("scored_pass"));
		$col++;

		array_push($datarow, $this->getTest()->lng->txt("pass"));
		$col++;

		$data =& $this->getTest()->getCompleteEvaluationData(TRUE, $filterby, $filtertext);
		$headerrow = $datarow;
		$counter = 1;
		foreach ($data->getParticipants() as $active_id => $userdata) 
		{
			$datarow = $headerrow;
			$remove = FALSE;
			if ($passedonly)
			{
				if ($data->getParticipant($active_id)->getPassed() == FALSE)
				{
					$remove = TRUE;
				}
			}
			if (!$remove)
			{
				$datarow2 = array();
				if ($this->getTest()->getAnonymity())
				{
					array_push($datarow2, $counter);
				}
				else
				{
					array_push($datarow2, $data->getParticipant($active_id)->getName());
					array_push($datarow2, $data->getParticipant($active_id)->getLogin());
					$userfields = ilObjUser::_lookupFields($userdata->getUserID());
					array_push($datarow2, $userfields['matriculation']);
				}
				if (count($additionalFields))
				{
					$userfields = ilObjUser::_lookupFields($userdata->getUserID());
					foreach ($additionalFields as $fieldname)
					{
						if (strcmp($fieldname, "gender") == 0)
						{
							array_push($datarow2, $this->getTest()->lng->txt("gender_" . $userfields[$fieldname]));
						}
						else
						{
							array_push($datarow2, $userfields[$fieldname]);
						}
					}
				}
				array_push($datarow2, $data->getParticipant($active_id)->getReached());
				array_push($datarow2, $data->getParticipant($active_id)->getMaxpoints());
				array_push($datarow2, $data->getParticipant($active_id)->getMark());
				if ($this->getTest()->ects_output)
				{
					array_push($datarow2, $data->getParticipant($active_id)->getECTSMark());
				}
				array_push($datarow2, $data->getParticipant($active_id)->getQuestionsWorkedThrough());
				array_push($datarow2, $data->getParticipant($active_id)->getNumberOfQuestions());
				array_push($datarow2, $data->getParticipant($active_id)->getQuestionsWorkedThroughInPercent() / 100.0);
				$time = $data->getParticipant($active_id)->getTimeOfWork();
				$time_seconds = $time;
				$time_hours    = floor($time_seconds/3600);
				$time_seconds -= $time_hours   * 3600;
				$time_minutes  = floor($time_seconds/60);
				$time_seconds -= $time_minutes * 60;
				array_push($datarow2, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
				$time = $data->getParticipant($active_id)->getQuestionsWorkedThrough() ? $data->getParticipant($active_id)->getTimeOfWork() / $data->getParticipant($active_id)->getQuestionsWorkedThrough() : 0;
				$time_seconds = $time;
				$time_hours    = floor($time_seconds/3600);
				$time_seconds -= $time_hours   * 3600;
				$time_minutes  = floor($time_seconds/60);
				$time_seconds -= $time_minutes * 60;
				array_push($datarow2, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
				
				$fv = $data->getParticipant($active_id)->getFirstVisit();
				$lv = $data->getParticipant($active_id)->getLastVisit();
				foreach(array($fv, $lv) as $ts)
				{
					if($ts)
					{
						$visit = ilFormat::formatDate(date('Y-m-d H:i:s', $ts), "datetime", false, false);
						array_push($datarow2, $visit);
					}
					else
					{
						array_push($datarow2, "");
					}
				}

				$median = $data->getStatistics()->getStatistics()->median();
				$pct = $data->getParticipant($active_id)->getMaxpoints() ? $median / $data->getParticipant($active_id)->getMaxpoints() * 100.0 : 0;
				$mark = $this->getTest()->mark_schema->getMatchingMark($pct);
				$mark_short_name = "";
				if (is_object($mark))
				{
					$mark_short_name = $mark->getShortName();
				}
				array_push($datarow2, $mark_short_name);
				array_push($datarow2, $data->getStatistics()->getStatistics()->rank($data->getParticipant($active_id)->getReached()));
				array_push($datarow2, $data->getStatistics()->getStatistics()->rank_median());
				array_push($datarow2, $data->getStatistics()->getStatistics()->count());
				array_push($datarow2, $median);
				if ($this->getTest()->getPassScoring() == SCORE_BEST_PASS)
				{
					array_push($datarow2, $data->getParticipant($active_id)->getBestPass() + 1);
				}
				else
				{
					array_push($datarow2, $data->getParticipant($active_id)->getLastPass() + 1);
				}
				for ($pass = 0; $pass <= $data->getParticipant($active_id)->getLastPass(); $pass++)
				{
					$finishdate = $this->getTest()->getPassFinishDate($active_id, $pass);
					if ($finishdate > 0)
					{
						if ($pass > 0)
						{
							for ($i = 1; $i < $col-1; $i++) 
							{
								array_push($datarow2, "");
								array_push($datarow, "");
							}
							array_push($datarow, "");
						}
						array_push($datarow2, $pass+1);
						if (is_object($data->getParticipant($active_id)) && is_array($data->getParticipant($active_id)->getQuestions($pass)))
						{
							$sortedquestions = array ();
							$sortedquestions = $data->getParticipant($active_id)->getQuestions($pass);
							asort($sortedquestions);
							
							foreach ($sortedquestions as $question)
							{
								$question_data = $data->getParticipant($active_id)->getPass($pass)->getAnsweredQuestionByQuestionId($question["id"]);
								array_push($datarow2, $question_data["reached"]);
								array_push($datarow, preg_replace("/<.*?>/", "", $data->getQuestionTitle($question["id"])));
							}
						}
						if ($this->getTest()->isRandomTest() || ($counter == 1 && $pass == 0))
						{
							array_push($rows, $datarow);
						}
						$datarow = array();
						array_push($rows, $datarow2);
						$datarow2 = array();
					}
				}
				$counter++;
			}
		}
		$csv = "";
		$separator = ";";
		foreach ($rows as $evalrow)
		{
			$csvrow =& $this->getTest()->processCSVRow($evalrow, TRUE, $separator);
			$csv .= join($csvrow, $separator) . "\n";
		}

		ilUtil::makeDirParents(dirname($filename->getPathname('csv', 'matriculation')));
		file_put_contents($filename->getPathname('csv', 'matriculation'), $csv);
	
	}
}
<?php
/**
 * Download scores report for SCORM
 * @package	scormreport
 * @subpackage download_Scores
 * @author	 Christian Perfect for Newcastle University
 * @license	http://www.apache.org/licenses/LICENSE-2.0.html Apache License, version 2.0
 */

defined('MOODLE_INTERNAL') || die();

class report extends mod_scorm\report {
	/**
	 * displays the full report
	 * @param stdClass $scorm full SCORM object
	 * @param stdClass $cm - full course_module object
	 * @param stdClass $course - full course object
	 * @param string $download - type of download being requested
	 */

	function get_objective_number($element) {
		preg_match('/^cmi.objectives.(\d+)/',$element,$matches);
		return $matches[1];
	}

	function define_headers() {
		// Define table columns
		$this->headers = array(
			get_string('username'),
			get_string('name'),
			get_string('attempt', 'scorm'),
			get_string('started', 'scorm'),
			get_string('last', 'scorm'),
			get_string('totalscore', 'scormreport_download_scores'),
			get_string('percentage', 'scormreport_download_scores')
		);
	
		foreach($this->all_question_ids as $id) {
			$this->headers[] = get_string('objectivename','scormreport_download_scores',$id);
		}

	}

	function output_ODS() {
		global $CFG;
		require_once("$CFG->libdir/odslib.class.php");

		$workbook = new MoodleODSWorkbook("-");
		$workbook->send($this->filename . '.ods');

		$this->output_workbook($workbook);
	}

	function output_Excel() {
		global $CFG;
		require_once("$CFG->libdir/excellib.class.php");

		// Creating a workbook
		$workbook = new MoodleExcelWorkbook("-");
		// Sending HTTP headers
		$workbook->send($this->filename . '.xls');

		$this->output_workbook($workbook);
	}

	function output_workbook($workbook) {
		$sheettitle = get_string('report', 'scorm');
		$worksheet = $workbook->add_worksheet($sheettitle);
		$format = $workbook->add_format();
		$format->set_bold(0);
		$formatbc = $workbook->add_format();
		$formatbc->set_bold(1);
		$formatbc->set_align('center');

		$colnum = 0;
		foreach ($this->headers as $item) {
			$worksheet->write(0, $colnum, $item, $formatbc);
			$colnum++;
		}
		$rownum=1;

		//for each row
		foreach($this->rows as $row) {
			$colnum = 0;
			foreach ($row as $item) {
				$worksheet->write($rownum, $colnum, $item, $format);
				$colnum++;
			}
			$rownum++;
		}

		$workbook->close();
	}

	function output_CSV() {
		global $CFG;
		require_once($CFG->libdir . '/csvlib.class.php');
		$csvexport = new csv_export_writer("tab");
		$csvexport->delimiter = ',';
		$csvexport->set_filename($this->filename, ".csv");
		$csvexport->add_data($this->headers);

		//for each row
		foreach($this->rows as $row) {
			$csvexport->add_data($row);
		}

		$csvexport->download_file();
	}

	function display($scorm, $cm, $course, $download) {
		global $OUTPUT, $PAGE;
		if(!$download) {
			echo $OUTPUT->single_button(new moodle_url($PAGE->url,array('download'=>'ODS')),get_string('downloadods'));
			echo $OUTPUT->single_button(new moodle_url($PAGE->url,array('download'=>'Excel')),get_string('downloadexcel'));
			echo $OUTPUT->single_button(new moodle_url($PAGE->url,array('download'=>'CSV')),get_string('downloadcsv','scormreport_download_scores'));
		} else {
			$coursecontext = $this->coursecontext = context_course::instance($course->id);
			$shortname = format_string($course->shortname, true, array('context' => $coursecontext));
			$this->filename = clean_filename("$shortname ".format_string($scorm->name, true));

			$this->scorm = $scorm;
			$this->download = $download;
			$this->get_data($scorm,$cm,$course);
			switch($download) {
			case 'ODS':
				$this->output_ODS();
				break;
			case 'Excel':
				$this->output_Excel();
				break;
			case 'CSV':
				$this->output_CSV();
				break;
			}
		}
	}

	function get_data($scorm, $cm, $course) {
		global $DB, $OUTPUT;

		$modulecontext = $this->modulecontext = context_module::instance($cm->id);
		$coursecontext = $this->coursecontext = context_course::instance($course->id);

		/// !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
		$this->num_objectives = 20;

		// find out current groups mode
		$currentgroup = groups_get_activity_group($cm, true);

		// select the students

		$nostudents = false;
		if (empty($currentgroup)) {
			// all users who can attempt scoes
			if (!$students = get_users_by_capability($modulecontext, 'mod/scorm:savetrack', 'u.id', '', '', '', '', '', false)) {
				echo $OUTPUT->notification(get_string('nostudentsyet'));
				$nostudents = true;
			} else {
				$allowedlist = array_keys($students);
				unset($students);
			}
		} else {
			// all users who can attempt scoes and who are in the currently selected group
			if (!$groupstudents = get_users_by_capability($modulecontext, 'mod/scorm:savetrack', 'u.id', '', '', '', $currentgroup, '', false)) {
				echo $OUTPUT->notification(get_string('nostudentsingroup'));
				$nostudents = true;
			} else {
				$allowedlist = array_keys($groupstudents);
				unset($groupstudents);
			}
		}

		if ( $nostudents ) {
			echo $OUTPUT->notification(get_string('noactivity', 'scorm'));
			return;
		}

		$params = array();
		list($usql, $params) = $DB->get_in_or_equal($allowedlist, SQL_PARAMS_NAMED);
						// Construct the SQL
		$select = 'SELECT DISTINCT '.$DB->sql_concat('u.id', '\'#\'', 'COALESCE(st.attempt, 0)').' AS uniqueid, ';
		$select .= 'st.scormid AS scormid, st.attempt AS attempt, ' .
				user_picture::fields('u', array('idnumber','username'), 'userid') .
				get_extra_user_fields_sql($coursecontext, 'u', '', array('idnumber')) . ' ';

		// This part is the same for all cases - join users and scorm_scoes_track tables
		$from = 'FROM {user} u ';
		$from .= 'LEFT JOIN {scorm_scoes_track} st ON st.userid = u.id AND st.scormid = '.$scorm->id;

		$where = ' WHERE u.id ' .$usql;

		// Fix some wired sorting
		if (empty($sort)) {
			$sort = ' ORDER BY uniqueid';
		} else {
			$sort = ' ORDER BY '.$sort;
		}

		// Fetch the attempts
		// gets uniqueid, scormid, attempt number, and all user fields. One row for each user-attempt
		$attempts = $DB->get_records_sql($select.$from.$where.$sort, $params);

		$this->all_question_ids = array();

		if (!$attempts) {
			return;
		}

		foreach ($attempts as $scouser) {
			if (!empty($scouser->attempt)) {
				$scouser->timetracks = scorm_get_sco_runtime($scorm->id, false, $scouser->userid, $scouser->attempt);
			} else {
				$scouser->timetracks = '';
			}

			if (!empty($scouser->timetracks->start)) {

				$objective_id_records = $DB->get_records_sql('SELECT element,value FROM {scorm_scoes_track} WHERE attempt = :attempt AND userid = :userid AND scormid = :scormid AND element LIKE \'cmi.objectives.%.id\'',array('attempt' => $scouser->attempt, 'userid' => $scouser->userid, 'scormid' => $scouser->scormid));
				$objective_ids = array();
				foreach($objective_id_records as $objective) {
					$n = $this->get_objective_number($objective->element);
					$id = $objective->value;
					$objective_ids[$n] = $id;
					if(!in_array($id,$this->all_question_ids)) {
						$this->all_question_ids[] = $id;
					}
				}

				$objectives = array();
				$objective_score_records = $DB->get_records_sql('SELECT element,value FROM {scorm_scoes_track} WHERE attempt = :attempt AND userid = :userid AND scormid = :scormid AND element LIKE \'cmi.objectives.%.score.raw\' ORDER BY element',array('attempt' => $scouser->attempt, 'userid' => $scouser->userid, 'scormid' => $scouser->scormid));
				foreach($objective_score_records as $objective) {
					$n = $this->get_objective_number($objective->element);
					$id = $objective_ids[$n];
					if(isset($objectives[$id])) {
						$objectives[$id] = max($objectives[$id],$objective->value);
					} else {
						$objectives[$id] = $objective->value;
					}
				}
				$scouser->objectives = $objectives;
			}

		}

		ksort($this->all_question_ids);

		$this->define_headers();
		$this->rows = array();

		foreach ($attempts as $scouser) {
			$row = array();

			$row[] = $scouser->username;

			$row[] = fullname($scouser);

			if (empty($scouser->timetracks->start)) {
				$row += ['','','','',''];
				foreach($this->all_question_ids as $id) {
					$row[] = '';
				}
			} else {
				$row[] = $scouser->attempt;

				if ($this->download =='ODS' || $this->download =='Excel' ) {
					$row[] = userdate($scouser->timetracks->start, get_string("strftimedatetime", "langconfig"));
					$row[] = userdate($scouser->timetracks->finish, get_string('strftimedatetime', 'langconfig'));
				} else {
					$row[] = userdate($scouser->timetracks->start);
					$row[] = userdate($scouser->timetracks->finish);
				}

				$row[] = scorm_grade_user_attempt($scorm, $scouser->userid, $scouser->attempt,false,true);

				$percentage = scorm_grade_user_attempt($scorm, $scouser->userid, $scouser->attempt,true,true);
				$row[] = number_format((float)$percentage*100, 2, '.', '');;

				foreach($this->all_question_ids as $id) {
					if(isset($scouser->objectives[$id])) {
						$row[] = $scouser->objectives[$id];
					} else {
						$row[] = '';
					}
				}
			}

			$this->rows[] = $row;
		}
	}
}

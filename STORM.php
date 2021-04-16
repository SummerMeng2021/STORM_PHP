<?php 
    # This program generates storm backup summary graph starting from year Jan_1st_2010 to current 
	$start_date = '1 Jan 2010';	
    $di = new Date_Iterator($start_date,false, Date_Iterator::YEAR);	
	//$di = new Date_Iterator($start_date);
	//$di->unit = 'YEAR';
	$a = array();
	foreach ($di as $value) $a[$value] = array(
		$di->unit=>$value, 		
		'Volume equal to Zero'=> 0,
		'Volume greater than Zero' => 0,
		'Total volume' => 0,		
	); 

	global  $oracle;
	global  $excel;
    $oracle = new Oracle_Connection();	
    $excel = new Excel_Workbook();	
	
	 $sql = <<<EOQ
select 
WORK_ORDER_NO as WORK_ORDER,
INF_WORK_ORDER_HISTORY.MAP_REFERENCE,
m.MAP_REFERENCE as SSA,
m.STRUCT_1,
m.STRUCT_2,
m.ACTIVITY_CODE,
PROBLEM_ADDRESS as ADDRESS,
STREET,
PROB_ZIP_CODE_NUMBER,
REQUESTED_DATE,
m.COMPLETED_DATE,
WO_DATE,
LOCATION,
CAUSE,
VOLUME,
PRIV,
BODY_OF_WATER,
CLEAR_WATER,
ACCESS_TO_BASEMENT
from INF_WORK_ORDER_HISTORY
left join INF_WORK_ORDER_MULTIPLE m using (WORK_ORDER_NO)
left join (
	select STREET_CODE as PROBLEM_STREET_CODE, DESCRIPTION as STREET
	from STREET_CODES
) using (PROBLEM_STREET_CODE)
left join (
	select SEQUENCE, WO_LOG_LOOKUP_VALUE as CAUSE
	from INF_WO_LOGGING
	where WO_LOG_CODE = 'BSCAUSE'
) using (SEQUENCE)
left join (
	select SEQUENCE, WO_LOG_LOOKUP_VALUE as LOCATION
	from INF_WO_LOGGING
	where WO_LOG_CODE = 'LOCBCKUP'
) using (SEQUENCE)
left join (
	select SEQUENCE, WO_LOG_OPEN_VALUE as VOLUME
	from INF_WO_LOGGING
	where WO_LOG_CODE = 'DESH2O'
) using (SEQUENCE)
left join (
	select SEQUENCE, WO_LOG_YN as PRIV
	from INF_WO_LOGGING
	where WO_LOG_CODE = 'PRIV'
) using (SEQUENCE)
left join (
	select SEQUENCE, WO_LOG_OPEN_VALUE as BODY_OF_WATER
	from INF_WO_LOGGING
	where WO_LOG_CODE = 'OVWAT'
) using (SEQUENCE)
left join (
	select SEQUENCE, WO_LOG_LOOKUP_VALUE as CLEAR_WATER
	from INF_WO_LOGGING
	where WO_LOG_CODE = 'CLORSW'
) using (SEQUENCE)
left join (
	select SEQUENCE, WO_LOG_YN as ACCESS_TO_BASEMENT
	from INF_WO_LOGGING
	where WO_LOG_CODE = 'HOMACC'
) using (SEQUENCE)
where m.ACTIVITY_CODE in('SCLNUPB','SFLBS')
	 and WO_DATE > '$start_date'
	 and WORK_ORDER_NO <> 'REMAY201828'
order by WO_DATE desc
EOQ;
	$rs = $oracle->query($sql);	
	
#Following is for storm Backups (Including High Water) per Year blanks should be there
	$matches = null;
	foreach ($oracle->query($sql) as $row) {
		extract($row);  
		$pattern = '/(\d*(\.\d+)?)/';
		$key = $di->getValue($WO_DATE);		
			if (preg_match($pattern, str_ireplace(",","",$VOLUME), $matches) && $matches[1] != '' && $matches[1] != 0 && $matches[1] != 'UNK' && $matches[1] != 'UNKNOWN') {
				$a[$key]['Volume greater than Zero']++;
				$a[$key]['Total volume'] += $matches[1];
			} else {
				$a[$key]['Volume equal to Zero']++;
			}		
		
	    
	}
	    
	
	$sheet = $excel->toSheet($rs);
	foreach (array('G', 'H', 'I') as $col) $sheet->columns($col)->numberFormat = 'd Mmm YYYY';
	#$sheet->pageSize11x17();
	$excel->Title = 'Basement Backups';
	$excel->Paragraph = 'P7';
	
	$chart = $excel->toBarChart($a, $di->unit, 'No of Basement backups', 'Basement Backups (Including High Water)');
	//$chart->colors(46,41,18);
	
	$chart->colors(3,23,50);
	
	for($i = 1; $i <= 2; $i++) {
		$chart->seriesCollection($i)->hasDataLabels = true;
		$chart->seriesCollection($i)->dataLabels->font->italic;
	}
			
	$chart->seriesCollection(3)->axisGroup = xlSecondary;
	$chart->seriesCollection(3)->chartType = xlLine;
	$chart->seriesCollection(3)->border->weight = xlThick;
	
	$chart->chart->axes(2, xlSecondary)->hasTitle = true;
	$chart->chart->axes(2, xlSecondary)->axisTitle->text ='Basement Backup Total Volume (gal)';
	$chart->chart->axes(2, xlSecondary)->tickLabels->numberFormat = '###,###';
	
	
		
	
	Highwater($oracle,$sql,$excel);	
	Cleaning($oracle,$sql,$excel);
	Grease($oracle,$sql,$excel);
	Roots($oracle,$sql,$excel);
	Unknown_Other($oracle,$sql,$excel);
	Exclude_HighWater($oracle,$sql,$excel);
	Average_Discharge($oracle,$sql,$excel);
	$excel->save('H:\Report\Backup_Summary.xls'); 

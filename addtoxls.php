<?php

set_time_limit (600);
 $workdir = "d:\\rosvid";

// $workdir = str_replace("\\", "/" , $work_dir);

$arr_sokr_kekv = array(1110, 1120, 1132, 1133, 1160, 1171, 1172, 1340, 5000);

$sheet1 = "otchet";


function addtoxls ($zv_type)
{
	global $workdir, $arr_sokr_kekv, $sheet1;


if ($zv_type == 'zf') { $filename = $workdir."\\zf.xls"; $fond = "=1"; } 
if ($zv_type == 'sf') { $filename = $workdir."\\sf.xls"; $fond = ">1"; }

$excel_app = new COM("Excel.application") or Die ("Did not connect");

$excel_app->Visible = 0;

$Workbook = $excel_app->Workbooks->Open("$filename") or Die("Did not open $filename $Workbook");
$Worksheet = $Workbook->Worksheets($sheet1);
$Worksheet->activate;


$z=23;

if ($db_sqlite = sqlite_open($workdir."\\bases\\rospis.sdb")) 
	{
    $query = sqlite_query($db_sqlite, 'select * from budgets group by BUDK order by KODBUDGET');
	while ($entry = sqlite_fetch_array($query, SQLITE_ASSOC)) {
    
	$coord_budk = "A".$z;

	$excel_cell_budk = $Worksheet->Range($coord_budk);
	$excel_cell_budk->activate;
	$excel_cell_budk->value = $entry['BUDK'];
	
	// echo $entry['BUDK']."\n";
	
	$y=$z+1;
	$z=$z+10;

	for($i=0; $i<count($arr_sokr_kekv); $i++)
	{
	
		$coord_kekv = "A".$y;
		//$coord_summa = "B".$y;
		$coord_s1 = "C".$y;
		$coord_s2 = "D".$y;
		$coord_s3 = "E".$y;
		$coord_s4 = "F".$y;
		$coord_s5 = "G".$y;
		$coord_s6 = "H". $y;
		$coord_s7 = "I".$y;
		$coord_s8 = "J".$y;
		$coord_s9 = "K".$y;
		$coord_s10 = "L".$y;
		$coord_s11 = "M".$y;
		$coord_s12 = "N".$y;	


	$query_kekv = sqlite_query($db_sqlite, 'select KEKV, KKFN, KODBUDGET, sum(SUMM) as SUMMA, sum(S1) as S1, sum(S2) as S2, sum(S3) as S3, sum(S4) as S4, sum(S5) as S5, sum(S6) as S6, sum(S7) as S7, sum(S8) as S8, sum(S9) as S9, sum(S10) as S10, sum(S11) as S11, sum(S12) as S12 from rov where KKFN'.$fond.' and kodbudget='.$entry['KODBUDGET'].' and kekv='.$arr_sokr_kekv[$i].' group by KODBUDGET');

			if (sqlite_num_rows($query_kekv) == 0)
		{
		
		$excel_cell_kekv = $Worksheet->Range($coord_kekv);
		$excel_cell_kekv->activate;
		$excel_cell_kekv->value = $arr_sokr_kekv[$i];
		
		//$excel_cell_summa = $Worksheet->Range($coord_summa);
		//$excel_cell_summa->activate;
		//$excel_cell_summa->value = 0;
		
		$excel_cell_s1 = $Worksheet->Range($coord_s1);
		$excel_cell_s1->activate;
		$excel_cell_s1->value = 0;
			
		$excel_cell_s2 = $Worksheet->Range($coord_s2);
		$excel_cell_s2->activate;
		$excel_cell_s2->value = 0;
		
		$excel_cell_s3 = $Worksheet->Range($coord_s3);
		$excel_cell_s3->activate;
		$excel_cell_s3->value = 0;

		$excel_cell_s4 = $Worksheet->Range($coord_s4);
		$excel_cell_s4->activate;
		$excel_cell_s4->value = 0;
		
		$excel_cell_s5 = $Worksheet->Range($coord_s5);
		$excel_cell_s5->activate;
		$excel_cell_s5->value = 0;

		$excel_cell_s6 = $Worksheet->Range($coord_s6);
		$excel_cell_s6->activate;
		$excel_cell_s6->value = 0;

		$excel_cell_s7 = $Worksheet->Range($coord_s7);
		$excel_cell_s7->activate;
		$excel_cell_s7->value = 0;

		$excel_cell_s8 = $Worksheet->Range($coord_s8);
		$excel_cell_s8->activate;
		$excel_cell_s8->value = 0;


		$excel_cell_s9 = $Worksheet->Range($coord_s9);
		$excel_cell_s9->activate;
		$excel_cell_s9->value = 0;

		$excel_cell_s10 = $Worksheet->Range($coord_s10);
		$excel_cell_s10->activate;
		$excel_cell_s10->value = 0;

		$excel_cell_s11 = $Worksheet->Range($coord_s11);
		$excel_cell_s11->activate;
		$excel_cell_s11->value = 0;

		$excel_cell_s12 = $Worksheet->Range($coord_s12);
		$excel_cell_s12->activate;
		$excel_cell_s12->value = 0;

		
		//echo $arr_sokr_kekv[$i].";0;0;0;0;0;0;0;0;0;0;0;0;0;0\n";
		
		$y=$y+1;
		}	
		else
		{
			while ($entry_kekv = sqlite_fetch_array($query_kekv, SQLITE_ASSOC)) 
	{
	
		$excel_cell_kekv = $Worksheet->Range($coord_kekv);
		$excel_cell_kekv->activate;
		$excel_cell_kekv->value = $entry_kekv['KEKV'];
		
		//$excel_cell_summa = $Worksheet->Range($coord_summa);
		//$excel_cell_summa->activate;
		//$excel_cell_summa->value = round($entry_kekv['SUMMA']/100000, 3);
		
		$excel_cell_s1 = $Worksheet->Range($coord_s1);
		$excel_cell_s1->activate;
		$excel_cell_s1->value = round($entry_kekv['S1']/100000, 3);
			
		$excel_cell_s2 = $Worksheet->Range($coord_s2);
		$excel_cell_s2->activate;
		$excel_cell_s2->value = round($entry_kekv['S2']/100000, 3);
		
		$excel_cell_s3 = $Worksheet->Range($coord_s3);
		$excel_cell_s3->activate;
		$excel_cell_s3->value = round($entry_kekv['S3']/100000, 3);

		$excel_cell_s4 = $Worksheet->Range($coord_s4);
		$excel_cell_s4->activate;
		$excel_cell_s4->value = round($entry_kekv['S4']/100000, 3);
		
		$excel_cell_s5 = $Worksheet->Range($coord_s5);
		$excel_cell_s5->activate;
		$excel_cell_s5->value = round($entry_kekv['S5']/100000, 3);

		$excel_cell_s6 = $Worksheet->Range($coord_s6);
		$excel_cell_s6->activate;
		$excel_cell_s6->value = round($entry_kekv['S6']/100000, 3);

		$excel_cell_s7 = $Worksheet->Range($coord_s7);
		$excel_cell_s7->activate;
		$excel_cell_s7->value = round($entry_kekv['S7']/100000, 3);

		$excel_cell_s8 = $Worksheet->Range($coord_s8);
		$excel_cell_s8->activate;
		$excel_cell_s8->value = round($entry_kekv['S8']/100000, 3);


		$excel_cell_s9 = $Worksheet->Range($coord_s9);
		$excel_cell_s9->activate;
		$excel_cell_s9->value = round($entry_kekv['S9']/100000, 3);

		$excel_cell_s10 = $Worksheet->Range($coord_s10);
		$excel_cell_s10->activate;
		$excel_cell_s10->value = round($entry_kekv['S10']/100000, 3);

		$excel_cell_s11 = $Worksheet->Range($coord_s11);
		$excel_cell_s11->activate;
		$excel_cell_s11->value = round($entry_kekv['S11']/100000, 3);

		$excel_cell_s12 = $Worksheet->Range($coord_s12);
		$excel_cell_s12->activate;
		$excel_cell_s12->value = round($entry_kekv['S12']/100000, 3);


	//echo $entry_kekv['KEKV'].";".$entry_kekv['SUMMA'].";".$entry_kekv['S1'].";".$entry_kekv['S2'].";".$entry_kekv['S3'].";".$entry_kekv['S4'].";".$entry_kekv['S5'].";".$entry_kekv['S6'].";".$entry_kekv['S7'].";".$entry_kekv['S8'].";".$entry_kekv['S9'].";".$entry_kekv['S10'].";".$entry_kekv['S11'].";".$entry_kekv['S12']."\n";
	
	$y=$y+1;
	}


		}


	}


}

	} 
	else 
		{
    echo "error!\n";
		}


sqlite_close($db_sqlite);


// closing excel

$excel_app->ActiveWorkbook->Save();

$excel_app->Quit();

// free the object
//$excel_app->Release();

$excel_app = null;

}

 addtoxls('zf');
 addtoxls('sf');

echo "מעקועû ספמנלטנמגאםû!";
?>

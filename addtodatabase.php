<?
set_time_limit (36000);

$arr_sokr_kekv = array(1110, 1120, 1132, 1133, 1160, 1171, 1172, 1340, 5000);

$f_year = "2010";

//$workdir = "d:\\rosvid";

$workdir = str_replace("\\", "/" , $work_dir);

//$tmpdbfdir = str_replace("\\", "/" , $tmp_dbf_dir);

// запись в массив данных справичника бюджетов

if ($db_sqlite = sqlite_open($workdir."\\bases\\rospis.sdb")) 
	{ 
    $result = sqlite_query($db_sqlite, 'select IDBUDGET, NAMETERR, KODBUDGET from budgets');
	$idbud_rows = sqlite_num_rows($result);   
	} 
	else 
	{
    echo "error!\n";
	}

// открытие папки с распакованными файлами данных dbf

if ($handle = opendir($workdir."\\tmpdbfdir")) 
	{
	
    while (false !== ($file = readdir($handle))) 
		{
		
		if ($file!=='.' and $file !=='..')
			{
			
			if (substr($file, 0, 3)=="ROT" or substr($file, 0, 3)=="RZT" ){$ftable = "ROT"; }
			if (substr($file, 0, 3)=="ROF" or substr($file, 0, 3)=="RZF"){$ftable = "ROF"; }
			if (substr($file, 0, 3)=="ROV" or substr($file, 0, 3)=="RZV"){$ftable = "ROV"; }

			$fl_pieces = explode(".", $file);
			

			$f_d = substr($fl_pieces[0], -3, 2);
			$f_m = substr($fl_pieces[0], -5, 2);
			
			$f_date = $f_year.$f_m.$f_d;

			echo $fl_pieces[1]."\n";
			
			$result = sqlite_query($db_sqlite, 'select * from budgets where IDBUDGET='.$fl_pieces[1]);
			//echo $idbud_rows = sqlite_num_rows($result);
			
			while ($entry = sqlite_fetch_array($result, SQLITE_ASSOC)) 
				{
				
						$db_dbf = dbase_open($workdir."\\tmpdbfdir\\".$file, 2);
						if ($db_dbf) 
						  {
						
							$record_numbers = dbase_numrecords($db_dbf);
							for ($i = 1; $i <= $record_numbers; $i++) 
								{
        						$row = dbase_get_record_with_names($db_dbf, $i); 


			if (substr($file, 0, 3)=="ROT" or substr($file, 0, 3)=="ROF" or substr($file, 0, 3)=="ROV"){$KKFN = $row['KKFN']; }
			if (substr($file, 0, 3)=="RZT" or substr($file, 0, 3)=="RZF" or substr($file, 0, 3)=="RZV"){$KKFN = $row['KKFB']; }
			

									if (in_array($row['KEKV'], $arr_sokr_kekv) != true)
								    {
								
								  	
									if (substr($row['KEKV'], 0, -1)==111)
										{
											$row['KEKV']=1110;
										}
									    elseif (substr($row['KEKV'], 0, -1)==116)
										{   
										$row['KEKV']=1160;
										}
										elseif (substr($row['KEKV'], 0, -1)==134)
										{
										$row['KEKV']=1340;
										}
										else
										{
										$row['KEKV']=5000;
										}
									
									}
									
									if ($row['MONTH']>0)
									{
									$sm='S'.$row['MONTH'];
									$s=$row['SUMM'];
									$row['SUMM']=0;
									$cols = "DATE, KEKV, KKFN, KODBUDGET, MONTH, ".$sm;
									$values = $f_date.", ".$row['KEKV'].", ".$KKFN.", ".$entry['KODBUDGET'].", ".$row['MONTH'].", ".$s;
									}
									else
										{
										$cols = "DATE, KEKV, KKFN, KODBUDGET, MONTH, SUMM";
										$values = $f_date.", ".$row['KEKV'].", ".$KKFN.", ".$entry['KODBUDGET'].", ".$row['MONTH'].", ".$row['SUMM'];
										}

						
								 $query = "insert into ".$ftable." (".$cols.") values (".$values.")";
								 //echo "\n";
								 sqlite_query($db_sqlite, $query);
								}
							
							}
							dbase_close($db_dbf);
				}
			}     
		} 
	}      
	else 
	{
	echo "error!";
	}
closedir($handle);
sqlite_close ($db_sqlite);
echo "Распакованные файлы обработаны. Данные добавлены в базу."

?>
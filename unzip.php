<?

$dbfile = str_replace("\\", "/" , $zip_file);
$extrdir = str_replace("\\", "/" , $extr_dir);

$zip = new ZipArchive;
if ($zip->open($zip_file) === TRUE) 
	{
    $zip->extractTo($extrdir."/tmpdbfdir/");
    $zip->close();
    echo "ОК! Все файлы распакованы в папку ".$extr_dir."\\tmpdbfdir\\";
	} 
	else 
	{
    echo "failed!\n";
	}
?>
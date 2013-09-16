<?php 
session_start();
$sessioncount=$_SESSION['count'];
$sessioncount++;
$_SESSION['count']=$sessioncount;
if($sessioncount>2000)
exit;
$arr=$_SESSION['arr'];
$id=$arr[$sessioncount];

set_time_limit(0);
if(isset($_GET['id']))
	$count=$_GET['id'];
else
	{
		echo 'id not set';
		exit;
	}

$curl = curl_init();
$ch=$curl;
// Set some options - we are passing in a useragent too here
curl_setopt_array($curl, array(
    CURLOPT_RETURNTRANSFER => 1,
    CURLOPT_URL => 'https://bigfuture.collegeboard.org/college-university-search/print-college-profile?id=' . $count,
    CURLOPT_USERAGENT => 'Mozilla Firefox'
));

curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false); 
//Send the request & save response to $resp
$resp = curl_exec($curl);
if($resp ==false)
	echo 'Curl error: ' . curl_error($curl);


include_once('simple_html_dom.php');

$name = explode('<title>', $resp);
$name2= explode('</title>', $name[1]);
$name3= explode('College Search - ',$name2[0]);


$collegename=$name3[1];



$pieces = explode('</script>', $name2[1]);
$pieces1= explode("<script",$pieces[1]);
$html = str_get_html($pieces1[0]);
$place=$html->find('h2',0)->innertext;

$h1tags = explode('<div class="mainTabSeperator"></div>',$html);
$cnth1tags=count($h1tags)-1;

curl_close($curl);

/** Error reporting */
error_reporting(E_ALL);

date_default_timezone_set('Europe/London');

/** Include PHPExcel */
require_once '/Classes/PHPExcel.php';

// Create new PHPExcel object
//echo date('H:i:s') , " Create new PHPExcel object" , PHP_EOL;
$objPHPExcel = new PHPExcel();

// Set document properties
//echo date('H:i:s') , " Set document properties" , PHP_EOL;
$objPHPExcel->getProperties()->setCreator("Kaushik Wavhal")
							 ->setLastModifiedBy("Kaushik Wavhal")
							 ->setTitle("US Colleges")
							 ->setSubject("US Colleges")
							 ->setDescription("US Colleges Data")
							 ->setKeywords("college")
							 ->setCategory("College Data file");
							 
$objPHPExcel->getDefaultStyle()->getFont()->setSize(12); 							 
							 
// Create a first sheet, representing sales data

$objPHPExcel->setActiveSheetIndex(0);


//Actual data
$objPHPExcel->getActiveSheet()->setCellValue('F1', 'Admission %');
$objPHPExcel->getActiveSheet()->setCellValue('G1', 'Testimonial');
$objPHPExcel->getActiveSheet()->setCellValue('H1', 'Selective Category (Super Selective/Selective/Medium/General)');
$objPHPExcel->getActiveSheet()->setCellValue('I1', 'Recommended Program');
$objPHPExcel->getActiveSheet()->setCellValue('J1', 'MMC Partner');
$objPHPExcel->getActiveSheet()->setCellValue('K1', 'Foundation Program(Y/N/Y by xxx)');
$objPHPExcel->getActiveSheet()->setCellValue('L1', 'Youtube Video');
$objPHPExcel->getActiveSheet()->getStyle('F1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('G1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('H1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('I1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('J1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('K1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('L1')->getFont()->setBold(true);

$objPHPExcel->getActiveSheet()->setCellValue('B1', $collegename);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setSize(20);
$objPHPExcel->getActiveSheet()->setCellValue('B2', $place);
$objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->setSize(17);

//set width
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(55);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(55);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(15);

$row=12;
$column=1;
$columnchar='A';

//==================================================================================================================================================================================


$cnt=1;
foreach($h1tags as $x)
{
	if($cnt==1)
	{
	$try=explode('>', $x);
	$len=count($try);
	$heading=substr($try[$len-2],0,count($try[$len-2])-5);
	$objPHPExcel->getActiveSheet()->setCellValue('A'.$row, $heading);
	$objPHPExcel->getActiveSheet()->getStyle('A'.$row)->getFont()->setBold(true);
	$objPHPExcel->getActiveSheet()->getStyle('A'.$row)->getFont()->setSize(20);
	$row++;
	$cnt++;
	}
	else if($cnt>1)
	{
	$finalhtml=str_get_html($x);



	foreach($finalhtml->find('tr') as $tblrow)
	{
	 $rowcontent=str_get_html($tblrow->innertext);
	  $tdcount=0;
	  $trrowtop=$row;
	  $trmax=$row;
	 	foreach($rowcontent->find('td') as $tbldata)
	    {
		$row=$trrowtop;
		$tdcount++;
		
		if($tdcount==1)
			$columnchar='A';
		else if ($tdcount==2)
			$columnchar='B';
		else if ($tdcount==3)
			$columnchar='C';
		else if ($tdcount==4)
			$columnchar='D';
		else if ($tdcount==5)
			$columnchar='E';		
		else if ($tdcount==6)
			$columnchar='F';
		else if ($tdcount==7)
			$columnchar='G';				
				
		$columncontent=str_get_html($tbldata->innertext);
		$innerheading=$columncontent->find('h2');
		     if(isset($innerheading[0]))
		     {
		     	$objPHPExcel->getActiveSheet()->setCellValue($columnchar . $row, $innerheading[0]->innertext);
		     	$objPHPExcel->getActiveSheet()->getStyle($columnchar . $row)->getFont()->setBold(true);
             		$objPHPExcel->getActiveSheet()->getStyle($columnchar . $row)->getFont()->setSize(17);
		     	$row++;
		  
		     }
		  
		  	 foreach($columncontent->find('p') as $para)
	         {
			   if(explode('<br/>',$para->innertext))
			   {
        			 $brtag= explode('<br/>',$para->innertext);	
			   				foreach($brtag as $br)
							{
							 
							 
							 if($brcontent=str_get_html($br))
								 {
							 	$boldtext=$brcontent->find('b');
								 if(isset($boldtext[0]))
							 		{
							 	$xyz=$brcontent->find('b');
								 $objPHPExcel->getActiveSheet()->getStyle($columnchar . $row)->getFont()->setBold(true);
							 	$objPHPExcel->getActiveSheet()->setCellValue($columnchar . $row, $boldtext[0]->innertext);
									 }
								 else
							 		{
							 	$objPHPExcel->getActiveSheet()->setCellValue($columnchar . $row, $br);
							 		}
							 	$row++;
							 	if($trmax<$row)
			                 				$trmax=$row;
								}
							}
			   		 
			   }
			   else
			   {
			   $objPHPExcel->getActiveSheet()->setCellValue($columnchar . $row, $para->innertext);		
			   $row++;
			   	if($trmax<$row)
			   		$trmax=$row;
			   }
			   
			   
			   
		     }
		  
		  
		  $row=$trmax+1;
		} //td for loop end
	 
	 
	 
	 $row++;
	}//tr for loop end
       
	   
	   
	   
	   
	   
	if($cnt<$cnth1tags) 
		{	   
		$try=explode('>', $x);
		$len=count($try);
		$heading=substr($try[$len-2],0,count($try[$len-2])-5);
		$objPHPExcel->getActiveSheet()->setCellValue('A'.$row, $heading);
		$objPHPExcel->getActiveSheet()->getStyle('A'.$row)->getFont()->setBold(true);
		$objPHPExcel->getActiveSheet()->getStyle('A'.$row)->getFont()->setSize(20);
		$row++;
		}


}//ifelse end
$cnt++;



}//main foreach h1 tag end





//==================================================================================================================================================================================

// Set header and footer. When no different headers for odd/even are used, odd header is assumed.
//echo date('H:i:s') , " Set header/footer" , PHP_EOL;
$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddHeader('&L&BPersonal cash register&RPrinted on &D');
$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' . $objPHPExcel->getProperties()->getTitle() . '&RPage &P of &N');


// Set page orientation and size
//echo date('H:i:s') , " Set page orientation and size" , PHP_EOL;
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);


// Rename worksheet
//echo date('H:i:s') , " Rename worksheet" , PHP_EOL;
$objPHPExcel->getActiveSheet()->setTitle('College Data');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Redirect output to a client�s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="' . $collegename . '.xls"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
header("Refresh: 0; url=http://localhost/met_tnp 2nd Feb-no cache/admin/curldemo.php?id=" . $id);
//exit;							 
?>

<?php

include('simple_html_dom.php');
include('product_url_config.php');
include('mysql_config.php');
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/Writer/Excel2007.php';
require_once 'Classes/PHPExcel/IOFactory.php';

set_time_limit(0);
error_reporting(E_ALL ^ E_NOTICE);
ini_set("memory_limit","512M");

$spreadsheet = PHPExcel_IOFactory::load("pavi.xlsx");
$spreadsheet->setActiveSheetIndex(0);
$worksheet = $spreadsheet->getActiveSheet();
$writer = new PHPExcel_Writer_Excel2007($spreadsheet);

//set the directory for the cookie using defined document root var
$path = $_SERVER["DOCUMENT_ROOT"].'/2d'.'/pavi';
$cookie_file_path = $path."/cookies.txt";

$con = mysqli_connect($host, $user, $password, $database);
if (!$con)  {
  die('Could not connect: ' . mysql_error());
}


$html = new simple_html_dom(); 
$html1 = new simple_html_dom();

for($u=0; $u < sizeof($plinks); $u++){
$url = $plinks[$u];
preg_match('@aspx[^\w]=(.*)@', $url, $mcategory);
$ecat = explode('-', $mcategory[1]);
$catno = (int)$ecat[0];
if($catno == 12)
	$category1 = 'Baby';
if($catno == 27)
	$category1 = 'Beauty';
if($catno == 18)
	$category1 = 'Beverages';
if($catno == 20)
	$category1 = 'Cellar';
if($catno == 25)
	$category1 = 'Chilled';
if($catno == 24)
	$category1 = 'Confectionery';
if($catno == 21)
	$category1 = 'Counters';
if($catno == 28)
	$category1 = 'Frozen';
if($catno == 14)
	$category1 = 'Groceries';
if($catno == 15)
	$category1 = 'Health';
if($catno == 13)
	$category1 = 'House';
if($catno == 16)
	$category1 = 'Pets';
if($catno == 26)
	$category1 = 'Simply & Pavi';
if($catno == 29)
	$category1 = 'XMAS';
preg_match("@aspx[^\w]=(.*)@", $url, $cmatch);
$categoryid = $cmatch[1];
/***************************************Phantomjs*********************************************************/
/***************************************First Page********************************************************/
$clenght = 0;
exec('phantomjs pageclick.js '.escapeshellarg($url).' > pavi.html');
for($j=0; $j < 1000; $j++){
	$contents = file_get_contents('pavi.html');
	$clenght = strlen($contents);
	if($clenght > 100000)
		break;
	sleep(5);
}
$contents = file_get_contents('pavi.html');
$clenght = 0;
$html->load($contents);
//$mcat = $html->find('#ctl00_ASPxMnuProducts_DXI0_T');
$subcat = $html->find('#ctl00_cphNestedMasterPage_pnlHeaderMain_lblHeaderMain');
//$category1 = $mcat[0]->title;
$category2 = $subcat[0]->plaintext;
echo $category1.'-'.$category2.'</br>';
$sprange = $html->find('b[class="dxp-lead"]');
$prange = $sprange[0]->plaintext;
//echo $prange.'</br>';
preg_match('@\((.*) items\)@', $prange, $match);
$pnum = (int)$match[1];
if($pnum == 0)
	continue;
//echo $pnum.'</br>';
preg_match_all("@pid=(.*)'}[^\w]@", $contents, $mpid);
$szp = sizeof($mpid[0]).'</br>';
$pid = array();
$href = array();
$id = array();

for($p=0; $p < $szp; $p++){
	$pid[$p] = str_replace("'})", '', $mpid[0][$p]);
	$href[$p] = 'http://www.pavi.com.mt/forms/ProductDetails.aspx?'.$pid[$p];
	preg_match('@pid=(.*)@', $pid[$p], $idmatch);
	$id[$p] = $idmatch[1];
	//echo $href[$p].' '.$id[$p].'</br>';
}

/***************************************Products less than 50*********************************************/
if($pnum < 50){
	echo 'Total no. of products < 50'.'</br>';
	//$szp
	for($i=0; $i < $szp; $i++){
	$h = $i;
	//echo $href[$h].' '.$id[$h].'</br>';
	$row = $spreadsheet->getActiveSheet()->getHighestRow()+1;
	$tr = $html->find('#ctl00_cphNestedMasterPage_gvItems_DXDataRow'.$i);
	if(sizeof($tr) == 0){
		echo 'Exited after inserting '.$i. 'records'.'</br>';
		exit;
	}	
	$span = $tr[0]->find('span.dxeBase');
	$product_name = $span[0]->plaintext;
	$img = $tr[0]->find('img.#ctl00_cphNestedMasterPage_gvItems_cell'.$i.'_3_btnImage_BImg');
	$imgurl = 'http://www.pavi.com.mt/'.$img[0]->src;
	$sptd = $tr[0]->find('td.dxgv');
	$stock = $sptd[2]->plaintext;
	$sprice = $sptd[3]->find('strike');
	$fprice = $sptd[3]->find('font');
	if(sizeof($sprice)> 0){
		$price_before = $sprice[0]->plaintext;
	}
	else{
		$price_before = ' ';
	}
	if(sizeof($fprice)> 0){
		$current_price = $fprice[0]->plaintext;
	}
	else{
		$current_price = $sptd[3]->plaintext;
	}
/***********************************Curl***********************************************************/
for($try = 0; $try < 3; $try++){
	$ch = curl_init();
	curl_setopt($ch, CURLOPT_HEADER, false);
	curl_setopt($ch, CURLOPT_NOBODY, false);
	curl_setopt($ch, CURLOPT_URL, $href[$h]);
	curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, 0);
	curl_setopt ($ch, CURLOPT_CAINFO, $path.'/CA.cer');
	curl_setopt($ch, CURLOPT_USERAGENT,
    "Mozilla/5.0 (Windows; U; Windows NT 5.0; en-US; rv:1.7.12) Gecko/20050915 Firefox/1.0.7");
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	curl_setopt($ch, CURLOPT_REFERER, 'http://www.pavi.com.mt/OnlineShop.aspx');
	$result = curl_exec($ch);
	curl_close($ch); 
	//echo $result;
	$html1->load($result);
	//echo strlen($result).'</br>';
	$table = $html1->find('.product_details_table');
	if(sizeof($table) == 0){
		$itemcode = ' ';
		$brand =' ';
		continue;
	}
	else{
		$tr = $table[0]->find('tr');
		$td = $tr[1]->find('td');
		$itemcode = $td[2]->plaintext;
		$td = $tr[4]->find('td');
		$brand = $td[2]->plaintext;
		$itemcode = trim($itemcode);
		//echo $product_name.' | '.$brand.' | '.$itemcode.' | '.$id[$h].' | '.$category1.' | '.$category2.' | '.$price_before.' | '.$current_price.' | '.$stock.' | '.$href[$h].' | '.$imgurl.'</br>';
		$sql = "INSERT INTO pavi (product_name, brand, itemcode, categoryid, productid, category1, category2, current_price, price_before, stock, product_page, image_url) VALUES ( '".$product_name."', '".$brand."' , '".$itemcode."' , '".$categoryid."', '".$productid."', '".$category1."', '".$category2."', '".$current_price."', '".$price_before."', '".$stock."', '".$href[$h]."', '".$imgurl."'  ) ";
		if (mysqli_query($con, $sql)) {
			//echo "New record created successfully";
		} else {
			echo "Error: " . $sql . "<br>" . mysqli_error($con);
		}
		$worksheet->SetCellValueByColumnAndRow(0, $row, $product_name);
		$worksheet->SetCellValueByColumnAndRow(1, $row, $brand);
		$worksheet->SetCellValueByColumnAndRow(2, $row, $itemcode);
		$worksheet->SetCellValueByColumnAndRow(3, $row, $categoryid);
		$worksheet->SetCellValueByColumnAndRow(4, $row, $id[$h]);
		$worksheet->SetCellValueByColumnAndRow(5, $row, $category1);
		$worksheet->SetCellValueByColumnAndRow(6, $row, $category2);
		$worksheet->SetCellValueByColumnAndRow(7, $row, $current_price);
		$worksheet->SetCellValueByColumnAndRow(8, $row, $price_before);
		$worksheet->SetCellValueByColumnAndRow(9, $row, $stock);
		$worksheet->SetCellValueByColumnAndRow(10, $row, $href[$h]);
		$worksheet->SetCellValueByColumnAndRow(11, $row, $imgurl);
		break;
	}
	
/***********************************Curl***********************************************************/	
	

/***************************************Phantomjs*********************************************************/
	
	
}
}
	
	$writer->save('pavi.xlsx');
	$html->clear();
	$html1->clear();		
	echo 'Done';

}

/***************************************Products less than 50*********************************************/
/***************************************First Page********************************************************/
else{
/***************************************Products more than 50*********************************************/
//echo 'Total no.of products > 50'.'</br>';
$psel =array();
if($pnum%50 == 0)
	$tpg = $pnum/50;
else
	$tpg = ($pnum/50)+1;
$ltpg = (int)$tpg+3;
$lip = 0;
for($pg=3; $pg < $ltpg; $pg++){
$plink = '#ctl00_cphNestedMasterPage_gvItems_DXPagerTop > a:nth-child('.$pg.')';
exec('phantomjs pageclickv1.js '.escapeshellarg($url).' '.escapeshellarg($plink).'> pavi.html');
for($j=0; $j < 100000; $j++){
	$contents = file_get_contents('pavi.html');
	$clenght = strlen($contents);
	if($clenght > 10000)
		break;
	sleep(5);
}
$contents = file_get_contents('pavi.html');
$html->load($contents);
$sprange = $html->find('b[class="dxp-lead"]');
$prange = $sprange[0]->plaintext;
//echo $prange.'</br>';
preg_match('@\((.*) items\)@', $prange, $match);
$pnum = (int)$match[1];
//echo $pnum.'</br>';
preg_match_all("@pid=(.*)'}[^\w]@", $contents, $mpid);
$szp = sizeof($mpid[0]).'</br>';
$pid = array();
$href = array();
$id = array();
for($p=0; $p < $szp; $p++){
	$pid[$p] = str_replace("'})", '', $mpid[0][$p]);
	$href[$p] = 'http://www.pavi.com.mt/forms/ProductDetails.aspx?'.$pid[$p];
	preg_match('@pid=(.*)@', $pid[$p], $idmatch);
	$id[$p] = $idmatch[1];
}
$h =0;
$liplimit = $szp+$lip;
//echo $liplimit.'--LIMIT ';
for($i=$lip; $i < $liplimit; $i++){
	$h = $i-$lip;
	$row = $spreadsheet->getActiveSheet()->getHighestRow()+1;
	$tr = $html->find('#ctl00_cphNestedMasterPage_gvItems_DXDataRow'.$i);
	if(sizeof($tr) == 0){
		echo 'Exited after inserting '.$i. 'records'.'</br>';
		exit;
	}	
	$span = $tr[0]->find('span.dxeBase');
	$product_name = $span[0]->plaintext;
	$img = $tr[0]->find('img.#ctl00_cphNestedMasterPage_gvItems_cell'.$i.'_3_btnImage_BImg');
	$imgurl = 'http://www.pavi.com.mt/'.$img[0]->src;
	$sptd = $tr[0]->find('td.dxgv');
	$stock = $sptd[2]->plaintext;
	$sprice = $sptd[3]->find('strike');
	$fprice = $sptd[3]->find('font');
	if(sizeof($sprice)> 0){
		$price_before = $sprice[0]->plaintext;
	}
	else{
		$price_before = ' ';
	}
	if(sizeof($fprice)> 0){
		$current_price = $fprice[0]->plaintext;
	}
	else{
		$current_price = $sptd[3]->plaintext;
	}
/***********************************Curl***********************************************************/
for($try = 0; $try < 3; $try++){
	$ch = curl_init();
	curl_setopt($ch, CURLOPT_HEADER, false);
	curl_setopt($ch, CURLOPT_NOBODY, false);
	curl_setopt($ch, CURLOPT_URL, $href[$h]);
	curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, 0);
	curl_setopt ($ch, CURLOPT_CAINFO, $path.'/CA.cer');
	curl_setopt($ch, CURLOPT_USERAGENT,
    "Mozilla/5.0 (Windows; U; Windows NT 5.0; en-US; rv:1.7.12) Gecko/20050915 Firefox/1.0.7");
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	curl_setopt($ch, CURLOPT_REFERER, 'http://www.pavi.com.mt/OnlineShop.aspx');
	$result = curl_exec($ch);
	curl_close($ch); 
	$html1->load($result);
	$table = $html1->find('.product_details_table');
	if(sizeof($table) == 0){
		$itemcode = ' ';
		$brand =' ';
		continue;
	}
	else{
		$tr = $table[0]->find('tr');
		$td = $tr[1]->find('td');
		$itemcode = $td[2]->plaintext;
		$td = $tr[4]->find('td');
		$brand = $td[2]->plaintext;
		$itemcode = trim($itemcode);
		//echo $product_name.' | '.$brand.' | '.$itemcode.' | '.$id[$h].' | '.$category1.' | '.$category2.' | '.$price_before.' | '.$current_price.' | '.$stock.' | '.$href[$h].' | '.$imgurl.'</br>';
		//( product_name, brand, itemcode, categoryid, productid, category1, category2, current_price, price_before, stock, product_page, image_url)
		$sql = "INSERT INTO pavi (product_name, brand, itemcode, categoryid, productid, category1, category2, current_price, price_before, stock, product_page, image_url) VALUES ( '".$product_name."', '".$brand."' , '".$itemcode."' , '".$categoryid."', '".$productid."', '".$category1."', '".$category2."', '".$current_price."', '".$price_before."', '".$stock."', '".$href[$h]."', '".$imgurl."'  ) ";
		if (mysqli_query($con, $sql)) {
			//echo "New record created successfully";
		} else {
			echo "Error: " . $sql . "<br>" . mysqli_error($con);
		}
		$worksheet->SetCellValueByColumnAndRow(0, $row, $product_name);
		$worksheet->SetCellValueByColumnAndRow(1, $row, $brand);
		$worksheet->SetCellValueByColumnAndRow(2, $row, $itemcode);
		$worksheet->SetCellValueByColumnAndRow(3, $row, $categoryid);
		$worksheet->SetCellValueByColumnAndRow(4, $row, $id[$h]);
		$worksheet->SetCellValueByColumnAndRow(5, $row, $category1);
		$worksheet->SetCellValueByColumnAndRow(6, $row, $category2);
		$worksheet->SetCellValueByColumnAndRow(7, $row, $current_price);
		$worksheet->SetCellValueByColumnAndRow(8, $row, $price_before);
		$worksheet->SetCellValueByColumnAndRow(9, $row, $stock);
		$worksheet->SetCellValueByColumnAndRow(10, $row, $href[$h]);
		$worksheet->SetCellValueByColumnAndRow(11, $row, $imgurl);
		break;
	}
	
/***********************************Curl***********************************************************/	
	

/***************************************Phantomjs*********************************************************/
	
	
}
}
$lip = $lip+50;
}	
	$writer->save('pavi.xlsx');	
}	
	$html->clear();
	$html1->clear();		
	echo 'Done';
}	
/***************************************Products more than 50*********************************************/
?>
 <!DOCTYPE html>
<html>
<?php
require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../vendor/codeplex/phpexcel/PHPExcel.php');


 Bigcommerce::configure(array(

        'store_url' => 'https://silverforte.com',
        'username' => 'scripting2',
        'api_key' => '736d29852c9c69fbf7e6f84682e9e2e33577679e'
    ));

use Bigcommerce\Api\Client as Bigcommerce;



Bigcommerce::verifyPeer(false);
//$start = microtime(true);

function get_order_products($xl_order_id)
{//func start
try {//try	

//text vars
$text_invoice = "Invoice # ";
$text_company = "SILVER FORTE";
$text_our_address = "640 S Hill St#758, Los Angeles, CA 90014, USA";
$text_bil_det = "Billing Details";
$text_ship_det = "Shipping Details";
$text_phone = "Phone: ";
$text_email = "Email: ";
$text_payment_method = "Payment Method: ";
$text_order_date = "Order Date: ";
$text_order = "Order: ";
$text_order_items = "ORDER ITEMS: ";
$text_shipping_method = "Shipping Method: ";
$text_qty = "Qty";
$text_sku = "SKU";
$text_name = "Product Name";
$text_price = "Price";
$text_price_total = "Total";
$text_subtotal = "Subtotal: ";
$text_shipping = "Shipping: ";
$text_grand_total = "Grand total: ";
$text_qty_total = "Qty Total: ";
$text_weight_total = "Weight Total: ";
$text_order_sf = "Order Online at www.silverforte.com";
$text_agreement = "THANK YOU FOR YOUR ORDER! RETURNS MUST BE MADE WITHIN 5 DAYS OF RECEIPT AND ARE SUBJECT TO APPROVAL BY SILVERFORTE. SILVERFORTE ONLY ACCEPTS MERCHANDISE THAT IS 1) SENT ERRONEOUSLY, 2) WITH MANUFACTURE DEFECTS OR 3) OF UNSATISFACTORY QUALITY. WE DO NOT ACCEPT DAMAGED RETURNS DUE TO WEAR AND TEAR. ACCTS 60 DAYS PAST DUE ARE SUBJECT TO 1.5% MONTHLY FINANCE CHARGES. ACCTS 90 DAYS PAST DUE ARE SENT TO COLLECTIONS AND ARE RESPONSIBLE FOR ALL LEGAL AND COLLECTION FEES INCURRED.";
$text_usd = "USD";
$text_tel = "tel: 213-266-8882";
$text_fax = "fax: 213-266-8921";
$text_web = "www.silverforte.com";
$text_email_body = "silver@silverforte.com";
$text_country_origin = "Country of origin: USA";

//END of text vars

$order = Bigcommerce::getOrder($xl_order_id);
//$customer_id = $order->customer_id;
//$date_shipped = $order->date_shipped;
//$how_many_items_shipped = $order->items_shipped;
$order_date = $order->date_created;
$order_date = date("l jS \of F Y h:i:s A", strtotime($order_date));
$subtotal = $order->subtotal_inc_tax;
$shipping = $order->shipping_cost_inc_tax;
$grand_total = $order->total_inc_tax;
$payment_method = $order->payment_method;
$shipping_method = $order->shipping_method;
//$items_total = $order->items_total;
//billing address
$billing_address_first_name = $order->billing_address->first_name;
$billing_address_last_name = $order->billing_address->last_name;
$billing_address_company = $order->billing_address->company;
$billing_address_street_1 = $order->billing_address->street_1;
$billing_address_street_2 = $order->billing_address->street_2;
$billing_address_city = $order->billing_address->city; 
$billing_address_state = $order->billing_address->state;  
$billing_address_zip = $order->billing_address->zip;
$billing_address_country = $order->billing_address->country;
$billing_address_phone = $order->billing_address->phone;
$billing_address_email = $order->billing_address->email;
//END of billing address
//shipping address
$shipping_addresses = $order->shipping_addresses;
$shipping_addresses_first_name = array_map(create_function('$fs', 'return $fs->first_name;'), $shipping_addresses);
$shipping_addresses_last_name = array_map(create_function('$fs', 'return $fs->last_name;'), $shipping_addresses);
$shipping_addresses_company = array_map(create_function('$fs', 'return $fs->company;'), $shipping_addresses);
$shipping_addresses_street_1 = array_map(create_function('$fs', 'return $fs->street_1;'), $shipping_addresses);
$shipping_addresses_street_2 = array_map(create_function('$fs', 'return $fs->street_2;'), $shipping_addresses);
$shipping_addresses_city = array_map(create_function('$fs', 'return $fs->city;'), $shipping_addresses);
$shipping_addresses_zip = array_map(create_function('$fs', 'return $fs->zip;'), $shipping_addresses);
$shipping_addresses_country = array_map(create_function('$fs', 'return $fs->country;'), $shipping_addresses);
$shipping_addresses_state = array_map(create_function('$fs', 'return $fs->state;'), $shipping_addresses);
$shipping_addresses_email = array_map(create_function('$fs', 'return $fs->email;'), $shipping_addresses);
$shipping_addresses_phone = array_map(create_function('$fs', 'return $fs->phone;'), $shipping_addresses);

foreach ($shipping_addresses_first_name as $shipping_first_name);
foreach ($shipping_addresses_last_name as $shipping_last_name);
foreach ($shipping_addresses_company as $shipping_company);
foreach ($shipping_addresses_street_1 as $shipping_street_1);
foreach ($shipping_addresses_street_2 as $shipping_street_2);
foreach ($shipping_addresses_city as $shipping_city);
foreach ($shipping_addresses_zip as $shipping_zip);
foreach ($shipping_addresses_country as $shipping_country);
foreach ($shipping_addresses_state as $shipping_state);
foreach ($shipping_addresses_phone as $shipping_phone);
foreach ($shipping_addresses_email as $shipping_email);
//END of shipping address

//INIT
/* LOAD FILE
$inputFileName = 'invoice.xlsx'; //Excel file to write
$objPHPExcel = PHPExcel_IOFactory::load($inputFileName);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
*/
//Create NEW FILE
$objPHPExcel = new PHPExcel();
// Set properties
$objPHPExcel->getProperties()->setTitle("Invoice ".$xl_order_id.".xlsx");
//				->setSubject("Office 2007 XLSX Test Document")
//				->setDescription("Test doc for Office 2007 XLSX, generated by PHPExcel.")
//				->setKeywords("office 2007 openxml php")
//				->setCategory("Test result file");

$objPHPExcel = new PHPExcel();
//$objPHPExcel->getActiveSheet()->setTitle('test');

$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
$cacheSettings = array(' memoryCacheSize ' => '8MB');
PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
//END INIT
//merging cells
$objPHPExcel->getActiveSheet()->mergeCells('B8:D8');
$objPHPExcel->getActiveSheet()->mergeCells('A1:E6');
$objPHPExcel->getActiveSheet()->mergeCells('B9:F9');
$objPHPExcel->getActiveSheet()->mergeCells('B10:C10');
$objPHPExcel->getActiveSheet()->mergeCells('B11:C11');
$objPHPExcel->getActiveSheet()->mergeCells('B12:D12');
$objPHPExcel->getActiveSheet()->mergeCells('B13:D13');
$objPHPExcel->getActiveSheet()->mergeCells('B15:C15');
$objPHPExcel->getActiveSheet()->mergeCells('I15:J15');
$objPHPExcel->getActiveSheet()->mergeCells('B27:C27');
$objPHPExcel->getActiveSheet()->mergeCells('I26:J26');
$objPHPExcel->getActiveSheet()->mergeCells('I27:J27');
$objPHPExcel->getActiveSheet()->mergeCells('E31:F31');
$objPHPExcel->getActiveSheet()->mergeCells('B16:D16');
$objPHPExcel->getActiveSheet()->mergeCells('B17:D17');
$objPHPExcel->getActiveSheet()->mergeCells('B18:D18');
$objPHPExcel->getActiveSheet()->mergeCells('B19:D19');
$objPHPExcel->getActiveSheet()->mergeCells('B20:D20');
$objPHPExcel->getActiveSheet()->mergeCells('B21:D21');
$objPHPExcel->getActiveSheet()->mergeCells('B22:D22');
$objPHPExcel->getActiveSheet()->mergeCells('B23:D23');
$objPHPExcel->getActiveSheet()->mergeCells('I16:K16');
$objPHPExcel->getActiveSheet()->mergeCells('I17:K17');
$objPHPExcel->getActiveSheet()->mergeCells('I18:K18');
$objPHPExcel->getActiveSheet()->mergeCells('I19:K19');
$objPHPExcel->getActiveSheet()->mergeCells('I20:K20');
$objPHPExcel->getActiveSheet()->mergeCells('I21:K21');
$objPHPExcel->getActiveSheet()->mergeCells('I22:K22');
$objPHPExcel->getActiveSheet()->mergeCells('I23:K23');
//END of merging static cells
//Billing Details
$objPHPExcel->getActiveSheet()->setCellValue('B16', $billing_address_first_name." ".$billing_address_last_name);
$objPHPExcel->getActiveSheet()->setCellValue('B17', $billing_address_company);
$objPHPExcel->getActiveSheet()->setCellValue('B18', $billing_address_street_1);
$objPHPExcel->getActiveSheet()->setCellValue('B19', $billing_address_street_2);
$objPHPExcel->getActiveSheet()->setCellValue('B20', $billing_address_city.", ".$billing_address_state." ".$billing_address_zip);
$objPHPExcel->getActiveSheet()->setCellValue('B21', $billing_address_country);
$objPHPExcel->getActiveSheet()->setCellValue('B22', $text_phone." ".$billing_address_phone);
$objPHPExcel->getActiveSheet()->setCellValue('B23', $text_email." ".$billing_address_email);
//END billing details

//Shipping Details
$objPHPExcel->getActiveSheet()->setCellValue('I16', $shipping_first_name." ".$shipping_last_name);
$objPHPExcel->getActiveSheet()->setCellValue('I17', $shipping_company);
$objPHPExcel->getActiveSheet()->setCellValue('I18', $shipping_street_1);
$objPHPExcel->getActiveSheet()->setCellValue('I19', $shipping_street_2);
$objPHPExcel->getActiveSheet()->setCellValue('I20', $shipping_city.", ".$shipping_state." ".$shipping_zip);
$objPHPExcel->getActiveSheet()->setCellValue('I21', $shipping_country);
$objPHPExcel->getActiveSheet()->setCellValue('I22', $text_phone." ".$shipping_phone);
$objPHPExcel->getActiveSheet()->setCellValue('I23', $text_email." ".$shipping_email);
//END shipping details


//static cell text
$objPHPExcel->getActiveSheet()->setCellValue('B8', $text_company);
$objPHPExcel->getActiveSheet()->setCellValue('B9', $text_our_address);
$objPHPExcel->getActiveSheet()->setCellValue('B10', $text_tel);
$objPHPExcel->getActiveSheet()->setCellValue('B11', $text_fax);
$objPHPExcel->getActiveSheet()->setCellValue('B12', $text_web);
$objPHPExcel->getActiveSheet()->setCellValue('B13', $text_email_body);
$objPHPExcel->getActiveSheet()->setCellValue('B15', $text_bil_det);
$objPHPExcel->getActiveSheet()->setCellValue('I15', $text_ship_det);
$objPHPExcel->getActiveSheet()->setCellValue('B26', $text_order);
$objPHPExcel->getActiveSheet()->setCellValue('B27', $text_payment_method);
$objPHPExcel->getActiveSheet()->setCellValue('B28', $text_country_origin);
$objPHPExcel->getActiveSheet()->setCellValue('I26', $text_order_date);
$objPHPExcel->getActiveSheet()->setCellValue('I27', $text_shipping_method);
//$objPHPExcel->getActiveSheet()->setCellValue('B29', $text_order_items);
$objPHPExcel->getActiveSheet()->setCellValue('B31', $text_qty);
$objPHPExcel->getActiveSheet()->setCellValue('C31', $text_sku);
$objPHPExcel->getActiveSheet()->setCellValue('E31', $text_name);
$objPHPExcel->getActiveSheet()->setCellValue('J31', $text_price);
$objPHPExcel->getActiveSheet()->setCellValue('L31', $text_price_total);
//END static cell text

//order details
$objPHPExcel->getActiveSheet()->setCellValue('D26', "#".$xl_order_id);
$objPHPExcel->getActiveSheet()->setCellValue('D27', $payment_method);
$objPHPExcel->getActiveSheet()->setCellValue('K26', $order_date);
$objPHPExcel->getActiveSheet()->setCellValue('K27', $shipping_method);
//END order details
/*
//styling static
$objPHPExcel->getActiveSheet()->getStyle("B8")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B9")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B10")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B11")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B12")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B13")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B15")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("I15")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B26")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B27")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B28")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("I26")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("I27")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B29")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B31")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("C31")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("E31")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("I31")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("K31")->getFont()->setBold(true);
*/
$objPHPExcel->getActiveSheet()->getStyle("B15:I15")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle("B31:L31")->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(8);
//END styling static

//logo drawing
$logo_image = imagecreatefromjpeg('logo.jpg');
$objDrawing = new PHPExcel_Worksheet_MemoryDrawing();
$objDrawing->setName('SF');$objDrawing->setDescription('SilverForte');
$objDrawing->setImageResource($logo_image);
$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
$objDrawing->setHeight(100);
$objDrawing->setOffsetX(7);
$objDrawing->setOffsetY(8);
$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
$objDrawing->setCoordinates('A1');
//END of LOgo Drawing

//ORDER ITEMS
//$array = (array)$products;
$products = $order->products;
global $total_weight;
$total_weight = 0;
global $items_total;
$items_total = 0;
global $counter;
if($counter == 0){
$counter = 32;
}
elseif($counter != 0){
$counter = $counter;
}
foreach ($products as $product) {
/*
	$file = 'test.txt';
			$output = print_r($product, true);	
			file_put_contents($file, $output."\r\n", FILE_APPEND | LOCK_EX);
*/
	global $total_weight;
	global $items_total;
	 $product_id = $product->product_id;
     $product_sku = $product->sku;
	 //echo $product_sku." ";
	 $product_name = $product->name;
	 $product_price = $product->price_inc_tax." ".$text_usd;
	 $product_price_total = $product->total_inc_tax." ".$text_usd;
	 $product_weight = $product->weight;//for post-processing
	 $product_quantity = $product->quantity;//to count how many pieces we have in order

	$product_all_data = Bigcommerce::getProduct($product_id);	
	if (!empty($product_all_data->primary_image->thumbnail_url)) {
$item_url = $product_all_data->primary_image->thumbnail_url;
$img_name = 'C:\wamp64\www\bigcommerceDemo\Get_by_Filter\invoice_images\\'."ID".$product_id.".jpg";
file_put_contents($img_name, file_get_contents($item_url));
	}
	else{
	$img_name = "imageless.jpg";
	}

$total_weight = $total_weight + $product_weight;
$items_total = $items_total + $product_quantity;

$product_price = round($product_price, 2);
$product_price_total = round($product_price_total, 2);
global $counter;
$objPHPExcel->getActiveSheet()->mergeCells("E".($counter).":I".($counter));

$objPHPExcel->getActiveSheet()->getRowDimension($counter)->setRowHeight(40);
$gdImage = imagecreatefromjpeg($img_name);
$objDrawing = new PHPExcel_Worksheet_MemoryDrawing();
$objDrawing->setName('SF');$objDrawing->setDescription('SilverForte');
$objDrawing->setImageResource($gdImage);
$objDrawing->setRenderingFunction(PHPExcel_Worksheet_MemoryDrawing::RENDERING_JPEG);
$objDrawing->setMimeType(PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT);
$objDrawing->setHeight(40);
$objDrawing->setOffsetX(7);
$objDrawing->setOffsetY(8);
$objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
$objDrawing->setCoordinates('A'.$counter);

$objPHPExcel->getActiveSheet()->setCellValue('B'.$counter, $product_quantity);
$objPHPExcel->getActiveSheet()->setCellValue('C'.$counter, $product_sku);
$objPHPExcel->getActiveSheet()->setCellValue('E'.$counter, $product_name);
$objPHPExcel->getActiveSheet()->setCellValue('J'.$counter, $product_price);
$objPHPExcel->getActiveSheet()->setCellValue('L'.$counter, $product_price_total);
//$objPHPExcel->getActiveSheet()->getStyle('B'.$counter.$objPHPExcel->getActiveSheet()->getHighestRow())->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_CENTER);


$counter = $counter+1;
}

$next_lines = $counter + 1;
$next_lines_2 = $next_lines + 1;
$next_lines_3 = $next_lines_2 + 1;
$next_lines_4 = $next_lines_3 + 2;
$next_lines_5 = $next_lines_4 + 2;
$next_lines_5_shift = $next_lines_5 + 5;

$objPHPExcel->getActiveSheet()->getStyle('E1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('E1:I'.$objPHPExcel->getActiveSheet()->getHighestRow())->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_TOP);
$objPHPExcel->getActiveSheet()->mergeCells("B".($next_lines_2).":C".($next_lines_2));
$objPHPExcel->getActiveSheet()->mergeCells("I".($next_lines_2).":J".($next_lines_2));
$objPHPExcel->getActiveSheet()->mergeCells("I".($next_lines_3).":J".($next_lines_3));
$objPHPExcel->getActiveSheet()->mergeCells("D".($next_lines_4).":J".($next_lines_4));
$objPHPExcel->getActiveSheet()->mergeCells("B".($next_lines_5).":K".($next_lines_5_shift));
//variables
$objPHPExcel->getActiveSheet()->setCellValue('D'.$next_lines, $items_total);
$objPHPExcel->getActiveSheet()->setCellValue('D'.$next_lines_2, $total_weight);
$objPHPExcel->getActiveSheet()->setCellValue('K'.$next_lines, $subtotal);
$objPHPExcel->getActiveSheet()->setCellValue('K'.$next_lines_2, $shipping);
$objPHPExcel->getActiveSheet()->setCellValue('K'.$next_lines_3, $grand_total);
//END variables
//text vars
$objPHPExcel->getActiveSheet()->setCellValue('B'.$next_lines, $text_qty_total);
$objPHPExcel->getActiveSheet()->setCellValue('B'.$next_lines_2, $text_weight_total);
$objPHPExcel->getActiveSheet()->setCellValue('I'.$next_lines, $text_subtotal);
$objPHPExcel->getActiveSheet()->setCellValue('I'.$next_lines_2, $text_shipping);
$objPHPExcel->getActiveSheet()->setCellValue('I'.$next_lines_3, $text_grand_total);
//END text vars

$objPHPExcel->getActiveSheet()->setCellValue('D'.$next_lines_4, $text_order_sf);
$objPHPExcel->getActiveSheet()->setCellValue('B'.$next_lines_5, $text_agreement);
$objPHPExcel->getActiveSheet()->getStyle('B'.$next_lines_5)->getAlignment()->setWrapText(true);

$objPHPExcel->getActiveSheet()->getStyle('B'.$next_lines)->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('B'.$next_lines_2)->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('I'.$next_lines)->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('I'.$next_lines_2)->getFont()->setBold(true);
$objPHPExcel->getActiveSheet()->getStyle('I'.$next_lines_3)->getFont()->setBold(true);
//echo "Counter: ".$counter."\r\n";
//END of THE ORDER ITEMS
$objWriter->save("Invoice ".$xl_order_id.".xlsx");

$filename = 'http://localhost/bigcommerceDemo/Get_by_Filter/'."Invoice ".$xl_order_id.".xlsx";   

echo '
 <body>
<h1 style="color: #5e9ca0;">Done! Your excel Invoice was created!</h1>
<h2 style="color: #2e6c80;">&nbsp;</h2>
<ol style="list-style: none; font-size: 25px; line-height: 32px; font-weight: bold;">
<li style="clear: both;"><img style="float: left;" src="https://html-online.com/img/6-table-div-html.png" alt="html table div" width="45" /></li>
<a href="' . $filename . '">Download it from here!</a>
</ol>
<p>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</p>
<h2 style="color: #2e6c80;">&nbsp;</h2>
<p><strong>&nbsp;</strong></p>

  </body>
</html>
';

		}//end try
					        catch(Bigcommerce\Api\Error $error) {
							echo $error->getCode();
							echo $error->getMessage();
																}

}//end function
/*
	    $myfile = "Orders2.xlsx";//file orders IDs
		$excelReader = PHPExcel_IOFactory::createReaderForFile($myfile);
		$excelObj = $excelReader->load($myfile);
		$worksheet = $excelObj->getSheet(0);
		$lastRow = $worksheet->getHighestRow();

//WARNING!Put how many rows you have in file!!!
//$progressBar = new \ProgressBar\Manager(0,3);

		for ($row = 2; $row <= $lastRow; $row++) 
{	
//  	$progressBar->update($row);

		$xl_order_id = $worksheet->getCell('A'.$row)->getValue();
//call the function
		get_order_products  ($xl_order_id,
                            $row);
}								
*/
$xl_order_id = trim($_REQUEST['xl_order_id']);
get_order_products  ($xl_order_id);

?>


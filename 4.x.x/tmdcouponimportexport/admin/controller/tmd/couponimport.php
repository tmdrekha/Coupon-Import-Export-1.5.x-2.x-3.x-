<?php 
namespace Opencart\Admin\Controller\Extension\Tmdcouponimportexport\Tmd;
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/Psr/autoloader.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/myclabs/Enum.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/ZipStream/autoloader.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/ZipStream/ZipStream.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/PhpSpreadsheet/autoloader.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xls;

class Couponimport extends \Opencart\System\Engine\Controller {	
	
	public function index() {	
		
		$totalnewcoupon=0;
		$totalupdatecoupon=0;
		$this->language->load('extension/tmdcouponimportexport/tmd/couponimport');

		$this->document->setTitle($this->language->get('heading_title1'));
		
		$this->load->model('extension/tmdcouponimportexport/tmd/couponimport');
				
		if ($this->request->server['REQUEST_METHOD'] == 'POST' && $this->user->hasPermission('modify', 'extension/tmdcouponimportexport/tmd/couponimport')) {
			
			if (is_uploaded_file($this->request->files['import']['tmp_name'])) {
				$content = file_get_contents($this->request->files['import']['tmp_name']);
			} else {
				$content = false;
			}
			
			if ($content) {
				////////////////////////// Started Import work  //////////////
			////////////////////////// Started Import work  //////////////
				$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
				$spreadsheet = $reader->load($_FILES['import']['tmp_name']);
				$spreadsheet->setActiveSheetIndex(0);
				$sheetDatas = $spreadsheet->getActiveSheet()->toArray(null,true,true,true);
				$i=0;
				/*
				@ arranging the data according to our need
				*/
				foreach($sheetDatas as $sheetData){
					if($i!=0)
					{
					
					/* Step Customer Collect Data */
					$coupon_id=$sheetData['A'];
					if(!empty($coupon_id)){
						
						$coupon_id=$this->model_extension_tmdcouponimportexport_tmd_couponimport->getCouponByCode($sheetData['C']);
					}
					$name=$sheetData['B'];
					$code=$sheetData['C'];
					$type=$sheetData['D'];
					$discount=$sheetData['E'];
					$total=$sheetData['F'];
					
					$logged=$sheetData['G'];
					$shipping=$sheetData['H'];
					
					$product=array();
					$products=$sheetData['I'];
					if(!empty($products))
					{
							$products=explode(',',$products);
							foreach($products as $productid)
							{	
									if(!empty($productid)){
										$product[]=$productid;
									}
							}
					}
					$category=array();
					$categories=$sheetData['K'];
					
					if(!empty($categories))
					{
							$categories=explode(',',$categories);
							foreach($categories as $categoryid)
							{	
									if(!empty($categoryid)){
										$category[]=$categoryid;
									}
							}
					}
					
					
					$date_start=$sheetData['M'];
					$date_end=$sheetData['N'];
					
					$uses_total=$sheetData['O'];
					$uses_customer=$sheetData['P'];
					
					
					//////////////// status
					$status=$sheetData['Q'];
					if(strtolower($status)=='enabled'){
						$status=1;
					}else{$status=0;}
					
					/* Step Customer Collect Data */
			
				$data=[];
				$data=[
				'coupon_id'=>$coupon_id,
				'name'=>$name,
				'code'=>$code,
				'type'=>$type,
				'discount'=>$discount,
				'total'=>$total,
				'logged'=>$logged,
				'shipping'=>$shipping,
				'coupon_product'=>$products,
				'coupon_category'=>$category,
				'date_start'=>$date_start,
				'date_end'=>$date_end,
				'uses_customer'=>$uses_customer,
				'uses_total'=>$uses_total,
				'status'=>$status,
				];
				if(empty($coupon_id)){ 
					if(!empty($code)){
						$this->model_extension_tmdcouponimportexport_tmd_couponimport->addCoupon($data);
						$totalnewcoupon++;
						$this->session->data['success']= $totalnewcoupon .' :: Total New Coupon added';
					}
				}else{
					$this->model_extension_tmdcouponimportexport_tmd_couponimport->editCoupon($data,$coupon_id);
					$totalupdatecoupon++;
					$this->session->data['success']= $totalupdatecoupon .' :: Total Coupon update';

				}
		       }
					$i++;
					
			}
				
				////////////////////////// Started Import work  //////////////
				$this->response->redirect($this->url->link('extension/tmdcouponimportexport/tmd/couponimport', 'user_token=' . $this->session->data['user_token'], true));
			} else {
				$json['error']['warning'] = $this->language->get('error_empty');
			}
		
		}
		if (isset($this->session->data['error'])) {
			$data['error_warning'] = $this->session->data['error'];
			unset($this->session->data['error']);
		} elseif (isset($json['error']['warning'])) {
			$data['error_warning'] = $json['error']['warning'];
		} else {
			$data['error_warning'] = '';
		}


		  $data['heading_title1'] = $this->language->get('heading_title1');
		
		
		if (isset($this->session->data['success'])) {
			$data['success'] = $this->session->data['success'];
		
			unset($this->session->data['success']);
		} else {
			$data['success'] = '';
		}
		
  		$data['breadcrumbs'] = array();

   		$data['breadcrumbs'][] = array(
       		'text'      => $this->language->get('text_home'),
			'href'      => $this->url->link('common/home', 'user_token=' . $this->session->data['user_token'], true),     		
   		);

   		$data['breadcrumbs'][] = array(
       		'text'      => $this->language->get('heading_title'),
			'href'      => $this->url->link('extension/tmdcouponimportexport/tmd/couponimport', 'user_token=' . $this->session->data['user_token'], true),
   		);
		
		$data['import'] = $this->url->link('extension/tmdcouponimportexport/tmd/couponimport', 'user_token=' . $this->session->data['user_token'], true);

		$data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');
				
		
		$this->response->setOutput($this->load->view('extension/tmdcouponimportexport/tmd/couponimport', $data));
	}
	
	
}
?>
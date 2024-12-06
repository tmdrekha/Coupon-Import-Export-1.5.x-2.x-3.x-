<?php 
// Lib Include 
require_once(DIR_SYSTEM.'/library/tmd/Psr/autoloader.php');
require_once(DIR_SYSTEM.'/library/tmd/myclabs/Enum.php');
require_once(DIR_SYSTEM.'/library/tmd/ZipStream/autoloader.php');
require_once(DIR_SYSTEM.'/library/tmd/ZipStream/ZipStream.php');
require_once(DIR_SYSTEM.'/library/tmd/PhpSpreadsheet/autoloader.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// Lib Include 

class ControllerExtensioncouponimport extends Controller { 
	private $error = array();
	
	public function index() {	
		$totalnewcoupon=0;
		$totalupdatecoupon=0;
		$this->language->load('extension/couponimport');

		$this->document->setTitle($this->language->get('heading_title'));
		
		$this->load->model('extension/couponimport');
				
		if ($this->request->server['REQUEST_METHOD'] == 'POST' && $this->user->hasPermission('modify', 'extension/couponimport')) {
			
			if (is_uploaded_file($this->request->files['import']['tmp_name'])) {
				$content = file_get_contents($this->request->files['import']['tmp_name']);
			} else {
				$content = false;
			}
			
			if ($content) {
				$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
				$path_parts = pathinfo($this->request->files['import']['name']);
				$extension = $path_parts['extension'];
				
				if('xls' == $extension) {     
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
						
						$coupon_id=$this->model_extension_couponimport->getCouponByCode($sheetData['C']);
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
					//////////////// status
					
					/* Step Customer Collect Data */
			
				$data='';
				$data=array(
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
				);
						
						if(empty($coupon_id)){ 
							if(!empty($code)){
								$this->model_extension_couponimport->addCoupon($data);
								$totalnewcoupon++;
								 $this->session->data['success']=$totalnewcoupon .':: Total New Coupon added';
				

							}
						}else{
							$this->model_extension_couponimport->editCoupon($data,$coupon_id);
							$totalupdatecoupon++;
							$this->session->data['success']=$totalupdatecoupon.' :: Total Coupon update ';
						}
		}
					$i++;
					
				
				}
				
				////////////////////////// Started Import work  //////////////
				$this->response->redirect($this->url->link('extension/couponimport', 'user_token=' . $this->session->data['user_token'], true));
				} else{     
					$this->error['warning'] = $this->language->get('error_empty');
				}

				
			} else {
				$this->error['warning'] = $this->language->get('error_empty');
			}
		}

		$data['heading_title'] = $this->language->get('heading_title');
		$data['button_import'] = $this->language->get('button_import');
		$data['entry_import'] = $this->language->get('entry_import');
		
		
		
		if (isset($this->session->data['error'])) {
    		$data['error_warning'] = $this->session->data['error'];
    unset($this->session->data['error']);
 		}elseif (isset($this->session->data['warning'])) {
    	$data['error_warning'] = $this->session->data['warning'];
    unset($this->session->data['warning']);
 		}
 elseif (isset($this->error['warning'])) {
			$data['error_warning'] = $this->error['warning'];
		} else {
			$data['error_warning'] = '';
		}
		
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
			'href'      => $this->url->link('extension/couponimport', 'user_token=' . $this->session->data['user_token'], true),
   		);
		
		$data['import'] = $this->url->link('extension/couponimport', 'user_token=' . $this->session->data['user_token'], true);
				
		$data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');

		$this->response->setOutput($this->load->view('extension/couponimport', $data));
	}
	
	
}
?>
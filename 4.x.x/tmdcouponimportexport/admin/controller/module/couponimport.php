<?php 
namespace Opencart\Admin\Controller\Extension\Tmdcouponimportexport\Module;
/* EXPORT STARTS */
// Lib Include 
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/system.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/Psr/autoloader.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/myclabs/Enum.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/ZipStream/autoloader.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/ZipStream/ZipStream.php');
require_once(DIR_EXTENSION.'/tmdcouponimportexport/system/library/tmd/PhpSpreadsheet/autoloader.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xls;	
// /* EXPORT ENDS */

class Couponimport extends \Opencart\System\Engine\Controller {	
	public function index() {
		$this->registry->set('tmd', new  \Tmdcouponimportexport\System\Library\Tmd\System($this->registry));
		$keydata=array(
		'code'=>'tmdkey_couponimport',
		'eid'=>'MjUzODA=',
		'route'=>'extension/tmdcouponimportexport/module/couponimport',
		);
		$couponimport=$this->tmd->getkey($keydata['code']);
		$data['getkeyform']=$this->tmd->loadkeyform($keydata);
		
		$this->load->language('extension/tmdcouponimportexport/module/couponimport');
		$this->document->setTitle($this->language->get('heading_title1'));
		$data['heading_title1'] = $this->language->get('heading_title');
		
		$this->load->model('setting/setting');
		
		if (isset($this->session->data['warning'])) {
			$data['error_warning'] = $this->session->data['warning'];
		
			unset($this->session->data['warning']);
		} else {
			$data['error_warning'] = '';
		}

		$data['breadcrumbs'] = [];

		$data['breadcrumbs'][] = [
			'text' => $this->language->get('text_home'),
			'href' => $this->url->link('common/dashboard', 'user_token=' . $this->session->data['user_token'])
		];

		$data['breadcrumbs'][] = [
			'text' => $this->language->get('text_extension'),
			'href' => $this->url->link('marketplace/extension', 'user_token=' . $this->session->data['user_token'] . '&type=total')
		];

		$data['breadcrumbs'][] = [
			'text' => $this->language->get('heading_title1'),
			'href' => $this->url->link('extension/tmdcouponimportexport/module/couponimport', 'user_token=' . $this->session->data['user_token'])
		];


		if(VERSION>='4.0.2.0'){
		$data['save'] = $this->url->link('extension/tmdcouponimportexport/module/couponimport.save', 'user_token=' . $this->session->data['user_token']);
		}else{
		$data['save'] = $this->url->link('extension/tmdcouponimportexport/module/couponimport|save', 'user_token=' . $this->session->data['user_token']);

		}
		
		$data['back'] = $this->url->link('marketplace/extension', 'user_token=' . $this->session->data['user_token'] . '&type=module');
		
		$this->load->model('localisation/language');
		$data['languages'] = $this->model_localisation_language->getLanguages();

		$data['module_couponimport_status'] = $this->config->get('module_couponimport_status');
		
		$data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');

		$this->response->setOutput($this->load->view('extension/tmdcouponimportexport/module/couponimport', $data));
	}

	public function save(): void {
		$this->load->language('extension/tmdcouponimportexport/module/couponimport');

		$json = [];

		if (!$this->user->hasPermission('modify', 'extension/tmdcouponimportexport/module/couponimport')) {
			$json['error'] = $this->language->get('error_permission');
		}
		
		$couponimport=$this->config->get('tmdkey_couponimport');
		if (empty(trim($couponimport))) {			
		$json['error'] ='Module will Work after add License key!';
		}
	
		if (!$json) {
			$this->load->model('setting/setting');

			$this->model_setting_setting->editSetting('module_couponimport', $this->request->post);

			$json['success'] = $this->language->get('text_success');
		}

		$this->response->addHeader('Content-Type: application/json');
		$this->response->setOutput(json_encode($json));
	}
	public function keysubmit() {
		$json = array(); 
		
      	if (($this->request->server['REQUEST_METHOD'] == 'POST')) {
			$keydata=array(
			'code'=>'tmdkey_couponimport',
			'eid'=>'MjUzODA=',
			'route'=>'extension/tmdcouponimportexport/module/couponimport',
			'moduledata_key'=>$this->request->post['moduledata_key'],
			);
			$this->registry->set('tmd', new  \Tmdcouponimportexport\System\Library\Tmd\System($this->registry));
		
            $json=$this->tmd->matchkey($keydata);       
		} 
		
		$this->response->addHeader('Content-Type: application/json');
		$this->response->setOutput(json_encode($json));
	}

	public function install(){
		$this->load->model('setting/event');
		$this->load->model('user/user_group');

		$this->model_user_user_group->addPermission($this->user->getGroupId(), 'access', 'extension/tmdcouponimportexport/module/couponimport');
		$this->model_user_user_group->addPermission($this->user->getGroupId(), 'modify', 'extension/tmdcouponimportexport/module/couponimport');
		$this->model_user_user_group->addPermission($this->user->getGroupId(), 'access', 'extension/tmdcouponimportexport/tmd/couponimport');
		$this->model_user_user_group->addPermission($this->user->getGroupId(), 'modify', 'extension/tmdcouponimportexport/tmd/couponimport');
			

		// TMD admin coupon events
		$this->model_setting_event->deleteEventByCode('module_tmdimportexportcoupon');

		if(VERSION>='4.0.2.0'){
			$eventaction='extension/tmdcouponimportexport/module/couponimport.tmdimportexportcoupon';
		}else{
			$eventaction='extension/tmdcouponimportexport/module/couponimport|tmdimportexportcoupon';
		}

		$eventrequest=[
			'code'=>'module_tmdimportexportcoupon',
			'description'=>'TMD Import Export coupon',
			'trigger'=>'admin/view/marketing/coupon/before',
			'action'=>$eventaction,
			'status'=>'1',
			'sort_order'=>'1',
		];
				
		if(VERSION=='4.0.0.0')
		{
		$this->model_setting_event->addEvent('module_tmdimportexportcoupon', 'TMD Import Export coupon', 'admin/view/marketing/coupon/before', 'extension/tmdcouponimportexport/module/couponimport|tmdimportexportcoupon', 1);
		}else{
			$this->model_setting_event->addEvent($eventrequest);
		}
	}
			
	public function uninstall(){
		$this->load->model('setting/event');
		$this->load->model('user/user_group');    
		$this->model_setting_event->deleteEventByCode('module_tmdimportexportcoupon');
		$this->model_user_user_group->removePermission($this->user->getGroupId(), 'access', 'extension/tmdcouponimportexport/module/couponimport');
		$this->model_user_user_group->removePermission($this->user->getGroupId(), 'modify', 'extension/tmdcouponimportexport/module/couponimport');
		$this->model_user_user_group->removePermission($this->user->getGroupId(), 'access', 'extension/tmdcouponimportexport/tmd/couponimport');
		$this->model_user_user_group->removePermission($this->user->getGroupId(), 'modify', 'extension/tmdcouponimportexport/tmd/couponimport');
	}

	public function tmdimportexportcoupon(string&$route, array&$args, mixed&$output):void {
		$modulestatus=$this->config->get('module_couponimport_status');
		$this->load->language('extension/tmdcouponimportexport/module/couponimport');
		$this->load->language('marketing/coupon');
		if(!empty($modulestatus)){
			$url='';
			if(VERSION>='4.0.2.0'){
				$args['export'] = $this->url->link('extension/tmdcouponimportexport/module/couponimport.export', 'user_token=' . $this->session->data['user_token'] . $url);
			}else{
				$args['export'] = $this->url->link('extension/tmdcouponimportexport/module/couponimport|export', 'user_token=' . $this->session->data['user_token'] . $url);
			}
			$args['import'] = $this->url->link('extension/tmdcouponimportexport/tmd/couponimport', 'user_token=' . $this->session->data['user_token'] . $url);
			///xml///

			
        	$template_buffer = $this->getTemplateBuffer($route,$output);

			if(VERSION>='4.1.0.0'){
	            $find='<a href="{{ add }}" data-bs-toggle="tooltip" title="{{ button_add }}" class="btn btn-primary"><i class="fa-solid fa-plus"></i></a>';
				$replace='<a href="{{ import }}" data-bs-toggle="tooltip" title="{{ button_importcoupons }}" class="btn btn-primary"><i class="fa fa-upload"></i> {{ button_importcoupons }}</a> 
				<a href="{{ export }}" class="btn btn-primary" data-bs-toggle="tooltip" title="{{ button_exportcoupons }}"><i class="fa fa-download"></i> {{ button_exportcoupons }}</a> <a href="{{ add }}" data-bs-toggle="tooltip" title="{{ button_add }}" class="btn btn-primary"><i class="fa-solid fa-plus"></i></a>';
             }else{
				$find='	<a href="{{ add }}" data-bs-toggle="tooltip" title="{{ button_add }}" class="btn btn-primary"><i class="fa-solid fa-plus"></i></a>';
				$replace='<a href="{{ import }}" data-bs-toggle="tooltip" title="{{ button_importcoupons }}" class="btn btn-primary"><i class="fa fa-upload"></i> {{ button_importcoupons }}</a> 
				<a href="{{ export }}" class="btn btn-primary" data-bs-toggle="tooltip" title="{{ button_exportcoupons }}"><i class="fa fa-download"></i> {{ button_exportcoupons }}</a><a href="{{ add }}" data-bs-toggle="tooltip" title="{{ button_add }}" class="btn btn-primary"><i class="fa-solid fa-plus"></i></a>';
		  }
			$output = str_replace( $find, $replace, $template_buffer );
		}
	}
	
	protected function getTemplateBuffer($route, $event_template_buffer) {

		// if there already is a modified template from view/*/before events use that one
		if ($event_template_buffer) {
		return $event_template_buffer;
		}

		// load the template file (possibly modified by ocmod and vqmod) into a string buffer

		if ($this->config->get('config_theme') == 'default') {
		$theme = $this->config->get('theme_default_directory');
		} else {
		$theme = $this->config->get('config_theme');
		}
		$dir_template = DIR_TEMPLATE;

		$template_file = $dir_template.$route.'.twig';
		if (file_exists($template_file) && is_file($template_file)) {

		return file_get_contents($template_file);
		}
	
		$dir_template  = DIR_TEMPLATE.'default/template/';
		$template_file = $dir_template.$route.'.twig';
		if (file_exists($template_file) && is_file($template_file)) {

		return file_get_contents($template_file);
		}
		trigger_error("Cannot find template file for route '$route'");
		exit;
	}

	public function export() {

		

			$this->load->model('marketing/coupon');
			$this->load->language('marketing/coupon');
			$this->load->language('extension/tmdcouponimportexport/module/couponimport');

			 if (isset($this->request->get['sort'])) {
				$sort = $this->request->get['sort'];
			} else {
				$sort = 'name';
			}

			if (isset($this->request->get['order'])) {
				$order = $this->request->get['order'];
			} else {
				$order = 'ASC';
			}

			$data['coupons'] = [];

			$filter_data = [
				'sort'                     => $sort,
				'order'                    => $order,
			];

			$results = $this->model_marketing_coupon->getCoupons($filter_data);

			foreach ($results as $result) {
				$result= $this->model_marketing_coupon->getCoupon($result['coupon_id']);
				$product_ids='';
				$product_names='';
				$categories_ids='';
				$categories_names='';
				$coupon_id=$result['coupon_id'];
				$datapro=$this->model_marketing_coupon->getProducts($coupon_id);
				
				if(isset($datapro)){
					foreach($datapro as $product){
						if(!empty($product)){
							$sql =$this->db->query("SELECT name FROM " . DB_PREFIX . "product p LEFT JOIN " . DB_PREFIX . "product_description pd ON (p.product_id = pd.product_id) where p.product_id='".$product."'");
							$product_ids .=$product.',';
							$product_names .=$sql->row['name'].',';
						}
					}
				}
				
				$datacat=$this->model_marketing_coupon->getCategories($coupon_id);

				if(isset($datacat)){
					foreach($datacat as $category){
						if(!empty($category)){
							$sql =$this->db->query("SELECT name FROM " . DB_PREFIX . "category c LEFT JOIN " . DB_PREFIX . "category_description cd ON (c.category_id = cd.category_id) where c.category_id='".$category."'");
							$categories_ids .=$category.',';
							$categories_names .=$sql->row['name'].',';
						}
					}
				}

				$data['coupons'][] = [
					'coupon_id'  => $result['coupon_id'],
					'name'       => $result['name'],
					'code'       => $result['code'],
					'discount'   => $result['discount'],
					'date_start'     => $result['date_start'],
					'date_end'     => $result['date_end'],
					'status'         => ($result['status'] ? $this->language->get('text_enabled') : $this->language->get('text_disabled')),
					'type'     => $result['type'],
					'total'     => $result['total'],
					'logged'         => ($result['logged'] ? $this->language->get('text_yes') : $this->language->get('text_no')),
					'shipping'         => ($result['shipping'] ? $this->language->get('text_yes') : $this->language->get('text_no')),
					'uses_total'     => $result['uses_total'],
					'uses_customer'     => $result['uses_customer'],
					'date_added'     => $result['date_added'],
					'product_ids'     => $product_ids,
					'product_names'     => $product_names,
					'categories_ids'     => $categories_ids,
					'categories_names'     => $categories_names,
				];
			}

			$coupons = $this->request->clean($data['coupons']);
			
			$spreadsheet = new Spreadsheet();

			//Columns
			$i=1;
			$spreadsheet->getActiveSheet()->SetCellValue('A'.$i, $this->language->get('column_coupon_id'));
			$spreadsheet->getActiveSheet()->SetCellValue('B'.$i, $this->language->get('column_name'));
			$spreadsheet->getActiveSheet()->SetCellValue('C'.$i, $this->language->get('column_code'));
			$spreadsheet->getActiveSheet()->SetCellValue('D'.$i, $this->language->get('column_type'));
			$spreadsheet->getActiveSheet()->SetCellValue('E'.$i, $this->language->get('column_discount'));
			$spreadsheet->getActiveSheet()->SetCellValue('F'.$i, $this->language->get('column_total'));
			$spreadsheet->getActiveSheet()->SetCellValue('G'.$i, $this->language->get('column_shipping'));
			$spreadsheet->getActiveSheet()->SetCellValue('H'.$i, $this->language->get('column_products_ids'));
			$spreadsheet->getActiveSheet()->SetCellValue('I'.$i, $this->language->get('column_products_names'));
			$spreadsheet->getActiveSheet()->SetCellValue('J'.$i, $this->language->get('column_products_names'));
			$spreadsheet->getActiveSheet()->SetCellValue('K'.$i, $this->language->get('column_categories_ids'));
			$spreadsheet->getActiveSheet()->SetCellValue('L'.$i, $this->language->get('column_categories_names'));
			$spreadsheet->getActiveSheet()->SetCellValue('M'.$i, $this->language->get('column_start_date'));
			$spreadsheet->getActiveSheet()->SetCellValue('N'.$i, $this->language->get('column_end_date'));
			$spreadsheet->getActiveSheet()->SetCellValue('O'.$i, $this->language->get('column_uses_per_couopn'));
			$spreadsheet->getActiveSheet()->SetCellValue('P'.$i, $this->language->get('column_uses_per_customer'));
			$spreadsheet->getActiveSheet()->SetCellValue('Q'.$i, $this->language->get('column_status'));
			$i=2;

			foreach($coupons as $coupon) {					
				$spreadsheet->getActiveSheet()->SetCellValue('A'.$i, $coupon['coupon_id']);
				$spreadsheet->getActiveSheet()->SetCellValue('B'.$i, $coupon['name']);
				$spreadsheet->getActiveSheet()->SetCellValue('C'.$i, $coupon['code']);
				$spreadsheet->getActiveSheet()->SetCellValue('D'.$i, $coupon['type']);
				$spreadsheet->getActiveSheet()->SetCellValue('E'.$i, $coupon['discount']);
				$spreadsheet->getActiveSheet()->SetCellValue('F'.$i, $coupon['total']);
				$spreadsheet->getActiveSheet()->SetCellValue('G'.$i, $coupon['logged']);
				$spreadsheet->getActiveSheet()->SetCellValue('H'.$i, $coupon['shipping']);
				$spreadsheet->getActiveSheet()->SetCellValue('I'.$i, $coupon['product_ids']);
				$spreadsheet->getActiveSheet()->SetCellValue('J'.$i, $coupon['product_names']);
				$spreadsheet->getActiveSheet()->SetCellValue('K'.$i, $coupon['categories_ids']);
				$spreadsheet->getActiveSheet()->SetCellValue('L'.$i, $coupon['categories_names']);
				$spreadsheet->getActiveSheet()->SetCellValue('M'.$i, $coupon['date_start']);
				$spreadsheet->getActiveSheet()->SetCellValue('N'.$i, $coupon['date_end']);
				$spreadsheet->getActiveSheet()->SetCellValue('O'.$i, $coupon['uses_total']);
				$spreadsheet->getActiveSheet()->SetCellValue('P'.$i, $coupon['uses_customer']);
				$spreadsheet->getActiveSheet()->SetCellValue('Q'.$i, $coupon['status']);
				$i++;
			}
				
		$filename = 'coupon_export.xls';
		$spreadsheet->getActiveSheet()->setTitle('All Coupons');
		$writer =new \PhpOffice\PhpSpreadsheet\Writer\Xls($spreadsheet);
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="'. urlencode($filename).'"');
		$writer->save('php://output');
	}
}
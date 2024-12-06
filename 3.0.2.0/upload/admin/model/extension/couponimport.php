<?php
/**
 * TMD(http://opencartextensions.in/)
 *
 * Copyright (c) 2016 - 2017 TMD
 * This package is Copyright so please us only one domain 
 * 
 */
class ModelExtensionCouponImport extends Model {
	
	public function addCoupon($data) {
		if($data['logged'] == 'Yes'){
          $logged = '1';
		
	    }else{
            $logged = '0';
	    }
		if($data['shipping'] == 'Yes'){
		  $shipping = '1';
		}else{
		    $shipping = '0';
		}
		if(!empty($data['date_added'])){
			$date_added = $data['date_added'];
		}else{
		  $date_added 	='NOW()';
		}

  	$this->db->query("INSERT INTO " . DB_PREFIX . "coupon SET coupon_id= '". $data['coupon_id'] . "', name = '" . $this->db->escape($data['name']) . "', code = '" . $this->db->escape($data['code']) . "', discount = '" . (float)$data['discount'] . "', type = '" . $this->db->escape($data['type']) . "', total = '" . (float)$data['total'] . "', logged = '" . (int)$logged . "', shipping = '" . (int)$shipping . "', date_start = '" . $this->db->escape($data['date_start']) . "', date_end = '" . $this->db->escape($data['date_end']) . "', uses_total = '" . (int)$data['uses_total'] . "', uses_customer = '" . (int)$data['uses_customer'] . "', status = '" . (int)$data['status'] . "', date_added = '". $date_added ."'");

		$coupon_id = $this->db->getLastId();

		if (isset($data['coupon_product'])) {
			foreach ($data['coupon_product'] as $product_id) {
				$this->db->query("INSERT INTO " . DB_PREFIX . "coupon_product SET coupon_id = '" . (int)$coupon_id . "', product_id = '" . (int)$product_id . "'");
			}
		}

		if (isset($data['coupon_category'])) {
			foreach ($data['coupon_category'] as $category_id) {
				$this->db->query("INSERT INTO " . DB_PREFIX . "coupon_category SET coupon_id = '" . (int)$coupon_id . "', category_id = '" . (int)$category_id . "'");
			}
		}		
	}

	public function editCoupon($data,$coupon_id) {
		if($data['logged'] == 'Yes'){
            $logged = '1';
		}else{
            $logged = '0';
	    }
		if($data['shipping'] == 'Yes'){
		  $shipping = '1';
		}else{
		    $shipping = '0';
		}
		
	$this->db->query("UPDATE " . DB_PREFIX . "coupon SET name = '" . $this->db->escape($data['name']) . "', code = '" . $this->db->escape($data['code']) . "', discount = '" . (float)$data['discount'] . "', type = '" . $this->db->escape($data['type']) . "', total = '" . (float)$data['total'] . "', logged = '" . (int)$logged . "', shipping = '" . (int)$shipping . "', date_start = '" . $this->db->escape($data['date_start']) . "', date_end = '" . $this->db->escape($data['date_end']) . "', uses_total = '" . (int)$data['uses_total'] . "', uses_customer = '" . (int)$data['uses_customer'] . "', status = '" . (int)$data['status'] . "' WHERE coupon_id = '" . (int)$coupon_id . "'");



		$this->db->query("DELETE FROM " . DB_PREFIX . "coupon_product WHERE coupon_id = '" . (int)$coupon_id . "'");

		if (isset($data['coupon_product'])) {
			foreach ($data['coupon_product'] as $product_id) {
				$this->db->query("INSERT INTO " . DB_PREFIX . "coupon_product SET coupon_id = '" . (int)$coupon_id . "', product_id = '" . (int)$product_id . "'");
			}
		}	

		$this->db->query("DELETE FROM " . DB_PREFIX . "coupon_category WHERE coupon_id = '" . (int)$coupon_id . "'");

		if (isset($data['coupon_category'])) {
			foreach ($data['coupon_category'] as $category_id) {
				$this->db->query("INSERT INTO " . DB_PREFIX . "coupon_category SET coupon_id = '" . (int)$coupon_id . "', category_id = '" . (int)$category_id . "'");
			}
		}				
	}

	
	public function getCouponByCode($code) {
		$query = $this->db->query("SELECT * FROM " . DB_PREFIX . "coupon WHERE code = '" . $this->db->escape($code) . "'");
		if(isset($query->row['coupon_id'])){
		return $query->row['coupon_id'];
		}
	}
	
		
	
}
?>
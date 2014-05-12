<?php
class Integra_xls {
	private static $tbl_blocks='tmp_catalog_blocks';
	private static $tbl_blocks_cont='tmp_catalog_blocks_cont';
	private static $tbl_photos='tmp_catalog_photos';
	private static $tbl_photos_cont='tmp_catalog_photos_cont';
	private static $updated=0;
	private static $added=0;
	private static $i=1;
	private static $AllSettings;
	private static $exist_img=array();
	private static $file_stat=array(
					'filename'=>'',
					'filerow'=>'',
					'filehash'=>'',
					'datetime'=>'',
					'added'=>'',
					'updated'=>'',
				);
	private static function createTmptables(){
		
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_sprav_citys`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_sprav_orientation`"); 
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_sprav_residence`"); 
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_sprav_type_estate`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_sprav_type_residence`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_catalog_blocks`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_catalog_blocks_cont`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_catalog_index`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_catalog_index_cont`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_catalog_photos`");
		Page::$DB->query("DROP TABLE IF EXISTS `tmp_catalog_photos_cont`");
		
		Page::$DB->query("CREATE TABLE `tmp_sprav_citys` LIKE `sprav_citys`;");
		Page::$DB->query("CREATE TABLE `tmp_sprav_orientation` LIKE `sprav_orientation`;");
		Page::$DB->query("CREATE TABLE `tmp_sprav_residence` LIKE `sprav_residence`;");
		Page::$DB->query("CREATE TABLE `tmp_sprav_type_estate` LIKE `sprav_type_estate`;");
		Page::$DB->query("CREATE TABLE `tmp_sprav_type_residence` LIKE `sprav_type_residence`;");
		Page::$DB->query("CREATE TABLE `tmp_catalog_blocks` LIKE `catalog_blocks`;");
		Page::$DB->query("CREATE TABLE `tmp_catalog_blocks_cont` LIKE `catalog_blocks_cont`;");
		Page::$DB->query("CREATE TABLE `tmp_catalog_index` LIKE `catalog_index`;");
		Page::$DB->query("CREATE TABLE `tmp_catalog_index_cont` LIKE `catalog_index_cont`;");
		Page::$DB->query("CREATE TABLE `tmp_catalog_photos` LIKE `catalog_photos`;");
		Page::$DB->query("CREATE TABLE `tmp_catalog_photos_cont` LIKE `catalog_photos_cont`;");
		
		Page::$DB->query("INSERT INTO `tmp_sprav_citys` (SELECT * FROM `sprav_citys`)");
		Page::$DB->query("INSERT INTO `tmp_sprav_orientation`(SELECT * FROM `sprav_orientation`)");
		Page::$DB->query("INSERT INTO `tmp_sprav_residence` (SELECT * FROM `sprav_residence`)");
		Page::$DB->query("INSERT INTO `tmp_sprav_type_estate` (SELECT * FROM `sprav_type_estate`)");
		Page::$DB->query("INSERT INTO `tmp_sprav_type_residence` (SELECT * FROM `sprav_type_residence`)");
		
		Page::$DB->query("INSERT INTO `tmp_catalog_blocks` (SELECT * FROM `catalog_blocks`)");
		Page::$DB->query("UPDATE `tmp_catalog_blocks` SET `block_show` = '0';");
		Page::$DB->query("INSERT INTO `tmp_catalog_blocks_cont` (SELECT * FROM `catalog_blocks_cont`)");
		
		Page::$DB->query("INSERT INTO `tmp_catalog_index` (SELECT * FROM `catalog_index`)");
		Page::$DB->query("INSERT INTO `tmp_catalog_index_cont` (SELECT * FROM `catalog_index_cont`)");
		
		Page::$DB->query("INSERT INTO `tmp_catalog_photos` (SELECT * FROM `catalog_photos`)");
		Page::$DB->query("INSERT INTO `tmp_catalog_photos_cont` (SELECT * FROM `catalog_photos_cont`)");
	}
	private static function finishImport(){
		Page::$DB->query("DROP TABLE IF EXISTS `old_sprav_citys`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_sprav_orientation`"); 
		Page::$DB->query("DROP TABLE IF EXISTS `old_sprav_residence`"); 
		Page::$DB->query("DROP TABLE IF EXISTS `old_sprav_type_estate`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_sprav_type_residence`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_catalog_blocks`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_catalog_blocks_cont`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_catalog_index`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_catalog_index_cont`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_catalog_photos`");
		Page::$DB->query("DROP TABLE IF EXISTS `old_catalog_photos_cont`");
		
		Page::$DB->query("RENAME TABLE `sprav_citys` TO `old_sprav_citys`, `tmp_sprav_citys` TO `sprav_citys`;");
		Page::$DB->query("RENAME TABLE `sprav_orientation` TO `old_sprav_orientation`, `tmp_sprav_orientation` TO `sprav_orientation`;");
		Page::$DB->query("RENAME TABLE `sprav_residence` TO `old_sprav_residence`, `tmp_sprav_residence` TO `sprav_residence`;");
		Page::$DB->query("RENAME TABLE `sprav_type_estate` TO `old_sprav_type_estate`, `tmp_sprav_type_estate` TO `sprav_type_estate`;");
		Page::$DB->query("RENAME TABLE `sprav_type_residence` TO `old_sprav_type_residence`, `tmp_sprav_type_residence` TO `sprav_type_residence`;");
		
		Page::$DB->query("RENAME TABLE `catalog_blocks` TO `old_catalog_blocks`, `tmp_catalog_blocks` TO `catalog_blocks`;");
		Page::$DB->query("RENAME TABLE `catalog_blocks_cont` TO `old_catalog_blocks_cont`, `tmp_catalog_blocks_cont` TO `catalog_blocks_cont`;");
		
		Page::$DB->query("RENAME TABLE `catalog_index` TO `old_catalog_index`, `tmp_catalog_index` TO `catalog_index`;");
		Page::$DB->query("RENAME TABLE `catalog_index_cont` TO `old_catalog_index_cont`, `tmp_catalog_index_cont` TO `catalog_index_cont`;");
		
		Page::$DB->query("RENAME TABLE `catalog_photos` TO `old_catalog_photos`, `tmp_catalog_photos` TO `catalog_photos`;");
		Page::$DB->query("RENAME TABLE `catalog_photos_cont` TO `old_catalog_photos_cont`, `tmp_catalog_photos_cont` TO `catalog_photos_cont`;");
    }
	
	private static function get_id ($val,$table) {
		if ((isset($val)&&!empty($val))&&(isset($table)&&!empty($table))){
			$sql = "SELECT `id`
					FROM `tmp_sprav_".$table."`
					WHERE `value` = '".Page::$DB->escape($val)."';";
			$res = Page::$DB->query($sql);
			$num = $res->num_rows();
				if ($num == 0){
						$sql= "INSERT INTO `tmp_sprav_".$table."`
							  SET `value` = '".Page::$DB->escape($val)."';";
							  
						Page::$DB->query($sql);
						$id = Page::$DB->insert_id();
						if ($table=='residence'){
							self::makeResidents($id,$val);
						}
					}else{
						$id=$res->get_one(0, 0);
				}
				return $id;
			}else{
				return false;
		}
	}
	
	private static function preparePrice($val){
		$val=preg_replace('#[^0-9]+#','',$val);
		$val= floatval($val);
		return $val;
	}
	
	public static function saveLog($val){
		
		require_once USES_FOLDER.'/core/class_Settings_model.php';
		Settings_model::singleton();
		require_once USES_FOLDER.'/core/class_Settings.php';
		Settings::singleton();
		
		self::$AllSettings = Settings_model::getAllSettings();

		if (is_array($val) && !empty($val)){
			
				$body= 'Уважаемый администратор, <br />'.strftime('%R %d.%m.%Y', time()).' поступил Лог разбора файла: <br />
					Имя файла: '.$val['filename'].'<br />
					Кол-во разобранных строк: '.$val['filerow'].'<br />
					Уникальный номер: '.$val['filehash'].'<br />
					Дата разбора файла: '.$val['datetime'].'<br />
					Кол-во добавленных строк: '.$val['added'].'<br />
					Кол-во измененных строк: '.$val['updated'].'<br />
					С уважением, Ваш сайт.
					';
					Mailsend::factory()
					->setFromEmail(self::$AllSettings['Main']['admin_mail'][1]['s_val'])
					->setFromName(self::$AllSettings['Main']['admin_name'][1]['s_val'])
					->setSubject('Протокол разбора файлов')
					->setBody($body)
					->setTo(self::$AllSettings['Main']['email_log'][1]['s_val'])
					->send();
					$sql= "INSERT INTO `files_log`
							  SET `log_status` = '1',
								  `log_filename` = '".Page::$DB->escape($val['filename'])."',
								  `log_filehash` = '".Page::$DB->escape($val['filehash'])."',
								  `log_added` = '".Page::$DB->escape($val['added'])."',
								  `log_updated` = '".Page::$DB->escape($val['updated'])."',
								  `log_filerow` = '".Page::$DB->escape($val['filerow'])."',
								  `log_datetime` = NOW()
							  ;";
								  
				Page::$DB->query($sql);
		}else{	
				$sql= "INSERT INTO `files_log`
							  SET `log_status` = '0',
								  `log_status_desc` = '".Page::$DB->escape($val)."',
								  `log_datetime` = NOW()
							  ;";
								  
				Page::$DB->query($sql);
				
				$body= 'Уважаемый администратор, <br />'.strftime('%R %d.%m.%Y', time()).' поступила Лог разбора файла<br />
					При разборе файлов возникли ошибки: <br />
					<b>'.$val.'</b><br />
					С уважением, Ваш сайт.
					';
					Mailsend::factory()
					->setFromEmail(self::$AllSettings['Main']['admin_mail'][1]['s_val'])
					->setFromName(self::$AllSettings['Main']['admin_name'][1]['s_val'])
					->setSubject('Протокол разбора файлов')
					->setBody($body)
					->setTo(self::$AllSettings['Main']['email_log'][1]['s_val'])
					->send();
		}
	}
	
	private static function makeResidents($id_pos,$name){
		$str=iconv("UTF-8", "CP1251//TRANSLIT//IGNORE", $name);
		$str=iconv("CP1251", "UTF-8", $str);
		$str=preg_replace('#[^a-z0-9.-_]#ui', ' ', $str);
		$str=preg_replace('#\\s+#u', '-', $str);
		
		$str=UTF8::strtolower($str);
		$data = array();
			
			$data['index_level'] = 2;
			$data['ci_url'] = $str;
			$data['index_show'] = 1;
			$data['index_order'] = $id_pos;
			$data['index_is_show'] = '1';
			$data['index_is_edit'] = '1';
			$data['index_is_del'] = '1';
			$data['index_is_add_childs'] = '0';
			$data['index_is_add_blocks'] = '0';
			$data['sitemap_changefreq'] = 'daily';
			$data['sitemap_priority'] = '0.5';
			$data['noindex'] = '0';
			$data['target_blank'] = '0';
			$data['comments_allowed'] = '0';
			$data['comm_count'] = '0';
			$data['cic_res_id'] = $id_pos;
			
		$last_id =  Menu::insertNode('tmp_Catalog', $data, 8);
		
		for ($lang = 1; $lang <= 3; $lang++) {
		$sql_ins="INSERT INTO `tmp_catalog_index_cont` (
								`ci_id`,
								`lang_id`,
								`cic_name`,
								`cic_page_name`,
								`created`,
								`uu_fio`
							) VALUES (
								".$last_id.",
								".$lang.",
								'".Page::$DB->escape($name)."',
								'".Page::$DB->escape($name)."',
								NOW(),
								'auto_created'
							);";
							
		Page::$DB->query($sql_ins);
		
		}
		
	}
	private static function saveValue($val){
		
		if (isset($val)&&is_array($val)){
			
			
			$sql = "SELECT `cb_id`
					FROM `".self::$tbl_blocks."`
					WHERE `obj_cod` = '".Page::$DB->escape($val[0])."' 
					AND `auto_man` = '0' LIMIT 1;";
			
			$res = Page::$DB->query($sql);
			$num = $res->num_rows();
				if ($num > 0){
						$id=$res->get_one(0, 0);
						$sql= "UPDATE `".self::$tbl_blocks."`
							   SET
							   `ci_id` = 7,
							   `auto_man` = '0',
							   `block_show` = '1',
							   `city` = '".Page::$DB->escape($val[1])."',
							   `residence` = '".Page::$DB->escape($val[2])."',
							   `type_residence` = '".Page::$DB->escape($val[3])."',
							   `type_estate` = '".Page::$DB->escape($val[4])."',
							   `num_rooms` = '".Page::$DB->escape($val[5])."',
							   `living_space` = '".Page::$DB->escape($val[6])."',
							   `terrace` = '".Page::$DB->escape($val[7])."',
							   `footage` = '".Page::$DB->escape($val[8])."',
							   `floor` = '".Page::$DB->escape($val[9])."',
							   `num_storeys` = '".Page::$DB->escape($val[10])."',
							   `orientation` = '".Page::$DB->escape($val[11])."',
							   `cb_price1` = '".Page::$DB->escape(self::preparePrice($val[12]))."',
							   `quarter` = ".(int)Page::$DB->escape($val[13]).",
							   `year` = ".(int)Page::$DB->escape($val[14])."
								WHERE `cb_id` = ".$id."
							   ;";
						
						Page::$DB->query($sql);
						self::$updated++;
					}else{
						$sql= "INSERT INTO `".self::$tbl_blocks."`
							   SET
							   `ci_id` = 7,
							   `auto_man` = '0',
							   `block_order` = 1,
							   `block_show` = '1',
							   `block_is_show` = '1',
							   `block_is_edit` = '1',
							   `block_is_del` = '1',
							   `block_img` = '',
							   `obj_cod` = '".Page::$DB->escape($val[0])."',
							   `city` = '".Page::$DB->escape($val[1])."',
							   `residence` = '".Page::$DB->escape($val[2])."',
							   `type_residence` = '".Page::$DB->escape($val[3])."',
							   `type_estate` = '".Page::$DB->escape($val[4])."',
							   `num_rooms` = '".Page::$DB->escape($val[5])."',
							   `living_space` = '".Page::$DB->escape($val[6])."',
							   `terrace` = '".Page::$DB->escape($val[7])."',
							   `footage` = '".Page::$DB->escape($val[8])."',
							   `floor` = '".Page::$DB->escape($val[9])."',
							   `num_storeys` = '".Page::$DB->escape($val[10])."',
							   `orientation` = '".Page::$DB->escape($val[11])."',
							   `cb_price1` = '".Page::$DB->escape(self::preparePrice($val[12]))."',
							   `quarter` = ".(int)Page::$DB->escape($val[13]).",
							   `year` = ".(int)Page::$DB->escape($val[14])."
							   ;";
						Page::$DB->query($sql);
						$id = Page::$DB->insert_id();
						
						$sql= "INSERT INTO `".self::$tbl_blocks_cont."`
							   SET
							   `cb_id` = ".$id.",
							   `lang_id` = 1,
							   `created` = NOW(),
							   `modified` = NOW(),
							   `uu_fio` = 'auto_uploaded'
							 ;";
						Page::$DB->query($sql);
						self::$added++;
				}
				
				$image=self::getExistPhotos($val[0]);
				
				if ($image!=false){
						$sql = "SELECT `cb_id`
								FROM `".self::$tbl_blocks."`
								WHERE `obj_cod` = '".$val[0]."' 
								";
						$res = Page::$DB->query($sql);
						$num = $res->num_rows();
						if ($num > 0){
							$p_id=$res->get_one(0, 0);
							self::$exist_img[$p_id]=$image;
						}
					}
			}else{
				return false;
		}
		
	}


	private static function removeFile($filepath,$file){
		$new_dirname=DOC_ROOT.$filepath.'produced/'.date("Ymd").'/';
		if (!file_exists($new_dirname)) mkdir($new_dirname);
		rename(DOC_ROOT.$filepath.$file,$new_dirname.$file);
	}
	
	public static function scanExcelFile($filepath,$file){
		if (is_file(DOC_ROOT . $filepath.$file)){
			$objPHPExcelReader = PHPExcel_IOFactory::createReader('Excel5');
			$objPHPExcel = $objPHPExcelReader->load(DOC_ROOT.$filepath.$file);
			$worksheet = $objPHPExcel->getActiveSheet();
			self::createTmptables();
			foreach ($worksheet->toArray() as $file_row => $cells){
				$reading_row = array();
				foreach ($cells as $file_cell => $value){
					$reading_row[]=$value;
				}
				
				
				if (count($reading_row)<=15){
					$reading_row[1]=(int)self::get_id($reading_row[1],'citys');
					$reading_row[2]=(int)self::get_id($reading_row[2],'residence');
					$reading_row[3]=(int)self::get_id($reading_row[3],'type_residence');
					$reading_row[4]=(int)self::get_id($reading_row[4],'type_estate');
					$reading_row[11]=(int)self::get_id($reading_row[11],'orientation');
					if (empty($reading_row[0])||$reading_row[0]==''){
						Integra_xls::saveLog('Пустое значение кода объекта в файле '.$file.' строка '.self::$i);
						return false;
					}else{
						
						self::saveValue($reading_row);
						self::$i++;
					}
				}else{
					
					//Integra_xls::saveLog('Неверное количество столбцов в файле для разбора '.$file);
					return false;
				}
			}
				$file_size=filesize(DOC_ROOT.$filepath.$file);
				self::$file_stat['filename']=$file;
				self::$file_stat['filehash']=md5($file.$file_size);
				self::$file_stat['datetime']=date("Y-m-d H:i:s",time());
				self::$file_stat['added']=self::$added;
				self::$file_stat['updated']=self::$updated;
				self::$file_stat['filerow']=self::$i;
				
				self::makePhotos();
				self::finishImport();
				self::saveLog(self::$file_stat);
				self::removeFile($filepath,$file);
				
				
					$worksheet->disconnectCells();
					unset($worksheet);
					$objPHPExcel->disconnectWorksheets();
					unset($objPHPExcel);
					unset($objPHPExcelReader);
			return true;
		}
		return false;
	}
	
	private static function getExistPhotos($name){
		$img_path = DOC_ROOT.'public/produce/image/';
		$img_array=array();
		$w = 0;
		for ($i = 0; $i <= 50; $i++) {
			if (is_file($img_path.$name.'_'.$i.'.jpg')){
						$img_array[$i]=$name.'_'.$i.'.jpg';
						$w = 0;
						continue;
				}elseif(is_file($img_path.$name.'_'.$i.'.gif')){
						$img_array[$i]=$name.'_'.$i.'.gif';
						$w = 0;
						continue;
					}elseif(is_file($img_path.$name.'_'.$i.'.png')){
						$img_array[$i]=$name.'_'.$i.'.png';
						$w = 0;
						continue;
						}elseif(is_file(UTF8::strtolower($img_path.$name.'_'.$i.'.jpg'))){
							$img_array[$i]=UTF8::strtolower($name.'_'.$i.'.jpg');
							$w = 0;
							continue;
							}elseif(is_file(UTF8::strtolower($img_path.$name.'_'.$i.'.gif'))){
								$img_array[$i]=UTF8::strtolower($name.'_'.$i.'.gif');
								$w = 0;
								continue;
							}elseif(is_file(UTF8::strtolower($img_path.$name.'_'.$i.'.png'))){
								$img_array[$i]=UTF8::strtolower($name.'_'.$i.'.png');
								$w = 0;
								continue;
						} else {
							$w++;
						}
			if ($w >= 4){
				break;
			}
		}
		if (!empty ($img_array)){
			return $img_array;
		}else{
			return false;
		}
		
		
		
	}
	private static function savePhoto($block,$name,$header) {
		
			$sql= "INSERT INTO `".self::$tbl_photos."`
								   SET
								   `cb_id` = ".$block.",
								   `photo_file` = '".$name."',
								   `photo_preview` = 'small_".$name."',
								   `photo_show` = '1',
								   `photo_order` = 1,
								   `photo_is` = 'photo'
								 ;";
			Page::$DB->query($sql);
			
			$id = Page::$DB->insert_id();
			
			$sql= "INSERT INTO `".self::$tbl_photos_cont."`
								   SET
								   `cp_id` = ".$id.",
								   `lang_id` = 1,
								   `cpc_header` = '".$header."',
								   `created` = NOW(),
								   `modified` = NOW(),
								   `uu_fio` = 'auto_upload'
								 ;";
			
			Page::$DB->query($sql);
	}
	
	private static function makePhotos() {
		foreach (self::$exist_img as $block=>$images){
				foreach ($images as $id=>$img){
					$ext=Gen::getFileExtension($img);
					
					Images::Resize_SaveAs(
								DOC_ROOT . 'public/produce/image/',
								$img,
								DOC_ROOT . 'public/content/catalog/',
								'64_' .  $block.'_'.$id.'.'.$ext,
								NULL,
								64,
								64,
								'wh'
							);
							Images::Resize_SaveAs(
								DOC_ROOT . 'public/produce/image/',
								$img,
								DOC_ROOT . 'public/content/catalog/',
								'small_' .  $block.'_'.$id.'.'.$ext,
								NULL,
								238,
								165,
								'w'
							);
							Images::Resize_SaveAs(
								DOC_ROOT . 'public/produce/image/',
								$img,
								DOC_ROOT . 'public/content/catalog/',
								'middle_' .  $block.'_'.$id.'.'.$ext,
								NULL,
								346,
								240,
								'w'
							);
							Images::Resize_SaveAs(
								DOC_ROOT . 'public/produce/image/',
								$img,
								DOC_ROOT . 'public/content/catalog/',
								'big_' .  $block.'_'.$id.'.'.$ext,
								NULL,
								800,
								0,
								'w'
							);
						self::savePhoto($block,$block.'_'.$id.'.'.$ext,$img);
						self::removeFile('public/produce/image/',$img);
					}
			}
	}
	
	
	
	
}

?>
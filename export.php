<?php
	$dir=dirname(__FILE__);//查找当前脚本所在路径
	require $dir."/PHPExcel/PHPExcel.php";//引入PHPExcel
	require $dir."/db.php";//引入mysql操作类文件
	
	$db = new db($phpexcel);//实例化db类
	$objPHPExcel=new PHPExcel();//实例化phpexcel类   等同于创建一个excel
	for ($i=1; $i <8 ; $i++) { 
		if($i>1){
			$objPHPExcel->createSheet();
		}
		$objPHPExcel->setActiveSheetIndex($i-1);
		$objSheet=$objPHPExcel->getActiveSheet();//获取当前活动sheet
		
		//判断第几天
                if($i == '1'){
                    $weeknum = '一'; 
                }
                if($i == '2'){
                    $weeknum = '二'; 
                }
                if($i == '3'){
                    $weeknum = '三'; 
                }
                if($i == '4'){
                    $weeknum = '四'; 
                }
                if($i == '5'){
                    $weeknum = '五'; 
                }
                if($i == '6'){
                    $weeknum = '六'; 
                }
                if($i == '7'){
                    $weeknum = '日'; 
                }
        $objSheet->setTitle("周".$weeknum);
		$data=$db->getDataByWeek($weeknum);
		//print_r($data);die;

//输出需求表
		$objSheet->setCellValue("A1","教室号")->setCellValue("B1","一二节")->setCellValue("C1","三四节")->setCellValue("D1","七八节")->setCellValue("E1","九十节")->setCellValue("F1","十一~十三节");
		$j=2;
		$jiaoshi=$db->getDataByROOM($ROOM);
		foreach ($jiaoshi as $key => $val) {
			
			$objSheet->setCellValue("A".$j,$val['ROOM']);
				
			$j++;
		}
 	
		// foreach ($data as $key => $val) {
			
		// 	for ($i=0;; $i++) { 
		// 		if($data[$i][ROOM]==)
				
		// 		print_r($data[$i][ROOM]);echo '<br>';
		// 	}
			//die;
			//$objSheet->setCellValue("A".$j,$val['ID']);
				
			
		//}

	}
	
	$objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');//生成excel文件
	$objWriter->save($dir."/export_1.xls");//保存文件
	// $date = date("Y_m_d",time());
	// $fileName = "{$date}.xlsx";

?>
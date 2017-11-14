<?php
	$dir=dirname(__FILE__);//查找当前脚本所在路径
	require $dir."/PHPExcel/PHPExcel.php";//引入PHPExcel
	require $dir."/db.php";//引入mysql操作类文件
 
	require_once 'PHPExcel\PHPExcel\IOFactory.php';
	require_once 'PHPExcel\PHPExcel\Reader\Excel5.php'; 
	require_once 'PHPExcel\PHPExcel\Reader\Excel2007.php'; 

	$servername = "localhost";
	$username ="root";
	$password = "";
	$dbname="kebiao";
	
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
		//print_r($data);
		
		//输出需求表
		$objSheet->setCellValue("A1","教室号")->setCellValue("B1","一二节")->setCellValue("C1","三四节")->setCellValue("D1","七八节")->setCellValue("E1","九十节")->setCellValue("F1","十一~十三节");
		$j=2;
		$jiaoshi=$db->getDataByROOM($ROOM);

		foreach ($jiaoshi as $key => $val) {
			$objSheet->setCellValue("A".$j,$val['ROOM']);//填充第一列教室号
			$jshvalue[]=$val['ROOM'];//将第一列教室号放进数组	
			$j++;
		}	
		//print_r($jshvalue);

		foreach ($data as $key => $val) {
			for ($h=0;$h<150; $h++) { //
				for ($k=0; $k <50 ; $k++) { 
					if($data[$h][ROOM]==$jshvalue[$k]){//判断行
						if ($data[$h][ROOM]=="1201") {
							$m=2;
						}else if ($data[$h][ROOM]=="1202") {
							$m=3;
						}else if ($data[$h][ROOM]=="1203") {
							$m=4;
						}else if ($data[$h][ROOM]=="1301") {
							$m=5;
						}else if ($data[$h][ROOM]=="1302") {
							$m=6;
						}else if ($data[$h][ROOM]=="1303") {
							$m=7;
						}
						
						if($data[$h][JIE]=="1-2"){
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("B".$m,$data[$h][XINGMING]);
						}else if ($data[$h][JIE]=="1-3") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("B".$m,$data[$h][XINGMING].'3');
						}else if ($data[$h][JIE]=="1-4") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("B".$m,$data[$h][XINGMING]);
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("C".$m,$data[$h][XINGMING]);
						}else if ($data[$h][JIE]=="7-10") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("D".$m,$data[$h][XINGMING]);
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("E".$m,$data[$h][XINGMING]);
						}else if ($data[$h][JIE]=="3-4") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("C".$m,$data[$h][XINGMING]);
						}else if ($data[$h][JIE]=="7-8") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("D".$m,$data[$h][XINGMING]);
						}else if ($data[$h][JIE]=="7-9") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("D".$m,$data[$h][XINGMING].'3');
						}else if ($data[$h][JIE]=="9-10") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("E".$m,$data[$h][XINGMING]);
						}else if ($data[$h][JIE]=="11-12") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("F".$m,$data[$h][XINGMING]);
						}else if ($data[$h][JIE]=="11-13") {
							$objPHPExcel->setActiveSheetIndex($i-1)->setCellValue("F".$m,$data[$h][XINGMING].'3');
						}
						
					}

				}
			}
			
		}

	}

		//删除MySQL中kebiao数据库
		$conn = mysqli_connect($servername,$username,$password,$dbname);
		mysqli_set_charset($conn,'utf8_general_ci');//不写这句话的话会造成乱码
			if(!$conn){
				die("连接失败：".mysqli_connect_error());
			}
		
			$sql = "DROP DATABASE kebiao";
			if(!(mysqli_query($conn,$sql))){
				echo "数据表删除失败:" . mysqli_error($conn);
			}
		
// 设置活动单指数到第一个表,所以Excel打开这是第一个表
$objPHPExcel->setActiveSheetIndex(0);		


	// //生成excel文件
	// $objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');
	// // //保存文件
	// $objWriter->save($dir."/export_1.xls");
	
$fileName = "123.xlsx";

// 将输出重定向到一个客户端web浏览器(Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename='.$fileName);
header('Cache-Control: max-age=0');

//要是输出为Excel2007,使用 Excel2007对应的类，生成的文件名为.xlsx.如果是Excel2005,使用Excel5,对应生成.xls文件
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
// $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

//支持浏览器下载生成的文档
$objWriter->save('php://output');

?>
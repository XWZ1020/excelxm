<?php
/**
 *文件编码需为UTF-8，否则会存在生成的文档内容乱码
 */

 /** 引入需要的类库*/
require_once 'PHPExcel\PHPExcel.php'; 
require_once 'PHPExcel\PHPExcel\IOFactory.php';
require_once 'PHPExcel\PHPExcel\Reader\Excel5.php'; 
require_once 'PHPExcel\PHPExcel\Reader\Excel2007.php'; 

$objPHPExcel = new PHPExcel();

//设置生成的Excel文件名
$date = date("Y_m_d",time());
$fileName = "{$date}.xlsx";

//数据库连接导入数据库里面
 $link = mysqli_connect('localhost','root','','kebiao');
 mysqli_set_charset($link,"utf8");//不写这句话的话会造成乱码
 if(!$link){
	 echo "数据库连接失败"; 
	 exit;
 }
 $sql = "SELECT * from kecheng";
 $conn= mysqli_query($link,$sql);
 //$row = mysqli_fetch_array($wp);
 while($row = mysqli_fetch_assoc($conn)) {
        $arr[]=$row;;
    }
// var_dump($arr);die;

$objPHPExcel->setActiveSheetIndex(0)->setCellValue("A1","编号")->setCellValue("B1","姓名")->setCellValue("C1","教室")->setCellValue("D1","周次")->setCellValue("E1","节次");
//适合把表中数据导入Excel文件中，多数据循环设置值			
foreach($arr as $key=> $value) {
	$key+=2;
	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('A'.$key,$value['ID'])
	            ->setCellValue('B'.$key,$value['XINGMING'])
	            ->setCellValue('C'.$key,$value['ROOM'])
	            ->setCellValue('D'.$key,$value['WEEK'])
	            ->setCellValue('E'.$key,$value['CLASS']);;
}

// 重命名表
// $objPHPExcel->getActiveSheet()->setTitle('Simple');

// 设置活动单指数到第一个表,所以Excel打开这是第一个表
$objPHPExcel->setActiveSheetIndex(0);

// 将输出重定向到一个客户端web浏览器(Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename='.$fileName);
header('Cache-Control: max-age=0');

//要是输出为Excel2007,使用 Excel2007对应的类，生成的文件名为.xlsx.如果是Excel2005,使用Excel5,对应生成.xls文件
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
// $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

//支持浏览器下载生成的文档
$objWriter->save('php://output');

//支持保存生成的文件在当前目录下,直接文件名做为参数
// $objWriter->save('test.xlsx');
   
?>
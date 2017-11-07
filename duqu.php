<?php

require_once 'PHPExcel\PHPExcel.php'; 
require_once 'PHPExcel\PHPExcel\IOFactory.php';
require_once 'PHPExcel\PHPExcel\Reader\Excel5.php'; 
require_once 'PHPExcel\PHPExcel\Reader\Excel2007.php'; 

//上传文件
//print_r($_FILES);die;
$filename=$_FILES['myFile']['name'];
$type=$_FILES['myFile']['type'];
$tmp_name=$_FILES['myFile']['tmp_name'];
$size=$_FILES['myFile']['size'];
$error=$_FILES['myFile']['error'];
move_uploaded_file($tmp_name,$filename);

$objPHPExcel = new PHPExcel();
$conn = mysqli_connect('127.0.0.1','root','','kebiao');
$reader = PHPExcel_IOFactory::createReader('Excel2007'); //设置以Excel2007格式
$PHPExcel = $reader->load($filename); // 载入上传的excel文件   选择本地文件上传导入
$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
$highestRow = $sheet->getHighestRow(); // 取得总行数
$highestColumm = $sheet->getHighestColumn(); // 取得总列数
$conn->query("set names utf8");//设置编码



/** 循环读取并插入每个单元格的数据 */
for ($row = 5; $row <= $highestRow ; $row++){//行数是以第5行开始
    for ($column = 'A'; $column < $highestColumm; $column++) {//列数是以A列开始
        
        $dataset[] = $sheet->getCell($column.$row)->getValue();//单元格值放进数组
        //echo $column.$row.":".$sheet->getCell($column.$row)->getValue()."<br />";
       
    }
    //利用正则截取教室数据
    if($dataset[9]){
        $jiaoshi='/\d+/';
        $string=$dataset[9];
        if(preg_match($jiaoshi, $string,$matches)){
            //var_dump($matches);
            $dataset[9]=$matches[0];
        }else{
            echo '没有匹配到';
        }
    }
    //利用正则截取周次数据
    //
    if($dataset[11]){
        $zhouci='/\W/';
        $string=$dataset[11];
        if(preg_match($zhouci, $string,$matches)){
            var_dump($matches);
            $dataset[14]=$matches[0];
        }else{
            echo '没有匹配到';
        }
    //利用正则截取节次数据
        $jieci='/\d+-\d+/';
        $string=$dataset[11];
        if(preg_match($jieci, $string,$matches)){
            //var_dump($matches);
            $dataset[13]=$matches[0];
        }else{
            echo '没有匹配到';
        }
    }
    
    $sqli="INSERT into kecheng (XINGMING,ROOM,WEEK,CLASS) values('$dataset[5]','$dataset[9]','$dataset[14]','$dataset[13]')";
    $conn ->query($sqli);
    
    //打印
    // for($i = 0;$i < count($dataset);$i++){
    //     echo $dataset[$i].",";//输出每个单元格值
    // }
    // 
    // 打印错误
    //echo "Error: " . $sqli . "<br>" . $conn->error;
   
    echo "<br>";
    unset($dataset);//重置数组

}

?>

<?php
require_once "PHPExcel/IOFactory.php";

$conn = mysqli_connect('127.0.0.1','root','','kebiao');
$reader = PHPExcel_IOFactory::createReader('Excel2007'); //设置以Excel2007格式
$PHPExcel = $reader->load("kebiao.xlsx"); // 载入excel文件   选择本地文件上传导入
$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
$highestRow = $sheet->getHighestRow(); // 取得总行数
$highestColumm = $sheet->getHighestColumn(); // 取得总列数
$conn->query("set names utf8");//设置编码

/** 循环读取并插入每个单元格的数据 */
for ($row = 5; $row <= $highestRow ; $row++){//行数是以第2行开始
    for ($column = 'A'; $column < $highestColumm; $column++) {//列数是以A列开始
        $dataset[] = $sheet->getCell($column.$row)->getValue();//单元格值放进数组
        //echo $column.$row.":".$sheet->getCell($column.$row)->getValue()."<br />";
    }
    
    $sqli="INSERT into kecheng (XINGMING,ROOM,WEEK,CLASS) values('$dataset[5]','$dataset[9]','$dataset[11]','$dataset[12]')";
    $conn ->query($sqli);
    for($i = 0;$i < count($dataset);$i++){
        echo $dataset[$i].",";//输出每个单元格值
    }
    
   
    echo "<br>";
    unset($dataset);//重置数组

}

?>

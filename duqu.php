<?php
header("Content-type:text/html;charset=utf-8");  //头部
require_once 'PHPExcel\PHPExcel.php'; 
require_once 'PHPExcel\PHPExcel\IOFactory.php';
require_once 'PHPExcel\PHPExcel\Reader\Excel5.php'; 
require_once 'PHPExcel\PHPExcel\Reader\Excel2007.php'; 

$servername = "localhost";
$username ="root";
$password = "";

//上传文件
//print_r($_FILES);die;
$filename=$_FILES['myFile']['name'];
$type=$_FILES['myFile']['type'];
$tmp_name=$_FILES['myFile']['tmp_name'];
$size=$_FILES['myFile']['size'];
$error=$_FILES['myFile']['error'];
move_uploaded_file($tmp_name,$filename);

//链接mysql   自动建库建表
$conn = mysqli_connect($servername,$username,$password);
mysqli_set_charset($conn,'utf8_general_ci');//不写这句话的话会造成乱码
$sql1 = "CREATE DATABASE kebiao CHARACTER SET 'utf8' COLLATE 'utf8_general_ci'";
if(mysqli_query($conn,$sql1)){
    $dbname = "kebiao";
}else{
    echo "数据库创建失败:" . mysqli_error($conn);
}

$conn = mysqli_connect($servername,$username,$password,$dbname);
mysqli_set_charset($conn,'utf8_general_ci');//不写这句话的话会造成乱码
if(!$conn){
die("连接失败：".mysqli_connect_error());
}
//创建数据库(不能使用name，这是关键字)
$sql = "CREATE TABLE kecheng (
ID INT(6) UNSIGNED AUTO_INCREMENT PRIMARY KEY,
XINGMING VARCHAR(50) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
ROOM INT(4) NOT NULL,
WEEK VARCHAR(50) NOT NULL,
JIE VARCHAR(10) NOT NULL
);";
if(!(mysqli_query($conn,$sql))){
echo "数据表创建失败:" . mysqli_error($conn);
}

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
        $zhouci='/[\x{4e00}-\x{9fff}]/u';
        $string=$dataset[11];
        if(preg_match($zhouci, $string,$matches)){
            
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
    
    $sqli="INSERT into kecheng (XINGMING,ROOM,WEEK,JIE) values('$dataset[5]','$dataset[9]','$dataset[14]','$dataset[13]')";
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

header("refresh:0;url=http://localhost:82/excelxm/export");
?>

<?php 

require_once "IPLimit.php";

$Sender = "901";
$Receiver = $_GET["receiver"];
$Sms= "����һ��OA���������뾡�촦��";//$_GET["msg"];

$cmd = 0x1001;


$ObjApi= new COM("Rtxserver.rtxobj");
$objProp= new COM("Rtxserver.collection");
$Name = "SMSObject";
$ObjApi->Name = $Name;


$objProp->Add("SENDER", $Sender);
$objProp->Add("RECEIVER", $Receiver);
$objProp->Add("SMS", $Sms);
$Result = @$ObjApi->Call2($cmd, $objProp);

$dal=$_GET['callback'];
$errstr = $php_errormsg;
if(strcmp($nullstr, $errstr) != 0)
{
	echo $dal."({\"msg\":\"".$errstr."\"})";
}
echo $dal."({\"msg\":\"�����ɹ�!\"})";
?>


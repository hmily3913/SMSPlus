<?PHP

require_once "IPLimit.php";

$receiver = $_GET["username"];

//ȡ���û�״̬
$state = 12;
$ObjApi= new COM("Rtxserver.rtxobj");
$objProp= new COM("Rtxserver.collection");
$objProp->Add("Username", $receiver);
$Result = @$ObjApi->Call2(0x2001, $objProp);
$errstr = $php_errormsg;
if(strcmp($nullstr, $errstr) == 0)
{
  $state = intVal($Result);
}

$dal=$_GET['callback'];
$obj->state = $state;
echo $dal."(".json_encode($obj).")";

?>
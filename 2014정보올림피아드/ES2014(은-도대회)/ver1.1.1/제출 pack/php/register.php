<?
    $id= $_GET['id'];
    $pw = $_GET['pw'];

    $data = $id."/".$pw;

    $flocal = "users/".$id.".ini";
    if(file_exists($flocal) == 1){
        echo "FAIL";
    }
    else {
        $fopen = fopen($flocal,"a+");
        fwrite($fopen,$data);
        $fclose = fclose($fopen);

        echo "SUCCESS";
    }
?>

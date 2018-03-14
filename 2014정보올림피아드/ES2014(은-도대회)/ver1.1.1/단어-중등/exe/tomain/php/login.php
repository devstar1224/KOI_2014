<?
    $id= $_GET['id'];
    $pw = $_GET['pw']; 
    $flocal = "users/".$id.".ini"; 
    if(file_exists($flocal) == 1){
        $fopen = fopen($flocal,"r"); 
        $fread = fread($fopen,filesize($flocal)+10);
        $fclose = fclose($fopen);
        $data = explode("/",$fread); 
        if($data[1] == $pw){ 
            echo "1";
        }
        else { 
            echo "FAIL";
        }
    }
    else { 
        echo "FAIL"; 
    }
?>




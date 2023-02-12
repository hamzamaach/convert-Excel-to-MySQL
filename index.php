<?php
?>
<!doctype html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>excel</title>
</head>
<body>
<form action="#" method="post" enctype="multipart/form-data">
    <input type="file" name="excel"/><br><br>
    <input type="submit" name="submit"/>
</form>
<?php
if (isset($_FILES['excel']['name'])) {
    try {
        $conn = new PDO("mysql:host=localhost;dbname=excel", 'root', '');
        // set the PDO error mode to exception
        $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        echo "Connected successfully";
    } catch (PDOException $e) {
        echo "Connection failed: " . $e->getMessage();
    }
    include 'xlsx.php';
    if ($conn) {
        $excel = SimpleXLSX::parse(($_FILES['excel']['tmp_name']));
        echo "<pre>";
        /*
                print_r($excel->rows());
                echo "</pre>";*/
        for ($sheet = 0; $sheet < sizeof($excel->sheetNames()); $sheet++) {
            $i = 0;
            $rowCol = $excel->dimension($sheet);
            if ($rowCol[0] != 1 && $rowCol[1] != 1) {
                foreach ($excel->rows($sheet) as $key => $row) {
                    $q = "";
                    foreach ($row as $key => $cell) {
//                print_r($cell);
//                echo "<br>";
                        if ($i == 0) {
                            $q .= $cell . " varchar(50),";
                        } else {
                            $q .= "'".$cell . "',";
                        }
                    }
                    if ($i == 0) {
                        $query = "CREATE table " . $excel->sheetName($sheet) . " (" . rtrim($q, ",") . ");";
                    } else {
                        $query = "INSERT INTO " . $excel->sheetName($sheet) . " values (" . rtrim($q, ",") . ");";
                    }
                    $i++;
//                    echo $query;
//                    echo "<br>";
                    $stm = $conn->prepare($query);
                    $stm->execute();
                }
            }
        }
    }
}
?>
</body>
</html>

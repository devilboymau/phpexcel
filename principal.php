<!DOCTYPE html>

<?php
require  '../phpexcel/conexion1.php';


//consulta  a la base de datos
$query = "select id, usuario, contrasenia,sesion,ubicacion,fecha from usuario";
$resultado = pg_query($link, $query) or die("Error en la Consulta SQL");
$numReg = pg_num_rows($resultado);//regresa el numero de resultado de la consulta

if($numReg>0){
    echo "<table border='1' align='center'>
            <tr bgcolor='skyblue'>
            <th>ID</th>
            <th>Usuario</th>
            <th>Constraseña</th>
            <th>Sesion</th>
            <th>Ubicacion</th>
            <th>Fecha</th></tr>";

    while ($fila=pg_fetch_array($resultado)) {
          echo "<tr><td>".$fila['id']."</td>";
          echo "<td>".$fila['usuario']."</td>";
          echo "<td>".$fila['contrasenia']."</td>";
          echo "<td>".$fila['sesion']."</td>";
          echo "<td>".$fila['ubicacion']."</td>";
          echo "<td>".$fila['fecha']."</td></tr>";
      }
               echo "</table>";
        }else{
               echo "No hay Registros";
      }


$query="select nombre_ente, num_sesiones from sesiones";
$resultado = pg_query($link, $query) or die("Error en la Consulta SQL");
$numReg = pg_num_rows($resultado);//regresa el numero de 

if($numReg>0){
  echo "<table border='1' align='center'>
        <tr bgcolor='skyblue'>
        <th>Ente</th>
        <th>sesiones</th></tr>";


  while ($array=pg_fetch_array($resultado)) {
      echo $array['nombre_ente'];
      echo $array['num_sesiones'];
      echo "<tr><td>".$array['nombre_ente']."</td>";
      echo "<td>".$array['num_sesiones']."</td></tr>";
     
  }
        echo "</table>";
    }else{
        echo "No hay Registros";
    }
//$objPHPExcel->getActiveSheet()->fromArray($array, null, 'A1');
pg_close($link);

?>


<html>
   <head>
      <meta charset="utf-8" />
      <title>Reporte de excel con PHP y postgres</title>
   </head>
   <body>
      <div>
        <header>
           <h1>REPORTE A EXCEL</h1>
        </header>
        <div>
           <a href="http://localhost/phpexcel/index.php"> Haz clic para descargar el reporte</a>
        </div>
      </div>

  
   </body>
</html>
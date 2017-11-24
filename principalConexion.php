<!DOCTYPE html>

<?php
require  '../phpexcel/conexionSise.php';

//consulta  a la base de datos
$query = "select cve_directorio, cve_cargo_com, nombre, puesto, telefono, correo, id, suplente from directorio limit 5";
$resultado = pg_query($link, $query) or die("Error en la Consulta SQL");
$numReg = pg_num_rows($resultado);//regresa el numero de resultado de la consulta

if($numReg>0){
    echo "<table border='1' align='center'>
            <tr bgcolor='skyblue'>
            <th>cve_directorio</th>
            <th>cve_cargo_com</th>
            <th>nombre</th>
            <th>puesto</th>
            <th>telefono</th>
            <th>correo</th>
            <th>id</th>
            <th>suplente</th></tr>";

    while ($fila=pg_fetch_array($resultado)) {
          echo "<tr><td>".$fila['cve_directorio']."</td>";
          echo "<td>".$fila['cve_directorio']."</td>";
          echo "<td>".$fila['nombre']."</td>";
          echo "<td>".$fila['puesto']."</td>";
          echo "<td>".$fila['telefono']."</td>";
          echo "<td>".$fila['correo']."</td>";
          echo "<td>".$fila['id']."</td>";
          echo "<td>".$fila['suplente']."</td></tr>";
      }
               echo "</table>";
        }else{
               echo "No hay Registros";
      }


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
           <a href="http://localhost/phpexcel/indexNuevo.php"> Haz clic para descargar el reporte</a>
        </div>
      </div>

 

   </body>
</html>
<?php
	$usuario='user=usuario1';
	$password='password=12345';
	
	$database='dbname=dbcom';
	$host='host=127.0.0.1';
	$puerto='port=5432';
	$link = pg_connect($host." ".$puerto." ".$database." ".$usuario." ".$password);
	if (isset($link)) { 
	//echo "Conexion Exitosa<br>"; 
	}else{ 
	//	echo "Conexion Fallida<br>"; 
	} 
?>
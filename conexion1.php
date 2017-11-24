<?php

$usuario='user=postgres';
	$password='password=dir12345';
	
	$database='dbname=php_postgresql';
	$host='host=localhost';
	$puerto='port=5432';
	
	$link = pg_connect($host." ".$puerto." ".$database." ".$usuario." ".$password);
	 
//Cierra la conexion con la BD


?>
<?php

/** INCLUIMOS LA LIBRERIA PHP EXCEL*/
session_start();


//DEFINIMOS LA FECHA PARA NO TENER PROBLEMAS CON LA BASE DE DATOS
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('America/Mexico_City');
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
date_default_timezone_set('America/Mexico_City');


//IMPORTAMOS LAS CLASES NECESARIAS Y LA CONEXION A LA BD
require_once  '../phpexcel/Classes/PHPExcel.php';
require  '../phpexcel/conexionSise.php';


//OBJETOS DE PHP EXCEL
$objPHPExcel = new PHPExcel();
$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(10);
//OBJETO DE LA GRAFICA
$objWorksheet = $objPHPExcel->getActiveSheet(0);


// Establecer propiedades de la hoja de calculo
$objPHPExcel->getProperties()->setCreator("DCIGP")
							 ->setLastModifiedBy("DCIGP")
							 ->setTitle("Reporte")
							 ->setSubject("Reporte")
							 ->setDescription("Reporte")
							 ->setKeywords("reporte")
							 ->setCategory("Reporte");

$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' .$objPHPExcel->getProperties()->getTitle() . '&RPagina &P de &N');


//CREACION Y ACOMODO DE IMAGENES IZQUIERDA
$objDrawing1 = new PHPExcel_Worksheet_Drawing();
$objDrawing1->setWorksheet($objPHPExcel->setActiveSheetIndex(0));
$objDrawing1->setPath("../phpexcel/images/oaxaca.png");
$objDrawing1->setCoordinates('A1');
$objDrawing1->setHeight(110);
$objDrawing1->setWidth(212);


//CREACION Y ACOMODO DE IMAGENES DERECHA
$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setWorksheet($objPHPExcel->setActiveSheetIndex(0));
$objDrawing->setPath("../phpexcel/images/contraloria.png");
$objDrawing->setCoordinates('F1');
$objDrawing->setHeight(80);
$objDrawing->setWidth(120);


//TAMAÑO DE SESION ACTIVA DE LA HOJA DE CALCULO 
$objPHPExcel->getActiveSheet()->mergeCells('B3:B3');


//ESTILOS DE LA HOJA DE CALCULO SOMBREADO.ORIENTACION, COLOR 
$styleArray = array('font' => array('bold' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,'horizontal'=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,),),'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'rotation' => 90,'startcolor' => array('argb' => 'DCDCDC',),),);

//ESTILO DE LOS BORDES 
$styleArray8 = array('font' => array('regular' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_DOTTED,'color' => array('argb' => '7A7A7A'),),),);



$style = array('font' => array('bold' => true,),);


//SE PASA EL OBJETO QUE CONTENDRA LOS ESTILOS A LA HOJA ACTIVA
$objPHPExcel->getActiveSheet()->getStyle('A3:B3')->applyFromArray($style);
$objPHPExcel->getActiveSheet()->getStyle('A5:B3')->applyFromArray($styleArray);



//ANCHO DE LAS CELDAS
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(60);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20);

//ANCHO DE LAS CELDAS
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(60);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);


// INFORMACION DE LAS CABEZERAS
$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A5', 'ENTE')
			->setCellValue('B5', 'NUM. ACUERDOS TOTALES')


			->setCellValue('E5', 'ENTE')
			->setCellValue('F5', 'NUM. ACUERDOS CUMPLIDOS');
		

//AGREGAR ESTILO A LAS COLUMNAS QUE CONTENDRAN LOS ENCABEZADOS
$objPHPExcel->getActiveSheet()->getStyle('A3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);	
$objPHPExcel->getActiveSheet()->getStyle('A5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('B5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('E5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('F5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);


//DATOS DE LA PRIMERA SERIE
$sql="select des_organo,count(se_cumplio) numero
from cat_organismos
inner join comites on cat_organismos.cve_organismo =comites.cve_organismo
inner join sesiones on comites.id = sesiones.id
inner join acuerdos on sesiones.cve_sesion = acuerdos.cve_sesion
group by des_organo
order by numero desc";

$query=pg_query($sql);

	$x=6;

	while($array=pg_fetch_array($query)){	
		$objPHPExcel->getActiveSheet()->SetCellValue("A".$x, $ente=$array['des_organo']);		
		$objPHPExcel->getActiveSheet()->SetCellValue("B".$x, $num_acuerdos=$array['numero']); 
		

			$data[]=array($ente,$num_acuerdos);
			$objWorksheet->fromArray($data,' ', 'A6');//celda desde donde se va iniciar el arreglo
		$x++;
		
	}




//DATOS DE LA SEGUNDA SERIE
$sql2="select des_organo,count(se_cumplio) numero
from cat_organismos
inner join comites on cat_organismos.cve_organismo =comites.cve_organismo
inner join sesiones on comites.id = sesiones.id
inner join acuerdos on sesiones.cve_sesion = acuerdos.cve_sesion
where se_cumplio='1'
group by des_organo
order by numero desc";

$query2=pg_query($sql2);

	$x2=6;

	while($array2=pg_fetch_array($query2)){	
		$objPHPExcel->getActiveSheet()->SetCellValue("E".$x2, $ente2=$array2['des_organo']);		
		$objPHPExcel->getActiveSheet()->SetCellValue("F".$x2, $num_acuerdos2=$array2['numero']); 
		

			$data2[]=array($ente2,$num_acuerdos2);
			$objWorksheet->fromArray($data2,' ', 'E6');//celda desde donde se va iniciar el arreglo
		$x++;
		
	}



	

//DATOS DE LA ETIQUETAS DE LOS DATOS DE LA GRAFICA TOTAL DE ACUERDOS
$dataseriesLabels1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$7', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$8', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$9', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$10', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$11', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$12', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$13', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$14', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$15', null, 1),
		

	);



//DATOS DE EJE X DE LA GRAFICA
/*
$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$7', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$8', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$9', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$10', null, 1),
	
	

);
*/

//SERIE DE DATOS A GRAFICAR
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$7', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$8', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$9', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$10', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$11', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$12', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$13', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$14', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$15', null, 1),	
	
	
);



$series1 = new PHPExcel_Chart_DataSeries(					// ASGNAMOS LOS DISTINTOS OBJETOS QUE CONTRUYEN LA GRAFICA
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,				// TIPO DE GRAFICA
	PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
	range(0, count($dataSeriesValues1)-1),					// CONTAMOS LOS DATOS DE LA GRAFICA
	$dataseriesLabels1,										// PASAMOS LA ETIQUETA DE DATOS A LA GRAFICA
	null,													// ASIGNAMOS EL PLANO X A LA GRAFICA
	$dataSeriesValues1										// SE AGREGAN LA SERIE DE DATOS A LA GRAFICA
);


//	UBICAMOS LOS DATOS EN LE AREA DE LA GRAFICA
$plotarea1 = new PHPExcel_Chart_PlotArea(null, array($series1));
//	CREAMOS LA POSICION DEL TITULO DE LA GRAFICA
$legend1 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_TOPRIGHT, null, false);
// TITULO DE LA GRAFICA
$title1 = new PHPExcel_Chart_Title('10 Entes con más Acuerdos');
// INFORMACION EN EL EJE Y
$yAxisLabel1 = new PHPExcel_Chart_Title('Acuerdos');
// INFORMACION EN EL EJE X
$xAxisLabel1 = new PHPExcel_Chart_Title('Entes');


//	MATERIALIZAMOS LA GRAFICA EN LA HOJA DE CALCULO 
$chart1 = new PHPExcel_Chart(
	'chart1',		// NOMBRE DE LA GRAFICA
	$title1,		// ASIGANMOS EL TITULO DE LA GRAFICA
	$legend1,		// POSICION DEL TITULO
	$plotarea1,		// DATOS QUE CONTENDRA LA GRAFICA
	true,			// HACEMOS VISIBLE EL AREA DE LA GRAFICA
	0,				
	$xAxisLabel1,	// PASAMOS LA INFORMACION DEL EJE X
	$yAxisLabel1	// PASAMOS LA INFORMACION DEL EJE Y
);


//	ASIGNAMOS LA POSICION DE LAS CELDAS DONDE APARECERA LA GRAFICA
$chart1->setTopLeftPosition('A100');
$chart1->setBottomRightPosition('D130');

$objWorksheet->addChart($chart1);

//	AGREGAMOS LA GRAFICA A LA HOJA ACTIVA







//DATOS DE LA ETIQUETAS DE LOS DATOS DE LA GRAFICA ACUERDOS REALIZADOS
$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$7', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$8', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$9', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$10', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$11', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$12', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$13', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$14', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$E$15', null, 1),
		

	);



//DATOS DE EJE X DE LA GRAFICA
$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$7', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$8', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$9', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$10', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$11', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$12', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$13', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$14', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$15', null, 1),
	
	

);

//SERIE DE DATOS A GRAFICAR
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$7', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$8', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$9', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$10', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$11', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$12', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$13', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$14', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$F$15', null, 1),	
		
	
);



$series2 = new PHPExcel_Chart_DataSeries(					// ASGNAMOS LOS DISTINTOS OBJETOS QUE CONTRUYEN LA GRAFICA
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,				// TIPO DE GRAFICA
	PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
	range(0, count($dataSeriesValues2)-1),					// CONTAMOS LOS DATOS DE LA GRAFICA
	$dataseriesLabels2,										// PASAMOS LA ETIQUETA DE DATOS A LA GRAFICA
	null,													// ASIGNAMOS EL PLANO X A LA GRAFICA
	$dataSeriesValues2										// SE AGREGAN LA SERIE DE DATOS A LA GRAFICA
);


//	UBICAMOS LOS DATOS EN LE AREA DE LA GRAFICA
$plotarea2 = new PHPExcel_Chart_PlotArea(null, array($series2));
//	CREAMOS LA POSICION DEL TITULO DE LA GRAFICA
$legend2 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_TOPRIGHT, null, false);
// TITULO DE LA GRAFICA
$title2 = new PHPExcel_Chart_Title('10 Entes con más Acuerdos Cumplidos');
// INFORMACION EN EL EJE Y
$yAxisLabel2 = new PHPExcel_Chart_Title('Acuerdos');
// INFORMACION EN EL EJE X
$xAxisLabel2 = new PHPExcel_Chart_Title('Entes');


//	MATERIALIZAMOS LA GRAFICA EN LA HOJA DE CALCULO 
$chart2 = new PHPExcel_Chart(
	'chart2',		// NOMBRE DE LA GRAFICA
	$title2,		// ASIGANMOS EL TITULO DE LA GRAFICA
	$legend2,		// POSICION DEL TITULO
	$plotarea2,		// DATOS QUE CONTENDRA LA GRAFICA
	true,			// HACEMOS VISIBLE EL AREA DE LA GRAFICA
	0,				
	$xAxisLabel2,	// PASAMOS LA INFORMACION DEL EJE X
	$yAxisLabel2	// PASAMOS LA INFORMACION DEL EJE Y
);


//	ASIGNAMOS LA POSICION DE LAS CELDAS DONDE APARECERA LA GRAFICA
$chart2->setTopLeftPosition('E100');
$chart2->setBottomRightPosition('H130');


//	AGREGAMOS LA GRAFICA A LA HOJA ACTIVA



$objWorksheet->addChart($chart2);






// MANDAMOS LA HOJA ACTIVA A LA PRIMERA PESTAÑA DE EXCEL
$objPHPExcel->setActiveSheetIndex(0);


// SE MODIFICAN LOS ENCABEZADOS DEL HTML 
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//NOMBRE DEL ARCHIVO DE EXCEL
header('Content-Disposition: attachment;filename="Reporte Acuerdos.xlsx"');
header('Cache-Control: max-age=0');


// EL OBJETO QUE TENEMOS LO CONVERTIREMOS UN UN NUEVO OBJETO PRA PODERLO DESCARGAR
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel, 'Excel2007');


// INCLUIMNOS LA GRAFICA A LA HOJA
$objWriter->setIncludeCharts(TRUE);

// ASIGNAMOS LA SALIDA DEL ARCHIVO A DESCARGAR
$objWriter->save('php://output');


exit;

 ?>



 


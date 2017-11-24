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
require  '../phpexcel/conexion1.php';


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
$objDrawing1->setCoordinates('B1');
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
$objPHPExcel->getActiveSheet()->mergeCells('A3:F3');


//ESTILOS DE LA HOJA DE CALCULO SOMBREADO.ORIENTACION, COLOR 
$styleArray = array('font' => array('bold' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,'horizontal'=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,),),'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'rotation' => 90,'startcolor' => array('argb' => 'DCDCDC',),),);

//ESTILO DE LOS BORDES 
$styleArray8 = array('font' => array('regular' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_DOTTED,'color' => array('argb' => '7A7A7A'),),),);



$style = array('font' => array('bold' => true,),);


//SE PASA EL OBJETO QUE CONTENDRA LOS ESTILOS A LA HOJA ACTIVA
$objPHPExcel->getActiveSheet()->getStyle('A3:F3')->applyFromArray($style);
$objPHPExcel->getActiveSheet()->getStyle('A5:F5')->applyFromArray($styleArray);



//ANCHO DE LAS CELDAS
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(22);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(18);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(10);



// INFORMACION DE LAS CABEZERAS
$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A3', 'GRAFICA')
			->setCellValue('A5', 'ID')
			->setCellValue('B5', 'ENTE')
			->setCellValue('C5', 'NUM. SESIONES')
                        ->setCellValue('D5', 'FECHA')
			->setCellValue('E5', 'EXTRA_SESIONES')
			->setCellValue('F5', 'ID');
		

//AGREGAR ESTILO A LAS COLUMNAS QUE CONTENDRAN LOS ENCABEZADOS
$objPHPExcel->getActiveSheet()->getStyle('A3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);	
$objPHPExcel->getActiveSheet()->getStyle('A5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('B5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('C5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('D5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('E5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('F5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);



//CONSULTA  A LA BASE DE DATOS
$sql="select id_ente, nombre_ente, num_sesiones,fec_sesiones,extra_sesiones,id from sesiones";
$query=pg_query($sql);

	$x=6;

	while($array=pg_fetch_array($query)){	
		$objPHPExcel->getActiveSheet()->SetCellValue("A".$x, $id=$array['id_ente']);		
		$objPHPExcel->getActiveSheet()->SetCellValue("B".$x, $ente=$array['nombre_ente']); 
		$objPHPExcel->getActiveSheet()->SetCellValue("C".$x, $sesion=$array['num_sesiones']);
		$objPHPExcel->getActiveSheet()->SetCellValue("D".$x, $fecha=$array['fec_sesiones']); 
		$objPHPExcel->getActiveSheet()->SetCellValue("E".$x, $extra=$array['extra_sesiones']); 
		$objPHPExcel->getActiveSheet()->SetCellValue("F".$x, $ids=$array['id']);

			$data[]=array($id,$ente,$sesion,$fecha,$extra,$ids);
			$objWorksheet->fromArray($data,' ', 'A6');//celda desde donde se va iniciar el arreglo
		$x++;
		
	}

	

//DATOS DE LA ETIQUETAS DE LOS DATOS DE LA GRAFICA
$dataseriesLabels1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$7', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$8', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$9', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$10', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$11', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$12', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$13', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$14', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$15', null, 1),
	
														
);

//DATOS DE EJE X DE LA GRAFICA
$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$6:$B$15', null, 10),

);

//SERIE DE DATOS A GRAFICAR
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$6', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$7', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$8', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$9', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$10', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$11', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$12', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$13', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$14', null, 1),
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$15', null, 1),
);



$series1 = new PHPExcel_Chart_DataSeries(					// ASGNAMOS LOS DISTINTOS OBJETOS QUE CONTRUYEN LA GRAFICA
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,				// TIPO DE GRAFICA
	PHPExcel_Chart_DataSeries::GROUPING_STANDARD,			// TIPO DE AGRUPAMIENTO DE LA GRAFICA
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
$title1 = new PHPExcel_Chart_Title('HISTORICO DE SESIONES');
// INFORMACION EN EL EJE Y
$yAxisLabel1 = new PHPExcel_Chart_Title('Sesiones');
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
$chart1->setTopLeftPosition('A20');
$chart1->setBottomRightPosition('K43');


//	AGREGAMOS LA GRAFICA A LA HOJA ACTIVA

$objWorksheet->addChart($chart1);


// AGREGAMOS A TITULO A LA HOJA ACTIVA

//$objWorksheet->getStyle('A6:F6'.$x);


// MANDAMOS LA HOJA ACTIVA A LA PRIMERA PESTAÑA DE EXCEL
$objPHPExcel->setActiveSheetIndex(0);


// SE MODIFICAN LOS ENCABEZADOS DEL HTML 
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//NOMBRE DEL ARCHIVO DE EXCEL
header('Content-Disposition: attachment;filename="PruebaGraficas.xlsx"');
header('Cache-Control: max-age=0');


// EL OBJETO QUE TENEMOS LO CONVERTIREMOS UN UN NUEVO OBJETO PRA PODERLO DESCARGAR
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel, 'Excel2007');


// INCLUIMNOS LA GRAFICA A LA HOJA
$objWriter->setIncludeCharts(TRUE);

// ASIGNAMOS LA SALIDA DEL ARCHIVO A DESCARGAR
$objWriter->save('php://output');


exit;

 ?>



 


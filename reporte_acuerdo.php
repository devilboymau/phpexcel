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
$objDrawing->setCoordinates('B1');
$objDrawing->setHeight(80);
$objDrawing->setWidth(120);


//TAMAÑO DE SESION ACTIVA DE LA HOJA DE CALCULO 
$objPHPExcel->getActiveSheet()->mergeCells('A3:D3');


//ESTILOS DE LA HOJA DE CALCULO SOMBREADO.ORIENTACION, COLOR 
$styleArray = array('font' => array('bold' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,'horizontal'=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,),),'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'rotation' => 90,'startcolor' => array('argb' => 'DCDCDC',),),);

//ESTILO DE LOS BORDES 
$styleArray8 = array('font' => array('regular' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_DOTTED,'color' => array('argb' => '7A7A7A'),),),);



$style = array('font' => array('bold' => true,),);


//SE PASA EL OBJETO QUE CONTENDRA LOS ESTILOS A LA HOJA ACTIVA
$objPHPExcel->getActiveSheet()->getStyle('A3:D3')->applyFromArray($style);
$objPHPExcel->getActiveSheet()->getStyle('A5:D3')->applyFromArray($styleArray);



//ANCHO DE LAS CELDAS
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(60);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20);



// INFORMACION DE LAS CABEZERAS
$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A3', 'GRAFICA')
			->setCellValue('A5', 'ENTE')
			->setCellValue('B5', 'NUM. ACUERDOS');
		

//AGREGAR ESTILO A LAS COLUMNAS QUE CONTENDRAN LOS ENCABEZADOS
$objPHPExcel->getActiveSheet()->getStyle('A3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);	
$objPHPExcel->getActiveSheet()->getStyle('A5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('B5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);



$sql="select des_organo,count(se_cumplio) numero
from cat_organismos
inner join comites on cat_organismos.cve_organismo =comites.cve_organismo
inner join sesiones on comites.id = sesiones.id
inner join acuerdos on sesiones.cve_sesion = acuerdos.cve_sesion
group by des_organo";

$query=pg_query($sql);

	$x=6;

	while($array=pg_fetch_array($query)){	
		$objPHPExcel->getActiveSheet()->SetCellValue("A".$x, $ente=$array['des_organo']);		
		$objPHPExcel->getActiveSheet()->SetCellValue("B".$x, $num_acuerdos=$array['numero']); 
		

			$data[]=array($ente,$num_acuerdos);
			$objWorksheet->fromArray($data,' ', 'A6');//celda desde donde se va iniciar el arreglo
		$x++;
		
	}

	

//DATOS DE LA ETIQUETAS DE LOS DATOS DE LA GRAFICA
$dataseriesLabels1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$6', null, 1),
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$7', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$8', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$9', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$10', null, 1),	

	);



//DATOS DE EJE X DE LA GRAFICA
$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$6:$A$10', null, 5),

);

//SERIE DE DATOS A GRAFICAR
$dataSeriesValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$6', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$7', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$8', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$9', null, 1),	
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$10', null, 1),	
	
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
$title1 = new PHPExcel_Chart_Title('Acuerdos por Ente');
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
$chart1->setTopLeftPosition('A20');
$chart1->setBottomRightPosition('N45');


//	AGREGAMOS LA GRAFICA A LA HOJA ACTIVA

$objWorksheet->addChart($chart1);





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



 


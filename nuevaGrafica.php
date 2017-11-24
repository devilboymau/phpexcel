<?php

set_include_path(get_include_path() . PATH_SEPARATOR . '../Classes/');
/** PHPExcel */
include '../phpexcel/Classes/PHPExcel.php';
require  '../phpexcel/conexion1.php';
$objPHPExcel = new PHPExcel();
//agregado
$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(10);



$objWorksheet = $objPHPExcel->getSheet(0); //PUEDE LLEVAR 0



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
$objWorksheet = $objPHPExcel->setActiveSheetIndex(0)
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





//Zona del arreglo

$arr="select  id_ente, nombre_ente, num_sesiones, fec_sesiones, extra_sesiones, id from sesiones";
  $query=pg_query($arr);
  //$array=pg_fetch_array($query, 0, PGSQL_NUM);

	$x=6;
 
while ($array=pg_fetch_array($query)) {
  		
  		$objPHPExcel->getActiveSheet()->SetCellValue("A".$x, $id_ente=$array['id_ente']); 
   		$objPHPExcel->getActiveSheet()->SetCellValue("B".$x, $ente=$array['nombre_ente']); 
   		$objPHPExcel->getActiveSheet()->SetCellValue("C".$x, $sesion=$array['num_sesiones']);
   	    $objPHPExcel->getActiveSheet()->SetCellValue("D".$x, $fecha=$array['fec_sesiones']); 
   		$objPHPExcel->getActiveSheet()->SetCellValue("E".$x, $extra=$array['extra_sesiones']); 
   		$objPHPExcel->getActiveSheet()->SetCellValue("F".$x, $id=$array['id']); 
   	
			$data[]=array($id_ente, $ente, $sesion, $fecha, $extra, $id);
			$objWorksheet->fromArray($data,' ', 'A6');

		$x++;	
   }


//Creacion del la hoja de calculo
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


$xAxisTickValues1 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$6:$B$15', null, 10),	

);


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


	$series1 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_BARCHART,				// plotType
	PHPExcel_Chart_DataSeries::GROUPING_STANDARD,			// plotGrouping
	range(0, count($dataSeriesValues1)-1),					// plotOrder
	$dataseriesLabels1,										// plotLabel
	null,													// plotCategory
	$dataSeriesValues1										// plotValues
);


//	Set the series in the plot area
$plotarea1 = new PHPExcel_Chart_PlotArea(null, array($series1));
//	Set the chart legend
$legend1 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_BOTTOM, null, false);
$title1 = new PHPExcel_Chart_Title('HISTORICO DE SESIONES');
$yAxisLabel1 = new PHPExcel_Chart_Title('Sesiones');
$xAxisLabel1 = new PHPExcel_Chart_Title('Entes');


//	Create the chart
$chart1 = new PHPExcel_Chart(
	'chart1',		// name
	$title1,		// title
	$legend1,		// legend
	$plotarea1,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	$xAxisLabel1,	// xAxisLabel
	$yAxisLabel1
);

//	Set the position where the chart should appear in the worksheet
$chart1->setTopLeftPosition('A20');
$chart1->setBottomRightPosition('S43');
//	Add the chart to the worksheet



//$objWorksheet->getActiveSheet(0);
$objWorksheet->getStyle('A6:F6'.$x);


/*
$objWorksheet = new \PHPExcel_Worksheet($objPHPExcel, 'Summary');
$objPHPExcel->addSheet($objWorksheet, 0);
$objWorksheet->setTitle('Summary');
*/

$objWorksheet->addChart($chart1);
//$objWorksheet->setTitle('Reporte');
//$objPHPExcel->setTitle('Reporte');
//$objPHPExcel->getActiveSheet()->getStyle('A6:F6'.$x)->applyFromArray($styleArray8);

//$objWorksheet = $objPHPExcel->setActiveSheetIndex(0);


// MANDAMOS LA HOJA ACTIVA A LA PRIMERA PESTAÑA DE EXCEL
$objPHPExcel->setActiveSheetIndex(0);
//$objWorksheet->getActiveSheet()->setTitle('Reporte');


header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="PruebaGraficas.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$objWriter->save('php://output');



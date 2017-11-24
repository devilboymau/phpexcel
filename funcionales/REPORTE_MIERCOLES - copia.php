<?php
session_start();
require("../phpexcel/limpia_fecha.php");


/** Error reporting */
//error_reporting(E_ALL);

date_default_timezone_set('America/Mexico_City');

/** Include PHPExcel */
require_once '../phpexcel/Classes/PHPExcel.php';
require('../phpexcel/conexionSise.php');


$area="DCIGP";
$anio="2017";
$condicionA="";
$cdepto="";
if(isset($_GET['DCIGP'])){
$cdepto=trim($_GET['DCIGP']);
}
if($area!='T' && ($cdepto=='' || $cdepto=='COV')){
	$ccondicion=" and trim(func_direccion.cve_area_resp)='".$area."' ";
}else if($area!='T' && $cdepto!=''){
	$ccondicion=" and trim(func_direccion.cve_area_resp)='".$area."' and trim(func_direccion.cve_depto_resp)='".trim($cdepto)."'";
}

$condicion=" Where (borrado='0' or borrado is null) and anio=".$anio.$condicionA;
$condicion2=" and anio=".$anio.$condicionA;

//TOMAMOS EL AÑO Y EL AREA DEL SISE 2017 DCGIP																										

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(10);

$objWorksheet = $objPHPExcel->getActiveSheet(0);

// Set document properties
$objPHPExcel->getProperties()->setCreator("DCIGP")
							 ->setLastModifiedBy("DCIGP")
							 ->setTitle("Reporte")
							 ->setSubject("Reporte")
							 ->setDescription("Reporte")
							 ->setKeywords("reporte")
							 ->setCategory("Reporte");
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' .$objPHPExcel->getProperties()->getTitle() . '&RPagina &P de &N');

$objDrawing1 = new PHPExcel_Worksheet_Drawing();
$objDrawing1->setWorksheet($objPHPExcel->setActiveSheetIndex(0));
$objDrawing1->setPath("../phpexcel/images/oaxaca.png");
$objDrawing1->setCoordinates('A1');
$objDrawing1->setHeight(110);
$objDrawing1->setWidth(212);

$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setWorksheet($objPHPExcel->setActiveSheetIndex(0));
$objDrawing->setPath("../phpexcel/images/contraloria.png");
$objDrawing->setCoordinates('G1');
$objDrawing->setHeight(80);
$objDrawing->setWidth(120);

$objPHPExcel->getActiveSheet()->mergeCells('A3:J3');

$styleArray = array('font' => array('bold' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,'horizontal'=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,),),'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'rotation' => 90,'startcolor' => array('argb' => 'DCDCDC',),),);
//allborders
$styleArray8 = array('font' => array('regular' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_DOTTED,'color' => array('argb' => '7A7A7A'),),),);
$style = array('font' => array('bold' => true,),);

$objPHPExcel->getActiveSheet()->getStyle('A3:J3')->applyFromArray($style);
$objPHPExcel->getActiveSheet()->getStyle('A5:J6')->applyFromArray($styleArray);

$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Logo');
$objDrawing->setDescription('Logo');
$objDrawing->setPath('../phpexcel/images/EncabezadoSCyTG.png');
$objDrawing->setHeight(36);


$objPHPExcel->getActiveSheet()->mergeCells('A5:A6');
$objPHPExcel->getActiveSheet()->mergeCells('B5:B6');
$objPHPExcel->getActiveSheet()->mergeCells('C5:C6');
$objPHPExcel->getActiveSheet()->mergeCells('D5:D6');
$objPHPExcel->getActiveSheet()->mergeCells('E5:J5');


$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(4);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(35);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(35);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(14);



// Add some data

$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A3', 'REPORTE SESIONES ASISTIDAS')
			->setCellValue('A5', 'No')
			->setCellValue('B5', 'MES')
            ->setCellValue('C5', 'AREA RESPONSABLE')
			->setCellValue('D5', 'DEPARTAMENTO RESPONSABLE')
			->setCellValue('E5', 'SESIONES ORDINARIAS')
			->setCellValue('E6', '1 ORDINARIA')
			->setCellValue('F6', '2 ORDINARIA')
			->setCellValue('G6', '3 ORDINARIA')
			->setCellValue('H6', '4 ORDINARIA')
			->setCellValue('I6', '5 ORDINARIA')
			->setCellValue('J6', 'TOTAL');
			
			
$objPHPExcel->getActiveSheet()->getStyle('A3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);	
$objPHPExcel->getActiveSheet()->getStyle('A5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('B5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('C5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('D5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('E5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('F5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('G5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('H5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('H6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('I5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('I6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('J5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('J6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
	
	


$sql="select cve_area_resp,desc_area_resp,cve_depto_resp,desc_depto_resp,tipo_sesion,num_sesion,fecha_sesion,count(mes) as cta,mes
from(
SELECT 	func_direccion.id,trim(func_direccion.cve_area_resp) as cve_area_resp,trim(func_direccion.cve_depto_resp) as cve_depto_resp,
trim(sesiones.tipo_sesion) as tipo_sesion,sesiones.num_sesion,
Case 
	 When to_char(sesiones.fecha_sesion, 'MM')='01' then 'ENERO'
 	 When to_char(sesiones.fecha_sesion, 'MM')='02' then 'FEBRERO' 
	 When to_char(sesiones.fecha_sesion, 'MM')='03' then 'MARZO'
	 When to_char(sesiones.fecha_sesion, 'MM')='04' then 'ABRIL'
	 When to_char(sesiones.fecha_sesion, 'MM')='05' then 'MAYO'
	 When to_char(sesiones.fecha_sesion, 'MM')='06' then 'JUNIO'
	 When to_char(sesiones.fecha_sesion, 'MM')='07' then 'JULIO'
	 When to_char(sesiones.fecha_sesion, 'MM')='08' then 'AGOSTO'
	 When to_char(sesiones.fecha_sesion, 'MM')='09' then 'SEPTIEMBRE'
	 When to_char(sesiones.fecha_sesion, 'MM')='10' then 'OCTUBRE'
	 When to_char(sesiones.fecha_sesion, 'MM')='11' then 'NOVIEMBRE'
	 When to_char(sesiones.fecha_sesion, 'MM')='12' then 'DICIEMBRE'
 else 'SIN FECHA'
 
 End as fecha_sesion,extract(month from sesiones.fecha_sesion) as mes,
 trim(cat_area_resp.desc_area_resp) as desc_area_resp,trim(cat_depto_resp.desc_depto_resp) as desc_depto_resp 
 
 FROM FUNC_DIRECCION 

left join comites on func_direccion.id=comites.id
left join sesiones on func_direccion.id=sesiones.id
left join status_sesiones on sesiones.cve_sesion=status_sesiones.cve_sesion
left join cat_area_resp on trim(func_direccion.cve_area_resp)=trim(cat_area_resp.cve_area_resp)
left join cat_depto_resp on trim(func_direccion.cve_depto_resp)=trim(cat_depto_resp.cve_depto_resp)

where  (comites.borrado='0' or comites.borrado is null) and status_sesiones.asistio='1' and anio=".$anio.$ccondicion."  order by func_direccion.cve_depto_resp,sesiones.fecha_sesion)
	as s1

group by cve_area_resp, desc_area_resp,cve_depto_resp,desc_depto_resp,tipo_sesion,fecha_sesion,mes,num_sesion
order by  mes,tipo_sesion,cve_area_resp, desc_area_resp,cve_depto_resp,desc_depto_resp,num_sesion";



$query=pg_query($link, $sql);
$x=6;
//$j=6;
$no=0;
//$vins=0;
$suma=0;
$suma1=0; //SUMA 1 ORDINARIAS
$suma2=0; //SUMA 2 ORDINARIAS
$suma3=0; //SUMA 3 ORDINARIAS
$suma4=0; //SUMA 4 ORDINARIAS
$suma5=0; //SUMA 5 ORDINARIAS
$sumaT=0; //SUMA TOTAL
$depto="";
$areaD=""; //AREA DEPARTAMENTO
$mes="";

$sumaN=0;




while($array=pg_fetch_array($query)){	
		
		
		if(trim($array['fecha_sesion'])!=$mes || trim($array['cve_area_resp'])!=$areaD  || trim($array['cve_depto_resp'])!=$depto ){
			if($no!=0){
				$objPHPExcel->getActiveSheet()->SetCellValue("J".$x, $suma);
			}
			$no++;
			$x++; //REALIZA UNA SUMA Y SALTA DE RENGLON PARA INSERTAR LA NUEVA SUMA
			
				
				$objPHPExcel->getActiveSheet()->SetCellValue("A".$x, $no); 

				$objPHPExcel->getActiveSheet()->getStyle('B'.$x)->getAlignment()->setWrapText(true);
				$objPHPExcel->getActiveSheet()->SetCellValue("B".$x, $array['fecha_sesion']);

				$objPHPExcel->getActiveSheet()->getStyle('C'.$x)->getAlignment()->setWrapText(true);
				$objPHPExcel->getActiveSheet()->SetCellValue("C".$x, trim($array['desc_area_resp']));

				$objPHPExcel->getActiveSheet()->getStyle('D'.$x)->getAlignment()->setWrapText(true);
				$objPHPExcel->getActiveSheet()->SetCellValue("D".$x, trim($array['desc_depto_resp']));

				
			$sumaT+=$suma;
			$mes=trim($array['fecha_sesion']);
			$areaD=trim($array['cve_area_resp']);
			$depto=trim($array['cve_depto_resp']);

			$suma=0;
			
		}
		

		if(trim($array['tipo_sesion'])=='O' && (trim($array['num_sesion'])=='1') ){
			$objPHPExcel->getActiveSheet()->SetCellValue("E".$x, $array['cta']);
			$suma1+=$array['cta'];
		}

		else if(trim($array['tipo_sesion'])=='O' && (trim($array['num_sesion'])=='2') ){
			$objPHPExcel->getActiveSheet()->SetCellValue("F".$x, $array['cta']);
			$suma2+=$array['cta'];
		}


		else if(trim($array['tipo_sesion'])=='O'  && trim($array['num_sesion'])=='3'){
			$objPHPExcel->getActiveSheet()->SetCellValue("G".$x, $array['cta']);
			$suma3+=$array['cta'];
		}


		else if(trim($array['tipo_sesion'])=='O'  && trim($array['num_sesion'])=='4'){
			$objPHPExcel->getActiveSheet()->SetCellValue("H".$x, $array['cta']);
			$suma4+=$array['cta'];
		}

		else if(trim($array['tipo_sesion'])=='O'  && trim($array['num_sesion'])=='5'){
			$objPHPExcel->getActiveSheet()->SetCellValue("I".$x, $array['cta']);
			$suma5+=$array['cta'];
		}

		
		$suma=$array['cta']+$suma;
	
			
		}
		
		$x++;


		$objPHPExcel->getActiveSheet()->getStyle('D'.$x.':J'.$x)->applyFromArray($style);
		$objPHPExcel->getActiveSheet()->SetCellValue('D'.$x, "TOTAL");
		$objPHPExcel->getActiveSheet()->SetCellValue("E".$x, $suma1);
		$objPHPExcel->getActiveSheet()->SetCellValue("F".$x, $suma2);
		$objPHPExcel->getActiveSheet()->SetCellValue("G".$x, $suma3);
		$objPHPExcel->getActiveSheet()->SetCellValue("H".$x, $suma4);
		$objPHPExcel->getActiveSheet()->SetCellValue("I".$x, $suma5);


		for ($x=7; $x<=14; $x++) { 
			$objPHPExcel->getActiveSheet()->setCellValue("J".$x,'=SUM(E'.$x.':I'.$x.')');}
			$objPHPExcel->getActiveSheet()->setCellValue("J".$x,'=SUM(E'.$x.':I'.$x.')'); 


$objPHPExcel->getActiveSheet()->getStyle('A7:J'.$x)->applyFromArray($styleArray8);
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="Reporte Ordinarias.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->setIncludeCharts(TRUE);
$objWriter->save('php://output');
exit;


	

<?php
session_start();
require("../phpexcel/limpia_fecha.php");


/** Error reporting */
//error_reporting(E_ALL);

date_default_timezone_set('America/Mexico_City');

/** Include PHPExcel */
require_once '../phpexcel/Classes/PHPExcel.php';
require('../phpexcel/conexionSise.php');


//$area=$_GET['area'];
//$anio=$_GET['anio'];

$area="DCIGP";
$anio="2017";

$condicionA="";

$cdepto=$_GET['xdepto'];
//$cdepto=$_GET['DCIGP'];



if($area!='T'){
	$condicionA=" and trim(func_direccion.cve_area_resp)='".$area."' ";
} 

if($cdepto!='' && $cdepto!='COV' ){
	$condicionA=$condicionA." and trim(func_direccion.cve_depto_resp)='".$cdepto."' ";
}


$condicion=" Where (borrado='0' or borrado is null) and anio=".$anio.$condicionA;
$condicion2=" and anio=".$anio.$condicionA;

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(10);
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
$objDrawing1->setPath("./images/oaxaca.png");
$objDrawing1->setCoordinates('A1');
$objDrawing1->setHeight(110);
$objDrawing1->setWidth(212);

$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setWorksheet($objPHPExcel->setActiveSheetIndex(0));
$objDrawing->setPath("./images/contraloria.png");
$objDrawing->setCoordinates('G1');
$objDrawing->setHeight(80);
$objDrawing->setWidth(120);

$objPHPExcel->getActiveSheet()->mergeCells('A3:H3');

$styleArray = array('font' => array('bold' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,'horizontal'=>PHPExcel_Style_Alignment::HORIZONTAL_CENTER,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,),),'fill' => array('type' => PHPExcel_Style_Fill::FILL_SOLID,'rotation' => 90,'startcolor' => array('argb' => 'DCDCDC',),),);
//allborders
$styleArray8 = array('font' => array('regular' => true,),'alignment' => array('vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP,),'borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_DOTTED,'color' => array('argb' => '7A7A7A'),),),);
$style = array('font' => array('bold' => true,),);

$objPHPExcel->getActiveSheet()->getStyle('A3:P3')->applyFromArray($style);


$objDrawing = new PHPExcel_Worksheet_Drawing();
$objDrawing->setName('Logo');
$objDrawing->setDescription('Logo');
$objDrawing->setPath('./images/EncabezadoSCyTG.png');
$objDrawing->setHeight(36);


$objPHPExcel->getActiveSheet()->mergeCells('A5:A6');
$objPHPExcel->getActiveSheet()->mergeCells('B5:B6');
$objPHPExcel->getActiveSheet()->mergeCells('C5:C6');
$objPHPExcel->getActiveSheet()->mergeCells('D5:D6');
$objPHPExcel->getActiveSheet()->mergeCells('E5:E6');
$objPHPExcel->getActiveSheet()->mergeCells('F5:F6');
$objPHPExcel->getActiveSheet()->mergeCells('G5:G6');

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(8);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(24);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(32);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(34);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(26);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(14);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('L')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('M')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('N')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('O')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('P')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('S')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('T')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('U')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('V')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('W')->setWidth(16);
$objPHPExcel->getActiveSheet()->getColumnDimension('X')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('Y')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('Z')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AA')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('AB')->setWidth(20);


// Add some data

$objPHPExcel->setActiveSheetIndex(0)
			->setCellValue('A3', 'REPORTE DETALLADO')
			->setCellValue('A5', 'ID')
            ->setCellValue('B5', 'AREA RESPONSABLE')
			->setCellValue('C5', 'TIPO DE ÓRGANISMO PÚBLICO EN QUE PARTICIPA')
			->setCellValue('D5', 'NOMBRE DE LA DEPENDENCIA Y/O ENTIDAD')
			->setCellValue('E5', 'TIPO DE COMITE QUE PARTICIPA LA DIRECCIÓN')
			->setCellValue('F5', 'FUNCIÓN QUE SE DESEMPEÑA EN LAS SESIONES DEL COMITÉ')
			->setCellValue('G5', 'NUM. DE SESIONES')
			->setCellValue('H5', 'FECHAS')
			->setCellValue('H6', 'INSTALACION DE COMITÉ');
			
			
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
$objPHPExcel->getActiveSheet()->getStyle('I6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('J6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('K6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('L6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('M6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('N6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('O6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);

			
	$sql1="select MAX(c) as m from (select count(sesiones.id) as c from sesiones left join func_direccion on sesiones.id=func_direccion.id left join comites on comites.id=sesiones.id WHERE (borrado='0' or borrado is null) and tipo_sesion='O' ".$condicion2." group by sesiones.id ) t";
	$query1=pg_query($link,$sql1);
		$data1=pg_fetch_array($query1);
	if($data1['m']==0){
		$data1['m']=4;
	}
	$total=$data1['m']+8;
	$col=1;		
	for ($i=8;$i<$total; $i++){
			$colString = PHPExcel_Cell::stringFromColumnIndex($i);
			$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($i,6,"ORDINARIA ".$col);
			$objPHPExcel->getActiveSheet()->getStyle($colString.'6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
			$col++;
	}
	
	$sql1a="select MAX(c) as m from (select count(sesiones.id) as c from sesiones left join func_direccion on sesiones.id=func_direccion.id left join comites on comites.id=sesiones.id WHERE (borrado='0' or borrado is null) and tipo_sesion='E' ".$condicion2." group by sesiones.id ) t";
	$query1a=pg_query($link,$sql1a);
	$data1a=pg_fetch_array($query1a);
	if($data1a['m']==0){
		$data1a['m']=3;

	}
	$total2=$data1a['m']+$total;
	$col=1;		
	for ($s=$total;$s<$total2; $s++){
			$colString = PHPExcel_Cell::stringFromColumnIndex($s);
			$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($s,6,"EXTRAORDINARIA ".$col);
			$objPHPExcel->getActiveSheet()->getStyle($colString.'6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
			//$objPHPExcel->getActiveSheet()->getCellByColumnAndRow($s,5)->setValue("EXTRAORDINARIA ".$col);
			$col++;
	}
			

$objPHPExcel->getActiveSheet()->mergeCells('H5:'.$colString.'5');

$sql="SELECT comites.id,cat_organismos.tip_organo,TRIM(cat_tip_org.desc_tip_org) AS tip_org,func_direccion.cve_area_resp,trim(cat_area_resp.desc_area_resp) AS area_resp,comites.cve_organismo,TRIM(cat_organismos.des_organo)AS des_org,TRIM(cat_organismos.tip_organo) AS cve_tipo_com, TRIM(cat_tipo_com.des_tip_com) AS tipo_com, fun_des_ses, TRIM(desc_func) AS desc_func, num_sesiones, anio,cve_funcion,status,ya_instalado
	FROM comites 
	LEFT JOIN cat_organismos ON TRIM(comites.cve_organismo)=TRIM(cat_organismos.cve_organismo)
	LEFT JOIN cat_tip_org ON cat_organismos.tip_organo=cat_tip_org.cve_tip_org
	LEFT JOIN func_direccion ON comites.id=func_direccion.id
	LEFT JOIN cat_area_resp ON func_direccion.cve_area_resp=cat_area_resp.cve_area_resp
	LEFT JOIN cat_func_des_ses ON TRIM(func_direccion.fun_des_ses)=TRIM(cat_func_des_ses.cve_func)
	LEFT JOIN cat_tipo_com ON comites.cve_tipo_com =cat_tipo_com.cve_tip_com ".$condicion." ORDER BY func_direccion.cve_area_resp,cat_organismos.tip_organo ";

$query=pg_query($link, $sql);
$x=7;


while($array=pg_fetch_array($query)){	
	if($array['status']=='0'){
			$colString1 = PHPExcel_Cell::stringFromColumnIndex($total2);
			cellColor('A'.$x.':'.$colString1.$x, 'D8D8D8');
	}
	$objPHPExcel->getActiveSheet()->SetCellValue("A".$x, $array['id']); 
	$objPHPExcel->getActiveSheet()->getStyle('B'.$x)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->SetCellValue("B".$x, $array['area_resp']); 
	$objPHPExcel->getActiveSheet()->getStyle('C'.$x)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->SetCellValue("C".$x, $array['tip_org']);
	$objPHPExcel->getActiveSheet()->getStyle('D'.$x)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->SetCellValue("D".$x, $array['des_org']); 
	$objPHPExcel->getActiveSheet()->getStyle('E'.$x)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->SetCellValue("E".$x, $array['tipo_com']); 
	$objPHPExcel->getActiveSheet()->getStyle('F'.$x)->getAlignment()->setWrapText(true);
	$objPHPExcel->getActiveSheet()->SetCellValue("F".$x, $array['desc_func']); 
	$objPHPExcel->getActiveSheet()->SetCellValue("G".$x, $array['num_sesiones']); 
	
	if($array['cve_area_resp']=='DCIGP'){
		$tipoSesion=" (tipo_sesion='OG1' or tipo_sesion='OG') ";
	}else{
		$tipoSesion=" tipo_sesion='OG' ";
	}
	
	$sql2="select func_direccion.id,tipo_sesion,num_sesion,fecha_sesion,confirmado,asistio,acta from  func_direccion left join sesiones  on func_direccion.id=sesiones.id left join status_sesiones on func_direccion.cve_funcion=status_sesiones.cve_funcion  and sesiones.cve_sesion=status_sesiones.cve_sesion where func_direccion.id=".$array['id']." and ".$tipoSesion.$condicionA;
	$query2=pg_query($link,$sql2);
	$data2=pg_fetch_array($query2);
	if($data2['confirmado']=='1'){
			cellColor('H'.$x, '92D050');
	}if($data2['asistio']=='1'){
			cellColor('H'.$x, '29A8FF');
	}if($data2['acta']=='1'){
			cellColor('H'.$x, 'F67519');
	}
	
	$objPHPExcel->getActiveSheet()->SetCellValue("H".$x, limpiafecha2($data2['fecha_sesion'])); 
	
	$sql3="select func_direccion.id,tipo_sesion,num_sesion,fecha_sesion,confirmado,asistio,acta from  func_direccion left join sesiones  on func_direccion.id=sesiones.id left join status_sesiones on func_direccion.cve_funcion=status_sesiones.cve_funcion  and sesiones.cve_sesion=status_sesiones.cve_sesion where func_direccion.id=".$array['id'].$condicionA." and tipo_sesion='O' order by num_sesion"; 
	$query3=pg_query($link,$sql3);
	$iit=7;
	while($data3=pg_fetch_array($query3)){
		$ii=$iit+$data3['num_sesion'];
		$colString = PHPExcel_Cell::stringFromColumnIndex($ii);
		if($data3['confirmado']=='1'){
			cellColor($colString.$x, '92D050');
		}if($data3['asistio']=='1'){
			cellColor($colString.$x, '29A8FF');
		}if($data3['acta']=='1'){
			cellColor($colString.$x, 'F67519');
		}
		$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($ii,$x,limpiafecha2($data3['fecha_sesion']));
		//$ii++;
	}
	
	$sql4="select func_direccion.id,tipo_sesion,num_sesion,fecha_sesion,confirmado,asistio,acta from  func_direccion left join sesiones  on func_direccion.id=sesiones.id left join status_sesiones on func_direccion.cve_funcion=status_sesiones.cve_funcion  and sesiones.cve_sesion=status_sesiones.cve_sesion where func_direccion.id=".$array['id'].$condicionA." and tipo_sesion='E' order by num_sesion";
	$query4=pg_query($link,$sql4);
	$iiit=($total-1);
	while($data4=pg_fetch_array($query4)){
		$iii=$iiit+$data4['num_sesion'];
		$colString = PHPExcel_Cell::stringFromColumnIndex($iii);
		if($data4['confirmado']=='1'){
			cellColor($colString.$x, '92D050');
		}if($data4['asistio']=='1'){
			cellColor($colString.$x, '29A8FF');
		}if($data4['acta']=='1'){
			cellColor($colString.$x, 'F67519');
		}
		$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($iii,$x,limpiafecha2($data4['fecha_sesion']));
		//$iii++;
	}
	
	if($array['ya_instalado']=='1'){
		$iii2=$total2;
		$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($iii2,$x,"*");
	}
	
	$x++;
			
}
// Rename worksheet
$x--;

$objPHPExcel->getActiveSheet()->setTitle('Reporte');
$a=$total2-1;
$colString = PHPExcel_Cell::stringFromColumnIndex($a);
$objPHPExcel->getActiveSheet()->getStyle('A5:'.$colString.'6')->applyFromArray($styleArray);
$objPHPExcel->getActiveSheet()->getStyle('A7:'.$colString.$x)->applyFromArray($styleArray8);
$x=$x+3;
cellColor("A".$x.":B".$x, '92D050');
$objPHPExcel->getActiveSheet()->mergeCells('A'.$x.':B'.$x);
$objPHPExcel->getActiveSheet()->SetCellValue("A".$x,"FECHAS CONFIRMADAS"); 
$x++;
cellColor("A".$x.":B".$x, '29A8FF');
$objPHPExcel->getActiveSheet()->mergeCells('A'.$x.':B'.$x);
$objPHPExcel->getActiveSheet()->SetCellValue("A".$x,"ASISTIDAS"); 
$x++;
cellColor("A".$x.":B".$x, 'F67519');
$objPHPExcel->getActiveSheet()->mergeCells('A'.$x.':B'.$x);
$objPHPExcel->getActiveSheet()->SetCellValue("A".$x,"CUENTAN CON ACTA"); 
if($_SESSION['arearesp']=='DCIGP'){
	$x++;
	cellColor("A".$x.":B".$x, 'D8D8D8');
	$objPHPExcel->getActiveSheet()->mergeCells('A'.$x.':B'.$x);
	$objPHPExcel->getActiveSheet()->SetCellValue("A".$x,"NO PUEDEN SER INSTALADAS"); 
	$x++;
	$objPHPExcel->getActiveSheet()->mergeCells('A'.$x.':D'.$x);
	$objPHPExcel->getActiveSheet()->SetCellValue("A".$x,"* ORGANISMOS QUE YA TENIAN INSTALADO SU COMITE DE CONTROL INTERNO"); 
}
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="Reporte Detallado.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save('php://output');
exit;


	

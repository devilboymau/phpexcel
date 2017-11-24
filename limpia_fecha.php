<?php
	function limpiafecha($nfecha){
		if(trim($nfecha)==''){
			return $nfecha;
		}else{
			return strftime("%d/%m/%Y",strtotime($nfecha));
		}
	}	


	function limpiafecha2($nfecha){
		  if($nfecha==''){
			  return $nfecha;
		  }else{
			  return strftime("%d/%m/%Y %H:%M",strtotime($nfecha));
		  }
	}	

?>
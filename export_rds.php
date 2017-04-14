<?php
	/* ������� ������������ ��� � ������� XLS */
		
	function f_export_rds(&$DB) {
		/* ������� ���� � ����� ������ */
		$xls =& new Spreadsheet_Excel_Writer($_SERVER['DOCUMENT_ROOT']."/export_rds.xls");
		/* ������ */
		$F_txt =& $xls->addFormat();
		$F_txt->setFontFamily('Times New Roman');
		$F_txt->setSize('10');
		$F_txt->setAlign('left');
		#$F_txt->setTextWrap();
		$F_txt->setAlign('vcenter');
			
		$WorksheetName = '���';
		// Create worksheet
		$cart =& $xls->addWorksheet($WorksheetName);
		/* ������� */
		$cart->setLandscape();
		$cart->hideGridlines();
		/* ������ ����� */
		$cart->write(0,0,"�������� ��������",$F_txt);
		$cart->write(0,1,"����� �������, ���������, ��������",$F_txt);
		$cart->write(0,2,"�����-�������������",$F_txt);
		$cart->write(0,3,"�����-���",$F_txt);
		$cart->write(0,4,"��� ������ (� ������. ������� ������)",$F_txt);
		$cart->write(0,5,"��� ��������� (� ������. ������� ������)",$F_txt);
			
		$rows = 0;
			
		/* ������ ������ � �� */
		$q_sel = $DB->SELECT("
						SELECT 
							p.name,
							d.dosage_name,
							d.dosage_code,
							p.product_id, 
							d.dosage_id
						FROM `omni_products` p
						INNER JOIN `omni_products_dosage` d ON d.product_id = p.product_id"
		);
						
		/* ��������� ������� */			
		foreach ($q_sel as $key => $val){
			$rows++;
			$cart->write($rows,0,$val['name'],$F_txt);
			$cart->write($rows,1,$val['dosage_name'],$F_txt);
			$cart->write($rows,3,$val['dosage_code'],$F_txt);
			$cart->write($rows,4,$val['product_id'],$F_txt);
			$cart->write($rows,5,$val['dosage_id'],$F_txt);
		}
			
		$xls->close();
	}
?>
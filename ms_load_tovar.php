<?php
	/* ���������� ��������������� ������ */
	require_once "ms_system.php";
	/* ������� � ����� ������ ERP */
	$dbh_av = ibase_connect($ht_av, $ur_av, $pd_av, "win1251");		
	
	/* ��������� ����� � MS */
	
	/* ������ ����������  */
	$connector = new COM("Cleverence.Warehouse.StorageConnector") or die("error create StorageConnector");
	/* ������������ � ���� MS */
	$connector->SelectCurrentApp($ms_base); 
	 /* �������� �������� */
	try {
		$connector->BeginUploadProducts(true, true, true); 
		/* ������ ������� */
		$products = new COM("Cleverence.Warehouse.ProductCollection") or die("error create ProductCollection");
		$q_select = "select * from R_MS_TOVAR";
		$res = ibase_query($dbh_av, $q_select);
		while ($row = ibase_fetch_assoc($res)) {
			/* ��������� �� id */ 
			$uploadProduct = $products->FindById($row['ITEMID']);
			if($uploadProduct == null) {
				/* ��������� �� 1000, ����� �������������� �������� */
				if($products->Count >= 1000) {
					/* ��������� � ���� */
					$connector->UploadProducts($products);
					$products = new COM("Cleverence.Warehouse.ProductCollection") or die("error create ProductCollection");
				}
				/* ���������� �� ������ */
				$uploadProduct = new COM("Cleverence.Warehouse.Product") or die("error create Product");
				$uploadProduct->SetField("Id", $row['ITEMID']);
				$uploadProduct->SetField("Name", $row['NAME']);
				$uploadProduct->SetField("Marking", $row['MARKING']);
				$uploadProduct->SetField("Barcode", $row['BARCODE']);	
				$uploadProduct->SetField("���������", $row['NDS']);
				$uploadProduct->SetField("�������������", $row['VENDOR']);
				$uploadProduct->SetField("���������������", $row['VENDOR_ID']);
				$uploadProduct->SetField("������", $row['CNTR']);
				$uploadProduct->SetField("��������", $row['CNTR_ID']);
				$uploadProduct->SetField("withsn", '1');
				$uploadProduct->SetField("withserial", '1');
				
				/* ���������� �� �������� */
				$uploadPacking = new COM("Cleverence.Warehouse.Packing") or die("error create Packing");
				$uploadPacking->SetField("Id", $row['ITEMID']);
				$uploadPacking->SetField("Name", $row['UNIT']);

				$uploadProduct->Packings->Add($uploadPacking); /* ��������� � ����� �������� */
				
				$products->Add($uploadProduct); /* ��������� ����� � ������ */
			}
		}
		ibase_free_result($res);

		if($products->Count > 0) {
			$connector->UploadProducts($products);
		}
		/* ��������� �������� */
		$connector->EndUploadProducts();
	} catch (Exception $e) {
		/* ������ �������� */
		$connector->ResetUploadProducts();
	}

	ibase_close($dbh_av);
?>
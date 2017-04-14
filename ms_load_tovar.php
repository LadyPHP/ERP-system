<?php
	/* подключаем регистрационные данные */
	require_once "ms_system.php";
	/* коннект с базай данных ERP */
	$dbh_av = ibase_connect($ht_av, $ur_av, $pd_av, "win1251");		
	
	/* Загружаем товар в MS */
	
	/* объект соединения  */
	$connector = new COM("Cleverence.Warehouse.StorageConnector") or die("error create StorageConnector");
	/* подключаемся к базе MS */
	$connector->SelectCurrentApp($ms_base); 
	 /* начинаем загрузку */
	try {
		$connector->BeginUploadProducts(true, true, true); 
		/* массив товаров */
		$products = new COM("Cleverence.Warehouse.ProductCollection") or die("error create ProductCollection");
		$q_select = "select * from R_MS_TOVAR";
		$res = ibase_query($dbh_av, $q_select);
		while ($row = ibase_fetch_assoc($res)) {
			/* проверяем по id */ 
			$uploadProduct = $products->FindById($row['ITEMID']);
			if($uploadProduct == null) {
				/* загружаем по 1000, чтобы контролировать нагрузку */
				if($products->Count >= 1000) {
					/* загружаем в базу */
					$connector->UploadProducts($products);
					$products = new COM("Cleverence.Warehouse.ProductCollection") or die("error create ProductCollection");
				}
				/* информация по товару */
				$uploadProduct = new COM("Cleverence.Warehouse.Product") or die("error create Product");
				$uploadProduct->SetField("Id", $row['ITEMID']);
				$uploadProduct->SetField("Name", $row['NAME']);
				$uploadProduct->SetField("Marking", $row['MARKING']);
				$uploadProduct->SetField("Barcode", $row['BARCODE']);	
				$uploadProduct->SetField("СтавкаНДС", $row['NDS']);
				$uploadProduct->SetField("Производитель", $row['VENDOR']);
				$uploadProduct->SetField("ПроизводительИд", $row['VENDOR_ID']);
				$uploadProduct->SetField("Страна", $row['CNTR']);
				$uploadProduct->SetField("СтранаИд", $row['CNTR_ID']);
				$uploadProduct->SetField("withsn", '1');
				$uploadProduct->SetField("withserial", '1');
				
				/* информация по упаковке */
				$uploadPacking = new COM("Cleverence.Warehouse.Packing") or die("error create Packing");
				$uploadPacking->SetField("Id", $row['ITEMID']);
				$uploadPacking->SetField("Name", $row['UNIT']);

				$uploadProduct->Packings->Add($uploadPacking); /* добавляем в товар упаковку */
				
				$products->Add($uploadProduct); /* добавляем товар в массив */
			}
		}
		ibase_free_result($res);

		if($products->Count > 0) {
			$connector->UploadProducts($products);
		}
		/* завершаем загрузку */
		$connector->EndUploadProducts();
	} catch (Exception $e) {
		/* отмена загрузки */
		$connector->ResetUploadProducts();
	}

	ibase_close($dbh_av);
?>
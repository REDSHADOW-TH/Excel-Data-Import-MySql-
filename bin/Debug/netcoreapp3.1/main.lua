filePath = 'D:\\Git\\Excel-Data-Import-MySql-\\bin\\Debug\\netcoreapp3.1\\File\\raw_upro.xlsx'

excelConfig = {
    path = filePath, sheet = 'Sheet1', startRow = 2, endRowIgnore = 1
}


data = {
    importData:MakeImportData('docdate','Document Date'),
    importData:MakeImportData('plant','Plant'),
    importData:MakeImportData('storage','Storage Location'),
    importData:MakeImportData('sales_document','Sales Document'),
    importData:MakeImportData('billing_document','Billing Document'),
    importData:MakeImportData('billing_date','Billing Date'),
    importData:MakeImportData('requested_deliv_date','Requested deliv.date'),
    importData:MakeImportData('delivery','Delivery'),
    importData:MakeImportData('purchase_order_number','Purchase order number'),
    importData:MakeImportData('sold_to_party','Sold-to party'),
    importData:MakeImportData('ship_to_party','Ship-to party'),
    importData:MakeImportData('ship_to_party_name','Ship-to party Name'),
    importData:MakeImportData('material','Material'),
    importData:MakeImportData('material_description','Material Description'),
    importData:MakeImportData('order_quantity','Order Quantity'),
    importData:MakeImportData('cumul_confirm_qty','Cumul.confirmed qty'),
    importData:MakeImportData('net_value','Net value'),
    importData:MakeImportData('pricing_date','Pricing date'),
    importData:MakeImportData('delivery_quantity','Delivery quantity'),
    importData:MakeImportData('net_price','Net Price'),
    importData:MakeImportData('billing_status','Billing status'),
    importData:MakeImportData('loc_sal','Loc.Sal. Workforce 1'),
    importData:MakeImportData('loc_sal_workforce1','Loc.Sal. Workforce 1 Name'),
    importData:MakeImportData('create_by','Created by'),
    importData:MakeImportData('key_customer4','Key Customer 4'),
    importData:MakeImportData('key_customer4_name','Key Customer 4 Name'),
    importData:MakeImportData('loc_sal_workforce4_name','Loc.Sal. Workforce 4 Name'),
    importData:MakeImportData('reason_for_rejection','Reason for rejection'),
    importData:MakeImportData('loc_sal_regions4_name','Loc.Sal. Regions 4 Name'),
    importData:MakeImportData('loc_sal_regions4','Loc.Sal. Regions 4'),
    importData:MakeImportData('loc_sal_workforce3_name','Loc.Sal. Workforce 3 Name'),
    importData:MakeImportData('loc_sal_workforce2','Loc.Sal. Workforce 2'),
    importData:MakeImportData('loc_sal_workforce3_names','Loc.Sal. Workforce 3 Name'),
    importData:MakeImportData('loc_sal_workforce3','Loc.Sal. Workforce 3'),
    importData:MakeImportData('loc_sal_workforce4','Loc.Sal. Workforce 4'),
    importData:MakeImportData('bill_to_party_name','Bill-to party Name'),
    importData:MakeImportData('bill_to_party','Bill-to party'),
    importData:MakeImportData('payer_name','Payer Name'),
    importData:MakeImportData('payer','Payer'),
    importData:MakeImportData('after_ppr','After PPR'),
    importData:MakeImportData('total_tpr','Total TPR'),
    importData:MakeImportData('total_ppr','Total PPR'),
    importData:MakeImportData('list_price_per_sales_uom','List Price Per Sales UOM'),
    importData:MakeImportData('list_price','List Price')
}


system:import(excelConfig, data, 'daily_order_upro_raw')

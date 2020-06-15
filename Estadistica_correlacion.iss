Sub Main
	IgnoreWarning(True)
	Set db = Client.OpenDatabase("Tabladin�mica10.IMD")
	Set task = db.Correlation
	task.AddFieldForCorrelation "COD_PROD_01"
	task.AddFieldForCorrelation "COD_PROD_02"
	task.AddFieldForCorrelation "COD_PROD_03"
	task.AddFieldForCorrelation "COD_PROD_04"
	task.AddFieldForCorrelation "COD_PROD_05"
	task.AddFieldForCorrelation "COD_PROD_06"
	task.AddFieldForCorrelation "TOTAL"
	resultName = db.UniqueResultName("Correlaci�n01")
	task.ResultName = resultName
	task.PerformTask
	
	Set db = Client.OpenDatabase("Tabladin�mica10.IMD")
	Set task = db.Correlation
	task.AddFieldForCorrelation "COD_PROD_01"
	task.AddFieldForCorrelation "COD_PROD_02"
	task.AddFieldForCorrelation "COD_PROD_03"
	task.AddFieldForCorrelation "COD_PROD_04"
	task.AddFieldForCorrelation "COD_PROD_05"
	task.AddFieldForCorrelation "COD_PROD_06"
	task.AddFieldForCorrelation "TOTAL"
	task.CreateResult = FALSE
	dbName = "Correlaci�n01.IMD"
	task.OutputDBName = dbName
	task.PerformTask
	
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase(dbName)
	Client.RefreshFileExplorer
End Sub


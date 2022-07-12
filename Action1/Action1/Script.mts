
On Error Resume Next
Dim strTestDataFilePath
Dim iRowCount,iLoop,libraryPath, currentRow

strTestDataFilePath="C:\Liberty Africa\LibertyAfricaTestData\iLabTestData.xlsx"
DataTable.ImportSheet strTestDataFilePath,"Data","Global"


If Err.Number<>0  Then

	Reporter.ReportEvent micWarning,"Fail to load test data","Error number is "& err.number & "and description is : " & err.description
	
	Else
		    iRowCount=DataTable.GetRowCount
	  	  For iLoop = 1 To iRowCount Step 1
	        DataTable.SetCurrentRow(iLoop)
	        Displayed_Results=""
    
     

	If (Lcase(DataTable.Value("in_Run"))="yes") Then
			    currentRow= DataTable.GetCurrentRow
				Call LogIn() 
				DataTable.SetCurrentRow(currentRow)
			End if
			wait 2
		Call Apply()
Next
End If






  
  

  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  

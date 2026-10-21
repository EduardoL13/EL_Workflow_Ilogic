Sub Main ()
    
	Dim esteDoc As AssemblyDocument = ThisDoc.Document
	
	Try
		
		skDocName = esteDoc.PropertySets.Item("Inventor User Defined Properties").Item("Skeleton Document").Value
		
	Catch

		Dim oDlg As New System.Windows.Forms.OpenFileDialog 'Archivo y no carpeta
	    Dim dlgResult As New System.Windows.Forms.DialogResult
	
    	With oDlg

    		'.ShowNewFolderButton = True
    		.Title = "Seleccionar archivo de Inventor"
            .Filter = "Archivos de Inventor (*.ipt; *.iam; *.idw; *.dwg)|*.ipt;*.iam;*.idw;*.dwg|Todos los archivos (*.*)|*.*"   
            .InitialDirectory = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
			
    		If .ShowDialog = dlgResult.Cancel
	    		Exit Sub
			End If
	        
			skDocName = .FileName
			
    	End With
		esteDoc.PropertySets.Item("Inventor User Defined Properties").Add(skDocName,"Skeleton Document")
		'esteDoc.PropertySets.Item("Inventor User Defined Properties").Item("Skeleton Document").Value = skDocName
		'skDocName = dlgResult
	End Try
	
	'Dim skDoc As PartDocument = ThisApplication.Documents.ItemByName("C:\Users\Owner\Desktop\8 Stations New Canopy\Canopy Param model 8.ipt") 'Cambiar en el futuro por un buscador en la carpeta
    Dim skDoc As PartDocument = ThisApplication.Documents.ItemByName(skDocName)
	Dim assemParams As UserParameters = esteDoc.ComponentDefinition.Parameters.UserParameters
	cf_in = 2.54 ' Conversion factor cm/in (length values must be given in cm in ilogic)
	cf_deg = PI/180 ' Conversion factor rad/Deg (angle values must be given in rad in ilogic)
	cf_mm = 1 / 10 ' Conversion factor mm->cm
	
	Dim assemParamsNames As New List(Of String)
	
	For Each aParameter In assemParams
		assemParamsNames.Add(aParameter.Name)
	Next
	
	For Each uParameter In skDoc.ComponentDefinition.Parameters.UserParameters
		If uParameter.IsKey Then
            Try
		        If uParameter.Units = "in"
			        uParameter.Value = Parameter(uParameter.Name) * cf_in
				Else If uParameter.Units = "mm"
					uParameter.Value = Parameter(uParameter.Name) * cf_mm	
				Else If uParameter.Units = "deg"
					uParameter.Value = Parameter(uParameter.Name) * cf_deg
			    End If
				
			Catch
				
				assemParams.AddByValue(uParameter.Name,uParameter.Value,uParameter.Units)
				MsgBox("Skeleton key parameter " & uParameter.Name & " has been added into Top Leven Assembly Parameters  ")
				
			End Try
			
		End If
	Next
	
	esteDoc.Update2
	
End Sub

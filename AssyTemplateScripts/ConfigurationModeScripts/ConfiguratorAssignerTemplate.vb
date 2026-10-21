Sub Main ()
    
	'Set Parameters and features access
	'MsgBox("Conf triggered")
	Dim currentDoc As AssemblyDocument = ThisDoc.Document
    Dim patternsList As OccurrencePatterns = currentDoc.ComponentDefinition.OccurrencePatterns
	
	'Expressions to add if required
    'Expression = "(kepWidth + noSigns * kepExcessTop + ( noSigns - 1 ul ) * kepGapBetSigns)/2"
	
	'Configuration list
	configuration1 = "G1"
	configuration2 = "N1"
	configuration3 = "N3"
	configuration4 = "T1"
	
	configuration5 = "N1LL"	
	configuration6 = "N3_8"
	
	'Changes configuration
	

    If parameter("currentConfiguration") = configuration1 Then
	' Listar Supresiones y cambios en kep parameters
	    Parameter("kepLength") = 36
        Parameter("kepHeight") = 16 
		Parameter("kepOffsetAnchorsFromEdge") = 6
		currentDoc.ComponentDefinition.ModelStates.Item("[Primary]").Activate
		MsgBox(configuration1 & " conf activa")
		
    Else If parameter("currentConfiguration") = configuration2 Then
		
	    Parameter("kepLength") = 96
        Parameter("kepHeight") = 8 
		Parameter("kepOffsetAnchorsFromEdge") = 6
		currentDoc.ComponentDefinition.ModelStates.Item("TrimConnLong").Activate
		MsgBox(configuration2 & " conf activa")

    Else If parameter("currentConfiguration") = configuration3 Then
		
	    Parameter("kepLength") = 96
        Parameter("kepHeight") = 8 
		Parameter("kepOffsetAnchorsFromEdge") = 0.001
		currentDoc.ComponentDefinition.ModelStates.Item("TrimConnLong-SF").Activate
		MsgBox(configuration3 & " conf activa")
		
    Else If parameter("currentConfiguration") = configuration4 Then
		
	    Parameter("kepLength") = 30
        Parameter("kepHeight") = 14 
		Parameter("kepOffsetAnchorsFromEdge") = 0.25
		currentDoc.ComponentDefinition.ModelStates.Item("TrimConn").Activate
		MsgBox(configuration4 & " conf activa")	
		
    Else If parameter("currentConfiguration") = configuration5 Then
		
	    Parameter("kepLength") = 96
        Parameter("kepHeight") = 12
		Parameter("kepOffsetAnchorsFromEdge") = 6
		currentDoc.ComponentDefinition.ModelStates.Item("TrimConnlong").Activate
		MsgBox(configuration5 & " conf activa")	
		
    Else If parameter("currentConfiguration") = configuration6 Then
		
	    Parameter("kepLength") = 96
        Parameter("kepHeight") = 12
		Parameter("kepOffsetAnchorsFromEdge") = 0.001
		currentDoc.ComponentDefinition.ModelStates.Item("TrimConnLong-SF").Activate
		MsgBox(configuration6 & " conf activa")			
		
	End If


End Sub.

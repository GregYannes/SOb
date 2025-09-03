Attribute VB_Name = "Test"



' #############
' ## Testing ##
' #############

' Test the lifecycle of a simulated "Dix" object.
Public Sub Test()
	' ### Construction ###
	' Dim dix As Object:	 SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Collection: SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Object:	 Set dix = SOb.New_Obj("Dix")
	Dim dix As Object: Set dix = DixSOb.New_Dix()
	
	
	' ### Classification ###
	Dim copy As String: SOb.Obj_ClassKey copy
	Debug.Print "Obj_ClassKey() = """ & copy & """"
	Debug.Print "Obj_HasClass(dix) = " & SOb.Obj_HasClass(dix)
	Debug.Print "Obj_Class(dix) = """ & SOb.Obj_Class(dix) & """"
	
	Debug.Print
	
	
	' ### Typology 1 ###
	Debug.Print "IsObj(dix) = " & SOb.IsObj(dix)
	Debug.Print "IsObj(dix, """ & DixSOb.DIX_CLS & """) = " & SOb.IsObj(dix, DixSOb.DIX_CLS)
	Debug.Print "IsObj(dix, ""Other"") = " & SOb.IsObj(dix, "Other")
	
	Debug.Print
	
	
	' ### Fields 1 ###
	SOb.Obj_Field(dix, DixSOb.Dix_Field.Count) = 42
	
	SOb.Obj_FieldKey copy, DixSOb.Dix_Field.Count
	Debug.Print "Obj_FieldKey(Dix_Field.Count) = """ & copy & """"
	Debug.Print "Obj_HasField(dix, Dix_Field.Count) = " & SOb.Obj_HasField(dix, DixSOb.Dix_Field.Count)
	Debug.Print "Obj_Field(dix, Dix_Field.Count) = " & SOb.Obj_Field(dix, DixSOb.Dix_Field.Count)
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	
	' ### Typology 2 ###
	Debug.Print "IsDix(dix) = " & DixSOb.IsDix(dix)
	
	
	' ### Fields 2 ###
	DixSOb.Dix_Count(dix) = 7
	Debug.Print "Dix_Count(dix) = " & DixSOb.Dix_Count(dix)
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	
	' ### Visualization ###
	Test_Print dix
End Sub


' Test printing for a simulated "Dix" object.
Private Sub Test_Print(ByRef dix As Object)
	Dim depth As Integer: depth = 1
	Dim plain As Boolean: plain = False
	Dim pointer As Boolean: pointer = True
	Dim summary As String: summary = "" & DixSOb.Dix_Count(dix)
	Dim detail As String: detail = "" & SOb.Obj_FormatFields0( _
		"Keys", "Collection[" & DixSOb.Dix_Keys(dix).Count & "]", _
		"Items", "Collection[" & DixSOb.Dix_Items(dix).Count & "]", _
		"Count", DixSOb.Dix_Count(dix) _
	)
	Dim preview As Boolean: preview = True
	Dim indent As String: indent = VBA.vbTab  ' & "----"
	Dim orphan As Boolean: orphan = True
	
	Debug.Print ">> Obj_Print(dix, ...)"
	Debug.Print
	SOb.Obj_Print dix, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		summary := summary, _
		details := detail, _
		preview := preview, _
		indent := indent, _
		orphan := orphan
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	Debug.Print ">> Dix_Print(dix, ...)"
	Debug.Print
	DixSOb.Dix_Print dix, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		preview := preview, _
		indent := indent, _
		orphan := orphan
End Sub



' ...

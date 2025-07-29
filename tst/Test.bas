Attribute VB_Name = "Test"

' ###############
' ## Constants ##
' ###############

' Class name for simulated Dix object.
Const CLS_DIX As String = "Dix"



' ##################
' ## Enumerations ##
' ##################

' Field support for simulated Dix object.
Private Enum DixField
	keys
	items
	count
End Enum



' #############
' ## Testing ##
' #############

' .
Private Sub Test()
	' Dim dix As Object:	 SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Collection: SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Object:	 Set dix = SOb.New_Obj("Dix")
	Dim dix As Object: Set dix = New_Dix()
	
	SOb.Obj_Field(dix, DixField.count) = 42
	
	
	Dim copy As String: SOb.Obj_ClassKey copy
	Debug.Print "Obj_ClassKey() = """ & copy & """"
	Debug.Print "Obj_HasClass(dix) = " & SOb.Obj_HasClass(dix)
	Debug.Print "Obj_Class(dix) = """ & SOb.Obj_Class(dix) & """"
	
	Debug.Print
	
	Debug.Print "IsObj(dix) = " & SOb.IsObj(dix)
	Debug.Print "IsObj(dix, """ & CLS_DIX & """) = " & SOb.IsObj(dix, CLS_DIX)
	Debug.Print "IsObj(dix, ""Other"") = " & SOb.IsObj(dix, "Other")
	
	Debug.Print
	
	SOb.Obj_FieldKey copy, DixField.count  ' obj := dix
	Debug.Print "Obj_FieldKey(DixField.Count) = """ & copy & """"
	Debug.Print "Obj_HasField(dix, DixField.Count) = " & SOb.Obj_HasField(dix, DixField.count)
	Debug.Print "Obj_Field(dix, DixField.Count) = " & SOb.Obj_Field(dix, DixField.count)
End Sub


' .
Private Sub Test_Format()
	Dim flds() As Variant: flds = Array( _
		"x", "True", _
		"y", "1", _
		"z", """one""" _
	)
	
	Dim name As String: name = "Dix"
	Dim dep As Integer: dep = 1
	Dim pln As Boolean: pln = True
	Dim ptr As String: ptr = "1234567890"
	Dim sum As String: sum = "1:9"
	Dim dtl As String: dtl = Obj_FormatFields(flds)
	Dim ind As String: ind = "----"
	
	' Debug.Print SOb.Obj_FormatDetails()
	' Debug.Print SOb.Obj_FormatDetails(flds)
	Debug.Print SOb.Obj_FormatStr( _
		name:=name, _
		dep:=dep, _
		pln:=pln, _
		ptr:=ptr, _
		sum:=sum, _
		dtl:=dtl, _
		ind:=ind _
	)
End Sub



' ...



' ###################
' ## Dixionary SOB ##
' ###################

' Constructor.
Public Function New_Dix() As Object
	Const CLS_NAME = CLS_DIX
	
	Dim dix As Object: Set dix = SOb.New_Obj(CLS_NAME)
	SOb.Dix_Initialize dix
	
	Set New_Dix = dix
End Function


' Initializer.
Private Sub Dix_Initialize(ByRef dix As Object)
	Const CLS_NAME = CLS_DIX
	SOb.Obj_Initialize dix, CLS_NAME
	
	If Not SOb.Obj_HasField(dix, DixField.keys) Then
		Dim keys As Collection: Set keys = New Collection
		Set SOb.Obj_Field(dix, DixField.keys) = keys
	End If
	
	If Not SOb.Obj_HasField(dix, DixField.items) Then
		Dim items As Collection: Set items = New Collection
		Set SOb.Obj_Field(dix, DixField.items) = items
	End If
End Sub


' ' Validator.
' Private Function Dix_Validate(ByRef dix As Object) Then
'	 ' ...
' End Function

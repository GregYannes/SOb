Attribute VB_Name = "Test"



' ###############
' ## Constants ##
' ###############

' Class name for simulated Dix object.
Private Const CLS_DIX As String = "Dix"



' ##################
' ## Enumerations ##
' ##################

' Field support for simulated Dix object.
Private Enum DixField
	Keys
	Items
	Count
End Enum



' #############
' ## Testing ##
' #############

' Test the lifecycle of a simulated "Dix" object.
Private Sub Test()
	' Dim dix As Object:	 SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Collection: SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Object:	 Set dix = SOb.New_Obj("Dix")
	Dim dix As Object: Set dix = New_Dix()
	
	SOb.Obj_Field(dix, DixField.Count) = 42
	
	
	Dim copy As String: SOb.Obj_ClassKey copy
	Debug.Print "Obj_ClassKey() = """ & copy & """"
	Debug.Print "Obj_HasClass(dix) = " & SOb.Obj_HasClass(dix)
	Debug.Print "Obj_Class(dix) = """ & SOb.Obj_Class(dix) & """"
	
	Debug.Print
	
	Debug.Print "IsObj(dix) = " & SOb.IsObj(dix)
	Debug.Print "IsObj(dix, """ & CLS_DIX & """) = " & SOb.IsObj(dix, CLS_DIX)
	Debug.Print "IsObj(dix, ""Other"") = " & SOb.IsObj(dix, "Other")
	
	Debug.Print
	
	SOb.Obj_FieldKey copy, DixField.Count  ' obj := dix
	Debug.Print "Obj_FieldKey(DixField.Count) = """ & copy & """"
	Debug.Print "Obj_HasField(dix, DixField.Count) = " & SOb.Obj_HasField(dix, DixField.Count)
	Debug.Print "Obj_Field(dix, DixField.Count) = " & SOb.Obj_Field(dix, DixField.Count)
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
' ## Dixionary SOb ##
' ###################

' ##############################
' ## Dixionary SOb | Creation ##
' ##############################

' Constructor.
Public Function New_Dix() As Object
	Const CLS_NAME As String = CLS_DIX
	
	Dim dix As Object: Set dix = SOb.New_Obj(CLS_NAME)
	Dix_Initialize dix
	
	Set New_Dix = dix
End Function


' Initializer.
Private Sub Dix_Initialize(ByRef dix As Object)
	Const CLS_NAME As String = CLS_DIX
	
	SOb.Obj_Initialize dix, CLS_NAME
	
	If Not SOb.Obj_HasField(dix, DixField.Keys) Then
		Dim keys As Collection: Set keys = New Collection
		Set Dix_Keys(dix) = keys
	End If
	
	If Not SOb.Obj_HasField(dix, DixField.Items) Then
		Dim items As Collection: Set items = New Collection
		Set Dix_Items(dix) = items
	End If
	
	' Initialize to count of keys.
	If Not SOb.Obj_HasField(dix, DixField.Count) Then
		Dim count As Long: count = Dix_Keys(dix).Count
		Dix_Count(dix) = count
	End If
End Sub



' ##############################
' ## Dixionary SOb | Typology ##
' ##############################

' Test for simulated "Dix" object.
Public Function IsDix(ByRef x As Variant) As Boolean
	Const CLS_NAME As String = CLS_DIX
	Const N_FLDS As String = 3
	
	IsDix = SOb.IsObj(x, CLS_NAME)
	
	If IsDix Then
		Dim obj As Object: Set obj = x
		IsDix = SOb.Obj_FieldCount(x) = N_FLDS
	End If
	
	If IsDix Then
		Dix_Keys
		Dix_Items
		Dix_Count
	End If
End Function


' Cast to simulated "Dix" object.
Public Function AsDix(ByRef x As Variant) As Object
	Const CLS_NAME As String = CLS_DIX
	
	Set AsDix = SOb.AsObj(x, CLS_NAME)
	
	Dix_Initialize AsDix
End Function



' ############################
' ## Dixionary SOb | Fields ##
' ############################

' The ".Keys" field: the user may neither read...
Private Property Get Dix_Keys(ByRef dix As Object) As Collection
	Set Dix_Keys = SOb.Obj_Field(dix, DixField.Keys)
End Property

' ...nor write.
Private Property Set Dix_Keys(ByRef dix As Object, _
	ByRef val As Collection _
)
	Set SOb.Obj_Field(dix, DixField.Keys) = val
End Property


' The ".Items" field: the user may neither read...
Private Property Get Dix_Items(ByRef dix As Object) As Collection
	Set Dix_Items = SOb.Obj_Field(dix, DixField.Keys)
End Property

' ...nor write.
Private Property Set Dix_Items(ByRef dix As Object, _
	ByRef val As Collection _
)
	Set SOb.Obj_Field(dix, DixField.Keys) = val
End Property


' The ".Count" property: the user may read...
Public Property Get Dix_Count(ByRef dix As object) As Long
	Dix_Count = SOb.Obj_Field(dix, DixField.Count)
	' Dix_Count = Dix_Keys(dix).Count
End Property

' ...but not write.
Private Property Let Dix_Count(ByRef dix As Object, _
	ByVal val As Long _
)
	SOb.Obj_Field(dix, DixField.Count) = val
End Property



' #############################
' ## Dixionary SOb | Methods ##
' #############################

' ...




' ###################################
' ## Dixionary SOb | Visualization ##
' ###################################

' Print a simulated "Dix" object.
Public Function Dix_Print(ByRef dix As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	Dix_Print = Dix_Format(dix, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		preview := preview, _
		indent := indent, _
		orphan := orphan _
	)
	
	Debug.Print Dix_Print
End Function


' Format a simulated "Dix" object for printing.
Public Function Dix_Format(ByRef dix As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	Dix_Format = SOb.Obj_Format(dix, _
		sum := Dix_Count(dix), _
		dtl := SOb.Obj_FormatFields0( _
			"Keys",  "Collection[" & Dix_Keys(dix).Count & "]", _
			"Items", "Collection[" & Dix_Items(dix).Count & "]", _
			"Count", Dix_Count(dix) _
		), _
		dep := depth, _
		pln := plain, _
		ptr := pointer, _
		pvw := preview, _
		ind := indent, _
		orf := orphan _
	)
End Function

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
Private Enum Dix_Field
	Keys
	Items
	Count
End Enum



' #############
' ## Testing ##
' #############

' Test the lifecycle of a simulated "Dix" object.
Public Sub Test()
	' ### Construction ###
	' Dim dix As Object:	 SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Collection: SOb.Obj_Initialize dix, "Dix"
	' Dim dix As Object:	 Set dix = SOb.New_Obj("Dix")
	Dim dix As Object: Set dix = New_Dix()
	
	
	' ### Classification ###
	Dim copy As String: SOb.Obj_ClassKey copy
	Debug.Print "Obj_ClassKey() = """ & copy & """"
	Debug.Print "Obj_HasClass(dix) = " & SOb.Obj_HasClass(dix)
	Debug.Print "Obj_Class(dix) = """ & SOb.Obj_Class(dix) & """"
	
	Debug.Print
	
	
	' ### Typology 1 ###
	Debug.Print "IsObj(dix) = " & SOb.IsObj(dix)
	Debug.Print "IsObj(dix, """ & CLS_DIX & """) = " & SOb.IsObj(dix, CLS_DIX)
	Debug.Print "IsObj(dix, ""Other"") = " & SOb.IsObj(dix, "Other")
	
	Debug.Print
	
	
	' ### Fields 1 ###
	SOb.Obj_Field(dix, Dix_Field.Count) = 42
	
	SOb.Obj_FieldKey copy, Dix_Field.Count
	Debug.Print "Obj_FieldKey(Dix_Field.Count) = """ & copy & """"
	Debug.Print "Obj_HasField(dix, Dix_Field.Count) = " & SOb.Obj_HasField(dix, Dix_Field.Count)
	Debug.Print "Obj_Field(dix, Dix_Field.Count) = " & SOb.Obj_Field(dix, Dix_Field.Count)
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	
	' ### Typology 2 ###
	Debug.Print "IsDix(dix) = " & IsDix(dix)
	
	
	' ### Fields 2 ###
	Dix_Count(dix) = 7
	Debug.Print "Dix_Count(dix) = " & Dix_Count(dix)
	
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
	Dim summary As String: summary = "" & Dix_Count(dix)
	Dim detail As String: detail = "" & SOb.Obj_FormatFields0( _
		"Keys", "Collection[" & Dix_Keys(dix).Count & "]", _
		"Items", "Collection[" & Dix_Items(dix).Count & "]", _
		"Count", Dix_Count(dix) _
	)
	Dim preview As Boolean: preview = True
	Dim indent As String: indent = VBA.vbTab  ' & "----"
	Dim orphan As Boolean: orphan = True
	
	Debug.Print ">> Obj_Print(dix, ...)"
	Debug.Print
	SOb.Obj_Print dix, _
		dep := depth, _
		pln := plain, _
		ptr := pointer, _
		sum := summary, _
		dtl := detail, _
		pvw := preview, _
		ind := indent, _
		orf := orphan
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	Debug.Print ">> Dix_Print(dix, ...)"
	Debug.Print
	Dix_Print dix, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		preview := preview, _
		indent := indent, _
		orphan := orphan
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
	
	If Not SOb.Obj_HasField(dix, Dix_Field.Keys) Then
		Dim keys As Collection: Set keys = New Collection
		Set Dix_Keys(dix) = keys
	End If
	
	If Not SOb.Obj_HasField(dix, Dix_Field.Items) Then
		Dim items As Collection: Set items = New Collection
		Set Dix_Items(dix) = items
	End If
	
	' Initialize to count of keys.
	If Not SOb.Obj_HasField(dix, Dix_Field.Count) Then
		Dim count As Long: count = Dix_Keys(dix).Count
		Dix_Count(dix) = count
	End If
End Sub



' ##############################
' ## Dixionary SOb | Typology ##
' ##############################

' Test for simulated "Dix" object.
Public Function IsDix(ByRef x As Variant, _
	Optional ByVal strict As Boolean = False _
) As Boolean
	Const CLS_NAME As String = CLS_DIX
	
	' Ensure an accurate class with its proper set of fields...
	IsDix = SOb.IsObj(x, cls := CLS_NAME, strict := strict, flds := Array( _
		Dix_Field.Keys, _
		Dix_Field.Items, _
		Dix_Field.Count _
	))
	If Not IsDix Then Exit Function
	
	' ...and that their accessors work.
	On Error GoTo CHECK_ERROR
	If IsDix Then SOb.Obj_Check _
		Dix_Keys(obj), _
		Dix_Items(obj), _
		Dix_Count(obj)
	
	' Return the result in lieu of errors.
	Exit Function
	
' Handle inaccessibility.
CHECK_ERROR:
	IsDix = SOb.Obj_Error(typ := True)
End Function


' Cast to simulated "Dix" object.
Public Function AsDix(ByRef x As Variant) As Object
	' Cast the input to a (generic) simulated object...
	Dim obj As Object: Set obj = SOb.AsObj(x)
	
	' ...and extract its fields into a new "Dix" object.
	Set AsDix = New_Dix()
	
	Set Dix_Keys(AsDix) = Dix_Keys(obj)
	Set Dix_Items(AsDix) = Dix_Items(obj)
	Let Dix_Count(AsDix) = Dix_Count(obj)
End Function



' ############################
' ## Dixionary SOb | Fields ##
' ############################

' The ".Keys" field: the user may neither read...
Private Property Get Dix_Keys(ByRef dix As Object) As Collection
	SOb.Obj_Get Dix_Keys, dix, Dix_Field.Keys
End Property

' ...nor write.
Private Property Set Dix_Keys(ByRef dix As Object, _
	ByRef val As Collection _
)
	Set SOb.Obj_Field(dix, Dix_Field.Keys) = val
End Property


' The ".Items" field: the user may neither read...
Private Property Get Dix_Items(ByRef dix As Object) As Collection
	SOb.Obj_Get Dix_Items, dix, Dix_Field.Items
End Property

' ...nor write.
Private Property Set Dix_Items(ByRef dix As Object, _
	ByRef val As Collection _
)
	Set SOb.Obj_Field(dix, Dix_Field.Items) = val
End Property


' The ".Count" property: the user may read...
Public Property Get Dix_Count(ByRef dix As object) As Long
	SOb.Obj_Get Dix_Count, dix, Dix_Field.Count
	' Dix_Count = Dix_Keys(dix).Count
End Property

' ...but not write.
Private Property Let Dix_Count(ByRef dix As Object, _
	ByVal val As Long _
)
	SOb.Obj_Field(dix, Dix_Field.Count) = val
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

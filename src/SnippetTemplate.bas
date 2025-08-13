' ###############
' ## Constants ##
' ###############

' Class name for the simulated "*" object.
' TODO: Name your object.
Private Const *_CLS As String = "*"





' ...





' ##################
' ## Enumerations ##
' ##################

' Fields for the simulated "*" object.
Private Enum *__Field
	' TODO: Enumerate all fields in your object.
	FieldOne    ' The 1st field.
	FieldTwo    ' The 2nd field.
	FieldThree  ' The 3rd field.
	' ...
End Enum





' ...





' #######
' ## * ##
' #######

' ##################
' ## * | Creation ##
' ##################

' Construct a simulated "*" object.
Public Function New_*() As Object
	*_Initialize New_*
End Function



' Initialize a simulated "*" object.
Private Sub *_Initialize(ByRef * As Object)
	Const CLS_NAME As String = *_CLS
	
	Obj_Initialize *, CLS_NAME
	
	
	' TODO: Initialize any missing fields to appropriate values.
	If Not Obj_HasField(*, *__Field.FieldOne) Then
		Dim f1 As Boolean: ' Let f1 = ...
		Let *_FieldOne(*) = f1
	End If
	
	If Not Obj_HasField(*, *__Field.FieldTwo) Then
		Dim f2 As Range: ' Set f2 = ...
		Set *_FieldTwo(*) = f2
	End If
	
	If Not Obj_HasField(*, *__Field.FieldThree) Then
		Dim f3 As Variant: ' Assign f3, ...
		Assign *_FieldThree(*), f3
	End If
	
	' ...
End Sub



' ##################
' ## * | Typology ##
' ##################

' Identify a simulated "*" object.
Public Function Is*(ByRef x As Variant, _
	Optional ByVal strict As Boolean = False _
) As Boolean
	Const CLS_NAME As String = *_CLS
	
	
	' ### Class and Fields ###
	
	' Ensure an accurate class with its proper set of fields.
	' TODO: List all fields for "*" within this 'Array()'.
	Is* = IsObj(x, class := CLS_NAME, strict := strict, fields := Array( _
		*__Field.FieldOne, _
		*__Field.FieldTwo, _
		*__Field.FieldThree _
		 _
		 _
		 _
	))
	If Not Is* Then Exit Function
	
	
	' ' ### Accessors ###
	' 
	' ' Treat as an object moving forward.
	' Dim obj As Object: Set obj = x
	' 
	' ' Ensure the field accessors all work.
	' On Error GoTo CHECK_ERROR
	' 
	' ' TODO: Call all your field accessors within this 'Check()'.
	' If Is* Then Obj_Check _
	' 	*_FieldOne(obj), _
	' 	*_FieldTwo(obj), _
	' 	*_FieldThree(obj) _
	' 	 _
	' 	 _
	' 	 _
	' 
	' On Error GoTo 0
	' If Not Is* Then Exit Function
	' 
	' 
	' ' ...
	' If Not Is* Then Exit Function
	' 
	' 
	' ' TODO: Any further validation you desire.
	' 
	' 
	' ' ...
	' If Not Is* Then Exit Function
	
	
	' Return the result in lieu of errors.
	Exit Function
	
' ' Handle inaccessibility.
' CHECK_ERROR:
' 	' TODO: Specify (TRUE) which validation errors (like type) you desire to report as FALSE.
' 	Is* = Obj_CheckError(type_ := True)
End Function



' Cast to a simulated "*" object.
Public Function As*(ByRef x As Variant) As Object
	' Cast the input to a (generic) simulated object...
	Dim obj As Object: Set obj = AsObj(x)
	
	' ...and extract its fields into a new "*" object.
	Set As* = New_*()
	
	' TODO: Assign each field from 'obj' to its corresponding field in 'As*'.
	Let *_FieldOne(As*) = *_FieldOne(obj)
	Set *_FieldTwo(As*) = *_FieldTwo(obj)
	Assign *_FieldThree(As*), *_FieldThree(obj)
	' ...
End Function



' ################
' ## * | Fields ##
' ################

' A simulated (scalar) field which your user may read AND write.
Public Property Get *_FieldOne(ByRef * As Object) As Boolean
	Obj_Get *_FieldOne, *, *__Field.FieldOne
End Property

Public Property Let *_FieldOne(ByRef * As Object, ByVal val As Boolean)
	Let Obj_Field(*, *__Field.FieldOne) = val
End Property



' A simulated (objective) field which your user may read but NOT write.
Public Property Get *_FieldTwo(ByRef * As Object) As Range
	Obj_Get *_FieldTwo, *, *__Field.FieldTwo
End Property

Private Property Set *_FieldTwo(ByRef * As Object, ByRef val As Range)
	Set Obj_Field(*, *__Field.FieldTwo) = val
End Property



' A simulated (variant) field which your user may NEITHER read NOR write.
Private Property Get *_FieldThree(ByRef * As Object) As Variant
	Obj_Get *_FieldThree, *, *__Field.FieldThree
End Property

Private Property Let *_FieldThree(ByRef * As Object, ByVal val As Variant)
	Let Obj_Field(*, *__Field.FieldThree) = val
End Property

Private Property Set *_FieldThree(ByRef * As Object, ByRef val As Object)
	Set Obj_Field(*, *__Field.FieldThree) = val
End Property



' TODO: Accessors for any further fields you enumerated.
' ...



' #################
' ## * | Methods ##
' #################

' ' An external method which your user may call for a return value.
' Public Function *_MethodOne(ByRef * As Object, _
' 	 _
' 	 _
' 	 _
' ) As Integer
' 	' ...
' 	
' 	' Let MethodOne = ...
' End Function



' ' An internal method which your user may NOT call.
' Private Sub *_MethodTwo(ByRef * As Object, _
' 	 _
' 	 _
' 	 _
' )
' 	' ...
' End Sub



' TODO: Procedures for any further methods you desire.
' ...



' #######################
' ## * | Visualization ##
' #######################

' Print a simulated "*" object.
Public Function *_Print(ByRef * As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	*_Print = *_Format(*, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		preview := preview, _
		indent := indent, _
		orphan := orphan _
	)
	
	Obj_Print0 *_Print
End Function



' Format a simulated "*" object for printing.
Public Function *_Format(ByRef * As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	' TODO: Create any 'summary' or 'details' you desire for 'Obj_Format()'.
	' ...
	
	' Adjust settings to your satisfaction.
	' TODO: Pass any such 'summary' or 'details' to 'Obj_Format()'.
	*_Format = Obj_Format(*, _
		 _
		 _
		 _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		pvw := preview, _
		ind := indent, _
		orf := orphan _
	)
End Function

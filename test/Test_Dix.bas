Attribute VB_Name = "Test_Dix"



' ###############
' ## Constants ##
' ###############

' Class name for the simulated "Dix" object.
Private Const Dix_CLS As String = "Dix"





' ...





' ##################
' ## Enumerations ##
' ##################

' Fields for the simulated "Dix" object.
Private Enum Dix_Field
	' TODO: Enumerate all fields in your object.
	FieldOne    ' The 1st field.
	FieldTwo    ' The 2nd field.
	FieldThree  ' The 3rd field.
	' ...
End Enum





' ...





' #######
' ## Dix ##
' #######

' ##################
' ## Dix | Creation ##
' ##################

' Construct a simulated "Dix" object.
Public Function New_Dix() As Object
	Dix_Initialize New_Dix
End Function



' Initialize a simulated "Dix" object.
Public Sub Dix_Initialize(ByRef Dix As Object)
	Const CLS_NAME As String = Dix_CLS
	
	SOb.Obj_Initialize Dix, CLS_NAME
	
	
	' TODO: Initialize any missing fields to appropriate values.
	If Not SOb.Obj_HasField(Dix, Dix_Field.FieldOne) Then
		Dim f1 As Boolean: ' Let f1 = ...
		Let Dix_FieldOne(Dix) = f1
	End If
	
	If Not SOb.Obj_HasField(Dix, Dix_Field.FieldTwo) Then
		Dim f2 As Range: ' Set f2 = ...
		Set Dix_FieldTwo(Dix) = f2
	End If
	
	If Not SOb.Obj_HasField(Dix, Dix_Field.FieldThree) Then
		Dim f3 As Variant: ' SOb.Assign f3, ...
		SOb.Assign Dix_FieldThree(Dix), f3
	End If
	
	' ...
End Sub



' ##################
' ## Dix | Typology ##
' ##################

' Identify a simulated "Dix" object.
Public Function IsDix(ByRef x As Variant, _
	Optional ByVal strict As Boolean = False _
) As Boolean
	Const CLS_NAME As String = Dix_CLS
	
	
	' ### Class and Fields ###
	
	' Ensure an accurate class with its proper set of fields.
	' TODO: Enumerate all fields for "Dix" within this 'Array()'.
	IsDix = SOb.IsObj(x, cls := CLS_NAME, strict := strict, flds := Array( _
		Dix_Field.FieldOne, _
		Dix_Field.FieldTwo, _
		Dix_Field.FieldThree _
		 _
		 _
		 _
	))
	If Not IsDix Then Exit Function
	
	
	' ' ### Accessors ###
	' 
	' ' Treat as an object moving forward.
	' Dim obj As Object: Set obj = x
	' 
	' ' Ensure the field accessors all work.
	' On Error GoTo CHECK_ERROR
	' 
	' ' TODO: Call all your field accessors within this 'Check()'.
	' If IsDix Then SOb.Obj_Check _
	' 	Dix_FieldOne(obj), _
	' 	Dix_FieldTwo(obj), _
	' 	Dix_FieldThree(obj) _
	' 	 _
	' 	 _
	' 	 _
	' 
	' On Error GoTo 0
	' If Not IsDix Then Exit Function
	' 
	' 
	' ' ...
	' If Not IsDix Then Exit Function
	' 
	' 
	' ' TODO: Any further validation you desire.
	' 
	' 
	' ' ...
	' If Not IsDix Then Exit Function
	
	
	' Return the result in lieu of errors.
	Exit Function
	
' ' Handle inaccessibility.
' CHECK_ERROR:
' 	IsDix = SOb.Obj_Error(typ := True)
End Function



' Cast to a simulated "Dix" object.
Public Function AsDix(ByRef x As Variant) As Object
	' Cast the input to a (generic) simulated object...
	Dim obj As Object: Set obj = SOb.AsObj(x)
	
	' ...and extract its fields into a new "Dix" object.
	Set AsDix = New_Dix()
	
	' TODO: Assign each field from 'obj' to its corresponding field in 'AsDix'.
	Let Dix_FieldOne(AsDix) = Dix_FieldOne(obj)
	Set Dix_FieldTwo(AsDix) = Dix_FieldTwo(obj)
	SOb.Assign Dix_FieldThree(AsDix), Dix_FieldThree(obj)
	' ...
End Function



' ################
' ## Dix | Fields ##
' ################

' A simulated (scalar) field which your user may read AND write.
Public Property Get Dix_FieldOne(ByRef Dix As Object) As Boolean
	Let Dix_FieldOne = SOb.Obj_Field(Dix, Dix_Field.FieldOne)
End Property

Public Property Let Dix_FieldOne(ByRef Dix As Object, ByVal val As Boolean)
	Let SOb.Obj_Field(Dix, Dix_Field.FieldOne) = val
End Property



' A simulated (objective) field which your user may read but NOT write.
Public Property Get Dix_FieldTwo(ByRef Dix As Object) As Range
	Set Dix_FieldTwo = SOb.Obj_Field(Dix, Dix_Field.FieldTwo)
End Property

Private Property Set Dix_FieldTwo(ByRef Dix As Object, ByRef val As Range)
	Set SOb.Obj_Field(Dix, Dix_Field.FieldTwo) = val
End Property



' A simulated (variant) field which your user may NEITHER read NOR write.
Private Property Get Dix_FieldThree(ByRef Dix As Object) As Variant
	SOb.Assign Dix_FieldThree, SOb.Obj_Field(Dix, Dix_Field.FieldThree)
End Property

Private Property Let Dix_FieldThree(ByRef Dix As Object, ByVal val As Variant)
	Let SOb.Obj_Field(Dix, Dix_Field.FieldThree) = val
End Property

Private Property Set Dix_FieldThree(ByRef Dix As Object, ByRef val As Object)
	Set SOb.Obj_Field(Dix, Dix_Field.FieldThree) = val
End Property



' TODO: Accessors for any further fields you enumerated.
' ...



' #################
' ## Dix | Methods ##
' #################

' ' An external method which your user may call for a return value.
' Public Function Dix_MethodOne(ByRef Dix As Object, _
' 	 _
' 	 _
' 	 _
' ) As Integer
' 	' ...
' 	
' 	' Let MethodOne = ...
' End Function



' ' An internal method which your user may NOT call.
' Private Sub Dix_MethodTwo(ByRef Dix As Object, _
' 	 _
' 	 _
' 	 _
' )
' 	' ...
' End Sub



' TODO: Procedures for any further methods you desire.
' ...



' #######################
' ## Dix | Visualization ##
' #######################

' Print a simulated "Dix" object.
Public Function Dix_Print(ByRef Dix As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	Dix_Print = Dix_Format(Dix, _
		depth := depth, _
		plain := plain, _
		pointer := pointer, _
		preview := preview, _
		indent := indent, _
		orphan := orphan _
	)
	
	SOb.Obj_Print0 Dix_Print
End Function



' Format a simulated "Dix" object for printing.
Public Function Dix_Format(ByRef Dix As Object, _
	Optional ByVal depth = 1, _
	Optional ByVal plain As Boolean = False, _
	Optional ByVal pointer As Boolean = False, _
	Optional ByVal preview As Boolean = False, _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal orphan As Boolean = True _
) As String
	' TODO: Create any summary ('sum') or detail ('dtl') you desire for 'Obj_Format()'.
	' ...
	
	' Adjust settings to your satisfaction.
	' TODO: Pass any such summary ('sum') or detail ('dtl') to 'Obj_Format()'.
	Dix_Format = SOb.Obj_Format(Dix, _
		 _
		 _
		 _
		dep := depth, _
		pln := plain, _
		ptr := pointer, _
		pvw := preview, _
		ind := indent, _
		orf := orphan _
	)
End Function

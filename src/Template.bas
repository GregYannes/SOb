' ###############
' ## Constants ##
' ###############

' Class name for simulated "*" object.
Private Const CLS_* As String = "*"



' ...



' ##################
' ## Enumerations ##
' ##################

' Fields for simulated "*" object.
Private Enum *_Field
	Field1
	Field2
	' ...
End Enum



' ...



' #######
' ## * ##
' #######

' ##################
' ## * | Creation ##
' ##################

' Constructor.
Public Function New_*() As Object
	Const CLS_NAME As String = *_CLS
	
	*_Initialize New_*, CLS_NAME
	
	' ...
End Function


' Initializer.
Public Sub *_Initialize(ByRef * As Object)
	Const CLS_NAME As String = *_CLS
	
	Obj_Initialize *, CLS_*
	
	' ...
End Sub



' ##################
' ## * | Typology ##
' ##################

' Identify simulated "*" objects.
Public Function Is*(ByRef x As Variant) As Boolean
	Const CLS_NAME As String = *_CLS
	
	Is* = IsObj(x, CLS_NAME)
	
	' ...
End Function


' Cast to simulated "*" object.
Public Function As*(ByRef x As Variant) As Object
	Const CLS_NAME As String = *_CLS
	
	Set As* = AsObj(x, CLS_NAME)
	
	' ...
End Function



' ################
' ## * | Fields ##
' ################

' The scalar "Field1" field: the user may read it...
Public Property Get *_Field1(ByRef * As Object) As Integer
	Let *_Field1 = Obj_Field(*, *_Field.Field1)
End Property

' ...and also write it.
Public Property Let *_Field1(ByRef * As Object, ByVal val As Integer)
	Let Obj_Field(*, *_Field.Field1) = val
End Property


' The objective "Field2" field: the user may read it...
Public Property Get *_Field2(ByRef * As Object) As Range
	Set *_Field2 = Obj_Field(*, *_Field.Field2)
End Property

' ...but never write it.
Private Property Set *_Field2(ByRef * As Object, ByRef val As Range) As Range
	Set Obj_Field(*, *_Field.Field2) = val
End Property


' ...



' #################
' ## * | Methods ##
' #################

' .
Function *_Method1(ByRef * As Object, ...) As Variant
	' ...
End Function


' .
Sub *_Method2(ByRef * As Object, ...)
	' ...
End Sub


' ...



' #######################
' ## * | Visualization ##
' #######################

' .
Public Function *_Print(ByRef * As Object, ...) As Variant
	' ...
End Function


' .
Public Function *_Format(ByRef * As Object, ...) As String
	' ...
End Sub

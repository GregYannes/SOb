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
	*_Initialize New_*
End Function


' Initializer.
Public Sub *_Initialize(ByRef * As Object)
	Obj_Initialize *, CLS_*
	
	' ...
End Sub



' ##################
' ## * | Typology ##
' ##################

' Identifier.
Public Function Is*(ByRef x As Variant) As Boolean
	Is* = IsObj(x, CLS_*)
	
	' ...
End Function


' Caster.
Public Function As*(ByRef x As Variant) As Object
	Set As* = AsObj(x, CLS_*)
	
	' ...
End Function



' ################
' ## * | Fields ##
' ################

' The "Field1" field.
Public Property Get *_Field1(ByRef * As Object) As Integer
	Let *_Field1 = Obj_Field(*, *_Field.Field1)
End Property

Public Property Let *_Field1(ByRef * As Object, ByVal val As Integer)
	Let Obj_Field(*, *_Field.Field1) = val
End Property


' The "Field2" field.
Public Property Get *_Field2(ByRef * As Object) As Range
	Set *_Field2 = Obj_Field(*, *_Field.Field2)
End Property

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

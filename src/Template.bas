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
Private Enum *Fields
	Fld1
	Fld2
	' ...
	FldN
End Enum



' ...



' #######
' ## * ##
' #######

' Constructor.
Public Function New_*() As Object
	*_Initialize New_*
End Function


' Initializer.
Public Sub *_Initialize(ByRef * As Object)
	Obj_Initialize *, CLS_*
	
	' ...
End Sub


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

' The "Fld1" field.
Public Property Get *_Fld1(ByRef * As Object) As Integer
	Let *_Field1 = Obj_Field(*, *Fields.Fld1)
End Property

Public Property Let *_Fld1(ByRef * As Object, ByVal val As Integer)
	Let Obj_Field(*, *Fields.Fld1) = val
End Property


' The "Fld2" field.
Public Property Get *_Fld2(ByRef * As Object) As Range
	Set *_Fld2 = Obj_Field(*, *Fields.Fld2)
End Property

Public Property Set *_Fld2(ByRef * As Object, ByRef val As Range) As Range
	Set Obj_Field(*, *Fields.Fld2) = val
End Property


' ...



' #################
' ## * | Methods ##
' #################

' .
Public Function *_Mtd1(ByRef * As Object, ...) As Variant
	' ...
End Function


' .
Public Sub *_Mtd2(ByRef * As Object, ...)
	' ...
End Sub

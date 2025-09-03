Attribute VB_Name = "Ex_Validation"



' ##################
' ## Enumerations ##
' ##################

Private Enum Foo__Fields
	Bar
	Baz
	Qux
End Enum



' ################
' ## Procedures ##
' ################

Property Get Foo_Bar(foo As Object) As Integer
	Obj_Get Foo_Bar, foo, Bar
End Property

Property Let Foo_Bar(foo As Object, val As Integer)
	Let Obj_Field(foo, Bar) = val
End Property


Property Get Foo_Baz(foo As Object) As String
	Obj_Get Foo_Baz, foo, Baz
End Property

Property Let Foo_Baz(foo As Object, val As String)
	Let Obj_Field(foo, Baz) = val
End Property


Property Get Foo_Qux(foo As Object) As Range
	Obj_Get Foo_Qux, foo, Qux
End Property

Property Set Foo_Qux(foo As Object, val As Range)
	Set Obj_Field(foo, Qux) = val
End Property


Sub Check_Example(foo As Object)
	Debug.Print "Catching..."
	On Error GoTo CHECK_ERROR
	
	Debug.Print "Checking..."
	Obj_Check Foo_Bar(foo), Foo_Baz(foo), Foo_Qux(foo)
	
	Debug.Print "Succeeding..."
	
CHECK_ERROR:
	Debug.Print "Handling..."
	Debug.Print Obj_CheckError(type_ := True)
End Sub



' ##############
' ## Examples ##
' ##############

Public Sub Validation()
	Debug.Print "################"
	Debug.Print "## Validation ##"
	Debug.Print "################"
	
	
	Debug.Print
	Debug.Print "```"
	VALIDATION_1_START:
Dim foo As Object: Set foo = New_Obj("Foo")

Foo_Bar(foo) = 10
Foo_Baz(foo) = "Twenty"
Set Foo_Qux(foo) = [A1:B2]

Check_Example foo
	VALIDATION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VALIDATION_2_START:
Obj_Field(foo, Bar) = "Ten"

Check_Example foo
	VALIDATION_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VALIDATION_3_START:
Set Obj_Field(foo, Bar) = New Collection

Check_Example foo
	VALIDATION_3_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VALIDATION_4_START:
Set Obj_Field(foo, Qux) = New Collection

Check_Example foo
	VALIDATION_4_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	VALIDATION_5_START:
	Debug.Print "Catching..."
	On Error GoTo CHECK_ERROR
	
	Debug.Print "Checking..."
	Dim num As Double: num = 1 / 0
	
	Debug.Print "Succeeding..."
	
CHECK_ERROR:
	Debug.Print "Handling..."
	Debug.Print Obj_CheckError(type_ := True)
	VALIDATION_5_END:
	Debug.Print "```"
End Sub

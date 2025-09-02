Attribute VB_Name = "Ex_Field"



' ##################
' ## Enumerations ##
' ##################

Enum Foo__Fields
	Bar
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



' ##############
' ## Examples ##
' ##############

Public Sub Field()
	Debug.Print "###########"
	Debug.Print "## Field ##"
	Debug.Print "###########"
	
	Debug.Print
	Debug.Print "############################"
	Debug.Print "## Field | Backend Access ##"
	Debug.Print "############################"
	
	
	Debug.Print
	Debug.Print "```"
	BACKEND_ACCESS_1_START:
Debug.Print "Creating..."
Dim foo1 As Object: Set foo1 = New_Obj("Foo")
Debug.Print Obj_Field(foo1, Bar)

Debug.Print "Initializing..."
Obj_Field(foo1, Bar) = 42
Debug.Print Obj_Field(foo1, Bar)
	BACKEND_ACCESS_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "####################################"
	Debug.Print "## Creation | Implement Accessors ##"
	Debug.Print "####################################"
	
	
	Debug.Print
	Debug.Print "```"
	IMPLEMENT_ACCESSORS_1_START:
Foo_Bar(foo1) = -1
Debug.Print Foo_Bar(foo1)
	IMPLEMENT_ACCESSORS_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	IMPLEMENT_ACCESSORS_2_START:
Foo_Bar(foo1) = "Forty-two"
	IMPLEMENT_ACCESSORS_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	IMPLEMENT_ACCESSORS_3_START:
Dim foo2 As Object: Set foo2 = New_Obj("Foo")
Debug.Print Foo_Bar(foo2)
	IMPLEMENT_ACCESSORS_3_END:
	Debug.Print "```"
End Sub

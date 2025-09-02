Attribute VB_Name = "Ex_Fields"



' ##################
' ## Enumerations ##
' ##################

Private Enum Foo__Fields
	Bar
	Baz
	Qux
End Enum



' ##############
' ## Examples ##
' ##############

Public Sub Fields()
	Debug.Print "############"
	Debug.Print "## Fields ##"
	Debug.Print "############"
	
	Debug.Print
	Debug.Print "###########################"
	Debug.Print "## Fields | Count Fields ##"
	Debug.Print "###########################"
	
	
	Debug.Print
	Debug.Print "```"
	COUNT_FIELDS_1_START:
Debug.Print "Creating..."
Dim foo As Object: Set foo = New_Obj("Foo")
Debug.Print Obj_FieldCount(foo)

Debug.Print "Initializing..."
Obj_Field(foo, Bar) = 10
Obj_Field(foo, Baz) = "Twenty"
Obj_Field(foo, Qux) = 30
Debug.Print Obj_FieldCount(foo)
	COUNT_FIELDS_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	COUNT_FIELDS_2_START:
Debug.Print "Removing..."
foo.Remove foo.Count
Debug.Print Obj_Field(foo, Qux)
Debug.Print Obj_FieldCount(foo)

Debug.Print "Adding..."
foo.Add "IMPOSTER"
Debug.Print Obj_FieldCount(foo)
	COUNT_FIELDS_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "##############################"
	Debug.Print "## Fields | Check Existence ##"
	Debug.Print "##############################"
	
	
	Debug.Print
	Debug.Print "```"
	CHECK_EXISTENCE_1_START:
Debug.Print Obj_HasField(foo, Bar)
Debug.Print Obj_HasField(foo, Baz)
Debug.Print Obj_HasField(foo, Qux)
	CHECK_EXISTENCE_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	CHECK_EXISTENCE_2_START:
Dim f1 As Variant: f1 = Array(Bar, Baz)
Debug.Print Obj_HasFields(foo, f1)

Dim f2 As Variant: f2 = Array(Bar, Baz, Qux)
Debug.Print Obj_HasFields(foo, f2)
	CHECK_EXISTENCE_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	CHECK_EXISTENCE_3_START:
Debug.Print Obj_HasFields0(foo, Bar, Baz)
Debug.Print Obj_HasFields0(foo, Bar, Baz, Qux)
	CHECK_EXISTENCE_3_END:
	Debug.Print "```"
End Sub

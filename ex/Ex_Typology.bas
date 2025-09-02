Attribute VB_Name = "Ex_Typology"



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

Public Sub Typology()
	Debug.Print "##############"
	Debug.Print "## Typology ##"
	Debug.Print "##############"
	
	Debug.Print
	Debug.Print "#########################"
	Debug.Print "## Typology | Creation ##"
	Debug.Print "#########################"
	
	
	Debug.Print
	Debug.Print "```"
	TYPOLOGY_CREATION_1_START:
Dim foo As Object
Debug.Print Obj_Class(foo)

Set foo = New_Obj("Foo")
Debug.Print Obj_Class(foo)
	TYPOLOGY_CREATION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "######################################"
	Debug.Print "## Typology | Simple Identification ##"
	Debug.Print "######################################"
	
	
	Debug.Print
	Debug.Print "```"
	TYPOLOGY_SIMPLE_IDENTIFICATION_1_START:
Debug.Print IsObj(foo)

Dim clx As Collection, obj As Object
Debug.Print IsObj(clx), IsObj(obj)

Set clx = New Collection
Set obj = New Collection
Debug.Print IsObj(clx), IsObj(obj)
	TYPOLOGY_SIMPLE_IDENTIFICATION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	TYPOLOGY_SIMPLE_IDENTIFICATION_2_START:
Debug.Print IsObj(foo, "Foo")
Debug.Print IsObj(foo, "foO")
Debug.Print IsObj(foo, "Snaf")
	TYPOLOGY_SIMPLE_IDENTIFICATION_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "########################################"
	Debug.Print "## Typology | Enhanced Identification ##"
	Debug.Print "########################################"
	
	
	Debug.Print
	Debug.Print "```"
	TYPOLOGY_ENHANCED_IDENTIFICATION_1_START:
Dim aFields As Variant: aFields = Array(Bar, Baz, Qux)

Debug.Print "Some fields..."
Obj_Field(foo, Bar) = 10
Obj_Field(foo, Baz) = "Twenty"
Debug.Print IsObj(foo, "Foo", aFields)

Debug.Print "All fields..."
Obj_Field(foo, Qux) = 30
Debug.Print IsObj(foo, "Foo", aFields)
	TYPOLOGY_ENHANCED_IDENTIFICATION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	TYPOLOGY_ENHANCED_IDENTIFICATION_2_START:
Debug.Print "Exact..."
Debug.Print IsObj(foo, "Foo", aFields, strict := False)
Debug.Print IsObj(foo, "Foo", aFields, strict := True)

foo.Add "IMPOSTER"

Debug.Print "Extra..."
Debug.Print IsObj(foo, "Foo", aFields, strict := False)
Debug.Print IsObj(foo, "Foo", aFields, strict := True)
	TYPOLOGY_ENHANCED_IDENTIFICATION_2_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "########################"
	Debug.Print "## Typology | Casting ##"
	Debug.Print "########################"
	
	
	Debug.Print
	Debug.Print "```"
	TYPOLOGY_CASTING_1_START:
Debug.Print "Declaring..."
Dim cSnaf As Collection: Set cSnaf = New Collection
Dim oSnaf As Object: Set oSnaf = New Collection

Debug.Print IsObj(cSnaf, "Snaf"), IsObj(oSnaf, "Snaf")
Debug.Print Obj_Class(cSnaf), Obj_Class(oSnaf)
Debug.Print

Debug.Print "Casting..."
Set cSnaf = AsObj(cSnaf, "Snaf")
Set oSnaf = AsObj(oSnaf, "Snaf")

Debug.Print IsObj(cSnaf, "Snaf"), IsObj(oSnaf, "Snaf")
Debug.Print Obj_Class(cSnaf), Obj_Class(oSnaf)
	TYPOLOGY_CASTING_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	TYPOLOGY_CASTING_2_START:
Set foo = AsObj(foo, "Snaf")

Debug.Print IsObj(foo, "Snaf")
Debug.Print Obj_Class(foo)
	TYPOLOGY_CASTING_2_END:
	Debug.Print "```"
End Sub

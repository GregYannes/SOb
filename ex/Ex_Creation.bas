Attribute VB_Name = "Ex_Creation"



' ##############
' ## Examples ##
' ##############

Public Sub Creation()
	Debug.Print "##############"
	Debug.Print "## Creation ##"
	Debug.Print "##############"
	
	Debug.Print
	Debug.Print "#########################"
	Debug.Print "## Creation | Creation ##"
	Debug.Print "#########################"
	
	
	Debug.Print
	Debug.Print "```"
	CREATION_1_START:
Dim foo As Object: Set foo = New_Obj("Foo")

Debug.Print Obj_Class(foo)
	CREATION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "###############################"
	Debug.Print "## Creation | Initialization ##"
	Debug.Print "###############################"
	
	Debug.Print
	Debug.Print "```"
	INITIALIZATION_1_START:
Debug.Print "Declaring..."
Dim cSnaf As Collection, oSnaf As Object

Debug.Print cSnaf Is Nothing, oSnaf Is Nothing
Debug.Print Obj_Class(cSnaf), Obj_Class(oSnaf)
Debug.Print

Debug.Print "Initializing..."
Obj_Initialize cSnaf, "Snaf"
Obj_Initialize oSnaf, "Snaf"

Debug.Print cSnaf Is Nothing, oSnaf Is Nothing
Debug.Print Obj_Class(cSnaf), Obj_Class(oSnaf)
	INITIALIZATION_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	INITIALIZATION_2_START:
Obj_Initialize foo, "Snaf"

Debug.Print Obj_Class(foo)
	INITIALIZATION_2_END:
	Debug.Print "```"
End Sub

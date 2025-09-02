Attribute VB_Name = "Ex_Setup"



' ################
' ## Procedures ##
' ################

	Public Function Foo_Format(ByRef foo As Object, _
		Optional ByVal depth = 1, _
		Optional ByVal plain As Boolean = False, _
		Optional ByVal pointer As Boolean = False, _
		Optional ByVal preview As Boolean = False, _
		Optional ByVal indent As String = VBA.vbTab, _
		Optional ByVal orphan As Boolean = True _
	) As String
		Dim summary As String: summary = "..."
		Dim details As String: details = "..." & vbNewLine & "..." & vbNewLine & "..."
   	
		' Adjust settings to your satisfaction.
		Foo_Format = Obj_Format(foo, _
 			summary := summary, _
 			details := details, _
			depth := depth, _
			plain := plain, _
			pointer := pointer, _
			preview := preview, _
			indent := indent, _
			orphan := orphan _
		)
	End Function



' ##############
' ## Examples ##
' ##############

Public Sub Setup()
	Debug.Print "###########"
	Debug.Print "## Setup ##"
	Debug.Print "###########"
	
	Debug.Print
	Debug.Print "######################"
	Debug.Print "## Setup | Template ##"
	Debug.Print "######################"
	
	Debug.Print
	Debug.Print "################################"
	Debug.Print "## Setup | Template | TODO #7 ##"
	Debug.Print "################################"
	
	
	Debug.Print
	Debug.Print "```"
	SETUP_TEMPLATE_TODO_7_1_START:
Dim foo As Object: Set foo = New_Obj("Foo")
Debug.Print Foo_Format(foo, depth := 0)
	SETUP_TEMPLATE_TODO_7_1_END:
	Debug.Print "```"
	
	
	Debug.Print
	Debug.Print "```"
	SETUP_TEMPLATE_TODO_7_2_START:
Debug.Print Foo_Format(foo, depth := 1)
	SETUP_TEMPLATE_TODO_7_2_END:
	Debug.Print "```"
End Sub

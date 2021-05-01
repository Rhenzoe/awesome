
' ==== GENERIC CALL A FUNCTION SCRIPT ====

' Declare variant variables
Dim strPathVBS, strFunctionName, strArguments

' Resize dimension of an array to get the total number of parameters in a function
ReDim strArguments(WScript.Arguments.Count - 4)

' Assign variable to get "vbscript file path" which is Item(0)
strPathVBS = WScript.Arguments.Item(0)

' Condition statement to identify vbs function with parameters and without parameters
' If function has paramters then it will get each parameter in a function
If WScript.Arguments.Count > 3  Then

	' Assign variable to get "function name" in the vbscript which is Item(1)
	strFunctionName = WScript.Arguments.Item(1)
	j = 0

	For i = 2 To WScript.Arguments.Count - 2

		strArguments(j) = """" & WScript.Arguments.Item(i) & """"

		j = j + 1

	Next 

	' Concatenate function name and parameters
	strFunctionName = strFunctionName & " " & Join(strArguments, ", " )
	' MsgBox "FUNCTION WITH PARAMETERS"

Else

	' If function has no parameters then it will just get the function name which is Item(1)
	strFunctionName = WScript.Arguments.Item(1)
	' MsgBox "FUNCTION WITHOUT PARAMETERS"

End If	

	' Call StoreFunction to open the vbscript file and read all the information inside the script and perform execute global to be able to call any procedures (functions/Sub).
	StoreFunction(Trim(strPathVBS))
	' Execute the function in a form of a string from (Global stored string)
	Execute strFunctionName

'Function to store and read all the script from the vbscript file
Function StoreFunction(strScriptPath)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FileExists(strScriptPath) Then
	
		Set objFSOFile = objFSO.OpenTextFile(strScriptPath)
		ExecuteGlobal objFSOFile.ReadAll
		objFSOFile.Close
		
	End If
	
	Set objFSOFile = Nothing
	Set objFSO = Nothing 


End Function
	
	




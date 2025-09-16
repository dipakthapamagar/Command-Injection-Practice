Dim shell, command, output, exec, args, allowedCommands, userCommand

' Create a WScript.Shell object
Set shell = CreateObject("WScript.Shell")

' Function to display help options
Sub ShowHelp()
    WScript.Echo "Available commands:"
    WScript.Echo "1. ipconfig - Displays network configuration"
    WScript.Echo "2. ping [hostname] - Pings a specified hostname"
    WScript.Echo "3. exit - Quit the program"
End Sub

' Main loop
Do
    ' Display help options
    ShowHelp()

    ' Prompt for user input
    WScript.StdOut.Write "Enter a command: "
    userCommand = WScript.StdIn.ReadLine()

   'Array of allowed commands
    allowedCommands = Array("ipconfig", "ping", "exit")    

    Dim isAllowed, i
    isAllowed = False

    ' Check if user command contains command in allowedCommands array
    For i = LBound(allowedCommands) To UBound(allowedCommands)
    	  If LCase(usercommand) = allowedCommands(i) Or LCase(Left(userCommand, Len(allowedCommands(i)))) = allowedCommands(i) Then
    		  isAllowed = True
    		  Exit For
      	End If
    Next
    ' Check if the user wants to exit
    If LCase(userCommand) = "exit" Then
        WScript.Echo "Exiting the program."
        Exit Do
    End If

    If isAllowed Then
    	command = "cmd.exe /c " & userCommand

    	' Display loading message
    	WScript.Echo "Executing command, please wait..."

    	' Execute the command
    	Set exec = shell.Exec(command)

    	' Read the output
   	output = ""
    	Do While Not exec.StdOut.AtEndOfStream
        	output = output & exec.StdOut.ReadLine() & vbCrLf
    	Loop

    	' Display the output
    	WScript.Echo output

    	' Clean up
    	Set exec = Nothing
    Else
	WScript.Echo vbCrLf & "Command not supported." & vbCrLf
    End If
Loop

' Clean up
Set shell = Nothing

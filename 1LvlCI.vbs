Dim shell, command, output, exec, userCommand

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

    ' Check if the user wants to exit
    If LCase(userCommand) = "exit" Then
        WScript.Echo "Exiting the program."
        Exit Do
    End If

    ' Prepare the command
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
Loop

' Clean up
Set shell = Nothing

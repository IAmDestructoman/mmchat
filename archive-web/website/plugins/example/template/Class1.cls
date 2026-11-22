VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'OK. this is the template for a mmchat plugin
'one very important thing is that the name of the class MUST MUST MUST be plugin
'if you dont name it properly, the plugin WILL NOT load in the server.

Public Function GetName() As String
    GetName = "Your Plugin name goes here. " & App.Major & "." & App.Minor
End Function

Public Function GetCommands() As String
'this is the syntax to make the server recognize what commands are
'assoiated with your plugin.
'it MUST start and end with a , and all commands and perm lvls are seperated by ,
'the syntax is command permlvl command permlvl and so on
'so the following example has 3 commands.
'command1 is perm lvl 0
'command2 is perm lvl 4
'command3 is perm lvl 1
    
    GetCommands = ",command1,0,command2,4,command3,1,"
End Function

Public Function main(dadata As String, othervars As String) As String
'this is the main function called when a command is executed
'there can be multipule results seperated by ~
'the first char of each result tells if you should:
'0 chatper - result is sent to user who initiated command
'1 chatall not per - this text is sent to everyone on the server except the user who initiated command
'2 cchatall not per - this text is sent to everyone in the channel of the user who initiated command except the user
'3 chatall all - chats to everyone on server
'4 cchatall all - chats to everyone in users channel
'5 chatper to another user 5username,message - send text to another user
'6 chatper to another user 6ocn,message - same thing but uses the origional chat name instead of their current name

abold = Chr(27) & "[1m"
areset = Chr(27) & "[0m"
ablack = Chr(27) & "[30m"
ablue = Chr(27) & "[34m"
acyan = Chr(27) & "[36m"
agreen = Chr(27) & "[32m"
amagenta = Chr(27) & "[35m"
ared = Chr(27) & "[31m"
ayellow = Chr(27) & "[33m"
awhite = Chr(27) & "[37m"

'main code goes here

main = "0DIROCHAT PLUGINS ROCK"

'would simply spit back that text to anyone who typed in any of the plugins commands
'you have to look at the "dadata" string that is passed to get the actual command they typed as well as any paramaters

End Function

Function getvar(varname As String, othervars As String) As String

'this function simple gets string based variables from the main chat server.
'current available ones are:
'path       - returns the path of the server.. duh
'servername - the chatname of the server
'username   - current user name of the person who issued the command
'ocn        - the origional chat name of the userwho issued the command
'perm       - the permission level of the user who issued the command
'IF YOU NEED MORE VARIABLES PASSED TO THE PLUGIN.. CONTACT DIROCUS AND I CAN
'INCLUSE THAT VARIABLE IN THE NEXT SERVER RELEASE

If InStr(othervars, varname & "=") = 0 Then
    getvar = vbNullString
Else
    daloc = InStr(othervars, varname & "=")
    daloc2 = InStr(daloc, othervars, ",")
    getvar = Mid(othervars, daloc + Len(varname) + 1, daloc2 - daloc - Len(varname) - 1)
End If
End Function

Dim blnVariable

Set fso = CreateObject("Scripting.FileSystemObject")

Set ws = CreateObject("WScript.Shell")

Set dic = CreateObject("Scripting.Dictionary")

' True = create it even if it already exists; False = if it already exists do not overwrite
fso.CreateTextFile "C:\Users\everis\Desktop\Test.txt", True
WScript.Sleep(5000) ' Miliseconds

' False: 0. True: -1
blnVariable = fso.FileExists("C:\Users\everis\Desktop\Test.txt")
WScript.Sleep(10000) 

fso.CopyFile "C:\Users\everis\Desktop\Test.txt", "C:\Users\everis\Desktop\Test_2.txt"
WScript.Sleep(5000) 

fso.MoveFile "C:\Users\everis\Desktop\Test.txt", "C:\Users\everis\Documents\Test.txt"
WScript.Sleep(5000) 

fso.DeleteFile "C:\Users\everis\Documents\Test.txt", True

' Change Attributes
Set myAttributes = fso.GetFile("C:\Users\everis\Desktop\Test_2.txt")

' Normal = 0; Read = 1; Hidden = 2; System = 4; Volume = 8
myAttributes.Attributes = 0 ' 3 = Read + Hidden

' Size
Set mySize = fso.GetFile("C:\Users\everis\Desktop\Test_2.txt")

' WScript.Echo(fix(mySize.size/1020))

ws.Run "C:\Users\everis\Desktop\Test_2.txt"

' True = create it even if it already exists; False = if it already exists do not overwrite
Set myFile = fso.CreateTextFile("C:\Users\everis\Desktop\Test_3.txt", True)

myFile.Write("AQUI")

myFile.Close

WScript.Quit(1)
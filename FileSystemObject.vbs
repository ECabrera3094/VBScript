' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object

' Actions over the File
Set myFile = CreateObject("Scripting.FileSystemObject")

' True = create it even if it already exists; False = if it already exists do not overwrite
myFile.CreateTextFile "C:\Users\everis\Desktop\Test.txt", True



' Actions over Text insede the File
Set myTextStream = myFile.CreateTextFile("C:\Users\everis\Desktop\Test.txt", True)

myTextStream.Write("AQUI")

' Actios over the Folder 
Set myFolder = myFile.GetFolder("C:\Users\everis\Desktop")

' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object

' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/textstream-object

' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folder-object
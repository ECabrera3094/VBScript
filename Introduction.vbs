Option Explicit

Dim myMessage, intBD, strText, intAge_1, intAge_2, intIndex, myCall

intBD = 30

strText = "" & Chr(34) & "Yo naci el " & intBD & " de Abril" & Chr(34) & "."

WScript.Echo(strText)

intAge_1 = InputBox("Ingresa la Edad de Emiliano: ", "Emiliano", "Ingrese la Edad")
intAge_2 = InputBox("Ingresa la Edad de Gabriela: ", "Gabriela", "Ingrese la Edad")

if intAge_1 > intAge_2 Then
    WScript.Echo("El es Mayor de Edad.")
elseif intAge_1 < intAge_2 Then
    WScript.Echo("Ella es Mayor de Edad.")
elseif intAge_1 = intAge_2 Then
    WScript.Echo("Ambos tienen la Misma Edad.")
End if

intIndex = 10

Do
    WScript.Echo(intIndex)
    intIndex = intIndex - 1

Loop Until intIndex = 0


For intIndex = 10 To 0 Step -2

    WScript.Echo(intIndex)

Next

intIndex = 20

myCall = myfunction(intIndex)

Function myfunction(intIndex)

    WScript.Echo("Dentro de la Funcion. " & intIndex & ".")
    myfunction = intIndex + 100 ' Return

End Function

WScript.Echo(myCall)

myMessage = MsgBox("Mi Mensaje", 4, "Titulo")

WScript.Quit()

Class myExample
    
    Private m_userName

    ' Constructor
    Private Sub Class_Initialize()
        ' Can Be Empty
    End Sub

    ' Destructor
    Private Sub Class_Terminate()
        ' Can Be Empty
    End Sub

    Public Property Let UserName(strUserName)
        m_userName = strUserName
    End Property

    Public Property Get UserName
        'UserName = wscript.Echo(m_userName)
        UserName = m_userName ' Return
    End Property

End Class ' Test1

Dim objUser, strVariable

Set objUser = new myExample

objUser.UserName = InputBox("Ingrese su Nombre: ")

strVariable = objUser.UserName

WScript.Echo(strVariable)

' https://www.4guysfromrolla.com/webtech/092399-1.2.shtml

' https://www.devguru.com/content/technologies/vbscript/objects-class.html

' https://www.oreilly.com/library/view/designing-active-server/0596000448/ch04s02.html#:~:text=Inheritance%20is%20an%20OOP%20technique,to%20the%20relationships%20among%20objects.&text=Since%20VBScript%20classes%20are%20not,defined%20within%20that%20ASP%20script.
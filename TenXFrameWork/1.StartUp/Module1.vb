
Option Strict Off
Option Explicit On

Module SubMain

    Public objMain As MainCls
    Public IsNull As String
    Public GetDate As String
    Public HanaLen As String
    Public HanaInt As String
    Public Hana As String
    Public Sql As String
    Public Concate As String
    Public SelectCase As String
    Public DraftMessage As String
    Public Sub Main()
        objMain = New MainCls
        If (objMain.Initialise()) Then

            System.Windows.Forms.Application.Run()
            'Else
            objMain.objApplication.StatusBar.SetText("Error in Connection", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub

#Region "Close Application"
    Public Sub CloseApp()
        System.Windows.Forms.Application.Exit()
    End Sub
#End Region

End Module
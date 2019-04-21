VERSION 5.00
Begin VB.Form frmSysInfo 
   Caption         =   "Sys Info"
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "frmSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public i As Integer
Public MTime As Double
Public strId As String
Dim CON As New ADODB.Connection

Private Sub Form_Load()
        On Error GoTo Ext
        With CON
            If .State = 1 Then .Close
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & App.Path & "\SysInfo.mdb" & "';Mode=ReadWrite|Share Deny None;Persist Security Info=False;Jet OLEDB:Database Password="
            .Open
        End With
        strId = "": strId = AutoID()
        CON.Execute "INSERT INTO SysInfo (OoiID, OoiStartDate, OoiStartTime, OoiMinute) VALUES ('" & strId & "', #" & FormatDateTime(Date, vbShortDate) & "#, '" & FormatDateTime(Time, vbLongTime) & "', 0)"
        i = 0: Timer1.Enabled = True
        Exit Sub
Ext:
        End
End Sub

Private Sub Timer1_Timer()
        On Error Resume Next
        Dim rsCheck As New ADODB.Recordset
        i = i + 1
        If i = 150 Then
           MTime = MTime + 2.5
           CON.Execute "UPDATE SysInfo SET OoiExitDate = #" & FormatDateTime(Date, vbShortDate) & "#, OoiExitTime = '" & FormatDateTime(Time, vbLongTime) & "', OoiMinute = " & MTime & " WHERE OoiID = '" & strId & "'"
           i = 0
        End If
End Sub

Public Function SQL(strSQL As String, rs As Recordset) As Recordset
       Debug.Print strSQL
       If rs.State = adStateOpen Then
          rs.Close
       End If
       rs.ActiveConnection = CON
       rs.CursorLocation = adUseClient
       rs.CursorType = adOpenDynamic
       rs.LockType = adLockOptimistic
       rs.Source = strSQL
       rs.Open
       Set SQL = rs
End Function

Public Function AutoID() As String
       Dim sTr As String
       Dim rsID As New ADODB.Recordset
       sTr = "": sTr = "OOI" & "-" & Year(Now) & "-" & Format(Month(Date), "00") & "-"
       Set rsID = SQL("Select * from SysInfo where OoiID like '%" & sTr & "%' Order By Right(OoiID,5) DESC", rsID)
       If rsID.RecordCount > 0 Then
          sTr = sTr & Format(Right(rsID("OoiID"), 5) + 1, "00000")
       Else
          sTr = sTr & "00001"
       End If
       AutoID = sTr
End Function


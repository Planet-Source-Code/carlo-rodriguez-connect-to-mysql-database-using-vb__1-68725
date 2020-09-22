Attribute VB_Name = "mysqlConn"
Option Explicit
Public conn As New ADODB.Connection 'Connection to be used to mysql database

Public Function OpenConn(srvIP As String, dbNAme As String, dbUser As String, dbPass As String, dbPORT As Long)
    On Error GoTo errH
    
    If conn.State <> 0 Then conn.Close 'Check if currently connected if yes, disconnect.
    
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & srvIP & ";DATABASE=" & dbNAme & ";" _
                                 & "UID=" & dbUser & ";PWD=" & dbPass & "; PORT=" & dbPORT & "; OPTION=3"
    conn.Open  'open the connection
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "ERROR!"
        
End Function


Public Function getTables(listname As ListBox) 'Retrieve tables from database then add to list
Dim rs1 As New ADODB.Recordset

        
        Set rs1 = conn.OpenSchema(adSchemaTables)
        listname.Clear
        Do While Not rs1.EOF
                        
            Debug.Print rs1.Fields("TABLE_NAME")
            listname.AddItem (rs1.Fields("TABLE_NAME"))
            rs1.MoveNext
            
        Loop
    
End Function


Public Function getFields(rs As ADODB.Recordset, lstBox As ListBox) 'Retrieve fields then add to list
Dim numFlds As Integer
Dim i As Integer
    numFlds = rs.Fields.Count - 1
    lstBox.Clear
    For i = 0 To numFlds
        Debug.Print rs.Fields(i).Name
        lstBox.AddItem rs.Fields(i).Name
    Next i
    
End Function

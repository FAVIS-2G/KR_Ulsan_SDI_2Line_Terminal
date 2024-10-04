Attribute VB_Name = "ModWriteLog"
Option Explicit
Public Function ResultWriteLog(Data As String)

On Error GoTo err

Dim i, j As Integer
Dim SHDate As String
Dim SHTime As String
Dim strSaveData As String
Dim strField As String
Dim strSpec As String
Dim FileName_Result As String
Dim FileNumber_Result As Integer

    SHDate = Format(Date, "yy-mm-dd")
    SHTime = Format(Time, "hh:mm:ss")
    
    FileName_Result = App.path & "AutoRunLog" & ".csv"
    FileNumber_Result = FreeFile
        
    Open FileName_Result For Append As FileNumber_Result
        
        
        strSaveData = SHDate & "," & SHTime & "," & Data
        Print #FileNumber_Result, strSaveData
            
    Close #FileNumber_Result
    

Exit Function

err:

Close #FileNumber_Result

End Function

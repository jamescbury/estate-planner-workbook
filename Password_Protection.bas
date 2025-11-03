Attribute VB_Name = "Password_Protection"
Option Explicit

' Simple Workbook Password Protection Module
' Allows users to set/remove passwords on the entire workbook

Public Sub ShowPasswordSettings()
    ' Main function to handle workbook password - call this from your button
    
    If ActiveWorkbook.HasPassword Then
        RemoveWorkbookPassword
    Else
        SetWorkbookPassword
    End If
    
End Sub

Private Sub SetWorkbookPassword()
    ' Set password protection on the workbook
    
    Dim newPassword As String
    Dim confirmPassword As String
    
    newPassword = InputBox("Enter a password to protect this workbook:" & vbCrLf & vbCrLf & _
                          "?? Important: Write this password down! If you forget it, " & _
                          "you will not be able to open this file.", _
                          "Set Workbook Password")
    
    If newPassword = "" Then Exit Sub ' User cancelled
    
    confirmPassword = InputBox("Please confirm your password:", "Confirm Password")
    
    If newPassword <> confirmPassword Then
        MsgBox "Passwords do not match. Please try again.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo PasswordError
    
    ActiveWorkbook.SaveAs ActiveWorkbook.FullName, Password:=newPassword
    MsgBox "Workbook password protection has been set successfully!" & vbCrLf & vbCrLf & _
           "?? Remember: You will need this password to open the file in the future.", _
           vbInformation, "Password Set"
    
    On Error GoTo 0
    Exit Sub
    
PasswordError:
    MsgBox "Error setting password protection.", vbCritical
    On Error GoTo 0
    
End Sub

Private Sub RemoveWorkbookPassword()
    ' Remove password protection from the workbook
    
    Dim currentPassword As String
    
    If MsgBox("This workbook is password protected." & vbCrLf & vbCrLf & _
             "Do you want to remove the password protection?", _
             vbYesNo + vbQuestion, "Remove Password Protection") = vbNo Then
        Exit Sub
    End If
    
    currentPassword = InputBox("Enter the current password to remove protection:", _
                             "Remove Password Protection")
    
    If currentPassword = "" Then Exit Sub ' User cancelled
    
    On Error GoTo PasswordError
    
    ActiveWorkbook.SaveAs ActiveWorkbook.FullName, Password:=""
    MsgBox "Password protection has been removed successfully!", vbInformation
    
    On Error GoTo 0
    Exit Sub
    
PasswordError:
    MsgBox "Incorrect password. Cannot remove protection.", vbCritical
    On Error GoTo 0
    
End Sub


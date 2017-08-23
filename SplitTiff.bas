Attribute VB_Name = "SplitTiff"
Option Explicit

Public Sub splitTiff()

    Dim Img As ImageFile
    Dim myPage As ImageFile
    Dim v As Vector
    Dim lp As Long
    Dim strPath As String
    Dim intChoice As Integer
    
    Application.ScreenUpdating = False

        Set Img = New ImageFile
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        intChoice = Application.FileDialog(msoFileDialogOpen).Show
        
        If intChoice = False Then
            MsgBox "Operation Cancelled"
            Exit Sub
        End If
    
        If intChoice <> 0 Then
           strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
           Img.LoadFile strPath
                For lp = 1 To Img.FrameCount
                    Img.ActiveFrame = lp
                    Set v = Img.ARGBData
                    Set myPage = v.ImageFile(Img.Width, Img.Height)
                    myPage.SaveFile strPath & lp & ".tif"
                Next
        End If
        
        MsgBox "Done!"
    Application.ScreenUpdating = True
End Sub


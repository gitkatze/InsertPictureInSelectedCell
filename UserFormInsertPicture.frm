VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormInsertPicture 
   Caption         =   "画像の貼り付け"
   ClientHeight    =   5110
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   4400
   OleObjectBlob   =   "UserFormInsertPicture.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserFormInsertPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ThumbnailImageName As String = "ImageThumbnail"
Const NumberOfPictures As Integer = 4

Enum AlignVirtical
    Top
    Bottom
    Center
End Enum

Enum AlignHolizontal
    Left
    Right
    Center
End Enum

Dim PicturePaths(0 To NumberOfPictures) As String

Function GetThumbnailImageName(index)
    GetThumbnailImageName = ThumbnailImageName & index
End Function

Function GetPicturePathsCount()
    Dim i, count
    For i = 1 To NumberOfPictures
        If PicturePaths(i) <> "" Then count = count + 1
    Next
    GetPicturePathsCount = count
End Function

Function ClearPicturePaths()
    Dim i
    For i = LBound(PicturePaths) To UBound(PicturePaths)
        PicturePaths(i) = ""
    Next
    ClearThumbnails
End Function

Function ClearThumbnails()
    Dim i, obj, controlName

    For i = 1 To NumberOfPictures
        controlName = GetThumbnailImageName(i)
        
        Set obj = Me.Controls(controlName)
        Set obj.Picture = Nothing
    Next
End Function

Function UpdateThumbnailImages()
    Dim i
    Dim controlName As String
    Dim msg As String
    
    ClearThumbnails
    
    On Error GoTo UpdateThumbnailImages_Err
    
    Dim obj
    For i = 1 To NumberOfPictures
        Set obj = Nothing
        controlName = GetThumbnailImageName(i)
        Set obj = Me.Controls(controlName)
        If Not obj Is Nothing Then
            obj.Picture = LoadPicture(PicturePaths(i))
        End If
    Next
UpdateThumbnailImages_Err:
    Me.Repaint
    Set obj = Nothing
    If Err <> 0 Then
        msg = "Error in 'UpdateThumbnailImages'" & vbCrLf & "ErrorNumber: " & Err.Number & vbCrLf & vbCrLf & Err.Description
        MsgBox msg
        Err = 0
        Exit Function
    End If
    
End Function

Function SelectFile(Optional filter As String = "Excel File(*.xls),*.xls", Optional filterIndex As Integer = 1, Optional title As String = "ファイルを選択してください", Optional isMultiSelect As Boolean = False) As Variant
    SelectFile = Application.GetOpenFilename(filter, filterIndex, title, , isMultiSelect)
End Function

Function InsertPictureInRange(targetSheet As Worksheet, picturePath As String, targetRange As Range, virticalAlign As AlignVirtical, holizontalAlign As AlignHolizontal, Optional parcentageInRange As Double = 1#)
    If picturePath = "" Or Dir(picturePath) = "" Then Exit Function
    
    Dim pic As Object
    Dim picX, picY, picAspect
    Dim rangeWidth, rangeHeight, rangeAspect
    
    targetRange.Select
    rangeWidth = targetRange.width
    rangeHeight = targetRange.height
    
    Set pic = targetSheet.Pictures.Insert(picturePath)
    
    picAspect = pic.height / pic.width
    rangeAspect = rangeHeight / rangeWidth
        
    With pic.ShapeRange
        .LockAspectRatio = msoCTrue
        
        If picAspect > rangeAspect Then
            .height = rangeHeight * parcentageInRange
        Else
            .width = rangeWidth * parcentageInRange
        End If
        
        Select Case virticalAlign
            Case AlignVirtical.Top
                picY = 0
            Case AlignVirtical.Bottom
                picY = rangeHeight - .height
            Case Else
                picY = (rangeHeight - .height) / 2
        End Select
        Select Case holizontalAlign
            Case AlignHolizontal.Left
                picX = 0
            Case AlignHolizontal.Right
                 picX = rangeWidth - .width
            Case Else
                picX = (rangeWidth - .width) / 2
        End Select
        .IncrementLeft picX
        .IncrementTop picY
    End With
    
End Function

Function Insert4Pictures()
    InsertPictureInRange ActiveSheet, PicturePaths(1), Selection, AlignVirtical.Top, AlignHolizontal.Left, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(2), Selection, AlignVirtical.Top, AlignHolizontal.Right, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(3), Selection, AlignVirtical.Bottom, AlignHolizontal.Left, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(4), Selection, AlignVirtical.Bottom, AlignHolizontal.Right, 0.5
End Function

Private Sub ButtonClearThumbnail_Click()
    ClearPicturePaths
    Me.Repaint
End Sub

Private Sub buttonInsertPicture_Click()
    Dim pictureCount
    
    pictureCount = GetPicturePathsCount
    
    If pictureCount = 0 Then
        MsgBox "画像を選択してください"
        Exit Sub
    End If
    
    Dim obj, picturePath As String
    Dim i, j
    
    
    If LCase(TypeName(Selection)) = "range" Then
        
        If CheckBox4Corners.value Then
            Insert4Pictures
            Exit Sub
        End If
        
        Dim truncatedPaths(0 To NumberOfPictures) As String
        j = 1
        For i = 1 To NumberOfPictures
            If PicturePaths(i) <> "" Then
                truncatedPaths(j) = PicturePaths(i)
                j = j + 1
            End If
        Next
        
        Select Case pictureCount
        Case 1
            picturePath = truncatedPaths(1)
            If Dir(picturePath) = "" Then Exit Sub
            
            InsertPictureInRange ActiveSheet, picturePath, Selection, AlignVirtical.Center, AlignHolizontal.Center, 1
        Case 2
            If (PicturePaths(1) <> "" And PicturePaths(2) <> "") Or (PicturePaths(3) <> "" And PicturePaths(4) <> "") Then
                    InsertPictureInRange ActiveSheet, truncatedPaths(1), Selection, AlignVirtical.Center, AlignHolizontal.Left, 0.5
                    InsertPictureInRange ActiveSheet, truncatedPaths(2), Selection, AlignVirtical.Center, AlignHolizontal.Right, 0.5
            Else
                If (PicturePaths(1) <> "" And PicturePaths(3) <> "") Or (PicturePaths(2) <> "" And PicturePaths(4) <> "") Then
                    InsertPictureInRange ActiveSheet, truncatedPaths(1), Selection, AlignVirtical.Top, AlignHolizontal.Center, 0.5
                    InsertPictureInRange ActiveSheet, truncatedPaths(2), Selection, AlignVirtical.Bottom, AlignHolizontal.Center, 0.5
                Else
                    Insert4Pictures
                End If
            End If
        Case Else
            Insert4Pictures
        End Select
        
    End If
    
    Set obj = Nothing
End Sub

Private Sub ButtonSelectPicture_Click()
    Dim paths, path
    Dim i
    
    paths = SelectFile("jpeg,*.jpg;*.jpeg,gif,*.gif,tiff,*.tif;*.tiff", 1, , True)
    
    If IsArray(paths) Then
        ClearPicturePaths
        i = 0
        For Each path In paths
            i = i + 1
            If i > NumberOfPictures Then Exit For
            PicturePaths(i) = path
        Next
    Else
        Exit Sub
    End If
    
    UpdateThumbnailImages
End Sub

Function SelectImage(index)
    Dim path
    path = SelectFile("jpeg,*.jpg;*.jpeg,gif,*.gif,tiff,*.tif;*.tiff", 1, , False)
    If path <> False Then
        PicturePaths(index) = path
        UpdateThumbnailImages
    End If
End Function

Private Sub ImageThumbnail1_Click()
    SelectImage 1
End Sub

Private Sub ImageThumbnail2_Click()
    SelectImage 2
End Sub

Private Sub ImageThumbnail3_Click()
    SelectImage 3
End Sub

Private Sub ImageThumbnail4_Click()
    SelectImage 4
End Sub

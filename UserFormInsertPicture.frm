VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormInsertPicture 
   Caption         =   "画像の貼り付け"
   ClientHeight    =   5590
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

Enum ImagePlacementTypes
    One
    Two_Holizontal
    Two_Virtical
    Four
End Enum

Dim PicturePaths(0 To NumberOfPictures) As String
Dim ImagePlacementType As ImagePlacementTypes

Function GetThumbnailImageName(index)
    GetThumbnailImageName = ThumbnailImageName & index
End Function

Sub GetThumbnailImage(ByRef obj, index)
    Set obj = Me.Controls(GetThumbnailImageName(index))
End Sub

Function GetPicturePathsCount()
    Dim i, count
    For i = 1 To NumberOfPictures
        If PicturePaths(i) <> "" Then count = count + 1
    Next
    GetPicturePathsCount = count
End Function

Sub ClearPicturePaths()
    Dim i
    For i = LBound(PicturePaths) To UBound(PicturePaths)
        PicturePaths(i) = ""
    Next
    ClearThumbnails
End Sub

Sub ClearThumbnails()
    Dim i, obj, controlName

    For i = 1 To NumberOfPictures
        controlName = GetThumbnailImageName(i)
        
        Set obj = Me.Controls(controlName)
        Set obj.Picture = Nothing
        obj.ControlTipText = ""
    Next
End Sub

Sub SetImagePlacementType(ptype As ImagePlacementTypes)
    placementType = ptype
End Sub

Sub SetVisibleThumbnailsByPlacementType()
    Select Case ImagePlacementType
        Case ImagePlacementTypes.One
            SetVisibleThumbnails True, False, False, False
            
        Case ImagePlacementTypes.Two_Holizontal
            SetVisibleThumbnails True, True, False, False
            
        Case ImagePlacementTypes.Two_Virtical
            SetVisibleThumbnails True, False, True, False
            
        Case ImagePlacementTypes.Four
            SetVisibleThumbnails True, True, True, True
            
    End Select
End Sub

Sub SetVisibleThumbnails(ParamArray visibles())
    Dim i As Integer
    Dim j As Integer
    Dim obj
    j = 1
    For i = LBound(visibles) To UBound(visibles)
        If j > NumberOfPictures Then Exit For
        GetThumbnailImage obj, j
        If Not obj Is Nothing Then obj.Visible = visibles(i)
        j = j + 1
        Set obj = Nothing
    Next
End Sub

Sub UpdateThumbnailImages()
    Dim i
    Dim controlName As String
    Dim msg As String
    
    ClearThumbnails
    
    On Error GoTo UpdateThumbnailImages_Err
    
    Dim obj
    For i = 1 To NumberOfPictures
        Set obj = Nothing
        GetThumbnailImage obj, i
        If Not obj Is Nothing Then
            obj.Picture = LoadPicture(PicturePaths(i))
            obj.ControlTipText = PicturePaths(i)
        End If
    Next
UpdateThumbnailImages_Err:
    Me.Repaint
    Set obj = Nothing
    If Err <> 0 Then
        msg = "Error in 'UpdateThumbnailImages'" & vbCrLf & "ErrorNumber: " & Err.Number & vbCrLf & vbCrLf & Err.Description
        MsgBox msg
        Err = 0
        Exit Sub
    End If
    
End Sub

Sub SetThumbnailsVisible(ParamArray thumbnailVisibles())
    Dim i As Integer
    For i = 0 To UBound(thumbnailVisibles)
    
    Next
    
End Sub

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
        .LockAspectRatio = msoTrue
        
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

Sub Insert1Picture()
    InsertPictureInRange ActiveSheet, PicturePaths(1), Selection, AlignVirtical.Center, AlignHolizontal.Center, 1
End Sub

Sub Insert2PicturesHorizontal()
    InsertPictureInRange ActiveSheet, PicturePaths(1), Selection, AlignVirtical.Center, AlignHolizontal.Left, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(2), Selection, AlignVirtical.Center, AlignHolizontal.Right, 0.5
End Sub

Sub Insert2PicturesVirtical()
    InsertPictureInRange ActiveSheet, PicturePaths(1), Selection, AlignVirtical.Top, AlignHolizontal.Center, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(2), Selection, AlignVirtical.Bottom, AlignHolizontal.Center, 0.5
End Sub

Sub Insert4Pictures()
    InsertPictureInRange ActiveSheet, PicturePaths(1), Selection, AlignVirtical.Top, AlignHolizontal.Left, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(2), Selection, AlignVirtical.Top, AlignHolizontal.Right, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(3), Selection, AlignVirtical.Bottom, AlignHolizontal.Left, 0.5
    InsertPictureInRange ActiveSheet, PicturePaths(4), Selection, AlignVirtical.Bottom, AlignHolizontal.Right, 0.5
End Sub

Private Sub ButtonClearThumbnail_Click()
    If CheckBoxNotConfirmBeforeClear.Value = False Then
        If MsgBox("すべての画像をクリアしますか？", vbYesNo, "確認") <> vbYes Then Exit Sub
    End If
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
        
    If LCase(TypeName(Selection)) = "range" Then
                
        Select Case ImagePlacementType
        Case ImagePlacementTypes.One
            Insert1Picture
        
        Case ImagePlacementTypes.Two_Holizontal
            Insert2PicturesHorizontal
        
        Case ImagePlacementTypes.Two_Virtical
            Insert2PicturesVirtical
            
        Case ImagePlacementTypes.Four
            Insert4Pictures
            
        End Select
        
    End If
    
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

Private Sub ComboBoxPlacementType_Change()
    Select Case ComboBoxPlacementType.ListIndex
        Case 0
            ImagePlacementType = ImagePlacementTypes.One
            
        Case 1
            ImagePlacementType = ImagePlacementTypes.Two_Holizontal
        
        Case 2
            ImagePlacementType = ImagePlacementTypes.Two_Virtical
        
        Case 3
            ImagePlacementType = ImagePlacementTypes.Four
    End Select
    
    SetVisibleThumbnailsByPlacementType
End Sub

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

Private Sub UserForm_Initialize()

    ComboBoxPlacementType.Clear
    
    ComboBoxPlacementType.AddItem "1"
    ComboBoxPlacementType.AddItem "2 (横)"
    ComboBoxPlacementType.AddItem "2 (縦)"
    ComboBoxPlacementType.AddItem "4"
    
    ComboBoxPlacementType.ListIndex = 0
End Sub

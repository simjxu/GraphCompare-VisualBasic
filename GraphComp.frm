VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GraphComp 
   Caption         =   "Graph Compare"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3510
   OleObjectBlob   =   "GraphComp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GraphComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Private functions for getting the color of the clicked pixel.
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Type POINT
    x As Long
    y As Long
End Type

' Initialized static variables
Dim xvalarray() As String
Dim yvalarray() As String
Dim namearray() As String
Dim labelarray() As String
Dim colorarray() As Variant
Dim legendloc(1 To 4) As Double

' Pick your own color for the plots you want to look at
Private Sub ColorPalette_Click()
Dim plocation As POINT
Dim lColour, lDC As Long

lDC = GetWindowDC(0)
Call GetCursorPos(plocation)
lColour = GetPixel(lDC, plocation.x, plocation.y)

    With ListBox1
        cnt = 1
        For i = 1 To .ListCount
            If .Selected(i - 1) Then

                ActiveChart.SeriesCollection(i).Border.Color = lColour
                colorarray(i) = lColour

            End If
        Next i
    End With
End Sub

Sub UserForm_Initialize()
Dim seriescount As Integer
Dim serieslength As Double
Dim currstring As String

    ' Count number of series in the current chart
    seriescount = ActiveChart.SeriesCollection.count
    serieslength = UBound(ActiveChart.SeriesCollection(1).XValues)
    ReDim xvalarray(1 To seriescount) As String
    ReDim yvalarray(1 To seriescount) As String
    ReDim namearray(1 To seriescount) As String
    ReDim labelarray(1 To seriescount) As String
    ReDim colorarray(1 To seriescount) As Variant
    
    ' Save the data
    ' Sample Formula: SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$35,Sheet1!$B$2:$B$35,1)
    For i = 1 To seriescount
        currstring = ActiveChart.SeriesCollection(i).Formula
        spltstring = Split(currstring, ",")
        xvalarray(i) = spltstring(1)
        yvalarray(i) = spltstring(2)
        namearray(i) = spltstring(0)
        colorarray(i) = ActiveChart.SeriesCollection(i).Border.Color
        
        ' Label array depends on whether it references cell or typed in name
        If InStr(spltstring(0), "!") > 0 Then
            labelarray(i) = Range(Right(spltstring(0), Len(spltstring(0)) - 8)).Value
        Else
            temp = Right(spltstring(0), Len(spltstring(0)) - 9)
            labelarray(i) = Left(temp, Len(temp) - 1)
        End If
    Next i
    
    ' In case you get that error for automatically colored lines, use this color array
    'colorarray2 = Array(0, 9264438, 3422095, 4098930, 7816797, _
    9731891, 3042745, 9856058, 3685272, 4363385, 8342627, 10324022, 3241157, 10316093, _
    3882656, 4627584, 8802408, 10850106, 3439311, 10907457, 4145832, 4891782, 9262190, _
    11441725, 3637722, 11367492, 4343216, 5155725, 9721971, 11967808, 3835876, 11893063, _
    4540599, 5354131, 10116216, 12428355, 4034029, 12287562, 4737982, 5552536, 10510461, _
    12954182, 4231926, 12618597, 6514115, 7126690, 11039883, 13218146, 6200566, 13015419, _
    8027336, 8438955, 11634584, 13482105, 7776246, 13346190, 9277134, 9620149, 12162981, _
    13811852, 9089526, 13742238, 10329811, 10604733, 12690864, 14075804, 10140406, _
    14006954, 11119319, 11326661, 13152441, 14339497, 10994422, 14336951, 11974620, _
    12180172, 13614019, 14603190, 11913975, 14667203, 12764128, 12902356, 14075596, _
    14932418, 12702455, 14996941, 13487845, 13624539)
        
'    If seriescount > 87 Then
        
'        For i = 85 To seriescount
'            colorarray(i) = 13828936 - (i - 1) * 30000000
'        Next
        
'    End If
    
    ' Save legend location
    legendloc(1) = ActiveChart.Legend.Left
    legendloc(2) = ActiveChart.Legend.Top
    legendloc(3) = ActiveChart.Legend.Width
    legendloc(4) = ActiveChart.Legend.Height
    
    ' Populate listbox with name array
    ListBox1.List = labelarray
    
    
End Sub

Private Sub CommandButton1_Click()

Dim newseries(1 To 99) As Integer
Dim count As Integer
Dim cnt As Integer
Dim seriescount As Integer
Dim serieslength As Double
Dim NewString As String
Dim chttype As String

    
    ' Record Chart Type
    chttype = ActiveChart.ChartType
    
    ' Save the selections made in ListBox
    With ListBox1
        cnt = 1
        For i = 1 To .ListCount
            If .Selected(i - 1) Then

                ActiveChart.SeriesCollection(i).Border.Color = colorarray(i)

            Else
                ActiveChart.SeriesCollection(i).Format.Line.Visible = msoFalse
            End If
        Next i
    End With
    ActiveChart.Legend.Left = legendloc(1)
    ActiveChart.Legend.Top = legendloc(2)
    ActiveChart.Legend.Width = legendloc(3)
    ActiveChart.Legend.Height = legendloc(4)
    
'    ActiveChart.ChartType = chttype

End Sub

Private Sub CommandButton2_Click()
    
    With ListBox1
        For i = 1 To .ListCount
            ActiveChart.SeriesCollection(i).Format.Line.Visible = msoTrue
        Next i
    End With
    
    ActiveChart.ClearToMatchStyle
    
    Call UserForm_Initialize

End Sub




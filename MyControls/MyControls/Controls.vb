Option Strict On

Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports System.Runtime.InteropServices

Public Class TabBrowser
    Inherits System.Windows.Forms.Control
    Public Class PageChangingArgs : Inherits EventArgs
        Public NewIndex As Integer
        Public Sub New(ByVal tNewIndex As Integer)
            NewIndex = tNewIndex
        End Sub
    End Class

    Event AddBoxClicked(ByVal sender As TabBrowser, ByVal e As EventArgs)
    Event PageChanged(ByVal sender As TabBrowser, ByVal e As EventArgs)
    Event PageChanging(ByVal sender As TabBrowser, ByVal e As PageChangingArgs)
    Public Shared HighlightColor As SolidBrush, HighlightColor2 As SolidBrush
    Public BorderColor As Pen
    Public FontColor As SolidBrush
    Public BorderThickness As Integer
    Public Pages As New List(Of Page)
    Public CenterTabText As Boolean = False, EnableAddBox As Boolean = False, EnableExitButton As Boolean = False
    Public LastPage As Page
    Public HHeight As Integer, PaintCounter As Integer = 0
    Private TabHeight As Integer = 20, TotalHeight As Integer = 0
    Private TabWidth As Integer = 100, TabWidthMax As Integer = 100, TabWidthMin As Integer = 20, TabLBuffer As Integer = 3
    Private DoubleBorder As Integer
    Private ExitButtonXW As Integer = 7, ExitButtonXH As Integer = 7, ExitButtonW As Integer = 11, ExitButtonH As Integer = 11, ExitButtonTBuffer As Integer, ExitButtonLBuffer As Integer, ExitButtonColor As New Pen(Color.Black)
    Private HighlightedExitButton As Integer = -1, HighlightedHeader As Integer = -1
    Private AddBox As NewPageTab
    Private ScrollArrows(1) As ScrollArrow
    Private ControlBrush As New SolidBrush(Color.FromKnownColor(KnownColor.Control))
    Private ItemsShownMax As Integer = 0, TrueTabHeaderWidth As Integer = 0
    Private MinBuffer As Single, ScrollPosition As Integer = 0
    Private TabOrientation As Byte = 0 '0 = left oriented, 1 = right oriented
    Private OMGLOL As Integer = 0
    Private BackBuffer As Bitmap

    Private _SelectedIndex As Integer = -1
    Public Property SelectedIndex() As Integer
        Get
            Return _SelectedIndex
        End Get
        Set(ByVal value As Integer)
            RaiseEvent PageChanging(Me, New PageChangingArgs(value))
            If _SelectedIndex <> value Then
                If value = -1 Then
                    _SelectedIndex = -1
                    _SelectedPage = Nothing
                Else
                    Pages(value).Visible = True
                    If Not _SelectedPage Is Nothing Then Pages(_SelectedIndex).Visible = False
                    _SelectedIndex = value
                    _SelectedPage = Pages(value)
                    UpdateView()
                    RaiseEvent PageChanged(Me, New EventArgs)
                End If
            End If
        End Set
    End Property

    Private _SelectedPage As Page
    Public Property SelectedPage() As Page
        Get
            Return _SelectedPage
        End Get
        Set(ByVal value As Page)
            Dim indx As Integer = Pages.IndexOf(value)
            If indx <> -1 And indx <> _SelectedIndex Then
                SelectedIndex = indx
            End If
        End Set
    End Property

    Public Sub New(ByVal tbName As String, ByVal prnt As Control, ByVal l As Integer, ByVal t As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal sPage As String = "Blank", Optional ByVal tEnableAddBox As Boolean = True)
        Name = tbName
        If HighlightColor Is Nothing Then
            HighlightColor = New SolidBrush(Color.FromArgb(40, Color.LightBlue))
            HighlightColor2 = New SolidBrush(Color.FromArgb(80, Color.LightBlue))
        End If
        ScrollArrows(0) = New ScrollArrow(0, 0, TabHeight, 0)
        Parent = prnt
        Font = New Font("Microsoft Sans Serif", 8.25)
        FontColor = New SolidBrush(Parent.ForeColor)
        SetBorderStyle(1, Color.DarkBlue)
        RelocateTo(l, t, w, h)
        AddPage(sPage, Nothing)
    End Sub

    Public Sub RelocateTo(ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        If Pages.Count = 4 Then
            Dim omg2 As Integer = 3
        End If
        SetBounds(x, y, w, TabHeight)
        TotalHeight = h
        HHeight = h - TabHeight
        ExitButtonTBuffer = CInt((TabHeight - ExitButtonH) / 2) - 1
        TrueTabHeaderWidth = Width - (ScrollArrows(0).rect.Width * 2) - TabHeight
        Dim omg As Single = CSng(TrueTabHeaderWidth / TabWidthMin)
        ItemsShownMax = CInt(Math.Ceiling(omg))
        MinBuffer = CSng(Math.Ceiling(ItemsShownMax - omg)) * TabWidthMin
        If MinBuffer = 0 Then MinBuffer = TabWidthMin
        BackBuffer = New Bitmap(w, TabHeight)
        ScrollArrows(1) = New ScrollArrow(Width - (TabHeight * 2), 0, TabHeight, 0)
        SetAddBoxSettings(EnableAddBox)
        For Each pg As Page In Pages
            pg.SetBounds(Left, Bottom, Width, HHeight)
        Next
    End Sub

    Public Sub SetTabWidth(ByVal w As Integer)
        TabWidth = w
        TabWidthMax = w
        Me.Refresh()
    End Sub

    Public Sub SetWidth(ByVal tWidth As Integer)
        If tWidth = 0 Then Exit Sub
        Width = tWidth
        BackBuffer = New Bitmap(tWidth, TabHeight)
        For i As Integer = 0 To Pages.Count - 1
            Pages(i).Width = tWidth
        Next
    End Sub

    Public Sub SetBorderStyle(ByVal thickness As Integer, Optional ByVal clr As Color = Nothing)
        BorderThickness = thickness
        DoubleBorder = BorderThickness * 2
        If clr <> Nothing Then BorderColor = New Pen(clr)
    End Sub

    Public Sub SetAddBoxSettings(ByVal enabled As Boolean)

        EnableAddBox = enabled
        Dim vis As Boolean = False
        If Not ScrollArrows(1) Is Nothing Then
            vis = ScrollArrows(1).visible
        End If
        ScrollArrows(1).visible = vis
    End Sub

    Public Function IsSelected(ByVal tPage As Page) As Boolean
        If _SelectedPage Is tPage Then Return True
        Return False
    End Function

    Public Sub SetPageName(ByVal tPage As Page, ByVal sName As String)
        tPage.Header = New Page.TabHeader(tPage, sName)
        Me.Refresh()
    End Sub

    Public Sub AddPage(ByVal sName As String, Optional ByVal tFont As Font = Nothing, Optional ByVal tFontColor As Color = Nothing, Optional ByVal tBGColor As Color = Nothing, Optional ByVal bRefresh As Boolean = True, Optional ByVal bChangeSelection As Boolean = True)
        If Pages.Count = 1 AndAlso Pages(0).Header.Text = "Blank" AndAlso Pages(0).Controls.Count = 0 Then
            Pages(0).Dispose()
            If Not Pages(0) Is Nothing Then Pages(0) = Nothing
            Pages.RemoveAt(0)
        End If
        Pages.Add(New Page(Me, sName, tFont, tFontColor, tBGColor))
        LastPage = Pages(Pages.Count - 1)
        If _SelectedIndex = -1 Or _SelectedPage Is Nothing Then
            SelectedIndex = 0
        Else
            If bChangeSelection = True Then
                _SelectedIndex = Pages.Count - 1
            Else
                LastPage.Visible = False
            End If
        End If
        If ScrollArrows(0).visible = True AndAlso (_SelectedIndex * TabWidth) + TabWidth + ScrollPosition > ScrollArrows(1).rect.Left Then
            ScrollPosition = -1 * ((Pages.Count * TabWidth) - TrueTabHeaderWidth - ScrollArrows(0).rect.Width)
        End If
        UpdateView()
        If bRefresh = True Then Refresh()
    End Sub

    Public Sub SelectPage(ByVal index As Integer, Optional ByVal bRefresh As Boolean = True)
        SelectedIndex = index
        If bRefresh = True Then Me.Refresh()
    End Sub

    Private Sub TB_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles Me.MouseLeave
        Dim OHH As Integer = HighlightedHeader, OHEB As Integer = HighlightedExitButton
        HighlightedExitButton = -1
        HighlightedHeader = -1
        If HighlightedExitButton <> OHEB Or HighlightedHeader <> OHH Then Me.Refresh()
    End Sub

    Private Sub TB_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseMove
        If e.Y > TabHeight Then Exit Sub
        Dim index As Integer = CInt(Math.Floor((e.X - ScrollPosition) / TabWidth))
        If index >= Pages.Count Then
            If HighlightedHeader <> -1 Then
                HighlightedHeader = -1
                Me.Refresh()
            End If
            Exit Sub
        End If

        Dim OHH As Integer = HighlightedHeader
        HighlightedHeader = index
        If EnableExitButton Then
            Dim OHEB As Integer = HighlightedExitButton
            If e.X >= (index * TabWidth) + ExitButtonLBuffer + ScrollPosition AndAlso e.X <= (index * TabWidth) + TabWidth - 1 + ScrollPosition And e.Y >= ExitButtonTBuffer And e.Y <= ExitButtonTBuffer + ExitButtonH Then
                If HighlightedExitButton <> index Then HighlightedExitButton = index
            Else
                If HighlightedExitButton <> -1 Then HighlightedExitButton = -1
            End If
            If HighlightedExitButton <> OHEB Then
                Me.Refresh()
                Exit Sub
            End If
        End If
        If HighlightedHeader = _SelectedIndex Then
            If OHH <> -1 Then Me.Refresh()
            HighlightedHeader = -1
        ElseIf HighlightedHeader <> OHH Then
            Me.Refresh()
        End If
    End Sub

    Private Sub TB_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseDown
        Me.Focus()
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.Y > TabHeight Then Exit Sub
        If EnableAddBox AndAlso AddBox.rect.Contains(e.X, e.Y) Then
            RaiseEvent AddBoxClicked(Me, New EventArgs)
            Exit Sub
        End If
        If ScrollArrows(0).visible = True Then
            For i As Integer = 0 To 1
                If ScrollArrows(i).rect.Contains(e.X, e.Y) Then
                    If i = 0 Then 'left arrow clicked
                        If TabOrientation = 1 Then ScrollPosition += CInt(MinBuffer) Else ScrollPosition += TabWidthMin
                        ScrollPosition = Math.Min(ScrollPosition, ScrollArrows(0).rect.Width)
                        TabOrientation = 0
                    Else 'right arrow clicked
                        If TabOrientation = 0 Then ScrollPosition -= CInt(MinBuffer) Else ScrollPosition -= TabWidthMin
                        ScrollPosition = Math.Max(ScrollPosition, -1 * ((Pages.Count * TabWidth) - TrueTabHeaderWidth - ScrollArrows(0).rect.Width))
                        TabOrientation = 1
                    End If
                    Me.Refresh()
                    Exit Sub
                End If
            Next
        End If
        Dim index As Integer = CInt(Math.Floor((e.X - ScrollPosition) / TabWidth))
        If index >= Pages.Count Then Exit Sub
        If EnableExitButton = True AndAlso Pages.Count > 1 Then
            If e.X >= (index * TabWidth) + ExitButtonLBuffer + ScrollPosition AndAlso e.X <= (index * TabWidth) + TabWidth - 1 + ScrollPosition And e.Y >= ExitButtonTBuffer And e.Y <= ExitButtonTBuffer + ExitButtonH Then
                RemovePage(index)
                Exit Sub
            End If
        End If
        SelectPage(index, True)
    End Sub

    Public Sub RemovePage(ByVal index As Integer)
        Dim pg As Page = Pages(index)
        For Each cntrl As Control In pg.Controls
            RecursiveDispose(cntrl)
        Next

        Pages(index).Dispose()
        If Not Pages(index) Is Nothing Then Pages(index) = Nothing
        Pages.RemoveAt(index)
        LastPage = Pages(Pages.Count - 1)

        Dim newIndex As Integer = _SelectedIndex
        _SelectedIndex = -1

        If newIndex = index Then
            If index = 0 Then
                SelectPage(0, True)
            Else
                SelectPage(index - 1, True)
            End If
        Else
            If index < newIndex Then
                SelectPage(newIndex - 1, True)
            Else
                UpdateView()
                Me.Refresh()
            End If
        End If
    End Sub

    Private Sub RecursiveDispose(ByVal tControl As Control)
        For Each cntrl As Control In tControl.Controls
            RecursiveDispose(cntrl)
        Next
        tControl.Parent.Controls.Remove(tControl)
        tControl.Dispose()
        If Not tControl Is Nothing Then tControl = Nothing
    End Sub

    Private Sub UpdateView()
        Dim SM As Integer = TabWidthMax * Pages.Count + TabHeight - 1
        If SM > Width Then
            If EnableAddBox Then
                TabWidth = CInt((Width - TabHeight) / (TabWidthMax * Pages.Count)) * TabWidthMax
                AddBox = New NewPageTab(Width - TabHeight - 1, 0, TabHeight)
            Else
                TabWidth = CInt(Width / (TabWidthMax * Pages.Count)) * TabWidthMax
            End If
            If TabWidth < TabWidthMin Then
                TabWidth = TabWidthMin
                ScrollArrows(0).visible = True
                ScrollArrows(1).visible = True
                If ScrollPosition = 0 Then ScrollPosition = ScrollArrows(0).rect.Width
            Else
                If ScrollPosition <> 0 Then
                    ScrollArrows(0).visible = False
                    ScrollArrows(1).visible = False
                    ScrollPosition = 0
                End If
            End If
        Else
            TabWidth = TabWidthMax
            If EnableAddBox Then AddBox = New NewPageTab(Pages.Count * TabWidth, 0, TabHeight)
        End If
        ExitButtonLBuffer = TabWidth - ExitButtonW - 3
        If EnableAddBox AndAlso AddBox.rect.Right >= Width - 2 Then AddBox.rect.X = Width - AddBox.rect.Width - 1
    End Sub

    Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
        PaintCounter += 1
        Try
            Using gx As Graphics = Graphics.FromImage(BackBuffer)
                gx.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
                Dim CX As Integer = 1 + ScrollPosition
                gx.FillRectangle(ControlBrush, 0, 0, Width, TabHeight - 1)
                gx.DrawLine(BorderColor, 1, TabHeight - 1, Width - 1, TabHeight - 1) 'Bottom tab line
                For i As Integer = 0 To Pages.Count - 1
                    Dim pg As Page = Pages(i)
                    Dim rect As Rectangle = New Rectangle(CX, 1, TabWidth, TabHeight - 2)
                    If CX + TabWidth > 0 Then
                        gx.FillRectangle(pg.BGColor, rect)
                        gx.DrawString(pg.Header.Text, pg.Font, pg.Header.clr, rect.Left + pg.Header.TextRect.Left, pg.Header.TextRect.Top)
                        If _SelectedIndex = i Then
                            gx.FillRectangle(HighlightColor2, CX, 1, TabWidth, TabHeight - 2)
                        Else
                            If HighlightedHeader = i Then gx.FillRectangle(HighlightColor, CX, 1, TabWidth, TabHeight - 2)
                        End If
                        If EnableExitButton Then
                            Dim ER As Rectangle = New Rectangle(CX + ExitButtonLBuffer, ExitButtonTBuffer, ExitButtonW, ExitButtonH)
                            If HighlightedExitButton = i Then gx.DrawRectangle(BorderColor, ER)
                            gx.DrawLine(BorderColor, ER.Left + 2, ER.Top + 2, ER.Right - 2, ER.Bottom - 2)
                            gx.DrawLine(BorderColor, ER.Right - 2, ER.Top + 2, ER.Left + 2, ER.Bottom - 2)
                        End If
                    End If
                    CX += TabWidth
                    Try
                        If (ScrollArrows(0).visible = True AndAlso (CX > ScrollArrows(1).rect.Left Or CX > Width - AddBox.rect.Width)) Then
                        Else
                            gx.DrawLine(BorderColor, rect.Right - 1, 1, rect.Right - 1, TabHeight - 1)
                        End If
                    Catch ex As Exception

                    End Try
                Next
                If Not _SelectedPage Is Nothing Then gx.DrawLine(New Pen(_SelectedPage.BGColor.Color), _SelectedIndex * TabWidth + 1, TabHeight - 1, (_SelectedIndex * TabWidth) + TabWidth - 1, TabHeight - 1) 'removes the border line under selected tab header
                If Pages.Count * TabWidth >= Width Then gx.DrawLine(BorderColor, Width - 1, 0, Width - 1, Height - 1)
                gx.DrawLine(BorderColor, 1, 0, Pages.Count * TabWidth, 0) 'h line, top line of tab header
                For i As Integer = 0 To BorderThickness - 1
                    gx.DrawLine(BorderColor, i, 0, i, Height - 1)
                    gx.DrawLine(BorderColor, Width - 1 - i, TabHeight, Width - 1 - i, Height - 1)
                Next
                If EnableAddBox = True Then
                    If AddBox.rect.Right < Width - 1 Then gx.FillRectangle(ControlBrush, AddBox.rect.Right + 1, 0, Width - AddBox.rect.Right - 1, TabHeight - 1)
                    gx.FillRectangle(ControlBrush, AddBox.rect)
                    gx.DrawRectangle(BorderColor, AddBox.rect)
                    gx.DrawLine(BorderColor, AddBox.L1X1, AddBox.L1Y, AddBox.L1X2, AddBox.L1Y)
                    gx.DrawLine(BorderColor, AddBox.L2X, AddBox.L2Y1, AddBox.L2X, AddBox.L2Y2)
                End If
                If ScrollArrows(0).visible = True Then
                    For i As Integer = 0 To 1
                        gx.FillRectangle(ControlBrush, ScrollArrows(i).rect)
                        gx.DrawRectangle(BorderColor, ScrollArrows(i).rect)
                    Next
                    gx.DrawLine(BorderColor, ScrollArrows(0).X1, ScrollArrows(0).Y2, ScrollArrows(0).X2, ScrollArrows(0).Y1)
                    gx.DrawLine(BorderColor, ScrollArrows(0).X1, ScrollArrows(0).Y2, ScrollArrows(0).X2, ScrollArrows(0).Y3)
                    gx.DrawLine(BorderColor, ScrollArrows(1).X2, ScrollArrows(1).Y2, ScrollArrows(1).X1, ScrollArrows(1).Y1)
                    gx.DrawLine(BorderColor, ScrollArrows(1).X2, ScrollArrows(1).Y2, ScrollArrows(1).X1, ScrollArrows(1).Y3)
                End If
            End Using
            e.Graphics.DrawImageUnscaled(BackBuffer, 0, 0)
            If _SelectedIndex <> -1 AndAlso _SelectedPage.Visible = False Then
                _SelectedPage.Visible = True
                _SelectedPage.Refresh()
            End If
        Catch ex As Exception

        End Try
        MyBase.OnPaint(e)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
    End Sub

    Public Sub SetBoundsTo(ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        HHeight = h - TabHeight
        SetBounds(x, y, w, TabHeight)
        For i As Integer = 0 To Pages.Count - 1
            Pages(i).SetBoundsTo(Left, Bottom, Width, HHeight)
        Next
        _SelectedPage.Refresh()
    End Sub

    Public Shared Function MeasureAString(ByVal Width0Height1Both2 As Byte, ByVal str As String, ByVal tFont As Font) As Object
        Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
            Dim SFormat As New System.Drawing.StringFormat
            Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
            Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
            SFormat.SetMeasurableCharacterRanges(range)
            Dim regions() As Region = gx.MeasureCharacterRanges(str, tFont, rect, SFormat)
            rect = regions(0).GetBounds(gx)

            If Width0Height1Both2 = 0 Then Return rect.Right + 1 'gx.MeasureString(str, Font, 50000000, lolz).Width
            If Width0Height1Both2 = 1 Then Return rect.Bottom + 1 'gx.MeasureString(str, Font, 50000000, lolz).Height
            If Width0Height1Both2 = 2 Then Return New SizeF(rect.Right + 1, rect.Bottom + 1) 'gx.MeasureString(str, Font, 50000000, lolz)
        End Using
        Return -1
    End Function

    Public Class ScrollArrow
        Public rect As Rectangle, visible As Boolean
        Public X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer, Y3 As Integer
        Public Sub New(ByVal tX As Integer, ByVal tY As Integer, ByVal sz As Integer, ByVal iType As Integer)
            rect = New Rectangle(tX, tY, sz - 1, sz - 1)
            Dim midB As Integer = CInt(sz / 2)
            Dim lineSZ As Integer = CInt(sz / 4) - 2
            X1 = midB - lineSZ + tX
            X2 = midB + lineSZ + tX
            Y2 = midB + tY
            Y1 = midB - lineSZ + tY
            Y3 = midB + lineSZ + tY
            visible = False
        End Sub
    End Class

    Public Class NewPageTab
        Public rect As Rectangle
        Public L1X1, L1X2, L1Y, L2Y1, L2X, L2Y2 As Integer
        Public Sub New(ByVal x As Integer, ByVal y As Integer, ByVal sz As Integer)
            rect = New Rectangle(x, y, sz - 1, sz - 1)
            Dim midB As Integer = CInt(sz / 2)
            Dim lineSZ As Integer = CInt((sz - midB) / 2) - 1
            L1X1 = midB - lineSZ + x
            L1X2 = midB + lineSZ + x
            L1Y = midB + y
            L2X = midB + x
            L2Y1 = midB - lineSZ + y
            L2Y2 = midB + lineSZ + y
        End Sub
    End Class

    Public Class Page : Inherits System.Windows.Forms.Panel
        Implements IDisposable
        Public Event PageShown(ByVal sender As Object, ByVal e As EventArgs)

        Private Shadows disposed As Boolean = False
        Protected Overrides Sub Dispose( _
           ByVal disposing As Boolean)
            If Not Me.disposed Then
                If disposing Then
                    If TB._SelectedPage Is Me Then TB._SelectedPage = Nothing
                    If TB.LastPage Is Me Then TB.LastPage = Nothing
                    Header.Parent = Nothing
                    Header = Nothing
                    TB = Nothing
                    Parent = Nothing
                End If

                ' Free your own state (unmanaged objects).
                ' Set large fields to null.
            End If
            Me.disposed = True
        End Sub

#Region "Border"
        Private Const WM_NCCALCSIZE As Integer = 131
        Private Const WM_NCPAINT As Integer = 133
        Private Const WM_NCHITTEST As Integer = 132
        Private Const WM_NCLBUTTONDOWN As Integer = 161
        Private Const WM_NCMOUSEMOVE As Integer = &HA0

        <System.Runtime.InteropServices.DllImport("User32.dll")> _
        Friend Shared Function GetWindowDC(ByVal hWnd As IntPtr) As IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("User32.dll")> _
        Friend Shared Function ReleaseDC(ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Integer
        End Function

        Private Sub AdjustClientRect(ByRef rcClient As RECT)
            rcClient.Left += BorderThickness
            rcClient.Right -= BorderThickness
            rcClient.Bottom -= BorderThickness
        End Sub

        Private Structure RECT
            Public Left As Integer, Top As Integer, Right As Integer, Bottom As Integer
            Public Sub New(ByVal tLeft As Integer, ByVal tTop As Integer, ByVal tRight As Integer, ByVal tBottom As Integer)
                Left = tLeft : Top = tTop : Right = tRight : Bottom = tBottom
            End Sub
        End Structure

        Private Structure NCCALCSIZE_PARAMS
            Public rcNewWindow As RECT
            Public rcOldWindow As RECT
            Public rcClient As RECT
            Private lppos As IntPtr

            Public Sub New(trcNewWindow As RECT, trcOldWindow As RECT, trcClient As RECT, tlppos As IntPtr)
                rcNewWindow = trcNewWindow : rcOldWindow = trcOldWindow : rcClient = trcClient : lppos = tlppos
            End Sub
        End Structure

#End Region

        Public Header As TabHeader
        Public TB As TabBrowser
        Public BorderThickness As Integer = 1, BorderColor As Pen, WWidth As Integer, WHeight As Integer
        Public BGColor As SolidBrush, PaintCount As Integer = 0
        Public Shared Count As Integer = 0 ', bPaintOnce As Boolean = False

        Private DoubleBorder As Integer
        Private bCovered As Boolean = False

        Public Sub New(ByVal prnt As TabBrowser, ByVal sName As String, Optional ByVal tFont As Font = Nothing, Optional ByVal tFontColor As Color = Nothing, Optional ByVal tBGColor As Color = Nothing)
            SetStyle(ControlStyles.Selectable Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)
            Count += 1
            Name = prnt.Name & "Page" & Count
            Parent = prnt.Parent
            TB = prnt
            If tFont Is Nothing Then Font = Parent.Font Else Font = tFont
            If tFontColor = Nothing Then ForeColor = TB.FontColor.Color Else ForeColor = tFontColor
            If tBGColor = Nothing Then BGColor = New SolidBrush(Color.FromKnownColor(KnownColor.Control)) Else BGColor = New SolidBrush(tBGColor)
            BorderColor = TB.BorderColor

            DoubleBorder = BorderThickness * 2
            SetBoundsTo(TB.Left, TB.Bottom, TB.Width, TB.HHeight)
            Header = New TabHeader(Me, sName)
        End Sub

        Public Sub SetBoundsTo(ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
            WWidth = w - DoubleBorder
            WHeight = h - BorderThickness
            SetBounds(x, y, w, h)
        End Sub

        Public Sub SetName(ByVal sName As String)
            Header = New Page.TabHeader(Me, sName)
            TB.Refresh()
        End Sub

        Private Sub Page_ControlAdded(sender As Object, e As System.Windows.Forms.ControlEventArgs) Handles Me.ControlAdded
            If TypeOf e.Control Is TabBrowser AndAlso e.Control.Left = 0 AndAlso e.Control.Top = 0 Then 'container page
                BorderThickness = 0
            End If
        End Sub

        Private Sub Page_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
            Me.Focus()
        End Sub

        Protected Overrides Sub OnPaintBackground(e As System.Windows.Forms.PaintEventArgs)
            PaintCount += 1
            MyBase.OnPaintBackground(e)

        End Sub

        Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
        End Sub

        <DebuggerStepThrough()> _
        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            Select Case m.Msg
                Case WM_NCCALCSIZE
                    If m.WParam <> IntPtr.Zero Then
                        Dim rcsize As NCCALCSIZE_PARAMS = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(m.LParam, GetType(NCCALCSIZE_PARAMS)), NCCALCSIZE_PARAMS)
                        AdjustClientRect(rcsize.rcNewWindow)
                        System.Runtime.InteropServices.Marshal.StructureToPtr(rcsize, m.LParam, False)
                    Else
                        Dim rcsize As RECT = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(m.LParam, GetType(RECT)), RECT)
                        AdjustClientRect(rcsize)
                        System.Runtime.InteropServices.Marshal.StructureToPtr(rcsize, m.LParam, False)
                    End If
                    m.Result = New IntPtr(1)
                    Return
                Case WM_NCPAINT
                    Dim hdc As IntPtr = GetWindowDC(m.HWnd)
                    If hdc <> IntPtr.Zero Then
                        PaintBorder()
                    End If
                    Return
                Case Else
                    MyBase.WndProc(m)
            End Select
        End Sub

        Private Sub PaintBorder()
            Dim hDC As IntPtr = GetWindowDC(Me.Handle) '  m.HWnd)
            Dim g As Graphics = Graphics.FromHdc(hDC)
            For i As Integer = 0 To BorderThickness - 1
                g.DrawLine(BorderColor, i, 0, i, Height - 1)
                g.DrawLine(BorderColor, Width - 1 - i, 0, Width - 1 - i, Height - 1)
                g.DrawLine(BorderColor, BorderThickness, Height - 1 - i, Width - BorderThickness, Height - 1 - i)
            Next
            ReleaseDC(Me.Handle, hDC)
            g.Dispose()
        End Sub

        Public Class TabHeader
            Public Shared Height As Integer = 20
            Public Parent As Page
            Public Text As String, TextRect As Rectangle
            Public clr As SolidBrush

            Public Sub New(ByVal prnt As Page, ByVal sName As String)
                Parent = prnt
                Text = sName
                clr = New SolidBrush(Parent.ForeColor)
                Dim SS As SizeF = DirectCast(TabBrowser.MeasureAString(2, Text, Parent.Font), SizeF)
                Dim HBuffer As Integer = 1 + CInt((Parent.TB.TabHeight - SS.Height) / 2)
                Dim w As Integer = Parent.TB.TabWidth - Parent.TB.TabLBuffer - Parent.TB.ExitButtonXW - 3
                If Parent.TB.CenterTabText = True Then
                Else
                    TextRect = New Rectangle(Parent.TB.TabLBuffer, HBuffer, CInt(Math.Min(SS.Width, w)), CInt(SS.Height))
                End If
                If TextRect.Width = w Then
                    Dim pct As Single = w / SS.Width
                    Dim newLine As String = Text.Substring(0, CInt(Math.Min(pct * Text.Length + 1, Text.Length)))
                    If newLine.Length < 4 Then
                        newLine = Text.Chars(0) & "."
                        Text = newLine
                    Else
                        Text = newLine.Remove(newLine.Length - 3, 3) & "..."
                    End If
                End If
                TextRect.Width += 5
            End Sub
        End Class

        Private Sub Page_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
            If Visible Then RaiseEvent PageShown(Me, New EventArgs)
        End Sub
    End Class

End Class

Public Class ScrollBarSet
    Public Event ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public HSize As Integer = -1, VSize As Integer = -1
    Public WithEvents HBar As ScrollBars.HBar, VBar As ScrollBars.VBar

    Private _HBarEnabled As Boolean
    Public Property HBarEnabled() As Boolean
        Get
            Return _HBarEnabled
        End Get
        Set(ByVal value As Boolean)
            If _HBarEnabled <> value Then
                _HBarEnabled = value
                UpdateScrollFields(VBar.TotalSize, HBar.TotalSize)
            End If
        End Set
    End Property

    Private _VBarEnabled As Boolean
    Public Property VBarEnabled() As Boolean
        Get
            Return _VBarEnabled
        End Get
        Set(ByVal value As Boolean)
            _VBarEnabled = value
        End Set
    End Property

    Private _VisualStyle As ScrollBars.VisualStyles = ScrollBars.VisualStyles.None
    Public Property VisualStyle() As ScrollBars.VisualStyles
        Get
            Return _VisualStyle
        End Get
        Set(ByVal value As ScrollBars.VisualStyles)
            VBar.VisualStyle = value
            HBar.VisualStyle = value
        End Set
    End Property

    Public Sub New(ByVal tControl As Control, ByVal bVBarEnabled As Boolean, ByVal bHBarEnabled As Boolean)
        VBar = New ScrollBars.VBar(tControl, "VBar", tControl.ClientSize.Width - ScrollBars.BarSize, 0, ScrollBars.BarSize, tControl.ClientSize.Height)
        HBar = New ScrollBars.HBar(tControl, "HBar", 0, tControl.ClientSize.Height - ScrollBars.BarSize, tControl.ClientSize.Width, ScrollBars.BarSize)
        _VBarEnabled = bVBarEnabled
        _HBarEnabled = bHBarEnabled
        VBar.Visible = False
        HBar.Visible = False
        AddHandler tControl.Resize, AddressOf Parent_Resized
    End Sub

    Public Sub SetColors(ByVal tBGColor As Color, Optional ByVal tBarColor As Color = Nothing, Optional ByVal tBarOutlineColor As Color = Nothing)
        If tBGColor <> Nothing Then
            VBar.BGColor = New SolidBrush(tBGColor)
            HBar.BGColor = New SolidBrush(tBGColor)
        End If
        If tBarColor <> Nothing Then
            VBar.BarColor = New SolidBrush(tBarColor)
            HBar.BarColor = New SolidBrush(tBarColor)
        End If
        If tBarOutlineColor <> Nothing Then
            VBar.BarOutline = New Pen(tBarOutlineColor)
            HBar.BarOutline = New Pen(tBarOutlineColor)
        End If
    End Sub

    Private Sub Parent_Resized(ByVal sender As Object, ByVal e As EventArgs)
        VBar.Left = VBar.Parent.ClientSize.Width - VBar.Width
        HBar.Top = HBar.Parent.ClientSize.Height - HBar.Height
        UpdateScrollFields(VBar.TotalSize, HBar.TotalSize)
    End Sub

    Public Sub UpdateScrollFields(ByVal VerticalMax As Integer, ByVal HorizontalMax As Integer, Optional ByVal VItemSize As Integer = -1, Optional ByVal HItemSize As Integer = -1)
        Dim bHBar As Boolean = False, bVBar As Boolean = False, THSize As Integer = HSize, TVSize As Integer = VSize
        If VItemSize <> -1 Then VBar.ItemSize = VItemSize
        If HItemSize <> -1 Then HBar.ItemSize = HItemSize
        If THSize = -1 Then THSize = VBar.Parent.ClientSize.Width
        If TVSize = -1 Then TVSize = VBar.Parent.ClientSize.Height
        If _VBarEnabled Then
            If VerticalMax > TVSize Then
                bVBar = True
            ElseIf _HBarEnabled AndAlso HorizontalMax > THSize AndAlso VerticalMax > TVSize - ScrollBars.BarSize Then 'hbar shown
                bHBar = True
                bVBar = True
            End If
        End If
        If Not bHBar AndAlso _HBarEnabled Then
            If bVBar Then
                If HorizontalMax > THSize - ScrollBars.BarSize Then bHBar = True
            Else
                If HorizontalMax > THSize Then bHBar = True
            End If
        End If
        Dim NewLength As Integer = -1
        If bVBar Then
            If bHBar Then NewLength = TVSize - ScrollBars.BarSize Else NewLength = TVSize
            VBar.SetBar(VerticalMax, NewLength, VBar.ItemSize, NewLength)
            VBar.Visible = True
        Else
            VBar.Visible = False
            VBar.SetBar(VerticalMax, VBar.ClientSize.Height, VBar.ItemSize, -1)
        End If

        NewLength = -1
        If bHBar Then
            If bVBar Then NewLength = THSize - ScrollBars.BarSize Else NewLength = THSize
            HBar.SetBar(HorizontalMax, NewLength, HBar.ItemSize, NewLength)
            HBar.Visible = True
        Else
            HBar.Visible = False
            HBar.SetBar(HorizontalMax, HBar.ClientSize.Width, HBar.ItemSize, NewLength)
        End If

    End Sub

    Private Sub Bar_ValueChanged(sender As Object, e As System.EventArgs) Handles HBar.ValueChanged, VBar.ValueChanged
        RaiseEvent ValueChanged(sender, e)
    End Sub

End Class

Public MustInherit Class ScrollBars : Inherits System.Windows.Forms.Control
    Private Shadows disposed As Boolean = False
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                ' Free other state (managed objects).
            End If
            BackBuffer.Dispose()
            BGx.Dispose()
            MDownTimer = Nothing
            MDownOrigin = Nothing
            MDownBarOrigin = Nothing
            BGColor.Dispose()
            BarColor.Dispose()
            BarOutline.Dispose()
            BarRect = Nothing
            Arrows(0) = Nothing
            Arrows(1) = Nothing
            Arrows = Nothing
            Parent = Nothing
            ' Free your own state (unmanaged objects).
            ' Set large fields to null.
        End If
        MyBase.Dispose(disposing)
        Me.disposed = True
    End Sub

    Public Enum VisualStyles
        None = 0
        Styled = 1
    End Enum

    Public Event Scrolling(ByVal sender As Object, ByVal e As EventArgs)
    Public Event ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Shared BarSize As Integer = 16, MDownWait As Integer = 600
    Public BGColor As New SolidBrush(Color.FromArgb(255, 238, 238, 239)), BarColor As New SolidBrush(Color.FromArgb(255, 213, 213, 216)), BarOutline As New Pen(Color.FromArgb(255, 147, 147, 157))
    Public Arrows(1) As Arrow, ArrowImgIndex As Integer = 0, BarRect As Rectangle = Nothing
    Public TotalSize As Integer, ViewSize As Integer, ObscuredSize As Integer, ItemSize As Integer, Min As Integer, Max As Integer, ScrollPosition As Integer = 0
    Public FocusControl As Control
    Private bVertical As Boolean = False, BackBuffer As Bitmap, BGx As Graphics
    Private bMouseDown As Boolean = False, MDownOrigin As Point, MDownBarOrigin As Point
    Private ValueMax As Integer = 0, ArrowDownIndex As Integer = 0, MDownTimer As System.Threading.Thread
    Private bIgnoreResize As Boolean = False
    Public Focuzed As Boolean = False
    Private _VisualStyle As VisualStyles = VisualStyles.Styled
    Public Property VisualStyle() As VisualStyles
        Get
            Return _VisualStyle
        End Get
        Set(ByVal value As VisualStyles)
            If value <> _VisualStyle Then
                _VisualStyle = value
                If Me.Visible Then Me.Refresh()
            End If
        End Set
    End Property

    Private _Value As Integer
    Public Property Value As Integer
        Get
            Return _Value
        End Get
        Set(ByVal value As Integer)
            value = Math.Max(0, Math.Min(ValueMax, value))
            If _Value <> value Then
                _Value = value
                If ValueMax <> 0 Then ScrollPosition = CInt((_Value / ValueMax) * Max)
                If TypeOf Me Is VBar Then
                    BarRect.Y = BarSize + ScrollPosition
                Else
                    BarRect.X = BarSize + ScrollPosition
                End If
                RaiseEvent ValueChanged(Me, New EventArgs)
            End If
        End Set
    End Property

    Public Sub UpdateValue()
        Value = Math.Min(ValueMax, Value)
    End Sub

    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        SetStyle(ControlStyles.Selectable, False)
        UpdateStyles()
        Parent = prnt
        FocusControl = Parent
        Name = sName
        Me.SetBounds(x, y, w, h)
    End Sub

    Private Sub ArrowDownLong()
        System.Threading.Thread.Sleep(MDownWait)
        If Not bMouseDown Then Exit Sub
        Do
            Try
                Me.Invoke(New Action(AddressOf ArrowScroll))
                System.Threading.Thread.Sleep(200)
                If Not bMouseDown Then Exit Sub
            Catch ex As Exception
                Exit Sub
            End Try
        Loop
    End Sub

    Private Sub ArrowScroll()
        If ArrowDownIndex = 0 Then
            Value = Math.Max(0, _Value - 1)
        Else
            Value = Math.Min(_Value + 1, ValueMax)
        End If
    End Sub



    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        BackBuffer = New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
        BGx = Graphics.FromImage(BackBuffer)
        If BarRect <> Nothing Then
            If Not bIgnoreResize Then SetBarRect()
        End If
        Resized()
        MyBase.OnResize(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        Focuzed = True
        FocusControl.Focus()
        bMouseDown = True
        Dim pt As New Point(e.X, e.Y)
        ArrowDownIndex = -1
        If BarRect.Contains(pt) Then
            MDownOrigin = pt
            MDownBarOrigin = New Point(BarRect.Left, BarRect.Top)
        Else
            ArrowDownIndex = ArrowDown(pt)
            If ArrowDownIndex = 2 Then
                ArrowDownIndex = -1
            Else
                ArrowScroll()
                MDownTimer = New System.Threading.Thread(AddressOf ArrowDownLong)
                MDownTimer.IsBackground = True
                MDownTimer.Start()
            End If
        End If
        MyBase.OnMouseDown(e)
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        Focuzed = False
        If ArrowDownIndex <> -1 Then 'stop teh thread
            ArrowDownIndex = -1
            MDownTimer = Nothing
        End If
        bMouseDown = False
        MDownOrigin = Nothing
        MDownBarOrigin = Nothing
        MyBase.OnMouseUp(e)
    End Sub

    Protected Overrides Sub OnLostFocus(ByVal e As System.EventArgs)
        bMouseDown = False
        MDownOrigin = Nothing
        MDownBarOrigin = Nothing
        MyBase.OnLostFocus(e)
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        If bMouseDown = False Or MDownOrigin = Nothing Then
            MyBase.OnMouseMove(e)
            Exit Sub
        End If
        If UpdateBarRect(e.X, e.Y) Then
            Dim i_Value As Integer = _Value
            _Value = Math.Min(ValueMax, CInt(Math.Ceiling(((ScrollPosition / Max) * ObscuredSize) / ItemSize)))
            RaiseEvent Scrolling(Me, New EventArgs)
            If i_Value <> _Value Then
                RaiseEvent ValueChanged(Me, New EventArgs)
            Else
                Me.Refresh()
            End If
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        BGx.FillRectangle(BGColor, Me.ClientRectangle)
        Dim pts1() As Drawing.Point = Nothing, pts2() As Drawing.Point = Nothing
        Dim midp As Integer = CInt(Math.Ceiling(Arrow.ArrowSize / 2))
        If TypeOf Me Is HBar Then
            pts1 = {New Point(0, midp), New Point(BarSize, -1), New Point(BarSize, BarSize)}
            pts2 = {New Point(Me.Width - 1, midp), New Point(Arrows(1).x, BarSize), New Point(Arrows(1).x, -1)}

        Else
            pts1 = {New Point(midp, 0), New Point(-1, BarSize), New Point(BarSize, BarSize)}
            pts2 = {New Point(midp, Me.Height - 1), New Point(-1, Arrows(1).y), New Point(BarSize, Arrows(1).y)}
        End If
        BGx.FillPolygon(BarColor, pts1)
        BGx.FillPolygon(BarColor, pts2)
        If _VisualStyle = VisualStyles.Styled Then
            BGx.DrawPolygon(BarOutline, pts1)
            BGx.DrawPolygon(BarOutline, pts2)
        End If

        If BarRect <> Nothing Then DrawBar()
        e.Graphics.DrawImageUnscaled(BackBuffer, 0, 0)
        MyBase.OnPaint(e)
    End Sub

    Private Sub DrawBar()
        If _VisualStyle = VisualStyles.None Then
            BGx.FillRectangle(BarColor, BarRect.X + 1, BarRect.Top + 1, BarRect.Width - 2, BarRect.Height - 2)
        ElseIf _VisualStyle = VisualStyles.Styled Then
            BGx.DrawLine(BarOutline, BarRect.X, BarRect.Y + 1, BarRect.X, BarRect.Bottom - 2)
            BGx.DrawLine(BarOutline, BarRect.Right - 1, BarRect.Y + 1, BarRect.Right - 1, BarRect.Bottom - 2)
            BGx.DrawLine(BarOutline, BarRect.X + 1, BarRect.Y, BarRect.Right - 2, BarRect.Y)
            BGx.DrawLine(BarOutline, BarRect.X + 1, BarRect.Bottom - 1, BarRect.Right - 2, BarRect.Bottom - 1)
            BGx.FillRectangle(BarColor, BarRect.X + 1, BarRect.Top + 1, BarRect.Width - 2, BarRect.Height - 2)
            If TypeOf Me Is HBar Then
                If BarRect.Width > 20 Then
                    Dim mid As Integer = BarRect.Left + CInt(BarRect.Width / 2), dist As Integer = 3
                    For i As Integer = mid - dist To mid + dist Step dist
                        BGx.DrawLine(BarOutline, i, 4, i, Height - 5)
                    Next
                End If
            Else
                If BarRect.Height > 20 Then
                    Dim mid As Integer = BarRect.Top + CInt(BarRect.Height / 2), dist As Integer = 3
                    For i As Integer = mid - dist To mid + dist Step dist
                        BGx.DrawLine(BarOutline, 4, i, Width - 5, i)
                    Next
                End If
            End If
        End If
    End Sub

    Public Sub SetBar(ByVal tTotalSize As Integer, ByVal tViewSize As Integer, ByVal tItemSize As Integer, Optional ByVal tNewLength As Integer = -1)
        If tNewLength <> -1 Then
            bIgnoreResize = True
            If TypeOf Me Is VBar Then Height = tNewLength Else Width = tNewLength
            bIgnoreResize = False
        End If
        TotalSize = tTotalSize
        ViewSize = tViewSize
        ObscuredSize = TotalSize - ViewSize
        ItemSize = tItemSize
        Min = 0
        ScrollPosition = 0
        SetBarRect()
    End Sub

    Private Sub ScrollBars_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
        Value = 0
        ScrollPosition = 0
    End Sub

    Public Sub Clear()
        BarRect = Nothing
        Value = 0
        ScrollPosition = 0
        Me.Visible = False
    End Sub

    Protected Overridable Function ArrowDown(ByVal pt As Point) As Integer
        Return 0
    End Function

    Protected Overridable Sub Resized()

    End Sub

    Protected Overridable Sub SetBarRect()

    End Sub

    Protected Overridable Function UpdateBarRect(ByVal x As Integer, ByVal y As Integer) As Boolean
        Return False
    End Function

    Public Class VBar : Inherits ScrollBars
        Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
            MyBase.New(prnt, sName, x, y, w, h)
            ArrowImgIndex = 2 'images 2 and 3, top and down arrows
            Arrows(0) = New Arrow(2, 0, 0)
            Arrows(1) = New Arrow(3, 0, Me.ClientSize.Height - BarSize)
        End Sub

        Protected Overrides Function UpdateBarRect(ByVal x As Integer, ByVal y As Integer) As Boolean
            If y = MDownOrigin.Y Then Return False
            Dim oy As Integer = BarRect.Y
            BarRect.Y = Math.Min(Math.Max(BarSize, MDownBarOrigin.Y + (y - MDownOrigin.Y)), Max + BarSize)
            If oy <> BarRect.Y Then
                ScrollPosition = BarRect.Y - BarSize
                Return True
            Else
                Return False
            End If
        End Function

        Protected Overrides Function ArrowDown(ByVal pt As Point) As Integer
            If pt.Y <= BarSize Then Return 0 Else If pt.Y >= Me.ClientSize.Height - BarSize Then Return 1 Else Return 2
        End Function

        Protected Overrides Sub Resized()
            If Not Arrows(1) Is Nothing Then Arrows(1).y = Me.ClientSize.Height - BarSize
        End Sub

        Protected Overrides Sub SetBarRect()
            If TotalSize = 0 Then Exit Sub
            Dim tBarHeight As Integer = Math.Max(10, CInt((ViewSize / TotalSize) * (Me.ClientSize.Height - BarSize - BarSize)))
            Dim oMax As Integer = ValueMax
            Max = Me.ClientSize.Height - BarSize - BarSize - tBarHeight
            ValueMax = CInt(Math.Max(0, Math.Ceiling(ObscuredSize / ItemSize)))
            If oMax <> 0 AndAlso BarRect <> Nothing Then
                Dim sPct As Double = Value / oMax
                BarRect = New Rectangle(0, BarSize, BarSize, tBarHeight)
                Value = CInt(sPct * ValueMax)
            Else
                BarRect = New Rectangle(0, BarSize, BarSize, tBarHeight)
            End If
        End Sub
    End Class

    Public Class HBar : Inherits ScrollBars
        Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
            MyBase.New(prnt, sName, x, y, w, h)
            ArrowImgIndex = 0 'images 0 and 1, left and right arrows
            Arrows(0) = New Arrow(0, 0, 0)
            Arrows(1) = New Arrow(1, Me.ClientSize.Width - BarSize, 0)
        End Sub

        Protected Overrides Function UpdateBarRect(ByVal x As Integer, ByVal y As Integer) As Boolean
            If x = MDownOrigin.X Then Return False
            Dim ox As Integer = BarRect.X
            BarRect.X = Math.Min(Math.Max(BarSize, MDownBarOrigin.X + (x - MDownOrigin.X)), Max + BarSize)
            If ox <> BarRect.X Then
                ScrollPosition = BarRect.X - BarSize
                Return True
            Else
                Return False
            End If
        End Function

        Protected Overrides Sub SetBarRect()
            If TotalSize = 0 Then Exit Sub
            Dim tBarWidth As Integer = Math.Max(10, CInt((ViewSize / TotalSize) * (Me.ClientSize.Width - BarSize - BarSize)))
            Dim oMax As Integer = ValueMax
            Max = Me.ClientSize.Width - BarSize - BarSize - tBarWidth
            ValueMax = CInt(Math.Max(0, Math.Ceiling(ObscuredSize / ItemSize)))
            If oMax <> 0 AndAlso BarRect <> Nothing Then
                Dim sPct As Double = Value / oMax
                BarRect = New Rectangle(BarSize, 0, tBarWidth, BarSize)
                Value = CInt(sPct * ValueMax)
            Else
                BarRect = New Rectangle(BarSize, 0, tBarWidth, BarSize)
            End If
        End Sub

        Protected Overrides Function ArrowDown(ByVal pt As Point) As Integer
            If pt.X <= BarSize Then Return 0 Else If pt.X >= Me.ClientSize.Width - BarSize Then Return 1 Else Return 2
        End Function

        Protected Overrides Sub Resized()
            If Not Arrows(1) Is Nothing Then Arrows(1).x = Me.ClientSize.Width - BarSize
        End Sub
    End Class

    Public Class Arrow
        Public Shared ArrowImgs(3) As Bitmap, ArrowSize As Integer = 16
        Public ImgIndx As Integer, x As Integer, y As Integer
        Public Sub New(ByVal tImgIndex As Integer, ByVal tX As Integer, ByVal tY As Integer)
            ImgIndx = tImgIndex
            x = tX
            y = tY
        End Sub
    End Class

End Class

Public Class ListItems : Inherits System.Windows.Forms.Control
#Region "Events and Dispose"
    Public Event ItemsChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event ItemClicked(ByVal sender As Object, ByVal e As ItemClickedArgs)
    Public Class ItemClickedArgs : Inherits System.EventArgs
        Public MEvent As MouseEventArgs, PreviousIndex As Integer, NewIndex As Integer
        Public Sub New(ByVal tPreviousIndex As Integer, ByVal tNewIndex As Integer, ByVal e As MouseEventArgs)
            PreviousIndex = tPreviousIndex
            NewIndex = tNewIndex
            MEvent = e
        End Sub
    End Class
    Public Delegate Sub _SelectedIndexChanged(ByVal sender As Object, ByVal e As SelectedIndexChangedArgs)
    Private delSelectedIndexChanged As New List(Of _SelectedIndexChanged)

    Public Custom Event SelectedIndexChanged As _SelectedIndexChanged
        AddHandler(ByVal value As _SelectedIndexChanged)
            Me.Events.AddHandler("SelectedIndexChanged", value)
            delSelectedIndexChanged.Add(value)
        End AddHandler
        RemoveHandler(ByVal value As _SelectedIndexChanged)
            Me.Events.RemoveHandler("SelectedIndexChanged", value)
            delSelectedIndexChanged.Remove(value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As SelectedIndexChangedArgs)
            If Not delSelectedIndexChanged.Count = 0 Then CType(Me.Events("SelectedIndexChanged"), _SelectedIndexChanged).Invoke(sender, e)
        End RaiseEvent
    End Event

    Public Class SelectedIndexChangedArgs : Inherits System.EventArgs
        Public oldIndex As Integer, SetByClick As Boolean
        Public Sub New(ByVal tOldIndex As Integer, ByVal bSetByClick As Boolean)
            oldIndex = tOldIndex
            SetByClick = bSetByClick
        End Sub
    End Class

    Private Shadows disposed As Boolean = False
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                ' Free other state (managed objects).
                _Items.Dispose()
            End If
            ' Free your own state (unmanaged objects).
            ' Set large fields to null.
            If Not delSelectedIndexChanged Is Nothing Then
                For i As Integer = delSelectedIndexChanged.Count - 1 To 0 Step -1
                    RemoveHandler Me.SelectedIndexChanged, delSelectedIndexChanged(i)
                Next
                delSelectedIndexChanged = Nothing
            End If
            BGx.Dispose()
            BackBuffer.Dispose()
            HighlightBrush.Dispose()
            SelectionBrush.Dispose()
            FontBrush.Dispose()
            SelectionFontBrush.Dispose()
            BGBrush.Dispose()
            If Not _Scrolling.VBar Is Nothing Then _Scrolling.VBar.Dispose()
            If Not _Scrolling.HBar Is Nothing Then _Scrolling.HBar.Dispose()
        End If
        MyBase.Dispose(disposing)
        Me.disposed = True
    End Sub
#End Region

#Region "Variables and Properties"

    Public ItemsShown As Integer = 0, TotalHeight As Integer = 0
    Public HighlightedItem As Integer = -1, HighlightBrush As New SolidBrush(Color.FromArgb(255, Color.AliceBlue)), SelectionBrush As New SolidBrush(Color.FromArgb(255, 51, 153, 255)), SelectionFontBrush As New SolidBrush(Color.White)
    Public ShowErrorMessages As Boolean = True
    Public MultiSelect As Boolean = False, DeselectOnEmptySpaceClick As Boolean = False, DeselectLastItemOnCtrlClick As Boolean = True
    Public Tag2 As Object

    Private BackBuffer As Bitmap, BGx As Graphics, BGBrush As New SolidBrush(Color.AliceBlue), FontBrush As New SolidBrush(Color.Black), TextHeight As Integer
    Private BufferV As Integer = 2, BufferH As Integer = 2, BufferItemSpacing As Integer = 0, ItemHeight As Integer
    Private VScrollPosition As Integer, HScrollPosition As Integer
    Private ItemsDrawn As Integer = 0, LongestWidth As Integer
    Private KDThread As System.Threading.Thread, _Keys As New List(Of System.Windows.Forms.Keys)
    Private KeyDownWait As Integer = 600, KeyDownSpeed As Integer = 200, _NewValue As Integer
    Private LastItemClicked As Integer = 0, bSetByClick As Boolean = False
    Private bMinimized As Boolean = False
    Private ScrollLock As New Object, _ScrollThreadActive As Boolean = False

    Private _Scrolling As ScrollBarSet
    Public ReadOnly Property Scrolling() As ScrollBarSet
        Get
            Return _Scrolling
        End Get
    End Property

    Private _SelectedIndices As List(Of Integer)
    Public ReadOnly Property SelectedIndices() As List(Of Integer)
        Get
            Return _SelectedIndices
        End Get
    End Property

    Public ReadOnly Property SelectedItems() As List(Of Object)
        Get
            Dim lst As New List(Of Object)
            For i As Integer = 0 To _SelectedIndices.Count - 1
                lst.Add(_Items(_SelectedIndices(i)))
            Next
            Return lst
        End Get
    End Property

    Private _SelectedIndex As Integer = -1
    Public Property SelectedIndex() As Integer
        Get
            Return _SelectedIndex
        End Get
        Set(ByVal value As Integer)
            value = Math.Min(value, _Items.Count - 1)
            If value <> _SelectedIndex Then 'if a new item is selected
                Dim args As New SelectedIndexChangedArgs(_SelectedIndex, bSetByClick)
                _SelectedIndex = value
                LastItemClicked = value
                If value > -1 Then
                    _SelectedIndices.Remove(args.oldIndex)
                    _SelectedIndices.Add(value)
                    _SelectedIndices.Sort()
                    _SelectedItem = _Items(value)
                    If _Scrolling.VBar.Visible Then
                        Dim lolzo As Integer = CInt(Math.Floor(ItemsShown / 2))
                        _Scrolling.VBar.Value = Math.Max(0, value - lolzo)
                    End If
                Else
                    _SelectedItem = Nothing
                    _SelectedIndices = New List(Of Integer)
                    If _Scrolling.VBar.Visible Then
                        _Scrolling.VBar.UpdateValue()
                    End If
                End If
                Me.Refresh()
                RaiseEvent SelectedIndexChanged(Me, args)
            End If
        End Set
    End Property

    Private _StartIndex As Integer
    Public Property StartIndex() As Integer
        Get
            Return _StartIndex
        End Get
        Set(ByVal value As Integer)
            _StartIndex = value
        End Set
    End Property

    Private _EndIndex As Integer
    Public Property EndIndex() As Integer
        Get
            Return _EndIndex
        End Get
        Set(ByVal value As Integer)
            _EndIndex = value
        End Set
    End Property

    Private _SelectedItem As Object
    Public ReadOnly Property SelectedItem() As Object
        Get
            Return _SelectedItem
        End Get
    End Property

    Private _SelectedText As String
    Public ReadOnly Property SelectedText() As String
        Get
            Return _SelectedItem.ToString
        End Get
    End Property

    Private _Items As ItemList
    Public ReadOnly Property Items() As ItemList
        Get
            Return _Items
        End Get
    End Property

    Public ReadOnly Property IsLastItemSelected() As Boolean
        Get
            Return _SelectedIndices.Contains(_Items.Count - 1)
        End Get
    End Property

#End Region

#Region "Border"
    Private Const WM_NCCALCSIZE As Integer = 131
    Private Const WM_NCPAINT As Integer = 133
    Private Const WM_NCHITTEST As Integer = 132
    Private Const WM_NCLBUTTONDOWN As Integer = 161
    Private Const WM_NCMOUSEMOVE As Integer = &HA0

    <System.Runtime.InteropServices.DllImport("User32.dll")> _
    Friend Shared Function GetWindowDC(ByVal hWnd As IntPtr) As IntPtr
    End Function

    <System.Runtime.InteropServices.DllImport("User32.dll")> _
    Friend Shared Function ReleaseDC(ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Integer
    End Function

    Private Sub AdjustClientRect(ByRef rcClient As RECT)
        rcClient.Left += BorderThickness
        rcClient.Top += BorderThickness
        rcClient.Right -= BorderThickness
        rcClient.Bottom -= BorderThickness
    End Sub

    Private Structure RECT
        Public Left As Integer, Top As Integer, Right As Integer, Bottom As Integer
        Public Sub New(ByVal tLeft As Integer, ByVal tTop As Integer, ByVal tRight As Integer, ByVal tBottom As Integer)
            Left = tLeft : Top = tTop : Right = tRight : Bottom = tBottom
        End Sub
    End Structure

    Private Structure NCCALCSIZE_PARAMS
        Public rcNewWindow As RECT
        Public rcOldWindow As RECT
        Public rcClient As RECT
        Private lppos As IntPtr

        Public Sub New(trcNewWindow As RECT, trcOldWindow As RECT, trcClient As RECT, tlppos As IntPtr)
            rcNewWindow = trcNewWindow : rcOldWindow = trcOldWindow : rcClient = trcClient : lppos = tlppos
        End Sub
    End Structure

    Private BorderThickness As Integer = 1, BorderPen As New Pen(Color.Black)
    Private _BorderStyle As BorderStyle = Windows.Forms.BorderStyle.FixedSingle
    Public Property BorderStyle() As BorderStyle
        Get
            Return _BorderStyle
        End Get
        Set(ByVal value As BorderStyle)
            If _BorderStyle <> value Then
                If _BorderStyle = Windows.Forms.BorderStyle.None Then 'was 0 border
                    SetBorder(1)
                ElseIf _BorderStyle = Windows.Forms.BorderStyle.None Then 'going to 0 border
                    SetBorder(0)
                End If
                _BorderStyle = value
                Me.Refresh()
                PaintBorder()
            End If
        End Set
    End Property

#End Region

    Public Sub New(ByVal prnt As Control, ByVal tName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Me.SetStyle(ControlStyles.Selectable Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)
        Me.UpdateStyles()
        w = Math.Max(64, w) : h = Math.Max(64, h)
        BackColor = Color.White
        Parent = prnt
        Name = tName
        _Items = New ItemList(Me) : _SelectedIndices = New List(Of Integer) : _SelectedIndex = -1
        SetFont(New Font(Font.FontFamily, 8.25, FontStyle.Regular), Color.Black)
        SetBounds(x, y, w, h)
        SetBorder(1, Color.Black, Color.White)
        Me.MinimumSize = New Size(65, 65)
        Me.Refresh()
        AddHandler _Scrolling.ValueChanged, AddressOf VBar_ValueChanged
        AddHandler _Scrolling.VBar.VisibleChanged, AddressOf VBar_VisibleChanged
    End Sub

    Private Sub VBar_ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
        If TypeOf sender Is ScrollBars.VBar Then
            StartIndex = _Scrolling.VBar.Value
            EndIndex = Math.Min(_Items.Count - 1, StartIndex + ItemsShown - 1)
        End If
        Me.Refresh()
    End Sub

    Private Sub VBar_VisibleChanged(sender As Object, e As System.EventArgs)
        If _SelectedIndex > StartIndex Then
            Dim lolzo As Integer = CInt(Math.Floor(ItemsShown / 2))
            _Scrolling.VBar.Value = Math.Max(0, _SelectedIndex - lolzo)
        End If
    End Sub

    Private Sub SetItemsShown()
        Dim iLol As Integer = Me.ClientSize.Height - BufferV
        ItemsShown = CInt(Math.Floor(iLol / ItemHeight))
        If ItemsShown * ItemHeight + TextHeight <= iLol Then ItemsShown += 1
        TotalHeight = ItemHeight * _Items.Count - BufferItemSpacing
        EndIndex = Math.Min(_Items.Count - 1, StartIndex + ItemsShown - 1)
    End Sub

    Private Sub SetBorder(ByVal tThickness As Integer, Optional ByVal tBorderColor As Color = Nothing, Optional ByVal tBackgroundColor As Color = Nothing)
        BorderThickness = Math.Min(Math.Max(0, tThickness), 10)
        If Not tBorderColor = Nothing Then BorderPen = New Pen(tBorderColor)
        If Not tBackgroundColor = Nothing Then BGBrush = New SolidBrush(tBackgroundColor)
    End Sub

    Public Sub SetFont(ByVal tFont As Font, Optional ByVal tColor As Color = Nothing)
        Font = tFont
        TextHeight = MeasureStringHeight("yj0") - 2
        ItemHeight = TextHeight + BufferItemSpacing
        SetItemsShown()
        If Not tColor = Nothing Then FontBrush = New SolidBrush(tColor)
    End Sub

    Public Sub UpdateItems(ByVal bRefresh As Boolean, Optional ByVal tLongestWidth As Integer = -1, Optional ByVal Cleared As Boolean = False)
        TotalHeight = ItemHeight * _Items.Count - BufferItemSpacing
        EndIndex = Math.Min(_Items.Count - 1, StartIndex + ItemsShown - 1)
        UpdateScrolling(tLongestWidth)
        If Cleared Then ClearSelections()
        If bRefresh Then Me.Refresh()
        RaiseEvent ItemsChanged(Me, New EventArgs)
    End Sub

    Private Sub ClearSelections()
        _SelectedIndices.Clear() : _SelectedIndex = -1
        LastItemClicked = 0
    End Sub

    Private Sub UpdateScrolling(Optional ByVal tLongestWidth As Integer = -1)
        Dim bVBar As Boolean = CheckForVScrolling()
        Dim bHBar As Boolean = CheckForHScrolling(tLongestWidth)
        _Scrolling.UpdateScrollFields(TotalHeight, LongestWidth, ItemHeight, 1)
    End Sub

    Public Sub Clear(Optional ByVal bRefresh As Boolean = True)
        _Items.Clear()
        ItemsDrawn = 0
        HighlightedItem = -1
        _SelectedItem = Nothing
        _SelectedIndex = -1
        _SelectedIndices = New List(Of Integer)
        Me.Refresh()
    End Sub

    Private Function CheckForHScrolling(Optional ByVal tLongestWidth As Integer = -1) As Boolean
        If Not _Scrolling.HBarEnabled Then Return False
        Dim MaxWidth As Integer = _Scrolling.VBar.Left
        If LongestWidth > 0 OrElse tLongestWidth <> -1 Then
            LongestWidth = Math.Max(tLongestWidth, LongestWidth)
        Else
            LongestWidth = 0
            For i As Integer = 0 To Items.Count - 1
                Try
                    LongestWidth = Math.Max(LongestWidth, MeasureStringWidth(Items(i).ToString()))
                Catch ex As Exception
                    LongestWidth = Math.Max(LongestWidth, MeasureStringWidth("ERROR"))
                End Try
            Next
            LongestWidth += 5
        End If
        If LongestWidth > MaxWidth Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CheckForVScrolling() As Boolean
        If TotalHeight > Me.ClientSize.Height - BorderThickness - BufferV Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub SetBackBuffer(ByVal w As Integer, ByVal h As Integer)
        If Not BackBuffer Is Nothing Then BackBuffer.Dispose()
        BackBuffer = New Bitmap(w, h)
        BGx = Graphics.FromImage(BackBuffer)
        BGx.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
    End Sub

    Private Sub UpdateSI()
        SelectedIndex = _NewValue
    End Sub

    Private Sub ScrollThread(ByVal state As Object)
        SyncLock ScrollLock
            System.Threading.Thread.Sleep(KeyDownWait)
            Do
                If _Keys.Count = 0 Or Not _ScrollThreadActive Then Exit Do
                Try
                    If _Keys(0) = Keys.Down Then
                        If _SelectedIndex < _Items.Count - 1 Then
                            _NewValue = _SelectedIndex + 1
                            Me.Invoke(New Action(AddressOf UpdateSI))
                        Else
                            Exit Do
                        End If
                    Else
                        If _SelectedIndex > 0 Then
                            _NewValue = _SelectedIndex - 1
                            Me.Invoke(New Action(AddressOf UpdateSI))
                        Else
                            Exit Do
                        End If
                    End If
                    System.Threading.Thread.Sleep(KeyDownSpeed)
                Catch ex As Exception
                    Exit Do
                End Try
            Loop
            _ScrollThreadActive = False
        End SyncLock
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        If keyData = Keys.Down Or keyData = Keys.Up Then
            Return True
        Else
            Return MyBase.ProcessCmdKey(msg, keyData)
        End If
    End Function

    Protected Overrides Sub OnPreviewKeyDown(ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs)
        If e.KeyCode = Keys.Down Then
            If Not _Keys.Contains(e.KeyCode) Then
                _Keys.Add(e.KeyCode)
                If _SelectedIndex < _Items.Count - 1 Then
                    If _ScrollThreadActive Then Exit Sub
                    SelectedIndex += 1
                    _ScrollThreadActive = True
                    ThreadPool.QueueUserWorkItem(AddressOf ScrollThread)
                End If
            End If
        ElseIf e.KeyCode = Keys.Up Then
            If Not _Keys.Contains(e.KeyCode) Then
                _Keys.Add(e.KeyCode)
                If _SelectedIndex > 0 Then
                    If _ScrollThreadActive Then Exit Sub
                    SelectedIndex -= 1

                    _ScrollThreadActive = True
                    ThreadPool.QueueUserWorkItem(AddressOf ScrollThread)
                End If
            End If
        Else
            MyBase.OnPreviewKeyDown(e)
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim lolz As Integer = _Keys.Count
        Dim indx As Integer = _Keys.IndexOf(e.KeyCode)
        If indx <> -1 Then _Keys.RemoveAt(indx)
        _ScrollThreadActive = False
        MyBase.OnKeyUp(e)
    End Sub

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        PaintBorder()
        If ClientSize.Width = 0 AndAlso ClientSize.Height = 0 Then
            bMinimized = True
            Exit Sub
        ElseIf bMinimized Then
            bMinimized = False
            Exit Sub
        End If
        SetItemsShown()
        SetBorder(BorderThickness, Nothing, Nothing)
        SetBackBuffer(ClientSize.Width, ClientSize.Height)
        StartIndex = 0
        If _Scrolling Is Nothing Then _Scrolling = New ScrollBarSet(Me, True, False)
        _Scrolling.VSize = Me.ClientSize.Height - BorderThickness - BufferV
        Me.Refresh()
        MyBase.OnResize(e)
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
        If Not HighlightedItem = -1 Then
            HighlightedItem = -1
            Me.Refresh()
        End If
        MyBase.OnMouseLeave(e)
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        If (_Items.Count = 0 Or ItemsDrawn = 0) Then

        Else
            Dim indx As Integer = CInt(Math.Max(0, Math.Floor((e.Y - BufferV) / ItemHeight)))
            Dim CurrentIndex As Integer = indx + StartIndex
            If CurrentIndex = HighlightedItem Then Exit Sub
            If CurrentIndex = _SelectedIndex Then
                If HighlightedItem <> -1 Then HighlightedItem = -1 : Me.Refresh()
                Exit Sub
            End If
            If indx >= ItemsShown Then
                If Not (HighlightedItem = -1 Or HighlightedItem = (_SelectedIndex - StartIndex)) Then
                    HighlightedItem = -1
                    Me.Refresh()
                End If
            ElseIf CurrentIndex >= _Items.Count Then
                If Not (HighlightedItem = -1 Or HighlightedItem = (_SelectedIndex - StartIndex)) Then
                    HighlightedItem = -1
                    Me.Refresh()
                End If
            Else
                If HighlightedItem <> CurrentIndex Then 'new item is highlighted
                    HighlightedItem = CurrentIndex
                    Me.Refresh()
                End If
            End If
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        If _Items.Count = 0 Then MyBase.OnMouseDown(e) : Exit Sub
        Dim indx As Integer = StartIndex + CInt(Math.Max(0, Math.Floor((e.Y - BufferV) / ItemHeight)))
        If indx >= _Items.Count Then 'Empty white space clicked
            If DeselectOnEmptySpaceClick AndAlso Not ModifierKeys.HasFlag(Keys.Control) AndAlso Not ModifierKeys.HasFlag(Keys.Shift) Then
                If _SelectedIndices.Count > 0 Then _SelectedIndices.Clear() : SelectedIndex = -1
            End If
        Else
            Me.Focus()
            Dim ICArgs As New ItemClickedArgs(_SelectedIndex, indx, e)
            If Not _SelectedIndices.Contains(indx) Then
                If MultiSelect Then
                    If ModifierKeys.HasFlag(Keys.Control) Then
                        If ModifierKeys.HasFlag(Keys.Shift) Then SelectRange(indx, LastItemClicked, True) Else SelectItem(indx, True, True)
                    ElseIf _SelectedIndices.Count > 0 AndAlso ModifierKeys.HasFlag(Keys.Shift) Then
                        _SelectedIndices.Clear()
                        SelectRange(indx, LastItemClicked, False)
                    Else
                        SelectItem(indx, True, False)
                    End If
                Else
                    SelectItem(indx, True, False)
                End If
            Else
                If e.Button = Windows.Forms.MouseButtons.Right And MultiSelect Then

                ElseIf ModifierKeys.HasFlag(Keys.Control) Then
                    If _SelectedIndices.Count > 1 Or DeselectLastItemOnCtrlClick Then DeselectItem(indx)
                ElseIf _SelectedIndices.Count > 1 AndAlso Not ModifierKeys.HasFlag(Keys.Shift) Then
                    DeselectAll()
                    SelectItem(indx, True, True)
                End If
            End If
            RaiseEvent ItemClicked(Me, ICArgs)
            If Not ModifierKeys.HasFlag(Keys.Shift) Then LastItemClicked = indx
        End If
        MyBase.OnMouseDown(e)
    End Sub

    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
        If _Scrolling.VBar.BarRect <> Nothing Then 'scrolling available
            Dim oV As Integer = _Scrolling.VBar.Value
            If e.Delta > 0 Then '
                _Scrolling.VBar.Value -= 1
            Else '
                _Scrolling.VBar.Value += 1
            End If
        End If
        MyBase.OnMouseWheel(e)
    End Sub

    Protected Overrides Sub OnLostFocus(ByVal e As System.EventArgs)
        _Keys = New List(Of System.Windows.Forms.Keys)
        MyBase.OnLostFocus(e)
    End Sub

    Protected Overrides Sub OnVisibleChanged(ByVal e As System.EventArgs)
        MyBase.OnVisibleChanged(e)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)

    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        BGx.FillRectangle(New SolidBrush(BackColor), Me.ClientRectangle)
        Dim sLine As String = "", RealRight As Integer = 0
        If _Scrolling.VBar.Visible Then RealRight = _Scrolling.VBar.Left Else RealRight = Me.ClientSize.Width
        If _Items.Count > 0 Then
            Dim ItemTop As Integer = BufferV, ItemLeft As Integer = BufferH - _Scrolling.HBar.Value
            For i As Integer = StartIndex To EndIndex
                Try
                    sLine = Items(i).ToString()
                Catch ex As Exception
                    sLine = "ERROR"
                End Try
                If _SelectedIndices.Contains(i) Then 'selected item
                    BGx.FillRectangle(SelectionBrush, 0, ItemTop, RealRight, TextHeight)
                    BGx.DrawString(sLine, Font, SelectionFontBrush, ItemLeft, ItemTop)
                ElseIf i = HighlightedItem Then
                    BGx.FillRectangle(HighlightBrush, 0, ItemTop, RealRight, TextHeight)
                    BGx.DrawString(sLine, Font, FontBrush, ItemLeft, ItemTop)
                Else
                    BGx.DrawString(sLine, Font, FontBrush, ItemLeft, ItemTop)
                End If
                ItemTop += ItemHeight
            Next
            ItemsDrawn += 1
        End If
        If _Scrolling.HBar.Visible AndAlso _Scrolling.VBar.Visible Then
            BGx.FillRectangle(_Scrolling.HBar.BGColor, _Scrolling.HBar.Right, _Scrolling.VBar.Bottom, ScrollBars.BarSize, ScrollBars.BarSize)
        End If
        e.Graphics.DrawImageUnscaled(BackBuffer, 0, 0)
        MyBase.OnPaint(e)
    End Sub

    <DebuggerStepThrough>
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_NCCALCSIZE
                If m.WParam <> IntPtr.Zero Then
                    Dim rcsize As NCCALCSIZE_PARAMS = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(m.LParam, GetType(NCCALCSIZE_PARAMS)), NCCALCSIZE_PARAMS)
                    AdjustClientRect(rcsize.rcNewWindow)
                    System.Runtime.InteropServices.Marshal.StructureToPtr(rcsize, m.LParam, False)
                Else
                    Dim rcsize As RECT = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(m.LParam, GetType(RECT)), RECT)
                    AdjustClientRect(rcsize)
                    System.Runtime.InteropServices.Marshal.StructureToPtr(rcsize, m.LParam, False)
                End If
                m.Result = New IntPtr(1)
                Return
            Case WM_NCPAINT
                Dim hdc As IntPtr = GetWindowDC(m.HWnd)
                If hdc <> IntPtr.Zero Then
                    PaintBorder()
                End If
                Return
            Case Else
                MyBase.WndProc(m)
        End Select
    End Sub

    Private Sub PaintBorder()
        If _BorderStyle = Windows.Forms.BorderStyle.None Then Exit Sub
        Dim hDC As IntPtr = GetWindowDC(Me.Handle) '  m.HWnd)
        Dim g As Graphics = Graphics.FromHdc(hDC)
        If _BorderStyle = Windows.Forms.BorderStyle.FixedSingle Then
            For i As Integer = 0 To BorderThickness - 1
                g.DrawRectangle(BorderPen, i, i, Me.Width - (i * 2) - 1, Me.Height - (i * 2) - 1)
            Next
        ElseIf _BorderStyle = Windows.Forms.BorderStyle.Fixed3D Then
            Dim BGPen1 As New Pen(Color.FromArgb(255, 160, 160, 160))
            For i As Integer = 0 To BorderThickness - 1
                Dim w As Integer = Me.Width - i - 1, h As Integer = Me.Height - i - 1
                g.DrawLine(BGPen1, i, i, w - 1, i)
                g.DrawLine(BGPen1, i, i, i, h - 1)
                g.DrawLine(Pens.White, w, i, w, h)
                g.DrawLine(Pens.White, i, h, w, h)
            Next
        End If
        ReleaseDC(Me.Handle, hDC)
        g.Dispose()
    End Sub

    Private Sub SelectRange(ByVal num1 As Integer, ByVal num2 As Integer, Continuous As Boolean)
        If num1 > num2 Then
            If Continuous Then num2 += 1
            For i As Integer = num2 To num1
                If _SelectedIndices.Contains(i) Then _SelectedIndices.Remove(i) Else _SelectedIndices.Insert(_SelectedIndices.Count, i)
            Next
        Else
            If Continuous Then num2 -= 1
            For i As Integer = num1 To num2
                If _SelectedIndices.Contains(i) Then _SelectedIndices.Remove(i) Else _SelectedIndices.Insert(_SelectedIndices.Count, i)
            Next
        End If
        _SelectedIndex = _SelectedIndices(0)
        _SelectedIndices.Sort()
        Me.Refresh()
    End Sub

    Private Sub DeselectAll()
        _SelectedIndices = New List(Of Integer)
        _SelectedItem = -1
        _SelectedItem = Nothing
    End Sub

    Private Sub DeselectItem(ByVal tIndex As Integer)
        bSetByClick = True
        _SelectedIndices.Remove(tIndex)
        If _SelectedIndices.Count = 0 Then
            SelectedIndex = -1
        ElseIf _SelectedIndex <> _SelectedIndices(0) Then
            SelectedIndex = _SelectedIndices(0)
        Else
            Me.Refresh()
        End If
        bSetByClick = False
    End Sub

    Private Sub SelectItem(ByVal tIndex As Integer, ByVal SetByClick As Boolean, ByVal bAdd As Boolean)
        Dim args As New SelectedIndexChangedArgs(_SelectedIndex, SetByClick)
        _SelectedIndex = tIndex
        If bAdd Then
            _SelectedIndices.Add(tIndex)
            _SelectedIndices.Sort()
        Else
            _SelectedIndices = New List(Of Integer) From {tIndex}
        End If
        _SelectedItem = _Items(_SelectedIndices(0))
        Me.Refresh()
        RaiseEvent SelectedIndexChanged(Me, args)
    End Sub

    Private Function MeasureStringHeight(ByVal str As String) As Integer
        Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
            Dim SFormat As New System.Drawing.StringFormat
            Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
            Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
            SFormat.SetMeasurableCharacterRanges(range)
            Dim regions() As Region = gx.MeasureCharacterRanges(str, Font, rect, SFormat)
            rect = regions(0).GetBounds(gx)
            Return CInt(rect.Bottom + 1)
        End Using
    End Function

    Private Function MeasureStringWidth(ByVal str As String) As Integer
        Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
            Dim SFormat As New System.Drawing.StringFormat
            Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
            Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
            SFormat.SetMeasurableCharacterRanges(range)
            Dim regions() As Region = gx.MeasureCharacterRanges(str, Font, rect, SFormat)
            rect = regions(0).GetBounds(gx)
            Return CInt(rect.Right + 1)
        End Using
    End Function

    Public Class ItemList
        Implements IDisposable, IList(Of Object)

#Region "IDisposable Support"
        Private disposedValue As Boolean
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If
                LI = Nothing
                _Items = Nothing
                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

        Protected _Items As New List(Of Object), LI As ListItems
        Default Property Item(ByVal indx As Integer) As Object Implements IList(Of Object).Item
            Get
                If indx < _Items.Count Then Return _Items(indx) Else Return Nothing
            End Get
            Set(ByVal value As Object)
                If value IsNot _Items(indx) Then
                    _Items(indx) = value
                    If LI.SelectedIndex = indx Then LI._SelectedItem = value
                    LI.UpdateItems(True)
                End If
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements ICollection(Of Object).Count
            Get
                Return _Items.Count
            End Get
        End Property

        Public Sub Sort()
            _Items.Sort()
            LI.UpdateItems(True)
        End Sub

        Public Sub Clear() Implements ICollection(Of Object).Clear
            _Items.Clear()
            LI.SelectedIndex = -1
            LI.UpdateItems(True)
        End Sub

        Public Sub New(ByVal tListItems As ListItems)
            LI = tListItems
        End Sub

        Public Sub SetItems(ByVal tItems As Generic.IEnumerable(Of Object))
            If tItems.Count <> _Items.Count Then
                MsgBox("Error: Item count must be equal to replace current list of items.")
            Else
                _Items = Nothing
                _Items = tItems.ToList()
                If LI.SelectedIndex <> -1 Then LI._SelectedItem = _Items(LI.SelectedIndex)
                LI.Refresh()
            End If
        End Sub

        Public Function Contains(ByVal tObj As Object) As Boolean Implements ICollection(Of Object).Contains
            Return _Items.Contains(tObj)
        End Function

        Public Function ContainsText(ByVal sText As String, Optional ByVal bIgnoreCase As Boolean = True) As Boolean
            If bIgnoreCase Then
                Dim UText As String = sText.ToUpper
                For i As Integer = 0 To _Items.Count - 1
                    If _Items(i).ToString.ToUpper = UText Then Return True
                Next
                Return False
            Else
                For i As Integer = 0 To _Items.Count - 1
                    If _Items(i).ToString = sText Then Return True
                Next
                Return False
            End If
        End Function

        Public Function IndexOf(ByVal tObject As Object) As Integer Implements IList(Of Object).IndexOf
            Return _Items.IndexOf(tObject)
        End Function

        Public Sub AppendAllText(ByVal lst As Generic.IEnumerable(Of String))
            Try
                For i As Integer = 0 To lst.Count - 1
                    _Items(i) = CStr(_Items(i)) & lst(i)
                Next
                LI.Refresh()
            Catch ex As Exception
                Throw
            End Try
        End Sub

        Public Sub RemoveAllTextFromIndexOfChar(ByVal tChar As Char)
            Try
                Dim indx As Integer = 0
                For i As Integer = 0 To _Items.Count - 1
                    indx = _Items(i).ToString.IndexOf(tChar)
                    If indx <> -1 Then _Items(i) = _Items(i).ToString.Substring(0, indx).Trim()
                Next
                LI.Refresh()
            Catch ex As Exception
                Throw
            End Try
        End Sub

        Public Overloads Sub Add(ByVal item As Object, ByVal bRefresh As Boolean)
            _Items.Add(item)
            LI.UpdateItems(bRefresh)
        End Sub

        Public Overloads Sub Add(item As Object) Implements ICollection(Of Object).Add
            _Items.Add(item)
            LI.UpdateItems(True)
        End Sub

        Public Overloads Sub AddRange(ByVal lstItems As Generic.IEnumerable(Of Object), Optional ByVal Clear As Boolean = False)
            If Clear Then _Items.Clear()
            If lstItems Is Nothing Then Exit Sub
            _Items.AddRange(lstItems)
            Dim LongestWidth As Integer = -1
            If LI.Scrolling.HBarEnabled = True Then
                For Each obj As Object In lstItems
                    LongestWidth = Math.Max(LongestWidth, MeasureStringWidth(obj.ToString()))
                Next
                LongestWidth += 5
            End If
            LI.UpdateItems(True, LongestWidth, Clear)
        End Sub

        Private Function MeasureStringWidth(ByVal str As String) As Integer
            Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
                Dim SFormat As New System.Drawing.StringFormat
                Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
                Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
                SFormat.SetMeasurableCharacterRanges(range)
                Dim regions() As Region = gx.MeasureCharacterRanges(str, LI.Font, rect, SFormat)
                rect = regions(0).GetBounds(gx)
                Return CInt(rect.Right + 1)
            End Using
        End Function

        Public Overloads Sub Insert(index As Integer, item As Object) Implements IList(Of Object).Insert
            Insert(index, item, True)
        End Sub

        Public Overloads Sub Insert(ByVal Index As Integer, ByVal obj As Object, Optional ByVal bRefresh As Boolean = True)
            _Items.Insert(Index, obj)
            If LI.SelectedIndex >= Index Then LI.SelectedIndex = LI.SelectedIndex + 1
            LI.UpdateItems(bRefresh)
        End Sub

        Public Overloads Sub Insert(ByVal bChangeIndex As Boolean, ByVal Index As Integer, ByVal obj As Object)
            _Items.Insert(Index, obj)
            LI.UpdateItems(True)
        End Sub

        Public Sub InsertRange(ByVal Index As Integer, ByVal lstItems As List(Of Object), Optional ByVal bRefresh As Boolean = True)
            _Items.InsertRange(Index, lstItems)
            If LI.SelectedIndex >= Index Then LI.SelectedIndex = LI.SelectedIndex + lstItems.Count
            LI.UpdateItems(bRefresh)
        End Sub

        Public Function Find(ByVal match As Predicate(Of Object)) As Object
            Return _Items.Find(match)
        End Function

        Public Function FindIndex(ByVal match As Predicate(Of Object)) As Integer
            Return _Items.FindIndex(match)
        End Function

        Public Overloads Function Remove(ByVal obj As Object) As Boolean Implements ICollection(Of Object).Remove
            Dim indx As Integer = _Items.IndexOf(obj)
            If indx <> -1 Then Return Remove(indx) Else Return False
        End Function

        Public Overloads Function Remove(ByVal tIndex As Integer) As Boolean
            Try
                _Items.RemoveAt(tIndex)
                LI.UpdateItems(False)
                If Count <> 0 Then
                    If LI.SelectedIndex = tIndex Then
                        LI.SelectedIndex = -1
                    ElseIf LI.SelectedIndex >= tIndex Then
                        LI.SelectedIndex = LI.SelectedIndex - 1
                    End If
                Else
                    LI.SelectedIndex = -1
                End If
                LI.Refresh()
                Return True
            Catch ex As Exception
                If LI.ShowErrorMessages Then MsgBox("Error removing item at index " & tIndex & ": " & ex.Message)
                Return False
            End Try
        End Function

        Public Sub RemoveAt(ByVal Index As Integer) Implements IList(Of Object).RemoveAt
            _Items.RemoveAt(Index)
            LI.UpdateItems(True)
        End Sub

        Public Sub RemoveSelected()
            Remove(LI.SelectedIndex)
        End Sub

        Public Sub RemoveList(ByVal lst As List(Of Integer))
            Dim RemoveSelected As Boolean = False, NewIndex As Integer = LI.SelectedIndex
            If LI.SelectedIndices.Count > 1 Then
                For i As Integer = LI.SelectedIndices.Count - 1 To 0 Step -1
                    If lst.Contains(LI.SelectedIndices(i)) Then LI.SelectedIndices.Remove(i)
                Next
            Else
                If lst.Contains(LI.SelectedIndex) Then
                    RemoveSelected = True
                End If
            End If
            lst.Sort()
            If NewIndex = -1 Then
                For i As Integer = lst.Count - 1 To 0 Step -1
                    _Items.RemoveAt(lst(i))
                Next
            Else
                For i As Integer = lst.Count - 1 To 0 Step -1
                    Dim intz As Integer = lst(i)
                    If intz < NewIndex Then NewIndex -= 1
                    _Items.RemoveAt(intz)
                Next
            End If
            LI.UpdateItems(False)
            If RemoveSelected Then
                LI.SelectedIndex = -1
            Else
                If NewIndex <> LI.SelectedIndex Then
                    LI.SelectedIndex = NewIndex
                Else
                    LI.Refresh()
                End If
            End If
        End Sub

        Public Sub RemoveRange(ByVal tIndex As Integer, ByVal tCount As Integer)
            Try
                _Items.RemoveRange(tIndex, tCount)
                LI.UpdateItems(False)
                If LI.SelectedIndex >= tIndex AndAlso LI.SelectedIndex < tIndex + tCount Then 'selected item is in the removal range
                    LI.SelectedIndex = -1
                ElseIf LI.SelectedIndex >= tIndex Then
                    LI.SelectedIndex = LI.SelectedIndex - tCount
                Else
                    LI.Refresh()
                End If
            Catch ex As Exception
                If LI.ShowErrorMessages Then MsgBox("Error removing items at index " & tIndex & " with count " & tCount & ": " & ex.Message)
            End Try
        End Sub

        Public Sub CopyTo(array() As Object, arrayIndex As Integer) Implements ICollection(Of Object).CopyTo
            _Items.CopyTo(array, arrayIndex)
        End Sub

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of Object).IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Function GetEnumerator() As IEnumerator(Of Object) Implements IEnumerable(Of Object).GetEnumerator
            Return _Items.GetEnumerator()
        End Function

        Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

    End Class

End Class

<DebuggerStepThrough>
Public Class SearchListItems
    Public Event SelectedIndexChanged(ByVal sender As Object, ByVal e As ListItems.SelectedIndexChangedArgs)
    Public WithEvents lblTitle As LabelX, txtSearch As TextBoxX, lstItems As ListItems, Items As ListItems.ItemList
    Public lstObj As New List(Of Object), CaseSensitive As Boolean = False
    Protected bWorking As Boolean = False, bSetting As Boolean = False, _Bounds As Rectangle
    Protected _Index As Integer = -1, objList As List(Of ObjNum)

    Public Property Index() As Integer
        Get
            Return _Index
        End Get
        Set(value As Integer)
            If value <> _Index Then
                If txtSearch.Text <> "" Then txtSearch.Text = ""
                lstItems.SelectedIndex = value
            End If
        End Set
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return lstObj.Count
        End Get
    End Property

    Public ReadOnly Property SItem() As Object
        Get
            Return lstItems.SelectedItem
        End Get
    End Property

    Public ReadOnly Property Bounds() As Rectangle
        Get
            Return _Bounds
        End Get
    End Property

    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sTitle As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        lblTitle = New LabelX(prnt, "lblTitle", sTitle, x, y, -1, -1)
        txtSearch = New TextBoxX(prnt, "txtSearch", "", lblTitle.Right, y, w - lblTitle.Width, 22) : txtSearch.ShortcutsEnabled = False
        lstItems = New ListItems(prnt, sName, x, txtSearch.Bottom + 2, w, h - 24)
        lstItems.Tag2 = Me
        Items = lstItems.Items
        lstObj = New List(Of Object)
        objList = New List(Of ObjNum)
        _Bounds = New Rectangle(x, y, w, h)
    End Sub

    Public Sub ReplaceItem(ByVal tIndex As Integer, ByVal tObj As Object)
        lstObj(tIndex) = tObj
        lstItems.Items(tIndex) = tObj
    End Sub

    Public Function ContainsText(ByVal sText As String, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        Return Items.ContainsText(sText, IgnoreCase)
    End Function

    Public Sub Clear()
        Try
            lstObj.Clear()
            If txtSearch.Text <> "" Then txtSearch.Text = "" Else lstItems.Clear()
            txtSearch.ShortcutsEnabled = False
        Catch ex As Exception

        End Try
    End Sub

    Public Sub Add(ByVal tObj As Object)
        lstObj.Add(tObj)
        If txtSearch.Text = "" Then Items.Add(tObj) Else txtSearch.Text = ""
    End Sub

    Public Sub AddRange(ByVal tObj As Generic.IEnumerable(Of Object), Optional ByVal Clear As Boolean = False)
        Dim wof As List(Of Object) = tObj.ToList
        lstObj.AddRange(wof)
        If txtSearch.Text = "" Then Items.AddRange(wof, Clear) Else txtSearch.Text = ""
    End Sub

    Public Sub Insert(ByVal tIndex As Integer, ByVal tObj As Object)
        If _Index <> -1 AndAlso tIndex < tIndex Then _Index += 1
        lstObj.Insert(tIndex, tObj)
        If txtSearch.Text = "" Then Items.Insert(tIndex, tObj) Else txtSearch.Text = ""
    End Sub

    Public Overloads Sub Remove(ByVal tIndex As Integer)
        Try
            lstObj.RemoveAt(tIndex)
            If txtSearch.Text = "" Then Items.Remove(tIndex) Else txtSearch.Text = ""
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Overloads Sub Remove(ByVal obj As Object)
        Try
            lstObj.Remove(obj)
            If txtSearch.Text = "" Then Items.Remove(obj) Else txtSearch.Text = ""
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub txtSearch_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        If lstObj Is Nothing OrElse lstObj.Count = 0 Then e.SuppressKeyPress = True
    End Sub

    Private Sub txtSearch_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyUp
        If e.KeyCode = Keys.Return AndAlso Items.Count = 1 Then
            SelectIndex(objList(0).num)
            Exit Sub
        End If
    End Sub

    Private Sub lstItems_ItemsChanged(sender As Object, e As System.EventArgs) Handles lstItems.ItemsChanged
        If Items.Count = 0 Then txtSearch.ShortcutsEnabled = False Else txtSearch.ShortcutsEnabled = True
        If txtSearch.Text = "" Or bWorking Then Exit Sub
        txtSearch.Text = ""
        Items.Clear()
        Items.AddRange(lstObj)
        lstItems.SelectedIndex = _Index
    End Sub

    Private Sub SelectIndex(ByVal tIndex As Integer)
        txtSearch.Text = ""
        bWorking = True
        lstItems.Clear()
        Items.AddRange(lstObj)
        bWorking = False
        bSetting = True
        lstItems.SelectedIndex = tIndex
    End Sub

    Private Sub lstItems_SelectedIndexChanged(sender As Object, e As ListItems.SelectedIndexChangedArgs) Handles lstItems.SelectedIndexChanged
        If txtSearch.Text <> "" Then
            If Not lstItems.SelectedIndex = -1 Then
                Dim lolzo As ObjNum = objList.Find(Function(x) x.obj Is lstItems.SelectedItem)
                SelectIndex(lolzo.num)
            End If
        Else
            If _Index <> lstItems.SelectedIndex Then
                _Index = lstItems.SelectedIndex
                RaiseEvent SelectedIndexChanged(Me, e)
            End If
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As System.EventArgs) Handles txtSearch.TextChanged
        If txtSearch.Text = "" Then 'backspaced/deleted
            bWorking = True
            Items.AddRange(lstObj, True)
            lstItems.SelectedIndex = _Index
            bWorking = False
        Else
            objList.Clear()
            Dim lstObjects As New List(Of Object)
            StuffChanged(lstObjects)
            bWorking = True
            Items.AddRange(lstObjects, True)
            bWorking = False
        End If
    End Sub

    Protected Overridable Sub StuffChanged(ByVal lstObjects As List(Of Object))
        If CaseSensitive Then
            Dim sText As String = txtSearch.Text
            For i As Integer = 0 To lstObj.Count - 1
                If lstObj(i).ToString.Contains(sText) Then
                    objList.Add(New ObjNum(i, lstObj(i)))
                    lstObjects.Add(lstObj(i))
                End If
            Next
        Else
            Dim UText As String = txtSearch.Text.Trim().ToUpper
            For i As Integer = 0 To lstObj.Count - 1
                If lstObj(i).ToString.ToUpper.Contains(UText) Then
                    objList.Add(New ObjNum(i, lstObj(i)))
                    lstObjects.Add(lstObj(i))
                End If
            Next
        End If
    End Sub

    Protected Structure ObjNum
        Public num As Integer, obj As Object
        Public Sub New(ByVal number As Integer, ByVal tobj1 As Object)
            num = number
            obj = tobj1
        End Sub
    End Structure
End Class

Public Class ACTextBox
    Inherits RichTextBox

#Region "Events and Dispose"
    Public Delegate Sub _WordAdded(ByVal sender As Object, ByVal e As WordAddedArgs)
    Private evWordAdded As New List(Of _WordAdded)

    Public Event WordUsedChanged(ByVal sender As Object, ByVal e As WordUsedChangedArgs)
    Public Class WordUsedChangedArgs : Inherits System.EventArgs
        Public Added As Boolean = False, Item As Item
        Public Sub New(ByVal bAdded As Boolean, ByVal tItem As Item)
            Added = bAdded
            Item = tItem
        End Sub
    End Class
    Custom Event WordAdded As _WordAdded
        AddHandler(ByVal value As _WordAdded)
            Me.Events.AddHandler("WordAdded", value)
            evWordAdded.Add(value)
        End AddHandler
        RemoveHandler(ByVal value As _WordAdded)
            Me.Events.RemoveHandler("WordAdded", value)
            evWordAdded.Remove(value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As WordAddedArgs)
            If evWordAdded.Count > 0 Then CType(Me.Events("WordAdded"), _WordAdded).Invoke(sender, e)
        End RaiseEvent
    End Event

    Public Class WordAddedArgs
        Inherits System.EventArgs
        Public Word As String
        Public Sub New(ByVal sWord As String)
            Word = sWord
        End Sub
    End Class
#End Region

    Public lstItems As New List(Of Item)
    Private lstShow As List(Of Item), CWord As String = "", CWordStart As Integer = 0
    Private IL As ItemWindow
    Public StrictWords As Boolean = True, SingleUsage As Boolean = True, AlwaysShotListIfAvailable As Boolean = False, ShowOnRightClick As Boolean = True, EnterOnLostFocus As Boolean = True
    Private bSingleWord As Boolean = False, bNeverAcceptTab As Boolean = False
    Public WordsEntered As Integer = 0, UsedList As New List(Of Item)
    Private CKey As Integer = -1, Handling As Boolean = False, BackspacedChar As Char = Nothing, KDControl As Boolean = False, KDShift As Boolean = False, KDAlt As Boolean = False
    Private Delegates(0) As ArrayList, bLocalChange As Boolean = False
    Public Tag2 As Object

    Public Property SingleWord As Boolean
        Get
            Return bSingleWord
        End Get
        Set(ByVal value As Boolean)
            bSingleWord = value
            If bSingleWord = True Then
                bNeverAcceptTab = True
                SingleUsage = False
                StrictWords = True
            End If
        End Set
    End Property

    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal tColor As Color = Nothing)
        Parent = prnt
        Name = sName
        AcceptsTab = False
        SetBounds(x, y, w, h)
        If tColor <> Nothing Then BackColor = tColor
        IL = New ItemWindow(Me)
        AddHandler Me.KeyDown, AddressOf AC_KeyDown
        AddHandler Me.KeyUp, AddressOf AC_KeyUp
        AddHandler Me.KeyPress, AddressOf AC_KeyPress
        AddHandler Me.LostFocus, AddressOf AC_LostFocus
        AddHandler Me.GotFocus, AddressOf AC_GotFocus
        AddHandler Me.MouseUp, AddressOf AC_MouseUp
    End Sub

    Private Sub ClearUsed()
        For i As Integer = UsedList.Count - 1 To 0 Step -1
            SetUsed(UsedList(i), False)
        Next
        UsedList.Clear()
    End Sub

    Public Sub SetItems(ByVal lst As List(Of Integer))
        ClearUsed()
        If lst Is Nothing OrElse lst.Count = 0 Then Me.Text = "" : Exit Sub
        Dim sLine As New System.Text.StringBuilder()
        For i As Integer = 0 To lst.Count - 1
            Dim tItem As Item = lstItems(lst(i))
            SetUsed(tItem, True)
            sLine.Append(tItem.Text & ", ")
        Next
        sLine = sLine.Remove(sLine.Length - 2, 2)
        Me.Text = sLine.ToString()
    End Sub

    Public Sub SetText(ByVal str As String)
        ClearUsed()
        str = str.Trim
        If str <> "" Then
            If str.EndsWith(",") Then str = str.Remove(str.Length - 1, 1)
            If SingleUsage = False Or SingleWord = True Then
                ScanText(str)
            Else
                ScanText(str)
            End If
        End If
        Me.Text = str
    End Sub

    Private Sub ScanText(ByVal str As String)
        Dim indx As Integer = 0, pindex As Integer = 0, lstWords As New List(Of String)
        Do
            indx = str.IndexOf(",", pindex)
            If indx = -1 Then
                lstWords.Add(str.Substring(pindex).Trim)
                Exit Do
            Else
                lstWords.Add(str.Substring(pindex, indx - pindex).Trim)
            End If
            pindex = indx + 1
        Loop
        If lstWords.Count = 0 Then lstWords.Add(str)
        For i As Integer = 0 To lstWords.Count - 1
            Dim itm As Item = ContainsWord(lstWords(i))
            If Not itm Is Nothing Then SetUsed(itm, True)
        Next
    End Sub

    Private Sub AC_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        If Me.SelectionLength > 0 Then
            SelectionDelete()
        Else
            If ShowOnRightClick = True AndAlso e.Button = Windows.Forms.MouseButtons.Right Then
                RecreateList()
                If lstShow.Count > 0 Then ShowIL()
            ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
                HideIL()
            End If
        End If
    End Sub

    Private Sub AC_GotFocus(ByVal sender As Object, ByVal e As EventArgs)
        If AlwaysShotListIfAvailable = False Then Exit Sub
        If Me.Text = "" Then
            RecreateList()
            ShowIL()
        End If
    End Sub

    Private Sub AC_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
        HideIL(True)
        CWord = ""
        Me.Text = Me.Text.Trim()
        If Me.Text <> "" Then
            If Me.Text.EndsWith(",") Then
                Me.Text = Me.Text.Remove(Me.Text.Length - 1, 1)
            ElseIf EnterOnLostFocus Then
                Dim indx As Integer = GetCurrentWord(Me.Text, Me.Text.Length)
                If indx = -1 Then indx = 0
                If CWord <> "" Then
                    Dim itm As Item = ContainsWord(CWord)
                    If itm Is Nothing Then
                        If lstShow.Count > 0 Then itm = lstShow(0)
                    End If
                    CWord = ""
                    If itm.Used AndAlso SingleUsage Then Exit Sub
                    If Not itm Is Nothing Then
                        Me.Text = Me.Text.Remove(indx)
                        If Me.Text = "" Then Me.Text &= itm.Text Else Me.Text &= ", " & itm.Text
                        SetUsed(itm, True)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub AC_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)
        KDShift = False
        KDControl = False
        KDAlt = False
    End Sub

    Private Sub AC_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        CKey = e.KeyCode
        If CKey = Keys.Delete Then
            If Me.SelectionLength > 0 Then
                SelectionDelete()
                e.SuppressKeyPress = True
            Else
                If Me.SelectionStart <> Me.Text.Length Then
                    If TextChangedEvent(Me.Text.Remove(Me.SelectionStart, 1), Me.SelectionStart) = False Then
                        e.SuppressKeyPress = True
                    End If
                End If
            End If
        ElseIf CKey = Keys.Back Then
            If Me.SelectionLength > 0 Then
                SelectionDelete()
                e.SuppressKeyPress = True
            Else
                If Me.SelectionStart > 0 Then
                    Dim iStart As Integer = Me.SelectionStart
                    BackspacedChar = Me.Text.Chars(Me.SelectionStart - 1)
                    If BackspacedChar = "," Then
                        e.SuppressKeyPress = True
                    Else
                        If TextChangedEvent(Me.Text.Remove(iStart - 1, 1), iStart - 1) = False Then
                            e.SuppressKeyPress = True
                        End If
                    End If
                End If
            End If
        ElseIf CKey = Keys.Enter Then
            e.SuppressKeyPress = True
        ElseIf CKey = Keys.ShiftKey Then
            KDShift = True
        ElseIf CKey = Keys.ControlKey Then
            KDControl = True
        ElseIf CKey = Keys.Alt Then
            KDAlt = True
        End If
    End Sub

    Private Sub AC_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        If lstItems.Count = 0 Then
            e.Handled = True
            Exit Sub
        End If
        If StrictWords = False Then Exit Sub
        If KDAlt = True Or KDControl = True Then Exit Sub
        Handling = True
        Dim iChar As Integer = Asc(e.KeyChar)
        Dim iStart As Integer = Me.SelectionStart
        If iChar = 9 Then 'Tab
            EnterItem()
            e.Handled = True
            If bSingleWord = True Then Me.Parent.SelectNextControl(Me, True, False, False, True)
        ElseIf iChar = 8 Then 'Note: Backspace already executed prior to this sub firing, unlike other key presses
            If Me.Text = "" Then CWord = ""
        ElseIf e.KeyChar = " " Or e.KeyChar = "," Then
            e.Handled = True
            If SingleWord = False AndAlso Me.Text <> "" AndAlso iStart <> 0 Then
                If CWord = "" Then GetCurrentWord(Me.Text, iStart)
                If CWord <> "" Then
                    Dim lolz As Item = ContainsWord(CWord)
                    If Not lolz Is Nothing Then
                        If CWord <> lolz.Text Then Me.Text = Me.Text.Replace(CWord, lolz.Text)
                        Me.Text = Me.Text.Insert(iStart, ", ")
                        Me.SelectionStart = iStart + 2
                        SetUsed(lolz, True)
                        HideIL()
                        RaiseEvent WordAdded(Me, New WordAddedArgs(lolz.Text))
                        CWord = ""
                    End If
                Else
                    Me.SelectionStart = iStart
                End If
            Else
                Me.SelectionStart = iStart
            End If
        Else 'regular key presses
            e.Handled = True
            If Me.Text <> "" AndAlso iStart <> 0 AndAlso Me.Text.Chars(iStart - 1) = "," Then
                Me.Text = Me.Text.Insert(iStart, " ")
                iStart += 1
            End If
            If TextChangedEvent(Me.Text.Insert(Me.SelectionStart, e.KeyChar), iStart + 1) = True Then
                If lstShow.Count = 1 AndAlso lstShow(0).Text.ToUpper = CWord.ToUpper Then
                    Dim iStartso As Integer = iStart - CWord.Length + 1
                    Me.Text = Me.Text.Remove(iStartso, CWord.Length - 1)
                    Me.Text = Me.Text.Insert(iStartso, lstShow(0).Text)
                    RaiseEvent WordAdded(Me, New WordAddedArgs(lstShow(0).Text))
                Else
                    Me.Text = Me.Text.Insert(iStart, e.KeyChar)
                End If
                Me.SelectionStart = iStart + 1
            End If
        End If
        Handling = False
    End Sub

    Public Sub EnterItem(Optional ByVal indx As Integer = 0)
        If bSingleWord = True AndAlso Me.Text <> "" Then
            Me.Text = lstShow(indx).Text
            For i As Integer = UsedList.Count - 1 To 0 Step -1
                SetUsed(UsedList(i), False)
            Next
            SetUsed(lstShow(indx), True)
        Else
            Dim TS As Integer = Me.SelectionStart - CWord.Length
            Me.Text = Me.Text.Remove(TS, CWord.Length)
            If TS <> 0 AndAlso Me.Text.Chars(TS - 1) <> " " Then
                Me.Text = Me.Text.Insert(TS, " ")
                TS += 1
                If TS > 2 AndAlso Me.Text.Chars(TS - 3) <> "," Then
                    Me.Text = Me.Text.Insert(TS - 1, ",")
                    TS += 1
                End If
            End If
            If SingleWord = True Then
                Me.Text = Me.Text.Insert(TS, lstShow(indx).Text)
                Me.SelectionStart = TS + lstShow(indx).Text.Length
            Else
                Me.Text = Me.Text.Insert(TS, lstShow(indx).Text & ", ")
                Me.SelectionStart = TS + lstShow(indx).Text.Length + 2
            End If
            SetUsed(lstShow(indx), True)
            CWord = ""
            WordsEntered += 1
        End If
        RaiseEvent WordAdded(Me, New WordAddedArgs(lstShow(indx).Text))
        HideIL()
    End Sub

    Private Sub SelectionDelete()
        Dim lolz As String = Me.SelectedText
        Dim omg As Item = ContainsWord(lolz)
        If Not omg Is Nothing Then
            SetUsed(omg, False)
            Dim iStart As Integer = Me.SelectionStart
            Me.Text = Me.Text.Remove(iStart, Me.SelectionLength)
            Me.Text = Me.Text.Trim.Replace(" ,", " ").Replace("  ", " ")
            If Me.Text.StartsWith(",") Then Me.Text = Me.Text.Remove(0, 1).Trim
            Me.SelectionStart = iStart
            CWord = ""
            HideIL()
        End If
    End Sub

    Private Function ContainsWord(ByVal sWord As String) As Item
        sWord = sWord.ToUpper
        For i As Integer = 0 To lstItems.Count - 1
            If lstItems(i).Text.ToUpper = sWord Then
                Return lstItems(i)
            End If
        Next
        Return Nothing
    End Function

    Private Function GetCurrentWord(ByVal sText As String, ByVal iStart As Integer) As Integer
        CWord = ""
        Dim indx As Integer = 0
        For i As Integer = iStart - 1 To 0 Step -1
            indx = i
            If i = 0 Then
                CWord = sText.Substring(0, iStart).Trim
                Exit For
            ElseIf sText.Substring(i, 1) = "," Then
                CWord = sText.Substring(i + 1, iStart - i - 1).Trim
                Exit For
            End If
        Next
        For i As Integer = iStart To sText.Length - 1
            If i = sText.Length - 1 Then
                CWord &= sText.Substring(iStart, sText.Length - iStart).Trim
                Exit For
            ElseIf sText.Substring(i, 1) = "," Then
                CWord &= sText.Substring(iStart, i - iStart).Trim
                Exit For
            End If
        Next
        Return indx
    End Function

    Private Function TextChangedEvent(ByVal sText As String, ByVal iStart As Integer) As Boolean
        If sText = "" Or lstItems.Count = 0 Then
            HideIL()
            Return True
        End If
        If sText.Length > 1 Then
            Dim omg As String = ""
        End If
        Dim OWord As String = CWord     'Get Current Word
        GetCurrentWord(sText, iStart)
        Dim UWord As String = CWord.ToUpper
        Dim NList As New List(Of Item)
        For i As Integer = 0 To lstItems.Count - 1
            If SingleUsage = True AndAlso lstItems(i).Used = True Then Continue For
            If lstItems(i).Text.ToUpper.Contains(UWord) Then
                NList.Add(lstItems(i))
            End If
        Next
        If NList.Count > 0 Then
            lstShow = NList
            ShowIL()
            Return True
        Else
            CWord = OWord
            Return False
        End If
    End Function

    Private Sub RecreateList()
        lstShow = New List(Of Item)
        For i As Integer = 0 To lstItems.Count - 1
            If lstItems(i).Used = True Then Continue For
            lstShow.Add(lstItems(i))
        Next
    End Sub

    Private Sub HideIL(Optional ByVal HideOverride As Boolean = False)
        If HideOverride = False AndAlso AlwaysShotListIfAvailable = True Then
            RecreateList()
            If lstShow.Count > 0 Then ShowIL()
        Else
            Me.AcceptsTab = False
            IL.Visible = False
        End If
    End Sub

    Private Sub ShowIL()
        IL.SetItems(lstShow)
        If bNeverAcceptTab = False Then Me.AcceptsTab = True
    End Sub

    Public Overridable Overloads Sub SetAutoCompletionList(ByVal lst As List(Of String), ByVal bStrictWords As Boolean, Optional ByVal bSingleUsage As Boolean = True, Optional ByVal tSingleWord As Boolean = False, Optional ByVal bAlwaysShowListIfAvailable As Boolean = False, Optional ByVal bShowListOnRightClick As Boolean = True)
        SetAutoCompletionList(lst)
        StrictWords = bStrictWords
        SingleUsage = bSingleUsage
        bSingleWord = tSingleWord
        AlwaysShotListIfAvailable = bAlwaysShowListIfAvailable
        ShowOnRightClick = bShowListOnRightClick
    End Sub

    Public Overridable Overloads Sub SetAutoCompletionList(ByVal lst As List(Of String))
        lstItems = New List(Of Item)
        'UsedList.Clear()
        If lst Is Nothing Then Exit Sub
        For i As Integer = 0 To lst.Count - 1
            lstItems.Add(New Item(lst(i), i))
        Next
    End Sub

    Public Function MeasureStrHeight(ByRef str As String) As Integer
        Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
            Dim SFormat As New System.Drawing.StringFormat
            Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
            Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
            SFormat.SetMeasurableCharacterRanges(range)
            Dim regions() As Region = gx.MeasureCharacterRanges(str, Font, rect, SFormat)
            rect = regions(0).GetBounds(gx)
            Return CInt(rect.Bottom + 1)
        End Using
        Return -1
    End Function

    Public Sub RemoveACItem(ByVal tIndex As Integer)
        If lstItems(tIndex).Used Then
            Dim txt As String = lstItems(tIndex).Text
            Dim indx As Integer = 0, dlength As Integer = 0
            Do
                indx = Text.IndexOf(txt, indx)
                If indx = -1 Then Exit Do 'check left to make sure whole word
                If indx = 0 Then
                ElseIf Text.Substring(indx - 2, 2) = ", " Then
                Else
                    indx += 1
                    Continue Do
                End If
                If indx + txt.Length = Text.Length Then
                    dlength = txt.Length
                ElseIf Text.Substring(indx + txt.Length, 2) = ", " Then
                    dlength = txt.Length + 2
                Else
                    indx += 1
                    Continue Do
                End If
                Me.Text = Me.Text.Remove(indx, dlength)
            Loop
            SetUsed(lstItems(tIndex), False)
        End If
        lstItems.RemoveAt(tIndex)
        For i As Integer = tIndex To lstItems.Count - 1
            lstItems(i).Index -= 1
        Next
    End Sub

    Public Sub InsertACItem(ByVal tIndex As Integer, ByVal txt As String)
        For i As Integer = tIndex To lstItems.Count - 1
            lstItems(i).Index += 1
        Next
        lstItems.Insert(tIndex, New Item(txt, tIndex))
    End Sub

    Public Sub SetUsed(ByVal tItem As Item, ByVal tUsed As Boolean)
        If tItem.Used = False AndAlso tUsed = True Then
            UsedList.Add(tItem)
            tItem.Used = tUsed
            RaiseEvent WordUsedChanged(Me, New WordUsedChangedArgs(True, tItem))
        ElseIf tItem.Used = True AndAlso tUsed = False Then
            UsedList.Remove(tItem)
            tItem.Used = tUsed
            RaiseEvent WordUsedChanged(Me, New WordUsedChangedArgs(False, tItem))
        End If
    End Sub

    Public Class Item
        Public Text As String, Used As Boolean = False, Index As Integer = 0
        Public Sub New(ByVal sText As String, ByVal tIndex As Integer)
            Text = sText
            Index = tIndex
        End Sub
    End Class

    Public Class ItemWindow
        Inherits System.Windows.Forms.Control

        Public tb As ACTextBox, lstItems As List(Of ACTextBox.Item)
        Public Shared MyFont As Font = Nothing, FontBrush As SolidBrush, HighlightBrush As SolidBrush, BackBrush As New SolidBrush(Color.LightGray)
        Private Shared ItemHeight As Integer = 0, VBuffer As Integer = 0, HBuffer As Integer = 2
        Private HighlightedItem As Integer = -1, InvRect As Rectangle, LongestWord As Integer = 0

        Public Sub New(ByVal ACTB As ACTextBox, Optional ByVal items As List(Of ACTextBox.Item) = Nothing)
            Visible = False
            tb = ACTB
            Parent = tb.FindForm
            If MyFont Is Nothing Then SetFont(New Font("Microsoft Sans MS", 8, FontStyle.Regular))
            AddHandler Me.Paint, AddressOf IL_Paint
            If Not items Is Nothing Then SetItems(items)
        End Sub

        Private Sub IL_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseMove
            Dim HI As Integer = CInt(Math.Floor(e.Y / ItemHeight))
            If HI <> HighlightedItem Then
                HighlightedItem = HI
                Me.Refresh()
            End If
        End Sub

        Private Sub IL_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseDown
            tb.EnterItem(HighlightedItem)
            HighlightedItem = -1
            Me.Visible = False
        End Sub

        Public Function MeasureStrWidth(ByRef str As String) As Integer
            Try
                Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
                    Dim SFormat As New System.Drawing.StringFormat
                    Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
                    Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
                    SFormat.SetMeasurableCharacterRanges(range)
                    Dim regions() As Region = gx.MeasureCharacterRanges(str, Font, rect, SFormat)
                    rect = regions(0).GetBounds(gx)
                    Return CInt(rect.Right + 1)
                End Using
            Catch ex As Exception
                Return -1
            End Try
        End Function

        Public Sub SetItems(ByVal items As List(Of ACTextBox.Item))
            lstItems = items
            LongestWord = 0
            Dim zasto As Form = tb.FindForm
            Dim loc As Point = zasto.PointToClient(tb.Parent.PointToScreen(tb.Location))
            Dim x As Integer = loc.X, y As Integer = loc.Y + tb.Height, w As Integer, h As Integer, DHB As Integer
            For i As Integer = 0 To lstItems.Count - 1
                LongestWord = Math.Max(LongestWord, MeasureStrWidth(lstItems(i).Text))
            Next
            If LongestWord > tb.Width Then w = Math.Min(LongestWord, zasto.ClientSize.Width - 10) Else w = tb.Width
            h = (VBuffer * 2) + (ItemHeight * lstItems.Count)
            DHB = loc.Y + tb.Height + h
            If h >= zasto.ClientSize.Height Then
                y = 0

            ElseIf DHB > zasto.ClientSize.Height Then
                If loc.Y < h Then 'center it
                    y = CInt((zasto.ClientSize.Height - h) / 2)
                Else 'top orient it
                    y = loc.Y - h
                End If
            End If

            DHB = loc.X + w
            If DHB > zasto.ClientSize.Width Then
                x = Math.Max(5, zasto.ClientSize.Width - LongestWord)
            End If
            Me.SetBounds(x, y, w, h)
            Me.BringToFront()
            Visible = True
            Me.Refresh()
        End Sub

        Public Sub SetFont(ByVal sFont As Font, Optional ByVal sFontColor As Color = Nothing, Optional ByVal sHighlightColor As Color = Nothing)
            MyFont = sFont
            If sFontColor <> Nothing Then FontBrush = New SolidBrush(sFontColor) Else If FontBrush Is Nothing Then FontBrush = New SolidBrush(Color.DarkRed) Else 
            If sHighlightColor <> Nothing Then HighlightBrush = New SolidBrush(Color.FromArgb(100, sHighlightColor)) Else If HighlightBrush Is Nothing Then HighlightBrush = New SolidBrush(Color.FromArgb(100, Color.LightBlue))
            ItemHeight = tb.MeasureStrHeight("j0yf")
        End Sub

        Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)
        End Sub

        Private Sub IL_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
            If lstItems.Count = 0 Then
                e.Graphics.FillRectangle(BackBrush, Me.ClientRectangle)
            Else
                Using img As New Bitmap(Me.Width, Me.Height)
                    Using gx As Graphics = Graphics.FromImage(img)
                        gx.FillRectangle(BackBrush, Me.ClientRectangle)
                        For i As Integer = 0 To lstItems.Count - 1
                            gx.DrawString(lstItems(i).Text, MyFont, FontBrush, HBuffer, VBuffer + (i * ItemHeight))
                        Next
                        If HighlightedItem <> -1 Then gx.FillRectangle(HighlightBrush, 0, VBuffer + (HighlightedItem * ItemHeight), Me.Width, ItemHeight)
                    End Using
                    e.Graphics.DrawImage(img, 0, 0)
                End Using
            End If
        End Sub

    End Class

End Class

Public Class PromptBox : Inherits Form
    Public Buttons As New List(Of String), panButtons As New Panel, BResult As Integer = 0, Title As String, CenterForm As Form
    Private ButtonPressed As Boolean = False

    Public Sub New(Optional ByVal tTitle As String = "", Optional ByVal tCenterForm As Form = Nothing, Optional ByVal btns As List(Of String) = Nothing)
        Me.MaximizeBox = False : Me.MinimizeBox = False : Me.ShowInTaskbar = False
        If btns Is Nothing OrElse btns.Count = 0 Then Buttons = New List(Of String) From {"OK", "Cancel"} Else Buttons = btns
        Title = tTitle
        If tCenterForm IsNot Nothing Then
            CenterForm = tCenterForm
            CenterForm.AddOwnedForm(Me)
        End If
    End Sub

    Public Function ShowPromptBox() As Integer
        Me.ShowDialog()
        Return BResult
    End Function

    Private Sub PromptBox_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If CenterForm IsNot Nothing Then CenterForm.RemoveOwnedForm(Me)
    End Sub

    Private Sub PromptBox_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim btnWidth As Integer = 0, btns As New List(Of Button)
        Dim iRect As Rectangle = New Rectangle(90000, 90000, 0, 0)
        For Each ctrl As Control In Me.Controls
            If Not ctrl.Visible Then Continue For
            iRect.X = Math.Min(iRect.X, ctrl.Left)
            iRect.Y = Math.Min(iRect.Y, ctrl.Top)
            iRect.Height = Math.Max(iRect.Height, ctrl.Bottom)
            iRect.Width = Math.Max(iRect.Width, ctrl.Right)
        Next
        For i As Integer = 0 To Buttons.Count - 1
            Dim btn As New Button()
            btn.AutoSize = True : btn.Text = Buttons(i)
            Dim sSizeos As Integer = btn.Width : btn.AutoSize = False
            btn.Size = New Size(Math.Min(86, sSizeos + 5), 24)
            btnWidth += btn.Width + 10
            btn.Tag = i + 1
            AddHandler btn.Click, AddressOf btnClick
            btns.Add(btn)
        Next
        btnWidth += 50
        If iRect.X < 1000 Then iRect.Width += iRect.X
        iRect.Width = Math.Max(iRect.Width, btnWidth)
        panButtons.Parent = Me : panButtons.BackColor = Color.FromKnownColor(KnownColor.Control)
        panButtons.SetBounds(0, iRect.Height, iRect.Width, 49) : iRect.Height += 49
        Dim iLeft As Integer = panButtons.Width - btns(btns.Count - 1).Width - 10
        For i As Integer = btns.Count - 1 To 0 Step -1
            btns(i).Parent = panButtons
            btns(i).Location = New Point(iLeft, 14)
            iLeft = iLeft - btns(Buttons.Count - 1).Width - 10
        Next
        Dim w As Integer = iRect.Width + 16, h As Integer = iRect.Height + 40
        Me.Text = Title
        If CenterForm Is Nothing Then
            Me.SetBounds(CInt((Screen.PrimaryScreen.WorkingArea.Width - w) / 2), CInt((Screen.PrimaryScreen.WorkingArea.Height - h) / 2), w, h)
        Else
            Me.SetBounds(CenterForm.Left + CInt((CenterForm.Width - w) / 2), CenterForm.Top + CInt((CenterForm.Height - h) / 2), w, h)
        End If
    End Sub

    Private Sub btnClick(ByVal sender As Object, ByVal e As EventArgs)
        Dim btn As Button = DirectCast(sender, Button)
        If btn.Text = "Cancel" Then BResult = 0 Else BResult = CInt(btn.Tag)
        Me.Close()
    End Sub

    Public Enum Result
        Cancel = 0
        Ok = 1
    End Enum

End Class

Public Class ProgressBarX : Inherits Control
#Region "Events"
    Public Delegate Sub _BarClicked(ByVal sender As Object, ByVal e As BarClickedArgs)
    Public Delegate Sub _ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Private evBarClicked As New List(Of _BarClicked), evValueChanged As New List(Of _ValueChanged)

    Private Shadows disposed As Boolean = False
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            For i As Integer = evBarClicked.Count - 1 To 0 Step -1
                RemoveHandler BarClicked, evBarClicked(i)
            Next
            For i As Integer = evValueChanged.Count - 1 To 0 Step -1
                RemoveHandler ValueChanged, evValueChanged(i)
            Next
        End If
        Me.disposed = True
    End Sub

    Public Class BarClickedArgs : Inherits System.EventArgs
        Public Handled As Boolean = False, ValueChanging As Boolean, Value As Integer
        Public Sub New(ByVal tValueChanging As Boolean, ByVal tValue As Integer)
            ValueChanging = tValueChanging
            Value = tValue
        End Sub
    End Class

    Public Custom Event BarClicked As _BarClicked
        AddHandler(ByVal value As _BarClicked)
            Me.Events.AddHandler("BarClicked", value)
            evBarClicked.Add(value)
        End AddHandler
        RemoveHandler(ByVal value As _BarClicked)
            Me.Events.RemoveHandler("BarClicked", value)
            evBarClicked.Remove(value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As BarClickedArgs)
            If evBarClicked.Count > 0 Then CType(Me.Events("BarClicked"), _BarClicked).Invoke(sender, e)
        End RaiseEvent
    End Event

    Public Custom Event ValueChanged As _ValueChanged
        AddHandler(ByVal value As _ValueChanged)
            Me.Events.AddHandler("ValueChanged", value)
            evValueChanged.Add(value)
        End AddHandler
        RemoveHandler(ByVal value As _ValueChanged)
            Me.Events.RemoveHandler("ValueChanged", value)
            evValueChanged.Remove(value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As System.EventArgs)
            If evValueChanged.Count > 0 Then CType(Me.Events("ValueChanged"), _ValueChanged).Invoke(sender, e)
        End RaiseEvent
    End Event

#End Region

    Public Minimum, Maximum, UpdateStep As Integer
    Private GradientColor As Color, bMouseDown As Boolean = False

    Private _Value As Integer
    Public Property Value() As Integer
        Get
            Return _Value
        End Get
        Set(ByVal value As Integer)
            If _Value <> value Then
                _Value = value
                Me.Invalidate()
                RaiseEvent ValueChanged(Me, New System.EventArgs)
            End If
        End Set
    End Property

    Public Sub New(ByVal tParent As Control, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.ResizeRedraw, True)
        UpdateStyles()
        Parent = tParent
        SetBounds(x, y, w, h)
        Minimum = 0
        Maximum = 100
        UpdateStep = 1
        BackColor = Color.LightGray
        ForeColor = Color.FromArgb(255, 43, 134, 218)
    End Sub

    Protected Overrides Sub OnMouseDown(e As System.Windows.Forms.MouseEventArgs)
        Mousez(e)
        bMouseDown = True
        MyBase.OnMouseDown(e)
    End Sub

    Private Sub Mousez(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim iValue As Integer = 0
        If e.X > -1 AndAlso Maximum > 0 Then
            iValue = Math.Min(Maximum, CInt((e.X / Me.ClientSize.Width) * Maximum))
        End If
        Dim args As New BarClickedArgs(True, iValue)
        If iValue = _Value Then args.ValueChanging = False
        RaiseEvent BarClicked(Me, args)
        If Not args.Handled Then Value = iValue
    End Sub

    Protected Overrides Sub OnMouseUp(e As System.Windows.Forms.MouseEventArgs)
        bMouseDown = False
        MyBase.OnMouseUp(e)
    End Sub

    Protected Overrides Sub OnMouseMove(e As System.Windows.Forms.MouseEventArgs)
        If bMouseDown Then
            If MouseButtons = Windows.Forms.MouseButtons.Left Then
                Mousez(e)
            Else
                bMouseDown = False
            End If
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnForeColorChanged(e As System.EventArgs)
        If ForeColor = Nothing Then
            ForeColor = Color.FromKnownColor(KnownColor.Control)
        Else
            GradientColor = Color.FromArgb(255, Math.Max(0, ForeColor.R - 3), Math.Max(0, ForeColor.G - 9), Math.Max(0, ForeColor.B - 14))
            MyBase.OnForeColorChanged(e)
        End If
    End Sub

    Protected Overrides Sub OnPaintBackground(pevent As System.Windows.Forms.PaintEventArgs)
    End Sub

    Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
        Dim iWidth As Integer = 0, iHeight As Integer = Me.ClientSize.Height - 4, iTop As Integer = 2
        Dim bbrush As New SolidBrush(Parent.BackColor)
        e.Graphics.FillRectangle(bbrush, 0, 0, Me.ClientSize.Width, iTop)
        If Maximum > 0 Then
            iWidth = CInt((Value / Maximum) * Me.ClientSize.Width)
            Dim gbrush As New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(0, iHeight), Me.ForeColor, Me.GradientColor)
            e.Graphics.FillRectangle(gbrush, 0, iTop, iWidth, iHeight)
        End If
        e.Graphics.FillRectangle(New SolidBrush(Me.BackColor), iWidth, iTop, Me.ClientSize.Width - iWidth, iHeight)
        e.Graphics.FillRectangle(bbrush, 0, Me.ClientSize.Height - iTop, Me.ClientSize.Width, iTop)
    End Sub

End Class

Public Class ImageViewer
    Inherits System.Windows.Forms.Form
    Implements IDisposable

#Region "Events and Dispose"
    Private Shadows disposed As Boolean = False
    Private Delegates(2) As ArrayList
    Public Delegate Sub _ImageChanged(ByVal sender As Object, ByVal e As ImageChangedEventArgs)
    Private evImageChanged As New List(Of _ImageChanged)

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            If Not Delegates(0) Is Nothing Then
                For i As Integer = evImageChanged.Count - 1 To 0 Step -1
                    RemoveHandler ImageChanged, evImageChanged(i)
                Next
                Delegates(0) = Nothing
            End If
        End If
        Me.disposed = True
    End Sub
    Public Custom Event ImageChanged As _ImageChanged
        AddHandler(ByVal value As _ImageChanged)
            Me.Events.AddHandler("ImageChanged", value)
            evImageChanged.Add(value)
        End AddHandler
        RemoveHandler(ByVal value As _ImageChanged)
            Me.Events.RemoveHandler("ImageChanged", value)
            evImageChanged.Remove(value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As ImageChangedEventArgs)
            If evImageChanged.Count > 0 Then CType(Me.Events("ImageChanged"), _ImageChanged).Invoke(sender, e)
        End RaiseEvent
    End Event

    Public Event ImageChanging(ByVal sender As Object, ByVal e As ImageChangingArgs)
    Public Event ImageEndReached(ByVal sender As Object, ByVal e As ImageEndReachedArgs)
    Public Event ImageDeletion(ByVal sender As Object, ByVal e As ImageDeletionArgs)
    Public Event ViewerKeyDown(ByVal sender As Object, ByVal e As ViewerKeyDownArgs)

    Public Class ViewerKeyDownArgs : Inherits System.EventArgs
        Public KChar As Char, Events As KeyEventArgs
        Public Sub New(ByVal tChar As Char, ByVal eEvent As KeyEventArgs)
            KChar = tChar
            Events = eEvent
        End Sub
    End Class

    Public Class ImageChangingArgs : Inherits System.EventArgs
        Public Direction As Boolean, Index As Integer, Handled As Boolean
        Public Sub New(ByVal tDirection As Boolean, ByVal tIndex As Integer)
            Direction = tDirection
            Index = tIndex
        End Sub
    End Class

    Public Class ImageEndReachedArgs : Inherits System.EventArgs
        Public Handled As Boolean, Direction As String = ""
        Public Sub New(ByVal sDirection As String)
            Handled = False
            Direction = sDirection
        End Sub
    End Class

    Public Class ImageDeletionArgs : Inherits System.EventArgs
        Public Handled As Boolean, FileName As String = ""
        Public Sub New(ByVal sFileName As String)
            Handled = False
            FileName = sFileName
        End Sub
    End Class

    Public Class ImageChangedEventArgs
        Inherits System.EventArgs
        Public SetByProgram As Boolean, PreviousDirectory As String, FileName As String, Index As Integer
        Public Sub New(ByVal tFile As String, ByVal tIndex As Integer, ByVal tPreviousDirectory As String, ByVal tSetByProgram As Boolean)
            FileName = tFile
            Index = tIndex
            PreviousDirectory = tPreviousDirectory
            SetByProgram = tSetByProgram
        End Sub
    End Class
#End Region

    Public img As Bitmap, ZImage As Bitmap, BGBrush As New SolidBrush(Color.FromArgb(255, 199, 228, 197)), imgList As New List(Of String)
    Public ZoomMode As ZoomModes = 0, di As System.IO.DirectoryInfo, CurrentImage As String, PreviousImage As String, ImageName As String
    Public WithEvents CM As New ContextMenuStrip()
    Public ZoomFixture As Single = 1, ZIncrement As Decimal = CDec(0.3)
    Private TempRect As Rectangle, DrawRect As Rectangle, CenterRect As Rectangle, BGRects As List(Of Rectangle)
    Private CZoom As Single = 1, ZMax As Single = 3.1, ZMin As Single = 1
    Private MHandlersAdded As Boolean = False, HScrollMax As Integer = 0, VScrollMax As Integer = 0, PMouseX As Integer = -1, PMouseY As Integer = -1
    Private EventMouseDown As New MouseEventHandler(AddressOf IV_MouseDown), EventMouseUp As New MouseEventHandler(AddressOf IV_MouseUp), EventMouseMove As New MouseEventHandler(AddressOf IV_MouseMove)
    Private Comparerz As New MyComparer
    Private TImage As Bitmap = New Bitmap(1, 1), ZDifX As Integer = 0, ZDifY As Integer = 0
    Private MButtonDown As MouseButtons, bResizable As Boolean = False
    Private bMinimizing As Boolean = False, bSpecialList As Boolean = False
    Private _Animating As Boolean = False, AEvent As New EventHandler(AddressOf Me.OnFrameChanged)

    Private _indx As Integer
    Public Property indx As Integer
        Get
            Return _indx
        End Get
        Set(ByVal value As Integer)
            _indx = value
        End Set
    End Property

    Private _StandAloneMode As Boolean, _ImageCount As Integer = 0
    Public ReadOnly Property ImageCount() As Integer
        Get
            Return _ImageCount
        End Get
    End Property

    Public Sub SetToStandAloneList(ByVal tCount As Integer)
        _ImageCount = tCount
        _StandAloneMode = True
        bSpecialList = True
    End Sub

    Public Enum ZoomModes
        Mouse = 0
        Center = 1
    End Enum

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            cp.Style = cp.Style Or &H2000000 And Not 33554432
            Return cp
        End Get
    End Property

    Public Sub New(ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal Resizable As Boolean = False)
        TempRect = New Rectangle(x, y, w, h)
        bResizable = Resizable
        Me.SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.ResizeRedraw, True)
        Me.UpdateStyles()
        AddHandler Me.MouseUp, AddressOf IV_MouseUp
        AddHandler Me.MouseDown, AddressOf IV_MouseDown
    End Sub

    Private Sub CM_ItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles CM.ItemClicked
        If e.ClickedItem.Text = "Open File Location" Then
            OpenFileInExplorer(CurrentImage)
        ElseIf e.ClickedItem.Text = "Delete File" Then
            DeleteFile()
        End If
    End Sub

    Private Sub DeleteFile()
        Dim sFile As String = di.FullName & "\" & imgList(_indx)
        Dim ev As New ImageDeletionArgs(sFile)
        RaiseEvent ImageDeletion(Me, ev)
        If ev.Handled = False Then
            Try
                My.Computer.FileSystem.DeleteFile(sFile, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
                If Not System.IO.File.Exists(sFile) Then
                    RemoveCurrentImage()
                End If
            Catch ex As Exception
                If Not ex.Message = "The operation was canceled." Then MsgBox("Error deleting file: " & ex.Message)
            End Try
        End If
    End Sub

    Public Function OpenFileInExplorer(ByRef sFile As String) As Boolean
        Try
            Call Shell("explorer /select," & sFile, AppWinStyle.NormalFocus)
            Return True
        Catch ex As Exception
            MsgBox("Error opening file in explorer: " & ex.Message)
            Return False
        End Try
    End Function

    Public Sub RemoveCurrentImage()
        imgList.RemoveAt(_indx)
        _indx -= 1
        NextImage(True)
    End Sub

    Protected Overrides Sub SetBoundsCore(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal specified As System.Windows.Forms.BoundsSpecified)
        MyBase.SetBoundsCore(x, y, width, height, specified)
        TempRect.Width = Me.ClientSize.Width
        TempRect.Height = Me.ClientSize.Height
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        If Me.ClientSize.Width = 0 Or Me.ClientSize.Height = 0 Then Exit Sub
        If Me.WindowState = FormWindowState.Minimized Then
            bMinimizing = True
            Exit Sub
        End If
        If bMinimizing = True Then
            bMinimizing = False
        Else
            TempRect.Width = Me.ClientSize.Width
            TempRect.Height = Me.ClientSize.Height
            SetImageWorker(di.FullName & "\" & imgList(_indx))
        End If
    End Sub

    Private Sub DisposeImages()
        If Not img Is Nothing Then
            img.Dispose()
            img = Nothing
        End If
        If Not ZImage Is Nothing Then
            ZImage.Dispose()
            ZImage = Nothing
        End If
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        Me.Visible = False
        DisposeImages()
        e.Cancel = True
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, False) ' Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, False)
        Me.SetBounds(TempRect.Left, TempRect.Top, TempRect.Width, TempRect.Height)
        If bResizable Then Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable Else Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle

    End Sub

    Protected Overrides Sub OnLostFocus(e As EventArgs)
        MyBase.OnLostFocus(e)
        MButtonDown = Nothing
    End Sub

    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        MButtonDown = Nothing
        If e.KeyCode = Keys.Left Then
            NextImage(False)
        ElseIf e.KeyCode = Keys.Right Then
            NextImage(True)
        ElseIf e.KeyCode = Keys.Delete Then
            DeleteFile()
        ElseIf e.KeyCode = Keys.ShiftKey OrElse e.KeyCode = Keys.ControlKey OrElse e.KeyCode = 18 Then '18 = alt

        Else
            RaiseEvent ViewerKeyDown(Me, New ViewerKeyDownArgs(KeyCodeToAscii.GetAsciiCharacter(e.KeyCode), e))
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
        MyBase.OnKeyUp(e)
    End Sub

    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        MyBase.OnMouseWheel(e)
        If img Is Nothing Then Exit Sub
        If Not Me.ClientRectangle.Contains(e.X, e.Y) Then Exit Sub
        Dim OZ As Single = CZoom
        If e.Delta > 0 Then 'scroll up
            CZoom = CSng(Math.Round(Math.Min(CZoom + ZIncrement, ZMax), 1))
        Else
            CZoom = CSng(Math.Round(Math.Max(CZoom - ZIncrement, ZMin), 1))
        End If
        If CZoom <> OZ Then Zoom(CZoom, Math.Min(Math.Max(e.X, DrawRect.Left), DrawRect.Right), Math.Min(Math.Max(e.Y, DrawRect.Top), DrawRect.Bottom))
    End Sub

    Private Sub IV_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        MButtonDown = e.Button
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        PMouseX = e.X
        PMouseY = e.Y
    End Sub

    Private Sub IV_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        If MButtonDown = Nothing Then Exit Sub
        Dim lolzo As String = CurrentImage
        If MButtonDown = Windows.Forms.MouseButtons.XButton1 Then 'Back
            NextImage(False)
        ElseIf MButtonDown = Windows.Forms.MouseButtons.XButton2 Then 'Forward
            NextImage(True)
        ElseIf MButtonDown = Windows.Forms.MouseButtons.Right Then
            CM.Items.Clear()
            CM.Items.Add("Open File Location")
            CM.Items.Add("Delete File")
            CM.Show(New Point(Cursor.Position.X, Cursor.Position.Y))
        End If
        MButtonDown = Nothing
    End Sub

    Private Sub IV_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        If MButtonDown = Windows.Forms.MouseButtons.Left Then 'click and drag operation
            Dim ORect As Rectangle = DrawRect
            Dim nx As Integer = e.X - PMouseX, ny As Integer = e.Y - PMouseY
            If VScrollMax <> 0 Then
                DrawRect.Y = Math.Max(Math.Min(0, DrawRect.Y + ny), VScrollMax)
                PMouseY = e.Y
            End If
            If HScrollMax <> 0 Then
                DrawRect.X = Math.Max(Math.Min(0, DrawRect.X + nx), HScrollMax)
                PMouseX = e.X
            End If
            If DrawRect <> ORect Then Me.Refresh()
        End If
    End Sub

    Private Sub Zoom(ByVal z As Single, ByVal x As Integer, ByVal y As Integer)
        If ZoomMode = ZoomModes.Center Then
            Dim NR As New Rectangle(CenterRect.X, CenterRect.Y, CInt(CenterRect.Width * z), CInt(CenterRect.Height * z))
            Dim WDif As Integer = NR.Width - CenterRect.Width, HDif As Integer = NR.Height - CenterRect.Height
            NR.X -= CInt(WDif / 2)
            NR.Y -= CInt(HDif / 2)
            DrawRect = NR
        Else
            Dim NR As New Rectangle(DrawRect.X, DrawRect.Y, CInt(CenterRect.Width * z), CInt(CenterRect.Height * z))
            Dim WDif As Integer = NR.Width - CenterRect.Width, HDif As Integer = NR.Height - CenterRect.Height
            If NR.Width > Me.TempRect.Width Then
                Dim MouseP As Single = CInt(x / Me.TempRect.Width)
                If NR.Width >= DrawRect.Width Then NR.X = Math.Min(Math.Max(DrawRect.X - CInt(MouseP * ZDifX), Me.TempRect.Width - NR.Width), 0) Else NR.X = Math.Min(Math.Max(DrawRect.X + CInt(MouseP * ZDifX), Me.TempRect.Width - NR.Width), 0)
            Else
                NR.X = CenterRect.X - CInt(WDif / 2)
            End If
            If NR.Height > Me.TempRect.Height Then
                Dim MouseP As Single = CInt(y / Me.TempRect.Height)
                If NR.Height >= DrawRect.Height Then NR.Y = Math.Min(Math.Max(DrawRect.Y - CInt(MouseP * ZDifY), Me.TempRect.Height - NR.Height), 0) Else NR.Y = Math.Min(Math.Max(DrawRect.Y + CInt(MouseP * ZDifY), Me.TempRect.Height - NR.Height), 0)
            Else
                NR.Y = CenterRect.Y - CInt(HDif / 2)
            End If
            DrawRect = NR
        End If
        If DrawRect.Height > Me.TempRect.Height Or DrawRect.Width > Me.TempRect.Width Then 'Enable horizontal moving
            VScrollMax = Math.Min(0, Me.TempRect.Height - DrawRect.Height)
            HScrollMax = Math.Min(0, Me.TempRect.Width - DrawRect.Width)
            If MHandlersAdded = False Then
                MHandlersAdded = True
                AddHandler Me.MouseMove, EventMouseMove
            End If
        Else
            If MHandlersAdded = True Then
                RemoveHandler Me.MouseMove, EventMouseMove
                MHandlersAdded = False
            End If
        End If
        CreateZImage(DrawRect)
        GenerateBGRects()
        Me.Refresh()
    End Sub

    Private Sub GenerateBGRects()
        BGRects = New List(Of Rectangle)
        If DrawRect.Left > 0 Then BGRects.Add(New Rectangle(0, -1, DrawRect.Left, Me.TempRect.Height + 1))
        If DrawRect.Right < Me.TempRect.Width Then BGRects.Add(New Rectangle(DrawRect.Right - 1, -1, Me.TempRect.Width - DrawRect.Right + 1, Me.TempRect.Height + 1))
        If DrawRect.Top > 0 Then BGRects.Add(New Rectangle(DrawRect.Left - 1, 0, DrawRect.Width + 2, DrawRect.Top))
        If DrawRect.Bottom < Me.TempRect.Height Then BGRects.Add(New Rectangle(DrawRect.Left - 1, DrawRect.Bottom, DrawRect.Width + 2, Me.TempRect.Height - DrawRect.Bottom))
    End Sub

    Public Sub RefreshList(ByVal tIndex As Integer)
        PopulateImageList("")
        If tIndex > -1 AndAlso tIndex < _ImageCount Then
            _indx = tIndex
            Me.Text = imgList(_indx)
            'SetImageWorker(di.FullName & "\" & imgList(_indx), True)
        End If
    End Sub

    Public Sub SetImageList(ByVal sList As List(Of String))
        bSpecialList = True
        di = Nothing
        imgList = sList
        _ImageCount = imgList.Count
    End Sub

    Public Sub LoadDirectory(ByVal sPath As String)
        If sPath.EndsWith("\") Then sPath = sPath.Remove(sPath.Length - 1, 1)
        If di IsNot Nothing AndAlso di.FullName = sPath Then Exit Sub
        indx = -1
        imgList.Clear()
        di = New DirectoryInfo(sPath)
        Dim arrFiles() As FileInfo = di.GetFiles()
        For Each fi As System.IO.FileInfo In arrFiles
            If fi.Extension = ".jpg" Or fi.Extension = ".png" Or fi.Extension = ".bmp" Or fi.Extension = ".gif" Then
                imgList.Add(fi.Name)
            End If
        Next
        _ImageCount = imgList.Count
    End Sub

    Public Sub ReloadImage()
        Dim sName As String = CurrentImage
        CurrentImage = ""
        SetImageWorker(sName, True)
    End Sub

    Public Function SetImage(ByVal sFileName As String, Optional ByVal MaximizeWindow As Boolean = True) As Boolean
        Dim sw As New Stopwatch()
        sw.Start()
        Try
            bSpecialList = False
            If sFileName = "" Then
                imgList = New List(Of String)
                di = Nothing
                DisposeImages()
                Return False
            End If
            If Not System.IO.File.Exists(sFileName) Then
                MsgBox("Error: File does not exist.")
                Return False
            End If
            Dim sPath As String = sFileName.Substring(0, sFileName.LastIndexOf("\")), tPreviousDirectory As String = ""
            If Not di Is Nothing AndAlso di.FullName = sPath Then
                If sFileName = CurrentImage Then Return True
                Dim sName As String = sFileName.Substring(sFileName.LastIndexOf("\") + 1)
                If di.FullName = tPreviousDirectory Then Return False
                If sName = imgList(_indx) Then Return True
                _indx = imgList.IndexOf(sName)
                If _indx = -1 Then PopulateImageList(sFileName)
                tPreviousDirectory = di.FullName
            Else
                If di Is Nothing Then tPreviousDirectory = "" Else tPreviousDirectory = di.FullName
                di = New System.IO.DirectoryInfo(sPath)
                PopulateImageList(sFileName)
            End If
            SetImageWorker(sFileName)
            RaiseEvent ImageChanged(Me, New ImageChangedEventArgs(imgList(_indx), _indx, tPreviousDirectory, True))
            If MaximizeWindow = True Then If Me.WindowState = FormWindowState.Minimized Then Me.WindowState = FormWindowState.Normal
            Me.BringToFront()
            sw.Stop()
            Return True
        Catch ex As Exception
            MsgBox("Error setting image: " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub PopulateImageList(ByVal sFileName As String)
        imgList.Clear()
        Try
            Dim arrFi() As System.IO.FileInfo = di.GetFiles()
            _indx = -1
            For Each fi As System.IO.FileInfo In arrFi
                Dim FiExtension As String = fi.Extension.ToLower()
                If FiExtension = ".jpg" Or FiExtension = ".png" Or FiExtension = ".bmp" Or FiExtension = ".gif" Then
                    imgList.Add(fi.Name)
                End If
            Next
            imgList.Sort(Comparerz)
            _ImageCount = imgList.Count
            If sFileName <> "" Then
                Dim sName As String = sFileName.Substring(sFileName.LastIndexOf("\") + 1)
                _indx = imgList.IndexOf(sName)
                If _indx = -1 Then
                    If _ImageCount > 0 Then _indx = 0
                End If
            End If
        Catch ex As Exception
            MsgBox("Error populating image list: " & ex.Message)
        End Try
    End Sub

    Public Overloads Sub SetImageSpecial(ByVal tIndex As Integer)
        _indx = tIndex
        SetImageWorker(imgList(_indx))
    End Sub

    Public Overloads Sub SetImageSpecial(ByVal tIndex As Integer, ByVal sImage As String)
        _indx = tIndex
        SetImageWorker(sImage)
    End Sub

    Public Sub AnimateImage()
        If Not _Animating Then
            ImageAnimator.Animate(img, AEvent)
            _Animating = True
        End If
    End Sub

    Private Sub OnFrameChanged(ByVal o As Object, ByVal e As EventArgs)
        If _Animating Then Me.Invalidate()
    End Sub

    Private Sub SetImageWorker(ByVal sImg As String, Optional ByVal bRefreshed As Boolean = False)
        If _Animating Then
            _Animating = False
            ImageAnimator.StopAnimate(img, AEvent)
        End If
        Try
            If Not System.IO.File.Exists(sImg) Then
                PopulateImageList("")
                If _indx = -1 Then
                    DisposeImages()
                    Exit Sub
                Else
                    SetImageWorker(imgList(0))
                End If
                Exit Sub
            End If
            If bRefreshed Or img Is Nothing Or CurrentImage <> sImg Then
                Try
                    DisposeImages()
                    Dim sImg2 As String = sImg.ToUpper
                    If sImg2.EndsWith(".GIF") Then
                        img = New Bitmap(sImg)
                    Else
                        Using ThisMemoryStream As New System.IO.MemoryStream(My.Computer.FileSystem.ReadAllBytes(sImg))
                            img = New Bitmap(ThisMemoryStream)
                        End Using
                    End If
                Catch ex As Exception
                    MsgBox("Error loading image: " & ex.Message)
                    Exit Sub
                End Try
                Me.Text = sImg.Substring(sImg.LastIndexOf("\") + 1)
                PreviousImage = CurrentImage
                CurrentImage = sImg
                ImageName = Me.Text.Substring(0, Me.Text.IndexOf("."))
            End If
            Me.Show()
            VScrollMax = 0
            HScrollMax = 0
            MButtonDown = Nothing
            CenterImage()
            ZDifX = CInt(ZIncrement * CenterRect.Width)
            ZDifY = CInt(ZIncrement * CenterRect.Height)
            CZoom = ZoomFixture
            If ZoomFixture <> 1 Then
                Zoom(CZoom, 0, 0)
            Else
                GenerateBGRects()
            End If
            If CurrentImage.EndsWith(".gif") Then
                AnimateImage()
            End If
            Me.Invalidate()
        Catch ex As Exception
            MsgBox("Error in ImageViewer.SetImageWorker: " & ex.Message)
        End Try
    End Sub

    Private Sub NextImage(ByVal blnNext As Boolean)
        If _ImageCount < 2 Then Exit Sub
        Dim ev As New ImageChangingArgs(blnNext, _indx)
        If blnNext = True Then
            _indx += 1
            If _indx = _ImageCount Then
                _indx -= 1
                Dim lolz As New ImageEndReachedArgs("Next")
                RaiseEvent ImageEndReached(Me, lolz)
                If lolz.Handled = True Then Exit Sub Else _indx = 0
            Else
                RaiseEvent ImageChanging(Me, ev)
            End If
        Else
            _indx -= 1
            If _indx = -1 Then
                _indx = 0
                Dim lolz As New ImageEndReachedArgs("Previous")
                RaiseEvent ImageEndReached(Me, lolz)
                If lolz.Handled = True Then Exit Sub Else _indx = _ImageCount - 1
            Else
                RaiseEvent ImageChanging(Me, ev)
            End If
        End If
        If bSpecialList Then
            If Not ev.Handled Then
                SetImageWorker(imgList(_indx))
            Else
                RaiseEvent ImageChanged(Me, New ImageChangedEventArgs(CurrentImage, _indx, "", False))
            End If
        Else
            If Not ev.Handled Then SetImageWorker(di.FullName & "\" & imgList(_indx))
            RaiseEvent ImageChanged(Me, New ImageChangedEventArgs(imgList(_indx), _indx, di.FullName, False))
        End If
    End Sub

    Public Sub SetImageByIndex(ByVal tIndex As Integer)
        Try
            indx = tIndex
            SetImageWorker(di.FullName & "\" & imgList(tIndex))
        Catch ex As Exception
            MsgBox("Error setting image by index: " & ex.Message)
        End Try
    End Sub

    Private Sub CenterImage()
        'Determine limiting factor
        If img.Width < Me.TempRect.Width And img.Height <= Me.TempRect.Height Then 'Just center normally
            CenterRect = New Rectangle(CInt((TempRect.Width - img.Width) / 2), CInt((TempRect.Height - img.Height) / 2), img.Width, img.Height)
        Else
            Dim PWidth As Single = CSng(img.Width / Me.TempRect.Width), PHeight As Single = CSng(img.Height / Me.TempRect.Height)
            Dim PWD As Integer = Me.TempRect.Width - img.Width, PHD As Integer = Me.TempRect.Height - img.Height 'difference between each
            Dim NH As Integer = 0, NW As Integer = 0
            If PWidth > PHeight Then 'Width is limiting factor
                NH = CInt(Me.TempRect.Width / img.Width * img.Height)
                CenterRect = New Rectangle(0, CInt((Me.TempRect.Height - NH) / 2), Me.TempRect.Width, NH)
            Else 'Height is limiting factor
                NW = CInt(Me.TempRect.Height / img.Height * img.Width)
                CenterRect = New Rectangle(CInt((Me.TempRect.Width - NW) / 2), 0, NW, Me.TempRect.Height)
            End If
        End If
        CreateZImage(CenterRect)
        DrawRect = CenterRect
    End Sub

    Private Sub CreateZImage(ByVal rect As Rectangle, Optional ByVal PixelFormat As System.Drawing.Imaging.PixelFormat = System.Drawing.Imaging.PixelFormat.Format32bppPArgb)
        Try
            If Not TImage Is Nothing Then
                TImage.Dispose()
                TImage = Nothing
            End If
            TImage = New Bitmap(rect.Width, rect.Height, PixelFormat) 'Format32bppPArgb is faster than default
            Using gxTemp As Graphics = Graphics.FromImage(TImage)
                gxTemp.SmoothingMode = Drawing2D.SmoothingMode.HighSpeed
                gxTemp.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
                gxTemp.DrawImage(img, 0, 0, rect.Width, rect.Height)
            End Using
            ZImage = TImage
        Catch ex As Exception
            MsgBox("Error resizing image: " & ex.Message)
        End Try
    End Sub

    Public Sub Clear()
        imgList = New List(Of String)
        di = Nothing
        indx = 0
        Me.Visible = False
        DisposeImages()
        Me.Refresh()
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
        If ZImage Is Nothing Then
            e.Graphics.FillRectangle(BGBrush, Me.ClientRectangle)
        Else
            e.Graphics.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighSpeed
            For i As Integer = 0 To BGRects.Count - 1
                e.Graphics.FillRectangle(BGBrush, BGRects(i))
            Next
            If _Animating Then
                ImageAnimator.UpdateFrames()
                e.Graphics.DrawImage(img, DrawRect)
            Else
                e.Graphics.DrawImageUnscaled(ZImage, DrawRect)
            End If
        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub



End Class

Public Class LoadingWindow : Inherits Form
    Public Event Starting(ByVal sender As Object, ByVal e As EventArgs)
    Public Event Finished(ByVal sender As Object, ByVal e As FinishedArgs)
    Public Class FinishedArgs : Inherits System.EventArgs
        Public Closed As Boolean
        Public Sub New(ByVal tClosed As Boolean)
            Closed = tClosed
        End Sub
    End Class
    Public WithEvents PLabel As Label, PBar As ProgressBar, btnCancel As Button
    Public CloseOnCompletion As Boolean = True
    Private tRect As Rectangle, bFinished As Boolean = False
    Private bClosing As Boolean = False

    Private _EnableClose As Boolean = False
    Public Property EnableClose() As Boolean
        Get
            Return _EnableClose
        End Get
        Set(ByVal value As Boolean)
            _EnableClose = value
            If value Then
                btnCancel.Enabled = True
                Me.ControlBox = True
            Else
                btnCancel.Enabled = False
                Me.ControlBox = False
            End If
        End Set
    End Property

    Private _Title As String = "Loading..."
    Public Property Title() As String
        Get
            Return _Title
        End Get
        Set(ByVal value As String)
            _Title = value
            If Me.Text <> "" Then Me.Text = _Title & " " & (Math.Round(PBar.Value / PBar.Maximum, 2) * 100).ToString & "%"
        End Set
    End Property

    Public Sub New(ByVal fWidth As Integer, fHeight As Integer, Optional ByVal x As Integer = -1, Optional ByVal y As Integer = -1, Optional ByVal tEnableClose As Boolean = True)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False : Me.MinimizeBox = False
        Me.MinimumSize = New Size(150, 200)
        PLabel = New Label() : PLabel.Parent = Me : PLabel.AutoSize = False : PLabel.BorderStyle = BorderStyle.Fixed3D
        PBar = New ProgressBar() : PBar.Parent = Me : PBar.Style = ProgressBarStyle.Continuous
        btnCancel = New Button() : btnCancel.Parent = Me : btnCancel.Text = "Cancel"
        tRect = New Rectangle(x, y, fWidth, fHeight)
        EnableClose = tEnableClose
        bClosing = False
    End Sub

    Private Sub LoadingWindow_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not bClosing Then
            If Not _EnableClose Then e.Cancel = True Else CloseWindow()
        End If
    End Sub

    Private Sub LoadingWindow_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If tRect.Left = -1 Then
            SetBounds(CInt((Screen.PrimaryScreen.WorkingArea.Width - tRect.Width) / 2), CInt((Screen.PrimaryScreen.WorkingArea.Height - tRect.Height) / 2), tRect.Width, tRect.Height)
        Else
            SetBounds(tRect.Left, tRect.Top, tRect.Width, tRect.Height)
        End If
        btnCancel.SetBounds(Me.ClientSize.Width - 95, Me.ClientSize.Height - 32, 80, 22)
        PBar.SetBounds(15, btnCancel.Top - 27, Me.ClientSize.Width - 30, 22)
        PLabel.SetBounds(15, 15, PBar.Width, PBar.Top - 20)
        RaiseEvent Starting(Me, New System.EventArgs())
    End Sub

    Private Sub UpdateStuff()
        Me.Text = _Title & " " & (Math.Round(PBar.Value / PBar.Maximum, 2) * 100).ToString & "%"
        If PBar.Value = PBar.Maximum Then
            bFinished = True
            If CloseOnCompletion Then CloseWindow() Else btnCancel.Text = "Close"
        End If
    End Sub

    Public Sub Increment()
        PBar.Value = Math.Min(PBar.Maximum, PBar.Value + PBar.Step)
        UpdateStuff()
    End Sub

    Public Sub UpdateValue(ByVal tValue As Integer)
        PBar.Value += tValue
        UpdateStuff()
    End Sub

    Public Sub SetParams(ByVal tMin As Integer, ByVal tMax As Integer, Optional ByVal tValue As Integer = 0, Optional ByVal tStep As Integer = 1)
        PBar.Minimum = tMin : PBar.Maximum = tMax : PBar.Value = tValue : PBar.Step = tStep
    End Sub

    Public Sub ShowWindow(ByVal tMin As Integer, ByVal tMax As Integer, Optional ByVal tValue As Integer = 0, Optional ByVal tStep As Integer = 1)
        PBar.Minimum = tMin : PBar.Maximum = tMax : PBar.Value = tValue : PBar.Step = tStep
        Me.Text = _Title & " " & (Math.Round(PBar.Value / PBar.Maximum, 2) * 100).ToString & "%"
        Me.ShowDialog()
    End Sub

    Public Sub CloseWindow()
        RaiseEvent Finished(Me, New FinishedArgs(Not bFinished))
        bClosing = True
        Me.Dispose()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        CloseWindow()
    End Sub
End Class

Public Class TransparentPanel
    Inherits Panel
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Gets the creation parameters.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim createparams__1 As CreateParams = MyBase.CreateParams
            createparams__1.ExStyle = createparams__1.ExStyle Or &H20
            ' WS_EX_TRANSPARENT
            Return createparams__1
        End Get
    End Property

    ''' <summary>
    ''' Skips painting the background.
    ''' </summary>
    ''' <param name="e">E.</param>
    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        'do nothing
    End Sub
End Class

#Region "Control Extensions"
Public Class LabelX : Inherits Label
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal tAlign As ContentAlignment = ContentAlignment.TopLeft)
        Me.Parent = prnt
        Me.TextAlign = tAlign
        If sName <> "" Then Name = sName
        If w <> -1 Then
            SetBounds(x, y, w, h)
        Else
            Left = x
            Top = y
            AutoSize = True
        End If
        Text = sText
    End Sub
End Class

Public Class DecimalTextBox : Inherits TextBox
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
        Text = sText
        ShortcutsEnabled = False
    End Sub

    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
        Dim lol As Integer = e.KeyValue
        If (lol > 47 AndAlso lol < 58) OrElse lol = 190 OrElse lol = 8 OrElse lol = 46 OrElse (lol > 36 AndAlso lol < 41) Then
            'MsgBox("Success!")
        Else
            'MsgBox("Fail! " & lol.ToString())
            e.SuppressKeyPress = True
            e.Handled = True
        End If
    End Sub

    Public Function GetSingle() As Single
        Try
            Return CSng(Me.Text)
        Catch ex As Exception
            Return 0
        End Try
    End Function
End Class

Public Class DataGridViewX : Inherits System.Windows.Forms.DataGridView
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, keyData As System.Windows.Forms.Keys) As Boolean
        If (keyData < 48 Or keyData > 57) AndAlso (keyData < 37 Or keyData > 40) AndAlso keyData <> 13 AndAlso keyData <> 190 AndAlso keyData <> 8 AndAlso keyData <> 46 AndAlso keyData <> 65552 AndAlso keyData <> 65589 Then
            Return True
        Else
            Return MyBase.ProcessCmdKey(msg, keyData)
        End If
    End Function
End Class

Public Class RichTextBoxX : Inherits RichTextBox
    '#Region "Extensions"
    '    Private Overloads Declare Function SendMessage Lib "user32" (ByVal hWnd As HandleRef, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    '    Private Overloads Declare Function SendMessage Lib "user32" (ByVal hWnd As HandleRef, ByVal msg As Integer, ByVal wParam As Integer, ByRef lp As PARAFORMAT) As Integer

    '    Public Sub BeginUpdate()
    '        ' Deal with nested calls.
    '        updating = (updating + 1)
    '        If (updating > 1) Then
    '            Return
    '        End If
    '        ' Prevent the control from raising any events.
    '        oldEventMask = SendMessage(New HandleRef(Me, Handle), EM_SETEVENTMASK, 0, 0)
    '        ' Prevent the control from redrawing itself.
    '        SendMessage(New HandleRef(Me, Handle), WM_SETREDRAW, 0, 0)
    '    End Sub

    '    ''' <summary>
    '    ''' Resumes drawing and event handling.
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' This method should be called every time a call is made
    '    ''' made to BeginUpdate. It resets the event mask to it's
    '    ''' original value and enables redrawing of the control.
    '    ''' </remarks>
    '    Public Sub EndUpdate()
    '        ' Deal with nested calls.
    '        updating = (updating + 1)
    '        If (updating > 0) Then
    '            Return
    '        End If
    '        ' Allow the control to redraw itself.
    '        SendMessage(New HandleRef(Me, Handle), WM_SETREDRAW, 1, 0)
    '        ' Allow the control to raise event messages.
    '        SendMessage(New HandleRef(Me, Handle), EM_SETEVENTMASK, 0, oldEventMask)
    '    End Sub

    '    ''' <summary>
    '    ''' Gets or sets the alignment to apply to the current
    '    ''' selection or insertion point.
    '    ''' </summary>
    '    ''' <remarks>
    '    ''' Replaces the SelectionAlignment from
    '    ''' <see cref="RichTextBox"/>.
    '    ''' </remarks>
    '    Public Shadows Property SelectionAlignment As TextAlign
    '        Get
    '            Dim fmt As PARAFORMAT = New PARAFORMAT
    '            fmt.cbSize = Marshal.SizeOf(fmt)
    '            ' Get the alignment.
    '            SendMessage(New HandleRef(Me, Handle), EM_GETPARAFORMAT, SCF_SELECTION, fmt)
    '            ' Default to Left align.
    '            If ((fmt.dwMask And PFM_ALIGNMENT) _
    '                        = 0) Then
    '                Return TextAlign.Left
    '            End If
    '            Return CType(fmt.wAlignment, TextAlign)
    '        End Get
    '        Set(value As TextAlign)
    '            Dim fmt As PARAFORMAT = New PARAFORMAT
    '            fmt.cbSize = Marshal.SizeOf(fmt)
    '            fmt.dwMask = PFM_ALIGNMENT
    '            fmt.wAlignment = CType(value, Short)
    '            ' Set the alignment.
    '            SendMessage(New HandleRef(Me, Handle), EM_SETPARAFORMAT, SCF_SELECTION, fmt)
    '        End Set
    '    End Property

    '    ''' <summary>
    '    ''' This member overrides
    '    ''' <see cref="Control"/>.OnHandleCreated.
    '    ''' </summary>
    '    Protected Overrides Sub OnHandleCreated(ByVal e As EventArgs)
    '        MyBase.OnHandleCreated(e)
    '        ' Enable support for justification.
    '        SendMessage(New HandleRef(Me, Handle), EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY)
    '    End Sub

    '    Private updating As Integer = 0

    '    Private oldEventMask As Integer = 0

    '    ' Constants from the Platform SDK.
    '    Private Const EM_SETEVENTMASK As Integer = 1073

    '    Private Const EM_GETPARAFORMAT As Integer = 1085

    '    Private Const EM_SETPARAFORMAT As Integer = 1095

    '    Private Const EM_SETTYPOGRAPHYOPTIONS As Integer = 1226

    '    Private Const WM_SETREDRAW As Integer = 11

    '    Private Const TO_ADVANCEDTYPOGRAPHY As Integer = 1

    '    Private Const PFM_ALIGNMENT As Integer = 8

    '    Private Const SCF_SELECTION As Integer = 1

    '    ' It makes no difference if we use PARAFORMAT or
    '    ' PARAFORMAT2 here, so I have opted for PARAFORMAT2.
    '    <StructLayout(LayoutKind.Sequential)> _
    '    Private Structure PARAFORMAT

    '        Public cbSize As Integer

    '        Public dwMask As UInteger

    '        Public wNumbering As Short

    '        Public wReserved As Short

    '        Public dxStartIndent As Integer

    '        Public dxRightIndent As Integer

    '        Public dxOffset As Integer

    '        Public wAlignment As Short

    '        Public cTabCount As Short

    '        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> _
    '        Public rgxTabs() As Integer

    '        ' PARAFORMAT2 from here onwards.
    '        Public dySpaceBefore As Integer

    '        Public dySpaceAfter As Integer

    '        Public dyLineSpacing As Integer

    '        Public sStyle As Short

    '        Public bLineSpacingRule As Byte

    '        Public bOutlineLevel As Byte

    '        Public wShadingWeight As Short

    '        Public wShadingStyle As Short

    '        Public wNumberingStart As Short

    '        Public wNumberingStyle As Short

    '        Public wNumberingTab As Short

    '        Public wBorderSpace As Short

    '        Public wBorderWidth As Short

    '        Public wBorders As Short
    '    End Structure

    '    Public Enum TextAlign

    '        Left = 1

    '        Right = 2

    '        Center = 3

    '        Justify = 4
    '    End Enum
    '#End Region
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
        Text = sText
    End Sub

    Private Sub RichTextBoxX_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyValue = 220 Then e.SuppressKeyPress = True
    End Sub
End Class

Public Class TextBoxX : Inherits TextBox
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
        Text = sText
        Parent = prnt
    End Sub

    Private Sub TextBoxX_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyValue = 220 Then e.SuppressKeyPress = True
    End Sub
End Class

Public Class GXBTextbox : Inherits TextBoxX
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        MyBase.New(prnt, sName, sText, x, y, w, h)
    End Sub

    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
        If (e.KeyValue < 48 Or e.KeyValue > 57) AndAlso e.KeyValue <> 8 AndAlso e.KeyValue <> 46 AndAlso e.KeyValue <> 39 AndAlso e.KeyValue <> 37 Then
            If Not Text.Contains("x") Then
                If Text <> "" Then
                    Dim into As Integer = SelectionStart
                    Text = Text.Insert(into, "x")
                    SelectionStart = into + 1
                End If
            End If
            e.SuppressKeyPress = True
            e.Handled = True
        End If
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnLostFocus(e As System.EventArgs)
        If Text = "" Then Exit Sub 'MsgBox("Format correctly!")
        Dim indx As Integer = Text.IndexOf("x")
        If indx = -1 Or indx = 0 Then
            Text = "" : Exit Sub
        End If
        If indx = Text.Length - 1 Then
            Text &= "0" : Exit Sub
        End If
        MyBase.OnLostFocus(e)
    End Sub

    Public Function GetLeft() As Byte
        Try
            Return CByte(Text.Substring(0, Text.IndexOf("x")))
        Catch
            Return CByte(1)
        End Try
    End Function

    Public Function GetRight() As Byte
        Try
            Return CByte(Text.Substring(Text.IndexOf("x") + 1))
        Catch
            Return CByte(0)
        End Try
    End Function
End Class

Public Class ButtonX : Inherits System.Windows.Forms.Button
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        If w <> -1 Then
            SetBounds(x, y, w, h)
        Else
            Left = x
            Top = y
            AutoSize = True
        End If
        TextAlign = ContentAlignment.MiddleCenter
        Text = sText
    End Sub
End Class

Public Class ComboBoxX : Inherits System.Windows.Forms.ComboBox
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
        DropDownStyle = ComboBoxStyle.DropDownList
    End Sub
End Class

Public Class ListViewX : Inherits System.Windows.Forms.ListView
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)

    End Sub
End Class

Public Class ListBoxX : Inherits System.Windows.Forms.ListBox
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
    End Sub
End Class

Public Class TreeViewX : Inherits System.Windows.Forms.TreeView
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
        Indent = 5
    End Sub
End Class

Public Class NumericUpDownX : Inherits System.Windows.Forms.NumericUpDown
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal min As Decimal = CDec(-1.77), Optional ByVal max As Decimal = CDec(-1.77))
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
        Me.ReadOnly = True
        If min <> CDec(-1.77) Then Minimum = min
        If max <> CDec(-1.77) Then Maximum = max
    End Sub
End Class

Public Class CheckBoxX : Inherits System.Windows.Forms.CheckBox
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal bChecked As Boolean = False)
        Parent = prnt
        If sName <> "" Then Name = sName
        If w <> -1 Then
            SetBounds(x, y, w, h)
        Else
            Left = x
            Top = y
            AutoSize = True
        End If
        Text = sText
        Checked = bChecked
    End Sub
End Class

Public Class RadioButtonX : Inherits System.Windows.Forms.RadioButton
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal sText As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal bChecked As Boolean = False)
        Parent = prnt
        If sName <> "" Then Name = sName
        If w <> -1 Then
            SetBounds(x, y, w, h)
        Else
            Left = x
            Top = y
            AutoSize = True
        End If
        Text = sText
        Checked = bChecked
    End Sub
End Class

Public Class PanelX : Inherits System.Windows.Forms.Panel
    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        Parent = prnt
        If sName <> "" Then Name = sName
        SetBounds(x, y, w, h)
    End Sub
End Class
#End Region

Public Module NewMsgBox
    Public Function NewMsg(ByVal tObject As Object, Optional ByVal tTitle As String = "", Optional ByVal tButtons As List(Of String) = Nothing, Optional ByVal tIcon As MessageBoxIcon = MessageBoxIcon.None, Optional ByVal msgParams As MessageBoxXParams = Nothing) As String
        If msgParams Is Nothing Then
            msgParams = New MessageBoxXParams(Nothing, 1, 0)
        End If
        If tButtons Is Nothing Then msgParams.Buttons = New List(Of String) From {"OK"} Else msgParams.Buttons = tButtons
        msgParams.tIcon = tIcon
        Dim mlolz As New MessageBoxX(tObject, msgParams)
        mlolz.Text = tTitle
        mlolz.ShowDialog()
        Return mlolz.Result
    End Function

    Public Class MessageBoxXParams
        Public Columns As Integer, Bounds As Rectangle, CenterOnForm As Form, Buttons As List(Of String), LineCountPerColumn As Integer
        Public tIcon As MessageBoxIcon, ColumnHeader As String

        Public Sub New(ByVal tCenterOnForm As Form, ByVal tColumns As Integer, ByVal tLineCountPerColumn As Integer, Optional ByVal tColumnHeader As String = "")
            CenterOnForm = tCenterOnForm
            Columns = tColumns
            LineCountPerColumn = tLineCountPerColumn
            ColumnHeader = tColumnHeader
        End Sub
    End Class

    Private Class MessageBoxX : Inherits Form
        Public Buttons As New List(Of Button), PIcon As PictureBox, Labels As New List(Of Label), panButtons As Panel, panText As Panel
        Public Result As String = ""
        Private msgBounds As Rectangle, obj As Object, msgParams As MessageBoxXParams

        Public Sub New(ByVal tObject As Object, ByVal tMsgParams As MessageBoxXParams)
            Me.ShowInTaskbar = False
            Me.MinimumSize = New Size(155, 140)
            Me.MaximumSize = New Size(600, Screen.PrimaryScreen.WorkingArea.Height)
            Me.MaximizeBox = False : Me.MinimizeBox = False
            Me.Font = SystemFonts.MessageBoxFont
            Me.BackColor = Color.White
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
            obj = tObject
            msgParams = tMsgParams
            If msgParams.Buttons Is Nothing Then
                msgParams.Buttons = New List(Of String) From {"OK"}
            End If
        End Sub

        Private Sub MessageBoxX_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            If Result = "" Then
                If Buttons.Count = 1 AndAlso Buttons(0).Text <> "Cancel" Then Result = "Ok" Else Result = "Cancel"
            End If
        End Sub

        Private Sub btnClick(ByVal sender As Object, ByVal e As EventArgs)
            Result = DirectCast(sender, Button).Text
            Me.Close()
        End Sub

        Protected Overrides Sub OnLoad(e As System.EventArgs)
            msgBounds = Rectangle.Empty
            Dim btnWidth As Integer = 0, lolz As IList = Nothing
            Dim MaxTextWidth As Integer = Me.MaximumSize.Width - -24 - 16, MaxHeights As Integer = Me.MaximumSize.Height - 40
            Dim bmp As Bitmap = Nothing
            If msgParams.tIcon <> MessageBoxIcon.None Then
                If msgParams.tIcon = MessageBoxIcon.Asterisk Then
                    bmp = SystemIcons.Asterisk.ToBitmap
                ElseIf msgParams.tIcon = MessageBoxIcon.Error Then
                    bmp = SystemIcons.Error.ToBitmap
                ElseIf msgParams.tIcon = MessageBoxIcon.Exclamation Then
                    bmp = SystemIcons.Exclamation.ToBitmap
                ElseIf msgParams.tIcon = MessageBoxIcon.Hand Then
                    bmp = SystemIcons.Hand.ToBitmap
                ElseIf msgParams.tIcon = MessageBoxIcon.Error Then
                    bmp = SystemIcons.Information.ToBitmap
                ElseIf msgParams.tIcon = MessageBoxIcon.Question Then
                    bmp = SystemIcons.Question.ToBitmap
                ElseIf msgParams.tIcon = MessageBoxIcon.Stop Then
                    bmp = SystemIcons.Exclamation.ToBitmap
                ElseIf msgParams.tIcon = MessageBoxIcon.Warning Then
                    bmp = SystemIcons.Warning.ToBitmap
                End If
                PIcon = New PictureBox() : PIcon.Parent = Me : PIcon.Size = New Size(bmp.Width, bmp.Height)
                PIcon.Image = bmp
            End If
            For i As Integer = 0 To msgParams.Buttons.Count - 1
                Dim btn As New Button()
                btn.AutoSize = True : btn.Text = msgParams.Buttons(i)
                Dim sSizeos As Integer = btn.Width : btn.AutoSize = False
                btn.Size = New Size(Math.Min(86, sSizeos + 5), 24)
                btnWidth += btn.Width + 10
                AddHandler btn.Click, AddressOf btnClick
                Buttons.Add(btn)
            Next
            btnWidth = Math.Min(btnWidth + 50, Me.MaximumSize.Width)
            Dim objType As Type = obj.GetType()

            If objType.Name.Contains("List") Then
                Dim genericType = GetType(List(Of )).MakeGenericType(objType)
                lolz = DirectCast(Activator.CreateInstance(genericType), IList)
                lolz = DirectCast(obj, IList)
                msgParams.Columns = Math.Min(msgParams.Columns, lolz.Count)
            End If
            If msgParams.Columns > 1 Then
                Dim iColCount As Integer = msgParams.Columns
                iColCount = Math.Min(iColCount, CInt(Math.Ceiling(lolz.Count / msgParams.LineCountPerColumn)))
                Dim ColumnWidths As Integer = CInt(MaxTextWidth / iColCount)
                Dim ColumnText(iColCount - 1) As System.Text.StringBuilder, CIndex As Integer = 0
                Dim iCount As Integer = 0, iLines As Integer = msgParams.LineCountPerColumn - 1
                Dim ColumnTexts(iColCount - 1) As String
                For i As Integer = 0 To iColCount - 1
                    ColumnText(i) = New System.Text.StringBuilder()
                Next
                Dim MaxCount As Integer = Math.Min((msgParams.LineCountPerColumn * iColCount), lolz.Count)
                For i As Integer = 0 To MaxCount - 1 ' lolz.Count - 1
                    Dim str As String = lolz(i).ToString()
                    Dim ilength As Integer = str.Length
                    If ilength > MaxTextWidth Then
                        Dim tLines As Integer = CInt(Math.Ceiling(ilength / MaxTextWidth))
                        For i2 As Integer = 0 To tLines - 2
                            ColumnText(CIndex).Append(str.Substring((i2 * MaxTextWidth), MaxTextWidth) & Chr(10))
                            If iCount = iLines Then iCount = 0 : CIndex += 1 Else iCount += 1
                        Next
                        ColumnText(CIndex).Append(str.Substring(((tLines - 1) * MaxTextWidth)) & Chr(10))
                        If iCount = iLines Then iCount = 0 : CIndex += 1 Else iCount += 1
                    Else
                        ColumnText(CIndex).Append(str & Chr(10))
                        If iCount = iLines Then iCount = 0 : CIndex += 1 Else iCount += 1
                    End If

                Next
                Dim iTotalWidth As Integer = 0, HighestHeight As Integer = 0, HighestWidth As Integer = 0
                For i As Integer = 0 To iColCount - 1
                    ColumnTexts(i) = ColumnText(i).ToString()
                    Dim ssizeo As SizeF = DirectCast(MeasureAString(2, ColumnTexts(i), ColumnWidths, MaxHeights - 49), SizeF)
                    iTotalWidth += CInt(ssizeo.Width)
                    HighestHeight = Math.Max(CInt(ssizeo.Height), HighestHeight)
                    HighestWidth = Math.Max(CInt(ssizeo.Width), HighestWidth)
                Next
                iTotalWidth = ((HighestWidth + 10) * iColCount)
                If iTotalWidth < btnWidth Then
                    iTotalWidth = btnWidth ': HighestHeight = 0
                End If
                ' iTotalWidth = Math.Max(iTotalWidth, btnWidth)
                If iTotalWidth < Me.MinimumSize.Width - 16 - 24 Then
                    iTotalWidth = Me.MinimumSize.Width - 16 - 24
                    ColumnWidths = CInt(iTotalWidth / iColCount)
                End If
                Dim iLefts As Integer = 0
                'If iTotalWidth > MaxTextWidth OrElse HighestHeight > MaxHeights Then
                '    'omg problem
                'Else
                msgBounds.Width = iTotalWidth + 16 + 24
                If Not PIcon Is Nothing Then
                    PIcon.Location = New Point(12, CInt(((HighestHeight + 31) - bmp.Height) / 2))
                    msgBounds.Width += PIcon.Width + 10
                    iLefts = PIcon.Right + 10
                Else
                    iLefts = 12
                End If
                msgBounds.Height = HighestHeight + 40 + 49 + 58
                msgBounds.X = CInt((Screen.PrimaryScreen.WorkingArea.Width - msgBounds.Width) / 2)
                msgBounds.Y = CInt((Screen.PrimaryScreen.WorkingArea.Height - msgBounds.Height) / 2)
                ' End If
                Me.SetBounds(msgBounds.Left, msgBounds.Top, msgBounds.Width, msgBounds.Height)
                ColumnWidths = CInt(iTotalWidth / iColCount)
                For i As Integer = 0 To iColCount - 1
                    Dim lbl As New Label() : lbl.Parent = Me : lbl.AutoSize = False ': lbl.BorderStyle = BorderStyle.FixedSingle
                    lbl.SetBounds(iLefts, 29, ColumnWidths, HighestHeight + 2)
                    lbl.Text = ColumnTexts(i)
                    iLefts = lbl.Right
                    Labels.Add(lbl)
                Next
            Else
                Dim lolz2 As String = ""
                If lolz Is Nothing Then
                    lolz2 = obj.ToString()
                Else
                    Dim sb As New System.Text.StringBuilder()
                    For Each obj As Object In lolz
                        sb.Append(obj.ToString() & Chr(10))
                    Next
                    lolz2 = sb.ToString()
                End If
                Dim sizeo As SizeF = DirectCast(MeasureAString(2, lolz2, MaxTextWidth, MaxHeights), SizeF)
                msgBounds.Height = CInt(Math.Min(Math.Max(Me.MinimumSize.Height, sizeo.Height + 107 + 40), Me.MaximumSize.Height))
                If PIcon Is Nothing Then
                    msgBounds.Width = Math.Max(CInt(sizeo.Width) + 24 + 16, btnWidth)
                Else
                    PIcon.Location = New Point(12, CInt((sizeo.Height + 29) / 2))
                    msgBounds.Width = Math.Max(CInt(sizeo.Width) + PIcon.Width + 10 + 24 + 16, btnWidth)
                End If
                msgBounds.Width = Math.Min(Math.Max(Me.MinimumSize.Width, msgBounds.Width), Me.MaximumSize.Width)
                msgBounds.X = CInt((Screen.PrimaryScreen.WorkingArea.Width - msgBounds.Width) / 2)
                msgBounds.Y = CInt((Screen.PrimaryScreen.WorkingArea.Height - msgBounds.Height) / 2)
                Me.SetBounds(msgBounds.Left, msgBounds.Top, msgBounds.Width, msgBounds.Height)
                Dim lbl As New Label() : lbl.Parent = Me : lbl.AutoSize = False ': lbl.BackColor = Color.DarkRed
                If PIcon Is Nothing Then
                    lbl.SetBounds(12, 29, Me.ClientSize.Width - 24, Me.ClientSize.Height - 49 - 58)
                Else
                    lbl.SetBounds(PIcon.Right + 10, 29, Me.ClientSize.Width - 24, Me.ClientSize.Height - 49 - 58)
                End If
                'lbl.Font = SystemFonts.MessageBoxFont
                lbl.Text = lolz2
            End If
            panButtons = New Panel : panButtons.Parent = Me : panButtons.SetBounds(0, Me.ClientSize.Height - 49, Me.ClientSize.Width, 49)
            panButtons.BackColor = Color.FromKnownColor(KnownColor.Control)
            Dim iLeft As Integer = panButtons.Width - Buttons(Buttons.Count - 1).Width - 10
            For i As Integer = Buttons.Count - 1 To 0 Step -1
                Buttons(i).Parent = panButtons
                Buttons(i).SetBounds(iLeft, 14, Buttons(i).Width, Buttons(i).Height)
                iLeft = iLeft - Buttons(Buttons.Count - 1).Width - 10
            Next
            If msgParams.ColumnHeader <> "" Then
                Dim lblHeader As New Label() : lblHeader.AutoSize = False : lblHeader.Parent = Me
                lblHeader.SetBounds(0, 0, Me.ClientSize.Width, 29) : lblHeader.TextAlign = ContentAlignment.MiddleCenter
                lblHeader.Text = msgParams.ColumnHeader
            End If
            MyBase.OnLoad(e)
        End Sub

        Public Function MeasureAString(ByVal Width0Height1Both2 As Byte, ByVal str As String, ByVal w As Integer, ByVal h As Integer) As Object
            Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
                Dim SFormat As New System.Drawing.StringFormat
                Dim rect As New System.Drawing.RectangleF(0, 0, w, h)
                Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
                SFormat.SetMeasurableCharacterRanges(range)
                Dim regions() As Region = gx.MeasureCharacterRanges(str, Font, rect, SFormat)
                rect = regions(0).GetBounds(gx)
                If Width0Height1Both2 = 0 Then Return rect.Right + 1 'gx.MeasureString(str, Font, 50000000, lolz).Width
                If Width0Height1Both2 = 1 Then Return rect.Bottom + 1 'gx.MeasureString(str, Font, 50000000, lolz).Height
                If Width0Height1Both2 = 2 Then Return New SizeF(rect.Right + 1, rect.Bottom + 1) 'gx.MeasureString(str, Font, 50000000, lolz)
            End Using
            Return -1
        End Function

    End Class

End Module

Public Module NewInputBox
    Private TextLimits As Integer = 1, PB As PromptBox, txtTexto As TextBoxX, bShift As Boolean, bControl As Boolean, bAlt As Boolean, CurrentLetter As String

    Public Class TextStuff
        Public bShift As Boolean, bControl As Boolean, bAlt As Boolean, CurrentLetter As String
        Public Sub New()

        End Sub
    End Class

    Public Function Show(ByVal Message As String, Optional ByVal Title As String = "", Optional ByVal InitialText As String = "", Optional ByVal TextLimit As Integer = 1) As InputBoxResult
        PB = New PromptBox(Title)
        Dim lblTexto As New LabelX(PB, "", Message, 15, 15, -1, -1, ContentAlignment.MiddleCenter)
        txtTexto = New TextBoxX(PB, "", InitialText, lblTexto.Right + 5, lblTexto.Top, 200, 22)
        txtTexto.ReadOnly = True
        AddHandler txtTexto.KeyDown, AddressOf PB_KeyDown
        TextLimits = TextLimit
        Dim Cancel As Boolean = (PB.ShowPromptBox = 0)
        If Cancel Then
            Show = New InputBoxResult("", "", True)
        Else
            Dim Modifiers As String = ""
            If bShift Then Modifiers = "S"
            If bControl Then Modifiers &= "C"
            If bAlt Then Modifiers &= "A"
            Show = New InputBoxResult(CurrentLetter, Modifiers, False)
        End If
    End Function

    Public Function GetLetter(ByVal KeyCode As Integer) As String
        If KeyCode > 64 AndAlso KeyCode < 91 Then
            Dim pChar As Char = Chr(KeyCode)
            If pChar <> Nothing Then Return pChar Else Return ""
        ElseIf KeyCode > 47 AndAlso KeyCode < 58 Then
            Dim pChar As Char = Chr(KeyCode)
            If pChar <> Nothing Then Return pChar Else Return ""
        ElseIf KeyCode = Keys.OemMinus Then
            Return "-"
        ElseIf KeyCode = Keys.Oemplus Then
            Return "+"
        ElseIf KeyCode = Keys.Oemtilde Then
            Return "~"
        ElseIf KeyCode = Keys.Oem4 Then
            Return "["
        ElseIf KeyCode = Keys.Oem6 Then
            Return "]"
        ElseIf KeyCode = Keys.Oem2 Then
            Return "/"
        ElseIf KeyCode = Keys.Oemcomma Then
            Return ","
        ElseIf KeyCode = Keys.OemPeriod Then
            Return "."
        ElseIf KeyCode = Keys.Oem7 Then
            Return "'"
        ElseIf KeyCode = Keys.Oem1 Then
            Return ";"
        Else
            Return ""
        End If
    End Function

    Private Sub PB_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If TextLimits = 1 Then
            bShift = True = e.Shift
            bAlt = e.Alt
            bControl = e.Control
            CurrentLetter = ""
            Dim NewText As String = ""
            If e.Shift Then ' Shift Key
                NewText = "Shift+"
            End If
            If e.Control Then ' Control Key
                NewText &= "Control+"
            End If
            If e.Alt Then ' Alt Key
                NewText &= "Alt+"
            End If
            CurrentLetter = GetLetter(e.KeyValue)
            If CurrentLetter <> "" Then NewText &= CurrentLetter
            txtTexto.Text = NewText
        End If
    End Sub

    Public Class InputBoxResult
        Public Text As String, Modifiers As String, Cancel As Boolean = False
        Public Sub New(ByVal tText As String, ByVal tModifiers As String, ByVal bCancel As Boolean)
            Text = tText : Modifiers = tModifiers : Cancel = bCancel
        End Sub
        Public Sub New()

        End Sub
    End Class
End Module



Public Class KeyCodeToAscii
    <DllImport("User32.dll")> _
    Public Shared Function ToAscii(ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByVal lpbKeyState As Byte(), ByVal lpChar As Byte(), ByVal uFlags As Integer) As Integer
    End Function

    <DllImport("User32.dll")> _
    Public Shared Function GetKeyboardState(ByVal pbKeyState As Byte()) As Integer
    End Function

    Public Shared Function GetAsciiCharacter(ByVal uVirtKey As Integer) As Char
        Dim lpKeyState As Byte() = New Byte(255) {}
        GetKeyboardState(lpKeyState)
        Dim lpChar As Byte() = New Byte(1) {}
        If ToAscii(uVirtKey, 0, lpKeyState, lpChar, 0) = 1 Then
            Return Convert.ToChar((lpChar(0)))
        Else
            Return New Char()
        End If
    End Function
End Class

Public Class MyComparer
    Implements IComparer(Of String)

    Declare Unicode Function StrCmpLogicalW Lib "shlwapi.dll" ( _
        ByVal s1 As String, _
        ByVal s2 As String) As Int32

    Public Function Compare(ByVal x As String, ByVal y As String) As Integer _
        Implements System.Collections.Generic.IComparer(Of String).Compare
        Return StrCmpLogicalW(x, y)
    End Function
End Class
Option Strict On

Imports System
Imports System.Threading
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Linq.Expressions
Imports MyControls

Public Class Form1 : Inherits Form
    Public Shared w As Integer = 900, h As Integer = 500
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            cp.Style = cp.Style Or &H2000000 And Not 33554432
            Return cp
        End Get
    End Property

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.ResizeRedraw, True)
        UpdateStyles()
        Me.SetBounds(CInt((Screen.PrimaryScreen.WorkingArea.Width - w) / 2), CInt((Screen.PrimaryScreen.WorkingArea.Height - h) / 2), w, h)
        Me.MinimumSize = New Size(600, 430)
        Me.MaximumSize = New Size(Screen.PrimaryScreen.WorkingArea.Width, h)
    End Sub
End Class

Public Class AddSSForm : Inherits Form
    Public WithEvents txtName, txtAmount As TextBoxX, btnCreate, btnCancel As ButtonX
    Private bCreate As Boolean = False

    Public Sub New()
        FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False
        Dim lblTemp As New LabelX(Me, "", "Serving Name:", 10, 10, 90, 20)
        lblTemp.TextAlign = ContentAlignment.MiddleRight
        txtName = New TextBoxX(Me, "txtName", "", lblTemp.Right + 2, lblTemp.Top, 150, 22)
        lblTemp = New LabelX(Me, "", "Serving Amount:", 10, lblTemp.Bottom + 10, 90, 20)
        lblTemp.TextAlign = ContentAlignment.MiddleRight
        txtAmount = New TextBoxX(Me, "txtAmount", "", lblTemp.Right + 2, lblTemp.Top, 150, 22)
        btnCreate = New ButtonX(Me, "btnCreate", "Create", lblTemp.Left, lblTemp.Bottom + 10, 118, 22)
        btnCancel = New ButtonX(Me, "btnCancel", "Cancel", btnCreate.Right + 5, btnCreate.Top, 118, 22)
        AcceptButton = btnCreate
        CancelButton = btnCancel
    End Sub

    Private Sub AddSSForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim w As Integer = 270, h As Integer = 135
        Me.SetBounds(CInt(frmMain.Left + (frmMain.Width - w) / 2), CInt(frmMain.Top + (frmMain.Height - h) / 2), w, h)
    End Sub

    Public Function GetServingSize(ByVal lst As List(Of ServingSize)) As ServingSize
        txtName.Text = ""
        txtAmount.Text = ""
        ShowDialog()
        If bCreate AndAlso txtName.Text <> "" AndAlso txtAmount.Text <> "" Then
            If Not IsNumeric(txtAmount.Text) Then
                MsgBox("Error: " & txtAmount.Text & " is not a numeric amount.")
                Return Nothing
            Else
                For i As Integer = 0 To lst.Count - 1
                    If lst(i).Name = txtName.Text Then
                        MsgBox("Error: Serving Size with the same name already exists.")
                        Return Nothing
                    End If
                Next
                Return New ServingSize(txtName.Text, CSng(txtAmount.Text))
            End If
        Else
            Return Nothing
        End If
    End Function

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As System.EventArgs) Handles btnCreate.Click
        bCreate = True
        Me.Close()
    End Sub

    Private Sub AddSSForm_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        txtName.Focus()
    End Sub
End Class

Public Class DataField : Inherits PanelX
    Public WithEvents DF As DataGridViewX, cmbSite As ComboBoxX, CM As ContextMenuStrip, btnAddManual, btnExpand As ButtonX
    Public CFood As FoodItem, lstSites As New List(Of SiteProfile), CSite As SiteProfile
    Public Rows As New Dictionary(Of Integer, Integer), Values As New Dictionary(Of Integer, Single)
    Public Changed As Boolean = False

    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        MyBase.New(prnt, sName, x, y, w, h)
        cmbSite = New ComboBoxX(Me, "cmbSite", 0, 0, 150, 22)
        btnAddManual = New ButtonX(Me, "btnAddManual", "Add Manual", cmbSite.Right + 10, cmbSite.Top, 120, 22)
        btnExpand = New ButtonX(Me, "btnExpand", "Expand All", btnAddManual.Right + 10, btnAddManual.Top, btnAddManual.Width, btnAddManual.Height)
        DF = New DataGridViewX(Me, "DF", 0, cmbSite.Bottom + 5, w, h - cmbSite.Height - 5)
        DF.AllowUserToAddRows = False
        DF.AllowUserToDeleteRows = False
        DF.AllowUserToOrderColumns = False
        DF.MultiSelect = False
        DF.RowHeadersWidth = 50
        DF.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DF.RowHeadersDefaultCellStyle.Font = New Font(DF.Font.FontFamily, 10)
        Me.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        DF.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        CM = New ContextMenuStrip()
    End Sub

    Public Sub LoadFoodItem(ByVal tFood As FoodItem)
        CSite = Nothing
        CFood = tFood
        DF.Columns.Clear()
        DF.Columns.Add("", "Nutrient")
        cmbSite.Items.Clear() : lstSites.Clear()
        For i As Integer = 0 To CFood.DataSites.Count - 1
            Dim st As SiteProfile = CFood.DataSites.Values(i)
            If st.Site.ID = 2 Then cmbSite.Items.Add("Manual") Else cmbSite.Items.Add(st.Site.Name)
            lstSites.Add(st)
        Next
        DF.Columns.Add("", CFood.ServingSizes(0).Name)
        For i As Integer = 1 To CFood.ServingSizes.Count - 1
            DF.Columns.Add("", CFood.ServingSizes(i).Name & " (" & CFood.ServingSizes(i).Amount.ToString() & " g)")
            DF.Columns(i + 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        Next
        DF.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

        If CFood.DataSites.ContainsKey(ManualSite.ID) Then btnAddManual.Enabled = False Else btnAddManual.Enabled = True
        If lstSites.Count > 0 Then cmbSite.SelectedIndex = 0
    End Sub

    Public Sub LoadProfile(ByVal tSite As SiteProfile)
        Dim swp As New Stopwatch()
        swp.Start()
        Try
            Rows.Clear()
            Values.Clear()
            CSite = tSite
            CSite.Properties.Sort()
            DF.Rows.Clear()
            Dim Props As List(Of FoodProperty) = CSite.Properties
            Dim CRow As System.Windows.Forms.DataGridViewRow = Nothing, CCategory As Item = Nothing
            If tSite.Site.ID = ManualSite.ID Then
                DF.ReadOnly = False
                For Each np As NutrientProperty In Nutrients
                    If np.Parent <> -1 Then Continue For
                    Dim sLine(CFood.ServingSizes.Count) As String
                    If np.Category IsNot CCategory Then
                        If CCategory IsNot Nothing Then
                            DF.Rows.Add()
                            DF.Rows(DF.Rows.Count - 1).ReadOnly = True
                        End If
                        CCategory = np.Category
                        DF.Rows.Add()
                        DF.Rows(DF.Rows.Count - 1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                        DF.Rows(DF.Rows.Count - 1).SetValues("" & CCategory.Name & "")
                        DF.Rows(DF.Rows.Count - 1).ReadOnly = True
                    End If
                    DF.Rows.Add(sLine)
                    CRow = DF.Rows(DF.Rows.Count - 1)
                    Rows.Add(np.ID, CRow.Index)
                    CRow.Tag = New List(Of NutrientProperty)
                    CRow.Cells(0).Value = np
                Next
                For Each np As NutrientProperty In Nutrients
                    If np.Parent = -1 Then Continue For
                    CRow = DF.Rows(Rows(np.Parent))
                    CRow.HeaderCell.Value = "+"
                    DirectCast(CRow.Tag, List(Of NutrientProperty)).Add(np)
                Next
                For i As Integer = 0 To Props.Count - 1
                    Values.Add(Props(i).Nutrient, Props(i).Amount)
                    If Not Rows.ContainsKey(Props(i).Nutrient) Then Continue For
                    CRow = DF.Rows(Rows(Props(i).Nutrient))
                    CRow.Cells(1).Value = Props(i).Amount.ToString()
                    For i2 As Integer = 1 To CFood.ServingSizes.Count - 1
                        CRow.Cells(i2 + 1).Value = Math.Round(((CFood.ServingSizes(i2).Amount / 100) * Props(i).Amount), 2).ToString()
                    Next
                Next
            Else
                If CSite.Properties.Count = 0 Then swp.Stop() : Exit Sub
                DF.ReadOnly = True
                For Each p As FoodProperty In Props
                    Dim np As NutrientProperty = DirectCast(Nutrients.ByID(p.Nutrient), NutrientProperty)
                    If np.Parent <> -1 Then Continue For
                    If np.Category IsNot CCategory Then
                        If CCategory IsNot Nothing Then DF.Rows.Add()
                        CCategory = np.Category
                        DF.Rows.Add()
                        DF.Rows(DF.Rows.Count - 1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                        DF.Rows(DF.Rows.Count - 1).SetValues("" & CCategory.Name & "")
                    End If
                    DF.Rows.Add(GetLine(p))
                    Rows.Add(p.Nutrient, DF.Rows.Count - 1)
                    CRow = DF.Rows(DF.Rows.Count - 1)
                    CRow.Tag = New List(Of FoodProperty)
                    CRow.Cells(0).Value = Nutrients.ByID(p.Nutrient)
                Next
                For Each p As FoodProperty In Props
                    Dim np As NutrientProperty = DirectCast(Nutrients.ByID(p.Nutrient), NutrientProperty)
                    If np.Parent = -1 Then Continue For
                    Try
                        CRow = DF.Rows(Rows(np.Parent))
                        CRow.HeaderCell.Value = "+"
                        DirectCast(CRow.Tag, List(Of FoodProperty)).Add(p)
                    Catch ex As Exception

                    End Try
                Next
            End If
        Catch ex As Exception
            MsgBox("Error loading " & CFood.Name & " profile " & CSite.Site.Name & ": " & ex.Message)
        End Try
        swp.Stop()
        frmMain.Text = swp.ElapsedMilliseconds & " ms to load " & CFood.Name & " profile."
    End Sub

    Private Function GetLine(ByVal tProp As FoodProperty) As String()
        Dim sLine(CFood.ServingSizes.Count) As String
        sLine(0) = ""
        sLine(1) = tProp.Amount.ToString()
        For i2 As Integer = 1 To CFood.ServingSizes.Count - 1
            sLine(i2 + 1) = Math.Round(((CFood.ServingSizes(i2).Amount / 100) * tProp.Amount), 2).ToString()
        Next
        Return sLine
    End Function

    Public Sub Save()
        SaveProfile()
    End Sub

    Public Sub SaveProfile()
        If CSite Is Nothing Then Exit Sub
        If CSite.Site.ID <> 2 Then Exit Sub
        CSite.Properties.Clear()
        For i As Integer = 0 To Values.Count - 1
            CSite.Properties.Add(New FoodProperty(Values.Keys(i), Values.Values(i)))
        Next
        CSite.Properties.Sort()
    End Sub

    Private Sub cmbSite_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbSite.SelectedIndexChanged
        If cmbSite.SelectedIndex = -1 Then Exit Sub
        SaveProfile()
        LoadProfile(lstSites(cmbSite.SelectedIndex))
    End Sub

    Private Sub DF_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DF.CellClick
        If e.ColumnIndex = -1 AndAlso e.RowIndex <> -1 Then
            Dim CRow As DataGridViewRow = DF.Rows(e.RowIndex)
            If CRow.Tag Is Nothing Then Exit Sub
            Dim indx As Integer = CRow.Index + 1
            If CRow.HeaderCell.Value Is "+" Then 'expand the list
                If CSite.Site.ID = ManualSite.ID Then
                    ExpandRowManual(CRow)
                Else
                    Dim lst As List(Of FoodProperty) = DirectCast(CRow.Tag, List(Of FoodProperty))
                    If lst.Count <> 0 Then
                        CRow.HeaderCell.Value = "-"
                        For i As Integer = lst.Count - 1 To 0 Step -1
                            DF.Rows.Insert(indx, GetLine(lst(i)))
                            DF.Rows(indx).Cells(0).Value = Nutrients.ByID(lst(i).Nutrient)
                            DF.Rows(indx).HeaderCell.Value = "."
                        Next
                    End If
                End If
            Else
                If CSite.Site.ID = ManualSite.ID Then
                    Dim lst As List(Of NutrientProperty) = DirectCast(CRow.Tag, List(Of NutrientProperty))
                    If lst.Count <> 0 Then
                        CRow.HeaderCell.Value = "+"
                        For i As Integer = 0 To lst.Count - 1
                            DF.Rows.RemoveAt(indx)
                        Next
                    End If
                Else
                    Dim lst As List(Of FoodProperty) = DirectCast(CRow.Tag, List(Of FoodProperty))
                    If lst.Count <> 0 Then
                        CRow.HeaderCell.Value = "+"
                        For i As Integer = 0 To lst.Count - 1
                            DF.Rows.RemoveAt(indx)
                        Next
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ExpandRowManual(ByVal CRow As DataGridViewRow)
        Dim lst As List(Of NutrientProperty) = DirectCast(CRow.Tag, List(Of NutrientProperty)), indx As Integer = CRow.Index + 1
        If lst.Count <> 0 Then
            CRow.HeaderCell.Value = "-"
            For i As Integer = lst.Count - 1 To 0 Step -1
                DF.Rows.Insert(indx, New DataGridViewRow)
                DF.Rows(indx).Cells(0).Value = lst(i)
                If Values.ContainsKey(lst(i).ID) Then
                    Dim sAmount As Single = Values(lst(i).ID)
                    DF.Rows(indx).Cells(1).Value = sAmount.ToString()
                    For i2 As Integer = 1 To CFood.ServingSizes.Count - 1
                        DF.Rows(indx).Cells(i2 + 1).Value = Math.Round(((CFood.ServingSizes(i2).Amount / 100) * sAmount), 2).ToString()
                    Next
                End If
                DF.Rows(indx).HeaderCell.Value = "."
            Next
        End If
    End Sub

    Private Sub DF_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DF.CellEndEdit
        Changed = True
        Dim txt As String = CStr(DF.CurrentCell.Value), CRow As DataGridViewRow = DF.Rows(e.RowIndex), np As NutrientProperty = DirectCast(DF.CurrentCell.OwningRow.Cells(0).Value, NutrientProperty)
        If txt = "" Then
            For i As Integer = 1 To CFood.ServingSizes.Count
                CRow.Cells(i).Value = ""
            Next
            If Values.ContainsKey(np.ID) Then
                Values.Remove(np.ID)
            End If
            Exit Sub
        End If
        Dim tDV As Single = -1
        If txt.EndsWith("%") Then
            tDV = np.NutrientData(NutrientDataPair.Fields.DV).Value
            If tDV = CSng(-1) Then
                MsgBox("Error: No Daily Value is set for this nutrient.")
                DF.CurrentCell.Value = Nothing
                Exit Sub
            End If
            txt = txt.Remove(txt.Length - 1, 1)
        End If
        If Not IsNumeric(txt) Then
            DF.CurrentCell.Value = Nothing
        Else
            Dim sValue As Single = CSng(txt), SS As ServingSize = CFood.ServingSizes(e.ColumnIndex - 1)
            If tDV <> -1 Then
                sValue = (sValue / 100) * tDV
                DF.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = sValue.ToString()
            End If
            If e.ColumnIndex > 1 Then 'additional SS changed
                DF.Rows(e.RowIndex).Cells(1).Value = Math.Round((100 / SS.Amount) * sValue, 2)
            Else
                For i As Integer = 2 To CFood.ServingSizes.Count
                    DF.Rows(e.RowIndex).Cells(i).Value = Math.Round((CFood.ServingSizes(i - 1).Amount / 100) * sValue, 2)
                Next
            End If
        End If
        If Values.ContainsKey(np.ID) Then
            Values(np.ID) = CSng(DF.Rows(e.RowIndex).Cells(1).Value)
        Else
            Values.Add(np.ID, CSng(DF.Rows(e.RowIndex).Cells(1).Value))
        End If
    End Sub

    Private Sub DF_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DF.CellMouseClick
        If Not e.Button = Windows.Forms.MouseButtons.Right Then Exit Sub
        CM.Items.Clear()
        If e.RowIndex <> -1 AndAlso e.ColumnIndex <> -1 Then
            DF.CurrentCell = DF.Rows(e.RowIndex).Cells(e.ColumnIndex)
            If Not TypeOf DF.Rows(e.RowIndex).Cells(0).Value Is NutrientProperty Then Exit Sub
            CM.Items.Add("Show Nutrient Profile")
            CM.Items(0).Tag = Nutrients.IndexOf(DirectCast(DF.Rows(e.RowIndex).Cells(0).Value, NutrientProperty))
        End If
        If e.ColumnIndex = DF.Columns.Count - 1 Then
            CM.Items.Add("Add Serving Size")
        ElseIf e.ColumnIndex > 1 Then
            CM.Items.Add("Insert Serving Size")
            CM.Items(CM.Items.Count - 1).Tag = e.ColumnIndex - 1
        End If
        If CM.Items.Count > 0 Then CM.Show(New Point(Cursor.Position.X, Cursor.Position.Y))
    End Sub

    Private Sub CM_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles CM.ItemClicked
        If e.ClickedItem.Text = "Show Nutrient Profile" Then
            F.tabNutrients.lstNutrients.lstItems.SelectedIndex = CInt(e.ClickedItem.Tag)
            F.TabMain.SelectPage(2)
        ElseIf e.ClickedItem.Text = "Insert Serving Size" Then
            Dim SS As ServingSize = frmSS.GetServingSize(CFood.ServingSizes)
            If Not SS Is Nothing Then
                Dim indx As Integer = CInt(e.ClickedItem.Tag) + 1
                CFood.ServingSizes.Insert(indx - 1, SS)
                Dim ddd As New DataGridViewColumn(DF.Columns(1).CellTemplate)
                ddd.HeaderText = SS.Name & " (" & SS.Amount & " g)"
                DF.Columns.Insert(indx, ddd)
                For i As Integer = 0 To DF.Rows.Count - 1
                    If DF.Rows(i).Cells(1).Value IsNot Nothing Then DF.Rows(i).Cells(indx).Value = (SS.Amount / 100) * CSng(DF.Rows(i).Cells(1).Value)
                Next
                Changed = True
            End If
        ElseIf e.ClickedItem.Text = "Add Serving Size" Then
            Dim SS As ServingSize = frmSS.GetServingSize(CFood.ServingSizes)
            If Not SS Is Nothing Then
                CFood.ServingSizes.Add(SS)
                DF.Columns.Add("", SS.Name & " (" & SS.Amount & " g)")
                Dim indx As Integer = DF.Columns.Count - 1
                For i As Integer = 0 To DF.Rows.Count - 1
                    If DF.Rows(i).Cells(1).Value IsNot Nothing Then DF.Rows(i).Cells(indx).Value = (SS.Amount / 100) * CSng(DF.Rows(i).Cells(1).Value)
                Next
                Changed = True
            End If
        End If
    End Sub

    Private Sub btnAddManual_Click(sender As Object, e As System.EventArgs) Handles btnAddManual.Click
        Dim st As New SiteProfile(ManualSite, "")
        CFood.DataSites.Add(st.Site.ID, st)
        cmbSite.Items.Add("Manual")
        lstSites.Add(st)
        btnAddManual.Enabled = False
        cmbSite.SelectedIndex = CFood.DataSites.Count - 1
    End Sub

    Private Sub btnExpand_Click(sender As Object, e As System.EventArgs) Handles btnExpand.Click
        If btnExpand.Text = "Expand All" Then
            For Each CRow As DataGridViewRow In DF.Rows
                If CRow.HeaderCell.Value Is "+" Then
                    CRow.HeaderCell.Value = "-"
                    Dim lst As List(Of FoodProperty) = DirectCast(CRow.Tag, List(Of FoodProperty)), indx As Integer = CRow.Index + 1
                    For i As Integer = lst.Count - 1 To 0 Step -1
                        DF.Rows.Insert(indx, GetLine(lst(i)))
                        DF.Rows(indx).Cells(0).Value = Nutrients.ByID(lst(i).Nutrient)
                        DF.Rows(indx).HeaderCell.Value = "."
                    Next
                End If
            Next
            btnExpand.Text = "Collapse All"
        Else
            For Each CRow As DataGridViewRow In DF.Rows
                If CRow.HeaderCell.Value Is "-" Then
                    CRow.HeaderCell.Value = "+"
                    Dim lst As List(Of FoodProperty) = DirectCast(CRow.Tag, List(Of FoodProperty)), indx As Integer = CRow.Index + 1
                    For i As Integer = lst.Count - 1 To 0 Step -1
                        DF.Rows.RemoveAt(indx)
                    Next
                End If
            Next
            btnExpand.Text = "Expand All"
        End If
    End Sub

End Class

Public Class NewNutrientForm : Inherits Form
    Public WithEvents txtName, tLabel, tAltNames, tParent As TextBoxX, cmbCategory, cmbUnit As ComboBoxX, btnCreate, btnCancel As ButtonX, chkUseParent As CheckBoxX
    Private bCreate As Boolean = False, DSite As DatabaseSite, LastCategory As Integer = 0, LastParent As Integer = -1, lblLastParent As LabelX

    Public Sub New()
        Dim w1 As Integer = 75, w2 As Integer = 120
        Dim lblTemp As New LabelX(Me, "", "Name:", 10, 10, w1, 22)
        txtName = New TextBoxX(Me, "txtName", "", lblTemp.Right + 2, lblTemp.Top, w2, 22)
        lblTemp = New LabelX(Me, "", "Alt Names:", lblTemp.Left, lblTemp.Bottom + 5, w1, 22)
        tAltNames = New TextBoxX(Me, "tAltNames", "", lblTemp.Right + 2, lblTemp.Top, w2, 22)
        lblTemp = New LabelX(Me, "", "Site Label:", lblTemp.Left, lblTemp.Bottom + 5, w1, 22)
        tLabel = New TextBoxX(Me, "tLabel", "", lblTemp.Right + 2, lblTemp.Top, w2, 22)
        lblTemp = New LabelX(Me, "", "Category:", lblTemp.Left, lblTemp.Bottom + 5, w1, 22)
        cmbCategory = New ComboBoxX(Me, "cmbCategory", lblTemp.Right + 2, lblTemp.Top, w2, 22)
        cmbCategory.Items.AddRange(NCategories.ToArray)
        lblTemp = New LabelX(Me, "", "Unit:", lblTemp.Left, lblTemp.Bottom + 5, w1, 22)
        cmbUnit = New ComboBoxX(Me, "cmbUnit", lblTemp.Right + 2, lblTemp.Top, w2, 22)
        cmbUnit.Items.AddRange(Units.ToArray)
        lblTemp = New LabelX(Me, "", "Parent:", lblTemp.Left, lblTemp.Bottom + 5, w1, 22)
        tParent = New TextBoxX(Me, "tParent", "", lblTemp.Right + 2, lblTemp.Top, w2, 22)
        chkUseParent = New CheckBoxX(Me, "chkUseParent", "Use Parent", lblTemp.Left, lblTemp.Bottom + 5, 85, 22)
        lblLastParent = New LabelX(Me, "lblLastParent", "N/A", chkUseParent.Right + 2, chkUseParent.Top, w2, 22)
        btnCreate = New ButtonX(Me, "btnCreate", "Create", chkUseParent.Left, chkUseParent.Bottom + 5, w1, 22)
        btnCancel = New ButtonX(Me, "btnCancel", "Cancel", btnCreate.Right + 10, chkUseParent.Bottom + 5, w1, 22)
    End Sub

    Private Sub NewNutrient_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim w As Integer = 270, h As Integer = 270
        Me.SetBounds(CInt(frmMain.Left + (frmMain.Width - w) / 2), CInt(frmMain.Top + (frmMain.Height - h) / 2), w, h)
    End Sub

    Public Function GetNutrient(ByVal sName As String, ByVal tSite As DatabaseSite, ByVal tUnit As Unit) As NutrientProperty
        Dim nu As New NutrientProperty(sName, GetNextID), indx As Integer = sName.IndexOf("(")
        bCreate = False
        DSite = tSite
        tLabel.Text = sName
        If indx <> -1 Then
            txtName.Text = sName.Substring(0, indx).Trim()
            tAltNames.Text = sName.Substring(indx + 1, sName.IndexOf(")", indx) - indx - 1).Trim()
        Else
            txtName.Text = sName.Trim()
            tAltNames.Text = ""
        End If
        cmbCategory.SelectedIndex = LastCategory
        chkUseParent.Checked = False
        cmbUnit.SelectedIndex = Units.IndexOf(tUnit)
        If LastParent <> -1 Then
            lblLastParent.Text = Nutrients.ByID(LastParent).Name
            tParent.Text = LastParent.ToString()
        Else
            lblLastParent.Text = ""
            tParent.Text = ""
        End If
        ShowDialog()
        If bCreate Then
            nu.Name = txtName.Text.Trim()
            nu.AlternateNames = tAltNames.Text.Split(Chr(44)).ToList()
            If nu.AlternateNames.Count = 1 AndAlso nu.AlternateNames(0) = "" Then nu.AlternateNames = New List(Of String)
            nu.Category = DirectCast(cmbCategory.SelectedItem, Item)
            Dim tempUnit As Unit = DirectCast(cmbUnit.SelectedItem, Unit)
            For i As Integer = 0 To 3
                nu.NutrientData(i) = New NutrientDataPair(CByte(i), -1, tempUnit)
            Next
            If chkUseParent.Checked AndAlso tParent.Text <> "" Then
                Dim id As Integer = CInt(tParent.Text)
                nu.Parent = id
                LastParent = id
                DirectCast(Nutrients.ByID(id), NutrientProperty).SubProperties.Add(nu.ID)
            Else
                LastParent = nu.ID
            End If
            DSite.AddTerm(nu, tLabel.Text)
            LastCategory = cmbCategory.SelectedIndex
            Nutrients.Add(nu)
            Return nu
        Else
            Return Nothing
        End If
    End Function

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As System.EventArgs) Handles btnCreate.Click
        bCreate = True
        Me.Close()
    End Sub
End Class

Public Class FoodComponents : Inherits PanelX
    Public WithEvents btnAdd As ButtonX
    Public Lines As New List(Of FoodLine)
    Public SSWidth As Integer
    Public Changed As Boolean = False, Loading As Boolean = False
    Private LineHeight As Integer = 22, LineSpacingH As Integer = 0, LineIndent As Integer = 5

    Public Sub New(ByVal prnt As Control, ByVal sName As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, Optional ByVal tSSWidth As Integer = 130)
        ' Me.SetStyle(ControlStyles.DoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
        MyBase.New(prnt, sName, x, y, w, h)
        SSWidth = tSSWidth
        btnAdd = New ButtonX(Me, "btnAdd", "Add", Me.ClientSize.Width - 44, 0, 40, 20)
        AddLine()
        btnAdd.Left = Lines(0).tAmount.Right - btnAdd.Width
        AutoScroll = True
    End Sub

    Public Overloads Sub AddLine()
        Lines.Add(New FoodLine(Me, LineIndent + Me.AutoScrollPosition.X, ((LineHeight + LineSpacingH) * Lines.Count) + Me.AutoScrollPosition.Y, Me.ClientSize.Width - LineIndent - Me.AutoScrollPosition.X, LineHeight))
        btnAdd.Top = Lines(Lines.Count - 1).Bottom + 3
    End Sub

    Public Sub DeleteLine(ByVal tLine As FoodLine)
        Dim indx As Integer = Lines.IndexOf(tLine)
        Lines.RemoveAt(indx)
        tLine.Dispose()
        For i As Integer = indx To Lines.Count - 1
            Lines(i).Top -= LineHeight
        Next
        btnAdd.Top = Lines(Lines.Count - 1).Bottom + 3
        If Not Loading Then Changed = True
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
        AddLine()
        Changed = True
    End Sub

    Public Sub AddRecipe(ByVal tRecipe As Recipe)
        Dim LineTop As Integer = 0
        If Lines.Count > 0 Then LineTop = Lines(Lines.Count - 1).Bottom + LineSpacingH
        For Each fe As FoodEntry In tRecipe.Entries
            Dim FL As New FoodLine(Me, LineIndent, LineTop, Me.ClientSize.Width - LineIndent, LineHeight)
            FL.cmbFood.SelectedItem = fe.Food
            FL.cmbServingSize.SelectedIndex = fe.ServingSize
            FL.tAmount.Text = fe.Amount.ToString()
            Lines.Add(FL)
            LineTop = FL.Bottom + LineSpacingH
        Next
        btnAdd.Top = Lines(Lines.Count - 1).Bottom + 3
    End Sub

    Public Sub SetRecipes(ByVal tList As List(Of FoodEntry))
        Loading = True
        If tList.Count = 0 Then Clear()
        Try
            If tList.Count > Lines.Count Then
                For i As Integer = Lines.Count To tList.Count - 1 'if there are more items in the list, add more lines
                    Lines.Add(New FoodLine(Me, LineIndent, ((LineHeight + LineSpacingH) * Lines.Count), Me.ClientSize.Width - LineIndent, LineHeight))
                Next
            End If
        Catch ex As Exception
            MsgBox("Error in SetRecipes, adding new lines: " & ex.Message)
        End Try
        Try
            For i As Integer = 0 To tList.Count - 1
                Dim FLine As FoodLine = Lines(i)
                FLine.cmbFood.SelectedIndex = Foods.IndexOf(tList(i).Food)
                FLine.cmbServingSize.SelectedIndex = tList(i).ServingSize
                FLine.tAmount.Text = tList(i).Amount.ToString()
            Next
        Catch ex As Exception
            MsgBox("Error in SetRecipes, converting old lines: " & ex.Message)
        End Try
        Try
            If Lines.Count > tList.Count Then
                Lines(tList.Count).Clear()
                Dim count As Integer = Lines.Count - 1
                For i As Integer = count To tList.Count + 1 Step -1
                    DeleteLine(Lines(i))
                Next
            Else
                AddLine()
            End If
        Catch ex As Exception
            MsgBox("Error in SetRecipes, deleting old lines: " & ex.Message)
        End Try
        btnAdd.Top = Lines(Lines.Count - 1).Bottom + 3
        Loading = False
    End Sub

    Public Sub Clear()
        Lines.Clear()
        For i As Integer = Controls.Count - 1 To 0 Step -1
            If TypeOf Controls(i) Is FoodLine Then Controls.RemoveAt(i)
        Next
        AddLine()
    End Sub

    Public Class FoodLine : Inherits PanelX
        Public WithEvents cmbFood As ComboBoxX, tAmount As DecimalTextBox, cmbServingSize As ComboBoxX, btnExit As ButtonX
        Public FC As FoodComponents, Labels(3) As LabelX, CFood As FoodItem

        Public Sub New(ByVal prnt As FoodComponents, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
            MyBase.New(prnt, "", x, y, w, h)
            ' BorderStyle = Windows.Forms.BorderStyle.FixedSingle
            FC = prnt
            btnExit = New ButtonX(Me, "btnExit", "X", 0, 2, 20, h - 4)
            Labels(0) = New LabelX(Me, "", "Food:", btnExit.Right + 2, 3, -1, -1)
            cmbFood = New ComboBoxX(Me, "cmbFood", Labels(0).Right + 2, 1, 240, h - 2)
            cmbFood.MaxDropDownItems = 20
            cmbFood.Items.AddRange(Foods.ToArray)
            Labels(1) = New LabelX(Me, "", "Serving Size:", cmbFood.Right + 2, 3, -1, -1)
            cmbServingSize = New ComboBoxX(Me, "cmbServingSize", Labels(1).Right + 2, 1, FC.SSWidth, h - 2)
            Labels(2) = New LabelX(Me, "", "Amount:", cmbServingSize.Right + 2, 3, -1, -1)
            tAmount = New DecimalTextBox(Me, "tAmount", "", Labels(2).Right + 2, 1, 45, h - 2)
        End Sub

        Public Sub Clear()
            cmbFood.SelectedIndex = -1
            cmbServingSize.Items.Clear()
            tAmount.Text = ""
            CFood = Nothing
        End Sub

        Private Sub btnExit_Click(sender As Object, e As System.EventArgs) Handles btnExit.Click
            If FC.Lines.Count = 1 Then
                Clear()
            Else
                If cmbFood.SelectedIndex = -1 Then Exit Sub
                If MessageBox.Show("Remove " & cmbFood.SelectedText & " from the recipe?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                FC.DeleteLine(Me)
            End If
        End Sub

        Private Sub cmbFood_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFood.SelectedIndexChanged
            If cmbFood.SelectedIndex = -1 Then Exit Sub
            CFood = DirectCast(cmbFood.SelectedItem, FoodItem)
            cmbServingSize.Items.Clear()
            cmbServingSize.Items.AddRange(CFood.ServingSizes.ToArray)
            cmbServingSize.SelectedIndex = 0
            If Not FC.Loading Then FC.Changed = True
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            CFood = Nothing
            FC = Nothing
            Labels = Nothing
            MyBase.Dispose(disposing)
        End Sub

        Private Sub cmbServingSize_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbServingSize.SelectedIndexChanged, tAmount.TextChanged
            If Not FC.Loading Then FC.Changed = True
        End Sub
    End Class

End Class

Public Class MessageBox2

    Public Shared Function Show(ByVal Messages As List(Of String()), tMessageParameters As MsgParams) As Integer
        Dim MB As New MBox2(Messages, tMessageParameters)
        Show = MB.ShowPromptBox()
    End Function

    Public Class MsgParams
        Public Title As String = "", Format As String = ""
        Public CenterForm As Form
        Public HighlightPercents As Boolean = False
        Public InvertHighlights As Boolean = False
        Public PositiveVsNegative As Boolean = False
    End Class

    Protected Class MBox2 : Inherits PromptBox
        Protected tBox As RichTextBox2
        Protected Class RichTextBox2 : Inherits RichTextBox
            Public Sub AppendText2(tText As String)
                SelectionStart = TextLength
                SelectionLength = 0
                SelectionFont = New Font(Font.FontFamily, Font.Size, FontStyle.Underline)
                AppendText(tText)
                SelectionFont = Me.Font
            End Sub
        End Class

        Public Sub New(ByVal Messages As List(Of String()), tMessageParameters As MsgParams)
            MyBase.New(tMessageParameters.Title, tMessageParameters.CenterForm, New List(Of String) From {"OK"})
            tBox = New RichTextBox2() : tBox.Parent = Me : tBox.Location = New Point(15, 15)
            tBox.Font = New Font("Lucida Console", 10, FontStyle.Regular)
            tBox.ReadOnly = True : tBox.Text = "" : tBox.WordWrap = False
            Dim sb As New System.Text.StringBuilder(), longestwidth As Integer = 0
            Dim sw As Stopwatch = Stopwatch.StartNew(), SpaceChar As Char = Chr(32), indx As Integer = -1
            If tMessageParameters.Format = "" Then
                For Each Str As String() In Messages
                    Dim lol As String = String.Join(" ", Str)
                    longestwidth = Math.Max(longestwidth, lol.Length)
                    sb.Append(lol & Chr(10))
                Next
            Else
                For Each Str As String() In Messages
                    Dim lol As String = String.Format(tMessageParameters.Format, Str)
                    Dim lol2 As String = lol.Trim()
                    longestwidth = Math.Max(longestwidth, lol.Length)
                    If lol2.Length > 0 And lol2.Length < 20 Then
                        tBox.AppendText2(lol2)
                    Else
                        tBox.AppendText(lol)
                    End If
                    If tMessageParameters.HighlightPercents Then
                        indx = lol.IndexOf("%")
                        If indx <> -1 Then
                            For i2 As Integer = indx To 0 Step -1
                                If lol.Chars(i2) = SpaceChar Then
                                    Dim yesto As String = lol.Substring(i2 + 1, indx - i2 - 1).Trim()
                                    If IsNumeric(yesto) Then
                                        Dim omgs As Single = CSng(yesto), tColor As Color = Color.Red
                                        If tMessageParameters.PositiveVsNegative Then
                                            If omgs = 0 Then Exit For
                                            If omgs > 0 Then tColor = Color.Green
                                        Else
                                            If omgs >= 100 Then
                                                If Not tMessageParameters.InvertHighlights Then tColor = Color.Green
                                            Else
                                                If tMessageParameters.InvertHighlights Then tColor = Color.Green
                                            End If
                                        End If

                                        Dim lolo As Integer = lol.Length - indx + yesto.Length
                                        tBox.SelectionStart = tBox.TextLength - lolo
                                        tBox.SelectionLength = yesto.Length
                                        tBox.SelectionColor = tColor
                                        tBox.SelectionStart = tBox.TextLength
                                        tBox.SelectionColor = ForeColor
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    tBox.AppendText(Chr(10))
                Next
            End If
            sw.Stop() '120
            tBox.Size = New Size(longestwidth * 8 + 15, tBox.Font.Height * Messages.Count + 10)
        End Sub

    End Class
End Class

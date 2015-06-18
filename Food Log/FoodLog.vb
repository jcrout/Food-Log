Option Strict On

Imports System
Imports System.Text
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime
Imports System.Drawing
Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Linq.Expressions
Imports MyControls

Public Class FoodLog
    Public WithEvents TabMain As TabBrowser, TCM As ContextMenuStrip
    Public tabNutrients As NutrientEditorPage, tabFood As FoodEditorPage, tabEntries As FoodLogPage
    Public tabRecipes As RecipeEditorPage, tabEditor As EditorPage
    Public CurrentPage As FLPage, AllPages As List(Of FLPage)
    Private bExitWithoutSaving As Boolean = False

    Public Sub New()
        LoadArrays()
        LoadControls()
        tabEntries.InitialLoad()
    End Sub

    Private Sub LoadArrays()
        NCategories.Add(New Item("Macronutrients", 0))
        NCategories.Add(New Item("Vitamins", 1))
        NCategories.Add(New Item("Minerals", 2))
        NCategories.Add(New Item("Other", 3))
        NCategories.Add(New Item("Anthocyanidins", 4))
        NCategories.Add(New Item("Proanthocyanidin", 5))
        NCategories.Add(New Item("Flavan-3-ols", 6))
        NCategories.Add(New Item("Flavanones", 7))
        NCategories.Add(New Item("Flavones", 8))
        NCategories.Add(New Item("Flavonols", 9))
        NCategories.Add(New Item("Isoflavones", 10))

        FCategories.Add(New Item("N/A", 0))
        FCategories.Add(New Item("Vegetables", 1))
        FCategories.Add(New Item("Fruits", 2))
        FCategories.Add(New Item("Whole Grains", 3))
        FCategories.Add(New Item("Protein Foods", 4))
        FCategories.Add(New Item("Dairy", 5))
        FCategories.Add(New Item("Junk", 6))
        FCategories.Add(New Item("Processed Foods", 7))
        FCategories.Add(New Item("Condiments", 8))
        FCategoriesNA = FCategories(0)

        LoadUnits()
        LoadNutrients()
        LoadSites()
        LoadFoods()
        LoadRecipes()
        LoadLogs()
    End Sub

    Private Sub LoadControls()
        frmMain = New Form1
        TabMain = New TabBrowser("TabMain", frmMain, 0, 0, Form1.w - 16, Form1.h - 37, "Entries", False)
        TabMain.AddPage("Recipes", Nothing, Nothing, Nothing, False, False)
        TabMain.AddPage("Food Editor", Nothing, Nothing, Nothing, False, False)
        TabMain.AddPage("Nutrient Editor", Nothing, Nothing, Nothing, False, False)
        TabMain.AddPage("Editor", Nothing, Nothing, Nothing, False, False)
        TabMain.ContextMenuStrip = New ContextMenuStrip()
        TCM = TabMain.ContextMenuStrip
        TCM.Items.Add("Exit Without Saving")
        AllPages = New List(Of FLPage)
        tabNutrients = New NutrientEditorPage(TabMain.Pages(3)) : AllPages.Add(tabNutrients)
        tabFood = New FoodEditorPage(TabMain.Pages(2)) : AllPages.Add(tabFood)
        tabRecipes = New RecipeEditorPage(TabMain.Pages(1)) : AllPages.Add(tabRecipes)
        tabEntries = New FoodLogPage(TabMain.Pages(0)) : AllPages.Add(tabEntries)
        tabEditor = New EditorPage(TabMain.Pages(4)) : AllPages.Add(tabEditor)
        TabMain.SelectPage(0)
        AddHandler frmMain.Resize, AddressOf frmMain_Resize
    End Sub

    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As EventArgs)
        TabMain.SetWidth(frmMain.ClientSize.Width)
    End Sub

    Private Sub TCM_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles TCM.ItemClicked
        If e.ClickedItem.Text = "Exit Without Saving" Then
            bExitWithoutSaving = True
            frmMain.Close()
        End If
    End Sub

    Private Sub TabMain_DoubleClick(sender As Object, e As System.EventArgs) Handles TabMain.DoubleClick
        Process.Start("explorer.exe", Application.StartupPath)
    End Sub

    Private Sub TabMain_PageChanging(sender As TabBrowser, e As TabBrowser.PageChangingArgs) Handles TabMain.PageChanging
        If e.NewIndex = 0 Then tabEntries.UpdateItems()
        If CurrentPage IsNot Nothing Then CurrentPage.SaveLists()
        CurrentPage = AllPages.Find(Function(x) x.Page Is TabMain.Pages(e.NewIndex))
    End Sub

    Public Sub Save()
        If bExitWithoutSaving Then Exit Sub
        'SaveUnits()
        'SaveSites()
        If CurrentPage IsNot Nothing Then CurrentPage.SaveLists()
        tabEntries.Save()
        SaveLogs()
    End Sub

    Private Sub LoadRecipes()
        Dim sFile As String = CDirectory & "Recipes.txt"
        If Not File.Exists(sFile) Then Exit Sub
        Try
            Using sr As New StreamReader(sFile, System.Text.Encoding.UTF8)
                Do While sr.Peek <> -1
                    Recipies.Add(Recipe.FromLine(sr.ReadLine()))
                Loop
            End Using
        Catch ex As Exception
            MsgBox("Error loading Recipes: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadSites()
        Dim sFile As String = CDirectory & "Sites.txt"
        If Not File.Exists(sFile) Then Exit Sub
        Try
            Dim sr As New StreamReader(sFile, System.Text.Encoding.UTF8)
            sr.ReadLine()
            Do While sr.Peek <> -1
                Dim sParts() As String = sr.ReadLine().Split(Chr(124)), st As DatabaseSite
                If sParts(0) = "0" Then st = New DatabaseSite.USDA(sParts(1), CInt(sParts(0)), sParts(2)) Else st = New DatabaseSite.WHFoods(sParts(1), CInt(sParts(0)), sParts(2))
                If sParts(3) <> "" Then
                    For i As Integer = 3 To sParts.Length - 1
                        Dim indx As Integer = sParts(i).IndexOf(";")
                        st.AddTerm(DirectCast(Nutrients.ByID(CInt(sParts(i).Substring(0, indx))), NutrientProperty), sParts(i).Substring(indx + 1))
                    Next
                End If
                Sites.Add(st)
            Loop
            sr.Close()
            ManualSite = New DatabaseSite.Manual("Manual", 2, "")
            ComboSite = New DatabaseSite.Manual("Combo", 3, "")
            Sites.Add(ManualSite)
            Sites.Add(ComboSite)
        Catch ex As Exception
            MsgBox("Error loading Sites: " & ex.Message)
        End Try
    End Sub

    Private Sub SaveSites()
        Dim sFile As String = CDirectory & "Sites.txt", sTempFile As String = CDirectory & "TEMPSites.txt", sw As StreamWriter = Nothing, sLine As StringBuilder
        If File.Exists(sTempFile) Then
            If MessageBox.Show("TEMPSites.txt file exists, continue saving?", "Uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
            File.Delete(sTempFile)
        End If
        Try
            sw = New StreamWriter(sTempFile, False, System.Text.Encoding.UTF8)
            For i As Integer = 0 To Sites.Count - 1
                Dim st As DatabaseSite = DirectCast(Sites(i), DatabaseSite) : sLine = New StringBuilder(st.ID & "|" & st.Name & "|" & st.SiteBase & "|")
                If st.Terms.Count > 0 Then
                    For Each kvp As KeyValuePair(Of String, NutrientProperty) In st.Terms
                        sLine.Append(kvp.Value.ID & ";" & kvp.Key & "|")
                    Next
                    sLine.Remove(sLine.Length - 1, 1)
                End If
                sw.WriteLine(sLine.ToString())
            Next
        Catch ex As Exception
            MsgBox("Error saving TEMPSites: " & ex.Message)
        End Try
        sw.Close()
        Try
            If File.Exists(sFile) Then File.Delete(sFile)
            My.Computer.FileSystem.RenameFile(sTempFile, "Sites.txt")
        Catch ex As Exception
            MsgBox("Error renmaing TempSites: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadUnits()
        Dim sFile As String = CDirectory & "Units.txt"
        If Not File.Exists(sFile) Then Exit Sub
        Try
            Dim sr As New StreamReader(sFile, System.Text.Encoding.UTF8)
            Dim sParts() As String = sr.ReadLine().Split(Chr(124))
            For i As Integer = 1 To sParts.Length - 1 '2|0;0;Gram;g;1|
                Dim sSubParts() As String = sParts(i).Split(Chr(59))
                Units.Add(New Unit(CInt(sSubParts(1)), sSubParts(2), CInt(sSubParts(0)), sSubParts(3).Split(Chr(126)).ToList(), CSng(sSubParts(4))))
            Next
            sr.Close()
        Catch ex As Exception
            MsgBox("Error loading Units: " & ex.Message)
        End Try
    End Sub

    Private Sub SaveUnits()
        Dim sFile As String = CDirectory & "Units.txt", sTempFile As String = CDirectory & "TEMPUnits.txt", sw As StreamWriter = Nothing, sLine As StringBuilder
        If File.Exists(sTempFile) Then
            If MessageBox.Show("TEMPUnits.txt file exists, continue saving?", "Uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
            File.Delete(sTempFile)
        End If
        Try
            sw = New StreamWriter(sTempFile, False, System.Text.Encoding.UTF8)
            sLine = New StringBuilder()
            For i As Integer = 0 To Units.Count - 1
                Dim u As Unit = DirectCast(Units(i), Unit)
                sLine.Append("|" & u.ID & ";" & u.BUnit & ";" & u.Name & ";" & String.Join("~", u.Abbrev) & ";" & u.Relation)
            Next
            sLine.Remove(0, 1)
            sw.WriteLine(sLine.ToString())
        Catch ex As Exception
            MsgBox("Error saving TEMPUnits: " & ex.Message)
        End Try
        sw.Close()
        Try
            If File.Exists(sFile) Then File.Delete(sFile)
            My.Computer.FileSystem.RenameFile(sTempFile, "Units.txt")
        Catch ex As Exception
            MsgBox("Error renmaing TempUnits: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadNutrients()
        Dim sFile As String = CDirectory & "Nutrients.txt"
        If Not File.Exists(sFile) Then Exit Sub
        Try
            Using sr As New StreamReader(sFile)
                Do While sr.Peek <> -1
                    Nutrients.Add(NutrientProperty.FromLine(sr.ReadLine().Split(Chr(124))))
                Loop
            End Using
            ReportX.OmegaFats(0) = New List(Of Integer) From {71, 74, 75, 76}
            ReportX.OmegaFats(1) = New List(Of Integer) From {70, 73}
        Catch ex As Exception
            MsgBox("Error loading Nutrients: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadFoods()
        Dim sFile As String = CDirectory & "Foods.txt"
        If Not File.Exists(sFile) Then Exit Sub
        Dim swp As New Stopwatch(), fd As FoodItem = Nothing
        swp.Start()
        Try
            Using sr As New StreamReader(sFile, Encoding.UTF8)
                Do While sr.Peek <> -1
                    Foods.Add(FoodItem.FromSR(sr))
                Loop
            End Using
            For i As Integer = 0 To Foods.Count - 1
                Dim fo As FoodItem = DirectCast(Foods(i), FoodItem)
                If fo.Variations.Count > 0 Then
                    For i2 As Integer = 0 To fo.Variations.Count - 1
                        DirectCast(Foods(i + i2 + 1), FoodItem).Parent = fo
                    Next
                    i += fo.Variations.Count
                End If
            Next
            'SortFoods()
        Catch ex As Exception
            If fd Is Nothing Then
                MsgBox("Error loading Foods: " & ex.Message)
            Else
                MsgBox("Error loading food " & Chr(34) & fd.Name & Chr(34) & ": " & ex.Message)
            End If
        End Try
        swp.Stop()
    End Sub

    Public Sub SortFoods()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Foods.Sort()
        Dim doops As New ItemList()
        For Each fo As FoodItem In Foods
            If fo.Parent IsNot Nothing Then Continue For
            doops.Add(fo)
            If fo.Variations.Count > 0 Then
                For Each varz As Integer In fo.Variations
                    Dim fo2 As FoodItem = DirectCast(Foods.ByID(varz), FoodItem)
                    doops.Add(fo2)
                Next
            End If
        Next
        Foods = doops
        sw.Stop()
    End Sub

    Private Sub SaveLogs()
        Dim sFile As String = CDirectory & "Logs.txt", sTempFile As String = CDirectory & "TEMPLogs.txt"
        If File.Exists(sTempFile) Then
            If MessageBox.Show("TEMPLogs.txt file exists, continue saving?", "Uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
            File.Delete(sTempFile)
        End If
        Try
            Using sw As New StreamWriter(sTempFile, False, Encoding.UTF8)
                sw.WriteLine((Logs.Count - 2).ToString())
                For i As Integer = 0 To Logs.Count - 2 'don't include the last log, the Test Log
                    Logs(i).Save(sw)
                Next
            End Using
        Catch ex As Exception
            MsgBox("Error saving TEMPLogs: " & ex.Message) : Exit Sub
        End Try
        Try
            If File.Exists(sFile) Then File.Delete(sFile)
            My.Computer.FileSystem.RenameFile(sTempFile, "Logs.txt")
        Catch ex As Exception
            MsgBox("Error renmaing TempLogs: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadLogs()
        Dim sFile As String = CDirectory & "Logs.txt" '95 = _
        If Not File.Exists(sFile) Then Exit Sub
        Dim swp As New Stopwatch(), tLog As LogEntry = Nothing
        swp.Start()
        Try
            Using sr As New StreamReader(sFile, Encoding.UTF8)
                Dim LogCount As Integer = CInt(sr.ReadLine())
                For i As Integer = 0 To LogCount
                    Logs.Add(LogEntry.FromLine(sr))
                Next
            End Using
        Catch ex As Exception
            If tLog Is Nothing Then
                MsgBox("Error loading Logs: " & ex.Message)
            Else
                MsgBox("Error loading log " & Chr(34) & tLog.ToString() & Chr(34) & ": " & ex.Message)
            End If
        End Try
        swp.Stop()
    End Sub

    Public Class ReportX
        Public Shared OmegaFats(1) As List(Of Integer)
        Public LongestName As Integer = 0, CompletedLogs As Integer = 0
        Public Params As New Dictionary(Of Integer, ParamsList)
        Public Omega6to3Ratio As String

        Public Class ParamsList
            Public TotalAmount As Single = 0, Used As Boolean = False, MissingList As New List(Of FoodItem)
            Public DRI100Days, DRI75Days, DRILowDays As Integer, CurrentTotal As Single = 0
            Public Highest As New StringAndNum("", 0), Lowest As New StringAndNum("", 999999999)
        End Class

        Public Sub New(ByVal bSet As Boolean)
            If bSet Then
                For Each n As NutrientProperty In Nutrients
                    If n.Basic Then
                        LongestName = Math.Max(n.Name.Length, LongestName)
                        Params.Add(n.ID, New ParamsList())
                    End If
                Next
            End If
        End Sub

        Public Sub Report(ByVal Logz As List(Of LogEntry), ByVal Entries As List(Of FoodEntry), Optional ByVal Extended As Boolean = False)
            Dim sw As Stopwatch = Stopwatch.StartNew()
            If Logz Is Nothing Then Logz = Logs
            Dim tAmount As Single = 0, lstSites As New List(Of Integer) From {3, 2, 0}, MultipleLogs As Boolean = Logz.Count > 1
            For Each lg As LogEntry In Logz
                If MultipleLogs Then
                    If Not lg.Completed Or Not lg.AllMealsConsumed Then Continue For
                    CompletedLogs += 1
                End If
                Dim Entriez As List(Of FoodEntry)
                If Entries Is Nothing Then Entriez = lg.FoodEntries Else Entriez = Entries
                For i As Integer = 0 To Entriez.Count - 1
                    Dim tFE As FoodEntry = Entriez(i)
                    If tFE.Amount = 0 Then Continue For
                    Dim tFood As FoodItem = Entriez(i).Food
                    If tFE.ServingSize = 0 Then
                        tAmount = tFE.Amount
                    Else
                        tAmount = tFood.ServingSizes(tFE.ServingSize).Amount * tFE.Amount
                    End If
                    tAmount /= 100
                    Do
                        For i2 As Integer = 0 To lstSites.Count - 1
                            If tFood.DataSites.ContainsKey(lstSites(i2)) Then
                                Dim props As List(Of FoodProperty) = tFood.DataSites(lstSites(i2)).Properties
                                For i3 As Integer = 0 To props.Count - 1
                                    Dim ID As Integer = props(i3).Nutrient
                                    If Params.ContainsKey(ID) Then
                                        Dim PL As ParamsList = Params(ID)
                                        If PL.Used = True Then Continue For
                                        Dim tempAmount As Single = tAmount * props(i3).Amount
                                        PL.CurrentTotal += tempAmount
                                        PL.TotalAmount += tempAmount
                                        PL.Used = True
                                    End If
                                Next
                            End If
                        Next
                        If tFood.Parent Is Nothing Then Exit Do
                        tFood = tFood.Parent
                    Loop
                    For Each kvp As KeyValuePair(Of Integer, ParamsList) In Params
                        Dim PL As ParamsList = kvp.Value
                        If PL.Used = False Then
                            If Not PL.MissingList.Contains(tFood) Then PL.MissingList.Add(tFood)
                        Else
                            PL.Used = False
                        End If
                    Next
                Next
                If Extended Then
                    For Each kvp As KeyValuePair(Of Integer, ParamsList) In Params
                        Dim PL As ParamsList = kvp.Value
                        Dim n As NutrientProperty = DirectCast(Nutrients.ByID(kvp.Key), NutrientProperty)
                        Dim nDRI As NutrientDataPair = n.NutrientData(NutrientDataPair.Fields.DRI)
                        If nDRI.Value <> -1 Then
                            Dim tValue As Single = GetAdjustedValue(n, nDRI.Unit, PL.CurrentTotal)
                            Dim pct As Double = tValue / nDRI.Value
                            If pct >= 1 Then
                                PL.DRI100Days += 1
                            ElseIf pct >= 0.75 Then
                                PL.DRI75Days += 1
                            Else
                                PL.DRILowDays += 1
                            End If
                        End If
                        If PL.CurrentTotal > PL.Highest.Number Then PL.Highest = New StringAndNum(lg.EntryDate.ToString(), PL.CurrentTotal)
                        If PL.CurrentTotal < PL.Lowest.Number Then PL.Lowest = New StringAndNum(lg.EntryDate.ToString(), PL.CurrentTotal)
                        PL.CurrentTotal = 0
                    Next
                End If
            Next
            If Params.Count > 3 Then
                Dim Omegas(1) As Single
                For i As Integer = 0 To 1
                    For Each intz As Integer In OmegaFats(i)
                        Omegas(i) += Params(intz).TotalAmount
                    Next
                Next
                Dim ratio As Single = Omegas(1) / Omegas(0)
                If ratio > 100 Then
                    Omega6to3Ratio = ">100"
                Else
                    Omega6to3Ratio = Math.Round(Omegas(1) / Omegas(0), 2).ToString()
                End If
            End If
            sw.Stop()
        End Sub
    End Class

    Public MustInherit Class FLPage
        Public Page As TabBrowser.Page, Changed As Boolean = False, Loading As Boolean = False

        Public Overridable ReadOnly Property FileName() As String
            Get
                Return ""
            End Get
        End Property

        Public Sub SaveLists()
            Save()
            If Changed Then
                Changed = False
                Dim sFile As String = CDirectory & FileName & ".txt", sTempFile As String = CDirectory & "TEMP" & FileName & ".txt"
                If File.Exists(sTempFile) Then
                    If MessageBox.Show("TEMPFoods.txt file exists, continue saving?", "Uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                    File.Delete(sTempFile)
                End If
                Try
                    Using sw As New StreamWriter(sTempFile)
                        SaveList(sw)
                    End Using
                Catch ex As Exception
                    MsgBox("Error saving TEMP" & FileName & ": " & ex.Message) : Exit Sub
                End Try
                Try
                    If File.Exists(sFile) Then File.Delete(sFile)
                    My.Computer.FileSystem.RenameFile(sTempFile, FileName & ".txt")
                Catch ex As Exception
                    MsgBox("Error renmaing TEMP" & FileName & ": " & ex.Message)
                End Try
            End If
        End Sub

        Protected Overridable Sub SaveList(ByVal sw As StreamWriter)

        End Sub

        Public Overridable Sub Save()

        End Sub
    End Class

    Public Class EditorPage : Inherits FLPage
        Public WithEvents btnBackup, btnClean As ButtonX

        Public Sub New(ByVal tPage As TabBrowser.Page)
            Try
                Page = tPage
                LoadControls()
            Catch ex As Exception
                MsgBox("Error loading FoodLog page: " & ex.Message)
            End Try
        End Sub

        Private Sub LoadControls()
            btnBackup = New ButtonX(Page, "btnBackup", "Backup Files", 15, 15, 80, 22)
            btnClean = New ButtonX(Page, "btnClean", "Clean Data", 15, btnBackup.Bottom + 5, 80, 22)
        End Sub

        Private Sub btnBackup_Click(sender As Object, e As System.EventArgs) Handles btnBackup.Click
            Try
                Dim sw As New Stopwatch() : sw.Start()
                Dim BArchive As String = CDirectory & "0Archives\"
                With Date.Now
                    BArchive &= .Year.ToString() & FormatTwoDigits(.Month) & FormatTwoDigits(.Day) & "-" & _
                            FormatTwoDigits(.Hour) & FormatTwoDigits(.Minute) & "\"
                End With
                If Directory.Exists(BArchive) Then
                    MsgBox("Error: You can only back up files once per minute.")
                    Exit Sub
                End If
                Directory.CreateDirectory(BArchive)
                My.Computer.FileSystem.CopyFile(CDirectory & "Foods.txt", BArchive & "Foods.txt")
                My.Computer.FileSystem.CopyFile(CDirectory & "Logs.txt", BArchive & "Logs.txt")
                My.Computer.FileSystem.CopyFile(CDirectory & "Nutrients.txt", BArchive & "Nutrients.txt")
                My.Computer.FileSystem.CopyFile(CDirectory & "Recipes.txt", BArchive & "Recipes.txt")
                My.Computer.FileSystem.CopyFile(CDirectory & "Sites.txt", BArchive & "Sites.txt")
                My.Computer.FileSystem.CopyFile(CDirectory & "Units.txt", BArchive & "Units.txt")
                sw.Stop()
                MsgBox("Files successfully backed up (" & sw.ElapsedMilliseconds.ToString() & " ms)")
            Catch ex As Exception
                MsgBox("Error backing up files: " & ex.Message)
            End Try
        End Sub

        Private Sub btnClean_Click(sender As Object, e As System.EventArgs) Handles btnClean.Click
            Dim tFood As FoodItem = Nothing, iCount As Integer = 0
            For i As Integer = 0 To Foods.Count - 1
                tFood = DirectCast(Foods(i), FoodItem)
                For i2 As Integer = tFood.DataSites.Count - 1 To 0 Step -1
                    If tFood.DataSites.Values(i2).Site.ID <> ManualSite.ID Then Continue For
                    If tFood.DataSites.Values(i2).Properties.Count = 0 Then
                        tFood.DataSites.Remove(ManualSite.ID)
                        iCount += 1
                        Exit For
                    End If
                Next
            Next
            MsgBox("Removed " & iCount.ToString() & " empty site entries.")
        End Sub
    End Class

    Public Class FoodLogPage : Inherits FLPage
        Public WithEvents lblLogDate, lblFood As LabelX, lstFoods As ListItems, CM As ContextMenuStrip
        Public WithEvents cmbRecipe, cmbFood, cmbServing, cmbNutrientReport As ComboBoxX
        Public WithEvents tAmount As DecimalTextBox, tComments As RichTextBoxX, chkCompleted, chkAllMealsConsumed As CheckBoxX
        Public WithEvents btnDelete, btnNext, btnPrev, btnCal, btnShiftUp, btnShiftDown As ButtonX
        Public WithEvents btnDVReport, btnDRIReport, btnEARReport, btnULReport, btnFoodBreakdown As ButtonX
        Public WithEvents btnNutrientAverage, btnLastXDays As ButtonX
        Public AskDeletion As Boolean = False, CLIndex As Integer = -1
        Public CLog As LogEntry, CFE As FoodEntry, CFood As FoodItem, FoodEntries As List(Of FoodEntry)
        Private CFEIndex As Integer = -1, LastLog As LogEntry

        Public Sub New(ByVal tPage As TabBrowser.Page)
            Try
                Page = tPage
                LoadControls()
            Catch ex As Exception
                MsgBox("Error loading FoodLog page: " & ex.Message)
            End Try
        End Sub

        Private Sub LoadControls()
            Dim lblTemp As LabelX = Nothing, lblW As Integer = 66
            CM = New ContextMenuStrip()
            lblTemp = New LabelX(Page, "", "Food Entries:", 5, 5, -1, -1)
            lstFoods = New ListItems(Page, "lstFoods", 5, lblTemp.Bottom + 5, 150, 280)
            lstFoods.MultiSelect = True : lstFoods.DeselectOnEmptySpaceClick = True
            lstFoods.BorderStyle = BorderStyle.FixedSingle
            btnShiftUp = New ButtonX(Page, "btnShiftUp", "Shift Up", lstFoods.Left, lstFoods.Bottom, CInt(lstFoods.Width / 2) - 3, 22)
            btnShiftDown = New ButtonX(Page, "btnShiftDown", "Shift Down", btnShiftUp.Left, btnShiftUp.Bottom + 5, btnShiftUp.Width, btnShiftUp.Height)
            btnDelete = New ButtonX(Page, "btnDelete", "Delete", btnShiftUp.Right + 5, btnShiftUp.Top, btnShiftUp.Width, btnShiftUp.Height)

            btnPrev = New ButtonX(Page, "btnPrev", "<<", lstFoods.Right + 5, 5, 25, 20)
            btnCal = New ButtonX(Page, "btnCal", "", btnPrev.Right + 1, btnPrev.Top, btnPrev.Width, btnPrev.Height)
            btnNext = New ButtonX(Page, "btnNext", ">>", Page.Width - btnPrev.Width - 5, btnPrev.Top, btnPrev.Width, btnPrev.Height)
            lblLogDate = New LabelX(Page, "", "", btnCal.Right + 1, btnPrev.Top, btnNext.Left - btnCal.Right - 2, 20)
            lblLogDate.TextAlign = ContentAlignment.MiddleCenter : lblLogDate.AutoSize = False : lblLogDate.BorderStyle = BorderStyle.Fixed3D
            lblTemp = New LabelX(Page, "", "Add Recipe:", lstFoods.Right + 8, lblLogDate.Bottom + 5, lblW, 22)
            cmbRecipe = New ComboBoxX(Page, "cmbRecipe", lblTemp.Right + 5, lblTemp.Top - 2, 450, 22) : cmbRecipe.MaxDropDownItems = 25
            lblTemp = New LabelX(Page, "", "Add Food:", lblTemp.Left, lblTemp.Bottom + 5, lblW, 22)
            cmbFood = New ComboBoxX(Page, "cmbFood", lblTemp.Right + 5, lblTemp.Top - 2, 450, 22) : cmbFood.MaxDropDownItems = 25

            lblTemp = New LabelX(Page, "", "Food:", lstFoods.Right + 8, cmbFood.Bottom + 20, lblW, 22) : lblTemp.TextAlign = ContentAlignment.MiddleRight
            lblFood = New LabelX(Page, "lblFood", "", lblTemp.Right + 5, lblTemp.Top, 150, 42)
            lblFood.TextAlign = ContentAlignment.MiddleCenter : lblFood.AutoSize = False : lblFood.BorderStyle = BorderStyle.Fixed3D
            lblTemp = New LabelX(Page, "", "Serving:", lblTemp.Left, lblFood.Bottom + 10, lblW, 22) : lblTemp.TextAlign = ContentAlignment.MiddleRight
            cmbServing = New ComboBoxX(Page, "cmbServing", lblTemp.Right + 5, lblTemp.Top, 150, 22)
            lblTemp = New LabelX(Page, "", "Amount:", lblTemp.Left, cmbServing.Bottom + 10, lblW, 22) : lblTemp.TextAlign = ContentAlignment.MiddleRight
            tAmount = New DecimalTextBox(Page, "tAmount", "", lblTemp.Right + 5, lblTemp.Top - 2, 150, 22)

            chkCompleted = New CheckBoxX(Page, "chkCompleted", "Completed", lblTemp.Left, lstFoods.Bottom - 82, -1, -1)
            chkAllMealsConsumed = New CheckBoxX(Page, "chkAllMealsConsumed", "All Meals Consumed", chkCompleted.Right + 10, chkCompleted.Top, -1, -1)
            lblTemp = New LabelX(Page, "", "Comments:", lblTemp.Left, lstFoods.Bottom - 60, lblW, 22) : lblTemp.TextAlign = ContentAlignment.MiddleRight
            tComments = New RichTextBoxX(Page, "tComments", "", lblTemp.Right + 5, lblTemp.Top - 2, 550, 80)
            btnDVReport = New ButtonX(Page, "btnBasicReport", "DV Report", lblFood.Right + 55, lblFood.Top, 100, 20)
            btnDRIReport = New ButtonX(Page, "btnDRIReport", "DRI Report", btnDVReport.Left, btnDVReport.Bottom + 10, 100, 20)
            btnEARReport = New ButtonX(Page, "btnEARReport", "EAR Report", btnDVReport.Left, btnDRIReport.Bottom + 10, 100, 20)
            btnULReport = New ButtonX(Page, "btnULReport", "UL Report", btnDVReport.Left, btnEARReport.Bottom + 10, 100, 20)
            lblTemp = New LabelX(Page, "", "Nutrient Report:", btnDVReport.Right + 10, btnDVReport.Top, -1, -1)
            cmbNutrientReport = New ComboBoxX(Page, "cmbNutrientReport", lblTemp.Right + 2, lblTemp.Top, 140, 22)
            cmbNutrientReport.MaxDropDownItems = 20
            btnNutrientAverage = New ButtonX(Page, "btnAllTimeAverage", "Nutrient Averages", lblTemp.Left, cmbNutrientReport.Bottom + 5, btnDVReport.Width, btnDVReport.Height)
            btnLastXDays = New ButtonX(Page, "btnAllTimeAverage", "Last X Days", lblTemp.Left, btnNutrientAverage.Bottom + 5, btnDVReport.Width, btnDVReport.Height)
            btnFoodBreakdown = New ButtonX(Page, "btnFoodBreakdown", "Food Breakdown", btnLastXDays.Left, btnLastXDays.Bottom + 10, 100, 20)
            For Each n As NutrientProperty In Nutrients
                If n.Basic Then cmbNutrientReport.Items.Add(n)
            Next
        End Sub

        Public Sub InitialLoad()
            Dim iDate As Integer = GetDataNumber()
            If Not _Logs.ContainsKey(iDate) Then Logs.Add(New LogEntry(GetDataNumber()))
            CLIndex = Logs.Count - 1
            LastLog = New LogEntry(11111111)
            LastLog.Comments = "TEST"
            Logs.Add(LastLog)
            If Logs.Count = 2 Then btnPrev.Enabled = False
            LoadLogEntry(Logs(CLIndex))
        End Sub

        Public Overrides Sub Save()
            If CLog Is Nothing Or CLog Is LastLog Then Exit Sub
            SaveCFE()
            CLog.FoodEntries = FoodEntries
            CLog.Comments = tComments.Text
            CLog.Completed = chkCompleted.Checked
            CLog.AllMealsConsumed = chkAllMealsConsumed.Checked
        End Sub

        Private Sub ResetServings()
            cmbServing.Items.Clear()
            cmbServing.Items.Add("Grams")
            cmbServing.SelectedIndex = 0
        End Sub

        Public Sub LoadLogEntry(ByVal tLog As LogEntry)
            Save()
            CFEIndex = -1
            If CLIndex > 0 Then btnPrev.Enabled = True Else btnPrev.Enabled = False
            If CLIndex < Logs.Count - 1 Then btnNext.Enabled = True Else btnNext.Enabled = False
            CLog = tLog
            CFE = Nothing
            CFood = Nothing
            tAmount.Focus()
            tComments.Text = CLog.Comments
            chkCompleted.Checked = CLog.Completed
            chkAllMealsConsumed.Checked = CLog.AllMealsConsumed
            FoodEntries = CLog.FoodEntries
            If FoodEntries.Count > 0 Then
                lstFoods.Items.AddRange(FoodEntries, True)
                lstFoods.SelectedIndex = FoodEntries.Count - 1
            Else
                lstFoods.Items.Clear()
                CFood = Nothing
                CFE = Nothing
                tAmount.Text = ""
                lblFood.Text = ""
                ResetServings()
            End If
            tAmount.SelectionStart = tAmount.Text.Length
            If CLog Is LastLog Then
                lblLogDate.Text = "TEST"
                tComments.Enabled = False
            Else
                tComments.Enabled = True
                Dim tDateString As String = CLog.EntryDate.ToString()
                Dim tDate As New Date(CInt(tDateString.Substring(0, 4)), CInt(tDateString.Substring(4, 2)), CInt(tDateString.Substring(6, 2)))
                lblLogDate.Text = tDate.ToLongDateString
            End If
        End Sub

        Public Sub UpdateItems()
            cmbRecipe.Items.Clear()
            For i As Integer = 0 To Recipies.Count - 1
                cmbRecipe.Items.Add(Recipies(i))
            Next
            cmbFood.Items.Clear()
            For i As Integer = 0 To Foods.Count - 1
                cmbFood.Items.Add(Foods(i))
            Next
        End Sub

        Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click
            If CFE Is Nothing Then Exit Sub
            If AskDeletion AndAlso MessageBox.Show("Remove " & CFood.ToString() & " from the list?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
            Dim tIndex As Integer = CFEIndex
            If lstFoods.Items.Count = 1 Then
                CFE = Nothing : CFood = Nothing
                lstFoods.Items.RemoveAt(CFEIndex)
            Else
                CFE = Nothing
                If CFEIndex = 0 Then
                    lstFoods.Items.RemoveAt(0)
                    lstFoods.SelectedIndex = 0
                Else
                    lstFoods.Items.RemoveAt(tIndex)
                    lstFoods.SelectedIndex = tIndex - 1
                End If
            End If
            CFEIndex = lstFoods.SelectedIndex
            FoodEntries.RemoveAt(tIndex)
            Changed = True
        End Sub

        Private Sub SaveCFE()
            If CFE Is Nothing Or CFEIndex = -1 Then Exit Sub
            If tAmount.Text = "" Then tAmount.Text = "0"
            CFE.Amount = tAmount.GetSingle()
            CFE.ServingSize = cmbServing.SelectedIndex
            lstFoods.Refresh()
        End Sub

        Private Sub LoadFE(ByVal tFE As FoodEntry)
            SaveCFE()
            CFE = tFE
            CFood = tFE.Food
            lblFood.Text = CFood.ToString()
            cmbServing.Items.Clear()
            cmbServing.Items.AddRange(CFood.ServingSizes.ToArray)
            cmbServing.SelectedIndex = CFE.ServingSize
            tAmount.Text = CFE.Amount.ToString()
        End Sub

        Private Sub tAmount_LostFocus(sender As Object, e As System.EventArgs) Handles tAmount.LostFocus
            SaveCFE()
        End Sub

        Private Sub lstFoods_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lstFoods.MouseClick
            If e.Button = MouseButtons.Right Then
                CM.Items.Clear()
                If lstFoods.SelectedIndex <> -1 Then
                    If lstFoods.SelectedIndices.Count = 1 Then CM.Items.Add("View DRI report for this entry")
                    CM.Items.Add("View food profile")
                    CM.Items.Add("Compare with other foods")
                    CM.Items.Add("Replace all instances with...")
                Else
                    CM.Items.Add("Compare foods")
                End If
                CM.Show(New Point(Cursor.Position.X, Cursor.Position.Y))
            End If
        End Sub

        Private Sub CM_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles CM.ItemClicked
            If e.ClickedItem.Text = "View DRI report for this entry" Then
                DailyReport(0, True)
            ElseIf e.ClickedItem.Text = "Compare with other foods" Then
                Dim lst As New List(Of FoodEntry)
                For Each intz As Integer In lstFoods.SelectedIndices
                    lst.Add(DirectCast(lstFoods.Items(intz), FoodEntry))
                Next
                CompareFoods(lst)
            ElseIf e.ClickedItem.Text = "Compare foods" Then
                CompareFoods(New List(Of FoodEntry))
            ElseIf e.ClickedItem.Text = "View food profile" Then
                Dim lst As SearchListItems = F.tabFood.lstFood
                lst.Index = lst.lstItems.Items.FindIndex(Function(x) DirectCast(x, ObjectBox).Obj Is CFE.Food)
                F.TabMain.SelectedPage = F.tabFood.Page
            ElseIf e.ClickedItem.Text = "Replace all instances with..." Then
                ReplaceAllInstances()
            End If
        End Sub

        Private Sub ReplaceAllInstances()
            Dim PB As New PromptBox("Food to replace " & CFood.Name, frmMain)
            Dim lblTemp As New LabelX(PB, "", "Replacement Food:", 15, 15, -1, -1)
            Dim cmbFood As New ComboBoxX(PB, "", lblTemp.Right + 4, lblTemp.Top, 220, 22)
            cmbFood.Items.AddRange(Foods.ToArray)
            cmbFood.MaxDropDownItems = 25
            If PB.ShowPromptBox() = PromptBox.Result.Cancel Then Exit Sub
            If cmbFood.SelectedIndex = -1 Then Exit Sub
            Dim NewFood As FoodItem = DirectCast(cmbFood.SelectedItem, FoodItem)
            If NewFood Is CFood Then Exit Sub
            Dim Counter As Integer = 0
            'Replace all foods with NewFood
            For Each lg As LogEntry In Logs
                Dim FEs As List(Of FoodEntry) = lg.FoodEntries.FindAll(Function(x) x.Food Is CFood)
                For Each fe As FoodEntry In FEs
                    fe.Food = NewFood
                    Counter += 1
                Next
            Next
            MsgBox("Replaced " & Counter & " food entries.")
        End Sub

        Private Sub _CompareFoods_cmbRecipies_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim cmb As ComboBoxX = DirectCast(sender, ComboBoxX)
            Dim FC As FoodComponents = DirectCast(cmb.Tag, FoodComponents)
            Dim rp As Recipe = DirectCast(cmb.SelectedItem, Recipe)
            FC.AddRecipe(rp)
        End Sub

        Private Sub CompareFoods(ByVal FoodList As List(Of FoodEntry))
            Dim PB As New PromptBox("Food Comparison", frmMain)
            Dim cmbRecipies(1) As ComboBoxX
            Dim FCs(1) As FoodComponents
            Dim radPositives(1) As RadioButtonX, TTop As Integer = 15
            For i As Integer = 0 To 1
                Dim lblTemp As New LabelX(PB, "", "Recipe: ", 15, TTop, -1, -1)
                cmbRecipies(i) = New ComboBoxX(PB, "", lblTemp.Right + 3, TTop, 200, 22)
                FCs(i) = New FoodComponents(PB, "", 15, cmbRecipies(i).Bottom + 5, 700, 100, 200)
                radPositives(i) = New RadioButtonX(PB, "", "Positive", FCs(i).Right + 5, FCs(i).Top, -1, -1, True)
                TTop = FCs(i).Bottom + 10
                cmbRecipies(i).Items.AddRange(Recipies.ToArray())
                cmbRecipies(i).Tag = FCs(i)
                AddHandler cmbRecipies(i).SelectedIndexChanged, AddressOf _CompareFoods_cmbRecipies_SelectedIndexChanged
            Next
            FCs(0).SetRecipes(FoodList)
            FCs(1).Lines(0).cmbFood.SelectedIndex = Recipies.Count
            FCs(1).Lines(0).cmbServingSize.SelectedIndex = 0
            FCs(1).Lines(0).tAmount.Text = ""
            Dim L0 As FoodComponents.FoodLine = FCs(0).Lines(0)
            Dim indx As Integer = L0.cmbServingSize.SelectedIndex
            If indx = -1 Then

            ElseIf indx = 0 Then
                FCs(1).Lines(0).tAmount.Text = L0.tAmount.Text
            Else
                FCs(1).Lines(0).tAmount.Text = (FoodList(0).Amount * DirectCast(L0.cmbFood.SelectedItem, FoodItem).ServingSizes(FoodList(0).ServingSize).Amount).ToString()
            End If
            If PB.ShowPromptBox() = PromptBox.Result.Cancel Then Exit Sub
            Dim Entriez(1) As List(Of FoodEntry)
            For i As Integer = 0 To 1
                Dim FC As FoodComponents = FCs(i)
                Dim lst As List(Of FoodEntry) = (From ln As FoodComponents.FoodLine In FC.Lines Where ln.tAmount.Text <> "" Select New FoodEntry(ln.CFood, ln.cmbServingSize.SelectedIndex, CSng(ln.tAmount.Text))).ToList()
                Entriez(i) = lst
            Next

            Dim lstSites As New List(Of Integer) From {2, 0}, LongestName As Integer = 0
            Dim nList(1) As Dictionary(Of Integer, Single)
            Dim UsedList(1) As Dictionary(Of Integer, Boolean)
            Dim MissingList(1) As Dictionary(Of Integer, List(Of FoodItem))
            For i As Integer = 0 To 1
                nList(i) = New Dictionary(Of Integer, Single)
                UsedList(i) = New Dictionary(Of Integer, Boolean)
                MissingList(i) = New Dictionary(Of Integer, List(Of FoodItem))
            Next
            For Each n As NutrientProperty In Nutrients
                If n.Basic Then
                    For i As Integer = 0 To 1
                        nList(i).Add(n.ID, 0)
                        UsedList(i).Add(n.ID, False)
                        MissingList(i).Add(n.ID, New List(Of FoodItem))
                    Next
                    LongestName = Math.Max(n.Name.Length, LongestName)
                End If
            Next

            For i As Integer = 0 To 1
                For Each fe As FoodEntry In Entriez(i)
                    Dim tFood As FoodItem = fe.Food
                    Dim tAmounts As Single = 0
                    If fe.ServingSize = 0 Then
                        tAmounts = fe.Amount / 100
                    Else
                        tAmounts = tFood.ServingSizes(fe.ServingSize).Amount * fe.Amount / 100
                    End If
                    Do
                        For i2 As Integer = 0 To lstSites.Count - 1
                            If tFood.DataSites.ContainsKey(lstSites(i2)) Then
                                Dim props As List(Of FoodProperty) = tFood.DataSites(lstSites(i2)).Properties
                                For i3 As Integer = 0 To props.Count - 1
                                    If nList(i).ContainsKey(props(i3).Nutrient) Then
                                        If UsedList(i)(props(i3).Nutrient) = True Then Continue For
                                        nList(i)(props(i3).Nutrient) += tAmounts * props(i3).Amount
                                        UsedList(i)(props(i3).Nutrient) = True
                                    End If
                                Next
                            End If
                        Next
                        If tFood.Parent Is Nothing Then Exit Do
                        tFood = tFood.Parent
                    Loop
                    If Entriez(i).Count > 1 Then
                        For Each n As NutrientProperty In Nutrients
                            If n.Basic Then UsedList(i)(n.ID) = False
                        Next
                    End If
                Next
            Next
            Dim sCategory As Item = Nothing, Lines As New List(Of String()), TopPositive As Boolean = radPositives(0).Checked
            Lines.Add({"Property", "Amount", "% DRI", "Missing"})
            For Each kvp As KeyValuePair(Of Integer, Single) In nList(0)
                Dim sPart1, sPart2, sPart3, sPart4 As String
                Dim n As NutrientProperty = DirectCast(Nutrients.ByID(kvp.Key), NutrientProperty)
                If n.Category IsNot sCategory Then
                    sCategory = n.Category
                    If Lines.Count <> 0 Then Lines.Add({"", "", "", ""})
                    Lines.Add({sCategory.Name, "", "", ""})
                End If
                If n.Parent = -1 Then
                    sPart1 = n.Name
                Else
                    If nList(0).ContainsKey(n.Parent) Then sPart1 = "-" & n.Name Else sPart1 = n.Name
                End If
                Dim TVal2 As Single = nList(1)(kvp.Key)
                Dim TDif As Single = kvp.Value - TVal2
                If Not TopPositive Then TDif *= -1
                Dim ND As NutrientDataPair = n.NutrientData(NutrientDataPair.Fields.DRI)
                If TDif <= 0 Then
                    sPart2 = Math.Round(TDif, 2).ToString() & " " & ND.Unit.Abbrev(0)
                Else
                    sPart2 = "+" & Math.Round(TDif, 2).ToString() & " " & ND.Unit.Abbrev(0)
                End If
                If ND.Value = -1 Then
                    sPart3 = ""
                Else
                    Dim tValue2 As Double = Math.Round(GetAdjustedValue(n, ND.Unit, TDif), 2), sExt As String = ""
                    If TDif > 0 Then sExt = "+"
                    sPart3 = sExt & Math.Round(Math.Round(tValue2 / CDbl(ND.Value), 2) * 100, 2).ToString() & "%"
                End If
                Dim val As List(Of FoodItem) = MissingList(0)(n.ID)
                If val.Count > 0 Then sPart4 = val.Count.ToString() Else sPart4 = ""
                Lines.Add({sPart1, sPart2, sPart3, sPart4})
            Next
            Dim paramz As New MessageBox2.MsgParams()
            paramz.Format = "{0," & "-" & (LongestName + 2).ToString() & "}{1,-13}{2,-10}{3,-8}"
            paramz.CenterForm = frmMain : paramz.HighlightPercents = True : paramz.PositiveVsNegative = True
            MessageBox2.Show(Lines, paramz)
        End Sub

        Private Sub CompareFoods_CMBFoodChanged(ByVal sender As Object, ByVal e As EventArgs)
            Dim cmbF As ComboBoxX = DirectCast(sender, ComboBoxX)
            Dim cmbS As ComboBoxX = DirectCast(cmbF.Tag, ComboBoxX)
            Dim TempFood As FoodItem = DirectCast(cmbF.SelectedItem, FoodItem)
            cmbS.Items.Clear()
            cmbS.Items.Add("Grams")
            For i As Integer = 1 To TempFood.ServingSizes.Count - 1
                cmbS.Items.Add(TempFood.ServingSizes(i))
            Next
            cmbS.SelectedIndex = 0
        End Sub

        Private Sub SetEnabled(ByVal Enabled As Boolean)
            cmbServing.Enabled = Enabled
            tAmount.Enabled = Enabled
            If Not Enabled Then
                CFE = Nothing : CFood = Nothing : CFEIndex = -1
                tAmount.Text = ""
                cmbServing.Items.Clear()
                lblFood.Text = ""
            End If
        End Sub

        Private Sub lstFoods_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles lstFoods.SelectedIndexChanged
            If lstFoods.SelectedIndices.Count > 1 Then
                SetEnabled(False)
            ElseIf lstFoods.SelectedIndex = -1 Then
                SetEnabled(False)
            Else
                SetEnabled(True)
                If CFEIndex = lstFoods.SelectedIndex Then Exit Sub
                If lstFoods.SelectedIndex = -1 Then
                    lblFood.Text = ""
                    ResetServings()
                    tAmount.Text = ""
                    CFEIndex = -1
                Else
                    LoadFE(FoodEntries(lstFoods.SelectedIndex))
                End If
                CFEIndex = lstFoods.SelectedIndex
            End If
        End Sub

        Private Sub cmbFood_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbFood.SelectedIndexChanged
            If cmbFood.SelectedIndex = -1 Then Exit Sub
            Dim tFood As FoodItem = DirectCast(cmbFood.SelectedItem, FoodItem)
            Dim FE As New FoodEntry(tFood, 0, 0), indx As Integer = lstFoods.SelectedIndex : If indx = -1 Then indx = FoodEntries.Count Else indx += 1
            lstFoods.Items.Insert(indx, FE)
            FoodEntries.Insert(indx, FE)
            SaveCFE()
            lstFoods.SelectedIndices.Clear()
            lstFoods.SelectedIndex = indx
            tAmount.Focus()
            tAmount.SelectionStart = 1
            cmbFood.SelectedIndex = -1
        End Sub

        Private Sub cmbRecipe_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbRecipe.SelectedIndexChanged
            If cmbRecipe.SelectedIndex = -1 Then Exit Sub
            Dim tRecipe As Recipe = DirectCast(Recipies(cmbRecipe.SelectedIndex), Recipe), indx As Integer = lstFoods.SelectedIndex
            If indx = -1 Then indx = FoodEntries.Count Else indx += 1
            If tRecipe.Ratios Then
                Dim sInput As String = InputBox("Enter total weight:", "Weight", "")
                If sInput = "" OrElse Not IsNumeric(sInput) Then Exit Sub
                Dim sAmount As Single = CSng(sInput)
                For i As Integer = tRecipe.Entries.Count - 1 To 0 Step -1
                    Dim FE As New FoodEntry(tRecipe.Entries(i).Food, tRecipe.Entries(i).ServingSize, tRecipe.Entries(i).Amount * sAmount)
                    lstFoods.Items.Insert(indx, FE)
                    FoodEntries.Insert(indx, FE)
                Next
            Else
                For i As Integer = tRecipe.Entries.Count - 1 To 0 Step -1
                    Dim FE As New FoodEntry(tRecipe.Entries(i).Food, tRecipe.Entries(i).ServingSize, tRecipe.Entries(i).Amount)
                    lstFoods.Items.Insert(indx, FE)
                    FoodEntries.Insert(indx, FE)
                Next
            End If
            lstFoods.SelectedIndex = indx + tRecipe.Entries.Count - 1
            cmbRecipe.SelectedIndex = -1
        End Sub

        Private Sub btnPrev_Click(sender As Object, e As System.EventArgs) Handles btnPrev.Click
            If CLIndex <> 0 Then
                CLIndex -= 1
                LoadLogEntry(Logs(CLIndex))
            End If
        End Sub

        Private Sub btnNext_Click(sender As Object, e As System.EventArgs) Handles btnNext.Click
            If CLIndex <> Logs.Count - 1 Then
                CLIndex += 1
                LoadLogEntry(Logs(CLIndex))
            End If
        End Sub

        Private Sub btnCal_Click(sender As Object, e As System.EventArgs) Handles btnCal.Click
            Dim sInput As String = InputBox("Enter Date:", "Log Date", "")
            If sInput = "" Then Exit Sub
            Dim sParts() As String = sInput.Split(Chr(45))
            If sParts.Length <> 3 Then
                MsgBox("Error: Invalid date entered (" & sInput & ")" & Chr(10) & Chr(10) & "Format: YYYY-MM-DD")
            Else
                If sParts(1).Length = 1 Then sParts(1) = "0" & sParts(1)
                If sParts(2).Length = 1 Then sParts(2) = "0" & sParts(2)
                Dim iDate As Integer = CInt(sParts(0) & sParts(1) & sParts(2))
                If _Logs.ContainsKey(iDate) Then
                    Dim logo As LogEntry = _Logs(iDate)
                    CLIndex = Logs.IndexOf(logo)
                    LoadLogEntry(logo)
                Else
                    MsgBox("Error: No log entry for the entered date (" & sInput & ")")
                End If
            End If
        End Sub

        Private Sub btnShiftUp_Click(sender As Object, e As System.EventArgs) Handles btnShiftUp.Click
            If lstFoods.SelectedIndex < 1 Then Exit Sub
            Dim indx As Integer = lstFoods.SelectedIndex, fe As FoodEntry = CFE : CFE = Nothing
            FoodEntries.RemoveAt(indx)
            FoodEntries.Insert(indx - 1, fe)
            lstFoods.Items.AddRange(FoodEntries, True)
            lstFoods.SelectedIndex = indx - 1
        End Sub

        Private Sub btnShiftDown_Click(sender As Object, e As System.EventArgs) Handles btnShiftDown.Click
            If lstFoods.SelectedIndex = FoodEntries.Count - 1 Then Exit Sub
            Dim indx As Integer = lstFoods.SelectedIndex, fe As FoodEntry = CFE : CFE = Nothing
            FoodEntries.RemoveAt(indx)
            FoodEntries.Insert(indx + 1, fe)
            lstFoods.Items.AddRange(FoodEntries, True)
            lstFoods.SelectedIndex = indx + 1
        End Sub

        Private Sub DailyReport(ByVal mode As Byte, Optional ByVal SingleFood As Boolean = False)
            SaveCFE()
            Dim sw As Stopwatch = Stopwatch.StartNew()
            Dim Reportz As New ReportX(True), EntriezList As List(Of FoodEntry) = Nothing
            Dim paramz As New MessageBox2.MsgParams() : paramz.Title = "Nutrient Report"
            If lstFoods.SelectedIndices.Count < 2 Then
                If SingleFood Then EntriezList = New List(Of FoodEntry) From {CFE}
            Else
                EntriezList = (From intz As Integer In lstFoods.SelectedIndices Select FoodEntries(intz)).ToList()
            End If
            If mode = 3 Then paramz.InvertHighlights = True
            Reportz.Report(New List(Of LogEntry) From {CLog}, EntriezList)

            Dim sCategory As Item = Nothing, Lines As New List(Of String())
            Lines.Add({"Property", "Amount", NutrientDataPair.GetName(mode) & " %", "Missing"})
            For Each kvp As KeyValuePair(Of Integer, ReportX.ParamsList) In Reportz.Params
                Dim PL As ReportX.ParamsList = kvp.Value
                Dim tValue As Single = PL.TotalAmount, sPart1, sPart2, sPart3, sPart4 As String
                Dim n As NutrientProperty = DirectCast(Nutrients.ByID(kvp.Key), NutrientProperty)
                If n.Category IsNot sCategory Then
                    sCategory = n.Category
                    If Lines.Count <> 0 Then Lines.Add({"", "", "", ""})
                    Lines.Add({sCategory.Name, "", "", ""})
                End If
                If n.Parent = -1 Then
                    If n.Name = "Fat" Then sPart1 = n.Name & " (" & Reportz.Omega6to3Ratio & " ratio)" Else sPart1 = n.Name
                Else
                    If Reportz.Params.ContainsKey(n.Parent) Then sPart1 = "-" & n.Name Else sPart1 = n.Name
                End If
                Dim nUnitPair As NutrientDataPair = n.NutrientData(mode)
                tValue = GetAdjustedValue(n, nUnitPair.Unit, tValue)
                sPart2 = Math.Round(tValue, 2).ToString() & " " & nUnitPair.Unit.Abbrev(0)
                If nUnitPair.Value <> -1 Then
                    sPart3 = (Math.Round(tValue / nUnitPair.Value * 100, 1)).ToString() & "%"
                Else
                    sPart3 = "" : sPart4 = ""
                End If
                If PL.MissingList.Count > 0 Then sPart4 = PL.MissingList.Count.ToString() Else sPart4 = ""
                Lines.Add({sPart1, sPart2, sPart3, sPart4})
            Next
            paramz.Format = "{0," & "-" & (Reportz.LongestName + 2).ToString() & "}{1,-13}{2,-10}{3,-8}"
            paramz.CenterForm = frmMain : paramz.HighlightPercents = True
            sw.Stop()
            MessageBox2.Show(Lines, paramz)
        End Sub

        Public Sub NutrientAverage(Optional ByVal Data As Object = Nothing)
            SaveCFE()
            Dim Reportz As New ReportX(True)
            Dim Logz As List(Of LogEntry) ', Reportz.longestname As Integer = 0
            Dim paramz As New MessageBox2.MsgParams() : paramz.Title = "Nutrient Report"
            If Data Is Nothing Then Logz = Logs Else Logz = DirectCast(Data, List(Of LogEntry))
            Reportz.Report(Logz, Nothing, True)

            Dim sCategory As Item = Nothing, Lines As New List(Of String())
            If Data Is Nothing Then Lines.Add({"Total Completed: " & Reportz.CompletedLogs.ToString() & ",  Total Unfinished: " & (Logs.Count - Reportz.CompletedLogs).ToString(), "", "", "", ""})
            Lines.Add({"Property", "Amount", "DRI %", "100+% DRI", "Missing"})
            For Each kvp As KeyValuePair(Of Integer, ReportX.ParamsList) In Reportz.Params
                Dim PL As ReportX.ParamsList = kvp.Value
                Dim tValue As Single = PL.TotalAmount, sPart1, sPart2, sPart3, sPart4, sPart5 As String
                Dim n As NutrientProperty = DirectCast(Nutrients.ByID(kvp.Key), NutrientProperty)
                If n.Category IsNot sCategory Then
                    sCategory = n.Category
                    If Lines.Count <> 0 Then Lines.Add({"", "", "", "", ""})
                    Lines.Add({sCategory.Name, "", "", "", ""})
                End If
                If n.Parent = -1 Then
                    If n.Name = "Fat" Then sPart1 = n.Name & " (" & Reportz.Omega6to3Ratio & " ratio)" Else sPart1 = n.Name
                Else
                    If Reportz.Params.ContainsKey(n.Parent) Then sPart1 = "-" & n.Name Else sPart1 = n.Name
                End If
                Dim nDRI As NutrientDataPair = n.NutrientData(NutrientDataPair.Fields.DRI)
                tValue = GetAdjustedValue(n, nDRI.Unit, tValue / Reportz.CompletedLogs)
                sPart2 = Math.Round(tValue, 2).ToString() & " " & nDRI.Unit.Abbrev(0)
                If nDRI.Value <> -1 Then
                    sPart3 = (Math.Round(tValue / nDRI.Value * 100, 1)).ToString() & "%"
                    sPart4 = Math.Round(PL.DRI100Days / Reportz.CompletedLogs * 100, 1).ToString() & "%"
                Else
                    sPart3 = "" : sPart4 = ""
                End If
                If PL.MissingList.Count > 0 Then sPart5 = PL.MissingList.Count.ToString() Else sPart5 = ""
                Lines.Add({sPart1, sPart2, sPart3, sPart4, sPart5})
            Next
            paramz.Format = "{0," & "-" & (Reportz.LongestName + 2).ToString() & "}{1,-13}{2,-10}{3,-11}{4,-8}"
            paramz.CenterForm = frmMain : paramz.HighlightPercents = True
            MessageBox2.Show(Lines, paramz)
        End Sub

        Public Function FormatDateString(ByVal sDate As String) As String
            Return sDate.Substring(0, 4) & "-" & sDate.Substring(4, 2) & "-" & sDate.Substring(6, 2)
        End Function

        Private Sub btnDVReport_Click(sender As Object, e As System.EventArgs) Handles btnDVReport.Click
            DailyReport(1)
        End Sub

        Private Sub btnDRIReport_Click(sender As Object, e As System.EventArgs) Handles btnDRIReport.Click
            DailyReport(0)
        End Sub

        Private Sub btnEARReport_Click(sender As Object, e As System.EventArgs) Handles btnEARReport.Click
            DailyReport(2)
        End Sub

        Private Sub btnULReport_Click(sender As Object, e As System.EventArgs) Handles btnULReport.Click
            DailyReport(3)
        End Sub

        Private Sub btnFoodBreakdown_Click(sender As Object, e As System.EventArgs) Handles btnFoodBreakdown.Click
            Dim FB As New Dictionary(Of Integer, List(Of Single)), totalweight As Single = 0, totalcalories As Single = 0
            Dim CalID As Integer = Nutrients.FindByName("Calories").ID
            For i As Integer = 0 To FCategories.Count - 1
                FB.Add(i, New List(Of Single) From {0, 0, 0})
            Next
            For Each fe As FoodEntry In FoodEntries
                Dim fd As FoodItem = fe.Food
                Dim lst As List(Of Single) = FB(fd.Category.ID)
                Dim tAmount As Single = 0
                If fe.ServingSize = 0 Then
                    tAmount = fe.Amount
                Else
                    tAmount = fd.ServingSizes(fe.ServingSize).Amount * fe.Amount
                End If
                Do
                    For i2 As Integer = 0 To 2
                        If fd.DataSites.ContainsKey(i2) Then
                            Dim props As List(Of FoodProperty) = fd.DataSites(i2).Properties
                            Dim indx As FoodProperty = props.Find(Function(x) x.Nutrient = CalID)
                            If indx IsNot Nothing Then
                                Dim tAmounto As Single = (indx.Amount * (tAmount / 100))
                                totalcalories += tAmounto
                                lst(2) += tAmounto
                                Exit Do
                            End If
                        End If
                    Next
                    If fd.Parent Is Nothing Then Exit Do
                    fd = fd.Parent
                Loop
                lst(0) += 1
                lst(1) += tAmount
                totalweight += tAmount
            Next      'String.Format("{0,-27}", String.Format("{0," + ((27 + s.Length) / 2).ToString() +  "}", s))
            Dim Lines As New List(Of String())
            Lines.Add({"Category  ", "Entries", "Grams", "Calories", "% Total Calories"})
            For i As Integer = 0 To FCategories.Count - 1
                Dim pf As List(Of Single) = FB(i)

                Dim defto As String() = {FCategories(i).Name & ": ", GetSingleString(pf(0)), GetSingleString(pf(1)), GetSingleString(pf(2)), GetSingleString(pf(2) / totalcalories * 100) & "%"}
                Lines.Add(defto)
            Next
            Lines.Add({"Total: ", FoodEntries.Count.ToString(), Math.Round(totalweight, 1).ToString(), Math.Round(totalcalories, 1).ToString(), "100"})
            Dim paramz As New MessageBox2.MsgParams()
            paramz.Title = "Nutrient Report"
            paramz.Format = "{0,18}{1,-10}{2,-8}{3,-10}{4,-17}"
            paramz.CenterForm = frmMain
            MessageBox2.Show(Lines, paramz)
        End Sub

        Private Function GetSingleString(ByVal s As Single) As String
            Return Math.Round(s, 1).ToString()
        End Function

        Private Sub cmbNutrientReport_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbNutrientReport.SelectedIndexChanged
            If cmbNutrientReport.SelectedIndex = -1 Then Exit Sub
            Dim n As NutrientProperty = DirectCast(cmbNutrientReport.SelectedItem, NutrientProperty)
            Dim Reportz As New ReportX(False)
            Reportz.Params.Add(n.ID, New ReportX.ParamsList)
            Reportz.Report(Logs, Nothing, True)
            Dim PL As ReportX.ParamsList = Reportz.Params(n.ID)
            Dim tValue As Double = Math.Round(GetAdjustedValue(n, n.NutrientData(0).Unit, PL.TotalAmount / Reportz.CompletedLogs), 2)
            Dim ND As NutrientDataPair = n.NutrientData(NutrientDataPair.Fields.DRI)
            Dim nDRI As Single = ND.Value
            If PL.Highest.Number = 0 Then PL.Highest = New StringAndNum("N/A", 0)
            If PL.Lowest.Number = 999999999 Then PL.Lowest = New StringAndNum("N/A", 0)
            Dim sb As New System.Text.StringBuilder(n.Name & Chr(10) & "Total: " & PL.TotalAmount.ToString() & " " & n.Unit.Abbrev(0) & Chr(10) & "Completed Entries: " & Reportz.CompletedLogs.ToString() & Chr(10))
            If nDRI <> -1 Then
                sb.Append("DRI: " & nDRI.ToString() & " " & ND.Unit.Abbrev(0) & Chr(10))
                sb.Append("-Total Days 100+%:  " & PL.DRI100Days & " days (" & Math.Round(PL.DRI100Days / Reportz.CompletedLogs * 100, 1).ToString() & "%)" & Chr(10))
                sb.Append("-Total Days 75-99%:  " & PL.DRI75Days & " days (" & Math.Round(PL.DRI75Days / Reportz.CompletedLogs * 100, 1).ToString() & "%)" & Chr(10))
                sb.Append("-Total Days < 75%:  " & PL.DRILowDays & " days (" & Math.Round(PL.DRILowDays / Reportz.CompletedLogs * 100, 1).ToString() & "%)" & Chr(10))
            End If
            sb.Append("Average: " & tValue.ToString() & " " & ND.Unit.Abbrev(0))
            Dim NHighestN1 As Single = GetAdjustedValue(n, n.NutrientData(NutrientDataPair.Fields.DRI).Unit, PL.Highest.Number)
            Dim NLowestN1 As Single = GetAdjustedValue(n, n.NutrientData(NutrientDataPair.Fields.DRI).Unit, PL.Lowest.Number)
            Dim NHighestN As String = Math.Round(NHighestN1, 1).ToString(), NLowestN As String = Math.Round(NLowestN1, 1).ToString()
            If nDRI <> -1 Then
                sb.Append(" (" & Math.Round(tValue / n.NutrientData(0).Value * 100, 2) & "% DRI)" & Chr(10))
                sb.Append("Highest Amount: " & NHighestN & " " & ND.Unit.Abbrev(0) & " (" & Math.Round(NHighestN1 / nDRI * 100, 2) & "% DRI)" & ", " & FormatDateString(PL.Highest.Text) & Chr(10))
                sb.Append("Lowest Amount: " & NLowestN & " " & ND.Unit.Abbrev(0) & " (" & Math.Round(NLowestN1 / nDRI * 100, 2) & "% DRI)" & ", " & FormatDateString(PL.Lowest.Text) & Chr(10))
            Else
                sb.Append(Chr(10) & "Highest Amount: " & NHighestN & " " & ND.Unit.Abbrev(0) & ", " & FormatDateString(PL.Highest.Text) & Chr(10))
                sb.Append("Lowest Amount: " & NLowestN & " " & ND.Unit.Abbrev(0) & ", " & FormatDateString(PL.Lowest.Text) & Chr(10))
            End If
            MsgBox(sb.ToString(), MsgBoxStyle.OkOnly, n.Name)
            cmbNutrientReport.SelectedIndex = -1
        End Sub

        Private Sub btnNutrientAverage_Click(sender As Object, e As System.EventArgs) Handles btnNutrientAverage.Click
            NutrientAverage(Nothing)
        End Sub

        Private Sub btnLastXDays_Click(sender As Object, e As System.EventArgs) Handles btnLastXDays.Click
            Dim sInput As String = InputBox("Nutrient report for the last X days:", "Nutrient Report", "7")
            If Not IsNumeric(sInput) OrElse sInput.Contains(".") OrElse sInput = "0" OrElse sInput.Contains("-") OrElse sInput.Contains(",") Then
                MsgBox("Error: Must enter a positive integer.") : Exit Sub
            End If
            Dim Days As Integer = Math.Min(Logs.Count - 1, CInt(sInput)), LCount As Integer = Logs.Count - 2
            Dim lstLogs As New List(Of LogEntry), Count As Integer = 0
            For i As Integer = LCount To 0 Step -1
                Dim lg As LogEntry = Logs(i)
                If lg.Completed And lg.AllMealsConsumed Then
                    lstLogs.Add(lg)
                    Count += 1
                    If Count >= Days Then Exit For
                End If
            Next
            NutrientAverage(lstLogs)
        End Sub

        Private Sub lblFood_DoubleClick(sender As Object, e As System.EventArgs) Handles lblFood.DoubleClick
            If CFood Is Nothing Then Exit Sub
            Dim FDID As Integer = CFood.ID, Largest As New StringAndNum("", 0), Smallest As New StringAndNum("", 9999999)
            Dim Total As New PointF(0, 0), MostEntries As New StringAndNum("", 0), MostTotalDay As New StringAndNum("", 0)
            Dim FirstTime As String = "", LastTime As String = ""
            For Each lg As LogEntry In Logs
                Dim lstFE As List(Of FoodEntry) = lg.FoodEntries.FindAll(Function(x) x.Food.ID = FDID)
                If lstFE.Count > MostEntries.Number Then MostEntries = New StringAndNum(lg.EntryDate.ToString(), lstFE.Count)
                Dim TempWeight As Single = 0
                For Each fe As FoodEntry In lstFE
                    If FirstTime = "" Then FirstTime = lg.EntryDate.ToString()
                    LastTime = lg.EntryDate.ToString()
                    Dim tAmount As Single = 0
                    If fe.ServingSize = 0 Then
                        tAmount = fe.Amount
                    Else
                        tAmount = fe.Amount * CFood.ServingSizes(fe.ServingSize).Amount
                    End If
                    If tAmount > Largest.Number Then Largest = New StringAndNum(lg.EntryDate.ToString(), tAmount)
                    If tAmount < Smallest.Number Then Smallest = New StringAndNum(lg.EntryDate.ToString(), tAmount)
                    TempWeight += tAmount
                Next
                If TempWeight > MostTotalDay.Number Then MostTotalDay = New StringAndNum(lstFE.Count.ToString() & ":" & lg.EntryDate.ToString(), TempWeight)
                Total = New PointF(Total.X + TempWeight, Total.Y + lstFE.Count)
            Next
            Dim sb As New StringBuilder(CFood.Name & Chr(10))
            sb.Append("Total Weight: " & Math.Round(Total.X, 1).ToString() & "g" & Chr(10))
            sb.Append("Average Weight: " & Math.Round(Total.X / Total.Y, 1).ToString() & "g" & Chr(10))
            sb.Append("Total Entries: " & Math.Round(Total.Y, 1).ToString() & Chr(10))
            sb.Append("Average Entries: " & Math.Round(Total.Y / Logs.Count, 2).ToString() & " per day" & Chr(10))
            sb.Append("Most Entries, Day: " & MostEntries.Number & ", " & FormatDateString(MostEntries.Text) & Chr(10))
            Dim indx As Integer = MostTotalDay.Text.IndexOf(":")
            If indx <> -1 Then
                sb.Append("Most total weight, Day: " & MostTotalDay.Number & "g, " & MostTotalDay.Text.Substring(0, indx) & " entries, " & FormatDateString(MostTotalDay.Text.Substring(indx + 1)) & Chr(10))
            End If
            sb.Append("Largest entry: " & Math.Round(Largest.Number, 1).ToString() & "g, " & FormatDateString(Largest.Text) & Chr(10))
            sb.Append("Smallest entry: " & Math.Round(Smallest.Number, 1).ToString() & "g, " & FormatDateString(Smallest.Text) & Chr(10))
            sb.Append("First entry: " & FormatDateString(FirstTime) & Chr(10))
            sb.Append("Last entry: " & FormatDateString(LastTime))
            MsgBox(sb.ToString())
        End Sub

    End Class

    Public Class RecipeEditorPage : Inherits FLPage
        Public WithEvents lstRecipies As SearchListItems, tName As TextBoxX, tComments As RichTextBoxX, FC As FoodComponents, chkRatios As CheckBoxX
        Public WithEvents btnAdd, btnDelete, btnCreateFood As ButtonX
        Public CRecipe As Recipe
        Private CIndex As Integer = -1

        Public Overrides ReadOnly Property FileName() As String
            Get
                Return "Recipes"
            End Get
        End Property

        Public Sub New(ByVal tPage As TabBrowser.Page)
            Try
                Page = tPage
                LoadControls()
                LoadRecipesList()
            Catch ex As Exception
                MsgBox("Error loading RecipeEditor page: " & ex.Message)
            End Try
        End Sub

        Private Sub LoadControls()
            Dim lblTemp As LabelX, w As Integer = 70
            lstRecipies = New SearchListItems(Page, "lstRecipes", "Recipe:", 5, 5, 150, 300)
            lstRecipies.txtSearch.BorderStyle = BorderStyle.FixedSingle
            btnAdd = New ButtonX(Page, "btnAdd", "Add", lstRecipies.Bounds.Left, lstRecipies.Bounds.Bottom + 5, 70, 22)
            btnDelete = New ButtonX(Page, "btnDelete", "Delete", btnAdd.Right + 10, btnAdd.Top, 70, 22)

            lblTemp = New LabelX(Page, "", "Name:", lstRecipies.Bounds.Right + 5, lstRecipies.Bounds.Top, w, 22)
            tName = New TextBoxX(Page, "tName", "", lblTemp.Right + 2, lblTemp.Top, 200, 22)
            lblTemp = New LabelX(Page, "", "Comments:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            tComments = New RichTextBoxX(Page, "tComments", "", lblTemp.Right + 2, lblTemp.Top, 400, 44)
            btnCreateFood = New ButtonX(Page, "btnCreateFood", "Create Food", tComments.Right + 5, tComments.Top, 100, 22)
            chkRatios = New CheckBoxX(Page, "chkRatios", "Set as ratios", tName.Right + 10, tName.Top, -1, -1, False)
            FC = New FoodComponents(Page, "FC", lblTemp.Left, tComments.Bottom + 5, 605, Page.Height - tComments.Bottom - 10)
        End Sub

        Public Sub LoadRecipesList()
            If Recipies.Count = 0 Then
                FC.Visible = False
                lstRecipies.Items.Clear()
            Else
                FC.Visible = True
                lstRecipies.Clear()
                lstRecipies.AddRange(Recipies)
                lstRecipies.lstItems.SelectedIndex = 0
            End If
        End Sub

        Protected Overrides Sub SaveList(sw As System.IO.StreamWriter)
            For i As Integer = 0 To Recipies.Count - 1
                sw.WriteLine(DirectCast(Recipies(i), Recipe).SaveLine())
            Next
        End Sub

        Private Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
            Dim sInput As String = AddItem(lstRecipies)
            If sInput = "" Then Exit Sub
            Dim rp As New Recipe(sInput, GetNextID), insindex As Integer, itms As ListItems.ItemList = lstRecipies.lstItems.Items
            insindex = Recipies.Insert(rp)
            itms.Insert(False, insindex, rp)
            lstRecipies.Index = -1
            lstRecipies.Index = insindex
            If Recipies.Count = 1 Then FC.Visible = True
        End Sub

        Private Sub lstRecipes_SelectedIndexChanged(sender As Object, e As ListItems.SelectedIndexChangedArgs) Handles lstRecipies.SelectedIndexChanged
            If lstRecipies.Index <> -1 Then
                LoadRecipe(DirectCast(lstRecipies.SItem, Recipe))
                CIndex = lstRecipies.Index
            End If
        End Sub

        Private Sub LoadRecipe(ByVal tRecipe As Recipe)
            Loading = True
            SaveRecipe()
            CRecipe = tRecipe
            tName.Text = CRecipe.Name
            tComments.Text = CRecipe.Comments
            chkRatios.Checked = CRecipe.Ratios
            FC.SetRecipes(CRecipe.Entries)
            Loading = False
        End Sub

        Public Overrides Sub Save()
            SaveRecipe()
            If FC.Changed Then Changed = True : FC.Changed = False
        End Sub

        Private Sub SaveRecipe()
            If CRecipe Is Nothing Then Exit Sub 'whole sub takes well less than 0 ms (200~ ticks) for a recipe with several food entries
            If Not tName.Text = CRecipe.Name Then
                CRecipe.Name = tName.Text.Trim()
                lstRecipies.Items(CIndex) = CRecipe.Name
            End If
            CRecipe.Comments = tComments.Text
            CRecipe.Entries.Clear()
            For i As Integer = 0 To FC.Lines.Count - 1
                With FC.Lines(i)
                    If .CFood Is Nothing Or .cmbServingSize.SelectedIndex = -1 Then Continue For
                    CRecipe.Entries.Add(New FoodEntry(.CFood, .cmbServingSize.SelectedIndex, .tAmount.GetSingle()))
                End With
            Next
            If CRecipe.Entries.Count = 0 Then 'remove
                Dim indx As Integer = Recipies.IndexOf(CRecipe)
                Recipies.RemoveAt(indx)
                CRecipe = Nothing
                lstRecipies.lstItems.Items.RemoveAt(CIndex)
            End If
            CRecipe.Ratios = chkRatios.Checked
        End Sub

        Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click
            If CRecipe Is Nothing Then Exit Sub
            If MessageBox.Show("Delete recipe " & Chr(34) & CRecipe.Name & Chr(34) & "?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
            Dim rp As Integer = Recipies.IndexOf(CRecipe)
            Recipies.RemoveAt(rp)
            lstRecipies.Remove(rp)
            If rp > 0 Then
                lstRecipies.lstItems.SelectedIndex = rp - 1
            ElseIf lstRecipies.Count > 0 Then
                lstRecipies.lstItems.SelectedIndex = 0
            Else
                FC.Clear()
                tName.Text = ""
                CRecipe = Nothing
            End If
            Changed = True
        End Sub

        Private Sub tName_TextChanged(sender As Object, e As System.EventArgs) Handles tName.TextChanged, tComments.TextChanged, chkRatios.CheckedChanged
            If Not Loading Then Changed = True
        End Sub

        Private Sub btnCreateFood_Click(sender As Object, e As System.EventArgs) Handles btnCreateFood.Click
            Dim sInput As String = InputBox("Enter name:", "Food Name", "")
            If sInput = "" Then Exit Sub
            Try
                Dim tf As New FoodItem(sInput, GetNextID())
                Dim SP As New SiteProfile(ComboSite, "")
                tf.DataSites.Add(ComboSite.ID, SP)
                Dim tAmount As Single = 0, lstSites As New List(Of Integer) From {3, 2, 0}, PropsList As List(Of FoodProperty) = SP.Properties
                For Each fe As FoodEntry In CRecipe.Entries
                    tAmount = fe.Amount
                    Dim tFood As FoodItem = fe.Food
                    Dim UsedList As New Dictionary(Of Integer, Boolean) ' = Nutrients.Cast(Of NutrientProperty)().ToDictionary(Function(n) n.ID, Function(n) False), note that this is really slow
                    For Each n As NutrientProperty In Nutrients
                        UsedList.Add(n.ID, False)
                    Next
                    Do
                        For i2 As Integer = 0 To lstSites.Count - 1
                            If tFood.DataSites.ContainsKey(lstSites(i2)) Then
                                Dim props As List(Of FoodProperty) = tFood.DataSites(lstSites(i2)).Properties
                                For i3 As Integer = 0 To props.Count - 1
                                    Dim ID As Integer = props(i3).Nutrient
                                    Dim tempAmount As Single = tAmount * props(i3).Amount
                                    If UsedList(ID) Then
                                        Continue For
                                    Else
                                        UsedList(ID) = True
                                        Dim prop As FoodProperty = PropsList.Find(Function(x) x.Nutrient = ID)
                                        If prop Is Nothing Then
                                            prop = New FoodProperty(ID, tempAmount) : PropsList.Add(prop)
                                        Else
                                            prop.Amount += tempAmount
                                        End If
                                    End If
                                Next
                            End If
                        Next
                        If tFood.Parent Is Nothing Then Exit Do
                        tFood = tFood.Parent
                    Loop
                Next
                tf.IsComboFood = True
                InsertFood(tf)
            Catch ex As Exception
                MsgBox("Error creating new food from recipe: " & ex.Message)
            End Try
        End Sub

    End Class

    Public Class FoodEditorPage : Inherits FLPage
        Public WithEvents lstFood As SearchListItems, btnAdd, btnAddVariation, btnDelete, btnRip As ButtonX, DF As DataField
        Public WithEvents tName, tSiteProfile, tComments As TextBoxX, cmbSiteProfiles, cmbCategory As ComboBoxX, CM As ContextMenuStrip
        Public CFood As FoodItem
        Private SiteIndex As Integer = 0

        Public Overrides ReadOnly Property FileName() As String
            Get
                Return "Foods"
            End Get
        End Property

        Public Sub New(ByVal tPage As TabBrowser.Page)
            Try
                Page = tPage
                LoadControls()
                LoadFoodList()
                lstFood.lstItems.SelectedIndex = 0
            Catch ex As Exception
                MsgBox("Error loading Food Editor page: " & ex.Message)
            End Try
        End Sub

        Private Sub LoadControls()
            Dim lblTemp As LabelX = Nothing, w As Integer = 70, w2 As Integer = 400
            CM = New ContextMenuStrip()
            lstFood = New SearchListItems(Page, "lstFood", "Food:", 5, 5, 150, 320)
            lstFood.txtSearch.BorderStyle = BorderStyle.FixedSingle
            AddHandler lstFood.lstItems.ItemClicked, AddressOf lstFood_ItemClicked
            btnAdd = New ButtonX(Page, "btnAdd", "Add", lstFood.Bounds.Left, lstFood.Bounds.Bottom + 5, 70, 22)
            btnDelete = New ButtonX(Page, "btnDelete", "Delete", btnAdd.Right + 10, btnAdd.Top, 70, 22)
            btnAddVariation = New ButtonX(Page, "btnAddSubFood", "Add Variation", btnAdd.Left, btnAdd.Bottom + 5, btnAdd.Width, btnAdd.Height)
            btnRip = New ButtonX(Page, "btnRip", "Rip Data", btnAddVariation.Right + 10, btnAddVariation.Top, btnAdd.Width, btnAdd.Height)

            lblTemp = New LabelX(Page, "", "Name:", lstFood.Bounds.Right + 5, lstFood.Bounds.Top, w, 22)
            tName = New TextBoxX(Page, "tName", "", lblTemp.Right + 2, lblTemp.Top, w2, 22)
            lblTemp = New LabelX(Page, "", "Category:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            cmbCategory = New ComboBoxX(Page, "", lblTemp.Right + 2, lblTemp.Top, 100, 22)
            cmbCategory.Items.AddRange(FCategories.ToArray) : cmbCategory.MaxDropDownItems = 20
            Dim lblTemp2 As New LabelX(Page, "", "Comments:", cmbCategory.Right + 5, lblTemp.Top, -1, -1)
            tComments = New TextBoxX(Page, "tComments", "", lblTemp2.Right + 2, lblTemp2.Top, tName.Right - lblTemp2.Right - 2, 22)
            lblTemp = New LabelX(Page, "", "Site Profiles:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            AddHandler lblTemp.DoubleClick, AddressOf lblSite_DoubleClick
            cmbSiteProfiles = New ComboBoxX(Page, "cmbSiteLabels", lblTemp.Right + 2, lblTemp.Top, 80, 22)
            For i As Integer = 0 To Sites.Count - 1
                cmbSiteProfiles.Items.Add(Sites(i).Name)
            Next
            tSiteProfile = New TextBoxX(Page, "tSiteLabel", "", cmbSiteProfiles.Right + 2, lblTemp.Top, tName.Right - cmbSiteProfiles.Right - 2, 22)
            DF = New DataField(Page, "DF", lblTemp.Left, tSiteProfile.Bottom + 5, Page.Width - lblTemp.Left - 5, btnRip.Bottom - tSiteProfile.Bottom - 5)
        End Sub

        Private Sub lstFood_ItemClicked(ByVal sender As Object, ByVal e As ListItems.ItemClickedArgs)
            If e.MEvent.Button = MouseButtons.Right Then
                Dim fd As FoodItem = DirectCast(Foods(e.NewIndex), FoodItem)
                If fd.Parent Is Nothing Then
                    CM.Items.Clear()
                    CM.Items.Add("Set Parent") : CM.Items(CM.Items.Count - 1).Tag = fd
                    CM.Show(Cursor.Position.X, Cursor.Position.Y)
                End If
            End If
        End Sub

        Private Sub CM_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles CM.ItemClicked
            If e.ClickedItem.Text = "Set Parent" Then
                Dim PB As New PromptBox("Set Parent Food", frmMain)
                Dim cmbFood As New ComboBoxX(PB, "", 15, 20, 300, 20)
                Dim lstFoods As New List(Of FoodItem), tfd As FoodItem = DirectCast(e.ClickedItem.Tag, FoodItem)
                For Each fd As FoodItem In Foods
                    If fd.Parent Is Nothing Then
                        If fd IsNot tfd Then
                            lstFoods.Add(fd)
                        End If
                    End If
                Next
                cmbFood.MaxDropDownItems = 25
                cmbFood.Items.AddRange(lstFoods.ToArray)
                If PB.ShowPromptBox = PromptBox.Result.Ok Then
                    If cmbFood.SelectedIndex = -1 Then Exit Sub
                    Dim fd As FoodItem = DirectCast(cmbFood.SelectedItem, FoodItem)
                    fd.Variations.Add(tfd.ID) : tfd.Parent = fd
                    Dim index As Integer = lstFood.Items.FindIndex(Function(x) DirectCast(x, ObjectBox).Obj Is fd)
                    lstFood.Items.Insert(index + 1, New ObjectBox("-" & tfd.Name, tfd))
                    lstFood.Index = index + 1
                    Foods.Insert(index + 1, tfd)
                    index = lstFood.Items.FindIndex(Function(x) DirectCast(x, ObjectBox).Obj Is tfd)
                    lstFood.Items.RemoveAt(index) : Foods.Remove(tfd)
                    Changed = True
                End If
            End If
        End Sub

        Private Sub lblSite_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)
            If tSiteProfile.Text = "" Then
                Dim sURL As String = "http://www.google.com/search?q=%22" & CFood.Name & "%22+%22" & "site:" & DirectCast(Sites(SiteIndex), DatabaseSite).SiteBase & "%22 "
                System.Diagnostics.Process.Start("C:\Program Files (x86)\Mozilla Firefox\firefox.exe", "-new-tab " & Chr(34) & sURL & Chr(34))
            Else
                System.Diagnostics.Process.Start("C:\Program Files (x86)\Mozilla Firefox\firefox.exe", "-new-tab " & Chr(34) & tSiteProfile.Text & Chr(34))
            End If
        End Sub

        Private Sub _SortFoods()
            Dim newList As New List(Of FoodItem)
            Dim deleg = Function(F1 As FoodItem, F2 As FoodItem)
                            Return F1.Name.CompareTo(F2.Name)
                        End Function
            For i As Integer = 0 To Foods.Count - 1
                Dim fo As FoodItem = DirectCast(Foods(i), FoodItem)
                If fo.Parent Is Nothing Then
                    newList.Add(fo)
                End If
            Next
            newList.Sort(deleg)
            Dim newList2 As New List(Of FoodItem)
            For i As Integer = 0 To newList.Count - 1
                Dim fo As FoodItem = DirectCast(newList(i), FoodItem)
                If fo.Parent Is Nothing Then
                    newList2.Add(fo)
                    If fo.Variations.Count > 0 Then
                        Dim newList3 As New List(Of FoodItem)
                        For Each variation As Integer In fo.Variations
                            newList3.Add(DirectCast(Foods.Find(Function(F As Item)
                                                                   Return F.ID = variation
                                                               End Function), FoodItem))
                        Next
                        newList3.Sort(deleg)
                        newList2.AddRange(newList3)
                    End If
                Else
                    MessageBox.Show("Fail")
                End If
            Next
            For Each fi As FoodItem In Foods
                If newList2.Contains(fi) = False Then
                    MessageBox.Show("Fail")
                End If
            Next
            Foods = ItemList.FromList(newList2)
        End Sub

        Public Sub LoadFoodList()
            Dim obList As New List(Of ObjectBox)
            For i As Integer = 0 To Foods.Count - 1
                Dim fo As FoodItem = DirectCast(Foods(i), FoodItem)
                If fo.IsComboFood Then
                    obList.Add(New ObjectBox(_ComboFoodStart & fo.Name, fo))
                Else
                    obList.Add(New ObjectBox(fo.Name, fo))
                    If fo.Variations.Count > 0 Then
                        For i2 As Integer = 0 To fo.Variations.Count - 1
                            Try
                                Dim fo2 As FoodItem = DirectCast(Foods.ByID(fo.Variations(i2)), FoodItem)
                                obList.Add(New ObjectBox("-" & fo2.Name, fo2))
                            Catch ex As Exception
                                MsgBox("Error loading variations for " & fo.Name & ": " & ex.Message)
                            End Try
                        Next
                        i += fo.Variations.Count
                    End If
                End If
            Next


            lstFood.AddRange(obList)
        End Sub

        Public Overrides Sub Save()
            If Not CFood Is Nothing Then SaveFoodItem(CFood)
            If DF.Changed Then Changed = True : DF.Changed = False
        End Sub

        Protected Overrides Sub SaveList(sw As System.IO.StreamWriter)
            For Each fd As FoodItem In Foods
                fd.SaveFood(sw)
            Next
        End Sub

        Private Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
            Dim sInput As String = AddItem(lstFood)
            If sInput = "" Then Exit Sub
            Dim fd As New FoodItem(sInput, GetNextID), insindex As Integer = -1, itms As ListItems.ItemList = lstFood.lstItems.Items
            For i As Integer = 0 To itms.Count - 1
                Dim ob As ObjectBox = DirectCast(itms(i), ObjectBox)
                If ob.Text.StartsWith("-") Then Continue For
                If sInput < ob.Text Then
                    insindex = i
                    Exit For
                End If
            Next
            If insindex = -1 Then insindex = itms.Count
            Foods.Insert(insindex, fd)
            itms.Insert(False, insindex, New ObjectBox(fd.Name, fd))
            If lstFood.Index = insindex Then lstFood.lstItems.SelectedIndex = -1
            lstFood.lstItems.SelectedIndex = insindex
        End Sub

        Private Sub btnAddSubFood_Click(sender As Object, e As System.EventArgs) Handles btnAddVariation.Click
            Dim sInput As String = AddItem(lstFood)
            If sInput = "" Then Exit Sub
            Dim fd As New FoodItem(sInput, GetNextID), indx As Integer = lstFood.Index, insindex As Integer = -1
            fd.Parent = CFood
            For i As Integer = 0 To CFood.Variations.Count - 1
                If fd.Name < Foods.ByID(CFood.Variations(i)).Name Then
                    insindex = i
                    Exit For
                End If
            Next
            If insindex = -1 Then insindex = CFood.Variations.Count
            CFood.Variations.Insert(insindex, fd.ID)
            Foods.Insert(indx + insindex + 1, fd)
            lstFood.lstItems.Items.Insert(False, indx + insindex + 1, New ObjectBox("-" & fd.Name, fd))
            lstFood.lstItems.SelectedIndex = indx + insindex + 1
            Changed = True
        End Sub

        Private Sub lstFood_SelectedIndexChanged(byvalsender As Object, e As EventArgs) Handles lstFood.SelectedIndexChanged
            If Not lstFood.Index = -1 Then LoadFoodItem(DirectCast(DirectCast(lstFood.SItem, ObjectBox).Obj, FoodItem))
        End Sub

        Private Sub LoadFoodItem(ByVal tFood As FoodItem)
            Loading = True
            If Not CFood Is Nothing Then SaveFoodItem(CFood)
            CFood = tFood
            DF.LoadFoodItem(tFood)
            tName.Text = CFood.Name
            tSiteProfile.Text = ""
            tComments.Text = tFood.Comments
            cmbCategory.SelectedIndex = FCategories.IndexOf(tFood.Category)
            cmbSiteProfiles.SelectedIndex = -1
            cmbSiteProfiles.SelectedIndex = 0
            If lstFood.lstItems.SelectedItem.ToString().StartsWith("-") Then btnAddVariation.Enabled = False Else btnAddVariation.Enabled = True
            Loading = False
        End Sub

        Public Sub SaveFoodItem(ByVal tFood As FoodItem)
            tFood.Name = tName.Text
            tFood.Comments = tComments.Text
            Dim tCategory As Item = DirectCast(cmbCategory.SelectedItem, Item)
            If tFood.Category IsNot tCategory Then
                tFood.Category = tCategory
                If tFood.Parent Is Nothing Then
                    For Each intz As Integer In tFood.Variations
                        DirectCast(Foods.ByID(intz), FoodItem).Category = tFood.Category
                    Next
                End If
            End If
            SaveSite()
            DF.Save()
        End Sub

        Private Sub SaveSite()
            If tSiteProfile.Text <> "" Then
                Dim id As Integer = Sites(SiteIndex).ID
                If CFood.DataSites.ContainsKey(id) Then
                    CFood.DataSites(id).Link = tSiteProfile.Text
                    Exit Sub
                End If
                CFood.DataSites.Add(id, New SiteProfile(DirectCast(Sites(SiteIndex), DatabaseSite), tSiteProfile.Text))
            End If
        End Sub

        Private Sub cmbSiteLabels_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbSiteProfiles.SelectedIndexChanged
            If cmbSiteProfiles.SelectedIndex = -1 Then Exit Sub
            SaveSite()
            If CFood.DataSites.ContainsKey(Sites(cmbSiteProfiles.SelectedIndex).ID) Then tSiteProfile.Text = CFood.DataSites(cmbSiteProfiles.SelectedIndex).Link Else tSiteProfile.Text = ""
            SiteIndex = cmbSiteProfiles.SelectedIndex
        End Sub

        Private Sub btnRip_Click(sender As Object, e As System.EventArgs) Handles btnRip.Click
            If CFood Is Nothing Then Exit Sub
            SaveSite()
            Dim bRipped As Boolean = False
            For i As Integer = 0 To 0 ' CFood.NProfiles.Count - 1
                If Not CFood.DataSites.ContainsKey(Sites(i).ID) Then Continue For
                If CFood.DataSites(i).Link <> "" Then
                    CFood.DataSites(i).Site.Rip(CFood, i)
                    bRipped = True
                End If
            Next
            If bRipped Then
                DF.LoadFoodItem(CFood)
                Changed = True
            Else
                MsgBox("No new data acquired.")
            End If
        End Sub

        Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click
            If CFood Is Nothing Then Exit Sub
            If MessageBox.Show("Delete food " & Chr(34) & CFood.Name & Chr(34) & "?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
            Dim fo As Integer = Foods.IndexOf(CFood)
            Foods.RemoveAt(fo)
            lstFood.Remove(fo)
            For i As Integer = 0 To CFood.Variations.Count - 1
                Dim tFood As FoodItem = DirectCast(Foods.ByID(fo), FoodItem)
                Foods.Remove(tFood)
                Dim obj As Object = lstFood.Items.Find(Function(x) DirectCast(x, ObjectBox).Obj Is tFood)
                lstFood.Remove(obj)
            Next
            If Not CFood.Parent Is Nothing Then CFood.Parent.Variations.Remove(CFood.ID)
            If fo > 0 Then
                lstFood.lstItems.SelectedIndex = fo - 1
            ElseIf lstFood.Count > 0 Then
                lstFood.lstItems.SelectedIndex = 0
            Else
                MsgBox("omg y u do that")
            End If
            Changed = True
        End Sub

        Private Sub StuffChanged(sender As Object, e As System.EventArgs) Handles cmbCategory.SelectedIndexChanged, tComments.TextChanged, tName.TextChanged, tSiteProfile.TextChanged
            If Not Loading Then Changed = True
        End Sub

    End Class

    Public Class NutrientEditorPage : Inherits FLPage
        Public WithEvents lstNutrients As SearchListItems, tName, tDRI, tDV, tUL, tEAR, tSiteLabel As TextBoxX, tAltNames As RichTextBoxX, chkBasic As CheckBoxX
        Public WithEvents cmbDVUnit, cmbDRIUnit, cmbULUnit, cmbEARUnit, cmbSiteLabels, cmbCategory As ComboBoxX
        Public WithEvents btnAdd, btnDelete, btnShiftUp, btnShiftDown As ButtonX
        Public CNutrient As NutrientProperty, CIndex As Integer = -1, CSite As DatabaseSite, NCDataLists As List(Of ComboBoxX), NTDataLists As List(Of TextBoxX)
        Public Overrides ReadOnly Property FileName() As String
            Get
                Return "Nutrients"
            End Get
        End Property

        Public Sub New(ByVal tPage As TabBrowser.Page)
            Try
                Page = tPage
                LoadControls()
                If Nutrients.Count > 0 Then
                    LoadNutrientList()
                    lstNutrients.lstItems.SelectedIndex = 0
                End If
            Catch ex As Exception
                MsgBox("Error loading Nutrient Editor page: " & ex.Message)
            End Try
        End Sub

        Private Sub LoadControls()
            Dim lblTemp As LabelX = Nothing, w As Integer = 70, w2 As Integer = 245
            lstNutrients = New SearchListItems(Page, "lstNutrients", "Nutrients:", 5, 5, 150, 300)
            lstNutrients.txtSearch.BorderStyle = BorderStyle.FixedSingle
            btnAdd = New ButtonX(Page, "btnAdd", "Add", lstNutrients.Bounds.Left, lstNutrients.Bounds.Bottom + 5, 70, 22)
            btnDelete = New ButtonX(Page, "btnDelete", "Delete", btnAdd.Right + 10, btnAdd.Top, 70, 22)
            btnShiftUp = New ButtonX(Page, "btnShiftUp", "Shift Up", btnAdd.Left, btnAdd.Bottom + 5, btnAdd.Width, btnAdd.Height)
            btnShiftDown = New ButtonX(Page, "btnShiftDown", "Shift Down", btnDelete.Left, btnDelete.Bottom + 5, btnDelete.Width, btnDelete.Height)

            lblTemp = New LabelX(Page, "", "Name:", lstNutrients.Bounds.Right + 5, lstNutrients.Bounds.Top, w, 22)
            tName = New TextBoxX(Page, "tName", "", lblTemp.Right + 2, lblTemp.Top, w2, 22)
            lblTemp = New LabelX(Page, "", "Alt Names:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            tAltNames = New RichTextBoxX(Page, "tAltNames", "", lblTemp.Right + 2, lblTemp.Top, w2, 66)
            lblTemp = New LabelX(Page, "", "Category:", lblTemp.Left, tAltNames.Bottom + 5, w, 22)
            cmbCategory = New ComboBoxX(Page, "cmbCategory", lblTemp.Right + 2, lblTemp.Top, 190, 22)
            cmbCategory.Items.AddRange(NCategories.ToArray)
            chkBasic = New CheckBoxX(Page, "chkBasic", "Basic", cmbCategory.Right + 5, cmbCategory.Top, -1, -1, False)
            lblTemp = New LabelX(Page, "", "Unit:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            cmbDVUnit = New ComboBoxX(Page, "cmbUnit", lblTemp.Right + 2, lblTemp.Top, w2, 22)
            lblTemp = New LabelX(Page, "", "Site Labels:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            cmbSiteLabels = New ComboBoxX(Page, "cmbSiteLabels", lblTemp.Right + 2, lblTemp.Top, 80, 22)
            For i As Integer = 0 To Sites.Count - 1
                cmbSiteLabels.Items.Add(Sites(i).Name)
            Next
            tSiteLabel = New TextBoxX(Page, "tSiteLabel", "", cmbSiteLabels.Right + 2, lblTemp.Top, tName.Right - cmbSiteLabels.Right - 2, 22)
            lblTemp = New LabelX(Page, "", "DRI:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            tDRI = New TextBoxX(Page, "tDRI", "", lblTemp.Right + 2, lblTemp.Top, 100, 22)
            cmbDRIUnit = New ComboBoxX(Page, "cmbDRIUnit", tDRI.Right + 2, tDRI.Top, tSiteLabel.Right - tDRI.Right - 2, 22)
            lblTemp = New LabelX(Page, "", "Daily Value:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            tDV = New TextBoxX(Page, "tDV", "", lblTemp.Right + 2, lblTemp.Top, w2, 22)
            lblTemp = New LabelX(Page, "", "UL:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            tUL = New TextBoxX(Page, "tUL", "", lblTemp.Right + 2, lblTemp.Top, 100, 22)
            cmbULUnit = New ComboBoxX(Page, "cmbULUnit", cmbDRIUnit.Left, lblTemp.Top, tDRI.Width, tDRI.Height)
            lblTemp = New LabelX(Page, "", "EAR:", lblTemp.Left, lblTemp.Bottom + 5, w, 22)
            tEAR = New TextBoxX(Page, "tEAR", "", lblTemp.Right + 2, lblTemp.Top, 100, 22)
            cmbEARUnit = New ComboBoxX(Page, "cmbEARUnit", cmbDRIUnit.Left, lblTemp.Top, tDRI.Width, tDRI.Height)

            NCDataLists = New List(Of ComboBoxX) : NTDataLists = New List(Of TextBoxX)
            NCDataLists.Add(cmbDRIUnit) : NCDataLists.Add(cmbDVUnit)
            NCDataLists.Add(cmbEARUnit) : NCDataLists.Add(cmbULUnit)
            NTDataLists.Add(tDRI) : NTDataLists.Add(tDV)
            NTDataLists.Add(tEAR) : NTDataLists.Add(tUL)
            Dim unitList() As Item = Units.ToArray()
            For i As Integer = 0 To 3
                NCDataLists(i).Items.AddRange(unitList)
            Next
        End Sub

        Public Sub LoadNutrientList()
            lstNutrients.Clear() 'whole sub takes 1 ms, don't bother removing sorting stuff for now
            Dim lstStuff As New ItemList()
            Dim Categs As New List(Of List(Of NutrientProperty))
            For i As Integer = 0 To NCategories.Count - 1
                Categs.Add(New List(Of NutrientProperty))
            Next
            For Each n As NutrientProperty In Nutrients
                Categs(n.Category.ID).Add(n)
            Next
            For i As Integer = 0 To Categs.Count - 1
                Dim Catego As List(Of NutrientProperty) = Categs(i) '(From n As NutrientProperty In Catego Where n.Parent = -1 Select n).ToList(), note that this is way slower
                Dim props As New List(Of NutrientProperty)
                For Each n As NutrientProperty In Catego
                    If n.Parent <> -1 Then Continue For
                    props.Add(n)
                Next
                For Each n As NutrientProperty In props
                    lstStuff.Add(n)
                    If n.SubProperties.Count > 0 Then
                        For i3 As Integer = 0 To n.SubProperties.Count - 1
                            Dim n2 As NutrientProperty = DirectCast(Nutrients.ByID(n.SubProperties(i3)), NutrientProperty)
                            lstStuff.Add(n2)
                        Next
                    End If
                Next
            Next
            Nutrients = lstStuff
            lstNutrients.AddRange(Nutrients)
        End Sub

        Public Overrides Sub Save()
            If Not CNutrient Is Nothing Then SaveNutrient(CNutrient)
        End Sub

        Protected Overrides Sub SaveList(sw As System.IO.StreamWriter)
            For Each n As NutrientProperty In Nutrients
                sw.WriteLine(n.SaveLine())
            Next
        End Sub

        Private Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
            Dim sInput As String = AddItem(lstNutrients)
            If sInput = "" Then Exit Sub
            Dim nt As New NutrientProperty(sInput, GetNextID)
            Nutrients.Add(nt)
            lstNutrients.Items.Add(nt, True)
            lstNutrients.lstItems.SelectedIndex = lstNutrients.Items.Count - 1
        End Sub

        Private Sub lstNutrients_SelectedIndexChanged(sender As Object, e As ListItems.SelectedIndexChangedArgs) Handles lstNutrients.SelectedIndexChanged
            If lstNutrients.Index = -1 Then Exit Sub
            LoadNutrient(DirectCast(lstNutrients.SItem, NutrientProperty))
            CIndex = lstNutrients.Index
        End Sub

        Private Sub LoadNutrient(ByVal tNutrient As NutrientProperty)
            Loading = True
            If Not CNutrient Is Nothing Then SaveNutrient(CNutrient)
            CSite = Nothing
            CNutrient = tNutrient
            tName.Text = CNutrient.Name
            tAltNames.Text = String.Join(Chr(10), CNutrient.AlternateNames)
            cmbCategory.SelectedIndex = NCategories.IndexOf(tNutrient.Category)
            For i As Integer = 0 To 3
                Dim cmb As ComboBoxX = NCDataLists(i), txt As TextBoxX = NTDataLists(i), nd As NutrientDataPair = tNutrient.NutrientData(i)
                cmb.SelectedIndex = cmb.Items.IndexOf(nd.Unit) : txt.Text = nd.Value.ToString()
            Next
            chkBasic.Checked = CNutrient.Basic
            tSiteLabel.Text = ""
            cmbSiteLabels.SelectedIndex = -1
            cmbSiteLabels.SelectedIndex = 0
            Loading = False
        End Sub

        Public Sub SaveNutrient(ByVal tNutrient As NutrientProperty)
            If tNutrient.Name <> tName.Text Then
                tNutrient.Name = tName.Text
                If CIndex <> -1 Then
                    If tNutrient.Parent = -1 Then lstNutrients.Items(CIndex) = tNutrient.Name Else lstNutrients.Items(CIndex) = "-" & tNutrient.Name
                End If
            End If
            CNutrient.Basic = chkBasic.Checked
            tNutrient.AlternateNames = tAltNames.Text.Split(Chr(10)).ToList()
            tNutrient.Category = DirectCast(cmbCategory.SelectedItem, Item)
            SaveSite()
            For i As Integer = 0 To 3
                Dim cmb As ComboBoxX = NCDataLists(i), txt As TextBoxX = NTDataLists(i), nd As NutrientDataPair = tNutrient.NutrientData(i)
                If cmb.SelectedIndex = -1 Then tNutrient.NutrientData(i) = New NutrientDataPair(nd.Index, CSng(txt.Text), Nothing) Else tNutrient.NutrientData(i) = New NutrientDataPair(nd.Index, CSng(txt.Text), DirectCast(cmb.SelectedItem, Unit))
            Next
        End Sub

        Private Sub SaveSite()
            If Not CSite Is Nothing AndAlso tSiteLabel.Text <> "" Then
                If CSite.IDs.ContainsKey(CNutrient.ID) Then
                    CSite.IDs(CNutrient.ID) = tSiteLabel.Text.Trim()
                Else
                    CSite.AddTerm(CNutrient, tSiteLabel.Text.Trim())
                End If
            End If
        End Sub

        Private Sub cmbSiteLabels_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbSiteLabels.SelectedIndexChanged
            If cmbSiteLabels.SelectedIndex = -1 Then Exit Sub
            SaveSite()
            CSite = DirectCast(Sites(cmbSiteLabels.SelectedIndex), DatabaseSite)
            If CSite.IDs.ContainsKey(CNutrient.ID) Then tSiteLabel.Text = CSite.IDs(CNutrient.ID) Else tSiteLabel.Text = ""
        End Sub

        Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click
            If CNutrient Is Nothing Then Exit Sub
            If MessageBox.Show("Delete nutrient " & Chr(34) & CNutrient.Name & Chr(34) & "?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
            Dim nu As Integer = CNutrient.ID
            Nutrients.Remove(CNutrient)
            lstNutrients.Remove(lstNutrients.Index)
            For Each fo As FoodItem In Foods
                For i2 As Integer = 0 To fo.DataSites.Count - 1  'need to remove this item if it exists
                    For i3 As Integer = 0 To fo.DataSites(i2).Properties.Count - 1
                        If fo.DataSites(i2).Properties(i3).Nutrient = nu Then
                            fo.DataSites(i2).Properties.RemoveAt(i3)
                            Exit For
                        End If
                    Next
                Next
            Next
        End Sub

        Private Sub btnShiftUp_Click(sender As Object, e As System.EventArgs) Handles btnShiftUp.Click
            If CNutrient Is Nothing Or lstNutrients.Index = 0 Then Exit Sub
            Dim PNutrient As NutrientProperty = DirectCast(lstNutrients.Items(lstNutrients.Index - 1), NutrientProperty)
            If PNutrient.Parent <> CNutrient.Parent Then Exit Sub
            If CNutrient.Parent = -1 Then 'top nutrient
                ShiftNutrient(-1, CNutrient)
            Else 'sub-nutrient
                If CNutrient.Parent = PNutrient.ID Then Exit Sub 'top-most sub-nutrient
                Dim ParentNutrient As NutrientProperty = DirectCast(Nutrients.ByID(CNutrient.Parent), NutrientProperty), indx2 As Integer = ParentNutrient.SubProperties.IndexOf(CNutrient.ID)
                ParentNutrient.SubProperties.RemoveAt(indx2)
                ParentNutrient.SubProperties.Insert(indx2 - 1, CNutrient.ID)
                ShiftNutrient(-1, CNutrient)
            End If
        End Sub

        Private Sub ShiftNutrient(ByVal tDirection As Integer, ByVal tNutrient As NutrientProperty)
            Dim indx As Integer = lstNutrients.Index
            Nutrients.Remove(CNutrient)
            Nutrients.Insert(lstNutrients.Index + tDirection, CNutrient)
            lstNutrients.lstItems.Items.RemoveAt(indx)
            lstNutrients.lstItems.Items.Insert(False, indx + tDirection, tNutrient)
            lstNutrients.lstItems.SelectedIndex = indx + tDirection
        End Sub

        Private Sub btnShiftDown_Click(sender As Object, e As System.EventArgs) Handles btnShiftDown.Click
            Try
                If CNutrient Is Nothing Or lstNutrients.Index = lstNutrients.Items.Count - 1 Then Exit Sub
                Dim NNutrient As NutrientProperty = DirectCast(lstNutrients.Items(lstNutrients.Index + 1), NutrientProperty)
                If NNutrient.Parent <> CNutrient.Parent Then Exit Sub
                If CNutrient.Parent = -1 Then 'top nutrient
                    ShiftNutrient(1, CNutrient)
                Else
                    Dim ParentNutrient As NutrientProperty = DirectCast(Nutrients.ByID(CNutrient.Parent), NutrientProperty), indx2 As Integer = ParentNutrient.SubProperties.IndexOf(CNutrient.ID)
                    ParentNutrient.SubProperties.RemoveAt(indx2)
                    ParentNutrient.SubProperties.Insert(indx2 + 1, CNutrient.ID)
                    ShiftNutrient(1, CNutrient)
                End If
            Catch ex As Exception
                MsgBox("Error shifting nutrient down: " & ex.Message)
            End Try
        End Sub

        Private Sub tDRI_TextChanged(sender As Object, e As System.EventArgs) Handles tDRI.TextChanged, tAltNames.TextChanged, tDV.TextChanged, tEAR.TextChanged, tName.TextChanged, _
            tSiteLabel.TextChanged, tUL.TextChanged, cmbCategory.SelectedIndexChanged, cmbDRIUnit.SelectedIndexChanged, cmbEARUnit.SelectedIndexChanged, cmbULUnit.SelectedIndexChanged, _
            cmbDVUnit.SelectedIndexChanged, chkBasic.CheckedChanged
            If Not Loading Then Changed = True
        End Sub

    End Class

End Class

Public Class LogEntry
    Implements IComparable

    Public EntryDate As Integer '20140829 '20141101, for comparing dates
    Public FoodEntries As New List(Of FoodEntry)
    Public Comments As String
    Public Completed As Boolean = False, AllMealsConsumed As Boolean = False

    Public Sub New(ByVal tDate As Integer)
        EntryDate = tDate
        _Logs.Add(tDate, Me)
    End Sub

    Public Overrides Function ToString() As String
        Dim eDate As String = EntryDate.ToString()
        Return eDate.Substring(0, 4) & "-" & eDate.Substring(4, 2) & "-" & eDate.Substring(6, 2)
    End Function

    Public Function CompareTo(obj As Object) As Integer Implements System.IComparable.CompareTo
        Return EntryDate.CompareTo(DirectCast(obj, LogEntry).EntryDate)
    End Function

    Public Sub Save(ByVal sw As StreamWriter)
        sw.WriteLine(EntryDate.ToString() & "|" & Completed.ToString() & "|" & AllMealsConsumed.ToString() & "|" & Comments.Replace(Chr(10), _NL) & "|" & (FoodEntries.Count - 1).ToString())
        If FoodEntries.Count > 0 Then
            Dim TFE As FoodEntry = FoodEntries(0), sLine As New StringBuilder(TFE.Food.ID.ToString() & ";" & TFE.ServingSize.ToString() & ";" & TFE.Amount.ToString())
            For i2 As Integer = 1 To FoodEntries.Count - 1
                TFE = FoodEntries(i2) : sLine.Append("|" & TFE.Food.ID.ToString() & ";" & TFE.ServingSize.ToString() & ";" & TFE.Amount.ToString())
            Next
            sw.WriteLine(sLine.ToString())
        End If
    End Sub

    Public Shared Function FromLine(ByVal sr As StreamReader) As LogEntry
        Dim sParts() As String = sr.ReadLine().Split(Chr(124))
        Dim tLog As New LogEntry(CInt(sParts(0)))
        tLog.Completed = CBool(sParts(1))
        tLog.AllMealsConsumed = CBool(sParts(2))
        tLog.Comments = sParts(3).Replace(_NL, Chr(10))
        Dim FECount As Integer = CInt(sParts(4))
        If FECount > -1 Then
            sParts = sr.ReadLine().Split(Chr(124))
            For i2 As Integer = 0 To FECount
                Dim sParts2() As String = sParts(i2).Split(Chr(59))
                tLog.FoodEntries.Add(New FoodEntry(DirectCast(Foods.ByID(CInt(sParts2(0))), FoodItem), CInt(sParts2(1)), CSng(sParts2(2))))
            Next
        End If
        Return tLog
    End Function
End Class

Public Class FoodEntry
    Public Food As FoodItem
    Public Amount As Single '1 serving1, 1.75 serving2, etc.
    Public ServingSize As Integer

    Public Sub New(ByVal tFood As FoodItem, ByVal tSS As Integer, ByVal tAmount As Single)
        Food = tFood
        ServingSize = tSS
        Amount = tAmount
    End Sub

    Public Overrides Function ToString() As String
        Try
            If ServingSize = 0 Then
                Return Food.ToString() & "; " & Amount.ToString() & " g"
            Else
                Return Food.ToString() & "; " & (Amount * Food.ServingSizes(ServingSize).Amount).ToString() & " g"
            End If
        Catch ex As Exception
            Return "ERROR"
        End Try
    End Function

End Class

Public Class Recipe : Inherits Item
    Public Entries As New List(Of FoodEntry), Ratios As Boolean = False, Comments As String = ""

    Public Sub New(ByVal tName As String, ByVal tID As Integer)
        MyBase.New(tName, tID)
    End Sub

    Public Shared Function FromLine(ByVal sLine As String) As Recipe
        Dim sParts() As String = sLine.Split(Chr(124))
        Dim rp As New Recipe(sParts(1), CInt(sParts(0)))
        rp.Ratios = CBool(sParts(2))
        rp.Comments = sParts(3).Replace(_NL, Chr(10))
        sParts = sParts(4).Split(Chr(59))
        For i As Integer = 0 To sParts.Length - 1
            Dim sParts2() As String = sParts(i).Split(Chr(44))
            rp.Entries.Add(New FoodEntry(DirectCast(Foods.ByID(CInt(sParts2(0))), FoodItem), CInt(sParts2(1)), CSng(sParts2(2))))
        Next
        Return rp
    End Function

    Public Function SaveLine() As String
        Dim sLine As New StringBuilder()
        sLine.Append(ID.ToString() & "|" & Name & "|" & Ratios.ToString() & "|" & Comments.Replace(Chr(10), _NL) & "|")
        If Entries.Count > 0 Then
            sLine.Append(Entries(0).Food.ID.ToString() & "," & Entries(0).ServingSize.ToString() & "," & Entries(0).Amount.ToString())
            For i2 As Integer = 1 To Entries.Count - 1
                sLine.Append(";" & Entries(i2).Food.ID.ToString() & "," & Entries(i2).ServingSize.ToString() & "," & Entries(i2).Amount.ToString())
            Next
        End If
        Return sLine.ToString()
    End Function
End Class

Public Class FoodItem : Inherits Item
    Public Variations As New List(Of Integer), Parent As FoodItem
    Public DataSites As New Dictionary(Of Integer, SiteProfile)
    Public ServingSizes As New List(Of ServingSize)
    Public Refuse As Single = 0, RefuseDescription As String = ""
    Public Category As Item = FCategoriesNA
    Public Comments As String = ""
    Public IsComboFood As Boolean = False

    Public Sub New(ByVal tName As String, ByVal tID As Integer)
        MyBase.New(tName, tID)
        ServingSizes.Add(New ServingSize("Value per 100 g", 100))
    End Sub

    Public Overrides Function ToString() As String
        If Parent Is Nothing Then Return Name Else Return Parent.Name & "; " & Name
    End Function

    Public Sub SaveFood(ByVal sw As StreamWriter)
        Dim Props As List(Of FoodProperty) = Nothing '& "|" & BooleanSave(IsComboFood)
        Dim sLine As New StringBuilder(ID & "|" & Name & "|" & String.Join(",", Variations) & "|" & BooleanSave(IsComboFood) & "|" & Refuse.ToString() & "|" & RefuseDescription & "|" & Comments & "|" & Category.ID.ToString() & "|")
        If ServingSizes.Count > 1 Then 'Always contains "Value per 100 g"
            sLine.Append(ServingSizes(1).Name & "~" & ServingSizes(1).Amount)
            For i2 As Integer = 2 To ServingSizes.Count - 1 'write the other SSs, if they exist
                sLine.Append(";" & ServingSizes(i2).Name & "~" & ServingSizes(i2).Amount)
            Next
        End If
        sLine.Append("|")
        If DataSites.Count > 0 Then
            With DataSites
                sLine.Append(.Values(0).Site.ID.ToString() & "," & .Values(0).Link)
                For i2 As Integer = 1 To DataSites.Count - 1
                    sLine.Append(";" & .Values(i2).Site.ID.ToString() & "," & .Values(i2).Link)
                Next
            End With
        End If
        sw.WriteLine(sLine)
        For i2 As Integer = 0 To DataSites.Count - 1
            sLine = New StringBuilder()
            Props = DataSites.Values(i2).Properties
            If Props.Count > 0 Then
                Props.Sort()
                sLine.Append(Props(0).Nutrient & " " & CInt(Props(0).Imputed).ToString() & " " & Props(0).Amount.ToString())
                For i3 As Integer = 1 To Props.Count - 1
                    sLine.Append("~" & Props(i3).Nutrient & " " & CInt(Props(i3).Imputed).ToString() & " " & Props(i3).Amount.ToString())
                Next
            End If
            sw.WriteLine(sLine)
        Next
    End Sub

    Public Shared Function FromSR(ByVal sr As StreamReader) As FoodItem
        Dim sParts() As String = sr.ReadLine().Split(Chr(124))
        Dim fd As New FoodItem(sParts(1), CInt(sParts(0)))
        fd.Variations = GetIntListFromLine(sParts(2))
        fd.IsComboFood = BooleanGet(sParts(3))
        fd.Refuse = CSng(sParts(4))
        fd.RefuseDescription = sParts(5)
        fd.Comments = sParts(6)
        fd.Category = FCategories.ByID(CInt(sParts(7)))
        Dim sParts2() As String = sParts(8).Split(Chr(59))
        If sParts2(0) <> "" Then
            For i As Integer = 0 To sParts2.Length - 1
                Dim indx As Integer = sParts2(i).IndexOf("~")
                fd.ServingSizes.Add(New ServingSize(sParts2(i).Substring(0, indx), CSng(sParts2(i).Substring(indx + 1))))
            Next
        End If
        sParts2 = sParts(9).Split(Chr(59))
        If sParts2(0) <> "" Then
            For i As Integer = 0 To sParts2.Length - 1
                Dim indx As Integer = sParts2(i).IndexOf(",")
                Dim SP As New SiteProfile(DirectCast(Sites.ByID(CInt(sParts2(i).Substring(0, indx))), DatabaseSite), sParts2(i).Substring(indx + 1))
                fd.DataSites.Add(SP.Site.ID, SP)
                Dim sParts3() As String = sr.ReadLine().Split(Chr(126))
                If sParts3(0) = "" Then Continue For
                For i2 As Integer = 0 To sParts3.Length - 1
                    Dim sParts4() As String = sParts3(i2).Split(Chr(32))
                    Dim FP As New FoodProperty(CInt(sParts4(0)), CSng(sParts4(2)))
                    FP.Imputed = CBool(sParts4(1))
                    SP.Properties.Add(FP)
                Next
            Next
        End If
        Return fd
    End Function

End Class

Public Class ServingSize
    Public Name As String, Amount As Single
    Public Sub New(ByVal tName As String, ByVal tAmount As Single)
        Name = tName
        Amount = tAmount
    End Sub

    Public Overrides Function ToString() As String
        If Name.StartsWith("Value per") Then Return "Grams" Else Return Name.ToString() & " (" & Amount.ToString() & "g)"
    End Function
End Class

Public Class SiteProfile
    Public Site As DatabaseSite, Link As String
    Public Properties As New List(Of FoodProperty)
    Public Sub New(ByVal tSite As DatabaseSite, ByVal tLink As String)
        Site = tSite
        Link = tLink
    End Sub
End Class

Public Class FoodProperty
    Implements IComparable

    Public Amount As Single, Nutrient As Integer, Imputed As Boolean = False
    Public Sub New(ByVal tProp As Integer, ByVal tAmount As Single) ', ByVal tStatus As FPStatus)
        Nutrient = tProp
        Amount = tAmount
    End Sub

    Public Overrides Function ToString() As String
        Return Nutrients.ByID(Nutrient).Name & ": " & Amount.ToString()
    End Function

    Public Function CompareTo(obj As Object) As Integer Implements System.IComparable.CompareTo
        Dim nu As NutrientProperty = DirectCast(Nutrients.ByID(Nutrient), NutrientProperty), nu2 As NutrientProperty = DirectCast(Nutrients.ByID(DirectCast(obj, FoodProperty).Nutrient), NutrientProperty)
        Return Nutrients.IndexOf(nu).CompareTo(Nutrients.IndexOf(nu2))
    End Function
End Class

Public MustInherit Class DatabaseSite : Inherits Item
    Public SiteBase As String, Terms As New Dictionary(Of String, NutrientProperty), IDs As New Dictionary(Of Integer, String)

    Public Sub New(ByVal tName As String, ByVal tID As Integer, ByVal tSiteBase As String)
        MyBase.New(tName, tID)
        SiteBase = tSiteBase
    End Sub

    Public Sub AddTerm(ByVal tNutrient As NutrientProperty, ByVal tTerm As String)
        Terms.Add(tTerm, tNutrient)
        IDs.Add(tNutrient.ID, tTerm)
    End Sub

    Public Overridable Sub Rip(ByVal tFood As FoodItem, ByVal tIndex As Integer)
        tFood.DataSites(tIndex).Properties.Clear()
    End Sub

    Public Class USDA : Inherits DatabaseSite
        Public Sub New(ByVal tName As String, ByVal tID As Integer, ByVal tSiteBase As String)
            MyBase.New(tName, tID, tSiteBase)
        End Sub

        Public Overrides Sub Rip(ByVal tFood As FoodItem, ByVal tIndex As Integer)
            MyBase.Rip(tFood, tIndex)
            Const sNutrientNameStart As String = "<td style=" & Chr(34) & "line-height:110%"
            Const sRefuseIndexStart As String = "Refuse:"
            WC.Encoding = System.Text.Encoding.UTF8
            Dim sPage As String = WC.DownloadString(tFood.DataSites(tIndex).Link)
            Dim stp As New Stopwatch()
            stp.Start()
            'Get column names, serving sizes and stuff
            Dim Index As Integer = 0, PIndex As Integer = sPage.IndexOf("Value per 100 g"), EndIndex As Integer = sPage.IndexOf("</thead>")
            Dim SSizes As New List(Of ServingSize), Props As List(Of FoodProperty) = tFood.DataSites(tIndex).Properties
            SSizes.Add(New ServingSize("Value per 100 g", 100)) 'Get serving sizes on page
            Try
                Dim RefuseIndex As Integer = sPage.IndexOf(sRefuseIndexStart), tRefuse As String = "", tRefuseDescription As String = ""
                If RefuseIndex <> -1 Then
                    RefuseIndex = sPage.IndexOf("<span class=", RefuseIndex)
                    RefuseIndex = sPage.IndexOf(">", RefuseIndex) + 1
                    tRefuse = sPage.Substring(RefuseIndex, sPage.IndexOf("%", RefuseIndex) - RefuseIndex).Trim()
                    RefuseIndex = sPage.IndexOf("Refuse Description:")
                    If RefuseIndex <> -1 Then
                        RefuseIndex = sPage.IndexOf("<span class=", RefuseIndex, StringComparison.Ordinal)
                        RefuseIndex = sPage.IndexOf(">", RefuseIndex) + 1
                        tRefuseDescription = sPage.Substring(RefuseIndex, sPage.IndexOf("</div>", RefuseIndex) - RefuseIndex).Trim()
                    End If
                    tFood.Refuse = CSng(tRefuse)
                    tFood.RefuseDescription = tRefuseDescription
                End If
            Catch ex As Exception
                MsgBox("Error obtaining Refuse information for " & tFood.Name & ": " & ex.Message)
                stp.Stop()
                stp = New Stopwatch()
                stp.Start()
            End Try
            Do
                Index = sPage.IndexOf("value=" & Chr(34), PIndex) + 7
                If Index = 6 Or Index > EndIndex Then Exit Do
                PIndex = sPage.IndexOf(Chr(34), Index)
                If PIndex = -1 Then Exit Do
                Dim sValue As String = sPage.Substring(Index, PIndex - Index)
                Index = sPage.IndexOf("<br/>", PIndex) + 5
                If Index = 4 Then Exit Do
                PIndex = sPage.IndexOf("</th>", Index)
                If PIndex = -1 Then Exit Do
                Dim sText As String = sPage.Substring(Index, PIndex - Index).Trim()
                Dim sSubText As String = sText.Substring(sText.IndexOf("<br/>") + 5).Trim()
                sText = sValue & " " & sText.Remove(sText.IndexOf("<br/>")).Trim()
                If sSubText <> "" Then
                    sSubText = System.Net.WebUtility.HtmlDecode(sSubText.Substring(0, sSubText.IndexOf("g")))
                End If
                SSizes.Add(New ServingSize(System.Net.WebUtility.HtmlDecode(sText), CSng(sSubText)))
            Loop
            Dim sLineStart As String = "<tr  style=" & Chr(34), bRoute As Boolean = False   'Get nutrients
            PIndex = EndIndex
            Do
                Index = sPage.IndexOf(sLineStart, PIndex)
                If Index = -1 Then
                    If sLineStart <> "<tr style=" & Chr(34) Then
                        sLineStart = "<tr style=" & Chr(34)
                        bRoute = True
                        Continue Do
                    Else
                        Exit Do
                    End If
                End If
                PIndex = Index

                Dim sLines(2) As String
                Dim NutrientNameIndex As Integer = sPage.IndexOf(sNutrientNameStart, Index) 'nutrient name and the following values use seperate delineator
                If NutrientNameIndex = -1 Then
                    NutrientNameIndex = sPage.IndexOf("<td  style" & "=" & Chr(34) & "text-align", Index) 'nutrient name and the following values use seperate delineator
                End If
                Index = sPage.IndexOf(">", NutrientNameIndex) + 1
                sLines(0) = System.Net.WebUtility.HtmlDecode(sPage.Substring(Index, sPage.IndexOf("<", Index) - Index)).Trim()
                For i As Integer = 1 To 2
                    If i = 0 AndAlso bRoute Then Index = sPage.IndexOf("<td  style=" & Chr(34) & "text-align", PIndex) Else Index = sPage.IndexOf("<td style=" & Chr(34) & "text-align", PIndex)
                    Index = sPage.IndexOf(">", Index) + 1
                    sLines(i) = System.Net.WebUtility.HtmlDecode(sPage.Substring(Index, sPage.IndexOf("<", Index) - Index)).Trim()
                    PIndex = Index
                Next
                'sLines(0) = unit, sLines(1) = value per 100
                'previously, sLines(1) = amount sLines(2) = value per 100
                If sLines(1) = "kJ" Then Continue Do
                Dim tUnit As Unit = GetUnit(sLines(1))
                If tUnit Is Nothing Then
                    If MessageBox.Show("Error: Cannot find unit " & Chr(34) & sLines(1) & Chr(34) & "; abort?", "uh do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then Exit Sub
                    Continue Do
                End If
                If Terms.ContainsKey(sLines(0)) Then
                    Dim n As NutrientProperty = Terms(sLines(0))
                    Dim iFactor As Single = 1
                    If tUnit IsNot n.Unit Then
                        iFactor = tUnit.GetConversionFactor(n.Unit)
                    End If
                    Dim np As New FoodProperty(n.ID, (CSng(sLines(2)) * iFactor))
                    'If sLines(3) = "--" Then np.Imputed = True
                    Props.Add(np)
                Else
                    Dim nu As NutrientProperty = frmNN.GetNutrient(sLines(0), Me, tUnit)
                    If nu Is Nothing Then Exit Sub
                End If
            Loop
            For i As Integer = 0 To tFood.ServingSizes.Count - 1
                For i2 As Integer = SSizes.Count - 1 To 1 Step -1
                    If tFood.ServingSizes(i).Amount = SSizes(i2).Amount Then SSizes.RemoveAt(i2)
                Next
            Next
            For i As Integer = 1 To SSizes.Count - 1
                tFood.ServingSizes.Add(SSizes(i))
            Next
            stp.Stop()
            frmMain.Text = stp.ElapsedMilliseconds.ToString() & "ms for " & tFood.Name
        End Sub
    End Class

    Public Class WHFoods : Inherits DatabaseSite
        Public Sub New(ByVal tName As String, ByVal tID As Integer, ByVal tSiteBase As String)
            MyBase.New(tName, tID, tSiteBase)
        End Sub

        Public Overrides Sub Rip(ByVal tFood As FoodItem, ByVal tIndex As Integer)
            Dim sPage As String = ""
            WC.Encoding = Encoding.UTF8
            WC.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; rv:2.0) Gecko/20100101 Firefox/4.0")

            sPage = WC.DownloadString(tFood.DataSites(tIndex).Link)
            If sPage = "" Then
                MsgBox("Error retrieving source code on " & Name & ": blank string returned.")
                Exit Sub
            End If

            Dim Index As Integer = 0, PIndex As Integer = sPage.IndexOf("In-depth nutrient analysis"), EndIndex As Integer = sPage.IndexOf("The nutrient profiles provided")
            Dim SSizes As New List(Of ServingSize)
            SSizes.Add(New ServingSize("Value per 100 g", 100)) 'Get serving sizes on page
            PIndex = sPage.IndexOf("<tr><td>", PIndex) + 8
            Index = sPage.IndexOf("</td>", PIndex)
            Dim sPart As String = sPage.Substring(PIndex, Index - PIndex).Trim(), tUnit As Unit = Nothing
            Dim sNameo As String = "", iAmounto As Single = -1
            If sPart.Contains("<") Then
                sNameo = sPart.Substring(0, sPart.IndexOf("<"))
            Else
                sNameo = sPart
            End If
            If sPart.IndexOf("(") <> -1 Then
                Dim roflo As Integer = sPart.IndexOf("(") + 1
                Dim sPart2 As String = sPart.Substring(roflo, sPart.IndexOf(")", roflo) - roflo).Trim()
                If sPart.IndexOf(" ") <> -1 Then
                    iAmounto = CSng(sPart2.Substring(0, sPart2.IndexOf(" ")).Trim())
                End If
                tUnit = GetUnit(sPart2.Substring(sPart2.IndexOf(" ") + 1).Trim())
                iAmounto *= tUnit.Relation
            End If
            SSizes.Add(New ServingSize(sNameo, iAmounto))

            Dim sLineStart As String = "<tr  style=" & Chr(34) & "display"   'Get nutrients
            PIndex = EndIndex
            Do
                Index = sPage.IndexOf(sLineStart, PIndex)
                If Index = -1 Then Exit Do
                PIndex = Index
                Dim sLines(SSizes.Count + 1) As String
                For i As Integer = 0 To SSizes.Count + 1
                    Index = sPage.IndexOf("<td style=" & Chr(34) & "text-align", PIndex)
                    Index = sPage.IndexOf(">", Index) + 1
                    sLines(i) = sPage.Substring(Index, sPage.IndexOf("<", Index) - Index).Trim()
                    PIndex = Index
                Next
                tUnit = GetUnit(sLines(1))
            Loop
        End Sub
    End Class

    Public Class Manual : Inherits DatabaseSite
        Public Sub New(ByVal tName As String, ByVal tID As Integer, ByVal tSiteBase As String)
            MyBase.New(tName, tID, tSiteBase)
        End Sub
    End Class

End Class

Public Class NutrientProperty : Inherits Item
    Public Category As Item
    Public AlternateNames As New List(Of String)
    Public NutrientData(3) As NutrientDataPair '0 = DRI, 1 = DV, 2 = EAR, 3 = UL
    Public SubProperties As New List(Of Integer), Parent As Integer = -1
    Public Basic As Boolean = False

    Public ReadOnly Property Unit() As Unit
        Get
            Return NutrientData(1).Unit
        End Get
    End Property

    Public Sub New(ByVal tName As String, ByVal tID As Integer)
        MyBase.New(tName, tID)
    End Sub

    Public Function SaveLine() As String
        Dim sLine As New StringBuilder(ID & "|" & Name & "|" & String.Join(";", AlternateNames) & "|" & Category.ID.ToString() & "|" & Parent.ToString() & "|" & String.Join(",", SubProperties) & "|" & CInt(Basic).ToString())
        For Each nd As NutrientDataPair In NutrientData
            sLine.Append("|" & nd.Value.ToString() & "|" & nd.Unit.ID.ToString())
        Next
        Return sLine.ToString()
    End Function

    Public Shared Function FromLine(ByVal sParts() As String) As NutrientProperty
        Dim n As New NutrientProperty(sParts(1), CInt(sParts(0)))
        n.AlternateNames = sParts(2).Split(Chr(59)).ToList()
        If sParts(3) <> "" Then n.Category = NCategories.ByID(CInt(sParts(3)))
        If sParts(4) <> "" Then n.Parent = CInt(sParts(4))
        n.SubProperties = GetIntListFromLine(sParts(5))
        n.Basic = CBool(sParts(6))
        For i As Integer = 0 To 3
            n.NutrientData(i) = New NutrientDataPair(CByte(i), CSng(sParts(7 + (i * 2))), DirectCast(Units.ByID(CInt(sParts(8 + (i * 2)))), Unit))
        Next
        Return n
    End Function

    Public Overrides Function ToString() As String
        If Parent = -1 Then
            If Unit.Abbrev.Count > 0 Then Return Name & ", " & Unit.Abbrev(0) Else Return Name
        Else
            If Unit.Abbrev.Count > 0 Then Return "-" & Name & ", " & Unit.Abbrev(0) Else Return "-" & Name
        End If
    End Function
End Class

Public Structure NutrientDataPair
    Private Shared Names() As String = {"DRI", "DV", "EAR", "UL"}
    Public Enum Fields As Byte
        DRI = 0
        DV = 1
        EAR = 2
        UL = 3
    End Enum

    Public Value As Single, Unit As Unit, Index As Byte
    Public ReadOnly Property Name() As String
        Get
            Return Names(Index)
        End Get
    End Property

    Public Shared Function GetName(ByVal mode As Byte) As String
        Return [Enum].GetName(GetType(Fields), DirectCast(mode, Fields))
    End Function

    Public Sub New(ByVal tIndex As Byte, ByVal tValue As Single, ByVal tUnit As Unit)
        Index = tIndex
        Value = tValue
        Unit = tUnit
    End Sub

    Public Overrides Function ToString() As String
        Return Name & ": " & Value & " " & Unit.Abbrev(0)
    End Function
End Structure

Public Class Unit : Inherits Item
    Public Abbrev As New List(Of String), BUnit As Integer
    Public Relation As Single = 1

    Public Sub New(ByVal tBaseUnit As Integer, ByVal tName As String, ByVal tID As Integer, ByVal tAbbrev As List(Of String), ByVal tRelation As Single)
        MyBase.New(tName, tID)
        Abbrev = tAbbrev
        Relation = tRelation
        BUnit = tBaseUnit
    End Sub

    Public Function GetConversionFactor(ByVal tUnit As Unit) As Single
        If BUnit <> tUnit.BUnit Then
            If tUnit.BUnit = -1 Then
                Return Relation
            Else
                MsgBox("Error: Cannot convert " & Abbrev(0) & " to " & tUnit.Abbrev(0) & ".")
                Return -1
            End If
        ElseIf tUnit.ID = BUnit Then
            Return 1
        Else
            Return Relation / (tUnit.Relation)
        End If
    End Function

    Public Overrides Function ToString() As String
        Return Name & ", " & Abbrev(0)
    End Function
End Class

Public Class Item
    Implements IComparable, ICloneable

    Public Name As String

    Private _ID As Integer
    Public Property ID() As Integer
        Get
            Return _ID
        End Get
        Set(ByVal value As Integer)
            _ID = value
            If HighestID < value Then
                HighestID = value
            End If
        End Set
    End Property

    Public Sub New(ByVal tName As String, ByVal tID As Integer)
        Name = tName
        ID = tID
    End Sub

    Public Function CompareTo(obj As Object) As Integer Implements System.IComparable.CompareTo
        Return Name.CompareTo(DirectCast(obj, Item).Name)
    End Function

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function

    Public Overrides Function ToString() As String
        Return Name
    End Function
End Class

Public Class ItemList
    Implements ICloneable
    Implements IList(Of Item)
    Private _List As List(Of Item)

    Public Shared Function FromList(ByVal list As IEnumerable(Of Item)) As ItemList
        Dim ilist = New ItemList()
        If list.GetType() = GetType(List(Of Item)) Then ilist._List = DirectCast(list, List(Of Item)) Else ilist._List = list.ToList()
        Return ilist
    End Function

    Public Sub New()
        _List = New List(Of Item)
    End Sub

    Public Function Find(t As Predicate(Of Item)) As Item
        Return _List.Find(t)
    End Function

    Public Function FindIndex(t As Predicate(Of Item)) As Integer
        Return _List.FindIndex(t)
    End Function

    Public Function FindByName(ByVal sName As String) As Item
        Return _List.Find(Function(x) x.Name = sName)
    End Function

    Public Function GetInsertionIndex(ByVal sb As Item) As Integer
        If _List.Count = 0 OrElse sb.Name = "0Main" Then Return 0
        Dim sName As String = sb.Name.ToLower()
        For i As Integer = 0 To _List.Count - 1
            Dim lsto As Item = _List(i)
            If lsto.Name = "0Main" Then Continue For
            If lsto.Name.ToLower() > sName Then Return i
        Next
        Return _List.Count
    End Function

    Default Public Property Item(index As Integer) As Item Implements System.Collections.Generic.IList(Of Item).Item
        Get
            Return _List(index)
        End Get
        Set(value As Item)
            _List(index) = value
        End Set
    End Property

    Public Sub Add(item As Item) Implements System.Collections.Generic.ICollection(Of Item).Add
        _List.Add(item)
    End Sub

    Public Sub AddRange(ByVal lst As List(Of Item))
        _List.AddRange(lst)
    End Sub

    Public Function ByID(ByVal tID As Integer) As Item
        Return _List.Find(Function(x) x.ID = tID)
    End Function

    Public Sub Clear() Implements System.Collections.Generic.ICollection(Of Item).Clear
        _List.Clear()
    End Sub

    Public Function Contains(item As Item) As Boolean Implements System.Collections.Generic.ICollection(Of Item).Contains
        Return _List.Contains(item)
    End Function

    Public Sub CopyTo(array() As Item, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of Item).CopyTo
        _List.CopyTo(array)
    End Sub

    Public ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of Item).Count
        Get
            Return _List.Count
        End Get
    End Property

    Public ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of Item).IsReadOnly
        Get
            Return False
        End Get
    End Property

    Public Function Remove(item As Item) As Boolean Implements System.Collections.Generic.ICollection(Of Item).Remove
        Return _List.Remove(item)
    End Function

    Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of Item) Implements System.Collections.Generic.IEnumerable(Of Item).GetEnumerator
        Return _List.GetEnumerator
    End Function

    Public Function IndexOf(item As Item) As Integer Implements System.Collections.Generic.IList(Of Item).IndexOf
        Return _List.IndexOf(item)
    End Function

    Public Overloads Sub Insert(index As Integer, item As Item) Implements System.Collections.Generic.IList(Of Item).Insert
        _List.Insert(index, item)
    End Sub

    Public Sub RemoveAt(index As Integer) Implements System.Collections.Generic.IList(Of Item).RemoveAt
        _List.RemoveAt(index)
    End Sub

    Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return _List.GetEnumerator
    End Function

    Public Overloads Function Insert(item As Item) As Integer
        Dim itm As String = item.Name.ToLower()
        Dim indx As Integer = _List.FindIndex(Function(x) x.Name.ToLower() > itm)
        If indx = -1 Then indx = _List.Count
        _List.Insert(indx, item)
        Return indx
    End Function

    Public Function RenameItem(ByVal tIndex As Integer, ByVal tName As String) As Integer
        Dim tItem As Item = _List(tIndex)
        _List(tIndex).Name = tName
        _List.RemoveAt(tIndex)
        Dim indx As Integer = Insert(tItem)
        Return indx
    End Function

    Public Sub Sort()
        _List.Sort()
    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
End Class

Public Class ObjectBox
    Public Text As String, Obj As Object

    Public Sub New(ByVal tText As String, ByVal tSC As Object)
        Text = tText
        Obj = tSC
    End Sub

    Public Overrides Function ToString() As String
        Return Text
    End Function
End Class

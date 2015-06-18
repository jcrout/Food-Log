Option Strict On

Imports System
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Linq
Imports System.Linq.Expressions
Imports MyControls

Module General
    Public WithEvents frmMain As Form1, F As FoodLog, frmSS As AddSSForm, frmNN As NewNutrientForm
    Public Nutrients As New ItemList(), Foods As New ItemList(), Sites As New ItemList(), ManualSite As DatabaseSite, ComboSite As DatabaseSite
    Public Units As New ItemList(), NCategories As New ItemList(), FCategories As New ItemList(), FCategoriesNA As Item, Recipies As New ItemList()
    Public Logs As New List(Of LogEntry), _Logs As New Dictionary(Of Integer, LogEntry)
    Public WC As WebClient, CDirectory As String = "", HighestID As Integer = 0
    Public Const _NL As String = "*N*", _ComboFoodStart As String = "***"

    <STAThread()>
    Public Sub Main(ByVal args() As String)
        Dim createdNew As Boolean = True
        Using mutex As New Threading.Mutex(True, "Food Log", createdNew) 'prevents multiiple being open
            If createdNew Then
                CDirectory = Application.StartupPath & "\"
                WC = New WebClient
                F = New FoodLog()
                frmMain.ShowDialog()
            Else
                Dim current As Process = Process.GetCurrentProcess()
                For Each process__1 As Process In Process.GetProcessesByName(current.ProcessName)
                    If process__1.Id <> current.Id Then
                        SetForegroundWindow(process__1.MainWindowHandle)
                        Exit For
                    End If
                Next
            End If
        End Using
    End Sub

    Public Function GetNextID() As Integer
        HighestID += 1
        Return HighestID
    End Function

    Public Function BooleanSave(ByVal tBool As Boolean) As String
        If tBool Then Return "1" Else Return ""
    End Function

    Public Function BooleanGet(ByVal str As String) As Boolean
        If str = "1" Then Return True Else Return False
    End Function

    Public Sub InsertFood(ByVal tFood As FoodItem)
        Dim tName As String = tFood.Name.ToUpper(), InsIndex As Integer = -1
        For i As Integer = 0 To Foods.Count
            Dim fd As FoodItem = DirectCast(Foods(i), FoodItem)
            If fd.Parent IsNot Nothing Then Continue For
            If fd.Name.ToUpper() > tName Then
                InsIndex = i
                Exit For
            End If
        Next
        If InsIndex = -1 Then InsIndex = Foods.Count
        Foods.Insert(InsIndex, tFood)
        F.tabFood.lstFood.Insert(InsIndex, New ObjectBox(_ComboFoodStart & tFood.Name, tFood))
        F.tabFood.Changed = True
        F.tabFood.SaveLists()
    End Sub

    Public Function GetAdjustedValue(ByVal n As NutrientProperty, ByVal tUnit As Unit, ByVal tValue As Single) As Single
        If tUnit IsNot n.Unit Then
            If n.Unit Is Units(4) Then 'if IU
                If n.Name.StartsWith("Vitamin D") Then
                    Return CSng(tValue * 0.025)
                ElseIf n.Name.StartsWith("Vitamin A") Then
                    Return CSng(tValue * 0.3)
                Else
                    MsgBox("OMG")
                    Return tValue
                End If
            Else
                Return n.Unit.GetConversionFactor(tUnit) * tValue
            End If
        Else
            Return tValue
        End If
    End Function

    <DllImport("user32.dll")> _
    Private Function SetForegroundWindow(hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Public Function GetDataNumber() As Integer
        Return CInt(Date.Today.Year.ToString() & FormatTwoDigits(Date.Today.Month) & FormatTwoDigits(Date.Today.Day))
    End Function

    Public Function FormatTwoDigits(ByVal iNum As Integer) As String
        If iNum > 9 Then Return iNum.ToString() Else Return "0" & iNum.ToString()
    End Function

    Public Function GetUnit(ByVal tName As String) As Unit
        tName = tName.Replace("&micro;g", "µg")
        For i As Integer = 0 To Units.Count - 1
            Dim u As Unit = DirectCast(Units(i), Unit)
            If u.Name = tName Then
                Return u
            ElseIf u.Abbrev.Contains(tName) Then
                Return u
            End If
        Next
        Return Nothing
    End Function

    Public Function GetIntListFromLine(ByVal sLine As String) As List(Of Integer)
        If sLine = "" Then Return New List(Of Integer)
        Dim lst As New List(Of Integer), intString() As String = sLine.Split(CChar(","))
        For i As Integer = 0 To intString.Count - 1
            lst.Add(CInt(intString(i)))
        Next
        Return lst
    End Function

    Public Function AddItem(ByVal lstSearch As SearchListItems) As String
        Dim sInput As String = InputBox("Enter name:", "Food", "").Trim()
        If sInput = "" Then Return ""
        sInput = sInput.Substring(0, 1).ToUpper() & sInput.Substring(1)
        If lstSearch.Items.ContainsText(sInput, True) Then
            MsgBox("Error: Item with name " & Chr(34) & sInput & Chr(34) & " already exists.")
            Return ""
        End If
        Return sInput
    End Function

    Private Sub frmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles frmMain.FormClosing
        F.Save()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As System.EventArgs) Handles frmMain.Load
        frmSS = New AddSSForm()
        frmNN = New NewNutrientForm()
    End Sub

End Module

Public Structure StringAndNum
    Public ReadOnly Text As String, Number As Single
    Public Sub New(ByVal tText As String, ByVal tNumber As Single)
        Text = tText
        Number = tNumber
    End Sub

    Public Overrides Function ToString() As String
        Return Text & ": " & Number.ToString()
    End Function
End Structure

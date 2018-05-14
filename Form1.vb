Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Public Class Form1
    Public xlApp As Excel.Application ' = New Microsoft.Office.Interop.Excel.Application()
    Public WB As Excel.Workbook
    Public WS As Excel.Worksheet
    Public WS_range As Excel.Range
    Public excFileName As String = "C:\Users\User\downloads\таблица.xlsx"
    Public Part_Arr As String = "СПИСОК ДЕТАЛЕЙ"
    Public LastInd As Integer
    Public rw As Excel.Range
    Dim uchet As Integer

    Sub Create_EX_Doc(Visible As Double)
        xlApp = New Excel.Application()
        xlApp.Visible = Visible
        WB = xlApp.Workbooks.Add(1)
        WS = WB.Sheets(1)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'create a new TreeView
        Dim TreeView1 As TreeView
        TreeView1 = Me.TreeView1
        'TreeView1.Location = New Point(10, 10)
        'TreeView1.Size = New Size(150, 150)
        'Me.Controls.Add(TreeView1)
        'TreeView1.Nodes.Clear()
        'Creating the root node
        Dim root = New TreeNode("Application")
        TreeView1.Nodes.Add(root)
        TreeView1.Nodes(0).Nodes.Add(New TreeNode("Project 1"))
        'Creating child nodes under the first child
        For loopindex As Integer = 1 To 4
            TreeView1.Nodes(0).Nodes(0).Nodes.Add(New _
                TreeNode("Sub Project" & Str(loopindex)))
        Next loopindex
        ' creating child nodes under the root
        TreeView1.Nodes(0).Nodes.Add(New TreeNode("Project 6"))
        'creating child nodes under the created child node
        For loopindex As Integer = 1 To 3
            TreeView1.Nodes(0).Nodes(1).Nodes.Add(New _
                TreeNode("Project File" & Str(loopindex)))
        Next loopindex
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        xlApp = New Excel.Application()
        xlApp.Visible = Visible
        WB = xlApp.Workbooks.Open(excFileName)
        WS = WB.Sheets.Item(Part_Arr)

        Dim root = New TreeNode("Вентилятор")
        TreeView1.Nodes.Clear()
        TreeView1.Nodes.Add(root)
        recvAdd(root, 5, True, 233)
    End Sub
    Private Function recvAdd(node As TreeNode, rownum As Integer, isFirst As Boolean, totalNum As Integer)
        Dim numpp As String = WS.Cells(rownum, 1).Value()

        Dim ptrn1
        If isFirst Then
            ptrn1 = New Regex("^.+")
        Else
            ptrn1 = New Regex("^" + node.Text + ".+")
        End If
        While (rownum <= totalNum)
            If (numpp Is Nothing) Then
                numpp = node.Text + ".1"
            End If
            If Not ptrn1.IsMatch(numpp) Then
                Exit While
            End If
            Dim newNode As New TreeNode(numpp)
            node.Nodes.Add(newNode)
            rownum += 1
            Dim nextnumpp As String = WS.Cells(rownum, 1).Value()
            If (nextnumpp Is Nothing) Then
                If (isFirst) Then
                    nextnumpp = CInt(newNode.Text) + 1
                Else
                    Dim prevNum As Integer = CInt(Regex.Match(newNode.Text, "[0-9]+$").Value)
                    nextnumpp = node.Text + "." + CStr(prevNum + 1)
                End If
            End If
            Dim ptrn2 As New Regex("^" + numpp + "\.[0-9]+$")
            If ptrn2.IsMatch(nextnumpp) Then
                rownum = recvAdd(newNode, rownum, False, totalNum)
                numpp = WS.Cells(rownum, 1).Value()
            Else
                numpp = nextnumpp
            End If
            If Regex.IsMatch(numpp, "10$") Then
                totalNum += 1
                totalNum -= 1
            End If
        End While


        Return rownum

    End Function
End Class

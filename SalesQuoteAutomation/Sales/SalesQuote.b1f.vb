Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System.IO
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Windows.Forms

Namespace SalesQuoteAutomation
    <FormAttribute("SQUO", "Sales/SalesQuote.b1f")>
    Friend Class SalesQuote
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Public Filename As String = ""
        Dim objrs As SAPbobsCOM.Recordset
        Dim BankFileName = ""
        Public objfile As FileInfo
        Public WithEvents DBSource As SAPbouiCOM.DBDataSource
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("202").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("MtxData").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("lCustName").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("CusName").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton0 = CType(Me.GetItem("Item_5").Specific, SAPbouiCOM.LinkedButton)
            Me.EditText2 = CType(Me.GetItem("txtFName").Specific, SAPbouiCOM.EditText)
            Me.Button3 = CType(Me.GetItem("Import").Specific, SAPbouiCOM.Button)
            Me.StaticText1 = CType(Me.GetItem("QARef").Specific, SAPbouiCOM.StaticText)
            Me.StaticText2 = CType(Me.GetItem("lblRFQDate").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("RFQDate").Specific, SAPbouiCOM.EditText)
            Me.EditText5 = CType(Me.GetItem("txtDoc").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("DEntry").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("tdocentry").Specific, SAPbouiCOM.EditText)
            Me.ComboBox0 = CType(Me.GetItem("Series").Specific, SAPbouiCOM.ComboBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeBefore, AddressOf Me.Form_ResizeBefore
            AddHandler DataAddBefore, AddressOf Me.Form_DataAddBefore

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("SQUO", 0)
                'objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                DBSource = objform.DataSources.DBDataSources.Item("@MIPL_SQUO")
                objaddon.objglobalmethods.LoadSeries(objform, DBSource)
                objaddon.objapplication.Menus.Item("1300").Activate()
                'EditText6.Value = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_SQUO")
                'EditText5.Value = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_SQUO")
                objform.Items.Item("RFQDate").Specific.String = Now.Date.ToString("dd/MM/yy")
                objform.ActiveItem = "RFQDate"
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "Series", True, True, False)
                StaticText4.Item.Visible = False
                EditText5.Item.Visible = False
                objform.Settings.Enabled = True
                objform.Freeze(False)
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Finally
                objform.Freeze(False)
            End Try
            

        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton

        Private Sub DefaultSeries()
            Try
                Dim dftlseries As String = objaddon.objglobalmethods.default_series("SQUO", IIf(EditText4.String = "", Now.Date.ToString("dd/MM/yy"), EditText4.String))
                'Dim ocomboseries As SAPbouiCOM.ComboBox = objform.Items.Item("Series").Specific
                'If dftlseries.ToString <> "" Then ComboBox0.Select(dftlseries.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception
            End Try
        End Sub
        Private WithEvents EditText1 As SAPbouiCOM.EditText

        Private Sub EditText0_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.LostFocusAfter
            Try
                objaddon.objapplication.Menus.Item("1300").Activate()
                EditText6.Value = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_SQUO")
                EditText5.Value = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_SQUO")
                Matrix0.AddRow()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
            
        End Sub
        Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
        Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim excelRng As Microsoft.Office.Interop.Excel.Range
        Private Sub ReadExcel(ByVal FileName As String)
            Dim j As Integer = 1
            Try
                FileName = objform.Items.Item("txtFName").Specific.string
                Dim RowIndex As Integer
                ExcelWorkbook = ExcelApp.Workbooks.Open(FileName)
                ExcelWorkSheet = ExcelWorkbook.ActiveSheet
                'excelRng = ExcelWorkSheet.Range("A1")
                excelRng = ExcelWorkSheet.UsedRange

                objaddon.objapplication.SetStatusBarMessage("Excel Data Loading please wait... " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Long, False)
                'objform.Freeze(True)
                Matrix0.Clear()
                If objform.Items.Item("txtFName").Specific.String <> "" Then
                    For RowIndex = 2 To excelRng.Rows.Count
                        Matrix0.AddRow()
                        If ExcelWorkSheet.Cells(1, 1).Value = "Item Code" And ExcelWorkSheet.Cells(1, 2).Value = "SAP Item Description" And ExcelWorkSheet.Cells(1, 3).Value = "MPN" And ExcelWorkSheet.Cells(1, 4).Value = "Make" And ExcelWorkSheet.Cells(1, 5).Value = "Customer Description" And ExcelWorkSheet.Cells(1, 6).Value = "RFQ QTY" And ExcelWorkSheet.Cells(1, 7).Value = "QTY" And ExcelWorkSheet.Cells(1, 8).Value = "Remarks" And ExcelWorkSheet.Cells(1, 9).Value = "Customer Target Price" And ExcelWorkSheet.Cells(1, 10).Value = "CPN" And ExcelWorkSheet.Cells(1, 11).Value = "SQ Type" Then
                            Matrix0.Columns.Item("#").Cells.Item(j).Specific.String = j
                            Matrix0.Columns.Item("SItem").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 1).Value)
                            Matrix0.Columns.Item("SDesc").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 2).Value)
                            Matrix0.Columns.Item("MPN").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 3).Value)
                            Matrix0.Columns.Item("Make").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 4).Value)
                            Matrix0.Columns.Item("Desc").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 5).Value)
                            Matrix0.Columns.Item("RFQQty").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 6).Value)
                            Matrix0.Columns.Item("Quant").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 7).Value)
                            Matrix0.Columns.Item("Remarks").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 8).Value)
                            Matrix0.Columns.Item("Cusprice").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 9).Value)
                            Matrix0.Columns.Item("CPN").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 10).Value)
                            Matrix0.Columns.Item("SQType").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 11).Value)
                            j += 1
                        Else
                            objaddon.objapplication.SetStatusBarMessage("Incorrect Excel Format...Please update correct format as it's from Automation Screen...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            Exit Sub
                        End If
                    Next RowIndex
                    'objform.Freeze(False)
                    objaddon.objapplication.Menus.Item("1300").Activate()
                End If

                'ExcelApp.ActiveWorkbook.Close()
                objaddon.objapplication.StatusBar.SetText("Excel Data Successfully Loaded!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'objaddon.objapplication.SetStatusBarMessage("Excel Loaded..." & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Long, False)
            Catch ex As Exception
                'ExcelApp.ActiveWorkbook.Close()
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Finally
                objform.Freeze(False)
                ExcelApp.ActiveWorkbook.Close()
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub
        Private Sub test()
            Dim conn As OleDbConnection
            Dim dtr As OleDbDataReader
            Dim dta As OleDbDataAdapter
            Dim cmd As OleDbCommand
            Dim dts As DataSet
            Dim excel As String
            Try
                Dim j As Integer = 1
            excel = objform.Items.Item("txtFName").Specific.string
            conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excel + ";Extended Properties=Excel 12.0;")
            dta = New OleDbDataAdapter("Select * From [Sheet1$]", conn)
            dts = New DataSet
                dta.Fill(dts, "[Sheet1$]")

                For i As Integer = dts.Tables(0).Rows.Count - 1 To 0 Step -1
                    Dim row As DataRow = dts.Tables(0).Rows(i)
                    If row.Item(2) Is Nothing Then
                        dts.Tables(0).Rows.Remove(row)
                    ElseIf row.Item(2).ToString = "" Then
                        dts.Tables(0).Rows.Remove(row)
                    End If
                Next
                'objform.Freeze(True)
                For RowIndex = -1 To dts.Tables(0).Rows.Count - 1
                    If dts.Tables(0).Rows.Count > 0 Then
                        Matrix0.AddRow()
                        'dts.Tables(0).Columns.Count
                        Matrix0.Columns.Item("#").Cells.Item(j).Specific.String = j
                        If dts.Tables(0).Columns(0).ToString <> "" Then
                            Matrix0.Columns.Item("SItem").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(0).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 1).Value)
                        End If
                        If dts.Tables(0).Columns(1).ToString <> "" Then
                            Matrix0.Columns.Item("SDesc").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(1).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 2).Value)
                        End If
                        If dts.Tables(0).Columns(2).ToString <> "" Then
                            Matrix0.Columns.Item("MPN").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(2).ToString ' CStr(ExcelWorkSheet.Cells(RowIndex, 3).Value)
                        End If
                        If dts.Tables(0).Columns(3).ToString <> "" Then
                            Matrix0.Columns.Item("Make").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(3).ToString ' CStr(ExcelWorkSheet.Cells(RowIndex, 4).Value)
                        End If
                        If dts.Tables(0).Columns(4).ToString <> "" Then
                            Matrix0.Columns.Item("Desc").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(4).ToString ' CStr(ExcelWorkSheet.Cells(RowIndex, 5).Value)
                        End If
                        If dts.Tables(0).Columns(5).ToString <> "" Then
                            Matrix0.Columns.Item("SPQ").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(5).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 6).Value)
                        End If
                        If dts.Tables(0).Columns(6).ToString <> "" Then
                            Matrix0.Columns.Item("MOQ").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(6).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 7).Value)
                        End If
                        If dts.Tables(0).Columns(7).ToString <> "" Then
                            Matrix0.Columns.Item("Quant").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(7).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 8).Value)
                        End If
                        If dts.Tables(0).Columns(8).ToString <> "" Then
                            Matrix0.Columns.Item("UnitPr").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(8).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 9).Value)
                        End If
                        If dts.Tables(0).Columns(9).ToString <> "" Then
                            Matrix0.Columns.Item("LeadT").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(9).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 10).Value)
                        End If
                        If dts.Tables(0).Columns(10).ToString <> "" Then
                            Matrix0.Columns.Item("usd").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(10).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 11).Value)
                        End If
                        If dts.Tables(0).Columns(11).ToString <> "" Then
                            Matrix0.Columns.Item("value").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(11).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 12).Value)
                        End If
                        'If dts.Tables(0).Columns(12).ToString <> "" Then
                        '    Matrix0.Columns.Item("CPN").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(12).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 13).Value)
                        'End If
                        'If dts.Tables(0).Columns(13).ToString <> "" Then
                        '    Matrix0.Columns.Item("Remarks").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(13).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 14).Value)
                        'End If
                        'If dts.Tables(0).Columns(14).ToString <> "" Then
                        '    Matrix0.Columns.Item("Cusprice").Cells.Item(j).Specific.String = dts.Tables(0).Rows(RowIndex + 1)(14).ToString 'CStr(ExcelWorkSheet.Cells(RowIndex, 15).Value)
                        'End If
                        j += 1

                    End If
                Next RowIndex
                'objform.Freeze(False)
            conn.Close()
            objaddon.objapplication.Menus.Item("1300").Activate()
                objaddon.objapplication.StatusBar.SetText("Excel Data Successfully Loaded!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Finally
                ExcelApp.ActiveWorkbook.Close()
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Public DocEntry As String = ""
        Private Function SalesQuotationScreen() As Boolean
            Dim objSOform As SAPbouiCOM.Form
            Dim objMatrix As SAPbouiCOM.Matrix
            Dim objRS As SAPbobsCOM.Recordset
            Dim objSQType As SAPbouiCOM.ComboBox
            Dim ItemCode As String = ""
            Dim MPNCode As String = "", StrQuery As String = ""
            objaddon.objapplication.Menus.Item("2049").Activate()
            objSOform = objaddon.objapplication.Forms.ActiveForm
            objSOform.Visible = True
            objSOform.Height = 700
            objSOform.Width = 700
            Try
                objSOform = objaddon.objapplication.Forms.Item(objSOform.UniqueID)
                objMatrix = objSOform.Items.Item("38").Specific
                objSOform.Items.Item("4").Specific.String = objform.Items.Item("CusName").Specific.String
                objaddon.objapplication.SetStatusBarMessage("Excel Data Copying please wait... " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Long, False)
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objSOform.Freeze(True)
                'objSOform.Items.Item("12").Specific.String = DateTime.Today.AddDays(29).ToString("dd/MM/yy")
                objSOform.Items.Item("U_RFQ").Specific.String = objform.Items.Item("tdocentry").Specific.String
                objSOform.Items.Item("U_Date").Specific.String = objform.Items.Item("RFQDate").Specific.String
                For i As Integer = 1 To Matrix0.RowCount
                    'ItemCode = objaddon.objglobalmethods.getSingleValue("CALL ""GetValidItemCode""('" & Matrix0.Columns.Item("SItem").Cells.Item(i).Specific.String & "');")
                    'If ItemCode = "" Then
                    MPNCode = objaddon.objglobalmethods.getSingleValue("CALL ""GetValidMPN_MakeCode""('" & Matrix0.Columns.Item("MPN").Cells.Item(i).Specific.String & "','" & Matrix0.Columns.Item("Make").Cells.Item(i).Specific.String & "');")
                    StrQuery = "select ""U_SPQ"",""U_MOQ"",""U_Value"",""U_LeadTime"" from OITM where ""ItemCode""='" & MPNCode & "' and""U_OrderPN""='" & Matrix0.Columns.Item("MPN").Cells.Item(i).Specific.String & "' and ""U_Make""='" & Matrix0.Columns.Item("Make").Cells.Item(i).Specific.String & "'"
                    objRS.DoQuery(StrQuery)
                    objMatrix.Columns.Item("1").Cells.Item(i).Specific.String = MPNCode
                    If CDbl(Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String) = 0.0 Then
                        objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = "1" 'Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String
                    Else
                        objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = Matrix0.Columns.Item("Quant").Cells.Item(i).Specific.String
                    End If
                    objMatrix.Columns.Item("U_PrtNo").Cells.Item(i).Specific.String = Matrix0.Columns.Item("MPN").Cells.Item(i).Specific.String
                    objMatrix.Columns.Item("U_Make").Cells.Item(i).Specific.String = Matrix0.Columns.Item("Make").Cells.Item(i).Specific.String
                    objMatrix.Columns.Item("U_CDesc").Cells.Item(i).Specific.String = Matrix0.Columns.Item("Desc").Cells.Item(i).Specific.String
                    objMatrix.Columns.Item("U_CQTY").Cells.Item(i).Specific.String = Matrix0.Columns.Item("RFQQty").Cells.Item(i).Specific.String
                    objMatrix.Columns.Item("U_CustPartNo").Cells.Item(i).Specific.String = Matrix0.Columns.Item("CPN").Cells.Item(i).Specific.String
                    objMatrix.Columns.Item("U_Remarks").Cells.Item(i).Specific.String = Matrix0.Columns.Item("Remarks").Cells.Item(i).Specific.String
                    objMatrix.Columns.Item("U_CustPrice").Cells.Item(i).Specific.String = Matrix0.Columns.Item("Cusprice").Cells.Item(i).Specific.String
                    objSQType = objMatrix.Columns.Item("U_SQType").Cells.Item(i).Specific
                    If Matrix0.Columns.Item("SQType").Cells.Item(i).Specific.String <> "" Then
                        objSQType.Select(Matrix0.Columns.Item("SQType").Cells.Item(i).Specific.String, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If

                    If Matrix0.Columns.Item("SQType").Cells.Item(i).Specific.String = "Online" Or Matrix0.Columns.Item("SQType").Cells.Item(i).Specific.String = "ONLINE" Then
                        objMatrix.Columns.Item("U_SPQ").Cells.Item(i).Specific.String = "1" 'objRS.Fields.Item("U_SPQ").Value.ToString
                    Else
                        objMatrix.Columns.Item("U_SPQ").Cells.Item(i).Specific.String = objRS.Fields.Item("U_SPQ").Value.ToString
                    End If
                    objMatrix.Columns.Item("U_MOQ").Cells.Item(i).Specific.String = objRS.Fields.Item("U_MOQ").Value.ToString
                    objMatrix.Columns.Item("U_Value").Cells.Item(i).Specific.String = objRS.Fields.Item("U_Value").Value.ToString
                    objMatrix.Columns.Item("U_LeadTime").Cells.Item(i).Specific.String = objRS.Fields.Item("U_LeadTime").Value.ToString
                Next
                'objSOform.Freeze(False)

                objaddon.objapplication.StatusBar.SetText("Excel Data Copied to Sales Quotation Successfully!!! " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            Finally
                objSOform.Freeze(False)
                GC.WaitForPendingFinalizers()
                GC.Collect()
                objform.Refresh()
                objform.Update()
            End Try
            'objaddon.objapplication.SetStatusBarMessage("Excel data Copied to SO..." & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Long, False)
        End Function
        Private Sub releaseObject(ByVal obj As Object)
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
            Finally
                GC.Collect()
            End Try

        End Sub
        Public Function FindFile() As String

            Dim ShowFolderBrowserThread As Threading.Thread
            Try
                ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)

                If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                    ShowFolderBrowserThread.Start()
                ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                    ShowFolderBrowserThread.Start()
                    ShowFolderBrowserThread.Join()
                End If

                While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                    System.Windows.Forms.Application.DoEvents()
                    ' ShowFolderBrowserThread.Sleep(100)
                    Thread.Sleep(100)
                End While

                If BankFileName <> "" Then
                    Return BankFileName
                End If

            Catch ex As Exception

                objaddon.objapplication.MessageBox("File Find  Method Failed : " & ex.Message)
            End Try
            Return ""
        End Function

        Public Sub ShowFolderBrowser()
            Dim MyProcs() As System.Diagnostics.Process
            Dim nw As New NativeWindow

            Dim OpenFile As New OpenFileDialog
            Try
                ' Dim initialpath As String = objaddon.objglobalmethods.getSingleValue("select ""ExcelPath"" from oadm")
                Dim initialpath As String = System.Windows.Forms.Application.StartupPath + "\"
                OpenFile.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                OpenFile.Multiselect = False
                'OpenFile.ShowDialog()
                OpenFile.Filter = "All files(*.)|*.*" '   "|*.*"
                Dim filterindex As Integer = 0
                Try
                    filterindex = 0
                Catch ex As Exception
                End Try

                Dim form As New System.Windows.Forms.Form
                form.TopMost = True
                OpenFile.FilterIndex = filterindex
                OpenFile.RestoreDirectory = True
                'OpenFile.CheckFileExists = True
                'OpenFile.CheckPathExists = True
                MyProcs = Process.GetProcessesByName("SAP Business One")
                'nw.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")(0).MainWindowHandle)
                'NativeWindow.FromHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")(0).MainWindowHandle)
                'If MyProcs.Length = 1 Then
                If MyProcs.Length >= 1 Then
                    For i As Integer = 0 To MyProcs.Length - 1
                        Dim comname As String() = MyProcs(i).MainWindowTitle.ToString.Split("-")

                        'Open dialog only for the company where the button is clicked
                        Dim com As String = objaddon.objcompany.CompanyName.ToString.Trim.ToUpper
                        'If comname(1).ToString.Trim.ToUpper = com Then
                        Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)

                        'Dim ret As System.Windows.Forms.DialogResult = OpenFile.ShowDialog(MyWindow)
                        'Dim ret As System.Windows.Forms.DialogResult = OpenFile.
                        If OpenFile.ShowDialog(NativeWindow.FromHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")(0).MainWindowHandle)) <> System.Windows.Forms.DialogResult.Cancel Then
                            BankFileName = OpenFile.FileName
                            'OpenFile.Dispose()
                        Else
                            System.Windows.Forms.Application.ExitThread()
                        End If
                        'End If
                        Exit For
                    Next
                    '  Else
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message)
                BankFileName = ""
            Finally
                OpenFile.Dispose()
            End Try
        End Sub
        
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Button3 As SAPbouiCOM.Button
        Public Class WindowWrapper

            Implements System.Windows.Forms.IWin32Window
            Private _hwnd As IntPtr

            Public Sub New(ByVal handle As IntPtr)
                _hwnd = handle
            End Sub

            Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
                Get
                    Return _hwnd
                End Get
            End Property

        End Class

        Private Sub Button3_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If objform.Items.Item("txtFName").Specific.String <> "" Then
                        ReadExcel(Filename)
                        'test()
                    Else
                        objaddon.objapplication.SetStatusBarMessage("Please Choose a file...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        Exit Sub
                    End If
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Finally
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If EditText0.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Customer Code is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        objform.ActiveItem = "CusName"
                        BubbleEvent = False : Exit Sub
                    End If
                    If Matrix0.Columns.Item("MPN").Cells.Item(Matrix0.VisualRowCount).Specific.string = "" Then
                        Matrix0.DeleteRow(Matrix0.VisualRowCount)
                    End If
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
           

        End Sub

        Private Sub EditText0_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText0.ChooseFromListBefore

            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_1")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "C"
                ' oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND      
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub EditText0_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText0.Value = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub
       
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText

        Private Sub Button1_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            objform.Freeze(True)
            Matrix0.Clear()
            objform.Items.Item("CusName").Specific.string = ""
            objform.Items.Item("txtFName").Specific.string = ""
            objform.Freeze(False)
        End Sub

        Private Sub Button0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objaddon.objglobalmethods.LoadSeries(objform, DBSource)
                'EditText6.Value = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_SQUO")
                'EditText5.Value = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_SQUO")
                objform.Items.Item("RFQDate").Specific.String = Now.Date.ToString("dd/MM/yy")
            End If
        End Sub
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText

        Private Sub Form_ResizeBefore(pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean)
            Try
                EditText2.Item.Left = 214
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox

        Private Sub Form_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As System.Boolean)
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If SalesQuotationScreen() Then
                    Else
                        BubbleEvent = False : Exit Sub
                    End If
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Finally
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try

        End Sub
    End Class
End Namespace

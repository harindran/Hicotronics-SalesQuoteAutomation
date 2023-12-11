Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace SalesQuoteAutomation
    <FormAttribute("149", "Sales/FrmSalesQuotation.b1f")>
    Friend Class FrmSalesQuotation
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.StaticText0 = CType(Me.GetItem("RFQDate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("U_Date").Specific, SAPbouiCOM.EditText)
            Me.EditText1 = CType(Me.GetItem("U_RFQ").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub



        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("149", 0)
                EditText0.Item.Top = objform.Items.Item("46").Top + 18
                EditText0.Item.Left = objform.Items.Item("46").Left
                StaticText0.Item.Top = objform.Items.Item("86").Top + 18
                StaticText0.Item.Left = objform.Items.Item("86").Left
                EditText1.Item.Top = EditText0.Item.Top 'objform.Items.Item("46").Top + 15 'objform.Items.Item("U_Date").Top + 2
                EditText1.Item.Left = EditText0.Item.Left + EditText0.Item.Width + 2 'objform.Items.Item("U_Date").Left
                'StaticText1.Item.Top = objform.Items.Item("RFQDate").Top + 15
                'StaticText1.Item.Left = objform.Items.Item("RFQDate").Left
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                EditText0.Item.Top = objform.Items.Item("46").Top + 18
                EditText0.Item.Left = objform.Items.Item("46").Left
                StaticText0.Item.Top = objform.Items.Item("86").Top + 18
                StaticText0.Item.Left = objform.Items.Item("86").Left
                EditText1.Item.Top = EditText0.Item.Top 'objform.Items.Item("46").Top + 15 'objform.Items.Item("U_Date").Top + 2
                EditText1.Item.Left = EditText0.Item.Left + EditText0.Item.Width + 2 'objform.Items.Item("U_Date").Left
                'StaticText1.Item.Top = objform.Items.Item("RFQDate").Top + 15
                'StaticText1.Item.Left = objform.Items.Item("RFQDate").Left
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
    End Class
End Namespace

Imports SAPbouiCOM
Namespace SalesQuoteAutomation

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "SQUO"
                        SalesQuote_MenuEvent(pVal, BubbleEvent)

                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                Else
                    Select Case pval.MenuUID
                        Case "1281"
                        Case Else
                    End Select
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "SalesQuote"

        Private Sub SalesQuote_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim DBSource As SAPbouiCOM.DBDataSource
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"
                            BubbleEvent = False
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tdocentry").Enabled = True
                            objform.ActiveItem = "tdocentry"
                            objform.Items.Item("tdocentry").Click(BoCellClickType.ct_Regular)
                        Case "1282" ' Add Mode
                            DBSource = objform.DataSources.DBDataSources.Item("@MIPL_SQUO")
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource)
                            'objform.Items.Item("tdocentry").Specific.String = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_SQUO")
                            'objform.Items.Item("txtDoc").Specific.String = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_SQUO")
                        Case "1288", "1289", "1290", "1291"
                            objaddon.objapplication.Menus.Item("1300").Activate()

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#End Region




    End Class
End Namespace
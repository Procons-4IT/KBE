Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsProductionOrder
    Inherits clsBase
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Private oGrid As SAPbouiCOM.Grid
    Private oDtSplitList As SAPbouiCOM.DataTable
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private strQuery As String
    Dim oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_UpdateProductionOrder, frm_UpdateProductionOrder)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            Dim ostatic As SAPbouiCOM.StaticText
            ostatic = oForm.Items.Item("31").Specific
            ostatic.Caption = ""
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.UserDataSources.Add("frmOrd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("toOrd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("frmOrDt", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("toOrDt", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("frmDue", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("toDue", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("frmCard", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("toCard", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("frmItem", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("toItem", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("ItmsGrp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEdit = oForm.Items.Item("7").Specific
            oEdit.DataBind.SetBound(True, "", "frmOrd")
            oEdit.ChooseFromListUID = "CFL_6"
            oEdit.ChooseFromListAlias = "DocNum"
            oEdit = oForm.Items.Item("8").Specific
            oEdit.DataBind.SetBound(True, "", "toOrd")
            oEdit.ChooseFromListUID = "CFL_7"
            oEdit.ChooseFromListAlias = "DocNum"

            oEdit = oForm.Items.Item("10").Specific
            oEdit.DataBind.SetBound(True, "", "frmOrDt")

            oEdit = oForm.Items.Item("11").Specific
            oEdit.DataBind.SetBound(True, "", "toOrDt")

            oEdit = oForm.Items.Item("13").Specific
            oEdit.DataBind.SetBound(True, "", "frmDue")

            oEdit = oForm.Items.Item("14").Specific
            oEdit.DataBind.SetBound(True, "", "toDue")

            oCombo = oForm.Items.Item("22").Specific
            oCombo.DataBind.SetBound(True, "", "ItmsGrp")
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select ""ItmsGrpCod"",""ItmsGrpNam"" from OITB order by ""ItmsGrpCod""")
            oCombo.ValidValues.Add("", "")
            For introw As Integer = 0 To oRecordSet.RecordCount - 1
                oCombo.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value)
                oRecordSet.MoveNext()
            Next
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            addChooseFromListConditions(oForm)
            oForm.PaneLevel = 1

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)


            oCFL = oCFLs.Item("CFL_4")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "TreeType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_5")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "TreeType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SelectAll(ByVal aForm As SAPbouiCOM.Form, ByVal aflag As Boolean)
        oGrid = aForm.Items.Item("3").Specific
        aForm.Freeze(True)
        Dim oCheckBox As SAPbouiCOM.CheckBoxColumn
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheckBox = oGrid.Columns.Item("Check")
            oCheckBox.Check(intRow, aflag)
        Next
        aForm.Freeze(False)
    End Sub


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
                Case mnu_UpdateProductionOrder
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

    Private Function DataBind(aform As SAPbouiCOM.Form) As Boolean
        Dim strSQL As String
        Dim strCondition, strFrmOrdDate, strToOrdDate, strFrmDueDate, strToDueDate, strFrmDocNum, strToDocNum, strFrmCardcode, strToCardCode, strFrmItem, strToItem, stritemgroup As String
        Dim dtFrmOrdDate, dtToOrdDate, dtFrmDueeDate, dtToDuedate As Date

        Dim strStatus As String
        oCombo = aform.Items.Item("30").Specific
        Try
            strStatus = oCombo.Selected.Value
        Catch ex As Exception
            strStatus = ""
        End Try
        If strStatus = "" Then
            oApplication.Utilities.Message("Production Order is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If


        strFrmDocNum = oApplication.Utilities.getEditTextvalue(aform, "7")
        strToDocNum = oApplication.Utilities.getEditTextvalue(aform, "8")

        strFrmOrdDate = oApplication.Utilities.getEditTextvalue(aform, "10")
        strToOrdDate = oApplication.Utilities.getEditTextvalue(aform, "11")
        strFrmDueDate = oApplication.Utilities.getEditTextvalue(aform, "13")
        strToDueDate = oApplication.Utilities.getEditTextvalue(aform, "14")
        strFrmCardcode = oApplication.Utilities.getEditTextvalue(aform, "16")
        strToCardCode = oApplication.Utilities.getEditTextvalue(aform, "17")
        strFrmItem = oApplication.Utilities.getEditTextvalue(aform, "19")
        strToItem = oApplication.Utilities.getEditTextvalue(aform, "20")
        oCombo = aform.Items.Item("22").Specific
        stritemgroup = oCombo.Selected.Value
        If strFrmOrdDate = "" Then
            strCondition = " (1=1 "
        Else
            dtFrmOrdDate = oApplication.Utilities.GetDateTimeValue(strFrmOrdDate)
            strCondition = " ( T0.PostDate >='" & dtFrmOrdDate.ToString("yyyy-MM-dd") & "'"
        End If

        If strToOrdDate = "" Then
            strCondition = strCondition & " and 1=1 ) "
        Else
            dtToOrdDate = oApplication.Utilities.GetDateTimeValue(strToOrdDate)
            strCondition = strCondition & " and T0.PostDate <='" & dtToOrdDate.ToString("yyyy-MM-dd") & "')"
        End If


        If strFrmDueDate = "" Then
            strCondition = strCondition & "  and (1=1 "
        Else
            dtFrmDueeDate = oApplication.Utilities.GetDateTimeValue(strFrmDueDate)
            strCondition = strCondition & "  and  ( T0.DueDate >='" & dtFrmDueeDate.ToString("yyyy-MM-dd") & "'"
        End If

        If strToDueDate = "" Then
            strCondition = strCondition & " and 1=1 ) "
        Else
            dtToDuedate = oApplication.Utilities.GetDateTimeValue(strToDueDate)
            strCondition = strCondition & " and T0.DueDate <='" & dtToDuedate.ToString("yyyy-MM-dd") & "')"
        End If

        If strFrmCardcode = "" Then
            strCondition = strCondition & "  and (1=1 "
        Else
            strCondition = strCondition & "  and ( T1.CardCode >='" & strFrmCardcode & "'"

        End If

        If strToCardCode = "" Then
            strCondition = strCondition & "  and 1=1) "
        Else
            strCondition = strCondition & "  and  T1.CardCode <='" & strToCardCode & "')"
        End If


        If strFrmItem = "" Then
            strCondition = strCondition & "  and (1=1 "
        Else
            strCondition = strCondition & "  and ( T2.ItemCode >='" & strFrmItem & "'"

        End If

        If strToItem = "" Then
            strCondition = strCondition & "  and 1=1) "
        Else
            strCondition = strCondition & "  and  T2.ItemCode <='" & strToItem & "')"
        End If
        If stritemgroup = "" Then

        Else
            strCondition = strCondition & " and T3.ItmsGrpCod='" & stritemgroup & "'"

        End If

        strCondition = strCondition & " and T0.Status='" & strStatus & "'"
        strSQL = "SELECT  T0.[DocEntry],T0.[DocNum], T0.[ItemCode],T2.ItemName, T0.CardCode,T1.CardName, T0.[PostDate], T0.[DueDate], T0.[Warehouse], T0.[Status] FROM OWOR T0  Left Outer JOIN OCRD T1 ON T0.[CardCode] = T1.[CardCode] INNER JOIN OITM T2 ON T2.[ItemCode] = T0.[ItemCode] INNER JOIN OITB T3 ON T2.[ItmsGrpCod] = T3.[ItmsGrpCod]"
        strSQL = strSQL & " where " & strCondition
        Dim oEditText As SAPbouiCOM.EditTextColumn
        oGrid = aform.Items.Item("28").Specific
        oGrid.DataTable.ExecuteQuery(strSQL)

        oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Production Order Entry"
        oEditText = oGrid.Columns.Item("DocEntry")
        oEditText.LinkedObjectType = "202"
        oGrid.Columns.Item("DocNum").TitleObject.Caption = "Production Order Number"
        oEditText = oGrid.Columns.Item("DocNum")
        oEditText.LinkedObjectType = "202"
        oGrid.Columns.Item("ItemCode").TitleObject.Caption = "Product No"
        oEditText = oGrid.Columns.Item("ItemCode")
        oEditText.LinkedObjectType = "4"
        oGrid.Columns.Item("CardCode").TitleObject.Caption = "Customer Code"
        oEditText = oGrid.Columns.Item("CardCode")
        oEditText.LinkedObjectType = "2"
        oGrid.Columns.Item("CardName").TitleObject.Caption = "Customer Name"
        oGrid.Columns.Item("ItemName").TitleObject.Caption = "Product Name"
        oGrid.Columns.Item("PostDate").TitleObject.Caption = "Posting Date"
        oGrid.Columns.Item("DueDate").TitleObject.Caption = "Due Date"
        oGrid.Columns.Item("Warehouse").TitleObject.Caption = "Warehouse"
        oGrid.Columns.Item("Status").TitleObject.Caption = "Status"
        oGrid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        Dim oCombobox1 As SAPbouiCOM.ComboBoxColumn
        oCombobox1 = oGrid.Columns.Item("Status")
        oCombobox1.ValidValues.Add("P", "Planned")
        oCombobox1.ValidValues.Add("R", "Release")
        oCombobox1.ValidValues.Add("L", "Close")
        oCombobox1.ValidValues.Add("C", "Cancel")
        oCombobox1.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        oCombobox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oGrid.AutoResizeColumns()
        aform.Items.Item("28").Enabled = False
        If strStatus = "C" Or strStatus = "L" Then
            aform.Items.Item("5").Enabled = False
        Else
            aform.Items.Item("5").Enabled = True
        End If
        Return True
    End Function

    Private Function UpdateProductionOrder(aform As SAPbouiCOM.Form) As Boolean
        Dim strStatus As String

        oCombo = aform.Items.Item("30").Specific
        strStatus = oCombo.Selected.Value
        If strStatus = "P" Then
            If oApplication.SBO_Application.MessageBox("Do you want to change the Selected production order status from Planned to release ?", , "Continue", "Cancel") = 2 Then
                Return False
            End If
        End If

        If strStatus = "R" Then
            If oApplication.SBO_Application.MessageBox("Do you want to change the Selected production order status from Release to Close ?", , "Continue", "Cancel") = 2 Then
                Return False
            End If
        End If

        oGrid = aform.Items.Item("28").Specific
        Dim oPO As SAPbobsCOM.ProductionOrders
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        Dim ostatic As SAPbouiCOM.StaticText
        ostatic = aform.Items.Item("31").Specific
        oApplication.Company.StartTransaction()
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oPO = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
            ostatic.Caption = "Processing Production Order : " & oGrid.DataTable.GetValue("DocNum", intRow)
            If oPO.GetByKey(oGrid.DataTable.GetValue("DocEntry", intRow)) Then
                If strStatus = "P" Then
                    oPO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                ElseIf strStatus = "R" Then
                    oPO.ClosingDate = Now.Date
                    oPO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposClosed
                End If
                If oPO.Update <> 0 Then
                    ostatic.Caption = "Error Occured . Production order Number : " & oGrid.DataTable.GetValue("DocNum", intRow) & ": Error : " & oApplication.Company.GetLastErrorDescription

                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If oApplication.Company.InTransaction Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                End If
            End If
        Next
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        ostatic.Caption = "Process Completed"
        oApplication.Utilities.Message("Operation Completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_UpdateProductionOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CLICK

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID


                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "4"
                                        If oForm.PaneLevel = 2 Then
                                            If DataBind(oForm) = False Then
                                                Exit Sub
                                            End If
                                        End If
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                    Case "5"
                                        If UpdateProductionOrder(oForm) = True Then
                                            oForm.Close()
                                        End If
                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim val1, val, Val2 As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If pVal.ItemUID = "16" Or pVal.ItemUID = "17" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                                'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                '    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                'End If
                                            End Try
                                        End If
                                        If pVal.ItemUID = "19" Or pVal.ColUID = "20" Then
                                            val1 = oDataTable.GetValue("ItemCode", 0)
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                                'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                '    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                'End If
                                            End Try
                                        End If

                                        If pVal.ItemUID = "7" Or pVal.ColUID = "8" Then
                                            val1 = oDataTable.GetValue("DocNum", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val1)
                                            Catch ex As Exception
                                                'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                '    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                'End If
                                            End Try
                                        End If
                                    End If
                                Catch ex As Exception
                                End Try
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        If oForm.TypeEx = frm_Production Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    oDBDataSource = oForm.DataSources.DBDataSources.Item(0)
                    If oDBDataSource.GetValue("Status", 0).ToString = "P" And oDBDataSource.GetValue("U_Split", 0).ToString <> "Y" And oDBDataSource.GetValue("U_BaseProd", 0).ToString = "" Then
                        If Not oMenuItem.SubMenus.Exists(mnu_GenerateSplit) Then
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_GenerateSplit
                            oCreationPackage.String = "Split Production Order"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If
                    ElseIf oDBDataSource.GetValue("Status", 0).ToString <> "P" Or oDBDataSource.GetValue("U_BaseProd", 0).ToString <> "" Then
                        If oMenuItem.SubMenus.Exists(mnu_GenerateSplit) Then
                            Try
                                oMenuItem.SubMenus.RemoveEx(mnu_GenerateSplit)
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub
#End Region

#Region "Function"

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub callSplipProductionList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim objCommCharge As clsSplitList
            objCommCharge = New clsSplitList
            Dim strPONo As String = (oForm.Items.Item("18").Specific.value)
            oDBDataSource = oForm.DataSources.DBDataSources.Item(0)
            strPONo = oDBDataSource.GetValue("DocEntry", 0)

            Dim dblPlannedQty As Double = CDbl(oForm.Items.Item("12").Specific.value)
            Dim oRecordset As SAPbobsCOM.Recordset
            oRecordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select OrdrMulti,ItemCode,ItemName,U_MinPer From OITM Where ItemCode = '" + oApplication.Utilities.getEditTextvalue(oForm, 6) + "'"
            oRecordset.DoQuery(strQuery)
            If oRecordset.Fields.Item(0).Value <= 0 Then
                oApplication.Utilities.Message("Order Multiple Qty is not defined for this Item.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            objCommCharge.LoadForm(strPONo, dblPlannedQty)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class

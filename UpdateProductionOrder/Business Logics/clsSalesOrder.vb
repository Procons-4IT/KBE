Public Class clsSalesOrder
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
            'oForm = oApplication.SBO_Application.Forms.ActiveForm()
            'oForm.Freeze(True)
            'initialize(oForm)
            'Dim ostatic As SAPbouiCOM.StaticText
            'ostatic = oForm.Items.Item("31").Specific
            'ostatic.Caption = ""
            'oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    


    
    


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            'Select Case pVal.MenuUID
            '    Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            '    Case mnu_ADD
            '    Case mnu_UpdateProductionOrder
            '        If pVal.BeforeAction = False Then
            '            LoadForm()
            '        End If
            'End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

    



    Private Sub AddControl(aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            oApplication.Utilities.AddControls(aform, "_10", "2", SAPbouiCOM.BoFormItemTypes.it_STATIC, "RIGHT", 0, 0, "2", "Item Group", 100)
            oApplication.Utilities.AddControls(aform, "_11", "_10", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, "RIGHT", 0, 0, "_10", , 120)
            oApplication.Utilities.AddControls(aform, "_12", "_11", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "_11", "Update Allow procurement", 130)
            aform.DataSources.UserDataSources.Add("Itms", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aform.Items.Item("_11").AffectsFormMode = False
            oCombo = aform.Items.Item("_11").Specific
            oCombo.DataBind.SetBound(True, "", "Itms")
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("Select ItmsGrpCod,ItmsGrpNam from OITB order by ItmsGrpCod")
            oCombo.ValidValues.Add("", "")
            For intRow As Integer = 0 To oTest.RecordCount - 1
                oCombo.ValidValues.Add(oTest.Fields.Item(0).Value, oTest.Fields.Item(1).Value)
                oTest.MoveNext()
            Next
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            aform.Items.Item("_11").DisplayDesc = True
            aform.Freeze(False)
        Catch ex As Exception
            aform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub PopulateMainItem(omatrix As SAPbouiCOM.Matrix, aRow As Integer, aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strMainItem, strAddon As String
            oCombo = omatrix.Columns.Item("U_Z_MItemType").Cells.Item(aRow).Specific
            strAddon = oCombo.Selected.Value
            strMainItem = ""
            If strAddon = "A" Then
                If aRow > 1 Then
                    For intRow As Integer = aRow - 1 To 1 Step -1
                        Try
                            oCombo = omatrix.Columns.Item("U_Z_MItemType").Cells.Item(intRow).Specific
                            If oCombo.Selected.Value = "M" Then
                                strMainItem = oApplication.Utilities.getMatrixValues(omatrix, "1", intRow)
                                Exit For
                            End If
                        Catch ex As Exception

                        End Try
                       
                    Next
                End If
                Try
                    If strMainItem <> "" Then
                        oApplication.Utilities.SetMatrixValues(omatrix, "U_Z_MRelateItem", aRow, strMainItem)
                    End If
                Catch ex As Exception

                End Try
            End If
            aform.Freeze(False)
        Catch ex As Exception
            aform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub procurement(aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim strItemGroup, stritem As String
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim oCheckbox As SAPbouiCOM.CheckBox
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = aform.Items.Item("38").Specific
            oCombo = aform.Items.Item("_11").Specific
            strItemGroup = oCombo.Selected.Value
            If strItemGroup <> "" Then
                For intLoop As Integer = 1 To oMatrix.RowCount
                    stritem = oApplication.Utilities.getMatrixValues(oMatrix, "1", intLoop)
                    oTest.DoQuery("Select * from OITM T0 inner Join OITB T1 on T1.ItmsGrpCod=T0.ItmsGrpCod where T0.ItemCode='" & stritem & "' and T1.ItmsGrpCod=" & strItemGroup)
                    If oTest.RecordCount > 0 Then
                        Try
                            oCheckbox = oMatrix.Columns.Item("234000353").Cells.Item(intLoop).Specific
                            oCheckbox.Checked = True
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            aform.Freeze(False)
        Catch ex As Exception
            aform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ORDR Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_MRelateItem" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_MRelateItem" And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddControl(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "U_Z_MItemType" Then
                                    Dim oMatrix As SAPbouiCOM.Matrix
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row) <> "" Then
                                        oCombo = oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific
                                        If oCombo.Selected.Value = "A" Then
                                            PopulateMainItem(oMatrix, pVal.Row, oForm)
                                        End If
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID

                                    Case "_12"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            If oApplication.SBO_Application.MessageBox("Do you want to update allow procurement for selected itemgroup?", , "Yes", "No") = 2 Then
                                                Exit Sub
                                            Else
                                                procurement(oForm)
                                            End If
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

    Private Sub UpdateProcrument(aForm As SAPbouiCOM.Form)

    End Sub

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

    

#End Region

End Class

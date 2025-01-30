using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace SBOAddonProject
{
    [FormAttribute("SBOAddonProject.FormPoint", "FormPoint.b1f")]
    class FormPoint : UserFormBase
    {
        public FormPoint()
        {
            oForm = (SAPbouiCOM.Form)Program.sboApp.Forms.ActiveForm;
            dtPoint = oForm.DataSources.DataTables.Add("dtPoint");
            dtItem = oForm.DataSources.DataTables.Add("dtItem");
            dtVoucher = oForm.DataSources.DataTables.Add("dtVoucher");
            try
            {
                cmbPeriod.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                SAPbobsCOM.Recordset rstPeriod;
                rstPeriod = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rstPeriod.DoQuery("SELECT DISTINCT T0.[U_Period] AS [Period] FROM [dbo].[@GLACCOUNT_H]  T0 WHERE T0.[U_Period] IS NOT NULL ORDER BY T0.[U_Period] DESC");
                if (rstPeriod.RecordCount > 0)
                {
                    rstPeriod.MoveFirst();
                    for (int row = 0; row < rstPeriod.RecordCount; row++)
                    {
                        cmbPeriod.ValidValues.Add(rstPeriod.Fields.Item("Period").Value.ToString(), rstPeriod.Fields.Item("Period").Value.ToString());
                        rstPeriod.MoveNext();
                    }
                    cmbPeriod.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                this.CreateChooseFromListWithFilter(Program.sboApp, oForm);

                oForm.DataSources.UserDataSources.Add("item", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.DataSources.UserDataSources.Add("voucher", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                optItem.DataBind.SetBound(true, Alias: "item");
                optVoucher.DataBind.SetBound(true, Alias: "voucher");
                optVoucher.GroupWith("optItem");
                optItem.Selected = true;
            }
            catch (Exception ex)
            {
            }
        }

        public override void OnInitializeComponent()
        {
            this.lblCCode = ((SAPbouiCOM.StaticText)(this.GetItem("lblCCode").Specific));
            this.lblCName = ((SAPbouiCOM.StaticText)(this.GetItem("lblCName").Specific));
            this.lblPeriod = ((SAPbouiCOM.StaticText)(this.GetItem("lblPeriod").Specific));
            this.lblPoint = ((SAPbouiCOM.StaticText)(this.GetItem("lblPoint").Specific));
            this.lblFNumber = ((SAPbouiCOM.StaticText)(this.GetItem("lblFNumber").Specific));
            this.txtCCode = ((SAPbouiCOM.EditText)(this.GetItem("txtCCode").Specific));
            this.txtCCode.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.txtCCode_ChooseFromListAfter);
            this.txtCName = ((SAPbouiCOM.EditText)(this.GetItem("txtCName").Specific));
            this.txtPoint = ((SAPbouiCOM.EditText)(this.GetItem("txtPoint").Specific));
            this.txtFNumber = ((SAPbouiCOM.EditText)(this.GetItem("txtFNumber").Specific));
            this.lkbCCode = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkbCCode").Specific));
            this.cmbPeriod = ((SAPbouiCOM.ComboBox)(this.GetItem("cmbPeriod").Specific));
            this.grdPoint = ((SAPbouiCOM.Grid)(this.GetItem("grdPoint").Specific));
            this.grdItem = ((SAPbouiCOM.Grid)(this.GetItem("grdItem").Specific));
            this.grdVoucher = ((SAPbouiCOM.Grid)(this.GetItem("grdVoucer").Specific));
            this.optItem = ((SAPbouiCOM.OptionBtn)(this.GetItem("optItem").Specific));
            this.optItem.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.optItem_ClickBefore);
            this.optVoucher = ((SAPbouiCOM.OptionBtn)(this.GetItem("optVoucher").Specific));
            this.optVoucher.ClickBefore += new SAPbouiCOM._IOptionBtnEvents_ClickBeforeEventHandler(this.optVoucher_ClickBefore);
            this.btnCheck = ((SAPbouiCOM.Button)(this.GetItem("btnCheck").Specific));
            this.btnCheck.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnCheck_ClickBefore);
            this.btnRedeem = ((SAPbouiCOM.Button)(this.GetItem("btnRedeem").Specific));
            this.btnRedeem.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.btnRedeem_ClickBefore);
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
        }

        public void CreateChooseFromListWithFilter(SAPbouiCOM.Application app, SAPbouiCOM.Form form)
        {
            try
            {
                // Add DB Data Source
                oForm.DataSources.DBDataSources.Add("OCRD");
                // Add User Data Source
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                // Add DataBind to Edit Text Field
                txtCCode.DataBind.SetBound(true, "OCRD", "CardCode");
                // Create a ChooseFromList collection
                SAPbouiCOM.ChooseFromListCollection cflCollection = form.ChooseFromLists;
                // Create the ChooseFromList creation structure
                SAPbouiCOM.ChooseFromListCreationParams cflParams = (SAPbouiCOM.ChooseFromListCreationParams)app.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                // Set ChooseFromList properties
                cflParams.MultiSelection = false;
                cflParams.ObjectType = "2"; // "2" for Business Partners (OCRD)
                cflParams.UniqueID = "cflOCRD";
                // Add the ChooseFromList to the form
                SAPbouiCOM.ChooseFromList cfl = cflCollection.Add(cflParams);

                SAPbouiCOM.Conditions conditions = (SAPbouiCOM.Conditions)app.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                SAPbouiCOM.Condition condition = conditions.Add();
                condition.Alias = "CardType";
                condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                condition.CondVal = "C";
                condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                SAPbouiCOM.Condition condition2 = conditions.Add();
                condition2.Alias = "validFor";
                condition2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                condition2.CondVal = "Y";
                cfl.SetConditions(conditions);

                // Now the CFL is ready to be linked to an item in the form
                txtCCode.ChooseFromListUID = "cflOCRD";
                txtCCode.ChooseFromListAlias = "CardCode"; // The field to populate
            }
            catch (Exception ex)
            {
                app.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void txtCCode_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                SAPbouiCOM.DataTable dt = cflEvent.SelectedObjects;
                grdPoint.DataTable = null;
                grdItem.DataTable = null;
                grdVoucher.DataTable = null;
                txtPoint.Value = "";
                txtCName.Value = dt.GetValue("CardName", 0).ToString();
                txtCCode.Value = dt.GetValue("CardCode", 0).ToString();
            }
            catch (Exception ex)
            {
            }
        }

        private void btnCheck_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //GRID POINT
                dtPoint.ExecuteQuery("EXEC Q_CHECK_POINT '" + txtCCode.Value + "', '" + cmbPeriod.Value + "'");
                grdPoint.DataTable = dtPoint;
                for (int Index = 0; Index < this.dtPoint.Columns.Count; ++Index)
                {
                    switch (Index)
                    {
                        default:
                            grdPoint.Columns.Item((object)Index).Editable = false;
                            break;
                    }
                }
                double totalSum = 0;
                int columnIndex = -1;
                for (int i = 0; i < grdPoint.Columns.Count; i++)
                {
                    if (grdPoint.Columns.Item(i).TitleObject.Caption == "Point")
                    {
                        columnIndex = i;
                        break;
                    }
                }
                if (columnIndex == -1)
                {
                    Program.sboApp.SetStatusBarMessage("Error : Column Point Not Found", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                for (int i = 0; i < grdPoint.Rows.Count; i++)
                {
                    string cellValue = grdPoint.DataTable.GetValue(columnIndex, i).ToString();
                    double numericValue;
                    if (double.TryParse(cellValue, out numericValue))
                    {
                        totalSum += numericValue;
                    }
                }
                txtPoint.Value = totalSum.ToString();
                //GRID ITEM
                dtItem.ExecuteQuery("SELECT 'N' AS 'Check', T1.U_poin AS 'Point', T1.U_itemcode AS 'Itemcode', T1.U_itemname AS 'Itemname', T1.U_quantity AS 'Available', 1 AS Quantity FROM [@REDEEM_PM] T0 INNER JOIN [@REDEEM_P] T1 ON T0.Code = T1.Code WHERE T0.U_period = '" + cmbPeriod.Value + "' ");
                grdItem.DataTable = dtItem;
                grdItem.Columns.Item((object)"Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                for (int Index = 0; Index < this.dtItem.Columns.Count; ++Index)
                {
                    switch (Index)
                    {
                        case 0:
                            grdItem.Columns.Item((object)Index).Editable = true;
                            break;
                        case 5:
                            grdItem.Columns.Item((object)Index).Editable = true;
                            break;
                        default:
                            grdItem.Columns.Item((object)Index).Editable = false;
                            break;
                    }
                }
                //GRID VOUCHER
                dtVoucher.ExecuteQuery("SELECT 'N' AS 'Check', T1.U_poin AS 'Point', T1.U_itemcode AS 'Itemcode', T1.U_itemname AS 'Itemname', T1.U_diskon AS 'Discount', 1 AS Quantity FROM [@REDEEM_PM] T0 INNER JOIN [@REDEEM_PD] T1 ON T0.Code = T1.Code WHERE T0.U_period = '" + cmbPeriod.Value + "' ");
                grdVoucher.DataTable = dtVoucher;
                grdVoucher.Columns.Item((object)"Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                for (int Index = 0; Index < this.dtVoucher.Columns.Count; ++Index)
                {
                    switch (Index)
                    {
                        case 0:
                            grdVoucher.Columns.Item((object)Index).Editable = true;
                            break;
                        case 5:
                            grdVoucher.Columns.Item((object)Index).Editable = true;
                            break;
                        default:
                            grdVoucher.Columns.Item((object)Index).Editable = false;
                            break;
                    }
                }

                Program.sboApp.SetStatusBarMessage("Success", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                Program.sboApp.SetStatusBarMessage("Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void btnRedeem_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string strTypeTransaction = "";
            if (grdItem.Item.Enabled == true)
            {
                strTypeTransaction = "Sales Order";
            }
            else
            {
                strTypeTransaction = "A/R Credit Memo";
            }
            // Total Point Item
            double totalSumItem = 0;
            int columnIndexItem = -1;
            for (int i = 0; i < grdItem.Columns.Count; i++)
            {
                if (grdItem.Columns.Item(i).TitleObject.Caption == "Point")
                {
                    columnIndexItem = i;
                    break;
                }
            }
            if (columnIndexItem == -1)
            {
                Program.sboApp.SetStatusBarMessage("Error : Column Point Not Found", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            for (int rowIndex = 0; rowIndex < dtItem.Rows.Count; ++rowIndex)
            {
                if (dtItem.GetValue("Check", rowIndex).ToString() == "Y")
                {
                    string cellValue = grdItem.DataTable.GetValue(columnIndexItem, rowIndex).ToString();
                    double numericValue;
                    if (double.TryParse(cellValue, out numericValue))
                    {
                        totalSumItem += numericValue;
                    }
                }
            }
            // Total Point Voucher
            double totalSumVoucher = 0;
            int columnIndexVoucher = -1;
            for (int i = 0; i < grdVoucher.Columns.Count; i++)
            {
                if (grdVoucher.Columns.Item(i).TitleObject.Caption == "Point")
                {
                    columnIndexVoucher = i;
                    break;
                }
            }
            if (columnIndexVoucher == -1)
            {
                Program.sboApp.SetStatusBarMessage("Error : Column Point Not Found", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            for (int rowIndex = 0; rowIndex < dtVoucher.Rows.Count; ++rowIndex)
            {
                if (dtVoucher.GetValue("Check", rowIndex).ToString() == "Y")
                {
                    string cellValue = grdVoucher.DataTable.GetValue(columnIndexVoucher, rowIndex).ToString();
                    double numericValue;
                    if (double.TryParse(cellValue, out numericValue))
                    {
                        totalSumVoucher += numericValue;
                    }
                }
            }
            // Check Point
            if (((totalSumItem <= double.Parse(txtPoint.Value)) || (totalSumVoucher <= double.Parse(txtPoint.Value))) && (double.Parse(txtPoint.Value) > 0))
            {
                // Check Total Point Item Or Total Point Voucher
                if (totalSumItem > 0 || totalSumVoucher > 0)
                {
                    // Check Form Number
                    if (txtFNumber.Value != "")
                    {
                        int ithReturnValue;
                        ithReturnValue = Program.sboApp.MessageBox("Redeem Point to " + strTypeTransaction + " ?", 1, "Continue", "Cancel", "");
                        if (ithReturnValue == 1)
                        {
                            BubbleEvent = true;
                            DateTime now = DateTime.Now;
                            if (grdItem.Item.Enabled == true)
                            {
                                // Sales Order
                                SAPbobsCOM.Documents businessObject = (SAPbobsCOM.Documents)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                businessObject.CardCode = txtCCode.Value;
                                businessObject.CardName = txtCName.Value;
                                businessObject.DocDueDate = now;
                                businessObject.DocDate = now;
                                businessObject.Comments = cmbPeriod.Value + " : " + txtPoint.Value + " Point";
                                businessObject.UserFields.Fields.Item("U_FORMNO").Value = txtFNumber.Value;
                                for (int rowIndex = 0; rowIndex < dtItem.Rows.Count; ++rowIndex)
                                {
                                    string strItemCode = dtItem.GetValue("Itemcode", rowIndex).ToString();
                                    string strItemName = dtItem.GetValue("Itemname", rowIndex).ToString();
                                    string strPoint = dtItem.GetValue("Point", rowIndex).ToString();
                                    string strQuantity = dtItem.GetValue("Quantity", rowIndex).ToString();
                                    if (dtItem.GetValue("Check", rowIndex).ToString() == "Y")
                                    {
                                        businessObject.Lines.ItemCode = strItemCode;
                                        businessObject.Lines.ItemDescription = strItemName;
                                        businessObject.Lines.Quantity = double.Parse(strQuantity);
                                        businessObject.Lines.Price = 0.0;
                                        businessObject.Lines.VatGroup = "PPNO 0";
                                        businessObject.Lines.COGSCostingCode = "199";
                                        businessObject.Lines.WarehouseCode = "WH22";
                                        businessObject.Lines.Add();
                                    }
                                }
                                if (businessObject.Add() != 0)
                                {
                                    Program.sboApp.SetStatusBarMessage("Error : " + Program.oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                }
                                else
                                {
                                    Program.sboApp.MessageBox(strTypeTransaction + " Success Created", 1, "Ok", "", "");
                                }
                            }
                            else
                            {
                                // A/R Credit Memo
                                SAPbobsCOM.Documents businessObject = (SAPbobsCOM.Documents)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                                businessObject.CardCode = txtCCode.Value;
                                businessObject.CardName = txtCName.Value;
                                businessObject.DocDueDate = now;
                                businessObject.DocDate = now;
                                businessObject.Comments = cmbPeriod.Value + " : " + txtPoint.Value + " Point";
                                businessObject.UserFields.Fields.Item("U_FORMNO").Value = txtFNumber.Value;
                                for (int rowIndex = 0; rowIndex < dtVoucher.Rows.Count; ++rowIndex)
                                {
                                    string strItemCode = dtVoucher.GetValue("Itemcode", rowIndex).ToString();
                                    string strItemName = dtVoucher.GetValue("Itemname", rowIndex).ToString();
                                    string strPoint = dtVoucher.GetValue("Point", rowIndex).ToString();
                                    string strQuantity = dtVoucher.GetValue("Quantity", rowIndex).ToString();
                                    if (dtVoucher.GetValue("Check", rowIndex).ToString() == "Y")
                                    {
                                        businessObject.Lines.ItemCode = strItemCode;
                                        businessObject.Lines.ItemDescription = strItemName;
                                        businessObject.Lines.Quantity = double.Parse(strQuantity);
                                        businessObject.Lines.Price = 0.0;
                                        businessObject.Lines.VatGroup = "PPNO 0";
                                        businessObject.Lines.COGSCostingCode = "199";
                                        businessObject.Lines.WarehouseCode = "WH22";
                                        businessObject.Lines.Add();
                                    }
                                }
                                if (businessObject.Add() != 0)
                                {
                                    Program.sboApp.SetStatusBarMessage("Error : " + Program.oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                }
                                else
                                {
                                    Program.sboApp.MessageBox(strTypeTransaction + " Success Created", 1, "Ok", "", "");
                                }
                            }
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        oForm.Items.Item("txtFNumber").Click();
                        Program.sboApp.SetStatusBarMessage("Error : Empty Form Number", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }
                else
                {
                    oForm.Items.Item("txtPoint").Click();
                    Program.sboApp.SetStatusBarMessage("Error : Choose Item Or Voucher", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            else
            {
                oForm.Items.Item("txtPoint").Click();
                Program.sboApp.SetStatusBarMessage("Error : Points Are Not Enough", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void optItem_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            for (int rowIndex = 0; rowIndex < dtVoucher.Rows.Count; ++rowIndex)
            {
                if (dtVoucher.GetValue("Check", rowIndex).ToString() == "Y")
                {
                    dtVoucher.SetValue("Check", rowIndex, "N");
                }
            }
            oForm.Items.Item("txtCCode").Click();
            grdItem.Item.Enabled = true;
            grdVoucher.Item.Enabled = false;
        }

        private void optVoucher_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            for (int rowIndex = 0; rowIndex < dtItem.Rows.Count; ++rowIndex)
            {
                if (dtItem.GetValue("Check", rowIndex).ToString() == "Y")
                {
                    dtItem.SetValue("Check", rowIndex, "N");
                }
            }
            oForm.Items.Item("txtCCode").Click();
            grdVoucher.Item.Enabled = true;
            grdItem.Item.Enabled = false;
        }

        public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.DataTable dtPoint;
        public SAPbouiCOM.DataTable dtItem;
        public SAPbouiCOM.DataTable dtVoucher;
        private SAPbouiCOM.StaticText lblCCode;
        private SAPbouiCOM.StaticText lblCName;
        private SAPbouiCOM.StaticText lblPeriod;
        private SAPbouiCOM.StaticText lblPoint;
        private SAPbouiCOM.StaticText lblFNumber;
        private SAPbouiCOM.EditText txtCCode;
        private SAPbouiCOM.EditText txtCName;
        private SAPbouiCOM.EditText txtPoint;
        private SAPbouiCOM.EditText txtFNumber;
        private SAPbouiCOM.LinkedButton lkbCCode;
        private SAPbouiCOM.ComboBox cmbPeriod;
        private SAPbouiCOM.Grid grdPoint;
        private SAPbouiCOM.Grid grdItem;
        private SAPbouiCOM.Grid grdVoucher;
        private SAPbouiCOM.OptionBtn optItem;
        private SAPbouiCOM.OptionBtn optVoucher;
        private SAPbouiCOM.Button btnCheck;
        private SAPbouiCOM.Button btnRedeem;
    }
}

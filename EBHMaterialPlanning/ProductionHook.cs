using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// SAP forms the add-on listens to:
    ///   - Production Order (form 65211): "Purchase Requests" button beside Cancel, opens the list of the requests raised from it
    ///   - Purchase Request (form 1470000113): after it was updated (closed, cancelled) the Production Orders it refers to are refreshed,
    ///     so the quantity of a cancelled / closed request can be requested again
    /// </summary>
    public static class ProductionHook
    {
        #region Declaretion
        public const string FormProduction = "65211";
        private const string FormRequest = "1470000113", ButtonUID = "EBH_PRQ";
        private static bool IsRunning;
        #endregion

        public static void Start()
        {
            if (IsRunning) return;
            IsRunning = true;
            Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
        }

        #region Production Order button
        /// <summary>Item events of the Production Order form, called by the add-on's one item event handler (PlanForm).</summary>
        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx != FormProduction) return;
            try
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
                    AddButton(Application.SBO_Application.Forms.Item(FormUID));
                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction && pVal.ItemUID == ButtonUID)
                    ButtonPressed(Application.SBO_Application.Forms.Item(FormUID));
            }
            catch (Exception ex)
            {
                PlanForm.Alert(ex.Message);
            }
        }

        /// <summary>"Purchase Requests" button to the right of Cancel (item 2), same height.</summary>
        private static void AddButton(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oCancel = oForm.Items.Item("2");
            SAPbouiCOM.Item oItem = oForm.Items.Add(ButtonUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = oCancel.Left + oCancel.Width + 5;
            oItem.Top = oCancel.Top;
            oItem.Width = oCancel.Width + 40;
            oItem.Height = oCancel.Height;
            oItem.LinkTo = "2";
            ((SAPbouiCOM.Button)oItem.Specific).Caption = "Purchase Requests";
        }

        private static void ButtonPressed(SAPbouiCOM.Form oForm)
        {
            int entry;
            string key = oForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && int.TryParse(key, out entry) && entry > 0)
                LinkedRequestForm.Show(entry);
            else
                Status("Open a saved Production Order first.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
        #endregion

        #region Purchase Request changed
        private static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx != FormRequest || pVal.BeforeAction || !pVal.ActionSuccess) return;
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) return;
            try
            {
                Match m = Regex.Match(pVal.ObjectKey ?? "", @"<DocEntry>\s*(\d+)\s*</DocEntry>");
                if (!m.Success) return;
                int docEntry = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                // Off the SAP event thread: a form event waiting for the DI lock would keep SAP from answering this call
                ThreadPool.QueueUserWorkItem(delegate { Refresh(docEntry); });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("EBH Material Planning purchase request event : " + ex.Message);
            }
        }

        private static void Refresh(int docEntry)
        {
            try
            {
                lock (Db.DiLock)
                {
                    var orders = new List<int>();
                    foreach (var r in Db.Rows("Select Distinct \"" + PlanStore.FieldPdo + "\" as \"Pdo\" from PRQ1 Where \"DocEntry\" = " + docEntry + " and \"" + PlanStore.FieldPdo + "\" > 0"))
                        orders.Add(Convert.ToInt32(r["Pdo"], CultureInfo.InvariantCulture));
                    if (orders.Count > 0) PlanService.Store.Sync(orders);
                }
            }
            catch (Exception ex)
            {
                // Not fatal: the list of the Production Order refreshes it again when it is opened
                System.Diagnostics.Trace.WriteLine("EBH Material Planning refresh after purchase request " + docEntry + " : " + ex.Message);
            }
        }
        #endregion

        private static void Status(string message, SAPbouiCOM.BoStatusBarMessageType type)
        {
            Application.SBO_Application.StatusBar.SetText("EBH Material Planning: " + message, SAPbouiCOM.BoMessageTime.bmt_Short, type);
        }
    }
}

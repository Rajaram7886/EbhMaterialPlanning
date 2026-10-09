using SAPbouiCOM.Framework;
using System;
using ERPBotHubDLL;
using System.Threading;
using System.Collections.Generic;

namespace EBHMaterialPlanning
{
    [FormAttribute("Framework.DBSetup", "DataBaseFrm/DBSetup.b1f")]
    class DBSetup : UserFormBase
    {
        #region Declaretion
        private SAPbouiCOM.Button btnAllTab, btnBaseTab, btnUDF, btnLic, btnRemove;
        public SAPbobsCOM.Company oCompany;
        ERPBotHub_Commonfile EBH = new ERPBotHub_Commonfile();
        #endregion

        #region Standard Code

        #region Text File Read
        private System.Windows.Forms.OpenFileDialog saveFileDialog1;
        private void TextFileRead()
        {
            string selectedPath = "";
            string text = "";
            try
            {
                Thread t = new Thread((ThreadStart)(() =>
                {
                    saveFileDialog1 = new System.Windows.Forms.OpenFileDialog();
                    saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
                    saveFileDialog1.FilterIndex = 2;
                    saveFileDialog1.RestoreDirectory = true;

                    if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        selectedPath = saveFileDialog1.FileName;
                    }
                }));
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();
                if (selectedPath == "") return;
                Program.SBO_App.StatusBar.SetText("Please wait process is runing...!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                EBH.TextFileRead(selectedPath, text);
            }
            catch (Exception ex)
            {
                Program.SBO_App.StatusBar.SetText("Error : License file not uploaded : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Form Event
        public DBSetup() { }
        public override void OnInitializeComponent()
        {
            this.Declare();
            this.Event();
            // SAP may still owe the "database structure has been modified" question from an earlier setup run and
            // asks it as soon as a window opens: answered like the ones during a step
            Watch();
        }
        public override void OnInitializeFormEvents() { }
        #endregion

        #region Declare
        private void Declare()
        {
            try
            {
                oCompany = Program.oCompany;
                this.btnAllTab = ((SAPbouiCOM.Button)(this.GetItem("btnAllTab").Specific));
                this.btnBaseTab = ((SAPbouiCOM.Button)(this.GetItem("btnBaseTab").Specific));
                this.btnUDF = ((SAPbouiCOM.Button)(this.GetItem("btnUDF").Specific));
                this.btnLic = ((SAPbouiCOM.Button)(this.GetItem("btnLic").Specific));
                this.btnRemove = ((SAPbouiCOM.Button)(this.GetItem("btnRemove").Specific));
            }
            catch (Exception ex) { Error("Database Setup form", ex); }
        }
        #endregion

        #region Event
        private void Event()
        {
            try
            {
                btnAllTab.PressedAfter += BtnAllTab_PressedAfter;
                btnBaseTab.PressedAfter += BtnBaseTab_PressedAfter;
                btnLic.PressedAfter += BtnLic_PressedAfter;
                btnUDF.PressedAfter += BtnUDF_PressedAfter;
                btnRemove.PressedAfter += BtnRemove_PressedAfter;
            }
            catch (Exception ex) { Error("Database Setup events", ex); }
        }

        private static void Error(string what, Exception ex)
        {
            Log("ERROR " + what + " : " + ex.Message);
            Program.SBO_App.StatusBar.SetText("Error : " + what + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }

        /// <summary>The status bar cuts long messages: every setup step and error is also written to %TEMP%\EBHMaterialPlanning\dbsetup.log.</summary>
        private static void Log(string text)
        {
            try
            {
                string folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "EBHMaterialPlanning");
                System.IO.Directory.CreateDirectory(folder);
                System.IO.File.AppendAllText(System.IO.Path.Combine(folder, "dbsetup.log"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + text + Environment.NewLine);
            }
            catch (System.IO.IOException ex) { System.Diagnostics.Trace.WriteLine("DBSetup log : " + ex.Message); }
        }

        /// <summary>
        /// Progress on the button. SAP closes every open window (this form too) when the database structure
        /// changes, so the button may be gone: the caption is then simply not shown.
        /// </summary>
        private static void SetCaption(SAPbouiCOM.Button button, string caption)
        {
            try { button.Caption = caption; }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine("DBSetup caption '" + caption + "' : " + ex.Message); }
        }

        private static int Running;
        private static bool Subscribed;
        private static DateTime LastStep = DateTime.MinValue;

        /// <summary>
        /// For a new or removed table / field SAP asks: "The database structure has been modified ... all open windows
        /// will be closed. Do you want to continue ...?" - during the step and once more right after it has finished.
        /// The user already started the setup, so that question (and only that one) is answered with Yes
        /// while a step runs and for one minute afterwards.
        /// </summary>
        private static void ConfirmStructureChange(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx != "0" || pVal.BeforeAction || pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_LOAD) return;
            if (Running == 0 && (DateTime.Now - LastStep).TotalSeconds > 60) return;
            try
            {
                SAPbouiCOM.Form oMsg = Program.SBO_App.Forms.Item(FormUID);
                bool structure = false;
                for (int i = 0; i < oMsg.Items.Count && !structure; i++)
                {
                    SAPbouiCOM.Item oItem = oMsg.Items.Item(i);
                    if (oItem.Type == SAPbouiCOM.BoFormItemTypes.it_STATIC)
                        structure = ((SAPbouiCOM.StaticText)oItem.Specific).Caption.IndexOf("database structure", StringComparison.OrdinalIgnoreCase) >= 0;
                }
                if (structure) oMsg.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine("DBSetup system message : " + ex.Message); }
        }

        private static void Watch()
        {
            if (!Subscribed) { Program.SBO_App.ItemEvent += ConfirmStructureChange; Subscribed = true; }
            LastStep = DateTime.Now;
        }

        private static void Begin()
        {
            Watch();
            Running++;
        }

        private static void End()
        {
            Running--;
            LastStep = DateTime.Now;
        }

        /// <summary>One setup step with the running / done caption on its button. Returns false when the step failed.</summary>
        private bool Step(SAPbouiCOM.Button button, string caption, Action work)
        {
            Begin();
            Log("START " + caption);
            try
            {
                Program.SBO_App.StatusBar.SetText(caption + " : please wait process is runing...!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SetCaption(button, caption + " - runing...");
                work();
                SetCaption(button, caption + " - created.");
                Log("DONE  " + caption);
                Program.SBO_App.StatusBar.SetText(caption + " : completed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return true;
            }
            catch (Exception ex)
            {
                SetCaption(button, caption + " - FAILED");
                Error(caption, ex);
                return false;
            }
            finally
            {
                End();
            }
        }

        private static void Done(bool ok)
        {
            Log(ok ? "FINISHED OK" : "FINISHED WITH ERROR");
            LastStep = DateTime.Now;
            if (ok) Program.SBO_App.MessageBox("Database Setup completed. Restart SAP Business One before using EBH Material Planning.");
            else Program.SBO_App.MessageBox("Database Setup stopped with an error. See the System Messages Log, or the file %TEMP%\\EBHMaterialPlanning\\dbsetup.log for the full text.");
        }

        #region Create All Environment
        private static bool Busy;

        /// <summary>
        /// Runs setup work on its own thread and returns to SAP at once. This matters: SAP shows its "database structure
        /// has been modified" question and refreshes its metadata only when the add-on is not inside a button event.
        /// A step that depends on the previous one (the object needs its tables, the invoice link needs the object)
        /// fails when both run inside the same event.
        /// </summary>
        private void RunAsync(string what, Func<bool> work)
        {
            if (Busy)
            {
                Program.SBO_App.StatusBar.SetText("Database Setup is still running. Please wait for the completion message.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return;
            }
            Busy = true;
            Thread t = new Thread((ThreadStart)(() =>
            {
                bool ok = false;
                try { ok = work(); }
                catch (Exception ex) { Error(what, ex); }
                finally { Busy = false; }
                Done(ok);
            }));
            t.SetApartmentState(ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
        }

        /// <summary>Gives SAP time to ask about the changed structure (answered by ConfirmStructureChange) and to refresh.</summary>
        private static void Settle()
        {
            LastStep = DateTime.Now;
            Thread.Sleep(5000);
        }

        private void BtnAllTab_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            // Order: base table → system UDF. A failed step stops the rest.
            RunAsync("Create All Environment", () =>
            {
                if (!Step(btnBaseTab, "Base Tables", () => EBH.CreateBaseTable(oCompany))) return false;
                Settle();
                return Step(btnUDF, "Create System UDF", SystemTableCreate);
            });
        }
        #endregion

        private void BtnBaseTab_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal) { RunAsync("Base Tables", () => Step(btnBaseTab, "Base Tables", () => EBH.CreateBaseTable(oCompany))); }
        private void BtnLic_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal) { TextFileRead(); }
        private void BtnUDF_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal) { RunAsync("Create System UDF", () => Step(btnUDF, "Create System UDF", SystemTableCreate)); }
        #endregion

        #endregion
        #region Metadata helpers
        /// <summary>True when the user field (alias without U_) exists on the table.</summary>
        private bool FieldExists(string TableID, string Alias)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery("Select \"FieldID\" from CUFD Where \"TableID\"='" + TableID + "' and \"AliasID\"='" + Alias + "'");
                return oRs.RecordCount > 0;
            }
            finally
            {
                if (oRs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
            }
        }

        /// <summary>
        /// SAP refuses to add / remove a user object while another DI metadata object is still referenced
        /// ("Ref count for this object is higher then 0"): released COM wrappers are collected first.
        /// </summary>
        private static void FreeMetadata()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void LastError(string what)
        {
            int ErrCode; string ErrDesc;
            oCompany.GetLastError(out ErrCode, out ErrDesc);
            throw new Exception(what + " : " + ErrDesc + " (" + ErrCode + ")");
        }
        #endregion


        //Modified Below code as per the requirement
        #region System Table Create
        /// <summary>
        /// Fields that tie a Purchase Request line to its Production Order component. Only what is missing is created.
        ///   PRQ1.U_EBH_PRODUCTIONNO  Production Order (DocEntry) the material was fetched from - golden arrow to the Production Order
        ///   PRQ1.U_EBH_PRODLINE      its component line (WOR1.LineNum), so the same item in two orders stays separate
        ///   PRQ1.U_EBH_ValidTill     Valid Until Date of the row
        ///   OPRQ.U_EBH_Vendor / U_EBH_Whse   header vendor and warehouse (not standard Purchase Request header fields)
        ///   WOR1.U_EBH_PRNO / U_EBH_REQQTY   latest open Purchase Request and quantity on open requests, for the component
        /// </summary>
        private void SystemTableCreate()
        {
            #region Purchase Request row
            // EBH_PRODUCTIONNO holds the Production Order DocEntry (needed for the golden arrow); EBH_PRODNO is the
            // same order's user-facing Document Number, plain numeric so it displays without following the link.
            CreateLinkField("PRQ1", "EBH_PRODUCTIONNO", "Production Order No.", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulProductionOrders);
            CreateLinkField("PRQ1", "EBH_PRODLINE", "Production Order Line", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 6, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            CreateLinkField("PRQ1", "EBH_PRODNO", "Production Order Document No.", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            CreateLinkField("PRQ1", "EBH_ValidTill", "Valid Until Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            #endregion

            #region Purchase Request header
            CreateLinkField("OPRQ", "EBH_Vendor", "Vendor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulBusinessPartners);
            CreateLinkField("OPRQ", "EBH_Whse", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            #endregion

            #region Production Order component row
            CreateLinkField("WOR1", "EBH_PRNO", "Purchase Request No.", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            CreateLinkField("WOR1", "EBH_REQQTY", "Open Requested Qty", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            #endregion
        }

        /// <summary>User field, with the link arrow when a system object is given. A refused link (some SAP versions) falls back to the plain field.</summary>
        private void CreateLinkField(string TableName, string Alias, string Description, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int Size, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum LinkedObject)
        {
            if (FieldExists(TableName, Alias)) return;
            string linkError = AddField(TableName, Alias, Description, Type, SubType, Size, LinkedObject);
            if (linkError == null) return;
            if (LinkedObject == SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone) throw new Exception("Field " + TableName + ".U_" + Alias + " : " + linkError);
            string plainError = AddField(TableName, Alias, Description, Type, SubType, Size, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            if (plainError != null) throw new Exception("Field " + TableName + ".U_" + Alias + " : " + plainError);
            Program.SBO_App.StatusBar.SetText("Field: " + TableName + ".U_" + Alias + " was created successfully (without link arrow).", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            Log(TableName + ".U_" + Alias + " created without link : " + linkError);
        }

        /// <summary>Adds one user field; returns null on success, otherwise the SAP error text.</summary>
        private string AddField(string TableName, string Alias, string Description, SAPbobsCOM.BoFieldTypes Type, SAPbobsCOM.BoFldSubTypes SubType, int Size, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum LinkedObject)
        {
            SAPbobsCOM.UserFieldsMD oUDF = null;
            try
            {
                GC.Collect();
                oUDF = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUDF.TableName = TableName;
                oUDF.Name = Alias;
                oUDF.Description = Description;
                oUDF.Type = Type;
                oUDF.SubType = SubType;
                if (Type != SAPbobsCOM.BoFieldTypes.db_Date) oUDF.EditSize = Size;
                if (LinkedObject != SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone) oUDF.LinkedSystemObject = LinkedObject;
                if (oUDF.Add() != 0)
                {
                    int ErrCode; string ErrDesc;
                    oCompany.GetLastError(out ErrCode, out ErrDesc);
                    return ErrDesc + " (" + ErrCode + ")";
                }
                Program.SBO_App.StatusBar.SetText("Field: " + TableName + ".U_" + Alias + " was created successfully..!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return null;
            }
            finally
            {
                if (oUDF != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                oUDF = null;
                GC.Collect();
            }
        }
        #endregion

        #region Remove Add-On Environment
        /// <summary>Everything this add-on creates (the fields above) and only that. ERPBotHub base / license tables are shared with other ERPBotHub add-ons and are never removed.</summary>
        private List<string[]> RemovalList()
        {
            var list = new List<string[]>();     // { table, field }
            foreach (string[] f in new[] { new[] { "PRQ1", "EBH_PRODUCTIONNO" }, new[] { "PRQ1", "EBH_PRODLINE" }, new[] { "PRQ1", "EBH_PRODNO" }, new[] { "PRQ1", "EBH_ValidTill" },
                new[] { "OPRQ", "EBH_Vendor" }, new[] { "OPRQ", "EBH_Whse" }, new[] { "WOR1", "EBH_PRNO" }, new[] { "WOR1", "EBH_REQQTY" } })
                if (FieldExists(f[0], f[1])) list.Add(f);
            return list;
        }

        private void BtnRemove_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                List<string[]> list = RemovalList();
                string udf = string.Join(", ", list.ConvertAll(x => x[0] + ".U_" + x[1]).ToArray());
                if (Program.SBO_App.MessageBox("Remove Add-On Environment will PERMANENTLY DELETE the user fields EBH Material Planning created, with their data " +
                    "(the link between Purchase Request lines and Production Orders):" + Environment.NewLine +
                    "- UDF : " + (udf.Length == 0 ? "(none)" : udf) + Environment.NewLine +
                    "Not removed: ERPBotHub base / license tables." + Environment.NewLine +
                    "Do you want to continue?", 2, "Yes", "No") != 1) return;
                if (Program.SBO_App.MessageBox("This cannot be undone. Are you sure you want to remove the EBH Material Planning environment from company '" + Program.oCompany.CompanyName + "'?", 2, "Yes", "No") != 1) return;

                Program.SBO_App.StatusBar.SetText("Please wait process is runing...!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SetCaption(btnRemove, "Remove Add-On Environment - runing...");
                int ErrCount = 0;
                // SAP asks about the changed database structure for every removal as well: answered like during creation
                Begin();
                Log("START Remove Add-On Environment");
                try { foreach (string[] x in list) if (!RemoveSystemField(x[0], x[1])) ErrCount++; }
                finally { End(); }
                Log("DONE  Remove Add-On Environment, errors " + ErrCount);
                SetCaption(btnRemove, "Remove Add-On Environment - completed.");
                if (ErrCount == 0)
                    Program.SBO_App.StatusBar.SetText("Add-On environment removed successfully...!! Restart SAP Business One.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                else
                    Program.SBO_App.MessageBox("Add-On environment removal completed with " + ErrCount + " error(s). Check the status bar messages log.");
            }
            catch (Exception ex) { Error("Remove Add-On Environment", ex); }
        }

        
        private bool RemoveSystemField(string TableName, string FieldName)
        {
            SAPbobsCOM.UserFieldsMD oUDF = null;
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                GC.Collect();
                int FieldID = -1;
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery("Select \"FieldID\" from CUFD Where \"TableID\"='" + TableName + "' and \"AliasID\"='" + FieldName + "'");
                if (oRs.RecordCount > 0) FieldID = Convert.ToInt32(oRs.Fields.Item("FieldID").Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                oRs = null;

                if (FieldID < 0)
                {
                    Program.SBO_App.StatusBar.SetText("Field: " + TableName + ".U_" + FieldName + " not found - skipped.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                }

                oUDF = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                if (!oUDF.GetByKey(TableName, FieldID))
                {
                    Program.SBO_App.StatusBar.SetText("Field: " + TableName + ".U_" + FieldName + " not found - skipped.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                }
                if (oUDF.Remove() != 0)
                {
                    int ErrCode; string ErrDesc;
                    oCompany.GetLastError(out ErrCode, out ErrDesc);
                    Log("ERROR Field " + TableName + ".U_" + FieldName + " : " + ErrDesc); Program.SBO_App.StatusBar.SetText("Error : Field " + TableName + ".U_" + FieldName + " : " + ErrDesc, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                Program.SBO_App.StatusBar.SetText("Field: " + TableName + ".U_" + FieldName + " was removed successfully..!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return true;
            }
            catch (Exception ex)
            {
                Log("ERROR Field " + TableName + ".U_" + FieldName + " : " + ex.Message); Program.SBO_App.StatusBar.SetText("Error : Field " + TableName + ".U_" + FieldName + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            finally
            {
                if (oRs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                if (oUDF != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDF);
                oUDF = null;
                GC.Collect();
            }
        }
        #endregion
    }
}

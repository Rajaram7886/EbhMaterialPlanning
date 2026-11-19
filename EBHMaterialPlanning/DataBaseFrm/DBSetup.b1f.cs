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
        /// <summary>True when the user field (alias without U_) exists on the table. A custom (user-defined) table's
        /// own fields were found catalogued in CUFD under the "@"-prefixed physical name rather than the plain name
        /// passed to UserFieldsMD.TableName when adding them (a known SAP quirk, version/install-dependent) - both
        /// forms are checked, or an existing field is missed and Add() retries it, failing with "This entry already
        /// exists" (ODBC -2035) instead of being skipped. Harmless no-op for a system table (PRQ1, OPRQ, ...), whose
        /// name never starts with "@" either way.</summary>
        private bool FieldExists(string TableID, string Alias)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string alt = TableID.StartsWith("@") ? TableID.Substring(1) : "@" + TableID;
                oRs.DoQuery("Select \"FieldID\" from CUFD Where (\"TableID\"='" + TableID + "' or \"TableID\"='" + alt + "') and \"AliasID\"='" + Alias + "'");
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
        ///   PRQ1.U_EBH_PRODLINE      component line(s) of that same order (WOR1.LineNum), comma-separated when the line was
        ///                            consolidated from more than one - see ConvertProdLineFieldToText
        ///   PRQ1.U_EBH_ValidTill     Valid Until Date of the row
        ///   OPRQ.U_EBH_Vendor / U_EBH_Whse   header vendor and warehouse (not standard Purchase Request header fields)
        ///   WOR1.U_EBH_PRNO / U_EBH_REQQTY   latest open Purchase Request and quantity on open requests, for the component
        ///   @EBH_PRQMAP              Purchase Request line to Production Order component map - see CreateMappingTable,
        ///                            the field every "already requested" / progress / Linked Documents lookup reads
        ///
        /// U_EBH_PRODNO (Production Order Document No.) and U_EBH_PDOMAP (the old, fixed-size-text predecessor of
        /// @EBH_PRQMAP) are deliberately NOT created any more: PRODNO was only a display convenience (the same
        /// number is one click away via the EBH_PRODUCTIONNO golden arrow) and PDOMAP's job is now done properly by
        /// @EBH_PRQMAP. Both were also seen to fail on at least one company with "Data too large for the field",
        /// which - before the fix below - aborted every field after them, including @EBH_PRQMAP itself.
        ///
        /// Each field/table below is created independently (see CreateLinkField / CreateMappingTable): one failing
        /// does not stop the rest, and every failure is collected into one error reported at the end, so a single
        /// bad field never again blocks fields that have nothing to do with it.
        /// </summary>
        private void SystemTableCreate()
        {
            var errors = new List<string>();
            Action<string, Action> step = (what, work) =>
            {
                try { work(); }
                catch (Exception ex) { errors.Add(what + " : " + ex.Message); Log("ERROR " + what + " : " + ex.Message); }
            };

            #region Production Order component map (the real tracking table)
            // Created first, and filled from the existing Purchase Request lines (PlanStore.BackfillLegacyMap) BEFORE
            // U_EBH_PRODLINE is converted below - the conversion erases the old line numbers the backfill reads.
            bool mapReady = false;
            step("@EBH_PRQMAP", () => { CreateMappingTable(); mapReady = true; });
            if (mapReady)
                step("@EBH_PRQMAP backfill", () =>
                {
                    FreeMetadata();
                    int added = PlanStore.BackfillLegacyMap();
                    Log("@EBH_PRQMAP backfill : " + added + " existing Purchase Request line(s) copied");
                });
            #endregion

            #region Purchase Request row
            // EBH_PRODUCTIONNO holds the Production Order DocEntry, needed for the golden arrow.
            step("PRQ1.U_EBH_PRODUCTIONNO", () => CreateLinkField("PRQ1", "EBH_PRODUCTIONNO", "Production Order No.", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulProductionOrders));
            // Was a single Numeric line number; now holds every merged component's own line number that belongs to
            // the SAME order as EBH_PRODUCTIONNO above, comma-separated ("3,4,7") - a component from a DIFFERENT
            // order cannot be named here unambiguously (which order is line 7 from?) and is tracked only through
            // @EBH_PRQMAP, the table every lookup in this add-on actually reads. Converting an existing Numeric
            // field to text replaces it - see ConvertProdLineFieldToText. Never attempted unless the map table is
            // in place and backfilled, so the old line numbers are never lost without having been copied first.
            if (mapReady) step("PRQ1.U_EBH_PRODLINE", ConvertProdLineFieldToText);
            else errors.Add("PRQ1.U_EBH_PRODLINE : not converted, because @EBH_PRQMAP could not be set up (its existing line numbers would otherwise be lost).");
            step("PRQ1.U_EBH_ValidTill", () => CreateLinkField("PRQ1", "EBH_ValidTill", "Valid Until Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            #endregion

            #region Purchase Request header
            step("OPRQ.U_EBH_Vendor", () => CreateLinkField("OPRQ", "EBH_Vendor", "Vendor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulBusinessPartners));
            step("OPRQ.U_EBH_Whse", () => CreateLinkField("OPRQ", "EBH_Whse", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            #endregion

            #region Remark written by the add-on on the documents it creates (instead of the standard Comments / Remarks)
            step("OPRQ.U_EBH_REMARK", () => CreateLinkField("OPRQ", "EBH_REMARK", "EBH Remark", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("OPOR.U_EBH_REMARK", () => CreateLinkField("OPOR", "EBH_REMARK", "EBH Remark", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            #endregion

            #region Production Order component row
            step("WOR1.U_EBH_PRNO", () => CreateLinkField("WOR1", "EBH_PRNO", "Purchase Request No.", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("WOR1.U_EBH_REQQTY", () => CreateLinkField("WOR1", "EBH_REQQTY", "Open Requested Qty", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            // Manually priced components: when set on the component row, Material Request Planning loads these
            // instead of defaulting Price / Discount to 0
            step("WOR1.U_EBH_Unit_Price", () => CreateLinkField("WOR1", "EBH_Unit_Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("WOR1.U_EBH_Discount", () => CreateLinkField("WOR1", "EBH_Discount", "Discount %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 10, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("WOR1.U_EBH_Line_Total", () => CreateLinkField("WOR1", "EBH_Line_Total", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            #endregion

            #region Manufacturer catalog number: OITM is the source, PRQ1 / POR1 carry a copy onto each line
            step("OITM.U_EBH_MfgCatlog", () => CreateLinkField("OITM", "EBH_MfgCatlog", "Manufacturer Catalog No.", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("PRQ1.U_EBH_CATLOG", () => CreateLinkField("PRQ1", "EBH_CATLOG", "Manufacturer Catalog No.", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("POR1.U_EBH_CATLOG", () => CreateLinkField("POR1", "EBH_CATLOG", "Manufacturer Catalog No.", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            #endregion

            #region Job code: OITB flags whether its Item Group tracks one; when it does, the Production Order's own
            // product (OWOR.ItemCode) is copied onto every PRQ1 / POR1 line raised for its components
            step("OITB.U_EBH_JOBCODNO", () => CreateLinkField("OITB", "EBH_JOBCODNO", "Track Job Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Checkbox, 1, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("PRQ1.U_EBH_JOBCODNO", () => CreateLinkField("PRQ1", "EBH_JOBCODNO", "Job Code (Product No.)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulItems));
            step("POR1.U_EBH_JOBCODNO", () => CreateLinkField("POR1", "EBH_JOBCODNO", "Job Code (Product No.)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulItems));
            #endregion

            #region Planned price (marketing document rows): the Purchase Request carries the planning grid's Price After
            // Discount / Total here, not in the standard Price / Discount / Line Total, which stay blank on the request;
            // the Purchase Order Planning grid caps its own Price After Discount at this value
            step("PRQ1.U_EBH_PRICEAFTERDISCOUNT", () => CreateLinkField("PRQ1", "EBH_PRICEAFTERDISCOUNT", "Price After Discount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("POR1.U_EBH_PRICEAFTERDISCOUNT", () => CreateLinkField("POR1", "EBH_PRICEAFTERDISCOUNT", "Price After Discount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("PRQ1.U_EBH_LINETOTAL", () => CreateLinkField("PRQ1", "EBH_LINETOTAL", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            step("POR1.U_EBH_LINETOTAL", () => CreateLinkField("POR1", "EBH_LINETOTAL", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 15, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            #endregion

            if (errors.Count > 0) throw new Exception(errors.Count + " field(s)/table(s) failed - " + string.Join(" | ", errors.ToArray()));
        }

        /// <summary>True when a user table (its plain name, without the "@") already exists.</summary>
        private bool TableExists(string tableName)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery("Select \"TableName\" from OUTB Where \"TableName\"=" + "'" + tableName + "'");
                return oRs.RecordCount > 0;
            }
            finally { if (oRs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs); }
        }

        /// <summary>
        /// The real source of truth for "which Production Order component(s) does this Purchase Request line cover,
        /// and with how much quantity" - one row per component, however many components a line was consolidated
        /// from (Material Request Planning's merged rows). Unlike the fixed-size text fields this replaces
        /// (U_EBH_PDOMAP, still written for human reference only - see SystemTableCreate), a table has no size limit
        /// to silently truncate past, and its numeric columns cannot be misread the way a comma-separated number
        /// list can. Every "already requested" / progress / Linked Documents lookup in this add-on reads this table.
        /// </summary>
        private void CreateMappingTable()
        {
            const string table = "EBH_PRQMAP";
            if (!TableExists(table))
            {
                SAPbobsCOM.UserTablesMD oTable = null;
                try
                {
                    oTable = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                    oTable.TableName = table;
                    // SAP caps TableDescription at 30 characters - the full "Purchase Request Line - Production
                    // Order Component Map" (57 chars) was rejected outright with "Data too large for the field",
                    // the real cause behind every @EBH_PRQMAP failure so far (not any of its UDFs, all of which
                    // already have their own short, safe sizes).
                    oTable.TableDescription = "PR Line - PDO Map";
                    oTable.TableType = SAPbobsCOM.BoUTBTableType.bott_NoObject;
                    if (oTable.Add() != 0) LastError("Table " + table);
                }
                finally { if (oTable != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oTable); }
                FreeMetadata();
            }
            var fieldErrors = new List<string>();
            Action<string, Action> field = (what, work) =>
            {
                try { work(); }
                catch (Exception ex) { fieldErrors.Add(what + " : " + ex.Message); Log("ERROR " + what + " : " + ex.Message); }
            };
            field(table + ".U_PrEntry", () => CreateLinkField(table, "PrEntry", "Purchase Request DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            field(table + ".U_PrLine", () => CreateLinkField(table, "PrLine", "Purchase Request Line", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 6, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            field(table + ".U_PdoEntry", () => CreateLinkField(table, "PdoEntry", "Production Order DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            field(table + ".U_PdoLine", () => CreateLinkField(table, "PdoLine", "Production Order Line", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 6, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            // Float/Quantity UDFs in this company only proved reliable at size 10 (see WOR1.EBH_REQQTY) - size 19
            // failed outright with "Data too large for the field", so this matches that known-good precedent rather
            // than a guessed-larger one.
            field(table + ".U_Qty", () => CreateLinkField(table, "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 10, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            // Reference only, written once at Create time (see PlanStore.InsertMapRow) - a status the user sees when
            // inspecting this table directly, matching P0."DocStatus" ('O' Open / 'C' Closed) with 'X' for a
            // cancelled document. Every actual "already requested" / progress lookup still reads the LIVE status
            // from PRQ1 / OPRQ via the join, not this column, so a Purchase Request cancelled or closed afterwards
            // is still picked up correctly without this field ever needing to be kept in sync.
            field(table + ".U_PrStatus", () => CreateLinkField(table, "PrStatus", "Purchase Request Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone));
            if (fieldErrors.Count > 0) throw new Exception(string.Join(" | ", fieldErrors.ToArray()));
        }

        /// <summary>
        /// PRQ1.U_EBH_PRODLINE started as a single Numeric line number (one Production Order component per line) and
        /// is now a comma-separated Alpha list (several components can share one line after consolidation - see
        /// Material Request Planning's merged rows). A missing field is simply created as Alpha. An existing one
        /// cannot be changed in place (SAP's UDF metadata has no "alter type"): it has to be removed and re-added,
        /// which erases whatever line number is already stored on every existing Purchase Request line - their real
        /// traceability is unaffected, since every lookup in this add-on reads the @EBH_PRQMAP table (see
        /// CreateMappingTable), not this field. Confirmed with the user (by MessageBox) every time this runs against
        /// a field that still exists, so a routine re-run of Create System UDF never silently drops data the user
        /// did not mean to replace.
        /// </summary>
        private void ConvertProdLineFieldToText()
        {
            const string table = "PRQ1", alias = "EBH_PRODLINE";
            if (!FieldExists(table, alias))
            {
                string addErr = AddField(table, alias, "Production Order Line(s)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
                if (addErr != null) throw new Exception("Field " + table + ".U_" + alias + " : " + addErr);
                return;
            }
            if (Program.SBO_App.MessageBox(
                table + ".U_" + alias + " is being converted from a single Production Order line number to a comma-separated list. " +
                "This REPLACES the field: whatever line number is already stored on existing Purchase Request lines will be lost. " +
                "Their Production Order links have just been copied into the @EBH_PRQMAP table, which is what the add-on reads, so they are kept there. Continue?",
                2, "Yes", "No") != 1)
            {
                Log("SKIPPED converting " + table + ".U_" + alias + " - user declined");
                return;
            }
            if (!RemoveSystemField(table, alias)) throw new Exception("Could not remove the old " + table + ".U_" + alias + " field.");
            FreeMetadata();
            string err = AddField(table, alias, "Production Order Line(s)", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone);
            if (err != null) throw new Exception("Field " + table + ".U_" + alias + " : " + err);
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
        /// <summary>Everything this add-on creates (the fields above) and only that. ERPBotHub base / license tables are shared with other ERPBotHub add-ons and are never removed.
        /// EBH_PRODNO and EBH_PDOMAP are no longer created (see SystemTableCreate) but are still listed here so a
        /// company that has them from an earlier run of this add-on can still have them cleaned up.</summary>
        private List<string[]> RemovalList()
        {
            var list = new List<string[]>();     // { table, field }
            foreach (string[] f in new[] { new[] { "PRQ1", "EBH_PRODUCTIONNO" }, new[] { "PRQ1", "EBH_PRODLINE" }, new[] { "PRQ1", "EBH_PRODNO" }, new[] { "PRQ1", "EBH_ValidTill" }, new[] { "PRQ1", "EBH_PDOMAP" },
                new[] { "OPRQ", "EBH_Vendor" }, new[] { "OPRQ", "EBH_Whse" }, new[] { "OPRQ", "EBH_REMARK" }, new[] { "OPOR", "EBH_REMARK" }, new[] { "WOR1", "EBH_PRNO" }, new[] { "WOR1", "EBH_REQQTY" },
                new[] { "WOR1", "EBH_Unit_Price" }, new[] { "WOR1", "EBH_Discount" }, new[] { "WOR1", "EBH_Line_Total" },
                new[] { "OITM", "EBH_MfgCatlog" }, new[] { "PRQ1", "EBH_CATLOG" }, new[] { "POR1", "EBH_CATLOG" },
                new[] { "OITB", "EBH_JOBCODNO" }, new[] { "PRQ1", "EBH_JOBCODNO" }, new[] { "POR1", "EBH_JOBCODNO" },
                new[] { "PRQ1", "EBH_PRICEAFTERDISCOUNT" }, new[] { "POR1", "EBH_PRICEAFTERDISCOUNT" }, new[] { "PRQ1", "EBH_LINETOTAL" }, new[] { "POR1", "EBH_LINETOTAL" },
                new[] { "EBH_PRQMAP", "PrEntry" }, new[] { "EBH_PRQMAP", "PrLine" }, new[] { "EBH_PRQMAP", "PdoEntry" }, new[] { "EBH_PRQMAP", "PdoLine" }, new[] { "EBH_PRQMAP", "Qty" }, new[] { "EBH_PRQMAP", "PrStatus" } })
                if (FieldExists(f[0], f[1])) list.Add(f);
            return list;
        }

        /// <summary>Removes the @EBH_PRQMAP table itself, with its data - only once its own fields are gone (SAP
        /// refuses to drop a table that still has user fields on it).</summary>
        private bool RemoveMappingTable()
        {
            const string table = "EBH_PRQMAP";
            if (!TableExists(table)) return true;
            SAPbobsCOM.UserTablesMD oTable = null;
            try
            {
                oTable = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (!oTable.GetByKey(table))
                {
                    Program.SBO_App.StatusBar.SetText("Table: " + table + " not found - skipped.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                }
                if (oTable.Remove() != 0)
                {
                    int ErrCode; string ErrDesc;
                    oCompany.GetLastError(out ErrCode, out ErrDesc);
                    Log("ERROR Table " + table + " : " + ErrDesc);
                    Program.SBO_App.StatusBar.SetText("Error : Table " + table + " : " + ErrDesc, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                Program.SBO_App.StatusBar.SetText("Table: " + table + " was removed successfully..!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return true;
            }
            catch (Exception ex)
            {
                Log("ERROR Table " + table + " : " + ex.Message);
                Program.SBO_App.StatusBar.SetText("Error : Table " + table + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            finally
            {
                if (oTable != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oTable);
                GC.Collect();
            }
        }

        private void BtnRemove_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                List<string[]> list = RemovalList();
                bool hasMapTable = TableExists("EBH_PRQMAP");
                string udf = string.Join(", ", list.ConvertAll(x => x[0] + ".U_" + x[1]).ToArray());
                if (Program.SBO_App.MessageBox("Remove Add-On Environment will PERMANENTLY DELETE the user fields EBH Material Planning created, with their data " +
                    "(the link between Purchase Request lines and Production Orders):" + Environment.NewLine +
                    "- UDF : " + (udf.Length == 0 ? "(none)" : udf) + Environment.NewLine +
                    (hasMapTable ? "- Table : @EBH_PRQMAP (Production Order component map)" + Environment.NewLine : "") +
                    "Not removed: ERPBotHub base / license tables." + Environment.NewLine +
                    "Do you want to continue?", 2, "Yes", "No") != 1) return;
                if (Program.SBO_App.MessageBox("This cannot be undone. Are you sure you want to remove the EBH Material Planning environment from company '" + Program.oCompany.CompanyName + "'?", 2, "Yes", "No") != 1) return;

                Program.SBO_App.StatusBar.SetText("Please wait process is runing...!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SetCaption(btnRemove, "Remove Add-On Environment - runing...");
                int ErrCount = 0;
                // SAP asks about the changed database structure for every removal as well: answered like during creation
                Begin();
                Log("START Remove Add-On Environment");
                try
                {
                    foreach (string[] x in list) if (!RemoveSystemField(x[0], x[1])) ErrCount++;
                    // The table can only go once its own fields are gone (SAP refuses to drop a table still carrying
                    // user fields), so this runs last.
                    if (hasMapTable) { FreeMetadata(); if (!RemoveMappingTable()) ErrCount++; }
                }
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
                // A custom table's own fields can be catalogued under the "@"-prefixed physical name rather than the
                // plain name (see FieldExists) - both are checked, and the TableID actually stored is read back and
                // used for GetByKey below, since that call needs the exact form CUFD has on record, not necessarily
                // the one this method was called with.
                string actualTableId = TableName;
                string alt = TableName.StartsWith("@") ? TableName.Substring(1) : "@" + TableName;
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery("Select \"FieldID\", \"TableID\" from CUFD Where (\"TableID\"='" + TableName + "' or \"TableID\"='" + alt + "') and \"AliasID\"='" + FieldName + "'");
                if (oRs.RecordCount > 0)
                {
                    FieldID = Convert.ToInt32(oRs.Fields.Item("FieldID").Value);
                    actualTableId = Convert.ToString(oRs.Fields.Item("TableID").Value).Trim();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                oRs = null;

                if (FieldID < 0)
                {
                    Program.SBO_App.StatusBar.SetText("Field: " + TableName + ".U_" + FieldName + " not found - skipped.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                }

                oUDF = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                if (!oUDF.GetByKey(actualTableId, FieldID))
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

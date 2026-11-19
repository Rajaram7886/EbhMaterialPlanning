using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Goods Receipt PO (form 143): a "Bin Auto Allocations" button beside Cancel. One press reads the unsaved GRPO,
    /// allocates every row's full quantity to its project bin "&lt;Warehouse&gt;-&lt;Project&gt;" (BinAllocationService)
    /// and adds (posts) the GRPO through the DI API (GrpoDraftService). SAP's Bin Location Allocation window is never
    /// opened. On success the unsaved GRPO form is closed and the posted GRPO opened.
    /// </summary>
    public static class GrpoHook
    {
        public const string FormGrpo = "143";
        private const string ButtonAutoUID = "EBHBAUTO";
        /// <summary>Goods Receipt PO object (OPDN) - opens the posted GRPO.</summary>
        private const int GrpoObjectType = 20;

        /// <summary>GRPO FormUID -> GRPO number already posted from it, so a second press cannot post it twice.</summary>
        private static readonly Dictionary<string, int> Posted = new Dictionary<string, int>();

        private static void Status(string message, SAPbouiCOM.BoStatusBarMessageType type)
        {
            Program.SBO_App.StatusBar.SetText("EBH Material Planning: " + message, SAPbouiCOM.BoMessageTime.bmt_Short, type);
        }

        private static void Log(string text)
        {
            System.Diagnostics.Trace.WriteLine("EBH Material Planning GRPO : " + text);
        }

        #region Signals (SAP event threads -> main-thread pump)
        /// <summary>
        /// SAP delivers the events a UI API call causes on another thread and waits for them (see PlanForm). An event
        /// that waited for Db.DiLock while the lock holder was inside a UI call froze SAP (found 2026-10-30). So an
        /// event only posts a signal - it never waits and never calls the UI API - and this timer on the add-on's main
        /// thread does all the work, outside any SAP event.
        /// </summary>
        public static void Start()
        {
            pump = new System.Windows.Forms.Timer { Interval = 50 };
            pump.Tick += Pump;
            pump.Start();
        }

        private enum Sig { GrpoLoad, GrpoPlace, GrpoUnload, PressAuto }

        private struct Signal
        {
            public Sig Kind;
            public string FormUID;
            /// <summary>PressAuto only: the user already answered Yes (a retry after a busy DI lock must not ask again).</summary>
            public bool Confirmed;
        }

        /// <summary>Guards Signals only; never held across a UI API call. Everything else is touched by the pump alone.</summary>
        private static readonly object SignalLock = new object();
        private static List<Signal> Signals = new List<Signal>();
        private static System.Windows.Forms.Timer pump;
        /// <summary>A UI / DI call inside the pump can let the message loop run another Tick; that one is skipped.</summary>
        private static bool pumping;

        private static void Post(Sig kind, string formUid, bool confirmed = false)
        {
            lock (SignalLock) Signals.Add(new Signal { Kind = kind, FormUID = formUid, Confirmed = confirmed });
        }

        /// <summary>Called by PlanForm's dispatcher for form 143. Only records the event - must not block.</summary>
        public static void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.BeforeAction) return;
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: Post(Sig.GrpoLoad, FormUID); break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: Post(Sig.GrpoPlace, FormUID); break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: Post(Sig.GrpoUnload, FormUID); break;
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    if (pVal.ItemUID == ButtonAutoUID) Post(Sig.PressAuto, FormUID);
                    break;
            }
        }

        private static void Pump(object sender, EventArgs e)
        {
            if (pumping) return;
            pumping = true;
            try
            {
                List<Signal> batch;
                lock (SignalLock)
                {
                    if (Signals.Count == 0) return;
                    batch = Signals;
                    Signals = new List<Signal>();
                }
                // Signals raised while this batch runs (by our own UI calls) wait for the next Tick.
                bool pressHandled = false;
                foreach (Signal s in batch)
                {
                    try
                    {
                        switch (s.Kind)
                        {
                            case Sig.GrpoLoad: AddButton(Program.SBO_App.Forms.Item(s.FormUID)); break;
                            case Sig.GrpoPlace: PlaceButton(Program.SBO_App.Forms.Item(s.FormUID)); break;
                            case Sig.GrpoUnload: Posted.Remove(s.FormUID); break;
                            case Sig.PressAuto:
                                if (pressHandled) break;        // a double click queued two presses: one run is enough
                                pressHandled = true;
                                // The DI lock is busy (another add-on action is querying): try again on the next Tick rather than wait here.
                                bool confirmed = s.Confirmed;
                                if (!Run(s.FormUID, ref confirmed)) Post(s.Kind, s.FormUID, confirmed);
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log(s.Kind + " " + s.FormUID + " : " + ex);
                        if (s.Kind == Sig.PressAuto) PlanForm.Alert("Bin Auto Allocations stopped: " + ex.Message);
                    }
                }
            }
            finally { pumping = false; }
        }
        #endregion

        #region Button
        private static void AddButton(SAPbouiCOM.Form form)
        {
            try { form.Items.Item(ButtonAutoUID); }
            catch (COMException)
            {
                try
                {
                    SAPbouiCOM.Item it = form.Items.Add(ButtonAutoUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    ((SAPbouiCOM.Button)it.Specific).Caption = "Bin Auto Allocations";
                }
                catch (COMException ex) { Log("could not add button : " + ex.Message); return; }
            }
            PlaceButton(form);
        }

        /// <summary>Directly after Cancel. SAP re-flows its footer after load (Copy From / Copy To), so this runs on every resize / activate.</summary>
        private static void PlaceButton(SAPbouiCOM.Form form)
        {
            try
            {
                SAPbouiCOM.Item oCancel = form.Items.Item("2");
                SAPbouiCOM.Item it = form.Items.Item(ButtonAutoUID);
                it.Left = oCancel.Left + oCancel.Width + 5; it.Top = oCancel.Top; it.Width = oCancel.Width + 70; it.Height = oCancel.Height;
                it.Enabled = form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            }
            catch (COMException ex) { Log("could not place button : " + ex.Message); }
        }
        #endregion

        #region Bin Auto Allocations -> posted GRPO
        /// <summary>False only when the DI lock is busy and nothing was done - the caller retries (confirmed tells it the user already said Yes).</summary>
        private static bool Run(string formUid, ref bool confirmed)
        {
            SAPbouiCOM.Form form;
            try { form = Program.SBO_App.Forms.Item(formUid); }
            catch (COMException) { return true; }       // the GRPO was closed meanwhile

            int existing;
            if (Posted.TryGetValue(formUid, out existing))
            {
                PlanForm.Alert("Goods Receipt PO " + existing + " was already added from this form. Close this form instead of adding it again.");
                return true;
            }
            if (form.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                PlanForm.Alert("Bin Auto Allocations works on a new Goods Receipt PO that has not been added yet.");
                return true;
            }

            if (!confirmed)
            {
                if (Program.SBO_App.MessageBox("Are you sure you want to auto allocate the Bin Location quantities?" + Environment.NewLine +
                    "Every row's full quantity will be allocated to its project bin" + Environment.NewLine +
                    "and the Goods Receipt PO will be added.", 2, "Yes", "No") != 1)
                {
                    Status("Bin Auto Allocations cancelled. Nothing was allocated or added.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                }
                confirmed = true;
            }

            Status("Starting automatic project-based bin allocation...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            var clock = System.Diagnostics.Stopwatch.StartNew();
            var phases = new List<string>();
            GrpoDraft draft = GrpoDraftService.PrepareDraft(form);      // UI reads only - outside the DI lock
            phases.Add("read form " + clock.ElapsedMilliseconds + " ms");
            if (!draft.BaseLinkRead)
            {
                // Posting rows without their Purchase Order link would leave the PO open and allow receiving it twice
                PlanForm.Alert("Goods Receipt PO not added: its rows' link to the Purchase Order could not be read off the form. Add it the standard way.");
                return true;
            }

            if (!System.Threading.Monitor.TryEnter(Db.DiLock)) return false;
            CreatedDocument created;
            try
            {
                Status("Processing " + draft.Lines.Count + " GRPO line(s)...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                List<string> errors = BinAllocationService.ApplyBinAllocation(draft.Lines);
                phases.Add("bins " + clock.ElapsedMilliseconds + " ms");
                if (errors.Count > 0)
                {
                    Fail("Goods Receipt PO not added. Fix these row(s) and press Bin Auto Allocations again:", errors);
                    return true;
                }

                Status("Adding Goods Receipt PO...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                created = GrpoDraftService.AddGrpo(draft);      // verified and committed, or rolled back and thrown
                phases.Add("add + verify " + clock.ElapsedMilliseconds + " ms");
                Log("GRPO " + created.DocEntry + " / " + created.DocNum + " added from " + formUid);
            }
            finally { System.Threading.Monitor.Exit(Db.DiLock); }

            Posted[formUid] = created.DocNum;
            Log("timing (cumulative) : " + string.Join(", ", phases.ToArray()));
            string seconds = (clock.ElapsedMilliseconds / 1000.0).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
            int allocated = draft.Lines.FindAll(l => l.NeedsBin).Count;
            Status("Goods Receipt PO " + created.DocNum + " added successfully (" + allocated + " row(s) allocated to project bins) in " + seconds + " s.",
                SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            CloseAndOpenGrpo(form, created);
            return true;
        }

        private static void Fail(string title, List<string> errors)
        {
            Log(title + " " + string.Join(" | ", errors.ToArray()));
            PlanForm.Alert(title + Environment.NewLine + string.Join(Environment.NewLine, errors.ToArray()));
        }

        /// <summary>The GRPO is posted: close the unsaved form (SAP may ask to confirm losing its unsaved data) and show the posted GRPO.</summary>
        private static void CloseAndOpenGrpo(SAPbouiCOM.Form form, CreatedDocument created)
        {
            try { form.Close(); }
            catch (COMException ex) { Log("could not close the GRPO : " + ex.Message); }
            try { Program.SBO_App.OpenForm((SAPbouiCOM.BoFormObjectEnum)GrpoObjectType, "", created.DocEntry.ToString(System.Globalization.CultureInfo.InvariantCulture)); }
            catch (COMException ex)
            {
                Log("could not open GRPO " + created.DocEntry + " : " + ex.Message);
                Status("Goods Receipt PO " + created.DocNum + " added. Open it from Purchasing > Goods Receipt PO.", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
        }
        #endregion
    }
}

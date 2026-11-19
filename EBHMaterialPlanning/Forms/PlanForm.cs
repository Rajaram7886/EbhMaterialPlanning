using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Base of the add-on's native SAP Business One forms. The forms are built in code with the UI API
    /// (no .b1f), one object per open window; item events are routed here by FormUID.
    /// Item and data source ids are the same and at most 10 characters.
    /// </summary>
    internal abstract class PlanForm
    {
        #region Routing
        private static readonly Dictionary<string, PlanForm> Open = new Dictionary<string, PlanForm>();
        private static bool IsRunning;
        private static int Counter;
        protected static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        protected static SAPbouiCOM.Application App { get { return Program.SBO_App; } }
        protected static PlanStore Store { get { return PlanService.Store; } }

        public static void Start()
        {
            if (IsRunning) return;
            IsRunning = true;
            App.ItemEvent += SBO_Application_ItemEvent;
        }

        // SAP delivers events on several threads. While one thread works on a form (and waits for SAP inside a UI API
        // call), SAP sends the events that call causes (activate, resize, focus) on another thread and waits for them.
        // Such an event must never wait for the working thread: it is skipped. Found 2026-10-04 as a dead lock on the
        // first form that was opened.
        private static readonly object Gate = new object();
        private static int busyThread, busyDepth;

        /// <summary>False when another thread is inside a form operation.</summary>
        private static bool Begin()
        {
            int me = System.Threading.Thread.CurrentThread.ManagedThreadId;
            lock (Gate)
            {
                if (busyThread != 0 && busyThread != me) return false;
                busyThread = me;
                busyDepth++;
                return true;
            }
        }

        private static void End()
        {
            lock (Gate)
            {
                if (--busyDepth == 0) busyThread = 0;
            }
        }

        /// <summary>Support diagnostics: EBH_MRP_TRACE=1 writes the events of the add-on forms to %TEMP%\EBHMaterialPlanning\forms.log.</summary>
        private static readonly bool TraceOn = Environment.GetEnvironmentVariable("EBH_MRP_TRACE") == "1";

        private static void Trace(string text)
        {
            try
            {
                string folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "EBHMaterialPlanning");
                System.IO.Directory.CreateDirectory(folder);
                lock (Gate) System.IO.File.AppendAllText(System.IO.Path.Combine(folder, "forms.log"), DateTime.Now.ToString("HH:mm:ss.fff") + "  " + text + Environment.NewLine);
            }
            catch (System.IO.IOException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning log : " + ex.Message); }
        }

        private static PlanForm Find(string uid)
        {
            PlanForm f;
            lock (Open) return Open.TryGetValue(uid, out f) ? f : null;
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try { SBO_Application_ItemEvent_Inner(FormUID, ref pVal, ref BubbleEvent); }
            catch (Exception ex)
            {
                // SAP occasionally hands over an ItemEvent whose COM proxy is already gone - seen live as
                // System.InvalidCastException (HRESULT RPC_E_SYS_CALL_FAILED) reading pVal.FormTypeEx itself, before
                // any of this add-on's own code runs. A transient RPC hiccup, not something retrying this one event
                // could fix. Previously uncaught here (every branch below was only ever wrapped starting from
                // Find(FormUID) onwards), so one bad event took the whole dispatcher down - and with it every open
                // add-on form and both system-form hooks - instead of just that one event being skipped.
                System.Diagnostics.Trace.WriteLine("EBH Material Planning : dropped a malformed SAP event : " + ex);
            }
        }

        private static void SBO_Application_ItemEvent_Inner(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            if (pVal.FormTypeEx == ProductionHook.FormProduction)
            {
                ProductionHook.ItemEvent(FormUID, ref pVal, out BubbleEvent);
                return;
            }
            if (pVal.FormTypeEx == GrpoHook.FormGrpo)
            {
                GrpoHook.ItemEvent(FormUID, ref pVal, out BubbleEvent);
                return;
            }
            PlanForm f = Find(FormUID);
            if (f == null) return;
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                if (!pVal.BeforeAction) lock (Open) Open.Remove(FormUID);
                return;
            }
            if (!f.isBuilt) return;       // resize and activate fire while the items are still being added
            if (TraceOn) Trace(FormUID + " " + pVal.EventType + (pVal.BeforeAction ? " before" : " after") + " item '" + pVal.ItemUID + "' col '" + pVal.ColUID + "' row " + pVal.Row + " changed " + pVal.ItemChanged + " busy " + busyThread);
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE ||
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS || pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) return;
            if (!Begin()) return;
            try
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    f.Freeze(f.Relayout);
                    return;
                }
                bool bubble = true;
                // A button click (Search, Load, Create...) can take a few seconds (DI API / database round trips).
                // Every button is disabled for that click's own handler and restored right after, so an impatient
                // second click on the same or another button cannot re-enter a still-running action; the status bar
                // line says so without blocking anything.
                bool isButtonPress = pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction;
                if (isButtonPress)
                {
                    Status("Please wait, your work is in progress. Buttons auto-enable after the task finishes.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    f.SetButtonsEnabled(false);
                }
                try { lock (Db.DiLock) { if (!f.SortByTitle(pVal)) f.OnEvent(pVal, ref bubble); } }
                finally { if (isButtonPress) f.SetButtonsEnabled(true); }
                BubbleEvent = bubble;
            }
            catch (Exception ex)
            {
                f.Fail(ex);
            }
            finally
            {
                End();
            }
        }

        /// <summary>Open forms of one class (e.g. the draft list, to refresh it after a slip was saved).</summary>
        protected static IEnumerable<T> OpenForms<T>() where T : PlanForm
        {
            lock (Open) return Open.Values.OfType<T>().ToList();
        }

        /// <summary>Runs a menu / button action that opens a form: errors go to the status bar.</summary>
        public static void Run(Action open)
        {
            if (!Begin())
            {
                Status("Still working on the previous action. Try again in a moment.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return;
            }
            try
            {
                lock (Db.DiLock) open();
            }
            catch (Exception ex)
            {
                Alert(ex.Message);
            }
            finally
            {
                End();
            }
        }

        /// <summary>
        /// One instance per T of the flag below: keeps Show() static methods generic without a field per form class.
        /// </summary>
        private static class Opening<T> { public static bool Value; }

        /// <summary>
        /// Like Run(), for a form's "Show" entry point. Begin()'s same-thread reentrancy (needed so a button handler can
        /// open another add-on form while already inside an event) would otherwise let SAP calling the menu/shortcut
        /// twice for one click re-enter Show() before the first call has registered its new form in Open, opening it
        /// twice. This flag is set before Run() starts and closes that window: a reentrant Show() of the same form
        /// type just returns.
        /// </summary>
        protected static void RunShow<T>(Action open) where T : PlanForm
        {
            if (Opening<T>.Value) return;
            Opening<T>.Value = true;
            try { Run(open); }
            finally { Opening<T>.Value = false; }
        }
        #endregion

        #region Form
        protected SAPbouiCOM.Form Form;
        private bool isBuilt;
        public string Uid { get; private set; }

        /// <summary>
        /// Creates the window, lets the form add its items (build), lays it out and shows it. width / height are the
        /// size the form was designed for - its floor, never shrunk below. With autoSize, the window instead opens at
        /// roughly 90% of the SAP Desktop (still never below that floor, never past the Desktop itself), so a form
        /// meant to be worked in all day does not open tiny on a large monitor just because of a fixed pixel default;
        /// a small utility popup (autoSize false) keeps the exact size it asked for.
        /// </summary>
        protected void Create(string type, string title, int width, int height, bool modal, Action build, bool autoSize = false)
        {
            var p = (SAPbouiCOM.FormCreationParams)App.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            p.FormType = type;
            p.BorderStyle = modal ? SAPbouiCOM.BoFormBorderStyle.fbs_Fixed : SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
            if (modal) p.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
            for (int attempt = 0; Form == null; attempt++)
            {
                p.UniqueID = "EBP" + (++Counter).ToString(Inv);
                try { Form = App.Forms.AddEx(p); }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // The id is taken by a window left over from an earlier run of the add-on: next number
                    if (attempt > 200) throw;
                }
            }
            Uid = Form.UniqueID;
            lock (Open) Open[Uid] = this;
            try
            {
                Form.Freeze(true);
                Form.Title = title;
                int deskW = App.Desktop.Width, deskH = App.Desktop.Height;
                if (autoSize)
                {
                    // Grow past the designed default on a generous Desktop (the previous version stopped here, which
                    // on a small monitor left width/height pinned to the design default no matter how little room the
                    // SAP Desktop actually had - the window then opened partly off-window, its footer buttons and
                    // Total row pushed past the visible bottom edge, as seen on a 1024x768 display).
                    width = Math.Max(width, (int)(deskW * 0.9));
                    height = Math.Max(height, (int)(deskH * 0.9));
                }
                // Whatever that produced (or the plain design default for a form that does not opt into autoSize),
                // never exceed what the Desktop actually has room for: this is what makes a small monitor shrink the
                // window instead of clipping it, for every form, not only the ones above.
                // App.Desktop occasionally reports a tiny, unusable size for one call - seen when a form is created
                // right as another one opens or closes (e.g. the Switch button, or the automatic hand-off from a
                // created Purchase Request to Purchase Order Planning). Trusting that reading shrank the window to a
                // sliver, with items overlapping. No real SAP Desktop is this small, so below this it is treated as a
                // bad reading and ignored rather than honoured.
                const int minTrustedDeskW = 500, minTrustedDeskH = 400;
                if (deskW >= minTrustedDeskW) width = Math.Min(width, deskW - 20);
                if (deskH >= minTrustedDeskH) height = Math.Min(height, deskH - 60);
                Form.ClientWidth = width;
                Form.ClientHeight = height;
                build();
                Snapshot();
                Layout();
                Form.Left = Math.Max(0, (App.Desktop.Width - Form.Width) / 2);
                Form.Top = Math.Max(0, (App.Desktop.Height - Form.Height) / 3);
            }
            catch (Exception)
            {
                lock (Open) Open.Remove(Uid);
                Form.Freeze(false);
                Form.Close();
                throw;
            }
            Form.Freeze(false);
            isBuilt = true;
            Form.Visible = true;
        }

        protected abstract void OnEvent(SAPbouiCOM.ItemEvent e, ref bool bubble);

        /// <summary>Positions the items that follow the window size. Called after the build and on every resize.</summary>
        protected virtual void Layout() { }

        private readonly Dictionary<string, int[]> home = new Dictionary<string, int[]>();

        /// <summary>Remembers where the build put every item.</summary>
        private void Snapshot()
        {
            for (int i = 0; i < Form.Items.Count; i++)
            {
                SAPbouiCOM.Item it = Form.Items.Item(i);
                home[it.UniqueID] = new[] { it.Left, it.Top, it.Width, it.Height };
            }
        }

        /// <summary>
        /// After a resize SAP has moved and stretched every item in proportion (a search field ends up at the far right of a
        /// maximised window). Everything goes back to where it was built, then Layout() places the items that follow the size.
        /// </summary>
        private void Relayout()
        {
            foreach (var h in home)
            {
                SAPbouiCOM.Item it = Form.Items.Item(h.Key);
                it.Left = h.Value[0]; it.Top = h.Value[1]; it.Width = h.Value[2]; it.Height = h.Value[3];
            }
            Layout();
        }

        protected void Freeze(Action work)
        {
            Form.Freeze(true);
            try { work(); }
            finally { Form.Freeze(false); }
        }

        public void Select() { Form.Select(); }
        protected void Close() { Form.Close(); }

        protected static void Status(string message, SAPbouiCOM.BoStatusBarMessageType type)
        {
            App.StatusBar.SetText("EBH Material Planning: " + message, SAPbouiCOM.BoMessageTime.bmt_Short, type);
        }
        /// <summary>
        /// An error: shown in a message box the user has to close, and in the status bar so it stays in the
        /// System Messages Log. The box blocks the event that caused it until OK is pressed.
        /// </summary>
        public static void Alert(string message)
        {
            Status(message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            try { App.MessageBox(message, 1, "OK", "", ""); }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning message box : " + ex.Message); }
        }

        protected static void Ok(string message) { Status(message, SAPbouiCOM.BoStatusBarMessageType.smt_Success); }
        protected static void Warn(string message) { Status(message, SAPbouiCOM.BoStatusBarMessageType.smt_Warning); }

        /// <summary>
        /// A process has finished (Load, Apply, Create...): shown in a message box the user has to close, and in the
        /// status bar same as Ok(). A status bar line alone is easy to miss - it is replaced by the next action and
        /// never demands attention - so anything the user waited on gets this instead.
        /// </summary>
        protected static void Done(string message)
        {
            Status(message, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            try { App.MessageBox(message, 1, "OK", "", ""); }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning message box : " + ex.Message); }
        }

        /// <summary>Shows an error; a validation error puts the cursor on its field when the form maps it (FieldItem).</summary>
        protected virtual void Fail(Exception ex)
        {
            if (TraceOn) Trace("FAIL " + ex);
            Alert(ex.Message);
            var pe = ex as PlanException;
            if (pe == null) System.Diagnostics.Trace.WriteLine("EBH Material Planning : " + ex);
        }

        protected static bool Ask(string question, string yes, string no)
        {
            return App.MessageBox(question, 2, yes, no, "") == 1;
        }

        protected static bool Pressed(SAPbouiCOM.ItemEvent e, string uid)
        {
            return e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !e.BeforeAction && e.ItemUID == uid;
        }

        /// <summary>Enter pressed in an edit field (no Ctrl/Alt held, so it does not also match a Ctrl+Enter shortcut).</summary>
        protected static bool Enter(SAPbouiCOM.ItemEvent e, params string[] uids)
        {
            return Key(e, 13) && uids.Contains(e.ItemUID);
        }

        /// <summary>A key combination pressed anywhere on the form, whatever item has focus. CharPressed uses the same codes as virtual-key / ASCII (e.g. 13 Enter, 27 Esc, 114 F3, 'A'..'Z' for a letter).</summary>
        protected static bool Key(SAPbouiCOM.ItemEvent e, int charPressed, SAPbouiCOM.BoModifiersEnum modifiers = SAPbouiCOM.BoModifiersEnum.mt_None)
        {
            return e.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && !e.BeforeAction && e.CharPressed == charPressed && e.Modifiers == modifiers;
        }
        #endregion

        #region Items
        protected SAPbouiCOM.Item Item(string uid) { return Form.Items.Item(uid); }

        /// <summary>Base font size of every item the form builds, so the whole UI reads consistently larger than SAP's default.</summary>
        protected const int BaseFontSize = 14;

        protected SAPbouiCOM.Item Add(string uid, SAPbouiCOM.BoFormItemTypes type, int left, int top, int width, int height)
        {
            SAPbouiCOM.Item it = Form.Items.Add(uid, type);
            it.Left = left; it.Top = top; it.Width = width; it.Height = height;
            it.AffectsFormMode = false;
            try { it.FontSize = BaseFontSize; } catch (System.Runtime.InteropServices.COMException) { /* item type has no font (e.g. the section divider line) */ }
            return it;
        }

        protected void Place(string uid, int left, int top, int width, int height)
        {
            SAPbouiCOM.Item it = Item(uid);
            it.Left = left; it.Top = top;
            if (width > 0) it.Width = width;
            if (height > 0) it.Height = height;
        }

        /// <summary>Height of a single-line label/field/combo, sized for BaseFontSize.</summary>
        protected const int RowH = 18;
        /// <summary>Height of a button, sized for BaseFontSize.</summary>
        protected const int ButtonH = 22;
        /// <summary>Height of a section's bold caption; its divider line sits SectionH below "top".</summary>
        protected const int SectionH = 18;

        protected void Label(string uid, string caption, int left, int top, int width)
        {
            ((SAPbouiCOM.StaticText)Add(uid, SAPbouiCOM.BoFormItemTypes.it_STATIC, left, top, width, RowH).Specific).Caption = caption;
        }

        /// <summary>
        /// Bold section header with a thin divider line beneath it: pure visual grouping, carries no data.
        /// Adds two items, "uid" (the caption) and "uid" + "L" (the line) - keep uid at 9 characters or fewer.
        /// </summary>
        protected void Section(string uid, string caption, int left, int top, int width)
        {
            SAPbouiCOM.Item head = Add(uid, SAPbouiCOM.BoFormItemTypes.it_STATIC, left, top, width, SectionH);
            ((SAPbouiCOM.StaticText)head.Specific).Caption = caption;
            head.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_BOLD;
            head.ForeColor = GridTable.Rgb(37, 99, 235);
            SAPbouiCOM.Item line = Add(uid + "L", SAPbouiCOM.BoFormItemTypes.it_STATIC, left, top + SectionH, width, 2);
            ((SAPbouiCOM.StaticText)line.Specific).Caption = "";
            line.BackColor = GridTable.Rgb(226, 232, 240);
        }

        protected void Caption(string uid, string caption) { ((SAPbouiCOM.StaticText)Item(uid).Specific).Caption = caption; }

        private void Source(string uid, SAPbouiCOM.BoDataType type, int length)
        {
            Form.DataSources.UserDataSources.Add(uid, type, length);
        }

        /// <summary>Edit field bound to a user data source of the same id.</summary>
        protected SAPbouiCOM.Item Edit(string uid, int left, int top, int width, SAPbouiCOM.BoDataType type = SAPbouiCOM.BoDataType.dt_SHORT_TEXT, int length = 254)
        {
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_EDIT, left, top, width, RowH);
            Source(uid, type, length);
            ((SAPbouiCOM.EditText)it.Specific).DataBind.SetBound(true, "", uid);
            return it;
        }

        /// <summary>Read-only field: a disabled edit, so the text can be selected by SAP's own means and looks like any SAP field.</summary>
        protected void Field(string uid, int left, int top, int width, bool right = false)
        {
            SAPbouiCOM.Item it = Edit(uid, left, top, width);
            it.Enabled = false;
            if (right) it.RightJustified = true;
        }

        /// <summary>Multi-line read-only text.</summary>
        protected void Memo(string uid, int left, int top, int width, int height)
        {
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT, left, top, width, height);
            Source(uid, SAPbouiCOM.BoDataType.dt_LONG_TEXT, 4000);
            ((SAPbouiCOM.EditText)it.Specific).DataBind.SetBound(true, "", uid);
            it.Enabled = false;
        }

        /// <summary>Combo box; values = value, description, value, description ...</summary>
        protected SAPbouiCOM.ComboBox Combo(string uid, int left, int top, int width, params string[] values)
        {
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, left, top, width, RowH);
            it.DisplayDesc = true;
            Source(uid, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
            var combo = (SAPbouiCOM.ComboBox)it.Specific;
            combo.DataBind.SetBound(true, "", uid);
            for (int i = 0; i + 1 < values.Length; i += 2) combo.ValidValues.Add(values[i], values[i + 1]);
            return combo;
        }

        /// <summary>Replaces the choices of a combo box; pairs = value, description.</summary>
        protected void Choices(string uid, IEnumerable<KeyValuePair<string, string>> pairs)
        {
            var combo = (SAPbouiCOM.ComboBox)Item(uid).Specific;
            Set(uid, "");
            for (int i = combo.ValidValues.Count - 1; i >= 0; i--) combo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
            foreach (var p in pairs) combo.ValidValues.Add(p.Key, p.Value);
        }

        /// <summary>
        /// Choose From List on an edit field: Business Partner (object 2, vendors only), Project (63) or Warehouse (64).
        /// The chosen codes are read with ChosenCodes.
        /// </summary>
        protected void Cfl(string editUid, string objectType, string alias, bool multi)
        {
            var p = (SAPbouiCOM.ChooseFromListCreationParams)App.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            p.MultiSelection = multi;
            p.ObjectType = objectType;
            p.UniqueID = "cfl" + editUid;
            SAPbouiCOM.ChooseFromList cfl = Form.ChooseFromLists.Add(p);
            if (objectType == "2")
            {
                SAPbouiCOM.Conditions cons = cfl.GetConditions();
                SAPbouiCOM.Condition con = cons.Add();
                con.Alias = "CardType";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "S";
                cfl.SetConditions(cons);
            }
            var edit = (SAPbouiCOM.EditText)Item(editUid).Specific;
            edit.ChooseFromListUID = p.UniqueID;
            edit.ChooseFromListAlias = alias;
        }

        /// <summary>The Choose From List attached to an edit field by Cfl(), to change its conditions before it opens.</summary>
        protected SAPbouiCOM.ChooseFromList GetCfl(string editUid) { return Form.ChooseFromLists.Item("cfl" + editUid); }

        /// <summary>The codes picked in a Choose From List event of this edit field (empty when the list was cancelled).</summary>
        protected static List<string> ChosenCodes(SAPbouiCOM.ItemEvent e, string editUid, string alias)
        {
            var list = new List<string>();
            if (e.EventType != SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST || e.BeforeAction || e.ItemUID != editUid) return list;
            SAPbouiCOM.DataTable dt = ((SAPbouiCOM.IChooseFromListEvent)e).SelectedObjects;
            if (dt == null) return list;
            for (int i = 0; i < dt.Rows.Count; i++) list.Add(Convert.ToString(dt.GetValue(alias, i), Inv));
            return list;
        }

        /// <summary>
        /// Choose From List on a grid column (e.g. row level Warehouse or Vendor). Business Partner (object 2) is
        /// narrowed to vendors, the same as the single-field Cfl. One list serves every row of the column.
        /// </summary>
        protected void CflColumn(GridTable grid, string colUid, string objectType, string alias)
        {
            var p = (SAPbouiCOM.ChooseFromListCreationParams)App.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            p.MultiSelection = false;
            p.ObjectType = objectType;
            p.UniqueID = "cfl" + grid.Uid + colUid;
            SAPbouiCOM.ChooseFromList cfl = Form.ChooseFromLists.Add(p);
            if (objectType == "2")
            {
                SAPbouiCOM.Conditions cons = cfl.GetConditions();
                SAPbouiCOM.Condition con = cons.Add();
                con.Alias = "CardType";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "S";
                cfl.SetConditions(cons);
            }
            var col = (SAPbouiCOM.EditTextColumn)grid.Grid.Columns.Item(colUid);
            col.ChooseFromListUID = p.UniqueID;
            col.ChooseFromListAlias = alias;
        }

        /// <summary>The code picked in a Choose From List event raised from one grid column; "row" is the DataTable row it applies to (-1 when not this column / cancelled).</summary>
        protected static List<string> ChosenGridCodes(SAPbouiCOM.ItemEvent e, GridTable grid, string colUid, string alias, out int row)
        {
            row = -1;
            var list = new List<string>();
            if (e.EventType != SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST || e.BeforeAction || e.ItemUID != grid.Uid || e.ColUID != colUid) return list;
            SAPbouiCOM.DataTable dt = ((SAPbouiCOM.IChooseFromListEvent)e).SelectedObjects;
            if (dt == null) return list;
            for (int i = 0; i < dt.Rows.Count; i++) list.Add(Convert.ToString(dt.GetValue(alias, i), Inv));
            row = grid.Row(e.Row);
            return list;
        }

        protected void Check(string uid, string caption, int left, int top, int width)
        {
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, left, top, width, RowH);
            Source(uid, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            var check = (SAPbouiCOM.CheckBox)it.Specific;
            check.Caption = caption;
            check.ValOn = "Y"; check.ValOff = "N";
            check.DataBind.SetBound(true, "", uid);
        }

        protected void Button(string uid, string caption, int left, int top, int width)
        {
            ((SAPbouiCOM.Button)Add(uid, SAPbouiCOM.BoFormItemTypes.it_BUTTON, left, top, width, ButtonH).Specific).Caption = caption;
        }

        /// <summary>Button tinted with a background colour (see GridTable's Red/Yellow/Green/Grey/Blue), to tell actions apart at a glance.</summary>
        protected void Button(string uid, string caption, int left, int top, int width, int backColor)
        {
            Button(uid, caption, left, top, width);
            Item(uid).BackColor = backColor;
        }

        protected void ButtonCaption(string uid, string caption) { ((SAPbouiCOM.Button)Item(uid).Specific).Caption = caption; }

        /// <summary>Golden arrow in front of a field. The press is handled by the form (LinkPressed), SAP's own link is cancelled.</summary>
        protected void Link(string uid, string editUid, SAPbouiCOM.BoLinkedObject target)
        {
            SAPbouiCOM.Item edit = Item(editUid);
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON, edit.Left - 14, edit.Top, 12, 12);
            it.LinkTo = editUid;
            ((SAPbouiCOM.LinkedButton)it.Specific).LinkedObject = target;
        }

        /// <summary>
        /// A golden arrow was used. The form cancels the event (so SAP does not follow its own link) and acts when
        /// act is true: once per press, on the click.
        /// </summary>
        protected static bool LinkPressed(SAPbouiCOM.ItemEvent e, string uid, out bool act)
        {
            act = false;
            if (!e.BeforeAction || e.ItemUID != uid) return false;
            if (e.EventType != SAPbouiCOM.BoEventTypes.et_CLICK && e.EventType != SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) return false;
            act = e.EventType == SAPbouiCOM.BoEventTypes.et_CLICK;
            return true;
        }

        /// <summary>Enables / disables an item. SAP refuses this for the item that has the focus: the action behind it checks again.</summary>
        protected void Enable(string uid, bool on)
        {
            try
            {
                SAPbouiCOM.Item it = Item(uid);
                if (it.Enabled != on) it.Enabled = on;
            }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning enable " + uid + " : " + ex.Message); }
        }

        /// <summary>Disables (false) or restores (true) every button on the form. Called by the dispatcher around a
        /// button press's own handler - not meant to be called directly by form code.</summary>
        private void SetButtonsEnabled(bool on)
        {
            for (int i = 0; i < Form.Items.Count; i++)
            {
                SAPbouiCOM.Item it = Form.Items.Item(i);
                if (it.Type == SAPbouiCOM.BoFormItemTypes.it_BUTTON) Enable(it.UniqueID, on);
            }
        }

        protected void Visible(string uid, bool on)
        {
            try
            {
                SAPbouiCOM.Item it = Item(uid);
                if (it.Visible != on) it.Visible = on;
            }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning show " + uid + " : " + ex.Message); }
        }

        /// <summary>
        /// Editing a grid cell puts the form into Update mode, and SAP then renames button 2 to "Cancel".
        /// The add-on forms save through their own buttons: back to OK mode and to the caption the form gave the button.
        /// </summary>
        protected void KeepOkMode(string caption)
        {
            try
            {
                if (Form.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE) Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                var close = (SAPbouiCOM.Button)Item("2").Specific;
                if (close.Caption != caption) close.Caption = caption;
            }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning form mode : " + ex.Message); }
        }

        protected void Focus(string uid)
        {
            try { Form.ActiveItem = uid; }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning focus " + uid + " : " + ex.Message); }
        }
        #endregion

        #region Values
        /// <summary>Stored value of a field (combo, check box, read-only field).</summary>
        protected string Get(string uid) { return (Form.DataSources.UserDataSources.Item(uid).ValueEx ?? "").Trim(); }

        protected void Set(string uid, string value) { Form.DataSources.UserDataSources.Item(uid).ValueEx = value ?? ""; }

        /// <summary>
        /// What is typed in an edit field right now, in SAP's internal format (numbers with a point, dates yyyyMMdd).
        /// The data source only receives the text when the field loses the focus, which a button press or Enter does not cause.
        /// </summary>
        protected string Text(string uid) { return (((SAPbouiCOM.EditText)Item(uid).Specific).Value ?? "").Trim(); }

        protected decimal Number(string uid)
        {
            decimal d;
            return decimal.TryParse(Text(uid), NumberStyles.Number, Inv, out d) ? d : 0;
        }

        protected void Set(string uid, decimal value) { Set(uid, value.ToString(Inv)); }

        /// <summary>Date field as yyyy-MM-dd ("" when empty).</summary>
        protected string Date(string uid)
        {
            DateTime d;
            return DateTime.TryParseExact(Text(uid), "yyyyMMdd", Inv, DateTimeStyles.None, out d) ? d.ToString("yyyy-MM-dd") : "";
        }

        protected void SetDate(string uid, DateTime value) { Set(uid, value.ToString("yyyyMMdd")); }

        protected static DateTime ToDate(string iso)
        {
            DateTime d;
            return DateTime.TryParseExact(iso, "yyyy-MM-dd", Inv, DateTimeStyles.None, out d) ? d : new DateTime(1899, 12, 30);
        }

        /// <summary>Date as text in the short format of the PC (read-only fields).</summary>
        protected static string ShowDate(string iso)
        {
            DateTime d;
            return DateTime.TryParseExact(iso, "yyyy-MM-dd", Inv, DateTimeStyles.None, out d) ? d.ToShortDateString() : "";
        }

        protected static string Num(decimal v) { return v.ToString("0.######", Inv); }
        #endregion

        #region Grid
        protected GridTable Grid(string uid, int left, int top, int width, int height)
        {
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_GRID, left, top, width, height);
            return new GridTable((SAPbouiCOM.Grid)it.Specific, Form.DataSources.DataTables.Add(uid));
        }

        private readonly Dictionary<string, KeyValuePair<GridTable, Action>> sortedGrids = new Dictionary<string, KeyValuePair<GridTable, Action>>();

        /// <summary>Double-clicking a sortable column title of this grid re-sorts it and redraws it with show (which must apply grid.Sort).</summary>
        protected void SortWith(GridTable grid, Action show)
        {
            sortedGrids[grid.Uid] = new KeyValuePair<GridTable, Action>(grid, show);
        }

        /// <summary>A column title double-click (Row -1) on a sortable column: handled here, so the form never sees it.</summary>
        private bool SortByTitle(SAPbouiCOM.ItemEvent e)
        {
            KeyValuePair<GridTable, Action> g;
            if (e.EventType != SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK || e.BeforeAction || e.Row != -1 ||
                !sortedGrids.TryGetValue(e.ItemUID ?? "", out g) || !g.Key.ToggleSort(e.ColUID)) return false;
            g.Value();
            return true;
        }

        protected static bool Clicked(SAPbouiCOM.ItemEvent e, GridTable grid)
        {
            return e.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !e.BeforeAction && e.ItemUID == grid.Uid && e.Row >= 0;
        }

        protected static bool DoubleClicked(SAPbouiCOM.ItemEvent e, GridTable grid)
        {
            return e.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK && !e.BeforeAction && e.ItemUID == grid.Uid && e.Row >= 0;
        }
        #endregion
    }

    /// <summary>
    /// A grid on its own DataTable, filled from code. Only cells whose value changed are written, so a
    /// recalculation after every quantity edit stays quick (each cell is one call into the SAP client).
    /// </summary>
    internal sealed class GridTable
    {
        public readonly SAPbouiCOM.Grid Grid;
        public readonly SAPbouiCOM.DataTable Table;
        public readonly string Uid;
        private readonly List<string> cols = new List<string>();
        private readonly List<Action> setup = new List<Action>();
        private readonly List<string[]> cache = new List<string[]>();
        private readonly List<int> colors = new List<int>();
        private readonly Dictionary<string, string> titles = new Dictionary<string, string>();
        private readonly HashSet<string> sortable = new HashSet<string>();
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        /// <summary>Row of the DataTable the user clicked last (-1 = none).</summary>
        public int Current = -1;

        public const int NoColor = -1;
        public static readonly int Red = Rgb(255, 205, 205), Yellow = Rgb(255, 248, 196), Green = Rgb(214, 240, 214), Grey = Rgb(232, 232, 232), Blue = Rgb(214, 230, 247);
        // Named palette for Production Order completion progress (#D1FAE5 / #FEF3C7 / #E5E7EB), linked-document
        // highlighting (#DBEAFE) and the Materials grid's default row tint (#F3E8FF)
        public static readonly int MintGreen = Rgb(0xD1, 0xFA, 0xE5), PastelYellow = Rgb(0xFE, 0xF3, 0xC7), CoolGray = Rgb(0xE5, 0xE7, 0xEB), SkyBlue = Rgb(0xDB, 0xEA, 0xFE), Lavender = Rgb(0xF3, 0xE8, 0xFF);
        internal static int Rgb(int r, int g, int b) { return r | (g << 8) | (b << 16); }

        public GridTable(SAPbouiCOM.Grid grid, SAPbouiCOM.DataTable table)
        {
            Grid = grid; Table = table; Uid = grid.Item.UniqueID;
        }

        /// <summary>Adds a column. Text columns are 254 characters; numbers are right-aligned by SAP.</summary>
        public GridTable Column(string uid, string title, SAPbouiCOM.BoFieldsType type = SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, int width = 0, bool editable = false)
        {
            Table.Columns.Add(uid, type, type == SAPbouiCOM.BoFieldsType.ft_AlphaNumeric ? 254 : 0);
            cols.Add(uid);
            titles[uid] = title;
            setup.Add(() =>
            {
                SAPbouiCOM.GridColumn c = Grid.Columns.Item(uid);
                c.TitleObject.Caption = title;
                c.Editable = editable;
                if (width > 0) c.Width = width;
            });
            return this;
        }

        /// <summary>Turns a text column into a combo box column.</summary>
        public GridTable ComboColumn(string uid)
        {
            setup.Add(() =>
            {
                Grid.Columns.Item(uid).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                ((SAPbouiCOM.ComboBoxColumn)Grid.Columns.Item(uid)).DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Value;
            });
            return this;
        }

        /// <summary>Turns a text column holding "Y" / "N" into a check box column.</summary>
        public GridTable CheckColumn(string uid)
        {
            setup.Add(() => Grid.Columns.Item(uid).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox);
            return this;
        }

        /// <summary>Shows a golden arrow in the cells of a column. The form handles the press itself (et_MATRIX_LINK_PRESSED).</summary>
        public GridTable LinkColumn(string uid, string linkedObjectType)
        {
            setup.Add(() => ((SAPbouiCOM.EditTextColumn)Grid.Columns.Item(uid)).LinkedObjectType = linkedObjectType);
            return this;
        }

        /// <summary>
        /// Columns the user can sort by double-clicking their title (ascending, then descending). The form sorts
        /// its own list with Sort() before Fill, so grid row = DataTable row = list index stays true everywhere
        /// (row colours, locked cells, Current). SAP's own title sort is switched off for these columns.
        /// </summary>
        public GridTable Sortable(params string[] uids)
        {
            foreach (string uid in uids) sortable.Add(uid);
            setup.Add(() =>
            {
                foreach (string uid in uids)
                {
                    try { Grid.Columns.Item(uid).TitleObject.Sortable = false; }
                    catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning sortable : " + ex.Message); }
                }
            });
            return this;
        }

        /// <summary>Column the rows are sorted by (null = the form's own order) and its direction.</summary>
        public string SortColumn { get; private set; }
        public bool SortAscending { get; private set; } = true;

        /// <summary>A title double-click on a sortable column: ascending first, then flips. False for any other column.</summary>
        public bool ToggleSort(string uid)
        {
            if (!sortable.Contains(uid)) return false;
            SortAscending = SortColumn == uid ? !SortAscending : true;
            SortColumn = uid;
            Current = -1;       // the clicked row moves; a stale index would point at a different line
            foreach (string c in sortable)
            {
                try { Grid.Columns.Item(c).TitleObject.Caption = titles[c] + (c == SortColumn ? (SortAscending ? " ▲" : " ▼") : ""); }
                catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning sort caption : " + ex.Message); }
            }
            return true;
        }

        /// <summary>The rows in the chosen sort order (stable: equal values keep the form's own order). key(column, row) gives the text to sort on.</summary>
        public List<T> Sort<T>(List<T> rows, Func<string, T, string> key)
        {
            if (SortColumn == null) return rows;
            string col = SortColumn;
            return (SortAscending
                ? rows.OrderBy(r => key(col, r) ?? "", StringComparer.CurrentCultureIgnoreCase)
                : rows.OrderByDescending(r => key(col, r) ?? "", StringComparer.CurrentCultureIgnoreCase)).ToList();
        }

        /// <summary>Several rows can be selected with Ctrl / Shift (the Production Orders list).</summary>
        public GridTable MultiSelect()
        {
            setup.Add(() => Grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto);
            return this;
        }

        /// <summary>DataTable rows of the grid rows that are selected.</summary>
        public List<int> SelectedRows()
        {
            var list = new List<int>();
            SAPbouiCOM.SelectedRows sel = Grid.Rows.SelectedRows;
            if (sel == null) return list;
            for (int i = 0; i < sel.Count; i++) list.Add(Grid.GetDataTableRowIndex(sel.Item(i, SAPbouiCOM.BoOrderType.ot_RowOrder)));
            return list;
        }

        /// <summary>Binds the grid to the table and applies titles, widths and editability.</summary>
        public GridTable Bind()
        {
            Table.Rows.Clear();
            Grid.DataTable = Table;
            Grid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
            foreach (Action a in setup) a();
            return this;
        }

        public void Editable(string uid, bool on)
        {
            SAPbouiCOM.GridColumn c = Grid.Columns.Item(uid);
            if (c.Editable != on) c.Editable = on;
        }

        /// <summary>
        /// Widens one column (typically Description) to soak up whatever width the grid has left over once every
        /// other column has its designed width, instead of leaving a blank grey stripe down the right side of the
        /// grid on a wide window - call from Layout() after Place()'ing the grid to its final width.
        /// </summary>
        public void FillWidth(string uid, int gridWidth)
        {
            try
            {
                int used = 0;
                for (int i = 0; i < Grid.Columns.Count; i++)
                {
                    SAPbouiCOM.GridColumn c = Grid.Columns.Item(i);
                    if (c.UniqueID != uid) used += c.Width;
                }
                int w = gridWidth - used - 20;    // room for the row-selector strip and a vertical scrollbar
                SAPbouiCOM.GridColumn target = Grid.Columns.Item(uid);
                if (w > 60 && target.Width != w) target.Width = w;
            }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning column width : " + ex.Message); }
        }

        public void AddChoice(string uid, string value, string description)
        {
            ((SAPbouiCOM.ComboBoxColumn)Grid.Columns.Item(uid)).ValidValues.Add(value, description);
        }

        private static string Key(object v)
        {
            if (v == null) return "";
            if (v is DateTime) return ((DateTime)v).ToString("yyyyMMdd");
            return Convert.ToString(v, Inv);
        }

        private static object Native(object v)
        {
            if (v == null) return "";
            if (v is decimal) return (double)(decimal)v;
            return v;
        }

        /// <summary>Shows these rows (one object[] per row, in column order): string, int, decimal or DateTime values.</summary>
        public void Fill(IList<object[]> rows)
        {
            for (int i = cache.Count - 1; i >= rows.Count; i--)
            {
                Table.Rows.Remove(i);
                cache.RemoveAt(i); colors.RemoveAt(i);
            }
            if (rows.Count > cache.Count)
            {
                Table.Rows.Add(rows.Count - cache.Count);
                while (cache.Count < rows.Count) { cache.Add(new string[cols.Count]); colors.Add(NoColor); }
            }
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < cols.Count; c++)
                {
                    string key = Key(rows[r][c]);
                    if (cache[r][c] == key) continue;
                    Table.SetValue(cols[c], r, Native(rows[r][c]));
                    cache[r][c] = key;
                }
            if (Current >= rows.Count) Current = -1;
        }

        public int Count { get { return cache.Count; } }

        /// <summary>Value the user typed into a cell. The next Fill writes the cell again, whatever it holds.</summary>
        public decimal Typed(int row, string uid)
        {
            cache[row][cols.IndexOf(uid)] = null;
            return Convert.ToDecimal(Table.GetValue(uid, row), Inv);
        }

        /// <summary>Date the user typed into a cell as yyyy-MM-dd ("" when empty).</summary>
        public string TypedDate(int row, string uid)
        {
            cache[row][cols.IndexOf(uid)] = null;
            object v = Table.GetValue(uid, row);
            return v is DateTime && ((DateTime)v).Year > 1950 ? ((DateTime)v).ToString("yyyy-MM-dd") : "";
        }

        public string TypedText(int row, string uid)
        {
            cache[row][cols.IndexOf(uid)] = null;
            return Convert.ToString(Table.GetValue(uid, row), Inv).Trim();
        }

        /// <summary>Background of a row (NoColor = SAP default).</summary>
        public void Color(int row, int color)
        {
            if (colors[row] == color) return;
            try
            {
                Grid.CommonSetting.SetRowBackColor(row + 1, color);
                colors[row] = color;
            }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning row colour : " + ex.Message); }
        }

        /// <summary>Disables (greys out, non-clickable) one cell of one row - e.g. the checkbox of a closed Purchase
        /// Request line, so it cannot be ticked at all rather than being ticked and then reverted. SetCellEditable
        /// takes the column by position (0-based, matching Column() call order), not by UID.</summary>
        public void Lock(int row, string colUid, bool editable)
        {
            try { Grid.CommonSetting.SetCellEditable(row + 1, cols.IndexOf(colUid), editable); }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning cell lock : " + ex.Message); }
        }

        /// <summary>The grid row of a click event becomes the current row and is highlighted.</summary>
        public void Click(int gridRow)
        {
            Current = Grid.GetDataTableRowIndex(gridRow);
            try
            {
                Grid.Rows.SelectedRows.Clear();
                Grid.Rows.SelectedRows.Add(gridRow);
            }
            catch (System.Runtime.InteropServices.COMException ex) { System.Diagnostics.Trace.WriteLine("EBH Material Planning row select : " + ex.Message); }
        }

        /// <summary>Row of the DataTable for an event row.</summary>
        public int Row(int gridRow) { return Grid.GetDataTableRowIndex(gridRow); }
    }
}

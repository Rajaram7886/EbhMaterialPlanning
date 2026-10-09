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
            if (pVal.FormTypeEx == ProductionHook.FormProduction)
            {
                ProductionHook.ItemEvent(FormUID, ref pVal, out BubbleEvent);
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
                lock (Db.DiLock) f.OnEvent(pVal, ref bubble);
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
        #endregion

        #region Form
        protected SAPbouiCOM.Form Form;
        private bool isBuilt;
        public string Uid { get; private set; }

        /// <summary>Creates the window, lets the form add its items (build), lays it out and shows it.</summary>
        protected void Create(string type, string title, int width, int height, bool modal, Action build)
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

        /// <summary>Enter pressed in an edit field.</summary>
        protected static bool Enter(SAPbouiCOM.ItemEvent e, params string[] uids)
        {
            return e.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && !e.BeforeAction && e.CharPressed == 13 && uids.Contains(e.ItemUID);
        }
        #endregion

        #region Items
        protected SAPbouiCOM.Item Item(string uid) { return Form.Items.Item(uid); }

        protected SAPbouiCOM.Item Add(string uid, SAPbouiCOM.BoFormItemTypes type, int left, int top, int width, int height)
        {
            SAPbouiCOM.Item it = Form.Items.Add(uid, type);
            it.Left = left; it.Top = top; it.Width = width; it.Height = height;
            it.AffectsFormMode = false;
            return it;
        }

        protected void Place(string uid, int left, int top, int width, int height)
        {
            SAPbouiCOM.Item it = Item(uid);
            it.Left = left; it.Top = top;
            if (width > 0) it.Width = width;
            if (height > 0) it.Height = height;
        }

        protected void Label(string uid, string caption, int left, int top, int width)
        {
            ((SAPbouiCOM.StaticText)Add(uid, SAPbouiCOM.BoFormItemTypes.it_STATIC, left, top, width, 14).Specific).Caption = caption;
        }

        protected void Caption(string uid, string caption) { ((SAPbouiCOM.StaticText)Item(uid).Specific).Caption = caption; }

        private void Source(string uid, SAPbouiCOM.BoDataType type, int length)
        {
            Form.DataSources.UserDataSources.Add(uid, type, length);
        }

        /// <summary>Edit field bound to a user data source of the same id.</summary>
        protected SAPbouiCOM.Item Edit(string uid, int left, int top, int width, SAPbouiCOM.BoDataType type = SAPbouiCOM.BoDataType.dt_SHORT_TEXT, int length = 254)
        {
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_EDIT, left, top, width, 14);
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
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX, left, top, width, 14);
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
            SAPbouiCOM.Item it = Add(uid, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX, left, top, width, 14);
            Source(uid, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            var check = (SAPbouiCOM.CheckBox)it.Specific;
            check.Caption = caption;
            check.ValOn = "Y"; check.ValOff = "N";
            check.DataBind.SetBound(true, "", uid);
        }

        protected void Button(string uid, string caption, int left, int top, int width)
        {
            ((SAPbouiCOM.Button)Add(uid, SAPbouiCOM.BoFormItemTypes.it_BUTTON, left, top, width, 19).Specific).Caption = caption;
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
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        /// <summary>Row of the DataTable the user clicked last (-1 = none).</summary>
        public int Current = -1;

        public const int NoColor = -1;
        public static readonly int Red = Rgb(255, 205, 205), Yellow = Rgb(255, 248, 196), Green = Rgb(214, 240, 214), Grey = Rgb(232, 232, 232), Blue = Rgb(214, 230, 247);
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

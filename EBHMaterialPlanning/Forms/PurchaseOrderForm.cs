using System;
using System.Collections.Generic;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Purchase Order Planning: open Purchase Request lines of the selected Project / Vendor -> one Purchase Order.
    /// The picker lists one row per eligible line (a request can carry lines of more than one Project / Vendor).
    /// Filters apply as soon as a value is picked or committed (Enter), or on F5.
    /// Shortcuts: Ctrl+F focuses Purchase Request search, F5 searches, Alt+L focuses the material grid, Ctrl+Enter creates the Purchase Order.
    /// </summary>
    internal sealed class PurchaseOrderForm : PlanForm
    {
        private const int ObjPurchaseRequest = 1470000113, ObjPurchaseOrder = 22, ObjTaxCode = 128;

        private GridTable gReq, gRows;
        private List<PurchaseRequestLine> requests = new List<PurchaseRequestLine>();
        /// <summary>The requests actually shown in gReq: all of requests, or just those of the chosen Manufacturer. Grid
        /// row indexes (click handlers, ...) always index into this, never into requests, so a filtered-out line can
        /// never be touched by a stale index.</summary>
        private List<PurchaseRequestLine> reqView = new List<PurchaseRequestLine>();
        private List<PoRow> rows = new List<PoRow>();
        /// <summary>The rows actually shown in gRows: all of rows, or just those of the chosen Manufacturer. Grid row
        /// indexes (ReadRow, click handlers, ...) always index into this, never into rows, so a filtered-out row can
        /// never be touched by a stale index.</summary>
        private List<PoRow> view = new List<PoRow>();
        /// <summary>Stock of the loaded items for the In Stock / Project Stock / Main Loc Stock columns, read on Load.</summary>
        private StockMap stock = new StockMap();
        private bool multiBranch;
        private int lastPoEntry, lastPoNum;
        /// <summary>
        /// Share of the space between the two grids given to the Purchase Request Lines grid (the rest goes to
        /// Materials); 0.5 = 50:50. SAPbouiCOM has no draggable splitter item (checked BoFormItemTypes - the closest
        /// is it_RECTANGLE, a plain shape with no drag/resize event of its own), so the split is adjusted with the
        /// bSplitUp / bSplitDn buttons instead and recomputed on every resize via Layout().
        /// </summary>
        private double splitRatio = 0.5;
        /// <summary>True right after this form itself sets eProj (CFL pick, auto-init from Vendor / Purchase Request): the
        /// Validate event SAP raises for that field change must not also run ClearOnProjectChange.</summary>
        private bool settingProject;
        /// <summary>Same as settingProject, for eVend: suppresses the auto Search() the Validate event would otherwise
        /// trigger a second time right after this form's own code already set and searched for the new Vendor.</summary>
        private bool settingVendor;

        private static PurchaseOrderForm OpenOrFocus()
        {
            PurchaseOrderForm f = OpenForms<PurchaseOrderForm>().FirstOrDefault();
            if (f != null) { f.Select(); return f; }
            f = new PurchaseOrderForm();
            f.Create("EBHMRPPO", "Purchase Order Planning", 1300, 700, false, f.Build, autoSize: true);
            return f;
        }

        public static void Show()
        {
            RunShow<PurchaseOrderForm>(() => OpenOrFocus().Focus("eProj"));
        }

        /// <summary>
        /// Opens (or focuses) Purchase Order Planning already searched and loaded for one just-created Purchase
        /// Request, so punching the Purchase Order from it is a single click away - used right after Material
        /// Request Planning creates a request.
        /// Deferred one message-loop tick (a short-lived WinForms Timer, not run inline): creating / filling the grid
        /// of a second form from inside another form's own button-click event left the grid's DataTable correctly
        /// filled but visually blank until the user interacted with the new window themselves (e.g. clicking its own
        /// Search button) - a SAP B1 UI API repaint quirk for cross-form work done inside someone else's event. Letting
        /// the originating click's event handler return first, then doing the open/fill as its own top-level action,
        /// avoids it.
        /// </summary>
        public static void ShowForRequest(int prDocNum, string project)
        {
            var timer = new System.Windows.Forms.Timer { Interval = 150 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                RunShow<PurchaseOrderForm>(() => OpenOrFocus().PrefillFromRequest(prDocNum, project));
            };
            timer.Start();
        }

        private void PrefillFromRequest(int prDocNum, string project)
        {
            // SetProject, not a raw Set: SAP raises a Validate for eProj some time after this field change, and
            // without the suppression flag that later, asynchronous event ran ClearOnProjectChange() - which wipes
            // eFind and rows and re-Searches with Project only, overwriting everything built below with every open
            // line of the Project (no Purchase Request number filter at all, which is exactly the bug reported: all
            // lines of the Project shown, the Purchase Request field blank, and Materials emptied back out).
            SetProject(project ?? "");
            Set("eFind", prDocNum.ToString(Inv));
            // Not Search(): right after a brand new form is created, the eProj / eFind edit controls can lag a beat
            // behind the data source values just set above, so Search()'s own Text("eProj") / Text("eFind") read back
            // stale (empty) values and the grid stays blank until the user clicks Search themselves. Querying directly
            // with what was just passed in sidesteps that control/data-source race entirely.
            requests = Store.ListRequestLines(Codes(project ?? ""), new string[0], prDocNum.ToString(Inv));
            ShowRequests();
            RefreshVendorDisplay();
            if (requests.Count == 0) { Warn("No open Purchase Request line found for this search."); return; }
            SetAllRequestsSelected(true);
            // Same rule as ticking a line by hand: fill Vendor only when every loaded line agrees on one (PRQ1.LineVendor) -
            // a mixed-vendor request is left with Vendor blank so Load() below still picks up every one of its lines.
            List<string> vendors = requests.Select(r => r.Vendor).Where(v => v.Length > 0).Distinct().ToList();
            if (vendors.Count == 1) SetVendor(vendors[0]);
            RefreshVendorDisplay();
            Load();
        }

        #region Build
        private void Build()
        {
            // 1. Filters
            Section("hFilt", "Search Filters", 8, 6, 300);
            Label("lProj", "Project", 8, 30, 55);
            Edit("eProj", 66, 30, 200);
            Cfl("eProj", "63", "PrjCode", false);
            Label("lVend", "Vendor", 276, 30, 50);
            Edit("eVend", 330, 30, 150);
            Cfl("eVend", "2", "CardCode", false);
            Label("lFind", "Purchase Request", 490, 30, 115);
            Edit("eFind", 608, 30, 140);
            Cfl("eFind", ObjPurchaseRequest.ToString(Inv), "DocNum", true);
            // Manufacturer: filters the Materials grid to just that Manufacturer's rows (OITM.FirmCode). Applies
            // automatically when the combo itself changes; the button is a reliable, explicit way to re-apply it too.
            // Kept on the same line as the other search filters, right next to Purchase Request.
            Label("lMan", "Manufacture", 756, 30, 90);
            Combo("cMan", 850, 28, 150);
            Button("bManFilter", "Apply Manufacturer Filter", 1004, 28, 190, GridTable.Blue);
            Button("bClear", "Clear", 1202, 28, 70, GridTable.Grey);
            // Right-aligned by Layout(): jumps to the other planning screen without going through the SAP menu
            Button("bSwitch", "Switch to Material Request Planning", 910, 28, 220, GridTable.Blue);
            Visible("bSwitch", false);

            // 2. Purchase Request lines
            Section("hReq", "Purchase Request Lines", 8, 60, 300);
            Label("lReqCap", "[Tick the line(s) to draw a Purchase Order from]", 230, 62, 300);
            gReq = Grid("gReq", 8, 86, 964, 130)
                .Column("Sr", "#", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Sel", "Select", width: 46, editable: true).CheckColumn("Sel")
                .Column("No", "Request No.", SAPbouiCOM.BoFieldsType.ft_Integer, 70).LinkColumn("No", ObjPurchaseRequest.ToString(Inv))
                .Column("Line", "Row", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Ser", "Series", width: 80)
                .Column("Prj", "Project", width: 85)
                .Column("Vend", "Vendor", width: 85)
                .Column("Catalog", "Catalog No.", width: 90)
                .Column("Item", "Item No.", width: 95).LinkColumn("Item", "4")
                .Column("Desc", "Description", width: 160)
                .Column("Man", "Manufacturer", width: 110)
                .Column("Job", "Job Code", width: 70)
                .Column("Qty", "Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 65)
                .Column("Open", "Open Qty", SAPbouiCOM.BoFieldsType.ft_Quantity, 65)
                .Column("Whs", "Warehouse", width: 75)
                .Column("Stat", "Status", width: 60)
                .Column("Date", "Date", SAPbouiCOM.BoFieldsType.ft_Date, 75)
                .Sortable("Catalog", "Item", "Desc", "Man", "Vend", "Stat", "Whs", "Date")
                .Bind();
            SortWith(gReq, ShowRequests);
            Button("bLoad", "Load Selected Requests", 8, 222, 160, GridTable.Blue);
            Check("cReqAll", "Select / Deselect All", 176, 222, 160);
            // Own status label for this grid, separate from the Materials grid's "lFoot": the two grids' captions
            // used to share "lFoot", so re-searching Purchase Request Lines (e.g. picking a Vendor) overwrote the
            // Materials summary ("43 of 43 row(s) selected") with the request lines' own ("152 of 152 line(s)
            // selected"), reading as if Materials had just reloaded to 152 rows when it had not moved at all.
            Label("lReqFoot", "", 346, 225, 360);
            // Splitter, on the same row as bLoad / cReqAll, right-aligned: Layout() repositions all three on every
            // call, so their Build() position is a placeholder. A real drag handle does not exist in the UI API (no
            // mouse-move/drag event on any item); this is the supported substitute - nudge buttons plus a one-click
            // reset, per the design note above.
            Button("bSplitRst", "↕ click to reset ↕", 716, 222, 140, GridTable.Grey);
            Button("bSplitUp", "▲", 860, 222, 36, GridTable.Grey);
            Button("bSplitDn", "▼", 900, 222, 36, GridTable.Grey);

            // 3. Purchase Order header
            Section("hHead", "Purchase Order Details", 8, 250, 300);
            Label("lSeries", "Series", 8, 272, 48);
            Combo("cSeries", 60, 270, 115);

            Label("lNextNo", "Doc. No.", 184, 272, 55);
            Field("fNextNo", 242, 270, 65, true);

            Label("lBranch", "Branch", 316, 272, 48);
            Combo("cBranch", 368, 270, 125);

            Label("lDue", "Delivery Date", 502, 272, 85);
            Edit("eDue", 591, 270, 90, SAPbouiCOM.BoDataType.dt_DATE);

            Label("lVRef", "Vendor Ref.", 692, 272, 75);
            Edit("eVRef", 771, 270, 100);        // optional

            Label("lVend2", "Vendor", 8, 299, 48);
            Field("fVend", 60, 297, 150);

            // Picks one Tax Code and stamps it onto every currently-selected Materials row in one click, instead of
            // fixing the Tax column line by line after a rejection like "(line 1) you must choose IGST (-10)".
            Label("lTaxApply", "Tax Code", 240, 299, 60);
            Edit("eTaxApply", 304, 297, 90);
            Cfl("eTaxApply", ObjTaxCode.ToString(Inv), "Code", false);
            Button("bApplyTax", "Apply to Selected Rows", 398, 295, 160, GridTable.Blue);

            Check("cRowAll", "Select / Deselect All", 607, 296, 160);

            // 4. Material details
            Section("hMat", "Materials", 8, 324, 300);
            gRows = Grid("gRows", 8, 348, 964, 280)
                .Column("Sr", "#", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Sel", "Select", width: 46, editable: true).CheckColumn("Sel")
                .Column("Catalog", "Catalog No.", width: 90)
                .Column("Item", "Item No.", width: 110).LinkColumn("Item", "4")
                .Column("Desc", "Description", width: 200)
                .Column("Man", "Manufacturer", width: 120)
                .Column("Job", "Job Code", width: 70)
                .Column("Qty", "Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 70, true)
                .Column("Open", "Open Qty", SAPbouiCOM.BoFieldsType.ft_Quantity, 70)
                .Column("Price", "Price", SAPbouiCOM.BoFieldsType.ft_Price, 80, true)
                .Column("Disc", "Discount %", SAPbouiCOM.BoFieldsType.ft_Percent, 70, true)
                .Column("After","Price After Discount", SAPbouiCOM.BoFieldsType.ft_Price, 100)
                .Column("Total", "Total", SAPbouiCOM.BoFieldsType.ft_Sum, 90)
                .Column("Whs", "Warehouse", width: 90, editable: true)
                .Column("Stk", "In Stock", SAPbouiCOM.BoFieldsType.ft_Quantity, 70)
                .Column("PStk", "Project Stock", SAPbouiCOM.BoFieldsType.ft_Quantity, 80)
                .Column("MStk", "Main Loc Stock", SAPbouiCOM.BoFieldsType.ft_Quantity, 85)
                .Column("Prj", "Project", width: 90)
                .Column("Tax", "Tax Code", width: 90, editable: true)
                .Column("Base", "Base Document", width: 100).LinkColumn("Base", ObjPurchaseRequest.ToString(Inv))
                .Sortable("Catalog", "Item", "Desc", "Man", "Whs")
                .Bind();
            SortWith(gRows, ShowRows);
            CflColumn(gRows, "Whs", "64", "WhsCode");
            CflColumn(gRows, "Tax", ObjTaxCode.ToString(Inv), "Code");

            Button("bCreate", "Purchase Order", 8, 634, 150, GridTable.Green);
            Button("2", "Cancel", 162, 634, 70, GridTable.Grey);
            Button("bLastPo", "", 236, 634, 190, GridTable.Green);
            Visible("bLastPo", false);
            Label("lFoot", "", 436, 637, 260);
            Label("lTotal", "Total", 740, 637, 90);
            Field("eTotal", 834, 634, 130, true);

            // A blank first entry so the combo itself can clear the filter (it is a strict pick list, not free text)
            var manChoices = new List<KeyValuePair<string, string>> { new KeyValuePair<string, string>("", "(All Manufacturers)") };
            manChoices.AddRange(Store.Manufacturers());
            Choices("cMan", manChoices);
            Choices("cSeries", Store.SeriesFor(ObjPurchaseOrder));
            int series = Store.DefaultSeriesFor(ObjPurchaseOrder);
            if (series > 0) Set("cSeries", series.ToString(Inv));

            multiBranch = Store.MultiBranch;
            Visible("lBranch", multiBranch); Visible("cBranch", multiBranch);
            if (multiBranch)
            {
                Choices("cBranch", Store.Branches());
                int branch = Store.DefaultBranch();
                if (branch > 0) Set("cBranch", branch.ToString(Inv));
            }

            SetDate("eDue", DateTime.Today.AddDays(7));
            Set("eTotal", 0m);
            Set("cReqAll", "Y");
            Set("cRowAll", "Y");
            RefreshNextNo();
            RefreshVendorDisplay();
            ShowRows();
        }

        /// <summary>
        /// 50:50 (by default) split of the space between the two grids. Everything else between them (the Load /
        /// Select-All / splitter row, the Purchase Order Details fields) has a fixed height, so only the remaining
        /// space - what the two grids actually compete for - is divided by splitRatio.
        /// </summary>
        protected override void Layout()
        {
            // Small floors only to keep every width/height positive - the real limit is whatever Create() already fit
            // to the SAP Desktop, which on a small monitor can be well under the form's original design size.
            int w = Math.Max(600, Form.ClientWidth), h = Math.Max(400, Form.ClientHeight);
            Place("bSwitch", Math.Max(1280, w - 16 - 220), 28, 0, 0);    // never left of bClear (ends at 1272), however narrow w gets
            Place("hFiltL", 8, 24, w - 16, 0);
            Place("hReqL", 8, 78, w - 16, 0);

            const int top = 86;                 // gReq's top edge, fixed
            const int bottomReserve = 40;        // footer button row below gRows, fixed
            const int minGridH = 80;             // neither grid collapses below this
            const int gapToFoot = 6, gapToHead = 8, headToRow1 = 20, row1ToRow2 = 27, row2ToMat = 27, matToGrid = 24;
            const int midFixed = gapToFoot + ButtonH + gapToHead + headToRow1 + row1ToRow2 + row2ToMat + matToGrid;

            int avail = Math.Max(2 * minGridH, h - top - midFixed - bottomReserve);
            int reqH = Math.Max(minGridH, (int)Math.Round(avail * splitRatio));
            int rowsH = Math.Max(minGridH, avail - reqH);

            Place("gReq", 8, top, w - 16, reqH);
            gReq.FillWidth("Desc", w - 16);

            // Load / Select-All / splitter share one row: the splitter buttons sit right-aligned on it, parallel to
            // Select / Deselect All, same as the Apply row on Material Request Planning.
            int rowY = top + reqH + gapToFoot;
            Place("bLoad", 8, rowY, 0, 0);
            Place("cReqAll", 176, rowY, 0, 0);
            int splitRstX = w - 16 - 36 - 36 - 140;
            Place("lReqFoot", 346, rowY + 3, Math.Max(0, splitRstX - 10 - 346), 0);
            Place("bSplitRst", splitRstX, rowY, 140, 0);
            Place("bSplitUp", w - 16 - 36 - 36, rowY, 36, 0);
            Place("bSplitDn", w - 16 - 36, rowY, 36, 0);
            int y = rowY + ButtonH;

            int headY = y + gapToHead;
            Place("hHead", 8, headY, 300, 0);
            Place("hHeadL", 8, headY + SectionH, w - 16, 0);

            int row1Y = headY + headToRow1;
            Place("lSeries", 8, row1Y + 2, 0, 0);
            Place("cSeries", 60, row1Y, 0, 0);
            Place("lNextNo", 184, row1Y + 2, 0, 0);
            Place("fNextNo", 242, row1Y, 0, 0);
            Place("lBranch", 316, row1Y + 2, 0, 0);
            Place("cBranch", 368, row1Y, 0, 0);
            Place("lDue", 502, row1Y + 2, 0, 0);
            Place("eDue", 591, row1Y, 0, 0);
            Place("lVRef", 692, row1Y + 2, 0, 0);
            Place("eVRef", 771, row1Y, 0, 0);

            int row2Y = row1Y + row1ToRow2;
            Place("lVend2", 8, row2Y + 2, 0, 0);
            Place("fVend", 60, row2Y, 0, 0);
            Place("lTaxApply", 240, row2Y + 2, 0, 0);
            Place("eTaxApply", 304, row2Y, 0, 0);
            Place("bApplyTax", 398, row2Y - 2, 160, 0);
            Place("cRowAll", 607, row2Y - 1, 0, 0);

            int matY = row2Y + row2ToMat;
            Place("hMat", 8, matY, 300, 0);
            Place("hMatL", 8, matY + SectionH, w - 16, 0);

            int gRowsY = matY + matToGrid;
            Place("gRows", 8, gRowsY, w - 16, rowsH);
            gRows.FillWidth("Desc", w - 16);

            Place("bCreate", 8, h - 30, 0, 0);
            Place("2", 162, h - 30, 0, 0);
            Place("bLastPo", 236, h - 30, 0, 0);
            Place("lFoot", 436, h - 27, 0, 0);
            Place("lTotal", w - 240, h - 27, 0, 0);
            Place("eTotal", w - 146, h - 30, 0, 0);
        }

        /// <summary>Adjusts the split (clamped so neither grid can be squeezed away) and re-lays out immediately, without waiting for a resize.</summary>
        private void SetSplit(double ratio)
        {
            splitRatio = Math.Max(0.15, Math.Min(0.85, ratio));
            Freeze(Layout);
        }
        #endregion

        #region Search and load
        private static List<string> Codes(string text)
        {
            return text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(c => c.Trim()).Where(c => c.Length > 0).ToList();
        }

        /// <summary>Purchase Request numbers currently typed in eFind (only when every token parses as a number).</summary>
        private List<int> FindDocNums()
        {
            var nums = new List<int>();
            foreach (string t in Codes(Text("eFind")))
            {
                int n;
                if (!int.TryParse(t, out n)) return new List<int>();
                nums.Add(n);
            }
            return nums;
        }

        /// <summary>A brand new, empty Conditions object: SetConditions always replaces the Choose From List's filter
        /// with exactly what is passed, so building a fresh one each time avoids old conditions piling up (Conditions
        /// has no Remove method to take them back out).</summary>
        private SAPbouiCOM.Conditions NewConditions()
        {
            return (SAPbouiCOM.Conditions)App.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
        }

        /// <summary>
        /// Narrows the Purchase Request Choose From List (eFind) to open requests matching the Project / Vendor
        /// currently typed (either of them, both, or neither - same matching as Search itself).
        /// </summary>
        private void FilterFindCfl()
        {
            SAPbouiCOM.Conditions cons = NewConditions();
            string proj = Text("eProj"), vend = Text("eVend");
            if (proj.Length > 0 || vend.Length > 0)
            {
                // Conditions has no clean way to AND a status check with an OR'd list of document numbers, so the
                // matching numbers are found with the same store query Search uses (already limited to open requests).
                List<int> nums = Store.ListRequestLines(Codes(proj), Codes(vend), "").Select(r => r.PrNum).Distinct().ToList();
                if (nums.Count == 0) nums.Add(-1);
                for (int i = 0; i < nums.Count; i++)
                {
                    SAPbouiCOM.Condition con = cons.Add();
                    con.Alias = "DocNum";
                    con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    con.CondVal = nums[i].ToString(Inv);
                    if (i < nums.Count - 1) con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                }
            }
            else
            {
                SAPbouiCOM.Condition con = cons.Add();
                con.Alias = "DocStatus";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "O";
            }
            GetCfl("eFind").SetConditions(cons);
        }

        /// <summary>Narrows the Vendor Choose From List (eVend) to vendors of open Purchase Requests of the Project currently typed, if any. Vendor-type (CardType = S) stays applied, same as when the list was created.</summary>
        private void FilterVendorCfl()
        {
            SAPbouiCOM.Conditions cons = NewConditions();
            SAPbouiCOM.Condition typeCon = cons.Add();
            typeCon.Alias = "CardType";
            typeCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            typeCon.CondVal = "S";
            string proj = Text("eProj");
            if (proj.Length > 0)
            {
                List<string> vendors = Store.VendorsOfProjects(Codes(proj));
                if (vendors.Count > 0)
                {
                    typeCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    for (int i = 0; i < vendors.Count; i++)
                    {
                        SAPbouiCOM.Condition con = cons.Add();
                        con.Alias = "CardCode";
                        con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        con.CondVal = vendors[i];
                        if (i < vendors.Count - 1) con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                    }
                }
            }
            GetCfl("eVend").SetConditions(cons);
        }

        /// <summary>Narrows the Project Choose From List (eProj) to projects of open Purchase Requests of the Vendor currently typed, if any.</summary>
        private void FilterProjectCfl()
        {
            SAPbouiCOM.Conditions cons = NewConditions();
            string vend = Text("eVend");
            if (vend.Length > 0)
            {
                List<string> projects = Store.ProjectsOfVendors(Codes(vend));
                for (int i = 0; i < projects.Count; i++)
                {
                    SAPbouiCOM.Condition con = cons.Add();
                    con.Alias = "PrjCode";
                    con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    con.CondVal = projects[i];
                    if (i < projects.Count - 1) con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                }
            }
            GetCfl("eProj").SetConditions(cons);
        }

        /// <summary>Sets eProj without triggering ClearOnProjectChange for the Validate event SAP raises right after.</summary>
        private void SetProject(string project)
        {
            settingProject = true;
            Set("eProj", project);
        }

        /// <summary>Sets eVend without triggering a second, redundant Search() for the Validate event SAP raises right after.</summary>
        private void SetVendor(string vendor)
        {
            settingVendor = true;
            Set("eVend", vendor);
        }

        /// <summary>
        /// Initialises Project and Vendor from the given Purchase Requests (overwriting what is there) when they
        /// resolve to exactly one: a request commonly carries lines of more than one Project / Vendor, so picking one
        /// is left blank rather than refused - the line picker grid below is what narrows to a single Vendor.
        /// </summary>
        private void InitFromRequests(IEnumerable<int> docNums)
        {
            var nums = docNums.ToList();
            List<string> projects = Store.ProjectsOfRequests(nums).Distinct().ToList();
            List<string> vendors = Store.VendorsOfRequests(nums).Distinct().ToList();
            if (projects.Count == 1) SetProject(projects[0]);
            if (vendors.Count == 1) SetVendor(vendors[0]);
        }

        private void Search()
        {
            // Read what is typed, then write it straight back into the bound data source: a button press moves the
            // focus off the edit field, and without this the field can revert to its last committed (empty) value.
            string proj = Text("eProj"), vend = Text("eVend"), find = Text("eFind");
            Set("eProj", proj);
            Set("eVend", vend);
            Set("eFind", find);

            requests = Store.ListRequestLines(Codes(proj), Codes(vend), find);
            ShowRequests();
            RefreshVendorDisplay();
            if (requests.Count == 0) Warn("No open Purchase Request line found for this search.");
        }

        /// <summary>Mirrors the Vendor filter into the read-only header field, so the single vendor the Purchase Order will target stays visible once requests are loaded.</summary>
        private void RefreshVendorDisplay()
        {
            Set("fVend", Text("eVend"));
        }

        /// <summary>Requests gReq should currently show: all of them, or just those of the chosen Manufacturer.</summary>
        private List<PurchaseRequestLine> RequestsVisible()
        {
            string man = Get("cMan");
            return man.Length == 0 ? requests : requests.Where(r => r.Manufacturer == man).ToList();
        }

        private void ShowRequests()
        {
            reqView = gReq.Sort(RequestsVisible(), (c, r) => c == "Catalog" ? r.Catalog : c == "Item" ? r.ItemCode
                : c == "Desc" ? r.Description : c == "Man" ? r.ManufacturerName : c == "Vend" ? r.Vendor
                : c == "Stat" ? (r.Closed ? "Closed" : "Open") : c == "Whs" ? r.Warehouse
                : c == "Date" ? r.PostingDate : "");     // yyyy-MM-dd: text order = date order
            Freeze(() =>
            {
                gReq.Fill(reqView.Select((r, i) => new object[]
                {
                    // A closed line is always shown unticked, whatever Selected was last set to - it cannot be drawn
                    // from again, and RefreshClosedRequests() already forces Selected false the moment a line closes.
                    i + 1, r.Closed ? "N" : (r.Selected ? "Y" : "N"), r.PrNum, r.PrLine, r.SeriesName, r.Project, r.Vendor, r.Catalog,
                    r.ItemCode, r.Description, r.ManufacturerName, r.JobCode, r.Quantity, r.OpenQuantity, r.Warehouse,
                    r.Closed ? "Closed" : "Open", ToDate(r.PostingDate)
                }).ToList());
                // Closed (document closed, line closed, or nothing left open on it - already fully drawn into a
                // Purchase Order) always shows green, overriding the unselected-grey below: it is a stronger signal
                // ("do not draw a Purchase Order from this line again") than plain unselected. Its checkbox is also
                // actually disabled here (not just reverted after a click - see the gReq tick handler, kept as a
                // second line of defence), so a closed line visibly cannot be ticked in the first place.
                for (int i = 0; i < reqView.Count; i++)
                {
                    gReq.Color(i, reqView[i].Closed ? GridTable.Green : !reqView[i].Selected ? GridTable.Grey : GridTable.NoColor);
                    gReq.Lock(i, "Sel", !reqView[i].Closed);
                }
                string filtered = reqView.Count != requests.Count ? " (filtered: " + requests.Count + " line(s) in total)" : "";
                Caption("lReqFoot", requests.Count == 0 ? "No open Purchase Request line found for this search."
                    : reqView.Count(r => r.Selected) + " of " + reqView.Count + " line(s) selected" + filtered);
            });
        }

        /// <summary>Select All / Deselect All act only on the lines currently visible under the Manufacturer filter; a
        /// line hidden by the filter keeps whatever Selected state it already had.</summary>
        private void SetAllRequestsSelected(bool on)
        {
            if (reqView.Count == 0) return;
            foreach (PurchaseRequestLine r in reqView) r.Selected = on;
            ShowRequests();
        }

        private void Load()
        {
            var ticked = requests.Where(r => r.Selected).ToList();
            if (ticked.Count == 0) throw new PlanException(400, "Tick one or more Purchase Request lines first.");
            // The Purchase Order targets one Vendor: when one is typed, lines of any other Vendor are left out here
            // rather than rejecting the whole pick - a ticked set commonly spans more than one Vendor.
            string vendor = Text("eVend");
            var eligible = vendor.Length == 0 ? ticked : ticked.Where(r => string.Equals(r.Vendor, vendor, StringComparison.OrdinalIgnoreCase)).ToList();
            if (eligible.Count == 0) throw new PlanException(400, "None of the ticked lines belong to Vendor " + vendor + ".");
            var picked = eligible.Select(r => new KeyValuePair<int, int>(r.PrEntry, r.PrLine)).Distinct().ToList();
            rows = PlanService.LoadPoRows(picked);
            stock = Store.GetStock(rows.Select(r => r.ItemCode));
            HideLastPo();
            ShowRows();
            int skipped = ticked.Count - eligible.Count;
            if (rows.Count == 0) Alert("Nothing to order: the matching line(s) are no longer open.");
            else Done(rows.Count + " material row(s) loaded." + (skipped > 0 ? " " + skipped + " ticked line(s) of a different Vendor were skipped." : ""));
        }

        private void Clear()
        {
            Set("eProj", "");
            Set("eVend", "");
            Set("eFind", "");
            requests = new List<PurchaseRequestLine>();
            rows = new List<PoRow>();
            HideLastPo();
            ShowRequests();
            ShowRows();
            RefreshVendorDisplay();
            Ok("Filters and rows cleared.");
        }

        /// <summary>
        /// Project was typed or picked afresh: Vendor, Purchase Request and whatever was loaded no longer match it, so
        /// they are cleared (Project itself is left as the user just set it) and the grid reloads for the new Project
        /// on its own, same as pressing Search.
        /// </summary>
        private void ClearOnProjectChange()
        {
            Set("eVend", "");
            Set("eFind", "");
            rows = new List<PoRow>();
            HideLastPo();
            Search();
        }
        #endregion

        #region Rows
        private static string Base(PoRow r)
        {
            string b = r.PrNum + " / " + r.PrLine;
            return r.MergedFrom != null && r.MergedFrom.Count > 1 ? b + " (+" + (r.MergedFrom.Count - 1) + ")" : b;
        }

        /// <summary>
        /// Rows gRows should currently show: all of them (or just those of the chosen Manufacturer), with same-Item +
        /// same-Manufacturer rows combined into one summary row (quantity / open quantity summed; a different
        /// Manufacturer for the same item is kept as its own row; Price / Discount / Price After Discount from the member
        /// with the lowest Price After Discount, copied onto every member so each Purchase Order line is raised at that
        /// price, and Total = that price x the summed Quantity). A single matching row is returned unchanged - only
        /// an actual duplicate is ever wrapped. The underlying "rows" list is never touched here: each source line
        /// still becomes its own Purchase Order line at Create time (SAP draws one PO line from exactly one base
        /// Purchase Request line), so a merged row's edits are propagated back onto its members - see ReadRow.
        /// </summary>
        private List<PoRow> Visible()
        {
            string man = Get("cMan");
            var filtered = man.Length == 0 ? rows : rows.Where(r => r.Manufacturer == man).ToList();
            return filtered
                .GroupBy(r => r.ItemCode + "|" + r.Manufacturer)
                .OrderBy(g => g.First().ItemCode).ThenBy(g => g.First().ManufacturerName)
                .Select(g =>
                {
                    var members = g.ToList();
                    if (members.Count == 1) return members[0];
                    PoRow first = members[0];
                    PoRow cheapest = PlanService.Cheapest(members, m => m.PriceAfterDiscount);
                    foreach (PoRow m in members) { m.Price = cheapest.Price; m.Discount = cheapest.Discount; }
                    // One price for every member, so it must stay within the tightest member's cap (0 = no limit)
                    decimal cap = members.Select(m => m.MaxPriceAfterDiscount).Where(c => c > 0).DefaultIfEmpty(0).Min();
                    return new PoRow
                    {
                        Selected = members.Any(m => m.Selected),
                        PrEntry = first.PrEntry, PrNum = first.PrNum, PrLine = first.PrLine,
                        ItemCode = first.ItemCode, Description = first.Description, Uom = first.Uom,
                        Quantity = members.Sum(m => m.Quantity), OpenQuantity = members.Sum(m => m.OpenQuantity),
                        Price = cheapest.Price, Discount = cheapest.Discount, MaxPriceAfterDiscount = cap,
                        Warehouse = first.Warehouse, Vendor = first.Vendor, Project = first.Project, TaxCode = first.TaxCode,
                        Manufacturer = first.Manufacturer, ManufacturerName = first.ManufacturerName,
                        Catalog = first.Catalog, JobCode = first.JobCode, MergedFrom = members
                    };
                }).ToList();
        }

        private void ShowRows()
        {
            view = gRows.Sort(Visible(), (c, r) => c == "Catalog" ? r.Catalog : c == "Item" ? r.ItemCode
                : c == "Desc" ? r.Description : c == "Man" ? r.ManufacturerName : c == "Whs" ? r.Warehouse : "");
            foreach (PoRow r in rows) r.SrNo = 0;
            for (int i = 0; i < view.Count; i++)
            {
                view[i].SrNo = i + 1;
                if (view[i].MergedFrom != null) foreach (PoRow m in view[i].MergedFrom) m.SrNo = i + 1;
            }
            Freeze(() =>
            {
                gRows.Fill(view.Select((r, i) => new object[]
                {
                    i + 1, r.Selected ? "Y" : "N", r.Catalog, r.ItemCode, r.Description, r.ManufacturerName, r.JobCode, r.Quantity, r.OpenQuantity, r.Price, r.Discount, r.PriceAfterDiscount, r.Total,
                    r.Warehouse, stock.InStock(r.ItemCode, r.Warehouse), stock.ProjectStock(r.ItemCode, r.Warehouse, r.Project), stock.MainLocStock(r.ItemCode),
                    r.Project, r.TaxCode, Base(r)
                }).ToList());
                for (int i = 0; i < view.Count; i++)
                    gRows.Color(i, !view[i].Selected ? GridTable.Grey : view[i].Quantity > view[i].OpenQuantity ? GridTable.Red
                        : view[i].TaxCode.Length == 0 ? GridTable.Yellow : GridTable.NoColor);
                // The Total stays the grand total of every selected row (any Manufacturer): what Create would actually
                // order, not just what happens to be visible under the current filter.
                decimal total = rows.Where(r => r.Selected).Sum(r => r.Total);
                Set("eTotal", total);
                string man = Get("cMan");
                int manCount = man.Length == 0 ? rows.Count : rows.Count(r => r.Manufacturer == man);
                string filtered = man.Length > 0 && manCount != rows.Count
                    ? " (filtered: " + manCount + " of " + rows.Count + " row(s) match Manufacturer " + man + ")" : "";
                // Same-item rows are grouped into one row here for easier editing (see Visible()), but SAP still draws
                // one Purchase Order line per source Purchase Request line underneath - every grouped member becomes
                // its own line at Create time (see ReadRow / CreateOrder). Whenever that would create more lines than
                // are shown selected, say so here up front instead of surprising the user with a bigger Purchase
                // Order than the grid implied.
                int realLines = rows.Count(r => r.Selected);
                int shownSelected = view.Count(r => r.Selected);
                string grouped = realLines != shownSelected
                    ? " - will create " + realLines + " Purchase Order line(s) (rows above are grouped by item)" : "";
                if (rows.Count > 0) Caption("lFoot", shownSelected + " of " + view.Count + " row(s) selected" + filtered + grouped);
            });
        }

        /// <summary>Takes what the user typed in one visible grid row into the model. A merged row (see Visible) has
        /// no Purchase Order line of its own, so every edit is also propagated onto its member rows - Warehouse / Tax Code / Selected applied as typed to each member, Quantity redistributed across them in
        /// their original proportion (each member still capped by its own source line, enforced at Create time).
        /// Price / Discount % are taken as typed; the Production Order's limit (MaxPriceAfterDiscount) is checked only
        /// when the Purchase Order button is pressed (PlanService.ValidatePo).</summary>
        private void ReadRow(int tr)
        {
            PoRow r = view[tr];
            r.Selected = gRows.TypedText(tr, "Sel") == "Y";
            r.Quantity = gRows.Typed(tr, "Qty");
            r.Warehouse = gRows.TypedText(tr, "Whs");
            r.TaxCode = gRows.TypedText(tr, "Tax");
            r.Price = gRows.Typed(tr, "Price");
            r.Discount = gRows.Typed(tr, "Disc");
            if (r.MergedFrom != null)
            {
                decimal originalTotal = r.MergedFrom.Sum(m => m.Quantity);
                foreach (PoRow m in r.MergedFrom)
                {
                    m.Selected = r.Selected;
                    m.Quantity = originalTotal > 0 ? Math.Round(r.Quantity * (m.Quantity / originalTotal), 6) : r.Quantity / r.MergedFrom.Count;
                    m.Warehouse = r.Warehouse; m.TaxCode = r.TaxCode;
                    m.Price = r.Price; m.Discount = r.Discount;
                }
            }
        }

        /// <summary>The row whose Price was edited and not yet checked against the Production Order's limit - checked
        /// when its Discount % is changed or the cursor moves to another row (and again on the Purchase Order button).</summary>
        private string pricePending;

        /// <summary>Identifies a grid row across redraws (a merged row is rebuilt on every ShowRows): its first source line.</summary>
        private static string Key(PoRow r) { return r == null ? "" : r.PrEntry + "/" + r.PrLine; }

        /// <summary>Alerts when the row's Price less Discount % is above the Production Order's Price After Discount.</summary>
        private void CheckPrice(string key)
        {
            pricePending = null;
            PoRow r = view.FirstOrDefault(v => Key(v) == key);
            if (r == null || !PlanService.PriceOverCap(r.PriceAfterDiscount, r.MaxPriceAfterDiscount)) return;
            throw new PlanException(400, "Sr No " + r.SrNo + " (" + r.ItemCode + "): " + PlanService.PriceCapMessage(r.PriceAfterDiscount, r.MaxPriceAfterDiscount), "gRows");
        }

        private PoHeader ReadHeader()
        {
            int series, branch = 0;
            int.TryParse(Get("cSeries"), out series);
            if (multiBranch) int.TryParse(Get("cBranch"), out branch);
            return new PoHeader { Series = series, Branch = branch, DueDate = Date("eDue"), Vendor = Text("eVend"), VendorRef = Text("eVRef") };
        }

        /// <summary>Select All / Deselect All act only on the rows currently visible under the Manufacturer filter; a
        /// row hidden by the filter keeps whatever Selected state it already had.</summary>
        private void SetAllSelected(bool on)
        {
            if (view.Count == 0) return;
            foreach (PoRow r in view)
            {
                r.Selected = on;
                if (r.MergedFrom != null) foreach (PoRow m in r.MergedFrom) m.Selected = on;
            }
            ShowRows();
        }

        /// <summary>Stamps the Tax Code typed/picked in eTaxApply onto every currently-selected Materials row (and,
        /// for a merged row, onto its underlying member rows too - same propagation as a manual per-row edit in
        /// ReadRow) - one click instead of fixing the Tax column line by line after a rejection such as
        /// "(line 1) you must choose IGST (-10)".</summary>
        private void ApplyTaxCode()
        {
            string tax = Text("eTaxApply");
            if (tax.Length == 0) throw new PlanException(400, "Pick a Tax Code first.");
            if (!view.Any(r => r.Selected)) throw new PlanException(400, "Select at least one material row first.");
            foreach (PoRow r in view.Where(r => r.Selected))
            {
                r.TaxCode = tax;
                if (r.MergedFrom != null) foreach (PoRow m in r.MergedFrom) m.TaxCode = tax;
            }
            ShowRows();
            Done("Tax Code " + tax + " applied to " + view.Count(r => r.Selected) + " row(s).");
        }

        /// <summary>
        /// Fires automatically when the Manufacturer combo changes: re-filters both grids (Purchase Request Lines and,
        /// once loaded, Materials) to just that Manufacturer's rows (the blank entry shows everything again). Does
        /// not touch any row's Selected state - filtering is a display concern, ticks stay exactly as the user left
        /// them even for rows the filter now hides.
        /// </summary>
        private void FilterByManufacturer()
        {
            ShowRequests();
            ShowRows();
            string man = Get("cMan");
            if (man.Length == 0) Ok("Showing all line(s).");
            else if (reqView.Count == 0 && rows.Count == 0) Warn("No line matches Manufacturer " + man + ".");
            else Ok(reqView.Count + " of " + requests.Count + " Purchase Request line(s)" +
                (rows.Count > 0 ? ", " + view.Count + " of " + rows.Count + " material row(s)," : "") +
                " match Manufacturer " + man + ".");
        }

        private void CreateOrder()
        {
            for (int i = 0; i < view.Count; i++) ReadRow(i);
            pricePending = null;
            // Same-item rows are grouped into one row for display (see Visible()), but each one still becomes its own
            // Purchase Order line at Create time - SAP draws a line from exactly one source Purchase Request line, so
            // a merged row of 3 source lines becomes 3 Purchase Order lines, not 1. Confirm before creating whenever
            // that means more lines than the grid showed selected, so the real count is never a surprise.
            int realLines = rows.Count(r => r.Selected);
            int shownSelected = view.Count(r => r.Selected);
            if (realLines != shownSelected && !Ask(
                shownSelected + " row(s) are selected above, grouped by item. SAP draws one Purchase Order line per " +
                "source Purchase Request line, so this will create " + realLines + " Purchase Order line(s). Continue?",
                "Create", "Cancel")) return;
            var drawnKeys = rows.Where(r => r.Selected).Select(r => new KeyValuePair<int, int>(r.PrEntry, r.PrLine)).Distinct().ToList();
            CreatedOrder made = PlanService.CreatePo(ReadHeader(), rows);
            // One order per load: the rows are cleared so Create cannot be pressed twice for the same requests
            rows = new List<PoRow>();
            ShowRows();
            lastPoEntry = made.DocEntry; lastPoNum = made.DocNum;
            ButtonCaption("bLastPo", "Open Purchase Order " + made.DocNum);
            Visible("bLastPo", true);
            RefreshClosedRequests(drawnKeys);
            Done("Purchase Order " + made.DocNum + " was created with " + realLines + " line(s).");
        }

        /// <summary>
        /// After drawing a Purchase Order from some ticked Purchase Request lines, re-reads exactly those lines -
        /// whatever their status is now (see PlanService.RefreshRequestLines) - and updates the matching rows of
        /// "requests" in place. A line fully closed as a result is shown locked (green, unticked - see ShowRequests)
        /// instead of just vanishing the way a brand new search would: ListRequestLines only ever returns still-open
        /// lines, so the picker grid would otherwise lose track of a line the moment it closes.
        /// </summary>
        private void RefreshClosedRequests(List<KeyValuePair<int, int>> keys)
        {
            if (keys.Count == 0) return;
            var fresh = PlanService.RefreshRequestLines(keys).ToDictionary(r => r.PrEntry + "/" + r.PrLine);
            foreach (PurchaseRequestLine r in requests)
            {
                PurchaseRequestLine f;
                if (!fresh.TryGetValue(r.PrEntry + "/" + r.PrLine, out f)) continue;
                r.Status = f.Status; r.LineStatus = f.LineStatus; r.OpenQuantity = f.OpenQuantity; r.Quantity = f.Quantity;
                if (r.Closed) r.Selected = false;
            }
            ShowRequests();
        }

        private void HideLastPo()
        {
            lastPoEntry = 0;
            Visible("bLastPo", false);
        }

        private void RefreshNextNo()
        {
            int series;
            int.TryParse(Get("cSeries"), out series);
            int next = Store.NextNumberFor(series, ObjPurchaseOrder);
            Set("fNextNo", next > 0 ? next.ToString(Inv) : "");
        }
        #endregion

        #region Events
        protected override void OnEvent(SAPbouiCOM.ItemEvent e, ref bool bubble)
        {
            List<string> chosen;
            int gridRow;
            if (Key(e, 'F', SAPbouiCOM.BoModifiersEnum.mt_CTRL)) Focus("eFind");
            else if (Key(e, 116)) Search();                                                                   // F5
            else if (Key(e, 'L', SAPbouiCOM.BoModifiersEnum.mt_ALT)) Focus("gRows");
            else if (Key(e, 13, SAPbouiCOM.BoModifiersEnum.mt_CTRL)) CreateOrder();
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && e.BeforeAction && e.ItemUID == "eFind") FilterFindCfl();
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && e.BeforeAction && e.ItemUID == "eVend") FilterVendorCfl();
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && e.BeforeAction && e.ItemUID == "eProj") FilterProjectCfl();
            else if ((chosen = ChosenCodes(e, "eProj", "PrjCode")).Count > 0)
            {
                // Picking a Project narrows the request list to it right away, same as pressing Search
                SetProject(chosen[0]);
                ClearOnProjectChange();
            }
            else if ((chosen = ChosenCodes(e, "eVend", "CardCode")).Count > 0)
            {
                // Picking a Vendor filters the request list to it (and the Project, when it alone matches) right away
                SetVendor(chosen[0]);
                List<string> projects = Store.ProjectsOfVendors(new[] { chosen[0] }).Distinct().ToList();
                if (projects.Count == 1) SetProject(projects[0]);
                RefreshVendorDisplay();
                Search();
            }
            else if ((chosen = ChosenCodes(e, "eTaxApply", "Code")).Count > 0) Set("eTaxApply", chosen[0]);
            else if ((chosen = ChosenCodes(e, "eFind", "DocNum")).Count > 0)
            {
                Set("eFind", string.Join(",", chosen.ToArray()));
                var nums = chosen.Select(c => { int n; return int.TryParse(c, out n) ? n : 0; }).Where(n => n > 0).ToList();
                InitFromRequests(nums);
                RefreshVendorDisplay();
                Search();
            }
            else if (e.ItemUID == "eProj" && !e.BeforeAction && e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ItemChanged)
            {
                if (settingProject) settingProject = false;
                else ClearOnProjectChange();
            }
            else if (e.ItemUID == "eVend" && !e.BeforeAction && e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ItemChanged)
            {
                // Vendor typed by hand (not picked) and committed: reload the grid for Project + Vendor together, keeping the Project filter
                if (settingVendor) settingVendor = false;
                else { RefreshVendorDisplay(); Search(); }
            }
            else if ((chosen = ChosenGridCodes(e, gRows, "Whs", "WhsCode", out gridRow)).Count > 0 && gridRow >= 0 && gridRow < view.Count)
            {
                view[gridRow].Warehouse = chosen[0];
                if (view[gridRow].MergedFrom != null) foreach (PoRow m in view[gridRow].MergedFrom) m.Warehouse = chosen[0];
                ShowRows();
                KeepOkMode("Cancel");
            }
            else if ((chosen = ChosenGridCodes(e, gRows, "Tax", "Code", out gridRow)).Count > 0 && gridRow >= 0 && gridRow < view.Count)
            {
                view[gridRow].TaxCode = chosen[0];
                if (view[gridRow].MergedFrom != null) foreach (PoRow m in view[gridRow].MergedFrom) m.TaxCode = chosen[0];
                ShowRows();
                KeepOkMode("Cancel");
            }
            else if (Enter(e, "eFind", "eProj", "eVend")) Search();
            else if (Pressed(e, "bClear")) Clear();
            else if (Pressed(e, "bSwitch")) MaterialPlanForm.Show();
            else if (e.ItemUID == "cSeries" && !e.BeforeAction &&
                (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)) RefreshNextNo();
            else if (Pressed(e, "bLoad")) Load();
            else if (Pressed(e, "bSplitUp")) SetSplit(splitRatio + 0.05);
            else if (Pressed(e, "bSplitDn")) SetSplit(splitRatio - 0.05);
            else if (Pressed(e, "bSplitRst")) SetSplit(0.5);
            else if (Pressed(e, "cReqAll")) SetAllRequestsSelected(Get("cReqAll") == "Y");
            else if (Pressed(e, "cRowAll")) SetAllSelected(Get("cRowAll") == "Y");
            else if (e.ItemUID == "cMan" && !e.BeforeAction &&
                (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)) FilterByManufacturer();
            else if (Pressed(e, "bManFilter")) FilterByManufacturer();
            else if (Pressed(e, "bApplyTax")) ApplyTaxCode();
            else if (Pressed(e, "bCreate")) CreateOrder();
            else if (Pressed(e, "bLastPo")) { if (lastPoEntry > 0) Store.OpenPurchaseOrder(lastPoEntry); }
            else if (e.ItemUID == "gRows" && !e.BeforeAction && e.Row >= 0 && pricePending != null &&
                (e.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS || e.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) &&
                Key(RowAt(e.Row)) != pricePending)
            {
                // The cursor left the row whose Price was edited: check it now
                if (e.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) gRows.Click(e.Row);
                CheckPrice(pricePending);
            }
            else if (Clicked(e, gRows)) gRows.Click(e.Row);
            else if (e.ItemUID == "gReq" && !e.BeforeAction && e.Row >= 0 &&
                ((e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ItemChanged) || (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && e.ColUID == "Sel")))
            {
                int tr = gReq.Row(e.Row);
                if (tr < 0 || tr >= reqView.Count) return;
                reqView[tr].Selected = gReq.TypedText(tr, "Sel") == "Y";
                // A closed line's checkbox is already disabled (see ShowRequests -> gReq.Lock), so this is only a
                // second line of defence for however it still got ticked. Status bar only, not a blocking message
                // box - shown in error red (smt_Error), not Warn()'s yellow, since this blocks a real action (the
                // line really cannot be drawn from) rather than just flagging something to double-check.
                if (reqView[tr].Selected && reqView[tr].Closed)
                {
                    reqView[tr].Selected = false;
                    ShowRequests();
                    Status("Purchase Request " + reqView[tr].PrNum + " line " + reqView[tr].PrLine + " is closed and cannot be drawn from again.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
                if (reqView[tr].Selected)
                {
                    // Each line already carries its own Project / Vendor, loaded with the search - no need to re-query.
                    // Ticking lines of more than one Vendor is allowed here (a request commonly mixes them); Load
                    // Selected Requests is what narrows to the Vendor currently typed, not this tick.
                    var selected = requests.Where(r => r.Selected).ToList();
                    List<string> projects = selected.Select(r => r.Project).Where(p => p.Length > 0).Distinct().ToList();
                    List<string> vendors = selected.Select(r => r.Vendor).Where(v => v.Length > 0).Distinct().ToList();
                    // Suppressed: ticking rows must not itself trigger a re-Search, which would replace the very list being ticked
                    if (projects.Count == 1) SetProject(projects[0]);
                    if (vendors.Count == 1) SetVendor(vendors[0]);
                }
                ShowRequests();
                RefreshVendorDisplay();
                KeepOkMode("Cancel");
            }
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && e.BeforeAction && e.ItemUID == "gReq" && e.ColUID == "No")
            {
                // SAP's own link would look up the number shown; the Purchase Request is opened by its DocEntry
                bubble = false;
                int tr = gReq.Row(e.Row);
                if (tr >= 0 && tr < reqView.Count) Store.OpenPurchaseRequest(reqView[tr].PrEntry);
            }
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && e.BeforeAction && e.ItemUID == "gRows" && e.ColUID == "Base")
            {
                // SAP's own link would look up the number shown; the Purchase Request is opened by its DocEntry
                bubble = false;
                int tr = gRows.Row(e.Row);
                if (tr >= 0 && tr < view.Count) Store.OpenPurchaseRequest(view[tr].PrEntry);
            }
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && e.BeforeAction && e.ItemUID == "gReq" && e.ColUID == "Item")
            {
                bubble = false;
                int tr = gReq.Row(e.Row);
                if (tr >= 0 && tr < reqView.Count) Store.OpenItem(reqView[tr].ItemCode);
            }
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && e.BeforeAction && e.ItemUID == "gRows" && e.ColUID == "Item")
            {
                bubble = false;
                int tr = gRows.Row(e.Row);
                if (tr >= 0 && tr < view.Count) Store.OpenItem(view[tr].ItemCode);
            }
            else if (e.ItemUID == "gRows" && !e.BeforeAction && e.Row >= 0 &&
                ((e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ItemChanged) || (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && e.ColUID == "Sel")))
            {
                int tr = gRows.Row(e.Row);
                if (tr < 0 || tr >= view.Count) return;
                gRows.Click(e.Row);
                ReadRow(tr);
                string edited = Key(view[tr]);
                ShowRows();
                KeepOkMode("Cancel");
                if (e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ColUID == "Disc") CheckPrice(edited);
                else if (e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ColUID == "Price") pricePending = edited;
            }
        }

        /// <summary>The model row shown on the given grid row (null when outside the grid).</summary>
        private PoRow RowAt(int gridRow)
        {
            int tr = gRows.Row(gridRow);
            return tr >= 0 && tr < view.Count ? view[tr] : null;
        }

        /// <summary>A failed check puts the cursor on the field that has to be corrected; everything entered stays.</summary>
        protected override void Fail(Exception ex)
        {
            base.Fail(ex);
            var pe = ex as PlanException;
            if (pe != null && pe.Field.Length > 0 && pe.Field != "gRows") Focus(pe.Field);
        }
        #endregion
    }
}

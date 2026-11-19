using System;
using System.Collections.Generic;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Material Request Planning: Production Orders of the selected project -> their component materials ->
    /// one Purchase Request. Filters apply only when Search is pressed.
    /// Shortcuts: Ctrl+F focuses Production Order search, F5 searches, Alt+L focuses the material grid, Ctrl+Enter creates the Purchase Request.
    /// Navigation: a Production Order already having a Purchase Request is shown in green; double-click it to open its linked requests.
    /// </summary>
    internal sealed class MaterialPlanForm : PlanForm
    {
        private GridTable gOrd, gRows;
        private List<ProductionOrder> orders = new List<ProductionOrder>();
        private HashSet<int> ordersWithLinks = new HashSet<int>();     // DocEntry of orders with an existing Purchase Request (shown in green, double-click to see them)
        private List<PlanRow> rows = new List<PlanRow>();
        /// <summary>The rows actually shown in gRows: all of rows, or just those of the chosen Manufacturer. Grid row
        /// indexes (Current, ReadRow, ...) always index into this, never into rows, so a filtered-out row can never
        /// be touched by a stale index.</summary>
        private List<PlanRow> view = new List<PlanRow>();
        /// <summary>Stock of the loaded items for the In Stock / Project Stock / Main Loc Stock columns, read on Load.</summary>
        private StockMap stock = new StockMap();
        private bool multiBranch;
        private int lastPrEntry, lastPrNum;
        private string lastPrProject = "";

        public static void Show()
        {
            RunShow<MaterialPlanForm>(() =>
            {
                MaterialPlanForm open = OpenForms<MaterialPlanForm>().FirstOrDefault();
                if (open != null) { open.Select(); return; }
                var f = new MaterialPlanForm();
                f.Create("EBHMRPPLN", "Material Request Planning", 980, 726, false, f.Build, autoSize: true);
                f.Focus("eProj");
            });
        }

        #region Build
        private void Build()
        {
            // 1. Filters
            Section("hFilt", "Search Filters", 8, 6, 300);
            Label("lProj", "Project", 8, 30, 55);
            Edit("eProj", 66, 30, 230);
            Cfl("eProj", "63", "PrjCode", false);
            Label("lFind", "Production Order", 306, 30, 120);
            Edit("eFind", 432, 30, 160);
            Cfl("eFind", "202", "DocNum", true);
            Button("bFind", "Search", 600, 28, 70, GridTable.Blue);
            Button("bClear", "Clear", 676, 28, 70, GridTable.Grey);
            // Right-aligned by Layout(): jumps to the other planning screen without going through the SAP menu
            Button("bSwitch", "Switch to Purchase Order Planning", 760, 28, 200, GridTable.Blue);
            Visible("bSwitch", false);

            // 2. Production Orders
            Section("hOrd", "Production Orders", 8, 60, 300);
            Label("lOrdCap", "[Double-click a row to see PR]", 170, 62, 260);
            gOrd = Grid("gOrd", 8, 86, 964, 130)
                .Column("Sr", "#", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Sel", "Select", width: 46, editable: true).CheckColumn("Sel")
                .Column("No", "Order No.", SAPbouiCOM.BoFieldsType.ft_Integer, 70).LinkColumn("No", "202")
                .Column("Ser", "Series", width: 90)
                .Column("Prj", "Project", width: 90)
                .Column("Prod", "Product No.", width: 120)
                .Column("Desc", "Product Description", width: 220)
                .Column("Stat", "Status", width: 70)
                .Column("Due", "Due Date", SAPbouiCOM.BoFieldsType.ft_Date, 80)
                .Column("Prog", "Progress", width: 120)
                .Sortable("Prod", "Desc", "Stat", "Due")
                .Bind();
            SortWith(gOrd, ShowOrders);
            Button("bOSelAll", "Select All", 8, 222, 90, GridTable.Grey);
            Button("bODeselAl", "Deselect All", 104, 222, 100, GridTable.Grey);
            Button("bLoad", "Load Selected Materials", 210, 222, 150, GridTable.Blue);
            // Own status label for this grid, separate from the Materials grid's "lFoot": the two grids' captions
            // used to share "lFoot", so re-searching Production Orders overwrote the Materials summary ("N of N
            // row(s) selected") with the orders' own, reading as if Materials had just reloaded when it had not.
            Label("lOrdFoot", "", 372, 225, 380);
            // Label("lOrdHint", "Double-click a row to see its Purchase Requests", 376, 225, 280);

            // 3. Purchase Request header
            //Section("hReq", "Request Details", 8, 250, 300);
            //Label("lSeries", "Series", 8, 272, 50);
            //Combo("cSeries", 64, 272, 110);
            //Label("lNextNo", "Doc.No", 182, 272, 55);
            //Field("fNextNo", 240, 272, 60, true);
            //Label("lBranch", "Branch", 308, 272, 50);
            //Combo("cBranch", 362, 272, 130);
            //Label("lReq", "Required Date", 500, 272, 90);
            //Edit("eReq", 594, 272, 90, SAPbouiCOM.BoDataType.dt_DATE);
            //Label("lValid", "Valid Until Date", 688, 272, 105);
            //Edit("eValid", 798, 272, 90, SAPbouiCOM.BoDataType.dt_DATE);

            //Label("lVend", "Vendor", 8, 296, 50);
            //Edit("eVend", 62, 296, 125);
            //Cfl("eVend", "2", "CardCode", false);
            //Label("lWhs", "Warehouse", 196, 296, 78);
            //Edit("eWhs", 278, 296, 125);
            //Cfl("eWhs", "64", "WhsCode", false);
            //Button("bApply", "Apply to Selected Rows", 412, 294, 195, GridTable.Blue);
            //Button("bSelAll", "Select All", 616, 294, 80, GridTable.Grey);
            //Button("bDeselAll", "Deselect All", 706, 294, 100, GridTable.Grey);
            Section("hReq", "Request Details", 8, 250, 300);

            // Row 1: Document and date information
            Label("lSeries", "Series", 8, 272, 48);
            Combo("cSeries", 60, 270, 115);

            Label("lNextNo", "Doc. No.", 184, 272, 55);
            Field("fNextNo", 242, 270, 65, true);

            Label("lBranch", "Branch", 316, 272, 48);
            Combo("cBranch", 368, 270, 125);

            Label("lReq", "Required Date", 502, 272, 88);
            Edit("eReq", 594, 270, 90, SAPbouiCOM.BoDataType.dt_DATE);

            Label("lValid", "Valid Until", 692, 272, 75);
            Edit("eValid", 771, 270, 100, SAPbouiCOM.BoDataType.dt_DATE);


            // Row 2: Manufacturer filter first, then Vendor and Warehouse
            // Manufacturer: filters the Materials grid to just that Manufacturer's rows. Applies automatically when
            // the combo itself changes; the button is a reliable, explicit way to re-apply it too.
            Label("lMan", "Manufacture", 8, 299, 80);
            Combo("cMan", 90, 297, 230);
            Button("bManFilter", "Apply Manufacturer Filter", 326, 296, 160, GridTable.Blue);

            Label("lVend", "Vendor", 496, 299, 48);
            Edit("eVend", 548, 297, 125);
            Cfl("eVend", "2", "CardCode", false);

            Label("lWhs", "Warehouse", 682, 299, 72);
            Edit("eWhs", 758, 297, 125);
            Cfl("eWhs", "64", "WhsCode", false);

            // Row 3: Selection actions
            Button("bSelAll", "Select All", 8, 320, 85, GridTable.Grey);
            Button("bDeselAll", "Deselect All", 101, 320, 100, GridTable.Grey);
            Button("bApply", "Apply to Selected Rows", 209, 320, 185, GridTable.Blue);

            // 4. Material details
            Section("hMat", "Materials", 8, 350, 300);
            gRows = Grid("gRows", 8, 374, 964, 280)
                .Column("Sr", "#", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Sel", "Select", width: 46, editable: true).CheckColumn("Sel")
                .Column("Catalog", "Catalog No.", width: 90)
                .Column("Item", "Item No.", width: 100).LinkColumn("Item", "4")
                .Column("Desc", "Description", width: 180)
                .Column("Job", "Job Code", width: 90)
                .Column("Man", "Manufacturer", width: 120)
                .Column("Qty", "Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 70, true)
                .Column("Uom", "UOM", width: 50)
                .Column("After", "Price After Discount", SAPbouiCOM.BoFieldsType.ft_Price, 100)
                .Column("Total", "Total", SAPbouiCOM.BoFieldsType.ft_Sum, 90)
                .Column("Whs", "Warehouse", width: 80, editable: true)
                .Column("Stk", "In Stock", SAPbouiCOM.BoFieldsType.ft_Quantity, 70)
                .Column("PStk", "Project Stock", SAPbouiCOM.BoFieldsType.ft_Quantity, 80)
                .Column("MStk", "Main Loc Stock", SAPbouiCOM.BoFieldsType.ft_Quantity, 85)
                .Column("Vend", "Vendor", width: 90, editable: true)
                .Column("Req", "Required Date", SAPbouiCOM.BoFieldsType.ft_Date, 85, true)
                .Column("Valid", "Valid Until", SAPbouiCOM.BoFieldsType.ft_Date, 85, true)
                .Column("Base", "Base Document", width: 90).LinkColumn("Base", "202")
                .Sortable("Catalog", "Item", "Desc", "Man", "Vend", "Whs", "Req", "Valid")
                .Bind();
            SortWith(gRows, ShowRows);
            CflColumn(gRows, "Whs", "64", "WhsCode");
            CflColumn(gRows, "Vend", "2", "CardCode");

            Button("bCreate", "Purchase Request", 8, 574, 150, GridTable.Green);
            Button("2", "Cancel", 162, 574, 70, GridTable.Grey);
            Button("bLinked", "Linked Doc.", 236, 574, 120, GridTable.SkyBlue);
            Button("bLastPr", "", 366, 574, 190, GridTable.Green);
            Visible("bLastPr", false);
            Button("bCreatePo", "Create PO", 562, 574, 150, GridTable.Blue);
            Visible("bCreatePo", false);
            Label("lFoot", "", 718, 577, 180);
            Label("lTotal", "Total", 740, 577, 90);
            Field("eTotal", 834, 574, 130, true);

            // A blank first entry so the combo itself can clear the filter (it is a strict pick list, not free text)
            var manChoices = new List<KeyValuePair<string, string>> { new KeyValuePair<string, string>("", "(All Manufacturers)") };
            manChoices.AddRange(Store.Manufacturers());
            Choices("cMan", manChoices);
            Choices("cSeries", Store.Series());
            int series = Store.DefaultSeries();
            if (series > 0) Set("cSeries", series.ToString(Inv));

            multiBranch = Store.MultiBranch;
            Visible("lBranch", multiBranch); Visible("cBranch", multiBranch);
            if (multiBranch)
            {
                Choices("cBranch", Store.Branches());
                int branch = Store.DefaultBranch();
                if (branch > 0) Set("cBranch", branch.ToString(Inv));
            }

            SetDate("eReq", DateTime.Today);
            SetDate("eValid", DateTime.Today.AddMonths(1));
            Set("eWhs", "MAIN LOC");
            Set("eTotal", 0m);
            RefreshNextNo();
            ShowRows();
            // U_EBH_PRQMAP.U_PrStatus is only ever refreshed when this add-on happens to touch a given Production
            // Order (see PlanStore.Sync); a Purchase Request cancelled through its own standard document since the
            // last time could otherwise sit stale until something unrelated triggers a refresh. Brought current for
            // every row at once the moment the form opens, so what is on screen never starts out stale. Best-effort:
            // a failure here must not stop the form from opening.
            try { Store.SyncAllMapStatus(); } catch (Exception ex) { Warn("Status not refreshed : " + ex.Message); }
        }

        protected override void Layout()
        {
            // Small floors only to keep every width/height positive - the real limit is whatever Create() already fit
            // to the SAP Desktop, which on a small monitor can be well under the form's original design size.
            int w = Math.Max(600, Form.ClientWidth), h = Math.Max(400, Form.ClientHeight);
            Place("bSwitch", Math.Max(756, w - 16 - 200), 28, 0, 0);     // never left of bClear (ends at 746), however narrow w gets
            Place("hFiltL", 8, 24, w - 16, 0);
            Place("hOrdL", 8, 78, w - 16, 0);
            Place("gOrd", 8, 86, w - 16, 130);
            gOrd.FillWidth("Desc", w - 16);
            Place("hReqL", 8, 268, w - 16, 0);
            Place("hMatL", 8, 368, w - 16, 0);
            Place("gRows", 8, 374, w - 16, Math.Max(80, h - 374 - 40));
            gRows.FillWidth("Desc", w - 16);
            Place("bCreate", 8, h - 30, 0, 0);
            Place("2", 162, h - 30, 0, 0);
            Place("bLinked", 236, h - 30, 0, 0);
            Place("bLastPr", 366, h - 30, 0, 0);
            Place("bCreatePo", 562, h - 30, 0, 0);
            Place("lFoot", 718, h - 27, 0, 0);
            Place("lTotal", w - 240, h - 27, 0, 0);
            Place("eTotal", w - 146, h - 30, 0, 0);
        }
        #endregion

        #region Search and load
        private static List<string> Codes(string text)
        {
            return text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(c => c.Trim()).Where(c => c.Length > 0).ToList();
        }

        /// <summary>Production Order numbers currently typed in eFind (only when every token parses as a number).</summary>
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

        /// <summary>Narrows the Production Order Choose From List (eFind) to the project currently typed, if any.</summary>
        private void FilterFindCfl()
        {
            SAPbouiCOM.Conditions cons = NewConditions();
            string proj = Text("eProj");
            if (proj.Length > 0)
            {
                SAPbouiCOM.Condition con = cons.Add();
                con.Alias = "Project";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = proj;
            }
            GetCfl("eFind").SetConditions(cons);
        }

        /// <summary>
        /// Fills Project from the given Production Orders when they share exactly one and Project is still blank;
        /// refuses when they span more than one Project (the request can only target a single Project / Vendor / Warehouse header).
        /// </summary>
        private void CheckSameProject(IEnumerable<int> docNums)
        {
            List<string> projects = Store.ProjectsOfOrders(docNums).Distinct().ToList();
            if (projects.Count > 1) throw new PlanException(400, "Please select Production Orders of the same Project.");
            if (projects.Count == 1 && Text("eProj").Length == 0) Set("eProj", projects[0]);
        }

        /// <summary>Narrows the Project Choose From List (eProj) to the project(s) of the Production Order(s) currently typed, if any.</summary>
        private void FilterProjectCfl()
        {
            SAPbouiCOM.Conditions cons = NewConditions();
            List<int> nums = FindDocNums();
            if (nums.Count > 0)
            {
                List<string> projects = Store.ProjectsOfOrders(nums);
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

        private void Search()
        {
            // Read what is typed, then write it straight back into the bound data source: a button press moves the
            // focus off the edit field, and without this the field can revert to its last committed (empty) value.
            string proj = Text("eProj"), find = Text("eFind");
            Set("eProj", proj);
            Set("eFind", find);

            orders = Store.ListOrders(Codes(proj), find);
            ordersWithLinks = new HashSet<int>(Store.OrdersWithLinkedRequests(orders.Select(o => o.DocEntry)));
            ShowOrders();
            if (orders.Count == 0) Warn("No Production Order found for this search.");
        }

        private static string OrderStatus(ProductionOrder o)
        {
            return o.RequestsClosed ? "Completed" : o.Status == "R" ? "Released" : "Planned";
        }

        private void ShowOrders()
        {
            orders = gOrd.Sort(orders, (c, o) => c == "Prod" ? o.ItemCode : c == "Desc" ? o.ItemName : c == "Stat" ? OrderStatus(o) : c == "Due" ? o.DueDate : "");
            Freeze(() =>
            {
                gOrd.Fill(orders.Select((o, i) => new object[]
                {
                    i + 1, o.Selected ? "Y" : "N", o.DocNum, o.SeriesName, o.Project, o.ItemCode, o.ItemName,
                    OrderStatus(o), ToDate(o.DueDate),
                    o.TotalComponents == 0 ? "" : o.CompletedComponents + " of " + o.TotalComponents + " completed"
                }).ToList());
                // Completed (every Purchase Request line Closed) wins over the progress colours
                for (int i = 0; i < orders.Count; i++)
                    gOrd.Color(i, orders[i].RequestsClosed ? GridTable.Green
                        : orders[i].ProgressState == "Full" ? GridTable.MintGreen
                        : orders[i].ProgressState == "Partial" ? GridTable.PastelYellow : GridTable.CoolGray);
                int linked = orders.Count(o => ordersWithLinks.Contains(o.DocEntry));
                Caption("lOrdFoot", orders.Count == 0 ? "No Production Order found for this search."
                    : orders.Count(o => o.Selected) + " of " + orders.Count + " Production Order(s) selected" +
                    (linked > 0 ? " (" + linked + " already have a Purchase Request - double-click to see them)" : ""));
            });
        }

        private void SetAllOrdersSelected(bool on)
        {
            if (orders.Count == 0) return;
            foreach (ProductionOrder o in orders) o.Selected = on;
            ShowOrders();
        }

        private void Load()
        {
            var picked = orders.Where(o => o.Selected).Select(o => o.DocEntry).Distinct().ToList();
            if (picked.Count == 0) throw new PlanException(400, "Tick one or more Production Orders first.");
            int skipped;
            rows = PlanService.LoadRows(picked, out skipped);
            stock = Store.GetStock(rows.Select(r => r.ItemCode));
            HideLastPr();
            ShowRows();
            if (rows.Count == 0) Alert("Nothing to request: every component of the selected Production Orders is already requested.");
            else Done(rows.Count + " material row(s) loaded" + (skipped > 0 ? ", " + skipped + " component(s) already fully requested were left out." : "."));
        }

        private void Clear()
        {
            Set("eProj", "");
            Set("eFind", "");
            orders = new List<ProductionOrder>();
            ordersWithLinks = new HashSet<int>();
            rows = new List<PlanRow>();
            HideLastPr();
            ShowOrders();
            ShowRows();
            Ok("Filters and rows cleared.");
        }
        #endregion

        #region Rows
        private static string Base(PlanRow r)
        {
            string b = r.PdoNum + " / " + r.PdoLine;
            return r.MergedFrom != null && r.MergedFrom.Count > 1 ? b + " (+" + (r.MergedFrom.Count - 1) + ")" : b;
        }

        /// <summary>
        /// Rows gRows should currently show: all of them (or just the chosen Manufacturer's), with rows that share the
        /// same Catalog No. + Item No. + Vendor + Description combined into one summary row - Quantity, Balance and
        /// AlreadyRequested summed; Warehouse / Required Date / Valid Until taken from the first row that actually has
        /// one (these are per-component values, not additive); Price / Discount / Price After Discount from the member
        /// with the lowest Price After Discount (WOR1.U_EBH_PAFDI), copied onto every member so the request line is
        /// raised at that price, and Total = that price x the summed Quantity. A single
        /// matching row is returned unchanged - only an actual duplicate is ever wrapped. Each summary row becomes ONE
        /// Purchase Request line at Create time, with every member component recorded in @EBH_PRQMAP (see
        /// PlanStore.CreatePurchaseRequest); the underlying "rows" list is never replaced here, so a merged row's
        /// edits are propagated back onto its members - see ReadRow / Apply.
        /// <paramref name="ignoreFilter"/>: group every row whatever the Manufacturer filter shows (used by Create).
        /// </summary>
        private List<PlanRow> Visible(bool ignoreFilter = false)
        {
            string man = ignoreFilter ? "" : Get("cMan");
            var filtered = man.Length == 0 ? rows : rows.Where(r => r.Manufacturer == man).ToList();
            return filtered
                .GroupBy(r => r.Catalog + "|" + r.ItemCode + "|" + r.Vendor + "|" + r.Description)
                .OrderBy(g => g.First().ItemCode).ThenBy(g => g.First().Vendor)
                .Select(g =>
                {
                    var members = g.ToList();
                    if (members.Count == 1) return members[0];
                    PlanRow first = members[0];
                    PlanRow cheapest = PlanService.Cheapest(members, m => m.PriceAfterDiscount);
                    foreach (PlanRow m in members) { m.Price = cheapest.Price; m.Discount = cheapest.Discount; m.PriceAfterDiscount = cheapest.PriceAfterDiscount; }
                    return new PlanRow
                    {
                        Selected = members.Any(m => m.Selected),
                        PdoEntry = first.PdoEntry, PdoNum = first.PdoNum, PdoLine = first.PdoLine,
                        ItemCode = first.ItemCode, Description = first.Description, Uom = first.Uom,
                        Quantity = members.Sum(m => m.Quantity), Balance = members.Sum(m => m.Balance),
                        AlreadyRequested = members.Sum(m => m.AlreadyRequested),
                        HasLinkedRequest = members.Any(m => m.HasLinkedRequest),
                        Price = cheapest.Price, Discount = cheapest.Discount, PriceAfterDiscount = cheapest.PriceAfterDiscount,
                        Warehouse = members.Select(m => m.Warehouse).FirstOrDefault(w => w.Length > 0) ?? "",
                        Vendor = first.Vendor,
                        RequiredDate = members.Select(m => m.RequiredDate).FirstOrDefault(d => d.Length > 0) ?? "",
                        ValidUntil = members.Select(m => m.ValidUntil).FirstOrDefault(d => d.Length > 0) ?? "",
                        Project = first.Project, Manufacturer = first.Manufacturer, ManufacturerName = first.ManufacturerName,
                        Catalog = first.Catalog, PdoItemCode = first.PdoItemCode, TracksJobCode = first.TracksJobCode,
                        MergedFrom = members
                    };
                }).ToList();
        }

        private void ShowRows()
        {
            view = gRows.Sort(Visible(), (c, r) => c == "Catalog" ? r.Catalog : c == "Item" ? r.ItemCode
                : c == "Desc" ? r.Description : c == "Man" ? r.ManufacturerName : c == "Vend" ? r.Vendor
                : c == "Whs" ? r.Warehouse : c == "Req" ? r.RequiredDate : c == "Valid" ? r.ValidUntil : "");     // dates are yyyy-MM-dd: text order = date order
            Freeze(() =>
            {
                gRows.Fill(view.Select((r, i) => new object[]
                {
                    i + 1, r.Selected ? "Y" : "N", r.Catalog, r.ItemCode, r.Description, r.TracksJobCode ? r.PdoItemCode : "", r.ManufacturerName, r.Quantity, r.Uom, r.PriceAfterDiscount, r.Total,
                    r.Warehouse, stock.InStock(r.ItemCode, r.Warehouse), stock.ProjectStock(r.ItemCode, r.Warehouse, r.Project), stock.MainLocStock(r.ItemCode),
                    r.Vendor, ToDate(r.RequiredDate), ToDate(r.ValidUntil), Base(r)
                }).ToList());
                for (int i = 0; i < view.Count; i++)
                    gRows.Color(i, !view[i].Selected ? GridTable.Grey : view[i].Quantity > view[i].Balance ? GridTable.Red
                        : view[i].HasLinkedRequest ? GridTable.SkyBlue : GridTable.Lavender);
                // The Total stays the grand total of every selected row (any Manufacturer): what Create would actually
                // request, not just what happens to be visible under the current filter.
                decimal total = rows.Where(r => r.Selected).Sum(r => r.Total);
                Set("eTotal", total);
                int linked = view.Count(r => r.HasLinkedRequest);
                string man = Get("cMan");
                int manCount = man.Length == 0 ? rows.Count : rows.Count(r => r.Manufacturer == man);
                string filtered = man.Length > 0 && manCount != rows.Count
                    ? " (filtered: " + manCount + " of " + rows.Count + " row(s) match Manufacturer " + man + ")" : "";
                // Same-key rows are grouped into one row here for easier editing and are exactly what Create writes:
                // one Purchase Request line per group shown, quantity = the group's own total (see CreateRequest).
                string grouped = view.Count != rows.Count
                    ? " (grouped from " + rows.Count + " source component row(s) by Catalog No. + Item No. + Vendor + Description)" : "";
                if (rows.Count > 0) Caption("lFoot", view.Count(r => r.Selected) + " of " + view.Count + " row(s) selected" + filtered + grouped +
                    (linked > 0 ? " (" + linked + " shown in green: source Production Order already has a Purchase Request)" : ""));
            });
        }

        /// <summary>Takes what the user typed in one visible grid row into the model. A merged row (see Visible) has
        /// no Purchase Request line of its own, so every edit is also propagated onto its member rows - Warehouse / Vendor / Required Date / Valid Until / Selected applied as typed to each member,
        /// Quantity redistributed across them in their original proportion (each member still capped by its own
        /// source component's Balance at Create time).</summary>
        private void ReadRow(int tr)
        {
            PlanRow r = view[tr];
            r.Selected = gRows.TypedText(tr, "Sel") == "Y";
            r.Quantity = gRows.Typed(tr, "Qty");
            r.Warehouse = gRows.TypedText(tr, "Whs");
            r.Vendor = gRows.TypedText(tr, "Vend");
            r.RequiredDate = gRows.TypedDate(tr, "Req");
            r.ValidUntil = gRows.TypedDate(tr, "Valid");
            if (r.MergedFrom == null) return;
            decimal originalTotal = r.MergedFrom.Sum(m => m.Quantity);
            foreach (PlanRow m in r.MergedFrom)
            {
                m.Selected = r.Selected;
                m.Quantity = originalTotal > 0 ? Math.Round(r.Quantity * (m.Quantity / originalTotal), 6) : r.Quantity / r.MergedFrom.Count;
                m.Warehouse = r.Warehouse; m.Vendor = r.Vendor;
                m.RequiredDate = r.RequiredDate; m.ValidUntil = r.ValidUntil;
            }
        }

        private PlanHeader ReadHeader()
        {
            int series, branch = 0;
            int.TryParse(Get("cSeries"), out series);
            if (multiBranch) int.TryParse(Get("cBranch"), out branch);
            return new PlanHeader { Series = series, Branch = branch, RequiredDate = Date("eReq"), ValidUntil = Date("eValid"), Vendor = Text("eVend"), Warehouse = Text("eWhs") };
        }

        private void Apply()
        {
            if (!view.Any(r => r.Selected)) throw new PlanException(400, "Select at least one material row first.");
            if (!Ask("Vendor, Warehouse, Required Date and Valid Until Date of the selected rows will be replaced with the header values. Continue?", "Replace", "Cancel")) return;
            var selected = view.Where(r => r.Selected).ToList();
            PlanHeader h = ReadHeader();
            PlanService.ApplyHeader(h, selected);
            // "view" rows are a display summary rebuilt on every ShowRows() (see Visible()); a merged row's own
            // edit must also land on its real member rows, or it is lost the moment the grid next refreshes.
            foreach (PlanRow r in selected) if (r.MergedFrom != null) PlanService.ApplyHeader(h, r.MergedFrom);
            ShowRows();
            Done("Header values applied to " + selected.Count + " row(s).");
        }

        /// <summary>Select All / Deselect All act only on the rows currently visible under the Manufacturer filter; a
        /// row hidden by the filter keeps whatever Selected state it already had.</summary>
        private void SetAllSelected(bool on)
        {
            if (view.Count == 0) return;
            foreach (PlanRow r in view)
            {
                r.Selected = on;
                if (r.MergedFrom != null) foreach (PlanRow m in r.MergedFrom) m.Selected = on;
            }
            ShowRows();
        }

        /// <summary>
        /// Fires automatically when the Manufacturer combo changes: re-filters the grid to just that Manufacturer's
        /// rows (the blank entry shows everything again). Does not touch any row's Selected state - filtering is a
        /// display concern, ticks stay exactly as the user left them even for rows the filter now hides.
        /// </summary>
        private void FilterByManufacturer()
        {
            ShowRows();
            if (rows.Count == 0) return;
            string man = Get("cMan");
            if (man.Length == 0) Ok("Showing all " + rows.Count + " row(s).");
            else if (view.Count == 0) Warn("No row matches Manufacturer " + man + ".");
            else Ok(view.Count + " of " + rows.Count + " row(s) match Manufacturer " + man + ".");
        }

        private void CreateRequest()
        {
            for (int i = 0; i < view.Count; i++) ReadRow(i);
            // One Purchase Request line per selected group, quantity = the group's own combined total. Every selected
            // row is included, also one the Manufacturer filter currently hides (the footer Total already counts it).
            // Every source Production Order component a group stands for is recorded separately in @EBH_PRQMAP (see
            // PlanStore.CreatePurchaseRequest), so a component merged into someone else's line is still recognised as
            // "already requested" the next time its Production Order is loaded.
            var selected = Visible(true).Where(r => r.Selected).ToList();
            if (selected.Count == 0) throw new PlanException(400, "Select at least one material row.", "gRows");
            int sourceComponents = selected.Sum(r => r.MergedFrom != null && r.MergedFrom.Count > 0 ? r.MergedFrom.Count : 1);
            decimal totalQty = selected.Sum(r => r.Quantity);
            if (!Ask(
                selected.Count + " consolidated material group(s) are selected, from " + sourceComponents + " source Production Order component row(s). " +
                "One Purchase Request with " + selected.Count + " line(s) will be created (total quantity " + totalQty.ToString("0.######") + "). " +
                "Source Production Order traceability will be retained. Continue?", "Create", "Cancel")) return;
            CreatedRequest made = PlanService.Create(ReadHeader(), selected);
            // One request per load: the rows are cleared so Create cannot be pressed twice for the same materials
            rows = new List<PlanRow>();
            ShowRows();
            lastPrEntry = made.DocEntry; lastPrNum = made.DocNum; lastPrProject = Text("eProj");
            ButtonCaption("bLastPr", "Open Purchase Request " + made.DocNum);
            Visible("bLastPr", true);
            // Re-searches the Production Orders grid with whatever filters are already set, so its progress count
            // and colour (driven by @EBH_PRQMAP - see PlanStore.ComponentProgress) reflect the request just created
            // immediately, instead of looking stale until the user clicks Search themselves.
            Search();
            if (made.Warning.Length > 0) Warn(made.Warning);
            Done("Purchase Request " + made.DocNum + " was created with " + selected.Count + " line(s), from " + sourceComponents + " source component(s).");
        }

        /// <summary>Opens Purchase Order Planning already searched and loaded for the just-created Purchase Request, so punching the Purchase Order from it is a single click away.</summary>
        private void CreatePo()
        {
            if (lastPrNum <= 0) return;
            PurchaseOrderForm.ShowForRequest(lastPrNum, lastPrProject);
        }

        private void HideLastPr()
        {
            lastPrEntry = 0;
            Visible("bLastPr", false);
            Visible("bCreatePo", false);
        }

        private int CurrentPdo()
        {
            if (gRows.Current < 0 || gRows.Current >= view.Count) throw new PlanException(400, "Click a material row first.");
            return view[gRows.Current].PdoEntry;
        }

        private void RefreshNextNo()
        {
            int series;
            int.TryParse(Get("cSeries"), out series);
            int next = Store.NextNumber(series);
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
            else if (Key(e, 13, SAPbouiCOM.BoModifiersEnum.mt_CTRL)) CreateRequest();
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && e.BeforeAction && e.ItemUID == "eProj") FilterProjectCfl();
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && e.BeforeAction && e.ItemUID == "eFind") FilterFindCfl();
            else if ((chosen = ChosenCodes(e, "eProj", "PrjCode")).Count > 0) Set("eProj", chosen[0]);
            else if ((chosen = ChosenCodes(e, "eFind", "DocNum")).Count > 0)
            {
                Set("eFind", string.Join(",", chosen.ToArray()));
                var nums = chosen.Select(c => { int n; return int.TryParse(c, out n) ? n : 0; }).Where(n => n > 0).ToList();
                CheckSameProject(nums);
                Search();
            }
            else if ((chosen = ChosenCodes(e, "eVend", "CardCode")).Count > 0) Set("eVend", chosen[0]);
            else if ((chosen = ChosenCodes(e, "eWhs", "WhsCode")).Count > 0) Set("eWhs", chosen[0]);
            else if ((chosen = ChosenGridCodes(e, gRows, "Whs", "WhsCode", out gridRow)).Count > 0 && gridRow >= 0 && gridRow < view.Count)
            {
                view[gridRow].Warehouse = chosen[0];
                if (view[gridRow].MergedFrom != null) foreach (PlanRow m in view[gridRow].MergedFrom) m.Warehouse = chosen[0];
                ShowRows();
                KeepOkMode("Cancel");
            }
            else if ((chosen = ChosenGridCodes(e, gRows, "Vend", "CardCode", out gridRow)).Count > 0 && gridRow >= 0 && gridRow < view.Count)
            {
                view[gridRow].Vendor = chosen[0];
                if (view[gridRow].MergedFrom != null) foreach (PlanRow m in view[gridRow].MergedFrom) m.Vendor = chosen[0];
                ShowRows();
                KeepOkMode("Cancel");
            }
            else if (Pressed(e, "bFind") || Enter(e, "eFind", "eProj")) Search();
            else if (Pressed(e, "bClear")) Clear();
            else if (Pressed(e, "bSwitch")) PurchaseOrderForm.Show();
            else if (e.ItemUID == "cSeries" && !e.BeforeAction &&
                (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)) RefreshNextNo();
            else if (e.ItemUID == "cMan" && !e.BeforeAction &&
                (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)) FilterByManufacturer();
            else if (Pressed(e, "bManFilter")) FilterByManufacturer();
            else if (Pressed(e, "bLoad")) Load();
            else if (Pressed(e, "bOSelAll")) SetAllOrdersSelected(true);
            else if (Pressed(e, "bODeselAl")) SetAllOrdersSelected(false);
            else if (Pressed(e, "bApply")) Apply();
            else if (Pressed(e, "bSelAll")) SetAllSelected(true);
            else if (Pressed(e, "bDeselAll")) SetAllSelected(false);
            else if (Pressed(e, "bCreate")) CreateRequest();
            else if (Pressed(e, "bLastPr")) { if (lastPrEntry > 0) Store.OpenPurchaseRequest(lastPrEntry); }
            else if (Pressed(e, "bCreatePo")) CreatePo();
            else if (Pressed(e, "bLinked")) LinkedRequestForm.Show(CurrentPdo());
            else if (Clicked(e, gRows)) gRows.Click(e.Row);
            else if (DoubleClicked(e, gOrd))
            {
                // Quick navigation: jump straight to this order's Purchase Requests without loading materials first
                int tr = gOrd.Row(e.Row);
                if (tr >= 0 && tr < orders.Count) LinkedRequestForm.Show(orders[tr].DocEntry);
            }
            else if (e.ItemUID == "gOrd" && !e.BeforeAction && e.Row >= 0 &&
                ((e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ItemChanged) || (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && e.ColUID == "Sel")))
            {
                int tr = gOrd.Row(e.Row);
                if (tr < 0 || tr >= orders.Count) return;
                orders[tr].Selected = gOrd.TypedText(tr, "Sel") == "Y";
                if (orders[tr].Selected)
                {
                    var selected = orders.Where(o => o.Selected).Select(o => o.DocNum).ToList();
                    List<string> projects = Store.ProjectsOfOrders(selected).Distinct().ToList();
                    if (projects.Count > 1)
                    {
                        // Undo this pick: a single request can only target one Project / Vendor / Warehouse header
                        orders[tr].Selected = false;
                        ShowOrders();
                        throw new PlanException(400, "Please select Production Orders of the same Project.");
                    }
                    if (projects.Count == 1 && Text("eProj").Length == 0) Set("eProj", projects[0]);
                }
                ShowOrders();
                KeepOkMode("Cancel");
            }
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && e.BeforeAction && e.ItemUID == "gOrd" && e.ColUID == "No")
            {
                // SAP's own link would look up the number shown; the Production Order is opened by its DocEntry
                bubble = false;
                int tr = gOrd.Row(e.Row);
                if (tr >= 0 && tr < orders.Count) Store.OpenProductionOrder(orders[tr].DocEntry);
            }
            else if (e.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && e.BeforeAction && e.ItemUID == "gRows" && e.ColUID == "Base")
            {
                // SAP's own link would look up the number shown; the Production Order is opened by its DocEntry
                bubble = false;
                int tr = gRows.Row(e.Row);
                if (tr >= 0 && tr < view.Count) Store.OpenProductionOrder(view[tr].PdoEntry);
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
                ShowRows();
                KeepOkMode("Cancel");
            }
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

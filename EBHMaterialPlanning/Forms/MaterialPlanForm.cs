using System;
using System.Collections.Generic;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Material Request Planning: Production Orders of the selected project -> their component materials ->
    /// one Purchase Request. Filters apply only when Search is pressed.
    /// </summary>
    internal sealed class MaterialPlanForm : PlanForm
    {
        private GridTable gOrd, gRows;
        private List<ProductionOrder> orders = new List<ProductionOrder>();
        private List<PlanRow> rows = new List<PlanRow>();
        private bool multiBranch;
        private int lastPrEntry, lastPrNum;

        public static void Show()
        {
            Run(() =>
            {
                MaterialPlanForm open = OpenForms<MaterialPlanForm>().FirstOrDefault();
                if (open != null) { open.Select(); return; }
                var f = new MaterialPlanForm();
                f.Create("EBHMRPPLN", "Material Request Planning", 980, 600, false, f.Build);
                f.Focus("eProj");
            });
        }

        #region Build
        private void Build()
        {
            // 1. Filters
            Label("lProj", "Project", 8, 8, 50);
            Edit("eProj", 62, 8, 230);
            Cfl("eProj", "63", "PrjCode", false);
            Label("lFind", "Production Order", 306, 8, 110);
            Edit("eFind", 420, 8, 160);
            Cfl("eFind", "202", "DocNum", true);
            Button("bFind", "Search", 588, 5, 70);
            Button("bClear", "Clear", 664, 5, 70);

            // 2. Production Orders
            gOrd = Grid("gOrd", 8, 32, 964, 120)
                .Column("Sel", "Select", width: 46, editable: true).CheckColumn("Sel")
                .Column("No", "Order No.", SAPbouiCOM.BoFieldsType.ft_Integer, 70).LinkColumn("No", "202")
                .Column("Ser", "Series", width: 90)
                .Column("Prj", "Project", width: 90)
                .Column("Prod", "Product No.", width: 120)
                .Column("Desc", "Product Description", width: 220)
                .Column("Stat", "Status", width: 70)
                .Column("Due", "Due Date", SAPbouiCOM.BoFieldsType.ft_Date, 80)
                .Bind();
            Button("bLoad", "Load Selected Materials", 8, 158, 150);
            Button("bOSelAll", "Select All", 166, 158, 90);
            Button("bODeselAl", "Deselect All", 262, 158, 100);

            // 3. Purchase Request header
            Label("lSeries", "Target Document Series", 8, 182, 140);
            Combo("cSeries", 152, 182, 110);
            Label("lNextNo", "Doc.No", 266, 182, 55);
            Field("fNextNo", 324, 182, 60, true);
            Label("lBranch", "Branch", 392, 182, 45);
            Combo("cBranch", 440, 182, 100);
            Label("lReq", "Required Date", 548, 182, 85);
            Edit("eReq", 636, 182, 90, SAPbouiCOM.BoDataType.dt_DATE);
            Label("lValid", "Valid Until Date", 734, 182, 100);
            Edit("eValid", 838, 182, 90, SAPbouiCOM.BoDataType.dt_DATE);

            Label("lVend", "Vendor", 8, 204, 45);
            Edit("eVend", 56, 204, 120);
            Cfl("eVend", "2", "CardCode", false);
            Label("lWhs", "Warehouse", 184, 204, 62);
            Edit("eWhs", 250, 204, 120);
            Cfl("eWhs", "64", "WhsCode", false);
            Button("bApply", "Apply Header to Selected Rows", 378, 204, 170);
            Button("bApplyWhs", "Apply Warehouse to Selected Rows", 552, 204, 180);
            Button("bSelAll", "Select All", 736, 204, 80);
            Button("bDeselAll", "Deselect All", 820, 204, 100);

            // 4. Material details
            gRows = Grid("gRows", 8, 228, 964, 280)
                .Column("Sel", "Select", width: 46, editable: true).CheckColumn("Sel")
                .Column("Item", "Item No.", width: 100)
                .Column("Desc", "Description", width: 180)
                .Column("Qty", "Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 70, true)
                .Column("Uom", "UOM", width: 50)
                .Column("Price", "Price", SAPbouiCOM.BoFieldsType.ft_Price, 70, true)
                .Column("Disc", "Discount %", SAPbouiCOM.BoFieldsType.ft_Percent, 70, true)
                .Column("After", "Price After Discount", SAPbouiCOM.BoFieldsType.ft_Price, 100)
                .Column("Total", "Total", SAPbouiCOM.BoFieldsType.ft_Sum, 90)
                .Column("Whs", "Warehouse", width: 80, editable: true)
                .Column("Vend", "Vendor", width: 90, editable: true)
                .Column("Req", "Required Date", SAPbouiCOM.BoFieldsType.ft_Date, 85, true)
                .Column("Valid", "Valid Until", SAPbouiCOM.BoFieldsType.ft_Date, 85, true)
                .Column("Base", "Base Document", width: 90).LinkColumn("Base", "202")
                .Bind();
            CflColumn(gRows, "Whs", "64", "WhsCode");
            CflColumn(gRows, "Vend", "2", "CardCode");

            Button("bCreate", "Create Purchase Request", 8, 540, 150);
            Button("2", "Cancel", 162, 540, 70);
            Button("bLinked", "Linked Documents", 236, 540, 120);
            Button("bLastPr", "", 366, 540, 190);
            Visible("bLastPr", false);
            Label("lFoot", "", 566, 543, 180);
            Label("lTotal", "Estimated Total", 740, 543, 90);
            Field("eTotal", 834, 540, 130, true);

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
            SetDate("eValid", DateTime.Today.AddDays(30));
            Set("eTotal", 0m);
            RefreshNextNo();
            ShowRows();
        }

        protected override void Layout()
        {
            int w = Math.Max(900, Form.ClientWidth), h = Math.Max(420, Form.ClientHeight);
            Place("gOrd", 8, 32, w - 16, 120);
            Place("gRows", 8, 228, w - 16, h - 228 - 40);
            Place("bCreate", 8, h - 30, 0, 0);
            Place("2", 162, h - 30, 0, 0);
            Place("bLinked", 236, h - 30, 0, 0);
            Place("bLastPr", 366, h - 30, 0, 0);
            Place("lFoot", 566, h - 27, 0, 0);
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
            ShowOrders();
            if (orders.Count == 0) Warn("No Production Order found for this search.");
        }

        private void ShowOrders()
        {
            Freeze(() =>
            {
                gOrd.Fill(orders.Select(o => new object[]
                {
                    o.Selected ? "Y" : "N", o.DocNum, o.SeriesName, o.Project, o.ItemCode, o.ItemName, o.Status == "R" ? "Released" : "Planned", ToDate(o.DueDate)
                }).ToList());
                Caption("lFoot", orders.Count == 0 ? "No Production Order found for this search."
                    : orders.Count(o => o.Selected) + " of " + orders.Count + " Production Order(s) selected");
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
            HideLastPr();
            ShowRows();
            if (rows.Count == 0) Warn("Nothing to request: every component of the selected Production Orders is already requested.");
            else Ok(rows.Count + " material row(s) loaded" + (skipped > 0 ? ", " + skipped + " component(s) already fully requested were left out." : "."));
        }

        private void Clear()
        {
            Set("eProj", "");
            Set("eFind", "");
            orders = new List<ProductionOrder>();
            rows = new List<PlanRow>();
            HideLastPr();
            ShowOrders();
            ShowRows();
            Ok("Filters and rows cleared.");
        }
        #endregion

        #region Rows
        private static string Base(PlanRow r) { return r.PdoNum + " / " + r.PdoLine; }

        private void ShowRows()
        {
            Freeze(() =>
            {
                gRows.Fill(rows.Select(r => new object[]
                {
                    r.Selected ? "Y" : "N", r.ItemCode, r.Description, r.Quantity, r.Uom, r.Price, r.Discount, r.PriceAfterDiscount, r.Total,
                    r.Warehouse, r.Vendor, ToDate(r.RequiredDate), ToDate(r.ValidUntil), Base(r)
                }).ToList());
                for (int i = 0; i < rows.Count; i++)
                    gRows.Color(i, !rows[i].Selected ? GridTable.Grey : rows[i].Quantity > rows[i].Balance ? GridTable.Red
                        : rows[i].HasLinkedRequest ? GridTable.Green : GridTable.NoColor);
                decimal total = rows.Where(r => r.Selected).Sum(r => r.Total);
                Set("eTotal", total);
                int linked = rows.Count(r => r.HasLinkedRequest);
                if (rows.Count > 0) Caption("lFoot", rows.Count(r => r.Selected) + " of " + rows.Count + " row(s) selected" +
                    (linked > 0 ? " (" + linked + " shown in green: source Production Order already has a Purchase Request)" : ""));
            });
        }

        /// <summary>Takes what the user typed in one grid row into the model.</summary>
        private void ReadRow(int tr)
        {
            PlanRow r = rows[tr];
            r.Selected = gRows.TypedText(tr, "Sel") == "Y";
            r.Quantity = gRows.Typed(tr, "Qty");
            r.Price = gRows.Typed(tr, "Price");
            r.Discount = gRows.Typed(tr, "Disc");
            r.Warehouse = gRows.TypedText(tr, "Whs");
            r.Vendor = gRows.TypedText(tr, "Vend");
            r.RequiredDate = gRows.TypedDate(tr, "Req");
            r.ValidUntil = gRows.TypedDate(tr, "Valid");
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
            if (!rows.Any(r => r.Selected)) throw new PlanException(400, "Select at least one material row first.");
            if (!Ask("Vendor, Warehouse, Required Date and Valid Until Date of the selected rows will be replaced with the header values. Continue?", "Replace", "Cancel")) return;
            PlanService.ApplyHeader(ReadHeader(), rows.Where(r => r.Selected));
            ShowRows();
            Ok("Header values applied to " + rows.Count(r => r.Selected) + " row(s).");
        }

        private void ApplyWarehouseOnly()
        {
            if (!rows.Any(r => r.Selected)) throw new PlanException(400, "Select at least one material row first.");
            string whs = Text("eWhs");
            if (whs.Length == 0) throw new PlanException(400, "Enter the Warehouse first.", new List<Dictionary<string, string>> { new Dictionary<string, string> { { "field", "eWhs" } } });
            if (!Ask("Warehouse of the selected rows will be replaced with " + whs + ". Continue?", "Replace", "Cancel")) return;
            PlanService.ApplyWarehouse(whs, rows.Where(r => r.Selected));
            ShowRows();
            Ok("Warehouse applied to " + rows.Count(r => r.Selected) + " row(s).");
        }

        private void SetAllSelected(bool on)
        {
            if (rows.Count == 0) return;
            foreach (PlanRow r in rows) r.Selected = on;
            ShowRows();
        }

        private void CreateRequest()
        {
            for (int i = 0; i < rows.Count; i++) ReadRow(i);
            CreatedRequest made = PlanService.Create(ReadHeader(), rows);
            // One request per load: the rows are cleared so Create cannot be pressed twice for the same materials
            rows = new List<PlanRow>();
            ShowRows();
            lastPrEntry = made.DocEntry; lastPrNum = made.DocNum;
            ButtonCaption("bLastPr", "Open Purchase Request " + made.DocNum);
            Visible("bLastPr", true);
            if (made.Warning.Length > 0) Warn(made.Warning);
            Ok("Purchase Request " + made.DocNum + " was created.");
        }

        private void HideLastPr()
        {
            lastPrEntry = 0;
            Visible("bLastPr", false);
        }

        private int CurrentPdo()
        {
            if (gRows.Current < 0 || gRows.Current >= rows.Count) throw new PlanException(400, "Click a material row first.");
            return rows[gRows.Current].PdoEntry;
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
            if (e.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && e.BeforeAction && e.ItemUID == "eProj") FilterProjectCfl();
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
            else if ((chosen = ChosenGridCodes(e, gRows, "Whs", "WhsCode", out gridRow)).Count > 0 && gridRow >= 0 && gridRow < rows.Count)
            {
                rows[gridRow].Warehouse = chosen[0];
                ShowRows();
                KeepOkMode("Cancel");
            }
            else if ((chosen = ChosenGridCodes(e, gRows, "Vend", "CardCode", out gridRow)).Count > 0 && gridRow >= 0 && gridRow < rows.Count)
            {
                rows[gridRow].Vendor = chosen[0];
                ShowRows();
                KeepOkMode("Cancel");
            }
            else if (Pressed(e, "bFind") || Enter(e, "eFind", "eProj")) Search();
            else if (Pressed(e, "bClear")) Clear();
            else if (e.ItemUID == "cSeries" && !e.BeforeAction &&
                (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)) RefreshNextNo();
            else if (Pressed(e, "bLoad")) Load();
            else if (Pressed(e, "bOSelAll")) SetAllOrdersSelected(true);
            else if (Pressed(e, "bODeselAl")) SetAllOrdersSelected(false);
            else if (Pressed(e, "bApply")) Apply();
            else if (Pressed(e, "bApplyWhs")) ApplyWarehouseOnly();
            else if (Pressed(e, "bSelAll")) SetAllSelected(true);
            else if (Pressed(e, "bDeselAll")) SetAllSelected(false);
            else if (Pressed(e, "bCreate")) CreateRequest();
            else if (Pressed(e, "bLastPr")) { if (lastPrEntry > 0) Store.OpenPurchaseRequest(lastPrEntry); }
            else if (Pressed(e, "bLinked")) LinkedRequestForm.Show(CurrentPdo());
            else if (Clicked(e, gRows)) gRows.Click(e.Row);
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
                if (tr >= 0 && tr < rows.Count) Store.OpenProductionOrder(rows[tr].PdoEntry);
            }
            else if (e.ItemUID == "gRows" && !e.BeforeAction && e.Row >= 0 &&
                ((e.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && e.ItemChanged) || (e.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && e.ColUID == "Sel")))
            {
                int tr = gRows.Row(e.Row);
                if (tr < 0 || tr >= rows.Count) return;
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

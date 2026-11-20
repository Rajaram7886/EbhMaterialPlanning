using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// All SAP access of the add-on: read with Recordset (Db), Purchase Request and Production Order written through DI API.
    /// The link between the two documents is kept on the Purchase Request line (U_EBH_PRODUCTIONNO / U_EBH_PRODLINE);
    /// what is requested on a Production Order component is always read from the open Purchase Request lines,
    /// so a cancelled or closed request frees its quantity without any extra bookkeeping.
    /// </summary>
    public sealed class PlanStore
    {
        public const string FieldPdo = "U_EBH_PRODUCTIONNO", FieldPdoLine = "U_EBH_PRODLINE";
        /// <summary>Planned Price After Discount / Line Total of a Purchase Request (and Purchase Order) line - see CreatePurchaseRequest.</summary>
        public const string FieldPafd = "U_EBH_PRICEAFTERDISCOUNT", FieldLineTotal = "U_EBH_LINETOTAL";
        private const int ObjPurchaseRequest = 1470000113, ObjProductionOrder = 202, ObjPurchaseOrder = 22, ObjInventoryTransferRequest = 1250000001;
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        #region Helpers
        private static SAPbobsCOM.Company Company { get { return Db.Company; } }

        private static decimal Dec(Dictionary<string, object> row, string key)
        {
            object v;
            if (row == null || !row.TryGetValue(key, out v) || v == null || v is DBNull) return 0;
            try { return Convert.ToDecimal(v, Inv); }
            catch (FormatException) { return 0; }
        }

        private static int Int(Dictionary<string, object> row, string key) { return (int)Dec(row, key); }

        private static string Date(Dictionary<string, object> row, string key)
        {
            object v;
            if (row == null || !row.TryGetValue(key, out v) || v == null || v is DBNull) return "";
            DateTime d;
            if (v is DateTime) d = (DateTime)v;
            else if (!DateTime.TryParse(Convert.ToString(v), out d)) return "";
            return d.Year < 1950 ? "" : d.ToString("yyyy-MM-dd");
        }

        private static DateTime ToDate(string iso)
        {
            return DateTime.ParseExact(iso, "yyyy-MM-dd", Inv);
        }

        private static string Like(string search)
        {
            return Db.Quote("%" + (search ?? "").Trim().Replace("[", "").Replace("%", "").Replace("_", "") + "%");
        }

        private static void Check(int result, string what)
        {
            if (result == 0) return;
            int code; string msg;
            Company.GetLastError(out code, out msg);
            throw new PlanException(500, what + " : " + msg + " (" + code + ")");
        }

        /// <summary>The add-on's remark goes to the header UDF U_EBH_REMARK (OPRQ / OPOR), not the standard Comments / Remarks, which stay free for the user.</summary>
        private static void SetRemark(SAPbobsCOM.UserFields header, string text)
        {
            try { header.Fields.Item("U_EBH_REMARK").Value = Cut(text, 254); }
            catch (COMException)
            {
                throw new PlanException(500, "The field U_EBH_REMARK is not set up in this company. Run Database Setup (Create All Environment), then restart SAP Business One.");
            }
        }

        private static void Release(object com)
        {
            if (com != null) Marshal.ReleaseComObject(com);
        }

        /// <summary>Quoted, comma separated list of codes for an IN (...) clause. Codes are escaped; an empty list gives ''.</summary>
        private static string InList(IEnumerable<string> codes)
        {
            var list = codes.Select(c => (c ?? "").Trim()).Where(c => c.Length > 0).Distinct().Select(Db.Quote).ToList();
            return list.Count == 0 ? "''" : string.Join(",", list.ToArray());
        }

        private static string InList(IEnumerable<int> numbers)
        {
            var list = numbers.Distinct().Select(n => n.ToString(Inv)).ToList();
            return list.Count == 0 ? "-1" : string.Join(",", list.ToArray());
        }

        private bool? hasPafdi;
        /// <summary>
        /// SQL for the Price After Discount of a WOR1 row aliased <paramref name="w"/>: U_EBH_PAFDI as-is when set,
        /// otherwise U_EBH_Unit_Price less U_EBH_Discount %. U_EBH_PAFDI is not created by Database Setup, so whether
        /// it exists in this company is checked once (CUFD) instead of assumed - the query must not fail without it.
        /// </summary>
        private string PafdiExpr(string w)
        {
            if (hasPafdi == null)
                hasPafdi = Db.Rows("Select 1 as \"X\" from CUFD Where \"TableID\" = 'WOR1' and \"AliasID\" = 'EBH_PAFDI'").Count > 0;
            string computed = "IsNull(" + w + ".\"U_EBH_Unit_Price\", 0) * (1 - IsNull(" + w + ".\"U_EBH_Discount\", 0) / 100)";
            return hasPafdi.Value
                ? "(Case When IsNull(" + w + ".\"U_EBH_PAFDI\", 0) <> 0 Then " + w + ".\"U_EBH_PAFDI\" Else " + computed + " End)"
                : "(" + computed + ")";
        }
        #endregion

        #region Stock
        /// <summary>Warehouse on-hand and bin quantities (system bin / project bins) of the given items, every warehouse - see StockMap.</summary>
        public StockMap GetStock(IEnumerable<string> itemCodes)
        {
            var map = new StockMap();
            var items = itemCodes.Where(c => !string.IsNullOrEmpty(c)).Distinct().ToList();
            if (items.Count == 0) return map;
            foreach (var r in Db.Rows("Select \"ItemCode\", \"WhsCode\", \"OnHand\" from OITW Where \"ItemCode\" in (" + InList(items) + ")"))
                map.AddOnHand(Db.Str(r, "ItemCode"), Db.Str(r, "WhsCode"), Dec(r, "OnHand"));
            foreach (var r in Db.Rows(
                "Select Q.\"ItemCode\", Q.\"WhsCode\", B.\"SysBin\", IsNull(B.\"SL1Code\", '') as \"Project\", Sum(Q.\"OnHandQty\") as \"Qty\" " +
                "from OIBQ Q Inner Join OBIN B on B.\"AbsEntry\" = Q.\"BinAbs\" " +
                "Where Q.\"ItemCode\" in (" + InList(items) + ") " +
                "Group by Q.\"ItemCode\", Q.\"WhsCode\", B.\"SysBin\", B.\"SL1Code\""))
            {
                if (Db.Str(r, "SysBin") == "Y") map.AddSystemBin(Db.Str(r, "ItemCode"), Db.Str(r, "WhsCode"), Dec(r, "Qty"));
                else if (Db.Str(r, "Project").Length > 0) map.AddProjectBin(Db.Str(r, "ItemCode"), Db.Str(r, "WhsCode"), Db.Str(r, "Project"), Dec(r, "Qty"));
            }
            return map;
        }
        #endregion

        #region Lists for the form
        public string UserCode { get { return Company.UserName; } }

        public List<KeyValuePair<string, string>> Projects()
        {
            return Db.Rows("Select \"PrjCode\", \"PrjName\" from OPRJ Where \"Active\" = 'Y' Order by \"PrjCode\"")
                .Select(r => new KeyValuePair<string, string>(Db.Str(r, "PrjCode"), Db.Str(r, "PrjName"))).ToList();
        }

        /// <summary>Distinct, non-blank Project codes of the given Production Order numbers (to narrow the Project Choose From List).</summary>
        public List<string> ProjectsOfOrders(IEnumerable<int> docNums)
        {
            var nums = docNums.Distinct().ToList();
            if (nums.Count == 0) return new List<string>();
            return Db.Rows("Select Distinct \"Project\" from OWOR Where \"DocNum\" in (" + InList(nums) + ") and \"Project\" is not null and \"Project\" <> ''")
                .Select(r => Db.Str(r, "Project")).Where(p => p.Length > 0).ToList();
        }

        /// <summary>Numbering series of the given document object (e.g. the Purchase Request or Purchase Order).</summary>
        public List<KeyValuePair<string, string>> SeriesFor(int objectCode)
        {
            return Db.Rows("Select \"Series\", \"SeriesName\" from NNM1 Where \"ObjectCode\" = '" + objectCode + "' and \"Locked\" = 'N' Order by \"Series\"")
                .Select(r => new KeyValuePair<string, string>(Db.Str(r, "Series"), Db.Str(r, "SeriesName"))).ToList();
        }

        public int DefaultSeriesFor(int objectCode)
        {
            var row = Db.FirstRow("Select \"DfltSeries\" from ONNM Where \"ObjectCode\" = '" + objectCode + "'");
            return Int(row, "DfltSeries");
        }

        /// <summary>Numbering series of the Purchase Request (object 1470000113).</summary>
        public List<KeyValuePair<string, string>> Series() { return SeriesFor(ObjPurchaseRequest); }

        public int DefaultSeries() { return DefaultSeriesFor(ObjPurchaseRequest); }

        public List<KeyValuePair<string, string>> Warehouses()
        {
            return Db.Rows("Select \"WhsCode\", \"WhsName\" from OWHS Where \"Inactive\" = 'N' Order by \"WhsCode\"")
                .Select(r => new KeyValuePair<string, string>(Db.Str(r, "WhsCode"), Db.Str(r, "WhsName"))).ToList();
        }

        public List<KeyValuePair<string, string>> Vendors()
        {
            return Db.Rows("Select \"CardCode\", \"CardName\" from OCRD Where \"CardType\" = 'S' and \"frozenFor\" = 'N' Order by \"CardCode\"")
                .Select(r => new KeyValuePair<string, string>(Db.Str(r, "CardCode"), Db.Str(r, "CardName"))).ToList();
        }

        /// <summary>Item manufacturers (OMRC), to filter the Material Request Planning grid by OITM.FirmCode.</summary>
        public List<KeyValuePair<string, string>> Manufacturers()
        {
            return Db.Rows("Select \"FirmCode\", \"FirmName\" from OMRC Order by \"FirmName\"")
                .Select(r => new KeyValuePair<string, string>(Db.Str(r, "FirmCode"), Db.Str(r, "FirmName"))).ToList();
        }

        public bool VendorExists(string code)
        {
            return Db.FirstRow("Select \"CardCode\" from OCRD Where \"CardType\" = 'S' and \"CardCode\" = " + Db.Quote(code)) != null;
        }

        public bool WarehouseExists(string code)
        {
            return Db.FirstRow("Select \"WhsCode\" from OWHS Where \"WhsCode\" = " + Db.Quote(code)) != null;
        }

        /// <summary>Next number the chosen series will give the document (display only, SAP assigns the real one on Add).</summary>
        public int NextNumberFor(int series, int objectCode)
        {
            if (series <= 0) return 0;
            var row = Db.FirstRow("Select \"NextNumber\" from NNM1 Where \"Series\" = " + series + " and \"ObjectCode\" = '" + objectCode + "'");
            return Int(row, "NextNumber");
        }

        /// <summary>Next number the chosen series will give the Purchase Request (display only, SAP assigns the real one on Add).</summary>
        public int NextNumber(int series) { return NextNumberFor(series, ObjPurchaseRequest); }

        /// <summary>True when the company has the Branches feature active (a branch must then be set on the document).</summary>
        public bool MultiBranch
        {
            get { return Db.FirstRow("Select Top 1 \"BPLId\" from OBPL Where \"Disabled\" = 'N'") != null; }
        }

        public List<KeyValuePair<string, string>> Branches()
        {
            return Db.Rows("Select \"BPLId\", \"BPLName\" from OBPL Where \"Disabled\" = 'N' Order by \"BPLId\"")
                .Select(r => new KeyValuePair<string, string>(Db.Str(r, "BPLId"), Db.Str(r, "BPLName"))).ToList();
        }

        /// <summary>
        /// First active branch, used to pre-select the Branch field. Picking the user's own default branch would need
        /// the exact user-branch assignment table layout, which was not reliable across companies; the user can still
        /// change the combo before creating the request.
        /// </summary>
        public int DefaultBranch()
        {
            var row = Db.FirstRow("Select Top 1 \"BPLId\" from OBPL Where \"Disabled\" = 'N' Order by \"BPLId\"");
            return Int(row, "BPLId");
        }
        #endregion

        #region Production Orders
        /// <summary>
        /// Planned and released Production Orders of the projects (all projects when none is given), optionally
        /// narrowed by a Production Order number, product or description.
        /// </summary>
        public List<ProductionOrder> ListOrders(IEnumerable<string> projects, string find)
        {
            string where = "T0.\"Status\" in ('P','R')";
            var prj = projects.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
            if (prj.Count > 0) where += " and T0.\"Project\" in (" + InList(prj) + ")";
            string f = (find ?? "").Trim();
            if (f.Length > 0)
            {
                // Several Production Order numbers chosen from the Choose From List (comma / semicolon separated)
                var tokens = f.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim()).Where(t => t.Length > 0).ToList();
                var nums = new List<int>();
                bool allNumbers = tokens.Count > 0;
                foreach (string t in tokens)
                {
                    int n;
                    if (int.TryParse(t, out n)) nums.Add(n);
                    else { allNumbers = false; break; }
                }
                if (allNumbers)
                    where += " and T0.\"DocNum\" in (" + InList(nums) + ")";
                else
                {
                    int num;
                    where += " and (" + (int.TryParse(f, out num) ? "T0.\"DocNum\" = " + num + " or " : "") +
                        "T0.\"ItemCode\" like " + Like(f) + " or T0.\"ProdName\" like " + Like(f) + ")";
                }
            }
            var baseRows = Db.Rows("Select Top 500 T0.\"DocEntry\", T0.\"DocNum\", T0.\"Series\", N1.\"SeriesName\", T0.\"Project\", IsNull(J.\"PrjName\", '') as \"PrjName\", " +
                "T0.\"ItemCode\", T0.\"ProdName\", T0.\"Status\", T0.\"DueDate\" " +
                "from OWOR T0 Left Join NNM1 N1 on N1.\"Series\" = T0.\"Series\" and N1.\"ObjectCode\" = '" + ObjProductionOrder + "' " +
                "Left Join OPRJ J on J.\"PrjCode\" = T0.\"Project\" " +
                "Where " + where + " Order by T0.\"DocNum\" Desc");
            // See ComponentProgress: the line-wise U_EBH_PRODUCTIONNO / U_EBH_PRODLINE count behind the progress column.
            Dictionary<int, Progress> progress = ComponentProgress(baseRows.Select(r => Int(r, "DocEntry")));
            HashSet<int> closed = OrdersWithAllRequestsClosed(baseRows.Select(r => Int(r, "DocEntry")));
            return baseRows.Select(r =>
            {
                int entry = Int(r, "DocEntry");
                Progress p;
                progress.TryGetValue(entry, out p);
                return new ProductionOrder
                {
                    DocEntry = entry, DocNum = Int(r, "DocNum"), Series = Int(r, "Series"), SeriesName = Db.Str(r, "SeriesName"),
                    // Shows both, since the add-on's own grid cannot mirror a user's per-field Choose From List display
                    // setting (Project Code vs Project Name) the way SAP's native Production Order list can.
                    Project = Db.Str(r, "Project") + (Db.Str(r, "PrjName").Length > 0 ? " - " + Db.Str(r, "PrjName") : ""),
                    ItemCode = Db.Str(r, "ItemCode"), ItemName = Db.Str(r, "ProdName"), Status = Db.Str(r, "Status"), DueDate = Date(r, "DueDate"),
                    TotalComponents = p.Total, CompletedComponents = p.Done, RequestsClosed = closed.Contains(entry)
                };
            }).ToList();
        }

        /// <summary>
        /// Production Orders (of the given ones) whose Purchase Request lines are all Closed: at least one line linked
        /// through @EBH_PRQMAP, and every linked line of a not-cancelled request has PRQ1.LineStatus 'C'. Reads the live
        /// status, not the cached U_PrStatus. Cancelled requests are ignored.
        /// </summary>
        private HashSet<int> OrdersWithAllRequestsClosed(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.Distinct().ToList();
            var result = new HashSet<int>();
            if (ids.Count == 0) return result;
            // Purchase Request lines and (when set up) Inventory Transfer Request lines together: Completed only when
            // every linked line of either kind is Closed
            string sql =
                "Select M.\"U_PdoEntry\" as \"Pdo\", Count(*) as \"Lines\", " +
                "Sum(Case When P1.\"LineStatus\" = 'C' Then 1 Else 0 End) as \"Closed\" " +
                "from " + MapTable + " M " +
                "Inner Join PRQ1 P1 on P1.\"DocEntry\" = M.\"U_PrEntry\" and P1.\"LineNum\" = M.\"U_PrLine\" " +
                "Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Where M.\"U_PdoEntry\" in (" + InList(ids) + ") and P0.\"CANCELED\" = 'N' " +
                "Group by M.\"U_PdoEntry\"";
            if (HasItrMap)
                sql += " Union All " +
                    "Select M.\"U_PdoEntry\" as \"Pdo\", Count(*) as \"Lines\", " +
                    "Sum(Case When T1.\"LineStatus\" = 'C' Then 1 Else 0 End) as \"Closed\" " +
                    "from " + ItrMapTable + " M " +
                    "Inner Join WTQ1 T1 on T1.\"DocEntry\" = M.\"U_ItrEntry\" and T1.\"LineNum\" = M.\"U_ItrLine\" " +
                    "Inner Join OWTQ T0 on T0.\"DocEntry\" = T1.\"DocEntry\" " +
                    "Where M.\"U_PdoEntry\" in (" + InList(ids) + ") and T0.\"CANCELED\" = 'N' " +
                    "Group by M.\"U_PdoEntry\"";
            foreach (var g in Db.Rows(sql).GroupBy(r => Int(r, "Pdo")))
            {
                int lines = g.Sum(r => Int(r, "Lines"));
                if (lines > 0 && g.Sum(r => Int(r, "Closed")) == lines) result.Add(g.Key);
            }
            return result;
        }

        /// <summary>Item-type component count and how many are already fully covered by Purchase Requests (still open,
        /// or already drawn into a Purchase Order), per Production Order - feeds the Production Orders grid's progress
        /// count and colour (full / partial / none). Reads @EBH_PRQMAP (via RequestedMap), not U_EBH_PRODUCTIONNO /
        /// U_EBH_PRODLINE: those two only name a line's own order, so a component consolidated onto a DIFFERENT
        /// order's Purchase Request line would never be recognised there. This count matches the outstanding
        /// quantity GetComponents / Balance rely on.</summary>
        private struct Progress { public int Total, Done; }

        private Dictionary<int, Progress> ComponentProgress(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.Distinct().ToList();
            var result = new Dictionary<int, Progress>();
            if (ids.Count == 0) return result;
            Dictionary<string, decimal> requested = RequestedMap(ids);
            foreach (var r in Db.Rows("Select \"DocEntry\", \"LineNum\", \"PlannedQty\" from WOR1 Where \"DocEntry\" in (" + InList(ids) + ") and \"ItemType\" = 4"))
            {
                int entry = Int(r, "DocEntry"), line = Int(r, "LineNum");
                decimal req;
                requested.TryGetValue(entry + ":" + line, out req);
                Progress p;
                result.TryGetValue(entry, out p);
                p.Total++;
                if (Dec(r, "PlannedQty") <= req) p.Done++;
                result[entry] = p;
            }
            return result;
        }

        /// <summary>
        /// Item components (not resources) of the Production Orders with the quantity already on open Purchase Requests.
        /// One Component per Production Order line, so the same item in two orders stays on two rows.
        /// </summary>
        public List<Component> GetComponents(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.ToList();
            Dictionary<string, decimal> requested = RequestedMap(ids);
            var components = Db.Rows(
                "Select T0.\"DocEntry\", T0.\"DocNum\", T0.\"Project\", T0.\"ItemCode\" as \"PdoItemCode\", T1.\"LineNum\", T1.\"ItemCode\", I.\"ItemName\", T1.\"PlannedQty\", T1.\"wareHouse\", " +
                "IsNull(U.\"UomCode\", I.\"BuyUnitMsr\") as \"Uom\", " +
                ManufacturerExpr + " as \"Manufacturer\", " +
                "IsNull(M.\"FirmName\", '') as \"ManufacturerName\", IsNull(I.\"U_EBH_MfgCatlog\", '') as \"Catalog\", " +
                "IsNull(G.\"U_EBH_JOBCODNO\", 'N') as \"JobFlag\", IsNull(T1.\"U_EBH_Unit_Price\", 0) as \"Price\", IsNull(T1.\"U_EBH_Discount\", 0) as \"Discount\", " +
                PafdiExpr("T1") + " as \"Pafdi\" " +
                "from OWOR T0 Inner Join WOR1 T1 on T1.\"DocEntry\" = T0.\"DocEntry\" " +
                "Inner Join OITM I on I.\"ItemCode\" = T1.\"ItemCode\" " +
                "Left Join OUOM U on U.\"UomEntry\" = T1.\"UomEntry\" " +
                "Left Join OMRC M on M.\"FirmCode\" = I.\"FirmCode\" " +
                "Left Join OITB G on G.\"ItmsGrpCod\" = I.\"ItmsGrpCod\" " +
                "Where T0.\"DocEntry\" in (" + InList(ids) + ") and T1.\"ItemType\" = 4 and T0.\"Status\" in ('P','R') " +
                "Order by T0.\"DocNum\", T1.\"VisOrder\"")
                .Select(r =>
                {
                    int pdoEntry = Int(r, "DocEntry"), pdoLine = Int(r, "LineNum");
                    decimal req;
                    requested.TryGetValue(pdoEntry + ":" + pdoLine, out req);
                    return new Component
                    {
                        PdoEntry = pdoEntry, PdoNum = Int(r, "DocNum"), PdoLine = pdoLine, Project = Db.Str(r, "Project"),
                        PdoItemCode = Db.Str(r, "PdoItemCode"),
                        ItemCode = Db.Str(r, "ItemCode"), Description = Db.Str(r, "ItemName"), Planned = Dec(r, "PlannedQty"),
                        OpenRequested = req, Uom = Db.Str(r, "Uom"), Warehouse = Db.Str(r, "wareHouse"),
                        Manufacturer = Db.Str(r, "Manufacturer"), ManufacturerName = Db.Str(r, "ManufacturerName"), Catalog = Db.Str(r, "Catalog"),
                        TracksJobCode = Db.Str(r, "JobFlag") == "Y", Price = Dec(r, "Price"), Discount = Dec(r, "Discount"),
                        PriceAfterDiscount = Math.Round(Dec(r, "Pafdi"), 6)
                    };
                }).ToList();
            // The same item on several lines of one Production Order (e.g. Price After Discount 0 / 20 / 80) is planned
            // at the lowest priced one (20) on every line - even a line loaded alone, whose siblings are already requested
            foreach (var g in components.GroupBy(c => c.PdoEntry + "|" + c.ItemCode).Where(g => g.Count() > 1))
            {
                Component cheapest = PlanService.Cheapest(g.ToList(), c => c.PriceAfterDiscount);
                foreach (Component c in g) { c.Price = cheapest.Price; c.Discount = cheapest.Discount; c.PriceAfterDiscount = cheapest.PriceAfterDiscount; }
            }
            return components;
        }
        #endregion

        #region Production Order component map (@EBH_PRQMAP)
        /// <summary>
        /// The real source of truth for which Production Order component(s) a Purchase Request line covers, and with
        /// how much quantity each - one row of @EBH_PRQMAP per component, written when the line is created (see
        /// CreatePurchaseRequest). Replaces an earlier design that packed the same information into a fixed-size text
        /// field on the line itself (U_EBH_PDOMAP, no longer created or written): that field hit a real size limit
        /// and its comma-separated numbers could be misread. A table has neither problem. Requests created before
        /// this table existed are copied into it by BackfillLegacyMap.
        /// </summary>
        private const string MapTable = "\"@EBH_PRQMAP\"";

        private static bool UdfExists(string table, string alias)
        {
            string alt = table.StartsWith("@") ? table.Substring(1) : "@" + table;
            return Db.Rows("Select 1 as \"X\" from CUFD Where (\"TableID\" = " + Db.Quote(table) + " or \"TableID\" = " + Db.Quote(alt) + ") and \"AliasID\" = " + Db.Quote(alias)).Count > 0;
        }

        /// <summary>
        /// Copies Purchase Request lines raised by earlier versions of the add-on (before @EBH_PRQMAP existed) into
        /// the map: one row per PRQ1 line that names its Production Order in U_EBH_PRODUCTIONNO and component line in
        /// U_EBH_PRODLINE (a single number back then), with the line's whole Quantity, and only when that line has no
        /// map row yet - so it is safe to run any number of times. Must run before U_EBH_PRODLINE is converted to
        /// text (DBSetup.ConvertProdLineFieldToText), which erases the old line numbers. Returns the rows added;
        /// does nothing (0) while the map table or the old fields are missing.
        /// </summary>
        public static int BackfillLegacyMap()
        {
            if (!UdfExists("PRQ1", "EBH_PRODUCTIONNO") || !UdfExists("PRQ1", "EBH_PRODLINE") || !UdfExists("EBH_PRQMAP", "PrStatus")) return 0;
            string where =
                "from PRQ1 P1 Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Where P1.\"" + FieldPdo + "\" is not null and P1.\"" + FieldPdo + "\" > 0 " +
                "and Try_Cast(Cast(P1.\"" + FieldPdoLine + "\" as nvarchar(50)) as int) is not null " +
                "and not exists (Select 1 from " + MapTable + " X Where X.\"U_PrEntry\" = P1.\"DocEntry\" and X.\"U_PrLine\" = P1.\"LineNum\")";
            int missing = Int(Db.FirstRow("Select Count(*) as \"N\" " + where), "N");
            if (missing == 0) return 0;
            Db.Exec(
                "Insert into " + MapTable + " (\"Code\",\"Name\",\"U_PrEntry\",\"U_PrLine\",\"U_PdoEntry\",\"U_PdoLine\",\"U_Qty\",\"U_PrStatus\") " +
                "Select K.\"C\", K.\"C\", K.\"DocEntry\", K.\"LineNum\", K.\"Pdo\", K.\"PdoLine\", K.\"Quantity\", K.\"St\" from (" +
                "Select Replace(Cast(NewId() as nvarchar(36)), '-', '') as \"C\", P1.\"DocEntry\", P1.\"LineNum\", P1.\"" + FieldPdo + "\" as \"Pdo\", " +
                "Try_Cast(Cast(P1.\"" + FieldPdoLine + "\" as nvarchar(50)) as int) as \"PdoLine\", P1.\"Quantity\", " +
                "Case When P0.\"CANCELED\" = 'Y' Then 'X' Else P0.\"DocStatus\" End as \"St\" " + where + ") K");
            return missing;
        }

        /// <summary>Writes one @EBH_PRQMAP row. Runs on the same DI API connection as the Purchase Request's own
        /// Add() - inside CreatePurchaseRequest's transaction, so both commit or roll back together. U_PrStatus
        /// starts as 'O' (the line is freshly added and open at this point) - a reference column for a person
        /// inspecting the table directly, refreshed later by SyncMapStatus whenever the Purchase Request's real
        /// status changes (every actual "already requested" lookup reads the live status via the join regardless -
        /// see RequestedMap - so this column lagging briefly behind never causes a wrong result, only a stale display).</summary>
        private static void InsertMapRow(int prEntry, int prLine, int pdoEntry, int pdoLine, decimal qty)
        {
            string code = Guid.NewGuid().ToString("N");
            Db.Exec("Insert into " + MapTable + " (\"Code\",\"Name\",\"U_PrEntry\",\"U_PrLine\",\"U_PdoEntry\",\"U_PdoLine\",\"U_Qty\",\"U_PrStatus\") Values (" +
                Db.Quote(code) + "," + Db.Quote(code) + "," + prEntry + "," + prLine + "," + pdoEntry + "," + pdoLine + "," + qty.ToString(Inv) + "," + Db.Quote("O") + ")");
        }

        /// <summary>
        /// Brings U_PrStatus back in line with the real status of each row's Purchase Request line - 'O' Open, 'C'
        /// Closed, or 'X' Cancelled (OPRQ.CANCELED, a standard SAP document action this add-on has no event to react
        /// to the moment it happens). Called every time Sync() runs (after creating a new request, and whenever the
        /// Linked Documents screen reloads), so a Purchase Request cancelled through its own standard document is
        /// reflected here the next time this add-on touches its Production Order, even though it did not cause the
        /// cancellation itself. One UPDATE, not a row read/write loop in C#, so this stays cheap regardless of how
        /// many mapping rows exist.
        /// </summary>
        private void SyncMapStatus(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.Distinct().ToList();
            if (ids.Count == 0) return;
            ExecSyncMapStatus(" Where \"U_PdoEntry\" in (" + InList(ids) + ")");
            SyncItrMapStatus(" Where \"U_PdoEntry\" in (" + InList(ids) + ")");
        }

        /// <summary>SyncMapStatus for every row in the table at once - run when Material Request Planning opens (see
        /// MaterialPlanForm.Build), so U_PrStatus is already current the moment the form is on screen, not only for
        /// whichever Production Orders happen to get touched by something this add-on does afterwards. Checks first
        /// (one cheap "is anything actually stale" query) and only runs the UPDATE when something is - on a company
        /// with a large table and nothing to fix (the common case: most opens follow one that already synced), this
        /// is the difference between touching every row on every single form open and touching none of them.</summary>
        public void SyncAllMapStatus()
        {
            BackfillLegacyMap();
            bool stale = Db.Rows(
                "Select Top 1 M.\"Code\" from " + MapTable + " M " +
                "Inner Join PRQ1 P1 on P1.\"DocEntry\" = M.\"U_PrEntry\" and P1.\"LineNum\" = M.\"U_PrLine\" " +
                "Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Where M.\"U_PrStatus\" <> (Case When P0.\"CANCELED\" = 'Y' Then 'X' Else P0.\"DocStatus\" End)")
                .Count > 0;
            if (stale) ExecSyncMapStatus("");
            if (HasItrMap && Db.Rows(
                "Select Top 1 M.\"Code\" from " + ItrMapTable + " M Inner Join OWTQ T0 on T0.\"DocEntry\" = M.\"U_ItrEntry\" " +
                "Where M.\"U_ItrStatus\" <> (Case When T0.\"CANCELED\" = 'Y' Then 'X' Else T0.\"DocStatus\" End)").Count > 0)
                SyncItrMapStatus("");
        }

        private void ExecSyncMapStatus(string where)
        {
            Db.Exec(
                "Update " + MapTable + " Set \"U_PrStatus\" = " +
                "(Select Case When P0.\"CANCELED\" = 'Y' Then 'X' Else P0.\"DocStatus\" End " +
                "from PRQ1 P1 Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Where P1.\"DocEntry\" = " + MapTable + ".\"U_PrEntry\" and P1.\"LineNum\" = " + MapTable + ".\"U_PrLine\")" + where);
        }

        /// <summary>Quantity already covered for each "entry:line" component among the given Production Orders, summed
        /// across every not-cancelled Purchase Request line that covers it: what is still open on the line (open lines
        /// only) plus what has already been drawn from it into not-cancelled Purchase Orders, capped at the line's
        /// Quantity, then proportioned by the component's own share of the line (U_Qty / Quantity - SAP has no sub-line
        /// record of which portion of a merged line belongs to which component, so this is the best estimate). Counting
        /// the ordered part keeps a component from showing its balance again once its request is turned into a Purchase
        /// Order. A line closed by hand with nothing ordered counts as 0. What Component.OpenRequested /
        /// PlanRow.AlreadyRequested and the Production Orders grid's progress count are read from.</summary>
        private Dictionary<string, decimal> RequestedMap(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.Distinct().ToList();
            var result = new Dictionary<string, decimal>();
            if (ids.Count == 0) return result;
            string covered =
                "(Case When P1.\"LineStatus\" = 'O' and P0.\"DocStatus\" = 'O' Then P1.\"OpenQty\" Else 0 End + IsNull(PO.\"Ordered\", 0))";
            foreach (var r in Db.Rows(
                "Select M.\"U_PdoEntry\" as \"Pdo\", M.\"U_PdoLine\" as \"Line\", " +
                "Sum(M.\"U_Qty\" * ((Case When " + covered + " > P1.\"Quantity\" Then P1.\"Quantity\" Else " + covered + " End) / NullIf(P1.\"Quantity\", 0))) as \"Share\" " +
                "from " + MapTable + " M " +
                "Inner Join PRQ1 P1 on P1.\"DocEntry\" = M.\"U_PrEntry\" and P1.\"LineNum\" = M.\"U_PrLine\" " +
                "Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Left Join (Select Q.\"BaseEntry\", Q.\"BaseLine\", Sum(Q.\"Quantity\") as \"Ordered\" from POR1 Q " +
                "Inner Join OPOR O on O.\"DocEntry\" = Q.\"DocEntry\" " +
                "Where Q.\"BaseType\" = " + ObjPurchaseRequest + " and O.\"CANCELED\" = 'N' Group by Q.\"BaseEntry\", Q.\"BaseLine\") PO " +
                "on PO.\"BaseEntry\" = P1.\"DocEntry\" and PO.\"BaseLine\" = P1.\"LineNum\" " +
                "Where P0.\"CANCELED\" = 'N' and M.\"U_PdoEntry\" in (" + InList(ids) + ") " +
                "Group by M.\"U_PdoEntry\", M.\"U_PdoLine\""))
                result[Int(r, "Pdo") + ":" + Int(r, "Line")] = Dec(r, "Share");
            if (!HasItrMap) return result;
            // Inventory Transfer Request lines count the same way: still open on the request + already transferred
            // (Inventory Transfers drawn from it), capped at the line quantity, x the component's share of the line
            string itrCovered =
                "(Case When T1.\"LineStatus\" = 'O' and T0.\"DocStatus\" = 'O' Then T1.\"OpenQty\" Else 0 End + IsNull(TR.\"Moved\", 0))";
            foreach (var r in Db.Rows(
                "Select M.\"U_PdoEntry\" as \"Pdo\", M.\"U_PdoLine\" as \"Line\", " +
                "Sum(M.\"U_Qty\" * ((Case When " + itrCovered + " > T1.\"Quantity\" Then T1.\"Quantity\" Else " + itrCovered + " End) / NullIf(T1.\"Quantity\", 0))) as \"Share\" " +
                "from " + ItrMapTable + " M " +
                "Inner Join WTQ1 T1 on T1.\"DocEntry\" = M.\"U_ItrEntry\" and T1.\"LineNum\" = M.\"U_ItrLine\" " +
                "Inner Join OWTQ T0 on T0.\"DocEntry\" = T1.\"DocEntry\" " +
                "Left Join (Select Q.\"BaseEntry\", Q.\"BaseLine\", Sum(Q.\"Quantity\") as \"Moved\" from WTR1 Q " +
                "Inner Join OWTR O on O.\"DocEntry\" = Q.\"DocEntry\" " +
                "Where Q.\"BaseType\" = " + ObjInventoryTransferRequest + " and O.\"CANCELED\" = 'N' Group by Q.\"BaseEntry\", Q.\"BaseLine\") TR " +
                "on TR.\"BaseEntry\" = T1.\"DocEntry\" and TR.\"BaseLine\" = T1.\"LineNum\" " +
                "Where T0.\"CANCELED\" = 'N' and M.\"U_PdoEntry\" in (" + InList(ids) + ") " +
                "Group by M.\"U_PdoEntry\", M.\"U_PdoLine\""))
            {
                string key = Int(r, "Pdo") + ":" + Int(r, "Line");
                decimal had;
                result.TryGetValue(key, out had);
                result[key] = had + Dec(r, "Share");
            }
            return result;
        }
        #endregion

        #region Inventory Transfer Request component map (@EBH_ITRMAP)
        /// <summary>Same idea as @EBH_PRQMAP, for the Inventory Transfer Request lines Material Request Planning raises
        /// (ITRQTY): one row per Production Order component each ITR line covers, with that component's quantity.</summary>
        private const string ItrMapTable = "\"@EBH_ITRMAP\"";

        private bool hasItrMap;
        /// <summary>True once Database Setup has created @EBH_ITRMAP (checked until it exists, then remembered), so
        /// every lookup keeps working on a company where the setup has not been run again yet.</summary>
        private bool HasItrMap
        {
            get
            {
                if (!hasItrMap) hasItrMap = UdfExists("EBH_ITRMAP", "ItrStatus");
                return hasItrMap;
            }
        }

        private static void InsertItrMapRow(int itrEntry, int itrLine, int pdoEntry, int pdoLine, decimal qty)
        {
            string code = Guid.NewGuid().ToString("N");
            Db.Exec("Insert into " + ItrMapTable + " (\"Code\",\"Name\",\"U_ItrEntry\",\"U_ItrLine\",\"U_PdoEntry\",\"U_PdoLine\",\"U_Qty\",\"U_ItrStatus\") Values (" +
                Db.Quote(code) + "," + Db.Quote(code) + "," + itrEntry + "," + itrLine + "," + pdoEntry + "," + pdoLine + "," + qty.ToString(Inv) + "," + Db.Quote("O") + ")");
        }

        /// <summary>Refreshes U_ItrStatus ('O' / 'C' / 'X' cancelled) from the live ITR document, for the given
        /// Production Orders ("" = every row). Reference only - lookups read the live status through the join.</summary>
        private void SyncItrMapStatus(string where)
        {
            if (!HasItrMap) return;
            Db.Exec(
                "Update " + ItrMapTable + " Set \"U_ItrStatus\" = " +
                "(Select Case When T0.\"CANCELED\" = 'Y' Then 'X' Else T0.\"DocStatus\" End from OWTQ T0 " +
                "Where T0.\"DocEntry\" = " + ItrMapTable + ".\"U_ItrEntry\")" + where);
        }
        #endregion

        #region Linked documents
        /// <summary>Production Order DocEntries (of the given ones) that already have at least one Purchase Request line raised from them (any status) - what the Linked Documents button would show something for.</summary>
        public List<int> OrdersWithLinkedRequests(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.Distinct().ToList();
            if (ids.Count == 0) return new List<int>();
            string sql = "Select Distinct M.\"U_PdoEntry\" as \"Pdo\" from " + MapTable + " M Where M.\"U_PdoEntry\" in (" + InList(ids) + ")";
            if (HasItrMap) sql += " Union Select Distinct M.\"U_PdoEntry\" as \"Pdo\" from " + ItrMapTable + " M Where M.\"U_PdoEntry\" in (" + InList(ids) + ")";
            return Db.Rows(sql).Select(r => Int(r, "Pdo")).Distinct().ToList();
        }

        /// <summary>Purchase Request lines raised from a Production Order, with the request status - including a
        /// line merged from this component and others, reported with this component's own share of quantity, not
        /// the whole merged line's.</summary>
        public List<LinkedRequest> LinkedRequests(int pdoEntry)
        {
            List<LinkedRequest> list = LinkedPurchaseRequests(pdoEntry);
            if (HasItrMap) list.AddRange(LinkedTransferRequests(pdoEntry));
            return list;
        }

        /// <summary>Inventory Transfer Request lines raised from a Production Order (DocType "ITR"), this component's own share.</summary>
        private List<LinkedRequest> LinkedTransferRequests(int pdoEntry)
        {
            return Db.Rows(
                "Select T0.\"DocEntry\", T0.\"DocNum\", T0.\"Series\", N1.\"SeriesName\", T0.\"DocDate\", T0.\"DocStatus\", T0.\"CANCELED\", " +
                "T1.\"LineNum\" as \"ItrLine\", T1.\"LineStatus\", T1.\"ItemCode\", T1.\"Dscription\", T1.\"Quantity\", T1.\"OpenQty\", " +
                "M.\"U_PdoLine\" as \"PdoLine\", M.\"U_Qty\" as \"MemberQty\" " +
                "from " + ItrMapTable + " M " +
                "Inner Join WTQ1 T1 on T1.\"DocEntry\" = M.\"U_ItrEntry\" and T1.\"LineNum\" = M.\"U_ItrLine\" " +
                "Inner Join OWTQ T0 on T0.\"DocEntry\" = T1.\"DocEntry\" " +
                "Left Join NNM1 N1 on N1.\"Series\" = T0.\"Series\" and N1.\"ObjectCode\" = '" + ObjInventoryTransferRequest + "' " +
                "Where M.\"U_PdoEntry\" = " + pdoEntry + " Order by T0.\"DocNum\", M.\"U_PdoLine\"")
                .Select(r => new LinkedRequest
                {
                    DocType = "ITR",
                    PrEntry = Int(r, "DocEntry"), PrNum = Int(r, "DocNum"), PrLine = Int(r, "ItrLine"),
                    Series = Int(r, "Series"), SeriesName = Db.Str(r, "SeriesName"), PostingDate = Date(r, "DocDate"),
                    Status = Db.Str(r, "CANCELED") == "Y" ? "X" : Db.Str(r, "DocStatus"),
                    ItemCode = Db.Str(r, "ItemCode"), Description = Db.Str(r, "Dscription"), PdoLine = Int(r, "PdoLine"),
                    Quantity = Dec(r, "MemberQty"),
                    OpenQuantity = Db.Str(r, "CANCELED") == "N" && Db.Str(r, "LineStatus") == "O" && Dec(r, "Quantity") > 0
                        ? Dec(r, "OpenQty") * (Dec(r, "MemberQty") / Dec(r, "Quantity")) : 0
                }).ToList();
        }

        private List<LinkedRequest> LinkedPurchaseRequests(int pdoEntry)
        {
            return Db.Rows(
                "Select P0.\"DocEntry\", P0.\"DocNum\", P0.\"Series\", N1.\"SeriesName\", P0.\"DocDate\", P0.\"DocStatus\", P0.\"CANCELED\", " +
                "P1.\"LineNum\" as \"PrLine\", P1.\"LineStatus\", P1.\"ItemCode\", P1.\"Dscription\", P1.\"Quantity\", P1.\"OpenQty\", " +
                "M.\"U_PdoLine\" as \"PdoLine\", M.\"U_Qty\" as \"MemberQty\" " +
                "from " + MapTable + " M " +
                "Inner Join PRQ1 P1 on P1.\"DocEntry\" = M.\"U_PrEntry\" and P1.\"LineNum\" = M.\"U_PrLine\" " +
                "Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Left Join NNM1 N1 on N1.\"Series\" = P0.\"Series\" and N1.\"ObjectCode\" = '" + ObjPurchaseRequest + "' " +
                "Where M.\"U_PdoEntry\" = " + pdoEntry + " Order by P0.\"DocNum\", M.\"U_PdoLine\"")
                .Select(r => new LinkedRequest
                {
                    PrEntry = Int(r, "DocEntry"), PrNum = Int(r, "DocNum"), PrLine = Int(r, "PrLine"),
                    Series = Int(r, "Series"), SeriesName = Db.Str(r, "SeriesName"), PostingDate = Date(r, "DocDate"),
                    Status = Db.Str(r, "CANCELED") == "Y" ? "X" : Db.Str(r, "DocStatus"),
                    ItemCode = Db.Str(r, "ItemCode"), Description = Db.Str(r, "Dscription"), PdoLine = Int(r, "PdoLine"),
                    Quantity = Dec(r, "MemberQty"),
                    OpenQuantity = Db.Str(r, "CANCELED") == "N" && Db.Str(r, "LineStatus") == "O" && Dec(r, "Quantity") > 0
                        ? Dec(r, "OpenQty") * (Dec(r, "MemberQty") / Dec(r, "Quantity")) : 0
                }).ToList();
        }

        public int PdoNum(int pdoEntry)
        {
            return Int(Db.FirstRow("Select \"DocNum\" from OWOR Where \"DocEntry\" = " + pdoEntry), "DocNum");
        }

        public void OpenPurchaseRequest(int docEntry)
        {
            Program.SBO_App.OpenForm((SAPbouiCOM.BoFormObjectEnum)ObjPurchaseRequest, "", docEntry.ToString(Inv));
        }

        public void OpenTransferRequest(int docEntry)
        {
            Program.SBO_App.OpenForm((SAPbouiCOM.BoFormObjectEnum)ObjInventoryTransferRequest, "", docEntry.ToString(Inv));
        }

        public void OpenProductionOrder(int docEntry)
        {
            Program.SBO_App.OpenForm((SAPbouiCOM.BoFormObjectEnum)ObjProductionOrder, "", docEntry.ToString(Inv));
        }

        /// <summary>Opens Item Master Data (object type 4) at the given item - keyed by ItemCode, not a DocEntry.</summary>
        public void OpenItem(string itemCode)
        {
            if (itemCode.Length == 0) return;
            Program.SBO_App.OpenForm((SAPbouiCOM.BoFormObjectEnum)4, "", itemCode);
        }
        #endregion

        #region Purchase Requests (for Purchase Order Planning)
        /// <summary>Project as shown on the Purchase Order Planning form: the header value (OPRQ.Project) when set, else the first line that carries one.</summary>
        private const string ProjectExpr = "IsNull(NullIf(P0.\"Project\",''), (Select Top 1 LP.\"Project\" from PRQ1 LP Where LP.\"DocEntry\" = P0.\"DocEntry\" and LP.\"Project\" is not null and LP.\"Project\" <> ''))";

        private string vendorColumn;
        /// <summary>
        /// The actual PRQ1 column behind the DI API's Document_Lines.LineVendor (what CreatePurchaseRequest writes via
        /// oDoc.Lines.LineVendor). Detected once per session instead of hard-coded, because that DB column name is not
        /// documented and was found to be wrong once already (it is not "VendorNum" on every SAP Business One setup).
        /// Read from the table's schema, not from a data row, so it works while PRQ1 is still empty.
        /// </summary>
        private string VendorColumn()
        {
            if (vendorColumn != null) return vendorColumn;
            var candidates = new[] { "LineVendor", "VendorNum", "CardCode" };
            var present = new HashSet<string>(
                Db.Rows("Select \"COLUMN_NAME\" from INFORMATION_SCHEMA.COLUMNS Where \"TABLE_NAME\" = 'PRQ1' and \"COLUMN_NAME\" in (" + InList(candidates) + ")")
                    .Select(r => Db.Str(r, "COLUMN_NAME")), StringComparer.OrdinalIgnoreCase);
            foreach (string c in candidates)
                if (present.Contains(c)) return vendorColumn = c;
            throw new PlanException(500, "Could not find the line Vendor field on the Purchase Request row (PRQ1). Looked for LineVendor, VendorNum, CardCode.");
        }

        /// <summary>Manufacturer code of the item aliased I ('' when it has none) - shared by the queries that list items.</summary>
        private const string ManufacturerExpr = "Case When I.\"FirmCode\" is null or I.\"FirmCode\" = -1 Then '' Else Cast(I.\"FirmCode\" as nvarchar(20)) End";

        /// <summary>The Purchase Request lines matching <paramref name="where"/> (P0 = OPRQ, P1 = PRQ1), as listed on the
        /// Purchase Order Planning picker grid - shared by ListRequestLines and RefreshRequestLines.</summary>
        private List<PurchaseRequestLine> QueryRequestLines(string top, string where, string orderBy)
        {
            return Db.Rows("Select " + top + "P0.\"DocEntry\", P0.\"DocNum\", P0.\"Series\", N1.\"SeriesName\", P0.\"DocDate\", P0.\"DocStatus\", " +
                "P1.\"LineNum\", P1.\"LineStatus\", P1.\"ItemCode\", P1.\"Dscription\", P1.\"Quantity\", P1.\"OpenQty\", P1.\"Price\", P1.\"DiscPrcnt\", P1.\"WhsCode\", P1.\"TaxCode\", " +
                "IsNull(NullIf(P1.\"Project\",''), P0.\"Project\") as \"Project\", P1.\"" + VendorColumn() + "\" as \"LineVendor\", " +
                ManufacturerExpr + " as \"Manufacturer\", " +
                "IsNull(M.\"FirmName\", '') as \"ManufacturerName\", IsNull(I.\"U_EBH_MfgCatlog\", '') as \"Catalog\", " +
                "IsNull(P1.\"U_EBH_JOBCODNO\", '') as \"JobCode\" " +
                "from PRQ1 P1 Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Left Join NNM1 N1 on N1.\"Series\" = P0.\"Series\" and N1.\"ObjectCode\" = '" + ObjPurchaseRequest + "' " +
                "Left Join OITM I on I.\"ItemCode\" = P1.\"ItemCode\" " +
                "Left Join OMRC M on M.\"FirmCode\" = I.\"FirmCode\" " +
                "Where " + where + orderBy)
                .Select(r => new PurchaseRequestLine
                {
                    PrEntry = Int(r, "DocEntry"), PrNum = Int(r, "DocNum"), PrLine = Int(r, "LineNum"),
                    Series = Int(r, "Series"), SeriesName = Db.Str(r, "SeriesName"), Status = Db.Str(r, "DocStatus"), LineStatus = Db.Str(r, "LineStatus"), PostingDate = Date(r, "DocDate"),
                    ItemCode = Db.Str(r, "ItemCode"), Description = Db.Str(r, "Dscription"), Quantity = Dec(r, "Quantity"), OpenQuantity = Dec(r, "OpenQty"),
                    Price = Dec(r, "Price"), Discount = Dec(r, "DiscPrcnt"), Warehouse = Db.Str(r, "WhsCode"), TaxCode = Db.Str(r, "TaxCode"),
                    Project = Db.Str(r, "Project"), Vendor = Db.Str(r, "LineVendor"),
                    Manufacturer = Db.Str(r, "Manufacturer"), ManufacturerName = Db.Str(r, "ManufacturerName"), Catalog = Db.Str(r, "Catalog"),
                    JobCode = Db.Str(r, "JobCode")
                }).ToList();
        }

        /// <summary>
        /// One row per eligible Purchase Request line, optionally narrowed by Project, Vendor and/or Purchase
        /// Request number - a request with lines of more than one Project / Vendor reports each eligible line
        /// separately rather than one row standing in for the whole document. A closed line (document closed, line
        /// closed, or nothing left open on it - PurchaseRequestLine.Closed) is included too, not just a still-open
        /// one: the picker grid shows it locked (green, unticked - see PurchaseOrderForm.ShowRequests) rather than
        /// silently dropping it. Only a cancelled document is left out.
        /// <paramref name="find"/> takes Purchase Request numbers only (comma / semicolon separated); anything else is refused.
        /// </summary>
        public List<PurchaseRequestLine> ListRequestLines(IEnumerable<string> projects, IEnumerable<string> vendors, string find)
        {
            string where = "P0.\"CANCELED\" = 'N'";
            var prj = projects.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
            if (prj.Count > 0)
                where += " and (P1.\"Project\" in (" + InList(prj) + ") or P0.\"Project\" in (" + InList(prj) + "))";
            var vnd = vendors.Where(v => !string.IsNullOrWhiteSpace(v)).ToList();
            if (vnd.Count > 0)
                where += " and P1.\"" + VendorColumn() + "\" in (" + InList(vnd) + ")";
            var nums = new List<int>();
            // Several Purchase Request numbers chosen from the Choose From List (comma / semicolon separated)
            foreach (string t in (find ?? "").Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim()).Where(t => t.Length > 0))
            {
                int n;
                if (!int.TryParse(t, NumberStyles.Integer, Inv, out n))
                    throw new PlanException(400, "\"" + t + "\" is not a Purchase Request number. Enter numbers only, separated by commas.");
                nums.Add(n);
            }
            if (nums.Count > 0) where += " and P0.\"DocNum\" in (" + InList(nums) + ")";
            // The 1000-row cap only guards an unfiltered "show everything" search; a Project / Vendor / number search
            // must return every eligible line (a line can otherwise be ticked for loading without ever appearing here).
            bool capped = prj.Count == 0 && vnd.Count == 0 && nums.Count == 0;
            return QueryRequestLines(capped ? "Top 1000 " : "", where, " Order by P0.\"DocNum\" Desc, P1.\"LineNum\"");
        }

        /// <summary>
        /// The exact given Purchase Request lines (DocEntry, LineNum), re-read fresh from PRQ1 regardless of their
        /// current status - for refreshing lines the picker grid already shows after drawing a Purchase Order from
        /// some of them: a line that closed as a result still needs to be shown (locked, in green), not silently dropped.
        /// </summary>
        public List<PurchaseRequestLine> RefreshRequestLines(IEnumerable<KeyValuePair<int, int>> keys)
        {
            var list = keys.Distinct().ToList();
            if (list.Count == 0) return new List<PurchaseRequestLine>();
            string keyWhere = string.Join(" or ", list.Select(k => "(P0.\"DocEntry\" = " + k.Key.ToString(Inv) + " and P1.\"LineNum\" = " + k.Value.ToString(Inv) + ")").ToArray());
            return QueryRequestLines("", "(" + keyWhere + ")", "");
        }

        /// <summary>Distinct, non-blank Vendor codes of every line of the given Purchase Request numbers - a request with lines of more than one Vendor reports all of them.</summary>
        public List<string> VendorsOfRequests(IEnumerable<int> docNums)
        {
            var nums = docNums.Distinct().ToList();
            if (nums.Count == 0) return new List<string>();
            return Db.Rows("Select Distinct L.\"" + VendorColumn() + "\" as \"Vendor\" from PRQ1 L Inner Join OPRQ P0 on P0.\"DocEntry\" = L.\"DocEntry\" " +
                "Where P0.\"DocNum\" in (" + InList(nums) + ")")
                .Select(r => Db.Str(r, "Vendor")).Where(v => v.Length > 0).ToList();
        }

        /// <summary>Distinct, non-blank effective Project codes (header, falling back to a line's own Project) of the given Purchase Request numbers.</summary>
        public List<string> ProjectsOfRequests(IEnumerable<int> docNums)
        {
            var nums = docNums.Distinct().ToList();
            if (nums.Count == 0) return new List<string>();
            return Db.Rows("Select Distinct " + ProjectExpr + " as \"Project\" from OPRQ P0 Where P0.\"DocNum\" in (" + InList(nums) + ")")
                .Select(r => Db.Str(r, "Project")).Where(p => p.Length > 0).ToList();
        }

        /// <summary>Distinct Vendor codes of every still-open, still-orderable line of the open Purchase Requests of the given Projects (to narrow the Vendor Choose From List / auto-initialise Vendor).</summary>
        public List<string> VendorsOfProjects(IEnumerable<string> projects)
        {
            var prj = projects.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
            if (prj.Count == 0) return new List<string>();
            return Db.Rows("Select Distinct L.\"" + VendorColumn() + "\" as \"Vendor\" from PRQ1 L Inner Join OPRQ P0 on P0.\"DocEntry\" = L.\"DocEntry\" " +
                "Where P0.\"DocStatus\" = 'O' and P0.\"CANCELED\" = 'N' and L.\"LineStatus\" = 'O' and L.\"OpenQty\" > 0 and " +
                "(P0.\"Project\" in (" + InList(prj) + ") or L.\"Project\" in (" + InList(prj) + "))")
                .Select(r => Db.Str(r, "Vendor")).Where(v => v.Length > 0).ToList();
        }

        /// <summary>Distinct Project codes of open Purchase Requests with a still-open, still-orderable line of the given Vendors (to narrow the Project Choose From List / auto-initialise Project).</summary>
        public List<string> ProjectsOfVendors(IEnumerable<string> vendors)
        {
            var vnd = vendors.Where(v => !string.IsNullOrWhiteSpace(v)).ToList();
            if (vnd.Count == 0) return new List<string>();
            return Db.Rows("Select Distinct " + ProjectExpr + " as \"Project\" from OPRQ P0 " +
                "Where P0.\"DocStatus\" = 'O' and P0.\"CANCELED\" = 'N' and " +
                "Exists (Select 1 from PRQ1 L Where L.\"DocEntry\" = P0.\"DocEntry\" and L.\"LineStatus\" = 'O' and L.\"OpenQty\" > 0 and L.\"" + VendorColumn() + "\" in (" + InList(vnd) + "))")
                .Select(r => Db.Str(r, "Project")).Where(p => p.Length > 0).ToList();
        }

        /// <summary>
        /// The exact Purchase Request lines (DocEntry, LineNum) ticked on the picker grid, re-read fresh from PRQ1 so
        /// price / open quantity reflect what SAP has right now rather than what the grid last showed, to load into
        /// the Purchase Order grid. A line no longer open (closed / cancelled / fully ordered meanwhile) is silently
        /// left out rather than loaded with stale data.
        /// </summary>
        public List<PoRow> GetRequestLines(IEnumerable<KeyValuePair<int, int>> keys)
        {
            var list = keys.Distinct().ToList();
            if (list.Count == 0) return new List<PoRow>();
            string col = VendorColumn();
            string keyWhere = string.Join(" or ", list.Select(k => "(P0.\"DocEntry\" = " + k.Key.ToString(Inv) + " and P1.\"LineNum\" = " + k.Value.ToString(Inv) + ")").ToArray());
            return Db.Rows(
                "Select P0.\"DocEntry\", P0.\"DocNum\", P1.\"LineNum\", P1.\"ItemCode\", P1.\"Dscription\", P1.\"OpenQty\", P1.\"Price\", P1.\"DiscPrcnt\", " +
                "P1.\"WhsCode\", P1.\"Project\", P1.\"TaxCode\", P1.\"" + col + "\" as \"LineVendor\", " +
                "IsNull(P1.\"U_EBH_CATLOG\", '') as \"Catalog\", IsNull(P1.\"U_EBH_JOBCODNO\", '') as \"JobCode\", " +
                "IsNull(P1.\"" + FieldPafd + "\", 0) as \"PrPafd\", " +
                ManufacturerExpr + " as \"Manufacturer\", " +
                "IsNull(M.\"FirmName\", '') as \"ManufacturerName\", " +
                "W.\"WPrice\", W.\"WDisc\", W.\"WPafdi\" " +
                "from PRQ1 P1 Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Left Join OITM I on I.\"ItemCode\" = P1.\"ItemCode\" " +
                "Left Join OMRC M on M.\"FirmCode\" = I.\"FirmCode\" " +
                // Every line of the same item on the Production Order(s) this request line was raised from (@EBH_PRQMAP),
                // not only the mapped component itself: the lowest priced Price After Discount among them supplies the
                // line's price and limit; an unpriced (0) line never wins.
                "Outer Apply (Select Top 1 IsNull(T1.\"U_EBH_Unit_Price\", 0) as \"WPrice\", IsNull(T1.\"U_EBH_Discount\", 0) as \"WDisc\", " +
                PafdiExpr("T1") + " as \"WPafdi\" " +
                "from WOR1 T1 Where T1.\"ItemCode\" = P1.\"ItemCode\" and T1.\"DocEntry\" in (Select PM.\"U_PdoEntry\" from " + MapTable + " PM " +
                "Where PM.\"U_PrEntry\" = P1.\"DocEntry\" and PM.\"U_PrLine\" = P1.\"LineNum\") and " + PafdiExpr("T1") + " > 0 " +
                "Order by " + PafdiExpr("T1") + ") W " +
                "Where P1.\"LineStatus\" = 'O' and P1.\"OpenQty\" > 0 and (" + keyWhere + ") Order by P0.\"DocNum\", P1.\"LineNum\"")
                .Select(r =>
                {
                    // The cap: the lower of the planned Price After Discount stored on the request line and the lowest
                    // priced same-item line of its Production Order(s) - so a request raised at a higher price before the
                    // lowest-price rule is still held to it; else (not raised from a Production Order) the line's own price
                    decimal prPafd = Math.Round(Dec(r, "PrPafd"), 6), wPafdi = Math.Round(Dec(r, "WPafdi"), 6);
                    decimal ownPafd = PlanService.AfterDiscount(Dec(r, "Price"), Dec(r, "DiscPrcnt"));
                    decimal cap = prPafd > 0 && wPafdi > 0 ? Math.Min(prPafd, wPafdi) : prPafd > 0 ? prPafd : wPafdi > 0 ? wPafdi : ownPafd;
                    // Price / Discount % start blank: the user enters them on the grid, checked against the cap
                    return new PoRow
                    {
                        PrEntry = Int(r, "DocEntry"), PrNum = Int(r, "DocNum"), PrLine = Int(r, "LineNum"),
                        ItemCode = Db.Str(r, "ItemCode"), Description = Db.Str(r, "Dscription"), OpenQuantity = Dec(r, "OpenQty"), Quantity = Dec(r, "OpenQty"),
                        MaxPriceAfterDiscount = cap,
                        Warehouse = Db.Str(r, "WhsCode"), Project = Db.Str(r, "Project"),
                        Vendor = Db.Str(r, "LineVendor"), TaxCode = Db.Str(r, "TaxCode"), Catalog = Db.Str(r, "Catalog"), JobCode = Db.Str(r, "JobCode"),
                        Manufacturer = Db.Str(r, "Manufacturer"), ManufacturerName = Db.Str(r, "ManufacturerName")
                    };
                }).ToList();
        }

        public void OpenPurchaseOrder(int docEntry)
        {
            Program.SBO_App.OpenForm((SAPbouiCOM.BoFormObjectEnum)ObjPurchaseOrder, "", docEntry.ToString(Inv));
        }
        #endregion

        #region Create Purchase Order
        /// <summary>
        /// Adds one Purchase Order drawn from the selected Purchase Request lines (BaseType / BaseEntry / BaseLine), so
        /// SAP keeps the standard document trail and closes / reduces the open quantity of the source request lines itself.
        /// The document is added completely or not at all (one DI call inside a transaction).
        /// </summary>
        public CreatedOrder CreatePurchaseOrder(PoHeader h, IList<PoRow> rows)
        {
            if (!VendorExists(h.Vendor)) throw new PlanException(400, "Vendor " + h.Vendor + " is not a valid supplier.", Field("eVend"));
            bool multiBranch = MultiBranch;
            if (multiBranch && h.Branch <= 0) throw new PlanException(400, "Select the Branch.", Field("cBranch"));

            SAPbobsCOM.Documents oDoc = null;
            bool inTransaction = false;
            var result = new CreatedOrder();
            try
            {
                if (Company.InTransaction) throw new PlanException(409, "SAP is busy with another transaction. Try again in a moment.");
                Company.StartTransaction();
                inTransaction = true;

                oDoc = (SAPbobsCOM.Documents)Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                oDoc.Series = h.Series;
                // The branch property is declared on IDocuments, not on the Documents interface GetBusinessObject returns
                if (multiBranch) ((SAPbobsCOM.IDocuments)oDoc).BPL_IDAssignedToInvoice = h.Branch;
                oDoc.CardCode = h.Vendor;
                oDoc.DocDate = DateTime.Today;
                oDoc.DocDueDate = ToDate(h.DueDate);
                if (h.VendorRef.Length > 0) oDoc.NumAtCard = h.VendorRef;
                SetRemark(oDoc.UserFields, "Created by EBH Material Request Planning from Purchase Request " + string.Join(", ", rows.Select(r => r.PrNum).Distinct().ToArray()));

                for (int i = 0; i < rows.Count; i++)
                {
                    PoRow r = rows[i];
                    if (i > 0) oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(i);
                    oDoc.Lines.BaseType = ObjPurchaseRequest;
                    oDoc.Lines.BaseEntry = r.PrEntry;
                    oDoc.Lines.BaseLine = r.PrLine;
                    oDoc.Lines.Quantity = (double)r.Quantity;
                    oDoc.Lines.WarehouseCode = r.Warehouse;
                    oDoc.Lines.UnitPrice = (double)r.Price;
                    oDoc.Lines.DiscountPercent = (double)r.Discount;
                    oDoc.Lines.UserFields.Fields.Item(FieldPafd).Value = (double)r.PriceAfterDiscount;
                    oDoc.Lines.UserFields.Fields.Item(FieldLineTotal).Value = (double)r.Total;
                    if (r.Project.Length > 0) ((SAPbobsCOM.IDocument_Lines)oDoc.Lines).ProjectCode = r.Project;
                    // Overriding Warehouse / Price here stops SAP from copying the Tax Code from the base Purchase Request line on its own
                    // (needed for India GST: interstate vs intrastate - IGST vs CGST/SGST - is decided per line from this code)
                    if (r.TaxCode.Length > 0) ((SAPbobsCOM.IDocument_Lines)oDoc.Lines).TaxCode = r.TaxCode;
                    if (r.Catalog.Length > 0) oDoc.Lines.UserFields.Fields.Item("U_EBH_CATLOG").Value = r.Catalog;
                    if (r.JobCode.Length > 0) oDoc.Lines.UserFields.Fields.Item("U_EBH_JOBCODNO").Value = r.JobCode;
                }
                Check(oDoc.Add(), "Purchase Order was not created");
                result.DocEntry = int.Parse(Company.GetNewObjectKey(), Inv);
                Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                inTransaction = false;
            }
            catch
            {
                if (inTransaction && Company.InTransaction) Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                throw;
            }
            finally { Release(oDoc); }

            result.DocNum = Int(Db.FirstRow("Select \"DocNum\" from OPOR Where \"DocEntry\" = " + result.DocEntry), "DocNum");
            return result;
        }
        #endregion

        #region Create Purchase Request / Inventory Transfer Request
        /// <summary>
        /// Adds the Purchase Request (one line per row of <paramref name="prRows"/>, quantity = PRQTY) and/or the
        /// Inventory Transfer Request (one line per row of <paramref name="itrRows"/>, quantity = ITRQTY) in ONE SAP
        /// transaction: both are added, or neither. Either list may be empty (Purchase Request or ITR button alone).
        /// Afterwards the Production Order components are refreshed; that part is repeatable (Sync), so a failure there
        /// is reported as a warning and does not undo the documents.
        /// </summary>
        public CreatedDocuments CreateDocuments(PlanHeader h, IList<PlanRow> prRows, IList<PlanRow> itrRows)
        {
            prRows = prRows ?? new List<PlanRow>();
            itrRows = itrRows ?? new List<PlanRow>();
            if (prRows.Count == 0 && itrRows.Count == 0) throw new PlanException(400, "Nothing to create.");
            bool multiBranch = MultiBranch;
            if (prRows.Count > 0)
            {
                // Vendor is optional: a Purchase Request line can be raised with no preferred Vendor at all, and each row
                // may carry its own anyway (LineVendor). Only check it when one was actually typed / applied.
                if (h.Vendor.Length > 0 && !VendorExists(h.Vendor)) throw new PlanException(400, "Vendor " + h.Vendor + " is not a valid supplier.", Field("eVend"));
                // One lookup per distinct row Vendor, not per row
                foreach (string v in prRows.Select(r => r.Vendor).Where(v => v.Length > 0 && !string.Equals(v, h.Vendor, StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase))
                    if (!VendorExists(v))
                        throw new PlanException(400, "Vendor " + v + " (row for " + prRows.First(r => string.Equals(r.Vendor, v, StringComparison.OrdinalIgnoreCase)).ItemCode + ") is not a valid supplier.", "gRows");
                if (!WarehouseExists(h.Warehouse)) throw new PlanException(400, "Warehouse " + h.Warehouse + " does not exist.", Field("eWhs"));
                if (multiBranch && h.Branch <= 0) throw new PlanException(400, "Select the Branch.", Field("cBranch"));
            }
            int itrSeries = 0;
            if (itrRows.Count > 0)
            {
                // The ITR writes WTQ1.U_EBH_PRODUCTIONNO and @EBH_ITRMAP: stop with a clear message, not a raw SAP /
                // SQL error after the document was already added and rolled back
                if (!HasItrMap || !UdfExists("WTQ1", "EBH_PRODUCTIONNO"))
                    throw new PlanException(500, "Inventory Transfer Requests are not set up in this company yet. Run Database Setup (Create All Environment), then restart SAP Business One.");
                List<string> errors = CheckTransferRows(itrRows, r => r.ItemCode);
                if (errors.Count > 0) throw new PlanException(400, string.Join(Environment.NewLine, errors.Take(8).ToArray()), "gRows");
                itrSeries = DefaultSeriesFor(ObjInventoryTransferRequest);
            }

            var result = new CreatedDocuments();
            bool inTransaction = false;
            try
            {
                if (Company.InTransaction) throw new PlanException(409, "SAP is busy with another transaction. Try again in a moment.");
                Company.StartTransaction();
                inTransaction = true;
                if (prRows.Count > 0) result.Pr = new CreatedRequest { DocEntry = AddPurchaseRequest(h, prRows, multiBranch) };
                if (itrRows.Count > 0) result.Itr = new CreatedRequest { DocEntry = AddTransferRequest(h, itrRows, itrSeries) };
                Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                inTransaction = false;
            }
            catch
            {
                if (inTransaction && Company.InTransaction) Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                throw;
            }

            if (result.Pr != null) result.Pr.DocNum = Int(Db.FirstRow("Select \"DocNum\" from OPRQ Where \"DocEntry\" = " + result.Pr.DocEntry), "DocNum");
            if (result.Itr != null) result.Itr.DocNum = Int(Db.FirstRow("Select \"DocNum\" from OWTQ Where \"DocEntry\" = " + result.Itr.DocEntry), "DocNum");
            var allPdoEntries = Members(prRows).Concat(Members(itrRows)).Select(m => m.PdoEntry).Distinct();
            try { Sync(allPdoEntries); }
            catch (Exception ex)
            {
                result.Warning = "The document(s) were created, but the Production Order could not be refreshed (" + ex.Message + "). It is refreshed when the Production Order's linked documents are opened.";
            }
            return result;
        }

        /// <summary>The source component rows a grid row stands for: its merged members, or the row itself.</summary>
        private static IEnumerable<PlanRow> Members(IEnumerable<PlanRow> rows)
        {
            return rows.SelectMany(r => r.MergedFrom != null && r.MergedFrom.Count > 0 ? (IEnumerable<PlanRow>)r.MergedFrom : new[] { r });
        }

        /// <summary>Builds and adds the Purchase Request (quantity = PRQTY) plus its @EBH_PRQMAP rows. Runs inside
        /// CreateDocuments' transaction; returns the new DocEntry.</summary>
        private int AddPurchaseRequest(PlanHeader h, IList<PlanRow> rows, bool multiBranch)
        {
            SAPbobsCOM.Documents oDoc = null;
            try
            {
                oDoc = (SAPbobsCOM.Documents)Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);
                oDoc.Series = h.Series;
                // The branch property is declared on IDocuments, not on the Documents interface GetBusinessObject returns
                if (multiBranch) ((SAPbobsCOM.IDocuments)oDoc).BPL_IDAssignedToInvoice = h.Branch;
                oDoc.DocDate = DateTime.Today;
                oDoc.DocDueDate = ToDate(h.ValidUntil);
                oDoc.RequriedDate = ToDate(h.RequiredDate);
                oDoc.ReqType = 12;                                      // requester is a user
                oDoc.Requester = Company.UserName;
                SetRemark(oDoc.UserFields, "Created by EBH Material Request Planning from Production Order " + string.Join(", ", Members(rows).Select(m => m.PdoNum).Distinct().ToArray()));
                // Vendor and Warehouse are not kept on the header (OPRQ.U_EBH_Vendor / U_EBH_Whse): each line carries its own (LineVendor / WarehouseCode), read back for Purchase Order Planning.
                // Header Project (OPRQ.Project): the rows all share one Project, enforced when they were picked on the form
                string headerProject = rows.Select(r => r.Project).FirstOrDefault(p => p.Length > 0) ?? "";
                if (headerProject.Length > 0) ((SAPbobsCOM.IDocuments)oDoc).Project = headerProject;

                for (int i = 0; i < rows.Count; i++)
                {
                    PlanRow r = rows[i];
                    if (i > 0) oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(i);
                    oDoc.Lines.ItemCode = r.ItemCode;
                    oDoc.Lines.Quantity = (double)r.PrQuantity;
                    oDoc.Lines.WarehouseCode = r.Warehouse;
                    oDoc.Lines.LineVendor = r.Vendor;
                    // Project is declared on IDocument_Lines, not on the Document_Lines interface Lines returns
                    if (r.Project.Length > 0) ((SAPbobsCOM.IDocument_Lines)oDoc.Lines).ProjectCode = r.Project;
                    oDoc.Lines.RequiredDate = ToDate(r.RequiredDate);
                    // The planned price goes to the UDFs only: the standard Price / Discount % / Price After Discount /
                    // Line Total of the request line are deliberately left blank (UnitPrice / DiscountPercent never set)
                    oDoc.Lines.UserFields.Fields.Item(FieldPafd).Value = (double)r.PriceAfterDiscount;
                    oDoc.Lines.UserFields.Fields.Item(FieldLineTotal).Value = (double)r.Total;
                    var members = Members(new[] { r }).ToList();
                    oDoc.Lines.UserFields.Fields.Item(FieldPdo).Value = r.PdoEntry.ToString(Inv);
                    // Comma-separated line number(s) of the members that belong to the SAME order as EBH_PRODUCTIONNO
                    // above (one number for an ordinary, non-merged row). A member from a different order is still
                    // fully tracked in @EBH_PRQMAP below, never relied on through this field.
                    string sameOrderLines = string.Join(",", members.Where(m => m.PdoEntry == r.PdoEntry).Select(m => m.PdoLine).Distinct().OrderBy(l => l).Select(l => l.ToString(Inv)).ToArray());
                    oDoc.Lines.UserFields.Fields.Item(FieldPdoLine).Value = sameOrderLines;
                    oDoc.Lines.UserFields.Fields.Item("U_EBH_ValidTill").Value = ToDate(r.ValidUntil);
                    if (r.Catalog.Length > 0) oDoc.Lines.UserFields.Fields.Item("U_EBH_CATLOG").Value = r.Catalog;
                    if (r.TracksJobCode && r.PdoItemCode.Length > 0) oDoc.Lines.UserFields.Fields.Item("U_EBH_JOBCODNO").Value = r.PdoItemCode;
                }
                Check(oDoc.Add(), "Purchase Request was not created");
                int docEntry = int.Parse(Company.GetNewObjectKey(), Inv);
                // @EBH_PRQMAP: one row per component every line actually covers, with that component's own PRQTY -
                // written inside the same transaction as the document, so a rollback undoes these rows too.
                for (int i = 0; i < rows.Count; i++)
                    foreach (PlanRow m in Members(new[] { rows[i] }).Where(m => m.PrQuantity > 0))
                        InsertMapRow(docEntry, i, m.PdoEntry, m.PdoLine, m.PrQuantity);
                return docEntry;
            }
            finally { Release(oDoc); }
        }

        /// <summary>
        /// Builds and adds the Inventory Transfer Request (object 1250000001): one line per row, quantity = ITRQTY,
        /// from the row's From Warehouse to the row's Warehouse (the same warehouse is allowed, as in SAP). Writes the
        /// @EBH_ITRMAP rows. Runs inside CreateDocuments' transaction; returns the new DocEntry.
        /// </summary>
        private int AddTransferRequest(PlanHeader h, IList<PlanRow> rows, int series)
        {
            SAPbobsCOM.StockTransfer oTr = null;
            try
            {
                oTr = (SAPbobsCOM.StockTransfer)Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                if (series > 0) oTr.Series = series;
                oTr.DocDate = DateTime.Today;
                oTr.DueDate = h.RequiredDate.Length > 0 ? ToDate(h.RequiredDate) : DateTime.Today;
                oTr.FromWarehouse = rows[0].FromWarehouse;
                oTr.ToWarehouse = rows[0].Warehouse;
                SetRemark(oTr.UserFields, "Created by EBH Material Request Planning from Production Order " + string.Join(", ", Members(rows).Select(m => m.PdoNum).Distinct().ToArray()));

                for (int i = 0; i < rows.Count; i++)
                {
                    PlanRow r = rows[i];
                    if (i > 0) oTr.Lines.Add();
                    oTr.Lines.SetCurrentLine(i);
                    oTr.Lines.ItemCode = r.ItemCode;
                    oTr.Lines.Quantity = (double)r.ItrQuantity;
                    oTr.Lines.FromWarehouseCode = r.FromWarehouse;
                    oTr.Lines.WarehouseCode = r.Warehouse;
                    if (r.Project.Length > 0) oTr.Lines.ProjectCode = r.Project;
                    oTr.Lines.UserFields.Fields.Item(FieldPdo).Value = r.PdoEntry.ToString(Inv);
                }
                Check(oTr.Add(), "Inventory Transfer Request was not created");
                int docEntry = int.Parse(Company.GetNewObjectKey(), Inv);
                for (int i = 0; i < rows.Count; i++)
                    foreach (PlanRow m in Members(new[] { rows[i] }).Where(m => m.ItrQuantity > 0))
                        InsertItrMapRow(docEntry, i, m.PdoEntry, m.PdoLine, m.ItrQuantity);
                return docEntry;
            }
            finally { Release(oTr); }
        }

        /// <summary>
        /// SAP's own Inventory Transfer Request rules, checked up front so the user gets every problem named at once
        /// (rows with ITRQTY &gt; 0). <paramref name="label"/> names a row in the messages. Empty list = all good.
        /// From Warehouse and Warehouse may be the same, as SAP's own Inventory Transfer Request allows.
        /// </summary>
        public List<string> CheckTransferRows(IList<PlanRow> rows, Func<PlanRow, string> label)
        {
            var errors = new List<string>();
            var whsCodes = rows.SelectMany(r => new[] { r.FromWarehouse, r.Warehouse }).Where(w => w.Length > 0).Distinct().ToList();
            // WhsCode -> Inactive
            var whs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (whsCodes.Count > 0)
                foreach (var w in Db.Rows("Select \"WhsCode\", \"Inactive\" from OWHS Where \"WhsCode\" in (" + InList(whsCodes) + ")"))
                    whs[Db.Str(w, "WhsCode")] = Db.Str(w, "Inactive");
            // ItemCode -> [InvntItem, ManBtchNum, ManSerNum, frozenFor]
            var items = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            var itemCodes = rows.Select(r => r.ItemCode).Where(c => c.Length > 0).Distinct().ToList();
            if (itemCodes.Count > 0)
                foreach (var it in Db.Rows("Select \"ItemCode\", \"InvntItem\", \"ManBtchNum\", \"ManSerNum\", \"frozenFor\" from OITM Where \"ItemCode\" in (" + InList(itemCodes) + ")"))
                    items[Db.Str(it, "ItemCode")] = new[] { Db.Str(it, "InvntItem"), Db.Str(it, "ManBtchNum"), Db.Str(it, "ManSerNum"), Db.Str(it, "frozenFor") };

            foreach (PlanRow r in rows)
            {
                string at = label(r);
                string[] it;
                if (!items.TryGetValue(r.ItemCode, out it)) { errors.Add(at + ": item not found."); continue; }
                if (it[0] != "Y") errors.Add(at + ": " + r.ItemCode + " is not an inventory item, so it cannot be transferred.");
                if (it[3] == "Y") errors.Add(at + ": " + r.ItemCode + " is inactive.");
                if (it[1] == "Y" || it[2] == "Y") errors.Add(at + ": " + (it[1] == "Y" ? "batch" : "serial") + "-managed items cannot be put on an Inventory Transfer Request from here.");

                string inactive;
                if (r.FromWarehouse.Length == 0) errors.Add(at + ": From Warehouse is missing.");
                else if (!whs.TryGetValue(r.FromWarehouse, out inactive)) errors.Add(at + ": From Warehouse " + r.FromWarehouse + " does not exist.");
                else if (inactive == "Y") errors.Add(at + ": From Warehouse " + r.FromWarehouse + " is inactive.");
                if (r.Warehouse.Length == 0) errors.Add(at + ": Warehouse (To Warehouse) is missing.");
                else if (!whs.TryGetValue(r.Warehouse, out inactive)) errors.Add(at + ": Warehouse " + r.Warehouse + " does not exist.");
                else if (inactive == "Y") errors.Add(at + ": Warehouse " + r.Warehouse + " is inactive.");
            }
            return errors;
        }

        /// <summary>
        /// Rows whose ITRQTY is more than what is available in their From Warehouse (OITW.OnHand - IsCommited). SAP
        /// still accepts the request - only the actual transfer is blocked - so this is a warning to confirm, not an error.
        /// </summary>
        public List<string> TransferStockShortfalls(IList<PlanRow> rows, Func<PlanRow, string> label)
        {
            var result = new List<string>();
            var items = rows.Select(r => r.ItemCode).Distinct().ToList();
            if (items.Count == 0) return result;
            var whs = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var w in Db.Rows("Select \"ItemCode\", \"WhsCode\", \"OnHand\" - \"IsCommited\" as \"Avail\" from OITW Where \"ItemCode\" in (" + InList(items) + ")"))
                whs[Db.Str(w, "ItemCode") + "|" + Db.Str(w, "WhsCode")] = Dec(w, "Avail");
            // Several rows can draw the same item from the same warehouse: compare their total
            foreach (var g in rows.GroupBy(r => r.ItemCode + "|" + r.FromWarehouse, StringComparer.OrdinalIgnoreCase))
            {
                decimal want = g.Sum(r => r.ItrQuantity), have;
                PlanRow first = g.First();
                whs.TryGetValue(g.Key, out have);
                if (want > have)
                    result.Add(label(first) + ": ITRQTY " + want.ToString("0.######", Inv) + " is more than the " + have.ToString("0.######", Inv) +
                        " available in warehouse " + first.FromWarehouse + ".");
            }
            return result;
        }

        private static List<Dictionary<string, string>> Field(string uid)
        {
            return new List<Dictionary<string, string>> { new Dictionary<string, string> { { "field", uid } } };
        }

        private static string Cut(string value, int length)
        {
            value = value ?? "";
            return value.Length <= length ? value : value.Substring(0, length);
        }
        #endregion

        #region Production Order side
        /// <summary>
        /// Refreshes the display fields of the Production Order components from the open Purchase Request lines:
        /// U_EBH_REQQTY = quantity on open requests, U_EBH_PRNO = latest open request. Only changed values are written.
        /// Also called when a request was cancelled or closed, which frees its quantity again.
        /// </summary>
        public void Sync(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.Distinct().ToList();
            SyncMapStatus(ids);
            foreach (int pdo in ids)
            {
                var comps = GetComponents(new[] { pdo });
                // Latest open request per component line, from @EBH_PRQMAP - a merged line counts here too, not only
                // one where this is the primary component.
                var latest = Db.Rows(
                    "Select M.\"U_PdoLine\" as \"PdoLine\", Max(P0.\"DocNum\") as \"PrNum\" from " + MapTable + " M " +
                    "Inner Join PRQ1 P1 on P1.\"DocEntry\" = M.\"U_PrEntry\" and P1.\"LineNum\" = M.\"U_PrLine\" " +
                    "Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                    "Where P0.\"CANCELED\" = 'N' and P1.\"LineStatus\" = 'O' and M.\"U_PdoEntry\" = " + pdo + " " +
                    "Group by M.\"U_PdoLine\"")
                    .ToDictionary(r => Int(r, "PdoLine"), r => Int(r, "PrNum"));
                var shown = Db.Rows("Select \"LineNum\", \"U_EBH_PRNO\", \"U_EBH_REQQTY\" from WOR1 Where \"DocEntry\" = " + pdo)
                    .ToDictionary(r => Int(r, "LineNum"), r => new[] { Dec(r, "U_EBH_PRNO"), Dec(r, "U_EBH_REQQTY") });
                var changes = new List<Component>();
                foreach (Component c in comps)
                {
                    int prNo; latest.TryGetValue(c.PdoLine, out prNo);
                    decimal[] was; shown.TryGetValue(c.PdoLine, out was);
                    if (was == null || was[0] != prNo || was[1] != c.OpenRequested) changes.Add(c);
                }
                if (changes.Count == 0) continue;

                SAPbobsCOM.ProductionOrders oPdo = null;
                try
                {
                    oPdo = (SAPbobsCOM.ProductionOrders)Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
                    if (!oPdo.GetByKey(pdo)) continue;
                    // The DI line index is the position, not WOR1.LineNum (they differ once a component was deleted)
                    var byLine = changes.ToDictionary(c => c.PdoLine);
                    for (int i = 0; i < oPdo.Lines.Count; i++)
                    {
                        oPdo.Lines.SetCurrentLine(i);
                        Component c;
                        if (!byLine.TryGetValue(oPdo.Lines.LineNumber, out c)) continue;
                        int prNo; latest.TryGetValue(c.PdoLine, out prNo);
                        oPdo.Lines.UserFields.Fields.Item("U_EBH_PRNO").Value = prNo.ToString(Inv);
                        oPdo.Lines.UserFields.Fields.Item("U_EBH_REQQTY").Value = c.OpenRequested.ToString(Inv);
                    }
                    Check(oPdo.Update(), "Production Order " + pdo + " was not updated");
                }
                finally { Release(oPdo); }
            }
        }
        #endregion
    }
}

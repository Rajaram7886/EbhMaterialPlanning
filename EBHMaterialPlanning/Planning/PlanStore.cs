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
        public const string FieldPdo = "U_EBH_PRODUCTIONNO", FieldPdoLine = "U_EBH_PRODLINE", FieldPdoNum = "U_EBH_PRODNO";
        private const int ObjPurchaseRequest = 1470000113, ObjProductionOrder = 202;
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

        /// <summary>Numbering series of the Purchase Request (object 1470000113).</summary>
        public List<KeyValuePair<string, string>> Series()
        {
            return Db.Rows("Select \"Series\", \"SeriesName\" from NNM1 Where \"ObjectCode\" = '" + ObjPurchaseRequest + "' and \"Locked\" = 'N' Order by \"Series\"")
                .Select(r => new KeyValuePair<string, string>(Db.Str(r, "Series"), Db.Str(r, "SeriesName"))).ToList();
        }

        public int DefaultSeries()
        {
            var row = Db.FirstRow("Select \"DfltSeries\" from ONNM Where \"ObjectCode\" = '" + ObjPurchaseRequest + "'");
            return Int(row, "DfltSeries");
        }

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

        public bool VendorExists(string code)
        {
            return Db.FirstRow("Select \"CardCode\" from OCRD Where \"CardType\" = 'S' and \"CardCode\" = " + Db.Quote(code)) != null;
        }

        public bool WarehouseExists(string code)
        {
            return Db.FirstRow("Select \"WhsCode\" from OWHS Where \"WhsCode\" = " + Db.Quote(code)) != null;
        }

        /// <summary>Next number the chosen series will give the Purchase Request (display only, SAP assigns the real one on Add).</summary>
        public int NextNumber(int series)
        {
            if (series <= 0) return 0;
            var row = Db.FirstRow("Select \"NextNumber\" from NNM1 Where \"Series\" = " + series + " and \"ObjectCode\" = '" + ObjPurchaseRequest + "'");
            return Int(row, "NextNumber");
        }

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
            return Db.Rows("Select Top 500 T0.\"DocEntry\", T0.\"DocNum\", T0.\"Series\", N1.\"SeriesName\", T0.\"Project\", T0.\"ItemCode\", T0.\"ProdName\", T0.\"Status\", T0.\"DueDate\" " +
                "from OWOR T0 Left Join NNM1 N1 on N1.\"Series\" = T0.\"Series\" and N1.\"ObjectCode\" = '" + ObjProductionOrder + "' " +
                "Where " + where + " Order by T0.\"DocNum\" Desc")
                .Select(r => new ProductionOrder
                {
                    DocEntry = Int(r, "DocEntry"), DocNum = Int(r, "DocNum"), Series = Int(r, "Series"), SeriesName = Db.Str(r, "SeriesName"),
                    Project = Db.Str(r, "Project"), ItemCode = Db.Str(r, "ItemCode"), ItemName = Db.Str(r, "ProdName"), Status = Db.Str(r, "Status"), DueDate = Date(r, "DueDate")
                }).ToList();
        }

        /// <summary>
        /// Item components (not resources) of the Production Orders with the quantity already on open Purchase Requests.
        /// One Component per Production Order line, so the same item in two orders stays on two rows.
        /// </summary>
        public List<Component> GetComponents(IEnumerable<int> pdoEntries)
        {
            string ids = InList(pdoEntries);
            return Db.Rows(
                "Select T0.\"DocEntry\", T0.\"DocNum\", T0.\"Project\", T1.\"LineNum\", T1.\"ItemCode\", I.\"ItemName\", T1.\"PlannedQty\", T1.\"wareHouse\", " +
                "IsNull(U.\"UomCode\", I.\"BuyUnitMsr\") as \"Uom\", " +
                "IsNull((Select Sum(P1.\"OpenQty\") from PRQ1 P1 Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Where P0.\"CANCELED\" = 'N' and P1.\"LineStatus\" = 'O' and P1.\"" + FieldPdo + "\" = T0.\"DocEntry\" and P1.\"" + FieldPdoLine + "\" = T1.\"LineNum\"), 0) as \"Requested\" " +
                "from OWOR T0 Inner Join WOR1 T1 on T1.\"DocEntry\" = T0.\"DocEntry\" " +
                "Inner Join OITM I on I.\"ItemCode\" = T1.\"ItemCode\" " +
                "Left Join OUOM U on U.\"UomEntry\" = T1.\"UomEntry\" " +
                "Where T0.\"DocEntry\" in (" + ids + ") and T1.\"ItemType\" = 4 and T0.\"Status\" in ('P','R') " +
                "Order by T0.\"DocNum\", T1.\"VisOrder\"")
                .Select(r => new Component
                {
                    PdoEntry = Int(r, "DocEntry"), PdoNum = Int(r, "DocNum"), PdoLine = Int(r, "LineNum"), Project = Db.Str(r, "Project"),
                    ItemCode = Db.Str(r, "ItemCode"), Description = Db.Str(r, "ItemName"), Planned = Dec(r, "PlannedQty"),
                    OpenRequested = Dec(r, "Requested"), Uom = Db.Str(r, "Uom"), Warehouse = Db.Str(r, "wareHouse")
                }).ToList();
        }
        #endregion

        #region Linked documents
        /// <summary>Production Order DocEntries (of the given ones) that already have at least one Purchase Request line raised from them (any status) - what the Linked Documents button would show something for.</summary>
        public List<int> OrdersWithLinkedRequests(IEnumerable<int> pdoEntries)
        {
            var ids = pdoEntries.Distinct().ToList();
            if (ids.Count == 0) return new List<int>();
            return Db.Rows("Select Distinct \"" + FieldPdo + "\" as \"Pdo\" from PRQ1 Where \"" + FieldPdo + "\" in (" + InList(ids) + ")")
                .Select(r => Int(r, "Pdo")).ToList();
        }

        /// <summary>Purchase Request lines raised from a Production Order, with the request status.</summary>
        public List<LinkedRequest> LinkedRequests(int pdoEntry)
        {
            return Db.Rows(
                "Select P0.\"DocEntry\", P0.\"DocNum\", P0.\"Series\", N1.\"SeriesName\", P0.\"DocDate\", P0.\"DocStatus\", P0.\"CANCELED\", P1.\"LineNum\", P1.\"ItemCode\", P1.\"Dscription\", " +
                "P1.\"Quantity\", P1.\"OpenQty\", P1.\"" + FieldPdoLine + "\" as \"PdoLine\" " +
                "from PRQ1 P1 Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                "Left Join NNM1 N1 on N1.\"Series\" = P0.\"Series\" and N1.\"ObjectCode\" = '" + ObjPurchaseRequest + "' " +
                "Where P1.\"" + FieldPdo + "\" = " + pdoEntry + " Order by P0.\"DocNum\", P1.\"LineNum\"")
                .Select(r => new LinkedRequest
                {
                    PrEntry = Int(r, "DocEntry"), PrNum = Int(r, "DocNum"), PrLine = Int(r, "LineNum"),
                    Series = Int(r, "Series"), SeriesName = Db.Str(r, "SeriesName"), PostingDate = Date(r, "DocDate"),
                    Status = Db.Str(r, "CANCELED") == "Y" ? "X" : Db.Str(r, "DocStatus"),
                    ItemCode = Db.Str(r, "ItemCode"), Description = Db.Str(r, "Dscription"), PdoLine = Int(r, "PdoLine"),
                    Quantity = Dec(r, "Quantity"), OpenQuantity = Dec(r, "OpenQty")
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

        public void OpenProductionOrder(int docEntry)
        {
            Program.SBO_App.OpenForm((SAPbouiCOM.BoFormObjectEnum)ObjProductionOrder, "", docEntry.ToString(Inv));
        }
        #endregion

        #region Create Purchase Request
        /// <summary>
        /// Adds one Purchase Request with a row per selected material. The document is added completely or not at all
        /// (one DI call inside a transaction). Afterwards the Production Order components are refreshed; that part is
        /// repeatable (Sync), so a failure there is reported as a warning and does not undo the request.
        /// </summary>
        public CreatedRequest CreatePurchaseRequest(PlanHeader h, IList<PlanRow> rows)
        {
            if (!VendorExists(h.Vendor)) throw new PlanException(400, "Vendor " + h.Vendor + " is not a valid supplier.", Field("eVend"));
            if (!WarehouseExists(h.Warehouse)) throw new PlanException(400, "Warehouse " + h.Warehouse + " does not exist.", Field("eWhs"));
            bool multiBranch = MultiBranch;
            if (multiBranch && h.Branch <= 0) throw new PlanException(400, "Select the Branch.", Field("cBranch"));

            SAPbobsCOM.Documents oDoc = null;
            bool inTransaction = false;
            var result = new CreatedRequest();
            try
            {
                if (Company.InTransaction) throw new PlanException(409, "SAP is busy with another transaction. Try again in a moment.");
                Company.StartTransaction();
                inTransaction = true;

                oDoc = (SAPbobsCOM.Documents)Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);
                oDoc.Series = h.Series;
                // The branch property is declared on IDocuments, not on the Documents interface GetBusinessObject returns
                if (multiBranch) ((SAPbobsCOM.IDocuments)oDoc).BPL_IDAssignedToInvoice = h.Branch;
                oDoc.DocDate = DateTime.Today;
                oDoc.DocDueDate = ToDate(h.ValidUntil);
                oDoc.RequriedDate = ToDate(h.RequiredDate);
                oDoc.ReqType = 12;                                      // requester is a user
                oDoc.Requester = Company.UserName;
                oDoc.Comments = Cut("Created by EBH Material Request Planning from Production Order " + string.Join(", ", rows.Select(r => r.PdoNum).Distinct().ToArray()), 254);
                oDoc.UserFields.Fields.Item("U_EBH_Vendor").Value = h.Vendor;
                oDoc.UserFields.Fields.Item("U_EBH_Whse").Value = h.Warehouse;

                for (int i = 0; i < rows.Count; i++)
                {
                    PlanRow r = rows[i];
                    if (i > 0) oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(i);
                    oDoc.Lines.ItemCode = r.ItemCode;
                    oDoc.Lines.Quantity = (double)r.Quantity;
                    oDoc.Lines.WarehouseCode = r.Warehouse;
                    oDoc.Lines.LineVendor = r.Vendor;
                    // Project is declared on IDocument_Lines, not on the Document_Lines interface Lines returns
                    if (r.Project.Length > 0) ((SAPbobsCOM.IDocument_Lines)oDoc.Lines).ProjectCode = r.Project;
                    oDoc.Lines.RequiredDate = ToDate(r.RequiredDate);
                    oDoc.Lines.UnitPrice = (double)r.Price;
                    oDoc.Lines.DiscountPercent = (double)r.Discount;
                    oDoc.Lines.UserFields.Fields.Item(FieldPdo).Value = r.PdoEntry.ToString(Inv);
                    oDoc.Lines.UserFields.Fields.Item(FieldPdoLine).Value = r.PdoLine.ToString(Inv);
                    oDoc.Lines.UserFields.Fields.Item(FieldPdoNum).Value = r.PdoNum.ToString(Inv);
                    oDoc.Lines.UserFields.Fields.Item("U_EBH_ValidTill").Value = ToDate(r.ValidUntil);
                }
                Check(oDoc.Add(), "Purchase Request was not created");
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

            result.DocNum = Int(Db.FirstRow("Select \"DocNum\" from OPRQ Where \"DocEntry\" = " + result.DocEntry), "DocNum");
            try { Sync(rows.Select(r => r.PdoEntry).Distinct()); }
            catch (Exception ex)
            {
                result.Warning = "Purchase Request " + result.DocNum + " was created, but the Production Order could not be refreshed (" + ex.Message + "). It is refreshed when the Production Order's Purchase Requests are opened.";
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
            foreach (int pdo in pdoEntries.Distinct())
            {
                var comps = GetComponents(new[] { pdo });
                var latest = Db.Rows(
                    "Select P1.\"" + FieldPdoLine + "\" as \"PdoLine\", Max(P0.\"DocNum\") as \"PrNum\" from PRQ1 P1 Inner Join OPRQ P0 on P0.\"DocEntry\" = P1.\"DocEntry\" " +
                    "Where P0.\"CANCELED\" = 'N' and P1.\"LineStatus\" = 'O' and P1.\"" + FieldPdo + "\" = " + pdo + " Group by P1.\"" + FieldPdoLine + "\"")
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

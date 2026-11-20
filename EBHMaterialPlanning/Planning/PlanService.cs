using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Workflow and rules of the Material Request Planning form. Pure code except for the store calls, so the
    /// calculation and the checks before creation can be tried without SAP.
    /// </summary>
    public static class PlanService
    {
        public static PlanStore Store;

        #region Calculation
        /// <summary>Unit price after the discount percentage.</summary>
        public static decimal AfterDiscount(decimal price, decimal discount)
        {
            return Math.Round(price * (1 - discount / 100m), 6);
        }

        /// <summary>Row total before tax: quantity x price after discount.</summary>
        public static decimal Total(decimal quantity, decimal priceAfterDiscount)
        {
            return Math.Round(quantity * priceAfterDiscount, 2);
        }

        /// <summary>The member of a merged row whose price wins: the lowest Price After Discount among those actually
        /// priced (an unpriced 0 never wins over a real price); the first member when none is priced.</summary>
        public static T Cheapest<T>(IList<T> members, Func<T, decimal> priceAfterDiscount) where T : class
        {
            return members.Where(m => priceAfterDiscount(m) > 0).OrderBy(priceAfterDiscount).FirstOrDefault() ?? members[0];
        }
        #endregion

        #region Rows
        /// <summary>One row per Production Order component, quantity = what is still to be requested. Components with nothing left are skipped.</summary>
        public static List<PlanRow> LoadRows(IEnumerable<int> pdoEntries, out int skipped)
        {
            var rows = new List<PlanRow>();
            skipped = 0;
            var components = Store.GetComponents(pdoEntries).ToList();
            var linked = new HashSet<int>(Store.OrdersWithLinkedRequests(components.Select(c => c.PdoEntry).Distinct()));
            // Every loaded row starts with Required Date = today and Valid Until = today + 1 month (still editable per row)
            string today = DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            string inAMonth = DateTime.Today.AddMonths(1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            foreach (Component c in components)
            {
                if (c.Balance <= 0) { skipped++; continue; }
                rows.Add(new PlanRow
                {
                    PdoEntry = c.PdoEntry, PdoNum = c.PdoNum, PdoLine = c.PdoLine, Project = c.Project,
                    ItemCode = c.ItemCode, Description = c.Description, Uom = c.Uom,
                    Quantity = c.Balance, PrQuantity = c.Balance, ItrQuantity = 0, Balance = c.Balance, AlreadyRequested = c.OpenRequested, Warehouse = c.Warehouse,
                    Manufacturer = c.Manufacturer, ManufacturerName = c.ManufacturerName, HasLinkedRequest = linked.Contains(c.PdoEntry),
                    Price = c.Price, Discount = c.Discount, PriceAfterDiscount = c.PriceAfterDiscount, Catalog = c.Catalog, TracksJobCode = c.TracksJobCode, PdoItemCode = c.PdoItemCode,
                    RequiredDate = today, ValidUntil = inAMonth
                });
            }
            return rows;
        }

        /// <summary>Copies the header defaults into the rows. Existing row values are replaced.</summary>
        public static void ApplyHeader(PlanHeader h, IEnumerable<PlanRow> rows)
        {
            foreach (PlanRow r in rows)
            {
                if (h.Vendor.Length > 0) r.Vendor = h.Vendor;
                if (h.Warehouse.Length > 0) r.Warehouse = h.Warehouse;
                if (h.RequiredDate.Length > 0) r.RequiredDate = h.RequiredDate;
                if (h.ValidUntil.Length > 0) r.ValidUntil = h.ValidUntil;
            }
        }

        #endregion

        #region Checks
        /// <summary>Which document(s) a Create press makes: the Purchase Request button, the Inventory Transfer
        /// Request button, or Combine (both, in one SAP transaction).</summary>
        public enum CreateKind { PurchaseRequest, TransferRequest, Both }

        private static bool MakesPr(CreateKind k) { return k != CreateKind.TransferRequest; }
        private static bool MakesItr(CreateKind k) { return k != CreateKind.PurchaseRequest; }

        private static string RowName(PlanRow r, int no)
        {
            return "Row " + no + " (" + r.ItemCode + ", Production Order " + r.PdoNum + ")";
        }

        /// <summary>Selected rows that go onto the Purchase Request (PRQTY &gt; 0) for this press.</summary>
        public static List<PlanRow> PrRows(IList<PlanRow> rows, CreateKind kind)
        {
            return MakesPr(kind) ? rows.Where(r => r.Selected && r.PrQuantity > 0).ToList() : new List<PlanRow>();
        }

        /// <summary>Selected rows that go onto the Inventory Transfer Request (ITRQTY &gt; 0) for this press.</summary>
        public static List<PlanRow> ItrRows(IList<PlanRow> rows, CreateKind kind)
        {
            return MakesItr(kind) ? rows.Where(r => r.Selected && r.ItrQuantity > 0).ToList() : new List<PlanRow>();
        }

        /// <summary>
        /// The checks before creation. Throws a PlanException that names the field or row to correct. Row numbers
        /// are the numbers the user sees (1 = first row of the grid).
        /// Every selected row: PRQTY and ITRQTY not negative, not both 0, PRQTY + ITRQTY &lt;= Quantity.
        /// Purchase Request rows (PRQTY &gt; 0): warehouse, dates, price. Inventory Transfer Request rows (ITRQTY &gt; 0):
        /// SAP's own transfer rules (PlanStore.CheckTransferRows).
        /// </summary>
        public static void Validate(PlanHeader h, IList<PlanRow> rows, CreateKind kind)
        {
            var errors = new List<Dictionary<string, string>>();
            Action<string, string> fail = (field, message) => errors.Add(new Dictionary<string, string> { { "field", field }, { "message", message } });

            var picked = rows.Select((r, i) => new { Row = r, No = i + 1 }).Where(x => x.Row.Selected).ToList();
            if (picked.Count == 0) fail("gRows", "Select at least one material row.");

            bool pr = MakesPr(kind) && picked.Any(x => x.Row.PrQuantity > 0);
            bool itr = MakesItr(kind) && picked.Any(x => x.Row.ItrQuantity > 0);
            if (picked.Count > 0)
            {
                if (kind == CreateKind.PurchaseRequest && !pr) fail("gRows", "No selected row has a PRQTY - nothing to put on a Purchase Request.");
                if (kind == CreateKind.TransferRequest && !itr) fail("gRows", "No selected row has an ITRQTY - nothing to put on an Inventory Transfer Request.");
                if (kind == CreateKind.Both && !pr && !itr) fail("gRows", "No selected row has a PRQTY or an ITRQTY.");
            }

            if (h.RequiredDate.Length == 0) fail("eReq", "Enter the Required Date.");
            if (pr)
            {
                if (h.Series <= 0) fail("cSeries", "Select the Target Document Series.");
                if (Store.MultiBranch && h.Branch <= 0) fail("cBranch", "Select the Branch.");
                if (h.ValidUntil.Length == 0) fail("eValid", "Enter the Valid Until Date.");
                if (h.Warehouse.Length == 0) fail("eWhs", "Select the Warehouse.");
            }

            foreach (var x in picked)
            {
                PlanRow r = x.Row;
                string at = RowName(r, x.No);
                if (r.PdoEntry <= 0 || r.ItemCode.Length == 0) fail("gRows", at + " has no source Production Order.");
                if (r.PrQuantity < 0) fail("gRows", at + ": PRQTY cannot be negative.");
                if (r.ItrQuantity < 0) fail("gRows", at + ": ITRQTY cannot be negative.");
                if (r.PrQuantity == 0 && r.ItrQuantity == 0) fail("gRows", at + ": enter a PRQTY or an ITRQTY (both are 0).");
                else if (r.PrQuantity + r.ItrQuantity > r.Quantity)
                    fail("gRows", at + ": PRQTY " + Qty(r.PrQuantity) + " + ITRQTY " + Qty(r.ItrQuantity) + " is more than the Quantity " + Qty(r.Quantity) + ".");
                if (MakesPr(kind) && r.PrQuantity > 0)
                {
                    if (r.Price < 0) fail("gRows", at + ": price cannot be negative.");
                    if (r.Discount < 0 || r.Discount > 100) fail("gRows", at + ": discount must be between 0 and 100.");
                    if (r.Warehouse.Length == 0) fail("gRows", at + ": warehouse is missing.");
                    if (r.RequiredDate.Length == 0) fail("gRows", at + ": required date is missing.");
                    if (r.ValidUntil.Length == 0) fail("gRows", at + ": valid until date is missing.");
                }
            }
            if (itr && errors.Count == 0)
            {
                var numbers = picked.ToDictionary(x => x.Row, x => x.No);
                foreach (string e in Store.CheckTransferRows(ItrRows(rows, kind), r => RowName(r, numbers[r]))) fail("gRows", e);
            }
            if (errors.Count > 0)
                throw new PlanException(400, string.Join(Environment.NewLine, errors.Take(8).Select(e => e["message"]).ToArray()) +
                    (errors.Count > 8 ? Environment.NewLine + "... and " + (errors.Count - 8) + " more." : ""), errors);

            // The same quantity may not be requested twice in one press: rows of one source component add up
            foreach (var g in picked.GroupBy(x => x.Row.PdoEntry + "/" + x.Row.PdoLine))
                if (g.Count() > 1 && g.Sum(x => x.Row.PrQuantity + x.Row.ItrQuantity) > g.First().Row.Balance)
                    throw new PlanException(400, "Production Order " + g.First().Row.PdoNum + " line " + g.First().Row.PdoLine + " (" + g.First().Row.ItemCode + ") is on more than one row and the quantities exceed what is still to be requested.");
        }

        private static string Qty(decimal q) { return q.ToString("0.######", CultureInfo.InvariantCulture); }

        /// <summary>Inventory Transfer Request rows asking for more than is in stock in their From Warehouse
        /// (a warning to confirm: SAP accepts the request, only the transfer itself is blocked). Empty = none.</summary>
        public static List<string> TransferWarnings(IList<PlanRow> rows, CreateKind kind)
        {
            var itrRows = ItrRows(rows, kind);
            if (itrRows.Count == 0) return new List<string>();
            var numbers = rows.Select((r, i) => new { r, No = i + 1 }).ToDictionary(x => x.r, x => x.No);
            return Store.TransferStockShortfalls(itrRows, r => RowName(r, numbers[r]));
        }
        #endregion

        /// <summary>Validates and creates the Purchase Request and/or Inventory Transfer Request for the selected rows.</summary>
        public static CreatedDocuments Create(PlanHeader h, IList<PlanRow> rows, CreateKind kind)
        {
            Validate(h, rows, kind);
            return Store.CreateDocuments(h, PrRows(rows, kind), ItrRows(rows, kind));
        }

        #region Purchase Order
        /// <summary>One row per ticked Purchase Request line (DocEntry, LineNum), re-read fresh from SAP.</summary>
        public static List<PoRow> LoadPoRows(IEnumerable<KeyValuePair<int, int>> lineKeys)
        {
            return Store.GetRequestLines(lineKeys);
        }

        /// <summary>The given Purchase Request lines (DocEntry, LineNum) re-read fresh, whatever their current status -
        /// used to refresh the picker grid's own rows after drawing a Purchase Order from some of them.</summary>
        public static List<PurchaseRequestLine> RefreshRequestLines(IEnumerable<KeyValuePair<int, int>> lineKeys)
        {
            return Store.RefreshRequestLines(lineKeys);
        }

        /// <summary>
        /// The basic checks before creating the Purchase Order. Throws a PlanException that names the field or row to correct.
        /// Rows are named by the grid's "#" (PoRow.SrNo).
        /// </summary>
        public static void ValidatePo(PoHeader h, IList<PoRow> rows)
        {
            var errors = new List<Dictionary<string, string>>();
            Action<string, string> fail = (field, message) => errors.Add(new Dictionary<string, string> { { "field", field }, { "message", message } });

            if (h.Series <= 0) fail("cSeries", "Select the Target Document Series.");
            if (Store.MultiBranch && h.Branch <= 0) fail("cBranch", "Select the Branch.");
            if (h.Vendor.Length == 0) fail("eVend", "Select the Vendor.");
            if (h.DueDate.Length == 0) fail("eDue", "Enter the Delivery Date.");

            var picked = rows.Select((r, i) => new { Row = r, No = i + 1 }).Where(x => x.Row.Selected).ToList();
            if (picked.Count == 0) fail("gRows", "Select at least one material row.");
            foreach (var x in picked)
            {
                PoRow r = x.Row;
                string at = (r.SrNo > 0 ? "Sr No " + r.SrNo : "Row hidden by the Manufacturer filter") + " (" + r.ItemCode + ", Purchase Request " + r.PrNum + ")";
                if (r.PrEntry <= 0 || r.ItemCode.Length == 0) fail("gRows", at + " has no source Purchase Request.");
                if (r.Quantity <= 0) fail("gRows", at + ": quantity must be more than 0.");
                else if (r.Quantity > r.OpenQuantity) fail("gRows", at + ": quantity " + r.Quantity.ToString("0.######") + " is more than the " + r.OpenQuantity.ToString("0.######") + " still open.");
                if (r.Price <= 0) fail("gRows", at + ": enter a price more than 0.");
                if (r.Discount < 0 || r.Discount > 100) fail("gRows", at + ": discount must be between 0 and 100.");
                else if (PriceOverCap(r.PriceAfterDiscount, r.MaxPriceAfterDiscount)) fail("gRows", at + ": " + PriceCapMessage(r.PriceAfterDiscount, r.MaxPriceAfterDiscount));
                if (r.Warehouse.Length == 0) fail("gRows", at + ": warehouse is missing.");
                if (r.TaxCode.Length == 0) fail("gRows", at + ": tax code is missing (set it on the source Purchase Request line).");
                // A row can only have reached the grid under a different Vendor if it was loaded before the Vendor filter
                // changed; the Purchase Order targets a single Vendor (h.Vendor), so such a row must not go through.
                if (h.Vendor.Length > 0 && r.Vendor.Length > 0 && !string.Equals(r.Vendor, h.Vendor, StringComparison.OrdinalIgnoreCase))
                    fail("gRows", at + ": belongs to Vendor " + r.Vendor + ", not " + h.Vendor + ". Remove it or reload for the right Vendor.");
            }
            if (errors.Count > 0)
                throw new PlanException(400, string.Join(Environment.NewLine, errors.Take(8).Select(e => e["message"]).ToArray()) +
                    (errors.Count > 8 ? Environment.NewLine + "... and " + (errors.Count - 8) + " more." : ""), errors);
        }

        /// <summary>Price less Discount % is above the Production Order's planned Price After Discount (0 = no planned price, no limit).</summary>
        public static bool PriceOverCap(decimal priceAfterDiscount, decimal cap)
        {
            return cap > 0 && priceAfterDiscount > cap;
        }

        public static string PriceCapMessage(decimal priceAfterDiscount, decimal cap)
        {
            return "Price After Discount " + priceAfterDiscount.ToString("0.######", CultureInfo.InvariantCulture) +
                " is more than the Production Order's Price After Discount " + cap.ToString("0.######", CultureInfo.InvariantCulture) +
                ". Price is not allowed to be more than " + cap.ToString("0.######", CultureInfo.InvariantCulture) + ".";
        }

        /// <summary>Validates and creates the Purchase Order for the selected rows.</summary>
        public static CreatedOrder CreatePo(PoHeader h, IList<PoRow> rows)
        {
            ValidatePo(h, rows);
            return Store.CreatePurchaseOrder(h, rows.Where(r => r.Selected).ToList());
        }
        #endregion
    }
}

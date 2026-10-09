using System;
using System.Collections.Generic;
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
        public static decimal Total(decimal quantity, decimal price, decimal discount)
        {
            return Math.Round(quantity * AfterDiscount(price, discount), 2);
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
            foreach (Component c in components)
            {
                if (c.Balance <= 0) { skipped++; continue; }
                rows.Add(new PlanRow
                {
                    PdoEntry = c.PdoEntry, PdoNum = c.PdoNum, PdoLine = c.PdoLine, Project = c.Project,
                    ItemCode = c.ItemCode, Description = c.Description, Uom = c.Uom,
                    Quantity = c.Balance, Balance = c.Balance, AlreadyRequested = c.OpenRequested, Warehouse = c.Warehouse,
                    HasLinkedRequest = linked.Contains(c.PdoEntry)
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

        /// <summary>Copies only the warehouse into the rows (header vendor and dates are left untouched).</summary>
        public static void ApplyWarehouse(string warehouse, IEnumerable<PlanRow> rows)
        {
            if (string.IsNullOrEmpty(warehouse)) return;
            foreach (PlanRow r in rows) r.Warehouse = warehouse;
        }
        #endregion

        #region Checks
        /// <summary>
        /// The basic checks before creation. Throws a PlanException that names the field or row to correct.
        /// Row numbers are the numbers the user sees (1 = first row of the grid).
        /// </summary>
        public static void Validate(PlanHeader h, IList<PlanRow> rows)
        {
            var errors = new List<Dictionary<string, string>>();
            Action<string, string> fail = (field, message) => errors.Add(new Dictionary<string, string> { { "field", field }, { "message", message } });

            if (h.Series <= 0) fail("cSeries", "Select the Target Document Series.");
            if (Store.MultiBranch && h.Branch <= 0) fail("cBranch", "Select the Branch.");
            if (h.RequiredDate.Length == 0) fail("eReq", "Enter the Required Date.");
            if (h.ValidUntil.Length == 0) fail("eValid", "Enter the Valid Until Date.");
            if (h.Vendor.Length == 0) fail("eVend", "Select the Vendor.");
            if (h.Warehouse.Length == 0) fail("eWhs", "Select the Warehouse.");

            var picked = rows.Select((r, i) => new { Row = r, No = i + 1 }).Where(x => x.Row.Selected).ToList();
            if (picked.Count == 0) fail("gRows", "Select at least one material row.");
            foreach (var x in picked)
            {
                PlanRow r = x.Row;
                string at = "Row " + x.No + " (" + r.ItemCode + ", Production Order " + r.PdoNum + ")";
                if (r.PdoEntry <= 0 || r.ItemCode.Length == 0) fail("gRows", at + " has no source Production Order.");
                if (r.Quantity <= 0) fail("gRows", at + ": quantity must be more than 0.");
                else if (r.Quantity > r.Balance) fail("gRows", at + ": quantity " + r.Quantity.ToString("0.######") + " is more than the " + r.Balance.ToString("0.######") + " still to be requested.");
                if (r.Price < 0) fail("gRows", at + ": price cannot be negative.");
                if (r.Discount < 0 || r.Discount > 100) fail("gRows", at + ": discount must be between 0 and 100.");
                if (r.Warehouse.Length == 0) fail("gRows", at + ": warehouse is missing.");
                if (r.Vendor.Length == 0) fail("gRows", at + ": vendor is missing.");
                if (r.RequiredDate.Length == 0) fail("gRows", at + ": required date is missing.");
                if (r.ValidUntil.Length == 0) fail("gRows", at + ": valid until date is missing.");
            }
            if (errors.Count > 0)
                throw new PlanException(400, string.Join(Environment.NewLine, errors.Take(8).Select(e => e["message"]).ToArray()) +
                    (errors.Count > 8 ? Environment.NewLine + "... and " + (errors.Count - 8) + " more." : ""), errors);

            // The same quantity may not be requested twice in one request: rows of one source component add up
            foreach (var g in picked.GroupBy(x => x.Row.PdoEntry + "/" + x.Row.PdoLine))
                if (g.Count() > 1 && g.Sum(x => x.Row.Quantity) > g.First().Row.Balance)
                    throw new PlanException(400, "Production Order " + g.First().Row.PdoNum + " line " + g.First().Row.PdoLine + " (" + g.First().Row.ItemCode + ") is on more than one row and the quantities exceed what is still to be requested.");
        }
        #endregion

        /// <summary>Validates and creates the Purchase Request for the selected rows.</summary>
        public static CreatedRequest Create(PlanHeader h, IList<PlanRow> rows)
        {
            Validate(h, rows);
            return Store.CreatePurchaseRequest(h, rows.Where(r => r.Selected).ToList());
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Project-based bin of every GRPO row: BinCode = WhsCode + "-" + Project (e.g. "MAIN LOC-AI:2627035"), looked up
    /// in OBIN by BinCode AND WhsCode, the full row quantity going to its AbsEntry. No fallback to any other bin.
    /// Callers hold Db.DiLock.
    /// </summary>
    public static class BinAllocationService
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public static string DetermineExpectedBinCode(string warehouse, string project)
        {
            return warehouse + "-" + project;
        }

        /// <summary>Resolves the bin of every row. Returns one message per row that cannot be allocated; empty = all good.</summary>
        public static List<string> ApplyBinAllocation(IList<GrpoDraftLine> lines)
        {
            var errors = new List<string>();
            var items = ItemFlags(lines.Select(l => l.ItemCode));
            var binWarehouses = BinWarehouses(lines.Select(l => l.Warehouse));

            var expected = new List<string>();
            foreach (GrpoDraftLine l in lines)
            {
                string[] flags;
                if (!items.TryGetValue(l.ItemCode, out flags)) { errors.Add(Describe(l) + " : item not found in the Item Master."); continue; }
                if (flags[1] == "Y" || flags[2] == "Y")
                {
                    errors.Add(Describe(l) + " : " + (flags[1] == "Y" ? "batch" : "serial") +
                        "-managed items are not supported by Bin Auto Allocations. Allocate this GRPO manually.");
                    continue;
                }
                if (l.Warehouse.Length == 0) { errors.Add(Describe(l) + " : Warehouse is missing."); continue; }
                // A non-inventory item, or a warehouse without bin locations, takes no allocation at all
                l.NeedsBin = flags[0] == "Y" && binWarehouses.Contains(l.Warehouse);
                if (!l.NeedsBin) continue;
                if (l.Project.Length == 0) { errors.Add(Describe(l) + " : Project is missing, so its bin cannot be determined."); continue; }
                l.BinCode = DetermineExpectedBinCode(l.Warehouse, l.Project);
                expected.Add(l.BinCode);
            }

            var bins = Bins(expected);
            foreach (GrpoDraftLine l in lines.Where(x => x.NeedsBin && x.BinCode.Length > 0))
            {
                string error = ValidateBin(l, bins);
                if (error != null) { errors.Add(error); l.BinAbsEntry = 0; }
            }
            if (errors.Count == 0) errors.AddRange(ValidateAllocation(lines));
            return errors;
        }

        /// <summary>The expected bin must exist in the row's own warehouse and be enabled.</summary>
        public static string ValidateBin(GrpoDraftLine line, Dictionary<string, List<Dictionary<string, object>>> bins)
        {
            List<Dictionary<string, object>> found;
            if (!bins.TryGetValue(line.BinCode, out found))
                return Describe(line) + " : project bin '" + line.BinCode + "' not found.";
            var own = found.FirstOrDefault(b => string.Equals(Db.Str(b, "WhsCode"), line.Warehouse, StringComparison.OrdinalIgnoreCase));
            if (own == null)
                return Describe(line) + " : bin '" + line.BinCode + "' belongs to warehouse " + Db.Str(found[0], "WhsCode") + ", not " + line.Warehouse + ".";
            if (Db.Str(own, "Disabled") == "Y")
                return Describe(line) + " : bin '" + line.BinCode + "' is inactive.";
            line.BinAbsEntry = Convert.ToInt32(own["AbsEntry"], Inv);
            return null;
        }

        /// <summary>Every row needing a bin has one, and allocates its full inventory quantity (nothing remaining).</summary>
        public static List<string> ValidateAllocation(IEnumerable<GrpoDraftLine> lines)
        {
            var errors = new List<string>();
            foreach (GrpoDraftLine l in lines.Where(x => x.NeedsBin))
            {
                if (l.BinAbsEntry <= 0) errors.Add(Describe(l) + " : no bin allocated.");
                else if (l.InventoryQuantity <= 0) errors.Add(Describe(l) + " : quantity must be greater than zero.");
            }
            return errors;
        }

        public static string Describe(GrpoDraftLine l)
        {
            return "Row " + l.Row + " (" + l.ItemCode + ", " + (l.Warehouse.Length > 0 ? l.Warehouse : "no warehouse") +
                ", project " + (l.Project.Length > 0 ? l.Project : "none") + ")";
        }

        /// <summary>ItemCode -> [InvntItem, ManBtchNum, ManSerNum].</summary>
        private static Dictionary<string, string[]> ItemFlags(IEnumerable<string> codes)
        {
            var map = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in Db.Rows("Select \"ItemCode\", \"InvntItem\", \"ManBtchNum\", \"ManSerNum\" from OITM Where \"ItemCode\" in (" + InList(codes) + ")"))
                map[Db.Str(r, "ItemCode")] = new[] { Db.Str(r, "InvntItem"), Db.Str(r, "ManBtchNum"), Db.Str(r, "ManSerNum") };
            return map;
        }

        private static HashSet<string> BinWarehouses(IEnumerable<string> codes)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in Db.Rows("Select \"WhsCode\" from OWHS Where \"BinActivat\" = 'Y' and \"WhsCode\" in (" + InList(codes) + ")"))
                set.Add(Db.Str(r, "WhsCode"));
            return set;
        }

        /// <summary>BinCode -> every OBIN row with that code (BinCode is not assumed unique across warehouses).</summary>
        private static Dictionary<string, List<Dictionary<string, object>>> Bins(IEnumerable<string> codes)
        {
            var map = new Dictionary<string, List<Dictionary<string, object>>>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in Db.Rows("Select \"AbsEntry\", \"BinCode\", \"WhsCode\", \"Disabled\" from OBIN Where \"BinCode\" in (" + InList(codes) + ")"))
            {
                string code = Db.Str(r, "BinCode");
                List<Dictionary<string, object>> list;
                if (!map.TryGetValue(code, out list)) map[code] = list = new List<Dictionary<string, object>>();
                list.Add(r);
            }
            return map;
        }

        private static string InList(IEnumerable<string> codes)
        {
            var list = codes.Select(c => (c ?? "").Trim()).Where(c => c.Length > 0).Distinct().Select(Db.Quote).ToList();
            return list.Count == 0 ? "''" : string.Join(",", list.ToArray());
        }
    }
}

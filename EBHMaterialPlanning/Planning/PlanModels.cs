using System;
using System.Collections.Generic;

namespace EBHMaterialPlanning
{
    /// <summary>A Production Order as listed on the Material Request Planning form.</summary>
    public sealed class ProductionOrder
    {
        public bool Selected;
        public int DocEntry;
        public int DocNum;
        public int Series;
        public string SeriesName = "";
        public string Project = "";
        public string ItemCode = "";
        public string ItemName = "";
        public string Status = "";          // P planned, R released
        public string DueDate = "";         // yyyy-MM-dd
        public int TotalComponents;         // item-type (not resource) component lines of this order
        public int CompletedComponents;     // of those, how many are fully covered by open Purchase Request quantity
        /// <summary>Every (not cancelled) Purchase Request line raised from this order is Closed - shown as Completed, in green.</summary>
        public bool RequestsClosed;

        /// <summary>Not Started / Partially Completed / Completed, from CompletedComponents vs TotalComponents.</summary>
        public string ProgressState
        {
            get
            {
                if (TotalComponents <= 0 || CompletedComponents <= 0) return "None";
                return CompletedComponents >= TotalComponents ? "Full" : "Partial";
            }
        }
    }

    /// <summary>One component line of a Production Order with what is already requested on open Purchase Requests.</summary>
    public sealed class Component
    {
        public int PdoEntry;
        public int PdoNum;
        public int PdoLine;                 // WOR1.LineNum
        public string PdoItemCode = "";     // OWOR.ItemCode: the product this order is manufacturing (not the component)
        public string ItemCode = "";
        public string Description = "";
        public decimal Planned;
        public decimal OpenRequested;       // quantity on open Purchase Request lines of this component
        public string Uom = "";
        public string Warehouse = "";
        public string Project = "";         // of the source Production Order (OWOR.Project)
        public string Manufacturer = "";    // OITM.FirmCode ("" when the item has none)
        public string ManufacturerName = ""; // OMRC.FirmName ("" when the item has none)
        public string Catalog = "";         // OITM.U_EBH_MfgCatlog, copied onto the Purchase Request line
        public bool TracksJobCode;          // true when the component's Item Group (OITB.U_EBH_JOBCODNO) is set
        public decimal Price, Discount;     // WOR1.U_EBH_Unit_Price / U_EBH_Discount, when the component was manually priced there (else 0)
        public decimal PriceAfterDiscount;  // WOR1.U_EBH_PAFDI as-is (Price less Discount % when that is blank)
        public decimal Balance { get { return Math.Max(0, Planned - OpenRequested); } }
    }

    /// <summary>The five header values of the Purchase Request.</summary>
    public sealed class PlanHeader
    {
        public int Series;
        public int Branch;                  // OBPL.BplId; required only when the company uses Branches
        public string RequiredDate = "";    // yyyy-MM-dd
        public string ValidUntil = "";      // yyyy-MM-dd
        public string Vendor = "";
        public string Warehouse = "";
    }

    /// <summary>A material row of the form. Source fields are set when the row is loaded and never edited.</summary>
    public sealed class PlanRow
    {
        public bool Selected = true;
        public int PdoEntry, PdoNum, PdoLine;
        public string ItemCode = "", Description = "", Uom = "";
        /// <summary>Display only: what is still to be planned for the source component (= Balance when loaded).</summary>
        public decimal Quantity;
        /// <summary>PRQTY: the quantity that goes onto the Purchase Request (starts = Quantity).</summary>
        public decimal PrQuantity;
        /// <summary>ITRQTY: the quantity that goes onto the Inventory Transfer Request (starts = 0).</summary>
        public decimal ItrQuantity;
        /// <summary>Warehouse the Inventory Transfer Request takes the ITRQTY from (the row's Warehouse is the target; may be the same one).</summary>
        public string FromWarehouse = "";
        public decimal Balance;             // what may still be requested for the source component
        public decimal AlreadyRequested;    // quantity already on an open Purchase Request for this component
        public bool HasLinkedRequest;       // the source Production Order already has a Purchase Request raised from it (any line, any status)
        public decimal Price, Discount;
        public string Warehouse = "", Vendor = "", RequiredDate = "", ValidUntil = "";
        public string Project = "";
        public string Manufacturer = "";    // OITM.FirmCode of the source item; not saved, used only to filter/select rows on the form
        public string ManufacturerName = ""; // OMRC.FirmName of the source item, for display only
        public string Catalog = "";         // OITM.U_EBH_MfgCatlog, saved onto the Purchase Request line (PRQ1.U_EBH_CATLOG)
        public string PdoItemCode = "";     // OWOR.ItemCode of the source Production Order, saved onto the line (PRQ1.U_EBH_JOBCODNO) only when TracksJobCode
        public bool TracksJobCode;
        /// <summary>Set only on a display-merged summary row (same Catalog No. + Item No. + Vendor + Description on two
        /// or more loaded rows, combined into one grid row): the original rows it stands in for. The summary row
        /// becomes ONE Purchase Request line at Create time, and each member is recorded against that line in
        /// @EBH_PRQMAP with its own quantity. Null on an ordinary row.</summary>
        public List<PlanRow> MergedFrom;

        public decimal PriceAfterDiscount;  // WOR1.U_EBH_PAFDI of the source component (the lowest one on a merged row)
        /// <summary>Price After Discount x PRQTY - the value of what goes onto the Purchase Request.</summary>
        public decimal Total { get { return PlanService.Total(PrQuantity, PriceAfterDiscount); } }
    }

    /// <summary>A Purchase Request (DocType "PR") or Inventory Transfer Request (DocType "ITR") line raised from a
    /// Production Order component. PrEntry / PrNum / PrLine are that document's own DocEntry / DocNum / LineNum.</summary>
    public sealed class LinkedRequest
    {
        public string DocType = "PR";
        public int PrEntry, PrNum, PrLine;
        public int Series;
        public string SeriesName = "";
        public string Status = "";          // O open, C closed, X cancelled
        public string PostingDate = "";
        public string ItemCode = "", Description = "";
        public int PdoLine;
        public decimal Quantity, OpenQuantity;
    }

    /// <summary>Result of creating a Purchase Request or an Inventory Transfer Request.</summary>
    public sealed class CreatedRequest
    {
        public int DocEntry, DocNum;
        public string Warning = "";         // set when the Production Order side could not be refreshed
    }

    /// <summary>Result of one Create press on Material Request Planning: the Purchase Request and/or the Inventory
    /// Transfer Request added (null when that document was not part of the press).</summary>
    public sealed class CreatedDocuments
    {
        public CreatedRequest Pr, Itr;
        public string Warning = "";         // set when the Production Order side could not be refreshed
    }

    /// <summary>
    /// One still-open, still-orderable Purchase Request line as listed on the Purchase Order Planning form - the
    /// picker shows one row per eligible line, not one row per Purchase Request document, since a request can carry
    /// lines of more than one Project / Vendor.
    /// </summary>
    public sealed class PurchaseRequestLine
    {
        public bool Selected = true;
        public int PrEntry, PrNum, PrLine;
        public int Series;
        public string SeriesName = "";
        public string Project = "";
        public string Vendor = "";
        public string ItemCode = "", Description = "", Uom = "";
        public decimal Quantity;            // requested quantity
        public decimal OpenQuantity;         // what is still open on this line
        public decimal Price, Discount;
        public string Warehouse = "";
        public string TaxCode = "";
        public string Status = "";          // O open, C closed (of the Purchase Request document)
        public string LineStatus = "";      // O open, C closed (of this line itself, PRQ1.LineStatus)
        public string PostingDate = "";     // yyyy-MM-dd
        public string Manufacturer = "";    // OITM.FirmCode of the item; not saved, used only to filter/select rows on the form
        public string ManufacturerName = ""; // OMRC.FirmName of the item, for display only
        public string Catalog = "";         // OITM.U_EBH_MfgCatlog, for display only
        public string JobCode = "";         // PRQ1.U_EBH_JOBCODNO, for display only

        /// <summary>The document is closed, this line itself is closed, or nothing is left open on it (fully drawn
        /// into a Purchase Order already, by this or any other document) - cannot be drawn from again.</summary>
        public bool Closed { get { return Status == "C" || LineStatus == "C" || OpenQuantity <= 0; } }
    }

    /// <summary>The header values of the Purchase Order.</summary>
    public sealed class PoHeader
    {
        public int Series;
        public int Branch;                  // OBPL.BplId; required only when the company uses Branches
        public string Vendor = "";
        public string DueDate = "";         // yyyy-MM-dd, the document's delivery date
        public string VendorRef = "";       // optional, OPOR.NumAtCard
    }

    /// <summary>An open Purchase Request line, loaded into the Purchase Order grid. Source fields are set when the row is loaded and never edited.</summary>
    public sealed class PoRow
    {
        public bool Selected = true;
        public int PrEntry, PrNum, PrLine;
        public string ItemCode = "", Description = "", Uom = "";
        public decimal Quantity;
        public decimal OpenQuantity;        // what is still open on the source Purchase Request line
        public decimal Price, Discount;
        public string Warehouse = "", Vendor = "", Project = "";
        public string TaxCode = "";         // GST / VAT tax code of the source Purchase Request line (required to add the Purchase Order)
        public string Manufacturer = "";    // OITM.FirmCode of the item; not saved, used only to filter/select rows on the form
        public string ManufacturerName = ""; // OMRC.FirmName of the item, for display only
        public string Catalog = "";         // copied from the source Purchase Request line (PRQ1.U_EBH_CATLOG) onto POR1.U_EBH_CATLOG
        public string JobCode = "";         // copied from the source Purchase Request line (PRQ1.U_EBH_JOBCODNO) onto POR1.U_EBH_JOBCODNO
        /// <summary>Set only on a display-merged summary row (same Item + Manufacturer on two or more loaded lines,
        /// combined into one grid row): the original rows it stands in for, each still its own Purchase Order line
        /// at Create time (one PO line can only draw from one base Purchase Request line). Null on an ordinary row.</summary>
        public List<PoRow> MergedFrom;

        /// <summary>The highest Price After Discount this line may be ordered at: PRQ1.U_EBH_PRICEAFTERDISCOUNT of the
        /// source Purchase Request line (the Production Order's planned price). For a request raised before that field
        /// existed: the lowest WOR1.U_EBH_PAFDI of the components it covers (@EBH_PRQMAP), else the request line's own
        /// Price less Discount %. 0 = no planned price, so no limit.</summary>
        public decimal MaxPriceAfterDiscount;
        /// <summary>The grid's "#" this row is shown under (a merged row's members share it); 0 = not shown (filtered out).</summary>
        public int SrNo;
        /// <summary>Price less Discount % - both editable on the Purchase Order Planning grid.</summary>
        public decimal PriceAfterDiscount { get { return PlanService.AfterDiscount(Price, Discount); } }
        public decimal Total { get { return PlanService.Total(Quantity, PriceAfterDiscount); } }
    }

    /// <summary>
    /// Stock of the loaded items, for the In Stock / Project Stock / Main Loc Stock grid columns: OITW.OnHand per
    /// warehouse, and bin quantities (OIBQ) split into the warehouse's system bin and its project bins (OBIN.SL1Code =
    /// Project code). Holds every warehouse of each item, so a row's Warehouse can change without reloading.
    /// </summary>
    public sealed class StockMap
    {
        public const string MainLoc = "MAIN LOC";
        private readonly Dictionary<string, decimal> onHand = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, decimal> systemBin = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, decimal> projectBin = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        private static void Add(Dictionary<string, decimal> map, string key, decimal qty)
        {
            decimal q;
            map.TryGetValue(key, out q);
            map[key] = q + qty;
        }

        private static decimal Get(Dictionary<string, decimal> map, string key)
        {
            decimal q;
            return map.TryGetValue(key, out q) ? q : 0;
        }

        public void AddOnHand(string item, string whs, decimal qty) { Add(onHand, item + "|" + whs, qty); }
        public void AddSystemBin(string item, string whs, decimal qty) { Add(systemBin, item + "|" + whs, qty); }
        public void AddProjectBin(string item, string whs, string project, decimal qty) { Add(projectBin, item + "|" + whs + "|" + project, qty); }

        /// <summary>OITW.OnHand of the item in the given warehouse.</summary>
        public decimal InStock(string item, string whs) { return Get(onHand, item + "|" + whs); }
        /// <summary>Quantity in the warehouse's bins assigned to the given Project (OBIN.SL1Code).</summary>
        public decimal ProjectStock(string item, string whs, string project) { return project.Length == 0 ? 0 : Get(projectBin, item + "|" + whs + "|" + project); }
        /// <summary>Quantity in the system bin of the MAIN LOC warehouse, whatever the row's own warehouse is.</summary>
        public decimal MainLocStock(string item) { return Get(systemBin, item + "|" + MainLoc); }
    }

    /// <summary>Result of creating a Purchase Order.</summary>
    public sealed class CreatedOrder
    {
        public int DocEntry, DocNum;
    }
}

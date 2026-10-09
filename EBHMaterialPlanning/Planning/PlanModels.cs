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
    }

    /// <summary>One component line of a Production Order with what is already requested on open Purchase Requests.</summary>
    public sealed class Component
    {
        public int PdoEntry;
        public int PdoNum;
        public int PdoLine;                 // WOR1.LineNum
        public string ItemCode = "";
        public string Description = "";
        public decimal Planned;
        public decimal OpenRequested;       // quantity on open Purchase Request lines of this component
        public string Uom = "";
        public string Warehouse = "";
        public string Project = "";         // of the source Production Order (OWOR.Project)
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
        public decimal Quantity;
        public decimal Balance;             // what may still be requested for the source component
        public decimal AlreadyRequested;    // quantity already on an open Purchase Request for this component
        public bool HasLinkedRequest;       // the source Production Order already has a Purchase Request raised from it (any line, any status)
        public decimal Price, Discount;
        public string Warehouse = "", Vendor = "", RequiredDate = "", ValidUntil = "";
        public string Project = "";

        public decimal PriceAfterDiscount { get { return PlanService.AfterDiscount(Price, Discount); } }
        public decimal Total { get { return PlanService.Total(Quantity, Price, Discount); } }
    }

    /// <summary>A Purchase Request line that was raised from a Production Order component.</summary>
    public sealed class LinkedRequest
    {
        public int PrEntry, PrNum, PrLine;
        public int Series;
        public string SeriesName = "";
        public string Status = "";          // O open, C closed, X cancelled
        public string PostingDate = "";
        public string ItemCode = "", Description = "";
        public int PdoLine;
        public decimal Quantity, OpenQuantity;
    }

    /// <summary>Result of creating a Purchase Request.</summary>
    public sealed class CreatedRequest
    {
        public int DocEntry, DocNum;
        public string Warning = "";         // set when the Production Order side could not be refreshed
    }
}

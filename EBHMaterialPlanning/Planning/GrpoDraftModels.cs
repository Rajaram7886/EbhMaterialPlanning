using System.Collections.Generic;

namespace EBHMaterialPlanning
{
    /// <summary>A user-defined field read off the GRPO form, already converted to the DI API type.</summary>
    public sealed class UdfValue
    {
        public string Name = "";
        public object Value;
    }

    /// <summary>Header of the unsaved Goods Receipt PO, read from the form's OPDN data source.</summary>
    public sealed class GrpoDraftHeader
    {
        public int Series;
        public string CardCode = "", VendorRef = "", Comments = "", JournalMemo = "", Currency = "", Project = "", ShipTo = "";
        public string DocDate = "", DueDate = "", TaxDate = "";         // yyyyMMdd, as SAP keeps them in the data source
        public int Branch, SalesPerson = -1, Owner, ContactPerson;
        public decimal DocRate, DiscountPercent;
        public readonly List<UdfValue> Udfs = new List<UdfValue>();
    }

    /// <summary>One row of the unsaved Goods Receipt PO (PDN1), plus the bin chosen for it.</summary>
    public sealed class GrpoDraftLine
    {
        /// <summary>1-based row number as the user sees it on the GRPO.</summary>
        public int Row;
        public string ItemCode = "", Description = "", Warehouse = "", Project = "", TaxCode = "", FreeText = "";
        public readonly string[] Costing = new string[5];
        public decimal Quantity, InventoryQuantity, UnitPrice, DiscountPercent;
        public int BaseType = -1, BaseEntry, BaseLine = -1, UomEntry = -1;
        public readonly List<UdfValue> Udfs = new List<UdfValue>();

        /// <summary>False for a non-inventory item or a warehouse without bin locations: no allocation needed.</summary>
        public bool NeedsBin;
        public string BinCode = "";
        public int BinAbsEntry;

        public bool CopiedFromBase { get { return BaseType > 0 && BaseEntry > 0 && BaseLine >= 0; } }
    }

    public sealed class GrpoDraft
    {
        public readonly GrpoDraftHeader Header = new GrpoDraftHeader();
        public readonly List<GrpoDraftLine> Lines = new List<GrpoDraftLine>();
        /// <summary>False when the rows' Purchase Order link could not be read off the form: the GRPO is then not added (it would leave the PO open).</summary>
        public bool BaseLinkRead = true;
    }

    public sealed class CreatedDocument
    {
        public int DocEntry, DocNum;
    }
}

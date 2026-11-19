using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Adds (posts) the unsaved Goods Receipt PO on screen through the DI API (OPDN, ObjType 20), with the project bin
    /// allocation of every row - never opens SAP's Bin Location Allocation window.
    /// PrepareDraft reads the form (UI API, SAP UI thread); AddGrpo needs Db.DiLock.
    /// </summary>
    public static class GrpoDraftService
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        private const string HeaderSource = "OPDN", LineSource = "PDN1";

        #region Read the form
        /// <summary>Header and rows from the form's own data sources - the GRPO is not in the database yet.</summary>
        public static GrpoDraft PrepareDraft(SAPbouiCOM.Form form)
        {
            var draft = new GrpoDraft();
            SAPbouiCOM.DBDataSource head = form.DataSources.DBDataSources.Item(HeaderSource);
            SAPbouiCOM.DBDataSource rows = form.DataSources.DBDataSources.Item(LineSource);

            GrpoDraftHeader h = draft.Header;
            h.Series = Int(Get(head, "Series", 0));
            h.CardCode = Get(head, "CardCode", 0);
            h.VendorRef = Get(head, "NumAtCard", 0);
            h.DocDate = Get(head, "DocDate", 0);
            h.DueDate = Get(head, "DocDueDate", 0);
            h.TaxDate = Get(head, "TaxDate", 0);
            h.Comments = Get(head, "Comments", 0);
            h.JournalMemo = Get(head, "JrnlMemo", 0);
            h.Currency = Get(head, "DocCur", 0);
            h.DocRate = Dec(Get(head, "DocRate", 0));
            h.DiscountPercent = Dec(Get(head, "DiscPrcnt", 0));
            h.Project = Get(head, "Project", 0);
            h.ShipTo = Get(head, "ShipToCode", 0);
            h.Branch = Int(Get(head, "BPLId", 0));
            h.SalesPerson = Int(Get(head, "SlpCode", 0), -1);
            h.Owner = Int(Get(head, "OwnerCode", 0));
            h.ContactPerson = Int(Get(head, "CntctCode", 0));
            ReadHeaderUdfs(head, h.Udfs);
            if (h.CardCode.Length == 0) throw new PlanException(400, "Select the Vendor before Bin Auto Allocations.");

            // On a GRPO in Add mode SAP keeps the rows in the matrix only - the PDN1 data source stays empty until the
            // document is saved (seen 2026-10-30: "has no item rows" on a 5-row GRPO). Read PDN1 when it is filled, else the matrix.
            LineReader reader = LineReader.FromDataSource(rows);
            if (reader == null) reader = LineReader.FromMatrix((SAPbouiCOM.Matrix)form.Items.Item(RowMatrixUID).Specific, rows);

            int visible = 0;
            for (int i = reader.First; i <= reader.Last; i++)
            {
                string itemCode = reader.Text("ItemCode", i);
                if (itemCode.Length == 0) continue;            // the empty last row of the matrix
                var l = new GrpoDraftLine { Row = ++visible, ItemCode = itemCode };
                l.Description = reader.Text("Dscription", i);
                l.Warehouse = reader.Text("WhsCode", i);
                l.Project = reader.Text("Project", i);
                l.TaxCode = reader.Text("TaxCode", i);
                l.FreeText = reader.Text("FreeTxt", i);
                l.Quantity = reader.Number("Quantity", i);
                l.UnitPrice = reader.Number("PriceBefDi", i);
                l.DiscountPercent = reader.Number("DiscPrcnt", i);
                l.BaseType = reader.Int("BaseType", i, -1);
                l.BaseEntry = reader.Int("BaseEntry", i, 0);
                l.BaseLine = reader.Int("BaseLine", i, -1);
                l.UomEntry = reader.Int("UomEntry", i, -1);
                l.Costing[0] = reader.Text("OcrCode", i);
                l.Costing[1] = reader.Text("OcrCode2", i);
                l.Costing[2] = reader.Text("OcrCode3", i);
                l.Costing[3] = reader.Text("OcrCode4", i);
                l.Costing[4] = reader.Text("OcrCode5", i);
                // Bins hold the inventory unit: InvQty when SAP filled it, else Quantity x Items per Unit
                decimal invQty = reader.Number("InvQty", i);
                decimal perUnit = reader.Number("NumPerMsr", i);
                l.InventoryQuantity = invQty > 0 ? invQty : l.Quantity * (perUnit > 0 ? perUnit : 1);
                reader.Udfs(i, l.Udfs);
                if (l.Quantity <= 0) throw new PlanException(400, "Row " + l.Row + " (" + l.ItemCode + "): enter a quantity greater than zero.");
                draft.Lines.Add(l);
            }
            if (draft.Lines.Count == 0) throw new PlanException(400, "The Goods Receipt PO has no item rows.");
            draft.BaseLinkRead = reader.HasBaseLink;
            return draft;
        }

        private const string RowMatrixUID = "38";

        /// <summary>
        /// Row values by PDN1 field name, from the data source (database format) or from the matrix (display format).
        /// The matrix is read in ONE call (SerializeAsXML): reading it cell by cell is a cross-process UI API call per
        /// cell and took minutes for 5 rows (2026-10-30).
        /// </summary>
        private sealed class LineReader
        {
            private SAPbouiCOM.DBDataSource ds;
            /// <summary>Matrix rows (index 0 = row 1): column UID -> cell text. Null when reading the data source.</summary>
            private List<Dictionary<string, string>> cells;
            /// <summary>PDN1 field -> matrix column UID, from each column's DataBind.</summary>
            private Dictionary<string, string> columns;
            /// <summary>UDF columns of the matrix: name -> type, looked up once rather than per row.</summary>
            private List<KeyValuePair<string, SAPbouiCOM.BoFieldsType>> udfFields;
            public int First, Last;

            /// <summary>The GRPO matrix layout does not change within a session: mapped once (the column count guards Form Settings changes).</summary>
            private static Dictionary<string, string> cachedColumns;
            private static int cachedColumnCount = -1;

            public static LineReader FromDataSource(SAPbouiCOM.DBDataSource ds)
            {
                for (int i = 0; i < ds.Size; i++)
                    if (Get(ds, "ItemCode", i).Length > 0) return new LineReader { ds = ds, First = 0, Last = ds.Size - 1 };
                return null;
            }

            public static LineReader FromMatrix(SAPbouiCOM.Matrix matrix, SAPbouiCOM.DBDataSource ds)
            {
                var r = new LineReader { ds = ds };
                var xml = new System.Xml.XmlDocument();
                xml.LoadXml(matrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All));

                r.columns = ColumnsFromXml(xml);
                if (!r.columns.ContainsKey("ItemCode")) r.columns = ColumnsFromMatrix(matrix);
                if (!r.columns.ContainsKey("ItemCode"))
                    throw new PlanException(500, "The Item No. column of the Goods Receipt PO could not be read.");

                r.cells = new List<Dictionary<string, string>>();
                foreach (System.Xml.XmlNode row in xml.SelectNodes("//Rows/Row"))
                {
                    var values = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (System.Xml.XmlNode col in row.SelectNodes("Columns/Column"))
                    {
                        System.Xml.XmlNode id = col.SelectSingleNode("ID"), value = col.SelectSingleNode("Value");
                        if (id != null) values[id.InnerText.Trim()] = value == null ? "" : value.InnerText.Trim();
                    }
                    r.cells.Add(values);
                }
                r.First = 1;
                r.Last = r.cells.Count;

                r.udfFields = new List<KeyValuePair<string, SAPbouiCOM.BoFieldsType>>();
                foreach (string alias in r.columns.Keys)
                {
                    if (!alias.StartsWith("U_", StringComparison.OrdinalIgnoreCase)) continue;
                    try { r.udfFields.Add(new KeyValuePair<string, SAPbouiCOM.BoFieldsType>(alias, ds.Fields.Item(alias).Type)); }
                    catch (COMException) { }        // bound to a field this data source does not define
                }
                return r;
            }

            /// <summary>Column bindings straight from the serialised matrix, when SAP includes them (no extra UI calls).</summary>
            private static Dictionary<string, string> ColumnsFromXml(System.Xml.XmlDocument xml)
            {
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (System.Xml.XmlNode info in xml.SelectNodes("//ColumnsInfo/ColumnInfo"))
                {
                    System.Xml.XmlNode uid = info.SelectSingleNode("UniqueID"), table = info.SelectSingleNode("DataBind/TableName"),
                        alias = info.SelectSingleNode("DataBind/Alias");
                    if (uid == null || table == null || alias == null) continue;
                    string a = alias.InnerText.Trim();
                    if (a.Length > 0 && string.Equals(table.InnerText.Trim(), LineSource, StringComparison.OrdinalIgnoreCase) && !map.ContainsKey(a))
                        map[a] = uid.InnerText.Trim();
                }
                return map;
            }

            private static Dictionary<string, string> ColumnsFromMatrix(SAPbouiCOM.Matrix matrix)
            {
                int count = matrix.Columns.Count;
                if (cachedColumns != null && cachedColumnCount == count) return cachedColumns;
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < count; i++)
                {
                    SAPbouiCOM.Column c = matrix.Columns.Item(i);
                    try
                    {
                        SAPbouiCOM.DataBind bind = c.DataBind;
                        string alias = (bind.Alias ?? "").Trim();
                        if (alias.Length > 0 && string.Equals((bind.TableName ?? "").Trim(), LineSource, StringComparison.OrdinalIgnoreCase)
                            && !map.ContainsKey(alias)) map[alias] = c.UniqueID;
                    }
                    catch (COMException) { }        // unbound column (row number etc.)
                }
                cachedColumns = map;
                cachedColumnCount = count;
                return map;
            }

            public string Text(string field, int row)
            {
                if (cells == null) return Get(ds, field, row);
                string uid, value;
                if (!columns.TryGetValue(field, out uid) || row < 1 || row > cells.Count) return "";
                return cells[row - 1].TryGetValue(uid, out value) ? value : "";
            }

            /// <summary>The matrix shows amounts with currency and the client's separators (e.g. "INR 35,598.36").</summary>
            public decimal Number(string field, int row)
            {
                string text = Text(field, row);
                if (cells == null) return Dec(text);
                var keep = new System.Text.StringBuilder();
                foreach (char ch in text) if (char.IsDigit(ch) || ch == '.' || ch == ',' || ch == '-') keep.Append(ch);
                decimal d;
                string s = keep.ToString();
                if (decimal.TryParse(s, NumberStyles.Number, CultureInfo.CurrentCulture, out d)) return d;
                return decimal.TryParse(s, NumberStyles.Number, Inv, out d) ? d : 0;
            }

            public int Int(string field, int row, int fallback)
            {
                return Text(field, row).Length == 0 ? fallback : (int)Number(field, row);
            }

            /// <summary>False when the matrix has no columns bound to the Purchase Order link (BaseEntry / BaseLine).</summary>
            public bool HasBaseLink { get { return cells == null || (columns.ContainsKey("BaseEntry") && columns.ContainsKey("BaseLine")); } }

            public void Udfs(int row, List<UdfValue> target)
            {
                if (cells == null) { ReadUdfs(ds, row, target); return; }
                foreach (var f in udfFields)
                {
                    string text = Text(f.Key, row);
                    if (text.Length == 0) continue;
                    object value = ToFieldValue(f.Value, text, Number(f.Key, row));
                    if (value != null) target.Add(new UdfValue { Name = f.Key, Value = value });
                }
            }
        }

        private static string Get(SAPbouiCOM.DBDataSource ds, string field, int row)
        {
            try { return (ds.GetValue(field, row) ?? "").Trim(); }
            catch (COMException) { return ""; }        // field not in this SAP version's data source
        }

        /// <summary>Every filled U_ field, converted from the data source's text to the field's DI API type.</summary>
        private static void ReadUdfs(SAPbouiCOM.DBDataSource ds, int row, List<UdfValue> target)
        {
            for (int i = 0; i < ds.Fields.Count; i++)
            {
                SAPbouiCOM.Field f = ds.Fields.Item(i);
                if (!f.Name.StartsWith("U_", StringComparison.OrdinalIgnoreCase)) continue;
                string text = Get(ds, f.Name, row);
                if (text.Length == 0) continue;
                object value = ToFieldValue(f.Type, text, Dec(text));
                if (value != null) target.Add(new UdfValue { Name = f.Name, Value = value });
            }
        }

        /// <summary>
        /// Header UDFs from one GetAsXML call; walking the ~300 OPDN fields one by one costs two UI API calls per field.
        /// Falls back to that walk when the XML holds no cells.
        /// </summary>
        private static void ReadHeaderUdfs(SAPbouiCOM.DBDataSource ds, List<UdfValue> target)
        {
            var xml = new System.Xml.XmlDocument();
            try { xml.LoadXml(ds.GetAsXML()); }
            catch (Exception ex) when (ex is COMException || ex is System.Xml.XmlException)
            {
                System.Diagnostics.Trace.WriteLine("EBH Material Planning GRPO draft : header XML unreadable, reading field by field : " + ex.Message);
                ReadUdfs(ds, 0, target);
                return;
            }
            System.Xml.XmlNodeList rows = xml.SelectNodes("//row");
            System.Xml.XmlNodeList cells = rows.Count == 0 ? null : rows[0].SelectNodes(".//cell");
            if (cells == null || cells.Count == 0) { ReadUdfs(ds, 0, target); return; }
            foreach (System.Xml.XmlNode cell in cells)
            {
                System.Xml.XmlNode uid = cell.SelectSingleNode("uid"), value = cell.SelectSingleNode("value");
                if (uid == null || value == null) continue;
                string name = uid.InnerText.Trim(), text = value.InnerText.Trim();
                if (!name.StartsWith("U_", StringComparison.OrdinalIgnoreCase) || text.Length == 0) continue;
                object v = ToFieldValue(ds.Fields.Item(name).Type, text, Dec(text));
                if (v != null) target.Add(new UdfValue { Name = name, Value = v });
            }
        }

        /// <summary>A UDF's text as its DI API type; null for a date that cannot be read (left empty on the draft).</summary>
        private static object ToFieldValue(SAPbouiCOM.BoFieldsType type, string text, decimal number)
        {
            switch (type)
            {
                case SAPbouiCOM.BoFieldsType.ft_Date:
                    DateTime d;
                    if (DateTime.TryParseExact(text, "yyyyMMdd", Inv, DateTimeStyles.None, out d)) return d;
                    if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out d)) return d;
                    System.Diagnostics.Trace.WriteLine("EBH Material Planning GRPO draft : date '" + text + "' not understood, left empty");
                    return null;
                case SAPbouiCOM.BoFieldsType.ft_Integer:
                case SAPbouiCOM.BoFieldsType.ft_ShortNumber:
                    return (int)number;
                case SAPbouiCOM.BoFieldsType.ft_Float:
                case SAPbouiCOM.BoFieldsType.ft_Quantity:
                case SAPbouiCOM.BoFieldsType.ft_Price:
                case SAPbouiCOM.BoFieldsType.ft_Rate:
                case SAPbouiCOM.BoFieldsType.ft_Measure:
                case SAPbouiCOM.BoFieldsType.ft_Sum:
                case SAPbouiCOM.BoFieldsType.ft_Percent:
                    return (double)number;
                default:
                    return text;
            }
        }
        #endregion

        #region Add the GRPO
        /// <summary>
        /// Adds the GRPO with one bin allocation per row that needs one, inside a DI transaction: the document is read
        /// back before commit and rolled back unless every row carries its full quantity in its project bin, so a GRPO is
        /// posted completely or not at all. Throws with SAP's message (or the verification problems) otherwise.
        /// </summary>
        public static CreatedDocument AddGrpo(GrpoDraft draft)
        {
            GrpoDraftHeader h = draft.Header;
            string localCurrency = Db.Str(Db.FirstRow("Select \"MainCurncy\" from OADM"), "MainCurncy");
            SAPbobsCOM.Documents oDoc = null;
            bool inTransaction = false;
            try
            {
                if (Db.Company.InTransaction) throw new PlanException(409, "SAP is busy with another transaction. Try again in a moment.");
                Db.Company.StartTransaction();
                inTransaction = true;

                oDoc = (SAPbobsCOM.Documents)Db.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                if (h.Series > 0) oDoc.Series = h.Series;
                // The branch property is declared on IDocuments, not on the Documents interface GetBusinessObject returns
                if (h.Branch > 0) ((SAPbobsCOM.IDocuments)oDoc).BPL_IDAssignedToInvoice = h.Branch;
                oDoc.CardCode = h.CardCode;
                if (h.VendorRef.Length > 0) oDoc.NumAtCard = h.VendorRef;
                if (h.DocDate.Length > 0) oDoc.DocDate = Date(h.DocDate);
                if (h.DueDate.Length > 0) oDoc.DocDueDate = Date(h.DueDate);
                if (h.TaxDate.Length > 0) oDoc.TaxDate = Date(h.TaxDate);
                if (h.Comments.Length > 0) oDoc.Comments = h.Comments;
                if (h.JournalMemo.Length > 0) oDoc.JournalMemo = h.JournalMemo;
                if (h.Currency.Length > 0) oDoc.DocCurrency = h.Currency;
                if (h.Currency.Length > 0 && !string.Equals(h.Currency, localCurrency, StringComparison.OrdinalIgnoreCase) && h.DocRate > 0)
                    oDoc.DocRate = (double)h.DocRate;
                if (h.DiscountPercent != 0) oDoc.DiscountPercent = (double)h.DiscountPercent;
                if (h.Project.Length > 0) ((SAPbobsCOM.IDocuments)oDoc).Project = h.Project;
                if (h.ShipTo.Length > 0) oDoc.ShipToCode = h.ShipTo;
                if (h.SalesPerson >= 0) oDoc.SalesPersonCode = h.SalesPerson;
                if (h.Owner > 0) oDoc.DocumentsOwner = h.Owner;
                if (h.ContactPerson > 0) oDoc.ContactPersonCode = h.ContactPerson;
                SetUdfs(oDoc.UserFields.Fields, h.Udfs, "header");

                for (int i = 0; i < draft.Lines.Count; i++)
                {
                    GrpoDraftLine l = draft.Lines[i];
                    if (i > 0) oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(i);
                    if (l.CopiedFromBase)
                    {
                        // Keeps the link to the Purchase Order line, so this receipt closes that PO quantity
                        oDoc.Lines.BaseType = l.BaseType;
                        oDoc.Lines.BaseEntry = l.BaseEntry;
                        oDoc.Lines.BaseLine = l.BaseLine;
                    }
                    else
                    {
                        oDoc.Lines.ItemCode = l.ItemCode;
                        if (l.UomEntry > 0) oDoc.Lines.UoMEntry = l.UomEntry;
                    }
                    if (l.Description.Length > 0) oDoc.Lines.ItemDescription = l.Description;
                    oDoc.Lines.Quantity = (double)l.Quantity;
                    oDoc.Lines.WarehouseCode = l.Warehouse;
                    // A row copied from the PO already has the PO's price: only override it with a value actually read off the form
                    if (l.UnitPrice > 0 || !l.CopiedFromBase) oDoc.Lines.UnitPrice = (double)l.UnitPrice;
                    if (l.DiscountPercent != 0 || !l.CopiedFromBase) oDoc.Lines.DiscountPercent = (double)l.DiscountPercent;
                    if (l.Project.Length > 0) ((SAPbobsCOM.IDocument_Lines)oDoc.Lines).ProjectCode = l.Project;
                    // Set after Warehouse / Price, which would otherwise make SAP re-determine the GST code
                    if (l.TaxCode.Length > 0) ((SAPbobsCOM.IDocument_Lines)oDoc.Lines).TaxCode = l.TaxCode;
                    if (l.Costing[0].Length > 0) oDoc.Lines.CostingCode = l.Costing[0];
                    if (l.Costing[1].Length > 0) oDoc.Lines.CostingCode2 = l.Costing[1];
                    if (l.Costing[2].Length > 0) oDoc.Lines.CostingCode3 = l.Costing[2];
                    if (l.Costing[3].Length > 0) oDoc.Lines.CostingCode4 = l.Costing[3];
                    if (l.Costing[4].Length > 0) oDoc.Lines.CostingCode5 = l.Costing[4];
                    if (l.FreeText.Length > 0) oDoc.Lines.FreeText = l.FreeText;
                    SetUdfs(oDoc.Lines.UserFields.Fields, l.Udfs, "row " + l.Row);

                    if (l.NeedsBin)
                    {
                        // The line's first allocation record exists already: fill it, no BinAllocations.Add() (that would leave an empty second record)
                        SAPbobsCOM.DocumentLinesBinAllocations bin = oDoc.Lines.BinAllocations;
                        bin.BinAbsEntry = l.BinAbsEntry;
                        bin.Quantity = (double)l.InventoryQuantity;
                        bin.AllowNegativeQuantity = SAPbobsCOM.BoYesNoEnum.tNO;
                        bin.SerialAndBatchNumbersBaseLine = -1;
                    }
                }

                Check(oDoc.Add(), "Goods Receipt PO was not added");
                var created = new CreatedDocument { DocEntry = int.Parse(Db.Company.GetNewObjectKey(), Inv) };

                List<string> check = ValidateGrpo(created.DocEntry, draft);
                if (check.Count > 0)
                    throw new PlanException(500, "Goods Receipt PO not added - the bin allocation did not come through:" + Environment.NewLine +
                        string.Join(Environment.NewLine, check.ToArray()));
                created.DocNum = Int(Db.Str(Db.FirstRow("Select \"DocNum\" from OPDN Where \"DocEntry\" = " + created.DocEntry), "DocNum"));

                Db.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                inTransaction = false;
                return created;
            }
            catch
            {
                if (inTransaction && Db.Company.InTransaction) Db.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                throw;
            }
            finally { Release(oDoc); }
        }

        private static void SetUdfs(SAPbobsCOM.Fields fields, List<UdfValue> udfs, string where)
        {
            foreach (UdfValue u in udfs)
            {
                try { fields.Item(u.Name).Value = u.Value; }
                catch (COMException ex)
                {
                    throw new PlanException(500, "User field " + u.Name + " (" + where + ") could not be copied to the Goods Receipt PO: " + ex.Message);
                }
            }
        }
        #endregion

        #region Verify
        /// <summary>
        /// Reads the added GRPO back (still inside the transaction): same rows, and every row needing a bin carries exactly its inventory quantity in
        /// the expected bin. Returns the problems found; empty = verified.
        /// </summary>
        private static List<string> ValidateGrpo(int docEntry, GrpoDraft draft)
        {
            var errors = new List<string>();
            SAPbobsCOM.Documents oDoc = null;
            try
            {
                oDoc = (SAPbobsCOM.Documents)Db.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                if (!oDoc.GetByKey(docEntry)) { errors.Add("Goods Receipt PO " + docEntry + " could not be read back."); return errors; }
                if (oDoc.Lines.Count != draft.Lines.Count)
                {
                    errors.Add("The added document has " + oDoc.Lines.Count + " row(s), the GRPO has " + draft.Lines.Count + ".");
                    return errors;
                }
                for (int i = 0; i < draft.Lines.Count; i++)
                {
                    GrpoDraftLine l = draft.Lines[i];
                    if (!l.NeedsBin) continue;
                    oDoc.Lines.SetCurrentLine(i);
                    SAPbobsCOM.DocumentLinesBinAllocations bins = oDoc.Lines.BinAllocations;
                    decimal allocated = 0;
                    bool otherBin = false;
                    for (int j = 0; j < bins.Count; j++)
                    {
                        bins.SetCurrentLine(j);
                        if (bins.BinAbsEntry <= 0) continue;
                        if (bins.BinAbsEntry != l.BinAbsEntry) otherBin = true;
                        allocated += (decimal)bins.Quantity;
                    }
                    if (otherBin) errors.Add(BinAllocationService.Describe(l) + " : allocated to a bin other than " + l.BinCode + ".");
                    if (Math.Abs(allocated - l.InventoryQuantity) > 0.000001m)
                        errors.Add(BinAllocationService.Describe(l) + " : required " + l.InventoryQuantity.ToString("0.######", Inv) +
                            ", allocated " + allocated.ToString("0.######", Inv) + ", remaining " + (l.InventoryQuantity - allocated).ToString("0.######", Inv) + ".");
                }
            }
            finally { Release(oDoc); }
            return errors;
        }

        #endregion

        #region Helpers
        private static void Check(int result, string what)
        {
            if (result == 0) return;
            int code; string msg;
            Db.Company.GetLastError(out code, out msg);
            throw new PlanException(500, what + " : " + msg + " (" + code + ")");
        }

        private static void Release(object com)
        {
            if (com != null) Marshal.ReleaseComObject(com);
        }

        private static DateTime Date(string yyyymmdd)
        {
            return DateTime.ParseExact(yyyymmdd, "yyyyMMdd", Inv);
        }

        private static decimal Dec(string text)
        {
            decimal d;
            return decimal.TryParse(text, NumberStyles.Number, Inv, out d) ? d : 0;
        }

        private static int Int(string text, int fallback = 0)
        {
            decimal d;
            return decimal.TryParse(text, NumberStyles.Number, Inv, out d) ? (int)d : fallback;
        }
        #endregion
    }
}

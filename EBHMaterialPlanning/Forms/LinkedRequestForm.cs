using System;
using System.Collections.Generic;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Purchase Requests of one Production Order: every request line raised from it, with the request status.
    /// Shortcuts: F5 refreshes, Enter opens the selected Purchase Request, Esc closes the window.
    /// "Back to Material Planning" returns to (or opens) the main planning screen.
    /// </summary>
    internal sealed class LinkedRequestForm : PlanForm
    {
        private GridTable gList;
        private List<LinkedRequest> lines = new List<LinkedRequest>();
        private int pdoEntry, pdoNum;

        /// <summary>One grid row: either a document group header (Line == null) or one of its detail lines. DocType
        /// ("PR" / "ITR") + PrEntry name the document: the two kinds have their own DocEntry numbering.</summary>
        private sealed class DisplayRow
        {
            public string DocType = "PR";
            public int PrEntry;
            public LinkedRequest Line;      // null for a group header row
        }

        private static string DocName(string docType) { return docType == "ITR" ? "Inventory Transfer Request" : "Purchase Request"; }

        private static void OpenDocument(DisplayRow d)
        {
            if (d.DocType == "ITR") Store.OpenTransferRequest(d.PrEntry);
            else Store.OpenPurchaseRequest(d.PrEntry);
        }
        private List<DisplayRow> display = new List<DisplayRow>();

        private static readonly Dictionary<string, string> StatusNames = new Dictionary<string, string> { { "O", "Open" }, { "C", "Closed" }, { "X", "Cancelled" } };
        private static string StatusName(string code) { return StatusNames.ContainsKey(code) ? StatusNames[code] : code; }

        public static void Show(int pdoEntry)
        {
            RunShow<LinkedRequestForm>(() =>
            {
                LinkedRequestForm open = OpenForms<LinkedRequestForm>().FirstOrDefault(x => x.pdoEntry == pdoEntry);
                if (open != null) { open.Select(); open.Reload(); return; }
                var f = new LinkedRequestForm { pdoEntry = pdoEntry, pdoNum = Store.PdoNum(pdoEntry) };
                f.Create("EBHMRPLNK", "Requests of Production Order " + f.pdoNum, 880, 360, false, f.Build);
            });
        }

        private void Build()
        {
            gList = Grid("gList", 8, 8, 864, 300)
                .Column("Sr", "#", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Type", "Type", width: 40)
                .Column("Series", "Series", width: 80)
                .Column("PR", "Document No.", SAPbouiCOM.BoFieldsType.ft_Integer, 80)
                .Column("Line", "Row", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Date", "Date", SAPbouiCOM.BoFieldsType.ft_Date, 70)
                .Column("Stat", "Status", width: 70)
                .Column("Item", "Item No.", width: 100)
                .Column("Desc", "Description", width: 200)
                .Column("PdoLine", "Component Line", SAPbouiCOM.BoFieldsType.ft_Integer, 90)
                .Column("Qty", "Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 80)
                .Column("Open", "Open Qty", SAPbouiCOM.BoFieldsType.ft_Quantity, 80)
                .Sortable("Item", "Desc", "Stat", "Date")
                .Bind();
            SortWith(gList, ShowList);
            Button("bOpen", "Open Document", 8, 320, 140);
            Button("bRefresh", "Refresh", 152, 320, 70);
            Button("2", "Close", 226, 320, 70);
            Button("bBack", "Back to Material Planning", 304, 320, 170);
            Label("lFoot", "", 484, 323, 330);
            Reload();
        }

        protected override void Layout()
        {
            int w = Math.Max(560, Form.ClientWidth), h = Math.Max(160, Form.ClientHeight);
            Place("gList", 8, 8, w - 16, h - 8 - 40);
            Place("bOpen", 8, h - 30, 0, 0);
            Place("bRefresh", 152, h - 30, 0, 0);
            Place("2", 226, h - 30, 0, 0);
            Place("bBack", 304, h - 30, 0, 0);
            Place("lFoot", 484, h - 27, 0, 0);
        }

        public void Reload()
        {
            // A request cancelled or closed since the last visit frees its quantity: the Production Order is refreshed first
            try { Store.Sync(new[] { pdoEntry }); }
            catch (Exception ex) { Warn("Production Order not refreshed : " + ex.Message); }
            lines = Store.LinkedRequests(pdoEntry);
            ShowList();
        }

        private void ShowList()
        {
            // One banner row per Purchase Request ("Series / No.") above its detail lines; lines already arrive
            // grouped (sorted by DocNum then LineNum), so a change of PrEntry marks the start of the next group.
            // A title sort (Item No. / Description / Status / Date) orders the lines within each Purchase Request, never across them.
            // Purchase Requests first, then Inventory Transfer Requests, each document's lines under its own banner
            var ordered = lines.GroupBy(l => l.DocType + ":" + l.PrEntry)
                .OrderBy(g => g.First().DocType == "PR" ? 0 : 1).ThenBy(g => g.First().PrNum)
                .SelectMany(g => gList.Sort(g.ToList(), (c, l) => c == "Item" ? l.ItemCode : c == "Desc" ? l.Description : c == "Stat" ? StatusName(l.Status) : c == "Date" ? l.PostingDate : ""));
            display = new List<DisplayRow>();
            string lastKey = "";
            foreach (LinkedRequest l in ordered)
            {
                string key = l.DocType + ":" + l.PrEntry;
                if (key != lastKey)
                {
                    display.Add(new DisplayRow { DocType = l.DocType, PrEntry = l.PrEntry });
                    lastKey = key;
                }
                display.Add(new DisplayRow { DocType = l.DocType, PrEntry = l.PrEntry, Line = l });
            }

            Freeze(() =>
            {
                gList.Fill(display.Select((d, i) =>
                {
                    // The group header carries no LinkedRequest of its own; borrow Series / Number from its group's first line.
                    LinkedRequest first = d.Line ?? lines.First(l => l.DocType == d.DocType && l.PrEntry == d.PrEntry);
                    return d.Line == null
                        ? new object[] { i + 1, d.DocType, first.SeriesName, first.PrNum, 0, ToDate(first.PostingDate), "", "", DocName(d.DocType) + " " + first.SeriesName + " " + first.PrNum, 0, 0m, 0m }
                        : new object[]
                        {
                            i + 1, d.DocType, d.Line.SeriesName, d.Line.PrNum, d.Line.PrLine, ToDate(d.Line.PostingDate),
                            StatusName(d.Line.Status),
                            d.Line.ItemCode, d.Line.Description, d.Line.PdoLine, d.Line.Quantity, d.Line.OpenQuantity
                        };
                }).ToList());
                for (int i = 0; i < display.Count; i++)
                    gList.Color(i, display[i].Line == null ? GridTable.Yellow : display[i].Line.Status == "O" ? GridTable.Green : GridTable.Grey);
                int prCount = lines.Where(l => l.DocType == "PR").Select(l => l.PrEntry).Distinct().Count();
                int itrCount = lines.Where(l => l.DocType == "ITR").Select(l => l.PrEntry).Distinct().Count();
                Caption("lFoot", lines.Count == 0 ? "No request has been raised from this Production Order."
                    : prCount + " Purchase Request(s), " + itrCount + " Transfer Request(s), " + lines.Count(l => l.Status == "O") + " open row(s)");
            });
        }

        private DisplayRow CurrentRow()
        {
            if (gList.Current < 0 || gList.Current >= display.Count) throw new PlanException(400, "Click a document row first.");
            return display[gList.Current];
        }

        protected override void OnEvent(SAPbouiCOM.ItemEvent e, ref bool bubble)
        {
            if (Clicked(e, gList)) gList.Click(e.Row);
            else if (DoubleClicked(e, gList)) { gList.Click(e.Row); OpenDocument(CurrentRow()); }
            else if (Pressed(e, "bOpen")) OpenDocument(CurrentRow());
            else if (Pressed(e, "bRefresh")) Reload();
            else if (Pressed(e, "bBack")) MaterialPlanForm.Show();
            else if (Key(e, 116)) Reload();                               // F5
            else if (Key(e, 13) && gList.Current >= 0) OpenDocument(CurrentRow());
            else if (Key(e, 27)) Close();                                 // Esc
        }
    }
}

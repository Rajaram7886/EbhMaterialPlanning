using System;
using System.Collections.Generic;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>Purchase Requests of one Production Order: every request line raised from it, with the request status.</summary>
    internal sealed class LinkedRequestForm : PlanForm
    {
        private GridTable gList;
        private List<LinkedRequest> lines = new List<LinkedRequest>();
        private int pdoEntry, pdoNum;

        /// <summary>One grid row: either a Purchase Request group header (Line == null) or one of its detail lines.</summary>
        private sealed class DisplayRow
        {
            public int PrEntry;
            public LinkedRequest Line;      // null for a group header row
        }
        private List<DisplayRow> display = new List<DisplayRow>();

        private static readonly Dictionary<string, string> StatusNames = new Dictionary<string, string> { { "O", "Open" }, { "C", "Closed" }, { "X", "Cancelled" } };

        public static void Show(int pdoEntry)
        {
            Run(() =>
            {
                LinkedRequestForm open = OpenForms<LinkedRequestForm>().FirstOrDefault(x => x.pdoEntry == pdoEntry);
                if (open != null) { open.Select(); open.Reload(); return; }
                var f = new LinkedRequestForm { pdoEntry = pdoEntry, pdoNum = Store.PdoNum(pdoEntry) };
                f.Create("EBHMRPLNK", "Purchase Requests of Production Order " + f.pdoNum, 840, 360, false, f.Build);
            });
        }

        private void Build()
        {
            gList = Grid("gList", 8, 8, 824, 300)
                .Column("Series", "Series", width: 80)
                .Column("PR", "Purchase Request", SAPbouiCOM.BoFieldsType.ft_Integer, 80)
                .Column("Line", "Row", SAPbouiCOM.BoFieldsType.ft_Integer, 40)
                .Column("Date", "Date", SAPbouiCOM.BoFieldsType.ft_Date, 70)
                .Column("Stat", "Status", width: 70)
                .Column("Item", "Item No.", width: 100)
                .Column("Desc", "Description", width: 200)
                .Column("PdoLine", "Component Line", SAPbouiCOM.BoFieldsType.ft_Integer, 90)
                .Column("Qty", "Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 80)
                .Column("Open", "Open Qty", SAPbouiCOM.BoFieldsType.ft_Quantity, 80)
                .Bind();
            Button("bOpen", "Open Purchase Request", 8, 320, 140);
            Button("bRefresh", "Refresh", 152, 320, 70);
            Button("2", "Close", 226, 320, 70);
            Label("lFoot", "", 310, 323, 400);
            Reload();
        }

        protected override void Layout()
        {
            int w = Math.Max(560, Form.ClientWidth), h = Math.Max(160, Form.ClientHeight);
            Place("gList", 8, 8, w - 16, h - 8 - 40);
            Place("bOpen", 8, h - 30, 0, 0);
            Place("bRefresh", 152, h - 30, 0, 0);
            Place("2", 226, h - 30, 0, 0);
            Place("lFoot", 310, h - 27, 0, 0);
        }

        public void Reload()
        {
            // A request cancelled or closed since the last visit frees its quantity: the Production Order is refreshed first
            try { Store.Sync(new[] { pdoEntry }); }
            catch (Exception ex) { Warn("Production Order not refreshed : " + ex.Message); }
            lines = Store.LinkedRequests(pdoEntry);

            // One banner row per Purchase Request ("Series / No.") above its detail lines; lines already arrive
            // grouped (sorted by DocNum then LineNum), so a change of PrEntry marks the start of the next group.
            display = new List<DisplayRow>();
            int lastPr = -1;
            foreach (LinkedRequest l in lines)
            {
                if (l.PrEntry != lastPr)
                {
                    display.Add(new DisplayRow { PrEntry = l.PrEntry });
                    lastPr = l.PrEntry;
                }
                display.Add(new DisplayRow { PrEntry = l.PrEntry, Line = l });
            }

            Freeze(() =>
            {
                gList.Fill(display.Select(d =>
                {
                    // The group header carries no LinkedRequest of its own; borrow Series / Number from its group's first line.
                    LinkedRequest first = d.Line ?? lines.First(l => l.PrEntry == d.PrEntry);
                    return d.Line == null
                        ? new object[] { first.SeriesName, first.PrNum, 0, ToDate(first.PostingDate), "", "", "Purchase Request " + first.SeriesName + " " + first.PrNum, 0, 0m, 0m }
                        : new object[]
                        {
                            d.Line.SeriesName, d.Line.PrNum, d.Line.PrLine, ToDate(d.Line.PostingDate),
                            StatusNames.ContainsKey(d.Line.Status) ? StatusNames[d.Line.Status] : d.Line.Status,
                            d.Line.ItemCode, d.Line.Description, d.Line.PdoLine, d.Line.Quantity, d.Line.OpenQuantity
                        };
                }).ToList());
                for (int i = 0; i < display.Count; i++)
                    gList.Color(i, display[i].Line == null ? GridTable.Yellow : display[i].Line.Status == "O" ? GridTable.Green : GridTable.Grey);
                Caption("lFoot", lines.Count == 0 ? "No Purchase Request has been raised from this Production Order."
                    : lines.Select(l => l.PrNum).Distinct().Count() + " Purchase Request(s), " + lines.Count(l => l.Status == "O") + " open row(s)");
            });
        }

        private DisplayRow CurrentRow()
        {
            if (gList.Current < 0 || gList.Current >= display.Count) throw new PlanException(400, "Click a Purchase Request row first.");
            return display[gList.Current];
        }

        protected override void OnEvent(SAPbouiCOM.ItemEvent e, ref bool bubble)
        {
            if (Clicked(e, gList)) gList.Click(e.Row);
            else if (DoubleClicked(e, gList)) { gList.Click(e.Row); Store.OpenPurchaseRequest(CurrentRow().PrEntry); }
            else if (Pressed(e, "bOpen")) Store.OpenPurchaseRequest(CurrentRow().PrEntry);
            else if (Pressed(e, "bRefresh")) Reload();
        }
    }
}

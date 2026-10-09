using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace EBHMaterialPlanning
{
    /// <summary>Small DI API Recordset helpers. Callers build SQL only from validated values / constants.</summary>
    public static class Db
    {
        /// <summary>Serialises all DI API work (form events on the SAP thread, invoice link on a worker thread).</summary>
        public static readonly object DiLock = new object();

        public static SAPbobsCOM.Company Company
        {
            get
            {
                SAPbobsCOM.Company oCompany = Program.oCompany;
                if (oCompany == null || !oCompany.Connected) throw new PlanException(503, "SAP company is not connected.");
                return oCompany;
            }
        }

        public static string Text(SAPbobsCOM.Recordset oRs, string field)
        {
            object v = oRs.Fields.Item(field).Value;
            return v == null ? "" : Convert.ToString(v).Trim();
        }

        /// <summary>Runs a query and returns every row as field → object.</summary>
        public static List<Dictionary<string, object>> Rows(string sql)
        {
            var list = new List<Dictionary<string, object>>();
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oRs.DoQuery(sql);
                while (!oRs.EoF)
                {
                    var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < oRs.Fields.Count; i++) row[oRs.Fields.Item(i).Name] = oRs.Fields.Item(i).Value;
                    list.Add(row);
                    oRs.MoveNext();
                }
            }
            catch (COMException ex)
            {
                // Before Database Setup the add-on tables / fields do not exist: say what to do instead of showing the SQL error
                if (ex.Message.Contains("'U_EBH_"))
                    throw new PlanException(500, "The material planning fields are not set up in this company. Run Database Setup (Create All Environment), then restart SAP Business One.");
                throw new PlanException(500, "Database query failed: " + ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oRs);
            }
            return list;
        }

        public static Dictionary<string, object> FirstRow(string sql)
        {
            var rows = Rows(sql);
            return rows.Count > 0 ? rows[0] : null;
        }

        public static string Str(Dictionary<string, object> row, string key)
        {
            object v;
            return row != null && row.TryGetValue(key, out v) && v != null ? Convert.ToString(v).Trim() : "";
        }

        /// <summary>Escapes a value for a SQL string literal (single quotes doubled).</summary>
        public static string Quote(string value)
        {
            return "'" + (value ?? "").Replace("'", "''") + "'";
        }
    }
}

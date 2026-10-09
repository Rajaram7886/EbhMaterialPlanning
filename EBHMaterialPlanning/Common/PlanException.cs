using System;
using System.Collections.Generic;
using System.Linq;

namespace EBHMaterialPlanning
{
    /// <summary>
    /// Error meant for the user (validation, missing record, not authorised). Status follows the HTTP convention
    /// (400 input, 403 not authorised, 404 not found, 409 conflict, 500 SAP / database).
    /// </summary>
    public class PlanException : Exception
    {
        public int Status { get; private set; }
        /// <summary>Optional validation errors per field: { field, message }.</summary>
        public object Details { get; private set; }
        public PlanException(int status, string message) : base(message) { Status = status; }
        public PlanException(int status, string message, object details) : base(message) { Status = status; Details = details; }

        /// <summary>Field of the first validation error, "" when the error is not about one field.</summary>
        public string Field
        {
            get
            {
                var list = Details as IEnumerable<Dictionary<string, string>>;
                Dictionary<string, string> first = list == null ? null : list.FirstOrDefault();
                string f;
                return first != null && first.TryGetValue("field", out f) ? f : "";
            }
        }
    }
}

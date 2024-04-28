using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Graph.Mail.Client.Models
{
    public class EmailOrderby
    {
        public EmailOrderbyField? OrderbyField { get; set; }
        public EmailOrderbyType? OrderbyType { get; set; }
    }
    public enum EmailOrderbyField
    {
        FromEmail = 0,
        ToEmail = 1,
        Subject = 2,
        receivedDateTime = 3
    }

    public enum EmailOrderbyType
    {
        Asc = 0,
        Desc = 1,
    }
}

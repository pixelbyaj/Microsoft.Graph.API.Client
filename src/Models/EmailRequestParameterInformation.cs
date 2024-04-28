using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Graph.API.Client.Models
{
    public class EmailRequestParameterInformation
    {
        public bool? IsRead { get; set; } = false;
        public string? Filter { get; set; } = null;
        public string? Search { get; set; } = null;
        public IList<EmailOrderby>? EmailOrderby { get; set; } = new List<EmailOrderby>
        {
            new EmailOrderby
            {
                OrderbyField = EmailOrderbyField.receivedDateTime,
                OrderbyType = EmailOrderbyType.Desc
            }
        };

        public bool? IncludeAttachments { get; set; } = false;

    }

        
}

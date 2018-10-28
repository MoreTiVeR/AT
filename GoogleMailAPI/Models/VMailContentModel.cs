using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GoogleMailAPI.Models
{
    public class VMailContentModel
    {
        public string MailTo { get; set; }
        public string MailSubject { get; set; }
        public string MailBody{ get; set; }
    }
}
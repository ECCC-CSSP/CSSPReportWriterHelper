//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CSSPReportWriterHelper.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class Email
    {
        public int EmailID { get; set; }
        public int EmailTVItemID { get; set; }
        public string EmailAddress { get; set; }
        public int EmailType { get; set; }
        public System.DateTime LastUpdateDate_UTC { get; set; }
        public int LastUpdateContactTVItemID { get; set; }
    
        public virtual TVItem TVItem { get; set; }
    }
}

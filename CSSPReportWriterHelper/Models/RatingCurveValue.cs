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
    
    public partial class RatingCurveValue
    {
        public int RatingCurveValueID { get; set; }
        public int RatingCurveID { get; set; }
        public double StageValue_m { get; set; }
        public double DischargeValue_m3_s { get; set; }
        public System.DateTime LastUpdateDate_UTC { get; set; }
        public int LastUpdateContactTVItemID { get; set; }
    
        public virtual RatingCurve RatingCurve { get; set; }
    }
}

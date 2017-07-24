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
    
    public partial class MWQMSample
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public MWQMSample()
        {
            this.MWQMSampleLanguages = new HashSet<MWQMSampleLanguage>();
        }
    
        public int MWQMSampleID { get; set; }
        public int MWQMSiteTVItemID { get; set; }
        public int MWQMRunTVItemID { get; set; }
        public System.DateTime SampleDateTime_Local { get; set; }
        public Nullable<double> Depth_m { get; set; }
        public int FecCol_MPN_100ml { get; set; }
        public Nullable<double> Salinity_PPT { get; set; }
        public Nullable<double> WaterTemp_C { get; set; }
        public Nullable<double> PH { get; set; }
        public string SampleTypesText { get; set; }
        public int SampleType_old { get; set; }
        public Nullable<int> Tube_10 { get; set; }
        public Nullable<int> Tube_1_0 { get; set; }
        public Nullable<int> Tube_0_1 { get; set; }
        public string ProcessedBy { get; set; }
        public System.DateTime LastUpdateDate_UTC { get; set; }
        public int LastUpdateContactTVItemID { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MWQMSampleLanguage> MWQMSampleLanguages { get; set; }
        public virtual TVItem TVItem { get; set; }
        public virtual TVItem TVItem1 { get; set; }
    }
}

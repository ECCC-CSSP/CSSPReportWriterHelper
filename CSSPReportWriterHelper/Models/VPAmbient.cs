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
    
    public partial class VPAmbient
    {
        public int VPAmbientID { get; set; }
        public int VPScenarioID { get; set; }
        public int Row { get; set; }
        public double MeasurementDepth_m { get; set; }
        public double CurrentSpeed_m_s { get; set; }
        public double CurrentDirection_deg { get; set; }
        public double AmbientSalinity_PSU { get; set; }
        public double AmbientTemperature_C { get; set; }
        public int BackgroundConcentration_MPN_100ml { get; set; }
        public double PollutantDecayRate_per_day { get; set; }
        public double FarFieldCurrentSpeed_m_s { get; set; }
        public double FarFieldCurrentDirection_deg { get; set; }
        public double FarFieldDiffusionCoefficient { get; set; }
        public System.DateTime LastUpdateDate_UTC { get; set; }
        public int LastUpdateContactTVItemID { get; set; }
    
        public virtual VPScenario VPScenario { get; set; }
    }
}

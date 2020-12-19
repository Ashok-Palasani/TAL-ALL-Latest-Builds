//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace i_facility.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblgenericworkcode
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public tblgenericworkcode()
        {
            this.tblgenericworkentries = new HashSet<tblgenericworkentry>();
        }
    
        public int GenericWorkID { get; set; }
        public string GenericWorkCode { get; set; }
        public string GenericWorkDesc { get; set; }
        public string MessageType { get; set; }
        public int GWCodesLevel { get; set; }
        public Nullable<int> GWCodesLevel1ID { get; set; }
        public Nullable<int> GWCodesLevel2ID { get; set; }
        public int IsDeleted { get; set; }
        public System.DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public Nullable<System.DateTime> ModifiedOn { get; set; }
        public Nullable<int> ModifiedBy { get; set; }
        public Nullable<int> EndCode { get; set; }
        public Nullable<System.DateTime> DeletedDate { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<tblgenericworkentry> tblgenericworkentries { get; set; }
    }
}

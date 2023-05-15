using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.PickList
{
    public class AllocationItem
    {
        public int PickListDocEntry { get; set; }
        public int PickListLineNum { get; set; }
        public int SODocEntry { get; set; }
        public int SOLineNum { get; set; }
        public string ItemCode { get; set; }
        public string ItemDesc { get; set; }
        public string DistNumber { get; set; }
        public string WhsCode { get; set; }
        public decimal ActualPickQty { get; set; }
        public decimal DraftQty { get; set; }
        public decimal TotalQty { get; set; }
    }
}

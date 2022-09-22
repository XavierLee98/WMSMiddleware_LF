using SAPbobsCOM;
using System;

namespace IMAppSapMidware_NetCore.Helper
{
    public class ErpPropertyHelper
    {
        public int Id { get; set; }
        public string Server { get; set; }
        public string DbServer { get; set; }
        public string SAPCompany { get; set; }       
        public string SAPUser { get; set; }
        public string SAPPass { get; set; }
        public string PortNumber { get; set; }
        public string LicenseServer { get; set; }
        public int DBType { get; private set; } // <--- add in to determine the year of the SQL server

        //string _databaseServerType { get; set; }

        //public string DatabaseServerType
        //{
        //    get => _databaseServerType;            
        //    set
        //    {
        //        if (_databaseServerType != value)
        //        {
        //            _databaseServerType = value;
        //            BoDataServerTypes = (BoDataServerTypes)Enum.Parse(typeof(BoDataServerTypes), _databaseServerType);
        //        }
        //    }
        //}

    }
}

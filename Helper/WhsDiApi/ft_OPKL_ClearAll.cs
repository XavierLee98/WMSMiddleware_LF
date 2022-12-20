using IMAppSapMidware_NetCore.Helper.SQL;
using IMAppSapMidware_NetCore.Models.SAPModels;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace IMAppSapMidware_NetCore.Helper.WhsDiApi
{
    public class ft_OPKL_ClearAll : IDisposable
    {
        public void Dispose() => GC.Collect();
        public static string LastSAPMsg { get; set; } = string.Empty;

        static string currentKey = string.Empty;
        static string currentStatus = string.Empty;
        static string CurrentDocNum = string.Empty;

        public static string Erp_DBConnStr { get; set; } = string.Empty;

        static DataTable dt = null;
        static SAPParam par;
        static SAPCompany sap;
        static PickLists oPickLists = null;
        static SAPbobsCOM.Documents oDocument = null;

        static void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public static void ClearAll()
        {
            string request = "Clear Pick List";

            try
            {
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                dt = ft_General.LoadData("LoadClearOPKL_sp");

                if (dt.Rows.Count > 0)
                {
                    par = SAP.GetSAPUser();
                    sap = SAP.getSAPCompany(par);

                    if (!sap.connectSAP())
                    {
                        Log($"{sap.errMsg}");
                        throw new Exception(sap.errMsg);
                    }

                    string key = dt.Rows[0]["key"].ToString();
                    currentKey = key;
                    currentStatus = failed_status;

                    oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
                    CurrentDocNum = dt.Rows[0]["sapDocNumber"].ToString();

                    if (!oPickLists.GetByKey(int.Parse(CurrentDocNum)))
                    {
                        LastSAPMsg = sap.oCom.GetLastErrorDescription().ToString();
                        throw new Exception(LastSAPMsg);
                    }

                    if (oPickLists.Status == BoPickStatus.ps_Closed)
                    {
                        LastSAPMsg = $"PickList No {CurrentDocNum} is closed.";
                        throw new Exception(LastSAPMsg);
                    }

                    if (!sap.oCom.InTransaction)
                        sap.oCom.StartTransaction();

                    List<int> SOdocEntries = new List<int>();
                    //List<int> RIdocEntries = new List<int>();

                    for (int x = 0; x < oPickLists.Lines.Count; x++)
                    {
                        oPickLists.Lines.SetCurrentLine(x);
                        if (oPickLists.Lines.PickStatus == BoPickStatus.ps_Closed) continue;
                        if (oPickLists.Lines.BaseObjectType == "17")
                        {
                            SOdocEntries.Add(oPickLists.Lines.OrderEntry);
                            continue;
                        }
                        //if (oPickLists.Lines.BaseObjectType == "13")
                        //{
                        //    RIdocEntries.Add(oPickLists.Lines.OrderEntry);
                        //    continue;
                        //}
                    }

                    if (SOdocEntries != null && SOdocEntries.Count > 0)
                    {
                        var distinctSONo = SOdocEntries.Distinct().ToList();

                        for (int y = 0; y < distinctSONo.Count; y++)
                        {
                            oDocument = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                            if (!oDocument.GetByKey(distinctSONo[y]))
                            {
                                throw new Exception(sap.oCom.GetLastErrorDescription());
                            }

                            for (int z = 0; z < oDocument.Lines.Count; z++)
                            {
                                oDocument.Lines.SetCurrentLine(z);
                                if (oPickLists.Lines.PickStatus == BoPickStatus.ps_Closed) continue;

                                for (int x = 0; x < oPickLists.Lines.Count; x++)
                                {
                                    oPickLists.Lines.SetCurrentLine(x);
                                    if (oPickLists.Lines.PickStatus == BoPickStatus.ps_Closed) continue;

                                    if (oPickLists.Lines.OrderEntry == oDocument.Lines.DocEntry && oPickLists.Lines.OrderRowID == oDocument.Lines.LineNum && oPickLists.Lines.BaseObjectType == "17")
                                    {
                                        oPickLists.Lines.PickedQuantity = 0;

                                        if (!string.IsNullOrEmpty(oDocument.Lines.BatchNumbers.BatchNumber))
                                        {
                                            for (int i = 0; i < oPickLists.Lines.BatchNumbers.Count; i++)
                                            {
                                                oDocument.Lines.BatchNumbers.SetCurrentLine(i);
                                                oDocument.Lines.BatchNumbers.Quantity = 0;
                                            }
                                        }
                                    }
                                }
                            }

                            int result = oDocument.Update();

                            if (result != 0)
                            {
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                throw new Exception(sap.oCom.GetLastErrorDescription());
                            }

                        }
                    }
                    int retcode = oPickLists.Update();

                    if (retcode != 0)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        throw new Exception(sap.oCom.GetLastErrorDescription());
                    }
                    else
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        Log($"{key }\n {success_status }\n  { CurrentDocNum } \n");
                        ft_General.UpdateStatus(key, success_status, "", CurrentDocNum);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("Clear OPKL", ex.Message);
                Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
                ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
            }
        }
    }
}

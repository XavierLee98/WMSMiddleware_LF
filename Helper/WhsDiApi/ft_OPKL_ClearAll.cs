using Dapper;
using IMAppSapMidware_NetCore.Helper.SQL;
using IMAppSapMidware_NetCore.Models.PickList;
using IMAppSapMidware_NetCore.Models.SAPModels;
using Microsoft.Data.SqlClient;
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

        //Select batch From PickListAllocateTable (-)   

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
                    string key = dt.Rows[0]["key"].ToString();
                    currentKey = key;
                    currentStatus = failed_status;
                    CurrentDocNum = dt.Rows[0]["AbsEntry"].ToString();

                    par = SAP.GetSAPUser();
                    sap = SAP.getSAPCompany(par);

                    if (!sap.connectSAP())
                    {
                        Log($"{sap.errMsg}");
                        throw new Exception(sap.errMsg);
                    }

                    oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);

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

                    List<SOLine> SOdocEntries = new List<SOLine>();

                    for (int x = 0; x < oPickLists.Lines.Count; x++)
                    {
                        oPickLists.Lines.SetCurrentLine(x);
                        if (oPickLists.Lines.PickStatus == BoPickStatus.ps_Closed) continue;
                        oPickLists.Lines.PickedQuantity = 0;

                        if (oPickLists.Lines.BaseObjectType == "17")
                        {
                            SOdocEntries.Add(new SOLine { SODocEntry = oPickLists.Lines.OrderEntry, SOLineNum = oPickLists.Lines.OrderRowID });
                        }

                        if (oPickLists.Lines.BatchNumbers.Count - 1 > 0 || oPickLists.Lines.BatchNumbers.BatchNumber != "")
                        {
                            for (int y = 0; y < oPickLists.Lines.BatchNumbers.Count; y++)
                            {
                                oPickLists.Lines.BatchNumbers.SetCurrentLine(y);
                                oPickLists.Lines.BatchNumbers.Quantity = 0;
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

                    DataTable batchesTable = GetPickLineBatches(int.Parse(CurrentDocNum));

                    List<int> distinctSONo = null;

                    if (SOdocEntries != null && SOdocEntries.Count > 0 && batchesTable.Rows.Count >0)
                    {
                        distinctSONo = SOdocEntries.GroupBy(x => x.SODocEntry).Select(y => y.Key).ToList();

                        for (int i = 0; i < distinctSONo.Count; i++)
                        {
                            oDocument = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                            if (!oDocument.GetByKey(distinctSONo[i]))
                            {
                                throw new Exception(sap.oCom.GetLastErrorDescription());
                            }

                            for (int j = 0; j < oDocument.Lines.Count; j++)
                            {
                                oDocument.Lines.SetCurrentLine(j);

                                if (oDocument.Lines.LineStatus == BoStatus.bost_Close) continue;

                                if (!string.IsNullOrEmpty(oDocument.Lines.BatchNumbers.BatchNumber))
                                {
                                    foreach (DataRow row in batchesTable.Rows)
                                    {
                                        double qty = 0;
                                        bool isBatchFound = false;

                                        if (Convert.ToInt32(row["SODocEntry"]) == distinctSONo[i] && Convert.ToInt32(row["SOLineNum"]) == oDocument.Lines.LineNum)
                                        {
                                            for (int m = 0; m < oDocument.Lines.BatchNumbers.Count; m++)
                                            {
                                                oDocument.Lines.BatchNumbers.SetCurrentLine(m);

                                                if (row["DistNumber"].ToString() == oDocument.Lines.BatchNumbers.BatchNumber)
                                                {
                                                    isBatchFound = true;
                                                    qty = double.Parse(row["TotalQty"].ToString());
                                                    break;
                                                }
                                            }

                                                if (isBatchFound)
                                                {
                                                    if (oDocument.Lines.BatchNumbers.Quantity <= qty)
                                                    {
                                                        oDocument.Lines.BatchNumbers.Quantity = 0;
                                                    }
                                                    else
                                                    {
                                                        oDocument.Lines.BatchNumbers.Quantity -= qty;
                                                    }
                                                }
                                        }
                                    }
                                }
                            }

                            retcode = oDocument.Update();

                            if (retcode != 0)
                            {
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                throw new Exception(sap.oCom.GetLastErrorDescription());
                            }
                        }
                    }

                    var AllocatedPickList = new List<AllocationItem>();
                    foreach (var so in SOdocEntries)
                    {
                        AllocatedPickList.AddRange(LoadSOBatchTransaction(int.Parse(CurrentDocNum), so.SODocEntry, so.SOLineNum));
                    }

                    if (retcode != 0)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        throw new Exception(sap.oCom.GetLastErrorDescription());
                    }
                    else
                    {
                        if (sap.oCom.InTransaction)
                        {
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }

                        if (UpdateStatus(int.Parse(CurrentDocNum), out string errMsg) < 0)
                        { 
                            throw new Exception(errMsg);
                        }
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

        static int UpdateStatus(int absentry, out string errorMsg)
        {
            try
            {
                errorMsg = "";
                var conn = new SqlConnection(Program._DbMidwareConnStr);
                var result = conn.Execute("sp_PickList_ClearAllocateItem",
                             new { absentry = absentry },
                             commandType: CommandType.StoredProcedure);

                return result;
            }
            catch (Exception excep)
            {
                errorMsg = excep.ToString();
                return -1;
            }
        }

        static void InsertCheckRequest(int absentry, List<int> otherSOList)
        {
            try
            {
                foreach(var soNo in otherSOList)
                {
                    LoadSOBatchTransaction(absentry, soNo);
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
            }
        }

        static List<AllocationItem> LoadSOBatchTransaction(int absentry, int sodocentry, int solinenum = -1)
        {
            try
            {
                SqlConnection conn = new SqlConnection(Program._DbMidwareConnStr);

                var list = conn.Query<AllocationItem>($"sp_PickList_GetSOLineBatches",
                                new { absentry = absentry, docentry = sodocentry, solinenum = solinenum },
                                commandType: CommandType.StoredProcedure,
                                commandTimeout: 0).ToList();
                return list;
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                return null;
            }
        }

        static DataTable GetPickLineBatches(int absentry)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(Program._DbMidwareConnStr))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("sp_PickList_GetPickLineBatches", connection))
                {

                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@absentry", absentry);
                    //command.Parameters.AddWithValue("@linenum", -1);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable table = new DataTable();

                        adapter.Fill(table);

                        return table;
                    }
                }

            }
        }
    }
}



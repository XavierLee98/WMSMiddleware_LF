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
                        if (oPickLists.Lines.BaseObjectType == "17")
                        {
                            SOdocEntries.Add(new SOLine { SODocEntry = oPickLists.Lines.OrderEntry, SOLineNum = oPickLists.Lines.OrderRowID });
                            continue;
                        }
                    }

                    if (SOdocEntries != null && SOdocEntries.Count > 0)
                    {
                        var distinctSONo = SOdocEntries.Distinct().ToList();

                        for (int y = 0; y < distinctSONo.Count; y++)
                        {
                            oDocument = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                            if (!oDocument.GetByKey(distinctSONo[y].SODocEntry))
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

                                        DataTable batchTable = GetPickLineBatches(int.Parse(CurrentDocNum), x);
                                        oPickLists.Lines.PickedQuantity = 0;

                                        if (!string.IsNullOrEmpty(oDocument.Lines.BatchNumbers.BatchNumber))
                                        {
                                            bool isBatchFound = false;
                                            double qty = 0;

                                            for (int i = 0; i < oDocument.Lines.BatchNumbers.Count; i++)
                                            {
                                                oDocument.Lines.BatchNumbers.SetCurrentLine(i);
                                                foreach (DataRow row in batchTable.Rows)
                                                {
                                                    var count = batchTable.Rows;
                                                    var batch1 = row["Batch"].ToString();
                                                    var batch2 = oDocument.Lines.BatchNumbers.BatchNumber;
                                                    if (row["Batch"].ToString() == oDocument.Lines.BatchNumbers.BatchNumber)
                                                    {
                                                        isBatchFound = true;
                                                        qty = double.Parse(row["Quantity"].ToString());
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

                    #region Assigned Back Other Affected PickList
                    var AllocatedPickList = new List<AllocationItem>();

                    foreach(var so in SOdocEntries)
                    {
                        AllocatedPickList.AddRange(LoadSOBatchTransaction(int.Parse(CurrentDocNum), so.SODocEntry, so.SOLineNum));
                    }

                    var picklistnoList = AllocatedPickList.GroupBy(x => x.PickListDocEntry).Select(y=>y.Key).ToList();
                    if (picklistnoList != null && picklistnoList.Count > 0) 
                    {
                        foreach (var pickno in picklistnoList)
                        {
                            if (!oPickLists.GetByKey(pickno))
                            {
                                throw new Exception(sap.oCom.GetLastErrorDescription());
                            }

                            var currentPKBatches = AllocatedPickList.Where(x => x.PickListDocEntry == pickno && x.ActualPickQty > 0).ToList();
                            if (currentPKBatches == null || currentPKBatches.Count <= 0) continue;
                            for (int a = 0; a < oPickLists.Lines.Count; a++)
                            {
                                oPickLists.Lines.SetCurrentLine(a);
                                bool isPickLinefound = false;

                                for (int b = 0; b < currentPKBatches.Count; b++)
                                {
                                    if (oPickLists.Lines.LineNumber == currentPKBatches[b].PickListLineNum)
                                    {
                                        isPickLinefound = true;
                                        break;
                                    }

                                    if (isPickLinefound)
                                    {
                                        if (oPickLists.Lines.PickStatus == BoPickStatus.ps_Closed) continue;

                                        if (oPickLists.Lines.BatchNumbers.Count - 1 > 0 && oPickLists.Lines.BatchNumbers.BatchNumber != "")
                                        {
                                            var isPickBatchFound = false;

                                            for (int c = 0; c < oPickLists.Lines.BatchNumbers.Count; c++)
                                            {
                                                if (oPickLists.Lines.BatchNumbers.BatchNumber == currentPKBatches[b].DistNumber)
                                                {
                                                    isPickBatchFound = true;
                                                    break;
                                                }
                                            }

                                            if (isPickBatchFound)
                                            {
                                                oPickLists.Lines.BatchNumbers.Quantity = int.Parse(currentPKBatches[b].ActualPickQty.ToString());
                                            }
                                            else
                                            {
                                                oPickLists.Lines.BatchNumbers.Add();
                                                oPickLists.Lines.BatchNumbers.BatchNumber = currentPKBatches[b].DistNumber;
                                                oPickLists.Lines.BatchNumbers.Quantity = int.Parse(currentPKBatches[b].ActualPickQty.ToString());
                                                oPickLists.Lines.BatchNumbers.BaseLineNumber = oPickLists.Lines.LineNumber;
                                            }
                                        }
                                        else
                                        {
                                            oPickLists.Lines.BatchNumbers.BatchNumber = currentPKBatches[b].DistNumber;
                                            oPickLists.Lines.BatchNumbers.Quantity = int.Parse(currentPKBatches[b].ActualPickQty.ToString());
                                            oPickLists.Lines.BatchNumbers.BaseLineNumber = oPickLists.Lines.LineNumber;
                                        }
                                    }
                                }
                            }
                            retcode = oPickLists.Update();

                            if (retcode != 0)
                            {
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                throw new Exception(sap.oCom.GetLastErrorDescription());
                            }
                        }
                    }
                    #endregion

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

                        //if (UpdateStatus(int.Parse(CurrentDocNum), out string errMsg) < 0)
                        //{
                        //    throw new Exception(errMsg);
                        //}
                        ClearAllocateItemTable(int.Parse(CurrentDocNum));
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

        static List<AllocationItem> LoadSOBatchTransaction(int absentry, int sodocentry, int solinenum)
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

        //static int UpdateStatus(int absentry, out string errorMsg)
        //{
        //    try
        //    {
        //        errorMsg = "";
        //        var conn = new SqlConnection(Program._DbErpConnStr);
        //        var result = conn.Execute("zwa_IMApp_PickList_spResetPickListStatus",
        //                     new { AbsEntry = absentry },
        //                     commandType: CommandType.StoredProcedure);

        //        return result;
        //    }
        //    catch (Exception excep)
        //    {
        //        errorMsg = excep.ToString();
        //        return -1;
        //    }
        //}

        static void ClearAllocateItemTable(int absentry)
        {
            var conn = new SqlConnection(Program._DbMidwareConnStr);
            var result = conn.Execute("sp_PickList_ClearAllocateItem",
             new { AbsEntry = absentry },
             commandType: CommandType.StoredProcedure);
        }

        static DataTable GetPickLineBatches(int absentry, int linenum)
        {
            DataTable dataTable = new DataTable();

            string query = $"SELECT [SODocEntry], [SOLineNum], [PickListDocEntry], [PickListLineNum], [ItemCode], [ItemDesc], [Batch], SUM([Quantity]) [Quantity] FROM zmwPickListAllocateItem " +
                           $"WHERE PickListDocEntry = {absentry} AND PickListLineNum = {linenum} AND IsCancelled = 0 " +
                           $"GROUP BY [SODocEntry], [SOLineNum], [PickListDocEntry], [PickListLineNum], [ItemCode], [ItemDesc], [Batch] " +
                           $"HAVING SUM(Quantity) > 0;";
            using (SqlConnection connection = new SqlConnection(Program._DbMidwareConnStr))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
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
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
        static DataTable dtLines = null;
        static DataTable dtBatches = null;
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

                    DataTable dtLines = new DataTable("PickLines");
                    DataTable dtBatches = new DataTable("PickBatches");

                    dtLines.Columns.Add("PickNo", typeof(int));
                    dtLines.Columns.Add("PickLine", typeof(int));
                    dtLines.Columns.Add("OrderNo", typeof(int));
                    dtLines.Columns.Add("OrderLine", typeof(int));
                    dtLines.Columns.Add("PickQty", typeof(double));

                    dtBatches.Columns.Add("PickNo", typeof(int));
                    dtBatches.Columns.Add("PickLine", typeof(int));
                    dtBatches.Columns.Add("OrderNo", typeof(int));
                    dtBatches.Columns.Add("OrderLine", typeof(int));
                    dtBatches.Columns.Add("Batch", typeof(string));
                    dtBatches.Columns.Add("Quantity", typeof(double));

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

                    int batchRowCount = 0;
                    int lineRowCount = 0;

                    for (int x = 0; x < oPickLists.Lines.Count; x++)
                    {
                        oPickLists.Lines.SetCurrentLine(x);
                        if (oPickLists.Lines.PickStatus == BoPickStatus.ps_Closed) continue;

                        DataRow newLineRow = dtLines.NewRow();
                        DataRow newbatchRow = dtBatches.NewRow();

                        if (lineRowCount > 0)
                        {
                            newLineRow = dtLines.NewRow();
                            newbatchRow = dtBatches.NewRow();
                            batchRowCount = 0;
                        }

                        if (oPickLists.Lines.BaseObjectType == "17")
                        {
                            newLineRow["PickNo"] = int.Parse(CurrentDocNum);
                            newLineRow["PickLine"] = oPickLists.Lines.LineNumber;
                            newLineRow["OrderNo"] = oPickLists.Lines.OrderEntry;
                            newLineRow["OrderLine"] = oPickLists.Lines.OrderRowID;
                            newLineRow["PickQty"] = oPickLists.Lines.PickedQuantity;

                            dtLines.Rows.Add(newLineRow);
                            lineRowCount++;
                        }

                        oPickLists.Lines.PickedQuantity = 0;

                        //if (oPickLists.Lines.BaseObjectType == "17")
                        //{
                        //    SOdocEntries.Add(new SOLine { SODocEntry = oPickLists.Lines.OrderEntry, SOLineNum = oPickLists.Lines.OrderRowID });
                        //}

                        if (oPickLists.Lines.BatchNumbers.Count - 1 > 0 || oPickLists.Lines.BatchNumbers.BatchNumber != "")
                        {
                            for (int y = 0; y < oPickLists.Lines.BatchNumbers.Count; y++)
                            {
                                oPickLists.Lines.BatchNumbers.SetCurrentLine(y);

                                if (batchRowCount > 0)
                                {
                                    newbatchRow = dtBatches.NewRow();
                                }

                                newbatchRow["PickNo"] = int.Parse(CurrentDocNum);
                                newbatchRow["PickLine"] = oPickLists.Lines.LineNumber;
                                newbatchRow["OrderNo"] = oPickLists.Lines.OrderEntry;
                                newbatchRow["OrderLine"] = oPickLists.Lines.OrderRowID;
                                newbatchRow["Batch"] = oPickLists.Lines.BatchNumbers.BatchNumber;
                                newbatchRow["Quantity"] = oPickLists.Lines.BatchNumbers.Quantity;

                                dtBatches.Rows.Add(newbatchRow);
                                batchRowCount++;

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

                    //Console.WriteLine("Rows:");
                    //foreach (DataRow row in dtBatches.Rows)
                    //{
                    //    Console.WriteLine("PickNo: " + row["PickNo"] + ", PickLine: " + row["PickLine"] + ", OrderNo: " + row["OrderNo"] +
                    //                      ", OrderLine: " + row["OrderLine"] + ", Batch: " + row["Batch"] + ", Quantity: " + row["Quantity"]);
                    //}

                    List<int> distinctSONo = null;

                    if (dtLines != null && dtLines.Rows.Count > 0 && dtBatches != null && dtBatches.Rows.Count > 0) 
                    {
                        distinctSONo = dtLines.AsEnumerable()
                                       .GroupBy(row => row.Field<int>("OrderNo"))
                                       .Select(x => x.Key).ToList();

                        if (distinctSONo != null && distinctSONo.Count > 0) 
                        {
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
                                        DataRow[] lineRows = dtBatches.Select("OrderNo= " + oDocument.DocEntry + "AND OrderLine= " + oDocument.Lines.LineNum);

                                        if(lineRows != null && lineRows.Length > 0)
                                        {

                                            foreach (DataRow row in lineRows)
                                            {
                                                double qty = 0;
                                                bool isBatchFound = false;

                                                for (int m = 0; m < oDocument.Lines.BatchNumbers.Count; m++)
                                                {
                                                    oDocument.Lines.BatchNumbers.SetCurrentLine(m);

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

                                retcode = oDocument.Update();

                                if (retcode != 0)
                                {
                                    if (sap.oCom.InTransaction)
                                        sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    throw new Exception(sap.oCom.GetLastErrorDescription());
                                }
                            }
                        }
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
            finally
            {
                dtBatches = null;
                dtLines = null;
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


 
    }
}

//DataTable batchesTable = GetPickLineBatches(int.Parse(CurrentDocNum));

//if (SOdocEntries != null && SOdocEntries.Count > 0 && batchesTable.Rows.Count >0)
//{
//    distinctSONo = SOdocEntries.GroupBy(x => x.SODocEntry).Select(y => y.Key).ToList();

//    for (int i = 0; i < distinctSONo.Count; i++)
//    {
//        oDocument = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

//        if (!oDocument.GetByKey(distinctSONo[i]))
//        {
//            throw new Exception(sap.oCom.GetLastErrorDescription());
//        }

//        for (int j = 0; j < oDocument.Lines.Count; j++)
//        {
//            oDocument.Lines.SetCurrentLine(j);

//            if (oDocument.Lines.LineStatus == BoStatus.bost_Close) continue;

//            if (!string.IsNullOrEmpty(oDocument.Lines.BatchNumbers.BatchNumber))
//            {
//                foreach (DataRow row in batchesTable.Rows)
//                {
//                    double qty = 0;
//                    bool isBatchFound = false;

//                    if (Convert.ToInt32(row["SODocEntry"]) == distinctSONo[i] && Convert.ToInt32(row["SOLineNum"]) == oDocument.Lines.LineNum)
//                    {
//                        for (int m = 0; m < oDocument.Lines.BatchNumbers.Count; m++)
//                        {
//                            oDocument.Lines.BatchNumbers.SetCurrentLine(m);

//                            if (row["DistNumber"].ToString() == oDocument.Lines.BatchNumbers.BatchNumber)
//                            {
//                                isBatchFound = true;
//                                qty = double.Parse(row["TotalQty"].ToString());
//                                break;
//                            }
//                        }

//                            if (isBatchFound)
//                            {
//                                if (oDocument.Lines.BatchNumbers.Quantity <= qty)
//                                {
//                                    oDocument.Lines.BatchNumbers.Quantity = 0;
//                                }
//                                else
//                                {
//                                    oDocument.Lines.BatchNumbers.Quantity -= qty;
//                                }
//                            }
//                    }
//                }
//            }
//        }

//        retcode = oDocument.Update();

//        if (retcode != 0)
//        {
//            if (sap.oCom.InTransaction)
//                sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
//            throw new Exception(sap.oCom.GetLastErrorDescription());
//        }
//    }
//}

//static void InsertCheckRequest(int absentry, List<int> otherSOList)
//{
//    try
//    {
//        foreach (var soNo in otherSOList)
//        {
//            LoadSOBatchTransaction(absentry, soNo);
//        }
//    }
//    catch (Exception ex)
//    {
//        Log($"{ ex.Message } \n");
//    }
//}

//static DataTable GetPickLineBatches(int absentry)
//{
//    DataTable dataTable = new DataTable();

//    using (SqlConnection connection = new SqlConnection(Program._DbMidwareConnStr))
//    {
//        connection.Open();

//        using (SqlCommand command = new SqlCommand("sp_PickList_GetPickLineBatches", connection))
//        {

//            command.CommandType = CommandType.StoredProcedure;

//            command.Parameters.AddWithValue("@absentry", absentry);
//            //command.Parameters.AddWithValue("@linenum", -1);
//            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
//            {
//                DataTable table = new DataTable();

//                adapter.Fill(table);

//                return table;
//            }
//        }

//    }
//}

using Dapper;
using IMAppSapMidware_NetCore.Helper.SQL;
using IMAppSapMidware_NetCore.Models.PickList;
using IMAppSapMidware_NetCore.Models.SAPModels;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace IMAppSapMidware_NetCore.Helper.WhsDiApi
{
    public class ft_OPKL: IDisposable
    {
        public void Dispose() => GC.Collect();
        public static string LastSAPMsg { get; set; } = string.Empty;

        static string currentKey = string.Empty;
        static string currentStatus = string.Empty;
        static string CurrentDocNum = string.Empty;
        public static string Erp_DBConnStr { get; set; } = string.Empty;

        static DataTable dt = null;
        static DataTable dtDetails = null;
        static SAPParam par;
        static SAPCompany sap;
        static PickLists oPickLists = null;
        static PickLists_Lines oPickLists_Lines = null;

        static SAPbobsCOM.Documents oDocument = null;

        static int retcode = -1;

        static void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public async static void Post()
        {
            try
            {
                string request = "Update Pick List";
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;
                int retcode = 0;
                double totalWeight = 0;

                LoadDataToDataTable(request);

                if(dt.Rows.Count > 0)
                {
                    string key = dt.Rows[0]["key"].ToString();
                    currentKey = key;
                    currentStatus = failed_status;
                    CurrentDocNum = dt.Rows[0]["sapDocNumber"].ToString();

                    par = SAP.GetSAPUser();
                    sap = SAP.getSAPCompany(par);
                        
                    if (!sap.connectSAP())
                    {
                        Log($"{sap.errMsg}");
                        throw new Exception(sap.errMsg);
                    }

                    if (!sap.oCom.InTransaction)
                        sap.oCom.StartTransaction();

                    //Get Distinct SO No List
                    List<int> docEntryList = dt.AsEnumerable()
                                                .Select(row => row.Field<int>("BaseEntry"))
                                                .Distinct()
                                                .ToList();


                    foreach (var soDocNo in docEntryList)
                    {
                        //Load Transaction
                        var transactionList = LoadPickTransaction(int.Parse(CurrentDocNum.ToString())).ToList();

                        //Get Trasanction SO Line List
                        var lineList = transactionList.Where(z => z.SODocEntry == soDocNo).GroupBy(x => x.SOLineNum).Select(y => y.Key).ToList();

                        oDocument = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        if (!oDocument.GetByKey(soDocNo))
                        {
                            throw new Exception(sap.oCom.GetLastErrorDescription());
                        }

                        foreach (var line in lineList)
                        {
                            bool isSOLineFound = false;

                            for (int i = 0; i < oDocument.Lines.Count; i++)
                            {
                                oDocument.Lines.SetCurrentLine(i);

                                if (oDocument.Lines.LineNum == line)
                                {
                                    isSOLineFound = true;
                                    break;
                                }
                            }

                            if (isSOLineFound)
                            {
                                //Get Distinct Transaction SO Line
                                var batchList = transactionList.Where(x => x.SODocEntry == soDocNo && x.SOLineNum == line).ToList();

                                bool isNext = false;

                                foreach(var batch in batchList)
                                {
                                    if (batch.DraftQty <= 0) continue;

                                    if((oDocument.Lines.BatchNumbers.Count - 1 == 0 && oDocument.Lines.BatchNumbers.BatchNumber == "") || isNext == false)
                                    {
                                        oDocument.Lines.BatchNumbers.BatchNumber = batch.DistNumber;
                                        oDocument.Lines.BatchNumbers.Quantity = double.Parse(batch.DraftQty.ToString());
                                    }
                                    else
                                    {
                                        bool isBatchFound = false;

                                        for (int j = 0; j < oDocument.Lines.BatchNumbers.Count; j++)
                                        {
                                            oDocument.Lines.BatchNumbers.SetCurrentLine(j);

                                            if (oDocument.Lines.LineStatus == BoStatus.bost_Close) continue;

                                            if (oDocument.Lines.BatchNumbers.BatchNumber == batch.DistNumber)
                                            {
                                                isBatchFound = true;
                                                break;
                                            }
                                        }

                                        if (isBatchFound)
                                        {
                                            oDocument.Lines.BatchNumbers.Quantity += double.Parse(batch.DraftQty.ToString());
                                        }
                                        else
                                        {
                                            oDocument.Lines.BatchNumbers.Add();
                                            oDocument.Lines.BatchNumbers.BatchNumber = batch.DistNumber;
                                            oDocument.Lines.BatchNumbers.Quantity = double.Parse(batch.DraftQty.ToString());
                                        }
                                    }

                                    isNext = true;
                                }
                            }

                        }
                    }

                    retcode = oDocument.Update();

                    if (retcode != 0)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                        string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                        Log($"{key }\n {failed_status }\n { message } \n");
                        ft_General.UpdateStatus(key, failed_status, message, CurrentDocNum);
                    }

                    #region Update Existing PickList

                    oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);

                    if (!oPickLists.GetByKey(int.Parse(dt.Rows[0]["sapDocNumber"].ToString())))
                        {
                            throw new Exception(sap.oCom.GetLastErrorDescription());
                        }

                        for (int d = 0; d < dt.Rows.Count; d++)
                        {

                            for (int x = 0; x < oPickLists.Lines.Count; x++)
                            {
                                oPickLists.Lines.SetCurrentLine(x);
                                if (oPickLists.Lines.LineNumber == int.Parse(dt.Rows[d]["BaseLine"].ToString()))
                                    break;
                            }

                            oPickLists.Lines.PickedQuantity = double.Parse(dt.Rows[d]["quantity"].ToString());
                            oPickLists.Lines.UserFields.Fields.Item("U_Weight").Value = double.Parse(dt.Rows[d]["LineWeight"].ToString());
                            totalWeight += double.Parse(dt.Rows[d]["LineWeight"].ToString());


                            DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[d]["key"].ToString() + "' and itemcode='" + dt.Rows[d]["itemcode"].ToString() + "' and BaseLine = '" + dt.Rows[d]["BaseLine"].ToString() + "'");

                            if (dr.Length > 0)
                            {
                                for (int x = 0; x < dr.Length; x++)
                                {
                                    if (dr[x]["batchnumber"].ToString() != "")
                                    {
                                        if (batch_cnt > 0) oPickLists.Lines.BatchNumbers.Add();

                                        oPickLists.Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();
                                        oPickLists.Lines.BatchNumbers.Quantity = double.Parse(dr[x]["Quantity"].ToString());
                                        oPickLists.Lines.BatchNumbers.BaseLineNumber = int.Parse(dr[x]["BaseLine"].ToString());
                                        batch_cnt++;
                                    }
                                }
                            }
                            batch_cnt = 0;
                        }

                        retcode = oPickLists.Update();

                        if (retcode != 0)
                        {
                            if (sap.oCom.InTransaction)
                                sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                            string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                            Log($"{key }\n {failed_status }\n { message } \n");
                            ft_General.UpdateStatus(key, failed_status, message, CurrentDocNum);
                        }
                        else
                        {

                            if (sap.oCom.InTransaction)
                                sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            UpdateTotalWeight(CurrentDocNum);
                            UpdateAllocateItemToPicked(CurrentDocNum);
                            Log($"{key }\n {success_status }\n  { CurrentDocNum } \n");
                            ft_General.UpdateStatus(key, success_status, "", CurrentDocNum);
                        }

                        if (oPickLists != null) Marshal.ReleaseComObject(oPickLists);
                        oPickLists = null;
                        if (oDocument != null) Marshal.ReleaseComObject(oDocument);
                        oDocument = null;
                        #endregion
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("OPKL", ex.Message);
                Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
                ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
            }
            finally
            {
                dt = null;
                dtDetails = null;
            }
        }

        static void LoadDataToDataTable(string request)
        {
            dt = ft_General.LoadData("LoadOPKL_sp");
            dt.DefaultView.Sort = "key, BaseLine";
            dt = dt.DefaultView.ToTable();

            dtDetails = ft_General.LoadDataByRequest("LoadDetails_sp2", request);
        }

        static List<AllocationItem> LoadPickTransaction(int absentry)
        {
            try
            {
                SqlConnection conn = new SqlConnection(Program._DbMidwareConnStr);

                var  list = conn.Query<AllocationItem>($"sp_PickList_GetPickLineBatches",
                                                new { absentry = absentry },
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

        static List<AllocationItem> LoadSOBatchTransaction(int absentry, int pickentry, int sodocentry, int solinenum)
        {
            try
            {
                SqlConnection conn = new SqlConnection(Program._DbMidwareConnStr);

                var list = conn.Query<AllocationItem>($"sp_PickList_GetSOLineBatches",
                                new { absentry = absentry, picklinenum = pickentry, docentry = sodocentry, solinenum = solinenum },
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

        static void UpdateTotalWeight(string PickNo)
        {
            try
            {
                SqlConnection conn = new SqlConnection(Program._DbErpConnStr);
                int result =  conn.Execute("zwa_IMApp_PickList_spSetTotalWeight", new { PickNo = int.Parse(PickNo) }, commandType: CommandType.StoredProcedure);

            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
            }
        }

        static void UpdateAllocateItemToPicked(string PickNo)
        {
            try
            {
                SqlConnection conn = new SqlConnection(Program._DbMidwareConnStr);
                int result = conn.Execute("sp_UpdatePickListAllocateItemToPicked", new { absentry = int.Parse(PickNo) }, commandType: CommandType.StoredProcedure);

            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
            }
        }
    }
 }

#region oldcode
//public async static void Post123()
//{
//    string request = "Update Pick List";
//    try
//    {
//        string failed_status = "ONHOLD";
//        string success_status = "SUCCESS";
//        int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;
//        int retcode = 0;
//        double totalWeight = 0;

//        LoadDataToDataTable(request);


//        if (dt.Rows.Count > 0)
//        {
//            string key = dt.Rows[0]["key"].ToString();
//            currentKey = key;
//            currentStatus = failed_status;
//            CurrentDocNum = dt.Rows[0]["sapDocNumber"].ToString();

//            par = SAP.GetSAPUser();
//            sap = SAP.getSAPCompany(par);

//            if (!sap.connectSAP())
//            {
//                Log($"{sap.errMsg}");
//                throw new Exception(sap.errMsg);
//            }

//            if (!sap.oCom.InTransaction)
//                sap.oCom.StartTransaction();

//            oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
//            oPickLists.GetByKey(int.Parse(dt.Rows[0]["sapDocNumber"].ToString()));
//            oPickLists_Lines = oPickLists.Lines;

//            for (int i = 0; i < dt.Rows.Count; i++)
//            {

//                for (int x = 0; x < oPickLists_Lines.Count; x++)
//                {
//                    oPickLists_Lines.SetCurrentLine(x);
//                    if (oPickLists_Lines.LineNumber == int.Parse(dt.Rows[i]["BaseLine"].ToString()))
//                        break;
//                }

//                //oPickLists_Lines.SetCurrentLine(int.Parse(dt.Rows[i]["BaseLine"].ToString()));
//                oPickLists_Lines.PickedQuantity = double.Parse(dt.Rows[i]["quantity"].ToString());
//                oPickLists_Lines.UserFields.Fields.Item("U_Weight").Value = double.Parse(dt.Rows[i]["LineWeight"].ToString());
//                totalWeight += double.Parse(dt.Rows[i]["LineWeight"].ToString());


//                DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "' and BaseLine = '" + dt.Rows[i]["BaseLine"].ToString() + "'");

//                if (dr.Length > 0)
//                {
//                    for (int x = 0; x < dr.Length; x++)
//                    {
//                        if (dr[x]["batchnumber"].ToString() != "")
//                        {    
//                            if (batch_cnt > 0) oPickLists_Lines.BatchNumbers.Add();

//                            oPickLists_Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();
//                            oPickLists_Lines.BatchNumbers.Quantity = double.Parse(dr[x]["Quantity"].ToString());
//                            oPickLists_Lines.BatchNumbers.BaseLineNumber = int.Parse(dr[x]["BaseLine"].ToString());
//                            batch_cnt++;    
//                        }
//                    }
//                }
//                batch_cnt = 0;
//            }

//            retcode = oPickLists.Update();

//            if (retcode != 0)
//            {
//                if (sap.oCom.InTransaction)
//                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

//                string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
//                Log($"{key }\n {failed_status }\n { message } \n");
//                ft_General.UpdateStatus(key, failed_status, message, CurrentDocNum);
//            }
//            else
//            {

//                if (sap.oCom.InTransaction)
//                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
//                UpdateTotalWeight(CurrentDocNum);
//                //RemoveOnholdItem();
//                Log($"{key }\n {success_status }\n  { CurrentDocNum } \n");
//                ft_General.UpdateStatus(key, success_status, "", CurrentDocNum);
//            }

//            if (oPickLists != null) Marshal.ReleaseComObject(oPickLists);
//            oPickLists = null;

//        }

//    }
//    catch (Exception ex)
//    {
//        Log($"{ ex.Message } \n");
//        ft_General.UpdateError("OPKL", ex.Message);
//        Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
//        ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
//    }
//    finally
//    {
//        dt = null;
//        dtDetails = null;
//    }
//}
#endregion

#region oldcode2
//for (int i = 0; i < dt.Rows.Count; i++)
//{
//    #region Start assign SO
//    var currentBatchlist = LoadPickTransaction(int.Parse(CurrentDocNum), int.Parse(dt.Rows[i]["SourceLineNum"].ToString()));

//    if (currentBatchlist.Sum(x => x.DraftQty) > 0)
//    {
//        var currentSODocEntry = currentBatchlist.FirstOrDefault().SODocEntry;
//        var currentSODocLineNum = currentBatchlist.FirstOrDefault().SOLineNum;

//        oDocument = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
//        if (!oDocument.GetByKey(currentSODocEntry))
//        {
//            throw new Exception(sap.oCom.GetLastErrorDescription());
//        }

//        for (int j = 0; j < currentBatchlist.Count; j++)
//        {
//            if (currentBatchlist[j].DraftQty <= 0) continue;

//            bool isSOLineFound = false;

//            for (int k = 0; k < oDocument.Lines.Count; k++)
//            {
//                oDocument.Lines.SetCurrentLine(k);

//                if (oDocument.Lines.LineNum == currentBatchlist[j].SOLineNum)
//                {
//                    isSOLineFound = true;
//                    break;
//                }
//            }

//            if (isSOLineFound)
//            {
//                bool isBatchFound = false;
//                var test = oDocument.Lines.BatchNumbers.Count;
//                var test1 = oDocument.Lines.BatchNumbers.BatchNumber;

//                if (oDocument.Lines.BatchNumbers.Count - 1 == 0 && oDocument.Lines.BatchNumbers.BatchNumber == "")
//                {
//                    oDocument.Lines.BatchNumbers.BatchNumber = currentBatchlist[j].DistNumber;
//                    oDocument.Lines.BatchNumbers.Quantity = double.Parse(currentBatchlist[j].DraftQty.ToString());
//                }
//                else
//                {
//                    for (int l = 0; l < oDocument.Lines.BatchNumbers.Count; l++)
//                    {
//                        oDocument.Lines.BatchNumbers.SetCurrentLine(l);

//                        if (oDocument.Lines.LineStatus == BoStatus.bost_Close) continue;

//                        if (oDocument.Lines.BatchNumbers.BatchNumber == currentBatchlist[j].DistNumber)
//                        {
//                            isBatchFound = true;
//                            break;
//                        }
//                    }

//                    if (isBatchFound)
//                    {
//                        oDocument.Lines.BatchNumbers.Quantity += double.Parse(currentBatchlist[j].DraftQty.ToString());
//                    }
//                    else
//                    {
//                        oDocument.Lines.BatchNumbers.Add();
//                        oDocument.Lines.BatchNumbers.BatchNumber = currentBatchlist[j].DistNumber;
//                        oDocument.Lines.BatchNumbers.Quantity = double.Parse(currentBatchlist[j].DraftQty.ToString());
//                    }
//                }

//            }
//        }

//        retcode = oDocument.Update();

//        if (retcode != 0)
//        {
//            if (sap.oCom.InTransaction)
//                sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

//            string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
//            Log($"{key }\n {failed_status }\n { message } \n");
//            ft_General.UpdateStatus(key, failed_status, message, CurrentDocNum);
//        }
//    }
//    #endregion

//    oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);

//    //#region Assigned Back Other Affected PickList
//    //var pickLists = LoadSOBatchTransaction(int.Parse(CurrentDocNum), int.Parse(dt.Rows[i]["SourceLineNum"].ToString()), int.Parse(dt.Rows[i]["BaseEntry"].ToString()), int.Parse(dt.Rows[i]["BaseLine"].ToString()));
//    //if (pickLists != null && pickLists.Count > 0 && pickLists.Sum(x => x.ActualPickQty) > 0)
//    //{
//    //    var distinctPickNoList = pickLists.Select(x => x.SODocEntry).Distinct().ToList();

//    //    foreach (var picklistNo in distinctPickNoList)
//    //    {
//    //        if (!oPickLists.GetByKey(picklistNo))
//    //        {
//    //            throw new Exception(sap.oCom.GetLastErrorDescription());
//    //        }

//    //        var currentPKBatches = pickLists.Where(x => x.PickListDocEntry == picklistNo && x.ActualPickQty > 0).ToList();

//    //        if (currentPKBatches == null || currentPKBatches.Count <= 0) continue;

//    //        for(int a=0; a <oPickLists.Lines.Count;a++)
//    //        {
//    //            oPickLists.Lines.SetCurrentLine(a);
//    //            bool isPickLinefound = false;
//    //            for (int b = 0; b < currentPKBatches.Count; b++) 
//    //            {
//    //                if(oPickLists.Lines.LineNumber == currentPKBatches[b].PickListLineNum)
//    //                {
//    //                    isPickLinefound = true;
//    //                    break;
//    //                }

//    //                if (isPickLinefound)
//    //                {
//    //                    if (oPickLists.Lines.PickStatus == BoPickStatus.ps_Closed) continue;


//    //                    if (oPickLists.Lines.BatchNumbers.Count>0)
//    //                    {
//    //                        var isPickBatchFound = false;

//    //                        for(int c = 0; c< oPickLists.Lines.BatchNumbers.Count; c++)
//    //                        {
//    //                            if(oPickLists.Lines.BatchNumbers.BatchNumber == currentPKBatches[b].DistNumber)
//    //                            {
//    //                                isPickBatchFound = true;
//    //                                break;
//    //                            }
//    //                        }

//    //                        if (isPickBatchFound)
//    //                        {
//    //                            oPickLists.Lines.BatchNumbers.Quantity = int.Parse(currentPKBatches[b].ActualPickQty.ToString());
//    //                        }
//    //                        else
//    //                        {
//    //                            oPickLists.Lines.BatchNumbers.Add();
//    //                            oPickLists.Lines.BatchNumbers.BatchNumber = currentPKBatches[b].DistNumber;
//    //                            oPickLists.Lines.BatchNumbers.Quantity = int.Parse(currentPKBatches[b].ActualPickQty.ToString());
//    //                            oPickLists.Lines.BatchNumbers.BaseLineNumber = oPickLists.Lines.LineNumber;
//    //                        }

//    //                    }
//    //                    else
//    //                    {
//    //                        oPickLists.Lines.BatchNumbers.BatchNumber = currentPKBatches[b].DistNumber;
//    //                        oPickLists.Lines.BatchNumbers.Quantity = int.Parse(currentPKBatches[b].ActualPickQty.ToString());
//    //                        oPickLists.Lines.BatchNumbers.BaseLineNumber = oPickLists.Lines.LineNumber;
//    //                    }
//    //                }
//    //            }
//    //        }
//    //    }

//    //    retcode = oPickLists.Update();

//    //    if (retcode != 0)
//    //    {
//    //        if (sap.oCom.InTransaction)
//    //            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

//    //        string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
//    //        Log($"{key }\n {failed_status }\n { message } \n");
//    //        ft_General.UpdateStatus(key, failed_status, message, CurrentDocNum);
//    //    }
//    //}
//    //#endregion
//}

#endregion
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


                    #region Update SO

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

                        if (oDocument.DocumentStatus == BoStatus.bost_Close) continue;

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

                                //bool isFirst = true;
                                foreach (var batch in batchList)
                                {
                                    if (batch.DraftQty <= 0) continue;

                                    if (oDocument.Lines.BatchNumbers.Count - 1 == 0 && oDocument.Lines.BatchNumbers.BatchNumber == "")
                                    {
                                        oDocument.Lines.BatchNumbers.BatchNumber = batch.DistNumber;
                                        oDocument.Lines.BatchNumbers.Quantity += double.Parse(batch.DraftQty.ToString());
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
                                }
                            }
                        }

                        retcode = oDocument.Update();

                        if (retcode != 0)
                        {
                            if (sap.oCom.InTransaction)
                                sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                            string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");

                            throw new Exception(message);
                        }
                    }
                    #endregion


                    #region Reset To PickList
                    oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
                    if (!oPickLists.GetByKey(int.Parse(dt.Rows[0]["sapDocNumber"].ToString())))
                    {
                        throw new Exception(sap.oCom.GetLastErrorDescription());
                    }
                    for (int x = 0; x < dt.Rows.Count; x++)
                    {
                        oPickLists.Lines.SetCurrentLine(x);
                        oPickLists.Lines.PickedQuantity = 0;

                        if (oPickLists.Lines.BatchNumbers.Count - 1 > 0 || oPickLists.Lines.BatchNumbers.BatchNumber != "")
                        {
                            for (int y = 0; y < oPickLists.Lines.BatchNumbers.Count; y++)
                            {
                                oPickLists.Lines.BatchNumbers.SetCurrentLine(y);

                                oPickLists.Lines.BatchNumbers.Quantity = 0;
                            }
                        }
                    }

                    retcode = oPickLists.Update();

                    if (retcode != 0)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                        string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");

                        throw new Exception(message);
                        //Log($"{key }\n {failed_status }\n { message } \n");
                        //ft_General.UpdateStatus(key, failed_status, message, CurrentDocNum);
                    }
                    #endregion

                    #region Update Existing PickList
                    oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);

                    if (!oPickLists.GetByKey(int.Parse(dt.Rows[0]["sapDocNumber"].ToString())))
                    {
                        throw new Exception(sap.oCom.GetLastErrorDescription());
                    }


                    for (int d = 0; d < dt.Rows.Count; d++)
                    {
                        bool isPickLineFound = false;

                        for (int x = 0; x < oPickLists.Lines.Count; x++)
                        {
                            oPickLists.Lines.SetCurrentLine(x);
                            if (oPickLists.Lines.LineNumber == int.Parse(dt.Rows[d]["BaseLine"].ToString()))
                            {
                                isPickLineFound = true;
                                break;
                            }
                        }

                        if (isPickLineFound)
                        {
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
                                        var batchcount = oPickLists.Lines.BatchNumbers.Count;

                                        if (oPickLists.Lines.BatchNumbers.Count - 1 == 0 && oPickLists.Lines.BatchNumbers.BatchNumber == "")
                                        {
                                            if (batch_cnt > 0) oPickLists.Lines.BatchNumbers.Add();

                                            oPickLists.Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();
                                            oPickLists.Lines.BatchNumbers.Quantity = double.Parse(dr[x]["Quantity"].ToString());
                                            oPickLists.Lines.BatchNumbers.BaseLineNumber = int.Parse(dr[x]["BaseLine"].ToString());
                                            batch_cnt++;
                                        }
                                        else
                                        {
                                            bool isPickBatchFound = false;
                                            for (int y = 0; y < oPickLists.Lines.BatchNumbers.Count; y++) 
                                            {
                                                oPickLists.Lines.BatchNumbers.SetCurrentLine(y);

                                                if(dr[x]["batchnumber"].ToString() == oPickLists.Lines.BatchNumbers.BatchNumber)
                                                {
                                                    isPickBatchFound = true;
                                                    break;  
                                                }
                                            }

                                            if (isPickBatchFound)
                                            {
                                                Console.WriteLine(dr[x]["batchnumber"].ToString() +" "+ double.Parse(dr[x]["Quantity"].ToString()));
                                                oPickLists.Lines.BatchNumbers.Quantity = double.Parse(dr[x]["Quantity"].ToString());
                                                batch_cnt++;
                                            }
                                            else
                                            {
                                                if (batch_cnt > 0) oPickLists.Lines.BatchNumbers.Add();
                                                //oPickLists.Lines.BatchNumbers.SetCurrentLine(oPickLists.Lines.BatchNumbers.Count - 1);

                                                oPickLists.Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();
                                                oPickLists.Lines.BatchNumbers.Quantity = double.Parse(dr[x]["Quantity"].ToString());
                                                oPickLists.Lines.BatchNumbers.BaseLineNumber = int.Parse(dr[x]["BaseLine"].ToString());
                                                batch_cnt++;
                                            }
                                        }
                                    }
                                }
                            }

                            batch_cnt = 0;
                        }
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

                var list = conn.Query<AllocationItem>($"sp_PickList_GetPickLineBatches",
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


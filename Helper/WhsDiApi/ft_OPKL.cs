using Dapper;
using IMAppSapMidware_NetCore.Helper.SQL;
using IMAppSapMidware_NetCore.Models.SAPModels;
using Microsoft.Data.SqlClient;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;

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

            static void Log(string message)
            {
                LastSAPMsg += $"\n\n{message}";
                Program.FilLogger?.Log(message);
            }

            public async static void Post()
            {
                string request = "Update Pick List";
                try
                {
                    string failed_status = "ONHOLD";
                    string success_status = "SUCCESS";
                    int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;
                    int retcode = 0;
                    double totalWeight=0;

                    LoadDataToDataTable(request);


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

                        if (!sap.oCom.InTransaction)
                            sap.oCom.StartTransaction();

                        oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
                        oPickLists.GetByKey(int.Parse(dt.Rows[0]["sapDocNumber"].ToString()));
                        CurrentDocNum = dt.Rows[0]["sapDocNumber"].ToString();
                        oPickLists_Lines = oPickLists.Lines;

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            for (int x = 0; x < oPickLists_Lines.Count; x++)
                            {
                                oPickLists_Lines.SetCurrentLine(x);
                                if (oPickLists_Lines.LineNumber == int.Parse(dt.Rows[i]["BaseLine"].ToString()))
                                    break;
                            }

                            //oPickLists_Lines.SetCurrentLine(int.Parse(dt.Rows[i]["BaseLine"].ToString()));
                            oPickLists_Lines.PickedQuantity = double.Parse(dt.Rows[i]["quantity"].ToString());
                            oPickLists_Lines.UserFields.Fields.Item("U_Weight").Value = double.Parse(dt.Rows[i]["LineWeight"].ToString());
                            totalWeight += double.Parse(dt.Rows[i]["LineWeight"].ToString());


                            DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "' and BaseLine = '" + dt.Rows[i]["BaseLine"].ToString() + "'");

                            if (dr.Length > 0)
                            {
                                for (int x = 0; x < dr.Length; x++)
                                {
                                    if (dr[x]["batchnumber"].ToString() != "")
                                    {
                                        if (batch_cnt > 0) oPickLists_Lines.BatchNumbers.Add();

  
                                        oPickLists_Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();
                                        oPickLists_Lines.BatchNumbers.Quantity = double.Parse(dr[x]["Quantity"].ToString());
                                        oPickLists_Lines.BatchNumbers.BaseLineNumber = int.Parse(dr[x]["BaseLine"].ToString());
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
                            RemoveOnholdItem();
                            Log($"{key }\n {success_status }\n  { CurrentDocNum } \n");
                            ft_General.UpdateStatus(key, success_status, "", CurrentDocNum);
                        }

                    if (oPickLists != null) Marshal.ReleaseComObject(oPickLists);
                        oPickLists = null;

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

        static void RemoveOnholdItem()
            {
                try
                {
                    SqlConnection conn = new SqlConnection(Program._DbMidwareConnStr);
                    int result = 0;
                    string deleteQuery = "DELETE FROM zmwSOHoldPickItem WHERE PickListDocEntry = @docEntry and PickListLineNum = @LineNum";

                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            result = conn.Execute(deleteQuery, new { docEntry = dt.Rows[i]["SourceDocEntry"].ToString() , LineNum = dt.Rows[i]["SourceLineNum"] });
                        }
                    }

                }
                catch (Exception ex)
                {
                     Log($"{ ex.Message } \n");
                }
            }


    }
    }
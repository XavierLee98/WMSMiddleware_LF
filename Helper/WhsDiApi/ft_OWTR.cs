﻿using Dapper;
using IMAppSapMidware_NetCore.Models.SAPModels;
using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;

namespace IMAppSapMidware_NetCore.Helper.SQL
{
    class ft_OWTR
    {
        public static string LastSAPMsg { get; set; } = string.Empty;

        // added by jonny to track error when unexpected error
        // 20210411
        static string currentKey = string.Empty;
        static string currentStatus = string.Empty;
        static string CurrentDocNum = string.Empty;

        static void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public static void Post()
        {
            DataTable dt = null;
            DataTable dtDetails = null;
            DataTable dtBinFrom = null;
            DataTable dtBinTo = null;
            //string sapdb = "SBODEMOUS2";
            string sapdb = Program._ErpDbName;
            string request = "Create Transfer1";

            try
            {
                dt = ft_General.LoadData("LoadOWTR_sp");
                dtDetails = ft_General.LoadDataByRequest("LoadTransferDetails_sp", request);
                dtBinFrom = ft_General.LoadDataByRequest("LoadTransferFromBinDetails_sp", request);
                dtBinTo = ft_General.LoadDataByRequest("LoadTransferToBinDetails_sp", request);
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                string tablename = "OWTR";
                string docnum = "";
                string docEntry = "";
                int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;
                int retcode = 0;

                if (dt.Rows.Count > 0)
                {
                    SAPParam par = SAP.GetSAPUser();
                    SAPCompany sap = SAP.getSAPCompany(par);

                    if (!sap.connectSAP()) throw new Exception(sap.errMsg);

                    string key = dt.Rows[0]["key"].ToString();
                    // added by jonny to track error when unexpected error
                    // 20210411
                    currentKey = key;
                    currentStatus = failed_status;

                    //SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.StockTransfer oDoc = null;// (SAPbobsCOM.StockTransfer)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (!sap.oCom.InTransaction)
                            sap.oCom.StartTransaction();

                        if (cnt > 0)
                        {
                            oDoc.Lines.Add();
                            oDoc.Lines.SetCurrentLine(cnt);

                            if (key == dt.Rows[i]["key"].ToString()) goto details;

                            retcode = oDoc.Add();// Add record 
                            if (retcode != 0) // if error
                            {
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                                Log($"{key }\n {failed_status }\n { message } \n");
                                ft_General.UpdateStatus(key, failed_status, message, "");
                            }
                            else
                            {
                                sap.oCom.GetNewObjectCode(out docEntry);
                                docnum = ft_General.GetDocNum(sap.oCom, tablename, docEntry);
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                Log($" {key }\n {success_status }\n  { docnum } \n");
                                ft_General.UpdateStatus(key, success_status, "", docnum);
                            }

                            cnt = 0;
                            if (oDoc != null) Marshal.ReleaseComObject(oDoc);
                            oDoc = null;
                        }

                        if (!sap.oCom.InTransaction)
                            sap.oCom.StartTransaction();

                        oDoc = (SAPbobsCOM.StockTransfer)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                        //if (!string.IsNullOrEmpty(dt.Rows[i]["Series"].ToString()))
                        //    oDoc.Series = int.Parse(dt.Rows[i]["Series"].ToString());
                        //oDoc.CardCode = dt.Rows[i]["cardcode"].ToString();
                        oDoc.DocDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
                        oDoc.TaxDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
                        if (dt.Rows[i]["Comments"].ToString() != "")
                            oDoc.Comments = dt.Rows[i]["Comments"].ToString();
                        if (dt.Rows[i]["JrnlMemo"].ToString() != "")
                        oDoc.JournalMemo = dt.Rows[i]["JrnlMemo"].ToString();
                        if (dt.Rows[i]["PriceList"].ToString() != "")
                            oDoc.PriceList = ConvertPriceListToInt(dt.Rows[i]["PriceList"].ToString());
                    //oDoc.FromWarehouse = dt.Rows[i]["fromwarehouse"].ToString();

                    details:
                        oDoc.Lines.ItemCode = dt.Rows[i]["itemcode"].ToString();
                        oDoc.Lines.Quantity = double.Parse(dt.Rows[i]["quantity"].ToString());
                        oDoc.Lines.FromWarehouseCode = dt.Rows[i]["WarehouseFrom"].ToString();
                        oDoc.Lines.WarehouseCode = dt.Rows[i]["WarehouseTo"].ToString();


                        if (int.Parse(dt.Rows[i]["baseentry"].ToString()) > 0)
                        {
                            oDoc.Lines.BaseEntry = int.Parse(dt.Rows[i]["baseentry"].ToString());
                            oDoc.Lines.BaseLine = int.Parse(dt.Rows[i]["baseline"].ToString());
                            oDoc.Lines.BaseType = SAPbobsCOM.InvBaseDocTypeEnum.InventoryTransferRequest;
                        }

                        //DataTable dtBinBatchSerial = ft_General.LoadBinBatchSerial(dt.Rows[i]["key"].ToString(), dt.Rows[i]["itemcode"].ToString());
                        DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");
                        if (dr.Length > 0)
                        {
                            for (int x = 0; x < dr.Length; x++)
                            {
                                if (dr[x]["batchnumber"].ToString() != "")
                                {
                                    if (batch_cnt > 0) oDoc.Lines.BatchNumbers.Add();
                                    oDoc.Lines.BatchNumbers.SetCurrentLine(batch_cnt);
                                    oDoc.Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();
                                    oDoc.Lines.BatchNumbers.Quantity = double.Parse(decimal.Parse(dr[x]["quantity"].ToString()).ToString());

                                    DataRow[] drBinFrom = dtBinFrom.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                        "' and Batch ='" + dr[x]["batchnumber"].ToString() + "'");

                                    if (drBinFrom.Length > 0)
                                    {
                                        for (int y = 0; y < drBinFrom.Length; y++)
                                        {
                                            if (drBinFrom[y]["binabs"].ToString() == "-1") continue;
                                            if (batchbin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                            oDoc.Lines.BinAllocations.SetCurrentLine(batchbin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBinFrom[y]["binabs"].ToString());
                                            oDoc.Lines.BinAllocations.Quantity = double.Parse(drBinFrom[y]["qty"].ToString());
                                            oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = batch_cnt;
                                            oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batFromWarehouse;
                                            batchbin_cnt++;
                                        }
                                    }
                                    //batchbin_cnt = 0;

                                    DataRow[] drBinTo = dtBinTo.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                        "' and Batch ='" + dr[x]["batchnumber"].ToString() + "'");

                                    if (drBinTo.Length > 0)
                                    {
                                        for (int y = 0; y < drBinTo.Length; y++)
                                        {
                                            if (drBinTo[y]["binabs"].ToString() == "-1") continue;
                                            if (batchbin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                            oDoc.Lines.BinAllocations.SetCurrentLine(batchbin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBinTo[y]["binabs"].ToString());
                                            oDoc.Lines.BinAllocations.Quantity = double.Parse(drBinTo[y]["qty"].ToString());
                                            oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = batch_cnt;
                                            oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batToWarehouse;
                                            batchbin_cnt++;
                                        }
                                    }
                                    //batchbin_cnt = 0;
                                    batch_cnt++;
                                }
                                else if (dr[x]["serialnumber"].ToString() != "")
                                {
                                    if (serial_cnt > 0) oDoc.Lines.SerialNumbers.Add();
                                    oDoc.Lines.SerialNumbers.SetCurrentLine(serial_cnt);
                                    oDoc.Lines.SerialNumbers.InternalSerialNumber = dr[x]["serialnumber"].ToString();
                                    DataRow[] drBinFrom = dtBinFrom.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                        "' and serial ='" + dr[x]["serialnumber"].ToString() + "' ");

                                    if (drBinFrom.Length > 0 && drBinFrom[0]["binabs"].ToString() != "-1")
                                    {
                                        if (serial_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                        oDoc.Lines.BinAllocations.SetCurrentLine(serial_cnt);
                                        oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBinFrom[0]["binabs"].ToString());
                                        oDoc.Lines.BinAllocations.Quantity = double.Parse(drBinFrom[0]["qty"].ToString());
                                        oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = serial_cnt;
                                        oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batFromWarehouse;
                                    }

                                    DataRow[] drBinTo = dtBinTo.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                       "' and serial ='" + dr[x]["serialnumber"].ToString() + "' ");

                                    if (drBinTo.Length > 0 && drBinTo[0]["binabs"].ToString() != "-1")
                                    {
                                        if (serial_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                        oDoc.Lines.BinAllocations.SetCurrentLine(serial_cnt);
                                        oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBinTo[0]["binabs"].ToString());
                                        oDoc.Lines.BinAllocations.Quantity = double.Parse(drBinTo[0]["qty"].ToString());
                                        oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = serial_cnt;
                                        oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batToWarehouse;
                                    }

                                    serial_cnt++;
                                }
                                else
                                {
                                    DataRow[] drBinFrom = dtBinFrom.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");

                                    if (drBinFrom.Length > 0)
                                    {
                                        for (int y = 0; y < drBinFrom.Length; y++)
                                        {
                                            if (drBinFrom[y]["binabs"].ToString() == "-1") continue;
                                            if (bin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                            oDoc.Lines.BinAllocations.SetCurrentLine(bin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBinFrom[y]["binabs"].ToString());
                                            oDoc.Lines.BinAllocations.Quantity = double.Parse(drBinFrom[y]["qty"].ToString());
                                            oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batFromWarehouse;
                                            bin_cnt++;
                                        }
                                    }
                                    bin_cnt = 0;
                                    DataRow[] drBinTo = dtBinTo.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");

                                    if (drBinTo.Length > 0)
                                    {
                                        for (int y = 0; y < drBinTo.Length; y++)
                                        {
                                            if (drBinTo[y]["binabs"].ToString() == "-1") continue;
                                            if (bin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                            oDoc.Lines.BinAllocations.SetCurrentLine(bin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBinTo[y]["binabs"].ToString());
                                            oDoc.Lines.BinAllocations.Quantity = double.Parse(drBinTo[y]["qty"].ToString());
                                            oDoc.Lines.BinAllocations.BinActionType = SAPbobsCOM.BinActionTypeEnum.batToWarehouse;
                                            bin_cnt++;
                                        }
                                    }

                                }
                            }
                            bin_cnt = 0;
                            serial_cnt = 0;
                            batch_cnt = 0;
                            batchbin_cnt = 0;
                        }

                        key = dt.Rows[i]["key"].ToString();
                        cnt++;
                    }
                    retcode = oDoc.Add();
                    if (retcode != 0)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                        Log($"{key }\n {failed_status }\n { message } \n");
                        ft_General.UpdateStatus(key, failed_status, message, "");
                    }
                    else
                    {
                        sap.oCom.GetNewObjectCode(out docEntry);
                        docnum = ft_General.GetDocNum(sap.oCom, tablename, docEntry);
                        
                        // added by jonny to track error when unexpected error
                        // 20210411
                        CurrentDocNum = docnum;

                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                        Log($" {key }\n {success_status }\n  { docnum } \n");
                        ft_General.UpdateStatus(key, success_status, "", docnum);
                    }

                    if (oDoc != null) Marshal.ReleaseComObject(oDoc);
                    oDoc = null;
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("OWTR", ex.Message);

                // needed add in this to prevent unexpected error
                Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
                ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
            }
            finally
            {
                dt = null;
                dtDetails = null;
                dtBinFrom = null;
                dtBinTo = null;
            }
        }

        static int ConvertPriceListToInt(string pricename)
        {
            var conn = new Microsoft.Data.SqlClient.SqlConnection(Program._DbErpConnStr);
            string query = "SELECT ListNum FROM OPLN WHERE ListName = @pricename";

            return conn.Query<int>(query, new { pricename = pricename }).FirstOrDefault();
        }
    }
}

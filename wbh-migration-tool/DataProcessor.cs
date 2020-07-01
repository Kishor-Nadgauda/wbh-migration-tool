using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Net.Http;
using System.Dynamic;
using System.Data;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Text;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Flurl;
using Flurl.Http;
using Independentsoft.Msg;

namespace PRGX.Panoptic_Migration_Core
{
    class DataProcessor
    {
        public static ConcurrentDictionary<string, int> CompletedCountList =
        new ConcurrentDictionary<string, int>(new Dictionary<string, int>());
        public static int ProjectId;
        public static string MigrationClientCode;

        public void ProcessData()
        {
            try
            {
                Preprocessor.Preprocess();
                DataTable dtSource = null;
                Stopwatch sw = new Stopwatch();

                DataSet ds = Util.GetMigrationConfigurations();
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    DataTable ProjTbl = ds.Tables[0].DefaultView.ToTable(true, "ClientCode", "ProjectId")
                        .AsEnumerable().CopyToDataTable();
                    foreach (DataRow ProjRow in ProjTbl.Rows)
                    {
                        MigrationClientCode = ProjRow["ClientCode"].ToString();
                        ProjectId = Convert.ToInt32(ProjRow["ProjectId"]);
                        CompletedCountList.Clear();
                        var ModuleRS = ds.Tables[0].AsEnumerable().Where(x => x.Field<string>("ClientCode") == MigrationClientCode
                        && x.Field<int>("ProjectId") == ProjectId);
                        foreach (DataRow ModuleRow in ModuleRS)
                        {
                            string entityName = ModuleRow["Name"].ToString();
                            int iterationCount = Convert.ToInt32(ModuleRow["IterationCount"]);
                            int batchSize = Convert.ToInt32(ModuleRow["BatchSize"]);
                            CompletedCountList.TryAdd(entityName, 0);
                            sw.Restart();
                            Logger.log.Info($"Processing {entityName}s...");
                            int cnt = iterationCount;
                            int ttl = 0;
                            while (cnt > 0 && batchSize > 0)
                            {
                                dtSource = Util.GetMigrationData(entityName, MigrationClientCode, ProjectId, ModuleRow["SPName"].ToString(), batchSize).Tables[0];
                                int rowCount = dtSource.Rows.Count;
                                if (dtSource != null && rowCount > 0)
                                {
                                    Logger.log.Info($"Processing Interation {iterationCount - cnt + 1}, Count: {rowCount} ... ");
                                    ttl += dtSource.Rows.Count;
                                    Parallel.ForEach<DataRow>(dtSource.AsEnumerable(),
                                    new ParallelOptions()
                                    {
                                        //Sets the maximum number of concurrent tasks
                                        MaxDegreeOfParallelism = Convert.ToInt32(ModuleRow["ThreadCount"])
                                    },
                                    (row =>
                                    {
                                        ProcessRow(entityName, row, ModuleRow["CustomFields"].ToString());
                                        Console.Write($"\r{((CompletedCountList[entityName] - ((iterationCount - cnt) * batchSize)) * 100 / rowCount)}%");
                                    }));
                                }
                                else
                                {
                                    break;
                                }
                                --cnt;
                                Console.WriteLine();
                            }
                            if (ttl > 0)
                            {
                                Logger.log.Info($"{entityName}s {CompletedCountList[entityName]} of {ttl} Processed Successfully in {sw.Elapsed.TotalMinutes} Minutes.");
                            }
                            else
                            {
                                Logger.log.Warn($"No records found to Process {entityName}s");
                            }
                        }
                    }
                }
                else
                {
                    Logger.log.Warn($"No Client or Project Information found to Process");
                }
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"HTTP Exception Occurred while Processing Data. Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, "Error Occurred while Processing Data.");
                    }
                    return false; //break execution.              
                });
            }
        }

        private void ProcessRow(string EntityName, DataRow row, string CustomFields)
        {
            int statId = 0;
            int VendorId = Convert.ToInt32(row["VenId"]);
            string RecordId = row["RecordId"].ToString();
            string ParentId = string.Empty;
            if (row.Table.Columns.Contains("ParentId") && row["ParentId"] != null) ParentId = row["ParentId"].ToString();
            try
            {
                Logger.log.Debug($"Started Processing {EntityName}. Vendor Id: {VendorId}. Task: {Task.CurrentId}");
                statId = Util.AddMigrationStatusRecord(VendorId, ParentId, RecordId, EntityName, MigrationClientCode, ProjectId, StatusCode.Inprocess.ToString(), "");
                if (statId > 0)
                {
                    APIReturnValue retVal = new APIReturnValue();
                    switch (EntityName)
                    {
                        case "Solicitation": retVal = ProcessVendorSolicitation(row, CustomFields); break;
                        case "SolicitationComment":
                        case "CreditComment":
                        case "ClaimComment":
                            retVal = ProcessComments(EntityName, row, CustomFields); break;
                        case "SolicitationDocument":
                        case "CreditDocument":
                            retVal = ProcessDocuments(EntityName, row, CustomFields); break;
                        case "Credit": retVal = ProcessVendorCredit(row, CustomFields); break;
                        case "CreditClaim": retVal = ProcessCreditClaim(row, CustomFields); break;
                        case "AdhocClaim": retVal = ProcessAdhocClaim(row, CustomFields); break;
                    }
                    switch (retVal.Status)
                    {
                        case StatusCode.Complete:
                            {
                                Util.UpdateMigrationStatusRecord(statId, retVal.PanopticId, StatusCode.Complete.ToString(), null);
                                Logger.log.Debug($"{EntityName} Processed. Panoptic Id: {retVal.PanopticId}. Response: {retVal.StatusDesc}");
                                CompletedCountList.AddOrUpdate(EntityName, 1, (key, oldValue) => oldValue + 1);
                                break;
                            }
                        case StatusCode.PartiallyComplete:
                            {
                                Util.UpdateMigrationStatusRecord(statId, retVal.PanopticId, "Partially Complete", retVal.StatusDesc);
                                Logger.log.Debug($"{EntityName} Processed Partially. Panoptic Id: {retVal.PanopticId}. Response: {retVal.StatusDesc}");
                                break;
                            }
                        case StatusCode.Error:
                            {
                                Util.UpdateMigrationStatusRecord(statId, null, StatusCode.Error.ToString(), retVal.StatusDesc);
                                Logger.log.Debug($"{EntityName} Processing Failed. Response: {retVal.StatusDesc}");
                                break;
                            }
                    }
                }
                Logger.log.Debug($"Finished Processing {EntityName}. Vendor Id: {VendorId}. Task: {Task.CurrentId}");
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"Error Occurred while Processing {EntityName} for Vendor Id: {VendorId}. Response: {es}");
                        if (statId > 0) Util.UpdateMigrationStatusRecord(statId, null, StatusCode.Error.ToString(), $"{e.Message} ,Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, $"Error Occurred while Processing {EntityName} for Vendor Id: {VendorId}");
                        if (statId > 0) Util.UpdateMigrationStatusRecord(statId, null, StatusCode.Error.ToString(), e.Message);
                    }
                    return true; //continue execution... set to false to break.              
                });
            }
            catch (Exception e)
            {
                Logger.log.Error(e, $"Error Occurred while Processing {EntityName} for Vendor Id: {VendorId}");
                if (statId > 0) Util.UpdateMigrationStatusRecord(statId, null, StatusCode.Error.ToString(), e.Message);
            }
        }

        private APIReturnValue ProcessVendorSolicitation(DataRow VSRow, string CustomFields)
        {
            APIReturnValue retVal = new APIReturnValue();
            int VendorId = Convert.ToInt32(VSRow["venId"]);
            string strResp = string.Empty;
            VendorSolicitationRequest objVendSol = Util.CreateObjectFromRow<VendorSolicitationRequest>(VSRow);
            objVendSol.projectVendor.projectVendorStatusId = Preprocessor.projectVendorStatusList[VSRow["ProjectVendorStatus"].ToString().ToLower()];
            objVendSol.projectVendor.projectVendorId = Preprocessor.GetProjectVendorId(Preprocessor.vendorMasterInfoList[VendorId].appVendorId, ProjectId);
            if (VSRow.Table.Columns.Contains("ActionTakenValue") && VSRow["ActionTakenValue"] != null)
            {
                objVendSol.actionTakenValueId = Preprocessor.vendorStatementActionList[VSRow["ActionTakenValue"].ToString()];
            }
            HttpResponseMessage respMsg;
            if (!string.IsNullOrEmpty(CustomFields))
            {
                dynamic expando = Util.ConvertToExpando(objVendSol);
                List<String> lst = CustomFields.Split(",").ToList();
                Dictionary<string, object> tempProps = new Dictionary<string, object>();
                foreach (string str in lst)
                {
                    if (VSRow.Table.Columns.Contains(str) && VSRow[str] != null)
                    {
                        tempProps.Add(str, VSRow[str]);
                    }
                }
                Util.AddProperties(expando, tempProps);
                string strPL = JsonConvert.SerializeObject(expando);
                respMsg = Util.AuditAPIURL.AppendPathSegment("projectvendorstatementactions").WithOAuthBearerToken(Util.bearerToken)
                                        .WithHeader("Content-Type", "application/json").PostStringAsync(strPL).Result;
            }
            else
            {
                respMsg = Util.AuditAPIURL.AppendPathSegment("projectvendorstatementactions").WithOAuthBearerToken(Util.bearerToken)
                                                    .WithHeader("Content-Type", "application/json").PostJsonAsync(objVendSol).Result;
            }
            strResp = respMsg.Content.ReadAsStringAsync().Result;
            retVal.StatusDesc = strResp;
            if (respMsg.IsSuccessStatusCode)
            {
                IDictionary<string, object> objResp = JsonConvert.DeserializeObject<ExpandoObject>(strResp);
                retVal.Status = StatusCode.Complete;
                retVal.PanopticId = objResp["projectVendorStatementActionId"].ToString();
            }
            else
            {
                retVal.Status = StatusCode.Error;
            }
            return retVal;
        }

        private APIReturnValue ProcessComments(string EntityName, DataRow Row, string CustomFields)
        {
            APIReturnValue retVal = new APIReturnValue();
            int ParentId;
            string qry = string.Empty;
            switch (EntityName)
            {
                case "SolicitationComment":
                    {
                        ParentId = Convert.ToInt32(Preprocessor.GetProjectVendorId(Preprocessor.vendorMasterInfoList[Convert.ToInt32(Row["VenId"])].appVendorId, ProjectId));
                        qry = $"vendorSolicitation({ParentId})/comment-attach/comments";
                        break;
                    }
                case "CreditComment":
                    {
                        ParentId = Convert.ToInt32(Row["PanopticParentId"]);
                        qry = $"vendorCredits({ParentId})/comment-attach/comments";
                        break;
                    }
                case "ClaimComment":
                    {
                        ParentId = Convert.ToInt32(Row["PanopticParentId"]);
                        qry = $"claims({ParentId})/comment-attach/comments";
                        break;
                    }
            }
            string strResp = string.Empty;
            CommentRequest objComReq = Util.CreateObjectFromRow<CommentRequest>(Row);
            HttpResponseMessage respMsg;
            if (!string.IsNullOrEmpty(CustomFields))
            {
                dynamic expando = Util.ConvertToExpando(objComReq);
                List<String> lst = CustomFields.Split(",").ToList();
                Dictionary<string, object> tempProps = new Dictionary<string, object>();
                foreach (string str in lst)
                {
                    if (Row.Table.Columns.Contains(str) && Row[str] != null)
                    {
                        tempProps.Add(str, Row[str]);
                    }
                }
                Util.AddProperties(expando, tempProps);
                string strPL = JsonConvert.SerializeObject(expando);
                respMsg = Util.AuditAPIURL.AppendPathSegment(qry)
                    .WithOAuthBearerToken(Util.bearerToken).WithHeader("Content-Type", "application/json")
                    .PostStringAsync(strPL).Result;
            }
            else
            {
                respMsg = Util.AuditAPIURL.AppendPathSegment(qry)
                                .WithOAuthBearerToken(Util.bearerToken).WithHeader("Content-Type", "application/json")
                                .PostJsonAsync(objComReq).Result;
            }
            strResp = respMsg.Content.ReadAsStringAsync().Result;
            retVal.StatusDesc = strResp;
            if (respMsg.IsSuccessStatusCode)
            {
                IDictionary<string, object> objResp = JsonConvert.DeserializeObject<ExpandoObject>(strResp);
                retVal.Status = StatusCode.Complete;
                retVal.PanopticId = objResp["id"].ToString();
            }
            else
            {
                retVal.Status = StatusCode.Error;
            }
            return retVal;
        }

        private APIReturnValue ProcessDocuments(string EntityName, DataRow Row, string CustomFields)
        {
            List<string> lstFailedFiles = new List<string>();
            List<string> lstPanIds = new List<string>();
            //string tempFilePath = Path.Combine("temp", Path.GetFileName(Row["filePath"].ToString()));
            string tempFilePath = Row["filePath"].ToString();
            string APIPath = string.Empty;
            APIReturnValue retVal = new APIReturnValue();
            string AppVenId = Preprocessor.vendorMasterInfoList[Convert.ToInt32(Row["venId"])].appVendorId;
            string strResp = string.Empty;
            //File.Copy(Row["filePath"].ToString(), tempFilePath, true);
            FileMetadata metadata = Util.CreateObjectFromRow<FileMetadata>(Row);
            metadata.appVendorId = AppVenId;
            metadata.projectId = ProjectId;
            switch (EntityName)
            {
                case "SolicitationDocument":
                    {
                        APIPath = $"projects({ProjectId})/doc-attach/objectstore/documents";
                        metadata.projectVendorId = Preprocessor.GetProjectVendorId(AppVenId, ProjectId);
                        break;
                    }
                case "CreditDocument":
                    {
                        APIPath = $"vendorCredits({Convert.ToInt32(Row["PanopticParentId"])})/doc-attach/objectstore/documents";
                        break;
                    }
            }
            if (Row.Table.Columns.Contains("DocumentType") && Row["DocumentType"] != null)
            {
                metadata.docType = Preprocessor.documentTypeList[Row["DocumentType"].ToString().ToLower()];
            }

            Dictionary<string, Stream> files = null;

            if (Path.GetExtension(tempFilePath).ToLower() == ".msg")
            {
                //files = GetFileStreams(tempFilePath);
                files = ExtractMSG(tempFilePath);
                Logger.log.Debug($"Get File Stream finished for {tempFilePath}");
            }
            else
            {
                files = new Dictionary<string, Stream>();
                files.Add(Path.GetFileName(tempFilePath), new FileStream(tempFilePath, FileMode.Open, FileAccess.Read));
            }

            if (files != null)
            {
                foreach (KeyValuePair<string, Stream> kvp in files)
                {
                    try
                    {
                        metadata.fileName = kvp.Key;
                        metadata.aliasName = kvp.Key;
                        Dictionary<string, string> kvfn = new Dictionary<string, string>();
                        if (!string.IsNullOrEmpty(CustomFields))
                        {
                            dynamic expando = Util.ConvertToExpando(metadata);
                            List<String> lst = CustomFields.Split(",").ToList();
                            Dictionary<string, object> tempProps = new Dictionary<string, object>();
                            foreach (string str in lst)
                            {
                                if (Row.Table.Columns.Contains(str) && Row[str] != null)
                                {
                                    tempProps.Add(str, Row[str]);
                                }
                            }
                            Util.AddProperties(expando, tempProps);
                            kvfn.Add(metadata.fileName, JsonConvert.SerializeObject(expando));
                        }
                        else
                        {
                            kvfn.Add(metadata.fileName, JsonConvert.SerializeObject(metadata));
                        }
                        Dictionary<string, Dictionary<string, string>> fileMetadata = new Dictionary<string, Dictionary<string, string>>();
                        fileMetadata.Add("fileMetadata", kvfn);
                        string s = JsonConvert.SerializeObject(fileMetadata);
                        /*
                        using (FileStream outputFileStream = new FileStream(kvp.Key, FileMode.Create))
                        {
                            kvp.Value.CopyTo(outputFileStream);
                        }
                        */
                        HttpResponseMessage respMsg = Util.AuditAPIURL
                        .AppendPathSegment(APIPath).WithOAuthBearerToken(Util.bearerToken)
                        .PostMultipartAsync(mp => mp.AddStringParts(new { request = s }).AddFile("file", kvp.Value, kvp.Key)).Result;
                        strResp = respMsg.Content.ReadAsStringAsync().Result;
                        Logger.log.Debug($"Upload finished for {kvp.Key}");
                        if (respMsg.IsSuccessStatusCode)
                        {
                            List<Dictionary<string, object>> objResp = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(strResp);
                            retVal.Status = StatusCode.PartiallyComplete;
                            lstPanIds.Add(objResp[0]["documentId"].ToString());
                        }
                        else
                        {
                            retVal.StatusDesc += strResp;
                            lstFailedFiles.Add(kvp.Key);
                            Logger.log.Error($"Error Occurred while Processing File {kvp.Key} for Record Id: {Row["RecordId"]}, Response: {strResp}");
                        }
                    }
                    catch (AggregateException ae)
                    {
                        lstFailedFiles.Add(kvp.Key);
                        ae.Handle(e =>
                        {
                            if (e is FlurlHttpException)
                            {
                                string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                                retVal.StatusDesc += es;
                                Logger.log.Error(e, $"Error Occurred while Processing File {kvp.Key} for Record Id: {Row["RecordId"]}, Response: {es}");
                            }
                            else
                            {
                                Logger.log.Error(e, $"Error Occurred while Processing File {kvp.Key} for Record Id: {Row["RecordId"]}");
                                retVal.StatusDesc += strResp;
                            }
                            return true; //continue execution... set to false to break.              
                        });
                    }
                    catch (Exception e)
                    {
                        lstFailedFiles.Add(kvp.Key);
                        retVal.StatusDesc += strResp;
                        Logger.log.Error(e, $"Error Occurred while Processing File {kvp.Key} for Record Id: {Row["RecordId"]}");
                    }
                }
            }
            else
            {
                throw new Exception("Error Occurred while Reading File Stream");
            }
            if ((lstPanIds.Count > 0) && (lstFailedFiles.Count == 0))
            {
                retVal.Status = StatusCode.Complete;
            }
            else if ((lstPanIds.Count > 0) && (lstFailedFiles.Count > 0))
            {
                retVal.Status = StatusCode.PartiallyComplete;
            }
            else
            {
                retVal.Status = StatusCode.Error;
            }
            if (lstFailedFiles.Count > 0) retVal.StatusDesc = $"Failed Files: {string.Join(",", lstFailedFiles)}, Response: {retVal.StatusDesc}";
            if (lstPanIds.Count > 0) retVal.PanopticId = string.Join(",", lstPanIds);
            //File.Delete(tempFilePath);
            return retVal;
        }

        private APIReturnValue ProcessVendorCredit(DataRow VCRow, string CustomFields)
        {
            APIReturnValue retVal = new APIReturnValue();
            int VendorId = Convert.ToInt32(VCRow["venId"]);
            string strResp = string.Empty;
            VendorCreditRequest objVendCredit = Util.CreateObjectFromRow<VendorCreditRequest>(VCRow);
            objVendCredit.projectId = ProjectId;
            objVendCredit.vendor.appVendorId = Preprocessor.vendorMasterInfoList[VendorId].appVendorId;
            if (!string.IsNullOrEmpty(VCRow["AccountNumber"].ToString()))
            {
                VendorAccountEntity objVAE = new VendorAccountEntity();
                objVAE.appVendorAccountId = Preprocessor.vendorMasterInfoList[VendorId].vendorAccounts[VCRow["AccountNumber"].ToString().ToLower()];
                objVendCredit.vendorAccountEntity = objVAE;
            }
            objVendCredit.creditStatus.id = Preprocessor.auditFieldLookupValues["creditStatus"][VCRow["CreditStatusValue"].ToString().ToLower()];
            objVendCredit.problemType.id = Preprocessor.auditFieldLookupValues["problemType"][VCRow["ProblemTypeValue"].ToString().ToLower()];
            if (VCRow.Table.Columns.Contains("ProblemSubTypeValue") && VCRow["ProblemSubTypeValue"] != null)
            {
                objVendCredit.problemSubType = new ProblemSubType();
                objVendCredit.problemSubType.id = Preprocessor.auditFieldLookupValues["problemSubType"][VCRow["ProblemSubTypeValue"].ToString().ToLower()];
            }
            HttpResponseMessage respMsg;
            if (!string.IsNullOrEmpty(CustomFields))
            {
                dynamic expando = Util.ConvertToExpando(objVendCredit);
                List<String> lst = CustomFields.Split(",").ToList();
                Dictionary<string, object> tempProps = new Dictionary<string, object>();
                foreach (string str in lst)
                {
                    if (VCRow.Table.Columns.Contains(str) && VCRow[str] != null)
                    {
                        tempProps.Add(str, VCRow[str]);
                    }
                }
                Util.AddProperties(expando, tempProps);
                string strPL = JsonConvert.SerializeObject(expando);
                respMsg = Util.AuditAPIURL.AppendPathSegment("vendorCredits").WithOAuthBearerToken(Util.bearerToken)
                                        .WithHeader("Content-Type", "application/json").PostStringAsync(strPL).Result;
            }
            else
            {
                respMsg = Util.AuditAPIURL.AppendPathSegment("vendorCredits").WithOAuthBearerToken(Util.bearerToken)
                                        .WithHeader("Content-Type", "application/json").PostJsonAsync(objVendCredit).Result;
            }
            strResp = respMsg.Content.ReadAsStringAsync().Result;
            retVal.StatusDesc = strResp;
            if (respMsg.IsSuccessStatusCode)
            {
                IDictionary<string, object> objResp = JsonConvert.DeserializeObject<ExpandoObject>(strResp);
                retVal.Status = StatusCode.Complete;
                retVal.PanopticId = objResp["id"].ToString();
            }
            else
            {
                retVal.Status = StatusCode.Error;
            }
            return retVal;
        }

        private APIReturnValue ProcessCreditClaim(DataRow ClRow, string CustomFields)
        {
            APIReturnValue retVal = new APIReturnValue();
            int VendorId = Convert.ToInt32(ClRow["venId"]);
            string strResp = string.Empty;
            CreditClaimRequest objClaim = Util.CreateObjectFromRow<CreditClaimRequest>(ClRow);
            ClaimInput objInp = Util.CreateObjectFromRow<ClaimInput>(ClRow);
            objClaim.claimInput = objInp;
            objClaim.projectId = ProjectId;
            objClaim.claimInput.appVendorId = Preprocessor.vendorMasterInfoList[VendorId].appVendorId;
            objClaim.uniqueIds = ClRow["CreditIds"].ToString().Split(',').Select(c => Convert.ToInt32(c)).ToList();
            if (!string.IsNullOrEmpty(ClRow["ProblemTypeValue"].ToString()))
            {
                ProblemTypeId objPT = new ProblemTypeId();
                objPT.value = ClRow["ProblemTypeValue"].ToString();
                objPT.id = Preprocessor.claimFieldLookupValues["problemTypeId"][ClRow["ProblemTypeValue"].ToString().ToLower()];
                objClaim.claimInput.problemTypeId = objPT;
            }
            if (!string.IsNullOrEmpty(ClRow["RootCauseValue"].ToString()))
            {
                RootCauseId objRC = new RootCauseId();
                objRC.value = ClRow["RootCauseValue"].ToString();
                objRC.id = Preprocessor.claimFieldLookupValues["rootCauseId"][ClRow["RootCauseValue"].ToString().ToLower()];
                objClaim.claimInput.rootCauseId = objRC;
            }
            objClaim.claimInput.vendorName = Preprocessor.vendorMasterInfoList[VendorId].vendorName;
            objClaim.claimInput.vendorNumber = Preprocessor.vendorMasterInfoList[VendorId].vendorNumber;
            objClaim.claimInput.claimDate = DateTime.Parse(ClRow["ClaimDate"].ToString()).ToUniversalTime().ToString("o");
            dynamic claimExpando = Util.ConvertToExpando(objClaim);
            if (!string.IsNullOrEmpty(CustomFields))
            {
                dynamic expando = Util.ConvertToExpando(objClaim.claimInput);
                List<String> lst = CustomFields.Split(",").ToList();
                Dictionary<string, object> tempProps = new Dictionary<string, object>();
                foreach (string str in lst)
                {
                    if (ClRow.Table.Columns.Contains(str) && ClRow[str] != null)
                    {
                        tempProps.Add(str, ClRow[str]);
                    }
                }
                Util.AddProperties(expando, tempProps);
                Util.AddProperty(claimExpando, "claimInput", expando);
            }
            string strPL = JsonConvert.SerializeObject(claimExpando);
            HttpResponseMessage respMsg = Util.AuditAPIURL.AppendPathSegment("claim/create").WithOAuthBearerToken(Util.bearerToken)
                                    .WithHeader("Content-Type", "application/json").WithHeader("Accept", "application/json")
                                    .PostStringAsync(strPL).Result;

            strResp = respMsg.Content.ReadAsStringAsync().Result;
            retVal.StatusDesc = strResp;
            if (respMsg.IsSuccessStatusCode)
            {
                try
                {
                    retVal.Status = StatusCode.PartiallyComplete;
                    long tempId = GetCreditClaimId(objClaim.uniqueIds.FirstOrDefault());
                    if (tempId != 0)
                    {
                        retVal.PanopticId = tempId.ToString();
                        ChangeClaimStatus((long)tempId, ClRow["StageValue"].ToString(), ClRow["StatusValue"].ToString());
                        retVal.Status = StatusCode.Complete;
                    }
                    else
                    {
                        retVal.StatusDesc = $"Credit Claim Created but Further Processing Failed for Vendor Id: {VendorId}.  Claim Id: {ClRow["RecordId"].ToString()}";
                    }
                }
                catch (AggregateException ae)
                {
                    ae.Handle(e =>
                    {
                        if (e is FlurlHttpException)
                        {
                            string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                            retVal.StatusDesc = es;
                            Logger.log.Error(e, $"Error Occurred while Post Creation Credit Claim Processing for PanopticId: {retVal.PanopticId}. Record Id: {ClRow["RecordId"]}, Response: {es}");
                        }
                        else
                        {
                            Logger.log.Error(e, $"Error Occurred while Post Creation Credit Claim Processing for PanopticId: {retVal.PanopticId}. Record Id: {ClRow["RecordId"]}");
                            retVal.StatusDesc = e.Message;
                        }
                        return true; //continue execution... set to false to break.              
                    });
                }
                catch (Exception e)
                {
                    retVal.StatusDesc = e.Message;
                    Logger.log.Error(e, $"Error Occurred while Post Creation Credit Claim Processing for PanopticId: {retVal.PanopticId}. Record Id: {ClRow["RecordId"]}");
                }
            }
            else
            {
                retVal.Status = StatusCode.Error;
            }
            return retVal;
        }

        private APIReturnValue ProcessAdhocClaim(DataRow ClRow, string CustomFields)
        {
            APIReturnValue retVal = new APIReturnValue();
            int VendorId = Convert.ToInt32(ClRow["venId"]);
            string strResp = string.Empty;
            AdhocClaimRequest objClaim = Util.CreateObjectFromRow<AdhocClaimRequest>(ClRow);
            objClaim.projectId = ProjectId.ToString();
            objClaim.appVendorId = Preprocessor.vendorMasterInfoList[VendorId].appVendorId;
            objClaim.vendorId = VendorId.ToString();
            if (!string.IsNullOrEmpty(ClRow["ProblemTypeValue"].ToString()))
            {
                ProblemTypeId objPT = new ProblemTypeId();
                objPT.value = ClRow["ProblemTypeValue"].ToString();
                objPT.id = Preprocessor.claimFieldLookupValues["problemTypeId"][ClRow["ProblemTypeValue"].ToString().ToLower()];
                objClaim.problemTypeId = objPT;
            }
            if (!string.IsNullOrEmpty(ClRow["RootCauseValue"].ToString()))
            {
                RootCauseId objRC = new RootCauseId();
                objRC.value = ClRow["RootCauseValue"].ToString();
                objRC.id = Preprocessor.claimFieldLookupValues["rootCauseId"][ClRow["RootCauseValue"].ToString().ToLower()];
                objClaim.rootCauseId = objRC;
            }
            objClaim.vendorName = Preprocessor.vendorMasterInfoList[VendorId].vendorName;
            objClaim.vendorNumber = Preprocessor.vendorMasterInfoList[VendorId].vendorNumber;
            dynamic expando = Util.ConvertToExpando(objClaim);
            if (!string.IsNullOrEmpty(CustomFields))
            {
                List<String> lst = CustomFields.Split(",").ToList();
                Dictionary<string, object> tempProps = new Dictionary<string, object>();
                foreach (string str in lst)
                {
                    if (ClRow.Table.Columns.Contains(str) && ClRow[str] != null)
                    {
                        tempProps.Add(str, ClRow[str]);
                    }
                }
                Util.AddProperties(expando, tempProps);
            }
            string strPL = JsonConvert.SerializeObject(expando);
            HttpResponseMessage respMsg = Util.ClaimsAPIURL.AppendPathSegment("claims").WithOAuthBearerToken(Util.bearerToken)
                                    .WithHeader("Content-Type", "application/json").PostStringAsync(strPL).Result;

            strResp = respMsg.Content.ReadAsStringAsync().Result;
            retVal.StatusDesc = strResp;
            if (respMsg.IsSuccessStatusCode)
            {
                retVal.Status = StatusCode.PartiallyComplete;
                IDictionary<string, object> objResp = JsonConvert.DeserializeObject<ExpandoObject>(strResp);
                retVal.PanopticId = objResp["id"].ToString();
                try
                {
                    ChangeClaimStatus(Convert.ToInt64(retVal.PanopticId), ClRow["StageValue"].ToString(), ClRow["StatusValue"].ToString());
                    retVal.Status = StatusCode.Complete;
                }
                catch (AggregateException ae)
                {
                    ae.Handle(e =>
                    {
                        if (e is FlurlHttpException)
                        {
                            string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                            retVal.StatusDesc = es;
                            Logger.log.Error(e, $"Error Occurred while Post Creation Adhoc Claim Processing for PanopticId: {retVal.PanopticId}. Record Id: {ClRow["RecordId"]}, Response: {es}");
                        }
                        else
                        {
                            Logger.log.Error(e, $"Error Occurred while Post Creation Adhoc Claim Processing for PanopticId: {retVal.PanopticId}. Record Id: {ClRow["RecordId"]}");
                            retVal.StatusDesc = e.Message;
                        }
                        return true; //continue execution... set to false to break.              
                    });
                }
                catch (Exception e)
                {
                    retVal.StatusDesc = e.Message;
                    Logger.log.Error(e, $"Error Occurred while Post Creation Adhoc Claim Processing for PanopticId: {retVal.PanopticId}. Record Id: {ClRow["RecordId"]}");
                }
            }
            else
            {
                retVal.Status = StatusCode.Error;
            }
            return retVal;
        }

        private static void ChangeClaimStatus(long ClaimId, string TargetStage, string TargetStatus)
        {
            string currentStage = string.Empty;
            ExpandoObject retMsg = null;
            HttpResponseMessage respMsg = null;
            string nextStatusActId = string.Empty;
            while (currentStage != TargetStage)
            {
                string nextStageActId = string.Empty;
                bool bActDis = true;
                int c = 9;
                int m = 0;
                //check if workflow is ready for the next action
                while (bActDis)
                {
                    retMsg = Util.ClaimsAPIURL.AppendPathSegment($"{ClaimId}/claimWorkflowDetails")
                                .WithOAuthBearerToken(Util.bearerToken).GetJsonAsync().Result;
                    bActDis = (bool)retMsg.FirstOrDefault(x => x.Key == "actionsDisabled").Value;
                    if (bActDis)
                    {
                        System.Threading.Thread.Sleep(500);
                        --c;
                        m++;
                        if (c < 0) throw new Exception($"Panoptic Claim {ClaimId} is in undefined State and Approvals could not be processed. Retry Count: {m}");
                    }
                }
                //Get the ActionId for "approve" action.
                currentStage = ((ExpandoObject)((ExpandoObject)retMsg.FirstOrDefault(x => x.Key == "workflowVariables").Value)
                    .FirstOrDefault(x => x.Key == "stageId").Value).FirstOrDefault(x => x.Key == "value").Value.ToString();
                Logger.log.Info($"Claim {ClaimId} Processed to Stage {currentStage}.");
                string strStageId = ((ExpandoObject)((ExpandoObject)retMsg.FirstOrDefault(x => x.Key == "workflowVariables").Value)
                    .FirstOrDefault(x => x.Key == "stageId").Value).FirstOrDefault(x => x.Key == "id").Value.ToString();
                var acts = ((List<Object>)(retMsg.FirstOrDefault(x => x.Key == "workflowActions").Value));
                foreach (ExpandoObject o in acts)
                {
                    if (o.FirstOrDefault(x => x.Key == "userAction" && x.Value.ToString() == "approve").Key != null)
                    {
                        nextStageActId = o.FirstOrDefault(x => x.Key == "workflowActionId").Value.ToString();
                    }
                    if (o.FirstOrDefault(x => x.Key == "userAction" && x.Value.ToString() == "adjust-status").Key != null)
                    {
                        nextStatusActId = o.FirstOrDefault(x => x.Key == "workflowActionId").Value.ToString();
                    }
                }
                if (currentStage == TargetStage) break;
                if (string.IsNullOrEmpty(nextStageActId)) throw new Exception($"Panoptic Claim {ClaimId} Approvals could not be processed");
                WFActionRequest objAct = new WFActionRequest();
                objAct.approvalNotes = "APPROVED: Approved during Migration";
                objAct.workflowInput.userAction = "approve";
                objAct.workflowInput.userActionId = int.Parse(nextStageActId);

                respMsg = Util.ClaimsAPIURL.AppendPathSegment($"claims({ClaimId})/wf/action").WithOAuthBearerToken(Util.bearerToken)
                                        .WithHeader("Content-Type", "application/json").PostJsonAsync(objAct).Result;
            }
            string disp = string.Empty;
            string dispId = string.Empty;
            string statId = string.Empty;
            var stats = ((List<Object>)(retMsg.FirstOrDefault(x => x.Key == "workflowStatus").Value));
            foreach (IDictionary<string, object> o in stats)
            {
                if (((IDictionary<string, object>)o["statusId"])["value"].ToString() == TargetStatus)
                {
                    statId = ((IDictionary<string, object>)o["statusId"])["id"].ToString();
                    disp = ((IDictionary<string, object>)o["dispositionId"])["value"].ToString();
                    dispId = ((IDictionary<string, object>)o["dispositionId"])["id"].ToString();
                }
            }

            retMsg = Util.ClaimsAPIURL.AppendPathSegment($"claims({ClaimId})")
            .WithOAuthBearerToken(Util.bearerToken).WithHeader("$select", "status").GetJsonAsync().Result;
            if (retMsg.FirstOrDefault(x => x.Key == "status").Value.ToString() != TargetStatus)
            {
                if (string.IsNullOrEmpty(statId)) throw new Exception($"Panoptic Claim {ClaimId} Status Change could not be processed. Status {TargetStatus} is not available in Stage {currentStage}");
                if (string.IsNullOrEmpty(nextStatusActId)) throw new Exception($"Panoptic Claim {ClaimId} Status Change could not be processed. Change Status is not available.");
                WFChangeStatusRequest objCS = new WFChangeStatusRequest();
                objCS.approvalNotes = "Change Status during Migration";
                objCS.workflowInput.userAction = "adjust-status";
                objCS.workflowInput.userActionId = int.Parse(nextStatusActId);
                objCS.claimHeader.dispositionId.id = dispId;
                objCS.claimHeader.statusId.value = TargetStatus;
                objCS.claimHeader.statusId.id = statId;
                objCS.claimHeader.dispositionId.value = disp;
                objCS.claimHeader.dispositionId.id = dispId;
                respMsg = Util.ClaimsAPIURL.AppendPathSegment($"claims({ClaimId})/wf/action").WithOAuthBearerToken(Util.bearerToken)
                                            .WithHeader("Content-Type", "application/json").PostJsonAsync(objCS).Result;
                Logger.log.Info($"Panoptic Claim {ClaimId} Processed to Status {TargetStatus}.");
            }
        }

        private static long GetCreditClaimId(int CreditId)
        {
            long RetId = 0;
            object objMapping = null;
            int c = 9;
            int m = 0;
            //See if Claim is created
            while (objMapping == null || RetId == 0)
            {
                ExpandoObject retMsg = Util.AuditAPIURL.AppendPathSegment("vendorCredits")
                                .WithOAuthBearerToken(Util.bearerToken)
                                .SetQueryParam("$filter", $"id eq {CreditId}")
                                .SetQueryParam("$expand", "claimMapping")
                                .GetJsonAsync().Result;
                objMapping = ((ExpandoObject)((List<object>)retMsg.FirstOrDefault(x => x.Key == "value").Value)[0]).FirstOrDefault(x => x.Key == "claimMapping").Value;
                if (objMapping != null) RetId = Convert.ToInt64(((ExpandoObject)objMapping).FirstOrDefault(x => x.Key == "claimId").Value);
                if (objMapping == null || RetId == 0)
                {
                    Thread.Sleep(500);
                    --c;
                    m++;
                    if (c < 0) throw new Exception($"Panoptic Claim Id could not be obtained for Credit Id: {CreditId}. Retry count: {m}");
                }
            }
            return RetId;
        }

        private static void ChangeDocumentStatus(string DocId, string TargetStatus)
        {
            HttpResponseMessage respStatMsg = Util.AuditAPIURL.AppendPathSegment($"documentmetadatas({DocId})")
                  .WithOAuthBearerToken(Util.bearerToken).WithHeader("Content-Type", "application/json")
                  .PatchJsonAsync(new { documentStatus = TargetStatus, visibilityLevel = "1" }).Result;
            string strStatResp = respStatMsg.Content.ReadAsStringAsync().Result;
            if (!respStatMsg.IsSuccessStatusCode)
            {
                string msg = $"Solicitation Document Status Change Failed. Document Id: {DocId}.";
                Logger.log.Error(msg);
                throw new Exception(msg);
            }
        }

        /*
        public static Dictionary<string, Stream> GetFileStreams(String FilePath)
        {
            Dictionary<string, Stream> files = new Dictionary<string, Stream>();
            HttpResponseMessage respEmailMsg = Util.EmailAPIURL
                    .AppendPathSegment($"email/read")
                    .PostMultipartAsync(mp => mp.AddStringParts(new { target = string.Empty })
                    .AddFile("emailFile", FilePath)).Result;
            Logger.log.Debug($"Read Email finished for {FilePath}");
            if (respEmailMsg.IsSuccessStatusCode)
            {
                IDictionary<string, object> objEmailResp =
                    JsonConvert.DeserializeObject<ExpandoObject>(respEmailMsg.Content.ReadAsStringAsync().Result);
                files.Add(Path.GetFileName(Path.ChangeExtension(FilePath, "html")),
                    new MemoryStream(Encoding.UTF8.GetBytes(objEmailResp["bodyHtml"].ToString() ?? "")));
                int intC = Convert.ToInt32(objEmailResp["attachmentCount"]);
                if (intC > 0)
                {
                    List<Object> colAtt = (List<Object>)objEmailResp["attachments"];
                    foreach (ExpandoObject obj in colAtt)
                    {
                        string attN = obj.FirstOrDefault(x => x.Key == "name").Value.ToString();
                        HttpResponseMessage respAttMsg = Util.EmailAPIURL
                            .AppendPathSegment($"attachment/download")
                            .PostMultipartAsync(mp => mp.AddStringParts(new { target = attN })
                            .AddFile("emailFile", FilePath)).Result;
                        Logger.log.Debug($"Download Attachment finished for {attN} in {FilePath}");
                        if (respAttMsg.IsSuccessStatusCode)
                        {
                            files.Add(attN, respAttMsg.Content.ReadAsStreamAsync().Result);
                        }
                        else
                        {
                            throw new Exception($"Error Occurred while Downloading Attachment {attN}. Resp: {respEmailMsg.Content.ReadAsStringAsync().Result}");
                        }
                    }
                }
            }
            else
            {
                throw new Exception($"Error Occurred while Converting MSG to HTML. Resp: {respEmailMsg.Content.ReadAsStringAsync().Result}");
            }
            return files;
        }
        */

        public static Dictionary<string, Stream> ExtractMSG(String FilePath)
        {
            Dictionary<string, Stream> files = new Dictionary<string, Stream>();
            Stream memStream = null;
            string attName = string.Empty;
            Message msg = new Message(FilePath);
            files.Add(Path.GetFileName(Path.ChangeExtension(FilePath, "html")), new MemoryStream(msg.BodyHtml));
            foreach (Attachment att in msg.Attachments)
            {
                if (att.EmbeddedMessage != null)
                {
                    Message newMsg = new Message();
                    newMsg = (Message)att.EmbeddedMessage;
                    newMsg.IsEmbedded = false; // set the flag to false
                    attName = att.DisplayName ?? "EmbeddedMessage.msg";
                    if (!Path.HasExtension(attName)) attName = Path.ChangeExtension(attName, att.Extension ?? ".msg");
                    memStream = newMsg.GetStream();
                }
                else if ((!att.IsHidden && att.Data != null) && (string.IsNullOrEmpty(att.ContentId) || !(msg.Body.Contains("cid:" + att.ContentId))))
                {
                    attName = att.DisplayName ?? "EmailAttachment";
                    if (!Path.HasExtension(attName) && !string.IsNullOrEmpty(att.Extension)) attName = Path.ChangeExtension(attName, att.Extension);
                    memStream = new MemoryStream(att.Data);
                }
                else
                {
                    continue;
                }
                string name = attName;
                int i = 0;
                while (files.ContainsKey(name)) name = $"{Path.GetFileNameWithoutExtension(attName)}({++i}){Path.GetExtension(name)}";
                files.Add(name, memStream);
            }
            return files;
        }
    }
}
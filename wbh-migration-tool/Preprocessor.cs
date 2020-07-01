using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Data;
using Microsoft.Extensions.Configuration;
using Flurl;
using Flurl.Http;

namespace PRGX.Panoptic_Migration_Core
{
    static class Preprocessor
    {
        public static Dictionary<int, VendorMaster> vendorMasterInfoList = new Dictionary<int, VendorMaster>();
        public static Dictionary<string, int> documentTypeList = new Dictionary<string, int>();
        public static Dictionary<string, int> projectVendorStatusList = new Dictionary<string, int>();
        public static Dictionary<string, int> vendorStatementActionList = new Dictionary<string, int>();
        public static Dictionary<string, Dictionary<string, string>> auditFieldLookupValues = new Dictionary<string, Dictionary<string, string>>();
        public static Dictionary<string, Dictionary<string, string>> claimFieldLookupValues = new Dictionary<string, Dictionary<string, string>>();

        public static List<object> GetDocumentTypes()
        {
            List<object> ret = new List<object>();
            try
            {
                ExpandoObject resp = Util.AuditAPIURL
                .AppendPathSegment("documentTypes")
                .WithOAuthBearerToken(Util.bearerToken)
                .SetQueryParam("$select", "documentType,documentTypeId")
                .SetQueryParam("$top", 5000)
                .GetJsonAsync().Result;

                ret = (List<object>)resp.FirstOrDefault(x => x.Key == "value").Value;
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"Error Occurred while Getting the DocumentType List, Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, "Error Occurred while Getting the DocumentType List");
                    }
                    return false; //Break the execution... set to true to continue.              
                });
            }
            catch (Exception e)
            {
                Logger.log.Error(e, "Error Occurred while Getting the DocumentType List");
                throw e;
            }
            return ret;
        }

        public static void PopulateVendorInfo()
        {
            DataSet venDS = Util.GetVendorMasterData();
            DataTable venTbl = venDS.Tables[0].DefaultView.ToTable(true, "VendorId", "AppVendorId", "VendorName", "VendorNumber", "VendorNumber2")
                    .AsEnumerable().AsEnumerable().OrderBy(r => r.Field<int>("VendorId")).CopyToDataTable();
            foreach (DataRow row in venTbl.Rows)
            {
                int VendorId = Convert.ToInt32(row["VendorId"].ToString());
                try
                {
                    VendorMaster objVend = new VendorMaster();
                    objVend.appVendorId = Convert.ToString(row["AppVendorId"]);
                    objVend.vendorName = Convert.ToString(row["VendorName"]);
                    objVend.vendorNumber = Convert.ToString(row["VendorNumber"]);
                    objVend.vendorNumber2 = Convert.ToString(row["VendorNumber2"]);
                    Dictionary<string, string> tempD = new Dictionary<string, string>();
                    var TempRS = venDS.Tables[0].AsEnumerable().Where(x => x.Field<int>("VendorId") == VendorId);
                    foreach (DataRow accountRow in TempRS)
                    {
                        string ccode = Convert.ToString(accountRow["vendorCompanyCode"]);
                        if (!string.IsNullOrEmpty(ccode))
                        {
                            ccode = ccode.ToLower();
                            string ccid = Convert.ToString(accountRow["appVendorAccountId"]);
                            if (!objVend.vendorAccounts.ContainsKey(ccode)) objVend.vendorAccounts.Add(ccode, ccid);
                        }
                    }
                    vendorMasterInfoList.Add(VendorId, objVend);
                }
                catch (Exception e)
                {
                    Logger.log.Error(e, $"Error Occurred while Getting the Vendor Account Info for Vendor Id: {VendorId}");
                }
            }
        }

        public static int GetProjectVendorId(string AppVendorId, int ProjectId)
        {
            int ret = 0;
            try
            {
                ExpandoObject resp = Util.AuditAPIURL
                .AppendPathSegment("vendorselectiondetails")
                .WithOAuthBearerToken(Util.bearerToken)
                .SetQueryParam("projectId", ProjectId)
                .SetQueryParam("$filter", $"appVendorId eq {AppVendorId}")
                .SetQueryParam("$select", "projectVendorId")
                .SetQueryParam("$top", 500)
                .GetJsonAsync().Result;

                var strPVId = ((ExpandoObject)((List<object>)resp.FirstOrDefault(x => x.Key == "value").Value)[0])
                    .FirstOrDefault(x => x.Key == "projectVendorId").Value;
                if (strPVId == null || !(int.TryParse(strPVId.ToString(), out ret)))
                    throw new Exception($"Project Vendor for AppVendor {AppVendorId} not found in Vendor Selection");
                return ret;
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"Error Occurred while Getting the Project Vendor Id, Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, "Error Occurred while Getting the Project Vendor Id");
                    }
                    return false; //Break the execution... set to true to continue.              
                });
            }
            catch (Exception e)
            {
                Logger.log.Error(e, "Error Occurred while Getting the Project Vendor Id");
                throw e;
            }
            return ret;
        }

        public static List<object> GetAuditLookupValues(string ObjectName)
        {
            List<object> ret = new List<object>();
            try
            {
                ExpandoObject resp = Util.AuditAPIURL
                .AppendPathSegment("dataObjects")
                .WithOAuthBearerToken(Util.bearerToken)
                .SetQueryParam("$filter", $"name eq '{ObjectName}'")
                .GetJsonAsync().Result;
                string resId = ((ExpandoObject)((List<object>)resp.FirstOrDefault(x => x.Key == "value").Value)[0])
                    .FirstOrDefault(x => x.Key == "id").Value.ToString();
                resp = Util.DataAPIURL
                .AppendPathSegment($"dataObjects({resId})/dataObjectFields")
                .WithHeader("Accept", "application/json")
                .WithOAuthBearerToken(Util.bearerToken)
                .SetQueryParam("$expand", "dataObjectFieldValues")
                .SetQueryParam("$filter", $"name eq '{"creditStatus"}' or name eq '{"problemType"}' or name eq '{"problemSubType"}'")
                .SetQueryParam("$top", 200)
                .GetJsonAsync().Result;
                ret = (List<object>)resp.FirstOrDefault(x => x.Key == "value").Value;
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"Error Occurred while Getting the Lookup Values, Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, "Error Occurred while Getting the Lookup Values");
                    }
                    return false; //Break the execution... set to true to continue.              
                });
            }
            catch (Exception e)
            {
                Logger.log.Error(e, "Error Occurred while Getting the Lookup Values");
                throw e;
            }
            return ret;
        }

        public static List<object> GetClaimLookupValues(string ObjectName)
        {
            List<object> ret = new List<object>();
            try
            {
                ExpandoObject resp = Util.DataAPIURL
                .AppendPathSegment("dataObjects")
                .WithOAuthBearerToken(Util.bearerToken)
                .WithHeader("Accept", "application/json")
                .SetQueryParam("$filter", $"name eq '{ObjectName}'")
                .GetJsonAsync().Result;
                string resId = ((ExpandoObject)((List<object>)resp.FirstOrDefault(x => x.Key == "value").Value)[0])
                    .FirstOrDefault(x => x.Key == "id").Value.ToString();
                /*
                call API http://data-dictionary.qa.wbh.ocp.prgx.com/aG5k/dataObjects(db9499cc-e872-49d0-beb6-30f99beb63fc)/dataObjectFields?
                $top=1000&$expand=dataObjectFieldValues,dataObjectFieldType&$filter=dataObjectFieldType/inactiveDate eq null
                This gives all the fields, their values and field types. 
                */
                resp = Util.DataAPIURL
                .AppendPathSegment($"dataObjects({resId})/dataObjectFields")
                .WithHeader("Accept", "application/json")
                .WithOAuthBearerToken(Util.bearerToken)
                .SetQueryParam("$expand", "dataObjectFieldValues")
                .SetQueryParam("$filter", $"name eq '{"problemTypeId"}' or name eq '{"rootCauseId"}' or name eq '{"statusId"}' or name eq '{"stageId"}'")
                .SetQueryParam("$top", 1000)
                .GetJsonAsync().Result;
                ret = (List<object>)resp.FirstOrDefault(x => x.Key == "value").Value;
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"Error Occurred while Getting the Lookup Values, Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, "Error Occurred while Getting the Lookup Values");
                    }
                    return false; //Break the execution... set to true to continue.              
                });
            }
            catch (Exception e)
            {
                Logger.log.Error(e, "Error Occurred while Getting the Lookup Values");
                throw e;
            }
            return ret;
        }

        public static List<object> GetVendorStatementActions()
        {
            List<object> ret = new List<object>();
            try
            {
                ExpandoObject resp = Util.AuditAPIURL
                .AppendPathSegment("vendorstatementactions")
                .WithOAuthBearerToken(Util.bearerToken)
                .SetQueryParam("$select", "actionTaken,vendorStatementActionId")
                .GetJsonAsync().Result;

                ret = (List<object>)resp.FirstOrDefault(x => x.Key == "value").Value;
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"Error Occurred while Getting the StatementAction List, Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, "Error Occurred while Getting the StatementAction List");
                    }
                    return false; //Break the execution... set to true to continue.              
                });
            }
            catch (Exception e)
            {
                Logger.log.Error(e, "Error Occurred while Getting the StatementAction List");
                throw e;
            }
            return ret;
        }

        public static List<object> GetProjectVendorStatuses()
        {
            List<object> ret = new List<object>();
            try
            {
                ExpandoObject resp = Util.AuditAPIURL
                .AppendPathSegment("projectvendorstatuses")
                .WithOAuthBearerToken(Util.bearerToken)
                .SetQueryParam("$select", "vendorStatus,projectVendorStatusId")
                .GetJsonAsync().Result;

                ret = (List<object>)resp.FirstOrDefault(x => x.Key == "value").Value;
            }
            catch (AggregateException ae)
            {
                ae.Handle(e =>
                {
                    if (e is FlurlHttpException)
                    {
                        string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                        Logger.log.Error(e, $"Error Occurred while Getting the ProjectVendorStatus List, Response: {es}");
                    }
                    else
                    {
                        Logger.log.Error(e, "Error Occurred while Getting the ProjectVendorStatus List");
                    }
                    return false; //Break the execution... set to true to continue.              
                });
            }
            catch (Exception e)
            {
                Logger.log.Error(e, "Error Occurred while Getting the ProjectVendorStatus List");
                throw e;
            }
            return ret;
        }

        public static bool Preprocess()
        {
            PopulateVendorInfo();

            foreach (ExpandoObject o in GetAuditLookupValues("Vendor"))
            {
                string nm = o.FirstOrDefault(x => x.Key == "name").Value.ToString();
                List<Object> f = (List<Object>)o.FirstOrDefault(x => x.Key == "dataObjectFieldValues").Value;
                Dictionary<string, string> tempD = new Dictionary<string, string>();
                foreach (ExpandoObject obj in f)
                {
                    string k = obj.FirstOrDefault(x => x.Key == "value").Value.ToString().ToLower();
                    string v = obj.FirstOrDefault(x => x.Key == "id").Value.ToString();
                    if (!tempD.ContainsKey(k)) tempD.Add(k, v);
                }
                if (!auditFieldLookupValues.ContainsKey(nm)) auditFieldLookupValues.Add(nm, tempD);
            }

            foreach (ExpandoObject o in GetClaimLookupValues("Claim"))
            {
                string nm = o.FirstOrDefault(x => x.Key == "name").Value.ToString();
                List<Object> f = (List<Object>)o.FirstOrDefault(x => x.Key == "dataObjectFieldValues").Value;
                Dictionary<string, string> tempD = new Dictionary<string, string>();
                foreach (ExpandoObject obj in f)
                {
                    string k = obj.FirstOrDefault(x => x.Key == "value").Value.ToString().ToLower();
                    string v = obj.FirstOrDefault(x => x.Key == "id").Value.ToString();
                    if (!tempD.ContainsKey(k)) tempD.Add(k, v);
                }
                if (!claimFieldLookupValues.ContainsKey(nm)) claimFieldLookupValues.Add(nm, tempD);
            }

            foreach (ExpandoObject o in Preprocessor.GetDocumentTypes())
            {
                int v = int.Parse((o.FirstOrDefault(x => x.Key == "documentTypeId").Value.ToString()));
                string av = (o.FirstOrDefault(x => x.Key == "documentType").Value.ToString().ToLower());
                if (!documentTypeList.ContainsKey(av)) documentTypeList.Add(av, v);
            }

            foreach (ExpandoObject o in GetVendorStatementActions())
            {
                int v = int.Parse((o.FirstOrDefault(x => x.Key == "vendorStatementActionId").Value.ToString()));
                string av = (o.FirstOrDefault(x => x.Key == "actionTaken").Value.ToString());
                if (!vendorStatementActionList.ContainsKey(av)) vendorStatementActionList.Add(av, v);
            }

            foreach (ExpandoObject o in GetProjectVendorStatuses())
            {
                int v = int.Parse((o.FirstOrDefault(x => x.Key == "projectVendorStatusId").Value.ToString()));
                string av = (o.FirstOrDefault(x => x.Key == "vendorStatus").Value.ToString().ToLower());
                if (!projectVendorStatusList.ContainsKey(av)) projectVendorStatusList.Add(av, v);
            }

            return true;
        }
    }
}

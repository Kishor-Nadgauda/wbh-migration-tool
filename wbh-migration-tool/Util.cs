using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Dynamic;
using System.Net.Http;
using System.Reflection;
using System.Collections;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Flurl.Http;

namespace PRGX.Panoptic_Migration_Core
{
    static class Util
    {
        public static IConfiguration Configuration;
        public static string SQLConnString = string.Empty;
        public static int BatchId = Convert.ToInt32((DateTime.UtcNow - new DateTime(2020, 06, 24).ToUniversalTime()).TotalSeconds);
        public static string AuditAPIURL;
        public static string DataAPIURL;
        public static string ClaimsAPIURL;
        public static string AuthenticationAPIURL;
        public static string ClientId;
        public static string ClientSecret;
        public static string VendorMasterSPName;
        public static string MigrationConfigSPName;
        public static string bearerToken = string.Empty;
        public static string refreshToken = string.Empty;

        static bool ConfigError = false;

        public static void Initialize()
        {
            Configuration = new ConfigurationBuilder()
            .AddJsonFile("AppSettings.json", optional: false, reloadOnChange: true).Build();
            AuditAPIURL = Configuration["AuditAPIHostURL"];
            if (string.IsNullOrEmpty(AuditAPIURL))
            {
                Logger.log.Error("APIHostURL is not configured in Config File.");
                ConfigError = true;
            }
            bearerToken = Configuration["APIToken"];
            if (string.IsNullOrEmpty(bearerToken))
            {
                Logger.log.Error("APIToken is not configured in Config File.");
                ConfigError = true;
            }
            DataAPIURL = Configuration["DataDictionaryAPIHostURL"];
            if (string.IsNullOrEmpty(DataAPIURL))
            {
                Logger.log.Error("DataDictionaryAPIHostURL is not configured in Config File.");
                ConfigError = true;
            }
            ClaimsAPIURL = Configuration["ClaimsAPIHostURL"];
            if (string.IsNullOrEmpty(ClaimsAPIURL))
            {
                Logger.log.Error("ClaimsAPIHostURL is not configured in Config File.");
                ConfigError = true;
            }
            AuthenticationAPIURL = Configuration["AuthenticationAPIURL"];
            if (string.IsNullOrEmpty(AuthenticationAPIURL))
            {
                Logger.log.Error("AuthenticationAPIURL is not configured in Config File.");
                ConfigError = true;
            }
            ClientId = Configuration["client_id"];
            if (string.IsNullOrEmpty(ClientId))
            {
                Logger.log.Error("client_id is not configured in Config File.");
                ConfigError = true;
            }
            ClientSecret = Configuration["client_secret"];
            if (string.IsNullOrEmpty(ClientSecret))
            {
                Logger.log.Error("client_secret is not configured in Config File.");
                ConfigError = true;
            }
            SQLConnString = Configuration["SQLConnectionString"];
            if (string.IsNullOrEmpty(SQLConnString))
            {
                Logger.log.Error("SQLConnString is not configured in Config File.");
                ConfigError = true;
            }
            VendorMasterSPName = Configuration["VendorMasterSPName"];
            if (string.IsNullOrEmpty(VendorMasterSPName))
            {
                Logger.log.Error("VendorMasterSPName is not configured in Config File.");
                ConfigError = true;
            }
            MigrationConfigSPName = Configuration["MigrationConfigSPName"];
            if (string.IsNullOrEmpty(MigrationConfigSPName))
            {
                Logger.log.Error("MigrationConfigSPName is not configured in Config File.");
                ConfigError = true;
            }
            if (ConfigError) throw new Exception("One or more configuration errors found. Please look at the log for details.");
        }

        // function that creates an object from the given data row
        public static T CreateObjectFromRow<T>(DataRow row) where T : new()
        {
            // create a new object
            T item = new T();

            // go through each column
            foreach (DataColumn c in row.Table.Columns)
            {
                // find the property for the column
                PropertyInfo p = item.GetType().GetProperty(c.ColumnName);

                // if exists, set the value
                if (p != null && row[c] != DBNull.Value)
                {
                    p.SetValue(item, row[c], null);
                }
            }
            return item;
        }

        //Helper function added to create dynamic request object if needed
        public static ExpandoObject ConvertToExpando(object obj)
        {
            //Get Properties Using Reflections
            BindingFlags flags = BindingFlags.Public | BindingFlags.Instance;
            PropertyInfo[] properties = obj.GetType().GetProperties(flags);

            //Add Them to a new Expando
            ExpandoObject expando = new ExpandoObject();
            var expandoDict = expando as IDictionary<string, object>;
            foreach (PropertyInfo property in properties)
            {
                if (expandoDict.ContainsKey(property.Name))
                    expandoDict[property.Name] = property;
                else
                    expandoDict.Add(property.Name, property.GetValue(obj));
            }
            return expando;
        }

        public static void AddProperty(ExpandoObject expando, string propertyName, object propertyValue)
        {
            // ExpandoObject supports IDictionary so we can extend it like this
            var expandoDict = expando as IDictionary<string, object>;
            if (expandoDict.ContainsKey(propertyName))
                expandoDict[propertyName] = propertyValue;
            else
                expandoDict.Add(propertyName, propertyValue);
        }

        public static void AddProperties(ExpandoObject Expando, Dictionary<string, object> PropSet)
        {
            // ExpandoObject supports IDictionary so we can extend it like this
            var expandoDict = Expando as IDictionary<string, object>;
            foreach (string str in PropSet.Keys)
            {
                if (expandoDict.ContainsKey(str))
                    expandoDict[str] = PropSet[str];
                else
                    expandoDict.Add(str, PropSet[str]);
            }
        }

        public static DataSet GetMigrationConfigurations()
        {
            try
            {
                Logger.log.Debug("Started Reading Migration Configurations");
                using (SqlConnection conn = new SqlConnection(SQLConnString))
                {
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = Util.MigrationConfigSPName;
                            cmd.CommandType = CommandType.StoredProcedure;
                            da.SelectCommand = cmd;
                            DataSet ds = new DataSet();
                            conn.Open();
                            da.Fill(ds);
                            Logger.log.Debug($"Finished Reading Migration Configurations");
                            return ds;
                        }
                    }
                }
            }
            catch (SqlException se)
            {
                Logger.log.Error(se, $"SQL Error Occurred while Reading Migration Configurations. Error Code: {se.ErrorCode}");
                throw se;
            }
            catch (Exception e)
            {
                Logger.log.Error(e, $"Error Occurred while Reading Migration Configurations");
                throw e;
            }
        }

        public static DataSet GetVendorMasterData()
        {
            try
            {
                Logger.log.Debug("Started Reading Vendor Master Data");
                using (SqlConnection conn = new SqlConnection(SQLConnString))
                {
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = Util.VendorMasterSPName;
                            cmd.CommandType = CommandType.StoredProcedure;
                            da.SelectCommand = cmd;
                            DataSet ds = new DataSet();
                            conn.Open();
                            da.Fill(ds);
                            Logger.log.Debug($"Finished Reading Vendor Master Data");
                            return ds;
                        }
                    }
                }
            }
            catch (SqlException se)
            {
                Logger.log.Error(se, $"SQL Error Occurred while Reading Vendor Master Data. Error Code: {se.ErrorCode}");
                throw se;
            }
            catch (Exception e)
            {
                Logger.log.Error(e, $"Error Occurred while Reading Vendor Master Data");
                throw e;
            }
        }

        public static DataSet GetMigrationData(string EntityName, string MigrationClientCode, int ProjectId, string SPName, int BatchSize)
        {
            try
            {
                Logger.log.Debug("Started Reading Migration Data for Entity: " + EntityName);
                using (SqlConnection conn = new SqlConnection(SQLConnString))
                {
                    using (SqlDataAdapter da = new SqlDataAdapter())
                    {
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = SPName;
                            cmd.CommandType = CommandType.StoredProcedure;
                            da.SelectCommand = cmd;
                            SqlParameter param = new SqlParameter();
                            param.ParameterName = "@Batch_Size";
                            param.Value = BatchSize;
                            cmd.Parameters.Add(param);
                            param = new SqlParameter();
                            param.ParameterName = "MigrationClientCode";
                            param.Value = MigrationClientCode;
                            cmd.Parameters.Add(param);
                            param = new SqlParameter();
                            param.ParameterName = "ProjectId";
                            param.Value = ProjectId;
                            cmd.Parameters.Add(param);
                            DataSet ds = new DataSet();
                            conn.Open();
                            da.Fill(ds);
                            Logger.log.Debug($"Finished Reading Migration Data for Entity: {EntityName}");
                            return ds;
                        }
                    }
                }
            }
            catch (SqlException se)
            {
                Logger.log.Error(se, $"SQL Error Occurred while Reading Migration Data for Entity: {EntityName}. Error Code: {se.ErrorCode}");
                throw se;
            }
            catch (Exception e)
            {
                Logger.log.Error(e, $"Error Occurred while Reading Migration Data for Entity: {EntityName}");
                throw e;
            }
        }

        public static int AddMigrationStatusRecord(int VenId, string ParentId, string RecordId, string EntityName, string MigrationClientCode, int ProjectId, string Status, string StatusDesc)
        {
            int ret = 0;
            try
            {
                Logger.log.Debug($"Started Adding Migration Status to {Status}. Record Id: {RecordId}. Entity Name: {EntityName}");
                SqlConnection conn = new SqlConnection(SQLConnString);
                SqlDataAdapter da = new SqlDataAdapter();
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "sp_AddMigrationStatusRecord";
                cmd.CommandType = CommandType.StoredProcedure;
                da.SelectCommand = cmd;
                SqlParameter param = new SqlParameter();
                param.ParameterName = "@BatchId";
                param.Value = BatchId;
                cmd.Parameters.Add(param);
                param = new SqlParameter();
                param.ParameterName = "@VenId";
                param.Value = VenId;
                cmd.Parameters.Add(param);
                if (!string.IsNullOrEmpty(ParentId))
                {
                    param = new SqlParameter();
                    param.ParameterName = "@ParentId";
                    param.Value = ParentId;
                    cmd.Parameters.Add(param);
                }
                param = new SqlParameter();
                param.ParameterName = "@RecordId";
                param.Value = RecordId;
                cmd.Parameters.Add(param);
                param = new SqlParameter();
                param.ParameterName = "@EntityName";
                //param.SqlDbType = SqlDbType.VarChar;
                //param.Size = 20;
                param.Value = EntityName;
                cmd.Parameters.Add(param);
                param = new SqlParameter();
                param.ParameterName = "@Status";
                //param.SqlDbType = SqlDbType.VarChar;
                //param.Size = 20;
                param.Value = Status;
                cmd.Parameters.Add(param);
                param = new SqlParameter();
                param.ParameterName = "@MigrationClientCode";
                param.Value = MigrationClientCode;
                cmd.Parameters.Add(param);
                param = new SqlParameter();
                param.ParameterName = "@ProjectId";
                param.Value = ProjectId;
                cmd.Parameters.Add(param);
                param = new SqlParameter();
                param.ParameterName = "@StatusDesc";
                //param.SqlDbType = SqlDbType.VarChar;
                //param.Size = 240;
                param.Value = StatusDesc;
                cmd.Parameters.Add(param);
                param = new SqlParameter();
                param.SqlDbType = SqlDbType.Int;
                param.ParameterName = "@Id";
                param.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(param);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (cmd.Parameters["@Id"].Value != null) ret = Convert.ToInt32(cmd.Parameters["@Id"].Value);
                Logger.log.Debug($"Finished Adding Migration Status for Record Id: {RecordId}. Entity Name: {EntityName}");
            }
            catch (SqlException se)
            {
                Logger.log.Error(se, $"SQL Error Occurred while Adding Migration Status for Record Id: {RecordId}. Entity Name: {EntityName}. Error Code: {se.ErrorCode}");
            }
            catch (Exception e)
            {
                Logger.log.Error(e, $"Error Occurred while Adding Migration Status for Record Id: {RecordId}. Entity Name: {EntityName}");
            }
            return ret;
        }

        public static void UpdateMigrationStatusRecord(int Id, string PanopticId, string Status, string StatusDesc)
        {
            try
            {
                Logger.log.Debug($"Started Updating Migration Status. Id: {Id}");
                SqlConnection conn = new SqlConnection(SQLConnString);
                SqlDataAdapter da = new SqlDataAdapter();
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "sp_UpdateMigrationStatusRecord";
                cmd.CommandType = CommandType.StoredProcedure;
                da.SelectCommand = cmd;
                SqlParameter param = new SqlParameter();
                param.ParameterName = "@Id";
                param.Value = Id;
                cmd.Parameters.Add(param);
                if (!string.IsNullOrEmpty(PanopticId))
                {
                    param = new SqlParameter();
                    param.ParameterName = "@PanopticId";
                    param.Value = PanopticId;
                    cmd.Parameters.Add(param);
                }
                param = new SqlParameter();
                param.ParameterName = "@Status";
                param.Value = Status;
                cmd.Parameters.Add(param);
                if (!string.IsNullOrEmpty(StatusDesc))
                {
                    param = new SqlParameter();
                    param.ParameterName = "@StatusDesc";
                    param.Value = StatusDesc;
                    cmd.Parameters.Add(param);
                }
                DataSet ds = new DataSet();
                da.Fill(ds);
                Logger.log.Debug($"Finished Updating Migration Status. Id: {Id}");
            }
            catch (SqlException se)
            {
                Logger.log.Error(se, $"SQL Error Occurred while Updating Migration Status for Record Id: {Id}. Error Code: {se.ErrorCode}");
            }
            catch (Exception e)
            {
                Logger.log.Error(e, $"Error Occurred while Updating Migration Status for Record Id: {Id}");
            }
        }

        public static void SetBearerToken()
        {
            HttpResponseMessage respMsg = null;
            do
            {
                Console.WriteLine(Environment.NewLine);
                Console.Write("Please Enter the Username: ");
                string strUser = Console.ReadLine();
                Console.Write("Please Enter the Password: ");
                string strPass = string.Empty;
                ConsoleKeyInfo keyInfo;
                do
                {
                    keyInfo = Console.ReadKey(true);
                    // Skip if Backspace or Enter is Pressed
                    if (keyInfo.Key != ConsoleKey.Backspace && keyInfo.Key != ConsoleKey.Enter)
                    {
                        strPass += keyInfo.KeyChar;
                        Console.Write("*");
                    }
                    else
                    {
                        if (keyInfo.Key == ConsoleKey.Backspace && strPass.Length > 0)
                        {
                            // Remove last charcter if Backspace is Pressed
                            strPass = strPass.Substring(0, (strPass.Length - 1));
                            Console.Write("b b");
                        }
                    }
                }
                // Stops Getting Password Once Enter is Pressed
                while (keyInfo.Key != ConsoleKey.Enter);
                try
                {
                    respMsg = Util.AuthenticationAPIURL.WithHeader("Content-Type", "application/x-www-form-urlencoded")
                        .PostUrlEncodedAsync(new
                        {
                            client_id = Util.ClientId,
                            client_secret = ClientSecret,
                            grant_type = "password",
                            username = strUser,
                            password = strPass
                        }).Result;
                    if (respMsg.IsSuccessStatusCode)
                    {
                        IDictionary<string, object> objResp = JsonConvert.DeserializeObject<ExpandoObject>(respMsg.Content.ReadAsStringAsync().Result);
                        bearerToken = objResp["access_token"].ToString();
                        refreshToken = objResp["refresh_token"].ToString();
                    }
                    else
                    {
                        Logger.log.Error($"Authentication failed for User {strUser}. Please try again");
                    }
                }
                catch (AggregateException ae)
                {
                    Console.WriteLine(Environment.NewLine);
                    ae.Handle(e =>
                    {
                        if (e is FlurlHttpException)
                        {
                            string es = ((FlurlHttpException)e).GetResponseStringAsync().Result;
                            Logger.log.Error(e, $"Error Occurred during Authentication. Response: {es}");
                        }
                        else
                        {
                            Logger.log.Error(e, $"Error Occurred during Authentication. Response: {e.Message}");
                        }
                        return false;
                    });
                }
            }
            while (respMsg == null || !respMsg.IsSuccessStatusCode);
        }

        /*
                public static void ReadMigrationClientCode()
                {
                    do
                    {
                        Console.Write("Enter the Client Code for Migration: ");
                        MigrationClientCode = Console.ReadLine();
                        Console.Write($"You entered Client Code: {MigrationClientCode}. Please type 'Y' to confirm... ");
                    }
                    while (Console.ReadKey().KeyChar.ToString().ToLower() != "y");
                }

                public static void ReadProjectId()
                {
                    do
                    {
                        Console.Write(Environment.NewLine);
                        Console.Write("Enter the ProjectId for Migration: ");
                        while (true)
                        {
                            if (int.TryParse(Console.ReadLine(), out ProjectId)) break;
                            Console.Write("Please enter only Integer value: ");
                        }
                        Console.Write($"You entered ProjectId: {ProjectId}. Please type 'Y' to confirm... ");
                    }
                    while (Console.ReadKey().KeyChar.ToString().ToLower() != "y");
                }
                 */
    }

}

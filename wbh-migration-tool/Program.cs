using System;
using System.Collections.Generic;
using System.IO;
using Flurl.Http;

namespace PRGX.Panoptic_Migration_Core
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //ToDo: done to workaround local service certificate issue.
                //FlurlHttp.Configure(settings => settings.HttpClientFactory = new CustomHttpClientFactory());
                FlurlHttp.Configure(settings => settings.HttpClientFactory = new PollyHttpClientFactory());

                Util.Initialize();
                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("-----------------------------------             Panoptic Data Migration             ----------------------------------- ");
                Console.WriteLine(Environment.NewLine);
                //Util.SetBearerToken();
                //Util.ReadMigrationClientCode();
                //Util.ReadProjectId();
                Logger.log.Info($"Started Processing Batch {Util.BatchId}");
                (new DataProcessor()).ProcessData();
            }
            catch (Exception e)
            {
                Logger.log.Error(e, $"Error Occurred while Processing Batch {Util.BatchId}");
            }
            finally
            {
                Logger.log.Info($"Finished Processing Batch {Util.BatchId}");
                Console.Write("Press Enter to quit...");
                Console.ReadLine();
            }
        }


    }
}

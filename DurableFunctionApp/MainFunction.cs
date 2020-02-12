using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json.Linq;

namespace DurableFunctionApp
{
    public static class MainFunction
    {
        [FunctionName("MainFunction")]
        public static async Task<List<string>> RunOrchestrator(
            [OrchestrationTrigger] DurableOrchestrationContext context)
        {
            var outputs = new List<string>();

            //access SPO
            outputs.Add(await context.CallActivityAsync<string>("MainFunction_AccessSPO", "access spo"));

            for (int i=0;i<10;i++)
            {
                string msg = $"Task {i + 1} is processing";
                outputs.Add(await context.CallActivityAsync<string>("MainFunction_Process", msg));
                context.SetCustomStatus(new 
                {
                    Process=(i+1)*100/10,
                    Message=msg
                });
            }
            // Replace "hello" with the name of your Durable Activity Function.
            //outputs.Add(await context.CallActivityAsync<string>("MainFunction_Hello", "Tokyo"));
            //outputs.Add(await context.CallActivityAsync<string>("MainFunction_Hello", "Seattle"));
            //outputs.Add(await context.CallActivityAsync<string>("MainFunction_Hello", "London"));

            // returns ["Hello Tokyo!", "Hello Seattle!", "Hello London!"]
            return outputs;
        }

        [FunctionName("MainFunction_Process")]
        public static string SayHello([ActivityTrigger] string name, ILogger log)
        {
            log.LogInformation($"Saying hello to {name}.");
            System.Threading.Thread.Sleep(3000);
            return $"Hello {name}!";
        }

        [FunctionName("MainFunction_AccessSPO")]
        public static async Task<string> AccessSPOAsync([ActivityTrigger] string name, ILogger log)
        {
            HttpClient client = new HttpClient();
            string appId = Environment.GetEnvironmentVariable("AppId");
            string appSecret = Environment.GetEnvironmentVariable("AppSecret");
            string tenantId = Environment.GetEnvironmentVariable("TenantId");
            string tenantName = Environment.GetEnvironmentVariable("TenantName");
            string grant_type = "client_credentials";

            string clientId = string.Format("{0}@{1}",appId, tenantId);
            string clientSecret = appSecret;
            string resource = String.Format("00000003-0000-0ff1-ce00-000000000000/{0}@{1}",tenantName,tenantId);
            string url = String.Format("https://accounts.accesscontrol.windows.net/{0}/tokens/OAuth/2",tenantId);

            //get acs token
            string body = String.Format("grant_type={0}", grant_type);
            body += String.Format("&client_id={0}", WebUtility.UrlEncode(clientId));
            body += String.Format("&client_secret={0}", WebUtility.UrlEncode(clientSecret));
            body += String.Format("&resource={0}", WebUtility.UrlEncode(resource));
            //body = "grant_type=client_credentials&client_id=1f956aa8-e02c-4510-9cbb-c579cbe8cf6d%408a5ee357-7de0-4836-ab20-9173b12cdce9&client_secret=x0ouYhYb5t%2Ba%2BdzPuniYAWb5GuBbzVa%2FtS5LADYYSJY%3D&resource=00000003-0000-0ff1-ce00-000000000000%2Fm365x725618.sharepoint.com%408a5ee357-7de0-4836-ab20-9173b12cdce9";
            StringContent content = new StringContent(body);
            //content.Headers.ContentType = new conten"application/x-www-form-urlencoded";
            content.Headers.Clear();
            content.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            
            HttpResponseMessage responseToken = await client.PostAsync(url, content);

            string accessToken = null;
            if(responseToken.IsSuccessStatusCode)
            {
                string responseBody = await responseToken.Content.ReadAsStringAsync();
                JObject tokenResult = JObject.Parse(responseBody);
                accessToken = tokenResult["access_token"].ToString();
            }

            if(!String.IsNullOrEmpty(accessToken))
            {
                HttpRequestMessage requestWeb = new HttpRequestMessage() 
                {
                    RequestUri = new Uri(String.Format("https://{0}/sites/FrankTeam1/_api/web",tenantName)),
                    Method=HttpMethod.Get
                };
                //requestWeb.Headers.Clear();
                requestWeb.Headers.Accept.Clear();
                requestWeb.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                //requestWeb.Headers.Add("Accept", "application/json");
                requestWeb.Headers.Add("Authorization", String.Format("Bearer {0}",accessToken));
                HttpResponseMessage responseWeb = await client.SendAsync(requestWeb);

                if(responseWeb.IsSuccessStatusCode)
                {
                    string json = await responseWeb.Content.ReadAsStringAsync();
                    JObject responseWebObj = JObject.Parse(json);
                    Console.WriteLine(responseWebObj.ToString());
                }
            }

            log.LogInformation($"Saying hello to {name}.");
            System.Threading.Thread.Sleep(3000);
            return $"Hello {name}!";
        }


        [FunctionName("MainFunction_HttpStart")]
        public static async Task<HttpResponseMessage> HttpStart(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post")]HttpRequestMessage req,
            [OrchestrationClient]DurableOrchestrationClient starter,
            ILogger log)
        {
            string result = await req.Content.ReadAsStringAsync();
            // Function input comes from the request content.
            string instanceId = await starter.StartNewAsync("MainFunction", null);

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            return starter.CreateCheckStatusResponse(req, instanceId);
        }
        public static async void WriteBlob(string requestJson, ILogger log)
        {
            CloudStorageAccount blobLogsAccount = null;
            string blobLogsDsn = Environment.GetEnvironmentVariable("BlobLogs");
            string blobContainerName = Environment.GetEnvironmentVariable("BlobContainer");

            string blobFile = String.Format("O365Email-{0}", System.Guid.NewGuid().ToString());
            if (CloudStorageAccount.TryParse(blobLogsDsn, out blobLogsAccount))
            {
                CloudBlobClient blobClient = blobLogsAccount.CreateCloudBlobClient();
                CloudBlobContainer blobContainer = blobClient.GetContainerReference(blobContainerName);
                await blobContainer.CreateIfNotExistsAsync();
                CloudBlockBlob blockFile = blobContainer.GetBlockBlobReference(blobFile);
                await blockFile.UploadTextAsync(requestJson);
                log.LogInformation(String.Format("File {0} was created.", blockFile.StorageUri.ToString()));
            }
            else
            {
                log.LogError(String.Format("Canot find storage account"));
            }

        }
    }
}
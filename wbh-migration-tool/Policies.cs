using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using Flurl.Http.Configuration;
using Polly;
using Polly.Retry;
using Polly.Timeout;
using Polly.Wrap;

namespace PRGX.Panoptic_Migration_Core
{
    public static class Policies
    {
        private static List<int> PermFailCodesList = new List<int>(new int[] { 400, 401, 403, 405 });
        private static AsyncTimeoutPolicy<HttpResponseMessage> TimeoutPolicy
        {
            get
            {
                return Policy.TimeoutAsync<HttpResponseMessage>(5, (context, timeSpan, task) =>
                {
                    Logger.log.Warn($"[Timeout Policy]: Timeout delegate fired after {timeSpan.Seconds} seconds");
                    return Task.CompletedTask;
                });
            }
        }

        public static AsyncRetryPolicy<HttpResponseMessage> RetryPolicy
        {
            get
            {
                return Policy
                    .HandleResult<HttpResponseMessage>(r =>
                    {
                        //include StatusCode 415 for Doc virus scan error
                        return ((!r.IsSuccessStatusCode) && (!PermFailCodesList.Contains((int)r.StatusCode)));
                    })
                    .Or<TimeoutRejectedException>()
                    .WaitAndRetryAsync(new[]
                        {
                        TimeSpan.FromSeconds(1),
                        TimeSpan.FromSeconds(2),
                        TimeSpan.FromSeconds(5)
                        },
                        (delegateResult, retryCount) =>
                        {
                            Logger.log.Warn($"Status Code: {delegateResult.Result.StatusCode}. Retry delegate fired, waiting for {retryCount.Seconds} seconds");
                        });
            }
        }

        public static AsyncPolicyWrap<HttpResponseMessage> PolicyStrategy => Policy.WrapAsync(RetryPolicy, TimeoutPolicy);
    }

    public class PolicyHandler : DelegatingHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            //return Policies.PolicyStrategy.ExecuteAsync(ct => base.SendAsync(request, ct), cancellationToken);
            return Policies.RetryPolicy.ExecuteAsync(ct => base.SendAsync(request, ct), cancellationToken);
        }
    }

    public class PollyHttpClientFactory : DefaultHttpClientFactory
    {
        public override HttpMessageHandler CreateMessageHandler()
        {
            var httpMessageHandler = base.CreateMessageHandler();

            // By default, Flurl creates HttpClientHandlers as message handlers.
            // Confirm this is what it did, and then attach a custom behavior.
            if (httpMessageHandler is HttpClientHandler httpClientHandler)
            {
                httpClientHandler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;
            }
            else
            {
                Logger.log.Warn("HttpMessageHandler is type {MessageHandlerType}. Cannot set custom certification validation callback.");
            }
            return new PolicyHandler
            {
                InnerHandler = httpMessageHandler
            };
        }
    }

    public class CustomHttpClientFactory : DefaultHttpClientFactory
    {
        public override HttpMessageHandler CreateMessageHandler()
        {
            var httpMessageHandler = base.CreateMessageHandler();

            // By default, Flurl creates HttpClientHandlers as message handlers.
            // Confirm this is what it did, and then attach a custom behavior.
            if (httpMessageHandler is HttpClientHandler httpClientHandler)
            {
                httpClientHandler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;
            }
            else
            {
                Logger.log.Warn("HttpMessageHandler is type {MessageHandlerType}. Cannot set custom certification validation callback.");
            }
            return httpMessageHandler;
        }

    }
}

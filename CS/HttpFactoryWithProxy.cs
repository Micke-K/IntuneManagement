using System;
using System.Net;
using System.Net.Http;
using Microsoft.Identity.Client;

public class HttpFactoryWithProxy : IMsalHttpClientFactory
{
    private static HttpClient _httpClient;

    public HttpFactoryWithProxy(string proxyURI) : this(proxyURI, null, null)
    {

    }

    public HttpFactoryWithProxy(string proxyURI, string proxyUserName = null, string proxyPassword = null)
    {
        if (_httpClient == null) 
        {
            var proxy = new WebProxy
            {
                Address = new Uri(proxyURI),
                BypassProxyOnLocal = false,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(
                    userName: proxyUserName,
                    password: proxyPassword)
            };

            var httpClientHandler = new HttpClientHandler
            {
                Proxy = proxy,
            };

            _httpClient = new HttpClient(handler: httpClientHandler);
        }
    }

    public HttpClient GetHttpClient()
    {
        return _httpClient;
    }
}
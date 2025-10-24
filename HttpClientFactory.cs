using System.Net;

namespace RoomExporter;

public static class HttpClientFactory
{
    public static HttpClient Create()
    {
        var http = new HttpClient(new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        });
        
        http.DefaultRequestHeaders.UserAgent.ParseAdd("RoomLoadExporter/1.0");
        return http;
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace OpenExcel
{
    class Program
    {
        static void Main(string[] args)
        {
           
            System.IO.File.Move(@"F:\Excel\test.vexcel",@"F:\Excel\test.xlsx" );
            Excel.Application xlapp;
            Excel.Workbook xlworkbook;
            xlapp = new Excel.Application();

            xlworkbook = xlapp.Workbooks.Open(@"F:\Excel\test.xlsx", 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            xlapp.Visible = true;
            while (IsFileLocked(new FileInfo(@"F:\Excel\test.xlsx")))
            {
                
            }
            System.IO.File.Move(@"F:\Excel\test.xlsx", @"F:\Excel\test.vexcel");

        }
        protected  static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
               return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
        public static class HttpClientEx
        {
            public const string MimeJson = "application/json";

            public static Task<HttpResponseMessage> PatchAsync(this HttpClient client, string requestUri, HttpContent content)
            {
                HttpRequestMessage request = new HttpRequestMessage
                {
                    Method = new HttpMethod("PATCH"),
                    RequestUri = new Uri(client.BaseAddress + requestUri),
                    Content = content,
                };

                return client.SendAsync(request);
            }

            public static Task<HttpResponseMessage> PostJsonAsync(this HttpClient client, string requestUri, Type type, object value)
            {
                return client.PostAsync(requestUri, new ObjectContent(type, value, new JsonMediaTypeFormatter(), MimeJson));
            }

            public static Task<HttpResponseMessage> PutJsonAsync(this HttpClient client, string requestUri, Type type, object value)
            {
                return client.PutAsync(requestUri, new ObjectContent(type, value, new JsonMediaTypeFormatter(), MimeJson));
            }

            public static Task<HttpResponseMessage> PatchJsonAsync(this HttpClient client, string requestUri, Type type, object value)
            {
                return client.PatchAsync(requestUri, new ObjectContent(type, value, new JsonMediaTypeFormatter(), MimeJson));
            }
        }
    }
}

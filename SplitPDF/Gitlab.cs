using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using SplitPDF;
using System.Reflection;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using static System.Net.WebRequestMethods;

namespace SplitPDF
{
    class Gitlab
    {
        private string gitlabserver = "http://gitlab.internal.28b.co.uk/api/v3/projects/";
        public string project = "";
        private string issuespath = "issues";
        private string uploadpath = "uploads";
        private int markdownhashlength = 11;
        private int markdowntraillength = 2;
        private string token = "xLZvX4nTxYrYN2HdEVfs";

        public async void raiseIssue(string filepath, string title, string description)
        {
            if (project == "") return;
            var client = new HttpClient();
            MultipartFormDataContent form = new MultipartFormDataContent();
                try {
                    var stream = System.IO.File.Open(filepath, FileMode.Open, FileAccess.Read);
                    HttpContent content = new StreamContent(stream);
                    form.Add(content, "file");
                    content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
                    {
                        Name = "file",
                        FileName = stream.Name
                    };
                    form.Headers.Add("PRIVATE-TOKEN", token);
                    // Get the response.
                    HttpResponseMessage response = await client.PostAsync(gitlabserver + "/" + project + "/" + uploadpath, form);
                    HttpContent responseContent = response.Content;
                    // Get the stream of the content.
                    using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
                    {
                        string results = (await reader.ReadToEndAsync());
                        //Get the hash for the uploaded image
                        int markdownlength = results.IndexOf("markdown") + markdownhashlength;
                        var hash = results.Substring(markdownlength, (results.LastIndexOf("}") - (markdownlength)) - markdowntraillength);
                        //Raise issue and attach the file to it
                        var requestContent = new FormUrlEncodedContent(new[] {
                        new KeyValuePair<string, string>("title", title),
                        new KeyValuePair<string, string>("description", description + "\r\n " + hash)
                    });
                        requestContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");
                        requestContent.Headers.Add("PRIVATE-TOKEN", token);
                        var url = gitlabserver + "/" + project + "/" + issuespath;
                        var result = client.PostAsync(url, requestContent).Result;
                    }
            }
            catch (Exception e)
                    {
                        Console.Write(e.Message);
                        
                    }
                //http://172.20.1.25/api/v3/projects/3/uploads
            // Get the response content.  Response.ReasonPhrase = "Created" Status 201
        }


    }
}

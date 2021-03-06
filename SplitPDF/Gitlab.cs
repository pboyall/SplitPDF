﻿using System;
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
using System.Web;
using System.Net;

namespace SplitPDF
{
    class Gitlab
    {
        private string gitlabserver = "http://gitlab.internal.28b.co.uk/api/v3/projects/";
        public string project = "";
        private string issuespath = "issues";
        private string uploadpath = "uploads";
        private int markdownhashlength = 11;
        private int markdowntraillength = 1;
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
                //@"g:\Code\screens\SPA_HUM_AxialSPA_UK_EN_AbbVie care experience_LO.png"
                //System.Net.WebUtility.UrlEncode()
                //@"G:\PDFSplitting\output\Thumb16383 HUMIRA axSpA CLM DA JAN 2017 Update_CONSIDERATION Qs_REFERENCEONLY-p2.png"
                form.Headers.Add("PRIVATE-TOKEN", token);
                    // Get the response.
                    HttpResponseMessage response = await client.PostAsync(gitlabserver + project + "/" + uploadpath, form);
                    HttpContent responseContent = response.Content;
                    // Get the stream of the content.
                    using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
                    {
                        string results = (await reader.ReadToEndAsync());
                        //Get the hash for the uploaded image
                        int markdownlength = results.IndexOf("markdown") + markdownhashlength;
                        var hash = results.Substring(markdownlength, (results.LastIndexOf("}") - (markdownlength)) - markdowntraillength);
                    //Raise issue and attach the file to it
                    Dictionary<string, string> postParameters = new Dictionary<string, string>();
                    if (title == "") { title = "Default Title"; }
                    postParameters.Add("title", title);
                    postParameters.Add("description", description + "\r\n " + hash);
                    HttpPostRequest(gitlabserver + project + "/" + issuespath, postParameters);
                    /*System.Net.WebUtility.UrlEncode()
                        var requestContent = new FormUrlEncodedContent(new[] {
                        new KeyValuePair<string, string>("title", System.Net.WebUtility.UrlEncode(title)),
                        new KeyValuePair<string, string>("description", System.Net.WebUtility.UrlEncode(description) + "\r\n " + System.Net.WebUtility.UrlEncode(hash))
                    });
                        requestContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");
                        requestContent.Headers.Add("PRIVATE-TOKEN", token);
                        var url = gitlabserver + project + "/" + issuespath;
                        var result = client.PostAsync(url, requestContent).Result;
                        */
                }
            }
            catch (Exception e)
                    {
                        Console.Write(e.Message);
                        
                    }
                //http://172.20.1.25/api/v3/projects/3/uploads
            // Get the response content.  Response.ReasonPhrase = "Created" Status 201
        }


        private string HttpPostRequest(string url, Dictionary<string, string> postParameters)
        {
            string postData = "";

            foreach (string key in postParameters.Keys)
            {
                postData += System.Net.WebUtility.UrlEncode(key) + "="
                      + System.Net.WebUtility.UrlEncode(postParameters[key]) + "&";
            }

            HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            myHttpWebRequest.Method = "POST";
            myHttpWebRequest.Headers.Add("PRIVATE-TOKEN", token);

            byte[] data = Encoding.ASCII.GetBytes(postData);

            myHttpWebRequest.ContentType = "application/x-www-form-urlencoded";
            myHttpWebRequest.ContentLength = data.Length;

            Stream requestStream = myHttpWebRequest.GetRequestStream();
            requestStream.Write(data, 0, data.Length);
            requestStream.Close();

            HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

            Stream responseStream = myHttpWebResponse.GetResponseStream();

            StreamReader myStreamReader = new StreamReader(responseStream, Encoding.Default);

            string pageContent = myStreamReader.ReadToEnd();

            myStreamReader.Close();
            responseStream.Close();

            myHttpWebResponse.Close();

            return pageContent;
        }

    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SplitPDF;
using System.Reflection;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using static System.Net.WebRequestMethods;
using Newtonsoft.Json;


namespace SplitPDFUI
{
    public partial class Form1 : Form
    {

        private string PDFFolderPath;
        private string outputFolderPath;
        private string mydirectory = "";
        private string project = "";

//Gitlab Settings
        private string gitlabserver = "http://gitlab.internal.28b.co.uk/api/v3/projects/";
        private string issuespath = "issues";
        private string uploadpath = "uploads";
        int markdownhashlength = 11;
        int markdowntraillength = 1;
        string token = "xLZvX4nTxYrYN2HdEVfs";

        public Form1()
        {
            InitializeComponent();
            mydirectory = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]).ToString();

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            //Duplicated code
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPDFFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            //string sourcedirectory = "";
            //string targetdirectory = "";
            string comparisonfile = "";
            string[] systemvalues;
            int loopcounter = 1;
            splitPDF splitter = new splitPDF();

            if (comparisonfile == "") { comparisonfile = mydirectory + "\\2244bb38-5e6b-450a-80dd-c490ec6344b0.xlsx"; }//Just testing
            splitter.comparisonfile = comparisonfile;

            splitter.newProject();
            
            try
            {
                systemvalues = System.IO.File.ReadAllLines(PDFFolderPath + "\\Project.txt");
                foreach (string line in systemvalues)
                {
                    int equalspos = line.IndexOf(":");
                    if (equalspos > 0)
                    {
                        string value = line.Substring(equalspos + 1);
                        string key = line.Substring(0, equalspos);
                        splitter.setProjectProperty(key, value);
                    }
                }
            }
            catch (Exception ee)
            {

            }
            
//Magic
            splitter.renderer.exportDPI = 150;
            splitter.renderer.thumbnailheight = 150;
            splitter.renderer.thumbnailwidth = 200;

            splitter.createPDFs = chkPDFs.Checked;
            splitter.createThumbs = chkThumbs.Checked;
            splitter.consolidatePages = chkConsolidate.Checked;
            splitter.extractText = chkText.Checked;
            splitter.exportNav = chkNav.Checked;
            SplitPDF.gitlabupload exportGitStatus;
            Enum.TryParse<SplitPDF.gitlabupload>(cmbGit.SelectedValue.ToString(), out exportGitStatus);
            splitter.exportGit = exportGitStatus;
            
            splitter.outputfile = outputFolderPath;
            splitter.newProject();

            try { 
                //For each PDF in source directory, run the routine
                string[] dirs = Directory.GetFiles(PDFFolderPath, "*.pdf");
                foreach (string dir in dirs)
                {
                    splitter.inputfile = dir;
                    //Create Presentation
                    splitter.newPresentation(loopcounter, System.IO.Path.GetFileNameWithoutExtension(dir));
                    loopcounter++;
                    //Execute code
                    int returned = splitter.Split();
                    string excelfile = splitter.outputfile + "\\" + dir + ".xlsx";
                    splitter.ExportToExcel(excelfile, "Meta", "Meta");     //No tabname for now - that would be if updating.  Later Guid.NewGuid().ToString() 
                    //splitter.ExportToExcel(excelfile, "Nav", "Nav");     //No tabname for now - that would be if updating.  Later
                    //Metadata Export
                    splitter.ExportMetadata();
                    if (splitter.exportGit == SplitPDF.gitlabupload.New){
                        splitter.ExportToGit(project);
                    }

                }
            }catch(Exception ee)
            {
                MessageBox.Show(ee.Message);
            }


            MessageBox.Show("Complete");

        }

        private void cmdDefault_Click(object sender, EventArgs e)
        {
            txtPDFFolder.Text = "G:\\PDFSplitting\\";
            txtOutputFolder.Text = "G:\\PDFSplitting\\Output";
//            txtPDFFolder.Text = "E:\\PDFSplitter\\";
//            txtOutputFolder.Text = "E:\\PDFSplitter\\Output";
            txtProject.Text = "48";
        }

        private void txtPDFFolder_TextChanged(object sender, EventArgs e)
        {
            PDFFolderPath = txtPDFFolder.Text;
        }

        private void btnCurlTests_Click(object sender, EventArgs e)
        {

            //iterator

            string image = @"g:\Code\screens\SPA_HUM_AxialSPA_UK_EN_AbbVie care experience_LO.png";
            string title = Path.GetFileName(image);
            string description = title + "Test Description";
            raiseIssue(image, title, description);
        }

        private async void raiseIssue(string filepath, string title, string description)
        {
            var client = new HttpClient();
            MultipartFormDataContent form = new MultipartFormDataContent();
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
            //http://172.20.1.25/api/v3/projects/3/uploads
            // Get the response content.  Response.ReasonPhrase = "Created" Status 201
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

        private void txtOutputFolder_TextChanged(object sender, EventArgs e)
        {
            outputFolderPath = txtOutputFolder.Text;
            
        }

        private void label1_Click(object sender, EventArgs e){}
        private void label1_Click_1(object sender, EventArgs e){}

        private void Form1_Load(object sender, EventArgs e)
        {

            cmbGit.DataSource = Enum.GetValues(typeof(SplitPDF.gitlabupload));
        }

        private void cmbGit_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void txtProject_TextChanged(object sender, EventArgs e)
        {
            project = txtProject.Text;
        }
    }
}

/* 
                //curl--request POST --header "PRIVATE-TOKEN: 9koXpg98eAheJpvBs5tK" https://gitlab.example.com/api/v3/projects/4/issues?title=Issues%20with%20auth&labels=bug
                //await reader.ReadToEndAsync())
    dynamic issue = new Newtonsoft.Json.Linq.JObject();
                issue.Title= title;
                issue.description = description;
                /*                var sform = new StringContent(issue.ToString(), Encoding.UTF8, "application/json");
                                sform.Headers.Add("PRIVATE-TOKEN", "xLZvX4nTxYrYN2HdEVfs");
                                sform.Headers.Add("Content-Type", "applcation/x-www-form-urlencoded");



client.BaseAddress = new Uri("http://your.url.com/");
            MultipartFormDataContent form = new MultipartFormDataContent();
            HttpContent content = new StringContent("fileToUpload");
            form.Add(content, "fileToUpload");
            var stream = await file.OpenStreamForReadAsync();
            content = new StreamContent(stream);
            content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data")
            {
                Name = "fileToUpload",
                FileName = file.Name
            };
            form.Add(content);


                // Create the HttpContent for the form to be posted.
            var requestContent = new FormUrlEncodedContent(new[] {
                new KeyValuePair<string, string>("file", "This is a block of text"),
            });

    
    
    var response = await client.PostAsync("upload.php", form);
            return response.Content.ReadAsStringAsync().Result;
*/

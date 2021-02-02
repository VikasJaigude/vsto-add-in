using Microsoft.Web.WebView2.Core;
using System;
using System.Windows.Forms;

namespace OutlookAddIn2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            
            InitializeComponent();
            InitAsync();
            this.Resize += new System.EventHandler(this.Form_Resize);
            this.FormClosed += Form1_FormClosed;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
               // webView.Dispose();
                //webView.Container.Dispose();
            }
            catch(Exception ex)
            {

            }
            
        }

        private async void InitAsync()
        {
            var env = await CoreWebView2Environment.CreateAsync(null, @"C:\temp\edgechromium");
            await webView.EnsureCoreWebView2Async(env);
            //webView.Source = new Uri("https://smartemailer365.azurewebsites.net/");
            webView.Source = new Uri("https://www.microsoft.com/");
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            webView.Size = this.ClientSize - new System.Drawing.Size(webView.Location);
        }
    }
}

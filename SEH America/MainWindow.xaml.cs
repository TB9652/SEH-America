using Syncfusion.Presentation;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net;

namespace SEH_America
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnCreatePowerPoint_Click(object sender, RoutedEventArgs e)
        {
            //Create a new PowerPoint presentation
            IPresentation powerpointDoc = Presentation.Create();

            //Add a blank slide to the presentation
            ISlide slide = powerpointDoc.Slides.Add(SlideLayoutType.Blank);

            //Add a textbox to the slide
            IShape shape = slide.AddTextBox(400, 100, 500, 100);
            shape.TextBody.AddParagraph(titleAreaText.Text);

            //Add a text to the textbox.
            //shape.TextBody.AddParagraph(titleAreaText.Text);
            shape.TextBody.AddParagraph(textAreaText.Text);

            //Save the PowerPoint presentation
            powerpointDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            powerpointDoc.Close();

            //Open the PowerPoint presentation
            System.Diagnostics.Process.Start("Sample.pptx");
        }
    }
}

namespace BingSearchApisQuickstart
{
    class Program
    {
        // Replace the this string with valid access key.
        const string subscriptionKey = "enter your key here";
        const string uriBase = "https://api.cognitive.microsoft.com/bing/v7.0/images/search";
        const string searchTerm = "tropical ocean";

        struct SearchResult
        {
            public String jsonResult;
            public Dictionary<String, String> relevantHeaders;
        }
        static SearchResult BingImageSearch(string searchTerm)
        {
            var uriQuery = uriBase + "?q=" + Uri.EscapeDataString(searchTerm);
            WebRequest request = WebRequest.Create(uriQuery);
            request.Headers["Ocp-Apim-Subscription-Key"] = subscriptionKey;
            HttpWebResponse response = (HttpWebResponse)request.GetResponseAsync().Result;
            string json = new StreamReader(response.GetResponseStream()).ReadToEnd();

            // Create the result object for return
            var searchResult = new SearchResult()
            {
                jsonResult = json,
                relevantHeaders = new Dictionary<String, String>()
            };

            // Extract Bing HTTP headers
            foreach (String header in response.Headers)
            {
                if (header.StartsWith("BingAPIs-") || header.StartsWith("X-MSEdge-"))
                    searchResult.relevantHeaders[header] = response.Headers[header];
            }
            return searchResult;
        }
    }
}


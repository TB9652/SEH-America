using Syncfusion.Presentation;
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

            //Add a text to the textbox.
            shape.TextBody.AddParagraph("Hello World!!!");

            //Save the PowerPoint presentation
            powerpointDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            powerpointDoc.Close();

            //Open the PowerPoint presentation
            System.Diagnostics.Process.Start("Sample.pptx");
        }
    }
}

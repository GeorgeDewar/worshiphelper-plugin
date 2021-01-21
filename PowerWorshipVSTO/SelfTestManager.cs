using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerWorshipVSTO
{
    class SelfTestManager
    {
        Application app = Globals.ThisAddIn.Application;
        private readonly ScriptureManager scriptureManager = new ScriptureManager();
        private readonly SongManager songManager = new SongManager();

        public void SelfTest()
        {
            // Insert song as first item, then scripture, then song
            //ClearPresentation();
            //InsertSlide();
            //songManager.InsertSongFromFile(TestFilePath("TestSong1.pptx"));

            //var sel = app.ActiveWindow.Selection;
            var activeWindow = app.ActiveWindow;

            // If zero slides, insertIndex = 1
            // If selection range exists, use last selected
            // If not, do this trick to set the selection
            activeWindow.ViewType = PpViewType.ppViewSlide;
            activeWindow.ViewType = PpViewType.ppViewNormal;
            
            
            
            Debug.WriteLine(activeWindow.Panes.Count);
            Debug.WriteLine(activeWindow.ActivePane.ToString());

            //activeWindow.Panes;

            //Slide sel = app.ActiveWindow.View.Slide;
            //Debug.WriteLine(sel.SlideIndex);
        }

        private void ClearPresentation()
        {
            app.ActivePresentation.Slides.Range().Delete();
        }

        private void InsertSlide(string content = "Standard Slide")
        {
            var slide = app.ActivePresentation.Slides.Add(app.ActivePresentation.Slides.Count + 1, PpSlideLayout.ppLayoutTitleOnly);
            slide.Shapes[1].TextFrame.TextRange.Text = content;
        }

        private string TestFilePath(string fileName)
        {
            var enviroment = AppDomain.CurrentDomain.BaseDirectory;
            string projectDirectory = Directory.GetParent(enviroment).Parent.Parent.FullName;
            return $"{projectDirectory}\\TestFiles\\{fileName}";
        }
    }
}

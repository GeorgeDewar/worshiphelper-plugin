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
        private readonly Bible bible = OpenSongBibleReader.LoadTranslation("NASB");

        private readonly int DELAY = 500;

        async public void SelfTest()
        {
            // Insert song as first item, then scripture, then song
            ClearPresentation();
            InsertSlide();
            var activeWindow = app.ActiveWindow;
            songManager.InsertSongFromFile(TestFilePath("TestSong1.pptx"));
            await Task.Delay(DELAY);
            scriptureManager.addScripture(bible, "Genesis", 1, 1, 2);
            await Task.Delay(DELAY);
            songManager.InsertSongFromFile(TestFilePath("TestSong2.pptx"));

            var index = 1;
            assertSongContent(index++, "Standard Slide");
            assertSongContent(index++, "Song 1 Slide 1");
            assertSongContent(index++, "Song 1 Slide 2");
            assertSongContent(index++, "Song 1 Slide 3");
            assertScriptureContent(index++, "In the beginning", "Genesis 1:1-2 (NASB)");
            assertSongContent(index++, "Song 2 Slide 1");
            assertSongContent(index++, "Song 2 Slide 2");
            assertSongContent(index++, "Song 2 Slide 3");
            //var sel = app.ActiveWindow.Selection;


            // If zero slides, insertIndex = 1
            // If selection range exists, use last selected
            // If not, do this trick to set the selection
            activeWindow.ViewType = PpViewType.ppViewSlide;
            activeWindow.ViewType = PpViewType.ppViewNormal;
            

        }

        private void assertSongContent(int slideIndex, string content)
        {
            if (slideIndex > app.ActivePresentation.Slides.Count)
            {
                Debug.WriteLine($"Slide {slideIndex} does not exist");
            }
            
            try
            {
                var slide = app.ActivePresentation.Slides.Range(slideIndex);
                var text = slide.Shapes[1].TextFrame.TextRange.Text;

                if (text.Contains(content))
                {
                    Debug.WriteLine($"Slide {slideIndex} OK");
                }
                else
                {
                    Debug.WriteLine($"Expected slide {slideIndex} to contain \"{content}\", but it actually was \"{text}\"");
                }
            } catch (Exception e)
            {
                Debug.WriteLine($"Failed to verify content of slide {slideIndex}: {e.Message}");
            }
        }

        private void assertScriptureContent(int slideIndex, string content, string reference)
        {
            if (slideIndex > app.ActivePresentation.Slides.Count)
            {
                Debug.WriteLine($"Slide {slideIndex} does not exist");
            }

            try
            {
                var slide = app.ActivePresentation.Slides.Range(slideIndex);
                var text = slide.Shapes[2].TextFrame.TextRange.Text;

                if (text.Contains(content))
                {
                    Debug.WriteLine($"Slide {slideIndex} OK");
                }
                else
                {
                    Debug.WriteLine($"Expected slide {slideIndex} to contain \"{content}\", but it actually was \"{text}\"");
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine($"Failed to verify content of slide {slideIndex}: {e.Message}");
            }
        }

        private void ClearPresentation()
        {
            Debug.WriteLine("Clearing presentation");
            app.ActivePresentation.Slides.Range().Delete();
        }

        private void InsertSlide(string content = "Standard Slide")
        {
            Debug.WriteLine($"Inserting slide with content: {content}");
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

using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace WorshipHelperVSTO
{
    class SelfTestManager
    {
        Application app = Globals.ThisAddIn.Application;
        private readonly ScriptureManager scriptureManager = new ScriptureManager();
        private readonly SongManager songManager = new SongManager();
        private readonly Bible bible = OpenSongBibleReader.LoadTranslation("NASB");
        private readonly ScriptureTemplate template = new ScriptureTemplate($@"{ThisAddIn.appDataPath}\Templates\ScriptureTemplate.pptx");

        private readonly int DELAY = 500;

        async public void SelfTest()
        {
            await TestSequentialInsertWithSongFirst();
            await TestSequentialInsertWithScriptureFirst();
            await TestInsertScriptureInMiddle();
            await TestInsertSongInMiddle();
        }

        private async Task TestSequentialInsertWithSongFirst()
        {
            // Insert song as first item, then scripture, then song
            ClearPresentation();
            InsertSlide();
            songManager.InsertSongFromFile(TestFilePath("TestSong1.pptx"));
            await Task.Delay(DELAY);
            scriptureManager.addScripture(template, bible, "Genesis", 1, 1, 2);
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
        }

        private async Task TestSequentialInsertWithScriptureFirst()
        {
            ClearPresentation();
            scriptureManager.addScripture(template, bible, "Genesis", 1, 1, 2);
            await Task.Delay(DELAY);
            songManager.InsertSongFromFile(TestFilePath("TestSong1.pptx"));
            await Task.Delay(DELAY);
            songManager.InsertSongFromFile(TestFilePath("TestSong2.pptx"));

            var index = 1;
            assertScriptureContent(index++, "In the beginning", "Genesis 1:1-2 (NASB)");
            assertSongContent(index++, "Song 1 Slide 1");
            assertSongContent(index++, "Song 1 Slide 2");
            assertSongContent(index++, "Song 1 Slide 3");
            assertSongContent(index++, "Song 2 Slide 1");
            assertSongContent(index++, "Song 2 Slide 2");
            assertSongContent(index++, "Song 2 Slide 3");
        }

        private async Task TestInsertScriptureInMiddle()
        {
            ClearPresentation();
            songManager.InsertSongFromFile(TestFilePath("TestSong1.pptx"));
            await Task.Delay(DELAY);
            songManager.InsertSongFromFile(TestFilePath("TestSong2.pptx"));
            await Task.Delay(DELAY);

            new SelectionManager().GoToSlide(3);
            scriptureManager.addScripture(template, bible, "Genesis", 1, 1, 2);

            var index = 1;
            assertSongContent(index++, "Song 1 Slide 1");
            assertSongContent(index++, "Song 1 Slide 2");
            assertSongContent(index++, "Song 1 Slide 3");
            assertScriptureContent(index++, "In the beginning", "Genesis 1:1-2 (NASB)");
            assertSongContent(index++, "Song 2 Slide 1");
            assertSongContent(index++, "Song 2 Slide 2");
            assertSongContent(index++, "Song 2 Slide 3");
        }

        private async Task TestInsertSongInMiddle()
        {
            ClearPresentation();
            songManager.InsertSongFromFile(TestFilePath("TestSong1.pptx"));
            await Task.Delay(DELAY);
            
            scriptureManager.addScripture(template, bible, "Genesis", 1, 1, 2);
            await Task.Delay(DELAY);

            var index = 1;
            assertSongContent(index++, "Song 1 Slide 1");
            assertSongContent(index++, "Song 1 Slide 2");
            assertSongContent(index++, "Song 1 Slide 3");
            assertScriptureContent(index++, "In the beginning", "Genesis 1:1-2 (NASB)");

            new SelectionManager().GoToSlide(3);
            songManager.InsertSongFromFile(TestFilePath("TestSong2.pptx"));
            
            // Now the scripture should be at the end still, displaced by Song 2
            index = 1;
            assertSongContent(index++, "Song 1 Slide 1");
            assertSongContent(index++, "Song 1 Slide 2");
            assertSongContent(index++, "Song 1 Slide 3");
            assertSongContent(index++, "Song 2 Slide 1");
            assertSongContent(index++, "Song 2 Slide 2");
            assertSongContent(index++, "Song 2 Slide 3");
            assertScriptureContent(index++, "In the beginning", "Genesis 1:1-2 (NASB)");

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

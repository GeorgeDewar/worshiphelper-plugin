using Microsoft.Office.Interop.PowerPoint;
using System.Linq;
using static Microsoft.Office.Core.MsoTriState;

namespace PowerWorshipVSTO
{
    class ScriptureManager
    {
        int maxHeight = 400;

        public void addScripture(string bookName, int chapterNum, int verseNumStart, int verseNumEnd)
        {
            var verseCount = verseNumEnd - verseNumStart + 1;

            Application app = Globals.ThisAddIn.Application;
            
            // Copy the template from the template presentation, and close it
            Presentation templatePresentation = app.Presentations.Open($@"{ThisAddIn.appDataPath}\Templates\ScriptureTemplate.pptx", msoTrue, msoFalse, msoFalse);
            var currentSlide = newSlideFromTemplate(templatePresentation);
            templatePresentation.Close();

            var objBodyTextBox = currentSlide.Shapes[2];
            var objDescTextBox = currentSlide.Shapes[3];
            var originalFontSize = objBodyTextBox.TextFrame.TextRange.Font.Size;

            var translation = "NASB";
            var bibleFile = $@"{ThisAddIn.appDataPath}\Bibles\NASB.xmm";
            var bible = new OpenSongBibleReader().load(bibleFile); // TODO: Inefficient to do this every time

            var chapter = bible.books.Where(item => item.name == bookName).First().chapters.Where(item => item.number == chapterNum).First();
            var verseList = chapter.verses.Where(verse => verse.number >= verseNumStart && verse.number <= verseNumEnd).OrderBy(verse => verse.number).ToList();
            var reference = $"{bookName} {chapterNum}:{verseNumStart}-{verseNumEnd} ({translation})";

            objBodyTextBox.TextFrame.TextRange.Text = "";
            objDescTextBox.TextFrame.TextRange.Text = reference;

            var newSlide = true;
            for (int i = 0; i < verseCount; i++)
            {
                var originalText = objBodyTextBox.TextFrame.TextRange.Text;
                var verseText = "$" + verseList[i].number + "$ " + verseList[i].text + " ";
                objBodyTextBox.TextFrame.TextRange.Text = objBodyTextBox.TextFrame.TextRange.Text + verseText;
                if (objBodyTextBox.Height > maxHeight)
                {
                    if (originalText == "")
                    {
                        // The verse is so long it cannot fit on the slide - make it smaller
                        while(objBodyTextBox.Height > maxHeight)
                        {
                            objBodyTextBox.TextFrame.TextRange.Font.Size -= 1;
                        }
                    } else
                    {
                        // We have overshot the space available on our slide, so *undo* the extra text insertion
                        objBodyTextBox.TextFrame.TextRange.Text = originalText;

                        // ... and move to a new slide
                        currentSlide = currentSlide.Duplicate()[1];
                        currentSlide.MoveTo(app.ActivePresentation.Slides.Count);
                        objBodyTextBox = currentSlide.Shapes[2];
                        objDescTextBox = currentSlide.Shapes[3];

                        objBodyTextBox.TextFrame.TextRange.Font.Size = originalFontSize;
                        objBodyTextBox.TextFrame.TextRange.Text = "";
                        objDescTextBox.TextFrame.TextRange.Text = reference;
                        i--;
                    }
                }
            }

            // Find the verse numbers (prefixed with a $) and superscript them, and remove the $
            for (int slideIndex = 1; slideIndex <= app.ActivePresentation.Slides.Count; slideIndex++)
            {
                currentSlide = app.ActivePresentation.Slides[slideIndex];
                objBodyTextBox = currentSlide.Shapes[2];
                foreach (Verse verse in verseList)
                {
                    string toFind = "$" + verse.number + "$";
                    int verseIndex = objBodyTextBox.TextFrame.TextRange.Text.IndexOf(toFind);
                    if (verseIndex > -1)
                    {
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + 1, toFind.Length).Font.Superscript = msoTrue;
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + 1, 1).Delete();
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + toFind.Length - 1, 1).Delete();
                    }
                }
            }
        }

        private Slide newSlideFromTemplate(Presentation templatePresentation)
        {
            Application app = Globals.ThisAddIn.Application;

            templatePresentation.Slides[1].Copy();
            if (app.ActivePresentation.Slides.Count > 0)
            {
                app.ActivePresentation.Windows[1].View.GotoSlide(app.ActivePresentation.Slides.Count);
            }
            app.ActivePresentation.Slides.Paste();

            return app.ActivePresentation.Slides[app.ActivePresentation.Slides.Count];
        }
    }
}

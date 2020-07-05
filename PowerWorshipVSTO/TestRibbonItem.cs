using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using static Microsoft.Office.Core.MsoTriState;
//using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
//using Office = Microsoft.Office.Core;

namespace PowerWorshipVSTO
{
    public partial class TestRibbonItem
    {
        private void TestRibbonItem_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnInsertSong_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;

            FileDialog dialog = app.FileDialog[MsoFileDialogType.msoFileDialogOpen];
            dialog.Title = "Select Song or Presentation";
            if (dialog.Show() == -1) // If user selected a file
            {
                foreach(string item in dialog.SelectedItems)
                {
                    var sourcePresentation = app.Presentations.Open(item, msoTrue, msoFalse, msoFalse);
                    sourcePresentation.Slides.Range().Copy();
                    app.CommandBars.ExecuteMso("PasteSourceFormatting");
                }
            }
        }

        private void btnInsertScripture_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Presentation macroPresentation = app.Presentations.Open("C:\\Users\\Owner\\Documents\\BibleSlidePOC2.pptm", msoTrue, msoFalse, msoFalse);

            int maxHeight = 400;

            macroPresentation.Slides[1].Copy();
            if (app.ActivePresentation.Slides.Count > 0)
            {
                app.ActivePresentation.Windows[1].View.GotoSlide(app.ActivePresentation.Slides.Count);
            }
            app.ActivePresentation.Slides.Paste();
            var currentSlide = app.ActivePresentation.Slides[app.ActivePresentation.Slides.Count];

            var objBodyTextBox = currentSlide.Shapes[2];
            var objDescTextBox = currentSlide.Shapes[3];

            var reference = "Genesis 1:1-10 (HCSB)";
            var verseList = new string[] {
                "In the beginning God created the heavens and the earth."
                , "Now the earth was formless and empty, darkness covered the surface of the watery depths, and the Spirit of God was hovering over the surface of the waters."
                , "Then God said, \"Let there be light,\" and there was light."
                , "God saw that the light was good, and God separated the light from the darkness."
                , "God called the light \"day,\" and He called the darkness \"night.\" Evening came, and then morning: the first day."
                , "Then God said, \"Let there be an expanse between the waters, separating water from water.\""
                , "So God made the expanse and separated the water under the expanse from the water above the expanse. And it was so."
                , "God called the expanse \"sky.\" Evening came, and then morning: the second day."
                , "Then God said, \"Let the water under the sky be gathered into one place, and let the dry land appear.\" And it was so."
                , "God called the dry land \"earth,\" and He called the gathering of the water \"seas.\" And God saw that it was good."
            };

            objBodyTextBox.TextFrame.TextRange.Text = "";
            objDescTextBox.TextFrame.TextRange.Text = reference;

            for (int i = 1; i < 10; i++)
            {
                var originalText = objBodyTextBox.TextFrame.TextRange.Text;
                var verseText = "$" + i + " " + verseList[i] + " ";
                objBodyTextBox.TextFrame.TextRange.Text = objBodyTextBox.TextFrame.TextRange.Text + verseText;
                if (objBodyTextBox.Height > maxHeight) {
                    // We have overshot the space available on our slide, so *undo* the extra text insertion
                    objBodyTextBox.TextFrame.TextRange.Text = originalText;

                    // ... and move to a new slide
                    currentSlide = app.ActivePresentation.Slides[1].Duplicate()[1];
                    currentSlide.MoveTo(app.ActivePresentation.Slides.Count);
                    objBodyTextBox = currentSlide.Shapes[2];
                    objDescTextBox = currentSlide.Shapes[3];
                    objBodyTextBox.TextFrame.TextRange.Text = verseText;
                    objDescTextBox.TextFrame.TextRange.Text = reference;
                }
            }

            // Find the verse numbers (prefixed with a $) and superscript them, and remove the $
            for (int slideIndex = 1; slideIndex <= app.ActivePresentation.Slides.Count; slideIndex++) {
                currentSlide = app.ActivePresentation.Slides[slideIndex];
                objBodyTextBox = currentSlide.Shapes[2];
                for (int j = 1; j < 10; j++) {
                    string toFind = "$" + j.ToString();
                    int verseIndex = objBodyTextBox.TextFrame.TextRange.Text.IndexOf(toFind);
                    if (verseIndex > -1) {
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + 1, 2).Font.Superscript = msoTrue;
                        objBodyTextBox.TextFrame.TextRange.Characters(verseIndex + 1, 1).Delete();
                    }
                }
            }
        }
    }
}

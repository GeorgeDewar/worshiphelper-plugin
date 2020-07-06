using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerWorshipVSTO
{
    public partial class InsertScriptureForm : Form
    {
        Bible bible;

        public InsertScriptureForm()
        {
            InitializeComponent();
            bible = new OpenSongBibleReader().load(@"C:\PowerWorship\Bibles\NASB.xmm");

            var source = new AutoCompleteStringCollection();
            source.AddRange(bible.books.Select(book => book.name).ToArray());
            txtBook.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtBook.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txtBook.AutoCompleteCustomSource = source;
        }

        private void txtSearchBox_TextChanged(object sender, EventArgs e)
        {
            var text = (sender as TextBox).Text;
            btnInsert.Enabled = isValidReference();
        }

        private void txtSearchBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            var result = base.ProcessCmdKey(ref msg, keyData);

            // If tab is pressed while a book name is being typed, after autocompletion add an extra space
            //if (keyData == Keys.Tab && txtBook.Focused && txtBook.Text.Length > 0)
            //{
            //    txtBook.Text += " ";
            //    txtBook.SelectionStart = txtBook.Text.Length;
            //    txtBook.SelectionLength = 0;
            //}

            return result;
        }

        private bool isValidReference()
        {
            var bookNames = bible.books.Select(book => book.name.ToLower()).ToList();

            var validBook = bookNames.Contains(txtBook.Text.ToLower());
            var validReference = Regex.Match(txtReference.Text, "^[0-9]+(:[0-9]+(-[0-9]+)?)?$").Success;

            return validBook && validReference;
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            var book = bible.books.Find(bookItem => bookItem.name.ToLower() == txtBook.Text.ToLower());
            var referenceParts = txtReference.Text.Split(new char[] { ':', '-' });

            var chapterNum = Int32.Parse(referenceParts[0]);
            var chapter = book.chapters.Find(chapterItem => chapterItem.number == chapterNum);

            int verseNumStart;
            int verseNumEnd;
            if (referenceParts.Length > 2) {
                verseNumStart = Int32.Parse(referenceParts[1]);
                verseNumEnd = Int32.Parse(referenceParts[2]);
            } else if(referenceParts.Length > 1) {
                verseNumStart = Int32.Parse(referenceParts[1]);
                verseNumEnd = verseNumStart;
            } else {
                // No verses were specified, so use the whole range
                verseNumStart = 1;
                verseNumEnd = chapter.verses.Last().number;
            }

            new ScriptureManager().addScripture(book.name, chapterNum, verseNumStart, verseNumEnd);
            this.Close();
        }

        private void txtReference_TextChanged(object sender, EventArgs e)
        {
            btnInsert.Enabled = isValidReference();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

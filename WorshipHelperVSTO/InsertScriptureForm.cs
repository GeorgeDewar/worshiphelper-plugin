using Microsoft.Win32;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public partial class InsertScriptureForm : Form
    {
        Bible bible;

        public InsertScriptureForm()
        {
            InitializeComponent();

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            var lastBible = registryKey.GetValue("LastBibleTranslation") as string;

            // Get a list of installed bibles
            var installedBibleFiles = Directory.GetFiles($@"{ThisAddIn.appDataPath}\Bibles");
            foreach (var file in installedBibleFiles)
            {
                var bibleName = file.Split(new char[] { '\\' }).Last().Replace(".xmm", "");
                cmbTranslation.Items.Add(bibleName);
                if (bibleName == lastBible)
                {
                    cmbTranslation.SelectedItem = bibleName;
                }
            }
            if (cmbTranslation.SelectedItem == null)
            {
                cmbTranslation.SelectedIndex = 0;
            }

            // Initialise so that we can populate the books
            bible = OpenSongBibleReader.LoadTranslation(cmbTranslation.SelectedItem as string);

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

        private bool isValidReference()
        {
            var bookNames = bible.books.Select(book => book.name.ToLower()).ToList();

            var validBook = bookNames.Contains(txtBook.Text.ToLower());
            var validReference = Regex.Match(txtReference.Text, "^[0-9]+(:[0-9]+(-[0-9]+)?)?$").Success;

            if (validBook && validReference)
            {
                try
                {
                    ScriptureReference.parse(bible, txtBook.Text, txtReference.Text);
                    return true;
                } catch(Exception e)
                {
                    // This will happen if parsing fails due to a bad reference
                    return false;
                }
            } else
            {
                return false;
            }
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

            new ScriptureManager().addScripture(bible, book.name, chapterNum, verseNumStart, verseNumEnd);
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

        private void cmbTranslation_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var box = (sender as ComboBox);
            var translationName = box.SelectedItem as string;

            bible = OpenSongBibleReader.LoadTranslation(translationName);

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            registryKey.SetValue("LastBibleTranslation", translationName);
        }
    }
}

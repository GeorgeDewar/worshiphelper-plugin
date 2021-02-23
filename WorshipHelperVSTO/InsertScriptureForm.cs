using log4net;
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
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        Bible bible;

        public InsertScriptureForm()
        {
            log.Info("Loading InsertScriptureForm");
            InitializeComponent();

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            var lastTemplate = registryKey.GetValue("LastScriptureTemplate") as string;
            var lastBible = registryKey.GetValue("LastBibleTranslation") as string;

            // Get a list of available templates, populate list and set initial selection
            log.Debug("Loading scripture templates");
            var installedTemplateFiles = Directory.GetFiles($@"{ThisAddIn.appDataPath}\Templates", "*.pptx");
            Directory.CreateDirectory($@"{ThisAddIn.userDataPath}\UserTemplates\Scripture");
            var userTemplateFiles = Directory.GetFiles($@"{ThisAddIn.userDataPath}\UserTemplates\Scripture", "*.pptx");
            foreach (var file in installedTemplateFiles.Concat(userTemplateFiles))
            {
                var template = new ScriptureTemplate(file);
                cmbTemplate.Items.Add(template);
                if (template.name == lastTemplate)
                {
                    cmbTemplate.SelectedItem = template;
                }
            }
            if (cmbTemplate.SelectedItem == null)
            {
                cmbTemplate.SelectedIndex = 0;
            }

            // Get a list of installed bibles, populate list and set initial selection
            log.Debug("Loading bibles");
            var installedBibleFiles = Directory.GetFiles($@"{ThisAddIn.appDataPath}\Bibles", "*.xmm");
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
            log.Debug($"Loading default bible ({cmbTranslation.SelectedItem})");
            bible = OpenSongBibleReader.LoadTranslation(cmbTranslation.SelectedItem as string);

            var source = new AutoCompleteStringCollection();
            log.Debug("Adding books");
            source.AddRange(bible.books.Select(book => book.name).ToArray());
            txtBook.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtBook.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txtBook.AutoCompleteCustomSource = source;
        }

        private void txtSearchBox_TextChanged(object sender, EventArgs e)
        {
            btnInsert.Enabled = isValidReference();
        }

        private void txtSearchBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private bool isValidReference()
        {
            log.Debug($"Checking reference validity (book: {txtBook.Text}, reference: {txtReference.Text})");
            var bookNames = bible.books.Select(book => book.name.ToLower()).ToList();

            var validBook = bookNames.Contains(txtBook.Text.ToLower());
            var validReference = Regex.Match(txtReference.Text, "^[0-9]+(:[0-9]+(-[0-9]+)?)?$").Success;

            if (validBook && validReference)
            {
                try
                {
                    log.Debug("Book and reference appear structurally valid; parsing...");
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
            log.Info("About to insert scripture");
            log.Debug("Finding book");
            var book = bible.books.Find(bookItem => bookItem.name.ToLower() == txtBook.Text.ToLower());
            var referenceParts = txtReference.Text.Split(new char[] { ':', '-' });

            log.Debug("Finding chapter");
            var chapterNum = Int32.Parse(referenceParts[0]);
            var chapter = book.chapters.Find(chapterItem => chapterItem.number == chapterNum);

            log.Debug("Finding verses");
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

            log.Debug("Inserting");
            new ScriptureManager().addScripture(cmbTemplate.SelectedItem as ScriptureTemplate, bible, book.name, chapterNum, verseNumStart, verseNumEnd);
            log.Debug("Closing");
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
            log.Info($"Selecting translation: {translationName}");

            bible = OpenSongBibleReader.LoadTranslation(translationName);

            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            registryKey.SetValue("LastBibleTranslation", translationName);
        }

        private void cmbTemplate_SelectionChangeCommitted(object sender, EventArgs e)
        {
            var box = (sender as ComboBox);
            var template = box.SelectedItem as ScriptureTemplate;
            log.Info($"Selected template: {template.name}");
            var registryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\WorshipHelper");
            registryKey.SetValue("LastScriptureTemplate", template.name);
        }
    }

    public class ScriptureTemplate
    {
        public string name { get; }
        public string path { get; }

        public ScriptureTemplate(string path)
        {
            this.path = path;
            this.name = path.Split(new char[] { '\\' }).Last().Replace(".pptx", "");
        }

        override public string ToString()
        {
            return name;
        }
    }
}

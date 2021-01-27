using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public partial class AddContentLiveForm : Form
    {
        public AddContentLiveForm()
        {
            InitializeComponent();
        }

        private void btnScripture_Click(object sender, EventArgs e)
        {
            new InsertScriptureForm().Show();
            Close();
        }

        private void btnSong_Click(object sender, EventArgs e)
        {
            new SongManager().InsertSong();
            // After inserting the song, we are left with the main window in focus
            DocumentWindow presenterView = new WindowManager().GetPresenterView();
            presenterView.Activate();
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

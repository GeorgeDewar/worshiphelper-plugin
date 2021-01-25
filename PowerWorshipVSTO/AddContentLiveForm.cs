using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerWorshipVSTO
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
            Close();
        }

    }
}

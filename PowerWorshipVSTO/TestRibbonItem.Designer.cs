using Microsoft.Office.Core;
using System.IO;

namespace PowerWorshipVSTO
{
    partial class TestRibbonItem : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TestRibbonItem()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            ThisAddIn.PreInitialize();

            var mySlides = Factory.CreateRibbonGroup();
            mySlides.Label = "OneClick";
            tab1.Groups.Add(mySlides);

            var oneClickPath = $@"{ThisAddIn.appDataPath}\OneClick";
            if (Directory.Exists(oneClickPath))
            {
                foreach (string file in Directory.GetFiles($@"{ThisAddIn.appDataPath}\OneClick", "*.pptx"))
                {
                    var slideButton = Factory.CreateRibbonButton();
                    mySlides.Items.Add(slideButton);

                    slideButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
                    var pathParts = file.Split(new char[] { '\\' });
                    slideButton.Label = pathParts[pathParts.Length - 1].Replace(".pptx", "");
                    slideButton.Tag = file;
                    var img = Properties.Resources.microsoft_powerpoint_computer_icons_clip_art_presentation_slide_vector_graphics_png_favpng_1fbdUWQVUmj03uyMzadXbfFG8;
                    slideButton.Image = img;
                    slideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertOneClick_Click);
                }
            }
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnInsertScripture = this.Factory.CreateRibbonButton();
            this.btnInsertSong = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "PowerWorship";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnInsertScripture);
            this.group1.Items.Add(this.btnInsertSong);
            this.group1.Label = "Insert";
            this.group1.Name = "group1";
            // 
            // btnInsertScripture
            // 
            this.btnInsertScripture.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertScripture.Image = global::PowerWorshipVSTO.Properties.Resources._22633_200;
            this.btnInsertScripture.Label = "Add Scripture";
            this.btnInsertScripture.Name = "btnInsertScripture";
            this.btnInsertScripture.ShowImage = true;
            this.btnInsertScripture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertScripture_Click);
            // 
            // btnInsertSong
            // 
            this.btnInsertSong.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInsertSong.Image = global::PowerWorshipVSTO.Properties.Resources.music_note_6;
            this.btnInsertSong.Label = "Add Song";
            this.btnInsertSong.Name = "btnInsertSong";
            this.btnInsertSong.ShowImage = true;
            this.btnInsertSong.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertSong_Click);
            // 
            // TestRibbonItem
            // 
            this.Name = "TestRibbonItem";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TestRibbonItem_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertScripture;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertSong;
    }

    partial class ThisRibbonCollection
    {
        internal TestRibbonItem TestRibbonItem
        {
            get { return this.GetRibbon<TestRibbonItem>(); }
        }
    }
}

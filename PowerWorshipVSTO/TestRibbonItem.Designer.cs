using Microsoft.Office.Core;
using Microsoft.Win32;
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

            var favRegistryKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PowerWorship\Favourites");
            foreach (string file in favRegistryKey.GetValueNames())
            {
                var slideButton = Factory.CreateRibbonSplitButton();
                favouritesGroup.Items.Add(slideButton);

                slideButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
                var pathParts = file.Split(new char[] { '\\' });
                slideButton.Label = pathParts[pathParts.Length - 1].Replace(".pptx", "").Replace(".ppt", "");
                slideButton.Tag = file;
                var img = Properties.Resources.microsoft_powerpoint_computer_icons_clip_art_presentation_slide_vector_graphics_png_favpng_1fbdUWQVUmj03uyMzadXbfFG8;
                slideButton.Image = img;
                slideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertOneClick_Click);

                var insertButton = Factory.CreateRibbonButton();
                insertButton.Label = "Insert this item";
                insertButton.Tag = file;
                insertButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertOneClick_Click);
                slideButton.Items.Add(insertButton);

                var removeButton = Factory.CreateRibbonButton();
                removeButton.Label = "Remove from favourites";
                removeButton.Tag = file;
                removeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveOneClick_Click);
                slideButton.Items.Add(removeButton);
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
            this.favouritesGroup = this.Factory.CreateRibbonGroup();
            this.btnAddFavourite = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.favouritesGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.favouritesGroup);
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
            // favouritesGroup
            // 
            this.favouritesGroup.Items.Add(this.btnAddFavourite);
            this.favouritesGroup.Label = "Favourites";
            this.favouritesGroup.Name = "favouritesGroup";
            // 
            // btnAddFavourite
            // 
            this.btnAddFavourite.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddFavourite.Image = global::PowerWorshipVSTO.Properties.Resources.icon_plus;
            this.btnAddFavourite.Label = "Add Favourite";
            this.btnAddFavourite.Name = "btnAddFavourite";
            this.btnAddFavourite.ShowImage = true;
            this.btnAddFavourite.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddFavourite_Click);
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
            this.favouritesGroup.ResumeLayout(false);
            this.favouritesGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertScripture;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertSong;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup favouritesGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddFavourite;
    }

    partial class ThisRibbonCollection
    {
        internal TestRibbonItem TestRibbonItem
        {
            get { return this.GetRibbon<TestRibbonItem>(); }
        }
    }
}

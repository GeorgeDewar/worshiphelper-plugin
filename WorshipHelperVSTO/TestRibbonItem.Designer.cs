using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;

namespace PowerWorshipVSTO
{
    partial class TestRibbonItem : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private List<RibbonSplitButton> favouriteButtons = new List<RibbonSplitButton>();

        public TestRibbonItem()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            ThisAddIn.PreInitialize();

            // Add a set of predefined buttons for favourites. We add and remove favourites at runtime but we cannot add and remove
            // items from the Ribbon. We can, however, change their attributes and show and hide them...
            for (int i=0; i<5; i++)
            {
                var slideButton = Factory.CreateRibbonSplitButton();
                favouritesGroup.Items.Add(slideButton);

                slideButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
                slideButton.Visible = false;

                var img = Properties.Resources.microsoft_powerpoint_computer_icons_clip_art_presentation_slide_vector_graphics_png_favpng_1fbdUWQVUmj03uyMzadXbfFG8;
                slideButton.Image = img;
                slideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertOneClick_Click);

                var insertButton = Factory.CreateRibbonButton();
                insertButton.Label = "Insert this item";
                insertButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertOneClick_Click);
                slideButton.Items.Add(insertButton);

                var removeButton = Factory.CreateRibbonButton();
                removeButton.Label = "Remove from favourites";
                removeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveOneClick_Click);
                slideButton.Items.Add(removeButton);

                favouriteButtons.Add(slideButton);
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
            this.grpDebug = this.Factory.CreateRibbonGroup();
            this.btnSelfTest = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.favouritesGroup.SuspendLayout();
            this.grpDebug.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.favouritesGroup);
            this.tab1.Groups.Add(this.grpDebug);
            this.tab1.Label = "WorshipHelper";
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
            // grpDebug
            // 
            this.grpDebug.Items.Add(this.btnSelfTest);
            this.grpDebug.Label = "Debug";
            this.grpDebug.Name = "grpDebug";
            // 
            // btnSelfTest
            // 
            this.btnSelfTest.Label = "Self-Test";
            this.btnSelfTest.Name = "btnSelfTest";
            this.btnSelfTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelfTest_Click);
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
            this.grpDebug.ResumeLayout(false);
            this.grpDebug.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertScripture;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertSong;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup favouritesGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddFavourite;
        internal RibbonGroup grpDebug;
        internal RibbonButton btnSelfTest;
    }

    partial class ThisRibbonCollection
    {
        internal TestRibbonItem TestRibbonItem
        {
            get { return this.GetRibbon<TestRibbonItem>(); }
        }
    }
}

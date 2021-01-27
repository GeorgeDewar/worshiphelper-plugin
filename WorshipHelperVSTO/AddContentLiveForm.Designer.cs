namespace WorshipHelperVSTO
{
    partial class AddContentLiveForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddContentLiveForm));
            this.btnScripture = new System.Windows.Forms.Button();
            this.btnSong = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnScripture
            // 
            this.btnScripture.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnScripture.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnScripture.ImageIndex = 0;
            this.btnScripture.ImageList = this.imageList1;
            this.btnScripture.Location = new System.Drawing.Point(39, 40);
            this.btnScripture.Name = "btnScripture";
            this.btnScripture.Padding = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.btnScripture.Size = new System.Drawing.Size(250, 200);
            this.btnScripture.TabIndex = 0;
            this.btnScripture.Text = "&Scripture";
            this.btnScripture.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnScripture.UseVisualStyleBackColor = true;
            this.btnScripture.Click += new System.EventHandler(this.btnScripture_Click);
            // 
            // btnSong
            // 
            this.btnSong.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnSong.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSong.ImageIndex = 1;
            this.btnSong.ImageList = this.imageList1;
            this.btnSong.Location = new System.Drawing.Point(324, 40);
            this.btnSong.Name = "btnSong";
            this.btnSong.Padding = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.btnSong.Size = new System.Drawing.Size(250, 200);
            this.btnSong.TabIndex = 1;
            this.btnSong.Text = "Song or &Presentation";
            this.btnSong.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnSong.UseVisualStyleBackColor = true;
            this.btnSong.Click += new System.EventHandler(this.btnSong_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "bible.png");
            this.imageList1.Images.SetKeyName(1, "music-note.png");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(36, 263);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(326, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "The added content will be inserted after the currently displayed slide";
            // 
            // AddContentLiveForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(620, 312);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSong);
            this.Controls.Add(this.btnScripture);
            this.Name = "AddContentLiveForm";
            this.Text = "Add Content Live";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnScripture;
        private System.Windows.Forms.Button btnSong;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label label1;
    }
}
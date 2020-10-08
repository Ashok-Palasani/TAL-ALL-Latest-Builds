using Tools;
namespace FileWatcherForDDL
{
    partial class Form1
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
        /// 

        private void InitializeComponent(string path,string username, string password,string domainName,string path1, string username1, string password1, string domainName1)
        {
            //using (new Impersonator(username, domainName, password))
            {
                this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
                ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
                this.SuspendLayout();
                // 
                // fileSystemWatcher1
                // 
                this.fileSystemWatcher1.EnableRaisingEvents = true;
                this.fileSystemWatcher1.SynchronizingObject = this;
                this.fileSystemWatcher1.EnableRaisingEvents = true;
                //this.fileSystemWatcher1.Path = "\\\\SRKS_TECH-1\\Users\\Public\\FTP\\AutoUpdatedFile\\\\";
                this.fileSystemWatcher1.Path = path;
                //\\SRKS_TECH-2\Users\Tech-2\J\2016-03-16
                this.fileSystemWatcher1.SynchronizingObject = this;
                this.fileSystemWatcher1.Created += new System.IO.FileSystemEventHandler(this.fileSystemWatcher1_Created);
                // 
                // Form1
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(284, 261);
                this.Name = "Form1";
                this.Text = "Form1";
                ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
                this.ResumeLayout(false);
                this.ShowInTaskbar = false;
                this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                this.ShowIcon = false;
            }

            //using (new Impersonator(username1, domainName1, password1))
            {
                this.fileSystemWatcher2 = new System.IO.FileSystemWatcher();
                ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher2)).BeginInit();
                this.SuspendLayout();
                // 
                // fileSystemWatcher2
                // 
                this.fileSystemWatcher2.EnableRaisingEvents = true;
                this.fileSystemWatcher2.SynchronizingObject = this;
                this.fileSystemWatcher2.EnableRaisingEvents = true;
                //this.fileSystemWatcher2.Path = "\\\\SRKS_TECH-1\\Users\\Public\\FTP\\AutoUpdatedFile\\\\";
                this.fileSystemWatcher2.Path = path1;
                //\\SRKS_TECH-2\Users\Tech-2\J\2016-03-16
                this.fileSystemWatcher2.SynchronizingObject = this;
                this.fileSystemWatcher2.Created += new System.IO.FileSystemEventHandler(this.fileSystemWatcher2_Created);
                // 
                // Form1
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(284, 261);
                this.Name = "Form1";
                this.Text = "Form1";
                ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher2)).EndInit();
                this.ResumeLayout(false);
                this.ShowInTaskbar = false;
                this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                this.ShowIcon = false;
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.ControlBox = false;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "FileWatcher";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.ResumeLayout(false);

        }

        #endregion

        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.IO.FileSystemWatcher fileSystemWatcher2;
    }
}


namespace oneToPdf;
partial class Form1
{
    private System.ComponentModel.IContainer components = null;
    private Button btnExtractPdfs;

    private TextBox extractingInfo;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        this.btnExtractPdfs = new Button();
        this.extractingInfo = new TextBox();

        this.SuspendLayout();
        this.Size = new Size(600, 400);
        this.StartPosition = FormStartPosition.CenterScreen;
        // 
        // btnExtractPdfs
        // 

        this.btnExtractPdfs.Name = "btnExtractPdfs";
        this.btnExtractPdfs.Size = new System.Drawing.Size(200, 50);
        this.btnExtractPdfs.TabIndex = 0;
        this.btnExtractPdfs.Text = "Extract PDFs from OneNote";
        this.btnExtractPdfs.UseVisualStyleBackColor = true;
        this.btnExtractPdfs.Click += new System.EventHandler(this.btnExtractPdfs_Click);
        this.btnExtractPdfs.Location = new Point(
                (this.ClientSize.Width - btnExtractPdfs.Width) / 2,
                30
            );
        this.btnExtractPdfs.Anchor = AnchorStyles.Top;
        //
        // extractingInfo
        //
        this.extractingInfo.Text = "";
        this.extractingInfo.Multiline = true;
        this.extractingInfo.ScrollBars = ScrollBars.Both;
        this.extractingInfo.WordWrap = false;
        this.extractingInfo.Size = new Size(400, 200);
        this.extractingInfo.Location = new Point(
                (this.ClientSize.Width - extractingInfo.Width) / 2,
                this.btnExtractPdfs.Bottom + 20
            );
        this.extractingInfo.Anchor = AnchorStyles.Top;
        // 
        // Form1
        // 

        // Adjust layout on resize
        this.Resize += (sender, e) =>
        {
            this.btnExtractPdfs.Location = new Point(
                (this.ClientSize.Width - this.btnExtractPdfs.Width) / 2,
                30
            );
            this.extractingInfo.Location = new Point(
                (this.ClientSize.Width - this.extractingInfo.Width) / 2,
                this.btnExtractPdfs.Bottom + 20
            );
        };


        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.Controls.Add(this.btnExtractPdfs);
        this.Controls.Add(this.extractingInfo);
        this.Name = "Form1";
        this.Text = "OneNote PDF Extractor";
        this.ResumeLayout(false);
    }
}
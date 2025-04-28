namespace oneToPdf;
partial class Form1
{
    private System.ComponentModel.IContainer components = null;
    private Button btnExtractPdfs;

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
        this.SuspendLayout();
        // 
        // btnExtractPdfs
        // 
        this.btnExtractPdfs.Location = new System.Drawing.Point(94, 73);
        this.btnExtractPdfs.Name = "btnExtractPdfs";
        this.btnExtractPdfs.Size = new System.Drawing.Size(200, 50);
        this.btnExtractPdfs.TabIndex = 0;
        this.btnExtractPdfs.Text = "Extract PDFs from OneNote";
        this.btnExtractPdfs.UseVisualStyleBackColor = true;
        this.btnExtractPdfs.Click += new System.EventHandler(this.btnExtractPdfs_Click);
        // 
        // Form1
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(382, 203);
        this.Controls.Add(this.btnExtractPdfs);
        this.Name = "Form1";
        this.Text = "OneNote PDF Extractor";
        this.ResumeLayout(false);
    }
}
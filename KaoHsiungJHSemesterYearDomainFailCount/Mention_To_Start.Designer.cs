namespace KaoHsiungJHSemesterYearDomainFailCount
{
    partial class Mention_To_Start
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
            this.btn_StartCalculate = new DevComponents.DotNetBar.ButtonX();
            this.btn_Exit = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // btn_StartCalculate
            // 
            this.btn_StartCalculate.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btn_StartCalculate.BackColor = System.Drawing.Color.Transparent;
            this.btn_StartCalculate.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btn_StartCalculate.Location = new System.Drawing.Point(27, 52);
            this.btn_StartCalculate.Name = "btn_StartCalculate";
            this.btn_StartCalculate.Size = new System.Drawing.Size(75, 23);
            this.btn_StartCalculate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btn_StartCalculate.TabIndex = 3;
            this.btn_StartCalculate.Text = "開始計算";
            this.btn_StartCalculate.Click += new System.EventHandler(this.btn_StartCalculate_Click);
            // 
            // btn_Exit
            // 
            this.btn_Exit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btn_Exit.BackColor = System.Drawing.Color.Transparent;
            this.btn_Exit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btn_Exit.Location = new System.Drawing.Point(125, 52);
            this.btn_Exit.Name = "btn_Exit";
            this.btn_Exit.Size = new System.Drawing.Size(75, 23);
            this.btn_Exit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btn_Exit.TabIndex = 4;
            this.btn_Exit.Text = "取消";
            this.btn_Exit.Click += new System.EventHandler(this.btn_Exit_Click);
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(27, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(245, 23);
            this.labelX1.TabIndex = 5;
            this.labelX1.Text = "開始計算全校領域不及格人數";
            // 
            // Mention_To_Start
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(248, 87);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btn_Exit);
            this.Controls.Add(this.btn_StartCalculate);
            this.DoubleBuffered = true;
            this.MaximumSize = new System.Drawing.Size(264, 126);
            this.Name = "Mention_To_Start";
            this.Text = "全年級學年領域統計不及格人數報表";
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btn_StartCalculate;
        private DevComponents.DotNetBar.ButtonX btn_Exit;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}
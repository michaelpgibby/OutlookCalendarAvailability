namespace OutlookCalendarAvailability
{
    partial class MeetingInputForm
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.Label lblAttendees;
        private System.Windows.Forms.TextBox txtAttendees;
        private System.Windows.Forms.Label lblExample;
        private System.Windows.Forms.Label lblFrom;
        private System.Windows.Forms.Label lblTo;
        private System.Windows.Forms.DateTimePicker dtFrom;
        private System.Windows.Forms.DateTimePicker dtTo;

        private System.Windows.Forms.Label lblLength;
        private System.Windows.Forms.ComboBox cmbLength; // << fixed list (30/60/90/120)

        private System.Windows.Forms.ComboBox cmbLocalTZ;
        private System.Windows.Forms.ComboBox cmbClientTZ;
        private System.Windows.Forms.Label lblLocalTZ;
        private System.Windows.Forms.Label lblClientTZ;

        private System.Windows.Forms.TextBox txtLocalStart;
        private System.Windows.Forms.TextBox txtLocalEnd;
        private System.Windows.Forms.TextBox txtClientStart;
        private System.Windows.Forms.TextBox txtClientEnd;
        private System.Windows.Forms.Label lblLocalHours;
        private System.Windows.Forms.Label lblClientHours;

        private System.Windows.Forms.CheckBox chkTentativeBusy;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.Button btnHelp;


        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblAttendees = new System.Windows.Forms.Label();
            this.txtAttendees = new System.Windows.Forms.TextBox();
            this.lblExample = new System.Windows.Forms.Label();
            this.lblFrom = new System.Windows.Forms.Label();
            this.lblTo = new System.Windows.Forms.Label();
            this.dtFrom = new System.Windows.Forms.DateTimePicker();
            this.dtTo = new System.Windows.Forms.DateTimePicker();

            this.lblLength = new System.Windows.Forms.Label();
            this.cmbLength = new System.Windows.Forms.ComboBox();

            this.cmbLocalTZ = new System.Windows.Forms.ComboBox();
            this.cmbClientTZ = new System.Windows.Forms.ComboBox();
            this.lblLocalTZ = new System.Windows.Forms.Label();
            this.lblClientTZ = new System.Windows.Forms.Label();

            this.txtLocalStart = new System.Windows.Forms.TextBox();
            this.txtLocalEnd = new System.Windows.Forms.TextBox();
            this.txtClientStart = new System.Windows.Forms.TextBox();
            this.txtClientEnd = new System.Windows.Forms.TextBox();
            this.lblLocalHours = new System.Windows.Forms.Label();
            this.lblClientHours = new System.Windows.Forms.Label();

            this.chkTentativeBusy = new System.Windows.Forms.CheckBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();

            this.SuspendLayout();
            // 
            // lblAttendees
            // 
            this.lblAttendees.AutoSize = true;
            this.lblAttendees.Location = new System.Drawing.Point(12, 9);
            this.lblAttendees.Name = "lblAttendees";
            this.lblAttendees.Size = new System.Drawing.Size(199, 13);
            this.lblAttendees.Text = "Attendees (comma/semicolon/new line):";
            // 
            // txtAttendees
            // 
            this.txtAttendees.AcceptsReturn = true;
            this.txtAttendees.Location = new System.Drawing.Point(15, 25);
            this.txtAttendees.Multiline = true;
            this.txtAttendees.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtAttendees.Size = new System.Drawing.Size(500, 80);
            // 
            // lblExample
            // 
            this.lblExample.Location = new System.Drawing.Point(15, 110);
            this.lblExample.Size = new System.Drawing.Size(500, 32);
            this.lblExample.Text = "Examples…";
            // 
            // Dates
            // 
            this.lblFrom.AutoSize = true; this.lblFrom.Location = new System.Drawing.Point(12, 150); this.lblFrom.Text = "From";
            this.dtFrom.Location = new System.Drawing.Point(15, 166); this.dtFrom.Size = new System.Drawing.Size(200, 20);

            this.lblTo.AutoSize = true; this.lblTo.Location = new System.Drawing.Point(250, 150); this.lblTo.Text = "To";
            this.dtTo.Location = new System.Drawing.Point(253, 166); this.dtTo.Size = new System.Drawing.Size(200, 20);
            // 
            // Meeting length (ComboBox)
            // 
            this.lblLength.AutoSize = true;
            this.lblLength.Location = new System.Drawing.Point(12, 195);
            this.lblLength.Text = "Required meeting length (minutes)";

            this.cmbLength.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLength.Location = new System.Drawing.Point(15, 211);
            this.cmbLength.Size = new System.Drawing.Size(120, 21);
            this.cmbLength.Items.AddRange(new object[] { "30", "60", "90", "120" });
            this.cmbLength.SelectedIndex = 0;
            // 
            // Time zones
            // 
            this.lblLocalTZ.AutoSize = true; this.lblLocalTZ.Location = new System.Drawing.Point(12, 245); this.lblLocalTZ.Text = "Your time zone";
            this.cmbLocalTZ.Location = new System.Drawing.Point(15, 261); this.cmbLocalTZ.Size = new System.Drawing.Size(150, 21);

            this.lblClientTZ.AutoSize = true; this.lblClientTZ.Location = new System.Drawing.Point(220, 245); this.lblClientTZ.Text = "Client time zone";
            this.cmbClientTZ.Location = new System.Drawing.Point(223, 261); this.cmbClientTZ.Size = new System.Drawing.Size(150, 21);
            // 
            // Working hours
            // 
            this.lblLocalHours.AutoSize = true; this.lblLocalHours.Location = new System.Drawing.Point(12, 292);
            this.lblLocalHours.Text = "Your hours (HH:mm)";
            this.txtLocalStart.Location = new System.Drawing.Point(15, 308); this.txtLocalStart.Size = new System.Drawing.Size(60, 20); this.txtLocalStart.Text = "09:00";
            this.txtLocalEnd.Location = new System.Drawing.Point(81, 308); this.txtLocalEnd.Size = new System.Drawing.Size(60, 20); this.txtLocalEnd.Text = "18:00";

            this.lblClientHours.AutoSize = true; this.lblClientHours.Location = new System.Drawing.Point(220, 292);
            this.lblClientHours.Text = "Client hours (HH:mm)";
            this.txtClientStart.Location = new System.Drawing.Point(223, 308); this.txtClientStart.Size = new System.Drawing.Size(60, 20); this.txtClientStart.Text = "09:00";
            this.txtClientEnd.Location = new System.Drawing.Point(289, 308); this.txtClientEnd.Size = new System.Drawing.Size(60, 20); this.txtClientEnd.Text = "18:00";
            // 
            // Tentative
            // 
            this.chkTentativeBusy.AutoSize = true;
            this.chkTentativeBusy.Location = new System.Drawing.Point(15, 340);
            this.chkTentativeBusy.Text = "Treat Tentative as Busy";
            this.chkTentativeBusy.Checked = true;
            // 
            // OK/Cancel
            // 
            this.btnOk.Location = new System.Drawing.Point(359, 372); this.btnOk.Size = new System.Drawing.Size(75, 24);
            this.btnOk.Text = "OK"; this.btnOk.Click += new System.EventHandler(this.btnOk_Click);

            this.btnCancel.Location = new System.Drawing.Point(440, 372); this.btnCancel.Size = new System.Drawing.Size(75, 24);
            this.btnCancel.Text = "Cancel"; this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);


            // btnAbout
            this.btnAbout = new System.Windows.Forms.Button();
            this.btnAbout.Location = new System.Drawing.Point(15, 372);
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(80, 24);
            this.btnAbout.TabIndex = 200;
            this.btnAbout.Text = "About";
            this.btnAbout.UseVisualStyleBackColor = true;
            this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);

            // btnHelp
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnHelp.Location = new System.Drawing.Point(101, 372);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(110, 24);
            this.btnHelp.TabIndex = 201;
            this.btnHelp.Text = "Help / Feedback";
            this.btnHelp.UseVisualStyleBackColor = true;
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);

            // Add to form controls (ensure this is before ResumeLayout)
            this.Controls.Add(this.btnAbout);
            this.Controls.Add(this.btnHelp);



            // 
            // MeetingInputForm
            // 
            this.ClientSize = new System.Drawing.Size(532, 408);
            this.Controls.Add(this.btnCancel); this.Controls.Add(this.btnOk);
            this.Controls.Add(this.chkTentativeBusy);

            this.Controls.Add(this.txtClientEnd); this.Controls.Add(this.txtClientStart);
            this.Controls.Add(this.txtLocalEnd); this.Controls.Add(this.txtLocalStart);
            this.Controls.Add(this.lblClientHours); this.Controls.Add(this.lblLocalHours);

            this.Controls.Add(this.lblClientTZ); this.Controls.Add(this.cmbClientTZ);
            this.Controls.Add(this.lblLocalTZ); this.Controls.Add(this.cmbLocalTZ);

            this.Controls.Add(this.cmbLength); this.Controls.Add(this.lblLength);

            this.Controls.Add(this.dtTo); this.Controls.Add(this.dtFrom);
            this.Controls.Add(this.lblTo); this.Controls.Add(this.lblFrom);

            this.Controls.Add(this.lblExample);
            this.Controls.Add(this.txtAttendees);
            this.Controls.Add(this.lblAttendees);

            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false; this.MinimizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Find common availability";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}

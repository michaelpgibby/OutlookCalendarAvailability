using System;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarAvailability
{
    public partial class MeetingInputForm : Form
    {
        // change this to your internal domain:
        private const string INTERNAL_DOMAIN = "hcg.com";

        public string[] AttendeeEmails { get; private set; }
        public DateTime DateFrom => dtFrom.Value.Date;
        public DateTime DateTo => dtTo.Value.Date;

        // Reads selected value from the ComboBox (30/60/90/120), defaults to 30
        public int MeetingLengthMinutes
        {
            get
            {
                if (int.TryParse((cmbLength.SelectedItem ?? "30").ToString(), out var m))
                    return m;
                return 30;
            }
        }

        public string LocalTimeZone => (cmbLocalTZ.SelectedItem ?? "Eastern").ToString();
        public string ClientTimeZone => (cmbClientTZ.SelectedItem ?? "Eastern").ToString();
        public string LocalStartHHmm => txtLocalStart.Text.Trim();
        public string LocalEndHHmm => txtLocalEnd.Text.Trim();
        public string ClientStartHHmm => txtClientStart.Text.Trim();
        public string ClientEndHHmm => txtClientEnd.Text.Trim();
        public bool TreatTentativeAsBusy => chkTentativeBusy.Checked;


        private void btnAbout_Click(object sender, EventArgs e)
        {
            var aboutText =
        @"Outlook Calendar Availability
Created by Michael Gibby

Purpose:
Find overlapping free time across a set of coworkers over a date range. Note: This will not find calendar availability for external users unless you
have access to their availability from this same calendar. It is designed for internal users to easily share available times with clients. 

Inputs:
• Attendees: This app is designed for the part of the email before the @. For example, to find availability for mgibby@hcg.com you only
need to type 'mgibby' and not the '@hcg.com' part. 
• Your time zone & hours: The candidate time slots are built in this zone.
• Client time zone & hours: Results include only times within the client’s workday too.
• Meeting length: Only show blocks at least this long.
• Dates: Weekends are skipped automatically.

Tips:
• If an attendee can't be resolved, they will be listed under
  'Does not include (possible typos)'. 
• Settings (time zones, hours, length, attendees) are remembered for next time.";

            MessageBox.Show(aboutText, "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            // Prepare the message body first
            string body =
        $@"Hi Michael,

(Describe your question/feedback here)

Current selections:
- Attendees: {txtAttendees.Text}
- Your TZ / Hours: {(cmbLocalTZ.SelectedItem ?? "N/A")} {txtLocalStart.Text}-{txtLocalEnd.Text}
- Client TZ / Hours: {(cmbClientTZ.SelectedItem ?? "N/A")} {txtClientStart.Text}-{txtClientEnd.Text}
- Meeting length (min): {MeetingLengthMinutes}
- Dates: {dtFrom.Value:MM-dd-yyyy} to {dtTo.Value:MM-dd-yyyy}
";

            // Launch a background thread AFTER the dialog closes
            new Thread(() =>
            {
                Thread.Sleep(500); // wait half a second for modal to close

                try
                {
                    var app = Globals.ThisAddIn.Application;
                    var mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    mail.To = "mgibby@hcg.com";
                    mail.Subject = "Outlook Availability Add-in - Help / Feedback";
                    mail.Body = body;
                    mail.Display(false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open Outlook compose window.\r\n" + ex.Message,
                        "Help / Feedback", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }).Start();

            // Close the dialog immediately
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }







        public MeetingInputForm()
        {
            InitializeComponent();

            // 1) Populate dropdowns FIRST
            cmbLocalTZ.Items.Clear();
            cmbClientTZ.Items.Clear();
            cmbLocalTZ.Items.AddRange(new object[] { "Pacific", "Mountain", "Central", "Eastern", "India (IST)" });
            cmbClientTZ.Items.AddRange(new object[] { "Pacific", "Mountain", "Central", "Eastern", "India (IST)" });

            cmbLength.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLength.Items.Clear();
            cmbLength.Items.AddRange(new object[] { "30", "60", "90", "120" });

            // 2) Now LOAD saved settings (with safe fallbacks)
            // Time zones
            var localTz = Properties.Settings.Default.LastLocalTZ;
            var clientTz = Properties.Settings.Default.LastClientTZ;
            if (string.IsNullOrWhiteSpace(localTz)) localTz = "Eastern";
            if (string.IsNullOrWhiteSpace(clientTz)) clientTz = "Pacific";

            // Try to select saved items; if not present, fall back gracefully
            int idxLocal = cmbLocalTZ.FindStringExact(localTz);
            int idxClient = cmbClientTZ.FindStringExact(clientTz);
            cmbLocalTZ.SelectedIndex = idxLocal >= 0 ? idxLocal : 3; // 3 = Eastern in our list
            cmbClientTZ.SelectedIndex = idxClient >= 0 ? idxClient : 0; // 0 = Pacific

            // Hours
            txtLocalStart.Text = string.IsNullOrWhiteSpace(Properties.Settings.Default.LastLocalStart) ? "09:00" : Properties.Settings.Default.LastLocalStart;
            txtLocalEnd.Text = string.IsNullOrWhiteSpace(Properties.Settings.Default.LastLocalEnd) ? "18:00" : Properties.Settings.Default.LastLocalEnd;
            txtClientStart.Text = string.IsNullOrWhiteSpace(Properties.Settings.Default.LastClientStart) ? "09:00" : Properties.Settings.Default.LastClientStart;
            txtClientEnd.Text = string.IsNullOrWhiteSpace(Properties.Settings.Default.LastClientEnd) ? "18:00" : Properties.Settings.Default.LastClientEnd;

            // Meeting length
            int ml = Properties.Settings.Default.LastMeetingLength > 0 ? Properties.Settings.Default.LastMeetingLength : 30;
            int idxLen = cmbLength.FindStringExact(ml.ToString());
            cmbLength.SelectedIndex = idxLen >= 0 ? idxLen : 0;

            // Attendees (optional convenience)
            if (!string.IsNullOrWhiteSpace(Properties.Settings.Default.LastAttendees))
                txtAttendees.Text = Properties.Settings.Default.LastAttendees;

            // Dates (if you also want to persist, add settings for them; otherwise default here)
            dtFrom.Value = DateTime.Today;
            dtTo.Value = DateTime.Today.AddDays(7);

            // Example label
            lblExample.Text =
                "Examples:\r\n" +
                "  mgibby, abanks   (omit @, defaults to @hcg.com)";
        }


        private void btnOk_Click(object sender, EventArgs e)
        {
            // attendees required
            var raw = txtAttendees.Text.Trim();
            if (string.IsNullOrWhiteSpace(raw))
            {
                MessageBox.Show("Please enter at least one attendee.", "Missing info",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var parts = raw
                .Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s.Contains("@") ? s : $"{s}@{INTERNAL_DOMAIN}")
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (DateTo < DateFrom)
            {
                MessageBox.Show("End date must be on or after start date.", "Date range",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // very light time format validation (HH:mm)
            if (!IsHHmm(txtLocalStart.Text) || !IsHHmm(txtLocalEnd.Text) ||
                !IsHHmm(txtClientStart.Text) || !IsHHmm(txtClientEnd.Text))
            {
                MessageBox.Show("Please use 24-hour time like 09:00 or 17:30.", "Time format",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AttendeeEmails = parts;
            // Save current selections
            Properties.Settings.Default.LastLocalTZ = (cmbLocalTZ.SelectedItem ?? "Eastern").ToString();
            Properties.Settings.Default.LastClientTZ = (cmbClientTZ.SelectedItem ?? "Pacific").ToString();
            Properties.Settings.Default.LastLocalStart = txtLocalStart.Text.Trim();
            Properties.Settings.Default.LastLocalEnd = txtLocalEnd.Text.Trim();
            Properties.Settings.Default.LastClientStart = txtClientStart.Text.Trim();
            Properties.Settings.Default.LastClientEnd = txtClientEnd.Text.Trim();
            Properties.Settings.Default.LastMeetingLength = MeetingLengthMinutes;
            Properties.Settings.Default.LastAttendees = txtAttendees.Text.Trim();
            Properties.Settings.Default.Save();

            DialogResult = DialogResult.OK;
            Close();
        }

        private static bool IsHHmm(string s)
        {
            return DateTime.TryParseExact(s?.Trim(), "HH:mm", null,
                System.Globalization.DateTimeStyles.None, out _);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}

using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarAvailability
{
    public partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        // Constructor that accepts RibbonFactory
        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory()) // Pass the RibbonFactory
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Initialize ribbon components here if needed
        }

        // This method will be triggered when the button in the Ribbon is clicked
        private void btnGenerateAvailability_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Call the CheckAvailability method from ThisAddIn.cs
                Globals.ThisAddIn.CheckAvailability();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error");
            }
        }
    }
}





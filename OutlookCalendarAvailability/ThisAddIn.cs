using System;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using OutlookCalendarAvailability;

namespace OutlookCalendarAvailability
{
    public partial class ThisAddIn
    {
        private string storedLocalTimeZone = "Eastern"; // Persistent variable for local time zone
        private string storedClientTimeZone = "Pacific"; // Persistent variable for client time zone

        // Startup event handler
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // This is where the startup logic goes
        }

        // Shutdown event handler
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // This is where the shutdown logic goes (if needed)
        }




        public void CheckAvailability()
        {
            try
            {
                // Prompt user for necessary inputs
                string localTimeZone = PromptUser("Enter your time zone (Pacific, Mountain, Central, Eastern):", storedLocalTimeZone);
                string clientTimeZone = PromptUser("Enter the client time zone (Pacific, Mountain, Central, Eastern):", storedClientTimeZone);

                string startDate = PromptUser("Enter the start date (MM-DD-YYYY):", DateTime.Now.ToString("MM-dd-yyyy"));
                string endDate = PromptUser("Enter the end date (MM-DD-YYYY):", DateTime.Now.AddDays(7).ToString("MM-dd-yyyy"));

                string treatTentativeAsBusyInput = PromptUser("Should Tentative be treated as Busy? (Yes/No):", "Yes");
                bool treatTentativeAsBusy = treatTentativeAsBusyInput.Equals("Yes", StringComparison.OrdinalIgnoreCase);

                string localStartTime = PromptUser("Enter your local start time (24-hour format, e.g. 09:00):", "09:00");
                string localEndTime = PromptUser("Enter your local end time (24-hour format, e.g. 18:00):", "18:00");

                string clientStartTime = PromptUser("Enter the client start time (24-hour format, e.g. 09:00):", "09:00");
                string clientEndTime = PromptUser("Enter the client end time (24-hour format, e.g. 18:00):", "18:00");

                string emailsInput = PromptUser("Enter email addresses separated by commas:", "");
                string[] emails = emailsInput.Split(',').Select(email => email.Trim()).ToArray();

                // Ensure proper time format parsing
                int localStart = ParseTime(localStartTime);
                int localEnd = ParseTime(localEndTime);
                int clientStart = ParseTime(clientStartTime);
                int clientEnd = ParseTime(clientEndTime);

                // Generate the availability chart
                string availabilityChart = GenerateAvailabilityChart(
                    startDate, endDate, emails,
                    localTimeZone, clientTimeZone,
                    localStart, localEnd,
                    clientStart, clientEnd,
                    treatTentativeAsBusy);

                // Display the availability chart in a new email
                Outlook.Application outlookApp = Globals.ThisAddIn.Application;
                Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                mailItem.Subject = "Availability Chart";
                mailItem.Body = availabilityChart;
                mailItem.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error");
            }
        }

        // Parse time from 24-hour format (e.g. 09:00 -> 9, 17:00 -> 17)
        private int ParseTime(string time)
        {
            try
            {
                DateTime parsedTime = DateTime.ParseExact(time, "HH:mm", null);
                return parsedTime.Hour;
            }
            catch
            {
                MessageBox.Show("Invalid time format entered. Please use 24-hour format (HH:mm).", "Invalid Input");
                throw new ArgumentException("Invalid time format.");
            }
        }

        private string GenerateAvailabilityChart(
        string startDate, string endDate, string[] emails,
        string localTimeZone, string clientTimeZone,
        int localStartTime, int localEndTime,
        int clientStartTime, int clientEndTime,
        bool treatTentativeAsBusy)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Availability For: " + string.Join(", ", emails));
            sb.AppendLine();

            DateTime reportStartDate = DateTime.Parse(startDate);
            DateTime reportEndDate = DateTime.Parse(endDate);

            for (DateTime currentDate = reportStartDate; currentDate <= reportEndDate; currentDate = currentDate.AddDays(1))
            {
                // Exclude weekends
                if (currentDate.DayOfWeek == DayOfWeek.Saturday || currentDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    continue;
                }

                // Include the day of the week in the email
                sb.AppendLine($"{currentDate:dddd MM-dd-yyyy}"); // 'dddd' gives the full name of the day
                sb.AppendLine("--------------------");

                // Generate time blocks for the day
                var availableSlots = GenerateTimeBlocks(currentDate, localStartTime, localEndTime);

                foreach (var email in emails)
                {
                    availableSlots = FilterAvailableSlots(email, currentDate, treatTentativeAsBusy, localTimeZone, clientTimeZone, clientStartTime, clientEndTime, availableSlots);
                }

                if (availableSlots.Any())
                {
                    // Group and format time slots with both local and client time zones
                    var groupedSlots = GroupConsecutiveTimeSlots(availableSlots);
                    foreach (var (SlotStart, SlotEnd) in groupedSlots)
                    {
                        // Convert the time slot to both time zones
                        DateTime localSlotStart = TimeZoneInfo.ConvertTime(SlotStart, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(localTimeZone)));
                        DateTime localSlotEnd = TimeZoneInfo.ConvertTime(SlotEnd, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(localTimeZone)));
                        DateTime clientSlotStart = TimeZoneInfo.ConvertTime(SlotStart, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(clientTimeZone)));
                        DateTime clientSlotEnd = TimeZoneInfo.ConvertTime(SlotEnd, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(clientTimeZone)));

                        // Format the output to include both local and client times
                        sb.AppendLine($" {localSlotStart:hh:mm tt} - {localSlotEnd:hh:mm tt} {localTimeZone} / {clientSlotStart:hh:mm tt} - {clientSlotEnd:hh:mm tt} {clientTimeZone}");
                    }
                }
                else
                {
                    sb.AppendLine("No available times.");
                }

                sb.AppendLine();
            }

            return sb.ToString();
        }

        private List<DateTime> GenerateTimeBlocks(DateTime date, int startHour, int endHour)
        {
            return Enumerable.Range(startHour, endHour - startHour)
                             .SelectMany(hour => new[] {
                                 new DateTime(date.Year, date.Month, date.Day, hour, 0, 0),
                                 new DateTime(date.Year, date.Month, date.Day, hour, 30, 0)
                             })
                             .ToList();
        }

        private List<DateTime> FilterAvailableSlots(
      string email, DateTime date, bool treatTentativeAsBusy,
      string localTimeZone, string clientTimeZone,
      int clientStartTime, int clientEndTime, List<DateTime> availableSlots)
        {
            try
            {
                Outlook.Application outlookApp = Globals.ThisAddIn.Application;
                Outlook.Recipient recipient = outlookApp.Session.CreateRecipient(email);
                recipient.Resolve();

                if (recipient.Resolved)
                {
                    string freeBusy = recipient.AddressEntry.GetExchangeUser()?.GetFreeBusy(date, 30, true);

                    if (!string.IsNullOrEmpty(freeBusy))
                    {
                        List<DateTime> filteredSlots = new List<DateTime>();

                        foreach (var slot in availableSlots)
                        {
                            // Determine the index for the free/busy string
                            int slotIndex = (slot.Hour * 60 + slot.Minute) / 30;

                            // Check if the index is within bounds of the freeBusy string
                            if (slotIndex >= 0 && slotIndex < freeBusy.Length)
                            {
                                char status = freeBusy[slotIndex];

                                // Include slot only if it meets availability conditions
                                if ((status == '0' || (!treatTentativeAsBusy && status == '1')) &&
                                    IsWithinWorkingHours(slot, localTimeZone, clientTimeZone, clientStartTime, clientEndTime))
                                {
                                    filteredSlots.Add(slot);
                                }
                            }
                        }

                        return filteredSlots;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing availability for {email}: {ex.Message}", "Error");
            }

            // Return an empty list if an error occurs or no data is available
            return new List<DateTime>();
        }

        private bool IsWithinWorkingHours(DateTime slot, string localTimeZone, string clientTimeZone, int clientStartTime, int clientEndTime)
        {
            // Logic to check if the slot is within working hours for local and client
            return true;
        }

        // Helper function to prompt for user input
        private string PromptUser(string prompt, string defaultValue)
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox(prompt, "User Input", defaultValue);
            if (string.IsNullOrWhiteSpace(input))
            {
                input = defaultValue; // use default if user cancels or leaves blank
            }
            return input;
        }

        private string GetTimeZoneId(string timeZone)
        {
            switch (timeZone.ToLower())
            {
                case "pacific":
                    return "Pacific Standard Time";
                case "mountain":
                    return "Mountain Standard Time";
                case "central":
                    return "Central Standard Time";
                case "eastern":
                    return "Eastern Standard Time";
                default:
                    throw new ArgumentException("Unsupported time zone.");
            }
        }

        // Method to group consecutive time slots
        // Method to group consecutive time slots
        private IEnumerable<(DateTime SlotStart, DateTime SlotEnd)> GroupConsecutiveTimeSlots(List<DateTime> slots)
        {
            var groupedSlots = new List<(DateTime, DateTime)>();

            if (slots.Count == 0)
            {
                return groupedSlots;
            }

            // Sort slots by time to ensure consecutive slots are grouped
            var sortedSlots = slots.OrderBy(s => s).ToList();

            DateTime currentSlotStart = sortedSlots[0];
            DateTime currentSlotEnd = currentSlotStart.AddMinutes(30);

            for (int i = 1; i < sortedSlots.Count; i++)
            {
                DateTime slotStart = sortedSlots[i];
                DateTime slotEnd = slotStart.AddMinutes(30);

                // Check if the current slot is consecutive to the previous one
                if (slotStart == currentSlotEnd)
                {
                    // If consecutive, extend the end time of the current group
                    currentSlotEnd = slotEnd;
                }
                else
                {
                    // If not consecutive, finalize the current group and start a new one
                    groupedSlots.Add((currentSlotStart, currentSlotEnd));

                    // Start a new group
                    currentSlotStart = slotStart;
                    currentSlotEnd = slotEnd;
                }
            }

            // Add the last group
            groupedSlots.Add((currentSlotStart, currentSlotEnd));

            return groupedSlots;
        }

    }
}


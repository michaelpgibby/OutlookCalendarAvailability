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
                using (var dlg = new MeetingInputForm())
                {
                    if (dlg.ShowDialog() != DialogResult.OK) return;

                    // Inputs from the single form
                    var localTimeZone = dlg.LocalTimeZone;
                    var clientTimeZone = dlg.ClientTimeZone;

                    var startDate = dlg.DateFrom.ToString("MM-dd-yyyy");
                    var endDate = dlg.DateTo.ToString("MM-dd-yyyy");

                    bool treatTentativeAsBusy = dlg.TreatTentativeAsBusy;

                    string localStartTime = dlg.LocalStartHHmm;
                    string localEndTime = dlg.LocalEndHHmm;
                    string clientStartTime = dlg.ClientStartHHmm;
                    string clientEndTime = dlg.ClientEndHHmm;

                    string[] emails = dlg.AttendeeEmails;
                    int requiredMinutes = dlg.MeetingLengthMinutes;

                    // Parse times (HH:mm)
                    int localStart = ParseTime(localStartTime);
                    int localEnd = ParseTime(localEndTime);
                    int clientStart = ParseTime(clientStartTime);
                    int clientEnd = ParseTime(clientEndTime);

                    // Generate chart + collect "not found" emails
                    var availabilityChart = GenerateAvailabilityChart_Filtered(
                        startDate, endDate, emails,
                        localTimeZone, clientTimeZone,
                        localStart, localEnd, clientStart, clientEnd,
                        treatTentativeAsBusy,
                        requiredMinutes,
                        out var notFoundEmails // will be deduped below
                    );

                    // Deduplicate not-found list (case-insensitive)
                    var notFoundDistinct = notFoundEmails
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    // Compute valid emails = input minus not-found
                    var validEmails = emails
                        .Where(e => !notFoundDistinct.Contains(e, StringComparer.OrdinalIgnoreCase))
                        .ToArray();

                    // If none valid, stop here (avoid misleading "free all day" email)
                    if (validEmails.Length == 0)
                    {
                        MessageBox.Show(
                            "No calendars were found for any of the addresses you entered.\r\n" +
                            "Please check for typos and try again.",
                            "No valid attendees",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return;
                    }

                    // Build the final body:
                    //  - line 1: Availability For: <valid emails>
                    //  - line 2: Does not include (possible typos): <bad emails> (only if any)
                    var header = new StringBuilder();
                    header.AppendLine("Availability For: " + string.Join(", ", validEmails));
                    if (notFoundDistinct.Count > 0)
                    {
                        header.AppendLine("Does not include (possible typos): " +
                                          string.Join(", ", notFoundDistinct));
                    }

                    var finalBody = header.ToString() + "\r\n" + availabilityChart;

                    // Show the result in a new mail
                    Outlook.Application outlookApp = Globals.ThisAddIn.Application;
                    Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                    mailItem.Subject = "Availability Chart";
                    mailItem.Body = finalBody;
                    mailItem.Display(false);
                }
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


        public void ComposeFeedbackEmail(string body)
        {
            try
            {
                Outlook.Application app = Globals.ThisAddIn.Application;
                Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                mail.To = "mgibby@hcg.com";
                mail.Subject = "Outlook Availability Add-in - Help / Feedback";
                mail.Body = body ?? "";
                mail.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open Outlook compose window.\r\n" + ex.Message,
                    "Help / Feedback", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private string GenerateAvailabilityChart_Filtered(
    string startDate, string endDate, string[] emails,
    string localTimeZone, string clientTimeZone,
    int localStartTime, int localEndTime,
    int clientStartTime, int clientEndTime,
    bool treatTentativeAsBusy,
    int requiredMinutes,
    out List<string> notFoundEmails)
        {
            notFoundEmails = new List<string>();
            var sb = new System.Text.StringBuilder();
           
            sb.AppendLine();

            DateTime reportStartDate = DateTime.Parse(startDate);
            DateTime reportEndDate = DateTime.Parse(endDate);

            for (DateTime currentDate = reportStartDate; currentDate <= reportEndDate; currentDate = currentDate.AddDays(1))
            {
                if (currentDate.DayOfWeek == DayOfWeek.Saturday || currentDate.DayOfWeek == DayOfWeek.Sunday)
                    continue;

                sb.AppendLine($"{currentDate:dddd MM-dd-yyyy}");
                sb.AppendLine("--------------------");

                var availableSlots = GenerateTimeBlocks(currentDate, localStartTime, localEndTime, localTimeZone);


                bool hadAnyData = false;
                foreach (var email in emails)
                {
                    availableSlots = FilterAvailableSlots_withNotFound(
                        email, currentDate, treatTentativeAsBusy,
                        localTimeZone, clientTimeZone,
                        clientStartTime, clientEndTime,
                        availableSlots, notFoundEmails, ref hadAnyData);
                }


                var groupedSlots = GroupConsecutiveTimeSlots(availableSlots)
                    .Where(g => (g.SlotEnd - g.SlotStart).TotalMinutes >= requiredMinutes) // <-- length filter
                    .ToList();

                if (groupedSlots.Any())
                {
                    foreach (var (SlotStart, SlotEnd) in groupedSlots)
                    {
                        DateTime localSlotStart = TimeZoneInfo.ConvertTime(SlotStart, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(localTimeZone)));
                        DateTime localSlotEnd = TimeZoneInfo.ConvertTime(SlotEnd, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(localTimeZone)));
                        DateTime clientSlotStart = TimeZoneInfo.ConvertTime(SlotStart, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(clientTimeZone)));
                        DateTime clientSlotEnd = TimeZoneInfo.ConvertTime(SlotEnd, TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(clientTimeZone)));

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


        private List<DateTime> GenerateTimeBlocks(DateTime date, int startHour, int endHour, string selectedLocalTz)
        {
            var list = new List<DateTime>();
            for (int h = startHour; h < endHour; h++)
            {
                list.Add(FromSelectedLocalTzToSystemLocal(date, h, 0, selectedLocalTz));
                list.Add(FromSelectedLocalTzToSystemLocal(date, h, 30, selectedLocalTz));
            }
            return list;
        }



        private List<DateTime> FilterAvailableSlots_withNotFound(
    string email, DateTime date, bool treatTentativeAsBusy,
    string localTimeZone, string clientTimeZone,
    int clientStartTime, int clientEndTime,
    List<DateTime> availableSlots,
    List<string> notFoundEmails,
    ref bool hadAnyData)
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
                        hadAnyData = true; // <-- mark that at least one address returned data

                        List<DateTime> filteredSlots = new List<DateTime>();
                        foreach (var slot in availableSlots)
                        {
                            int slotIndex = (slot.Hour * 60 + slot.Minute) / 30;
                            if (slotIndex >= 0 && slotIndex < freeBusy.Length)
                            {
                                char status = freeBusy[slotIndex];
                                if ((status == '0' || (!treatTentativeAsBusy && status == '1')) &&
                                    IsWithinWorkingHours(slot, localTimeZone, clientTimeZone, clientStartTime, clientEndTime))
                                {
                                    filteredSlots.Add(slot);
                                }
                            }
                        }
                        return filteredSlots;
                    }
                    else
                    {
                        AddUniqueIgnoreCase(notFoundEmails, email);
                        return availableSlots;
                    }
                }
                else
                {
                    AddUniqueIgnoreCase(notFoundEmails, email);
                    return availableSlots;
                }
            }
            catch
            {
                AddUniqueIgnoreCase(notFoundEmails, email);
                return availableSlots;
            }

        }


        private static void AddUniqueIgnoreCase(List<string> list, string email)
        {
            if (!list.Exists(x => string.Equals(x, email, StringComparison.OrdinalIgnoreCase)))
                list.Add(email);
        }


        private bool IsWithinWorkingHours(
      DateTime slot,
      string localTimeZone, string clientTimeZone,
      int clientStartTime, int clientEndTime)
        {
            try
            {
                // Convert slot (system-local) into each chosen TZ
                var tzLocal = TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(localTimeZone));
                var tzClient = TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(clientTimeZone));

                var slotLocal = TimeZoneInfo.ConvertTime(slot, tzLocal);
                var slotClient = TimeZoneInfo.ConvertTime(slot, tzClient);

                // Check hours in both zones
                bool inLocal = slotLocal.Hour >= 0; // we already restricted slots by local hours when creating them
                bool inClient = slotClient.Hour >= clientStartTime && slotClient.Hour < clientEndTime;

                return inLocal && inClient;
            }
            catch
            {
                // If conversion fails, be safe and exclude
                return false;
            }
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


        private DateTime FromSelectedLocalTzToSystemLocal(DateTime date, int hour, int minute, string selectedLocalTz)
        {
            // Build a wall-clock time IN the selected local TZ (unspecified), then interpret/convert to system local.
            var tzLocal = TimeZoneInfo.FindSystemTimeZoneById(GetTimeZoneId(selectedLocalTz));
            var unspecified = new DateTime(date.Year, date.Month, date.Day, hour, minute, 0, DateTimeKind.Unspecified);
            // Treat 'unspecified' as time in tzLocal:
            var dtoInSelected = new DateTimeOffset(unspecified, tzLocal.GetUtcOffset(unspecified));
            // Convert to system local time
            var dtoLocal = TimeZoneInfo.ConvertTime(dtoInSelected, TimeZoneInfo.Local);
            return dtoLocal.LocalDateTime;
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
                case "india":
                case "ist":
                case "india (ist)":
                    return "India Standard Time";
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


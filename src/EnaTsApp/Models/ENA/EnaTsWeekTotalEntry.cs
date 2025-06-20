using System;
using System.Text;
using Com.Ena.Timesheet.Ena;

namespace Com.Ena.Timesheet.Ena
{
    /// <summary>
    /// An entry that represents a total for a week
    /// </summary>
    public class EnaTsWeekTotalEntry : EnaTsEntry
    {
        private string hoursLabel;
        private string chargeLabel;

        public EnaTsWeekTotalEntry(float entryId, string hoursLabel, float? totalHours, string chargeLabel, float? totalCharge)
        {
            this.hoursLabel = hoursLabel;
            this.chargeLabel = chargeLabel;
            this.EntryId = entryId;
            this.ProjectId = "";
            this.Activity = "";
            this.Hours = totalHours;
            this.Description = chargeLabel;
            this.Charge = totalCharge;
        }

        public void SetProjectId(string projectId)
        {
            this.ProjectId = projectId;
        }

        public void SetActivity(string activity)
        {
            this.Activity = activity;
        }

        public void SetHours(float? hours)
        {
            this.Hours = hours;
        }

        public void SetDescription(string description)
        {
            this.Description = description;
        }

        public void SetCharge(float? charge)
        {
            this.Charge = charge;
        }

        /// <summary>
        /// A hack to get the total hours entry to show up in the table
        /// </summary>
        public override string GetDate()
        {
            return hoursLabel;
        }

        public override string GetRate()
        {
            return "";
        }

        public override string HtmlClass()
        {
            return " class='week-total'";
        }
    }
}
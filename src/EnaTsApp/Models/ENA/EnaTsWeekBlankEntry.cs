using System;
using System.Text;
using Com.Ena.Timesheet.Ena;

namespace Com.Ena.Timesheet.Ena
{
    /// <summary>
    /// An entry that represents a blank line in the table
    /// </summary>
    public class EnaTsWeekBlankEntry : EnaTsEntry
    {
        public EnaTsWeekBlankEntry(float entryId)
        {
            this.EntryId = entryId;
            this.ProjectId = "";
            this.Activity = "";
            this.Description = "";
        }

        public void SetProjectId(string projectId)
        {
            this.ProjectId = projectId;
        }

        public void SetActivity(string activity)
        {
            this.Activity = activity;
        }

        public void SetDescription(string description)
        {
            this.Description = description;
        }

        public override string GetDate()
        {
            return "";
        }

        public override string FormattedHours()
        {
            return "";
        }

        public override string GetRate()
        {
            return "";
        }

        public override string GetCharge()
        {
            return "";
        }
    }
}
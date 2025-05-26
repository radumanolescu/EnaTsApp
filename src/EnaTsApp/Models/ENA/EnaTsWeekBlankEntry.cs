namespace Com.Ena.Timesheet.Ena
{
    /// <summary>
    /// An entry that represents a blank line in the table
    /// </summary>
    public class EnaTsWeekBlankEntry : EnaTsEntry
    {
        public EnaTsWeekBlankEntry(float entryId)
        {
            this.SetEntryId(entryId);
            this.SetProjectId("");
            this.SetActivity("");
            this.SetDescription("");
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
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
            this.SetEntryId(entryId);
            this.SetProjectId("");
            this.SetActivity("");
            this.SetHours(totalHours);
            this.SetDescription(chargeLabel);
            this.SetCharge(totalCharge);
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
    }
}
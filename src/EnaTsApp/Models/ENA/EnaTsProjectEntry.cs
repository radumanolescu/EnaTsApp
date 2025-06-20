using System;
using System.Globalization;

namespace Com.Ena.Timesheet.Ena
{
    public class EnaTsProjectEntry : IComparable<EnaTsProjectEntry>
    {
        private static readonly NumberFormatInfo decimalFormat = new NumberFormatInfo { NumberDecimalDigits = 2 };

        private string projectId = "";
        private string activity = "";
        protected float hours;

        public EnaTsProjectEntry(string projectId, string activity, float hours)
        {
            this.projectId = projectId;
            this.activity = activity;
            this.hours = hours;
        }

        public string ProjectId
        {
            get => projectId;
            set => projectId = value;
        }

        public string Activity
        {
            get => activity;
            set => activity = value;
        }

        public float Hours
        {
            get => hours;
            set => hours = value;
        }

        public string FormattedHours
        {
            get => hours.ToString("#.##", decimalFormat);
        }

        public void SetHours(float hours)
        {
            this.hours = hours;
        }

        public override bool Equals(object? obj)
        {
            if (obj is EnaTsProjectEntry that)
            {
                return projectId == that.projectId &&
                       activity == that.activity &&
                       hours == that.hours;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(projectId, activity, hours);
        }

        private string ProjectActivity()
        {
            return projectId + "#" + activity;
        }

        public int CompareTo(EnaTsProjectEntry? that)
        {
            if (that == null) throw new ArgumentNullException(nameof(that));
            return this.ProjectActivity().CompareTo(that.ProjectActivity());
        }

        public override string ToString()
        {
            return $"{projectId}, {activity}, {Hours}";
        }

        public string HtmlClass()
        {
            return " class='month-total'";
        }

        public string GetHtmlRow()
        {
            string htmlClass = HtmlClass();
            return $"<tr{htmlClass}><td>{projectId}</td><td>{activity}</td><td>{FormattedHours}</td></tr>";
        }
    }
}
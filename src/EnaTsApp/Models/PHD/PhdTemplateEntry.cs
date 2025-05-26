using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Newtonsoft.Json;

namespace Com.Ena.Timesheet.Phd
{
    /// <summary>
    /// Model for a single row in the PHD template
    /// </summary>
    public class PhdTemplateEntry
    {
        /// <summary>
        /// Apache POI row number, zero-based
        /// </summary>
        public int RowNum { get; set; }
        public string Client { get; set; }
        public string Task { get; set; }

        private Dictionary<int, double> effort = new Dictionary<int, double>();

        public PhdTemplateEntry(int rowNum, string client, string task)
        {
            RowNum = rowNum;
            Client = client;
            Task = task;
        }

        public int GetRowNum() => RowNum;
        public string GetClient() => Client;
        public string GetTask() => Task;
        public Dictionary<int, double> GetEffort() => Effort;

        public void SetClient(string client)
        {
            Client = client;
        }

        public void SetTask(string task)
        {
            Task = task;
        }

        public void SetEffort(Dictionary<int, double> effort)
        {
            Effort = effort;
        }

        private int? day;
        public int? Day
        {
            get => day;
            set => day = value;
        }
        public void SetDay(int? day)
        {
            this.day = day;
        }

        public string ToJson()
        {
            return JsonConvert.SerializeObject(this);
        }

        public Dictionary<int, double> Effort
        {
            get => effort;
            set => effort = value ?? new Dictionary<int, double>();
        }

        public double TotalHours()
        {
            return effort.Values.Sum();
        }

        //[JsonIgnore]
        public bool IsBlank()
        {
            return string.IsNullOrWhiteSpace(Client) && string.IsNullOrWhiteSpace(Task);
        }

        public string EntryType()
        {
            string cl = string.IsNullOrWhiteSpace(Client) ? "null" : "Client";
            string tk = string.IsNullOrWhiteSpace(Task) ? "null" : "Task";
            return $"{cl}_{tk}";
        }

        /// <summary>
        /// Valid CSV format requires quotes when string contains commas
        /// </summary>
        [Browsable(false)]
        public string ClientCommaTask()
        {
            return $"\"{Clean(Client)}\",\"{Clean(Task)}\"";
        }

        [Browsable(false)]
        public string ClientTaskEffort()
        {
            var clientTask = new System.Text.StringBuilder(ClientCommaTask());
            for (int i = 1; i <= 31; i++)
            {
                clientTask.Append(",").Append(effort.TryGetValue(i, out var val) ? val : 0.0);
            }
            return clientTask.ToString();
        }

        /// <summary>
        /// A concatenation of `client#task`, stripped of all quotes, separated by #
        /// </summary>
        [Browsable(false)]
        public string ClientHashTask()
        {
            return Unquote(Client) + "#" + Unquote(Task);
        }

        private string Clean(string s)
        {
            return Unquote(s).Replace(",", "");
        }

        private static string Unquote(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\"", "");
        }

        public override bool Equals(object obj)
        {
            if (obj is PhdTemplateEntry other)
            {
                return string.Equals(Client, other.Client) && string.Equals(Task, other.Task);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Client, Task);
        }
    }
}
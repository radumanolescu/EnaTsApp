using System;
using System.Collections.Generic;
using Com.Ena.Timesheet.Phd;
using Com.Ena.Timesheet.Xl;
using Ena.Timesheet.Util;

namespace Com.Ena.Timesheet.Phd
{
    /// <summary>
    /// A parser for a PHD template, which has the following structure (in terms of rows):
    /// HEADER,(ClientBlock)+,SUM
    /// where:
    /// HEADER = "CLIENT,TASK,1,2,3,...,31,TOTALS"
    /// ClientBlock = clientName,taskName,effort1,effort2,...,effort31,total
    /// [empty line, all cells are empty]
    /// SUM ="SUM,Total,1,2,3,...,31"
    /// Here, the number "31" stands for the last day of the month.
    /// </summary>
    public class Parser : ExcelParser
    {
        /// <summary>
        /// Once the project code is set for all entries, check that there are no duplicate client-task pairs.
        /// Such errors have been observed in the past, and they are likely to be caused by manual errors in the template.
        /// When such a case is detected, the task is modified to include the word "dup".
        /// This will ensure that no effort will be booked against that entry.
        /// </summary>
        public static void CheckDupClientTask(List<PhdTemplateEntry> entries)
        {
            var clientTaskSet = new HashSet<string>();
            foreach (var entry in entries)
            {
                string clientTask = entry.ClientCommaTask();
                if (clientTaskSet.Contains(clientTask))
                {
                    string errMsg = $"Duplicate client-task in row {entry.GetRowNum() + 1}: '{clientTask}'";
                    Console.WriteLine(errMsg);
                    entry.SetTask($"{entry.GetTask()} (dup)");
                }
                // No point in adding empty client-task pairs, which are represented as "","".
                if (!string.IsNullOrEmpty(clientTask) && clientTask.Length > 5)
                {
                    clientTaskSet.Add(clientTask);
                }
            }
        }

        /// <summary>
        /// Parse the PHD template and set the project codes for all the entries.
        /// Iterate over the entries in reverse order, starting from the last line.
        /// Use empty lines to separate groups of tasks.
        /// Inside each group, the project code is the first non-empty value in the "Client" field (iterating from end to start).
        /// Once the project code is found, set it (the client property) for all the entries in the group.
        /// </summary>
        public static void SetProjectCodes(List<PhdTemplateEntry> entries)
        {
            string projectCde = "";
            int projectBgn = -1;
            int projectEnd = -1;

            int numLin = entries.Count;
            while (numLin > 0 && entries[numLin - 1].IsBlank())
            {
                entries.RemoveAt(numLin - 1); // remove last entries if they are empty
                numLin--;
            }
            // Starting from the last line, find the project code and boundaries ([start,end] line IDs).
            // Once found, set the project code (the client property) for all the entries in the project.
            for (int lineNum = numLin - 1; lineNum >= 0; lineNum--)
            {
                P(entries[lineNum], projectCde, projectBgn, projectEnd);
                string entryType = entries[lineNum].EntryType();
                switch (entryType)
                {
                    case "null_null": // empty line between groups
                        // The first empty line after a project group indicates that the project begins at the next line.
                        projectBgn = lineNum + 1;
                        // So set the project code for all the entries in the project
                        S(projectCde, projectBgn, projectEnd, entries);
                        P(entries[lineNum], projectCde, projectBgn, projectEnd);
                        // Reset the project code and boundaries
                        projectBgn = -1;
                        projectEnd = -1;
                        projectCde = "";
                        break;
                    case "null_Task": // activity line
                        // If we have encountered an activity but the project is undefined,
                        // it means we have encountered a new project.
                        if (projectEnd < 0)
                        {
                            projectEnd = lineNum;
                        }
                        break;
                    case "Client_Task":
                        // The project code is the first word of the line,
                        // in the last entry of the group that has a non-empty value in that field.
                        if (string.IsNullOrEmpty(projectCde))
                        {
                            projectCde = entries[lineNum].GetClient();
                        }
                        if (projectEnd < 0)
                        {
                            projectEnd = lineNum;
                        }
                        break;
                    default:
                        throw new InvalidOperationException($"Unexpected entry type at row {lineNum + 1}: {entries[lineNum].ClientCommaTask()} of type {entryType}");
                }
                if (lineNum == 0)
                {
                    projectBgn = 1;
                    S(projectCde, projectBgn, projectEnd, entries);
                }
                P(entries[lineNum], projectCde, projectBgn, projectEnd);
            }
        }

        private static void P(PhdTemplateEntry w, string pC, int pB, int pE)
        {
            string ws = w.ToJson();
            int xl = w.GetRowNum() + 1;
            string ps = $"{{xl={xl},t={w.EntryType()},c=`{pC}`,b={pB},e={pE}}}";
            Console.WriteLine(ws + "\t\t" + ps);
        }

        private static void S(string pC, int pB, int pE, List<PhdTemplateEntry> entries)
        {
            if (string.IsNullOrEmpty(pC) || pB < 0 || pE < 0)
            {
                return;
            }
            for (int lineNum = pB; lineNum <= pE; lineNum++)
            {
                entries[lineNum].SetClient(pC);
            }
        }
    }
}
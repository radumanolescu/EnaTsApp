using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Text;

namespace Ena.Timesheet.Tests
{
    public static class TestLogger
    {
        public static ILogger<T> CreateLogger<T>() where T : class
        {
            var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder
                    .AddConsole()
                    .SetMinimumLevel(LogLevel.Debug);
            });

            return loggerFactory.CreateLogger<T>();
        }
    }
}

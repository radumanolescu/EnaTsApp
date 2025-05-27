using System;
using Ena.Timesheet.Ena;
using System.Windows.Forms;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using System.IO;
using Microsoft.Extensions.Logging;
using NLog.Extensions.Logging;
using NLog;
using NLog.Web;
using LogLevel = Microsoft.Extensions.Logging.LogLevel;


namespace EnaTsApp
{
    public static class AppEntryPoint
    {
        [STAThread]
        public static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var host = CreateHostBuilder().Build();
            
            using (var scope = host.Services.CreateScope())
            {
                var services = scope.ServiceProvider;
                var mainForm = services.GetRequiredService<MainForm>();
                Application.Run(mainForm);
            }
        }

        private static IHostBuilder CreateHostBuilder()
        {
            var builder = new HostBuilder()
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
                })
                .ConfigureServices((hostContext, services) =>
                {
                    // Configure NLog
                    var logsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "EnaTsApp", "logs");
                    var mainTarget = new NLog.Targets.FileTarget
                    {
                        Name = "main",
                        FileName = Path.Combine(logsDirectory, "enatsapp-${date:format=yyyy-MM-dd}.log"),
                        Layout = "${longdate} | ${level:uppercase=true} | ${logger} | ${message} ${exception:format=tostring}"
                    };

                    var debugTarget = new NLog.Targets.FileTarget
                    {
                        Name = "debug",
                        FileName = Path.Combine(logsDirectory, "enatsapp-debug-${date:format=yyyy-MM-dd}.log"),
                        Layout = "${longdate} | ${level:uppercase=true} | ${logger} | ${message} ${exception:format=tostring}"
                    };

                    var errorTarget = new NLog.Targets.FileTarget
                    {
                        Name = "error",
                        FileName = Path.Combine(logsDirectory, "enatsapp-error-${date:format=yyyy-MM-dd}.log"),
                        Layout = "${longdate} | ${level:uppercase=true} | ${logger} | ${message} ${exception:format=tostring}"
                    };

                    var consoleTarget = new NLog.Targets.ConsoleTarget("console")
                    {
                        Layout = "${longdate} | ${level:uppercase=true} | ${logger} | ${message} ${exception:format=tostring}"
                    };

                    try
                    {
                        // Enable NLog internal logging for debugging
                        NLog.Common.InternalLogger.LogLevel = NLog.LogLevel.Debug;
                        NLog.Common.InternalLogger.LogFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "EnaTsApp", "nlog-internal.log");

                        // Create configuration
                        var nlogConfig = new NLog.Config.LoggingConfiguration();

                        // Add targets to configuration
                        nlogConfig.AddTarget(mainTarget);
                        nlogConfig.AddTarget(debugTarget);
                        nlogConfig.AddTarget(errorTarget);
                        nlogConfig.AddTarget(consoleTarget);

                        // Configure rules
                        nlogConfig.AddRule(NLog.LogLevel.Debug, NLog.LogLevel.Fatal, mainTarget);  // All levels to main log
                        nlogConfig.AddRule(NLog.LogLevel.Debug, NLog.LogLevel.Debug, debugTarget);  // Only debug to debug log
                        nlogConfig.AddRule(NLog.LogLevel.Error, NLog.LogLevel.Fatal, errorTarget, loggerNamePattern: "*");  // Only error/fatal to error log
                        nlogConfig.AddRule(NLog.LogLevel.Debug, NLog.LogLevel.Fatal, consoleTarget); // All levels to console

                        // Apply configuration
                        LogManager.Configuration = nlogConfig;

                        Directory.CreateDirectory(logsDirectory);
                        var configLogger = LogManager.GetCurrentClassLogger();
                        configLogger.Info("Main target configured with filename: {0}", mainTarget.FileName);
                        configLogger.Info("Debug target configured with filename: {0}", debugTarget.FileName);
                        configLogger.Info("Error target configured with filename: {0}", errorTarget.FileName);

                        // Configure services
                        services.AddLogging(builder =>
                        {
                            builder.ClearProviders();
                            builder.SetMinimumLevel(LogLevel.Debug);
                            builder.AddNLog(hostContext.Configuration.GetSection("Logging"));
                        });

                        // Register services
                        services.AddSingleton<MainForm>();
                        services.AddScoped<ILogger<EnaTimesheet>>(sp => sp.GetRequiredService<ILoggerFactory>().CreateLogger<EnaTimesheet>());
                        services.AddScoped<ILogger<EnaTsEntry>>(sp => sp.GetRequiredService<ILoggerFactory>().CreateLogger<EnaTsEntry>());
                        var appLogger = LogManager.GetCurrentClassLogger();
                        appLogger.Info("Application started");

                        // Test log file creation
                        var testLogger = LogManager.GetCurrentClassLogger();
                        testLogger.Info("Testing log file creation");
                        testLogger.Error("Testing error log file creation");
                        testLogger.Debug("Testing debug log file creation");

                        if (!File.Exists(Path.Combine(logsDirectory, $"enatsapp-debug-{DateTime.Now:yyyy-MM-dd}.log")))
                        {
                            throw new Exception($"Debug log file was not created: {Path.Combine(logsDirectory, $"enatsapp-debug-{DateTime.Now:yyyy-MM-dd}.log")}");
                        }

                        if (!File.Exists(Path.Combine(logsDirectory, $"enatsapp-error-{DateTime.Now:yyyy-MM-dd}.log")))
                        {
                            throw new Exception($"Error log file was not created: {Path.Combine(logsDirectory, $"enatsapp-error-{DateTime.Now:yyyy-MM-dd}.log")}");
                        }
                    }
                    catch (Exception ex)
                    {
                        var setupLogger = LogManager.GetCurrentClassLogger();
                        setupLogger.Error(ex, "Failed to configure NLog");
                        throw;
                    }
                });

            return builder;
        }
    }
}

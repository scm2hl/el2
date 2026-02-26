using CompositeCommands.Core;
using ControlzEx.Theming;
using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using Lieferliste_WPF.Dialogs;
using Lieferliste_WPF.Dialogs.ViewModels;
using Lieferliste_WPF.Planning;
using Lieferliste_WPF.ViewModels;
using Lieferliste_WPF.Views;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using El2Core.Dialogs;
using El2Core.Dialogs.ViewModels;
using ModuleDeliverList.Views;
using ModuleMeasuring.Views;
using ModulePlanning.Dialogs;
using ModulePlanning.Dialogs.ViewModels;
using ModulePlanning.Planning;
using ModulePlanning.Views;
using ModuleProducts.Views;
using ModuleReport.Views;
using ModuleShift.Dialogs;
using ModuleShift.Services;
using ModuleShift.Views;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Navigation.Regions;
using Prism.Unity;
using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using Unity;

namespace Lieferliste_WPF
{

    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    internal class Bootstrapper : PrismBootstrapper
    {
        private ILogger? _Logger;

        protected override DependencyObject CreateShell()
        {
            var loggerFactory = Container.Resolve<ILoggerFactory>();
            loggerFactory.AddLog4Net();
            log4net.Config.XmlConfigurator.Configure(new FileInfo("Log4Net.config"));
            _Logger = loggerFactory.CreateLogger<Bootstrapper>();
            Application.Current.DispatcherUnhandledException += Current_DispatcherUnhandledException;
            Application.Current.Exit += Current_Exit;
            _Logger.LogInformation("Shell Create 1");

            var settingsService = Container.Resolve<UserSettingsService>();
            settingsService.Upgrade();

            ThemeManager.Current.ChangeTheme(App.Current, settingsService.Theme);
            App.GlobalFontSize = settingsService.FontSize;
            _Logger.LogInformation("Shell Create 2");
            return Container.Resolve<MainWindow>();
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {           
            _Logger?.LogInformation("Exit: {pc}--{id} Exitcode:{ec}", UserInfo.PC, UserInfo.Dbid, e.ApplicationExitCode);
        }


        private void Current_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            string para = string.Empty;
            var ex = e.Exception as ArgumentNullException;
            if (ex != null) para = ex.ParamName + " Source: " + ex.Source;
            _Logger?.LogCritical("Unhandled exception: {message} Parameter: {p}", e.Exception.ToString(), para);
        }
        protected override void OnInitialized()
        {
            base.OnInitialized();
        }
  
        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            
            //var builder = new ConfigurationBuilder()
            //    .SetBasePath(Directory.GetCurrentDirectory())
            //    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            //IConfiguration configuration = builder.Build();
            //var defaultconnection = configuration.GetConnectionString("ConnectionBosch");
            //var builderopt = new DbContextOptionsBuilder<DB_COS_LIEFERLISTE_SQLContext>()
            //    .UseSqlServer(defaultconnection)
            //    .EnableThreadSafetyChecks(true);

            
            //var builder = new ConfigurationBuilder()
            //    .SetBasePath(Directory.GetCurrentDirectory())
            //    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            //IConfiguration configuration = builder.Build();
            //var defaultconnection = configuration.GetConnectionString("ConnectionHome");
            //var builderopt = new DbContextOptionsBuilder<DB_COS_LIEFERLISTE_SQLContext>()
            //    .UseSqlServer(defaultconnection)
            //    .EnableThreadSafetyChecks(true);
            //containerRegistry.RegisterSingleton<DB_COS_LIEFERLISTE_SQLContext>();
            containerRegistry.Register<DbContext, DB_COS_LIEFERLISTE_SQLContext>();
            containerRegistry.RegisterSingleton<IApplicationCommands, ApplicationCommands>();
            containerRegistry.RegisterSingleton<IHolidayLogic, HolidayLogic>();
            containerRegistry.RegisterSingleton<IProcessStripeService, ProcessStripeService>();
            containerRegistry.RegisterSingleton<IRegionManager, RegionManager>();
            containerRegistry.RegisterSingleton<IUserSettingsService, UserSettingsService>();
            containerRegistry.RegisterSingleton<ILoggerFactory, LoggerFactory>();
            containerRegistry.Register<INotifyBroker, NotifyBroker>();

            containerRegistry.RegisterForNavigation<UserSettings>();
            containerRegistry.RegisterForNavigation<RoleEdit>();
            containerRegistry.RegisterForNavigation<MachinePlan>();
            containerRegistry.RegisterForNavigation<MachineEdit>();
            containerRegistry.RegisterForNavigation<UserEdit>();
            containerRegistry.RegisterForNavigation<Seclusions>();
            containerRegistry.RegisterForNavigation<Liefer>();
            containerRegistry.RegisterForNavigation<ShowWorkArea>();
            containerRegistry.RegisterForNavigation<MeasuringRoom>();
            containerRegistry.RegisterForNavigation<MeasuringDocuments>();
            containerRegistry.RegisterForNavigation<TimeLine>();
            containerRegistry.RegisterForNavigation<HolidayEdit>();
            containerRegistry.RegisterForNavigation<ShiftPlanEdit>();
            containerRegistry.RegisterForNavigation<ReportMainView>();
            containerRegistry.RegisterForNavigation<Products>();
            containerRegistry.RegisterForNavigation<EmployNote>();
            
            containerRegistry.RegisterSingleton<IPlanMachineFactory, PlanMachineFactory>();
            containerRegistry.RegisterSingleton<IPlanWorkerFactory, PlanWorkerFactory>();
            containerRegistry.RegisterDialog<Order>();
            containerRegistry.RegisterDialog<MachineView, MachineViewVM>();
            containerRegistry.RegisterDialog<WorkerView, WorkerViewVM>();
            containerRegistry.RegisterDialog<Projects>();
            containerRegistry.RegisterDialog<AddNewWorkArea, AddNewWorkAreaVM>();
            containerRegistry.RegisterDialog<HistoryDialog, HistoryDialogVM>();
            containerRegistry.RegisterDialog<DocumentDialog, DocumentDialogVM>();
            containerRegistry.RegisterDialog<CorrectionDialog, CorrectionDialogVM>();
            containerRegistry.RegisterDialog<DetailCoverDialog, DetailCoverVM>();
            containerRegistry.RegisterDialog<InputDialog, InputDialogVM>();
            containerRegistry.RegisterDialog<ProjectEdit, ProjectEditViewModel>();
            containerRegistry.RegisterDialog<AttachmentDialog, AttachmentDialogViewModel>();
            containerRegistry.RegisterDialog<InputStoppage, InputStoppageVM>();
            containerRegistry.RegisterDialog<ProcessTimeDialog, ProcessTimeDialogVM>();


            Globals gl = new(Container);
            containerRegistry.RegisterInstance(Globals.CreateUserInfo(Container));
            RuleInfo rule = new(gl.Rules);
            containerRegistry.RegisterInstance(rule);

        }
        protected virtual void ConfigureContainer()
        {

        }
        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            base.ConfigureModuleCatalog(moduleCatalog);

            moduleCatalog.AddModule<ModuleDeliverList.DeliverListModule>();
            moduleCatalog.AddModule<ModuleMeasuring.MeasuringModule>();
            moduleCatalog.AddModule<ModulePlanning.PlanningModule>();
            moduleCatalog.AddModule<ModuleReport.ReportModule>();
            moduleCatalog.AddModule<ModuleProducts.ProductsModule>();
            moduleCatalog.AddModule<ModuleShift.ShiftModule>();
        }
    }
}

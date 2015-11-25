using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Linq;

namespace EnterpriseLibrary.ConfigConsole
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ConfigCmd
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("50013d8f-6901-49c0-8ebf-0e48dfae7356");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigCmd"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ConfigCmd(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
                commandService.AddCommand(menuItem);
            }
        }

        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand oCommand = (OleMenuCommand)sender;
            oCommand.Visible = GetVisible();
        }
        private bool GetVisible()
        {
            DTE2 Application = (DTE2)ServiceProvider.GetService(typeof(DTE));
            return ((Application.SelectedItems.Count == 1)
                            && (Application.SelectedItems.Item(1).ProjectItem != null)
                            && (Application.SelectedItems.Item(1).ProjectItem.Name.Contains(".config")));
        }

        #region 方法
        public static ConfigCmd Instance
        {
            get;
            private set;
        }


        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }


        public static void Initialize(Package package)
        {
            Instance = new ConfigCmd(package);
        }
        #endregion


        private void MenuItemCallback(object sender, EventArgs e)
        {
            DTE2 Application = (DTE2)ServiceProvider.GetService(typeof(DTE));
            EnvDTE.ProjectItem oProjectItem = Application.SelectedItems.Item(1).ProjectItem;
            string configPath = (string)oProjectItem.Properties.Item("FullPath").Value;
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.Arguments = configPath;
            process.StartInfo.FileName = @"C:\Users\SZKL-ZRF\Documents\EntLib50Src\Blocks\bin\Debug\EntLibConfig.NET4.exe";
            process.Start();
        }
        private void GetCurrentProjectConfig()
        {
            DTE2 Application = (DTE2)ServiceProvider.GetService(typeof(DTE));
            EnvDTE.ProjectItem oProjectItem = Application.SelectedItems.Item(1).ProjectItem;
            var activeSolutionProjects = (Array)Application.ActiveSolutionProjects;
            EnvDTE.Project dteProject = (EnvDTE.Project)activeSolutionProjects.GetValue(0);
            string configpath = System.IO.Path.Combine(dteProject.FullName, "App.config");
        }
    }
}

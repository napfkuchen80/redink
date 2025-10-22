using System;
using System.ComponentModel;
using System.Security;
using System.Security.Permissions;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Outlook;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace RedInk.OutlookAddIn
{
    [StartupObject(0)]
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class ThisAddIn : OutlookAddInBase
    {
        internal CustomTaskPaneCollection CustomTaskPanes;
        internal Application Application;

        [EditorBrowsable(EditorBrowsableState.Never)]
        public ThisAddIn(Factory factory, IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            Globals.Factory = factory;
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void Initialize()
        {
            base.Initialize();
            Application = GetHostItem<Application>(typeof(Application), "Application");
            Globals.ThisAddIn = this;
            System.Windows.Forms.Application.EnableVisualStyles();
            InitializeCachedData();
            InitializeControls();
            InitializeComponents();
            InitializeData();
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void FinishInitialization()
        {
            OnStartup();
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void InitializeDataBindings()
        {
            BeginInitialization();
            BindToData();
            EndInitialization();
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeCachedData()
        {
            if (DataHost == null)
            {
                return;
            }

            if (DataHost.IsCacheInitialized)
            {
                DataHost.FillCachedData(this);
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeData()
        {
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void BindToData()
        {
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        private void StartCaching(string memberName)
        {
            DataHost.StartCaching(this, memberName);
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        private void StopCaching(string memberName)
        {
            DataHost.StopCaching(this, memberName);
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        private bool IsCached(string memberName)
        {
            return DataHost.IsCached(this, memberName);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void BeginInitialization()
        {
            BeginInit();
            CustomTaskPanes.BeginInit();
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void EndInitialization()
        {
            CustomTaskPanes.EndInit();
            EndInit();
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeControls()
        {
            CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeComponents()
        {
            InternalStartup();
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        private bool NeedsFill(string memberName)
        {
            return DataHost.NeedsFill(this, memberName);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        protected override void OnShutdown()
        {
            CustomTaskPanes.Dispose();
            base.OnShutdown();
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}

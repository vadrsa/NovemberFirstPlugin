using DevExpress.Xpf.Grid.Native;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Uniconta.ClientTools;
using Uniconta.ClientTools.Controls;
using Uniconta.API.Service;
using Uniconta.Common;
using NovemberFirstPlugin;
using CefSharp.Wpf;
using CefSharp;
using System.Reflection;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace NovemberFirstPage
{
    public class PluginControls : IPluginControl
    {
        private string probingPath = "";
        private string pluginPath = @"C:\Uniconta\PluginPath";

        public PluginControls()
        {
            if (Environment.Is64BitOperatingSystem)
                probingPath = "x64";
            else
                probingPath = "x86";
            
        }

        public string Name
        {
          get { return "PluginDev"; }
        }

        public event EventHandler OnExecute;

        public ErrorCodes Execute(UnicontaBaseEntity master, UnicontaBaseEntity currentRow, IEnumerable<UnicontaBaseEntity> source, string command, string args)
        {
            return ErrorCodes.Succes;
        }

        public string[] GetDependentAssembliesName()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            return null;
        }

        public string GetErrorDescription()
        {
            return null;

        }

        public void Intialize()
        {
            if (!Cef.IsInitialized)
            {
                var settings = new CefSettings();
                settings.BrowserSubprocessPath = Path.Combine(pluginPath, $"{probingPath}/CefSharp.BrowserSubprocess.exe");
                Cef.Initialize(settings, performDependencyCheck: false, browserProcessHandler: null);
            }
        }

        public List<PluginControl> RegisterControls()
        {

            var ctrls = new List<PluginControl>();
            ctrls.Add(new PluginControl() { UniqueName = "NovemberFirstPage", PageType = typeof(NFPage),
                AllowMultipleOpen = false, PageHeader = "November First" });
            return ctrls;
        }

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assyName = new AssemblyName(args.Name);
            var newPath = Path.Combine(probingPath, assyName.Name);
            if (!newPath.EndsWith(".dll"))
            {
                newPath = newPath + ".dll";
            }
            string fullPath = Path.Combine(pluginPath, newPath);
            if (File.Exists(fullPath))
            {
                var assy = Assembly.LoadFile(fullPath);
                return assy;
            }
            return null;
        }

        public void SetAPI(BaseAPI api)
        {

        }

        public void SetMaster(List<UnicontaBaseEntity> masters)
        {
        }
    }

}

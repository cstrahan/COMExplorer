using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Windows.Threading;
using Microsoft.Win32;

namespace COMExplorer.ViewModel
{
    public class TypeLibListViewModel
    {
        public ObservableCollection<TypeLibViewModel> TypeLibs { get; set; }

        public TypeLibListViewModel()
        {
            TypeLibs = new ObservableCollection<TypeLibViewModel>();
            var vms = GetTypeLibs().Select(lib => new TypeLibViewModel(lib));
            //vms.AsParallel().ForAll(vm => Dispatcher.CurrentDispatcher.BeginInvoke((Action)(() => TypeLibs.Add(vm))));
            vms.AsParallel().ForAll(vm => TypeLibs.Add(vm));
        }

        IEnumerable<ITypeLib> GetTypeLibs()
        {
            var typeLibKey = RegistryKey
                .OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default)
                .OpenSubKey("TypeLib");

            var clsids = typeLibKey.GetSubKeyNames();

            foreach (var clsid in clsids)
            {
                Guid guid;
                if (!Guid.TryParse(clsid.Trim(new[] { '{', '}' }), out guid)) continue;

                var clsidKey = typeLibKey.OpenSubKey(clsid);
                var versionStrings = clsidKey.GetSubKeyNames();
                foreach (var versionString in versionStrings)
                {
                    short major, minor;
                    var versionParts = versionString.Split(".".ToCharArray());
                    if (!(versionParts.Length == 2
                        && short.TryParse(versionParts[0], out major)
                        && short.TryParse(versionParts[1], out minor))) continue;

                    var versionKey = clsidKey.OpenSubKey(versionString);
                    var lcids = versionKey.GetSubKeyNames();
                    foreach (var lcidString in lcids)
                    {
                        int lcid;
                        if (!int.TryParse(lcidString, out lcid)) continue;

                        //string libPath = versionKey.OpenSubKey(lcidString + @"\win32").GetValue("").ToString();

                        ITypeLib typeLib = null;
                        try
                        {
                             typeLib = COMUtil.LoadRegTypeLib(ref guid, major, minor, lcid);
                        }
                        catch (COMException)
                        {
                        }

                        if (typeLib != null) yield return typeLib;
                    }
                }
            }
        }
    }
}
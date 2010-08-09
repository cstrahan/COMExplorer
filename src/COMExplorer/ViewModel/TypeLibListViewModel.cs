using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Windows.Data;
using System.Windows.Threading;
using Microsoft.Win32;

namespace COMExplorer.ViewModel
{
    public class TypeLibListViewModel : ViewModelBase
    {
        private ObservableCollection<TypeLibViewModel> _typeLibs;
        public ObservableCollection<TypeLibViewModel> TypeLibs
        {
            get
            {
                return _typeLibs;
            }
            set
            {
                _typeLibs = value;
                RaiseChanged(() => TypeLibs);
            }
        }
        private Dispatcher _dispatcher = Dispatcher.CurrentDispatcher;

        public TypeLibListViewModel()
        {
            var thread = new Thread(() =>
                                        {
                                            var viewModels = new ObservableCollection<TypeLibViewModel>(GetTypeLibs().AsParallel().Select(lib => new TypeLibViewModel(lib)));
                                            _dispatcher.BeginInvoke((Action)(() => TypeLibs = viewModels));
                                        });
            thread.Start();
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
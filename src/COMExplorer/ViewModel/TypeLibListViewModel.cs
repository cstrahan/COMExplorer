using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Win32;

namespace COMExplorer.ViewModel
{
    public class TypeLibListViewModel : ViewModelBase
    {
        private readonly char[] _curlies = new[] {'{', '}'};
        private readonly Dispatcher _dispatcher = Application.Current.Dispatcher;
        private readonly char[] _period = new[] {'.'};
        private bool _isLoading;
        private TypeLibViewModel _selectedTypeLib;
        private ObservableCollection<TypeLibViewModel> _typeLibs;

        public TypeLibListViewModel()
        {
            LoadAsync();
        }

        public ObservableCollection<TypeLibViewModel> TypeLibs
        {
            get { return _typeLibs; }
            set
            {
                _typeLibs = value;
                RaiseChanged(() => TypeLibs);
            }
        }

        public TypeLibViewModel SelectedTypeLib
        {
            get { return _selectedTypeLib; }
            set
            {
                _selectedTypeLib = value;
                RaiseChanged(() => SelectedTypeLib);
            }
        }

        public bool IsLoading
        {
            get { return _isLoading; }
            set
            {
                _isLoading = value;
                RaiseChanged(() => IsLoading);
            }
        }

        public void LoadAsync()
        {
            new Thread(()=>
                                {
                                    var viewModels = new ObservableCollection<TypeLibViewModel>(
                                        GetTypeLibs().AsParallel().Select(lib => new TypeLibViewModel(lib))
                                            .OrderBy(vm => vm.Description)
                                            .ThenBy(vm => vm.Name));
                                    _dispatcher.BeginInvoke((Action) (() =>
                                                                          {
                                                                              TypeLibs = viewModels;
                                                                              IsLoading = false;
                                                                          }));
                                }).Start();
        }

        private IEnumerable<ITypeLib> GetTypeLibs()
        {
            RegistryKey typeLibKey = RegistryKey
                .OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default)
                .OpenSubKey("TypeLib");

            string[] clsids = typeLibKey.GetSubKeyNames();

            foreach (string clsid in clsids)
            {
                Guid guid;
                if (!Guid.TryParse(clsid.Trim(_curlies), out guid)) continue;

                RegistryKey clsidKey = typeLibKey.OpenSubKey(clsid);
                string[] versionStrings = clsidKey.GetSubKeyNames();
                foreach (string versionString in versionStrings)
                {
                    short major, minor;
                    string[] versionParts = versionString.Split(_period);
                    if (!(versionParts.Length == 2
                          && short.TryParse(versionParts[0], out major)
                          && short.TryParse(versionParts[1], out minor))) continue;

                    RegistryKey versionKey = clsidKey.OpenSubKey(versionString);
                    string[] lcids = versionKey.GetSubKeyNames();
                    foreach (string lcidString in lcids)
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
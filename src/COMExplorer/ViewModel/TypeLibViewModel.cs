using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;

namespace COMExplorer.ViewModel
{
    public class TypeLibViewModel : ViewModelBase
    {
        private readonly Dispatcher _dispatcher = Application.Current.Dispatcher;
        private readonly ITypeLib _typeLib;
        private ObservableCollection<TypeViewModel> _types;

        public TypeLibViewModel(ITypeLib typeLib)
        {
            _typeLib = typeLib;
            TYPELIBATTR attr = COMUtil.GetTypeLibAttr(typeLib);
            CLSID = attr.guid.ToString();
            MajorVersion = attr.wMajorVerNum;
            MinorVersion = attr.wMinorVerNum;
            LCID = attr.lcid;
            //Name = Marshal.GetTypeLibName(typeLib);
            string name, docString, helpFile;
            int helpContext;
            typeLib.GetDocumentation(-1,
                                     out name,
                                     out docString,
                                     out helpContext,
                                     out helpFile);

            Path = COMUtil.GetTypeLibPath(typeLib);
            Name = name;
            Description = docString;
            // Remove the null char
            HelpFilePath = helpFile == null ? string.Empty : helpFile.Substring(0, helpFile.Length - 1);

            string asmName, asmCodeBase;
            var converter = new TypeLibConverter();
            converter.GetPrimaryInteropAssembly(
                attr.guid, MajorVersion, MinorVersion, LCID, out asmName, out asmCodeBase);

            PIAName = asmName;
            PIACodeBase = asmCodeBase;
        }

        public string CLSID { get; set; }
        public int MajorVersion { get; set; }
        public int MinorVersion { get; set; }
        public int LCID { get; set; }
        public string PIACodeBase { get; set; }
        public string PIAName { get; set; }

        public string Version
        {
            get { return MajorVersion + "." + MinorVersion; }
        }

        public string Name { get; set; }
        public string Description { get; set; }
        public string Path { get; set; }
        public string HelpFilePath { get; set; }
        private bool _isLoadingTypes;
        public bool IsLoadingTypes
        {
            get { return _isLoadingTypes; }
            set
            {
                _isLoadingTypes = value;
                RaiseChanged(() => IsLoadingTypes);
            }
        }

        private TypeViewModel _selectedType;
        public TypeViewModel SelectedType
        {
            get { return _selectedType; }
            set
            {
                _selectedType = value;
                RaiseChanged(() => SelectedType);
            }
        }

        public ObservableCollection<TypeViewModel> Types
        {
            get
            {
                if (_types == null)
                {
                    IsLoadingTypes = true;
                    _types = new ObservableCollection<TypeViewModel>();
                    LoadTypesAsync();
                }

                return _types;
            }
        }

        public void LoadTypesAsync()
        {
            Parallel.Invoke(() =>
                                {
                                    var tempTypes = new List<TypeViewModel>();
                                    int infoCount = _typeLib.GetTypeInfoCount();
                                    for (int i = 0; i < infoCount; i++)
                                    {
                                        try
                                        {
                                            var typeViewModel = new TypeViewModel(_typeLib, i);
                                            tempTypes.Add(typeViewModel);
                                        }
                                        catch (COMException)
                                        {
                                        }
                                    }

                                    _dispatcher.BeginInvoke((Action)(() =>
                                                                         {
                                                                             foreach (TypeViewModel types in tempTypes.OrderBy(t => t.Name))
                                                                             {

                                                                                 _types.Add(types);
                                                                             }
                                                                             RaiseChanged(() => Types);
                                                                             IsLoadingTypes = false;
                                                                         }));
                                });
        }
    }
}
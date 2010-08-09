using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;

namespace COMExplorer.ViewModel
{
    public class TypeLibViewModel
    {
        private readonly ITypeLib _typeLib;
        public string CLSID { get; set; }
        public int MajorVersion { get; set; }
        public int MinorVersion { get; set; }
        public string Version
        {
            get { return MajorVersion + "." + MinorVersion; }
        }

        public string Name { get; set; }
        public string Description { get; set; }
        public string Path { get; set; }
        public string HelpFilePath { get; set; }

        private IList<TypeViewModel> _types;
        public IList<TypeViewModel> Types
        {
            get
            {
                if (_types == null)
                {
                    _types = new List<TypeViewModel>();

                    var infoCount = _typeLib.GetTypeInfoCount();
                    for (int i = 0; i < infoCount; i++)
                    {
                        try
                        {
                            var typeViewModel = new TypeViewModel(_typeLib, i);
                            _types.Add(typeViewModel);
                        }
                        catch (COMException)
                        {
                        }
                    }
                }

                return _types;
            }
        }

        public TypeLibViewModel(ITypeLib typeLib)
        {
            _typeLib = typeLib;
            var attr = COMUtil.GetTypeLibAttr(typeLib);
            CLSID = attr.guid.ToString();
            MajorVersion = attr.wMajorVerNum;
            MinorVersion = attr.wMinorVerNum;
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
        }
    }
}
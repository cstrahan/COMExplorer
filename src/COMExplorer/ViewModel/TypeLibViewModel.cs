using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace COMExplorer.ViewModel
{
    public class TypeLibViewModel
    {
        public string CLSID { get; set; }
        public int MajorVersion { get; set; }
        public int MinorVersion { get; set; }
        public string Version
        {
            get { return MajorVersion + "." + MinorVersion; }
        }

        public string Name { get; set; }
        public string Description { get; set; }
        public string HelpFilePath { get; set; }

        public TypeLibViewModel(ITypeLib typeLib)
        {
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

            Name = name;
            Description = docString;
            HelpFilePath = helpFile;
        }
    }
}
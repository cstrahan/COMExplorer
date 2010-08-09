using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;

namespace COMExplorer
{
    public class COMUtil
    {
        [DllImport("ole32.dll")]
        public static extern int ProgIDFromCLSID([In]ref Guid clsid, [MarshalAs(UnmanagedType.LPWStr)]out string lplpszProgID);

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        public static extern void LoadTypeLibEx(string strTypeLibName,
            REGKIND regKind,
            out ITypeLib TypeLib);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern ITypeLib LoadRegTypeLib(ref Guid clsid, short majorVersion, short minorVersion, int lcid);


        public static TYPELIBATTR GetTypeLibAttr(ITypeLib typeLib)
        {
            IntPtr ppTLibAttr;
            typeLib.GetLibAttr(out ppTLibAttr);
            TYPELIBATTR tlibattr = (TYPELIBATTR)Marshal.PtrToStructure(ppTLibAttr, typeof(TYPELIBATTR));
            typeLib.ReleaseTLibAttr(ppTLibAttr);

            return tlibattr;
        }

        public static TYPEATTR GetTypeAttr(ITypeInfo typeInfo)
        {
            IntPtr ppAttr;
            typeInfo.GetTypeAttr(out ppAttr);
            TYPEATTR typeAttr = (TYPEATTR)Marshal.PtrToStructure(ppAttr, typeof(TYPEATTR));
            typeInfo.ReleaseTypeAttr(ppAttr);

            return typeAttr;
        }

        public static ITypeInfo[] GetTypeInfos(ITypeLib typeLib)
        {
            var infoCount = typeLib.GetTypeInfoCount();
            var infos = new ITypeInfo[infoCount];

            for (int i = 0; i < infoCount; i++)
            {
                ITypeInfo info;
                typeLib.GetTypeInfo(i, out info);
                infos[i] = info;
            }

            return infos;
        }

        private static string ProgIdFromTypeAttr(TYPEATTR typeattr)
        {
            string lplpszProgID;
            ProgIDFromCLSID(ref typeattr.guid, out lplpszProgID);
            return lplpszProgID;
        }

        public static IList<string> GetProgIds(ITypeLib typeLib)
        {
            var typeInfos = GetTypeInfos(typeLib);
            return typeInfos
                .Select(GetTypeAttr)
                .Where(attr => attr.typekind == TYPEKIND.TKIND_COCLASS)
                .Select(ProgIdFromTypeAttr)
                .ToList();
        }

        public static string GetTypeLibPath(ITypeLib typeLib)
        {
            var attr = GetTypeLibAttr(typeLib);
            return GetTypeLibPath(attr.guid, attr.wMajorVerNum, attr.wMinorVerNum, attr.lcid);
        }

        public static string GetTypeLibPath(Guid g, int major, int minor, int lcid)
        {
            string path = null;
            var key = RegistryKey
                .OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default)
                .OpenSubKey(@"TypeLib\{" + g + @"}\" + major + "." + minor + @"\" + lcid + @"\win32");
            if (key != null) path = (string)key.GetValue("");
            return path;
        }
    }
}

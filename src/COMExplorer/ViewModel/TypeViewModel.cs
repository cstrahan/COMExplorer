using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace COMExplorer.ViewModel
{
    public class TypeViewModel : ViewModelBase
    {
        public string Name { get; set; }
        public TYPEKIND Kind { get; set; }

        private List<MemberViewModel> _members;
        public List<MemberViewModel> Members
        {
            get { return _members; }
            set { _members = value; }
        }

        public TypeViewModel(ITypeLib typeLib, int index)
        {
            Members = new List<MemberViewModel>();
            ITypeInfo typeInfo;
            typeLib.GetTypeInfo(index, out typeInfo);

            string strName, strDocString, strHelpFile;
            int dwHelpContext;
            typeLib.GetDocumentation(index, out strName, out strDocString, out dwHelpContext, out strHelpFile);

            Name = strName;

            // This sometimes throws a COMException
            var attr = COMUtil.GetTypeAttr(typeInfo);

            Kind = attr.typekind;

            // Get the functions
            for (int i = 0; i < attr.cFuncs; i++)
            {
                // TODO: Figure out why this throws an AccessViolationException for the groove type library
                IntPtr ppFuncDesc;
                typeInfo.GetFuncDesc(i, out ppFuncDesc);

                FUNCDESC funcdesc = (FUNCDESC)Marshal.PtrToStructure(ppFuncDesc, typeof(FUNCDESC));
                typeInfo.ReleaseFuncDesc(ppFuncDesc);

                typeInfo.GetDocumentation(funcdesc.memid, out strName, out strDocString, out dwHelpContext, out strHelpFile);
                var member = new MemberViewModel(strName);
                Members.Add(member);
            }

            // Get the vars
            for (int i = 0; i < attr.cVars; i++)
            {
                IntPtr ppVarDesc;
                typeInfo.GetVarDesc(i, out ppVarDesc);

                VARDESC vardesc = (VARDESC)Marshal.PtrToStructure(ppVarDesc, typeof(VARDESC));
                typeInfo.ReleaseVarDesc(ppVarDesc);

                typeInfo.GetDocumentation(vardesc.memid, out strName, out strDocString, out dwHelpContext, out strHelpFile);
                var member = new MemberViewModel(strName);
                Members.Add(member);
            }
        }
    }
}
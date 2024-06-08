using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VbProjectParser.Compression;

namespace VbProjectParser.Data._PROJECTINFORMATION
{
    /// <summary>
    /// Page 31
    /// </summary>
    public class PROJECTINFORMATION : DataBase
    {
        public readonly PROJECTSYSKIND SysKindRecord;
        public readonly PROJECTLCID LcidRecord;
        public readonly PROJECTLCIDINVOKE LcidInvokeRecord;
        public readonly PROJECTCODEPAGE CodePageRecord;
        public readonly PROJECTNAME NameRecord;
        public readonly PROJECTDOCSTRING DocStringRecord;
        public readonly PROJECTHELPFILEPATH HelpFilePathRecord;
        public readonly PROJECTHELPCONTEXT HelpContextRecord;
        public readonly PROJECTLIBFLAGS LibFlagsRecord;
        public readonly PROJECTVERSION VersionRecord;
        public readonly PROJECTCONSTANTS ConstantsRecord;

        public PROJECTINFORMATION(XlBinaryReader Data)
        {
            this.SysKindRecord = new PROJECTSYSKIND(Data);

            // Todo: Rogerg updated to handle class.

            // Code assumes records are written to the project information in 
            // same order every time.
            // Newer versions of Office are puttin a PROJECTCOMPATVERSION record
            // after the SysKindRecord
            // Check if the next record is a PROJECTCOMPATVERSION and if so skip.

            int nextRecordType = Data.PeekUInt16();
            if (nextRecordType == 0x004A)
            {
                // read in record but currently don't use
                UInt16 recordId =  Data.ReadUInt16();
                UInt32 compatRecordSize = Data.ReadUInt32();
                UInt32 compatVersion = Data.ReadUInt32();
            }

            this.LcidRecord = new PROJECTLCID(Data);
            this.LcidInvokeRecord = new PROJECTLCIDINVOKE(Data);
            this.CodePageRecord = new PROJECTCODEPAGE(Data);
            this.NameRecord = new PROJECTNAME(this,Data);
            this.DocStringRecord = new PROJECTDOCSTRING(this, Data);
            this.HelpFilePathRecord = new PROJECTHELPFILEPATH(this, Data);
            this.HelpContextRecord = new PROJECTHELPCONTEXT(Data);
            this.LibFlagsRecord = new PROJECTLIBFLAGS(Data);
            this.VersionRecord = new PROJECTVERSION(Data);
            this.ConstantsRecord = new PROJECTCONSTANTS(this, Data);

            Validate();
        }
    }
}

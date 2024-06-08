using AbnfFramework;
using AbnfFramework.Tokens;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbProjectParser.Data.ABNF
{
    // p 26
    public class HostExtenders
    {
        public IList<HostExtenderRef> HostExtenderRef { get; set; }

        public HostExtenders()
        {
            this.HostExtenderRef = new List<HostExtenderRef>();
        }

        public static void Setup(ISyntax Syntax)
        {
            // Todo: rogerg (HostExtenderInfo).
            //   Original code would check for a PreFix of [Host Extender Info]. This causes
            //   only the first entry to be added since the second entry wouldn't have a Prefix
            //   of [Host Extender Info].
            //   Also updated the VBAProjectText.cs to have [Host Extender Info] for the match on HostExtenders.
            Syntax
                .Entity<HostExtenders>()
                .EnumerableProperty(x => x.HostExtenderRef)
                .ByRegisteredTypes(typeof(HostExtenderRef));
        }
    }
}

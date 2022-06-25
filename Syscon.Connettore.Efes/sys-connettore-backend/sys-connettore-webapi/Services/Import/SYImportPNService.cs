using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using sys_connettore_database.Model;

namespace sys_connettore_webapi.Services.Import
{
    
    public interface ISYImportPNService
    {
        List<GOCG_PRIMANOTA> GetPrimanota();

        public class SYImportPNService : ISYImportPNService
        {
            public List<GOCG_PRIMANOTA> GetPrimanota()
            {
                List<GOCG_PRIMANOTA> pn = new List<GOCG_PRIMANOTA>();

                using (var ctx = new AlyEfesEntities())
                {
                    pn = ctx.GOCG_PRIMANOTA.ToList();
                }

                return pn;
            }
        }

    }

}

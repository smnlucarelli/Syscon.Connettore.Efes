using Microsoft.AspNetCore.Mvc;
using sys_connettore_database.Model;
using sys_connettore_webapi.Services;
using sys_connettore_webapi.Services.Import;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace sys_connettore_webapi.Controllers.Import
{
    [ApiController]
    [Route("api/v1/ImportPN")]

    public class SYImportPNController: ControllerBase
    {
        private readonly ISYImportPNService _sys;

        public SYImportPNController(ISYImportPNService sys)
        {
            _sys = sys;
        }

        [HttpGet]
        public List<GOCG_PRIMANOTA> GetPrimanota() => _sys.GetPrimanota();
    }
}

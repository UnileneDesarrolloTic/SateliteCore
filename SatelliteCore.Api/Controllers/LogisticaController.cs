﻿using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class LogisticaController : ControllerBase
    {
        private readonly ILogisticaServices _logisticaServices;

        public LogisticaController(ILogisticaServices logisticaServices)
        {
            _logisticaServices = logisticaServices;
        }

    }
}

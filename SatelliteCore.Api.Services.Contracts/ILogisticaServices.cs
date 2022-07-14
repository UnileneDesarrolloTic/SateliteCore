﻿using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILogisticaServices
    {

        public bool RegistrarMaestroItem(DatosRequestMaestroItemModel dato);
    }
}

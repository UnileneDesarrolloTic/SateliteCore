using SatelliteCore.Api.Models.Request;
using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.DataAccess.Contracts
{
    public interface ILogisticaRepository
    {
        public bool RegistrarMaestroItem(DatosRequestMaestroItemModel dato);
    }
}

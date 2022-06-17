using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Entities
{
    public class ListarProcesoEntity
    {
        public int IdProceso { get; set; }
        public string DescripcionProceso { get; set; }
        public string DescripcionComercial { get; set; }
        public string DescripcionComercialDetalle { get; set; }
        public int CantItems { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoListarOrdenFabricacionModel
    {
       public List<FormatoEstructuraObtenerOrdenFabricacion> InformacionLote { get; set; }
       public List<FormatoEstructuraDetalleOrdenFabricacionkardexInterno> Detalle { get; set; }
    }
}

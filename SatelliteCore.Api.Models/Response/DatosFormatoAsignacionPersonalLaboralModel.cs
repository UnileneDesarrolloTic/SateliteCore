using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoAsignacionPersonalLaboralModel
    {
        public List<FormatoDeAsignacionPersonalLaboralModel> PersonalLaboral { get; set; }
        public List<DatosFormatoContarAreaModel> ContarArea { get; set; }
    }
}

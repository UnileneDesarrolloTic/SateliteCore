using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.GestionGuias
{
    public struct DatosFormatoGestionGuiasClienteModel
    {   
        public string idCliente { get; set; }
        public string cliente { get; set; }
        public string destino { get; set; }
        public int transportista { get; set; }
        public DateTime? fechaInicio { get; set; }
        public DateTime? fechaFin { get; set; }
        public Boolean activarFecha { get; set; }
        public string exportar { get; set; }
    }
}

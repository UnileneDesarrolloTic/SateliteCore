using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoRegistroPruebasAgujasModel
    {
        [Required]
        public string Especialidad { get; set; }

        [Required]
        public string Lote { get; set; }

        public DateTime FechaAnalisis { get; set; }
        public List<GuardarPruebaFlexionAgujaModel> Detalle { get; set; }
    }
}

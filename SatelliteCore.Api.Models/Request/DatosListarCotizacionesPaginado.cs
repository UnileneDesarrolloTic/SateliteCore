﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public struct DatosListarCotizacionesPaginado
    {
        public string NumeroDocumento { get; set; }
        public string ClienteNombre { get; set; }
        [Required]
        public int Pagina { get; set; }
        [Required]
        public int RegistrosPorPagina { get; set; }
    }
}

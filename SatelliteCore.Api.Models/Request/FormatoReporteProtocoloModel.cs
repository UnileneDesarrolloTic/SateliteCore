﻿using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public struct FormatoReporteProtocoloModel
    {
        public int Secuencia { get; set; }
        public string ItemNumeroParte { get; set; }
        public string NumeroLote { get; set; }
        public string DescripcionLocal { get; set; }
        public string Presentacion { get; set; }
        public decimal TamanoLote { get; set; }
        public string Marca { get; set; }
        public DateTime FechaFabricacion { get; set; }
        public DateTime FechaExpira { get; set; }
        public DateTime FechaAnalisis { get; set; }
        public string Prueba { get; set; }
        public string Especificacion { get; set; }
        public string Resultado { get; set; }
        public string Metodologia { get; set; }
        public string Leyenda { get; set; }
        public string MetodosEsterilizacion { get; set; }
        public string NormasISO { get; set; }
        public string OrdenFabricacion { get; set; }
        public string Caracteristicas { get; set; }
        public string CaracteristicasEmpaque { get; set; }
        public string ParametroPiePaginaSGC { get; set; }

    }
}

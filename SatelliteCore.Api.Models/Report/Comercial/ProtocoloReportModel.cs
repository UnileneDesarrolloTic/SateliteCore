using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Comercial
{
    public class ProtocoloCabeceraModel
    {
        public string OrdenFabricacion { get; set; }
        public string NumeroParte { get; set; }
        public string ItemDescripcion { get; set; }
        public string Presentacion { get; set; }
        public DateTime FechaFabricacion { get; set; }
        public string Lote { get; set; }
        public decimal TamanoLote { get; set; }
        public DateTime FechaExpiracion { get; set; }
        public string Marca { get; set; }
        public DateTime FechaAnalisis { get; set; }
        public string Leyenda { get; set; }
        public string MetodosEsterilizacion { get; set; }
        public string NormasISO { get; set; }
    }

    public class ProtocoloDetalleModel
    {
        public string Prueba { get; set; }
        public string Especificacion { get; set; }
        public string Resultado { get; set; }
        public string Metodologia { get; set; }
        public string OrdenFabricacion { get; set; }
    }

    public class ProtocoloReportModel
    {
        public ProtocoloCabeceraModel Cabecera { get; set; }

        public List<ProtocoloDetalleModel> Detalle { get; set; }
    }
}

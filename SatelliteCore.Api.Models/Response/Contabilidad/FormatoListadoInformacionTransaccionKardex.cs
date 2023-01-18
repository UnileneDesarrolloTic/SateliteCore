using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.Contabilidad
{
    public class FormatoListadoInformacionTransaccionKardex
    {
        public string TipoDocumento { get; set; }
        public string NumeroDocumento { get; set; }
        public string TransaccionCodigo { get; set; }
        public string ReferenciaTipoDocumento { get; set; }
        public string ReferenciaNumeroDocumento { get; set; }
        public int Secuencia { get; set; }
        public string Item { get; set; }
        public string Lote { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal MontoTotal { get; set; }
        public decimal MontoTotalReal { get; set; }
        public string CodigoUnico { get; set; }
    }
}

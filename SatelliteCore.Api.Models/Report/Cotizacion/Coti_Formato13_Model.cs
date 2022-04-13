using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato13_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Celular { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_VigProducto { get; set; }
        public string Prov_VigOferta { get; set; }
        public decimal Prov_ValorTotal { get; set; }
        public DateTime Fecha_Cotizacion { get; set; }
        public List<Coti_Formato13_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato13_Detalle {

        public decimal NroItem { get; set; }
        public string CodigoSap { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string Um { get; set; }
        public string PlazoEntrega { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public string MarcaModelo { get; set; }
        public string CumpleEETT { get; set; }
        public string VigOferta { get; set; }
        public string Rnp { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }

    }
}

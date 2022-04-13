using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato62_Model : CotizacionAbstract
    {
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_NroCotizacion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Fax { get; set; }
        public string Vig_Oferta { get; set; }
        public DateTime Fecha_Cotizacion { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Cargo { get; set; }
        public string Prov_Telefono2 { get; set; }
        public string Prov_Celular { get; set; }
        public string Prov_DatosAdicionales { get; set; }
        public string Prov_GarantiaMinima { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Vigcotizacion { get; set; }
        public List<Coti_Formato62_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato62_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodigoSap { get; set; }
        public string Denominacion { get; set; }
        public decimal Cantidad { get; set; }
        public string UndMedida { get; set; }
        public string Modelo { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public string Voferta { get; set; }
        public string CumpDemoninacion { get; set; }
        public string VigProducto { get; set; }
        public string CapAtencion { get; set; }
        public string PlazoEntrega { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal ValorTotal { get; set; }

    }
}

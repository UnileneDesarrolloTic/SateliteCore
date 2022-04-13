using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato21_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_PlazoEntrega { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Garantia { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Celular { get; set; }
        public string Prov_VigOferta { get; set; }
        public DateTime Fecha_Cotizacion { get; set; }
        public string Prov_VigProducto { get; set; }
        public decimal Prov_ValorTotal { get; set; }
        public List<Coti_Formato21_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato21_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodigoSap { get; set; }
        public string Descripcion { get; set; }
        public string UndMedida { get; set; }
        public decimal Cantidad { get; set; }
        public string Presentacion { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string ValidezOferta { get; set; }
        public string VigMinima { get; set; }
        public string CumpDescripcion { get; set; }
        public string AutSanitaria { get; set; }
        public string BuenasPracticas { get; set; }
        public string CertSanitario { get; set; }
        public string NrRegSanitario { get; set; }
        public string CertManufactura { get; set; }
        public string CertProductoTer { get; set; }
        public string MetodoAnalisis { get; set; }
        public string LogSolicitado { get; set; }
        public string PlazoEntrega { get; set; }
        public string CapacidadAtencion { get; set; }

    }
}
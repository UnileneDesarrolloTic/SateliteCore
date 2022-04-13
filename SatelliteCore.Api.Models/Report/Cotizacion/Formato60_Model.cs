using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato60_Model: CotizacionAbstract
    {
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_NroCotizacion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Fax { get; set; }
        public string Prov_Telefono2 { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public string Prov_DatosAdicionales { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Cargo { get; set; }
        public string Prov_Telefono3 { get; set; }
        public string Prov_Email2 { get; set; }
        public string Soli_AreaUsuaria { get; set; }
        public string Soli_Dependencia { get; set; }
        public string Soli_Contacto { get; set; }
        public string Soli_Email { get; set; }
        public string Soli_Telefono { get; set; }
        public List<Formato60_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato60_Detalle
    {
        public int NroItem { get; set; }
        public string CodigoSap { get; set; }
        public string Denominacion { get; set; }
        public string UnidadMedida { get; set; }
        public decimal CantidadSolicitada { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public string Vencimiento { get; set; }
        public string PlazoEntrega { get; set; }
        public string Vigencia { get; set; }
        public string FormaPago { get; set; }
        public string RNP { get; set; }
        public string Garantia { get; set; }
        public string EspTecnica { get; set; }
        public decimal CostoUnitario { get; set; }
        public decimal CostoTotal { get; set; }
    }
}

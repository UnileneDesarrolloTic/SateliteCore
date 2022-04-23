using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato_30_Model : CotizacionAbstract
    {

        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Email1 { get; set; }
        public string Prov_Telefono1 { get; set; }
        public string Prov_fax { get; set; }
        public string Prov_vigOferta { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Cargo { get; set; }
        public string Prov_Telefono2 { get; set; }
        public string Prov_Email2 { get; set; }
        public string Prov_Adicional { get; set; }
        public string Prov_formaPago { get; set; }
        public string Soli_Area { get; set; }
        public string Soli_Dependecia { get; set; }
        public string Soli_Contacto { get; set; }
        public string Soli_Email { get; set; }
        public string Soli_Telefono { get; set; }

        public List<Coti_Formato30_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato30_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodigoSAP { get; set; }
        public string Denominacion { get; set; }
        public string UndMedida { get; set; }
        public decimal Cantidad { get; set; }
        public string Marca { get; set; }
        public string PaisProcedencia { get; set; }
        public string FormaPresentacion { get; set; }
        public string Modelo { get; set; }
        public string Rnp { get; set; }
        public string CumpDenominacion { get; set; }
        public string VigMinimaMaterial { get; set; }
        public string RegSanitario { get; set; }
        public string EspTecnicas { get; set; }
        public string CbPracticasAlma { get; set; }
        public string CapAtencion { get; set; }
        public string CbPracticasManufacturas { get; set; }
        public string PlazoEntrega { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }

    }
}

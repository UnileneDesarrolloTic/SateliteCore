using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato24_Model  : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_VigOferta { get; set; }
        public string Prov_PlazoEntrega { get; set; }
        public string Prov_Garantia { get; set; }
        public DateTime Prov_FechaOferta { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_ContCargo { get; set; }
        public string Prov_ContCelular { get; set; }
        public string Prov_ContEmail { get; set; }
        public decimal Prov_ValorTotal { get; set; }

        public List<Coti_Formato24_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato24_Detalle
    {
        public decimal NroItem { get; set; }
        public string Solped { get; set; }
        public string CodigoSap { get; set; }
        public string Producto { get; set; }
        public string Um { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public string PlazoEntrega { get; set; }
        public string VigMaterial { get; set; }
        public string CapAtencion { get; set; }
        public string CumpPlazo { get; set; }
        public string CBuenasPracticas { get; set; }
        public string RegSanitario { get; set; }
        public string CertBPracticas { get; set; }
        public string CumpNormas { get; set; }

    }
}

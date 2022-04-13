using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato63_Model : CotizacionAbstract
    {
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_VigOferta { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_ContCargo { get; set; }
        public string Prov_ContCelular { get; set; }
        public string Prov_ContEmail { get; set; }
        public string Prov_RpmRrc { get; set; }

        public List<Coti_Formato63_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato63_Detalle
    {
        public decimal NroItem { get; set; }
        public string Denominacion { get; set; }
        public string Um { get; set; }
        public string ReqTotal { get; set; }
        public string PaisProcedencia { get; set; }
        public string Presentacion { get; set; }
        public string VigMaterial { get; set; }
        public string AutoSanitaria { get; set; }
        public string CumpTerminos { get; set; }
        public string RegSanitario { get; set; }
        public string CertAnalisis { get; set; }
        public string BuenasPracticas { get; set; }
        public string Rotulado { get; set; }
        public string PlazoEntrega { get; set; }
        public string MarcaModelo { get; set; }
        public string Procedencia { get; set; }
        public string FormaPago { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal ValorTotal { get; set; }

    }
}

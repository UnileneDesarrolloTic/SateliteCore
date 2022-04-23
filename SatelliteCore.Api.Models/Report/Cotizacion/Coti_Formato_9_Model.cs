using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato_9_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_telefono { get; set; }
        public string Prov_movil { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Resposanble { get; set; }
        public string Prov_AreaRequiriente { get; set; }
        public string Prov_nit { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public decimal Prov_valorTotal { get; set; }

        public List<Coti_Formato9_Detalle> Detalle { get; set; }

    }


    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato9_Detalle
    {

        public decimal NroItem { get; set; }
        public string Solpe { get; set; }
        public string Pos { get; set; }
        public string CodigoSAP { get; set; }
        public string Descripcion { get; set; }
        public string Um { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }
        public string Marca { get; set; }
        public string Laboratorio { get; set; }
        public string Presentacion { get; set; }
        public string Procedencia { get; set; }
        public string PlazoEntrega { get; set; }
        public string ResolAutSanitaria { get; set; }
        public string CertBPA { get; set; }
        public string RegSanitario { get; set; }
        public string CertBPM { get; set; }
        public string ProtocoloAnalisis { get; set; }
        public string VigMinimaProducto { get; set; }
        public string EspTecnicas { get; set; }
        public string RegNacionalProveedores { get; set; }
        public string Observaciones { get; set; }

    }

}

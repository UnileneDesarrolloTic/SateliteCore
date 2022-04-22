using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato27_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_PlazoEntrega { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Garantia { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_movil { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_vigOferta { get; set; }
        public string Prog_RegiSanitario { get; set; }
        public decimal Prov_ValorTotal { get; set; }
        public DateTime Prov_Fecha { get; set; }

        public List<Coti_Formato27_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato27_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodiMaterial { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string UnidMedida { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public string Paisprocedencia { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }


    }
}

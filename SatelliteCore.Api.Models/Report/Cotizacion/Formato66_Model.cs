using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato66_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_movil { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_vigProducto { get; set; }
        public string Prov_vigOferta { get; set; }
        public decimal Prov_ValorTotal { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public List<Coti_Formato66_Detalle> Detalle { get; set; }


    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato66_Detalle
    {

        public decimal NroItem { get; set; }
        public string CodigoSAP { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string Um { get; set; }
        public string PlazoEntrega { get; set; }
        public string Procedencia { get; set; }
        public string Marcamodelo { get; set; }
        public string CumpEETT { get; set; }
        public string VigOferta { get; set; }
        public string Rnp { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }

    }

}

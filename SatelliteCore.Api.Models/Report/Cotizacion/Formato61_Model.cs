using MongoDB.Bson.Serialization.Attributes;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato61_Model : CotizacionAbstract
    {
        public string Nro_Cotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_PlazaEntrega { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Correo { get; set; }
        public string Prov_Fijo { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Movil { get; set; }
        public string Prov_Vigencia { get; set; }
        public decimal Monto_Total { get; set; }

        public List<Formato61_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato61_Detalle
    {
        public int NroItem { get; set; }
        public string Sap { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string Um { get; set; }
        public string Vigencia { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public string Procedencia { get; set; }
        public string Observaciones { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal PrecioTotal { get; set; }

    }
}

using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;


namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato4_Model : CotizacionAbstract
    {
        //DATOS DEL PROVEEDOR 
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Pro_Contacto { get; set; }
        public string Prov_Email { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public string Prov_Celular { get; set; }

        public List<Coti_Formato4_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato4_Detalle : CotizacionAbstract
    {
        public int NroItem { get; set; }
        public string Denominacion { get; set; }
        public string EspTecnica { get; set; }
        public string IndicadorObs { get; set; }
        public string Um { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PreUnitarioIgv { get; set; }
        public decimal PreTotaligv { get; set; }
        public string RnpVigente { get; set; }
        public string RegSanitario { get; set; }
        public string VigProducto { get; set; }
        public string ValidezOferta { get; set; }
        public string MarcapaisProcedencia { get; set; }
        public string PresentaProducto { get; set; }
        public string PlazaEntrega { get; set; }

    }
}

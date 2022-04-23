using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;


namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato_67_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_movil { get; set; }
        public string Prov_vigOferta { get; set; }
        public DateTime Prov_Fecha { get; set; }

        public List<Coti_Formato67_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato67_Detalle
    {
        public decimal NroItem { get; set; }
        public string Item { get; set; }
        public string Descripcion { get; set; }
        public string Um { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }
        public string Rnp { get; set; }
        public string NRegistro { get; set; }
        public string VigProducto { get; set; }
        public string CertBPM { get; set; }
        public string ProAnalisis { get; set; }
        public string ValiOferta { get; set; }
        public string Marcapaisesprocedencia { get; set; }
        public string CumpleEETT { get; set; }
        public string PlazoEntrega { get; set; }

    }
}

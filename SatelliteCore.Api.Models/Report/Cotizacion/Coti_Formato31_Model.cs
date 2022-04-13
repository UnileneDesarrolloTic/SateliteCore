using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;


namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato31_Model : CotizacionAbstract
    {
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Email { get; set; }
        public DateTime Prov_FeOferta { get; set; }
        public string Prov_Nombre { get; set; }
        public string Prov_Cargo { get; set; }
        public string Prov_Movil { get; set; }
        public string Prov_EmailContacto { get; set; }
        public DateTime Prov_PlazoOferta { get; set; }
        public List<Coti_Formato31_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato31_Detalle
    {
        public decimal NrItem { get; set; }
        public string SoliPediDelegado { get; set; }
        public string Posicion { get; set; }
        public string CodigoSAP { get; set; }
        public string Productos { get; set; }
        public decimal Cantidad { get; set; }
        public string Um { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }
        public string MarcaLaboratorio { get; set; }
        public string Procedencia { get; set; }
        public string ForPresentProducto { get; set; }
        public string PlazoEntrega { get; set; }
        public string VigMinima { get; set; }
        public string CumpAnalisis { get; set; }
        public string CumpDemoninacion { get; set; }
        public string CapAtencion { get; set; }
        public string CumplPlazo { get; set; }
        public string CertBPractica { get; set; }
        public string RegiSanitario { get; set; }
        public string CumpNormas { get; set; }

    }
}

using MongoDB.Bson.Serialization.Attributes;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato5_Model : CotizacionAbstract
    {

        public string NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Vigencia { get; set; }
        public string Prov_Ruc { get; set; }
        public string Cont_Telefono { get; set; }
        public string Cont_Email { get; set; }
        public string Cont_Area { get; set; }
        public string Ctac_Nombres { get; set; }
        public string Ctac_Cargo { get; set; }
        public string Ctac_Celular { get; set; }
        public string Ctac_Email1 { get; set; }
        public string Ctac_Email2 { get; set; }
        public decimal Foot_Total { get; set; }
        public List<Coti_Formato5_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato5_Detalle
    {
        public int NroItem { get; set; }
        public string CodigoSap { get; set; }
        public string Denominacion { get; set; }
        public string Um { get; set; }
        public decimal Cantidad { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public string PlazoEntrega { get; set; }
        public string VigenciaMinima { get; set; }
        public string AutoSanitario { get; set; }
        public string BuenasPracAlmace { get; set; }
        public string RegistroSanitario { get; set; }
        public string CartaRepresentacion { get; set; }
        public string Cbpm { get; set; }
        public string CertificadoAnalisis { get; set; }
        public string MetodologiaAnalisis { get; set; }
        public string FichaTecnica { get; set; }
        public string Folleteria { get; set; }
        public string CompPlazoEntrega { get; set; }
        public string CompReposicionDefecto { get; set; }
        public string Muestra { get; set; }
        public string RotuladoESSALUD { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal PrecioTotal { get; set; }
    }
}

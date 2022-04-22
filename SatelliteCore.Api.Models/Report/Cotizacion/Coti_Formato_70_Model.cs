using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;



namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    public class Coti_Formato_70_Model : CotizacionAbstract
    {
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Email1 { get; set; }
        public string Prov_valiOferta { get; set; }
        public string Prov_Representante { get; set; }
        public string Prov_Cargo { get; set; }
        public string Prov_Email2 { get; set; }
        public string Prov_movil { get; set; }
        public string Prov_Rpm { get; set; }
        public string Prov_CondPago { get; set; }
        public DateTime Prov_Fecha { get; set; }


        public List<Coti_Formato70_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato70_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodigoSAP { get; set; }
        public string Denominacion { get; set; }
        public string Um { get; set; }
        public decimal Cantidad { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public string VigProducto { get; set; }
        public string PlazoEntrega { get; set; }
        public string RegSanitario { get; set; }
        public string CProtocoloTerminado { get; set; }
        public string Metodoanalisis { get; set; }
        public string RAutoSanitaria { get; set; }
        public string Cbpm { get; set; }
        public string Cbpa { get; set; }
        public string Ccditem { get; set; }

    }
}
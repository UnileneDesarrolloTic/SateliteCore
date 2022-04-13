using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;


namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato64_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_ValiCotizacion { get; set; }
        public string Prov_PlazoEntrega { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Garantia { get; set; }
        public string Prov_vigProducto { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_ContTelefono { get; set; }
        public string Prov_ContEmail { get; set; }
        public string Prov_Cbancaria { get; set; }
        public string Prov_InfAdicional { get; set; }
        public DateTime Fecha_Cotizacion { get; set; }
        public decimal Prov_ValorTotal { get; set; }

        public List<Coti_Formato64_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato64_Detalle
    {

        public decimal NroItem { get; set; }
        public string Item { get; set; }
        public string Detalle { get; set; }
        public string UndMedida { get; set; }
        public decimal Cantidad { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public string Procedencia { get; set; }
        public string AnioFabricacion { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }

    }
}
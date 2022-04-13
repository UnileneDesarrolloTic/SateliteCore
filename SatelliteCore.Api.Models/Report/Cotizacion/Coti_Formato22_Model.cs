using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato22_Model : CotizacionAbstract
    {

        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_PlazoEntrega { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_Garantia { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_ContTelefono { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_ContCelular { get; set; }
        public string Prov_VigOferta { get; set; }
        public decimal Prov_ValorTotal { get; set; }
        public DateTime Fecha_Cotizacion { get; set; }

        public List<Coti_Formato22_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato22_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodMaterial { get; set; }
        public string Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string UndMedida { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public string Procedencia { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }


    }
}

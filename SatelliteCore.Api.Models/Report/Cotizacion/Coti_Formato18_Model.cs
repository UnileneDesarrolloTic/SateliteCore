using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato18_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc1 { get; set; }
        public string Prov_Codigo { get; set; }
        public string Prov_Direccion { get; set; }
        public string Prov_Telefono1 { get; set; }
        public string Prov_Fax { get; set; }
        public string Prov_CondVenta { get; set; }
        public DateTime  Prov_Fecha { get; set; }
        public string Prov_AsesComercial { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Telefono2 { get; set; }
        public string Prov_VigOferta { get; set; }
        public string Prov_PlazoEntrega { get; set; }
        public string Prov_LugarEntrega { get; set; }
        public string Prov_Garantia { get; set; }
        public string Prov_vigProducto { get; set; }
        public string Prov_Refencia { get; set; }
        public decimal Prov_ValorTotal { get; set; }
        public string Prov_precio { get; set; }
        public string Prov_FormaPago { get; set; }
        public string Prov_laboratorio { get; set; }
        public string Prov_Ruc2 { get; set; }
        public DateTime  Prov_FechaVenci { get; set; }
        public string Prov_Atencion { get; set; }

        public List<Coti_Formato18_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato18_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodSut { get; set; }
        public string Descripcion { get; set; }
        public string Marca { get; set; }
        public string Presentacion { get; set; }
        public string Procedencia { get; set; }
        public string UnidMedida { get; set; }
        public decimal Cantidad { get; set; }
        public string PlazoEntrega { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }

    }

}

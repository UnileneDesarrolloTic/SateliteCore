using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato10_Model : CotizacionAbstract
    {
        public string Nro_cotizacion { get; set; }
        public DateTime Fecha { get; set; }
        public string Prov_razonSocial { get; set; }
        public string Prov_direccion { get; set; }
        public string Prov_ruc { get; set; }
        public string Prov_email { get; set; }
        public string Vigencia_oferta { get; set; }
        public string Prov_telefono { get; set; }
        public string Cont_nombre { get; set; }
        public string Cont_cargo { get; set; }
        public string Cont_telefono { get; set; }
        public string Cont_email { get; set; }
        public decimal Monto_total { get; set; }
        public List<Formato10_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato10_Detalle
    {
        public int NroItem { get; set; }
        public string Sap { get; set; }
        public string Denominacion { get; set; }
        public string Um { get; set; }
        public decimal Requerimiento_total { get; set; }
        public string Proveedor { get; set; }
        public string Plazo_entrega { get; set; }
        public string Marca_modelo { get; set; }
        public string Procedencia { get; set; }
        public string Registro_proveedores { get; set; }
        public string Cumple_tdr { get; set; }
        public string Garantia_meses { get; set; }
        public decimal ValorUnitario { get; set; }
        public decimal ValorTotal { get; set; }

    }
}
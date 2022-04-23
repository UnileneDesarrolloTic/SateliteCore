using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato28_Model : CotizacionAbstract
    {
        public string Nro_Cotizacion { get; set; }
        public DateTime Fecha_documento { get; set; }
        public string Razon_social { get; set; }
        public string Ruc { get; set; }
        public string Direccion { get; set; }
        public string Telefono { get; set; }
        public string Email_unilene { get; set; }
        public string Vigencia { get; set; }
        public string Representante { get; set; }
        public string Cargo { get; set; }
        public string Email_representante { get; set; }
        public string Celular { get; set; }
        public string Rpm_rpc { get; set; }
        public string Codicion_pago { get; set; }
        public decimal Monto_total { get; set; }
        public List<Formato28_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato28_Detalle
    {
        public int Nroitem { get; set; }
        public string Sol_pedido { get; set; }
        public string Pos { get; set; }
        public string Codigo_sap { get; set; }
        public string Denominacion { get; set; }
        public string Um { get; set; }
        public decimal Cantidad { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public string Autorizacion_funcionamiento { get; set; }
        public string Bp_almacenamiento { get; set; }
        public string Registro_sanitario { get; set; }
        public string Bp_manufactura { get; set; }
        public string Metodologia_analisis { get; set; }
        public string Condicion_almacenamiento { get; set; }
        public string Plazo_entrega { get; set; }
        public decimal Precio_unitario { get; set; }
        public decimal Total { get; set; }

    }
}

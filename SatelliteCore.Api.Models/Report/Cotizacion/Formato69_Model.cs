using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato69_Model: CotizacionAbstract
    {
        public string Nro_cotizacion { get; set; }
        public DateTime Fecha_documento { get; set; }
        public int Codigo_cliente { get; set; }
        public string Cliente_razonSocial { get; set; }
        public string Ruc_cliente { get; set; }
        public string Direccion_cliente { get; set; }
        public string Telefono_cliente { get; set; }
        public string Fax_cliente { get; set; }
        public string Atencion { get; set; }
        public string Cond_venta { get; set; }
        public DateTime Fecha { get; set; }
        public string Asesor_prov { get; set; }
        public string Email_prov { get; set; }
        public string Telefono_prov { get; set; }
        public string Validez_oferta { get; set; }
        public string Plazo_entrega { get; set; }
        public string Lugar_entrega { get; set; }
        public string Garantia { get; set; }
        public string Vigencia_producto { get; set; }
        public string Ref { get; set; }
        public decimal Monto_total { get; set; }
        public DateTime Fecha_vencimiento { get; set; }
        public List<Formato69_Detalle> Detalle { get; set; }

    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato69_Detalle
    {
        public int NroItem { get; set; }
        public string Cod_sut { get; set; }
        public string Descripcion { get; set; }
        public string Marca { get; set; }
        public string Presentacion { get; set; }
        public string Procendencia { get; set; }
        public string Unidad_medida { get; set; }
        public decimal Cantidad { get; set; }
        public string Plazo_entrega { get; set; }
        public decimal Precio_unitario { get; set; }
        public decimal Total { get; set; }
    }
}

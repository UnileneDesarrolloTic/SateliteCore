using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato65_Model : CotizacionAbstract
    {
        public string Nro_Cotizacion { get; set; }
        public DateTime Fecha_1 { get; set; }
        public int Codigo { get; set; }
        public string Seniores { get; set; }
        public string Ruc { get; set; }
        public string Direccion { get; set; }
        public string Telefono { get; set; }
        public string Fax { get; set; }
        public string Atencion { get; set; }
        public string Condicion { get; set; }
        public DateTime Fecha_2 { get; set; }
        public string Asesor { get; set; }
        public string Email_asesor { get; set; }
        public string Telefono_asesor { get; set; }
        public string Validez_oferta { get; set; }
        public string Plazo_entrega { get; set; }
        public string Lugar_entrega { get; set; }
        public string Garantia { get; set; }
        public string Prov_VigProducto { get; set; }

        public string Prov_Ref { get; set; }
        public decimal Monto_total { get; set; }
        public List<Formato65_Detalle> Detalle { get; set; }
    }

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Formato65_Detalle 
    {
        public int NroItem { get; set; }
        public string Codsut { get; set; }
        public string Descripcion { get; set; }
        public string Marca { get; set; }
        public decimal Cantidad { get; set; }
        public decimal Precio_unitario { get; set; }
        public decimal Total { get; set; }
        public string Presentacion { get; set; }    
        public string Procedencia { get; set; }
        public string Unidadmedida { get; set; }
        public string PlazoEntrega { get; set; }

    }

}
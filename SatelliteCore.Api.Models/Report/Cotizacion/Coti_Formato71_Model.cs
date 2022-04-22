using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;



namespace SatelliteCore.Api.Models.Report.Cotizacion
{

    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato71_Model : CotizacionAbstract
    {
        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Representante { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_Celular { get; set; }
        public string Prov_Telefono { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_PlazoEntrega { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public decimal Prov_valorTotal { get; set; }

        public List<Coti_Formato71_Detalle> Detalle { get; set; }
    }


    [BsonIgnoreExtraElements(ignoreExtraElements: true)]

    public class Coti_Formato71_Detalle
    {
        public decimal NroItem { get; set; }
        public string CodigoSAP { get; set; }
        public string Descripcion { get; set; }
        public string UndMedida { get; set; }
        public decimal Cantidad { get; set; }
        public string Marca { get; set; }
        public string Rsanitario { get; set; }
        public string Procedencia { get; set; }
        public string Fichatecnica { get; set; }
        public string Crsvigente { get; set; }
        public string Cbpm { get; set; }
        
        public string Cbpa { get; set; }
        public string Bpdt { get; set; }
        public string Cmanalisis { get; set; }

        public string MetodoAnalisis { get; set; }

        public string Fichatenicaproducto { get; set; }
        public string Pmps { get; set; }
        public string Cuentamanual { get; set; }
        public string ResAutoSanitaria { get; set; }
        public string Anexol { get; set; }
        public string CvigminimaEntrega { get; set; }
        public string CPlazoEntrega { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }


    }
}

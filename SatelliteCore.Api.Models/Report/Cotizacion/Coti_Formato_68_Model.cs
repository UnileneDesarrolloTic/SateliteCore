using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;


namespace SatelliteCore.Api.Models.Report.Cotizacion
{
    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato_68_Model : CotizacionAbstract
    {

        public string Prov_NroCotizacion { get; set; }
        public string Prov_RazonSocial { get; set; }
        public string Prov_Ruc { get; set; }
        public string Prov_Email { get; set; }
        public string Prov_Contacto { get; set; }
        public string Prov_movil { get; set; }
        public string Prov_telefono { get; set; }
        public string Prov_RepVentas { get; set; }
        public string Prov_NombreBanco { get; set; }
        public string Prov_cci { get; set; }
        public string Prov_EntidadBancaria { get; set; }
        public string Prov_valiOferta { get; set; }
        public string Prov_referencia { get; set; }
        public string Prov_atencion { get; set; }
        public DateTime Prov_Fecha { get; set; }
        public decimal Prov_valorTotal { get; set; }

        public List<Coti_Formato68_Detalle> Detalle { get; set; }
    }


    [BsonIgnoreExtraElements(ignoreExtraElements: true)]
    public class Coti_Formato68_Detalle
    {

        public decimal NroItem { get; set; }
        public string Descripcion { get; set; }
        public string UndMedida { get; set; }
        public decimal Cantidad { get; set; }
        public string Nombcomercial { get; set; }
        public string Marca { get; set; }
        public string Procedencia { get; set; }
        public string Presentacion { get; set; }
        public decimal PreUnitario { get; set; }
        public decimal PreTotal { get; set; }
        public string VigProducto { get; set; }
        public string CumpFichaTecnica { get; set; }
        public string Garantia { get; set; }
        public string Plazoentrega { get; set; }

    }
}
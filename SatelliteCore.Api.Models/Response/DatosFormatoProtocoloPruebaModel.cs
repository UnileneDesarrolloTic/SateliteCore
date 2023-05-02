using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoProtocoloPruebaModel
    {
        public string LOTE { get; set; }
        public string   ORDENFABRICACION { get; set; }
        public string NUMERODEPARTE { get; set; }
        public string   ITEM { get; set; }
        public decimal CANTIDADPRODUCIDA{ get; set; }
        public DateTime FECHAPRODUCCION { get; set; }
        public DateTime FECHAEXPIRACION { get; set; }
        public string ITEMDESCRIPCION { get; set; }
        public string MARCA { get; set; }
        public string Presentacion { get; set; }
        public DateTime FECHAANALISIS { get; set; }
        public string TECNICA { get; set; }
        public string METODO { get; set; }
        public string DETALLE { get; set; }
        public int ID_DETALLE { get; set; }
        public int ID_FORMATO{ get; set; }
        public string DESCRIPCION_PRUEBA { get; set; }
        public string UNIDAD_MEDIDA { get; set; }
        public string ESPECIFICACION { get; set; }
        public string VALOR { get; set; }
        public string RESULTADO { get; set; }
        public string DESCRIPCION_METODOLOGIA { get; set; }
        public int ORDEN { get; set; }
    }
}

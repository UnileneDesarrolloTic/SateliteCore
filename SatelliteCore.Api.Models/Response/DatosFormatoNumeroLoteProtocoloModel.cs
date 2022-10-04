using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoNumeroLoteProtocoloModel
    {
        public string ORDENFABRICACION { get; set; }
        public string REFERENCIANUMERO { get; set; }
        public string ITEM { get; set; }
        public int CANTIDADPRODUCIDA { get; set; }
        public DateTime FECHAPRODUCCION { get; set; }
        public DateTime FECHAEXPIRACION { get; set; }
        public string NUMERODEPARTE { get; set; }
        public string ITEMDESCRIPCION { get; set; }
        public string MARCA { get; set; }
        public string Presentacion { get; set; }
        public string FECHAANALISIS { get; set; }
        public string TECNICA { get; set; }
        public string METODO { get; set; }
        public string DETALLE { get; set; }
        public decimal DMinimo { get; set; }
        public decimal DMaximo { get; set; }
        public decimal R_PromedioMinimo { get; set; }
        public decimal R_individualMinimo { get; set; }
        public decimal S_PromedioMinimo { get; set; }
        public decimal S_individualMinimo { get; set; }
        public decimal DEC_DMinimo { get; set; }
        public decimal DEC_DMaximo { get; set; }
        public decimal DEC_R_PromedioMinimo { get; set; }
        public decimal DEC_R_IndividualMinimo { get; set; }
        public decimal DEC_S_PromedioMinimo { get; set; }
        public decimal DEC_S_IndividualMinimo { get; set; }
    }
}

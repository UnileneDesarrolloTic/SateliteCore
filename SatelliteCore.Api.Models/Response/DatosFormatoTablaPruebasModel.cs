using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoTablaPruebasModel
    {
        public int ID_PRUEBA { get; set; }
        public int ID_AGRUPADOR_HEBRA { get; set; }
        public int ID_METODOLOGIA { get; set; }
        public string CALIBRE_USP { get; set; }
        public string UNIDAD_MEDIDA { get; set; }
        public decimal VALOR { get; set; }
        public DateTime FECHA_REGISTRO { get; set; }
        public string USUARIO { get; set; }
        public string ESTADO { get; set; }
        public string VERSION { get; set; }
        public int ID_PRUE { get; set; }
        public string ESPECIFFICACIONLOCAL { get; set; }
        public string ESPECIFFICACIONINGLES { get; set; }
        public int ID_ESPECIFICACION { get; set; }
        public int MOSTRAR { get; set; }
        public decimal cant_decimales { get; set; }
        public string ID_MARCA { get; set; }
        public string DESCRIPCIONLOCAL { get; set; }
        public string DESCRIPCIONINGLES { get; set; }
        public string DescripcionAgrupador { get; set; }
        public string DescripcionMetodologia { get; set; }
        public string DescripcionCalibre { get; set; }
    }
}

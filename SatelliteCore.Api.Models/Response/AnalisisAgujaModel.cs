using System;

namespace SatelliteCore.Api.Models.Response
{
    public struct AnalisisAgujaModel
    {
        public string NumeroOrden { get; set; }
        public string ControlNumero { get; set; }
        public DateTime FechaPreparacion { get; set; }
        public int secuencia { get; set; }
        public string ITEM { get; set; }
        public int CantidadPedida { get; set; }
        public int CantidadRecibida { get; set; }
        public string COD_PROVEEDOR { get; set; }
        public string PROVEEDOR { get; set; }
        public string ESTADO { get; set; }
        public string DESCRIPCION_ITEM { get; set; }
        public string UnidadCodigo { get; set; }
        public string LoteAprobado { get; set; }
        public string LoteRechazado { get; set; }
        public int Prueba_Detalle { get; set; }
        public int ID_ANALISIS { get; set; }
        public string ORDEN_COMPRA { get; set; }
        public DateTime FECHA_ANALISIS { get; set; }
        public Int32 CANTIDAD { get; set; }
        public string LOTE { get; set; }
        public Int32 CANT_PRUEBAS { get; set; }
        public Int32 CICLOS { get; set; }
        public Int32 Tiene_detalle { get; set; }

        //para la lista de ciclos
        public string identificador { get; set; }
        public string detalle { get; set; }
        public int val_1 { get; set; }
        public int val_2 { get; set; }


    }
}

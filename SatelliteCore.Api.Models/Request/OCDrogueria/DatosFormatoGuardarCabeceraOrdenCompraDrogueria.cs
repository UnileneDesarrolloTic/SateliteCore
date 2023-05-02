using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request.OCDrogueria
{
    public struct DatosFormatoGuardarCabeceraOrdenCompraDrogueria
    {
        public int codigo { get; set; }
        public DateTime fecha { get; set; }
        public string estado { get; set; }
        public int  persona { get; set; }
        public string proveedor { get; set; }
        public decimal montoTotal { get; set; }
        public string moneda { get; set; }
        public string procedencia { get; set; }  
        public DateTime fechaPrometida { get; set; }
        public decimal diasespera { get; set; }
        public string viaEnvio { get; set; }
        public string incoterms { get; set; }
        public string paisOrigen { get; set; }
        public string puertoSalida { get; set; }
        public List<DatosFormatoGuardarDetalleOrdenCompra> detalle { get; set; }
    }
}

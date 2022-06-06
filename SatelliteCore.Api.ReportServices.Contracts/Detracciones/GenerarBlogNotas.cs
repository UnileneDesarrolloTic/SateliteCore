using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Collections;
using System.Text;
using SatelliteCore.Api.Models.Request;

namespace SatelliteCore.Api.ReportServices.Contracts.Detracciones
{
    public class GenerarBlogNotas
    {
        public string ProcesarGenerarBlogNotas(FormatoProcesoDetracciones dato)
        {
            StringBuilder informacionBlogNota = new StringBuilder();
            string valorunilene = "P20197705249UNILENE SAC";
            var cuenta = "00002007770";
            informacionBlogNota.Append(valorunilene.PadRight(47)).Append(dato.periodo.PadRight(6,'0') + dato.totalimporte.ToString("D15") + "\n");
            
            foreach (var detraccion in dato.proceso)
            {
                var numero = Int32.Parse(detraccion.Numero);
                var servicio = Int32.Parse(detraccion.BienServicio);
                var operacion = Int32.Parse(detraccion.Tipo);
                var tipo = Int32.Parse(detraccion.Tipo);

                informacionBlogNota.Append(detraccion.TipoDocumento + detraccion.Ruc.PadRight(46) + servicio.ToString("D12") + cuenta + detraccion.Importe.ToString("D13") + operacion.ToString("D4") + detraccion.Periodo + tipo.ToString("D2") + detraccion.Serie + numero.ToString("D8") + "\n");
            }
           
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(informacionBlogNota.ToString());


            return System.Convert.ToBase64String(plainTextBytes);
        }

      

    }
}

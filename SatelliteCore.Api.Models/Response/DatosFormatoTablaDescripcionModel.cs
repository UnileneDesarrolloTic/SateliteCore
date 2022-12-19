using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public struct DatosFormatoTablaDescripcionModel
    {
        public int ID_DESCRIPCION { get; set; }
        public string ID_MARCA { get; set; }
        public string ID_HEBRA { get; set; }
        public string DESCRIPCIONLOCAL { get; set; }
        public string DESCRIPCIONINGLES { get; set; }
        public string ESTADO { get; set; }
        public string USUARIO { get; set; }
        public DateTime ULTIMAFECHA { get; set; }
        public string DESCRIPCIONMARCA { get; set; }
        public string DESCRIPCIONHEBRA { get; set; }
    }
}

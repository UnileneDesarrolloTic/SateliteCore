﻿using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class FormatoDatosBusquedaItemsVentas
    {
        public string Item { get; set; }
        public string Codsut { get; set; }
        public string Descripcion { get; set; }
        public int Origen { get; set; }
        public string idAgrupador { get; set; }
        public string idmarca { get; set; }
    }
}

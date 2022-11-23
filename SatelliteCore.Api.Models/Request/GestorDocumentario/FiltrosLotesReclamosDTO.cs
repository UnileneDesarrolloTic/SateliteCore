namespace SatelliteCore.Api.Models.Request.GestorDocumentario
{
    public struct FiltrosLotesReclamosDTO
    {
        public int Cliente { get; set; }
        public string TipoFiltro { get; set; }
        public string ValorFiltro { get; set; }

        public bool Validacion()
        {

            if(Cliente < 1 || string.IsNullOrEmpty(TipoFiltro) || (ValorFiltro == "O" || ValorFiltro == "L" ))
                return false;
            return true;
        }
    }
}

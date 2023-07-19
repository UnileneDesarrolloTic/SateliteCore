namespace SatelliteCore.Api.Models.Entities
{
    public struct TBDAnalisisHebraEntity
    {
        public string OrdenCompra { get; set; }
        public string NumeroAnalisis { get; set; }
        public int Numero { get; set; }
        public decimal? Longitud { get; set; }
        public decimal Diametro { get; set; }
        public decimal Tension { get; set; }

        public bool ValidarDatos()
        {
            if(string.IsNullOrWhiteSpace(OrdenCompra) || string.IsNullOrWhiteSpace(NumeroAnalisis) || Numero < 1)
                return false;
            return true;
        }
    }
}

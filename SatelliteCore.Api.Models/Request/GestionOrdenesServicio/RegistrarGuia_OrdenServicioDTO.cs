namespace SatelliteCore.Api.Models.Request.GestionOrdenesServicio
{
    public struct RegistrarGuia_OrdenServicioDTO
    {
        public int Cabecera { get; set; }
        public string Serie { get; set; }
        public string Guia { get; set; }
        public decimal Peso { get; set; }
        public int Bultos { get; set; }
        public string Usuario { get; set; }

        public RegistrarGuia_OrdenServicioDTO(int cabecera, string serie, string guia, decimal peso, int bultos, string usuario)
        {
            Cabecera = cabecera;
            Serie = serie;
            Guia = guia;
            Peso = peso;
            Bultos = bultos;
            Usuario = usuario;
        }
    }
}

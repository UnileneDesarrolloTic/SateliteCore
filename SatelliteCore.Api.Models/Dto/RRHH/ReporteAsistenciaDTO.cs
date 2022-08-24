namespace SatelliteCore.Api.Models.Dto.RRHH
{
    public struct ReporteAsistenciaDTO
    {
        public int IdEmpleado { get; set; }
        public string Area { get; set; }
        public string NombreCompleto { get; set; }
        public string AreaHorario { get; set; }
        public string HoraInicioPlanta { get; set; }
        public string HoraInicioPlantaColor { get; set; }
        public string HoraInicioComida { get; set; }
        public string HoraInicioComidaColor { get; set; }
        public string HoraFinComida { get; set; }
        public string HoraFinComidaColor { get; set; }
        public string HoraFinPlanta { get; set; }
        public string HoraFinPlantaColor { get; set; }
        public string TipoPlanilla { get; set; }

    }
}

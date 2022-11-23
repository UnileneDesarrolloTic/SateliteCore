namespace SatelliteCore.Api.CrossCutting.Helpers
{
    public static class QueryScript
    {
        public static string ConsultaPaginada(string consulta, string queryOrder)
        {
            return $"IF OBJECT_ID('Tempdb..#Temp_Paginacion') IS NOT NULL DROP TABLE #Temp_Paginacion {consulta} " +
                $"SELECT * FROM #Temp_Paginacion {queryOrder} OFFSET (@pagina - 1) * @registrosPorPagina ROWS FETCH NEXT @registrosPorPagina ROWS ONLY; SELECT COUNT(1) CantidadRegistros FROM #Temp_Paginacion";
        }

        public static string ConsultaValidacion(string consulta)
        {
            return $"IF EXISTS ( {consulta} ) SELECT 1 Existe ELSE SELECT 0 Existe";
        }
    }
}

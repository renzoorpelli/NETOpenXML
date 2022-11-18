namespace Entidades.Modelo
{
    public class Dato
    {
        public int Id { get; set; }
        public string Nombre { get; set; } = null!;
        public string Descripcion { get; set; } = null!;
        public DateTime FechaCompra { get; set; }
        public decimal Valor { get; set; }
        public int AmortizacionAnual { get; set; }

        public string Total { get; set; } = null!;
    }
}
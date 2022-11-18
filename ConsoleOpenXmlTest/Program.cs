using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using Entidades.Modelo;
using Entidades.ServicioExcel;
using System;

namespace Application
{
    class Program
    {
        static void Main(string[] args)
        {
            
            //declaro la lista de objetos
            ListaDato listaDatos = new ListaDato();

            Dato a1 = new Dato
            {
                Id = 1,
                Nombre = "Inmueble",
                Descripcion = "Inmueble Local Venta",
                FechaCompra = DateTime.Now.Date,
                AmortizacionAnual = 2,
                Valor = 10000000,
            };
            a1.Total = $"={a1.Valor}-({a1.Valor}*{a1.AmortizacionAnual})/100";


            Dato a2 = new Dato
            {
                Id = 2,
                Nombre = "Rodado",
                Descripcion = "Rodado destinado a envios",
                FechaCompra = DateTime.Now.AddDays(3),
                AmortizacionAnual = 20,
                Valor = 2000000

            };

            a2.Total = $"={a2.Valor}-({a2.Valor}*{a2.AmortizacionAnual})/100";


            Dato a3 = new Dato
            {
                Id = 3,
                Nombre = "Muebles y Utiles",
                Descripcion = "Destinados a amoblamiento del local",
                FechaCompra = DateTime.Now.AddDays(1),
                AmortizacionAnual = 10,
                Valor = 300000

            };
            a3.Total = $"={a3.Valor}-({a3.Valor}*{a3.AmortizacionAnual})/100";

            listaDatos.Cuentas.Add(a1);
            listaDatos.Cuentas.Add(a2);
            listaDatos.Cuentas.Add(a3);

            ExcelService excelService = new ExcelService("./Generate/");
            excelService.CreateExcel(listaDatos, "datos");
        }



       
    }
}
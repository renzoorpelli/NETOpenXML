using Entidades.Modelo;
using System;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Linq.Expressions;
using NCalc;

namespace Entidades.Servicios
{
    public class PDFService
    {
        private string baseDirectory;

        public PDFService(string baseDirectory)
        {
            this.baseDirectory = baseDirectory;
        }

        public void CreatePDF(string filename, ListaDato datos)
        {
            var dateFromCreate = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss").Replace('/', '_').Replace(':', '_');

            string fileFullName = Path.Combine(baseDirectory, filename.Trim() + ".pdf");

            if (File.Exists(fileFullName))
            {
                fileFullName = Path.Combine(baseDirectory, "datos_de_" + dateFromCreate + ".pdf");
            }
            else
            {
                Directory.CreateDirectory(this.baseDirectory);
            }

            if (GenerarPDF(fileFullName, datos))
            {
                Console.WriteLine("PDF GENERADO " + DateTime.Now.ToShortDateString());
            }
        }


        private bool GenerarPDF(string fileFullName, ListaDato datos)
        {
            try
            {
                var pdfDoc = new Document(PageSize.LETTER, 40f, 40f, 60f, 60f);
                var objectKeys = ObtenerKeysObjeto(datos.Cuentas[1]);


                PdfWriter.GetInstance(pdfDoc, new FileStream(fileFullName, FileMode.OpenOrCreate));
                pdfDoc.Open();


                var spacer = new Paragraph("")
                {
                    SpacingBefore = 10f,
                    SpacingAfter = 10f,
                };
                pdfDoc.Add(spacer);

                var columnCount = objectKeys.Count;
                var columnWidths = new[] { 1f, 2f, 3f, 3f, 3f, 1f, 2f };

                var table = new PdfPTable(columnWidths)
                {
                    HorizontalAlignment = 0, //0=Left, 1=Centre, 2=Right
                    WidthPercentage = 100,
                    DefaultCell = { MinimumHeight = 22f }
                };

                var cell = new PdfPCell(new Phrase("Reportes Cuentas Afectadas"))
                {
                    Colspan = columnCount,
                    HorizontalAlignment = 1,  //0=Left, 1=Centre, 2=Right
                    MinimumHeight = 30f
                };

                table.AddCell(cell);

                //armo el header de la tabla, por cada key de el array genero una comlumna
                objectKeys.ForEach(key =>
                { 
                    string keyToString = key.ToString()!;

                    if (keyToString == "Total")
                    {
                        keyToString = "Formula Excel";
                    }
                    if(keyToString == "AmortizacionAnual")
                    {
                        keyToString = "A.A";
                    }

                    table.AddCell(keyToString);
                });

                //rows para cada objeto del cuerpo de la tabla
                datos.Cuentas.ForEach(cuenta =>
                {
                    //obtengo las propiedades del objeto
                    var objectKeys = cuenta.GetType().GetProperties();

                    foreach (var key in objectKeys)
                    {
                        //por cada valor de la KEY del objeto agrego una celda
                        var cellData = new PdfPCell(new Phrase(key.GetValue(cuenta)?.ToString()));
                        table.AddCell(cellData);
                    }

                });

                pdfDoc.Add(table);

                pdfDoc.Close();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        public List<string> ObtenerKeysObjeto(Dato objeto)
        {
            List<string> keys = new List<string>();
            var objectKeys = objeto.GetType().GetProperties();
            foreach (var key in objectKeys)
            {
                keys.Add(key.Name.ToString()!);
            }
            return keys;
        }



    }
}

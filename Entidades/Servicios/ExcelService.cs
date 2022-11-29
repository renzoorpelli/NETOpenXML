using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Entidades.Modelo;
using PageMargins = DocumentFormat.OpenXml.Spreadsheet.PageMargins;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace Entidades.Servicios
{
    public class ExcelService
    {
        private readonly string baseDirectory;

        public ExcelService(string baseDirectory)
        {
            this.baseDirectory = baseDirectory;
        }

        public bool CreateExcel(ListaDato data, string fileName)
        {
            var dateFromCreate = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss").Replace('/', '_').Replace(':', '_');

            string fileFullName = Path.Combine(baseDirectory, fileName.Trim() + ".xlsx");

            if (File.Exists(fileFullName))
            {
                fileFullName = Path.Combine(baseDirectory, "datos_de_" + dateFromCreate + ".xlsx");
            }
            else
            {
                Directory.CreateDirectory(this.baseDirectory);
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullName, SpreadsheetDocumentType.Workbook))
            {
                if (CreatePartsForExcel(package, data))
                {
                    return true;
                }
                return false;
            }
        }

        private bool CreatePartsForExcel(SpreadsheetDocument document, ListaDato data)
        {
            if (document is null || data is null)
            {
                return false;
            }
            SheetData partSheetData = GenerateSheetDtaForDetails(data);

            WorkbookPart wbPartOne = document.AddWorkbookPart();

            GenerateWorbookPartContent(wbPartOne);

            WorkbookStylesPart wbStylesPartOne = wbPartOne.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkBookStylesPartContent(wbStylesPartOne);

            WorksheetPart wsPartOne = wbPartOne.AddNewPart<WorksheetPart>("rId1");

            GenerateWorkSheetPartContent(wsPartOne, partSheetData);
            return true;
        }

        /// <summary>
        /// método encargado de generar contenido para el workbook de excel
        /// </summary>
        /// <param name="workbookPartOne">La parte 1 del worboook generado en CreatePartsFromExcel()</param>
        private void GenerateWorbookPartContent(WorkbookPart workbookPartOne)
        {
            Workbook wb = new Workbook();
            Sheets sh = new Sheets();

            Sheet sheetOne = new Sheet()
            {
                Name = "Hoja1",
                SheetId = (UInt32Value)1U,
                Id = "rId1"
            };
            sh.Append(sheetOne);
            wb.Append(sh);

            workbookPartOne.Workbook = wb;
        }

        /// <summary>
        /// método encargado de generar el contenido para la parte de la hoja de trabajo. 
        /// se encargará de darle el formato a la hoja
        /// </summary>
        /// <param name="wsPartOne">la parte de la hoja de trabajo</param>
        /// <param name="sdOne">los datos de la hoja </param>
        private void GenerateWorkSheetPartContent(WorksheetPart wsPartOne, SheetData sdOne)
        {
            Worksheet ws = new Worksheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" }
            };
            //prefix estandards
            ws.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            ws.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            ws.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetDimension sd = new SheetDimension()
            {
                Reference = "A1" // tamaño de la hoja
            };

            SheetViews svs = new SheetViews();

            SheetView sv = new SheetView()
            {
                TabSelected = true,
                WorkbookViewId = (UInt32Value)0U
            };

            Selection selection = new Selection()
            {
                ActiveCell = "A1",
                SequenceOfReferences = new ListValue<StringValue>()
                {
                    InnerText = "A1"
                }
            };
            sv.Append(selection);
            svs.Append(sv);

            SheetFormatProperties sfp = new SheetFormatProperties()
            {
                DefaultRowHeight = 15D,
                DyDescent = 0.25D
            };

            PageMargins pm = new PageMargins()
            {
                Left = 0.7D,
                Right = 0.7D,
                Top = 0.75D,
                Bottom = 0.75D,
                Header = 0.3D,
                Footer = 0.3D
            };

            ws.Append(sd);
            ws.Append(svs);
            ws.Append(sfp);
            ws.Append(sdOne);
            ws.Append(pm);
            wsPartOne.Worksheet = ws;

        }
        /// <summary>
        /// metodo encargado de generar el estilo para el workbook
        /// </summary>
        /// <param name="workbookStylesPartOne"></param>
        private void GenerateWorkBookStylesPartContent(WorkbookStylesPart workbookStylesPartOne)
        {
            Stylesheet styleSheet = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes()
                {
                    Ignorable = "x14ac"
                }
            };
            styleSheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styleSheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");


            #region font region
            Fonts fonts = new Fonts()
            {
                Count = (UInt32Value)2U,
                KnownFonts = true
            };

            Font fontOne = new Font();

            FontSize fontSizefOne = new FontSize()
            {
                Val = 11D
            };
            Color colorfOne = new Color()
            {
                Theme = (UInt32Value)1U
            };
            FontName fontNamefOne = new FontName()
            {
                Val = "Arial"
            };
            FontFamilyNumbering fontFNumberingfOne = new FontFamilyNumbering()
            {
                Val = 2
            };
            FontScheme fontSchemefOne = new FontScheme()
            {
                Val = FontSchemeValues.Minor
            };

            //agrego todas las propiedades de la fuente 1
            fontOne.Append(fontSizefOne);
            fontOne.Append(colorfOne);
            fontOne.Append(fontNamefOne);
            fontOne.Append(fontFNumberingfOne);
            fontOne.Append(fontSchemefOne);

            //fuente para la row Header
            Font fontTwo = new Font();

            Bold boldforFontTwo = new Bold();

            FontSize fontSizefTwo = new FontSize()
            {
                Val = 16D
            };
            Color colorfTwo = new Color()
            {
                Theme = (UInt32Value)1U
            };
            FontName fontNamefTwo = new FontName()
            {
                Val = "Arial"
            };
            FontFamilyNumbering fontFNumberingfTwo = new FontFamilyNumbering()
            {
                Val = 2
            };
            FontScheme fontSchmefTwo = new FontScheme()
            {
                Val = FontSchemeValues.Minor
            };

            fontTwo.Append(boldforFontTwo);
            fontTwo.Append(fontSizefTwo);
            fontTwo.Append(colorfTwo);
            fontTwo.Append(fontNamefTwo);
            fontTwo.Append(fontFNumberingfTwo);
            fontTwo.Append(fontSchmefTwo);

            fonts.Append(fontOne);
            fonts.Append(fontTwo);

            #endregion


            #region fills region
            Fills fills = new Fills()
            {
                Count = (UInt32Value)0U
            };

            Fill fillOne = new Fill();
            PatternFill patterfFillOne = new PatternFill()
            {
                PatternType = PatternValues.None
            };

            fillOne.Append(patterfFillOne);

            Fill fillTwo = new Fill();

            PatternFill patterfFillTwo = new PatternFill()
            {
                PatternType = PatternValues.Gray125
            };
           
            fillTwo.Append(patterfFillTwo);

            //style each cell of header tittle
            Fill fillThree = new Fill();
            PatternFill patternFillThree = new PatternFill() 
            { 
                PatternType = PatternValues.Solid 
            };

            ForegroundColor backgroundColor = new ForegroundColor() 
            { 
                Rgb = new HexBinaryValue("ffe000") 
            };

            patternFillThree.Append(backgroundColor);
            fillThree.Append(patternFillThree);


            //0 to N value
            fills.Append(fillOne);
            fills.Append(fillTwo);
            fills.Append(fillThree);


            #endregion


            #region borders
            Borders borders = new Borders()
            {
                Count = (UInt32Value)2U
            };

            Border borderOne = new Border();


            LeftBorder lbOne = new LeftBorder()
            {
                Style = BorderStyleValues.Thick
            };
            Color colorLbOne = new Color()
            {
                Indexed = (UInt32Value)64U
            };

            lbOne.Append(colorLbOne);

            RightBorder rbOne = new RightBorder()
            {
                Style = BorderStyleValues.Thick
            };

            Color colorRbOne = new Color()
            {
                Indexed = (UInt32Value)64U
            };

            rbOne.Append(colorRbOne);


            TopBorder tbOne = new TopBorder()
            {
                Style = BorderStyleValues.Thick
            };

            Color colortbOne = new Color()
            {
                Indexed = (UInt32Value)64U
            };

            tbOne.Append(colortbOne);

            BottomBorder bBorderOne = new BottomBorder()
            {
                Style = BorderStyleValues.Thick
            };

            Color colorBbOne = new Color()
            {
                Indexed = (UInt32Value)64U
            };
            bBorderOne.Append(colorBbOne);



            borderOne.Append(lbOne);
            borderOne.Append(rbOne);
            borderOne.Append(tbOne);
            borderOne.Append(bBorderOne);

            Border borderTwo = new Border();

            LeftBorder lbTwo = new LeftBorder()
            {
                Style = BorderStyleValues.Thick
            };
            Color colorLb = new Color()
            {
                Indexed = (UInt32Value)64U
            };

            lbTwo.Append(colorLb);

            RightBorder rbTwo = new RightBorder()
            {
                Style = BorderStyleValues.Thick
            };
            Color colorRb = new Color()
            {
                Indexed = (UInt32Value)64U
            };

            rbTwo.Append(colorRb);

            TopBorder tbTwo = new TopBorder()
            {
                Style = BorderStyleValues.Thick
            };
            Color colorTb = new Color()
            {
                Indexed = (UInt32Value)64U
            };
            tbTwo.Append(colorTb);

            BottomBorder bBorderTwo = new BottomBorder()
            {
                Style = BorderStyleValues.Thick
            };
            Color colorBb = new Color()
            {
                Indexed = (UInt32Value)64U
            };

            bBorderTwo.Append(colorBb);

            DiagonalBorder dbBorderTwo = new DiagonalBorder();

            borderTwo.Append(lbTwo);
            borderTwo.Append(rbTwo);
            borderTwo.Append(tbTwo);
            borderTwo.Append(bBorderTwo);
            borderTwo.Append(dbBorderTwo);

            borders.Append(borderOne);
            borders.Append(borderTwo);

            #endregion


            CellStyleFormats cellStyleFormats = new CellStyleFormats()
            {
                Count = (UInt32Value)1U
            };

            //manage the empty filds
            CellFormat cellFormatOne = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)0U
            };

            cellStyleFormats.Append(cellFormatOne);

            CellFormats cellFormats = new CellFormats()
            {
                Count = (UInt32Value)3U
            };

            //manage the empty fields
            CellFormat cellFormatTwo = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)0U,
                FormatId = (UInt32Value)0U
            };

            //manage the cell of data
            CellFormat cellFormatThree = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)0U,
                FillId = (UInt32Value)0U,
                BorderId = (UInt32Value)1U,
                FormatId = (UInt32Value)0U,
                ApplyBorder = BooleanValue.FromBoolean(true)
            };

            //manage the row header
            CellFormat cellFormatFour = new CellFormat()
            {
                NumberFormatId = (UInt32Value)0U,
                FontId = (UInt32Value)1U,
                FillId = (UInt32Value)2U,
                BorderId = (UInt32Value)1U,
                FormatId = (UInt32Value)0U,
                ApplyFont = BooleanValue.FromBoolean(true),
                ApplyBorder = BooleanValue.FromBoolean(true)
            };

            


            cellFormats.Append(cellFormatTwo);
            cellFormats.Append(cellFormatThree);
            cellFormats.Append(cellFormatFour);


            CellStyles cellStylesOne = new CellStyles()
            {
                Count = (UInt32Value)1U
            };
            CellStyle cellStyle = new CellStyle()
            {
                Name = "Normal",
                FormatId = (UInt32Value)0U,
                BuiltinId = (UInt32Value)0U
            };

            cellStylesOne.Append(cellStyle);


            DifferentialFormats differentialFormats = new DifferentialFormats()
            {
                Count = (UInt32Value)0U
            };

            TableStyles tableStyles = new TableStyles()
            {
                Count = (UInt32Value)0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };

            StylesheetExtensionList styleSheetExtensionList = new StylesheetExtensionList();

            StylesheetExtension styleSheetExtesionOne = new StylesheetExtension()
            {
                Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
            };
            styleSheetExtesionOne.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStylesOne = new X14.SlicerStyles()
            {
                DefaultSlicerStyle = "SlicerStyleLight1"
            };

            styleSheetExtesionOne.Append(slicerStylesOne);

            StylesheetExtension styleSheetExtensionTwo = new StylesheetExtension()
            {
                Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}"
            };
            styleSheetExtensionTwo.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyle = new X15.TimelineStyles()
            {
                DefaultTimelineStyle = "TimeSlicerStyleLight1"
            };


            styleSheetExtensionTwo.Append(timelineStyle);

            styleSheetExtensionList.Append(styleSheetExtesionOne);
            styleSheetExtensionList.Append(styleSheetExtensionTwo);

            styleSheet.Append(fonts);
            styleSheet.Append(fills);
            styleSheet.Append(borders);
            styleSheet.Append(cellStyleFormats);
            styleSheet.Append(cellFormats);
            styleSheet.Append(cellStylesOne);
            styleSheet.Append(differentialFormats);
            styleSheet.Append(tableStyles);
            styleSheet.Append(styleSheetExtensionList);

            workbookStylesPartOne.Stylesheet = styleSheet;
        }

        /// <summary>
        /// metodo encargado de agregar todos los datos al archivo excel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private SheetData GenerateSheetDtaForDetails(ListaDato data)
        {
            SheetData sheetData = new SheetData();
            sheetData.Append(CreateHeaderRowForExcel());
            foreach (var cuenta in data.Cuentas)
            {
                Row partRows = GenerateRowForChildPartDetail(cuenta);
                sheetData.Append(partRows);
            }
            return sheetData;
        }

        /// <summary>
        /// metodo encargado de realizar el header de cada uno de los atributos
        /// </summary>
        /// <returns></returns>
        private Row CreateHeaderRowForExcel()
        {
            Row headerRow = new Row();

            headerRow.Append(CreateCell("ID", 2U));
            headerRow.Append(CreateCell("Nombre", 2U));
            headerRow.Append(CreateCell("Descripción", 2U));
            headerRow.Append(CreateCell("Fecha Compra", 2U));
            headerRow.Append(CreateCell("Valor", 2U));
            headerRow.Append(CreateCell("Amortización anual", 2U));
            headerRow.Append(CreateCell("Total anual", 2U));
            return headerRow;
        }

        private Row GenerateRowForChildPartDetail(Dato cuenta)
        {
            Row tRow = new Row();
            tRow.Append(CreateCell(cuenta.Id.ToString()));
            tRow.Append(CreateCell(cuenta.Nombre));
            tRow.Append(CreateCell(cuenta.Descripcion));
            tRow.Append(CreateCell(cuenta.FechaCompra.ToShortDateString()));
            tRow.Append(CreateCell(cuenta.Valor.ToString()));
            tRow.Append(CreatePorcentCell(cuenta.AmortizacionAnual.ToString()));
            tRow.Append(CreateFormulaCell(cuenta.Total));


            return tRow;
        }

        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = (UInt32Value)1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);

            return cell;
        }

        private Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();

            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        /// <summary>
        /// metodo encargado de aplicar la formula a una celda
        /// NumberFormatId = 10 refiere al 0.00
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private Cell CreateFormulaCell(string text)
        {
            Cell cell = new Cell();
            CellFormula formula = new CellFormula(text);
            cell.StyleIndex = 1U;
            formula.CalculateCell = true;
            cell.Append(formula);
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(text);
            return cell;
        }
        /// <summary>
        /// metodo encargado de establecer el correcto formateo para valores de una formula
        /// NumberFormatId = 10 refiere al 0.00%
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private Cell CreatePorcentCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(text);
            return cell;
        }

        /// <summary>
        /// metodo encargado de verificar si el dato que se le pasa por parametro es un numero
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intValue;
            decimal decimalValue;
            if (!int.TryParse(text, out intValue) && !decimal.TryParse(text, out decimalValue))
            {
                return CellValues.String;
            }
            return CellValues.Number;
        }

        

    }
}

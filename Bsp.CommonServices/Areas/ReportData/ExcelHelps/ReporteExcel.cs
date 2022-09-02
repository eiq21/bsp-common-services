using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;

namespace Bsp.CommonServices.Areas.ReportData.ExcelHelps
{

    public class ReporteExcel
    {
        public CodePageUtil CENTER;
        public CodePageUtil CENTER_SELECTION;
        public CodePageUtil DISTRIBUTED;
        public CodePageUtil RIGHT;
        public CodePageUtil LEFT;

        /// <summary>
        /// libro contenedor de hojas
        /// </summary>
        public HSSFWorkbook libro { get; set; }

        /// <summary>
        /// hojas del libro
        /// </summary>
        public ISheet hoja { get; set; }

        /// <summary>
        /// creacion de la hoja
        /// </summary>
        /// <param name="nombrePagina"></param>
        /// <param name="indLibro"></param>
        public void creaHoja(string nombrePagina, string indLibro)
        {
            if (indLibro.Equals("S"))
            {
                libro = new HSSFWorkbook();

            }

            hoja = libro.CreateSheet(nombrePagina);


        }

        /// <summary>
        /// Agrega el titulo al archivo Excel
        /// </summary>
        /// <param name="filaInicio">fila inicial</param>
        /// <param name="filaFin">fila final</param>
        /// <param name="colInicial">columna Inicial</param>
        /// <param name="colFinal">columna final</param>
        /// <param name="valor">Nombre dek titulo</param>
        /// <param name="style">estilo del titulo</param>
        public void addTituloExcel(int filaInicio, int filaFin, int colInicial, int colFinal, String valor, ICellStyle style)
        {
            String error;
            IRow row;

            try
            {
                // crea la fila inicial
                row = hoja.CreateRow(filaInicio);
                // crea la celda
                addCelda(row, colInicial, valor, style, "S", false);
                // coloca el titulo en el centro de acuerdo a la cantidad de columnas
                hoja.AddMergedRegion(new CellRangeAddress(filaInicio, filaFin, colInicial, colFinal));
                filaInicio += 2;

            }
            catch (Exception e)
            {
                error = "Error al agregar el titulo a la hoja del archivo excel -> " + e;
                throw new Exception(error);
            }
        }

        /// <summary>
        /// aggrega una celda
        /// </summary>
        /// <param name="row">fila</param>
        /// <param name="col">posicion de la columna</param>
        /// <param name="valor">valor</param>
        /// <param name="style">estilo de la letra</param>
        /// <param name="TipoDato">Tipo de dato</param>
        public void addCelda(IRow row, int col, string valor, ICellStyle style, string TipoDato, bool WrapText = false)
        {
            ICell celda; //se crea la celda en la posicion que se indica
            celda = row.CreateCell(col);
            try
            {
                if (TipoDato.Equals("N"))
                {
                    celda.SetCellType(CellType.Numeric);
                    celda.SetCellValue(Convert.ToInt64(valor));

                }
                else
                {
                    celda.SetCellValue(valor);
                }
            }
            catch (Exception)
            {
                celda.SetCellValue(valor);
            }


            celda.CellStyle = style;

            if (WrapText == true)
            {
                celda.CellStyle.WrapText = true;
            }
            else
            {
                celda.CellStyle.WrapText = false;
            }

        }




        /// <summary>
        /// agrega el merge
        /// </summary>
        /// <param name="filaInicio"></param>
        /// <param name="filaFin"></param>
        /// <param name="colInicial"></param>
        /// <param name="colFinal"></param>
        public void addCellMerge(IRow row, int filaInicio, int filaFin, int colInicial, int colFinal, String valor, ICellStyle style)
        {
            String error;

            try
            {

                addCelda(row, colInicial, valor, style, "S", false);
                // coloca el titulo en el centro de acuerdo a la cantidad de columnas
                hoja.AddMergedRegion(new CellRangeAddress(filaInicio, filaFin, colInicial, colFinal));

            }
            catch (Exception e)
            {
                error = "Error al agregar el titulo a la hoja del archivo excel -> " + e;
                throw new Exception(error);
            }
        }


        /// <summary>
        /// crea la fila y devulve la fila creada
        /// </summary>
        /// <param name="numFila">numero de fila</param>
        /// <returns>Returna un Irow una fila</returns>
        public IRow addFila(int numFila)
        {
            IRow row;
            row = hoja.CreateRow(numFila);
            return row;
        }

        /// <summary>
        /// Adiciona estilo para cadenas string
        /// </summary>
        /// <param name="negrita"></param>
        /// <param name="point"></param>
        /// <param name="alineamiento"></param>
        /// <returns></returns>
        public ICellStyle addEstiloCadena(Boolean negrita, int point, String alineamiento, bool vColorCeldaCabecera = false)
        {

            IFont fontCadena = libro.CreateFont();
            ICellStyle styleCadena = libro.CreateCellStyle();


            fontCadena.Color = HSSFColor.Black.Index;

            if (negrita)
            {
                fontCadena.Boldweight = (short)FontBoldWeight.Bold;
            }

            fontCadena.FontName = HSSFFont.FONT_ARIAL;

            fontCadena.FontHeightInPoints = (short)point;

            if ("CENTER".Equals(alineamiento))
            {
                styleCadena.Alignment = HorizontalAlignment.Center;
            }
            if ("CENTER_SELECTION".Equals(alineamiento))
            {
                styleCadena.Alignment = HorizontalAlignment.CenterSelection;
            }
            if ("DISTRIBUTED".Equals(alineamiento))
            {
                styleCadena.Alignment = HorizontalAlignment.Distributed;
            }
            if ("LEFT".Equals(alineamiento))
            {
                styleCadena.Alignment = HorizontalAlignment.Left;
            }
            if ("RIGHT".Equals(alineamiento))
            {
                styleCadena.Alignment = HorizontalAlignment.Left;
            }



            styleCadena.SetFont(fontCadena);

            if (vColorCeldaCabecera)
            {
                fontCadena.Color = HSSFColor.White.Index;
                styleCadena.FillForegroundColor = IndexedColors.Red.Index;
                styleCadena.FillPattern = FillPattern.SolidForeground;
            }



            return styleCadena;
        }



        /// <summary>
        /// Estilo para numeros
        /// </summary>
        /// <param name="negrita"></param>
        /// <param name="point"></param>
        /// <param name="alineamiento"></param>
        /// <returns></returns>
        public ICellStyle addEstiloNumero(Boolean negrita, int point, String alineamiento)
        {

            IFont fontNumero = libro.CreateFont();
            ICellStyle styleNumero = libro.CreateCellStyle();

            fontNumero.Color = HSSFColor.Black.Index;

            if (negrita)
            {
                fontNumero.Boldweight = (short)FontBoldWeight.Bold;
            }


            fontNumero.FontName = HSSFFont.FONT_ARIAL;

            fontNumero.FontHeightInPoints = (short)point;

            if ("CENTER".Equals(alineamiento))
            {
                styleNumero.Alignment = HorizontalAlignment.Center;
            }
            if ("CENTER_SELECTION".Equals(alineamiento))
            {
                styleNumero.Alignment = HorizontalAlignment.CenterSelection;
            }
            if ("DISTRIBUTED".Equals(alineamiento))
            {
                styleNumero.Alignment = HorizontalAlignment.Distributed;
            }
            if ("LEFT".Equals(alineamiento))
            {
                styleNumero.Alignment = HorizontalAlignment.Left;
            }
            if ("RIGHT".Equals(alineamiento))
            {
                styleNumero.Alignment = HorizontalAlignment.Right;
            }

            if ("AJUSTA".Equals(alineamiento))
            {
                styleNumero.Alignment = HorizontalAlignment.Left;
                styleNumero.VerticalAlignment = VerticalAlignment.Justify;

            }

            styleNumero.SetFont(fontNumero);

            return styleNumero;
        }

        /// <summary>
        /// Estilo para el titulo
        /// </summary>
        /// <param name="negrita"></param>
        /// <param name="point"></param>
        /// <param name="alineamiento"></param>
        /// <returns></returns>
        public ICellStyle addEstiloTitulo(Boolean negrita, int point, String alineamiento, bool vColorCeldaCabecera = false)
        {

            IFont fontTitulo = libro.CreateFont();
            ICellStyle styleTitulo = libro.CreateCellStyle();

            fontTitulo.Color = HSSFColor.Black.Index;
            if (negrita)
            {
                fontTitulo.Boldweight = (short)FontBoldWeight.Bold;
            }

            fontTitulo.FontName = HSSFFont.FONT_ARIAL;

            fontTitulo.FontHeightInPoints = (short)point;

            if ("CENTER".Equals(alineamiento))
            {
                styleTitulo.Alignment = HorizontalAlignment.Center;
            }
            if ("CENTER_SELECTION".Equals(alineamiento))
            {
                styleTitulo.Alignment = HorizontalAlignment.CenterSelection;
            }
            if ("DISTRIBUTED".Equals(alineamiento))
            {
                styleTitulo.Alignment = HorizontalAlignment.Distributed;
            }
            if ("LEFT".Equals(alineamiento))
            {
                styleTitulo.Alignment = HorizontalAlignment.Left;
            }
            if ("RIGHT".Equals(alineamiento))
            {
                styleTitulo.Alignment = HorizontalAlignment.Left;
            }

            styleTitulo.SetFont(fontTitulo);

            if (vColorCeldaCabecera)
            {
                fontTitulo.Color = HSSFColor.White.Index;
                styleTitulo.FillForegroundColor = IndexedColors.Red.Index;
                styleTitulo.FillPattern = FillPattern.SolidForeground;
            }

            return styleTitulo;
        }


        /// <summary>
        /// estilo negrita
        /// </summary>
        /// <param name="point"></param>
        /// <param name="alineamiento"></param>
        /// <returns></returns>
        public ICellStyle addEstiloCadenaNegrita(int point, String alineamiento)
        {

            IFont fontCadenaNegrita = libro.CreateFont();
            ICellStyle styleCadenaNegrita = libro.CreateCellStyle();

            fontCadenaNegrita.Color = HSSFColor.Black.Index;
            fontCadenaNegrita.Boldweight = (short)FontBoldWeight.Bold;



            fontCadenaNegrita.FontName = HSSFFont.FONT_ARIAL;

            fontCadenaNegrita.FontHeightInPoints = (short)point;

            if ("CENTER".Equals(alineamiento))
            {
                styleCadenaNegrita.Alignment = HorizontalAlignment.Center;
            }
            if ("CENTER_SELECTION".Equals(alineamiento))
            {
                styleCadenaNegrita.Alignment = HorizontalAlignment.CenterSelection;
            }
            if ("DISTRIBUTED".Equals(alineamiento))
            {
                styleCadenaNegrita.Alignment = HorizontalAlignment.Distributed;
            }
            if ("LEFT".Equals(alineamiento))
            {
                styleCadenaNegrita.Alignment = HorizontalAlignment.Left;
            }
            if ("RIGHT".Equals(alineamiento))
            {
                styleCadenaNegrita.Alignment = HorizontalAlignment.Left;
            }

            styleCadenaNegrita.SetFont(fontCadenaNegrita);

            return styleCadenaNegrita;
        }

        /// <summary>
        /// imprime Excel
        /// </summary>
        /// <param name="exportData"></param>
        /// <returns></returns>
        public MemoryStream imprimeExcel(MemoryStream exportData)
        {
            libro.Write(exportData);
            return exportData;
        }



        public FileStream imprimeExcelFile(FileStream exportData)
        {
            libro.Write(exportData);
            return exportData;
        }
    }
}
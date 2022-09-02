using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Bsp.CommonServices.Areas.ReportData.ExcelHelps;
using Bsp.CommonServices.Areas.ReportData.Models;
using NPOI.SS.UserModel;

namespace Bsp.CommonServices.Areas.ReportData.Helps
{
    public class GenerateExcelFile
    {

        public int ExportaExcelDetalle(string IsExportExcelDetalle, string ReporteSelecc, string FechaVal, string FechaText,
            string CodigoEmpresa, string EmpresaText, string DescripcionReporte, DataResponse data
            )
        {
            string vTitulo = "";
            long numero = 0;
            string ExportaExcel = IsExportExcelDetalle;
            var vReporte = ReporteSelecc;
            int resp = 0;

            if (vReporte != null && vReporte != "")
            {

                if (ExportaExcel != null)
                {
                    bool isExportExcel;

                    if (bool.TryParse(ExportaExcel, out isExportExcel))
                    {
                        if (isExportExcel)
                        {
                            string vFechaVal;
                            string vFechaText;
                            string vCodigoEmpresa;
                            string vNombreEmpresaText;
                            string vNombreReporte;

                            List<Header> listaCabecera;
                            listaCabecera = new List<Header>();

                            vFechaVal = FechaVal;
                            vFechaText = FechaText;
                            vCodigoEmpresa = CodigoEmpresa;
                            vNombreEmpresaText = EmpresaText;
                            vNombreReporte = DescripcionReporte;

                            ICellStyle styleTitulo, styleCadena, styleNegrita, styleNumero;
                            ICellStyle styleNegritaColor;
                            IRow row;

                            ReporteExcel objGeneraExcel;
                            objGeneraExcel = new ReporteExcel();
                            //se crea la hoja
                            objGeneraExcel.creaHoja("Reporte", "S");
                            //se agregan los estilos
                            styleTitulo = objGeneraExcel.addEstiloTitulo(true, 14, "CENTER", true);
                            styleCadena = objGeneraExcel.addEstiloCadena(false, 10, "LEFT");
                            styleNegrita = objGeneraExcel.addEstiloCadenaNegrita(10, "LEFT");
                            styleNumero = objGeneraExcel.addEstiloNumero(false, 10, "RIGHT");

                            styleNegritaColor = objGeneraExcel.addEstiloCadena(true, 10, "LEFT", true);

                            int fila = 0;
                            int columna = 0;


                            if (vReporte.Equals(Constantes.CODIGO_REPORTE_CREDITICIO_DEUDORES.ToString()))
                            {
                                vTitulo = Constantes.EXCEL_TITULO_CREDITICIO_DE_DEUDORES_SISTEMA_FINANCIERO;
                                listaCabecera = data.Header.ToList(); // PortalBL.obtieneCabecera(objDeudor);
                                //cantidad de columnas
                                row = objGeneraExcel.addFila(fila++);
                                //se crea el titulo del reporte
                               
                                objGeneraExcel.addCellMerge(row, 0, 0, 0, (listaCabecera.Count() - 1), Constantes.TITULO_REPORTE_CREDITICION_DEUDORES_LEASING, styleNegritaColor);
                                row = objGeneraExcel.addFila(fila++);

                                objGeneraExcel.addCelda(row, 0, "Producto:  Leasing / Lease Back", styleCadena, "S");
                                row = objGeneraExcel.addFila(fila++);
                                objGeneraExcel.addCelda(row, 0, "Al: " + vFechaText, styleCadena, "S");

                                row = objGeneraExcel.addFila(fila++);
                                foreach (var item in listaCabecera)
                                {
                                    objGeneraExcel.addCelda(row, columna, item.TituloCabecera, styleNegritaColor, "S");
                                    columna++;
                                }

                                //ini comentado temporalmente
                                //List<ReporteDeudorProducto> listaDetalle;
                                //listaDetalle = new List<ReporteDeudorProducto>();

                                //listaDetalle = PortalBL.obtieneLeasing(objDeudor);

                                //foreach (var item in listaDetalle)
                                //{
                                //    columna = 0;
                                //    row = objGeneraExcel.addFila(fila++);
                                //    objGeneraExcel.addCelda(row, columna, item.CodigoSbsCliente, styleCadena, "S");
                                //    columna++;
                                //    objGeneraExcel.addCelda(row, columna, item.NumeroDocumentoTributario, styleCadena, "S");
                                //    columna++;
                                //    objGeneraExcel.addCelda(row, columna, item.RazonSocial, styleCadena, "S");
                                //    //columna++;
                                //    //objGeneraExcel.addCelda(row, columna, item.CodigoEmpresa, styleCadena, "S");
                                //    columna++;
                                //    objGeneraExcel.addCelda(row, columna, item.NombreBanco, styleCadena, "S");
                                //    //columna++;
                                //    //objGeneraExcel.addCelda(row, columna, item.CodigoCuentaContable, styleCadena, "S");
                                //    //columna++;
                                //    //objGeneraExcel.addCelda(row, columna, item.DescripcionCuenta, styleCadena, "S");
                                //    columna++;
                                //    objGeneraExcel.addCelda(row, columna, item.Saldo, styleCadena, "S");
                                //    //columna++;
                                //    //objGeneraExcel.addCelda(row, columna, item.IdTipoCredito, styleCadena, "S");
                                //    columna++;
                                //    objGeneraExcel.addCelda(row, columna, item.DesTipoCredito, styleCadena, "S");
                                //    columna++;
                                //    objGeneraExcel.addCelda(row, columna, item.IdClasificacionSBS, styleCadena, "S");
                                //    columna++;
                                //    objGeneraExcel.addCelda(row, columna, item.Condicion, styleCadena, "S");

                                //}


                                //fin comentado temporalmente
                            }
                            else
                            {
                                vTitulo = SetTitulo(vTitulo, vReporte);
                                listaCabecera = data.Header.ToList();
                                row = objGeneraExcel.addFila(fila++);                                
                                objGeneraExcel.addCellMerge(row, 0, 0, 0, (listaCabecera.Count() - 1), (vNombreReporte + " " + vFechaText), styleNegritaColor);

                                row = objGeneraExcel.addFila(fila++);
                                foreach (var item in listaCabecera)
                                {
                                    objGeneraExcel.addCelda(row, columna, item.TituloCabecera, styleNegritaColor, "S");
                                    columna++;
                                }

                                SetDataDetalle(data, ref numero, vReporte, styleCadena, styleNumero, ref row, objGeneraExcel, ref fila);

                            }

                            for (int i = 0; i < listaCabecera.Count(); i++)
                            {
                                objGeneraExcel.hoja.AutoSizeColumn(i);
                            }

                            string NombreArchivoExcel = @"D:\Archivos\"+ vTitulo + "_" + data.FechaReporteNumerico + ".xls";
                            using (FileStream file = new FileStream(NombreArchivoExcel, FileMode.Create, FileAccess.ReadWrite))
                            {
                                objGeneraExcel.imprimeExcelFile(file).Close();
                                resp = 1;
                            }
                        }

                    }
                }
                vReporte = null;
            }
            return resp;
        }

        private static void SetDataDetalle(DataResponse data, ref long numero, string vReporte, ICellStyle styleCadena, ICellStyle styleNumero, ref IRow row, ReporteExcel objGeneraExcel, ref int fila)
        {
            var nroFilas = data.Details.GroupBy(x => x.Fila).Select(grp => grp.First()).ToList();

            var lst = data.Details.OrderBy(x => x.Fila).ToArray();

            for (int i = 0; i < nroFilas.Count(); i++)
            {
                row = objGeneraExcel.addFila(fila++);

                var q = (from c in lst where c.Fila == i + 1 select new { col = c.Columna, val = c.Valor }).ToList();

                foreach (var r in q)
                {
                    var cadena = r.val.Replace(",", "");
                    var col = Convert.ToInt32(r.col) - 1;
                    var isNumeric = long.TryParse(cadena, out numero);

                    if (isNumeric)
                    {
                        if (vReporte.Equals(Constantes.Reporte_25.ToString()))
                        {
                            if (col == 5)
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "S");
                            else
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "N");

                        }
                        else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_DOS_O_MAS_TIPOS_DE_CREDITO.ToString()))
                        {
                            if (col == 3)
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "S");
                            else
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "N");

                        }
                        else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_REFINANCIADOS.ToString())
                            || vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_VENCIDOS.ToString())
                            || vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_JUDICIALES.ToString())
                            || vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_CASTIGADOS.ToString())
                            || vReporte.Equals(Constantes.CODIGO_DEUDORES_SALDOS_RENDIMIENTOS_DE_CREDITO_Y_RENTAS_EN_SUSPENSO.ToString())
                            )
                        {
                            if (col == 3)
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "S");
                            else
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "N");

                        }
                        else if (vReporte.Equals(Constantes.Reporte_18.ToString()))
                        {
                            if (col == 2)
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "S");
                            else
                                objGeneraExcel.addCelda(row, col, cadena, styleNumero, "N");
                        }
                        else
                            objGeneraExcel.addCelda(row, col, cadena, styleNumero, "N");

                    }
                    else
                        objGeneraExcel.addCelda(row, col, cadena, styleCadena, "S");

                }
            }
        }

        private static string SetTitulo(string vTitulo, string vReporte)
        {
            if (vReporte.Equals(Constantes.CODIGO_REPORTE_DEUDAS.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_INFORMACION_DEUDAS;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_CLASIFICACION.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_CLASIFICACION;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_DOS_O_MAS_TIPOS_DE_CREDITO.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_DOS_O_MAS_TIPOS_DE_CREDITO;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_BCP_BBVA_SCOTIABANK.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_BCP_BBVA_SCOTIABANK;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_REFINANCIADOS.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_CON_CREDITOS_REFINANCIADOS;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_VENCIDOS.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_CON_CREDITOS_VENCIDOS;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_JUDICIALES.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_CON_CREDITOS_JUDICIALES;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_CON_CREDITOS_CASTIGADOS.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_CON_CREDITOS_CASTIGADOS;
            }
            else if (vReporte.Equals(Constantes.CODIGO_DEUDORES_SALDOS_RENDIMIENTOS_DE_CREDITO_Y_RENTAS_EN_SUSPENSO.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_SALDOS_RENDIMIENTOS_DE_CREDITO_Y_RENTAS_EN_SUSPENSO;
            }

            else if (vReporte.Equals(Constantes.Reporte_18.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_INCREMENTO_LEASEBACK_RESPECTO_MES_ANTERIOR;
            }
            else if (vReporte.Equals(Constantes.Reporte_19.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_BANCO_SANTANDER_PARTICIPACIÓN_MAYOR_20_DEL_TOTAL_SALDO;
            }
            else if (vReporte.Equals(Constantes.Reporte_20.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_INCREMENTO_DEUDA_DIRECTA_Y_DISMINUCIÓN_DEUDA_INDIRECTA;
            }
            else if (vReporte.Equals(Constantes.Reporte_21.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_PRODUCTO_SOBREGIRO_MAYOR_5_DEUDA_DIRECTA;
            }
            else if (vReporte.Equals(Constantes.Reporte_22.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_INCREMENTO_MAYOR_A_20_RESPECTO_MES_ANTERIOR_Y_MAYOR_50_12_MESES;
            }
            else if (vReporte.Equals(Constantes.Reporte_23.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_INCREMENTO_MAYOR_A_20_PRESTAMO_MES_ANTERIOR_Y_MAYOR_50_12_MESES;
            }
            else if (vReporte.Equals(Constantes.Reporte_24.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_INCREMENTO_MAYOR_A_20_GT_PREFERIDAS_MES_ANTERIOR_Y_MAYOR_50_12_MESES;
            }
            else if (vReporte.Equals(Constantes.Reporte_25.ToString()))
            {
                vTitulo = Constantes.EXCEL_TITULO_DEUDORES_DIAS_DE_ATRASO;
            }

            return vTitulo;
        }

        //private static JsonGrid SetDataFormat(List<Details> listaIA)
        //{
        //    var objOrdenado = new ExpandoObject() as IDictionary<string, Object>;
        //    dynamic objFila = new ExpandoObject();

        //    List<dynamic> listaOrdenada = new List<dynamic>();
        //    int contColumnas = 0;
        //    string ColumnaOrdena = Constantes.COLUMNA_DETALLE;
        //    int contFilas = 0;
        //    int cantidadColumnas = 0;
        //    int nFilas = 0;
        //    if (listaIA != null)
        //    {
        //        if (listaIA.Count > 0)
        //        {
        //            var listaFilas = listaIA.GroupBy(x => x.Fila).Select(grp => grp.First()).ToList();

        //            if (listaFilas != null && listaFilas.Count > 0)
        //            {
        //                foreach (var itemFila in listaFilas)
        //                {
        //                    nFilas++;
        //                    var listaGrupos = listaIA.Where(a => a.Fila == itemFila.Fila).OrderBy(p => p.Columna).ToList();
        //                    if (nFilas == 1)
        //                    {
        //                        cantidadColumnas = listaGrupos.Count();
        //                    }

        //                    if (listaGrupos.Count() == cantidadColumnas)
        //                    {
        //                        contColumnas = 0;
        //                        objOrdenado = new ExpandoObject() as IDictionary<string, Object>;

        //                        foreach (var itemGrupo in listaGrupos)
        //                        {
        //                            if (contColumnas == 0)
        //                            {
        //                                contFilas++;
        //                                ColumnaOrdena = Constantes.COLUMNA_DETALLE + contColumnas;
        //                                objOrdenado.Add(ColumnaOrdena, contFilas);
        //                            }
        //                            contColumnas++;
        //                            ColumnaOrdena = Constantes.COLUMNA_DETALLE;
        //                            ColumnaOrdena = ColumnaOrdena + contColumnas;
        //                            objOrdenado.Add(ColumnaOrdena, itemGrupo.Valor);
        //                        }
        //                        listaOrdenada.Add(objOrdenado);
        //                    }

        //                }

        //            }
        //        }
        //    }


        //    return GridData(1, 1, listaOrdenada.Count(), listaOrdenada);
        //}

        //private static JsonGrid GridData(int pPageCount, int pCurrentPage, int pRecordCount, List<dynamic> listaInfo)
        //{
        //    JsonGrid objGridJonResponse;
        //    objGridJonResponse = new JsonGrid();

        //    objGridJonResponse.PageCount = pPageCount;
        //    objGridJonResponse.CurrentPage = pCurrentPage;
        //    objGridJonResponse.RecordCount = pRecordCount;
        //    objGridJonResponse.Items = new List<GridItem>();

        //    var _items = new List<GridItem>();


        //    int cantidadColumanas = 0;
        //    string nombreColumna = Constantes.COLUMNA_DETALLE;
        //    GridItem objJQGridItem;
        //    objJQGridItem = new GridItem();
        //    string valorColumna = string.Empty;
        //    int cantidadColumnas = listaInfo.Count();
        //    List<string> listaDatos;
        //    listaDatos = new List<string>();
        //    Object sDato;
        //    sDato = new Object();
        //    int nId = 0;
        //    foreach (var item in listaInfo)
        //    {

        //        cantidadColumanas = 0;
        //        if (cantidadColumanas == 0)
        //        {
        //            nombreColumna = Constantes.COLUMNA_DETALLE;
        //            nombreColumna = nombreColumna + "" + cantidadColumanas.ToString();

        //            valorColumna = "";

        //            var objDict = ((IDictionary<String, Object>)item);
        //            sDato = new Object();
        //            if (objDict.TryGetValue(nombreColumna, out sDato))
        //            {
        //                nId = Convert.ToInt32(sDato);
        //            }

        //        }
        //        cantidadColumanas = cantidadColumanas + 1;
        //        if (cantidadColumanas > 0)
        //        {
        //            var objDict2 = ((IDictionary<String, Object>)item);
        //            listaDatos = new List<string>();
        //            for (int i = 1; i <= objDict2.Count(); i++)
        //            {
        //                nombreColumna = Constantes.COLUMNA_DETALLE;
        //                nombreColumna = nombreColumna + "" + i.ToString();
        //                valorColumna = "";
        //                sDato = new Object();
        //                if (objDict2.TryGetValue(nombreColumna, out sDato))
        //                {
        //                    listaDatos.Add(sDato.ToString());
        //                }

        //            }

        //        }
        //        _items.Add(new GridItem(nId, listaDatos));


        //    }
        //    objGridJonResponse.Items = _items;

        //    return objGridJonResponse;
        //}

    }
}
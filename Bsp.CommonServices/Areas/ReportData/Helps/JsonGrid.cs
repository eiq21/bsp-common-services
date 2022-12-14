using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bsp.CommonServices.Areas.ReportData.Helps
{
    public class JsonGrid
    {
        #region Passive attributes.

        private int _pageCount;
        private int _currentPage;
        private int _recordCount;
        private List<GridItem> _items;

        #endregion

        #region Properties

        /// <summary>
        /// Cantidad de páginas del JQGrid.
        /// </summary>
        public int PageCount
        {
            get { return _pageCount; }
            set { _pageCount = value; }
        }
        /// <summary>
        /// Página actual del JQGrid.
        /// </summary>
        public int CurrentPage
        {
            get { return _currentPage; }
            set { _currentPage = value; }
        }
        /// <summary>
        /// Cantidad total de elementos de la lista.
        /// </summary>
        public int RecordCount
        {
            get { return _recordCount; }
            set { _recordCount = value; }
        }
        /// <summary>
        /// Lista de elementos del JQGrid.
        /// </summary>
        public List<GridItem> Items
        {
            get { return _items; }
            set { _items = value; }
        }

        #endregion

    }
}
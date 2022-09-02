using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bsp.CommonServices.Areas.ReportData.Helps
{
    public class GridItem
    {

        #region Passive attributes

        private long _id;
        private List<string> _row;

        //private List<Object> _row2;
        #endregion

        #region Properties

        /// <summary>
        /// RowId de la fila.
        /// </summary>
        public long ID
        {
            get { return _id; }
            set { _id = value; }
        }
        /// <summary>
        /// Fila del JQGrid.
        /// </summary>
        public List<string> Row
        {
            get { return _row; }
            set { _row = value; }
        }


        //public List<Object> Row2
        //{
        //    get { return _row2; }
        //    set { _row2 = value; }
        //}
        #endregion

        #region Active Attributes

        /// <summary>
        /// Contructor.
        /// </summary>
        public GridItem(long pId, List<string> pRow)
        {
            _id = pId;
            _row = pRow;
        }
        public GridItem()
        {
        }

        #endregion
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelApplicationFindAndCopy
{
 
    class myExcelRange
    {
        public int rowStart, columnStart, rowEnd, columnEnd;    // 1,1 -- 1,5
        public string adressString;                             // "A1:E1";
        public string adressRow;                                // 1:1
        public string adressColumn;                             // A:A

        public myExcelRange(int rowStart, int columnStart, int rowEnd, int columnEnd, myExcelAdressMaker adrMaker)
        {
            this.rowStart = rowStart;
            this.columnStart = columnStart;
            this.rowEnd = rowEnd;
            this.columnEnd = columnEnd;
            adressString = adrMaker.getCellStringAdress(rowStart, columnStart) + ":" + adrMaker.getCellStringAdress(rowEnd, columnEnd);
            adressRow = rowStart + ":" + rowStart;
            adressColumn = adrMaker.getLetter(columnStart) + ":" + adrMaker.getLetter(columnStart);
        }

    }
}

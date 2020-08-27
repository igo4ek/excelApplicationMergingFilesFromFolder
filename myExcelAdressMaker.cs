using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelApplicationFindAndCopy
{
    class myExcelAdressMaker
    {
        private List<String> list;
        public myExcelAdressMaker()
        {
            this.list = new List<String>();
            #region Формируем список, где начиная с нуля вбиты все буквенные адреса столбцов
            this.list.Add("A");
            this.list.Add("B");
            this.list.Add("C");
            this.list.Add("D");
            this.list.Add("E");
            this.list.Add("F");
            this.list.Add("G");
            this.list.Add("H");
            this.list.Add("I");
            this.list.Add("J");
            this.list.Add("K");
            this.list.Add("L");
            this.list.Add("M");
            this.list.Add("N");
            this.list.Add("O");
            this.list.Add("P");
            this.list.Add("Q");
            this.list.Add("R");
            this.list.Add("S");
            this.list.Add("T");
            this.list.Add("U");
            this.list.Add("V");
            this.list.Add("W");
            this.list.Add("X");
            this.list.Add("Y");
            this.list.Add("Z");

            this.list.Add("AA");
            this.list.Add("AB");
            this.list.Add("AC");
            this.list.Add("AD");
            this.list.Add("AE");
            this.list.Add("AF");
            this.list.Add("AG");
            this.list.Add("AH");
            this.list.Add("AI");
            this.list.Add("AJ");
            this.list.Add("AK");
            this.list.Add("AL");
            this.list.Add("AM");
            this.list.Add("AN");
            this.list.Add("AO");
            this.list.Add("AP");
            this.list.Add("AQ");
            this.list.Add("AR");
            this.list.Add("AS");
            this.list.Add("AT");
            this.list.Add("AU");
            this.list.Add("AV");
            this.list.Add("AW");
            this.list.Add("AX");
            this.list.Add("AY");
            this.list.Add("AZ");

            this.list.Add("BA");
            this.list.Add("BB");
            this.list.Add("BC");
            this.list.Add("BD");
            this.list.Add("BE");
            this.list.Add("BF");
            this.list.Add("BG");
            this.list.Add("BH");
            this.list.Add("BI");
            this.list.Add("BJ");
            this.list.Add("BK");
            this.list.Add("BL");
            this.list.Add("BM");
            this.list.Add("BN");
            this.list.Add("BO");
            this.list.Add("BP");
            this.list.Add("BQ");
            this.list.Add("BR");
            this.list.Add("BS");
            this.list.Add("BT");
            this.list.Add("BU");
            this.list.Add("BV");
            this.list.Add("BW");
            this.list.Add("BX");
            this.list.Add("BY");
            this.list.Add("BZ");

            this.list.Add("CA");
            this.list.Add("CB");
            this.list.Add("CC");
            this.list.Add("CD");
            this.list.Add("CE");
            this.list.Add("CF");
            this.list.Add("CG");
            this.list.Add("CH");
            this.list.Add("CI");
            this.list.Add("CJ");
            this.list.Add("CK");
            this.list.Add("CL");
            this.list.Add("CM");
            this.list.Add("CN");
            this.list.Add("CO");
            this.list.Add("CP");
            this.list.Add("CQ");
            this.list.Add("CR");
            this.list.Add("CS");
            this.list.Add("CT");
            this.list.Add("CU");
            this.list.Add("CV");
            this.list.Add("CW");
            this.list.Add("CX");
            this.list.Add("CY");
            this.list.Add("CZ");

            this.list.Add("DA");
            this.list.Add("DB");
            this.list.Add("DC");
            this.list.Add("DD");
            this.list.Add("DE");
            this.list.Add("DF");
            this.list.Add("DG");
            this.list.Add("DH");
            this.list.Add("DI");
            this.list.Add("DJ");
            this.list.Add("DK");
            this.list.Add("DL");
            this.list.Add("DM");
            this.list.Add("DN");
            this.list.Add("DO");
            this.list.Add("DP");
            this.list.Add("DQ");
            this.list.Add("DR");
            this.list.Add("DS");
            this.list.Add("DT");
            this.list.Add("DU");
            this.list.Add("DV");
            this.list.Add("DW");
            this.list.Add("DX");
            this.list.Add("DY");
            this.list.Add("DZ");
            #endregion
        }
        public string getCellStringAdress(int row, int column)
        {
            return this.list[column-1] + row;
        }
        public string getLetter(int column)
        {
            return this.list[column - 1];
        }
    }
}


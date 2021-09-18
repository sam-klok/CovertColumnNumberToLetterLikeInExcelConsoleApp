using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {
        public const int alphabetLength = 26;
        public static readonly char[] letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        static void Main(string[] args)
        {
            // convert number to letters (like in spreadsheet Excel)
            // 1 - A
            // 2 - B
            // 27 - AA
            // 28 - AB
            TestConversionFromStackOverflow();



            // convert number to letters (like in spreadsheet Excel)
            // 1 - A
            // 2 - B
            // 27 - AA
            // 28 - AB

            TestMyConversionStart1();

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("**************************** start from 0 **************************");

            // convert number to letters (like in spreadsheet Excel)
            // 0 - A
            // 1 - B
            // 26 - AA
            // 27 - AB
            TestMyConversionStarts0();

            Console.WriteLine("Press any key..");
            Console.ReadKey();
        }


        private static void TestConversionFromStackOverflow()
        {
            for (int i = 1; i < 800; i++)
            {
                var columnName = GetExcelColumnName(i);
                Console.Write($"{i} - {columnName}; ");

                //let's separate each alphabet iteration
                if (i % alphabetLength == 0)
                {
                    Console.WriteLine();
                    Console.WriteLine();
                }

                //let's separate at triple letters
                if (i % (alphabetLength * (alphabetLength + 1)) == 0)
                    Console.WriteLine("===================================================");
            }
        }

        private static void TestMyConversionStart1()
        {
            for (int i = 1; i < 800; i++)
            {
                var columnName = ConvertToLetters(i);
                Console.Write($"{i} - {columnName}; ");

                //let's separate each alphabet iteration
                if (i % alphabetLength == 0)
                {
                    Console.WriteLine();
                    Console.WriteLine();
                }

                //let's separate at triple letters
                if (i % (alphabetLength * (alphabetLength + 1)) == 0)
                    Console.WriteLine("===================================================");
            }
        }

        private static void TestMyConversionStarts0()
        {
            for (int i = 0; i < 800; i++)
            {
                var columnName = ConvertToLettersStart0(i);
                Console.Write($"{i} - {columnName}; ");

                // let's separate each alphabet iteration
                if ((i + 1) % alphabetLength == 0)
                {
                    Console.WriteLine();
                    Console.WriteLine();
                }

                //let's separate at triple letters
                if ((i + 1) % (alphabetLength * (alphabetLength + 1)) == 0)
                    Console.WriteLine("===================================================");
            }
        }


        // I used this function to figure out division conversions
        static int ConvertToGroups(int columnNumber)
        {
            var groupNumber = columnNumber / alphabetLength;
            Console.Write($"{columnNumber} => {groupNumber}; ");
            return groupNumber;
        }


        //Solution from Stack Overflow
        // https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
        static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int remainder = (columnNumber - 1) % 26;

                //columnName = Convert.ToChar('A' + remainder) + columnName;    // first option
                columnName = letters[remainder] + columnName;                   // second option

                columnNumber = (columnNumber - remainder) / 26;
            }

            return columnName;
        }

        // not most efficient solution, but I did it myself
        // solution for column numbers starts 1
        // 1 2 3
        static string ConvertToLetters(int columnNumber, string leftLetters = "")
        {
            var columnName = new StringBuilder().Append(leftLetters);

            // because coulumns numbered starting 1, but array starts with 0
            // let's use "-1"

            if (columnNumber == 0)
                return leftLetters + "Z";

            if (columnNumber / alphabetLength == 0 ||
                (columnNumber / alphabetLength == 1 && columnNumber % alphabetLength == 0))
            {
                columnName.Append(letters[columnNumber - 1]);
                return columnName.ToString();
            }
            else
            {
                int leftLetterIndex = (columnNumber - 1) / alphabetLength;
                int restOfNumber = columnNumber - (leftLetterIndex * alphabetLength);

                var newleftLetters = ConvertToLetters(leftLetterIndex, leftLetters);
                var tempLetters = ConvertToLetters(restOfNumber, newleftLetters);
                return tempLetters;
            }

        }


        // solution for column numbers starts 0
        // 0 1 2
        static string ConvertToLettersStart0(int columnNumber, string leftLetters = "")
        {
            var columnName = new StringBuilder().Append(leftLetters);

            //if (columnNumber == alphabetLength)  // 26
            //    return leftLetters + "A";

            if (columnNumber / alphabetLength == 0)
                //||
                //(columnNumber / alphabetLength == 1 && columnNumber % alphabetLength == 0))
            {
                columnName.Append(letters[columnNumber]);
                return columnName.ToString();
            }
            else
            {
                int leftLetterIndex = columnNumber / alphabetLength - 1;
                int restOfNumber = columnNumber - ((leftLetterIndex + 1) * alphabetLength);

                var newleftLetters = ConvertToLettersStart0(leftLetterIndex, leftLetters);
                var tempLetters = ConvertToLettersStart0(restOfNumber, newleftLetters);
                return tempLetters;
            }

        }
    }



}

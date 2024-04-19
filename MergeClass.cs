using Aron.Sinoai.OfficeHelper;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using RGiesecke.DllExport;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelMerge
{
    public class MergeClass
    {
        //1: (text) file name with tokens and their values(this file can be exported from a clarion app.)
        //2: xlsx template filename
        //3: target filename(after merge)
        //(4: executable filename(ie excel. exe) to open 3)

        private static String GENERATED_FILE_NAME;
        //private static String TEMPLATE_FILE_NAME;

        public static IEnumerable<List<object>> GetSample()
        {
            var random = new Random();

            for (int loop = 0; loop < 3000; loop++)
            {
                yield return new List<object> { "test", DateTime.Now.AddDays(random.NextDouble() * 100 - 50), loop };
            }

        }

        [DllExport("CreateGeneratedFile", CallingConvention = CallingConvention.StdCall)]
        public static void CreateGeneratedFile(string pTemplateFile,string pGeneratedFile)
        {
            GENERATED_FILE_NAME = String.Concat(pGeneratedFile) + @"\generated.xlsx";

            Trace.WriteLine("[CreateGeneratedFile]GENERATED_FILE_NAME[" + String.Concat(GENERATED_FILE_NAME) + "]");

            using (ExcelHelper helper = new ExcelHelper(pTemplateFile, GENERATED_FILE_NAME))
            {
                helper.Direction = ExcelHelper.DirectionType.TOP_TO_DOWN;

                helper.CurrentSheetName = "Sheet1";

                helper.CurrentPosition = new CellRef("C3");

                //the template xlsx should contain the named range "header"; use the command "insert"/"name".
                helper.InsertRange("header");

                //the template xlsx should contain the named range "sample1";
                //inside this range you should have cells with these values:
                //<name> , <value> and <comment>, which will be replaced by the values from the getSample()
                CellRangeTemplate sample1 = helper.CreateCellRangeTemplate("sample1", new List<string> { "name", "value", "comment" });

                helper.InsertRange(sample1, GetSample());

                //you could use here other named ranges to insert new cells and call InsertRange as many times you want, 
                //it will be copied one after another;
                //even you can change direction or the current cell/sheet before you insert

                //tipically you put all your "template ranges" (the names) on the same sheet and then you just delete it
                helper.DeleteSheet("Sheet3");
            }
        }


    }
}

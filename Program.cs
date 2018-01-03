using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.IO;
using System.Configuration;

namespace BinaryCompile
{
    class Program
    {
        public static void Main(string[] args)
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excel.Workbooks.Add();

            var project = workbook.VBProject;

            String output_path = Path.GetFullPath(".\\");

            include_files(get_vba_files(), ref project);

            workbook.Application.Visible = false;

            workbook.SaveAs(output_path + "PERSONAL.xlsb", XlFileFormat.xlExcel12);
            workbook.Close();
            
            excel.Quit();

        }

        static string[] get_vba_files()
        {
            String source_path = Path.GetFullPath("src\\");
            return Directory.GetFiles(source_path, "*.vba");
        }

        static void include_files(string[] vba_files, ref VBProject project)
        {
            foreach (string source in vba_files)
            {
                var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                module.CodeModule.AddFromFile(source);
                module.Name = Path.GetFileNameWithoutExtension(source);
                
            }
        }
    }
}

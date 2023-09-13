using QRGenerator.Entities;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Drawing.Imaging;




namespace QRGenerator
{
    internal class Program
    {
        #region properties

        private const string excelPath = "D:\\Users\\Descargas\\Calificaciones (2)";

        private static readonly Microsoft.Office.Interop.Excel.Application excelApp = new();
        private static Workbook? workBook;
        private static Worksheet? WorkSheet;

        #endregion

        static void Main(string[] args)
        {
            Console.WriteLine("Iniciando Programa");
            StudentsCards cards = new StudentsCards();
            List<Student> students = GetStudents();

            int row = 1;
            foreach (Student student in students)
            {
                if (cards.CreateCard(student))
                {
                    Console.WriteLine("Creacion de credencial exitosa para: {0}", student.Name);
                }
                else 
                {
                    Console.WriteLine("Error al crear la credencial para: {0}", student.Name);                 
                }               
            }

            CloseWorkBook();

        }
        internal static List<Student> GetStudents()
        {
            Console.WriteLine("Obteniendo Estudiantes.....");
            List<Student> students = new();

            //Get data from excel and store it in a list of students.

            workBook = excelApp.Workbooks.Open(excelPath);
            WorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            Console.WriteLine("Lista de estudiantes: ");
            int j = 0;
            for (int i = 2; i <= WorkSheet.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(WorkSheet.Cells[i, "A"].Text))
                {
                    students.Add(new()
                    {
                        Id = WorkSheet.Cells[i, "A"].Text,
                        Name = WorkSheet.Cells[i, "B"].Text,
                        Age = WorkSheet.Cells[i, "C"].Text,
                        Grade = WorkSheet.Cells[i, "D"].Text,
                        Gender = WorkSheet.Cells[i, "E"].Text
                    });

                    Console.WriteLine(students[j].Name);
                    j++;
                }
                else
                {
                    break;
                }

            }

            return students;
        }
        internal static void CloseWorkBook()
        {

            workBook.Close();
            excelApp.Workbooks.Close();
            excelApp.Quit();
        }


    }
}
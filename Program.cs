using QRGenerator.Entities;
using QRCoder;
using System.Drawing;

namespace QRGenerator
{
    internal class Program
    {
        const string path = "D:\\Users\\Descargas\\Calificaciones (2)";
        static void Main(string[] args)
        {
            Console.WriteLine("Iniciando Programa");
             List<Student> students = GetStudents();
            int row = 1;
            foreach (Student student in students)
            {
                
                student.HasQr = GenerateQR(student);
                if (student.HasQr)
                {
                    UpdateExcel(student, row);
                    Console.WriteLine("QR creado para: {0}", student.Name);
                    row++;
                }

            }
        }
        public static List<Student> GetStudents()
        {
            Console.WriteLine("Obteniendo Estudiantes.....");
            List<Student> students = new();
           
            //Get data from excel and store it in a list of students.
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };
            var workBook = excelApp.Workbooks.Open(path);
            var WorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            Console.WriteLine("Lista de estudiantes: ");
            for (int i = 1; i <= WorkSheet.Rows.Count; i++)
            {
                students.Add(new Student()
                {
                    Id = (int)WorkSheet.Cells[i, 0],
                    Name = (string)WorkSheet.Cells[i, 1],
                    Age = (int)WorkSheet.Cells[i, 2],
                    Grade = (int)WorkSheet.Cells[i, 3],
                    Gender = (string)WorkSheet.Cells[i, 4],
                });
                Console.WriteLine(students[i].Name);
            }
            workBook.Close();
            excelApp.Workbooks.Close();
            excelApp.Quit();
            return students;
        }
        public static bool GenerateQR(Student item)
        {
            try
            {
                //Generate the QR code
                QRCodeGenerator generator = new QRCodeGenerator();
                QRCodeData codeData;

                codeData = generator.CreateQrCode(item.ToString(), QRCodeGenerator.ECCLevel.Q);
                PngByteQRCode qrCode = new PngByteQRCode(codeData);

                byte[] qrCodeByteArr = qrCode.GetGraphic(20);

                //Save the qr in a PNG file 
                using var ms = new MemoryStream(qrCodeByteArr);

                Image image = Image.FromStream(ms);
                string? imageName = item.Id.ToString() + item.Name + ".jpg";
                image.Save(imageName, System.Drawing.Imaging.ImageFormat.Jpeg);
                imageName = string.Empty;

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("No fue posible crear el codigo QR debido a {0}", e.Message);
                return false;

            }
        }
        public static void UpdateExcel(Student student,int row)
        {
            //update has qr field in the excel and save it
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };
            var workBook = excelApp.Workbooks.Open(path);
            var WorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            WorkSheet.Cells[row, 4] = student.HasQr?"QR creado":"Sin QR creado";
            workBook.Save();
            workBook.Close();
            excelApp.Workbooks.Close();
            excelApp.Quit();

        }
    }
}
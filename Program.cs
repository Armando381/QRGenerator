using QRGenerator.Entities;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Drawing.Imaging;

using QRCoder;
using ZXing;
using ZXing.QrCode;


namespace QRGenerator
{
    internal class Program
    {
        #region properties
        private static int NAME_MAXLENGHT = 20;
        private const string excelPath = "D:\\Users\\Descargas\\Calificaciones (2)";
        private const string imagePath = $"D:\\Users\\Projects\\QRGenerator\\StudentsQR\\";
        private const string cardPath = $"D:\\Users\\Projects\\QRGenerator\\Card\\";
        private static readonly Microsoft.Office.Interop.Excel.Application excelApp = new();
        private static Workbook? workBook;
        private static Worksheet? WorkSheet;

        #endregion

        static void Main(string[] args)
        {
            Console.WriteLine("Iniciando Programa");

            List<Student> students = GetStudents();

            int row = 1;
            foreach (Student student in students)
            {



                SaveCard(student);
                Console.WriteLine("QR creado para: {0}", student.Name);
                row++;


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
        internal static Bitmap GenerateQR(Student item)
        {
            try
            {
                //Generate the QR code
                string? imageName = item.Id.ToString() + item.Name + ".jpg";
                string data = Newtonsoft.Json.JsonConvert.SerializeObject(item);//creates JSON with student data


                QRCodeGenerator generator = new QRCodeGenerator();
                QRCodeData codeData;

                codeData = generator.CreateQrCode(data, QRCodeGenerator.ECCLevel.L);
                // PngByteQRCode qrCode = new PngByteQRCode(codeData);
                QRCoder.BitmapByteQRCode qrCode = new QRCoder.BitmapByteQRCode(codeData);

                byte[] qrCodeByteArr = qrCode.GetGraphic(4);

                //Save the qr in a PNG file 
                using var ms = new MemoryStream(qrCodeByteArr);

                Bitmap bitmap = new Bitmap(ms);


                //var qr = new QRCodeWriter();
                //var matrix = qr.encode(data, BarcodeFormat.QR_CODE, 100, 100);
                //var writer = new ZXing.Windows.Compatibility.BarcodeWriter();
                //using var bitmap = writer.Write(matrix);

                return bitmap;
            }
            catch (Exception e)
            {
                Console.WriteLine("No fue posible crear el codigo QR debido a {0}", e.Message);
                throw;

            }
        }


        internal static string GetHTML(string name, string id, string QRPath)
        {
            string html = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>    <meta name='viewport' content='width=device-width, initial-scale=1.0'>    <title>Credencial de Estudiante con Código QR</title>    <style>        body {            font-family: Arial, sans-serif;            background-color: #f0f0f0;        }        .card {            width: 300px;            background-color: #fff;            border-radius: 10px;            box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.2);            margin: 20px auto;            padding: 20px;        }        .student-photo {            width: 100px;            height: 100px;            border-radius: 50%;            margin: 0 auto 10px;            background-color: #ccc;            background-image: url('./Card/template/student.png'); /* Reemplaza 'tu_foto_de_estudiante.jpg' con la URL de tu foto */            background-size: cover;            background-position: center;        }        .student-info {            text-align: center;        }        .student-name {            font-size: 24px;            font-weight: bold;            margin: 10px 0;        }        .student-id {            font-size: 16px;            margin-bottom: 10px;        }        .validity {            font-size: 12px;            color: #888;        }        .qr-code {            width: 100px;            height: 100px;            margin: 10px auto;        }        img{            width : 80%;            height: 80%;        }    </style></head><body>    <div class='card'>        <div class='student-photo'>                    </div>        <div class='student-info'>            <div class='student-name'>{StudentName}</div>            <div class='student-id'>ID de Estudiante: {StudentId}</div>            <div class='validity'>Válido hasta: 31 de diciembre de 2023</div>        </div>        <div class='qr-code'>            <img src='{QRCode}' alt='Código QR'>        </div>    </div></body></html>\"";
            html = html.Replace("{StudentName}", name);
            html = html.Replace("{StudentId}", id);
            html = html.Replace("{QRCode}", QRPath);
            return html;

        }
        internal static void SaveCard(Student student)
        {
            string studentQRPath = Path.Combine(imagePath, string.Concat(student.Id, student.Name.Replace(" ", ""), ".png"));
            string HTMLContent = GetHTML(student.Name, student.Id, studentQRPath);
            GenerateCard(string.Concat(student.Id, student.Name.Replace(" ", "_"), "Card.JPEG"), student);

        }
        static void GenerateCard(string HTMLContent, string FileName)
        {
            try
            {
                Console.WriteLine("Generando credencial");
                // Create a PDF document from the HTML template

                var renderer = new IronPdf.HtmlToPdf();
                var pdf = renderer.RenderHtmlAsPdf(HTMLContent);
                pdf.SaveAs(cardPath + FileName);


                //Creates a MS and save the card as image

                //using var ms = new MemoryStream(pdf.Stream.ToArray());
                //System.Drawing.Image image = System.Drawing.Image.FromStream(ms);
                //image.Save(Path.Combine(imagePath, FileName));



            }
            catch (Exception)
            {

                throw;
            }


        }
        static void GenerateCard(string FileName, Student item)
        {
            try
            {
                System.Drawing.Image templateImage = System.Drawing.Image.FromFile("D:\\Users\\Projects\\QRGenerator\\Card\\template\\StudentCard2.PNG");
                System.Drawing.Point studentNamePosition;
                System.Drawing.Point studentIDPosition;
                System.Drawing.Point QrPosition;
                System.Drawing.Point studentNameHeaderPosition;
                System.Drawing.Point studentIdHeaderPosition;

                using (Graphics graphics = Graphics.FromImage(templateImage))
                {
                    // Create a font and brush for the text
                    System.Drawing.Font font = new System.Drawing.Font("Roboto", 12, FontStyle.Bold);
                    SolidBrush brush = new SolidBrush(System.Drawing.Color.Black);



                    //Sets the positions for Nombre and matricula headers.

                    studentNameHeaderPosition = new System.Drawing.Point(100, 140);
                    studentIdHeaderPosition = new System.Drawing.Point(100, 180);

                    // Define the positions to overlay student data

                    //accordong name lenght define the positions of name
                    //if it's next to nombre label or below and move it in x axis in order to center it
                    studentNamePosition = 
                        item.Name.Length <= NAME_MAXLENGHT ? 
                            new System.Drawing.Point(170, 140): 
                            item.Name.Length >= 28 ? 
                            new System.Drawing.Point(30, 160) :
                            new System.Drawing.Point(100, 160);

                    //Define student ID and QR positions in the image
                        studentIDPosition =  new System.Drawing.Point(200, 180);
                        QrPosition = new System.Drawing.Point(100, 225);
         
               


                    graphics.DrawString("Nombre: ", font, brush, studentNameHeaderPosition);
                    graphics.DrawString("Matricula: ", font, brush, studentIdHeaderPosition);
                    //overlay Student Infotmation

                    graphics.DrawString( item.Name, font, brush, studentNamePosition);
                    graphics.DrawString(item.Id, font, brush, studentIDPosition);
                    //Create QRCode and DrawIt into the image.
                    Bitmap qrCode = GenerateQR(item); // Adjust QR code size as needed
                    graphics.DrawImage(qrCode, QrPosition);


                    // Dispose of the font and brush
                    font.Dispose();
                    brush.Dispose();
                }
                templateImage.Save(cardPath + FileName);
                templateImage.Dispose();
            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}
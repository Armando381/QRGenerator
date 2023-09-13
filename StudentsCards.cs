#region Additional Usings
using QRCoder;
using QRCoder;
using ZXing;
using ZXing.QrCode;
#endregion

#region System Usings
using QRGenerator.Entities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion

namespace QRGenerator
{
    internal class StudentsCards
    {
        #region properties
        private int NAME_MAXLENGHT = 20;
        private const string imagePath = $"D:\\Users\\Projects\\QRGenerator\\StudentsQR\\";
        private const string cardPath = $"D:\\Users\\Projects\\QRGenerator\\Card\\";
        #endregion
        private Bitmap GenerateQR(Student item)
        {
            try
            {
                //Generate the QR code
                string? imageName = item.Id.ToString() + item.Name + ".jpg";
                string data = Newtonsoft.Json.JsonConvert.SerializeObject(item);//creates JSON with student data

                //Creates a QR with students information
                QRCodeGenerator generator = new QRCodeGenerator();
                QRCodeData codeData;
                codeData = generator.CreateQrCode(data, QRCodeGenerator.ECCLevel.L);
                QRCoder.BitmapByteQRCode qrCode = new QRCoder.BitmapByteQRCode(codeData);
                byte[] qrCodeByteArr = qrCode.GetGraphic(3);

                //use a memory stream to set the qr code into a bitmap image 
                using var ms = new MemoryStream(qrCodeByteArr);
                Bitmap bitmap = new Bitmap(ms);

                return bitmap;
            }
            catch (Exception e)
            {
                Console.WriteLine("No fue posible crear el codigo QR debido a {0}", e.Message);
                throw;

            }
        }

        #region ToDo
        internal string GetHTML(string name, string id, string QRPath)
        {
            //a student card as html tameplate
            string html = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>    <meta name='viewport' content='width=device-width, initial-scale=1.0'>    <title>Credencial de Estudiante con Código QR</title>    <style>        body {            font-family: Arial, sans-serif;            background-color: #f0f0f0;        }        .card {            width: 300px;            background-color: #fff;            border-radius: 10px;            box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.2);            margin: 20px auto;            padding: 20px;        }        .student-photo {            width: 100px;            height: 100px;            border-radius: 50%;            margin: 0 auto 10px;            background-color: #ccc;            background-image: url('./Card/template/student.png'); /* Reemplaza 'tu_foto_de_estudiante.jpg' con la URL de tu foto */            background-size: cover;            background-position: center;        }        .student-info {            text-align: center;        }        .student-name {            font-size: 24px;            font-weight: bold;            margin: 10px 0;        }        .student-id {            font-size: 16px;            margin-bottom: 10px;        }        .validity {            font-size: 12px;            color: #888;        }        .qr-code {            width: 100px;            height: 100px;            margin: 10px auto;        }        img{            width : 80%;            height: 80%;        }    </style></head><body>    <div class='card'>        <div class='student-photo'>                    </div>        <div class='student-info'>            <div class='student-name'>{StudentName}</div>            <div class='student-id'>ID de Estudiante: {StudentId}</div>            <div class='validity'>Válido hasta: 31 de diciembre de 2023</div>        </div>        <div class='qr-code'>            <img src='{QRCode}' alt='Código QR'>        </div>    </div></body></html>\"";

            //replace the field for student data
            html = html.Replace("{StudentName}", name);
            html = html.Replace("{StudentId}", id);
            html = html.Replace("{QRCode}", QRPath);
            return html;

        }
        void GenerateCard(string HTMLContent, string FileName)
        {
            try
            {
                Console.WriteLine("Generando credencial");
                // Create a PDF document from the HTML template and save it

                var renderer = new IronPdf.HtmlToPdf();
                var pdf = renderer.RenderHtmlAsPdf(HTMLContent);



                //Creates a MS from the pdf and save it and save the card as image

                using var ms = new MemoryStream(pdf.Stream.ToArray());
                System.Drawing.Image image = System.Drawing.Image.FromStream(ms);

                image.Save(Path.Combine(imagePath, FileName));
                pdf.SaveAs(cardPath + FileName);


            }
            catch (Exception)
            {

                throw;
            }


        }
        #endregion
        public bool CreateCard(Student student)
        {

            //In a future implement correctly save the card directly from a HTML
            //string studentQRPath = Path.Combine(imagePath, string.Concat(student.Id, student.Name.Replace(" ", ""), ".png"));
            //string HTMLContent = GetHTML(student.Name, student.Id, studentQRPath); 
            string FileName = string.Concat(student.Id, student.Name.Replace(" ", "_"), "Card.JPEG");
            GenerateCard(FileName, student);
            
            return File.Exists(cardPath+FileName);

        }

      
       private void GenerateCard(string FileName, Student item)
        {
            try
            {
                #region variables
                System.Drawing.Image templateImage = System.Drawing.Image.FromFile("D:\\Users\\Projects\\QRGenerator\\Card\\template\\StudentCardHorizontal.PNG");
                System.Drawing.Point studentNamePosition;
                System.Drawing.Point studentIDPosition;
                System.Drawing.Point QrPosition;
                System.Drawing.Point studentNameHeaderPosition;
                System.Drawing.Point studentIdHeaderPosition;
                #endregion
                using (Graphics graphics = Graphics.FromImage(templateImage))
                {
                    // Create a font and brush for the text
                    System.Drawing.Font font = new System.Drawing.Font("Roboto", 12, FontStyle.Bold);
                    SolidBrush brush = new SolidBrush(System.Drawing.Color.Black);



                    //Sets the positions for Nombre and matricula headers.

                    studentNameHeaderPosition = new System.Drawing.Point(10, 180);
                    studentIdHeaderPosition = new System.Drawing.Point(10, 220);

                    // Define the positions to overlay student data

                    //accordong name lenght define the positions of name
                    //if it's next to nombre label or below and move it in x axis in order to center it
                    studentNamePosition = new System.Drawing.Point(10, 200);
                    //    item.Name.Length <= NAME_MAXLENGHT ?
                    //        new System.Drawing.Point(170, 140) :
                    //        item.Name.Length >= 28 ?
                    //        new System.Drawing.Point(30, 160) :
                    //        new System.Drawing.Point(100, 160);


                    //Define student ID and QR positions in the image
                    studentIDPosition = new System.Drawing.Point(10, 240);
                    QrPosition = new System.Drawing.Point(300, 150);




                    graphics.DrawString("Nombre: ", font, brush, studentNameHeaderPosition);
                    graphics.DrawString("Matricula: ", font, brush, studentIdHeaderPosition);
                    //overlay Student Infotmation

                    graphics.DrawString(item.Name.TrimStart(), font, brush, studentNamePosition);
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

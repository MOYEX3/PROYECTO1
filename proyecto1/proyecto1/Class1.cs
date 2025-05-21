using System;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using System.Data.SqlClient;



//gpt-4o-mini
//-proj-iQchx9SQwA6kQuvDj0q4UKbb_hSzDBeb7Wr8ifi9TukGXiW_vfvJZumjf6ATuLfZip_nzsSpL4T3BlbkFJchKlX31HjtD6_vx8bfKpcrufgVOmVmY2_2q-J6oqOwvU9-odT2IGHV1p51g_hXNuLB1cJ7dZMA

namespace proyecto1
{
    internal class Class1

    {
        private static readonly HttpClient client = new HttpClient();
        public async Task<string>BuscarenopeniaAsync(string prompt, string apikey)
        {
            string url = "https://api.openai.com/v1/chat/completions";
            var requestBody = new
            {
                model = "gpt-4o-mini",
                messages = new[]
                { new {role = "user", content = prompt}},
                max_tokens = 100,
            };
            string json = System.Text.Json.JsonSerializer.Serialize(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apikey);
            HttpResponseMessage response = await client.PostAsync(url, content);

            if (response.IsSuccessStatusCode)
            {
                string resultado = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(resultado);
                var texto = doc.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString();
                return texto?.Trim() ?? "sin respuesta";
            }
            else
            {
                return $"error: {response.StatusCode}";
            }

        }

        public string crearcarpetain()
        {
            string proyectos = "C:\\Users\\Erick\\Desktop\\investigacion";
            if (Directory.Exists(proyectos))
            {

                MessageBox.Show("la carpeta de la investigacion ya existe");
            }
            else
            {
                MessageBox.Show("acabas de crear la carpeta de la investigacion");
                Directory.CreateDirectory(proyectos);
            }
            return proyectos;
        }


        public void documentoword(string ruta, string titulo, string cuerpo)
        {
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }
            // El error se debe a que hay varias clases llamadas 'Path' en los espacios de nombres importados.
            // Para solucionarlo, usa la referencia completa del espacio de nombres para 'System.IO.Path':

            string rutaarchivo = System.IO.Path.Combine(ruta, "documento.docx");
            using (WordprocessingDocument doc = WordprocessingDocument.Create(rutaarchivo, WordprocessingDocumentType.Document))
            {
                var mainpart = doc.AddMainDocumentPart();
                mainpart.Document = new W.Document();
                var body = mainpart.Document.AppendChild(new W.Body());

                // Título con estilo Heading1
                var parrafotitulo = new W.Paragraph();
                parrafotitulo.AppendChild(new W.ParagraphProperties(new W.ParagraphStyleId() { Val = "Heading1" }));
                parrafotitulo.AppendChild(new W.Run(new W.Text(titulo)));
                body.Append(parrafotitulo);

                // Cuerpo
                var parrafoCuerpo = new W.Paragraph();
                parrafoCuerpo.AppendChild(new W.Run(new W.Text(cuerpo)));
                body.Append(parrafoCuerpo);

                mainpart.Document.Save();
            }
        }

        public void documentopptx(string ruta, string titulo, string cuerpo)
        {
            if (string.IsNullOrWhiteSpace(ruta))
            {
                MessageBox.Show("La ruta está vacía o es nula.");
                return;
            }

            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }

            string rutaarchivo = System.IO.Path.Combine(ruta, "presentacion.pptx");

            using (PresentationDocument presentationDoc = PresentationDocument.Create(rutaarchivo, PresentationDocumentType.Presentation))
            {
                PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                slideMasterPart.SlideMaster = new SlideMaster(new CommonSlideData(new ShapeTree()));

                SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                slideLayoutPart.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));

                slideMasterPart.SlideMaster.AppendChild(new SlideLayoutIdList(
                    new SlideLayoutId() { Id = 1U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }));

                presentationPart.Presentation.AppendChild(new SlideMasterIdList(
                    new SlideMasterId() { Id = 1U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }));

                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));
                slidePart.AddPart(slideLayoutPart);

                SlideIdList slideIdList = new SlideIdList();
                uint slideId = 256U;
                SlideId slideIdElement = new SlideId() { Id = slideId, RelationshipId = presentationPart.GetIdOfPart(slidePart) };
                slideIdList.Append(slideIdElement);
                presentationPart.Presentation.AppendChild(slideIdList);

                var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

                shapeTree.AppendChild(new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = 1U, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()));
                shapeTree.AppendChild(new GroupShapeProperties());

                // Título (arriba)
                shapeTree.Append(CreateTextShape(2U, "Título", titulo, 1000000L, 1000000L, 8000000L, 1000000L));

                // Cuerpo (abajo)
                shapeTree.Append(CreateTextShape(3U, "Cuerpo", cuerpo, 1000000L, 2200000L, 8000000L, 4000000L));

                presentationPart.Presentation.Save();
            }
        }
        private Shape CreateTextShape(uint id, string name, string text, long x, long y, long cx, long cy)
        {
            return new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties() { Id = id, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = x, Y = y },
                        new A.Extents() { Cx = cx, Cy = cy }
                    )
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.Text(text))
                    )
                )
            );
        }
        private string conexionSql = "Server=DESKTOP-21B0LCS\\SQLEXPRESS;Database=investigacionesDB;Integrated Security=True;";


        public void GuardarEnBaseDeDatos(string prompt, string respuesta)
        {
            using (SqlConnection conexion = new SqlConnection(conexionSql))
            {
                conexion.Open();
                string query = "INSERT INTO Investigaciones (Prompt, Respuesta) VALUES (@prompt, @respuesta)";
                using (SqlCommand comando = new SqlCommand(query, conexion))
                {
                    comando.Parameters.AddWithValue("@prompt", prompt);
                    comando.Parameters.AddWithValue("@respuesta", respuesta);
                    comando.ExecuteNonQuery();
                }
            }
        }





    }


}

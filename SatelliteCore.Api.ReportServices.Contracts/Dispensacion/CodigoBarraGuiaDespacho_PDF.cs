using iText.Barcodes;
using iText.Commons.Actions;
using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SatelliteCore.Api.ReportServices.Contracts.Dispensacion
{
    public class CodigoBarraGuiaDespacho_PDF
    {
        public string Exportar(string id)
        {
            string reporte = null;

            MemoryStream ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Autorización de sobretiempo");
            docInfo.SetAuthor("Sistema Satelite");
           

            Document document = new Document(pdf, PageSize.A8.Rotate());

            Barcode128 code128 = new Barcode128(pdf);
            code128.SetFont(null);
            code128.SetCode(id.PadLeft(10,'0'));
            code128.SetCodeType(Barcode128.CODE128);

            Image code128Image = new Image(code128.CreateFormXObject(pdf)).SetWidth(130).SetBorder(Border.NO_BORDER).SetAutoScale(false);

            Table cuerpoTabla = new Table(1).UseAllAvailableWidth().SetBorder(Border.NO_BORDER);
            cuerpoTabla.SetFixedLayout();

            Cell cellCodigoBarraResponsables = new Cell(1, 1).Add(code128Image).SetBorder(Border.NO_BORDER);
            cuerpoTabla.AddCell(cellCodigoBarraResponsables);

            cellCodigoBarraResponsables = new Cell(1, 1).Add(new Paragraph(id.PadLeft(10, '0')))
                .SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            cuerpoTabla.AddCell(cellCodigoBarraResponsables);


            document.Add(cuerpoTabla);

            document.Close();

            byte[] file = ms.ToArray();

            if (file == null || file.Length == 0)
                return reporte;

            reporte = Convert.ToBase64String(file, 0, file.Length);

            pdf.Close();
            writer.Close();
            ms.Close();

            return reporte;
        }
    }
}

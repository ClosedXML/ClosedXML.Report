using System.Drawing;
using System.IO;
using ClosedXML.Excel.Drawings;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    public class ImageTag: OptionTag
    {
        public string Value => Parameters.ContainsKey("value") ? Parameters["value"] : null;
        public string ImageName => Parameters.ContainsKey("imagename") ? Parameters["imagename"] : null;
        public double Scale => Parameters.ContainsKey("scale") ? Parameters["scale"].AsDouble() : 0;
        public int Width => Parameters.ContainsKey("width") ? Parameters["width"].AsInt() : 0;
        public int Height => Parameters.ContainsKey("height") ? Parameters["height"].AsInt() : 0;

        /*
                IXLPicture AddPicture(Stream stream, XLPictureFormat format);
                IXLPicture AddPicture(Stream stream, XLPictureFormat format, string name);
         */
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            if (!string.IsNullOrEmpty(Value))
            {
                IXLPicture picture;
                var imgValue = context.Evaluator.Evaluate(Value, new Parameter("item", context.Value));

                switch (imgValue)
                {
                    case Stream stream: picture = xlCell.Worksheet.AddPicture(stream); break;
                    case string path: picture = xlCell.Worksheet.AddPicture(path); break;
                    case Bitmap image: picture = xlCell.Worksheet.AddPicture(image); break;
                    default: throw new TemplateParseException("Unsupported image type.", xlCell.AsRange());
                };
                picture.MoveTo(xlCell);
                if (!string.IsNullOrEmpty(ImageName)) picture.Name = ImageName;
                if (Scale > 0) picture.Scale(Scale);
                if (Width > 0) picture.Width = Width;
                if (Height > 0) picture.Height = Height;
            }
        }
    }
}

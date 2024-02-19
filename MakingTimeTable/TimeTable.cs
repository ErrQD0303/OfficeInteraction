using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace MakingTimeTable
{
    public class TimeTable : IOfficeWorker
    {
        void IOfficeWorker.Make(params string[] urls)
        {
            var today = DateOnly.FromDateTime(DateTime.Now);
            if (today.DayOfWeek != DayOfWeek.Monday)
                today = today.AddDays(-(int)today.DayOfWeek);

            foreach (var url in urls)
            {
                if (string.IsNullOrWhiteSpace(url))
                    return;

                using (var templateDocument = WordprocessingDocument.Open(url, false))
                {
                    var outputFileName = url.Replace(".docx", "_output.docx");
                    using (var outputDocument = WordprocessingDocument.Create(outputFileName, WordprocessingDocumentType.Document))
                    {
                        var mainPart = outputDocument.AddMainDocumentPart();

                        mainPart.Document = new Document() { Body = new() };

                        for (var i = 0; i < 53; ++i)
                        {
                            string timeInterval = GetWeekIntervalString(today, i);

                            var clonedElements = CloneElements(templateDocument.MainDocumentPart!.Document.Body!, typeof(Paragraph), typeof(Table));

                            var ownerNameSection = clonedElements
                                .FirstOrDefault(x => x.InnerText.Contains("[OwnerName]"))
                                ?.Descendants<Text>()
                                ?.FirstOrDefault(y => y.InnerText == "[OwnerName]");

                            ownerNameSection!.Text = ownerNameSection.Text
                                .Replace("[OwnerName]", @"Nguyễn Quốc Đạt");

                            var timeIntervalSection = clonedElements
                                .FirstOrDefault(x => x.InnerText.Contains("[TimeInterval]"))
                                ?.Descendants<Text>()
                                ?.FirstOrDefault(y => y.InnerText == "[TimeInterval]");
                            timeIntervalSection!.Text = timeIntervalSection.Text
                                .Replace("[TimeInterval]", timeInterval);

                            foreach (var element in clonedElements)
                                mainPart.Document.Body!.AppendChild(element);
                        }

                        var otherPart = CloneElements(templateDocument.MainDocumentPart!.Document.Body!, typeof(SectionProperties));
                        foreach (var element in otherPart)
                            mainPart.Document.Body!.AppendChild(element);

                        outputDocument.Save();
                    }
                    /*var runs = GetRuns(templateDocument)
                        .Where(x => Regex.IsMatch(x.InnerText, @"^\[[\w\d]*\]$"));*/

                }
            }
        }

        private static string GetWeekIntervalString(DateOnly today, int i)
        {
            var sb = new StringBuilder();

            sb.Append(today.AddDays(i * 7).ToString("dd/MM/yyyy"));
            sb.Append(" - ");
            sb.Append(today.AddDays(i * 7 + 6).ToString("dd/MM/yyyy"));

            return sb.ToString();
        }

        private IEnumerable<Run> GetRuns(WordprocessingDocument myDocument)
        {
            var body = myDocument?.MainDocumentPart?.Document.Body;
            return body?.Descendants<Run>()!;
        }

        //Another Comment
        private OpenXmlElement[] CloneElements(Body body, params Type[] types)
        {
            IEnumerable<OpenXmlElement> parentElements = body.ChildElements;

            if (types.Length != 0)
                parentElements = parentElements.Where(e => types.Any(x => e.GetType() == x));

            var returnElements = parentElements
                .Select(e => e.CloneNode(true));

            foreach (var element in returnElements)
            {
                if (element is Paragraph paragraph)
                {
                    foreach (OpenXmlElement run in paragraph.Descendants<Run>())
                    {
                        var runProperties = new RunProperties(new RunFonts() { Ascii = "Arial" });
                        run.RemoveAllChildren<RunProperties>();
                        run.PrependChild<RunProperties>(runProperties);
                    }
                }
            }

            return returnElements.ToArray();
        }
    }
}

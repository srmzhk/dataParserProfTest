using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using DataParserProfTest.Model;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.OfficeChart;

namespace DataParserProfTest.Core
{
    internal class ChartsCreator
    {
        public void CreateChartInterests(Test test, Dictionary<int, int> scores, string email, string date, string name)
        {
            //Create a word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();
            section.PageSetup.Margins.All = 70f;

            //Create a new paragraph and append text
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText(name.Replace("-", " "));
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Heading3);

            paragraph = section.AddParagraph();
            paragraph.AppendText("\n" + email + "\n");
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Normal);

            //Create a new paragraph to append a bar chart shape
            paragraph = section.AddParagraph();
            Spire.Doc.Fields.ShapeObject shape = paragraph.AppendChart(ChartType.Bar, 500, 300);

            //Clear the default series of the chart
            Chart chart = shape.Chart;
            chart.Series.Clear();

            var sortedScores = scores.OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
            string[] categories = new string[15];
            double[] values = new double[15];
            //Specify chart data
            int counter = 0;
            foreach (var score in sortedScores)
            {
                switch (score.Key)
                {
                    case 1:
                        categories[counter] = "физика";
                        break;
                    case 2:
                        categories[counter] = "математика";
                        break;
                    case 3:
                        categories[counter] = "экономика и бизнес";
                        break;
                    case 4:
                        categories[counter] = "техника и электротехника";
                        break;
                    case 5:
                        categories[counter] = "химия";
                        break;
                    case 6:
                        categories[counter] = "биология и сельское хозяйство";
                        break;
                    case 7:
                        categories[counter] = "медицина";
                        break;
                    case 8:
                        categories[counter] = "география и геология";
                        break;
                    case 9:
                        categories[counter] = "история";
                        break;
                    case 10:
                        categories[counter] = "филология, журналистика";
                        break;
                    case 11:
                        categories[counter] = "искусство";
                        break;
                    case 12:
                        categories[counter] = "педагогика";
                        break;
                    case 13:
                        categories[counter] = "труд в сфере обслуживания";
                        break;
                    case 14:
                        categories[counter] = "военное дело";
                        break;
                    case 15:
                        categories[counter] = "спорт";
                        break;
                    default:
                        break;
                }
                values[counter++] = score.Value;
            }

            //Add data series to the chart
            chart.Series.Add("", categories, values);

            //Set chart title
            chart.Title.Text = test.Title;
            //Set the number format of the Y-axis tick labels to group digits with commas
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            //Save the result document
            string path = Path.GetFullPath("..\\..\\..\\results\\" + name + "-" + test.Title + "-" + date + ".docx");
            document.SaveToFile(path, FileFormat.Docx2016);
            document.Dispose();
        }

        public string CreateChartOrientation(Test test, Dictionary<int, int> scoresWant, Dictionary<int, int> scoresCan, string email, string date, string name)
        {
            string path;
            using (WordDocument document = new WordDocument())
            {
                // Add a section to the document.
                IWSection section = document.AddSection();

                //user data
                IWParagraph paragraphName = section.AddParagraph();
                paragraphName.AppendText(name.Replace("-", " "));
                IWParagraph paragraphEmail = section.AddParagraph();
                paragraphEmail.AppendText("\n" + email + "\n");

                //custom style for user data paragraph
                var myStyle = document.AddParagraphStyle("MyCustomStyle");
                myStyle.CharacterFormat.FontName = "Times New Roman";
                myStyle.CharacterFormat.FontSize = 14;
                myStyle.CharacterFormat.Bold = true;
                paragraphName.ApplyStyle("MyCustomStyle");

                //Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();

                //Create and append the chart to the paragraph.
                WChart chart = paragraph.AppendChart(520, 216);

                paragraph.AppendText("\n\n\n");

                //убираем линии на заднем фоне
                chart.PrimaryValueAxis.HasMajorGridLines = false;
                chart.PrimaryValueAxis.HasMinorGridLines = false;

                //размер подписи оси Ox
                var categoryAxis = chart.PrimaryCategoryAxis;
                categoryAxis.Font.Size = 8;

                //Set chart type.
                chart.ChartType = OfficeChartType.Column_Clustered;
                chart.ChartArea.Fill.FillType = OfficeFillType.SolidColor;
                //Assign data.
                chart.ChartData.SetValue(1, 2, "Хочу");
                chart.ChartData.SetValue(1, 3, "Могу");

                string[] categories = new string[7];

                int k = 0;
                foreach (var score in scoresWant)
                {
                    switch (score.Key)
                    {
                        case 1:
                            categories[k] = "Человек-\nчеловек";
                            break;
                        case 2:
                            categories[k] = "Человек-\nтехника";
                            break;
                        case 3:
                            categories[k] = "Человек-\nинформация ";
                            break;
                        case 4:
                            categories[k] = "Человек-\nискусство";
                            break;
                        case 5:
                            categories[k] = "Человек-\nприрода";
                            break;
                        case 6:
                            categories[k] = "Исполнительский\nхарактер труда";
                            break;
                        case 7:
                            categories[k] = "Творческий\nхарактер\nтруда";
                            break;
                        default:
                            categories[k] = "пу пу пу";
                            break;
                    }
                    k++;
                }

                int counter = 0;
                for (int i = 2; i < 9; i++)
                {
                    chart.ChartData.SetValue(i, 1, categories[counter++]);
                    chart.ChartData.SetValue(i, 2, scoresWant[counter]);
                    chart.ChartData.SetValue(i, 3, scoresCan[counter]);
                }

                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = test.Title;
                chart.ChartTitleArea.Size = 14;
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series.Add("Хочу");
                //Set the data range of chart series – start row, start column, end row, and end column.
                series1.Values = chart.ChartData[2, 2, 8, 2];
                IOfficeChartSerie series2 = chart.Series.Add("Могу");
                //Set the data range of chart series start row, start column, end row, and end column.
                series2.Values = chart.ChartData[2, 3, 8, 3];
                //Set the data range of the category axis.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 8, 1];

                //подписи сверху колонок
                series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series1.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;
                series2.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;

                //размер подписей
                series1.DataPoints.DefaultDataPoint.DataLabels.Size = 9;
                series2.DataPoints.DefaultDataPoint.DataLabels.Size = 9;

                //границы
                series1.SerieFormat.LineProperties.LineColor = Syncfusion.Drawing.Color.Black;
                series1.SerieFormat.LineProperties.LinePattern = OfficeChartLinePattern.Solid;
                series1.SerieFormat.LineProperties.LineWeight = OfficeChartLineWeight.Medium;

                series2.SerieFormat.LineProperties.LineColor = Syncfusion.Drawing.Color.Black;
                series2.SerieFormat.LineProperties.LinePattern = OfficeChartLinePattern.Solid;
                series2.SerieFormat.LineProperties.LineWeight = OfficeChartLineWeight.Medium;

                // Устанавливаем тип штриховки
                series1.SerieFormat.Fill.Pattern = OfficeGradientPattern.Pat_5_Percent;
                series1.SerieFormat.Fill.BackColor = Syncfusion.Drawing.Color.White;
                series1.SerieFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Black;

                series2.SerieFormat.Fill.Pattern = OfficeGradientPattern.Pat_Wide_Upward_Diagonal;
                series2.SerieFormat.Fill.BackColor = Syncfusion.Drawing.Color.White;
                series2.SerieFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Black;

                //Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;

                //Create a file stream.
                path = Path.GetFullPath(@"../../../results/" + name + "-" + test.Title + "-" + date + ".docx");
                using (FileStream outputFileStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
            return path;
        }

        public string CreateChartInlination(Test test, Dictionary<int, int> scores, string email, string date, string name)
        {
            //Create a word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();
            section.PageSetup.Margins.All = 70f;

            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText(name.Replace("-", " "));
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Heading3);

            paragraph = section.AddParagraph();
            paragraph.AppendText("\n" + email + "\n");
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Normal);

            //Create a new paragraph to append a bar chart shape
            paragraph = section.AddParagraph();
            Spire.Doc.Fields.ShapeObject shape = paragraph.AppendChart(ChartType.Bar, 500, 255);
            paragraph.AppendText("\n\n");

            //Clear the default series of the chart
            Chart chart = shape.Chart;
            chart.Series.Clear();

            string[] categories = new string[12];
            double[] values = new double[12];
            int counter = 0;
            foreach (var score in scores)
            {
                switch (score.Key)
                {
                    case 1:
                        categories[counter] = "спортивно–физическая";
                        break;
                    case 2:
                        categories[counter] = "аналитико-математическая";
                        break;
                    case 3:
                        categories[counter] = "конструкторско–техническая";
                        break;
                    case 4:
                        categories[counter] = "обращение со знаковыми системами";
                        break;
                    case 5:
                        categories[counter] = "филологическая";
                        break;
                    case 6:
                        categories[counter] = "художественная";
                        break;
                    case 7:
                        categories[counter] = "изобразительная";
                        break;
                    case 8:
                        categories[counter] = "музыкальная";
                        break;
                    case 9:
                        categories[counter] = "природоохранная и сельскохозяйственная";
                        break;
                    case 10:
                        categories[counter] = "коммуникативная";
                        break;
                    case 11:
                        categories[counter] = "организаторская";
                        break;
                    case 12:
                        categories[counter] = "социально–педагогическая";
                        break;
                    default:
                        categories[counter] = "пу пу пу";
                        break;
                }
                if (score.Key == 12)
                {
                    values[counter++] = score.Value + scores.ElementAt(12).Value + scores.ElementAt(13).Value;
                    break;
                }
                else
                    values[counter++] = score.Value;
            }

            //Add data series to the chart
            chart.Series.Add("", categories, values);

            //Set chart title
            chart.Title.Text = test.Title;
            //Set the number format of the Y-axis tick labels to group digits with commas
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            //Save the result document
            string path = Path.GetFullPath(@"../../../results/" + name + "-" + test.Title + "-" + date + ".docx");
            document.SaveToFile(path, FileFormat.Docx2016);
            document.Dispose();
            return path;
        }

        public string CreateChartThinkingType(Test test, Dictionary<int, int> scores, string email, string date, string name)
        {
            //Create a word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();
            section.PageSetup.Margins.All = 70f;

            //Create a new paragraph and append text
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText(name.Replace("-", " "));
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Heading3);

            paragraph = section.AddParagraph();
            paragraph.AppendText("\n" + email + "\n");
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Normal);

            //Create a new paragraph to append a bar chart shape
            paragraph = section.AddParagraph();
            Spire.Doc.Fields.ShapeObject shape = paragraph.AppendChart(ChartType.Bar, 350, 255);
            paragraph.AppendText("\n\n");

            //Clear the default series of the chart
            Chart chart = shape.Chart;
            chart.Series.Clear();

            string[] categories = new string[5];
            double[] values = new double[5];

            var sortedScores = scores.OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);

            int counter = 0;
            foreach (var score in sortedScores)
            {
                switch (score.Key)
                {
                    case 1:
                        categories[counter] = "Предметно-действенное";
                        break;
                    case 2:
                        categories[counter] = "Абстрактно-символическое";
                        break;
                    case 3:
                        categories[counter] = "Словесно-логическое";
                        break;
                    case 4:
                        categories[counter] = "Наглядно-образное";
                        break;
                    case 5:
                        categories[counter] = "Креативность";
                        break;
                    default:
                        break;
                }
                values[counter++] = score.Value;
            }

            //Add data series to the chart
            chart.Series.Add("", categories, values);

            //Set chart title
            chart.Title.Text = test.Title;
            //Set the number format of the Y-axis tick labels to group digits with commas
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            //Save the result document
            string path = Path.GetFullPath("..\\..\\..\\results\\" + name + "-" + test.Title + "-" + date + ".docx");
            document.SaveToFile(path, FileFormat.Docx2016);
            document.Dispose();
            return path;
        }

        public string CreateChartProfType(Test test, Dictionary<string, int> scores, string email, string date, string name)
        {
            //Create a word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();
            section.PageSetup.Margins.All = 70f;

            //Create a new paragraph and append text
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText(name.Replace("-", " "));
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Heading3);

            paragraph = section.AddParagraph();
            paragraph.AppendText("\n" + email + "\n");
            paragraph.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Normal);

            //Create a new paragraph to append a bar chart shape
            paragraph = section.AddParagraph();
            Spire.Doc.Fields.ShapeObject shape = paragraph.AppendChart(ChartType.Bar, 350, 200);
            paragraph.AppendText("\n\n");

            //Clear the default series of the chart
            Chart chart = shape.Chart;
            chart.Series.Clear();

            var sortedScores = scores.OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
            string[] categories = new string[6];
            double[] values = new double[6];

            int counter = 0;
            foreach (var item in sortedScores)
            {
                categories[counter] = item.Key;
                values[counter++] = item.Value;
            }

            //Add data series to the chart
            chart.Series.Add("", categories, values);

            //Set chart title
            chart.Title.Text = test.Title;
            //Set the number format of the Y-axis tick labels to group digits with commas
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            //Save the result document
            string path = Path.GetFullPath(@"../../../results/" + name + "-" + test.Title + "-" + date + ".docx");
            document.SaveToFile(path, FileFormat.Docx2016);
            document.Dispose();
            return path;
        }
    }
}

using Spire.Doc;
using DataParserProfTest.Model;

namespace DataParserProfTest.Core
{
    internal class DocsMerger
    {
        public static void ClearDoc(string path)
        {
            Document document = new Document();
            document.LoadFromFile(path);

            // Удаляем колонтитулы из всех секций документа
            foreach (Section section in document.Sections)
            {
                // Очищаем содержимое верхнего колонтитула
                if (section.HeadersFooters.Header != null)
                    section.HeadersFooters.Header.ChildObjects.Clear();

                // Очищаем содержимое нижнего колонтитула
                if (section.HeadersFooters.Footer != null)
                    section.HeadersFooters.Footer.ChildObjects.Clear();
            }

            if (document.Sections.Count > 0 && document.Sections[0].Paragraphs.Count > 0)
            {
                document.Sections[0].Paragraphs.RemoveAt(0);
            }

            // Удаляем последний параграф
            if (document.Sections.Count > 0 && document.Sections[document.Sections.Count - 1].Paragraphs.Count > 0)
            {
                document.Sections[document.Sections.Count - 1].Paragraphs.RemoveAt(document.Sections[document.Sections.Count - 1].Paragraphs.Count - 1);
            }

            // Сохраняем изменения
            document.SaveToFile(path, FileFormat.Docx);
        }

        public static void MergeDocsOrientation(string path, Dictionary<int, int> scoresWant, Dictionary<int, int> scoresCan)
        {
            string testPath = Path.GetFullPath(@"../../../tests/orientation/");

            int counter = 1;
            Dictionary<int, int> mergeDict = new Dictionary<int, int>();

            foreach (var item in scoresWant)
            {
                mergeDict.Add(item.Key, item.Value + scoresCan[counter]);
                counter++;
            }

            var sortedMergeDict = mergeDict.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);

            //two main skills

            string filePath1 = Path.GetFullPath(testPath + mergeDict.ElementAt(0).Key + ".docx");
            string filePath2 = Path.GetFullPath(testPath + mergeDict.ElementAt(1).Key + ".docx");
            //A or B
            string filePath3 = "";
            if (mergeDict.ElementAt(5).Value > mergeDict.ElementAt(6).Value)
                filePath3 = Path.GetFullPath(testPath + mergeDict.ElementAt(5).Key + ".docx");
            else
                filePath3 = Path.GetFullPath(testPath + mergeDict.ElementAt(6).Key + ".docx");

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Microsoft.Office.Interop.Word.Document targetDoc = wordApp.Documents.Open(path);
            Microsoft.Office.Interop.Word.Document sourceDoc1 = wordApp.Documents.Open(filePath1);
            Microsoft.Office.Interop.Word.Document sourceDoc2 = wordApp.Documents.Open(filePath2);
            Microsoft.Office.Interop.Word.Document sourceDoc3 = wordApp.Documents.Open(filePath3);
            Microsoft.Office.Interop.Word.Range endRange = targetDoc.Content;
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            //изменение полей страницы
            var pageSetup = targetDoc.PageSetup;
            float cmToPoints = 28.35f; // Коэффициент для преобразования сантиметров в пункты (точки)
            pageSetup.LeftMargin = 1.5f * cmToPoints; // Левое поле (2.5 см)
            pageSetup.RightMargin = 1f * cmToPoints; // Правое поле (2.5 см)
            pageSetup.TopMargin = 1f * cmToPoints; // Верхнее поле (2.0 см)
            pageSetup.BottomMargin = 1f * cmToPoints; // Нижнее поле (2.0 см)

            // Copy the content from the source document to the target document
            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc1.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
            endRange.InsertAfter("\n\n");
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc2.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
            endRange.InsertAfter("\n\n");
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc3.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }

            // Закрыть и сохранить исходный документ
            targetDoc.Save();
            targetDoc.Close();
            // Закрыть источник документа
            sourceDoc1.Close();
            sourceDoc2.Close();
            sourceDoc3.Close();

            // Закрыть приложение Word
            wordApp.Quit();
        }

        public static void MergeDocsInlination(string path, Dictionary<int, int> scores)
        {
            string filepath = Path.GetFullPath(@"../../../tests/inelination/");

            //слияние трёх соц-подагогоических значений
            Dictionary<int, int> cuttingScores = new Dictionary<int, int>();
            for (int i = 1; i < 12; i++)
            {
                cuttingScores.Add(i, scores[i]);
            }
            cuttingScores.Add(12, scores[12] + scores[13] + scores[14]);

            var sortedScores = cuttingScores.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);

            string filePath1 = Path.GetFullPath(filepath + sortedScores.ElementAt(0).Key + ".docx");
            string filePath2 = Path.GetFullPath(filepath + sortedScores.ElementAt(1).Key + ".docx");
            string filePath3 = Path.GetFullPath(filepath + sortedScores.ElementAt(2).Key + ".docx");

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Microsoft.Office.Interop.Word.Document sourceDoc1 = wordApp.Documents.Open(filePath1);
            Microsoft.Office.Interop.Word.Document sourceDoc2 = wordApp.Documents.Open(filePath2);
            Microsoft.Office.Interop.Word.Document sourceDoc3 = wordApp.Documents.Open(filePath3);
            Microsoft.Office.Interop.Word.Document targetDoc = wordApp.Documents.Open(path);
            Microsoft.Office.Interop.Word.Range endRange = targetDoc.Content;

            //изменение полей страницы
            var pageSetup = targetDoc.PageSetup;
            float cmToPoints = 28.35f; // Коэффициент для преобразования сантиметров в пункты (точки)
            pageSetup.LeftMargin = 1.5f * cmToPoints; // Левое поле (2.5 см)
            pageSetup.RightMargin = 1f * cmToPoints; // Правое поле (2.5 см)
            pageSetup.TopMargin = 1f * cmToPoints; // Верхнее поле (2.0 см)
            pageSetup.BottomMargin = 1f * cmToPoints; // Нижнее поле (2.0 см)

            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            // Copy the content from the source document to the target document
            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc1.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            // Copy the content from the source document to the target document
            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc2.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc3.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            if (sortedScores.ElementAt(2).Value == sortedScores.ElementAt(3).Value)
            {
                string filePath4 = Path.GetFullPath(filepath + sortedScores.ElementAt(3).Key + ".docx");
                Microsoft.Office.Interop.Word.Document sourceDoc4 = wordApp.Documents.Open(filePath4);
                foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc4.StoryRanges)
                {
                    range.Copy();
                    endRange.Paste();
                }
                endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                sourceDoc4.Close();
                if (sortedScores.ElementAt(3).Value == sortedScores.ElementAt(4).Value)
                {
                    string filePath5 = Path.GetFullPath(filepath + sortedScores.ElementAt(4).Key + ".docx");
                    Microsoft.Office.Interop.Word.Document sourceDoc5 = wordApp.Documents.Open(filePath5);
                    foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc5.StoryRanges)
                    {
                        range.Copy();
                        endRange.Paste();
                    }
                    endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                    sourceDoc5.Close();
                }
            }

            targetDoc.Save();
            targetDoc.Close();
            // Закрыть источник документа
            sourceDoc1.Close();
            sourceDoc2.Close();
            sourceDoc3.Close();
            // Закрыть приложение Word
            wordApp.Quit();
        }

        public static void MergeDocsSocialBrigs(Test test, Dictionary<string, int> scores, string email, string date, string name)
        {
            string path;
            if (test.ID == 4)
                path = Path.GetFullPath(@"../../../tests/brigs/");
            else
                path = Path.GetFullPath(@"../../../tests/socionalType/");
            string letters = "";

            if (scores.ElementAt(0).Value > scores.ElementAt(1).Value)
                letters += scores.ElementAt(0).Key;
            else
                letters += scores.ElementAt(1).Key;

            if (scores.ElementAt(2).Value > scores.ElementAt(3).Value)
                letters += scores.ElementAt(2).Key;
            else
                letters += scores.ElementAt(3).Key;

            if (scores.ElementAt(4).Value > scores.ElementAt(5).Value)
                letters += scores.ElementAt(4).Key;
            else
                letters += scores.ElementAt(5).Key;

            if (scores.ElementAt(6).Value > scores.ElementAt(7).Value)
                letters += scores.ElementAt(6).Key;
            else
                letters += scores.ElementAt(7).Key;

            string filePath1 = Path.GetFullPath(path + letters + ".docx");

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Microsoft.Office.Interop.Word.Document sourceDoc1 = wordApp.Documents.Open(filePath1);
            Microsoft.Office.Interop.Word.Document targetDoc = wordApp.Documents.Add();
            Microsoft.Office.Interop.Word.Range endRange = targetDoc.Content;
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            //изменение полей страницы
            var pageSetup = targetDoc.PageSetup;
            float cmToPoints = 28.35f; // Коэффициент для преобразования сантиметров в пункты (точки)
            pageSetup.LeftMargin = 1.5f * cmToPoints; // Левое поле (2.5 см)
            pageSetup.RightMargin = 1f * cmToPoints; // Правое поле (2.5 см)
            pageSetup.TopMargin = 1f * cmToPoints; // Верхнее поле (2.0 см)
            pageSetup.BottomMargin = 1f * cmToPoints; // Нижнее поле (2.0 см)

            //add new paragraph for user data
            endRange.InsertAfter(name.Replace("-", " "));
            Microsoft.Office.Interop.Word.Font font1 = endRange.Font;
            font1.Name = "Times New Roman";
            font1.Size = 14;      // Новый размер шрифта
            font1.Bold = 1;       // Жирный шрифт (1 - true, 0 - false)
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            endRange.InsertAfter("\n" + email + "\n\n");
            Microsoft.Office.Interop.Word.Font font2 = endRange.Font;
            font2.Name = "Times New Roman";
            font2.Size = 12;      // Новый размер шрифта
            font2.Bold = 0;       // Жирный шрифт (1 - true, 0 - false)
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            // Copy the content from the source document to the target document
            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc1.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }

            // Закрыть и сохранить исходный документ
            string targetPath = Path.GetFullPath(@"../../../results/" + name + "-" + test.Title + "-" + date + ".docx");
            targetDoc.SaveAs2(targetPath);
            targetDoc.Close();
            // Закрыть источник документа
            sourceDoc1.Close();

            // Закрыть приложение Word
            wordApp.Quit();
        }        

        public static void MergeDocsThinkingType(string path, Dictionary<int, int> scores)
        {
            string filePath = Path.GetFullPath(@"../../../tests/thinkingType/");

            var sortedScores = scores.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);

            string filePath1 = Path.GetFullPath(filePath + sortedScores.ElementAt(0).Key.ToString() + ".docx");
            string filePath2 = Path.GetFullPath(filePath + sortedScores.ElementAt(1).Key.ToString() + ".docx");

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Microsoft.Office.Interop.Word.Document sourceDoc1 = wordApp.Documents.Open(filePath1);
            Microsoft.Office.Interop.Word.Document sourceDoc2 = wordApp.Documents.Open(filePath2);
            Microsoft.Office.Interop.Word.Document targetDoc = wordApp.Documents.Open(path);
            Microsoft.Office.Interop.Word.Range endRange = targetDoc.Content;

            //изменение полей страницы
            var pageSetup = targetDoc.PageSetup;
            float cmToPoints = 28.35f; // Коэффициент для преобразования сантиметров в пункты (точки)
            pageSetup.LeftMargin = 1.5f * cmToPoints; // Левое поле (2.5 см)
            pageSetup.RightMargin = 1f * cmToPoints; // Правое поле (2.5 см)
            pageSetup.TopMargin = 1f * cmToPoints; // Верхнее поле (2.0 см)
            pageSetup.BottomMargin = 1f * cmToPoints; // Нижнее поле (2.0 см)

            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            // Copy the content from the source document to the target document
            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc1.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc2.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            if (sortedScores.ElementAt(1).Value == sortedScores.ElementAt(2).Value)
            {
                string filePath3 = Path.GetFullPath(filePath + sortedScores.ElementAt(2).Key.ToString() + ".docx");
                Microsoft.Office.Interop.Word.Document sourceDoc3 = wordApp.Documents.Open(filePath3);
                foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc3.StoryRanges)
                {
                    range.Copy();
                    endRange.Paste();
                }
                sourceDoc3.Close();
            }

            // Закрыть и сохранить исходный документ
            targetDoc.Save();
            targetDoc.Close();
            // Закрыть источник документа
            sourceDoc1.Close();
            sourceDoc2.Close();
            // Закрыть приложение Word
            wordApp.Quit();
        }

        public static void MergeDocsProfType(string path, Dictionary<string, int> scores)
        {
            string filePath = Path.GetFullPath(@"../../../tests/profType/");

            var sortedScores = scores.OrderByDescending(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);

            string filePath1 = Path.GetFullPath(filePath + sortedScores.ElementAt(0).Key + ".docx");
            string filePath2 = Path.GetFullPath(filePath + sortedScores.ElementAt(1).Key + ".docx");

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Microsoft.Office.Interop.Word.Document sourceDoc1 = wordApp.Documents.Open(filePath1);
            Microsoft.Office.Interop.Word.Document sourceDoc2 = wordApp.Documents.Open(filePath2);
            Microsoft.Office.Interop.Word.Document targetDoc = wordApp.Documents.Open(path);
            Microsoft.Office.Interop.Word.Range endRange = targetDoc.Content;

            //изменение полей страницы
            var pageSetup = targetDoc.PageSetup;
            float cmToPoints = 28.35f; // Коэффициент для преобразования сантиметров в пункты (точки)
            pageSetup.LeftMargin = 1.5f * cmToPoints; // Левое поле (2.5 см)
            pageSetup.RightMargin = 1f * cmToPoints; // Правое поле (2.5 см)
            pageSetup.TopMargin = 1f * cmToPoints; // Верхнее поле (2.0 см)
            pageSetup.BottomMargin = 1f * cmToPoints; // Нижнее поле (2.0 см)

            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            // Copy the content from the source document to the target document
            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc1.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            // Copy the content from the source document to the target document
            foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc2.StoryRanges)
            {
                range.Copy();
                endRange.Paste();
            }
            endRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

            if (sortedScores.ElementAt(1).Value == sortedScores.ElementAt(2).Value)
            {
                string filePath3 = Path.GetFullPath(filePath + sortedScores.ElementAt(2).Key + ".docx");
                Microsoft.Office.Interop.Word.Document sourceDoc3 = wordApp.Documents.Open(filePath3);
                foreach (Microsoft.Office.Interop.Word.Range range in sourceDoc3.StoryRanges)
                {
                    range.Copy();
                    endRange.Paste();
                }
                sourceDoc3.Close();
            }

            targetDoc.Save();
            targetDoc.Close();
            // Закрыть источник документа
            sourceDoc1.Close();
            sourceDoc2.Close();
            // Закрыть приложение Word
            wordApp.Quit();
        }
    }
}
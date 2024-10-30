using DataParserProfTest.Repositories;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using DataParserProfTest.Model;
using Google.Apis.Sheets.v4.Data;

namespace DataParserProfTest.Core
{
    internal class TestParser
    {
        string spreadsheetId = "";
        string[] Scopes = { SheetsService.Scope.Spreadsheets };
        SheetsService service;
        ChartsCreator chartsCreator = new ChartsCreator();
        AnswersRepository answersRepository = new AnswersRepository();
        QuestionsRepository questionsRepository = new QuestionsRepository();
        TestsRepository testsRepository = new TestsRepository();

        public TestParser()
        {
            service = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets { ClientId = "977788122907-rrun02hhc225st0aaorhphs336jod6os.apps.googleusercontent.com", ClientSecret = "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" }, Scopes, "user", CancellationToken.None, new FileDataStore("myToken")).Result, ApplicationName = "OProfTest" });
        }

        private void DeleteRowFromSpreadsheet(Test test, int endIndex)
        {
            var deleteRequest = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = test.SheetID, // ID листа, который нужно получить из Spreadsheet
                        Dimension = "ROWS", // Удаляем строки
                        StartIndex = 5, // Индекс первой строки для удаления (включительно)
                        EndIndex = endIndex + 10 // Индекс после последней строки для удаления (не включительно)
                    }
                }
            };

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new[] { deleteRequest }
            };

            var batchUpdate = service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId);
            batchUpdate.Execute();
        }

        private void InsertValuesBrigs(Dictionary<string, int> scores)
        {
            scores.Add("E", 0);
            scores.Add("I", 0);
            scores.Add("S", 0);
            scores.Add("N", 0);
            scores.Add("T", 0);
            scores.Add("F", 0);
            scores.Add("J", 0);
            scores.Add("P", 0);
        }

        private void InsertValuesSocialType(Dictionary<string, int> scores)
        {
            scores.Add("1", 0);
            scores.Add("2", 0);
            scores.Add("3", 0);
            scores.Add("4", 0);
            scores.Add("5", 0);
            scores.Add("6", 0);
            scores.Add("7", 0);
            scores.Add("8", 0);
        }

        private void InsertValuesThinkingtype(Dictionary<int, int> scores)
        {
            scores.Add(1, 0);
            scores.Add(2, 0);
            scores.Add(3, 0);
            scores.Add(4, 0);
            scores.Add(5, 0);
        }

        private void InsertValuesProfType(Dictionary<string, int> scores)
        {
            scores.Add("A", 0);
            scores.Add("C", 0);
            scores.Add("I", 0);
            scores.Add("E", 0);
            scores.Add("R", 0);
            scores.Add("S", 0);
        }

        public void CheckTestsInterests()
        {
            Test test = testsRepository.GetTestById(1);
            var questions = questionsRepository.GetAllQuestionsByTestID(test.ID);
            var answers = answersRepository.GetAllAnswersByTestID(test.ID);
            var values = service.Spreadsheets.Values.Get(spreadsheetId, $"{test.Title}{test.SheetRange}").Execute().Values;
            var firstRow = values[0];
            int counter = 0;

            //count all rows at the Google spreadsheet until there is no values
            while (counter != values.Count - 1)
            {
                string date = values[counter + 1][0].ToString().Substring(0, 10);
                string email = values[counter + 1][1].ToString();
                string name = values[counter + 1][2].ToString().Replace(" ", "-");
                Dictionary<int, int> scores = new Dictionary<int, int>();
                var dataRow = values[counter + 1];

                for (int i = 3, k = 0; i < values[counter + 1].Count && k < questions.Count; i++, k++)
                {
                    if (!firstRow[i].ToString().Contains(questions[k].Title))
                        throw new Exception("Wrong sequence of questions!");
                    if (!scores.ContainsKey(questions[k].QuestionType))
                        scores.Add(questions[k].QuestionType, 0);
                    scores[questions[k].QuestionType] += answers.FirstOrDefault(p => p.Title.Equals(dataRow[i])).Value;
                }

                counter++;
                chartsCreator.CreateChartInterests(test, scores, email, date, name);
            }
            if (counter != 0)
                DeleteRowFromSpreadsheet(test, counter);
            Console.WriteLine("Кол-во обработанных тестов \"Интересы\": " + counter);
        }

        public void CheckTestsOrientation()
        {
            Test test = testsRepository.GetTestById(2);
            var questions = questionsRepository.GetAllQuestionsByTestID(test.ID);
            var answers = answersRepository.GetAllAnswersByTestID(test.ID);
            var values = service.Spreadsheets.Values.Get(spreadsheetId, $"{test.Title}{test.SheetRange}").Execute().Values;
            var firstRow = values[0];
            int counter = 0;

            //count all rows at the Google spreadsheet until there is no values
            while (counter != values.Count - 1)
            {
                string date = values[counter + 1][0].ToString().Substring(0, 10);
                string email = values[counter + 1][1].ToString();
                string name = values[counter + 1][2].ToString().Replace(" ", "-");
                Dictionary<int, int> scoresWant = new Dictionary<int, int>();
                Dictionary<int, int> scoresCan = new Dictionary<int, int>();
                var dataRow = values[counter + 1];

                for (int i = 3, k = 0; i < values[counter + 1].Count && k < questions.Count; i++, k++)
                {
                    if (!firstRow[i].ToString().Contains(questions[k].Title))
                        throw new Exception("Wrong sequence of questions!");
                    if (k < 35)
                    {
                        if (!scoresWant.ContainsKey(questions[k].QuestionType))
                            scoresWant.Add(questions[k].QuestionType, 0);
                        scoresWant[questions[k].QuestionType] += answers.FirstOrDefault(p => p.Title.Equals(dataRow[i])).Value;
                    }
                    else
                    {
                        if (!scoresCan.ContainsKey(questions[k].QuestionType))
                            scoresCan.Add(questions[k].QuestionType, 0);
                        scoresCan[questions[k].QuestionType] += answers.FirstOrDefault(p => p.Title.Equals(dataRow[i])).Value;
                    }
                }

                counter++;
                string path = chartsCreator.CreateChartOrientation(test, scoresWant, scoresCan, email, date, name);
                DocsMerger.ClearDoc(path);
                DocsMerger.MergeDocsOrientation(path, scoresWant, scoresCan);
            }
            if (counter != 0)
                DeleteRowFromSpreadsheet(test, counter);
            Console.WriteLine("Кол-во обработанных тестов \"Ориентация\": " + counter);
        }

        public void CheckTestsInlination()
        {
            Test test = testsRepository.GetTestById(3);
            var questions = questionsRepository.GetAllQuestionsByTestID(test.ID);
            var answers = answersRepository.GetAllAnswersByTestID(test.ID);
            var values = service.Spreadsheets.Values.Get(spreadsheetId, $"{test.Title}{test.SheetRange}").Execute().Values;
            var firstRow = values[0];
            int counter = 0;

            //count all rows at the Google spreadsheet until there is no values
            while (counter != values.Count - 1)
            {
                string date = values[counter + 1][0].ToString().Substring(0, 10);
                string email = values[counter + 1][1].ToString();
                string name = values[counter + 1][2].ToString().Replace(" ", "-");
                Dictionary<int, int> scores = new Dictionary<int, int>();
                var dataRow = values[counter + 1];

                for (int i = 3, k = 0; i < values[counter + 1].Count && k < questions.Count; i++, k++)
                {
                    if (!firstRow[i].ToString().Contains(questions[k].Title))
                        throw new Exception("Wrong sequence of questions!");
                    if (!scores.ContainsKey(questions[k].QuestionType))
                        scores.Add(questions[k].QuestionType, 0);
                    scores[questions[k].QuestionType] += answers.FirstOrDefault(p => p.Title.Equals(dataRow[i])).Value;
                }

                counter++;
                string path = chartsCreator.CreateChartInlination(test, scores, email, date, name);
                DocsMerger.MergeDocsInlination(path, scores);
            }
            if (counter != 0)
                DeleteRowFromSpreadsheet(test, counter);
            Console.WriteLine("Кол-во обработанных тестов \"Склонности\": " + counter);
        }

        public void CheckTestsBrigs()
        {
            Test test = testsRepository.GetTestById(4);
            var questions = questionsRepository.GetAllQuestionsByTestID(test.ID);
            var values = service.Spreadsheets.Values.Get(spreadsheetId, $"{test.Title}{test.SheetRange}").Execute().Values;
            var firstRow = values[0];
            int counter = 0;

            //count all rows at the Google spreadsheet until there is no values
            while (counter != values.Count - 1)
            {
                string date = values[counter + 1][0].ToString().Substring(0, 10);
                string email = values[counter + 1][1].ToString();
                string name = values[counter + 1][2].ToString().Replace(" ", "-");
                Dictionary<string, int> scores = new Dictionary<string, int>();
                InsertValuesBrigs(scores);
                var dataRow = values[counter + 1];

                for (int i = 3, k = 0; i < values[counter + 1].Count && k < questions.Count; i++, k++)
                {
                    if (!firstRow[i].ToString().Contains(questions[k].Title))
                        throw new Exception("Wrong sequence of questions!");
                    var answers = answersRepository.GetAllAnswersByQuestion(questions[k].QuestionType, test.ID);
                    Answer currentAnswer = answers.FirstOrDefault(p => p.Title.Equals(dataRow[i]));
                    if (scores.ContainsKey(currentAnswer.ValueKey))
                        scores[currentAnswer.ValueKey] += currentAnswer.Value;
                }

                counter++;
                DocsMerger.MergeDocsSocialBrigs(test, scores, email, date, name);
            }
            if (counter != 0)
                DeleteRowFromSpreadsheet(test, counter);
            Console.WriteLine("Кол-во обработанных тестов \"Бригс\": " + counter);
        }

        public void CheckTestsSocialType()
        {
            Test test = testsRepository.GetTestById(5);
            var values = service.Spreadsheets.Values.Get(spreadsheetId, $"{test.Title}{test.SheetRange}").Execute().Values;
            var firstRow = values[0];
            int counter = 0;

            //count all rows at the Google spreadsheet until there is no values
            while (counter != values.Count - 1)
            {
                string date = values[counter + 1][0].ToString().Substring(0, 10);
                string email = values[counter + 1][1].ToString();
                string name = values[counter + 1][2].ToString().Replace(" ", "-");
                Dictionary<string, int> scores = new Dictionary<string, int>();
                InsertValuesSocialType(scores);
                var dataRow = values[counter + 1];

                for (int i = 3, k = 1; i < values[counter + 1].Count; i++, k++)
                {
                    var answers = answersRepository.GetAllAnswersByType(k, test.ID);
                    Answer currentAnswer = answers.FirstOrDefault(p => p.Title.Equals(dataRow[i]));
                    if (scores.ContainsKey(currentAnswer.ValueKey))
                        scores[currentAnswer.ValueKey] += currentAnswer.Value;
                }

                counter++;
                DocsMerger.MergeDocsSocialBrigs(test, scores, email, date, name);
            }
            if (counter != 0)
                DeleteRowFromSpreadsheet(test, counter);
            Console.WriteLine("Кол-во обработанных тестов \"Соционический тип\": " + counter);
        }
        
        public void CheckTestsThinkingType()
        {
            Test test = testsRepository.GetTestById(6);
            var questions = questionsRepository.GetAllQuestionsByTestID(test.ID);
            var answers = answersRepository.GetAllAnswersByTestID(test.ID);
            var values = service.Spreadsheets.Values.Get(spreadsheetId, $"{test.Title}{test.SheetRange}").Execute().Values;
            var firstRow = values[0];
            int counter = 0;

            //count all rows at the Google spreadsheet until there is no values
            while (counter != values.Count - 1)
            {
                string date = values[counter + 1][0].ToString().Substring(0, 10);
                string email = values[counter + 1][1].ToString();
                string name = values[counter + 1][2].ToString().Replace(" ", "-");
                Dictionary<int, int> scores = new Dictionary<int, int>();
                InsertValuesThinkingtype(scores);
                var dataRow = values[counter + 1];

                for (int i = 3, k = 0; i < values[counter + 1].Count && k < questions.Count; i++, k++)
                {
                    if (!firstRow[i].ToString().Contains(questions[k].Title))
                        throw new Exception("Wrong sequence of questions!");
                    if (!scores.ContainsKey(questions[k].QuestionType))
                        scores.Add(questions[k].QuestionType, 0);
                    scores[questions[k].QuestionType] += answers.FirstOrDefault(p => p.Title.Equals(dataRow[i])).Value;
                }

                counter++;
                string path = chartsCreator.CreateChartThinkingType(test, scores, email, date, name);
                DocsMerger.MergeDocsThinkingType(path, scores);
            }
            if (counter != 0)
                DeleteRowFromSpreadsheet(test, counter);
            Console.WriteLine("Кол-во обработанных тестов \"Тип мышления\": " + counter);
        }

        public void CheckTestsProfType()
        {
            Test test = testsRepository.GetTestById(7);
            var values = service.Spreadsheets.Values.Get(spreadsheetId, $"{test.Title}{test.SheetRange}").Execute().Values;
            var firstRow = values[0];
            int counter = 0;

            //count all rows at the Google spreadsheet until there is no values
            while (counter != values.Count - 1)
            {
                string date = values[counter + 1][0].ToString().Substring(0, 10);
                string email = values[counter + 1][1].ToString();
                string name = values[counter + 1][2].ToString().Replace(" ", "-");
                Dictionary<string, int> scores = new Dictionary<string, int>();
                InsertValuesProfType(scores);
                var dataRow = values[counter + 1];

                for (int i = 3, k = 1; i < values[counter + 1].Count; i++, k++)
                {
                    var answers = answersRepository.GetAllAnswersByType(k, test.ID);
                    Answer currentAnswer = answers.FirstOrDefault(p => p.Title.Equals(dataRow[i]));
                    if (scores.ContainsKey(currentAnswer.ValueKey))
                        scores[currentAnswer.ValueKey] += currentAnswer.Value;
                }

                counter++;
                string path = chartsCreator.CreateChartProfType(test, scores, email, date, name);
                DocsMerger.MergeDocsProfType(path, scores);
            }
            if (counter != 0)
                DeleteRowFromSpreadsheet(test, counter);
            Console.WriteLine("Кол-во обработанных тестов \"Проф.тип\": " + counter);
        }
    }
}
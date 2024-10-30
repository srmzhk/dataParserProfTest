using DataParserProfTest.Model;

namespace DataParserProfTest.Repositories
{
    internal class QuestionsRepository
    {
        private readonly AppDbContext _appDbContext;
        public QuestionsRepository()
        {
            _appDbContext = new AppDbContext();
        }

        public List<Question> GetAllQuestions()
        {
            return _appDbContext.Questions.ToList();
        }

        public List<Question> GetAllQuestionsByTestID(int testID)
        {
            return _appDbContext.Questions.Where(p => p.TestID == testID).ToList();
        }

        public Question GetQuestionById(int QuestionId)
        {
            return _appDbContext.Questions.Find(QuestionId);
        }

        public Question GetQuestionByTitle(string title)
        {
            return _appDbContext.Questions.FirstOrDefault(p => p.Title == title);
        }
    }
}

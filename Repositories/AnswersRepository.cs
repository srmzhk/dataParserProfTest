using DataParserProfTest.Model;

namespace DataParserProfTest.Repositories
{
    internal class AnswersRepository
    {
        private readonly AppDbContext _appDbContext;
        public AnswersRepository()
        {
            _appDbContext = new AppDbContext();
        }

        public List<Answer> GetAllAnswersByTestID(int id)
        {
            return _appDbContext.Answers.Where(p => p.TestID == id).ToList();
        }

        public List<Answer> GetAllAnswersByQuestion(int qt, int testID)
        {
            return _appDbContext.Answers.Where(p => p.AnswerType == qt && p.TestID == testID).ToList();
        }

        public List<Answer> GetAllAnswersByType(int type, int testID)
        {
            return _appDbContext.Answers.Where(p => p.AnswerType == type && p.TestID == testID).ToList();
        }

        public bool ContainsAnswerByTitle(string title, int testID)
        {
            return _appDbContext.Answers.Any(p => p.Title == title && p.TestID == testID);
        }

        public Answer GetAnswerById(int answerId)
        {
            return _appDbContext.Answers.Find(answerId);
        }
    }
}

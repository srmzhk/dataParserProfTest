using DataParserProfTest.Model;

namespace DataParserProfTest.Repositories
{
    internal class TestsRepository
    {
        private readonly AppDbContext _appDbContext;
        public TestsRepository()
        {
            _appDbContext = new AppDbContext();
        }

        public List<Test> GetAllTests()
        {
            return _appDbContext.Tests.ToList();
        }

        public Test GetTestById(int TestId)
        {
            return _appDbContext.Tests.Find(TestId);
        }

        public Test GetTestByTitle(string TestTitle)
        {
            return _appDbContext.Tests.Where(p => p.Title.Equals(TestTitle)).FirstOrDefault();
        }

        public Test GetTestBySheetId(int SheetID)
        {
            return _appDbContext.Tests.Where(p => p.SheetID.Equals(SheetID)).FirstOrDefault();
        }
    }
}
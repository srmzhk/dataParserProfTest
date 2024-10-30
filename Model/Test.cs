using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace DataParserProfTest.Model
{
    [Serializable]
    [Table("Tests")]
    public class Test
    {
        [Key]
        public int ID { get; set; }

        [StringLength(100)]
        public string Title { get; set; }

        public int SheetID { get; set; }

        [StringLength(10)]
        public string SheetRange { get; set; }

        public IEnumerable<Answer> Answers { get; set; }

        public IEnumerable<Question> Questions { get; set; }
    }
}

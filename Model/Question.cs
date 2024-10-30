using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace DataParserProfTest.Model
{
    [Serializable]
    [Table("Questions")]
    public class Question
    {
        [Key]
        public int ID { get; set; }

        [ForeignKey("Test")]
        public int TestID { get; set; }
        public virtual Test Test { get; set; }

        public int QuestionType { get; set; }

        public string Title { get; set; }
    }
}
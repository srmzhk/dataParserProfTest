using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace DataParserProfTest.Model
{
    [Serializable]
    [Table("Answers")]
    public class Answer
    {
        [Key]
        public int ID { get; set; }

        [ForeignKey("Test")]
        public int TestID { get; set; }
        public virtual Test Test { get; set; }

        public int AnswerType { get; set; }

        public string Title { get; set; }

        [StringLength(50)]
        public string ValueKey { get; set; }

        public int Value { get; set; }
    }
}

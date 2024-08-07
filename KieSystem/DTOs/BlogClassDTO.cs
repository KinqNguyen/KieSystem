namespace KieSystem.DTOs
{
    public class BlogDTO
    {
        public int Id { get; set; }
        public int UserId { get; set; }
        public int Userlevel { get; set; }
        public DateTime Releasedate { get; set; }
        public int  Totalview { get; set; }
        public string Title { get; set; }
        public string Body { get; set; }
        public string BonusTotal { get; set; }
}
}

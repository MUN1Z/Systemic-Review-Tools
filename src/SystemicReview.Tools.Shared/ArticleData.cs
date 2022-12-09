namespace SystemicReview.Tools.Shared
{
    public class ArticleData
    {
        public string Title { get; set; }
        public string Abstract { get; set; }
        public string Url { get; set; }

        public override bool Equals(object obj)
        {
            // If the passed object is null, return False
            if (obj == null)
                return false;

            // If the passed object is not Customer Type, return False
            if (!(obj is ArticleData))
                return false;

            return Title == ((ArticleData)obj).Title;
        }
    }
}
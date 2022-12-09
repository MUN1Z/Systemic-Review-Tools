using MiniExcelLibs;

namespace SystemicReview.Tools.Shared
{
    public class ExcelManager
    {
        private static ExcelManager _instance;

        public static ExcelManager Instance()
        {
            if (_instance == null)
                _instance = new ExcelManager();

            return _instance;
        }

        public IEnumerable<ArticleData> ReadExcelFile(string path)
            => MiniExcel.Query<ArticleData>(path).Where(c => c.Abstract != null);

        public void CreateExcelFile(string path, IEnumerable<ArticleData> articles)
            => MiniExcel.SaveAs(path, articles);

    }
}

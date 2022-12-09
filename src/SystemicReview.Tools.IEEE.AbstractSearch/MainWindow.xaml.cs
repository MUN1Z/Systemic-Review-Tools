using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using System.Windows;
using SystemicReview.Tools.Shared;
using Expression = System.Linq.Expressions.Expression;

namespace SystemicReview.Tools.IEEE.AbstractSearch
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            UpdateResult();
        }

        int _articlesFoundedCount = 0;
        int _articlesLoadedCount = 0;
        int _articlesMatchedCount = 0;
        int _articlesExcludedCount = 0;

        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txbSearchString.Text))
            {
                var articlesThread = new Thread(() =>
                {
                    App.Current.Dispatcher.Invoke(() =>
                    {
                        btnSearch.IsEnabled = false;
                    });

                    var data = GetArticles();
                    FilterArticles(data);

                    App.Current.Dispatcher.Invoke(() =>
                    {
                        btnSearch.IsEnabled = true;
                        MessageBoxResult result = MessageBox.Show("Finished, Excel files generated in application folder!!");
                    });
                });

                articlesThread.Start();
            }
        }

        private IEnumerable<ArticleData> GetArticles()
        {
            var data = new List<ArticleData>();
            var driver = new FirefoxDriver();
            var articlesHrefs = new List<string>();

            var defaultUrl = string.Empty;
            App.Current.Dispatcher.Invoke(() =>
            {
                var ranges = string.Empty;

                if (!string.IsNullOrEmpty(txbInitialYear.Text) && !string.IsNullOrEmpty(txbEndYear.Text))
                    ranges = $"&ranges={txbInitialYear.Text}_{txbEndYear.Text}_Year";

                defaultUrl = $"https://ieeexplore.ieee.org/search/searchresult.jsp?queryText={txbSearchString.Text}{ranges}";
            });

            driver.Navigate().GoToUrl(defaultUrl);

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            while (true)
            {
                wait.Until(driver => driver.FindElement(By.ClassName("List-results-items")));

                var articleListItemsResults = driver.FindElements(By.ClassName("List-results-items"));

                foreach (var articleListItem in articleListItemsResults)
                {
                    var idValue = string.Empty;
                    App.Current.Dispatcher.Invoke(() =>
                    {
                        idValue = articleListItem.GetAttribute("id");
                    });

                    var href = $"https://ieeexplore.ieee.org/document/{idValue}/";

                    if (articlesHrefs.Contains(href))
                        break;

                    articlesHrefs.Add(href);

                    _articlesFoundedCount++;
                    UpdateResult();
                }

                try
                {
                    wait.Until(driver => driver.FindElement(By.ClassName("next-btn")));

                    var queryNextButton = driver.FindElement(By.ClassName("next-btn"));

                    App.Current.Dispatcher.Invoke(() =>
                    {
                        driver.ExecuteScript("scroll(0,100000)");
                    });

                    if (queryNextButton != null)
                        App.Current.Dispatcher.Invoke(() =>
                        {
                            queryNextButton.Click();
                        });
                    else
                        break;
                }
                catch
                {
                    break;
                }
            }

            foreach (var href in articlesHrefs)
            {
                driver.Navigate().GoToUrl(href);

                wait.Until(driver => driver.FindElement(By.ClassName("u-mb-1")));

                var titleElement = driver.FindElement(By.ClassName("text-2xl-md-lh"));

                var titleValue = string.Empty;
                App.Current.Dispatcher.Invoke(() =>
                {
                    titleValue = titleElement.Text;
                });

                var abstractElement = driver.FindElement(By.ClassName("abstract-text"));
                var abstractElementClass = abstractElement.FindElement(By.ClassName("u-mb-1"));
                var abstractElementDiv = abstractElementClass.FindElement(By.TagName("div"));

                var abstractValue = string.Empty;
                App.Current.Dispatcher.Invoke(() =>
                {
                    abstractValue = abstractElementDiv.Text;
                });

                data.Add(new ArticleData
                {
                    Url = href,
                    Title = titleValue,
                    Abstract = abstractValue
                });

                _articlesLoadedCount++;

                UpdateResult();
            }

            ExcelManager.Instance().CreateExcelFile($"AllArticlesFounded-{Guid.NewGuid()}.xlsx", data);

            return data;
        }

        private void FilterArticles(IEnumerable<ArticleData> data)
        {
            var searchString = string.Empty;

            App.Current.Dispatcher.Invoke(() =>
            {
                searchString = txbSearchString.Text;
            });

            if (searchString.StartsWith("((") && searchString.EndsWith("))"))
                searchString = searchString.Replace("((", "(").Replace("))", ")");

            var andsConditions = searchString.Split(" AND ");

            Expression<Func<ArticleData, bool>> predicate = null;

            foreach (var andCondition in andsConditions)
            {
                Expression<Func<ArticleData, bool>> predicateAnd = null;

                var andConditionOrs = andCondition.Replace("(", "").Replace(")", "").Replace("\"", "").Split(" OR ");

                foreach (var andConditionOr in andConditionOrs)
                {
                    var entityProperties = typeof(ArticleData).GetProperties();
                    var prop = entityProperties.FirstOrDefault(c => c.Name.Equals("Abstract"));

                    var parameter = Expression.Parameter(typeof(ArticleData), "f");
                    var propertyAccess = Expression.MakeMemberAccess(parameter, prop);

                    var indexOf = Expression.Call(
                        propertyAccess,
                        "IndexOf",
                        null,
                        Expression.Constant(andConditionOr.Trim(), typeof(string)),
                        Expression.Constant(StringComparison.CurrentCultureIgnoreCase));

                    var like = Expression.GreaterThanOrEqual(indexOf, Expression.Constant(0));
                    var predicateToOR = Expression.Lambda<Func<ArticleData, bool>>(like, parameter);

                    if (predicateAnd == null)
                        predicateAnd = predicateToOR;
                    else
                        predicateAnd = predicateAnd.Or(predicateToOR);
                }

                if (predicate == null)
                    predicate = predicateAnd;
                else
                    predicate = predicate.And(predicateAnd);
            }

            var dataMatched = data.AsQueryable().Where(predicate).ToList();

            var titlesExcluded = data.Select(c => c.Title).Except(dataMatched.Select(c => c.Title));
            var dataExcluded = data.Where(c => titlesExcluded.Contains(c.Title)).ToList();

            _articlesMatchedCount = dataMatched.Count();
            _articlesExcludedCount = dataExcluded.Count();

            UpdateResult();

            ExcelManager.Instance().CreateExcelFile($"ArticlesMatched-{Guid.NewGuid()}.xlsx", dataMatched);
            ExcelManager.Instance().CreateExcelFile($"ArticlesExcluded-{Guid.NewGuid()}.xlsx", dataExcluded);
        }

        private void UpdateResult()
        {
            App.Current.Dispatcher.Invoke(() =>
            {
                lblResult.Content = $"Founded: {_articlesFoundedCount} - Loaded: {_articlesLoadedCount} \n" +
                                    $"Matched: {_articlesMatchedCount}, Excluded: {_articlesExcludedCount}";
            });
        }
    }
}

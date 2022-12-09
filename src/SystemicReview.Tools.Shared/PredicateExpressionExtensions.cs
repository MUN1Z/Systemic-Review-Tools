using System.Linq.Expressions;

namespace SystemicReview.Tools.Shared
{
    public static class PredicateExpressionExtensions
    {
        #region public methods implementations

        public static Expression<Func<T, bool>> And<T>(this Expression<Func<T, bool>> a, Expression<Func<T, bool>> b)
        {

            ParameterExpression p = a.Parameters[0];

            SubstExpressionVisitor visitor = new SubstExpressionVisitor();
            visitor.Subst[b.Parameters[0]] = p;

            Expression body = Expression.AndAlso(a.Body, visitor.Visit(b.Body));
            return Expression.Lambda<Func<T, bool>>(body, p);
        }

        public static Expression<Func<T, bool>> Or<T>(this Expression<Func<T, bool>> a, Expression<Func<T, bool>> b)
        {
            ParameterExpression p = a.Parameters[0];

            SubstExpressionVisitor visitor = new SubstExpressionVisitor();
            visitor.Subst[b.Parameters[0]] = p;

            Expression body = Expression.OrElse(a.Body, visitor.Visit(b.Body));
            return Expression.Lambda<Func<T, bool>>(body, p);
        }

        #endregion
    }

    internal class SubstExpressionVisitor : ExpressionVisitor
    {
        #region properties

        public Dictionary<Expression, Expression> Subst = new Dictionary<Expression, Expression>();

        #endregion

        #region protected methods implementations

        protected override Expression VisitParameter(ParameterExpression node)
        {
            Expression newValue;

            if (Subst.TryGetValue(node, out newValue))
                return newValue;

            return node;
        }

        #endregion
    }
}

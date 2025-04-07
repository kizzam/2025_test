//Repository Pattern
//A generic repository interface that we can use on the server to access a database
//and on the client to access an API or service calls to the server
//

namespace RP_2023.Data
{
    public interface IRepository<TEntity> where TEntity : class
    {
        Task<IEnumerable<TEntity>> GetAllAsync();
        Task<IEnumerable<TEntity>> GetAsync(string query);
        Task<IEnumerable<TEntity>> GetAsync(
            Expression<Func<TEntity, bool>> filter = null,
            Func<IQueryable<TEntity>, IOrderedQueryable<TEntity>> orderBy = null,
            string includeProperties = "");
        Task<TEntity> InsertAsync(TEntity entity);
        Task<TEntity> UpdateAsync(TEntity entityToUpdate);
        Task<bool> DeleteAsync(TEntity entityToDelete);
    }
}

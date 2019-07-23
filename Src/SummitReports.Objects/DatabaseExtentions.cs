using FastMember;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SummitReports.Objects
{
    public static class DatabaseExtensions
    {
        ////Thank you http://stackoverflow.com/questions/3511780/system-linq-dynamic-and-sql-in-operator
        //public static IQueryable<TEntity> WhereIn<TEntity, TValue>
        //(
        //    this ObjectQuery<TEntity> query,
        //    Expression<Func<TEntity, TValue>> selector,
        //    IEnumerable<TValue> collection
        //)
        //{
        //    if (selector == null) throw new ArgumentNullException("selector");
        //    if (collection == null) throw new ArgumentNullException("collection");
        //    ParameterExpression p = selector.Parameters.Single();

        //    if (!collection.Any()) return query;

        //    IEnumerable<Expression> equals = collection.Select(value =>
        //      (Expression)Expression.Equal(selector.Body,
        //       Expression.Constant(value, typeof(TValue))));

        //    Expression body = equals.Aggregate((accumulate, equal) =>
        //    Expression.Or(accumulate, equal));

        //    return query.Where(Expression.Lambda<Func<TEntity, bool>>(body, p));
        //}

        static Regex underscore = new Regex(@"(^|_)(.)");
        static string convertName(string s)
        {
            return underscore.Replace(s.ToLower(), m => m.Groups[0].ToString().ToUpper().Replace("_", ""));
        }

        static T ToObject<T>(this IDataRecord r) where T : new()
        {
            var indexMembers = new Dictionary<string, Member>();
            T obj = new T();
            var accessor = TypeAccessor.Create(typeof(T));
            var members = accessor.GetMembers();
            foreach (var item in members) indexMembers.Add(item.Name, item);

            for (int i = 0; i < r.FieldCount; i++)
            {
                var name = r.GetName(i);
                if (r.GetName(i).Equals("Amendment"))
                {
                    //var k = 10;
                }
                if (indexMembers.ContainsKey(name))
                {
                    var PropertyType = indexMembers[name].Type;
                    if (PropertyType == r[i].GetType())
                        accessor[obj, name] = r[i];
                    else
                    {
                        if (PropertyType.GenericTypeArguments.Contains(r[i].GetType()))
                        {
                            accessor[obj, name] = r[i];
                        }
                        else
                        {
                            var c = TypeDescriptor.GetConverter(r[i]);
                            if (c.CanConvertTo(PropertyType))
                                accessor[obj, name] = c.ConvertTo(r[i], PropertyType);
                        }
                    }

                }
            }
            return obj;
        }

        static string DRToJson(this IDataRecord r)
        {
            JObject returnJson = new JObject();
            for (int i = 0; i < r.FieldCount; i++)
            {
                var name = r.GetName(i);
                var c = TypeDescriptor.GetConverter(r[i]);
                if (c.CanConvertTo(r[i].GetType()))
                    returnJson.Add(new JProperty(name, r[i]));
            }
            return returnJson.ToString();
        }


        public static string ToJson(this IDataReader r)
        {
            var sb = new StringBuilder();
            sb.Append("[");
            var rsCount = 0;
            do
            {
                var recCount = 0;
                sb.Append(((rsCount > 0) ? ",[" : "["));
                while (r.Read())
                {
                    if (recCount > 0) sb.Append(",");
                    sb.Append(r.DRToJson());
                    recCount++;
                }
                sb.Append("]");
                rsCount++;
            } while (r.NextResult());
            sb.Append("]");
            return sb.ToString();
        }

        public static IEnumerable<T> GetObjects<T>(this IDbCommand c) where T : new()
        {
            using (IDataReader r = c.ExecuteReader())
            {
                while (r.Read())
                {
                    yield return r.ToObject<T>();
                }
            }
        }

        public static IEnumerable<T> GetObjects<T>(this IDataReader r) where T : new()
        {
            while (r.Read())
            {
                yield return r.ToObject<T>();
            }
        }

        public static IEnumerable<T> GetObjectsNextRS<T>(this IDataReader r) where T : new()
        {
            if (r.NextResult())
            {
                while (r.Read())
                {
                    yield return r.ToObject<T>();
                }
            }
        }
    }

    public static class MarsDb
    {
        private static string defaultConnection = SummitReportSettings.Instance.ConnectionString;

        public static async Task<SqlConnection> ConnectionAsync()
        {
            var connection = new SqlConnection(defaultConnection);
            await connection.OpenAsync();
            return connection;
        }


        public static SqlConnection Connection()
        {
            var connection = new SqlConnection(defaultConnection);
            connection.Open();
            return connection;
        }
        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Query<T>(string sql) where T : new()
        {
            return await Query<T>(sql, new List<SqlParameter>());
        }

        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static string QueryAsJson(string sql)
        {
            return QueryAsJson(sql, new List<SqlParameter>());
        }

        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static string QueryAsJson(string sql, List<SqlParameter> parmList)
        {
            var retList = "";
            using (var conn = MarsDb.Connection())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.CommandType = CommandType.Text;
                    foreach (var parm in parmList)
                    {
                        cmd.Parameters.Add(parm);
                    }
                    using (var reader = cmd.ExecuteReader())
                    {
                        retList = reader.ToJson();
                    }
                }
            }
            return retList;
        }

        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Query<T>(string sql, Dictionary<string, object> parmList) where T : new()
        {
            var sqlParmList = new List<SqlParameter>();
            foreach (var parmName in parmList.Keys)
                sqlParmList.Add(new SqlParameter(parmName, parmList[parmName]));
            return await Query<T>(sql, sqlParmList);
        }


        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <param name="parmList">Will assume the parameter names is </param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Query<T>(string sql, List<object> parmList) where T : new()
        {
            var sqlParmList = new List<SqlParameter>();
            var i = 0;
            foreach (var val in parmList)
                sqlParmList.Add(new SqlParameter(string.Format("p{0}", i), val));
            return await Query<T>(sql, sqlParmList);
        }

        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <param name="parmList">Will assume the parameter names is </param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Query<T>(string sql, params object[] parmList) where T : new()
        {
            var sqlParmList = new List<SqlParameter>();
            var i = 0;
            foreach (var val in parmList)
            {
                sqlParmList.Add(new SqlParameter(string.Format("p{0}", i), val));
                i++;
            }
            return await Query<T>(sql, sqlParmList);
        }


        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Query<T>(string sql, List<SqlParameter> parmList) where T : new()
        {
            var retList = new List<T>();
            using (var conn = await MarsDb.ConnectionAsync())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.CommandType = CommandType.Text;
                    foreach (var parm in parmList)
                    {
                        cmd.Parameters.Add(parm);
                    }
                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        retList = reader.GetObjects<T>().ToList();
                    }
                }
            }
            return retList;
        }

        /// <summary>
        /// Execute a qeury and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <param name="parmList">Will assume the parameter names is </param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Procedure<T>(string sql, params object[] parmList) where T : new()
        {
            var sqlParmList = new List<SqlParameter>();
            var i = 0;
            foreach (var val in parmList)
            {
                sqlParmList.Add(new SqlParameter(string.Format("p{0}", i), val));
                i++;
            }
            return await Procedure<T>(sql, sqlParmList);
        }

        /// <summary>
        /// Execute a stored procedure and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Procedure<T>(string sql, List<SqlParameter> parmList) where T : new()
        {
            var retList = new List<T>();
            using (var conn = await MarsDb.ConnectionAsync())
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.CommandType = CommandType.StoredProcedure;
                    foreach (var parm in parmList)
                    {
                        cmd.Parameters.Add(parm);
                    }
                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        retList = reader.GetObjects<T>().ToList();
                    }
                }
            }
            return retList;
        }

        /// <summary>
        /// Execute a stored procedure and have it return the result set as an objects
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static async Task<IEnumerable<T>> Procedure<T>(string sql) where T : new()
        {
            return await Procedure<T>(sql, new List<SqlParameter>());
        }
    }
}
